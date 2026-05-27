"""
Launcher & Image Router for Studio Vision (Access).
1. Setup: Reads original .lnk shortcut (target, args, cwd) and saves to config.ini.
2. Launch: Recreates the exact launch sequence via subprocess.Popen.
3. Routing: Runs a background watchdog pipeline (Observer -> Queue -> Worker) to 
   process images and update the Access DB (with 1.5s debounce & auto-reconnect).
"""
import os
import sys
import time
import queue
import shutil
import ctypes
import logging
import threading
import subprocess
import pythoncom
import configparser
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from datetime import datetime
from pathlib import Path
from watchdog.observers.polling import PollingObserver as Observer
from watchdog.events import FileSystemEventHandler

try:
    import win32com.client
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
try:
    import pyodbc
    PYODBC_AVAILABLE = True
except ImportError:
    PYODBC_AVAILABLE = False
try:
    import pystray
    from PIL import Image, ImageDraw
    TRAY_AVAILABLE = True
except ImportError:
    TRAY_AVAILABLE = False
import win32api
import win32event
import winerror

# Logging: file + stdout, UTF-8
_LOG_DIR  = Path(os.path.expanduser("~")) / "studiovision"
_LOG_DIR.mkdir(parents=True, exist_ok=True)
_LOG_FILE = _LOG_DIR / "image_router.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  [%(threadName)s]  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(_LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("image_router")

def get_db_path_from_com() -> "Path | None":
    """Return the absolute network path to the Studio Vision backend DB.
    Reads the .Connect property of the linked 'Documents' table via COM.
    Falls back to CurrentDb().Name (frontend path) if .Connect is unavailable."""
    if not WIN32_AVAILABLE:
        log.error("win32com non disponible — détection COM impossible.")
        return None
    try:
        access = win32com.client.GetActiveObject("Access.Application")
        db     = access.CurrentDb()
        try:
            connect_string = db.TableDefs("Documents").Connect
            if connect_string and "DATABASE=" in connect_string.upper():
                upper = connect_string.upper()
                idx   = upper.index("DATABASE=") + len("DATABASE=")
                db_path = Path(connect_string[idx:].strip())
                log.info(f"Backend DB détecté via .Connect : {db_path}")
                return db_path
            else:
                log.warning(
                    f"Propriété Connect présente mais sans 'DATABASE=' : "
                    f"'{connect_string}' — fallback sur CurrentDb().Name"
                )
        except Exception as e_connect:
            log.warning(
                f"Impossible de lire .Connect sur la table 'Documents' "
                f"({e_connect}) — fallback sur CurrentDb().Name"
            )
        db_name = db.Name
        db_path = Path(db_name)
        log.info(f"Base de données détectée via COM (fallback Frontend) : {db_path}")
        return db_path
    except Exception as e:
        log.debug(f"COM get_db_path échoué : {e}")
        return None

def lire_raccourci_lnk(lnk_path: str) -> "tuple[str, str, str] | tuple[None, None, None]":
    """Parse a Windows .lnk shortcut via WScript.Shell.
    Returns (target, args, cwd) or (None, None, None) on failure."""
    if not WIN32_AVAILABLE:
        log.error("win32com non disponible — lecture du raccourci impossible.")
        return None, None, None
    try:
        shell    = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(lnk_path)
        target   = shortcut.TargetPath      or ""
        args     = shortcut.Arguments       or ""
        cwd      = shortcut.WorkingDirectory or ""
        log.info(f"Raccourci .lnk lu : {lnk_path}")
        log.info(f"  TargetPath       : {target}")
        log.info(f"  Arguments        : {args}")
        log.info(f"  WorkingDirectory : {cwd}")
        return target, args, cwd
    except Exception as e:
        log.error(f"Lecture du raccourci .lnk échouée ({lnk_path}) : {e}")
        return None, None, None

def auto_detect_sv_shortcut(frontend_path: str) -> "str | None":
    """Scan user and public desktops for the .lnk that launches this specific frontend.
    Returns the absolute shortcut path only if exactly one candidate is found."""
    if not WIN32_AVAILABLE:
        return None

    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        user_desktop = shell.SpecialFolders("Desktop")
        public_desktop = shell.SpecialFolders("AllUsersDesktop")
        
        search_dirs = [user_desktop]
        if public_desktop and public_desktop not in search_dirs:
            search_dirs.append(public_desktop)

        candidates = []
        frontend_lower = str(frontend_path).lower()

        log.info(f"Auto-détection du raccourci pour : {frontend_path}")

        for directory in search_dirs:
            if not directory or not os.path.exists(directory):
                continue
            for filename in os.listdir(directory):
                if filename.lower().endswith(".lnk"):
                    lnk_path = os.path.join(directory, filename)
                    try:
                        shortcut = shell.CreateShortcut(lnk_path)
                        target = str(shortcut.TargetPath or "").lower()
                        args = str(shortcut.Arguments or "").lower()
                        
                        if frontend_lower in target or frontend_lower in args:
                            candidates.append(lnk_path)
                            log.info(f"Candidat trouvé : {lnk_path}")
                    except Exception as e:
                        log.debug(f"Erreur lecture .lnk {lnk_path}: {e}")

        if len(candidates) == 1:
            log.info(f"Succès auto-détection: {candidates[0]}")
            return candidates[0]
        elif len(candidates) > 1:
            log.warning("Multiples candidats trouvés, auto-détection annulée.")
        else:
            log.info("Aucun candidat trouvé sur les bureaux.")
            
        return None

    except Exception as e:
        log.warning(f"Erreur lors de l'auto-détection: {e}")
        return None


def _get_real_user_desktop() -> str:
    """Return the desktop path of the actual logged-in user, even when running as admin.
    Priority: USERPROFILE env var > WScript.Shell SpecialFolders("Desktop")."""
    user_profile = os.environ.get("USERPROFILE", "")
    if user_profile:
        candidate = Path(user_profile) / "Desktop"
        if candidate.exists():
            log.info(f"Bureau utilisateur réel (USERPROFILE) : {candidate}")
            return str(candidate)
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        desktop = shell.SpecialFolders("Desktop")
        log.info(f"Bureau via WScript.Shell : {desktop}")
        return desktop
    except Exception as e:
        log.warning(f"Impossible de résoudre le bureau via WScript : {e}")
        return str(Path.home() / "Desktop")

def create_desktop_shortcut(target_exe: Path) -> None:
    """Create the 'Studio Vision - Connected' desktop shortcut pointing to this exe.
    Always targets the real user's desktop, even when the process is elevated (admin)."""
    if not WIN32_AVAILABLE:
        log.warning("win32com non disponible — création du raccourci ignorée.")
        return
    try:
        shell      = win32com.client.Dispatch("WScript.Shell")
        desktop    = _get_real_user_desktop()
        lnk_path   = os.path.join(desktop, "Studio Vision - Connected.lnk")
        shortcut   = shell.CreateShortcut(lnk_path)
        shortcut.TargetPath       = str(target_exe)
        shortcut.WorkingDirectory = str(target_exe.parent)
        shortcut.Description      = "Lance Studio Vision avec le routeur d'images intégré"
        if ICON_PATH.exists():
            shortcut.IconLocation = str(ICON_PATH)
        else:
            shortcut.IconLocation = f"{target_exe}, 0"
        shortcut.save()
        log.info(f"Raccourci Bureau créé : {lnk_path}")
    except Exception as e:
        log.error(f"Impossible de créer le raccourci Bureau : {e}")

def configurer_via_interface(config_path: Path) -> None:
    """First-run setup wizard (silent auto-detection + minimal user prompts).
    Detects backend DB, .lnk shortcut, Photos folder, and source camera folder.
    Writes config.ini and creates the desktop shortcut, then exits."""
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    # Detect backend DB and frontend path via COM (requires Access to be open)
    backend_mdb = get_db_path_from_com()
    if backend_mdb is None:
        messagebox.showerror(
            "Studio Vision introuvable",
            "Pour terminer l'installation, suivez ces étapes :\n\n"
            "  1. Ouvrez Studio Vision en tant qu'administrateur\n"
            "     (clic droit sur le raccourci → Exécuter en tant qu'administrateur).\n\n"
            "  2. Relancez le fichier  INSTALLATION_AUTOMATIQUE.bat\n"
        )
        sys.exit(1)

    frontend_path = None
    try:
        access = win32com.client.GetActiveObject("Access.Application")
        frontend_path = access.CurrentDb().Name
    except Exception:
        log.warning("Impossible de lire CurrentDb().Name pour l'auto-détection.")

    # Try to auto-detect .lnk; fall back to manual selection
    sv_target, sv_args, sv_cwd = None, None, None
    lnk_path = None
    
    if frontend_path:
        lnk_path = auto_detect_sv_shortcut(frontend_path)

    if lnk_path:
        sv_target, sv_args, sv_cwd = lire_raccourci_lnk(lnk_path)
        
    if not sv_target:
        messagebox.showinfo(
            "Raccourci introuvable",
            "La détection automatique du raccourci a échoué.\n\n"
            "Veuillez sélectionner votre raccourci 'Studio Vision' habituel."
        )
        lnk_path_manual = filedialog.askopenfilename(
            title="Sélectionnez le raccourci Studio Vision",
            filetypes=[("Raccourci Windows", "*.lnk")],
        )
        if not lnk_path_manual:
            messagebox.showerror(
                "Installation annulée",
                "Aucun raccourci sélectionné.\nRelancez l'installation pour recommencer."
            )
            sys.exit(1)
            
        sv_target, sv_args, sv_cwd = lire_raccourci_lnk(lnk_path_manual)
        if not sv_target:
            messagebox.showerror(
                "Raccourci illisible",
                "Impossible de lire le raccourci.\nRelancez l'installation."
            )
            sys.exit(1)

    log.info(f"SV_TARGET : {sv_target}")
    log.info(f"SV_ARGS   : {sv_args}")
    log.info(f"SV_CWD    : {sv_cwd or '(vide)'}")

    # Locate Photos folder: check one and two levels above backend DB
    _candidate1 = backend_mdb.parent / "Photos"
    _candidate2 = backend_mdb.parent.parent / "Photos"
    if _candidate1.is_dir():
        dest_photos = _candidate1
        log.info(f"DEST_PHOTOS détecté (niveau 1) : {dest_photos}")
    elif _candidate2.is_dir():
        dest_photos = _candidate2
        log.info(f"DEST_PHOTOS détecté (niveau 2) : {dest_photos}")
    else:
        log.warning("Dossier Photos introuvable. Demande manuelle à l'utilisateur.")
        messagebox.showinfo(
            "Dossier Photos introuvable",
            "Le dossier 'Photos' de Studio Vision n'a pas pu être détecté automatiquement.\n\n"
            "Cliquez sur OK, puis sélectionnez ce dossier manuellement."
        )
        _photos_manual = filedialog.askdirectory(
            title="Sélectionnez le dossier Photos de Studio Vision"
        )
        if not _photos_manual:
            messagebox.showerror(
                "Installation annulée",
                "Aucun dossier Photos sélectionné.\nRelancez l'installation."
            )
            sys.exit(1)
        dest_photos = Path(_photos_manual)
        log.info(f"DEST_PHOTOS sélectionné manuellement : {dest_photos}")

    # Only mandatory user interaction: pick the camera/source folder
    source_dir = filedialog.askdirectory(
        title="Dernière étape : Sélectionnez le dossier de votre appareil photo"
    )
    if not source_dir:
        messagebox.showerror(
            "Installation annulée",
            "Aucun dossier sélectionné.\nRelancez l'installation pour recommencer."
        )
        sys.exit(1)

    _desktop = Path(win32com.client.Dispatch("WScript.Shell").SpecialFolders("Desktop"))
    orphan_dir = str(_desktop / "Orphelins")

    cfg = configparser.ConfigParser()
    cfg["GENERAL"] = {
        "BOX_NAME":          "StudioVision Monitor",
        "DEFAULT_EXAM_NAME": "Image",
    }
    cfg["PATHS"] = {
        "SOURCE_DIR":  source_dir,
        "ORPHAN_DIR":  orphan_dir,
        "DEST_PHOTOS": str(dest_photos),
        "BACKEND_MDB": str(backend_mdb),
        "SV_TARGET":   sv_target,
        "SV_ARGS":     sv_args,
        "SV_CWD":      sv_cwd,
    }
    cfg["TIMEOUTS"] = {"PATIENT_WAIT_TIMEOUT": "900"}
    with open(config_path, "w", encoding="utf-8") as f:
        cfg.write(f)
    log.info(f"config.ini écrit dans : {config_path}")

    own_exe = Path(sys.executable) if getattr(sys, "frozen", False) else Path(__file__).resolve()
    create_desktop_shortcut(own_exe)

    messagebox.showinfo(
        "Installation terminée !",
        "Installation terminée avec succès !\n\n"
        "ACTION REQUISE :\n"
        "Veuillez maintenant FERMER la fenêtre Studio Vision\n"
        "actuellement ouverte.\n\n"
        "Pour travailler, utilisez désormais uniquement\n"
        "le raccourci  'Studio Vision - Connected'\n"
        "créé sur votre Bureau."
    )
    root.destroy()
    # os._exit bypasses sys.exit (which Tkinter can intercept)
    os._exit(0)

# Resolve base directory: next to .exe when frozen, next to .py otherwise
if getattr(sys, "frozen", False):
    _base_dir = Path(sys.executable).parent
else:
    _base_dir = Path(__file__).resolve().parent
_config_path = _base_dir / "config.ini"
ICON_PATH    = _base_dir / "Studiov2000.ico"

# First run: launch setup wizard before loading any config
if not _config_path.exists():
    configurer_via_interface(_config_path)

config = configparser.ConfigParser()
config.read(_config_path, encoding="utf-8")

BOX_NAME             = config.get("GENERAL", "BOX_NAME",          fallback="StudioVision Monitor")
DEFAULT_EXAM_NAME    = config.get("GENERAL", "DEFAULT_EXAM_NAME", fallback="Image")
SOURCE_DIR           = Path(config.get("PATHS", "SOURCE_DIR"))
ORPHAN_DIR           = Path(config.get("PATHS", "ORPHAN_DIR"))
DEST_PHOTOS          = Path(config.get("PATHS", "DEST_PHOTOS"))
BACKEND_MDB          = Path(
    config.get("PATHS", "BACKEND_MDB",
               fallback=config.get("PATHS", "PUBLIC_MDB", fallback=""))
)
SV_TARGET = config.get("PATHS", "SV_TARGET", fallback="").strip()
SV_ARGS   = config.get("PATHS", "SV_ARGS",   fallback="").strip()
SV_CWD    = config.get("PATHS", "SV_CWD",    fallback="").strip()

PATIENT_WAIT_TIMEOUT = config.getint("TIMEOUTS", "PATIENT_WAIT_TIMEOUT", fallback=900)
_MSACCESS_EXE = "MSACCESS.EXE"

WATCHED_EXTENSIONS     = {".jpg", ".jpeg", ".jfif", ".png", ".bmp", ".tif", ".tiff",
                          ".dcm", ".pdf", ".rtf", ".doc", ".docx", ".odt", ".xps", ".html"}
FILE_LOCK_RETRY_DELAY  = 3
FILE_LOCK_MAX_ATTEMPTS = 15
PATIENT_POLL_INTERVAL  = 3
_NETWORK_SHARE_POLL    = 10
ACCESS_FIELD_CODE   = "Code patient"
ACCESS_FIELD_NOM    = "NOM"
ACCESS_FIELD_PRENOM = "Prénom"
SFDOC_SUBFORM_NAME  = "SFDoc"
_ICON_SIZE     = 64
_COLOR_READY   = (30, 144, 255)
_COLOR_ACTIVE  = (50, 205, 50)
_icon: "pystray.Icon | None" = None
_status_text: str             = "Starting..."
_stop_event: threading.Event  = threading.Event()
_mutex_handle = None

def _make_icon(color: tuple) -> "Image.Image":
    """Generate a simple colored circle as a fallback tray icon."""
    img  = Image.new("RGBA", (_ICON_SIZE, _ICON_SIZE), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    margin = 4
    draw.ellipse([margin, margin, _ICON_SIZE - margin, _ICON_SIZE - margin], fill=color)
    return img

def _set_status(text: str, processing: bool = False) -> None:
    """Update the tray tooltip/menu status text."""
    global _status_text
    _status_text = text
    if _icon is not None:
        try:
            _icon.update_menu()
        except Exception as e:
            log.debug(f"Tray update failed: {e}")

def _notify(title: str, message: str = "") -> None:
    """Send a system tray balloon notification."""
    if _icon is not None:
        try:
            _icon.notify(message if message else title, title)
        except Exception as e:
            log.debug(f"Notification failed: {e}")

def _open_logs(icon, item) -> None:  # noqa: ARG001
    """Tray menu action: open the log file in the default viewer."""
    try:
        os.startfile(str(_LOG_FILE))
    except Exception as e:
        log.warning(f"Could not open log file: {e}")

def _quit(icon, item) -> None:  # noqa: ARG001
    """Tray menu action: signal all threads to stop and exit."""
    log.info("Quit requested from tray menu.")
    _stop_event.set()
    icon.stop()

def wait_for_network_share() -> None:
    """Block until SOURCE_DIR is reachable (UNC paths only). No-op for local paths."""
    is_network = str(SOURCE_DIR).startswith("\\\\") or str(SOURCE_DIR).startswith("//")
    if not is_network:
        return
    attempt = 0
    while not SOURCE_DIR.is_dir():
        attempt += 1
        log.warning(
            f"Network share not reachable: {SOURCE_DIR}  "
            f"(attempt {attempt}, retrying in {_NETWORK_SHARE_POLL}s)"
        )
        time.sleep(_NETWORK_SHARE_POLL)
    if attempt:
        log.info(f"Network share is now accessible after {attempt} attempt(s): {SOURCE_DIR}")

def db_connect(mdb_path: Path):
    """Open a pyodbc connection to an MDB/ACCDB file via the Access ODBC driver."""
    return pyodbc.connect(
        f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={mdb_path};"
    )

def get_active_patient() -> "dict | None":
    """Read Code/NOM/Prénom from the currently active Access form via COM.
    Returns a dict or None if no patient form is open."""
    if not WIN32_AVAILABLE:
        return None
    try:
        access = win32com.client.GetActiveObject("Access.Application")
        form   = access.Screen.ActiveForm
        if form is None:
            return None
        target = {ACCESS_FIELD_CODE, ACCESS_FIELD_NOM, ACCESS_FIELD_PRENOM}
        data: dict = {}
        for i in range(form.Controls.Count):
            ctrl = form.Controls(i)
            try:
                if str(ctrl.Name) in target:
                    data[ctrl.Name] = ctrl.Value
            except Exception:
                pass
        if not target.issubset(data.keys()):
            return None
        return {
            "code":   str(data[ACCESS_FIELD_CODE]),
            "nom":    str(data[ACCESS_FIELD_NOM]),
            "prenom": str(data[ACCESS_FIELD_PRENOM]),
        }
    except Exception as e:
        log.debug(f"COM error: {e}")
        return None

def find_patient_folder(patient_code: str) -> "Path | None":
    """Resolve the patient's photo folder on disk by querying [Photo externe]
    from the Documents table in the backend DB."""
    if not PYODBC_AVAILABLE:
        log.error("pyodbc not available.")
        return None
    if not BACKEND_MDB.exists():
        log.error(f"Backend DB not found: {BACKEND_MDB}")
        return None
    try:
        conn   = db_connect(BACKEND_MDB)
        cursor = conn.cursor()
        cursor.execute(
            "SELECT TOP 1 [Photo externe] FROM Documents "
            "WHERE [code patient] = ? AND [Photo externe] IS NOT NULL",
            (int(patient_code),)
        )
        row = cursor.fetchone()
        conn.close()
        if not row or not row[0]:
            log.warning(f"No existing document found for patient {patient_code}.")
            return None
        parts = row[0].strip().strip("\\").split("\\")
        if len(parts) < 2:
            log.error(f"Unexpected Photo externe format: {row[0]}")
            return None
        folder = DEST_PHOTOS / parts[0] / parts[1]
        if not folder.is_dir():
            log.error(f"Folder found in DB but missing on disk: {folder}")
            return None
        log.info(f"Patient folder resolved: {folder}")
        return folder
    except Exception as e:
        log.error(f"DB folder lookup failed: {e}")
        return None

def insert_document(patient: dict, relative_path: str, description: str) -> bool:
    """Insert a new Documents row linking the file to the patient. Returns success flag."""
    if not PYODBC_AVAILABLE:
        log.warning("pyodbc not available, insert skipped.")
        return False
    if not BACKEND_MDB.exists():
        log.error("Backend DB not found, insert skipped.")
        return False
    try:
        conn   = db_connect(BACKEND_MDB)
        cursor = conn.cursor()
        cursor.execute(
            """
            INSERT INTO Documents
                ([code patient], [Date], DESCRIPTIONS, TEXTE, [Photo externe], TypeVW, NumDocExterne)
            VALUES (?, ?, ?, ?, ?, 99, NULL)
            """,
            (int(patient["code"]), datetime.now(), description, relative_path, relative_path)
        )
        conn.commit()
        conn.close()
        log.info(f"Insert OK: patient={patient['code']} path='{relative_path}' db={BACKEND_MDB.name}")
        return True
    except Exception as e:
        log.error(f"DB insert failed: {e}")
        return False

_AC_SUBFORM = 112  # Access ControlType constant for subforms

def _find_sfdoc(form):
    """Recursively search a form's control tree for the SFDoc subform."""
    for i in range(form.Controls.Count):
        ctrl = form.Controls(i)
        try:
            if ctrl.ControlType != _AC_SUBFORM:
                continue
            if ctrl.Name == SFDOC_SUBFORM_NAME:
                return ctrl.Form
            found = _find_sfdoc(ctrl.Form)
            if found is not None:
                return found
        except Exception:
            pass
    return None

def refresh_ui(expected_patient_code: "str | None" = None) -> None:
    """Refresh the active Access form and Requery the SFDoc subform via COM.
    Skips if the active patient has changed since the file was inserted.
    Falls back to Refresh() if Requery() fails repeatedly."""
    if not WIN32_AVAILABLE:
        return
    try:
        access = win32com.client.GetActiveObject("Access.Application")
        form   = access.Screen.ActiveForm
        if form is None:
            log.warning("Refresh skipped: no active form in Access.")
            return
        if expected_patient_code is not None:
            current = get_active_patient()
            current_code = current["code"] if current else None
            if current_code != expected_patient_code:
                log.warning(
                    f"Refresh skipped: active patient changed "
                    f"(expected={expected_patient_code}, current={current_code})."
                )
                return
        try:
            form.Refresh()
            log.info(f"Refresh() on parent form '{form.Name}'")
        except Exception as e_ref:
            log.warning(f"Refresh() on parent form failed ({e_ref}), continuing...")
        sfdoc = _find_sfdoc(form)
        if sfdoc is None:
            log.warning(
                f"Subform '{SFDOC_SUBFORM_NAME}' not found in the active form. "
                "SFDoc refresh skipped."
            )
            return
        try:
            if form.Dirty:
                log.info("Parent form is in edit mode (Dirty=True); clearing Dirty before Requery.")
                form.Dirty = False
        except Exception as e_dirty:
            log.debug(f"Dirty check/clear failed ({e_dirty}), continuing...")
        _REQUERY_ATTEMPTS = 3
        _REQUERY_DELAY    = 0.5
        requery_ok = False
        for attempt in range(1, _REQUERY_ATTEMPTS + 1):
            try:
                sfdoc.Requery()
                log.info(f"Requery() on '{SFDOC_SUBFORM_NAME}' (attempt {attempt})")
                requery_ok = True
                break
            except Exception as e_req:
                log.warning(
                    f"Requery() attempt {attempt}/{_REQUERY_ATTEMPTS} failed "
                    f"on '{SFDOC_SUBFORM_NAME}': {e_req}"
                )
                if attempt < _REQUERY_ATTEMPTS:
                    time.sleep(_REQUERY_DELAY)
        if not requery_ok:
            log.warning(
                f"All {_REQUERY_ATTEMPTS} Requery() attempts failed on "
                f"'{SFDOC_SUBFORM_NAME}'; falling back to Refresh()."
            )
            try:
                sfdoc.Refresh()
                log.info(f"Fallback Refresh() on '{SFDOC_SUBFORM_NAME}'")
            except Exception as e_ref2:
                log.warning(
                    f"Fallback Refresh() also failed on '{SFDOC_SUBFORM_NAME}': {e_ref2}"
                )
        try:
            sfdoc.Recordset.MoveLast()
            log.info(f"MoveLast() on '{SFDOC_SUBFORM_NAME}'")
        except Exception as e_ml:
            log.debug(f"MoveLast() failed on '{SFDOC_SUBFORM_NAME}': {e_ml}")
    except Exception as e:
        log.warning(f"COM refresh failed (non-blocking): {e}")

def wait_for_file(file: Path) -> bool:
    """Poll until the file is no longer locked by another process.
    Returns True when the file is writable, False after max attempts."""
    for attempt in range(1, FILE_LOCK_MAX_ATTEMPTS + 1):
        try:
            with file.open("ab"):
                return True
        except (PermissionError, OSError):
            log.debug(f"Fichier verrouillé ({attempt}/{FILE_LOCK_MAX_ATTEMPTS}), retrying...")
            time.sleep(FILE_LOCK_RETRY_DELAY)
            
    log.error(f"Fichier toujours verrouillé après {FILE_LOCK_MAX_ATTEMPTS} tentatives: {file}")
    return False

def move_file(source: Path, dest_folder: Path, label: str = "") -> "Path | None":
    """Move source to dest_folder, appending a timestamp on name collision.
    Returns the final destination path, or None on failure."""
    dest_folder.mkdir(parents=True, exist_ok=True)
    dest = dest_folder / source.name
    if dest.exists():
        ts   = int(time.time())
        dest = dest_folder / f"{source.stem}_{ts}{source.suffix}"
        log.info(f"Name conflict, renamed to {dest.name}")
    try:
        shutil.move(str(source), str(dest))
        tag = f"[{label}]  " if label else ""
        log.info(f"{tag}{source.name} -> {dest}")
        return dest
    except Exception as e:
        log.error(f"Move failed: {e}")
        return None

def orphan_file(file: Path) -> None:
    """Move an unmatched file to ORPHAN_DIR for manual review."""
    log.warning(f"Orphaning: {file.name}")
    move_file(file, ORPHAN_DIR, label="ORPHAN")

def prevent_sleep() -> None:
    """Keep the system awake (ES_CONTINUOUS | ES_SYSTEM_REQUIRED) for the process lifetime."""
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(
            0x80000000 |  # ES_CONTINUOUS
            0x00000001    # ES_SYSTEM_REQUIRED
        )
        log.info("Sleep prevention active.")
    except Exception as e:
        log.warning(f"Could not set execution state: {e}")

def worker(file_queue: queue.Queue) -> None:
    """Consumer thread: dequeues files, waits for a patient, copies to DB folder,
    inserts a Documents row, and batches UI refreshes with a 1.5s burst debounce."""
    pythoncom.CoInitialize()
    log.info("Worker started.")
    needs_refresh: bool           = False
    last_patient_code: "str | None" = None
    burst_count: int              = 0
    try:
        while True:
            try:
                file: Path = file_queue.get(timeout=1.5)
            except queue.Empty:
                # No new file for 1.5s: flush pending refresh
                if needs_refresh:
                    log.info("Burst complete — triggering batched UI refresh.")
                    refresh_ui(expected_patient_code=last_patient_code)
                    needs_refresh = False
                    last_patient_code = None
                    _notify(
                        "Transfer complete",
                        f"{burst_count} file(s) processed",
                    )
                    _set_status(f"{BOX_NAME} — Ready", processing=False)
                    burst_count = 0
                continue
            except Exception as e:
                log.error(f"Queue error: {e}")
                continue
            log.info(f"Processing: {file.name} ({file_queue.qsize()} pending)")
            if burst_count == 0 and not needs_refresh:
                _notify("Transfer in progress", file.name)
            _set_status("Transfer in progress...", processing=True)
            if not file.exists():
                log.warning(f"File gone before processing: {file}")
                file_queue.task_done()
                continue
            if not wait_for_file(file):
                log.error(f"Aborting, persistent lock: {file.name}")
                _notify("Error", f"File locked: {file.name}")
                file_queue.task_done()
                continue
            # Wait for the doctor to open a patient record in Access
            patient    = None
            start_time = time.monotonic()
            first_log  = True
            while True:
                patient = get_active_patient()
                if patient:
                    break
                elapsed = time.monotonic() - start_time
                if elapsed >= PATIENT_WAIT_TIMEOUT:
                    orphan_file(file)
                    _notify("Orphan file", file.name)
                    file_queue.task_done()
                    patient = None
                    break
                if first_log:
                    log.info(
                        f"No patient open, waiting "
                        f"(timeout in {PATIENT_WAIT_TIMEOUT // 60} min)"
                    )
                    first_log = False
                time.sleep(PATIENT_POLL_INTERVAL)
            if patient is None:
                continue
            log.info(
                f"Patient: {patient['nom']} {patient['prenom']} "
                f"(code {patient['code']})"
            )
            patient_folder = find_patient_folder(patient["code"])
            if not patient_folder:
                log.error(
                    f"Could not resolve folder for patient {patient['code']}. "
                    "Orphaning."
                )
                orphan_file(file)
                _notify("Orphan file", file.name)
                file_queue.task_done()
                continue
            dest = move_file(file, patient_folder)
            if dest is None:
                file_queue.task_done()
                continue
            group_name    = patient_folder.parent.name
            relative_path = f"\\{group_name}\\{patient_folder.name}\\{dest.name}"
            
            description = DEFAULT_EXAM_NAME
            if insert_document(patient, relative_path, description):
                needs_refresh     = True
                last_patient_code = patient["code"]
                burst_count      += 1
                log.debug("Insert OK — needs_refresh=True (refresh deferred to burst end).")
            else:
                log.warning("Insert failed, refresh flag unchanged.")
                _notify("DB Error", "Insert failed — check logs")
            file_queue.task_done()
    finally:
        if needs_refresh:
            log.info("Worker shutting down — flushing pending UI refresh.")
            refresh_ui(expected_patient_code=last_patient_code)
            if burst_count:
                _notify("Transfer complete", f"{burst_count} file(s) processed")
        _set_status(f"{BOX_NAME} — Stopped")
        pythoncom.CoUninitialize()

class ImageProducer(FileSystemEventHandler):
    """Watchdog handler: enqueues new files matching WATCHED_EXTENSIONS."""
    def __init__(self, file_queue: queue.Queue) -> None:
        super().__init__()
        self._queue = file_queue

    def on_created(self, event) -> None:
        if event.is_directory:
            return
        file = Path(event.src_path)
        if file.suffix.lower() not in WATCHED_EXTENSIONS:
            return
        log.info(f"Enqueued: {file.name} (queue size: {self._queue.qsize() + 1})")
        self._queue.put(file)

def _run_background(file_queue: queue.Queue) -> None:
    """Producer thread: runs the PollingObserver and restarts it on network drops."""
    _RECONNECT_WAIT = 15

    def _start_observer() -> Observer:
        obs = Observer()
        obs.schedule(ImageProducer(file_queue), str(SOURCE_DIR), recursive=True)
        obs.start()
        log.info("Observer started — watching for images.")
        return obs

    observer = _start_observer()
    _set_status(f"{BOX_NAME} — Ready", processing=False)
    try:
        while not _stop_event.is_set():
            if not observer.is_alive():
                log.warning("Observer has stopped (network drop?). Attempting reconnect...")
                _set_status(f"{BOX_NAME} — Reconnecting...", processing=False)
                try:
                    observer.stop()
                    observer.join(timeout=5)
                except Exception:
                    pass
                wait_for_network_share()
                log.info(f"Waiting {_RECONNECT_WAIT}s before restarting observer...")
                time.sleep(_RECONNECT_WAIT)
                observer = _start_observer()
                _set_status(f"{BOX_NAME} — Ready", processing=False)
            time.sleep(1)
    finally:
        observer.stop()
        observer.join()
        remaining = file_queue.qsize()
        if remaining:
            log.info(f"Waiting for {remaining} remaining file(s)...")
            file_queue.join()
        log.info("Background thread stopped.")
        if _icon is not None:
            _icon.stop()

def _is_sv_running() -> bool:
    """Return True if MSACCESS.EXE is currently in the process list."""
    try:
        result = subprocess.run(
            ["tasklist", "/FI", f"IMAGENAME eq {_MSACCESS_EXE}", "/NH"],
            capture_output=True,
            text=True,
            timeout=5,
        )
        return _MSACCESS_EXE.lower() in result.stdout.lower()
    except Exception as e:
        log.warning(f"Impossible de vérifier si MSACCESS.EXE tourne : {e}")
        return False

def _is_sv_frontend_loaded() -> bool:
    """Return True if the correct Studio Vision frontend (.mde) is loaded in Access.
    Checks via COM that CurrentDb().Name matches SV_ARGS (the .mde path).
    Falls back to True if COM is unavailable (to avoid killing a valid session)."""
    if not WIN32_AVAILABLE:
        return _is_sv_running()
    if not SV_ARGS:
        return _is_sv_running()
    # Extract the .mde path from SV_ARGS (format: /runtime path\file.mde /wrkgrp ...)
    mde_path = None
    for part in SV_ARGS.split():
        if part.lower().endswith(".mde") or part.lower().endswith(".mdb") or part.lower().endswith(".accdb"):
            mde_path = part.strip('"').lower()
            break
    if mde_path is None:
        return _is_sv_running()
    try:
        access = win32com.client.GetActiveObject("Access.Application")
        current_db = access.CurrentDb().Name.lower()
        if mde_path in current_db or current_db in mde_path:
            log.info(f"Studio Vision correctement chargé : {current_db}")
            return True
        else:
            log.warning(
                f"MSACCESS.EXE tourne mais avec une base différente : "
                f"'{current_db}' (attendu: '{mde_path}') — relancement nécessaire."
            )
            return False
    except Exception:
        # COM not responding = Access is open but the frontend is not loaded yet, or crashed
        log.warning("MSACCESS.EXE présent mais COM non disponible — considéré comme non prêt.")
        return False

def _kill_msaccess() -> None:
    """Force-terminate all MSACCESS.EXE processes."""
    try:
        subprocess.run(
            ["taskkill", "/F", "/IM", _MSACCESS_EXE],
            capture_output=True, timeout=10
        )
        log.info("MSACCESS.EXE terminé de force.")
        time.sleep(2)
    except Exception as e:
        log.warning(f"Impossible de terminer MSACCESS.EXE : {e}")

def _ensure_sv_running(force_relaunch: bool = False) -> None:
    """Launch Studio Vision if not running, or if the wrong frontend is loaded.
    When force_relaunch=True (duplicate instance), kill any zombie MSACCESS and restart.
    Replicates a double-click on the original shortcut using SV_TARGET/SV_ARGS/SV_CWD."""
    if not SV_TARGET:
        log.warning(
            "SV_TARGET non configuré — "
            "lancement automatique de Studio Vision désactivé."
        )
        return

    sv_ok = _is_sv_frontend_loaded()

    if sv_ok and not force_relaunch:
        log.info("Studio Vision est déjà en cours d'exécution avec le bon frontend.")
        return

    if not sv_ok and _is_sv_running():
        log.warning("MSACCESS.EXE zombie ou mauvais frontend détecté — fermeture forcée.")
        _kill_msaccess()
    elif force_relaunch and sv_ok:
        # Duplicate click but SV is fine — nothing to do
        log.info("Studio Vision déjà actif et fonctionnel — aucune action nécessaire.")
        return

    cmd = f'"{SV_TARGET}"'
    if SV_ARGS:
        cmd = f'{cmd} {SV_ARGS}'
    log.info(f"Lancement de Studio Vision via : {cmd}")
    log.info(f"  WorkingDirectory  : {SV_CWD or '(hérité)'}")
    try:
        subprocess.Popen(
            cmd,
            cwd=SV_CWD or None,
            shell=True,
        )
    except Exception as e:
        log.error(f"Impossible de relancer Studio Vision : {e}")
        return
    _SV_LAUNCH_TIMEOUT = 30
    for elapsed in range(_SV_LAUNCH_TIMEOUT):
        time.sleep(1)
        if _is_sv_frontend_loaded():
            log.info(f"Studio Vision démarré avec succès (après ~{elapsed + 1}s).")
            return
    log.warning(
        f"Studio Vision n'a pas répondu dans les {_SV_LAUNCH_TIMEOUT}s imparties. "
        "Le routeur va continuer et attendra la connexion COM."
    )

def main() -> None:
    global _icon, _mutex_handle
    # Single-instance guard: if already running, just ensure SV is up and exit
    _mutex_handle = win32event.CreateMutex(None, False, "ImageRouter_StudioVision_Mutex")
    if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        log.info("Instance déjà en cours — vérification et relancement de Studio Vision si nécessaire.")
        _ensure_sv_running(force_relaunch=True)
        log.info("Arrêt silencieux du processus doublon.")
        sys.exit(0)
    _ensure_sv_running(force_relaunch=False)
    prevent_sleep()
    if not SOURCE_DIR.exists():
        log.critical(f"Source folder not found: {SOURCE_DIR}")
        sys.exit(1)
    ORPHAN_DIR.mkdir(parents=True, exist_ok=True)
    log.info("Version 5 — Lanceur dynamique + Routeur d'images démarré")
    log.info(f"  Source      : {SOURCE_DIR}")
    log.info(f"  Dest        : {DEST_PHOTOS}")
    log.info(f"  BACKEND_MDB : {BACKEND_MDB}")
    log.info(f"  SV_TARGET   : {SV_TARGET or '(non configuré)'}")
    log.info(f"  SV_ARGS     : {SV_ARGS   or '(vide)'}")
    log.info(f"  SV_CWD      : {SV_CWD    or '(hérité)'}")
    log.info(f"  Orphans     : {ORPHAN_DIR}")
    log.info(f"  Log file    : {_LOG_FILE}")
    log.info(f"  Timeout     : {PATIENT_WAIT_TIMEOUT // 60} min")
    log.info(f"  Ext         : {', '.join(sorted(WATCHED_EXTENSIONS))}")
    file_queue: queue.Queue = queue.Queue()
    threading.Thread(
        target=worker, args=(file_queue,), name="Worker", daemon=True
    ).start()
    threading.Thread(
        target=_run_background, args=(file_queue,), name="Background", daemon=True
    ).start()
    if not TRAY_AVAILABLE:
        log.warning("pystray/Pillow not available — running without system tray.")
        try:
            while not _stop_event.is_set():
                time.sleep(1)
        except KeyboardInterrupt:
            log.info("Shutdown requested.")
        finally:
            _stop_event.set()
        return
    menu = pystray.Menu(
        pystray.MenuItem(
            text=lambda item: _status_text,
            action=None,
            enabled=False,
        ),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Open logs", _open_logs),
        pystray.MenuItem("Quit", _quit),
    )
    try:
        if ICON_PATH.exists():
            tray_image = Image.open(str(ICON_PATH))
        else:
            tray_image = _make_icon(_COLOR_READY)
    except Exception as e:
        log.warning(f"Impossible de charger l'icône {ICON_PATH} ({e}) — icône de secours utilisée.")
        tray_image = _make_icon(_COLOR_READY)
    _icon = pystray.Icon(
        name=BOX_NAME,
        icon=tray_image,
        title=BOX_NAME,
        menu=menu,
    )
    log.info("System tray icon started.")
    _icon.run()
    _stop_event.set()
    log.info("Application stopped.")

if __name__ == "__main__":
    main()