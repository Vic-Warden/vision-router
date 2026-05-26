"""
Routes incoming imaging files to the correct patient folder,
inserts a DB record, and refreshes the Access UI.

Pipeline: PollingObserver → file_queue → Worker → Access DB + UI refresh
          (1.5 s burst debounce, auto-reconnect on network drop)

Dependencies: watchdog, pyodbc, pywin32, pythoncom, pystray, Pillow, psutil
"""

import os
import sys
import time
import queue
import shutil
import ctypes
import logging
import threading
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
import psutil

# --- INITIALISATION DES LOGS ---
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

# ---------------------------------------------------------------------------
#  HELPERS DE CONFIGURATION AUTOMATIQUE
# ---------------------------------------------------------------------------

def get_db_path_from_com() -> "Path | None":
    """
    Retourne le chemin réseau absolu du Backend Studio Vision (agnostique du nom
    de fichier) en lisant la propriété .Connect de la table liée "Documents" via COM.

    La chaîne Connect est de la forme :
        ;DATABASE=\\\\serveur\\dossier\\fichier\\NOM_DE_LA_BASE.MDB
    On extrait la partie après "DATABASE=" pour obtenir le chemin réseau réel,
    quel que soit le nom du fichier backend (PUBLIC.MDB, BASE_2026.ACCDB, etc.).

    Fallback : si la table "Documents" n'est pas liée (installation locale sans
    table liée), on retourne access.CurrentDb().Name — comportement conservatif.
    """
    if not WIN32_AVAILABLE:
        log.error("win32com non disponible — détection COM impossible.")
        return None
    try:
        access = win32com.client.GetActiveObject("Access.Application")
        db     = access.CurrentDb()

        # Tentative de lecture du chemin Backend via la table liée "Documents"
        try:
            connect_string = db.TableDefs("Documents").Connect
            if connect_string and "DATABASE=" in connect_string.upper():
                # Extraction robuste : on cherche "DATABASE=" en insensible à la casse
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

        # Fallback : chemin du Frontend local (comportement original)
        db_name = db.Name
        db_path = Path(db_name)
        log.info(f"Base de données détectée via COM (fallback Frontend) : {db_path}")
        return db_path

    except Exception as e:
        log.debug(f"COM get_db_path échoué : {e}")
        return None


def get_frontend_path_from_com() -> "Path | None":
    """
    Retourne le chemin du fichier Frontend Access actuellement ouvert
    (ex : Ophprog.mde, StudVis.mde, …) via access.CurrentDb().Name.

    Ce chemin est utilisé pour relancer Studio Vision au démarrage du routeur
    via os.startfile(), ce qui délègue l'ouverture à Windows (association de
    fichier → MSACCESS.EXE) sans présumer du chemin d'Access lui-même.
    """
    if not WIN32_AVAILABLE:
        log.warning("win32com non disponible — détection Frontend COM impossible.")
        return None
    try:
        access        = win32com.client.GetActiveObject("Access.Application")
        frontend_name = access.CurrentDb().Name
        frontend_path = Path(frontend_name)
        log.info(f"Frontend Access détecté via COM : {frontend_path}")
        return frontend_path
    except Exception as e:
        log.warning(f"Détection Frontend via COM échouée : {e}")
        return None


def create_desktop_shortcut(target_exe: Path) -> None:
    """
    Crée un raccourci '.lnk' nommé 'Studio Vision - Connected'
    sur le Bureau de l'utilisateur courant.

    L'icône utilisée est celle embarquée dans target_exe (le routeur compilé).
    Aucune tentative d'emprunter l'icône d'un .mde / .mdb tiers, car ces
    fichiers ne contiennent pas d'icônes extractibles par WScript.Shell.

    Paramètres
    ----------
    target_exe : Path
        Chemin vers l'exécutable du routeur (studiovision_monitor_AL.exe).
    """
    if not WIN32_AVAILABLE:
        log.warning("win32com non disponible — création du raccourci ignorée.")
        return
    try:
        shell      = win32com.client.Dispatch("WScript.Shell")
        desktop    = shell.SpecialFolders("Desktop")
        lnk_path   = os.path.join(desktop, "Studio Vision - Connected.lnk")
        shortcut   = shell.CreateShortcut(lnk_path)
        shortcut.TargetPath       = str(target_exe)
        shortcut.WorkingDirectory = str(target_exe.parent)
        shortcut.Description      = "Lance Studio Vision avec le routeur d'images intégré"
        # Icône : icône par défaut de l'exécutable compilé (index 0)
        shortcut.IconLocation     = f"{target_exe}, 0"
        shortcut.save()
        log.info(f"Raccourci Bureau créé : {lnk_path}")
    except Exception as e:
        log.error(f"Impossible de créer le raccourci Bureau : {e}")


# ---------------------------------------------------------------------------
#  INTERFACE DE CONFIGURATION (PREMIÈRE INSTALLATION)
# ---------------------------------------------------------------------------

def configurer_via_interface(config_path: Path) -> None:
    """
    Première configuration — flux UX minimal : 3 interactions seulement.

      1. Une boîte de bienvenue  →  rappel Studio Vision ouvert + annonce filedialog
      2. filedialog               →  sélection du dossier SOURCE
      3. (si besoin)              →  sélection manuelle du dossier Photos si non détecté
      4. Popup de succès          →  confirmation + nom du raccourci créé

    Tout le reste (Backend DB, DEST_PHOTOS, studiovision.exe)
    est détecté automatiquement, sans solliciter le médecin.
    DOCUM.MDB n'est pas utilisé par le routeur et n'est plus configuré.
    """
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    # ------------------------------------------------------------------
    # ÉTAPE 1 — Bienvenue : une seule popup, texte court et aéré
    # ------------------------------------------------------------------
    messagebox.showinfo(
        "Configuration du Routeur d'Images",
        "Avant de continuer, vérifiez les deux points suivants :\n\n"
        "1.  Studio Vision doit être actuellement OUVERT\n"
        "     EN MODE ADMINISTRATEUR.\n"
        "     (Sans cela, la liaison avec la base de données\n"
        "     ne pourra pas être établie automatiquement.)\n\n"
        "2.  Cliquez sur OK pour sélectionner le dossier\n"
        "     où votre appareil photo envoie les images."
    )

    # ------------------------------------------------------------------
    # Détections automatiques (silencieuses pour l'utilisateur)
    # ------------------------------------------------------------------
    backend_mdb = get_db_path_from_com()

    if backend_mdb is None:
        messagebox.showerror(
            "Studio Vision introuvable",
            "Studio Vision ne répond pas.\n\n"
            "Ouvrez Studio Vision, puis relancez l'installation."
        )
        sys.exit(1)

    # ------------------------------------------------------------------
    # Détection heuristique de dest_photos (architecture variable selon cabinet)
    #   Test 1 : backend_mdb.parent / "Photos"        (base à la racine du partage)
    #   Test 2 : backend_mdb.parent.parent / "Photos"  (base dans un sous-dossier)
    #   Fallback : demander explicitement à l'utilisateur
    # ------------------------------------------------------------------
    _candidate1 = backend_mdb.parent / "Photos"
    _candidate2 = backend_mdb.parent.parent / "Photos"

    if _candidate1.is_dir():
        dest_photos = _candidate1
        log.info(f"DEST_PHOTOS détecté (niveau 1) : {dest_photos}")
    elif _candidate2.is_dir():
        dest_photos = _candidate2
        log.info(f"DEST_PHOTOS détecté (niveau 2) : {dest_photos}")
    else:
        log.warning(
            f"Dossier Photos introuvable en '{_candidate1}' ni en '{_candidate2}'. "
            "Demande manuelle à l'utilisateur."
        )
        messagebox.showinfo(
            "Dossier Photos introuvable",
            "Le dossier 'Photos' de Studio Vision n'a pas pu être détecté "
            "automatiquement.\n\n"
            "Cliquez sur OK, puis sélectionnez ce dossier manuellement."
        )
        _photos_manual = filedialog.askdirectory(
            title="Sélectionnez le dossier Photos de Studio Vision"
        )
        if not _photos_manual:
            messagebox.showerror(
                "Installation annulée",
                "Aucun dossier Photos sélectionné.\n"
                "Relancez l'installation pour recommencer."
            )
            sys.exit(1)
        dest_photos = Path(_photos_manual)
        log.info(f"DEST_PHOTOS sélectionné manuellement : {dest_photos}")

    # ------------------------------------------------------------------
    # Détection du Frontend Access via COM (silencieuse pour l'utilisateur)
    # Le chemin du .mde/.mdb Frontend est lu depuis la même instance COM
    # déjà ouverte — aucune recherche psutil nécessaire.
    # ------------------------------------------------------------------
    sv_frontend_path     = get_frontend_path_from_com()
    sv_frontend_path_str = str(sv_frontend_path) if sv_frontend_path else ""

    log.info(f"BACKEND_MDB      : {backend_mdb}")
    log.info(f"DEST_PHOTOS      : {dest_photos}")
    log.info(f"SV Frontend      : {sv_frontend_path_str or '(non détecté)'}")

    # ------------------------------------------------------------------
    # ÉTAPE 2 — Sélection du dossier SOURCE (seule action requise)
    # ------------------------------------------------------------------
    source_dir = filedialog.askdirectory(
        title="Sélectionnez le dossier de l'appareil photo"
    )
    if not source_dir:
        messagebox.showerror(
            "Installation annulée",
            "Aucun dossier sélectionné.\n"
            "Relancez l'installation pour recommencer."
        )
        sys.exit(1)

    _desktop = Path(win32com.client.Dispatch("WScript.Shell").SpecialFolders("Desktop"))
    orphan_dir = str(_desktop / "Orphelins")

    # ------------------------------------------------------------------
    # Écriture du config.ini  (silencieux)
    # ------------------------------------------------------------------
    cfg = configparser.ConfigParser()
    cfg["GENERAL"] = {
        "BOX_NAME":          "StudioVision Monitor",
        "DEFAULT_EXAM_NAME": "Image",
    }
    cfg["PATHS"] = {
        "SOURCE_DIR":        source_dir,
        "ORPHAN_DIR":        orphan_dir,
        "DEST_PHOTOS":       str(dest_photos),
        "BACKEND_MDB":       str(backend_mdb),
        "SV_FRONTEND_PATH":  sv_frontend_path_str,
    }
    cfg["TIMEOUTS"] = {"PATIENT_WAIT_TIMEOUT": "900"}

    with open(config_path, "w", encoding="utf-8") as f:
        cfg.write(f)
    log.info(f"config.ini écrit dans : {config_path}")

    # Création du raccourci Bureau  (silencieuse — icône par défaut du .exe)
    own_exe = Path(sys.executable) if getattr(sys, "frozen", False) else Path(__file__).resolve()
    create_desktop_shortcut(own_exe)

    # ------------------------------------------------------------------
    # ÉTAPE 3 — Confirmation : courte, lisible, actionnable
    # ------------------------------------------------------------------
    messagebox.showinfo(
        "Installation terminée !",
        "Installation terminée avec succès !\n\n"
        "ACTION REQUISE :\n"
        "Veuillez maintenant FERMER la fenêtre Studio Vision\n"
        "actuellement ouverte (mode administrateur).\n\n"
        "Pour travailler, utilisez désormais uniquement\n"
        "le raccourci  'Studio Vision - Connected'\n"
        "créé sur votre Bureau."
    )
    root.destroy()


# ---------------------------------------------------------------------------
#  CHARGEMENT DE LA CONFIGURATION
# ---------------------------------------------------------------------------

# Résolution du dossier de base : dossier du .exe en mode compilé (PyInstaller),
# dossier du .py en mode source. sys._MEIPASS (dossier temporaire de décompression)
# n'est volontairement PAS utilisé — le config.ini doit persister entre les lancements.
if getattr(sys, "frozen", False):
    _base_dir = Path(sys.executable).parent
else:
    _base_dir = Path(__file__).resolve().parent
_config_path = _base_dir / "config.ini"

if not _config_path.exists():
    configurer_via_interface(_config_path)

config = configparser.ConfigParser()
config.read(_config_path, encoding="utf-8")

BOX_NAME             = config.get("GENERAL", "BOX_NAME",          fallback="StudioVision Monitor")
DEFAULT_EXAM_NAME    = config.get("GENERAL", "DEFAULT_EXAM_NAME", fallback="Image")
SOURCE_DIR           = Path(config.get("PATHS", "SOURCE_DIR"))
ORPHAN_DIR           = Path(config.get("PATHS", "ORPHAN_DIR"))
DEST_PHOTOS          = Path(config.get("PATHS", "DEST_PHOTOS"))
# Rétrocompatibilité : anciens config.ini écrits avec la clé "PUBLIC_MDB"
BACKEND_MDB          = Path(
    config.get("PATHS", "BACKEND_MDB",
               fallback=config.get("PATHS", "PUBLIC_MDB", fallback=""))
)
# Chemin du Frontend Access (.mde/.mdb) — lancé via os.startfile().
# Rétrocompatibilité : anciens config.ini écrits avec "STUDIO_VISION_EXE_PATH".
SV_FRONTEND_PATH     = config.get(
    "PATHS", "SV_FRONTEND_PATH",
    fallback=config.get("PATHS", "STUDIO_VISION_EXE_PATH", fallback="")
)
PATIENT_WAIT_TIMEOUT = config.getint("TIMEOUTS", "PATIENT_WAIT_TIMEOUT", fallback=900)

# Processus à surveiller pour détecter si Access / Studio Vision est ouvert
_MSACCESS_EXE        = "msaccess.exe"
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
    img  = Image.new("RGBA", (_ICON_SIZE, _ICON_SIZE), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    margin = 4
    draw.ellipse([margin, margin, _ICON_SIZE - margin, _ICON_SIZE - margin], fill=color)
    return img

def _set_status(text: str, processing: bool = False) -> None:
    global _status_text
    _status_text = text
    if _icon is not None:
        try:
            _icon.icon = _make_icon(_COLOR_ACTIVE if processing else _COLOR_READY)
            _icon.update_menu()
        except Exception as e:
            log.debug(f"Tray update failed: {e}")


def _notify(title: str, message: str = "") -> None:
    if _icon is not None:
        try:
            _icon.notify(message if message else title, title)
        except Exception as e:
            log.debug(f"Notification failed: {e}")


def _open_logs(icon, item) -> None:  # noqa: ARG001
    try:
        os.startfile(str(_LOG_FILE))
    except Exception as e:
        log.warning(f"Could not open log file: {e}")


def _quit(icon, item) -> None:  # noqa: ARG001
    log.info("Quit requested from tray menu.")
    _stop_event.set()
    icon.stop()

def wait_for_network_share() -> None:
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
    return pyodbc.connect(
        f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={mdb_path};"
    )


def get_active_patient() -> "dict | None":
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


_AC_SUBFORM = 112


def _find_sfdoc(form):
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
    log.warning(f"Orphaning: {file.name}")
    move_file(file, ORPHAN_DIR, label="ORPHAN")


def prevent_sleep() -> None:
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(
            0x80000000 |  # ES_CONTINUOUS
            0x00000001    # ES_SYSTEM_REQUIRED
        )
        log.info("Sleep prevention active.")
    except Exception as e:
        log.warning(f"Could not set execution state: {e}")

def worker(file_queue: queue.Queue) -> None:
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


#  Watchdog producer
class ImageProducer(FileSystemEventHandler):
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


#  Background thread 
def _run_background(file_queue: queue.Queue) -> None:
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


# ---------------------------------------------------------------------------
#  LANCEMENT DE STUDIO VISION SI NÉCESSAIRE
# ---------------------------------------------------------------------------

def _is_sv_running() -> bool:
    """
    Retourne True si MSACCESS.EXE est actuellement en cours d'exécution.
    Studio Vision étant un Frontend Access (.mde/.mdb), c'est toujours
    MSACCESS.EXE qui est le processus réel à surveiller.
    """
    return any(
        (p.info.get("name") or "").lower() == _MSACCESS_EXE
        for p in psutil.process_iter(["name"])
    )


def _ensure_sv_running() -> None:
    """
    Vérifie si MSACCESS.EXE est ouvert.
    S'il ne l'est pas et que le chemin du Frontend est configuré,
    ouvre le fichier Frontend via os.startfile() (association Windows
    .mde/.mdb → Access) et attend jusqu'à 30 secondes qu'Access soit prêt.
    """
    if _is_sv_running():
        log.info("MSACCESS.EXE est déjà en cours d'exécution.")
        return

    if not SV_FRONTEND_PATH:
        log.warning(
            "SV_FRONTEND_PATH non configuré — "
            "lancement automatique de Studio Vision désactivé."
        )
        return

    sv_frontend = Path(SV_FRONTEND_PATH)
    if not sv_frontend.is_file():
        log.error(
            f"Frontend introuvable à l'emplacement configuré : {sv_frontend}\n"
            "Lancement automatique annulé."
        )
        return

    log.info(f"MSACCESS.EXE non détecté — ouverture du Frontend : {sv_frontend}")
    try:
        # os.startfile délègue à Windows l'association .mde/.mdb → MSACCESS.EXE,
        # sans avoir à connaître le chemin d'installation d'Office.
        os.startfile(str(sv_frontend))
    except Exception as e:
        log.error(f"Impossible d'ouvrir le Frontend Studio Vision : {e}")
        return

    # Attente active : jusqu'à 30 s pour qu'MSACCESS.EXE apparaisse
    _SV_LAUNCH_TIMEOUT = 30
    for elapsed in range(_SV_LAUNCH_TIMEOUT):
        time.sleep(1)
        if _is_sv_running():
            log.info(f"MSACCESS.EXE démarré avec succès (après ~{elapsed + 1}s).")
            return

    log.warning(
        f"MSACCESS.EXE n'a pas démarré dans les {_SV_LAUNCH_TIMEOUT}s imparties. "
        "Le routeur va continuer et attendra la connexion COM."
    )


# ---------------------------------------------------------------------------
#  POINT D'ENTRÉE PRINCIPAL
# ---------------------------------------------------------------------------

def main() -> None:
    global _icon, _mutex_handle

    # ------------------------------------------------------------------
    # 1. Garde single-instance : quitter silencieusement si déjà actif.
    #    Ce mutex est le seul mécanisme anti-doublon — il couvre tous les
    #    scénarios (double-clic sur le raccourci, relance, etc.).
    # ------------------------------------------------------------------
    _mutex_handle = win32event.CreateMutex(None, False, "ImageRouter_StudioVision_Mutex")
    if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        log.info("Instance déjà en cours — vérification de Studio Vision avant arrêt.")
        # Le routeur tourne déjà en tray, mais l'utilisateur a peut-être fermé
        # Studio Vision manuellement. On s'assure qu'il est rouvert avant de quitter.
        _ensure_sv_running()
        log.info("Arrêt silencieux du processus doublon.")
        sys.exit(0)

    # ------------------------------------------------------------------
    # 2. Lancer Studio Vision si nécessaire.
    #    Appelé dès le démarrage via le raccourci 'Studio Vision - Connected'.
    # ------------------------------------------------------------------
    _ensure_sv_running()

    # ------------------------------------------------------------------
    # 3. Vérifications de démarrage
    # ------------------------------------------------------------------
    prevent_sleep()

    if not SOURCE_DIR.exists():
        log.critical(f"Source folder not found: {SOURCE_DIR}")
        sys.exit(1)

    ORPHAN_DIR.mkdir(parents=True, exist_ok=True)

    log.info("Version 5 started")
    log.info(f"  Source      : {SOURCE_DIR}")
    log.info(f"  Dest        : {DEST_PHOTOS}")
    log.info(f"  BACKEND_MDB : {BACKEND_MDB}")
    log.info(f"  SV Frontend : {SV_FRONTEND_PATH or '(non configuré)'}")
    log.info(f"  Orphans     : {ORPHAN_DIR}")
    log.info(f"  Log file    : {_LOG_FILE}")
    log.info(f"  Timeout     : {PATIENT_WAIT_TIMEOUT // 60} min")
    log.info(f"  Ext         : {', '.join(sorted(WATCHED_EXTENSIONS))}")

    # ------------------------------------------------------------------
    # 4. Démarrage des threads de surveillance
    # ------------------------------------------------------------------
    file_queue: queue.Queue = queue.Queue()

    threading.Thread(
        target=worker, args=(file_queue,), name="Worker", daemon=True
    ).start()

    threading.Thread(
        target=_run_background, args=(file_queue,), name="Background", daemon=True
    ).start()

    # ------------------------------------------------------------------
    # 5. Icône système (system tray)
    # ------------------------------------------------------------------
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

    _icon = pystray.Icon(
        name=BOX_NAME,
        icon=_make_icon(_COLOR_READY),
        title=BOX_NAME,
        menu=menu,
    )

    log.info("System tray icon started.")
    _icon.run()

    _stop_event.set()
    log.info("Application stopped.")


if __name__ == "__main__":
    main()