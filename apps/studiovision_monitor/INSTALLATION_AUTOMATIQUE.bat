@echo off
title Installation Automatique - Routeur Studio Vision
cls

:: 1. VERIFICATION DES DROITS ADMINISTRATEUR
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo ====================================================================
    echo   ERREUR : Droits d'administrateur requis !
    echo ====================================================================
    echo.
    echo 1. Faites un CLIC DROIT sur ce fichier
    echo 2. Choisissez 'Executer en tant qu'administrateur'
    echo.
    pause
    exit /b
}

echo ====================================================================
echo   DEBUT DE L'INSTALLATION DU ROUTEUR D'IMAGES (STANDARD)
echo ====================================================================
echo.

:: 2. INSTALLATION SILENCIEUSE DU MOTEUR ACCESS
echo [1/4] Installation du composant systeme Microsoft Access...
if exist "accessdatabaseengine_X64.exe" (
    start /wait "" "accessdatabaseengine_X64.exe" /quiet
    echo [OK] Composant Access installe avec succes.
) else (
    echo [ALERTE] Fichier 'accessdatabaseengine_X64.exe' absent.
)
echo.

:: 3. CREATION DU DOSSIER ET COPIE DE L'APPLICATION
echo [2/4] Configuration du repertoire...
set "TARGET_DIR=C:\Routeur_Images"
if not exist "%TARGET_DIR%" mkdir "%TARGET_DIR%"

if exist "studiovision_monitor_AL.exe" (
    copy /y "studiovision_monitor_AL.exe" "%TARGET_DIR%\studiovision_monitor_AL.exe" >nul
    echo [OK] Application copiee dans : %TARGET_DIR%
) else (
    echo [ERREUR] Le fichier 'studiovision_monitor_AL.exe' est introuvable !
    pause
    exit /b
)
echo.

:: 4. INSCRIPTION DANS LE REGISTRE (DEMARRAGE AUTOMATIQUE)
echo [3/4] Inscription au demarrage de Windows...
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v "StudioVisionMonitor" /t REG_SZ /d "\"%TARGET_DIR%\studiovision_monitor_AL.exe\"" /f >nul
echo [OK] Le routeur demarrera automatiquement a chaque ouverture de session.
echo.

:: 5. FINALISATION
echo [4/4] Lancement de l'interface de configuration...
echo.
echo ====================================================================
echo   INSTALLATION REUSSIE !
echo ====================================================================
echo L'application va maintenant s'ouvrir.
echo Veuillez repondre aux questions a l'ecran pour finaliser.

start "" "%TARGET_DIR%\studiovision_monitor_AL.exe"

timeout /t 5 >nul
exit