@echo off
title Installation - Routeur Studio Vision
cls

:: ====================================================================
:: 1. VERIFICATION DES DROITS ADMINISTRATEUR
:: ====================================================================
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
)

echo ====================================================================
echo   DEBUT DE L'INSTALLATION DU ROUTEUR D'IMAGES
echo ====================================================================
echo.

:: ====================================================================
:: 2. INSTALLATION SILENCIEUSE DU MOTEUR ACCESS
:: ====================================================================
echo [1/3] Installation du composant systeme Microsoft Access...
if exist "accessdatabaseengine_X64.exe" (
    start /wait "" "accessdatabaseengine_X64.exe" /quiet
    echo [OK] Composant Access installe avec succes.
) else (
    echo [ALERTE] Fichier 'accessdatabaseengine_X64.exe' absent - etape ignoree.
)
echo.

:: ====================================================================
:: 3. CREATION DU DOSSIER ET COPIE DE L'APPLICATION
:: ====================================================================
echo [2/3] Configuration du repertoire d'installation...
set "TARGET_DIR=C:\Routeur_Images"
if not exist "%TARGET_DIR%" mkdir "%TARGET_DIR%"

if exist "studiovision_monitor_AL.exe" (
    copy /y "studiovision_monitor_AL.exe" "%TARGET_DIR%\studiovision_monitor_AL.exe" >nul
    echo [OK] Application copiee dans : %TARGET_DIR%
) else (
    cls
    echo.
    echo  ***********************************************************
    echo  *                                                         *
    echo  *   ATTENTION : FICHIER ZIP NON EXTRAIT !                 *
    echo  *                                                         *
    echo  *   Vous avez lance l'installation directement            *
    echo  *   depuis l'interieur du fichier ZIP.                    *
    echo  *                                                         *
    echo  *   QUE FAIRE :                                           *
    echo  *   1. Fermez cette fenetre.                              *
    echo  *   2. Faites un CLIC DROIT sur le fichier ZIP.           *
    echo  *   3. Choisissez  "Extraire tout..."                     *
    echo  *   4. Ouvrez le dossier extrait.                         *
    echo  *   5. Relancez  INSTALLATION_AUTOMATIQUE.bat             *
    echo  *                                                         *
    echo  ***********************************************************
    echo.
    pause
    exit /b
)
echo.

:: ====================================================================
:: 4. LANCEMENT POUR CONFIGURATION INITIALE
::    Le programme va :
::      - Se connecter a Studio Vision (ouvert) via COM
::      - Detecter automatiquement PUBLIC.MDB, DOCUM.MDB, studiovision.exe
::      - Demander uniquement le dossier SOURCE a l'utilisateur
::      - Creer le raccourci 'Studio Vision - Connected' sur le Bureau
:: ====================================================================
echo [3/3] Lancement de la configuration initiale...
echo.
echo  >>> Assurez-vous que Studio Vision est ouvert avant de continuer <<<
echo.
echo ====================================================================
echo   INSTALLATION DES FICHIERS TERMINEE
echo ====================================================================
echo.
echo L'application va s'ouvrir pour finaliser la configuration.
echo Repondez aux questions a l'ecran (une seule selection de dossier).
echo.
echo Un raccourci 'Studio Vision - Connected' sera automatiquement
echo cree sur le Bureau a la fin de la configuration.
echo.

start "" "%TARGET_DIR%\studiovision_monitor_AL.exe"

timeout /t 5 >nul
exit