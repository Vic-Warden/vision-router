@echo off
title Installation - Routeur Studio Vision
cls

:: Force le script a travailler dans son propre dossier
cd /d "%~dp0"

:: ====================================================================
:: 1. AUTO-ELEVATION ADMINISTRATEUR (Zero-Click)
:: ====================================================================
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo L'installation necessite les droits d'administrateur.
    echo Veuillez cliquer sur "Oui" dans la fenetre de securite qui va s'afficher...
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
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
    echo [ERREUR] Le fichier 'studiovision_monitor_AL.exe' est introuvable !
    echo Assurez-vous d'avoir telecharge les 4 fichiers dans le meme dossier.
    pause
    exit /b
)

if exist "Studiov2000.ico" (
    copy /y "Studiov2000.ico" "%TARGET_DIR%\Studiov2000.ico" >nul
    echo [OK] Icone copiee dans : %TARGET_DIR%
) else (
    echo [ALERTE] Fichier 'Studiov2000.ico' absent - icone ignoree.
)
echo.

:: ====================================================================
:: 4. LANCEMENT POUR CONFIGURATION INITIALE
:: ====================================================================
echo [3/3] Lancement de la configuration initiale...
echo.
echo ====================================================================
echo   INSTALLATION DES FICHIERS TERMINEE
echo ====================================================================
echo.
echo Veuillez suivre les dernieres instructions a l'ecran.
echo.

start "" "%TARGET_DIR%\studiovision_monitor_AL.exe"
exit