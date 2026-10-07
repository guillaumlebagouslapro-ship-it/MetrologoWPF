@echo off
setlocal EnableExtensions
title Reinstallation de Metrologo

REM ============================================================
REM  Reinstallation propre de Metrologo sur CE poste
REM  (a lancer depuis un poste ou le lecteur M: est connecte)
REM
REM  1. ferme Metrologo s'il est ouvert
REM  2. sauvegarde les donnees locales (%LocalAppData%\Metrologo)
REM  3. relance l'installateur depuis M: (derniere version publiee)
REM  4. remet les donnees locales en place, puis relance Metrologo
REM
REM  Pourquoi la sauvegarde : Velopack installe l'appli dans le MEME
REM  dossier que les donnees locales (Configuration, Presets,
REM  Utilisateurs...). Une reinstallation peut vider ce dossier.
REM ============================================================

REM --- Configuration ------------------------------------------
set "APP_ID=Metrologo"
set "FEED=M:\exe_spe\Data_Metrologo\SUITE ASERTI Guillaume\Metrologo"
set "SETUP=%FEED%\%APP_ID%-win-Setup.exe"
set "RACINE=%LOCALAPPDATA%\%APP_ID%"
REM ------------------------------------------------------------

for /f "usebackq delims=" %%d in (`powershell -NoProfile -Command "Get-Date -Format 'yyyyMMdd_HHmm'"`) do set "HORO=%%d"
set "SAUVEGARDE=%LOCALAPPDATA%\%APP_ID%_sauvegarde_%HORO%"

echo ============================================
echo   Reinstallation de Metrologo
echo ============================================
echo.

REM 1) Installateur accessible ?
if not exist "%SETUP%" goto NO_SETUP

echo Installateur : %SETUP%
echo.
choice /c ON /m "Reinstaller Metrologo maintenant (O = oui, N = non) "
if errorlevel 2 goto ANNULE

REM 2) Fermer Metrologo
echo.
echo == Fermeture de Metrologo ==
taskkill /im Metrologo.exe /f >nul 2>&1
timeout /t 2 /nobreak >nul

REM 3) Sauvegarde des donnees locales (sans les fichiers d'installation Velopack)
set "A_SAUVEGARDE=0"
if not exist "%RACINE%\" goto INSTALLER
echo.
echo == Sauvegarde des donnees locales ==
echo   vers %SAUVEGARDE%
robocopy "%RACINE%" "%SAUVEGARDE%" /E /XD "%RACINE%\current" "%RACINE%\packages" /XF Update.exe sq.version /R:1 /W:1 /NFL /NDL /NJH /NJS /NP
if errorlevel 8 goto ECHEC_SAUVEGARDE
set "A_SAUVEGARDE=1"

:INSTALLER
REM 4) Installation (l'installateur relance Metrologo a la fin)
echo.
echo == Installation de la derniere version ==
start "" /wait "%SETUP%"
if errorlevel 1 echo ATTENTION : l'installateur a signale une erreur ou a ete annule.

if "%A_SAUVEGARDE%"=="0" goto LANCER

REM 5) L'installateur demarre Metrologo tout seul : on attend qu'il apparaisse
REM    (15 s max) puis on le ferme le temps de remettre les donnees en place.
set /a N=0
:ATTENTE_APP
tasklist /fi "imagename eq Metrologo.exe" | find /i "Metrologo.exe" >nul
if not errorlevel 1 goto APP_LANCEE
set /a N+=1
if %N% geq 15 goto RESTAURER
timeout /t 1 /nobreak >nul
goto ATTENTE_APP

:APP_LANCEE
taskkill /im Metrologo.exe /f >nul 2>&1
timeout /t 2 /nobreak >nul

:RESTAURER
REM 6) Remise en place des donnees locales
echo.
echo == Restauration des donnees locales ==
robocopy "%SAUVEGARDE%" "%RACINE%" /E /R:1 /W:1 /NFL /NDL /NJH /NJS /NP
if errorlevel 8 goto ECHEC_RESTAURATION

:LANCER
REM 7) Relance
echo.
for /f "tokens=3" %%v in ('reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\%APP_ID%" /v DisplayVersion 2^>nul ^| find "DisplayVersion"') do echo Version installee : %%v
if exist "%RACINE%\current\Metrologo.exe" start "" "%RACINE%\current\Metrologo.exe"

echo.
echo ============================================
echo   Reinstallation terminee.
if "%A_SAUVEGARDE%"=="1" echo   Sauvegarde conservee (supprimable si tout va bien) :
if "%A_SAUVEGARDE%"=="1" echo   %SAUVEGARDE%
echo ============================================
echo.
pause
goto FIN

:NO_SETUP
echo ERREUR : installateur introuvable :
echo    %SETUP%
echo.
echo Connecte le lecteur M: avec le raccourci du bureau puis relance.
echo (Si M: est connecte, Metrologo n'a peut-etre pas encore ete publie :
echo  lancer publier-metrologo.bat.)
pause
goto FIN

:ANNULE
echo Reinstallation annulee, rien n'a ete modifie.
pause
goto FIN

:ECHEC_SAUVEGARDE
echo.
echo ERREUR pendant la sauvegarde des donnees locales.
echo Reinstallation ANNULEE pour ne rien perdre. Rien n'a ete supprime.
pause
goto FIN

:ECHEC_RESTAURATION
echo.
echo ERREUR pendant la restauration des donnees locales.
echo Tes donnees sont intactes dans :
echo    %SAUVEGARDE%
echo Recopie-les a la main dans %RACINE% (Metrologo ferme).
pause
goto FIN

:FIN
endlocal
