@echo off
setlocal
cd /d "%~dp0"

REM ============================================================
REM  Construit "Suite-ASERTI-Setup.exe" (installateur a cases a
REM  cocher) et le depose sur le reseau M:.
REM  Prerequis : Inno Setup 6 installe
REM              (winget install JRSoftware.InnoSetup)
REM ============================================================

set SORTIE=M:\exe_spe\Data_Metrologo\Suite

REM --- Localiser le compilateur Inno Setup (ISCC.exe) ---
set ISCC=
if exist "%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe" set "ISCC=%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"
if not defined ISCC if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not defined ISCC if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if not defined ISCC goto NO_ISCC

REM --- Dossier de sortie sur le reseau ---
if not exist "M:\" goto NO_RESEAU
if not exist "%SORTIE%\" mkdir "%SORTIE%"

echo ============================================
echo   Construction de la Suite ASERTI
echo ============================================
echo.
"%ISCC%" /O"%SORTIE%" suite-aserti.iss
if errorlevel 1 goto ECHEC

echo.
echo ============================================
echo   Termine.
echo   Installateur : %SORTIE%\Suite-ASERTI-Setup.exe
echo   (les utilisateurs lancent ce fichier et cochent ce qu'ils veulent)
echo ============================================
echo.
pause
goto FIN

:NO_ISCC
echo Inno Setup introuvable.
echo Installe-le avec :  winget install JRSoftware.InnoSetup
echo puis relance ce script.
pause
goto FIN

:NO_RESEAU
echo ERREUR : lecteur M: inaccessible. Connecte-le (raccourci du bureau) puis relance.
pause
goto FIN

:ECHEC
echo ECHEC de la compilation Inno Setup.
pause
goto FIN

:FIN
endlocal
