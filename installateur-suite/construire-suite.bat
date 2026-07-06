@echo off
setlocal
cd /d "%~dp0"

REM ============================================================
REM  Construit "Suite-ASERTI-Setup.exe" (installateur a cases a
REM  cocher) et le depose sur le reseau M:.
REM  Inno Setup est installe automatiquement s'il est absent.
REM ============================================================

set SORTIE=M:\exe_spe\Data_Metrologo\Suite

REM --- Localiser le compilateur Inno Setup (ISCC.exe) ---
call :FIND_ISCC
if defined ISCC goto ISCC_OK

REM --- Absent : installation automatique via winget ---
echo Inno Setup non detecte : installation automatique en cours...
echo (peut prendre 1 a 2 min ; une seule fois sur ce poste)
echo.
winget install --id JRSoftware.InnoSetup -e --accept-source-agreements --accept-package-agreements
echo.
call :FIND_ISCC
if defined ISCC goto ISCC_OK
goto NO_ISCC

:ISCC_OK
REM --- Dossier de sortie sur le reseau ---
if not exist "M:\exe_spe\Data_Metrologo\" goto NO_RESEAU
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
echo.
echo Inno Setup n'a pas pu etre installe automatiquement
echo (winget absent ou bloque sur ce poste).
echo Installe-le manuellement : https://jrsoftware.org/isdl.php
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

REM --- Sous-routine : cherche ISCC.exe aux emplacements habituels ---
:FIND_ISCC
set "ISCC="
if exist "%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe" set "ISCC=%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"
if not defined ISCC if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not defined ISCC if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
goto :eof

:FIN
endlocal
