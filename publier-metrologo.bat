@echo off
setlocal
cd /d "%~dp0"

REM ============================================================
REM  Publication d'une nouvelle version de Metrologo
REM  (a lancer depuis un poste ou le lecteur M: est connecte)
REM ============================================================

REM --- Configuration (a adapter si besoin) --------------------
set APP_ID=Metrologo
set MAIN_EXE=Metrologo.exe
set PROJET=Metrologo.csproj
set "SORTIE_RESEAU=M:\exe_spe\Data_Metrologo\SUITE ASERTI Guillaume\Metrologo"
set SPLASH=Resources\splash.gif
REM Dossier de travail : surtout PAS nomme TMP (TMP = dossier temporaire de Windows :
REM le compilateur et vpk y ecriraient leurs fichiers, qui finiraient dans le paquet
REM de mise a jour -> paquet pollue, MAJ qui ne s'applique pas / boucle).
set "PUBDIR=%~dp0_publish_tmp"
REM ------------------------------------------------------------

echo ============================================
echo   Publication d'une nouvelle version Metrologo
echo ============================================
echo.

REM 1) Verifier l'outil vpk (Velopack CLI)
where vpk >nul 2>&1
if errorlevel 1 goto NO_VPK

REM 2) Verifier l'acces au reseau, puis creer le dossier de sortie si absent
if not exist "M:\exe_spe\Data_Metrologo\" goto NO_RESEAU
if not exist "%SORTIE_RESEAU%\" mkdir "%SORTIE_RESEAU%"

REM 3) Numero de version : auto-calcule d'apres la date/heure (toujours croissant)
REM    Format 1.AAMMJJ.HMM  ->  ex 1.260706.1043 (le 06/07/2026 a 10h43),
REM    1.261007.912 a 9h12 (pas de zero en tete : interdit en SemVer).
REM    Une version tapee a la main doit rester PLUS GRANDE que la derniere
REM    publiee (ex 1.3 < 1.260706.1451 : jamais vue comme une MAJ).
for /f "usebackq delims=" %%v in (`powershell -NoProfile -Command "'1.' + (Get-Date -Format 'yyMMdd') + '.' + [int](Get-Date -Format 'HHmm')"`) do set AUTO_VER=%%v
echo Version proposee (automatique, datee) : %AUTO_VER%
:DEMANDE_VER
set "VER="
set /p VER="Entree = accepter, ou tape un numero manuel (ex 2.0.0) : "
if "%VER%"=="" set VER=%AUTO_VER%
powershell -NoProfile -ExecutionPolicy Bypass -File "outils\verifier-version.ps1" -Version "%VER%" -Feed "%SORTIE_RESEAU%"
if errorlevel 1 goto DEMANDE_VER
echo Version retenue : %VER%

REM 4) Compilation autonome (runtime .NET embarque -> aucun prerequis .NET sur les postes)
echo.
echo == Compilation self-contained win-x64 ==
REM Libere les fichiers encore tenus par le serveur de compilation
dotnet build-server shutdown >nul 2>&1
if exist "%PUBDIR%" rmdir /s /q "%PUBDIR%"
dotnet publish "%PROJET%" -c Release -r win-x64 --self-contained true -o "%PUBDIR%"
if errorlevel 1 goto ECHEC_BUILD

REM 5) Empaquetage Velopack + depot sur le reseau
echo.
echo == Empaquetage vers %SORTIE_RESEAU% ==
vpk pack --packId %APP_ID% --packVersion %VER% --packDir "%PUBDIR%" --mainExe %MAIN_EXE% --outputDir "%SORTIE_RESEAU%" --splashImage "%SPLASH%"
if errorlevel 1 goto ECHEC_PACK

REM 6) Nettoyage
rmdir /s /q "%PUBDIR%"

echo.
echo ============================================
echo   Version %VER% publiee avec succes.
echo.
echo   - Nouveau poste : lancer %SORTIE_RESEAU%\Metrologo-win-Setup.exe
echo   - Postes deja installes : mise a jour automatique
echo     au prochain demarrage de Metrologo.
echo ============================================
echo.
pause
goto FIN

:NO_VPK
echo Outil vpk absent : installation en cours ...
dotnet tool install -g vpk
echo.
echo IMPORTANT : ferme cette fenetre, rouvre-en une nouvelle,
echo puis relance ce script pour que le PATH soit pris en compte.
pause
goto FIN

:NO_RESEAU
echo ERREUR : dossier reseau inaccessible :
echo    %SORTIE_RESEAU%
echo Connecte le lecteur M: avec le raccourci du bureau puis relance.
pause
goto FIN

:ECHEC_BUILD
echo ECHEC de la compilation.
pause
goto FIN

:ECHEC_PACK
echo ECHEC de l'empaquetage.
pause
goto FIN

:FIN
endlocal
