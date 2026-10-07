@echo off
setlocal
cd /d "%~dp0"

REM ============================================================
REM  Publication d'une nouvelle version de Metrologo
REM  (a lancer depuis un poste ou le lecteur M: est connecte)
REM
REM  Numero de version : MAJOR.MINOR (2.0, 2.1, 2.2...)
REM    - Entree : MINOR + 1 (mise a jour courante)
REM    - M      : MAJOR + 1 (mise a jour importante, ex 3.0)
REM  Une ligne de description est demandee : elle apparait dans
REM  l'historique (Accueil > Mises a jour) sur chaque poste.
REM ============================================================

REM --- Configuration (a adapter si besoin) --------------------
set APP_ID=Metrologo
set MAIN_EXE=Metrologo.exe
set PROJET=Metrologo.csproj
set "SORTIE_RESEAU=M:\exe_spe\Data_Metrologo\SUITE ASERTI Guillaume\Metrologo"
set SPLASH=Resources\splash.gif
REM Icone (raccourci bureau, installateur, Applications installees) = celle de l'exe
set ICONE=Resources\logo.ico
REM Dossier de travail : surtout PAS nomme TMP (TMP = dossier temporaire de Windows :
REM le compilateur et vpk y ecriraient leurs fichiers, qui finiraient dans le paquet
REM de mise a jour -> paquet pollue, MAJ qui ne s'applique pas / boucle).
set "PUBDIR=%~dp0_publish_tmp"
set "INFO=%~dp0_publish_info"
REM Empaquetage fait EN LOCAL puis copie sur le reseau en une fois : vpk pack directement
REM sur M: y relisait/ecrivait ~250 Mo par petits morceaux (5-6 min au lieu de ~10 s).
set "RELDIR=%~dp0_publish_releases"
REM Nombre de versions gardees sur le reseau (les plus anciennes sont supprimees).
set GARDER=5
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

REM 3) Numero de version + description (lus/valides par outils\preparer-version.ps1 :
REM    superieur a la derniere version publiee, pas de zero en tete).
if exist "%INFO%" rmdir /s /q "%INFO%"
powershell -NoProfile -ExecutionPolicy Bypass -File "outils\preparer-version.ps1" -AppId %APP_ID% -Feed "%SORTIE_RESEAU%" -OutDir "%INFO%"
if errorlevel 1 goto ANNULE
if not exist "%INFO%\version.txt" goto ANNULE
set /p VER=<"%INFO%\version.txt"
set FULLVER=%VER%.0
set NOTES=
if exist "%INFO%\notes.md" set NOTES=--releaseNotes "%INFO%\notes.md"

REM 4) Compilation autonome (runtime .NET embarque -> aucun prerequis .NET sur les postes)
echo.
echo == Compilation self-contained win-x64 (version %FULLVER%) ==
REM Libere les fichiers encore tenus par le serveur de compilation
dotnet build-server shutdown >nul 2>&1
if exist "%PUBDIR%" rmdir /s /q "%PUBDIR%"
dotnet publish "%PROJET%" -c Release -r win-x64 --self-contained true -p:Version=%FULLVER% -o "%PUBDIR%"
if errorlevel 1 goto ECHEC_BUILD

REM 5) Empaquetage EN LOCAL (rapide), puis depot sur le reseau en une fois
echo.
if exist "%RELDIR%" rmdir /s /q "%RELDIR%"
echo == Recuperation de la derniere version publiee (base du delta) ==
vpk download local --path "%SORTIE_RESEAU%" --outputDir "%RELDIR%"
if errorlevel 1 echo    (aucune version sur le reseau : pas de delta, normal a la 1re publication)
echo.
echo == Empaquetage (local) ==
vpk pack --packId %APP_ID% --packVersion %FULLVER% --packDir "%PUBDIR%" --mainExe %MAIN_EXE% --outputDir "%RELDIR%" --splashImage "%SPLASH%" --icon "%ICONE%" %NOTES%
if errorlevel 1 goto ECHEC_PACK
echo.
echo == Depot sur %SORTIE_RESEAU% (garde les %GARDER% dernieres versions) ==
vpk upload local --path "%SORTIE_RESEAU%" --outputDir "%RELDIR%" --keepMaxReleases %GARDER%
if errorlevel 1 goto ECHEC_DEPOT

REM 6) Nettoyage
rmdir /s /q "%PUBDIR%"
rmdir /s /q "%RELDIR%"
rmdir /s /q "%INFO%"

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

:ANNULE
echo Publication annulee.
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

:ECHEC_DEPOT
echo ECHEC du depot sur le reseau. Le paquet est intact dans :
echo    %RELDIR%
echo Verifie le lecteur M: puis relance la publication.
pause
goto FIN

:ECHEC_PACK
echo ECHEC de l'empaquetage.
pause
goto FIN

:FIN
endlocal
