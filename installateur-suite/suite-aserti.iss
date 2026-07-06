; ============================================================
;  Installateur "Suite ASERTI"
;  Assistant avec cases a cocher : Asertools et/ou Metrologo.
;
;  Il N'INSTALLE PAS les applications lui-meme : pour chaque composant
;  coche, il lance le Setup.exe Velopack correspondant depuis le dossier
;  reseau M:. Chaque application garde ainsi sa mise a jour automatique
;  independante. L'installateur suite reste minuscule et installe toujours
;  la derniere version publiee.
; ============================================================

#define SuiteVersion "1.0.0"
#define FeedRoot "M:\exe_spe\Data_Metrologo\SUITE ASERTI Guillaume"

[Setup]
AppId={{B7E5B2A0-1C3D-4E5F-9A8B-1234567890AB}
AppName=Suite ASERTI
AppVersion={#SuiteVersion}
AppPublisher=ASERTI
DefaultDirName={autopf}\SuiteASERTI
DisableDirPage=yes
DisableProgramGroupPage=yes
Uninstallable=no
CreateUninstallRegKey=no
OutputBaseFilename=Suite-ASERTI-Setup
PrivilegesRequired=lowest
WizardStyle=modern
SetupIconFile=..\Resources\logo.ico
LanguageDetectionMethod=none

[Languages]
Name: "fr"; MessagesFile: "compiler:Languages\French.isl"

[Types]
Name: "complet"; Description: "Installation complete (Asertools + Metrologo)"
Name: "perso";   Description: "Installation personnalisee"; Flags: iscustom

[Components]
Name: "asertools"; Description: "Asertools (multimetre / calibrateur)"; Types: complet perso
Name: "metrologo"; Description: "Metrologo (frequencemetre / rubidium)"; Types: complet perso

[Run]
Filename: "{#FeedRoot}\Asertools\Asertools-win-Setup.exe"; Parameters: "--silent"; \
  Components: asertools; Check: DoitInstallerAsertools; \
  StatusMsg: "Installation d'Asertools en cours..."; Flags: waituntilterminated

Filename: "{#FeedRoot}\Metrologo\Metrologo-win-Setup.exe"; Parameters: "--silent"; \
  Components: metrologo; Check: DoitInstallerMetrologo; \
  StatusMsg: "Installation de Metrologo en cours..."; Flags: waituntilterminated

[Code]
{ Verifie que le dossier reseau est accessible avant de commencer. }
function InitializeSetup(): Boolean;
begin
  if not DirExists('{#FeedRoot}') then
  begin
    MsgBox('Le dossier reseau des installations n''est pas accessible :' + #13#10 +
           '{#FeedRoot}' + #13#10#13#10 +
           'Connecte le lecteur M: (raccourci du bureau) puis relance cet installateur.',
           mbCriticalError, MB_OK);
    Result := False;
    exit;
  end;
  Result := True;
end;

{ Version installee d'une app Velopack (via l'entree de desinstallation), '' si absente. }
function VersionInstallee(PackId: String): String;
var v: String;
begin
  Result := '';
  if RegQueryStringValue(HKCU, 'Software\Microsoft\Windows\CurrentVersion\Uninstall\' + PackId, 'DisplayVersion', v) then
    Result := v
  else if RegQueryStringValue(HKLM, 'Software\Microsoft\Windows\CurrentVersion\Uninstall\' + PackId, 'DisplayVersion', v) then
    Result := v;
end;

{ L'app est-elle installee ? (registre, ou dossier d'install Velopack) }
function EstInstalle(PackId: String): Boolean;
begin
  Result := (VersionInstallee(PackId) <> '') or
            DirExists(ExpandConstant('{localappdata}\' + PackId + '\current'));
end;

{ Derniere version disponible dans le dossier reseau (lue sur les noms de .nupkg). }
function DerniereVersionDispo(FeedDir, PackId: String): String;
var
  fr: TFindRec;
  nom, ver, prefixe: String;
  p: Integer;
begin
  Result := '';
  prefixe := PackId + '-';
  if FindFirst(FeedDir + '\' + PackId + '-*-full.nupkg', fr) then
  begin
    try
      repeat
        nom := fr.Name;
        ver := Copy(nom, Length(prefixe) + 1, Length(nom));
        p := Pos('-full.nupkg', ver);
        if p > 0 then ver := Copy(ver, 1, p - 1);
        if (ver <> '') and (ver > Result) then Result := ver;
      until not FindNext(fr);
    finally
      FindClose(fr);
    end;
  end;
end;

{ Decide s'il faut (re)installer un composant, avec garde-fous et messages. }
function DemandeInstallation(PackId, Nom, FeedDir: String): Boolean;
var
  setupPath, installe, dispo: String;
begin
  setupPath := FeedDir + '\' + PackId + '-win-Setup.exe';
  if not FileExists(setupPath) then
  begin
    MsgBox('Fichier d''installation introuvable :' + #13#10 + setupPath + #13#10#13#10 +
           'Ce composant sera ignore. Publie-le d''abord avec son script publier-*.bat.',
           mbInformation, MB_OK);
    Result := False;
    exit;
  end;

  if not EstInstalle(PackId) then
  begin
    Result := True;   { pas installe -> installation directe }
    exit;
  end;

  installe := VersionInstallee(PackId);
  dispo := DerniereVersionDispo(FeedDir, PackId);

  if (installe <> '') and (dispo <> '') and (installe = dispo) then
  begin
    Result := (MsgBox(Nom + ' est deja installe et a jour (version ' + installe + ').' + #13#10 + #13#10 +
                      'Reinstaller quand meme (forcer la reinstallation) ?',
                      mbConfirmation, MB_YESNO) = IDYES);
  end
  else if dispo <> '' then
  begin
    if installe <> '' then
      Result := (MsgBox(Nom + ' est deja installe (version ' + installe + ').' + #13#10 +
                        'Version disponible : ' + dispo + '.' + #13#10 + #13#10 +
                        'Installer cette version maintenant ?',
                        mbConfirmation, MB_YESNO) = IDYES)
    else
      Result := (MsgBox(Nom + ' est deja installe sur ce poste.' + #13#10 + #13#10 +
                        'Installer / remplacer par la derniere version (' + dispo + ') ?',
                        mbConfirmation, MB_YESNO) = IDYES);
  end
  else
  begin
    Result := (MsgBox(Nom + ' est deja installe sur ce poste.' + #13#10 + #13#10 +
                      'Reinstaller quand meme ?',
                      mbConfirmation, MB_YESNO) = IDYES);
  end;
end;

function DoitInstallerAsertools: Boolean;
begin
  Result := DemandeInstallation('Asertools', 'Asertools', '{#FeedRoot}\Asertools');
end;

function DoitInstallerMetrologo: Boolean;
begin
  Result := DemandeInstallation('Metrologo', 'Metrologo', '{#FeedRoot}\Metrologo');
end;
