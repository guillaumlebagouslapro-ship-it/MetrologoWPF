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
  Components: asertools; \
  Check: FichierPresent('{#FeedRoot}\Asertools\Asertools-win-Setup.exe'); \
  StatusMsg: "Installation d'Asertools en cours..."; Flags: waituntilterminated

Filename: "{#FeedRoot}\Metrologo\Metrologo-win-Setup.exe"; Parameters: "--silent"; \
  Components: metrologo; \
  Check: FichierPresent('{#FeedRoot}\Metrologo\Metrologo-win-Setup.exe'); \
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

{ Utilise en Check pour n'installer un composant que si son Setup.exe existe. }
function FichierPresent(Param: String): Boolean;
begin
  Result := FileExists(Param);
  if not Result then
    MsgBox('Fichier d''installation introuvable :' + #13#10 + Param + #13#10#13#10 +
           'Ce composant sera ignore. Publie-le d''abord avec son script publier-*.bat.',
           mbInformation, MB_OK);
end;
