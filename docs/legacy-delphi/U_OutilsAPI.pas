unit U_OutilsAPI;

interface

uses
    Dialogs, SysUtils, ShellAPI, Windows, Forms, Controls, ExtCtrls, Classes, DateUtils;

function cbCopyFile(spSource, spDest: string; bpDispWindow: Boolean; spText: string;
            bpMove: Boolean = False): Boolean;

function cbDeleteFile(spSource: string; bpDispWindow: Boolean; spText: string;
            bpDispErr: Boolean = True): Boolean;

function cbMessageDlg(const Caption: string; const Msg: string;
            DlgType: TMsgDlgType; Buttons: TMsgDlgButtons; DefaultBtn: TMsgDlgBtn): Word;

function DiskIsPresent(cpUnit: Char): Boolean;

function KillProcess(const ProcessName: string; const KillOnlyIfInactive: Boolean = False): Boolean;

procedure Modif(ctrl: TWinControl; Allow: Boolean);

function WaitAndCloseMessageBox(spCaption: string; ipIDCtrl: Integer; ipTimeOut: Integer): Boolean;

function GetFileTimes(const FileName: string; var Created: TDateTime;
            var Accessed: TDateTime; var Modified: TDateTime): Boolean;

procedure DisplaySystemError(lpErrID: DWORD; spMessage: string);

function GetTempoFileName(spPath, spPrefixe: String): string;

Function ApplicationVersion(): string;


type
    CloseThread = class(TThread)
    private
        FWindowName: string;
        FCtrlID: Integer;
        FTimeOut: Integer;

    public
        property WindowName: string read FWindowName write FWindowName;
        property CtrlID: Integer read FCtrlID write FCtrlID;
        property TimeOut: Integer read FTimeOut write FTimeOut;

    protected
        procedure Execute; override;

    end;

    PhaseCloseWindow = (AttenteFenetre, ClicBouton, RelacheButton, AttenteFermeture);


const
    BM_CLICK: Integer = $000000F5;
    BM_SETSTATE: Integer = $000000F3;
    BN_CLICKED: Integer = 0;
    CRLF = #13#10;

implementation

uses StrUtils, TlHelp32, TypInfo, JvDBSearchComboBox, AdvEdit, JvDBLookup,
    JvLabel, StdCtrls, E2MColumnGrid, E2MGrid, JvDBDateTimePicker, DBHTMLBtns,
    JvTransparentButton, JvRadioGroup;

var
    CloseWindowThread: CloseThread;


{*-------------------------------------------------------------------------------
  Procedure: cbCopyFile
  @Author:   cb le 06/08/2008
  @Param:    spSource, spDest: string; bpDispWindow: Boolean; spText: string
  @Result:   Boolean
-------------------------------------------------------------------------------}
function cbCopyFile(spSource, spDest: string; bpDispWindow: Boolean; spText: string; bpMove: Boolean = False): Boolean;
var
    llRC: integer;
    slFileOp: SHFILEOPSTRUCT;

begin

    with slFileOp do begin
        Wnd := Application.Handle;
        if bpMove then
            wFunc := FO_MOVE
        else
            wFunc := FO_COPY;
        pFrom := Pchar(spSource + chr(0) + chr(0));
        pTo := PChar(spDest + chr(0) + chr(0));
        hNameMappings := nil;

        if bpDispWindow then begin
            fFlags := FOF_SIMPLEPROGRESS or FOF_NOCONFIRMATION;
            lpszProgressTitle := PChar(spText);
        end
        else
            fFlags := FOF_NOCONFIRMATION or FOF_SILENT;

    end;

    llRc := SHFileOperation(slFileOp);

    if llRC <> 0 then begin
        if Length(spSource) > 128 then spSource := LeftStr(spSource, 128);
        MessageDlg('Erreur lors de la copie du (des) fichier(s) ' + spSource + ' en '
            + spDest, mtError, [mbOk], 0);
    end;

    Result := (llRC = 0);

end;

{*-------------------------------------------------------------------------------
  Procedure: cbDeleteFile
  @Author:   cb le 06/08/2008
  @Param:    spSource: string; bpDispWindow: Boolean; spText: string
  @Result:   Boolean
-------------------------------------------------------------------------------}
function cbDeleteFile(spSource: string; bpDispWindow: Boolean; spText: string; bpDispErr: Boolean = True): Boolean;
var
    llRC: integer;
    slFileOp: SHFILEOPSTRUCT;

begin

    with slFileOp do begin
        Wnd := Application.Handle;
        wFunc := FO_DELETE;
        pFrom := Pchar(spSource + chr(0) + chr(0));
        pTo := nil;
        hNameMappings := nil;

        if bpDispWindow then begin
            fFlags := FOF_SIMPLEPROGRESS or FOF_NOCONFIRMATION;
            lpszProgressTitle := PChar(spText);
        end
        else
            fFlags := FOF_NOCONFIRMATION or FOF_SILENT;

    end;

    llRc := SHFileOperation(slFileOp);

    if llRC <> 0 then begin
        if Length(spSource) > 128 then spSource := LeftStr(spSource, 128);
        MessageDlg('Erreur lors de la suppression du (des) fichier(s) ' + spSource, mtError, [mbOk], 0);
    end;

    Result := (llRC = 0);

end;

{*-------------------------------------------------------------------------------
  Procedure: cbMessageDlg
  @Author:   cb le 06/08/2008
  @Param:    const Caption: string; const Msg: string; DlgType: TMsgDlgType; Buttons: TMsgDlgButtons; DefaultBtn: TMsgDlgBtn
  @Result:   Word
-------------------------------------------------------------------------------}
function cbMessageDlg(const Caption: string; const Msg: string; DlgType:
    TMsgDlgType; Buttons: TMsgDlgButtons; DefaultBtn: TMsgDlgBtn): Word;
var
    ltMessage: MSGBOXPARAMS;

begin

    ltMessage.lpszCaption := PChar(Caption);
    ltMessage.lpszText := PChar(Msg);
    ltMessage.hwndOwner := Application.Handle;

    case DlgType of
        mtWarning: ltMessage.dwStyle := MB_ICONWARNING;
        mtError: ltMessage.dwStyle := MB_ICONERROR;
        mtInformation: ltMessage.dwStyle := MB_ICONINFORMATION;
        mtConfirmation: ltMessage.dwStyle := MB_ICONQUESTION;
    end;

    while True do {// Définit boutons à afficher} begin
        if Buttons = [mbYes..mbNo] then {// Yes No} begin
            ltMessage.dwStyle := ltMessage.dwStyle + MB_YESNO;
            if DefaultBtn = mbNo then
                ltMessage.dwStyle := ltMessage.dwStyle + MB_DEFBUTTON2;
            break;
        end;

        if Buttons = mbYesNoCancel then {// Yes No Cancel} begin
            ltMessage.dwStyle := ltMessage.dwStyle + MB_YESNOCANCEL;
            case DefaultBtn of
                mbNo: ltMessage.dwStyle := ltMessage.dwStyle + MB_DEFBUTTON2;
                mbCancel: ltMessage.dwStyle := ltMessage.dwStyle + MB_DEFBUTTON3;
            end;
            break;
        end;

        if Buttons = mbOKCancel then {// Ok Cancel} begin
            ltMessage.dwStyle := ltMessage.dwStyle + MB_OKCANCEL;
            if DefaultBtn = mbCancel then
                ltMessage.dwStyle := ltMessage.dwStyle + MB_DEFBUTTON2;
            break;
        end;

        if Buttons = mbAbortRetryIgnore then {// Abort Retry Ignore} begin
            ltMessage.dwStyle := ltMessage.dwStyle + MB_ABORTRETRYIGNORE;
            case DefaultBtn of
                mbRetry: ltMessage.dwStyle := ltMessage.dwStyle + MB_DEFBUTTON2;
                mbIgnore: ltMessage.dwStyle := ltMessage.dwStyle + MB_DEFBUTTON3;
            end;
            break;
        end;

        if Buttons = [mbRetry, mbCancel] then {// Retry Cancel} begin
            ltMessage.dwStyle := ltMessage.dwStyle + MB_RETRYCANCEL;
            if DefaultBtn = mbCancel then
                ltMessage.dwStyle := ltMessage.dwStyle + MB_DEFBUTTON2;
            break;
        end;

        ltMessage.dwStyle := ltMessage.dwStyle + MB_OK;
        break;

    end;

    //ltMessage.dwStyle := Buttons; //MB_YESNOCANCEL + MB_ICONERROR;
    ltMessage.cbSize := sizeof(ltMessage);

    Result := Word(MessageBoxIndirect(ltMessage));

end;

{*------------------------------------------------------------------------------
    Fonction indiquant si un disque est présent

@param cpUnit : Nom de l'unité (1 caractère) ou Espace pour unité courante
@Return

@Author CB
-------------------------------------------------------------------------------}

function DiskIsPresent(cpUnit: Char): Boolean;
var
    ErrorMode: Word;
    ilUnit: Integer;
    clUnit: Char;
    ordUnit: Cardinal;

begin

    Result := False;
    clUnit := UpCase(cpUnit);

    if clUnit = ' ' then
        ilUnit := 0
    else begin
        ordUnit := ord(clUnit);
        if (ordUnit < $41) or (ordUnit > $5A) then begin
            MessageDlg('Nom d''unité incorrect !', mtWarning, [mbOk], 0);
            exit;
        end;
        ilUnit := ordUnit - $40;
    end;

    ErrorMode := SetErrorMode(SEM_FAILCRITICALERRORS);
        // Désactive la gestion des erreurs
    try
        Result := DiskSize(ilUnit) <> -1; // DiskSize(0)= unité en cours, 1= A, 2= B

    finally
        SetErrorMode(ErrorMode); // Réactive la gestion des erreurs

    end;
end;

{*------------------------------------------------------------------------------
    Termine un processus portant le nom donné.
    @param ProcessName Nom du processus à quitter.
    @param KillOnlyIfInactive Quitte le process uniquement si celui-ci n'a pas de fenêtres ouvertes.
    @return Valeur indiquant si le process à été quitter ou non.
    @author JT
-------------------------------------------------------------------------------}

function KillProcess(const ProcessName: string; const KillOnlyIfInactive: Boolean =
    False): Boolean;
var
    ProcessEntry32: TProcessEntry32;
    HSnapShot: THandle;
    HProcess: THandle;
begin
    Result := False;

    HSnapShot := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if HSnapShot = 0 then exit;

    ProcessEntry32.dwSize := sizeof(ProcessEntry32);
    if Process32First(HSnapShot, ProcessEntry32) then
        repeat
            if CompareText(ProcessEntry32.szExeFile, ProcessName) = 0 then begin
                HProcess := OpenProcess(PROCESS_TERMINATE, False,
                    ProcessEntry32.th32ProcessID);
                if HProcess <> 0 then begin
                    Result := TerminateProcess(HProcess, 0);
                    CloseHandle(HProcess);
                end;
                Break;
            end;
        until not Process32Next(HSnapShot, ProcessEntry32);

    CloseHandle(HSnapshot);
end;

{*-------------------------------------------------------------------------------
  Procedure: Modif
  @Author:   cb le 06/08/2008
  @Param:    ctrl: TWinControl; Allow: Boolean
  @Result:   None
-------------------------------------------------------------------------------}
procedure Modif(ctrl: TWinControl; Allow: Boolean);
var
    i, j: Integer;
    txt: string;
begin
    with ctrl do
        for i := 0 to ControlCount - 1 do begin
            txt := Controls[i].Name;
            if (Controls[i].Tag = 1) or (Controls[i] is TLabel) or (Controls[i] is
                TJvLabel) then
                Continue

            else if Controls[i] is TE2MColumnGrid then
                for j := 0 to TE2MColumnGrid(Controls[i]).Columns.Count - 1 do begin
                    if TE2MColumnGrid(Controls[i]).Columns.Items[j].Tag = 1 then
                        Continue;
                    TE2MColumnGrid(Controls[i]).Columns.Items[j].ReadOnly := not
                        Allow;
                end

            else if (Controls[i] is TAdvEdit)
                or (Controls[i] is TJvRadioGroup) then
                SetPropValue(Controls[i], 'Readonly', not Allow)

            else if (Controls[i] is TJvDBLookupCombo)
                or (Controls[i] is TJvDBDateTimePicker)
                or (Controls[i] is TJvDBSearchComboBox)
                or (Controls[i] is TDBHTMLRadioGroup) then
                SetPropValue(Controls[i], 'Enabled', Allow)

            else if Controls[i] is TJvTransparentButton then
                Controls[i].Visible := Allow

            else if Controls[i].ClassType().InheritsFrom(TWinControl)
                and (TWinControl(Controls[i]).ControlCount > 0) then
                if Controls[i].Tag = 2 then
                    SetPropValue(Controls[i], 'Enabled', Allow)
                else
                    Modif(Controls[i] as TWinControl, Allow)

            else if (GetPropInfo(Controls[i].ClassInfo, 'Readonly') <> nil) then
                SetPropValue(Controls[i], 'Readonly', not Allow)

            else if GetPropInfo(Controls[i].ClassInfo, 'Enabled') <> nil then
                SetPropValue(Controls[i], 'Enabled', Allow);
        end;
end;


{*-------------------------------------------------------------------------------
  Procedure: WaitAndCloseMessageBox
  @Author:   cb le 08/08/2008
  @Param:    spCaption: string; ipIDCtrl, ipTimeOut: Integer
  @Result:   Boolean
-------------------------------------------------------------------------------}
function WaitAndCloseMessageBox(spCaption: string; ipIDCtrl: Integer; ipTimeOut: Integer): Boolean;
var
    ilThreadResult: Integer;

begin

    CloseWindowThread := CloseThread.Create(True);
    with CloseWindowThread do begin
        FreeOnTerminate := False;
        WindowName := spCaption;
        CtrlID := ipIDCtrl;
        TimeOut := ipTimeOut;
        Resume();

        ilThreadResult := WaitFor();
        Free;

        Result := (ilThreadResult = 0);
    end;

end;

{ CloseThread }


{*-------------------------------------------------------------------------------
  Procedure: CloseThread.Execute
  @Author:   cb le 08/08/2008
  @Param:    None
  @Result:   None
-------------------------------------------------------------------------------}
procedure CloseThread.Execute;
var
    ilNumPhase: PhaseCloseWindow;
    Fin: TDateTime;
    hlHandle: HWND;

begin

    ReturnValue := -1;
    hlHandle := 0;
    Fin := IncSecond(Now, TimeOut);
    ilNumPhase := AttenteFenetre;

    while (not Self.Terminated) and (SecondsBetween(Now, Fin) > 0) do begin
 //       Application.ProcessMessages;
        Sleep(500);
        case ilNumPhase of
            AttenteFenetre: begin
                    hlHandle := FindWindow(nil, PChar(WindowName)); /// Récupère handle de la fenêtre dont on a le caption
                    if hlHandle <> 0 then /// Si Handle OK, on vérifie si cette fenêtre contient bien le contrôle sur lequel on veut cliquer
                        if GetDlgItem(hlHandle, CtrlID) <> 0 then
                            Inc(ilNumPhase)
                        else
                            hlHandle := 0;
                end;

            ClicBouton: begin
                    SendDlgItemMessage(hlHandle, CtrlID, BM_CLICK, 0, 0);
//                    Application.ProcessMessages;
                    Inc(ilNumPhase);
                end;

            RelacheButton: begin
                    SendDlgItemMessage(hlHandle, CtrlID, BM_CLICK, 0, 0);
//                    Application.ProcessMessages;
                    Inc(ilNumPhase);
                end;

            AttenteFermeture:
                if not IsWindow(hlHandle) then begin
                    ReturnValue := 0;
                    Break;
                end;
        end;

    end;

end;

{*------------------------------------------------------------------------------
    Fonction retournant les différents temps d'un fichier.
    @comment Trouvée sur Developpez.com
    @param FileName Chemin d'accès complet du fichier.
    @param Created Date de création
    @param Accessed Date d'accès
    @param Modified Date de modification
    @return Valeur indiquant si les trois dates ont pu etre lues.
    @author jt
------------------------------------------------------------------------------*}
function GetFileTimes(const FileName: string; var Created: TDateTime;
    var Accessed: TDateTime; var Modified: TDateTime): Boolean;
var
    h: THandle;
    Info1, Info2, Info3: TFileTime;
    SysTimeStruct: SYSTEMTIME;
    TimeZoneInfo: TTimeZoneInformation;
    Bias: Double;
begin
    Result := False;
    Bias := 0;
    h := FileOpen(FileName, fmOpenRead or fmShareDenyNone);
    if h > 0 then begin
        try
            if GetTimeZoneInformation(TimeZoneInfo) <> $FFFFFFFF then
                Bias := TimeZoneInfo.Bias / 1440; // 60x24
            GetFileTime(h, @Info1, @Info2, @Info3);
            if FileTimeToSystemTime(Info1, SysTimeStruct) then
                Created := SystemTimeToDateTime(SysTimeStruct) - Bias;
            if FileTimeToSystemTime(Info2, SysTimeStruct) then
                Accessed := SystemTimeToDateTime(SysTimeStruct) - Bias;
            if FileTimeToSystemTime(Info3, SysTimeStruct) then
                Modified := SystemTimeToDateTime(SysTimeStruct) - Bias;
            Result := True;
        finally
            FileClose(h);
        end;
    end;
end;



{*-----------------------------------------------------------------------------
  Procedure:  DisplaySystemError
  @Author:    cb le 24/nov./2008
  @Param      lpErrID
  @Param      spMessage
  @Result:    None
-----------------------------------------------------------------------------*}
procedure DisplaySystemError(lpErrID: DWORD; spMessage: string);
var
    llErrID: DWORD;
    slTexte, slCodeErr: string;

begin

    if lpErrID = 0 then
        llErrID := GetLastError
    else
        llErrID := lpErrID;

    slTexte := StringOfChar(#0, 512);
    slCodeErr := 'Code Erreur : ' + IntToStr(llErrID);
    FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, nil, llErrID, LANG_NEUTRAL, PChar(slTexte), 512, nil);
    Application.MessageBox(PChar(spMessage + CRLF + slCodeErr + CRLF + slTexte), 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);

end;

{*-----------------------------------------------------------------------------

  @Author:    cb le 13/mars/2009
  @Param      spPath
  @Param      spPrefixe
  @Result:    string
-----------------------------------------------------------------------------*}
function GetTempoFileName(spPath, spPrefixe: String): string;
var
    slFic: string;

begin

    slFic := StringOfChar(#0, MAX_PATH);
    GetTempFileName(PChar(spPath), PChar(spPrefixe), 0, PChar(slFic));
    Application.ProcessMessages;
    if FileExists(slFic) then
        cbDeleteFile(slFic, False, '', True);

    Result := ChangeFileExt(slFic, '.ntx');

end;


{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.ApplicationVersion
  @Author:   cb le 07/05/2009
  @Param:    None
  @Result:   string
-------------------------------------------------------------------------------}
function ApplicationVersion: string;
Var
    VerInfoSize, VerValueSize, Dummy: DWord;
    VerInfo: Pointer;
    VerValue: PVSFixedFileInfo;

Begin

    Result := '';

    VerInfoSize := GetFileVersionInfoSize(PChar(ParamStr(0)), Dummy);

    If VerInfoSize <> 0 Then Begin //Les info de version sont inclues

        GetMem(VerInfo, VerInfoSize); //On alloue de la mémoire pour un pointeur sur les info de version :
        // On récupère ces informations
        GetFileVersionInfo(PChar(ParamStr(0)), 0, VerInfoSize, VerInfo);
        VerQueryValue(VerInfo, '\', Pointer(VerValue), VerValueSize);

        With VerValue^ Do Begin // On traite les informations ainsi récupérées
            Result := IntTostr(dwFileVersionMS Shr 16);
            Result := Result + '.' + IntTostr(dwFileVersionMS And $FFFF);
            Result := Result + '.' + IntTostr(dwFileVersionLS Shr 16);
            Result := Result + '.' + IntTostr(dwFileVersionLS And $FFFF);
        End;

        FreeMem(VerInfo, VerInfoSize); // On libère la place précédemment allouée

    End

    Else // Les infos de version ne sont pas inclues
        Raise EAccessViolation.Create('Les informations de version de sont pas inclues');

End;


end.

