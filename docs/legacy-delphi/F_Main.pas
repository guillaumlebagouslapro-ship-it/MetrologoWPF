Unit F_Main;

Interface

Uses
    ExcelXP, Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,ShellAPI,
    Dialogs, ToolWin, ComCtrls, JvExComCtrls, JvToolBar, ImgList,
    ExtCtrls, U_OutilsAPI, OleServer, Registry,
    U_DeclarationsMETROLOGO, dpib32, E2M_IniFile, F_Patience, uE2M,
    U_OutilsIEEE, F_MdpValidation, F_ChoixFreqGene, F_Configuration, F_VoiesMux,
    StdCtrls, F_ConfigStanford, F_ConfigRacal, JvExStdCtrls,
    DB, ADODB, AdvPanel, E2MPanel, F_SelectionGate,
    F_GetNbMesures, F_GetInfosDerive, E2MMutex, F_MtxBusy, jpeg,
    JvComponentBase, JvSearchFiles, F_ConfigEIP, E2MBackgroundWorker, Math,
    F_SaisieValFres, F_ParamsIncert, F_DispIncert, F_SelValidation, Grids,
    BaseGrid, AdvGrid, AdvCGrid, E2MColumnGrid, U_OutilsSystemes, Menus,
    AdvMenus, F_ChoixRubidium, IdAntiFreezeBase, IdAntiFreeze,
    IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdFTP, IdFTPCommon,
    StrUtils, AdvEdit, AdvEdBtn, JvDateTimePicker, F_SaisieMarcheHebdo,
    F_SaisieFreqAutres, F_DateDebutGraphe, IdExplicitTLSClientServerBase, pngimage;

Type
    TfrmMain = Class(TForm)
        lstIcoMenu: TImageList;
        tbMainMenu: TJvToolBar;
        btnExec: TToolButton;
        btnSep1: TToolButton;
        btnRelance: TToolButton;
        btnSep2: TToolButton;
        btnStop: TToolButton;
        btnSep3: TToolButton;
        btnInit: TToolButton;
        btnSep4: TToolButton;
        btnRecapitulation: TToolButton;
        btnSep6: TToolButton;
        btnParamIncert: TToolButton;
        btnSep7: TToolButton;
        btnLoadLast: TToolButton;
        btnSep8: TToolButton;
        btnValidation: TToolButton;
        btnExploitation: TToolButton;
        btnSep10: TToolButton;
        btnPrint: TToolButton;
        btnSep11: TToolButton;
        btnRubidium: TToolButton;
        btnSep12: TToolButton;
        btnQuit: TToolButton;
        imgFondAppli: TImage;
        XLApp: TExcelApplication;
        tmrDelai: TTimer;
        iniMetrologo: TE2M_IniFile;
        E2M: TE2M;
        cnxMetrologo: TADOConnection;
        qryMetrologo: TADOQuery;
        pnlCaptionMode: TE2MPanel;
        pnlCaptionRubidium: TE2MPanel;
        mtxCycleMesure: TE2MMutex;
        cnxEgide: TADOConnection;
        wbMetrologo: TExcelWorkbook;
        wsMesure: TExcelWorksheet;
        wsRecap: TExcelWorksheet;
        wsModele: TExcelWorksheet;
        dlgOpenLast: TOpenDialog;
        chartDerive: TExcelChart;
        pnlInfosGenerales: TE2MPanel;
        mmoInfosGenes: TMemo;
        splH: TSplitter;
        splV: TSplitter;
        mnuKillDerive: TAdvPopupMenu;
        optMenuKillDerive: TMenuItem;
        optDummy: TMenuItem;
        optMenuCancel: TMenuItem;
        btnSuiviRubidiums: TToolButton;
        lblDebug: TLabel;
        btnMajRubiAutre: TToolButton;
        tmrBesancon: TTimer;
    ftpe2m: TIdFTP;
        idntfrz1: TIdAntiFreeze;
        mnuMemo: TAdvPopupMenu;
        MenuItem2: TMenuItem;
        MenuItem3: TMenuItem;
        mnuClearMemo: TMenuItem;
        mnuEnregMemo: TMenuItem;
        spSaveGPSDaily: TADOStoredProc;
        tmrReminder: TTimer;
        pnlGlobal: TPanel;
        pnlInfosDerive: TE2MPanel;
        grdDerivesEnCours: TE2MColumnGrid;
        pnlConversion: TE2MPanel;
        dtpDate: TJvDateTimePicker;
        edtDateJulienne: TAdvEditBtn;
        btnGraphesderive: TToolButton;
        wbSuiviRubi: TExcelWorkbook;
        wsSuiviRubi: TExcelWorksheet;
        mnuGestFRef: TAdvPopupMenu;
        mnuForceLoadFTP: TMenuItem;
        mnuN1: TMenuItem;
        mnuModifValue: TMenuItem;
        mnuN2: TMenuItem;
        mnuAnnuler: TMenuItem;
        Procedure FormCreate(Sender: TObject);
        Procedure btnQuitClick(Sender: TObject);
        Procedure FormCloseQuery(Sender: TObject; Var CanClose: Boolean);
        Procedure tmrDelaiTimer(Sender: TObject);
        Function AddMacroComplementaire(spXLA: String): Boolean;
        Function InitAppareils(): Boolean;
        Procedure btnExecClick(Sender: TObject);
        Procedure btnValidationClick(Sender: TObject);
        Procedure btnExploitationClick(Sender: TObject);
        Procedure FormClose(Sender: TObject; Var Action: TCloseAction);
        Procedure btnRelanceClick(Sender: TObject);
        Procedure btnLoadLastClick(Sender: TObject);
        Procedure XLAppWorkbookBeforeClose(ASender: TObject; Const Wb: _Workbook; Var Cancel: WordBool);
        Procedure btnInitClick(Sender: TObject);
        Procedure btnStopClick(Sender: TObject);
        Procedure btnParamIncertClick(Sender: TObject);
        Procedure grdDerivesEnCoursRightClickCell(Sender: TObject; ARow, ACol: Integer);
        Procedure optMenuKillDeriveClick(Sender: TObject);
        Procedure btnRubidiumClick(Sender: TObject);
        Procedure GestionRubidiumACtif();
        Procedure tmrBesanconTimer(Sender: TObject);
        Procedure mnuClearMemoClick(Sender: TObject);
        Procedure mnuEnregMemoClick(Sender: TObject);
        Procedure ftpe2mAfterGet(ASender: TObject; VStream: TStream);
        Procedure tmrReminderTimer(Sender: TObject);
        Procedure dtpDateChange(Sender: TObject);
        Procedure edtDateJulienneClickBtn(Sender: TObject);
        Procedure btnSuiviRubidiumsClick(Sender: TObject);
        Procedure btnMajRubiAutreClick(Sender: TObject);
        Procedure btnGraphesderiveClick(Sender: TObject);
//        procedure pnlCaptionRubidiumDblClick(Sender: TObject);
        Procedure ProtectionClasseurMesure(bpProtect: Boolean);
        procedure pnlCaptionRubidiumMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
        procedure mnuForceLoadFTPClick(Sender: TObject);
        procedure mnuModifValueClick(Sender: TObject);


    Private
        bForceQuit: Boolean;                                                    // Si vrai, ne demande pas confirmation avant de quitter l'appli
        bXLConnected: Boolean;                                                  // Indique si XL est connecté
        bIEEEOk: Boolean;                                                       // Flag IEEE Ok
        bFlagRelanceMesure: Boolean;                                            // Relance mesure identique
        bArretDemande: Boolean;
        bGetDone: Boolean;                                                      // Flag indiquant la fin du Get du client FTP
        bReminderFreqRefAutres: Boolean;                                        // Indique s'il faut afficher régulièrement le message demandant de saisir les F fr réf. des rubidiums non actifs
        bReminderSaisieFRefAllouis: Boolean;                                    // Indique s'il faut afficher régulièrement le message demandant de saisir la marche hebdo pour calcul FRef allouis
        bCalculEnCours: Boolean;
        bForceLoadBesancon: Boolean;                                            // Flag sservant au forcage de le lecture du fichier de Besancon sur le ftp E2M

        iNbDeriveEnCours: Integer;                                              // Nb de dérives en cours

        dIncertResolution: Double;
        dIncertAutres: Double;
        dFLue: Double;

        MuxHP: TMultiplexeur;                                                   // Objet 'Mux'
        MesureEnCours: TMesure;                                                 // Infos relatives à la mesure en cours
        RubidiumEnCours: String;
        IDRubidiumEnCours: Integer;
        bAvecGPS: Boolean;

    Public

        Procedure ExecMesure(oMes: TMesure);

        Function GetRangeName(oWs: TExcelWorksheet; ipIdx: enZonesNommees): String;
        Procedure SetRangeValue(oWs: TExcelWorksheet; ipIdxRangeName: enZonesNommees; spValue: String); Overload;
        Procedure SetRangeValue(oWs: TExcelWorksheet; ipIdxRangeName: enZonesNommees; dpValue: Double); Overload;
        Procedure SetRangeValue(oWs: TExcelWorksheet; ipIdxRangeName: enZonesNommees; ipValue: Integer); Overload;

        Function GetRangeValueDouble(oWs: TExcelWorksheet; ipIdxRangeName: enZonesNommees): Double;
        Function GetRangeValueString(oWs: TExcelWorksheet; ipIdxRangeName: enZonesNommees): String;

        Procedure SetFormulavalue(oWs: TExcelWorksheet; ipIdxRangeName: enZonesNommees; spValue: String);

        Function GetRessourceMesure(): Boolean;

        Function AffectationValeur(spSection, spValeur: String; Var spDonnee: String): Boolean; Overload;
        Function AffectationValeur(spSection, spValeur: String; Var ipDonnee: Integer): Boolean; Overload;
        Function AffectationValeur(spSection, spValeur: String; Var ipDonnee: SmallInt): Boolean; Overload;
        Function AffectationValeur(spSection, spValeur: String; Var bpDonnee: Boolean): Boolean; Overload;
        Function AffectationValeur(spSection: String; ipIndex: enGates; lstpLibGates, lstpCdesGates, lstpIdxValGates: TStringList): Boolean; Overload;

        Procedure SetModeMetrologo(epMode: emModesMetrologo; bTesteIEEE: Boolean = True);
        Function ConfigurationSR620(opMesure: TMesure): TModalResult;
        Function Configuration1996(opMesure: TMesure): TModalResult;
        Function Configuration545(opMesure: TMesure): TModalResult;

        Function GestionFicResult(opMes: TMesure): Boolean;
        Function SheetExists(oWb: TExcelWorkbook; spNom: String; bpDispErr: Boolean; Var ipIndex: Integer): Boolean;
        Function TesteEcrasement(oWb: TExcelWorkbook; oWs: TExcelWorksheet; spNom: String): Boolean;

        Function GestionFeuillesNouveauClasseur(epTypeMes: enTypesMesures): Boolean;
        Function AffectationFeuillesTypes(epTypeMesure: enTypesMesures): Boolean;
        Function NouvelleFeuille(opMes: TMesure; opApp: TAppareilIEEE; ipIdxGate: enGates): Boolean;
        Function CreateFeuilleFreqAv_Finale(oWb: TExcelWorkbook; oWs: TExcelWorksheet; bpAvant: Boolean; spNom: String): Boolean;
        Function CreateFeuilleFrequence(oWb: TExcelWorkbook; oWs: TExcelWorksheet): Boolean;
        Procedure DuplicLignesMesure(oWs: TExcelWorksheet; ipNb: Integer);
        Procedure DuplicLignesDeriveRubidiums(oWs: TExcelWorksheet; ipNb: Integer; Var ipColDeb: Integer; Var ipRowDeb: Integer);

        Function ConfigAppareil(opMes: TMesure; opApp: TAppareilIEEE; Var ipErrNum: Integer): Boolean;

        Function Calculs(oWb: TExcelWorkbook; oWs: TExcelWorksheet; oMes: TMesure): Boolean;
        Function Incertitudes(oWb: TExcelWorkbook; oWs: TExcelWorksheet; oMes: TMesure; ipIdxgate: Integer; dpTempsMesure: double): Boolean;

        Procedure MajRecapFreq(oWF, oWR: TExcelWorksheet; oMes: TMesure);
        Procedure MajRecapStab(oWF, oWR: TExcelWorksheet; oMes: TMesure);
        Procedure MajRecapDerive(oWF, oWR: TExcelWorksheet; oMes: TMesure);

        Procedure gestionTimerDerive(Sender: TObject);

        Procedure DispInfosGenes(spMessage: String);

        Procedure AjoutDeriveDansListeEnCours(opMes: TMesure);
        Procedure SupprimeDeriveDeListeEnCours(opMes: TMesure);
        Procedure MajDeriveDansListeEnCours(opMes: TMesure);

        Procedure DelRowGrille(spNumFI: String);
        Procedure AddRowGrille(opMes: TMesure);
        Procedure MajRowGrille(opMes: TMesure);

        Procedure GetInfosDerivesBDR();

        Function MajEcartFreqMoyenneHebdo(ipIDRubidium: Integer; bpAvecGPS, bpEnregAutreRaccord: Boolean; DateJulRef: Integer; dpMarche: Double): Boolean;
        Procedure SetFreqRef(ipIDRubidium: Integer; bpAvecGPS, bpIsActive: Boolean);
        procedure MiseAJourBases(ipDateJul: Integer; dpDate: TDate; dpVal: Double);
        function GetMoyenneHebdo(ipIDRubActif, ipDateJul: Integer; var dpMoyenne: Double): Boolean;

    End;

Var
    frmMain: TfrmMain;

Implementation

Uses DateUtils, F_OutilsSQL, F_ModifValBesancon;

{$R *.dfm}

{*-------------------------------------------------------------------------------
  Procedure : TfrmMain.FormCreate
  @Author   : CB
  @Param    : Sender: TObject
  @Result   : None
  DateTime  : 06/05/2009
-------------------------------------------------------------------------------}

Procedure TfrmMain.FormCreate(Sender: TObject);
Begin

    Self.bForceQuit := True;

    Self.Caption := S_APP_CAPTION + ApplicationVersion;
    //Self.WindowState := wsMaximized;

    ListeFrequencemetres := TList.Create;
    ListeDerivesEnCours := TList.Create;

    MuxHP := TMultiplexeur.Create;
    MesureEnCours := TMesure.Create;
    MesureEnCours.onTimeDerive := gestionTimerDerive;

    Self.dIncertResolution := 0.0;
    Self.dIncertAutres := 0.0;
    Self.dFLue := 0.0;

    Self.bXLConnected := False;
    Self.bFlagRelanceMesure := False;

    cnxEgide.Open;
    sPathResult := IncludeTrailingPathDelimiter(GetPath.GetPathFromSQL('Metrologo', 'FicResultats', @cnxEgide)) + '%s';
    sPathDonFI := GetPath.GetPathFromSQL('dBase', 'DonFI', @cnxEgide);

{$IFDEF Debug}
    sPathModeles := 'L:\Projets Office\MetrologoV5\';
{$ELSE}
    sPathModeles := IncludeTrailingPathDelimiter(GetPath.GetPathFromSQL('Metrologo', 'Modele', @cnxEgide));
{$ENDIF}
    cnxEgide.Close;

    sModeleMetrologo := sPathModeles + S_MODELE_METROLOGO;
    sClasseurValidation := sPathModeles + S_CLASSEUR_VALIDATION;
    sClasseurSuiviRubi := sPathModeles + S_CLASSEUR_SUIVIRUBI;

    cnxMetrologo.Open;

    tmrdelai.Enabled := True;

End;

{*-------------------------------------------------------------------------------
  Procedure : TfrmMain.btnQuitClick
  @Author   : CB
  @Param    : Sender: TObject
  @Result   : None
  DateTime  : 06/05/2009
-------------------------------------------------------------------------------}

Procedure TfrmMain.btnQuitClick(Sender: TObject);
Begin

    Self.Close;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.FormCloseQuery
  @Author:   cb le 12/08/2009
  @Param:    Sender: TObject; Var CanClose: Boolean
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.FormCloseQuery(Sender: TObject; Var CanClose: Boolean);
Begin

    If bForceQuit Then
        CanClose := True
    Else
        CanClose := (Application.MessageBox('Quitter METROLOGO ?', 'Confirmation', MB_YESNO + MB_ICONQUESTION + MB_TOPMOST) = IDYES);

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.tmrDelaiTimer
  @Author:   cb le 12/08/2 009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.tmrDelaiTimer(Sender: TObject);
Var
    slIniFile: TFileName;
    ilErrNum, ilHndXL, ilInterval: Integer;
    dlHeureDeclench: TDateTime;

Begin

    tmrDelai.Enabled := False;
    Self.bCalculEnCours := False;
    dtpDate.DateTime := Now;
    dtpDateChange(dtpDate);

{$IFDEF Debug}
    lblDebug.Visible := True;
{$ELSE}
    lblDebug.Visible := False;
{$ENDIF}

    DispInfosGenes('Lancement de l''application');

    DispInfosGenes('Programmation du timer pour récupération des mesures corrigées par l''observatoire de Besançon');
    dlHeureDeclench := IncHour(DateOf(Now), 10);
    //dlHeureDeclench := IncMinute(DateOf(Now),590);
    While Now > dlHeureDeclench Do
        dlHeureDeclench := IncDay(dlHeureDeclench, 1);

    ilInterval := MilliSecondsBetween(Now, dlHeureDeclench);
//{$IFDEF  Debug}
//    tmrBesancon.Interval := 2000;
//{$ELSE}
    tmrBesancon.Interval := ilInterval;
//{$ENDIF}

    tmrBesancon.Enabled := True;
    DispInfosGenes(Format('Date et heure de la prochaine récupération : %s.', [FormatDateTime(S_FORMATDATETIME, dlHeureDeclench)]));

    // Récupère flags pour savoir s'il faut activer des reminders
    Self.bReminderSaisieFRefAllouis := False;
    ReadRegistryValue(HKEY_LOCAL_MACHINE, S_KEY_BDRMETROLOGO, S_KEY_REMINDERMARCHE, Self.bReminderSaisieFRefAllouis);

    Self.bReminderFreqRefAutres := False;
    ReadRegistryValue(HKEY_LOCAL_MACHINE, S_KEY_BDRMETROLOGO, S_KEY_REMINDERFREQAUTRES, Self.bReminderFreqRefAutres);

    // Active timer reminder si message à rappeller
    tmrReminder.Enabled := (Self.bReminderSaisieFRefAllouis Or Self.bReminderFreqRefAutres);

    // Récup des infos de dérive depuis la BdR
    iNbDeriveEnCours := 0;
    GetInfosDerivesBDR;

    // Get Mutex
    cnxEgide.Open;
    If Not mtxCycleMesure.InitMutex(S_MODULENAME, S_DENOMNAME, cnxEgide, 2000) Then Begin
        cnxEgide.Close;
        Application.MessageBox('Erreur à la création du mutex !' + CRLF + 'Abandon', 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
        Self.Close;
    End;
    cnxEgide.Close;

    // Initialisation IEEE
    AffichePatience('Initialisation du bus IEEE ...');
    Self.bIEEEOk := InitGPIB(I_DEF_IEEETIMEOUT, ilErrNum);
    If Self.bIEEEOk Then
        SetModeMetrologo(emExploitation, False)
    Else Begin
        Application.MessageBox(PChar(Format('Erreur n° %d lors de l''''initialisation du bus IEEE !' + CRLF
            + 'Le logiciel fonctionnera en mode ''''Simulation''''!', [ilErrNum])),
            'Attention', MB_OK + MB_ICONWARNING + MB_TOPMOST);

        SetModeMetrologo(emSimulation);
    End;

    // Recup infos rubidium
    ModifTxtPatience('Détermination du rubidium actif ...');
    Self.RubidiumEnCours := EmptyStr;

    GestionRubidiumACtif;

    With qryMetrologo Do Begin
        // Récupère infos pour calcul de l'incertitude accréditée

        SQL.Clear;
        SQL.Add(S_REQINFOS_INCERT);
        Open;
        MesureEnCours.iNbMesAccred := FieldByName('DAT_NBMES').AsInteger;
        MesureEnCours.iTmpsMesAccred := FieldByName('DAT_TMPS_MESURE').AsInteger;
        Close;

    End;

    // Lecture du fichier d'initialisation
    ModifTxtPatience('Lecture du fichier d''initialisation');
    slIniFile := ChangeFileExt(Application.ExeName, '.ini');
    If FileExists(slIniFile) Then
        iniMetrologo.LoadFromFile(slIniFile)
    Else
    Begin
        Application.MessageBox(PChar(Format('Fichier %s introuvable!' + CRLF + 'Abandon!', [slIniFile])),
            'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
        Self.Close;
    End;

    If Not InitAppareils() Then
        Self.Close;

    // Lancement d'EXCEL
    ModifTxtPatience('Lancement d''EXCEL');
    If Not FileExists(S_XLA) Then
    Begin
        Application.MessageBox(PChar(Format('Macro complémentaire %s introuvable!' + CRLF + 'Abandon!', [S_XLA])),
            'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
        Self.Close;
    End;

    // Chargement macro complémentaire
    Application.ProcessMessages;

    Try
        XLApp.Connect;

    Except
        Application.MessageBox('Erreur au lancement d''EXCEL !' + CRLF
            + 'Abandon!', 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
        Self.Close;
    End;

    ilHndXL := FindWindow(PChar('XLMAIN'), Nil);
    DeleteMenu(GetSystemMenu(ilHndXL, False), SC_CLOSE, MF_BYCOMMAND);

    bForceQuit := False;                                                        // A partir de maintenant, le programme demandera confirmation avant de quitter

    //XLApp.DisplayAlerts[LCID] := False;

    Try
        XLApp.Workbooks.Open(S_XLA, EmptyParam, OLEFalse, EmptyParam, EmptyParam, EmptyParam, OLETrue, EmptyParam, EmptyParam,
            EmptyParam, EmptyParam, EmptyParam, OLETrue, EmptyParam, EmptyParam, LCID);

    Except
        Application.MessageBox(PChar(Format('Erreur au chargement de la macro complémentaire %s !', [S_XLA])),
            'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
        Self.Close;

    End;

    bXLConnected := True;

    // Config. des valeurs par défaut
    ModifTxtPatience('Congiguration des valeurs par défaut ...');
    MesureEnCours.Init();

    FermePatience;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.AddMacroComplementaire
  @Author:   cb le 12/08/2009
  @Param:    spXLA: String
  @Result:   Boolean
-------------------------------------------------------------------------------}
Function TfrmMain.AddMacroComplementaire(spXLA: String): Boolean;
Var
    ilCnt: Integer;
    slSubKey: String;

Begin

    Result := False;

    With TRegistry.Create Do
    Begin                                                                       // Ouverture clef de la BdR
        RootKey := HKEY_CURRENT_USER;
        If Not OpenKey(S_KEYMACRO, False) Then Exit;

        Result := True;                                                         // True par défaut
        ilCnt := 0;                                                             // Recherche de la 1e clef OPEN libre
        slSubKey := 'OPEN';
        While ValueExists(slSubKey) Do
        Begin
            If Pos(spXLA, ReadString(slSubKey)) > 0 Then
            Begin
                CloseKey;
                Free;
                Exit;
            End;
            Inc(ilCnt);
            slSubKey := 'OPEN' + IntToStr(ilCnt);
        End;

        WriteString(slSubKey, spXLA);
        CloseKey;
        Free;

    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.InitAppareils
  @Author:   cb le 13/08/2009
  @Param:    None
  @Result:   Boolean
-------------------------------------------------------------------------------}
Function TfrmMain.InitAppareils: Boolean;
Var
    ilCnt: enAppareilsIEEE;
    lstLignes: TStringList;
    slSection: String;
    ilIndex: Integer;
    blOk: Boolean;
    ilCntGate: enGates;

Begin

    Result := False;
    blOk := False;

    For ilcnt := eaStanford To eaEIP Do Begin
        With iniMetrologo Do Begin
            slSection := S_NOMS_APPIEE[ilCnt];
            lstLignes := GetSection(slSection);
            If lstLignes = Nil Then Begin
                Application.MessageBox(PChar(Format('Section %s introuvable dans le fichier %s !' + CRLF
                    + 'Abandon !', [slSection, Fichier])), 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
                Exit;
            End;
        End;

        ilIndex := ListeFrequencemetres.Add(TAppareilIEEE.Create);
        With TAppareilIEEE(ListeFrequencemetres.Items[ilIndex]) Do Begin

            lstCdeGates := TStringList.Create;
            lstLibellesGates := TStringList.Create;
            lstIdxValeursGates := TStringList.Create;

            While True Do Begin
                blOk := False;
                sNom := slSection;
                If Not AffectationValeur(slSection, S_VAL_ADRESSE, iAdresse) Then Break;
                If Not AffectationValeur(slSection, S_VAL_WRITETERM, iWriteTerm) Then Break;
                If Not AffectationValeur(slSection, S_VAL_READTERM, iReadTerm) Then break;
                If Not AffectationValeur(slSection, S_VAL_HEADERREPONSE, iTailleHeaderReponse) Then break;
                If Not AffectationValeur(slSection, S_VAL_CHAINEINIT, sInit) Then break;
                If Not AffectationValeur(slSection, S_VAL_CONFENTREE, sConfEntree) Then break;
                If Not AffectationValeur(slSection, S_VAL_EXEMESURE, sExeMesure) Then break;
                If Not AffectationValeur(slSection, S_VAL_MONOCOUP, sMonocoup) Then break;
                For ilCntGate := I_GATE10MS To I_GATE100S Do
                    If Not AffectationValeur(slSection, ilCntGate, lstLibellesGates, lstCdeGates, lstIdxValeursGates) Then break;
                If Not AffectationValeur(slSection, S_VAL_GESTSRQ, bGereSRQ) Then Break;
                If Not AffectationValeur(slSection, S_VAL_SRQON, sSRQOn) Then Break;
                If Not AffectationValeur(slSection, S_VAL_SRQOFF, sSRQOff) Then Break;

                blOk := True;
                Break;
            End;
        End;

        // Si erreur de lecture d'un paramètre des appareils IEEE, exit
        If Not blOk Then Exit;

    End;

    With iniMetrologo Do
    Begin
        slSection := S_MUX_HP;
        lstLignes := GetSection(slSection);
        If lstLignes = Nil Then
        Begin
            Application.MessageBox(PChar(Format('Section %s introuvable dans le fichier %s !' + CRLF
                + 'Abandon !', [slSection, Fichier])), 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
            Exit;
        End;
    End;
    MuxHP.sNom := slSection;
    If Not AffectationValeur(slSection, S_VAL_ADRESSE, MuxHP.iAdresse) Then exit;
    If Not AffectationValeur(slSection, S_VAL_WRITETERM, MuxHP.iWriteTerm) Then exit;
    If Not AffectationValeur(slSection, S_VAL_READTERM, MuxHP.iReadTerm) Then exit;

    Result := True;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.AffectationValeur
  @Author:   cb le 13/08/2009
  @Param:    spSection, spValeur: String; var ipDonnee: Integer
  @Result:   Boolean
-------------------------------------------------------------------------------}
Function TfrmMain.AffectationValeur(spSection, spValeur: String; Var ipDonnee: Integer): Boolean;
Begin

    Result := TryStrToInt(iniMetrologo.ValueData(spSection, spValeur), ipDonnee);
    If Not Result Then
        Application.MessageBox(PChar(Format('Valeur incorrecte dans le fichier %s !' + CRLF
            + 'Section %s' + CRLF + 'Paramètre %s' + CRLF + CRLF
            + 'Abandon !', [iniMetrologo.Fichier, spsection, spValeur])), 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.AffectationValeur
  @Author:   cb le 13/08/2009
  @Param:    spSection, spValeur: String; var spDonnee: String
  @Result:   Boolean
-------------------------------------------------------------------------------}
Function TfrmMain.AffectationValeur(spSection, spValeur: String; Var spDonnee: String): Boolean;
Begin

    spDonnee := iniMetrologo.ValueData(spSection, spValeur);
    Result := True;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.AffectationValeur
  @Author:   cb le 02/09/2009
  @Param:    spSection, spValeur: String; lstpValGates, lstpLibellesGates: TStringList
  @Result:   Boolean
-------------------------------------------------------------------------------}
Function TfrmMain.AffectationValeur(spSection: String; ipIndex: enGates; lstpLibGates, lstpCdesGates, lstpIdxValGates: TStringList): Boolean;
Var
    slVal, slParam: String;

Begin

    slParam := AT_VAL_GATES[ipIndex].sValGate;
    slVal := iniMetrologo.ValueData(spSection, slParam);
    If Trim(slVal) <> '' Then
    Begin
        lstpCdesGates.Add(slVal);
        lstpLibGates.Add(AT_VAL_GATES[ipIndex].sLibGate);
        lstpIdxValGates.Add(IntToStr(Ord(ipIndex)));
    End;
    Result := True;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.AffectationValeur
  @Author:   cb le 13/08/2009
  @Param:    spSection, spValeur: String; var ipDonnee: SmallInt
  @Result:   Boolean
-------------------------------------------------------------------------------}
Function TfrmMain.AffectationValeur(spSection, spValeur: String; Var ipDonnee: SmallInt): Boolean;
Var
    ilVal: Integer;

Begin

    Result := TryStrToInt(iniMetrologo.ValueData(spSection, spValeur), ilVal);

    If Result Then
        ipDonnee := SmallInt(ilVal)
    Else
        Application.MessageBox(PChar(Format('Valeur incorrecte dans le fichier %s !' + CRLF
            + 'Section %s' + CRLF + 'Paramètre %s' + CRLF + CRLF
            + 'Abandon !', [iniMetrologo.Fichier, spsection, spValeur])), 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.AffectationValeur
  @Author:   cb le 15/09/2009
  @Param:    spSection, spValeur: String; Var bpDonnee: Boolean
  @Result:   Boolean
-------------------------------------------------------------------------------}
Function TfrmMain.AffectationValeur(spSection, spValeur: String; Var bpDonnee: Boolean): Boolean;
Var
    slVal: String;

Begin

    slVal := iniMetrologo.ValueData(spSection, spValeur);
    bpDonnee := (Trim(slVal) = '1');

    Result := True;

End;
{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.btnExecClick
  @Author:   cb le 24/08/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.btnExecClick(Sender: TObject);
Var
    mrResult: TModalResult;
    ilIdxObjetDerive: Integer;
    olMesureDerive, olMesureCourante: TMesure;

Begin

{$IFDEF Debug}
{
    With MesureEnCours Do Begin
        Init();
        // Config pour EIP
//        Frequencemetre := eaEIP;
//        TAppareilIEEE(ListeFrequencemetres.Items[ord(eaEIP)]).sConfEntree := 'B1';

        // Config pour Stanford
        Frequencemetre := eaStanford;
        TAppareilIEEE(ListeFrequencemetres.Items[ord(eaEIP)]).sConfEntree := 'term1,0;tcpl1,1';

        sNumFI := 'D9_00000';
        // Config pour stab.
//        TypeMesure := etStabilite;                                            // Types de mesures par défaut
//        IndexGate := I_IDX_PROCEDURE_100S;
//
        // Config pour fréquence
//        TypeMesure := etFrequence;                                            // Types de mesures par défaut
//        IndexGate := 0;

        // Config pour Dérive
        TypeMesure := etDerive;                                            // Types de mesures par défaut
        IndexGate := 0;
        IndexIntervalDerive := I_INTERVAL_1H;

        ExecMesure(MesureEnCours);
        btnRelance.Enabled := True;
        Exit;
    End;
}
{$ENDIF}

    btnExec.Enabled := False;
    Self.bFlagRelanceMesure := False;

    With TfrmConfiguration.Create(Self) Do
    Begin                                                                       // Configuration de la mesure (Freq, Mode direct, indirect, voie mux, coef. ...)
        oMesure := MesureEnCours;
        mrResult := ShowModal();
        Free;
    End;

    If mrResult = mrCancel Then
    Begin
        btnExec.Enabled := True;
        Exit;
    End;

    If (MesureEnCours.TypeMesure = etFrequence) Then
    Begin                                                                       // Si il y a une mesure de fréquence à effectuer
        With TfrmChoixFreqGene.Create(Self) Do
        Begin                                                                   // Affiche fenetre de choix fréquencemètre/générateur
            bChoixFreq := MesureEnCours.bMesureAvecFreq;
            mrResult := ShowModal();
            If mrResult = mrOk Then                                             // Si sortie par Ok, prend en compte le choix
                MesureEnCours.bMesureAvecFreq := bChoixFreq;
            Free;
        End;
        If mrResult = mrCancel Then
        Begin
            btnExec.Enabled := True;
            Exit;
        End;
    End;

    If MesureEnCours.Frequencemetre <> eaEIP Then Begin                         // Choix de la voie du multiplexeur si Freq <> EIP545
        With TfrmVoieMux.Create(Self) Do Begin
            oMesure := MesureEnCours;
            mrResult := ShowModal();
            Free;
        End;
        If mrResult = mrCancel Then Begin
            btnExec.Enabled := True;
            Exit;
        End;
    End;

    Case MesureEnCours.Frequencemetre Of                                        // Configuration du fréquencemètre sélectionné
        eaStanford: mrResult := ConfigurationSR620(MesureEnCours);
        eaRacal: mrResult := Configuration1996(MesureEnCours);
        eaEIP: mrResult := Configuration545(MesureEnCours);
    Else
        Begin
            Application.MessageBox('Féquencemètre inconnu !', 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
            Exit;
        End;
    End;
    If mrResult = mrCancel Then Begin
        btnExec.Enabled := True;
        Exit;
    End;

    If (MesureEnCours.TypeMesure = etInterval) Then                             // Si mesure en mode intervallomètre, simule Gate à 1s
        MesureEnCours.IndexGate := ord(I_GATE1S)
    Else Begin
        If MesureEnCours.TypeMesure <> etDerive Then Begin
            With TfrmSelectionGate.Create(Self) Do
            Begin                                                               // Selection du temps de Gate
                oMesure := MesureEnCours;
                mrResult := ShowModal();
                Free;
            End;
            If mrResult = mrCancel Then
            Begin
                btnExec.Enabled := True;
                Exit;
            End;
        End;
    End;

    With MesureEnCours Do Begin                                                 // Saisie du nombre de mesures pour les mesures autres que Dérive
        If TypeMesure In [etFreqAvantInterv, etStabilite, etFrequence, etFreqFinale, etInterval] Then
        Begin

            With TfrmGetNbMesures.Create(Self) Do
            Begin
                oMesure := MesureEnCours;
                mrResult := ShowModal();
                Free;
            End;
            If mrResult = mrCancel Then
            Begin
                btnExec.Enabled := True;
                Exit;
            End;
        End;
    End;

    If (MesureEnCours.TypeMesure = etDerive) Then Begin                         // Saisie des paramètres de la dérive
        With TfrmGetInfosDerive.Create(Self) Do
        Begin
            oMesure := MesureEnCours;
            mrResult := ShowModal();
            Free;
        End;
        If mrResult = mrCancel Then
        Begin
            btnExec.Enabled := True;
            Exit;
        End;

        // Ajout de l'objet TMesure dans la liste des dérives en cours
        // et création d'une nouvelle instance de MesureEnCours
        ilIdxObjetDerive := ListeDerivesEnCours.Add(MesureEnCours);
        MesureEnCours := TMesure.Create;
        MesureEnCours.onTimeDerive := gestionTimerDerive;
        MesureEnCours.Init;

        olMesureDerive := TMesure(ListeDerivesEnCours.Items[ilIdxObjetDerive]); // Pointe sur object Mesure
        With olMesureDerive Do Begin
            TimerDerive.Interval := AI_INTERVAL_DERIVE[IndexIntervalDerive] * 1000; // Intervalle du timer en ms
            TimerDerive.Enabled := True;
            DateDebutDerive := Now;
            DateNextCycleDerive := DateDebutDerive;                     // -- est incrémenté à la fin de la mesure -- IncSecond(DateDebutDerive, AI_INTERVAL_DERIVE[IndexIntervalDerive]);
            olMesureDerive.SaveToRegistry();                            // Sauvegarde des infos de dérive dans la base de registre
            AjoutDeriveDansListeEnCours(olMesureDerive);                // Met à jour grille des dérives en cours
        End;

        olMesureCourante := olMesureDerive;

    End
    Else
        olMesureCourante := MesureEnCours;

    DispInfosGenes('Attente obtention ressource Mesure');
    // Exécute la mesure
    If GetRessourceMesure Then Begin
        DispInfosGenes('Obtention ressource Mesure OK ...');
        ExecMesure(olMesureCourante);
        mtxCycleMesure.ReleaseRessource;
    End;

    btnRelance.Enabled := Not (olMesureCourante.TypeMesure = etDerive);

    btnExec.Enabled := True;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.SetModeValidation
  @Author:   cb le 25/08/2009
  @Param:    bpEtat: Boolean
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.SetModeMetrologo(epMode: emModesMetrologo; bTesteIEEE: Boolean = True);
Var
    ilErrNum: Integer;

Begin

    Case epMode Of
        emExploitation:
            Begin
                If bTesteIEEE Then
                Begin                                                           // S'il faut tester le bus IEEE
                    Self.bIEEEOk := InitGPIB(I_DEF_IEEETIMEOUT, ilErrNum);
                    If Not Self.bIEEEOk Then
                        Application.MessageBox(PChar(Format('Erreur n° %d lors de l''''initialisation du bus IEEE !' + CRLF
                            + 'Le logiciel restera dans le mode de fonctionnement courant !', [ilErrNum])),
                            'Attention', MB_OK + MB_ICONWARNING + MB_TOPMOST);

                End;
                If Self.bIEEEOk Then
                    ModeMetrologo := emExploitation
                Else
                    ModeMetrologo := emSimulation;

            End;
        emSimulation: ModeMetrologo := emSimulation;
        emValidation:
            Begin
                With TfrmMdpValidation.Create(Self) Do
                Begin
                    ExpectedPasswd := S_MDPVALIDATION;
                    ShowModal();
                    If PasswordOk Then
                        ModeMetrologo := emValidation;
                    Free;
                End;
            End;
    End;

    btnValidation.Enabled := ModeMetrologo In [emExploitation, emSimulation];

    btnExploitation.Enabled := ModeMetrologo In [emValidation, emSimulation];

    pnlCaptionMode.Text := Format(S_HTML_CAPTION, [AS_CAPTION_MODEMETROLOGO[ModeMetrologo]]);

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.btnValidationClick
  @Author:   cb le 25/08/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.btnValidationClick(Sender: TObject);
Begin

    SetModeMetrologo(emValidation);

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.btnExploitationClick
  @Author:   cb le 25/08/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.btnExploitationClick(Sender: TObject);
Begin

    SetModeMetrologo(emExploitation);

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.Configuration1996
  @Author:   cb le 28/08/2009
  @Param:    opMesure: TMesure
  @Result:   TModalResult
-------------------------------------------------------------------------------}
Function TfrmMain.Configuration1996(opMesure: TMesure): TModalResult;
Begin

    opMesure.bInitManu := (opMesure.TypeMesure = etInterval);

    With TfrmConfigRacal.Create(Self) Do
    Begin
        oMesure := opMesure;
        Result := ShowModal();
        Free;
    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.Configuration545
  @Author:   cb le 28/08/2009
  @Param:    opMesure: TMesure
  @Result:   TModalResult
-------------------------------------------------------------------------------}
Function TfrmMain.Configuration545(opMesure: TMesure): TModalResult;
Begin

    With TfrmConfigEIP.Create(Self) Do
    Begin
        oMesure := opMesure;
        Result := ShowModal();
        Free;
    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.ConfigurationSR620
  @Author:   cb le 28/08/2009
  @Param:    opMesure: TMesure
  @Result:   TModalResult
-------------------------------------------------------------------------------}
Function TfrmMain.ConfigurationSR620(opMesure: TMesure): TModalResult;
Begin

    If (opMesure.TypeMesure in [etInterval, etTachyContact, etTachyOptique, etStroboscope]) Then
    Begin
        opMesure.bInitManu := True;
        Result := mrOk;
        Exit;
    End;

    With TfrmConfigStanford.Create(Self) Do
    Begin
        oMesure := opMesure;
        Result := ShowModal();
        Free;
    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.FormClose
  @Author:   cb le 02/09/2009
  @Param:    Sender: TObject; var Action: TCloseAction
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.FormClose(Sender: TObject; Var Action: TCloseAction);
Var
    ilCnt: Integer;
    slAddInName: String;

Begin

    FermePatience;

    ListeDerivesEnCours.Free;

    MuxHP.Free;
    MesureEnCours.Free;

    If Assigned(ListeFrequencemetres) Then
    Begin
        With ListeFrequencemetres Do
        Begin
            For ilCnt := 0 To Count - 1 Do
            Begin
                TAppareilIEEE(Items[ilCnt]).lstCdeGates.Free;
                TAppareilIEEE(Items[ilCnt]).lstLibellesGates.Free;
                TAppareilIEEE(Items[ilCnt]).lstIdxValeursGates.Free;
                TAppareilIEEE(Items[ilCnt]).Free;
            End;
            Clear;
            Free;
        End;
    End;

    Try
        With XLApp Do
        Begin
            For ilCnt := 1 To AddIns.Count Do
            Begin
                slAddInName := AddIns.Item[ilCnt].Name;
                If Pos(S_XLA, slAddInName) > 0 Then
                    AddIns.Item[ilCnt].Installed := False;
            End;
//            If Workbooks.Count = 0 Then
//                Visible[LCID] := False;
            Quit;
            Disconnect;
        End;
    Except;
    End;

    cnxMetrologo.Close;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.ExecMesure
  @Author:   cb le 07/09/2009
  @Param:    oMes: TMesure
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.ExecMesure(oMes: TMesure);
Var
    ilCntMesure: Integer;
    dlValGateSecondes: Double;
    olFreq: TAppareilIEEE;
    dlMesure: double;
    XLRange: ExcelRange;
    ilOffsetRowXL, ilNoErrIEEE: Integer;
    adlMesures: Array Of Double;
    aslTime: Array Of String;
    ilCntValGate, ilIndexGateDebut, ilIndexGateFin: enGates;
    ilIndexGateCourant: enGates;
    blExit: Boolean;
    sCommande1996:string;
Begin

    XLApp.Visible[LCID] := True;
    btnStop.Enabled := True;

    While True Do Begin

        DispInfosGenes(Format('Lancement mesure FI N° %s', [oMes.sNumFI]));

        If Not GestionFicResult(oMes) Then Break;
        // A partir de ce point, wbMetrologo est connecté sur le fichier de mesure!!!
        // et wsRecap et wsModele sont connectés respectivement à la feuille recap et à la feuille Modele
        wsModele.Visible[LCID] := xlSheetVisible;

        With oMes Do Begin
            Case IndexGate Of
                I_IDX_PROCEDURE_100S:
                    Begin                                                       // Procedure de 10ms à 100s
                        ilIndexGateDebut := I_GATE10MS;
                        ilIndexGateFin := I_GATE100S;
                    End;

                I_IDX_PROCEDURE_10S:
                    Begin                                                       // Procedure de 10ms à 10s
                        ilIndexGateDebut := I_GATE10MS;
                        ilIndexGateFin := I_GATE10S;
                    End;

                I_IDX_PROCEDURE_EIP:
                    Begin                                                       // Prodédure 10ms, 100ms et 1s
                        ilIndexGateDebut := I_GATE10MS;
                        ilIndexGateFin := I_GATE50MS;                           // = 2 ::: Les procedures avec l'EIP545 ne contiennent que 3 valeurs de Gate !!!
                    End;
            Else
                Begin
                    ilIndexGateDebut := enGates(IndexGate);                     // 1 seule mesure de gate
                    ilIndexGateFin := enGates(IndexGate);
                End;
            End;

            olFreq := TAppareilIEEE(ListeFrequencemetres.Items[ord(Frequencemetre)]);

            //With olFreq do
            //EcritureIEEE(sInit, iAdresse, iWriteTerm);
            EcritureIEEE(olFreq.sInit, olFreq.iAdresse, olFreq.iWriteTerm);

            If bInitManu Then Begin
                SetRemoteLocal(olFreq.iAdresse, False);
                Application.MessageBox('Régler le fréquencemètre !', 'Information', MB_OK + MB_ICONINFORMATION + MB_TOPMOST);
                SetRemoteLocal(olFreq.iAdresse, True);
            End;

            If ConfigAppareil(oMes, olFreq, ilNoErrIEEE) Then Begin             // Config voie Mux, init appareil si pas en mode Manu et passe en mode SRQ si géré par Appareil
                Self.bArretDemande := False;

                If ModeMetrologo = emValidation Then Begin                      // On est en mode VALIDATION -> Choix de la zone nommée à lire
                    With TfrmSelValidation.Create(Self) Do Begin
                        oExcel := XLApp;
                        adMes := @adlMesures;
                        blExit := False;
                        If ShowModal() = mrCancel Then
                            blExit := True
                        Else If Length(adlMesures) < 1 Then Begin
                            Application.MessageBox('Le tableau des mesures de validation est vide !', 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
                            blExit := True;
                        End;
                        If blExit Then Begin
                            Free;
                            wbMetrologo.Close(OLEFalse);
                            wbMetrologo.Disconnect;
                            XLApp.Visible[LCID] := False;
                            btnStop.Enabled := False;
                            Exit;
                        End;

                    End;

                    NbMesures := Length(adlMesures);

                End;

                For ilCntValGate := ilIndexGateDebut To ilIndexGateFin Do Begin
                    // Détermine index de la Gate courante dans tableau AT_VAL_GATES
                    ilIndexGateCourant := enGates(StrToInt(olFreq.lstIdxValeursGates.Strings[Ord(ilCntValGate)]));

                    If Not NouvelleFeuille(oMes, olFreq, ilIndexGateCourant) Then Break;
                    dlValGateSecondes := AT_VAL_GATES[ilIndexGateCourant].dValSecondes;

                    // A partir de ce point, wsMesure est connecté à la feuille de mesure courante (Zones nommées et nb lignes mis à jour)

                    XLRange := wbMetrologo.Names.Item(wsMesure.Name + '!' + S_ZNMESURE1, EmptyParam, EmptyParam).RefersToRange; // Récup n° 1e ligne de mesure dans XL
                    ilOffsetRowXL := XLRange.Row;

                    If TypeMesure <> etInterval Then
                    EcritureIEEE(olFreq.lstCdeGates.Strings[Ord(ilCntValGate)], olFreq.iAdresse, olFreq.iWriteTerm); // Programmation de la gate de l'appareil

//                    XLApp.Calculation[LCID] := xlManual;
                    Application.ProcessMessages;

                    If ModeMetrologo = emvalidation Then Begin                  // On est en mode VALIDATION -> lecture des résultats dans classeur ValidationMetrologo.xls
                        For ilCntMesure := Low(adlMesures) To High(adlMesures) Do Begin
                            wsMesure.Cells.Item[ilOffsetRowXL + ilCntMesure, I_COLXL_HEURE].value := FormatDateTime(S_FORMAT_TIMESHEET, Now);
                            wsMesure.Cells.Item[ilOffsetRowXL + ilCntMesure, I_COLXL_MESURE].value := adlMesures[ilCntMesure];
                        End;
                    End
                    Else Begin
                        If dlValGateSecondes < 1.0 Then Begin                   // Boucle de mesures avec porte < 1s : on bufferise les mesures et
                            SetLength(adlMesures, NbMesures);                   // on remplit la feuille XL après
                            SetLength(aslTime, NbMesures);
                            For ilCntMesure := 0 To NbMesures - 1 Do Begin
                                Application.ProcessMessages;
                                If Self.bArretDemande Then Break;
//                                adlMesures[ilCntMesure] := olFreq.Lecture(olFreq.sExeMesure);
                                adlMesures[ilCntMesure] := olFreq.Lecture(oMes);
                                aslTime[ilCntMesure] := FormatDateTime(S_FORMAT_TIMESHEET, Now);
                            End;

                            If Not Self.bArretDemande Then Begin
                                For ilCntMesure := 0 To NbMesures - 1 Do Begin
                                    wsMesure.Cells.Item[ilOffsetRowXL + ilCntMesure, I_COLXL_HEURE].value := aslTime[ilCntMesure];
                                    wsMesure.Cells.Item[ilOffsetRowXL + ilCntMesure, I_COLXL_MESURE].value := adlMesures[ilCntMesure];
                                End;
                            End;
                        End

                        Else Begin                                              // On remplit la feuille XL au fur et à mesure
                            For ilCntMesure := 0 To NbMesures - 1 Do Begin
                                Application.ProcessMessages;
                                If Self.bArretDemande Then Break;
//                                dlMesure := olFreq.Lecture(olFreq.sExeMesure);
                                dlMesure := olFreq.Lecture(oMes);
                                wsMesure.Cells.Item[ilOffsetRowXL + ilCntMesure, I_COLXL_HEURE].value := FormatDateTime(S_FORMAT_TIMESHEET, Now);
                                wsMesure.Cells.Item[ilOffsetRowXL + ilCntMesure, I_COLXL_MESURE].value := dlMesure;
                            End;

                        End;
                    End;
//                    XLApp.Calculation[LCID] := xlAutomatic;
                    Application.ProcessMessages;

//////                    If olFreq.bGereSRQ Then                                       // Désactive SRQ si géré par appareil
//////                        EcritureIEEE(olFreq.sSRQOff, olFreq.iAdresse, olFreq.iWriteTerm);

// ************************* CB le 10/11/2011 : Correction bug blocage Racal 1996 entre cycle gate=10ms et début cycle gate=20ms
                    If olFreq.bGereSRQ and (ilCntValGate = ilIndexGateFin) Then         // Désactive SRQ si géré par appareil
                    EcritureIEEE(olFreq.sSRQOff, olFreq.iAdresse, olFreq.iWriteTerm) ;
// ************************* Fin modif

                    If Self.bArretDemande Then Break;

                    Calculs(wbMetrologo, wsMesure, oMes);
                    if not Incertitudes(wbMetrologo, wsMesure, oMes, Ord(ilIndexGateCourant), dlValGateSecondes) then
                            Application.MessageBox('Erreur lors de la détermination des incertitudes.' + CRLF
                            + 'Vérifiez que l''appareil est bien utilisé dans une plage de fréquence correcte !' + CRLF
                            + 'Les résultats seront générés quand même.',
                           'Attention', MB_OK + MB_ICONWARNING + MB_TOPMOST);

                    // Mise à jour des feuille Récap. et graphe (si dérive)
                    If IsFreqMeasurement Then
                        MajRecapFreq(wsMesure, wsRecap, oMes)
                    Else If (TypeMesure = etStabilite) Then
                        MajRecapStab(wsMesure, wsRecap, oMes)
                    Else If (TypeMesure = etDerive) Then
                        MajRecapDerive(wsMesure, wsRecap, oMes);

                End;
// ************************* CB le 10/11/2011 suite à modif ci-dessus, il faut donc annuler SRQ si ArretDemandé = true vu que ce n'est plaus fait en cours de boucle
                If olFreq.bGereSRQ and Self.bArretDemande Then         // Désactive SRQ si géré par appareil
                EcritureIEEE(olFreq.sSRQOff, olFreq.iAdresse, olFreq.iWriteTerm);
// ************************* Fin modif

                ProtectionClasseurMesure(True);

                If Not Self.bArretDemande Then Begin
                    wsModele.Visible[LCID] := xlSheetHidden;
                    wbMetrologo.Save(LCID);
                End;

            End
            Else
                Application.MessageBox(PChar(Format('Erreur n° %d lors de la configuration des appareils IEEE !', [ilNoErrIEEE])),
                    'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);

            // Gestion de la dérive
            If (TypeMesure = etDerive) Then Begin

                Inc(oMes.NbCyclesDeriveEffectuees);

                If oMes.NbCyclesDeriveEffectuees >= oMes.NbCyclesDerive Then Begin // Dernier cycle de la dérive
                    DispInfosGenes(Format('Fin du cycle de dérive de la FI n° %s.', [oMes.sNumFI]));
                    SupprimeDeriveDeListeEnCours(oMes);
                End
                Else
                    MajDeriveDansListeEnCours(oMes);

            End;

            Break;
        End;

    End;

    btnStop.Enabled := False;

    wsMesure.Disconnect;
    wsModele.Disconnect;
    wsRecap.Disconnect;

    if (oMes.TypeMesure = etDerive) Then
        wbMetrologo.Close(OLEFalse);

    wbMetrologo.Disconnect;

    if (oMes.TypeMesure = etDerive) Then
        XLApp.Visible[LCID] := False
    else begin
        ShowWindow(XLApp.Hwnd, SW_SHOWMINIMIZED);
        Application.ProcessMessages;
        ShowWindow(XLApp.Hwnd, SW_SHOWNORMAL);
        Application.ProcessMessages;
    end;


    //FermeMetrologoIEEE;
End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.GestionFicResult
  @Author:   cb le 08/09/2009
  @Param:    opMesure: TMesure
  @Result:   Boolean
-------------------------------------------------------------------------------}
Function TfrmMain.GestionFicResult(opMes: TMesure): Boolean;
Var
    slNomFic: String;
    olName: OleVariant;
    bExitOk, bSaveXL: Boolean;

Begin

    Result := False;
    bExitOk := False;
    bSaveXL := False;

    With opMes Do Begin
        If IsFreqMeasurement Then Begin                                         // Si l'on est dans le cas d'une mesure de fréquence
            slNomFic := sDossierResultats + '\' + NomFicFreq;
            If FileExists(slNomFic) Then Begin                                  // Si fichier existe déja, l'ouvre
                XLApp.Workbooks.Open(slNomFic, EmptyParam, OLEFalse, EmptyParam, EmptyParam, EmptyParam, OLETrue, EmptyParam, EmptyParam,
                    EmptyParam, EmptyParam, EmptyParam, OLETrue, EmptyParam, EmptyParam, LCID);
                wbMetrologo.ConnectTo(XLApp.ActiveWorkbook);
                wbMetrologo.Activate;
                ProtectionClasseurMesure(False);        // Ote protection pour trvailler dessus
                Application.ProcessMessages;
                If Not AffectationFeuillesTypes(opMes.TypeMesure) Then Begin
                    wbMetrologo.Close(OLEFalse);
                    Exit;
                End;

                bExitOk := True;
            End
            Else Begin                                                          // Création du fichier de mesure
                XLApp.Workbooks.Add(sModeleMetrologo, LCID);
                wbMetrologo.ConnectTo(XLApp.ActiveWorkbook);
                wbMetrologo.Activate;
                XLApp.DisplayAlerts[LCID] := False;
                Application.Processmessages;

                bExitOk := False;
                If GestionFeuillesNouveauClasseur(etFrequence) Then             // Supprime les feuilles du classeur selon type de mesure à effectuer
                    If AffectationFeuillesTypes(opMes.TypeMesure) Then Begin
                        bExitOk := True;
                        bSaveXL := True;
                        SetRangeValue(wsRecap, I_IIDX_ZNNOFICHERECAPFREQ, sNumFI); // Maj N° FI feuille recap
                        SetRangeValue(wsRecap, I_IDX_ZNRECAPF_NBMES, opMes.NbMesures); // Nb de mesures
                        SetRangeValue(wsRecap, I_IDX_ZNDATERECAPFREQ, FormatDateTime(S_FORMAT_DATESHEET, Now)); // Maj Date feuille recap
                    End;

            End;

        End
        Else If (TypeMesure = etStabilite) Then Begin                           // Si mesure de stab
            slNomFic := sDossierResultats + '\' + NomFicStab;
            If FileExists(slNomFic) Then Begin
                If Application.MessageBox(PChar(Format('Le fichier %s existe déja !' + CRLF + 'Ecraser ?', [slNomFic])),
                    'Confirmation', MB_YESNO + MB_ICONQUESTION + MB_TOPMOST) = IDNO Then
                    Exit
                Else
                Begin
                    If Not DeleteFile(slNomFic) Then
                    Begin
                        Application.MessageBox(PChar(Format('Erreur lors de la tentative de suppression du fichier %s !', [])), 'Erreur',
                            MB_OK + MB_ICONSTOP + MB_TOPMOST);
                        Exit;
                    End;
                End;
            End;

            XLApp.Workbooks.Add(sModeleMetrologo, LCID);
            wbMetrologo.ConnectTo(XLApp.ActiveWorkbook);
            wbMetrologo.Activate;
            XLApp.DisplayAlerts[LCID] := False;
            Application.Processmessages;

            bExitOk := False;
            If GestionFeuillesNouveauClasseur(etStabilite) Then Begin                                                               // Supprime les feuilles du classeur selon type de mesure à effectuer
                If AffectationFeuillesTypes(etStabilite) Then Begin
                    bExitOk := True;
                    bSaveXL := True;
                    SetRangeValue(wsRecap, I_IIDX_ZNNOFICHERECAPSTAB, sNumFI);  // Maj N° FI feuille recap
                    SetRangeValue(wsRecap, I_IDX_ZNDATERECAPSTAB, FormatDateTime(S_FORMAT_DATESHEET, Now)); // Maj Date feuille recap
                End;
            End;

        End
        Else If (TypeMesure = etDerive) Then Begin                              // Si nouvelle mesure de dérive
            slNomFic := sDossierResultats + '\' + NomFicDerive;
            If FileExists(slNomFic) Then Begin
                XLApp.Workbooks.Open(slNomFic, EmptyParam, OLEFalse, EmptyParam, EmptyParam, EmptyParam, OLETrue, EmptyParam, EmptyParam,
                    EmptyParam, EmptyParam, EmptyParam, OLETrue, EmptyParam, EmptyParam, LCID);
                wbMetrologo.ConnectTo(XLApp.ActiveWorkbook);
                wbMetrologo.Activate;
                ProtectionClasseurMesure(False);        // Ote protection pour trvailler dessus
                Application.ProcessMessages;
                If Not AffectationFeuillesTypes(opMes.TypeMesure) Then Begin
                    wbMetrologo.Close(OLEFalse);
                    Exit;
                End;
                bExitOk := True;
            End
            else begin
                XLApp.Workbooks.Add(sModeleMetrologo, LCID);
                wbMetrologo.ConnectTo(XLApp.ActiveWorkbook);
                wbMetrologo.Activate;
                XLApp.DisplayAlerts[LCID] := False;
                Application.Processmessages;

                bExitOk := False;
                If GestionFeuillesNouveauClasseur(etDerive) Then Begin              // Supprime les feuilles du classeur selon type de mesure à effectuer
                    If AffectationFeuillesTypes(etDerive) Then
                    Begin
                        bExitOk := True;
                        bSaveXL := True;
                        SetRangeValue(wsRecap, I_IIDX_ZNNOFICHERECAPDERIVE, sNumFI); // Maj N° FI feuille recap
                        SetRangeValue(wsRecap, I_IDX_ZNDATERECAPDERIVE, FormatDateTime(S_FORMAT_DATESHEET, Now)); // Maj Date feuille recap
                        SetRangeValue(wsRecap, I_IDX_ZNRECAPD_TPSMES, AT_VAL_GATES[enGates(opMes.IndexGate)].sLibGate); // Maj temps de mesure
                        SetRangeValue(wsRecap, I_IDX_ZNRECAPD_NBMESCYCLE, opMes.NbMesDerive); // Maj nb mes par cycle
                        SetRangeValue(wsRecap, I_IDX_ZNRECAPD_TPSCYCLE, AI_LIB_INTERVAL_DERIVE[IndexIntervalDerive]); // Maj intervalle dérive texte
                        SetRangeValue(wsRecap, I_IDX_ZNCOEFFDERIVEJOURNALIERE, AI_INTERVAL_DERIVE[IndexIntervalDerive]); // Maj intervalle dérive en secondes
                    End;
                End;
            end;
        End;

        If Not bExitOk Then
            wbMetrologo.Close(OLEFalse)
        Else Begin
            If bSaveXL Then Begin                                               // Le classeur vient d'être créé, on le sauve (avec protection activée)
                olName := slNomFic;
                ProtectionClasseurMesure(True);
                wbMetrologo.SaveAs(olName, EmptyParam, EmptyParam, EmptyParam, OLEFalse,
                    OLEFalse, xlNoChange, EmptyParam, OLETrue, EmptyParam, EmptyParam, EmptyParam, LCID);
            End;
            ProtectionClasseurMesure(False);        // Ote protection pour trvailler dessus
        End;

        Result := bExitOk;

    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.SheetExists
  @Author:   cb le 08/09/2009
  @Param:    oWb: TExcelWorkbook; spNom: String; bpDispErr: Boolean; var ipIndex: Integer
  @Result:   Boolean
-------------------------------------------------------------------------------}
Function TfrmMain.SheetExists(oWb: TExcelWorkbook; spNom: String; bpDispErr: Boolean; Var ipIndex: Integer): Boolean;
Var
    ilCnt: Integer;

Begin

    Result := True;
    ipIndex := -1;
    For ilCnt := 1 To oWb.Sheets.Count Do
        If (oWb.Sheets.Item[ilCnt] As _Worksheet).Name = spNom Then
        Begin
            ipIndex := ilCnt;
            Exit;
        End;

    If bpDispErr Then
        Application.MessageBox(PChar(Format('Feuille %s inexistante.' + CRLF
            + 'Abandon !', [spNom])), 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);

    Result := False;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.TesteEcrasement
  @Author:   cb le 08/09/2009
  @Param:    oWb: TExcelWorkbook; oWs: TExcelWorksheet; spNom: String
  @Result:   Boolean
-------------------------------------------------------------------------------}
Function TfrmMain.TesteEcrasement(oWb: TExcelWorkbook; oWs: TExcelWorksheet; spNom: String): Boolean;
Var
    ilIdx: Integer;

Begin

    Result := False;
    If SheetExists(oWb, spNom, False, ilIdx) Then
    Begin
        If Application.MessageBox(PChar(Format('La feuille %s existe déja !' + CRLF + 'Ecraser ?', [spNom])), 'Confirmation',
            MB_YESNO + MB_ICONQUESTION + MB_TOPMOST) = IDNO Then
            Exit;

        oWs.ConnectTo(oWb.Sheets.Item[ilIdx] As _WorkSheet);
        oWs.Select;
        oWs.Delete(LCID);
        oWs.Disconnect;

    End;
    Result := True;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.CreateFeuilleFreqAv_Finale
  @Author:   cb le 08/09/2009
  @Param:    oWb: TExcelWorkbook; oWs: TExcelWorksheet; spNom: String
  @Result:   Boolean
-------------------------------------------------------------------------------}
Function TfrmMain.CreateFeuilleFreqAv_Finale(oWb: TExcelWorkbook; oWs: TExcelWorksheet; bpAvant: Boolean; spNom: String): Boolean;
Var
    ilIndex: Integer;

Begin

    Result := False;

    If Not TesteEcrasement(oWb, oWs, spNom) Then exit;                              // Sort si refus écrasement

    wsModele.Select;
    If bpAvant Then Begin
        ilIndex := wsRecap.Index[LCID];                                             // Si Freq. avant, copie feuille après Recap
        wsModele.Copy(EmptyParam, oWb.Sheets.Item[ilIndex] As _Worksheet, LCID);
        ilIndex := wsRecap.Index[LCID] + 1;
    End
    Else Begin                                                                       // Sinon, en dernière position
        ilIndex := oWb.Sheets.Count;
        wsModele.Copy(EmptyParam, oWb.Sheets.Item[ilIndex], LCID);
        Inc(ilIndex);
    End;

    oWs.ConnectTo(oWb.Sheets.Item[ilIndex] As _Worksheet);                          // Selectionne nouvelle feuille
    oWs.Select;
    oWs.Name := spNom;

    Result := True;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.CreateFeuilleFrequence
  @Author:   cb le 08/09/2009
  @Param:    oWb: TExcelWorkbook; oWs: TExcelWorksheet; ipIdxModele, ipIdxRecap: Integer; spNom: String
  @Result:   Boolean
-------------------------------------------------------------------------------}
Function TfrmMain.CreateFeuilleFrequence(oWb: TExcelWorkbook; oWs: TExcelWorksheet): Boolean;
Var
    ilIndexLastFreq: Integer;                                                   // Index de la feuille Freq. la + recente
    ilLastFreqNum: Integer;                                                     // Nom de la feuille Freq. la + récente
    ilFreqNum: Integer;                                                         // Nom de la feuille Freq. courante
    ilIdxFreqAvant: Integer;                                                    // Index de la feuille Freq. avant Interv.
    ilCnt: Integer;
    slNomFeuille: String;

Begin

    ilIdxFreqAvant := -1;                                                       // Par défaut, pas de feuille Freq. avant interv.
    ilIndexLastFreq := -1;
    ilLastFreqNum := 0;

    For ilCnt := 1 To oWb.Sheets.Count Do Begin

        Try
            oWs.ConnectTo(oWb.Sheets.Item[ilCnt] As _Worksheet);
        Except
            Continue;                                                           // Si une erreur est générée c'est que la feuille est un graphique
        End;

        slNomFeuille := Trim(oWs.Name);

        If slNomFeuille = S_NOMSHEET_MODELE Then Continue;                      // Ignore feuilles autres que mesures freq
        If slNomFeuille = S_NOMSHEET_RECAP Then Continue;
        If slNomFeuille = S_NOMSHEET_FREQAVANT Then Begin
            ilIdxFreqAvant := ilCnt;
            Continue;
        End;
        If slNomFeuille = S_NOMSHEET_FREQFINALE Then Continue;
        If Not TryStrToInt(slNomFeuille, ilFreqNum) Then Continue;              // Ignore si nom feuille non numerique
        If ilFreqNum > ilLastFreqNum Then Begin                                 // mémorise N° le plus élevé des deux et l'index de la feuille XL
            ilLastFreqNum := ilFreqNum;
            ilIndexLastFreq := ilCnt;
        End;
    End;

    wsModele.Select;

    If ilLastFreqNum = 0 Then
    Begin                                                                       // Pas de feuille Fréquence trouvée
        If ilIdxFreqAvant = -1 Then                                             // Si pas de feuille Freq. avant Interv., place apres Recap
            ilIdxFreqAvant := wsRecap.Index[LCID];

        wsModele.Copy(EmptyParam, oWb.Sheets.Item[ilIdxFreqAvant] As _Worksheet, LCID);
        oWs.ConnectTo(oWb.Sheets.Item[ilIdxFreqAvant + 1] As _Worksheet);       // Selectionne nouvelle feuille
        oWs.Select;
        oWs.Name := '1';
    End
    Else Begin
        wsModele.Copy(oWb.Sheets.Item[ilIndexLastFreq] As _Worksheet, EmptyParam, LCID); // Duplique modèle AVANT mesure de Freq. la + récente
        oWs.ConnectTo(oWb.Sheets.Item[ilIndexLastFreq] As _Worksheet);          // Selectionne feuille créée
        oWs.Select;                                                             // (l'index ilIndexLastFreq correspond après insertion à la feuille créée)
        oWs.Name := IntToStr(ilLastFreqNum + 1);
    End;

    Result := True;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.SetFormulavalue
  @Author:   cb le 21/09/2009
  @Param:    oWs: TExcelWorksheet; ipIdxRangeName: enZonesNommees; spValue: String
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.SetFormulavalue(oWs: TExcelWorksheet; ipIdxRangeName: enZonesNommees; spValue: String);
Var
    XLRange: ExcelRange;
    slRangeName: String;

Begin

    slRangeName := GetRangeName(oWs, ipIdxRangeName);
    XLRange := wbMetrologo.Names.Item(slRangeName, EmptyParam, EmptyParam).RefersToRange;
    XLRange.FormulaR1C1 := spValue;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.SetRangeValue
  @Author:   cb le 09/09/2009
  @Param:    oWs: TExcelWorksheet; spRangeName, spValue: String
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.SetRangeValue(oWs: TExcelWorksheet; ipIdxRangeName: enZonesNommees; spValue: String);
Var
    XLRange: ExcelRange;
    slRangeName: String;

Begin

    slRangeName := GetRangeName(oWs, ipIdxRangeName);
    XLRange := wbMetrologo.Names.Item(slRangeName, EmptyParam, EmptyParam).RefersToRange;
    XLRange.Value2 := spValue;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.SetRangeValue
  @Author:   cb le 09/09/2009
  @Param:    oWs: TExcelWorksheet; spRangeName: string; dpValue: Double
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.SetRangeValue(oWs: TExcelWorksheet; ipIdxRangeName: enZonesNommees; dpValue: Double);
Var
    XLRange: ExcelRange;
    slRangeName: String;

Begin

    slRangeName := GetRangeName(oWs, ipIdxRangeName);
    XLRange := wbMetrologo.Names.Item(slRangeName, EmptyParam, EmptyParam).RefersToRange;
    XLRange.Value2 := dpValue;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.SetRangeValue
  @Author:   cb le 09/09/2009
  @Param:    oWs: TExcelWorksheet; spRangeName: string; ipValue: Integer
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.SetRangeValue(oWs: TExcelWorksheet; ipIdxRangeName: enZonesNommees; ipValue: Integer);
Var
    XLRange: ExcelRange;
    slRangeName: String;

Begin

    slRangeName := GetRangeName(oWs, ipIdxRangeName);
    XLRange := wbMetrologo.Names.Item(slRangeName, EmptyParam, EmptyParam).RefersToRange;
    XLRange.Value2 := ipValue;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.GetRangeValueDouble
  @Author:   cb le 21/09/2009
  @Param:    oWs: TExcelWorksheet; ipIdxRangeName: enZonesNommees
  @Result:   Double
-------------------------------------------------------------------------------}
Function TfrmMain.GetRangeValueDouble(oWs: TExcelWorksheet; ipIdxRangeName: enZonesNommees): Double;
Var
    XLRange: ExcelRange;
    slRangeName: String;

Begin

    slRangeName := GetRangeName(oWs, ipIdxRangeName);
    XLRange := wbMetrologo.Names.Item(slRangeName, EmptyParam, EmptyParam).RefersToRange;
    Result := XLRange.Value2;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.GetRangeValueString
  @Author:   cb le 22/09/2009
  @Param:    oWs: TExcelWorksheet; ipIdxRangeName2: enZonesNommees
  @Result:   string
-------------------------------------------------------------------------------}
Function TfrmMain.GetRangeValueString(oWs: TExcelWorksheet; ipIdxRangeName: enZonesNommees): String;
Var
    XLRange: ExcelRange;
    slRangeName: String;

Begin

    slRangeName := GetRangeName(oWs, ipIdxRangeName);
    XLRange := wbMetrologo.Names.Item(slRangeName, EmptyParam, EmptyParam).RefersToRange;
    Result := Trim(XLRange.Value2);

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.NouvelleFeuille
  @Author:   cb le 09/09/2009
  @Param:    None
  @Result:   Boolean
-------------------------------------------------------------------------------}
Function TfrmMain.NouvelleFeuille(opMes: TMesure; opApp: TAppareilIEEE; ipIdxGate: enGates): Boolean;
Var
    slLibelleMesure: String;

Begin

    Result := False;

    With opMes Do Begin
        If IsFreqMeasurement Then Begin
            slLibelleMesure := AS_LIBEL_TYPEMESURE[TypeMesure];
            If TypeMesure = etFreqAvantInterv Then Begin                                                               // Crée/Remplace feuille Freq. avant Interv (wsMesure est connecté à la feuille)
                If Not CreateFeuilleFreqAv_Finale(wbMetrologo, wsMesure, True, S_NOMSHEET_FREQAVANT) Then
                    Exit;
            End
            Else If TypeMesure in [etFrequence, etInterval, etTachyContact, etTachyOptique, etStroboscope] Then Begin
                If Not CreateFeuilleFrequence(wbMetrologo, wsMesure) Then
                    Exit;
            End
            Else Begin                                                               // Crée/Remplace feuille Freq. Finale (wsMesure est connecté à la feuille)
                If Not CreateFeuilleFreqAv_Finale(wbMetrologo, wsMesure, False, S_NOMSHEET_FREQFINALE) Then
                    Exit;
            End;
            Result := True;
        End
        Else Begin
            Result := CreateFeuilleFrequence(wbMetrologo, wsMesure);
            If (opMes.TypeMesure = etStabilite) Then
                slLibelleMesure := S_LIBEL_STAB
            Else
                slLibelleMesure := S_LIBEL_DERIVE;
        End;

        // Duplication des lignes de mesure
        DuplicLignesMesure(wsMesure, NbMesures);

        SetRangeValue(wsMesure, I_IDX_ZNNOFICHE, sNumFI);                       // Maj N° FI feuille recap
        SetRangeValue(wsMesure, I_IDX_ZNTYPEMESURE, slLibelleMesure);
        SetRangeValue(wsMesure, I_IDX_ZNDATE, FormatDateTime(S_FORMAT_DATESHEET, Now)); // Maj date

        SetRangeValue(wsMesure, I_IDX_ZNFREQUTILISE, Format(S_FORMAT_FREQUTILISE, [S_NOMS_APPIEE[opMes.Frequencemetre]])); // Maj date
        SetRangeValue(wsMesure, I_IDX_ZNRUBIDIUM, Self.RubidiumEnCours);        // Maj nom du rubidium utilisé

        If Not (TypeMesure = etInterval) Then Begin                             // Si mesure intervalle, ne pas indiquer temps de comptage
            SetRangeValue(wsMesure, I_IDX_ZNGATE, S_LIBEL_TPSCOMPTAGE);
            SetRangeValue(wsMesure, I_IDX_ZNLIBGATE, AT_VAL_GATES[ipIdxGate].sLibGate);
            SetRangeValue(wsMesure, I_IDX_ZNVALGATESECONDES, AT_VAL_GATES[ipIdxGate].dValSecondes);
        End;

        // Maj valeurs de nb mesures et temps de mesure accrédités
        SetRangeValue(wsMesure, I_IDX_ZNNBMESACCREDITE, opMes.iNbMesAccred);
        SetRangeValue(wsMesure, I_IDX_ZNTEMPSMESUREACCREDITE, opMes.iTmpsMesAccred);

        // Maj mode de mesure
        SetRangeValue(wsMesure, I_IDX_ZNMODEMESURE, AS_LIBEL_MODESMESURES[ModeMesure]);
        If ModeMesure = emIndirect Then Begin
            SetRangeValue(wsMesure, I_IDX_ZNCOEFFMULT, IndexMultiplicateur);
            SetRangeValue(wsMesure, I_IDX_ZNLIBFNOMINALE, S_LIB_FNOMINALE);
            SetRangeValue(wsMesure, I_IDX_ZNVALFNOMINALE, FNominale);
        End;

    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.AffectationFeuillesTypes
  @Author:   cb le 23/09/2009
  @Param:    epTypeMesure: TTypesMesure
  @Result:   Boolean
-------------------------------------------------------------------------------}
Function TfrmMain.AffectationFeuillesTypes(epTypeMesure: enTypesMesures): Boolean;
Var
    slSheet: String;

Begin

    Try                                                                         // Assignation feuille Recap à TExcelWorkSheet
        slSheet := S_NOMSHEET_RECAP;
        wsRecap.ConnectTo(wbMetrologo.Sheets.Item[S_NOMSHEET_RECAP] As _WorkSheet);
        slSheet := S_NOMSHEET_MODELE;
        wsModele.ConnectTo(wbMetrologo.Sheets.Item[S_NOMSHEET_MODELE] As _WorkSheet);
        slSheet := S_NOMSHEET_GRAPHE;
        If epTypeMesure = etDerive Then
            chartDerive.ConnectTo(wbMetrologo.Sheets.Item[S_NOMSHEET_GRAPHE] As _Chart);

        Result := True;

    Except
        Result := False;
        Application.MessageBox(PChar(Format('Erreur de connexion à la feuille %s du fichier %s!',
            [slSheet, wbMetrologo.Name])), 'Erreur', MB_OK + MB_ICONSTOP);

    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.GetRangeName
  @Author:   cb le 10/09/2009
  @Param:    ipIdx: Integer
  @Result:   String
-------------------------------------------------------------------------------}
Function TfrmMain.GetRangeName(oWs: TExcelWorksheet; ipIdx: enZonesNommees): String;
Begin

    With AT_ZNXL[ipIdx] Do
    Begin
        If bUnique Then
            Result := Nom
        Else
            Result := oWs.Name + '!' + Nom;
    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.DuplicLignesMesure
  @Author:   cb le 11/09/2009
  @Param:    oWs: TExcelWorksheet; ipNb: Integer
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.DuplicLignesMesure(oWs: TExcelWorksheet; ipNb: Integer);
Var
    XLRange, XLRangePointInsert: ExcelRange;
    slRangeName: String;
    ilCnt: Integer;

Begin

    slRangeName := oWs.Name + '!' + S_ZNMESURE2;
    XLRange := wbMetrologo.Names.Item(slRangeName, EmptyParam, EmptyParam).RefersToRange;

    Application.ProcessMessages;
    slRangeName := oWs.Name + '!' + S_ZNPOINTINSERT;
    XLRangePointInsert := wbMetrologo.Names.Item(slRangeName, EmptyParam, EmptyParam).RefersToRange;

    For ilCnt := 3 To ipNb Do
    Begin                                                                       // il existe déja 2 lignes de mesures dans une feuille vierge
        XLRange.Select;
        XLRange.Copy(EmptyParam);
        XLRangePointInsert.Select;
        XLRangePointInsert.Insert(xlShiftDown, EmptyParam);
    End;

    Application.ProcessMessages;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.ConfigAppareil
  @Author:   cb le 11/09/2009
  @Param:    epApp: enAppareilsIEEE
  @Result:   Boolean
-------------------------------------------------------------------------------}
Function TfrmMain.ConfigAppareil(opMes: TMesure; opApp: TAppareilIEEE; Var ipErrNum: Integer): Boolean;
Begin

    Result := True;

{$IFDEF Debug}
    Exit;
{$ENDIF}

    If ModeMetrologo In [emSimulation, emValidation] Then Exit;                 // Sort si Simulation ou validation

    Result := False;
    With opApp Do
    Begin
        SendIFC(GPIB0);                                                         // Envoi IFC sur bus IEEE
        If TraiteIbsta(ipErrNum) Then Exit;

        // Configure voie du Mux.
        EcritureIEEE(AS_CDESMUX[opMes.VoieMux], MuxHP.iAdresse, MuxHP.iWriteTerm);
        If TraiteIbsta(ipErrNum) Then Exit;

        // Init du frequencemetre
        If Not opMes.bInitManu Then
        Begin
//            EcritureIEEE(sInit, iAdresse, iWriteTerm);
            If TraiteIbsta(ipErrNum) Then Exit;
            EcritureIEEE(sConfEntree, iAdresse, iWriteTerm);

            If TraiteIbsta(ipErrNum) Then Exit;
        End;

        If bGereSRQ Then                                                        // Active gestion SRQ si pris en compte par appareil
        EcritureIEEE(sSRQOn, iAdresse, iWriteTerm);
        Result := True;

    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.btnRelanceClick
  @Author:   cb le 16/09/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}

Procedure TfrmMain.btnRelanceClick(Sender: TObject);
Begin

    If Not Assigned(MesureEnCours) Then Exit;

    btnExec.Enabled := False;
    btnRelance.Enabled := False;

    MesureEnCours.sNumFI := MesureEnCours.sNumFI;                               // Pour gestion des fichiers résultats

    If GetRessourceMesure Then
    Begin
        ExecMesure(MesureEnCours);
        mtxCycleMesure.ReleaseRessource;
    End;

    btnRelance.Enabled := True;
    btnExec.Enabled := True;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.GetRessourceMesure
  @Author:   cb le 16/09/2009
  @Param:    None
  @Result:   Boolean
-------------------------------------------------------------------------------}
Function TfrmMain.GetRessourceMesure: Boolean;
Begin

    // Attente obtention mutex sur mesure
    Result := mtxCycleMesure.GetRessource;
    If Not Result Then
    Begin
        frmMutexBusy := TfrmMutexBusy.Create(Self);
        frmMutexBusy.Show;
        Application.ProcessMessages;

        While True Do
        Begin
            If mtxCycleMesure.GetRessource Then
            Begin
                Result := True;
                Break;
            End;

            If frmMutexBusy.bGiveUp Then
            Begin
                Break;
            End;

            Sleep(200);                                                         // Attente 0,2s
            Application.ProcessMessages;
        End;

        frmMutexBusy.Close;
        frmMutexBusy.Free;

    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.btnLoadLastClick
  @Author:   cb le 16/09/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.btnLoadLastClick(Sender: TObject);
Begin

    If Not Assigned(MesureEnCours) Then Exit;
    With dlgOpenLast Do
    Begin
        InitialDir := MesureEnCours.sDossierResultats;
        If Not Execute() Then Exit;

        If Not XLApp.Visible[LCID] Then
            XLApp.Visible[LCID] := True;

        XLApp.Workbooks.Open(FileName, EmptyParam, OLEFalse, EmptyParam, EmptyParam, EmptyParam, OLETrue, EmptyParam, EmptyParam,
            EmptyParam, EmptyParam, EmptyParam, OLETrue, EmptyParam, EmptyParam, LCID);

        XLApp.ActiveWorkbook.Activate(LCID);
        XLApp.ActiveWorkbook.Save(LCID);

        ShowWindow(XLApp.Hwnd, SW_SHOWMINIMIZED);
        Application.ProcessMessages;
        ShowWindow(XLApp.Hwnd, SW_SHOWNORMAL);
        Application.ProcessMessages;

    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.XLAppWorkbookBeforeClose
  @Author:   cb le 16/09/2009
  @Param:    ASender: TObject; const Wb: _Workbook; var Cancel: WordBool
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.XLAppWorkbookBeforeClose(ASender: TObject; Const Wb: _Workbook; Var Cancel: WordBool);
Begin

    If XLApp.Workbooks.Count < 2 Then
        XLApp.Visible[LCID] := False;
    Cancel := False;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.btnInitClick
  @Author:   cb le 16/09/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.btnInitClick(Sender: TObject);
Begin

    Self.MesureEnCours.Init();
    btnRelance.Enabled := False;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.btnStopClick
  @Author:   cb le 16/09/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.btnStopClick(Sender: TObject);
Begin

    If Application.MessageBox('Arrêter la mesure ?', 'Confirmation', MB_YESNO
        + MB_ICONQUESTION + MB_TOPMOST) = IDYES Then
        Self.bArretDemande := True;

    Application.ProcessMessages;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.Calculs
  @Author:   cb le 17/09/2009
  @Param:    oWs: TExcelWorksheet; oMes: TMesure
  @Result:   Boolean
-------------------------------------------------------------------------------}
Function TfrmMain.Calculs(oWb: TExcelWorkbook; oWs: TExcelWorksheet; oMes: TMesure): Boolean;
Var
    XLRange: ExcelRange;
    ilRowFin, ilOffsetRowXL: Integer;
    slRangeName, slAdress: String;

Begin

    // Maj zone nommée nb de mesures et Fréq. de référence
    SetRangeValue(oWs, I_IDX_ZNNBMESURES, oMes.NbMesures);
    SetRangeValue(oWs, I_IDX_ZNFREQREF, oMes.FReference);

    // Récup n° 1e ligne de mesure dans XL
    XLRange := oWb.Names.Item(oWs.Name + '!' + S_ZNMESURE1, EmptyParam, EmptyParam).RefersToRange;
    ilOffsetRowXL := XLRange.Row;

    // Récup n° dernière ligne de mesure dans XL
    XLRange := oWb.Names.Item(oWs.Name + '!' + S_ZNPOINTINSERT, EmptyParam, EmptyParam).RefersToRange;
    if (oMes.TypeMesure = etInterval) and (oMes.NbMesures = 1) then
        ilRowFin := ilOffsetRowXL
    else
        ilRowFin := XLRange.Row - 1;

    // Crée zone nommée utilisée pour le calcul de la moyenne
    XLRange := xlApp.Range[oWs.Cells.item[ilOffsetRowXL, I_COLXL_FREELLE], oWs.Cells.item[ilRowFin, I_COLXL_FREELLE]];
    slRangeName := '_' + oWs.Name + S_ZNMOYENNE;
    slAdress := '=' + XLRange.Address[OLETrue, OLETrue, xlA1, OLEFalse, EmptyParam];
    oWb.Names.Add(slRangeName, slAdress, OLETrue, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
    // Ecrit formule de calcul de la moyenne
    SetFormulaValue(oWs, I_IDX_ZNFREQMOYREEL, Format(S_FORMULE_MOYENNE, [slRangeName]));
    Application.ProcessMessages;

    // Crée zone nommée utilisée pour calcul Somme des carrés de F(i) - F(i+1)
    XLRange := xlApp.Range[oWs.Cells.item[ilOffsetRowXL + 1, I_COLXL_DELTAF], oWs.Cells.item[ilRowFin, I_COLXL_DELTAF]];
    slRangeName := '_' + oWs.Name + S_ZNDELTAF;
    slAdress := '=' + XLRange.Address[OLETrue, OLETrue, xlA1, OLEFalse, EmptyParam];
    oWb.Names.Add(slRangeName, slAdress, OLETrue, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam, EmptyParam);
    // Ecrit formule de calcul de la moyenne
    SetFormulaValue(oWs, I_IDX_ZNVARIANCE, Format(S_FORMULE_VARIANCE, [slRangeName]));
    Application.ProcessMessages;

    Result := True;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.Incertitudes
  @Author:   cb le 23/09/2009
  @Param:    oWb: TExcelWorkbook; oWs: TExcelWorksheet; oMes: TMesure; ipIdxgate: Integer; dpTempsMesure: double
  @Result:   Boolean
-------------------------------------------------------------------------------}
Function TfrmMain.Incertitudes(oWb: TExcelWorkbook; oWs: TExcelWorksheet; oMes: TMesure; ipIdxgate: Integer; dpTempsMesure: double): Boolean;
Var
    ilIDApp, ilPlageID: Integer;
    dlMoyenne, dlCoefA, dlCoefB: Double;

Begin

    Result := False;
    dlCoefA := 0.0;
    dlCoefB := 0.0;

    // Lit valeur F Moyenne réelle et variance
    dlMoyenne := GetRangeValueDouble(oWs, I_IDX_ZNFREQMOYREEL);

    If dlMoyenne = 0.0 Then Exit;

    // Incertitudes
    If oMes.TypeMesure In [etFreqAvantInterv, etDerive, etFrequence, etFreqFinale] Then Begin
        If oMes.ModeMesure = emIndirect Then Begin
            ilIDApp := -1;
            ilPlageID := 5;
        End
        Else Begin
            ilIDApp := Ord(oMes.Frequencemetre);
            If oMes.Frequencemetre In [eaStanford, eaRacal] Then Begin   // Incertitudes pour Stanford et Racal
                If dlMoyenne <= 10000.0 Then
                    ilPlageID := 0
                Else If (dlMoyenne > 10000.0) And (dlMoyenne <= 1000000.0) Then
                    ilPlageID := 1
                Else If (dlMoyenne > 1000000.0) And (dlMoyenne <= 1010000000.0) Then
                    ilPlageID := 2
                Else
                    Exit;
            End
            Else Begin                                                               // Incertitudes pour EIP
                If dlMoyenne < 1000000000.0 Then
                    Exit
                Else If (dlMoyenne >= 1000000000.0) And (dlMoyenne <= 10000000000.0) Then
                    ilPlageID := 3
                Else If (dlMoyenne > 10000000000.0) And (dlMoyenne <= 18100000000.0) Then
                    ilPlageID := 4
                Else
                    Exit;
            End;
        End;

        // Requete incert frequence pour recup CoefA et CoefB
        With TADOQuery.create(Self) Do
        Begin
            Connection := cnxMetrologo;
            SQL.Clear;
            SQL.Add(Format(S_REQINCERT_FREQ, [Self.IDRubidiumEnCours, IfThen(Self.bAvecGPS, 1, 0), ilIDApp, ilPlageID]));
            Open;
            dlCoefA := FieldByName('CoeffA').AsFloat;
            dlCoefB := FieldByName('CoeffB').AsFloat;
            Close;
            Free;
        End;

    End
    Else If oMes.TypeMesure In [etInterval, etTachyContact, etTachyOptique, etStroboscope] Then
    Begin
        With TADOQuery.create(Self) Do
        Begin
            Connection := cnxMetrologo;
            SQL.Clear;
            SQL.Add(Format(S_REQ_DISPINCERT_AUTRES, [ord(oMes.TypeMesure)]));
            Open;
            dlCoefA := FieldByName('CoeffA').AsFloat;
            dlCoefB := FieldByName('CoeffB').AsFloat;
            Close;
            Free;
        End;
    End
    Else If (oMes.TypeMesure = etStabilite) Then
    Begin
        If oMes.ModeMesure = emIndirect Then
            ilIDApp := -1
        Else
        Begin
            If oMes.Frequencemetre = eaEIp Then
                Exit
            Else
                ilIDApp := Ord(oMes.Frequencemetre);
        End;

        dlCoefB := 0.0;
        // Requete incert Stab pour recup CoefB
        With TADOQuery.create(Self) Do
        Begin
            Connection := cnxMetrologo;
            SQL.Clear;
            SQL.Add(Format(S_REQINCERT_STAB, [ilIDApp, ipIdxgate]));
            Open;
            dlCoefA := FieldByName('CoeffA').AsFloat;
            Close;
            Free;
        End;

    End;

    // Calcul de l'incertitude relative
    // Incertitude accréditée := (dlCoefB / dlMoyenne) + dlCoefA  (Calculé par XL à partir des coeffs fournis)
    // dlIncertAccred := (dlCoefB / dlMoyenne) + dlCoefA ;
    SetRangeValue(oWs, I_IDX_ZNCOEFFA, dlCoefA);
    SetRangeValue(oWs, I_IDX_ZNCOEFFB, dlCoefB);

    // Récupère temps de mesure pour la gate considérée
    //    dlCoeffNbr := Sqrt(oMes.NbMesures / oMes.iNbMesAccred);
    //    dlCoeffT := oMes.iTmpsMesAccred / dpTempsMesure;
    // --- Ces valeurs sont renseignées au moment de la création de la feuille de mesure (voir NouvelleFeuille()) ---

    If (oMes.TypeMesure = etFrequence) And (oMes.bMesureAvecFreq) Then Begin
        XLApp.Visible[LCID] := False;
        Application.ProcessMessages;
        With TfrmSaisieValFreq.Create(Self) Do Begin
            dFreqLue := Self.dFLue;
            ShowModal();
            Self.dFLue := dFreqLue;
            Free;
        End;
        With TfrmParamsIncert.Create(Self) Do Begin
            dResol := Self.dIncertResolution;
            dIncertSupp := Self.dIncertAutres;
            ShowModal();
            Self.dIncertResolution := dResol;
            Self.dIncertAutres := dIncertSupp;
            Free;
        End;
        XLApp.Visible[LCID] := True;
        Application.ProcessMessages;
    End
    Else Begin
        Self.dIncertResolution := 0.0;
        Self.dIncertAutres := 0.0;
    End;

    // Maj valeurs pour calcul par EXCEL de l'incert globale
    SetRangeValue(oWs, I_IDX_ZNINCERTRESOL, Self.dIncertResolution);
    SetRangeValue(oWs, I_IDX_ZNINCERTSUP, Self.dIncertAutres);

    Result := True;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.btnParamIncertClick
  @Author:   cb le 22/09/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.btnParamIncertClick(Sender: TObject);
Begin

    With TfrmDispIncert.Create(Self) Do
    Begin
        CnxSQL := cnxMetrologo;
        ShowModal();
        Free;
    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.GestionFeuillesNouveauClasseur
  @Author:   cb le 23/09/2009
  @Param:    epTypeMes: enTypesMesures
  @Result:   Boolean
-------------------------------------------------------------------------------}
Function TfrmMain.GestionFeuillesNouveauClasseur(epTypeMes: enTypesMesures): Boolean;
Var
    olSheet: TExcelWorksheet;

Begin

    olSheet := TExcelWorksheet.Create(Self);

    Case epTypeMes Of

        etDerive:
            Begin
                olSheet.ConnectTo(wbMetrologo.Sheets.Item[S_NOMSHEET_RECAP_FREQ] As _WorkSheet);
                olSheet.Delete(LCID);
                Application.ProcessMessages;

                olSheet.ConnectTo(wbMetrologo.Sheets.Item[S_NOMSHEET_RECAP_STAB] As _WorkSheet);
                olSheet.Delete(LCID);
                Application.ProcessMessages;

                olSheet.ConnectTo(wbMetrologo.Sheets.Item[S_NOMSHEET_RECAP_DERIVE] As _WorkSheet);
                olSheet.Name := S_NOMSHEET_RECAP;
                olSheet.Select;

                Application.ProcessMessages;
            End;

        etStabilite:
            Begin
                olSheet.ConnectTo(wbMetrologo.Sheets.Item[S_NOMSHEET_RECAP_FREQ] As _WorkSheet);
                olSheet.Delete(LCID);
                Application.ProcessMessages;

                olSheet.ConnectTo(wbMetrologo.Sheets.Item[S_NOMSHEET_RECAP_DERIVE] As _WorkSheet);
                olSheet.Delete(LCID);
                Application.ProcessMessages;

                chartDerive.ConnectTo(wbMetrologo.Sheets.Item[S_NOMSHEET_GRAPHE] As _Chart);
                chartDerive.Delete(LCID);
                Application.ProcessMessages;

                olSheet.ConnectTo(wbMetrologo.Sheets.Item[S_NOMSHEET_RECAP_STAB] As _WorkSheet);
                olSheet.Name := S_NOMSHEET_RECAP;
                olSheet.Select;

                Application.ProcessMessages;
            End;
    Else
        Begin
            olSheet.ConnectTo(wbMetrologo.Sheets.Item[S_NOMSHEET_RECAP_STAB] As _WorkSheet);
            olSheet.Delete(LCID);
            Application.ProcessMessages;

            olSheet.ConnectTo(wbMetrologo.Sheets.Item[S_NOMSHEET_RECAP_DERIVE] As _WorkSheet);
            olSheet.Delete(LCID);
            Application.ProcessMessages;

            chartDerive.ConnectTo(wbMetrologo.Sheets.Item[S_NOMSHEET_GRAPHE] As _Chart);
            chartDerive.Delete(LCID);
            Application.ProcessMessages;

            olSheet.ConnectTo(wbMetrologo.Sheets.Item[S_NOMSHEET_RECAP_FREQ] As _WorkSheet);
            olSheet.Name := S_NOMSHEET_RECAP;

            Application.ProcessMessages;
        End;

    End;

    olSheet.Free;

    Result := True;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.MajRecapFreq
  @Author:   cb le 24/09/2009
  @Param:    oWF, oWR: TExcelWorksheet; oMes: TMesure
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.MajRecapFreq(oWF, oWR: TExcelWorksheet; oMes: TMesure);
Var
    XLRange, XLRangePointInsert: ExcelRange;
    slRangeName: String;
    ilRow: Integer;

Begin

    oWR.Select;
    // POinte zone de réf. à copier
    slRangeName := GetRangeName(oWr, I_IDX_ZNRECAPF_LIGNE0);
    XLRange := wbMetrologo.Names.Item(slRangeName, EmptyParam, EmptyParam).RefersToRange;
    // Point zone où effectuer l'insertion
    slRangeName := GetRangeName(oWr, I_IDX_ZNRECAPF_DEBZONE);
    XLRangePointInsert := wbMetrologo.Names.Item(slRangeName, EmptyParam, EmptyParam).RefersToRange;
    // Duplique zone
    XLRange.Select;
    XLRange.Copy(EmptyParam);
    XLRangePointInsert.Select;
    XLRangePointInsert.Insert(xlShiftDown, EmptyParam);

    Application.ProcessMessages;

    ilRow := XLRangePointInsert.Row - 1;                                        // Détermine n° de la ligne qui vient d'être créée

    slRangeName := '=' + QuotedStr(oWF.Name) + '!';

    oWR.Cells.Item[ilRow, I_COLXL_RECAPF_FREQNOM].FormulaR1C1 := slRangeName + AT_ZNXL[I_IDX_ZNFREQMOYREEL].Nom;
    oWR.Cells.Item[ilRow, I_COLXL_RECAPF_TEMPSMES].FormulaR1C1 := slRangeName + AT_ZNXL[I_IDX_ZNLIBGATE].Nom;
    oWR.Cells.Item[ilRow, I_COLXL_RECAPF_FREQCORR].FormulaR1C1 := slRangeName + AT_ZNXL[I_IDX_ZNFREQCORR].Nom;
    oWR.Cells.Item[ilRow, I_COLXL_RECAPF_ECARTTYPE].FormulaR1C1 := slRangeName + AT_ZNXL[I_IDX_ZNECARTTYPE].Nom;
    If oMes.bMesureAvecFreq Then
        oWR.Cells.Item[ilRow, I_COLXL_RECAPF_FREQINDIQUEE].Value := Self.dFLue
    Else
        oWR.Cells.Item[ilRow, I_COLXL_RECAPF_FREQINDIQUEE].Value := '''Géné.';

    oWR.Cells.Item[ilRow, I_COLXL_RECAPF_INCERTRESOL].FormulaR1C1 := slRangeName + AT_ZNXL[I_IDX_ZNINCERTRESOL].Nom;
    oWR.Cells.Item[ilRow, I_COLXL_RECAPF_INCERTSUP].FormulaR1C1 := slRangeName + AT_ZNXL[I_IDX_ZNINCERTSUP].Nom;
    oWR.Cells.Item[ilRow, I_COLXL_RECAPF_INCERTACCRRED].FormulaR1C1 := slRangeName + AT_ZNXL[I_IDX_ZNINCERTACCREDITEE].Nom;
    oWR.Cells.Item[ilRow, I_COLXL_RECAPF_INCERTGLOB].FormulaR1C1 := slRangeName + AT_ZNXL[I_IDX_ZNINCERTGLOBALE].Nom;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.MajRecapStab
  @Author:   cb le 25/09/2009
  @Param:    oWF, oWR: TExcelWorksheet; oMes: TMesure
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.MajRecapStab(oWF, oWR: TExcelWorksheet; oMes: TMesure);
Var
    XLRange, XLRangePointInsert: ExcelRange;
    slRangeName: String;
    ilRow: Integer;

Begin

    oWR.Select;
    // Pointe zone de réf. à copier
    slRangeName := GetRangeName(oWr, I_IDX_ZNRECAPS_LIGNE0);
    XLRange := wbMetrologo.Names.Item(slRangeName, EmptyParam, EmptyParam).RefersToRange;
    // Pointe zone où effectuer l'insertion
    slRangeName := GetRangeName(oWr, I_IDX_ZNRECAPS_DEBZONE);
    XLRangePointInsert := wbMetrologo.Names.Item(slRangeName, EmptyParam, EmptyParam).RefersToRange;
    // Duplique zone
    XLRange.Select;
    XLRange.Copy(EmptyParam);
    XLRangePointInsert.Select;
    XLRangePointInsert.Insert(xlShiftDown, EmptyParam);

    Application.ProcessMessages;

    ilRow := XLRangePointInsert.Row - 1;                                        // Détermine n° de la ligne qui vient d'être créée

    slRangeName := '=' + QuotedStr(oWF.Name) + '!';

    oWR.Cells.Item[ilRow, I_COLXL_RECAPS_TEMPSGATE].FormulaR1C1 := slRangeName + AT_ZNXL[I_IDX_ZNVALGATESECONDES].Nom;
    oWR.Cells.Item[ilRow, I_COLXL_RECAPS_FREQMOY].FormulaR1C1 := slRangeName + AT_ZNXL[I_IDX_ZNFREQMOYREEL].Nom;
    oWR.Cells.Item[ilRow, I_COLXL_RECAPS_ECARTTYPE].FormulaR1C1 := slRangeName + AT_ZNXL[I_IDX_ZNECARTTYPE].Nom;
    oWR.Cells.Item[ilRow, I_COLXL_RECAPS_INCERT].FormulaR1C1 := slRangeName + AT_ZNXL[I_IDX_ZNINCERTECARTTYPE].Nom;
    oWR.Cells.Item[ilRow, I_COLXL_RECAPS_INCERTACCRED].FormulaR1C1 := slRangeName + AT_ZNXL[I_IDX_ZNINCERTACCREDITEE].Nom;
    oWR.Cells.Item[ilRow, I_COLXL_RECAPS_INCERTGLOB].FormulaR1C1 := slRangeName + AT_ZNXL[I_IDX_ZNINCERTGLOBALE].Nom;
    // -- Formule -- oWR.Cells.Item[ilRow, I_COLXL_RECAPS_VALMAX].FormulaR1C1 := slRangeName + AT_ZNXL[I_IDX_ZNINCERTACCREDITEE].Nom;
    // -- Formule -- oWR.Cells.Item[ilRow, I_COLXL_RECAPS_VALMIN].FormulaR1C1 := slRangeName + AT_ZNXL[I_IDX_ZNINCERTGLOBALE].Nom;
    // -- Formule -- oWR.Cells.Item[ilRow, I_COLXL_RECAPS_ETOILE].FormulaR1C1 := slRangeName + AT_ZNXL[I_IDX_ZNINCERTGLOBALE].Nom;

    XLApp.Run('MajGraphStab');

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.MajRecapDerive
  @Author:   cb le 14/10/2009
  @Param:    oWF, oWR: TExcelWorksheet; oMes: TMesure
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.MajRecapDerive(oWF, oWR: TExcelWorksheet; oMes: TMesure);
Var
//    XLRange, XLRangePointInsert: ExcelRange;
    slRangeName: String;
    ilRow: Integer;

Begin

    oWR.Select;
    
     ilRow := XLApp.Run(wbMetrologo.Name + '!IncrementeZNRecapDerive');
    Application.ProcessMessages;

    slRangeName := '=' + QuotedStr(oWF.Name) + '!';

    oWR.Cells.Item[ilRow, I_COLXL_RECAPD_NUMCYCLE].Value := oWF.Name;
    oWR.Cells.Item[ilRow, I_COLXL_RECAPD_FREQMOY].FormulaR1C1 := slRangeName + AT_ZNXL[I_IDX_ZNFREQMOYREEL].Nom;
    oWR.Cells.Item[ilRow, I_COLXL_RECAPD_ECARTTYPE].FormulaR1C1 := slRangeName + AT_ZNXL[I_IDX_ZNECARTTYPE].Nom;
    // -- Formule -- oWR.Cells.Item[ilRow, I_COLXL_RECAPD_FREQLISSEE].FormulaR1C1 := ;
    oWR.Cells.Item[ilRow, I_COLXL_RECAPD_INCERTGLOB].FormulaR1C1 := slRangeName + AT_ZNXL[I_IDX_ZNINCERTGLOBALE].Nom;
    // -- Formule -- oWR.Cells.Item[ilRow, I_COLXL_RECAPD_FMOY_FLISSEE].FormulaR1C1 := ;

    If oMes.NbCyclesDeriveEffectuees > 2 then
        XLApp.Run('MajGraphDerive');

    oWR.Select;
    If oMes.NbCyclesDeriveEffectuees >= (oMes.NbCyclesDerive - 1) Then
        XLApp.Run('MajInfosFinalesDerive', DaysBetween(oMes.DateDebutDerive, Now), MinutesBetween(oMes.DateDebutDerive, Now));

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.gestionTimerDerive
  @Author:   cb le 30/09/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.gestionTimerDerive(Sender: TObject);
Var
    olMes: TMesure;

Begin

    olMes := TMesure(Sender);
    With olMes Do Begin
        // Reprogrammation systématique de l'intervalle du timer car celui-ci a pu être modifié si Metrologo a été relancé
        // (pour pouvoir retomber sur une date compatible (multiple) pour le cycle défini)
        TimerDerive.Enabled := False;
        TimerDerive.Interval := AI_INTERVAL_DERIVE[IndexIntervalDerive] * 1000;
        TimerDerive.Enabled := True;
        olMes.FReference := dFrequenceDeReference;                              // Si FRef a changé entre 2 cycles de dérive

        DispInfosGenes(Format('Attente obtention ressource Mesure pour relance dérive (N° FI %s).', [olMes.sNumFI]));
        // Exécute la mesure
        If GetRessourceMesure Then Begin
            DispInfosGenes('Obtention ressource Mesure OK ...');
            ExecMesure(olMes);
            mtxCycleMesure.ReleaseRessource;
        End;

    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.DispInfosGenes
  @Author:   cb le 30/09/2009
  @Param:    spMessage: String
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.DispInfosGenes(spMessage: String);
Begin

    mmoInfosGenes.Lines.Add(FormatDateTime(S_FORMATDATETIME, Now) + '   ' + spMessage);

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.DelRowGrille
  @Author:   cb le 30/09/2009
  @Param:    spNumFI: String
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.DelRowGrille(spNumFI: String);
Var
    ilIdx: Integer;

Begin

    With grdDerivesEnCours Do Begin
        ilIdx := Cols[0].IndexOf(spNumFI);
        If (ilIdx > 1) Or (Self.iNbDeriveEnCours > 1) Then
            RemoveRows(ilIdx, 1)
        Else
            Rows[ilIdx].Clear;
    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.MajRowGrille
  @Author:   cb le 01/10/2009
  @Param:    opMes: TMesure
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.MajRowGrille(opMes: TMesure);
Var
    ilIdx: Integer;

Begin

    With grdDerivesEnCours Do Begin
        ilIdx := Cols[0].IndexOf(opMes.sNumFI);
        If ilIdx <> -1 Then Begin
            Cells[2, RowCount - 1] := FormatDateTime(S_FORMATDATETIME, opMes.DateNextCycleDerive);
            Cells[4, RowCount - 1] := IntToStr(opMes.NbCyclesDeriveEffectuees);
        End;
    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.AddRowGrille
  @Author:   cb le 01/10/2009
  @Param:    opMes: TMesure
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.AddRowGrille(opMes: TMesure);
Begin

    With grdDerivesEnCours Do Begin
        If Self.iNbDeriveEnCours > 0 Then
            AddRow();
        Cells[0, RowCount - 1] := opMes.sNumFI;
        Cells[1, RowCount - 1] := FormatDateTime(S_FORMATDATETIME, opMes.DateDebutDerive);
        Cells[2, RowCount - 1] := FormatDateTime(S_FORMATDATETIME, opMes.DateNextCycleDerive);
        Cells[3, RowCount - 1] := IntToStr(opMes.NbCyclesDerive);
        Cells[4, RowCount - 1] := IntToStr(opMes.NbCyclesDeriveEffectuees);
        Cells[5, RowCount - 1] := IntToStr(opMes.VoieMux+1);

    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.GetInfosDerivesBDR
  @Author:   cb le 30/09/2009
  @Param:    None
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.GetInfosDerivesBDR;
Var
    lstClef: TStringList;
    ilCnt, ilNb, ilIndex, ilNbCyclesRates: Integer;

Begin

    DispInfosGenes('Lecture des informations sur les éventuelles dérives en cours ...');

    lstClef := TStringList.Create;
    If ReadRegistrylstKeys(HKEY_LOCAL_MACHINE, S_KEY_BDRMETROLOGO, lstClef) Then Begin

        ilNb := lstClef.Count;

        If ilNb > 0 Then Begin                                                  // Création et init des objets Mesure
            DispInfosGenes(Format('%d dérives en cours.', [ilNb]));
            For ilCnt := 0 To ilNb - 1 Do Begin
                ilIndex := ListeDerivesEnCours.Add(TMesure.Create());
                With TMesure(ListeDerivesEnCours.Items[ilIndex]) Do Begin
                    onTimeDerive := gestionTimerDerive;
                    Init;
                    sNumFI := lstClef.Strings[ilCnt];
                    If ReadFromRegistry() Then Begin
                        DispInfosGenes(Format('Relecture des infos de la dérive FI n° %s : OK', [sNumFI]));

                        // Calcul prochaine date de mesure de dérive
                        ilNbCyclesRates := 0;
                        While DateNextCycleDerive < incSecond(Now, 10) Do Begin // Si Date prochaine < Maintenant + 10 secondes (Temps d'init complet estimation tres large :o)
                            Inc(ilNbCyclesRates);
                            DateNextCycleDerive := IncSecond(DateNextCycleDerive, AI_INTERVAL_DERIVE[IndexIntervalDerive]);
                        End;

                        If ilNbCyclesRates > 0 Then                             // Indique nb cycles manqués
                            DispInfosGenes(Format('%d intervalles de cycles ont été manqués (FI n° %s)', [ilNbCyclesRates, sNumFI]));

                        // Ajoute dérive dans grille
                        AjoutDeriveDansListeEnCours(TMesure(ListeDerivesEnCours.Items[ilIndex]));

                        // Programme intervalle du timer et l'active
                        TimerDerive.Interval := (SecondsBetween(Now, DateNextCycleDerive)) * 1000; // Interval en ms
                        TimerDerive.Enabled := True;

                    End
                    Else Begin                                                  // Si erreur à la relecture des valeurs dans BdR
                        DispInfosGenes(Format('Erreur de lecture des informations de la Base De Registre : Abandon de la dérive FI n° %s', [sNumFI]));
                        Free;
                        ListeDerivesEnCours.Delete(ilIndex);
                    End;

                End;
            End;
        End;
    End;
    lstClef.Free;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.AjoutDeriveDansListeEnCours
  @Author:   cb le 01/10/2009
  @Param:    None
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.AjoutDeriveDansListeEnCours(opMes: TMesure);
Begin

    DispInfosGenes(Format('Lancement dérive (FI n° %s): Prochaine occurence : %s', // Affiche message de lancement de la dérive
        [opMes.sNumFI, FormatDateTime(S_FORMATDATETIME, opMes.DateNextCycleDerive)]));

    AddRowGrille(opMes);                                                        // Mise à jour de la grille des dérives en cours

    Inc(Self.iNbDeriveEnCours);                                                 // Incrémente nb dérives en cours

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.SupprimeDeriveDeListeEnCours
  @Author:   cb le 01/10/2009
  @Param:    None
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.SupprimeDeriveDeListeEnCours(opMes: TMesure);
Var
    ilIndex: Integer;

Begin

    DispInfosGenes(Format('Suppression de la dérive (FI n° %s).', [opMes.sNumFI]));

    opMes.RemoveFromRegistry();                                                 // Supprime infos dérive de la BdR
    DelRowGrille(opMes.sNumFI);
    opMes.TimerDerive.Enabled := False;                                         // Ihnibe timer
    opMes.Free;                                                                 // Free Objet Mesure
    Dec(Self.iNbDeriveEnCours);

    ilIndex := ListeDerivesEnCours.IndexOf(opMes);                              // Supprime objet Mesure de la TList
    If ilIndex <> -1 Then
        ListeDerivesEnCours.Delete(ilIndex);

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.MajDeriveDansListeEnCours
  @Author:   cb le 01/10/2009
  @Param:    opMes: TMesure
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.MajDeriveDansListeEnCours(opMes: TMesure);
Begin

    DispInfosGenes(Format('Mise à jour infos de la dérive (FI n° %s).', [opMes.sNumFI]));

    With opMes Do Begin
        DateNextCycleDerive := IncSecond(DateNextCycleDerive, AI_INTERVAL_DERIVE[IndexIntervalDerive]);
        opMes.UpdateRegistry();
        MajRowGrille(opMes);
    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.grdDerivesEnCoursRightClickCell
  @Author:   cb le 01/10/2009
  @Param:    Sender: TObject; ARow, ACol: Integer
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.grdDerivesEnCoursRightClickCell(Sender: TObject; ARow, ACol: Integer);
Begin

    If ARow < grdDerivesEnCours.FixedRows Then Exit;

    grdDerivesEnCours.Row := ARow;
    mnuKillDerive.PopupAtCursor();

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.optMenuKillDeriveClick
  @Author:   cb le 01/10/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.optMenuKillDeriveClick(Sender: TObject);
Var
    slFI: String;
    ilCnt: Integer;
    olMesure: TMesure;

Begin

    With grdDerivesEnCours Do Begin
        slFI := Cells[0, Row];
        If Application.MessageBox(PChar(Format('Etes-vous sûr(e) de vouloir supprimer la mesure de dérive de la FI n° %s ?', [slFI])),
            'Confirmation', MB_OKCANCEL + MB_ICONQUESTION + MB_TOPMOST) = IDCANCEL Then Exit;

        // Recherche index de la FI dans la TList contenant les objets de dMesure de dérive en cours
        For ilCnt := 0 To ListeDerivesEnCours.Count - 1 Do Begin
            olMesure := TMesure(ListeDerivesEnCours.Items[ilCnt]);
            If olMesure.sNumFI = slFI Then Begin
                SupprimeDeriveDeListeEnCours(olMesure);
                Exit;
            End;
        End;

        Application.MessageBox('Mesure non trouvée dans la liste !', 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);

    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.btnRubidiumClick
  @Author:   cb le 08/10/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.btnRubidiumClick(Sender: TObject);
Var
    blOk: Boolean;
    ilID: Integer;
    blGPS: Boolean;

Begin

    With TfrmChoixRubidium.Create(Self) Do Begin
        dstRubidiums.Connection := Self.cnxMetrologo;
        If ShowModal() = mrOk Then Begin

            blOk := True;
            ilID := iSelRubIDRubidium;
            blGPS := bSelRubAvecGPS;

            With spSaveGPSDaily Do Begin
                ProcedureName := 'Metrologo_SetRubidiumActif';
                Parameters.Refresh;

                Parameters[1].Value := ilID;
                Parameters[2].Value := blGPS;

                Try
                    ExecProc;                                                   // Execute procedure de maj de la BDD

                Except
                    On E: Exception Do Begin
                        blOk := False;
                        Application.MessageBox(PChar(Format('%s', [E.Message])), 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
                    End;
                End;
            End;

            If blOk Then
                GestionRubidiumACtif();

        End;

        Free();

    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.GestionRubidiumACtif
  @Author:   cb le 08/10/2009
  @Param:    None
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.GestionRubidiumACtif;
Begin

    With qryMetrologo Do Begin

        SQL.Clear;
        SQL.Add(S_REQ_RUBACTIF);
        Try
            Open();
            If RecordCount > 0 Then
            Begin
                dFrequenceDeReference := FieldByName(S_RUB_FMOYENNE).AsFloat;
                DispInfosGenes(Format('Fréquence de référence = %8.6f', [dFrequenceDeReference]));
                MesureEnCours.FReference := dFrequenceDeReference;
                Self.RubidiumEnCours := FieldByName(S_RUB_DESIGNATION).AsString;
                Self.IDRubidiumEnCours := FieldByName(S_RUB_ID).AsInteger;
                Self.bAvecGPS := FieldByName(S_RUB_AVECGPS).AsBoolean;
                If Self.bAvecGPS Then
                    Self.RubidiumEnCours := 'GPS + ' + Self.RubidiumEnCours;

                pnlCaptionRubidium.Text := Format(S_HTML_CAPTION, [Self.RubidiumEnCours]);
                btnSuiviRubidiums.Enabled := Not (Self.bAvecGPS);
            End
            Else
            Begin
                Application.MessageBox('Pas de rubidium actif détecté !', 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
                SetModeMetrologo(emSimulation);
            End;
            Close();

        Except
            Application.MessageBox('Erreur d''accès à la base de données !', 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
            SetModeMetrologo(emSimulation);
        End;

    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.tmrBesanconTimer
  @Author:   cb le 08/10/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.tmrBesanconTimer(Sender: TObject);
Var
    lstBesancon: TStringList;
    slFullFileName, slLigne, slDate, slVal, slCurrPath, slNameSave: String;
    blOk: Boolean;
    ilCnt, ilDateJulienne, ilIdx: Integer;
    dlVal: Double;
    ilPosFirstSpace, ilPosSecondSpace, ilPosPoint: Integer;
    dlCurrDate: TDateTime;
    ilJulianDate, ilDayOfTheWeek: Integer;

Begin

    if bForceLoadBesancon then
        bForceLoadBesancon := False
    else begin
        tmrBesancon.Enabled := False;
        tmrBesancon.Interval := 24 * 60 * 60 * 1000;                                // Nb de millisecondes danas 1 journée (pour retomber sur l'heure programmée au lanct)
        tmrBesancon.Enabled := True;
    end;

    DispInfosGenes('Attente obtention ressource Mesure (Pour éviter toute perturbation de la mesure éventuellement en cours)');
    // Exécute la mesure
    If GetRessourceMesure Then Begin

        DispInfosGenes('Obtention ressource Mesure OK ...');
        blOk := True;
        slCurrPath := ExtractFilePath(Application.ExeName);
        slFullFileName := slCurrPath + S_FIC_BESANCON;

        With ftpE2M Do Begin                                                    // Récupération du fichier sur le site E2M

            If Connected Then Disconnect;

            Connect;
            If Connected Then
            Try
                TransferType := ftBinary;
                Try
                    DispInfosGenes('Récupération du fichier de l''observatoire de Besançon ...');
                    Self.bGetDone := False;
                    Get(S_FIC_BESANCON, slFullFileName, True);
                    Application.ProcessMessages;

                    While Not Self.bGetDone Do
                        Application.ProcessMessages;

                    Try
                        DispInfosGenes('Suppression du fichier du ftp E2M ...');
                        Delete(S_FIC_BESANCON);
                        Application.ProcessMessages;

                    Except
                        DispInfosGenes(Format('Erreur à la suppression du fichier %s !', [S_FIC_BESANCON]));
                    End;

                Except
                    blOk := False;
                    DispInfosGenes(Format('Erreur à la récupération du fichier %s !', [S_FIC_BESANCON]));
                End;

            Finally
                Disconnect;
                 DispInfosGenes('FTP déconnecté !')
            End;

        End;

        Application.ProcessMessages;

        If blOk Then Begin
            If Not FileExists(ExtractFilePath(Application.ExeName) + S_FIC_BESANCON) Then Begin
                DispInfosGenes(Format('Erreur à la lecture du fichier %s récupéré ...!', [slFullFileName]));
                Exit;
            End;

            // Sauvegarde du fichier récupéré avec la date du jour
            ForceDirectories(slCurrPath + S_PATH_SAVEBESANCON);
            slNameSave := slCurrPath + S_PATH_SAVEBESANCON + '\' + FormatDateTime('yyyymmdd_hhnnss', Now) + '.txt';
            cbCopyFile(slFullFileName, slNameSave, True,'Sauvegarde fichier de Besançon ...', False);

            lstBesancon := TStringList.Create();
            With lstBesancon Do Begin
                LoadFromFile(slFullFileName);
                DispInfosGenes('Contenu du fichier : ');
                mmoInfosGenes.Lines.Add(Text);

                For ilCnt := 1 To Count - 1 Do Begin                            // Maj de la BDD avec valeurs du fichier (doublons testés dans procédure stockée)
                    slLigne := Strings[ilCnt];

                    ilPosFirstSpace := Pos(' ', slLigne);                       // Détecte position du 1er espace dans la ligne
                    If ilPosFirstSpace < 1 Then Continue;
                    slDate := Copy(slLigne, 1, ilPosFirstSpace - 1);            // Récup des caractères avant l'espace comme date julienne

                    ilPosPoint := Pos('.', slDate);                             // Supprime partie décimale si nécessaire
                    If ilPosPoint > 1 Then
                        slDate := LeftStr(slDate, ilPosPoint - 1);

                    If Not TryStrToInt(slDate, ilDateJulienne) Then             // et convertit en entier, boucle suivante sur erreur conversion
                        Continue;

                    ilIdx := ilPosFirstSpace + 1;
                    While slLigne[ilIdx] = ' ' Do                               // Recherche prochain caractère <> ' '
                        Inc(ilIdx);

                    ilPosSecondSpace := PosEx(' ', slLigne, ilIdx);             // Recherche position de ' ' après 2e valeur
                    If (ilPosSecondSpace < 1) Then Continue;

                    slVal := Copy(slLigne, ilIdx, ilPosSecondSpace - ilIdx);    // Convertit 2e valeur en double; boucle suivante sur erreur
                    slVal := StringReplace(slVal, '.', ',', []);                // Remplace séparateur décimal '.' par ','
                    If Not TryStrToFloat(slVal, dlVal) Then
                        Continue;

                    { En commentaire le 08/04/2022 par RF}
                    With spSaveGPSDaily Do Begin
                        ProcedureName := 'Metrologo_Add_Mesure_Journaliere_GPS';
                        Parameters.Refresh;
                        Parameters[1].Value := ilDateJulienne;
                        Parameters[2].Value := Self.IDRubidiumEnCours;
                        Parameters[3].Value := dlVal;

                        Try
                            ExecProc;                                           // Execute procedure de maj de la BDD
                        Except
                            On E: Exception Do
                                Application.MessageBox(PChar(Format('%s', [E.Message])), 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
                        End;
                    End;

                End;
            End;

            dlCurrDate := Now;
            ilJulianDate := Trunc(DateTimeToModifiedJulianDate(dlCurrDate));

            // Si l'on est un Mardi, effectue le calcul des moyennes pour les 3 mardis précédents
            // Chaque calcul est basé sur les 7 jours précédant la date du Mardi donné
            ilDayOfTheWeek := DayOfTheWeek(dlCurrDate);
            If ilDayOfTheWeek = DayTuesday Then Begin
                MajEcartFreqMoyenneHebdo(IDRubidiumEnCours, True, False, ilJulianDate - 21, 0.0);
                MajEcartFreqMoyenneHebdo(IDRubidiumEnCours, True, False, ilJulianDate - 14, 0.0);
                MajEcartFreqMoyenneHebdo(IDRubidiumEnCours, True, False, ilJulianDate - 7, 0.0);
            End
            Else Begin                                                          // Appelle proc de calcul avac date du dernier mardi passé
                Case ilDayOfTheWeek Of
                    DayMonday: ilJulianDate := ilJulianDate - 6;                // DayMonday = 1
                    DayTuesday: ;
                Else
                    ilJulianDate := ilJulianDate + DayTuesday - ilDayOfTheWeek;
                End;
            End;

            If MajEcartFreqMoyenneHebdo(IDRubidiumEnCours, True, Self.bAvecGPS, ilJulianDate, 0.0) Then Begin
                If Self.bAvecGPS Then Begin                                     // Si raccord = GPS récup. FRef Hebdo et active reminder saisie des fréquences de référence des rubidiums non actifs
                    SetFreqRef(Self.IDRubidiumEnCours, Self.bAvecGPS, True);    // Maj table TR_METROLOGO_RUBIDIUMS
                    GestionRubidiumActif();
                    Self.bReminderFreqRefAutres := True;
                    WriteRegistryValue(HKEY_LOCAL_MACHINE, S_KEY_BDRMETROLOGO, S_KEY_REMINDERFREQAUTRES, Self.bReminderFreqRefAutres);
                    tmrReminder.Enabled := True;
                End
                Else Begin                                                      //Sinon activer reminder de saisie marche hebdomadaire
                    Self.bReminderSaisieFRefAllouis := True;
                    WriteRegistryValue(HKEY_LOCAL_MACHINE, S_KEY_BDRMETROLOGO, S_KEY_REMINDERMARCHE, Self.bReminderSaisieFRefAllouis);
                    tmrReminder.Enabled := True;
                End;

            End;

        End;

        mtxCycleMesure.ReleaseRessource;
        DispInfosGenes('Ressource Mesure libérée !');
        
    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.mnuClearMemoClick
  @Author:   cb le 08/10/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.mnuClearMemoClick(Sender: TObject);
Begin

    mmoInfosGenes.Clear;
    mmoInfosGenes.Refresh;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.mnuEnregMemoClick
  @Author:   cb le 08/10/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.mnuEnregMemoClick(Sender: TObject);
Var
    slFicName: String;

Begin

    slFicName := ExtractFilePath(Application.ExeName) + FormatDateTime('yyyymmdd_hhnnss', Now) + '.log';
    mmoInfosGenes.Lines.SaveToFile(slFicName);
    Application.MessageBox(PChar(Format('Les informations ont été enregistrées dans le fichier %s.', [slFicName])),
        'Information', MB_OK + MB_ICONINFORMATION + MB_TOPMOST);

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.ftpE2MAfterGet
  @Author:   cb le 09/10/2009
  @Param:    ASender: TObject; VStream: TStream
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.ftpe2mAfterGet(ASender: TObject; VStream: TStream);
Begin

    Self.bGetDone := True;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.MajEcartFreqMoyenneHebdo : Maj écart Freq hebdo si date n'existe pas dans la base
  @Author:   cb le 09/10/2009
  @Param:    DateRef: TDateTime
  @Result:   Boolean
-------------------------------------------------------------------------------}
Function TfrmMain.MajEcartFreqMoyenneHebdo(ipIDRubidium: Integer; bpAvecGPS, bpEnregAutreRaccord: Boolean; DateJulRef: Integer; dpMarche: Double): Boolean;
Var
    blCreated: Boolean;

Begin

    blCreated := False;
    With spSaveGPSDaily Do Begin
        ProcedureName := 'Metrologo_Calcul_Ecart_Hebdo';
        Parameters.Refresh;
        Parameters[1].Value := ipIDRubidium;
        Parameters[2].Value := DateJulRef;
        Parameters[3].Value := bpAvecGPS;
        Parameters[4].Value := bpAvecGPS;
        Parameters[5].Value := dpMarche;
        Parameters[6].Value := 0;
        Try
            ExecProc();
            blCreated := Parameters[6].Value;
        Except
            On E: Exception Do
                Application.MessageBox(PChar(Format('%s', [E.Message])), 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
        End;

        If blCreated Then
            DispInfosGenes(Format('Enregistrement de l''écart de fréquence (ID Rubidium = %d, Raccord. = %s, date julienne = %d)',
                [ipIDRubidium, IfThen(bpAvecGPS, 'GPS', 'Allouis'), DateJulRef]));

    End;

    Result := blCreated;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.tmrReminderTimer
  @Author:   cb le 12/10/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.tmrReminderTimer(Sender: TObject);
Begin

    If Self.bReminderSaisieFRefAllouis Then
        DispInfosGenes('************************ Penser à renseigner la marche hebdomadaire !!!!!');

    If Self.bReminderFreqRefAutres Then
        DispInfosGenes('************************ Penser à saisir les fréquences corrigées des rubidiums non actifs !!!!!');

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.SetFreqRef
  @Author:   cb le 12/10/2009
  @Param:    ipIDRubidium: Integer; bpAvecGPS, bpIsActive: Boolean
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.SetFreqRef(ipIDRubidium: Integer; bpAvecGPS, bpIsActive: Boolean);
Begin

    With spSaveGPSDaily Do Begin
        ProcedureName := 'Metrologo_EnregFRef';
        Parameters.Refresh;
        Parameters[1].Value := ipIDRubidium;
        Parameters[2].Value := bpAvecGPS;
        Parameters[3].Value := bpIsActive;
        Try
            ExecProc;                                                           // Execute procedure de maj de la BDD
        Except
            On E: Exception Do
                Application.MessageBox(PChar(Format('%s', [E.Message])), 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
        End;
    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.dtpDateChange
  @Author:   cb le 12/10/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.dtpDateChange(Sender: TObject);
Var
    ilDateJulienne: Integer;

Begin

    If Self.bCalculEnCours Then Exit;

    ilDateJulienne := Trunc(DateTimeToModifiedJulianDate(dtpDate.DateTime));
    edtDateJulienne.Text := IntToStr(ilDateJulienne);

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.edtDateJulienneClickBtn
  @Author:   cb le 12/10/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.edtDateJulienneClickBtn(Sender: TObject);
Begin

    Self.bCalculEnCours := True;
    dtpDate.DateTime := ModifiedJulianDateToDateTime(StrToFloat(edtDateJulienne.Text));
    Self.bCalculEnCours := False;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.btnSuiviRubidiumsClick
  @Author:   cb le 12/10/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.btnSuiviRubidiumsClick(Sender: TObject);
Var
    omrResult: TModalResult;

Begin

    With TfrmSaisieMarcheHebdo.Create(Self) Do Begin
        lblRubActif.Caption := Self.RubidiumEnCours;
        omrResult := ShowModal();
        If omrResult = mrOk Then Begin                                          // Si sortie par OK
            If MajEcartFreqMoyenneHebdo(IDRubidiumEnCours, False, False, JulianDate, ValeurMarche) Then Begin
                SetFreqRef(Self.IDRubidiumEnCours, Self.bAvecGPS, True);        // Maj table TR_METROLOGO_RUBIDIUMS
                GestionRubidiumActif();
                Self.bReminderFreqRefAutres := True;
                Self.bReminderSaisieFRefAllouis := False;
                WriteRegistryValue(HKEY_LOCAL_MACHINE, S_KEY_BDRMETROLOGO, S_KEY_REMINDERFREQAUTRES, Self.bReminderFreqRefAutres);
                WriteRegistryValue(HKEY_LOCAL_MACHINE, S_KEY_BDRMETROLOGO, S_KEY_REMINDERMARCHE, Self.bReminderSaisieFRefAllouis);
                tmrReminder.Enabled := True;
            End;
        End;
        Free();
    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.btnMajRubiAutreClick
  @Author:   cb le 12/10/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.btnMajRubiAutreClick(Sender: TObject);
Var
    omrResult: TModalResult;
    ilDate: Integer;

Begin

    With TfrmGetFreqAutresRubis.Create(Self) Do Begin
        Cnx := cnxMetrologo;
        omrResult := ShowModal();
        If omrResult = mrOk Then Begin

            ilDate := Trunc(DateTimeToModifiedJulianDate(Now));

            With spSaveGPSDaily Do Begin
                ProcedureName := 'Metrologo_EnregFreqAutres';
                Parameters.Refresh;
                Parameters[1].Value := IDRubi1;
                Parameters[2].Value := ilDate;
                Parameters[3].Value := ValRubi1;
                Try
                    ExecProc;                                                   // Execute procedure de maj de la BDD
                Except
                    On E: Exception Do Begin
                        Application.MessageBox(PChar(Format('%s sur Rubidium d''ID %d', [E.Message, IDRubi1])), 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
                        Exit;
                    End;
                End;

                Application.ProcessMessages;

                Parameters[1].Value := IDRubi2;
                Parameters[2].Value := ilDate;
                Parameters[3].Value := ValRubi2;
                Try
                    ExecProc;                                                   // Execute procedure de maj de la BDD
                Except
                    On E: Exception Do Begin
                        Application.MessageBox(PChar(Format('%s sur Rubidium d''ID %d', [E.Message, IDRubi2])), 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
                        Exit;
                    End;
                End;

            End;

            Self.bReminderFreqRefAutres := False;
            WriteRegistryValue(HKEY_LOCAL_MACHINE, S_KEY_BDRMETROLOGO, S_KEY_REMINDERFREQAUTRES, Self.bReminderFreqRefAutres);
            tmrReminder.Enabled := Self.bReminderSaisieFRefAllouis;

        End;
        Free();

    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.btnGraphesderiveClick
  @Author:   cb le 12/10/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.btnGraphesderiveClick(Sender: TObject);
Var
    blGPS, blOk: Boolean;
    ilID, ilCnt, ilNb, ilOldSize, ilColDeb, ilRowDeb: Integer;
    ilDatedeb, ilDateFin, ilIdxArray, ilIdxListe, ilSemaine, ilAnnee: Integer;
    dlVal: Double;
    dlDate: TDateTime;
    slDate, slNomRubidium: String;
    atGrapheDerive: Array Of TGrapheDerive;
    lstInsertValues: TList;
    lstDatesMesure: TStringList;
    XLRange: ExcelRange;

Begin

    // Récupère ID du rubidium à tracer
    With TfrmChoixRubidium.Create(Self) Do Begin

        dstRubidiums.Connection := Self.cnxMetrologo;
        If ShowModal() = mrCancel Then Begin
            Free();
            Exit;
        End;
        slNomRubidium := NomRubidium;
        ilID := iSelRubIDRubidium;
        blGPS := bSelRubAvecGPS;
        If blGPS Then
            slNomRubidium := slNomRubidium + ' (Raccordement GPS)'
        Else
            slNomRubidium := slNomRubidium + ' (Raccordement ALLOUIS)';
        Free();

    End;

    // Récupère les dates de début et fin désirées pour le traçé du graphe
    With TfrmGetDatesGrapheDeriveRubidium.Create(Self) Do Begin
        If ShowModal() = mrCancel Then Begin
            Free();
            Exit;
        End;
        ilDatedeb := DateDebut;
        ilDateFin := DateFin;
        Free();
    End;

    // Récupère les données
    With TADOQuery.Create(Self) Do Begin
        Connection := cnxMetrologo;
        SQL.Clear;
        SQL.Add(Format(S_REQ_DATES_GRAPHES, [ilID, IfThen(blGPS, 1, 0), ilDatedeb, ilDateFin]));
        blOk := True;
        Try
            Open();
        Except
            blOk := False;
        End;

        If Not blOk Then Begin
            Free();
            Application.MessageBox('Erreur à l''ouverture de la base de données !', 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
            Exit;
        End;

        If RecordCount < 2 Then Begin
            Free();
            Application.MessageBox('Nombre de mesures insuffisant (<2) !', 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
            Exit;
        End;

        SetLength(atGrapheDerive, RecordCount);
        lstDatesMesure := TStringList.Create;
        ilIdxArray := 0;

        // Récupère les valeurs en ne conservant éventuellement que la mesure la plus récente pour une même semaine
        For ilCnt := 0 To RecordCount - 1 Do Begin
            dlDate := ModifiedJulianDateToDateTime(StrToFloat(FieldByName('DAT_ID').AsString)); // Convertit date julienne en float
            dlVal := FieldByName('SUV_ECARTF').AsFloat;
            ilSemaine := WeekOf(dlDate);
            ilAnnee := YearOf(dlDate);
            slDate := IntToStr(ilAnnee) + Format('%.2d', [ilSemaine]);          // Calcule date yyyyWW
            ilIdxListe := lstDatesMesure.IndexOf(slDate);
            If ilIdxListe = -1 Then Begin
                lstDatesMesure.Add(slDate);                                     // Le stocke si n'existe pas déja
                atGrapheDerive[ilIdxArray].iSemaine := ilSemaine;
                atGrapheDerive[ilIdxArray].iAnnee := ilAnnee;
                atGrapheDerive[ilIdxArray].dValeur := dlVal;
                atGrapheDerive[ilIdxArray].bMoyennee := False;
                Inc(ilIdxArray);
            End
            Else                                                                // Remplace la valeur existante
                atGrapheDerive[ilIdxListe].dValeur := dlVal;
            Next;
        End;
        Free();
    End;

    If lstDatesMesure.Count < 2 Then Begin
        lstDatesMesure.Free();
        Finalize(atGrapheDerive);
        Application.MessageBox('Nombre de mesures hebdomadaires uniques insuffisant (<2) !', 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
        Exit;
    End;

    // Readapte la taille du tableau en fonction du nb de semaines uniques dans la TStringList
    setlength(atGrapheDerive, lstDatesMesure.Count);

    // Recherche 'trous' dans les semaines et les comble avec valeurs moyennes de début et fin de la plage manquante (en partant de la fin à cause des insertions probables)
    lstInsertValues := TList.Create;

    For ilCnt := High(atGrapheDerive) Downto (Low(atGrapheDerive) + 1) Do Begin
        With atGrapheDerive[ilCnt] Do Begin
            If iSemaine = 1 Then Begin                                          // si la semaine pointée est la 1e de l'année, on calcule celle qui devrait précéder
                ilAnnee := iAnnee - 1;
                ilSemaine := WeeksInAYear(ilAnnee);                             // donc de l'année mémorisée - 1
            End
            Else Begin
                ilSemaine := iSemaine - 1;
                ilAnnee := iAnnee;
            End;
            If (atGrapheDerive[ilCnt - 1].iAnnee <> ilAnnee) Or (atGrapheDerive[ilCnt - 1].iSemaine <> ilSemaine) Then Begin // Si date <> de celle attendue
                ilIdxListe := lstInsertValues.Add(TInsertGraphDerive.Create()); // Maj liste de ajouts à effectuer
                With TInsertGraphDerive(lstInsertValues.Items[ilIdxListe]) Do Begin
                    IndexDeb := ilCnt - 1;
                    DateDeb := EncodeDateWeek(atGrapheDerive[ilCnt - 1].iAnnee, atGrapheDerive[ilCnt - 1].iSemaine);
                    DateFin := EncodeDateWeek(atGrapheDerive[ilCnt].iAnnee, atGrapheDerive[ilCnt].iSemaine);
                    ValMoy := (atGrapheDerive[ilCnt - 1].dValeur + atGrapheDerive[ilCnt].dValeur) / 2.0
                End;
            End;
        End;
    End;

    // Ajoute les dates manquantes
    For ilIdxListe := 0 To lstInsertValues.Count - 1 Do Begin
        With TInsertGraphDerive(lstInsertValues.Items[ilIdxListe]) Do Begin
            ilNb := WeeksBetween(DateDeb, DateFin) - 1;
            If ilNb < 1 Then Continue;
            ilOldSize := High(atGrapheDerive);
            SetLength(atGrapheDerive, High(atGrapheDerive) + ilNb + 1);         // Redimensionne tableau
            For ilCnt := ilOldSize Downto IndexDeb + 1 Do                       // Déplace vers le haut les enreg > IndexDeb
                atGrapheDerive[ilCnt + ilNb] := atGrapheDerive[ilCnt];

            For ilCnt := 0 To ilNb - 1 Do Begin                                 // Comble les trous en mettant un flag pour indiquer que la valeur est interpolée
                dlDate := IncWeek(DateDeb, ilCnt + 1);
                atGrapheDerive[IndexDeb + ilCnt + 1].iSemaine := WeekOf(dlDate);
                atGrapheDerive[IndexDeb + ilCnt + 1].iAnnee := YearOf(dlDate);
                atGrapheDerive[IndexDeb + ilCnt + 1].dValeur := ValMoy;
                atGrapheDerive[IndexDeb + ilCnt + 1].bMoyennee := True;
            End;

        End;
    End;

    // Libere objects TInsertGraphDerive et TList
    For ilCnt := 0 To lstInsertValues.Count - 1 Do
        TInsertGraphDerive(lstInsertValues.Items[ilCnt]).Free;
    lstInsertValues.Free;

    XLApp.Workbooks.Open(sClasseurSuiviRubi, EmptyParam, OLEFalse, EmptyParam, EmptyParam, EmptyParam, OLETrue, EmptyParam, EmptyParam,
        EmptyParam, EmptyParam, EmptyParam, OLETrue, EmptyParam, EmptyParam, LCID);
    XLApp.Visible[LCID] := True;

    wbSuiviRubi.ConnectTo(XLApp.ActiveWorkbook);
    wbSuiviRubi.Activate;
    Application.ProcessMessages;

    wsSuiviRubi.ConnectTo(wbSuiviRubi.Sheets.Item[S_NOMSHEET_DATAS_SUIVIRUBI] As _WorkSheet);
    ;
    wsSuiviRubi.Select;

    wbSuiviRubi.Application.calculation[LCID] := xlCalculationManual;           // Passe en mode recalcul Manuel

    XLRange := wbSuiviRubi.Names.Item(S_ZNRUBI_NOM, EmptyParam, EmptyParam).RefersToRange;
    Application.ProcessMessages;
    XLRange.Value2 := slNomRubidium;

    XLRange := wbSuiviRubi.Names.Item(S_ZNRUBI_DATE, EmptyParam, EmptyParam).RefersToRange;
    Application.ProcessMessages;
    XLRange.Value2 := FormatDateTime('dd/mm/yyyyy', Now);

    XLRange := wbSuiviRubi.Names.Item(S_ZNRUBI_NUMSEM, EmptyParam, EmptyParam).RefersToRange;
    Application.ProcessMessages;
    XLRange.Value2 := WeekOf(Now);

    XLRange := wbSuiviRubi.Names.Item(S_ZNRUBI_NBSEM, EmptyParam, EmptyParam).RefersToRange;
    Application.ProcessMessages;
    XLRange.Value2 := High(atGrapheDerive) + 1;

    DuplicLignesDeriveRubidiums(wsSuiviRubi, High(atGrapheDerive) + 1, ilColDeb, ilRowDeb);

    // Remplit tableau XL
    For ilCnt := Low(atGrapheDerive) To High(atGrapheDerive) Do Begin
        wsSuiviRubi.Cells.Item[ilRowDeb + ilCnt, ilColDeb].value := atGrapheDerive[ilCnt].iSemaine;
        If atGrapheDerive[ilCnt].bMoyennee Then
            wsSuiviRubi.Cells.Item[ilRowDeb + ilCnt, ilColDeb].interior.colorindex := 3
        Else
            wsSuiviRubi.Cells.Item[ilRowDeb + ilCnt, ilColDeb].interior.colorindex := 2;

        wsSuiviRubi.Cells.Item[ilRowDeb + ilCnt, ilColDeb + 1].value := atGrapheDerive[ilCnt].dValeur * 1E12;
        wsSuiviRubi.Cells.Item[ilRowDeb + ilCnt, ilColDeb + 4].value := ilCnt + 1;
    End;

    wbSuiviRubi.Application.calculation[LCID] := xlCalculationAutomatic;        // Passe en mode recalcul Automatique
    Application.ProcessMessages;

    wsSuiviRubi.Disconnect;
    wbSuiviRubi.Disconnect;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.DuplicLignesDeriveRubidiums
  @Author:   cb le 14/10/2009
  @Param:    oWs: TExcelWorksheet; ipNb: Integer
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMain.DuplicLignesDeriveRubidiums(oWs: TExcelWorksheet; ipNb: Integer; Var ipColDeb: Integer; Var ipRowDeb: Integer);
Var
    XLRange, XLRangePointInsert, XLRangeCellDeb, XLRangeTableau: ExcelRange;
    ilCnt, ilRowdeb, ilRowFin: Integer;

Begin

    XLRange := wbSuiviRubi.Names.Item(S_ZNLIGNE_MODELE, EmptyParam, EmptyParam).RefersToRange;
    Application.ProcessMessages;
    XLRangePointInsert := wbSuiviRubi.Names.Item(S_ZNCELL_DEB_COLLAGE, EmptyParam, EmptyParam).RefersToRange;
    Application.ProcessMessages;
    XLRangeCellDeb := wbSuiviRubi.Names.Item(S_ZNPREMIERE_CELLULE, EmptyParam, EmptyParam).RefersToRange;
    Application.ProcessMessages;

    // Suppression de toutes les lignes éventuelles entre la ligne de référence et la ligne contenant le point d'insertion
    ilRowDeb := XLRange.Row + 1;
    ilRowFin := XLRangePointInsert.Row;
    If ilRowFin <> ilRowdeb Then Begin
        With wsSuiviRubi Do Begin
            Range[Rows.item[ilRowDeb, EmptyParam], Rows.item[ilRowFin - 1, EmptyParam]].Select;
            (XLApp.Selection[LCID] As ExcelRange).Delete(EmptyParam);
        End;
    End;

    For ilCnt := 4 To ipNb Do Begin                                             // il existe déja 3 lignes de mesures dans une feuille vierge
        XLRange.Select;
        XLRange.Copy(EmptyParam);
        XLRangePointInsert.Select;
        XLRangePointInsert.Insert(xlShiftDown, EmptyParam);
    End;
    Application.ProcessMessages;

    XLRangeTableau := wbSuiviRubi.Names.Item(S_ZNTABLEAU, EmptyParam, EmptyParam).RefersToRange;
    XLRangeTableau.ClearContents;

    ipRowDeb := XLRangeCellDeb.Row;
    ipColDeb := XLRangeCellDeb.Column;

End;


{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.pnlCaptionRubidiumDblClick
  @Author:   cb le 13/11/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
//procedure TfrmMain.pnlCaptionRubidiumDblClick(Sender: TObject);
//begin
//
//    if (GetKeyState(VK_SHIFT) and $80) = $80 then begin
//        Self.bForceLoadBesancon := True;
//        tmrBesanconTimer(Self);
//    end;
//
//end;


{*-------------------------------------------------------------------------------
  Procedure: TfrmMain.ProtectionClasseurMesure
  @Author:   cb le 12/05/2011
  @Param:    bpProtect: Boolean
  @Result:   None
-------------------------------------------------------------------------------}
procedure TfrmMain.ProtectionClasseurMesure(bpProtect: Boolean);
begin

    Try
        if bpProtect then
            xlApp.Run('ProtectWorkBookMesure')
        else
            xlApp.Run('UnProtectWorkBookMesure')
    except

    end;

end;

procedure TfrmMain.pnlCaptionRubidiumMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin

    if (Button = mbRight) and ((GetKeyState(VK_SHIFT) and $80) = $80) then
        mnuGestFRef.PopupAtCursor;

end;


procedure TfrmMain.mnuForceLoadFTPClick(Sender: TObject);
begin

    if Application.MessageBox('Cette opération va forcer le chargement du fichier transmis par l''observatoire de Besançon' + CRLF
                + 'et effectuer toutes les opérations qui en découlent, à savoir :' + CRLF
                + '   - Mémorisation des valeurs de ce fichiers pour les jours qui ne sont pas encore renseignés.' + CRLF
                + '   - Si l''on est Mardi, la moyenne hebdomadaire va être recalculée ainsi que la Fréquence de référence.' + CRLF + CRLF
                + 'Etes-vous sûr(e) de vouloir continuer ?', 'Confirmation', MB_YESNO + MB_ICONQUESTION + MB_TOPMOST) = IDNO then
        Exit;

    Self.bForceLoadBesancon := True;
    tmrBesanconTimer(Self);

end;


procedure TfrmMain.mnuModifValueClick(Sender: TObject);
var
    dlVal: Double;
    ilDateJul: Integer;
    dlDate: TDate;
    blOk: Boolean;

begin

    with TfrmModifValBesancon.Create(Self, cnxMetrologo) do begin
        blOk := ShowModal = mrOk;
        if blOk then begin
            dlVal := ValJour;
            ilDateJul := DateJulienne;
            dlDate := DateOf(DateToChange);
        end;
        Free;
    end;

    if blOk then
        MiseAJourBases(ilDateJul, dlDate, dlVal);

end;


procedure TfrmMain.MiseAJourBases(ipDateJul: Integer; dpDate: TDate; dpVal: Double);
var
    olQry: TADOQuery;
    blOk: Boolean;
    ilRubActif: Integer;
    ilDayOfTheWeek: Word;
    slVal, slMessErr: string;
    dlMoyenne: Double;

begin

    olQry := TADOQuery.Create(Self);
    with olQry do begin
        Connection := cnxMetrologo;

        // Vérifie que l'on a un enregistrement à la date donnée, et récupère l'ID du rubidium actif si oui
        SQL.Clear;
        SQL.Add(Format('Select * from T_METROLOGO_DATESRUBIS where DAT_ID=%d', [ipDateJul]));
        Open;
        blOk := olQry.RecordCount > 0;
        if blOk then
            ilRubActif := FieldByName('RUB_ACTIF').AsInteger;
        Close;

        // Erreur si on n'a pas trouvé d'enregistrement à la date donnée
        if not blOk then begin
            Application.MessageBox(PChar(Format('La date julienne indiquée (%d) n''existe pas dans la base de données !', [ipDateJul])),
                    'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
            Exit;
        end;

        // Modification de la valeur
        slVal := StringReplace(Format('%g', [dpVal]), ',', '.', []);
        SQL.Clear;
        SQL.Add(Format('Update T_METROLOGO_DATESRUBIS set DAT_VALEUR = %s where DAT_ID=%d and RUB_ACTIF=%d', [slVal, ipDateJul, ilRubActif]));
        try
            slMessErr := '';
            blOk := ExecSQL = 1;
        except
            On E: Exception do begin
                slMessErr := E.Message;
                blOk := False;
            end;
        end;

        if not blOk then begin
            Application.MessageBox(PChar(Format('Erreur lors de l''écriture de la valeur %s à la date julienne %d !' + CRLF + '%s', [slVal, ipDateJul, slMessErr])),
                    'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
            Exit;
        end;

        // Recalcul de la valeur moyenne hebdo si mardi contenant la moyenne est déja dans la base
        ilDayOfTheWeek := DayOfTheWeek(dpDate);
        If ilDayOfTheWeek <> DayTuesday Then Begin               // Si la date julienne n'est pas un mardi, il faut trouver le mardi suivant contenant la moyenne
            Case ilDayOfTheWeek Of                               // des 7 jours précédents (la date julienne concernée faisant partie de ces 7 jours)
                DayMonday: ipDateJul := ipDateJul + 1;           // DayMonday = 1
                DayTuesday: ;
            Else
                ipDateJul := ipDateJul + DaySunday - ilDayOfTheWeek + DayTuesday;
            End;

            // Il faut vérifier que le mardi trouvé est bien présent dans la base de données
            SQL.Clear;
            SQL.Add(Format('Select * from T_METROLOGO_DATESRUBIS where DAT_ID=%d', [ipDateJul]));
            Open;
            blOk := olQry.RecordCount > 0;
            Close;
            if not blOk then begin
                Application.MessageBox('La moyenne hebdomadaire relative à la date concernée n''a pas encore été enregistrée dans la base de données.' + CRLF
                        + 'Le calcul de le fréquence de référence ne sera donc pas effectué !',
                        'Information', MB_OK + MB_ICONINFORMATION + MB_TOPMOST);
                Exit;
            end;
        End;

        // Calcul de la moyenne et recalcul de la fréquence de référence
        if GetMoyenneHebdo(ilRubActif, ipDateJul, dlMoyenne) then begin
            If Self.bAvecGPS Then Begin                                     // Si raccord = GPS récup. FRef Hebdo et active reminder saisie des fréquences de référence des rubidiums non actifs
                SetFreqRef(Self.IDRubidiumEnCours, Self.bAvecGPS, True);    // Maj table TR_METROLOGO_RUBIDIUMS
                GestionRubidiumActif();
            End;
        end;

        Free;

    end;

end;


function TfrmMain.GetMoyenneHebdo(ipIDRubActif, ipDateJul: Integer; var dpMoyenne: Double): Boolean;
var
    olQry: TADOQuery;
    ilDebut, ilFin, ilNb: Integer;
    slEcart, slDelta: string;

const
    S_REQCNT: string = 'Select Count(*) [Nb] from T_METROLOGO_DATESRUBIS where (RUB_ACTIF=%d) and (DAT_ID between %d and %d)';
    S_REQAVG: string = 'Select AVG(DAT_VALEUR) [Moy] from T_METROLOGO_DATESRUBIS where (RUB_ACTIF=%d) and (DAT_ID between %d and %d)';
    S_REQUPDATE: string = 'Update TJ_METROLOGO_SUIVIRUBI set SUV_ECARTF=%s, SUV_DELTATPS=%s where DAT_ID=%d and RUB_ID=%d';

begin

    Result := False;

    // Calcul dates juliennes des jours de début et fin de la plage de 7 jours précédant le jour passé en param (Mardi)
    ilDebut := ipDateJul - 7;
    ilFin := ipDateJul - 1;

    olQry := TADOQuery.Create(Self);
    with olQry do begin
        Connection := cnxMetrologo;
        SQL.Clear;
        SQL.Add(Format(S_REQCNT, [ipIDRubActif, ilDebut, ilFin]));
        Open;
        ilNb := FieldbyName('Nb').AsInteger;
        Close;
        if ilNb <> 7 then begin
            Application.MessageBox('La base de données ne contient pas la totalité des 7 mesures nécessaires pour calculer la moyenne hebdomadaire pour la date concernée !',
               'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
            Free;
            Exit;
        end;

        SQL.Clear;
        SQL.Add(Format(S_REQAVG, [ipIDRubActif, ilDebut, ilFin]));
        Open;
        dpMoyenne := FieldByName('Moy').AsFloat;
        Close;

        slEcart := StringReplace(Format('%g', [dpMoyenne]), ',', '.', []);
        slDelta := StringReplace(Format('%g', [dpMoyenne * 86400.0]), ',', '.', []);
        SQL.Clear;
        SQL.Add(Format(S_REQUPDATE, [slEcart, slDelta, ipDateJul, ipIDRubActif]));

        Result := ExecSQL > 0;

        Free;

    end;

end;

End.


