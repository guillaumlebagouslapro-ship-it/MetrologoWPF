Unit F_Configuration;

Interface

Uses
    Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
    Dialogs, StdCtrls, AdvEdit, JvExStdCtrls, JvRadioButton, JvCheckBox, Mask,
    JvButton, JvCtrls, JvExMask, JvToolEdit, JvMaskEdit, DateUtils,
    U_DeclarationsMETROLOGO, DB, ADODB,
    //ApoDSet, ApoEnv,
    JvComponentBase, StrUtils, ExtCtrls, ApoEnv, ApoDSet;

Type
    TfrmConfiguration = Class(TForm)
        grpFrequencemetre: TGroupBox;
        optStanford: TJvRadioButton;
        optRacalDana: TJvRadioButton;
        optEIP: TJvRadioButton;
        grpModeMesure: TGroupBox;
        optDirect: TJvRadioButton;
        optIndirect: TJvRadioButton;
        grpIndirect: TGroupBox;
        edtFNominale: TAdvEdit;
        cbxMultiplicateur: TComboBox;
        lblMultiplicateur: TLabel;
        btnOk: TJvImgBtn;
        btnCancel: TJvImgBtn;
        edtNoFI: TJvMaskEdit;
        lblNoFI: TLabel;
        tblInterv: TApolloTable;
        envMetrologo: TApolloEnv;
        rgbMesure: TRadioGroup;
    conASERi: TADOConnection;
        Procedure FormCreate(Sender: TObject);
        Procedure btnOkClick(Sender: TObject);
        Procedure FormCloseQuery(Sender: TObject; Var CanClose: Boolean);
        Procedure optDirectClick(Sender: TObject);
        Procedure optIndirectClick(Sender: TObject);
        Procedure optEIPClick(Sender: TObject);
        Procedure btnCancelClick(Sender: TObject);
        Procedure rgbMesureClick(Sender: TObject);
        Procedure optStanfordClick(Sender: TObject);

    Private
        bConfigEnCours: Boolean;
        bAllowQuit: Boolean;
        FMesure: TMesure;

        Function GetYearE2M(): String;
        Procedure ConfigMesure(Const Value: TMesure);

    Public
        Property oMesure: TMesure Read FMesure Write ConfigMesure;

    End;

Var
    frmConfiguration: TfrmConfiguration;

Implementation

{$R *.dfm}

{*-------------------------------------------------------------------------------
  Procedure : TfrmConfiguration.FormCreate
  @Author   : CB le 09/05/2009
  @Param    : Sender: TObject
  @Result   : None
-------------------------------------------------------------------------------}
Procedure TfrmConfiguration.FormCreate(Sender: TObject);
Begin

    //tblInterv.DatabaseName := sPathDonFI;
    Self.ballowQuit := True;
    Self.bConfigEnCours := False;

End;

{*-------------------------------------------------------------------------------
  Procedure : TfrmConfiguration.btnOkClick
  @Author   : CB le 09/05/2009
  @Param    : Sender: TObject
  @Result   : None
-------------------------------------------------------------------------------}
Procedure TfrmConfiguration.btnOkClick(Sender: TObject);
Var
    slNoFI, slPath, sFI: String;
	bExist: boolean;
{$IFNDEF Debug}
    slRacineFi: String;
{$ENDIF}

Begin

    Self.bAllowQuit := False;

    If rgbMesure.ItemIndex < 0 Then Begin
        Application.MessageBox('Il faut sélectionner au moins une mesure à effectuer !',
            'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
        Exit;
    End;

    // Test saisie du N° de FI
    slNoFI := Trim(edtNoFI.EditText);
    If Length(slNoFI) <> 8 Then Begin
        Application.MessageBox(PChar(Format('N° de FI incorrect : %s !', [slNoFI])),
            'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
        edtNoFI.SetFocus;
        Exit;
    End;

    // Test existence FI dans ASERi
    bExist := True;
    With TADODataSet.Create(nil) do begin
        sFI := StringReplace(slNoFI, '_', '/', []);
        Connection := conASERi;
        CommandText := Format(S_REQ_ASERI_FI, [QuotedStr(sFI)]);
        //Showmessage(commandtext);
        Open();
        if Eof Then begin
            Application.MessageBox(PChar(Format('La FI n° %s n''existe pas dans ASERi !', [sFI])), 'Erreur', MB_OK + MB_ICONSTOP);
            bExist := False;
        end;
        Active := False;
        Free();
    end;
    if not bExist then Exit;

    // Mémorise infos de mesure
    With FMesure Do Begin
        sNumFI := slNoFI;                                                       // Mémorise N° FI

{$IFNDEF Debug}
        slRacineFi := Leftstr(sNumFI, 8);
//        Try                                                                     // Recherche N° de FI dans Interv.dbf
//            tblInterv.Open;
//
//        Except
//            Application.MessageBox(PChar(Format('Erreur lors de l''ouverture de la table %s !', [tblInterv.TableName])), 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST + MB_TOPMOST);
//            Exit;
//        End;

        { If Not tblInterv.Seek(StringReplace(slRacineFi, '_', '/', [])) Then }
        { Begin }
            { tblInterv.Close; }
            { Application.MessageBox(PChar(Format('La FI n° %s n''existe pas dans la table %s !', [slRacineFi, tblInterv.TableName])), 'Erreur', MB_OK + MB_ICONSTOP); }
            { Exit; }
        { End; }

        //tblInterv.Close;
{$ENDIF}

        // Création du dossier résultat s'il n'existe pas
        slPath := Format(sPathResult, [slNoFI]);
        If Not DirectoryExists(slPath) Then
            If Not ForceDirectories(slPath) Then
            Begin
                Application.MessageBox(PChar(Format('Erreur à la création du dossier %s !', [slPath])), 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
                Exit;
            End;

        If optStanford.Checked Then                                             // Sélection fréquencemètre
            Frequencemetre := eaStanford
        Else If optRacalDana.Checked Then
            Frequencemetre := eaRacal
        Else
            Frequencemetre := eaEIP;

        If optDirect.Checked Then                                               // Mode de mesure
            ModeMesure := emDirect
        Else
            ModeMesure := emIndirect;

        IndexMultiplicateur := cbxMultiplicateur.ItemIndex;                     // valeur du coef. multiplicateur
        FNominale := StrToFloat(Trim(edtFNominale.Text));                       // Valeur de la F nominale

        TypeMesure := enTypesMesures(rgbMesure.ItemIndex);
    End;

    Self.bAllowQuit := True;

End;

{*-------------------------------------------------------------------------------
  Procedure : TfrmConfiguration.FormCloseQuery
  @Author   : CB le 09/05/2009
  @Param    : Sender: TObject; var CanClose: Boolean
  @Result   : None
-------------------------------------------------------------------------------}
Procedure TfrmConfiguration.FormCloseQuery(Sender: TObject; Var CanClose: Boolean);
Begin

    CanClose := Self.bAllowQuit;
    Self.bAllowQuit := True;

End;

{*-------------------------------------------------------------------------------
  Procedure : TfrmConfiguration.rb1Click
  @Author   : CB le 09/05/2009
  @Param    : Sender: TObject
  @Result   : None
-------------------------------------------------------------------------------}
Procedure TfrmConfiguration.optDirectClick(Sender: TObject);
Begin

    grpIndirect.Visible := False;

End;

{*-------------------------------------------------------------------------------
  Procedure : TfrmConfiguration.rbIndirectClick
  @Author   : CB le 09/05/2009
  @Param    : Sender: TObject
  @Result   : None
-------------------------------------------------------------------------------}
Procedure TfrmConfiguration.optIndirectClick(Sender: TObject);
Begin

    grpIndirect.Visible := True;

End;

{*-------------------------------------------------------------------------------
  Procedure : TfrmConfiguration.rbEIPClick
  @Author   : CB le 09/05/2009
  @Param    : Sender: TObject
  @Result   : None
-------------------------------------------------------------------------------}
Procedure TfrmConfiguration.optEIPClick(Sender: TObject);
Begin

    If Self.bConfigEnCours Then Exit;

    Self.bConfigEnCours := True;
    rgbMesure.Controls[ord(etDerive)].Enabled := False;
    rgbMesure.Controls[ord(etInterval)].Enabled := False;
    rgbMesure.Controls[ord(etTachyContact)].Enabled := False;
    rgbMesure.Controls[ord(etTachyOptique)].Enabled := False;
    rgbMesure.Controls[ord(etStroboscope)].Enabled := False;

    If enTypesMesures(rgbMesure.ItemIndex) In [etDerive, etInterval, etTachyContact, etTachyOptique, etStroboscope] Then
        rgbMesure.ItemIndex := Ord(etFrequence);

    Self.bConfigEnCours := False;

    optDirect.Checked := True;
    optIndirect.Enabled := False;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmConfiguration.GetYearE2M
  @Author:   cb le 25/08/2009
  @Param:    None
  @Result:   String
-------------------------------------------------------------------------------}
Function TfrmConfiguration.GetYearE2M: String;
Var
    ilYear: Word;

Begin

    ilYear := YearOf(Now);
    Result := Chr(Ord('D') + ((ilYear Div 10) - 200)) + LeftStr(IntToStr(ilYear - 2000), 1);

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmConfiguration.ConfigMesure
  @Author:   cb le 25/08/2009
  @Param:    const Value: TMesure
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmConfiguration.ConfigMesure(Const Value: TMesure);
Begin

    FMesure := Value;
    With FMesure Do
    Begin
        If Trim(sNumFI) = '' Then
        Begin                                                                   // Affichage du n° de FI
            edtNoFI.Text := GetYearE2M();
            edtNoFI.SelStart := 3;
        End
        Else
            edtNoFI.EditText := sNumFI;

        rgbMesure.ItemIndex := Ord(TypeMesure);

        cbxMultiplicateur.ItemIndex := IndexMultiplicateur;

        If ModeMesure = emDirect Then
            optDirect.Checked := True
        Else
            optIndirect.Checked := True;

        Case Frequencemetre Of
            eaStanford: optStanford.Checked := True;
            eaRacal: optRacalDana.Checked := True;
        Else
            optEIP.Checked := True;
        End;

        edtFNominale.Text := FloatToStr(FNominale);

    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmConfiguration.btnCancelClick
  @Author:   cb le 25/08/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmConfiguration.btnCancelClick(Sender: TObject);
Begin

    Self.bAllowQuit := True;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmConfiguration.rgb1Click
  @Author:   cb le 29/09/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmConfiguration.rgbMesureClick(Sender: TObject);
Var
    elMesure: enTypesMesures;

Begin

    elMesure := enTypesMesures(rgbMesure.ItemIndex);

    // EIP sélectionnable seulement en mesure de fréquence
    optEIP.Enabled := (elMesure = etFrequence);
    If (Not optEIP.Enabled) And (optEIP.Checked) Then
        optStanford.Checked := True;

    // Mode indirect sélectionnable seulement si différent de Intervalle, Tachymètrie ou Stroboscope
    optIndirect.Enabled := (Not (elMesure In [etInterval, etTachyContact, etTachyOptique, etStroboscope]))
        And (Not (optEIP.Checked));
    If (Not optIndirect.Enabled) And (optIndirect.Checked) Then
        optDirect.Checked := True;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmConfiguration.optStanfordClick
  @Author:   cb le 29/09/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmConfiguration.optStanfordClick(Sender: TObject);
Var
    ilCnt: Integer;

Begin

    optIndirect.Enabled := True;
    For ilcnt := 0 To rgbMesure.ControlCount - 1 Do
        rgbMesure.Controls[ilCnt].Enabled := True;

End;

End.
