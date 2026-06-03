Unit F_DispIncert;

Interface

Uses
    Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
    Dialogs, ComCtrls, JvExComCtrls, JvComCtrls, JvComponentBase, ADODB,
    StdCtrls, AdvEdit, DBAdvEd, DB, JvExStdCtrls, JvButton, JvCtrls,
    F_MdpValidation, U_DeclarationsMETROLOGO, ImgList, ExtCtrls, Grids,
    BaseGrid, AdvGrid, DBAdvGrid, AdvPanel, E2MPanel;

Type
    TfrmDispIncert = Class(TForm)
        pgMode: TJvPageControl;
        pntPages: TJvTabDefaultPainter;
        tsFrequence: TTabSheet;
        tsStabilite: TTabSheet;
        grpParamGlobaux: TGroupBox;
        edtMbMesAcc: TDBAdvEdit;
        edtTpsMesAcc: TDBAdvEdit;
        edtIncertAcc: TDBAdvEdit;
        lbl1: TLabel;
        lbl2: TLabel;
        lbl3: TLabel;
        dstIncertGen: TADODataSet;
        srcIncertGen: TDataSource;
        lstIcones: TImageList;
        pnlBtns: TPanel;
        btnModif: TJvImgBtn;
        btnOk: TJvImgBtn;
        pnlLibForme: TPanel;
        dstFreqFixes: TADODataSet;
        dstFreq620: TADODataSet;
        dstFreq1996: TADODataSet;
        dstFreq545: TADODataSet;
        srcFreqFixes: TDataSource;
        srcFreq620: TDataSource;
        srcFreq1996: TDataSource;
        srcFreq545: TDataSource;
        grpStabFixes: TGroupBox;
        grpStabSR620: TGroupBox;
        grpStab1996: TGroupBox;
        grdStabSR620: TDBAdvGrid;
        grdStabFixes: TDBAdvGrid;
        grdStab1996: TDBAdvGrid;
        dstStabFixes: TADODataSet;
        dstStabSR620: TADODataSet;
        dstStab1996: TADODataSet;
        srcStabFixes: TDataSource;
        srcStabSR620: TDataSource;
        srcStab1996: TDataSource;
        tbcRubidiums: TJvTabControl;
        grpFreqFixes: TGroupBox;
        grdFreqFixes: TDBAdvGrid;
        grpStanford: TGroupBox;
        grdFreqSR620: TDBAdvGrid;
        grpRacal: TGroupBox;
        grdFreq1996: TDBAdvGrid;
        grpEIP: TGroupBox;
        grdFreq545: TDBAdvGrid;
        tsIntervalle: TTabSheet;
        pnlLibForme2: TPanel;
        tsTachyContact: TTabSheet;
        pnlLibForme3: TPanel;
        tsTachyOptique: TTabSheet;
        pnlLibForme4: TPanel;
        tsStroboscope: TTabSheet;
        pnlLibForme5: TPanel;
        dstAutres: TADODataSet;
        srcAutres: TDataSource;
        pnlIncertInterval: TE2MPanel;
        edtLibelleIntervalle: TDBAdvEdit;
        edtCoeffAIntervalle: TDBAdvEdit;
        edtCoeffBIntervalle: TDBAdvEdit;
        pnlIncertTachyContact: TE2MPanel;
        edtLibelleTachyContact: TDBAdvEdit;
        edtCoefATachyContact: TDBAdvEdit;
        edtCoefBTachyContact: TDBAdvEdit;
        pnlIncertTachyOptique: TE2MPanel;
        edtlibelleTachyOptique: TDBAdvEdit;
        edtCoeffATachyOptique: TDBAdvEdit;
        edtCoeffBTachyOptique: TDBAdvEdit;
        pnlIncertStroboscope: TE2MPanel;
        edtLibelleStroboscope: TDBAdvEdit;
        edtCoeffAStroboscope: TDBAdvEdit;
        edtCoeffBStroboscope: TDBAdvEdit;
        Procedure FormClose(Sender: TObject; Var Action: TCloseAction);
        Procedure btnModifClick(Sender: TObject);
        Procedure FormResize(Sender: TObject);
        Procedure dstFreqFixesAfterOpen(DataSet: TDataSet);
        Procedure dstFreq620AfterOpen(DataSet: TDataSet);
        Procedure dstFreq1996AfterOpen(DataSet: TDataSet);
        Procedure dstFreq545AfterOpen(DataSet: TDataSet);
        Procedure grdFreqFixesGetAlignment(Sender: TObject; ARow, ACol: Integer; Var HAlign: TAlignment; Var VAlign: TVAlignment);
        Procedure FormShow(Sender: TObject);
        Procedure grdFreqFixesCanEditCell(Sender: TObject; ARow, ACol: Integer;
            Var CanEdit: Boolean);
        Procedure dstStabFixesAfterOpen(DataSet: TDataSet);
        Procedure dstStabSR620AfterOpen(DataSet: TDataSet);
        Procedure dstStab1996AfterOpen(DataSet: TDataSet);
        Procedure grdStabFixesCanEditCell(Sender: TObject; ARow, ACol: Integer;
            Var CanEdit: Boolean);
        Procedure grdStabFixesGetAlignment(Sender: TObject; ARow,
            ACol: Integer; Var HAlign: TAlignment; Var VAlign: TVAlignment);
        Procedure tbcRubidiumsChange(Sender: TObject);
        Procedure pgModeChange(Sender: TObject);
    Private
        FCnxSQL: TADOConnection;
        aiID: Array Of Integer;                                                 // Tableau pour mémo des IDs Rubidium relatif à un onglet
        Procedure setFCnxSQL(Const Value: TADOConnection);

    Public
        Procedure ConfigGrille(opGrille: TDBAdvGrid);
        Procedure ConfigGrilleStab(opGrille: TDBAdvGrid);
        Procedure CloseDataSetsFreq();
        Property CnxSQL: TADOConnection Read FCnxSQL Write setFCnxSQL;

    End;

Var
    frmDispIncert: TfrmDispIncert;

Implementation

{$R *.dfm}
Const
    AI_COLSWIDTHS: Array[0..3] Of Integer = (100, 70, 50, 50);
    AI_COLSWIDTHS_STAB: Array[0..1] Of Integer = (200, 220);

{*-------------------------------------------------------------------------------
  Procedure: TfrmDispIncert.setFCnxSQL
  @Author:   cb le 22/09/2009
  @Param:    const Value: TADOConnection
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmDispIncert.setFCnxSQL(Const Value: TADOConnection);
Begin

    FCnxSQL := Value;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmDispIncert.FormClose
  @Author:   cb le 22/09/2009
  @Param:    Sender: TObject; var Action: TCloseAction
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmDispIncert.FormClose(Sender: TObject; Var Action: TCloseAction);
Begin

    CloseDataSetsFreq;
    If dstAutres.Active Then dstAutres.Close();

    dstStabFixes.Close();
    dstStabSR620.Close();
    dstStab1996.Close();
    dstIncertGen.Close();

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmDispIncert.btnModifClick
  @Author:   cb le 22/09/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmDispIncert.btnModifClick(Sender: TObject);
Begin

    With TfrmMdpValidation.Create(Self) Do
    Begin
        ExpectedPasswd := S_MDPINCERT;
        ShowModal();
        If PasswordOk Then
            grpParamGlobaux.Enabled := True;
        Free;
    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmDispIncert.FormResize
  @Author:   cb le 22/09/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmDispIncert.FormResize(Sender: TObject);
Begin

    grpFreqFixes.Width := tbcRubidiums.ClientWidth Div 4;
    grpStanford.Width := tbcRubidiums.ClientWidth Div 4;
    grpRacal.Width := tbcRubidiums.ClientWidth Div 4;

    grpStabFixes.Width := tsStabilite.ClientWidth Div 3;
    grpStabSR620.Width := tsStabilite.ClientWidth Div 3;

    btnModif.Left := (Self.ClientWidth Div 2) - 80 - btnModif.Width;
    btnOk.Left := (Self.ClientWidth Div 2) + 80;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmDispIncert.dstFreqFixesAfterOpen
  @Author:   cb le 23/09/2009
  @Param:    DataSet: TDataSet
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmDispIncert.dstFreqFixesAfterOpen(DataSet: TDataSet);
Begin

    ConfigGrille(grdFreqFixes);

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmDispIncert.ConfigGrille
  @Author:   cb le 23/09/2009
  @Param:    opGrille: TDBAdvGrid
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmDispIncert.ConfigGrille(opGrille: TDBAdvGrid);
Var
    ilCnt: Integer;

Begin

    With opGrille Do
    Begin
        For ilCnt := Low(AI_COLSWIDTHS) To High(AI_COLSWIDTHS) Do
            Columns.Items[ilCnt].Width := AI_COLSWIDTHS[ilCnt];

        Columns.Items[1].Color := clWhite;
        Columns.Items[2].Color := clWhite;
        Columns.Items[3].Color := clWhite;

        For ilCnt := 1 To RowCount - 1 Do
        Begin
            Colors[0, ilCnt] := clWhite;
            ColorsTo[0, ilCnt] := clSilver;
        End;

        ColumnSize.StretchColumn := 0;
        ColumnSize.Stretch := True;
    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmDispIncert.dstFreq620AfterOpen
  @Author:   cb le 23/09/2009
  @Param:    DataSet: TDataSet
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmDispIncert.dstFreq620AfterOpen(DataSet: TDataSet);
Begin

    ConfigGrille(grdFreqSR620);

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmDispIncert.dstFreq1996AfterOpen
  @Author:   cb le 23/09/2009
  @Param:    DataSet: TDataSet
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmDispIncert.dstFreq1996AfterOpen(DataSet: TDataSet);
Begin

    ConfigGrille(grdFreq1996);

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmDispIncert.dstFreq545AfterOpen
  @Author:   cb le 23/09/2009
  @Param:    DataSet: TDataSet
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmDispIncert.dstFreq545AfterOpen(DataSet: TDataSet);
Begin

    ConfigGrille(grdFreq545);

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmDispIncert.grdFreqFixesGetAlignment
  @Author:   cb le 23/09/2009
  @Param:    Sender: TObject; ARow, ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmDispIncert.grdFreqFixesGetAlignment(Sender: TObject; ARow, ACol: Integer; Var HAlign: TAlignment; Var VAlign: TVAlignment);
Begin

    With TDBAdvGrid(Sender) Do
    Begin
        CellProperties[ACol, ARow].FontName := 'Comic Sans MS';
        VAlign := vtaCenter;
        If ARow = 0 Then
        Begin
            HAlign := taCenter;
            CellProperties[ACol, ARow].FontSize := 10;
        End
        Else
        Begin
            If ACol > 1 Then
                CellProperties[ACol, ARow].Editor := edFloat;

            HAlign := taLeftJustify;
            CellProperties[ACol, ARow].FontSize := 8;
        End;
    End;

End;
{*-------------------------------------------------------------------------------
  Procedure: TfrmDispIncert.FormShow
  @Author:   cb le 23/09/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmDispIncert.FormShow(Sender: TObject);
Var
    ilCnt, ilIdx, ilNbRubidiums, ilTabIndex: Integer;

Begin

    // Affectation des connexions
    dstIncertGen.Connection := CnxSQL;

    dstFreqFixes.Connection := CnxSQL;
    dstFreq620.Connection := CnxSQL;
    dstFreq1996.Connection := CnxSQL;
    dstFreq545.Connection := CnxSQL;
    dstAutres.Connection := CnxSQL;

    dstStabFixes.Connection := CnxSQL;
    dstStabSR620.Connection := CnxSQL;
    dstStab1996.Connection := CnxSQL;

    // Maj CommandText
    dstIncertGen.CommandText := S_REQINFOS_INCERT;
    dstStabFixes.CommandText := Format(S_REQ_DISPINCERT_STAB, [-1]);
    dstStabSR620.CommandText := Format(S_REQ_DISPINCERT_STAB, [Ord(eaStanford)]);
    dstStab1996.CommandText := Format(S_REQ_DISPINCERT_STAB, [Ord(eaRacal)]);

    // Ouverture des dataSets
    dstIncertGen.Open();
    dstStabFixes.Open();
    dstStabSR620.Open();
    dstStab1996.Open();

    // Génération des onglets (1 par Rubidium)
    With TADOQuery.Create(Self) Do
    Begin
        Connection := CnxSQL;
        sql.Clear;
        SQL.Add(S_REQ_RUBIDIUMS);
        Open;
        ilNbRubidiums := Recordcount;
        If ilNbRubidiums < 1 Then
        Begin
            Close;                                                              // Ferme query
            Application.MessageBox('Aucun rubidium trouvé dans la table SQL !', 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
            ModalResult := mrCancel;
            Self.Close;                                                         // Fereme fenetre
        End;

        SetLength(aiID, ilNbRubidiums);                                         // Dimensionne tableau de mémo des IDs de Rubidium
        ilTabIndex := 0;
        For ilCnt := 0 To ilNbRubidiums - 1 Do
        Begin
            ilIdx := tbcRubidiums.Tabs.Add(FieldByName(S_RUB_DESIGNATION).AsString);
            aiID[ilCnt] := FieldByName(S_RUB_ID).AsInteger;
            If FieldByName(S_RUB_ACTIF).AsBoolean Then                          // Mémorise index de l'onglet relatif au rubidium actif
                ilTabIndex := ilIdx;
            Next;
        End;

        tbcRubidiums.TabIndex := ilTabIndex;
        tbcRubidiumsChange(Self);
        tbcRubidiums.Refresh;

    End;

    pgMode.ActivePage := tsFrequence;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmDispIncert.grdFreqFixesCanEditCell
  @Author:   cb le 23/09/2009
  @Param:    Sender: TObject; ARow, ACol: Integer; var CanEdit: Boolean
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmDispIncert.grdFreqFixesCanEditCell(Sender: TObject; ARow, ACol: Integer; Var CanEdit: Boolean);
Begin

    CanEdit := False;
    If (ACol = 0) Or (ARow = 0) Then Exit;
    CanEdit := True;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmDispIncert.dstStabFixesAfterOpen
  @Author:   cb le 23/09/2009
  @Param:    DataSet: TDataSet
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmDispIncert.dstStabFixesAfterOpen(DataSet: TDataSet);
Begin

    ConfigGrilleStab(grdStabFixes);

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmDispIncert.dstStabSR620AfterOpen
  @Author:   cb le 23/09/2009
  @Param:    DataSet: TDataSet
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmDispIncert.dstStabSR620AfterOpen(DataSet: TDataSet);
Begin

    ConfigGrilleStab(grdStabSR620);

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmDispIncert.dstStab1996AfterOpen
  @Author:   cb le 23/09/2009
  @Param:    DataSet: TDataSet
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmDispIncert.dstStab1996AfterOpen(DataSet: TDataSet);
Begin

    ConfigGrilleStab(grdStab1996);

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmDispIncert.ConfigGrilleStab
  @Author:   cb le 23/09/2009
  @Param:    opGrille: TDBAdvGrid
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmDispIncert.ConfigGrilleStab(opGrille: TDBAdvGrid);
Var
    ilCnt: Integer;

Begin

    With opGrille Do
    Begin
        For ilCnt := Low(AI_COLSWIDTHS_STAB) To High(AI_COLSWIDTHS_STAB) Do
            Columns.Items[ilCnt].Width := AI_COLSWIDTHS_STAB[ilCnt];

        Columns.Items[1].Color := clWhite;

        For ilCnt := 1 To RowCount - 1 Do
        Begin
            Colors[0, ilCnt] := clWhite;
            ColorsTo[0, ilCnt] := clSilver;
        End;

        ColumnSize.StretchColumn := 0;
        ColumnSize.Stretch := True;
    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmDispIncert.grdStabFixesCanEditCell
  @Author:   cb le 23/09/2009
  @Param:    Sender: TObject; ARow, ACol: Integer; var CanEdit: Boolean
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmDispIncert.grdStabFixesCanEditCell(Sender: TObject; ARow, ACol: Integer; Var CanEdit: Boolean);
Begin

    CanEdit := False;
    If (ACol = 0) Or (ARow = 0) Then Exit;
    CanEdit := True;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmDispIncert.grdStabFixesGetAlignment
  @Author:   cb le 23/09/2009
  @Param:    Sender: TObject; ARow, ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmDispIncert.grdStabFixesGetAlignment(Sender: TObject; ARow, ACol: Integer; Var HAlign: TAlignment; Var VAlign: TVAlignment);
Begin

    With TDBAdvGrid(Sender) Do
    Begin
        CellProperties[ACol, ARow].FontName := 'Comic Sans MS';
        VAlign := vtaCenter;
        If ARow = 0 Then
        Begin
            HAlign := taCenter;
            CellProperties[ACol, ARow].FontSize := 10;
        End
        Else
        Begin
            If ACol > 0 Then
            Begin
                HAlign := taLeftJustify;
                CellProperties[ACol, ARow].Editor := edFloat;
            End
            Else
                HAlign := taCenter;
            CellProperties[ACol, ARow].FontSize := 8;
        End;
    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmDispIncert.tbcRubidiumsChange
  @Author:   cb le 29/09/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmDispIncert.tbcRubidiumsChange(Sender: TObject);
Var
    ilIdxID: Integer;

Begin

    CloseDataSetsFreq;
    ilIdxID := aiID[tbcRubidiums.TabIndex];

    // Maj commantext
    dstFreqFixes.CommandText := Format(S_REQ_DISPINCERT_FREQ, [ilIdxID, -1]);
    dstFreq620.CommandText := Format(S_REQ_DISPINCERT_FREQ, [ilIdxID, Ord(eaStanford)]);
    dstFreq1996.CommandText := Format(S_REQ_DISPINCERT_FREQ, [ilIdxID, Ord(eaRacal)]);
    dstFreq545.CommandText := Format(S_REQ_DISPINCERT_FREQ, [ilIdxID, Ord(eaEIP)]);

    // Ouverture DataSet
    dstFreqFixes.Open();
    dstFreq620.Open();
    dstFreq1996.Open();
    dstFreq545.Open();

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmDispIncert.CloseDataSets
  @Author:   cb le 29/09/2009
  @Param:    None
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmDispIncert.CloseDataSetsFreq;
Begin

    If dstFreqFixes.Active Then dstFreqFixes.Close();
    If dstFreq620.Active Then dstFreq620.Close();
    If dstFreq1996.Active Then dstFreq1996.Close();
    If dstFreq545.Active Then dstFreq545.Close();

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmDispIncert.pgModeChange
  @Author:   cb le 29/09/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmDispIncert.pgModeChange(Sender: TObject);
Var
    ilIDMesure: Integer;

Begin

    If pgMode.ActivePage = tsIntervalle Then
        ilIDMesure := Ord(etInterval)
    Else If pgMode.ActivePage = tsTachyContact Then
        ilIDMesure := Ord(etTachyContact)
    Else If pgMode.ActivePage = tsTachyOptique Then
        ilIDMesure := Ord(etTachyOptique)
    Else If pgMode.ActivePage = tsStroboscope Then
        ilIDMesure := Ord(etStroboscope)
    Else
        Exit;

    If dstAutres.Active Then dstAutres.Close();
    dstAutres.CommandText := Format(S_REQ_DISPINCERT_AUTRES, [ilIDMesure]);
    dstAutres.Open();

End;

End.
