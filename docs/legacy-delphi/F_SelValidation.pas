Unit F_SelValidation;

Interface

Uses
    Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
    Dialogs, OleServer, ExcelXP, U_DeclarationsMETROLOGO, StdCtrls,
    ExtCtrls, JvExStdCtrls, JvButton, JvCtrls;

Type
    TfrmSelValidation = Class(TForm)
        wbValidation: TExcelWorkbook;
        btnOk: TJvImgBtn;
        btnCancel: TJvImgBtn;
        rgbValidations: TRadioGroup;
        wsValidation: TExcelWorksheet;
        Procedure FormClose(Sender: TObject; Var Action: TCloseAction);
        Procedure btnOkClick(Sender: TObject);
    Private
        FExcel: TExcelApplication;
        FTabMes: PTabDouble;
        Procedure SetExcel(Const Value: TExcelApplication);
        Procedure setTabMes(Const Value: PTabDouble);

    Public
        Property oExcel: TExcelApplication Read FExcel Write SetExcel;
        Property adMes: PTabDouble Read FTabMes Write SetTabMes;

    End;

Var
    frmSelValidation: TfrmSelValidation;

Implementation

{$R *.dfm}

{ TfrmSelValidation }

{*-------------------------------------------------------------------------------
  Procedure: TfrmSelValidation.SetExcel
  @Author:   cb le 30/09/2009
  @Param:    const Value: TExcelApplication
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmSelValidation.SetExcel(Const Value: TExcelApplication);
Var
    ilCnt: enZNValidation;
    XLRange: ExcelRange;

Begin

    FExcel := Value;
    Self.oExcel.Workbooks.Open(sClasseurValidation, EmptyParam, OLEFalse, EmptyParam, EmptyParam, EmptyParam, OLETrue, EmptyParam, EmptyParam,
        EmptyParam, EmptyParam, EmptyParam, OLETrue, EmptyParam, EmptyParam, LCID);

    With wbValidation Do
    Begin
        ConnectTo(Self.oExcel.ActiveWorkbook);
        wsValidation.ConnectTo(wbValidation.Sheets[S_FEUILLE_VALIDATION] As _WorkSheet);

        For ilcnt := Low(AT_ZONESVALIDATION) To High(AT_ZONESVALIDATION) Do
        Begin
            XLRange := Names.Item(AT_ZONESVALIDATION[ilCnt].ZNLibelle, EmptyParam, EmptyParam).RefersToRange; // Récup nom mesure
            rgbValidations.Items.Add(XLRange.Value2);
        End;

    End;

    rgbValidations.ItemIndex := 0;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmSelValidation.FormClose
  @Author:   cb le 30/09/2009
  @Param:    Sender: TObject; var Action: TCloseAction
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmSelValidation.FormClose(Sender: TObject; Var Action: TCloseAction);
Begin

    wsValidation.Disconnect;
    wbValidation.Close(OLEFalse);
    wbValidation.Disconnect;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmSelValidation.setTabMes
  @Author:   cb le 30/09/2009
  @Param:    const Value: ^array of Double
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmSelValidation.setTabMes(Const Value: PTabDouble);
Begin

    FTabMes := Value;
    SetLength(FTabMes^, 0);

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmSelValidation.btnOkClick
  @Author:   cb le 30/09/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmSelValidation.btnOkClick(Sender: TObject);
Var
    XLRange: ExcelRange;
    ilCnt, ilNb, ilIdx: Integer;
    ilCol: Integer;

Begin

    XLRange := wbValidation.Names.Item(AT_ZONESVALIDATION[enZNValidation(rgbValidations.ItemIndex)].ZNDonnees, EmptyParam, EmptyParam).RefersToRange; // Récup plage de données
    ilCol := XLRange.Column;

    ilNb := XLRange.Rows.Count;
    SetLength(FTabMes^, ilNb);

    ilIdx := 0;
    For ilCnt := XLRange.Row To XLRange.Row + ilNb - 1 Do
    Begin
        FTabMes^[ilIdx] := wsValidation.Cells.Item[ilCnt, ilCol].Value2;
        inc(ilIdx);
    End;

End;

End.
