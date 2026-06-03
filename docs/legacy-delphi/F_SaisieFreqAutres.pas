Unit F_SaisieFreqAutres;

Interface

Uses
    Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
    Dialogs, ADODB, StdCtrls, AdvEdit, JvExStdCtrls, JvButton, JvCtrls;

Type
    TfrmGetFreqAutresRubis = Class(TForm)
        edtRubidium1: TAdvEdit;
        edtRubidium2: TAdvEdit;
        btnOk: TJvImgBtn;
        btnCancel: TJvImgBtn;
        lblTxt: TLabel;
        Procedure FormCreate(Sender: TObject);
        Procedure FormCloseQuery(Sender: TObject; Var CanClose: Boolean);
        Procedure btnCancelClick(Sender: TObject);
        Procedure btnOkClick(Sender: TObject);
        Procedure FormShow(Sender: TObject);
    Private
        bAllowQuit: Boolean;
        FCnx: TADOConnection;
        FValRubi1: Double;
        FValRubi2: Double;
        FIDRubi1: Integer;
        FIDRubi2: Integer;

    Public
        Property Cnx: TADOConnection Read FCnx Write FCnx;
        Property IDRubi1: Integer Read FIDRubi1 Write FIDRubi1;
        Property IDRubi2: Integer Read FIDRubi2 Write FIDRubi2;
        Property ValRubi1: Double Read FValRubi1 Write FValRubi1;
        Property ValRubi2: Double Read FValRubi2 Write FValRubi2;

    End;

Var
    frmGetFreqAutresRubis: TfrmGetFreqAutresRubis;

Implementation

{$R *.dfm}

{*-------------------------------------------------------------------------------
  Procedure: TfrmGetFreqAutresRubis.FormCreate
  @Author:   cb le 13/10/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmGetFreqAutresRubis.FormCreate(Sender: TObject);
Begin

    Self.bAllowQuit := True;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmGetFreqAutresRubis.FormCloseQuery
  @Author:   cb le 13/10/2009
  @Param:    Sender: TObject; var CanClose: Boolean
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmGetFreqAutresRubis.FormCloseQuery(Sender: TObject; Var CanClose: Boolean);
Begin

    CanClose := Self.bAllowQuit;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmGetFreqAutresRubis.btnCancelClick
  @Author:   cb le 13/10/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmGetFreqAutresRubis.btnCancelClick(Sender: TObject);
Begin

    Self.bAllowQuit := True;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmGetFreqAutresRubis.btnOkClick
  @Author:   cb le 13/10/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmGetFreqAutresRubis.btnOkClick(Sender: TObject);
Begin

    Self.bAllowQuit := False;
    If (Not TryStrToFloat(edtRubidium1.Text, FValRubi1)) Or (ValRubi1 = 0.0) Then Begin
        Application.MessageBox('Valeur incorrecte !', 'Attention', MB_OK + MB_ICONWARNING + MB_TOPMOST);
        edtRubidium1.SetFocus();
        Exit;
    End;

    If (Not TryStrToFloat(edtRubidium2.Text, FValRubi2)) Or (ValRubi2 = 0.0) Then Begin
        Application.MessageBox('Valeur incorrecte !', 'Attention', MB_OK + MB_ICONWARNING + MB_TOPMOST);
        edtRubidium2.SetFocus();
        Exit;
    End;

    FIDRubi1 := edtRubidium1.Tag;
    FIDRubi2 := edtRubidium2.Tag;
    
    Self.bAllowQuit := True;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmGetFreqAutresRubis.FormShow
  @Author:   cb le 13/10/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmGetFreqAutresRubis.FormShow(Sender: TObject);
Begin

    With TADOQuery.Create(Self) Do Begin
        Connection := Self.FCnx;
        SQL.Clear;
        SQL.Add('Select RUB_ID, RUB_DESIGNATION from TR_METROLOGO_RUBIDIUMS where RUB_ACTIF=0 order by RUB_ID');
        Open;

        If RecordCount <> 2 Then Begin
            btnOk.Enabled := False;

        End
        Else Begin
            btnOk.Enabled := True;
            edtRubidium1.LabelCaption := FieldByName('RUB_DESIGNATION').AsString;
            edtRubidium1.Tag := FieldByName('RUB_ID').AsInteger;
            Next;
            edtRubidium2.LabelCaption := FieldByName('RUB_DESIGNATION').AsString;
            edtRubidium2.Tag := FieldByName('RUB_ID').AsInteger;
        End;

        Close();
        Free();

    End;

End;

End.

