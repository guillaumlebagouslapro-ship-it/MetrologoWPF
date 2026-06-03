Unit F_GetNbMesures;

Interface

Uses
    Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
    Dialogs, StdCtrls, AdvEdit, JvExStdCtrls, JvButton, JvCtrls,
    U_DeclarationsMETROLOGO;

Type
    TfrmGetNbMesures = Class(TForm)
        btnOk: TJvImgBtn;
        btnCancel: TJvImgBtn;
        edtNbMesures: TAdvEdit;
        Procedure FormCloseQuery(Sender: TObject; Var CanClose: Boolean);
        Procedure btnCancelClick(Sender: TObject);
        Procedure btnOkClick(Sender: TObject);

    Private
        FMesure: TMesure;
        FbAllowQuit: Boolean;
        Procedure ConfigMesure(Const Value: TMesure);

    Public
        Property oMesure: TMesure Read FMesure Write ConfigMesure;
    End;

Var
    frmGetNbMesures: TfrmGetNbMesures;

Implementation

{$R *.dfm}

{*-------------------------------------------------------------------------------
  Procedure: TfrmGetNbMesures.ConfigMesure
  @Author:   cb le 03/09/2009
  @Param:    const Value: TMesure
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmGetNbMesures.ConfigMesure(Const Value: TMesure);
Begin

    FMesure := Value;
    FbAllowQuit := False;
    if FMesure.TypeMesure = etInterval then
        edtNbMesures.Text := '1'
    else
        edtNbMesures.Text := IntToStr(FMesure.NbMesures);

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmGetNbMesures.FormCloseQuery
  @Author:   cb le 03/09/2009
  @Param:    Sender: TObject; var CanClose: Boolean
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmGetNbMesures.FormCloseQuery(Sender: TObject; Var CanClose: Boolean);
Begin

    CanClose := FbAllowQuit;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmGetNbMesures.btnCancelClick
  @Author:   cb le 03/09/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmGetNbMesures.btnCancelClick(Sender: TObject);
Begin

    FbAllowQuit := True;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmGetNbMesures.btnOkClick
  @Author:   cb le 03/09/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmGetNbMesures.btnOkClick(Sender: TObject);
Var
    blOk: Boolean;
    ilVal: Integer;

Begin

    blOk := False;
    While True Do
    Begin
        If Not TryStrToInt(edtNbMesures.Text, ilVal) Then break;
        If ilVal < 1 Then Break;
        blOk := True;
        Break;
    End;

    If blOk Then
        FMesure.NbMesures := ilVal
    Else
    Begin
        Application.MessageBox(PChar(Format('Nombre incorrect : %s', [edtNbMesures.Text])), 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
        edtNbMesures.SetFocus;
    End;

    FbAllowQuit := blOk;

End;

End.
