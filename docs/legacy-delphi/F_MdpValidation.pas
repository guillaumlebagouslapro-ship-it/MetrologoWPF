Unit F_MdpValidation;

Interface

Uses
    Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
    Dialogs, StdCtrls, JvExStdCtrls, JvButton, JvCtrls, U_DeclarationsMETROLOGO,
    AdvEdit;

Type
    TfrmMdpValidation = Class(TForm)
        btnOk: TJvImgBtn;
        btnCancel: TJvImgBtn;
        edtMdp: TAdvEdit;
        Procedure btnOkClick(Sender: TObject);
        Procedure btnCancelClick(Sender: TObject);
        Procedure FormCloseQuery(Sender: TObject; Var CanClose: Boolean);

    Private
        FPasswd: Boolean;
        bAllowQuit: Boolean;
        FExpectedPasswd: String;

    Public
        Property PasswordOk: Boolean Read FPasswd;
        Property ExpectedPasswd: String Read FExpectedPasswd Write FExpectedPasswd;

    End;

Var
    frmMdpValidation: TfrmMdpValidation;

Implementation

{$R *.dfm}

{*-------------------------------------------------------------------------------
  Procedure: TfrmMdpValidation.btnOkClick
  @Author:   cb le 25/08/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMdpValidation.btnOkClick(Sender: TObject);
Begin

    If UpperCase(edtMdp.Text) = Self.FExpectedPasswd Then
    Begin
        Self.FPasswd := True;
        Self.bAllowQuit := True;
    End
    Else
    Begin
        Beep;
        edtMdp.Text := '';
        edtMdp.SetFocus;
        Self.bAllowQuit := False;
    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMdpValidation.btnCancelClick
  @Author:   cb le 25/08/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMdpValidation.btnCancelClick(Sender: TObject);
Begin

    Self.FPasswd := False;
    Self.bAllowQuit := True;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMdpValidation.FormCloseQuery
  @Author:   cb le 25/08/2009
  @Param:    Sender: TObject; var CanClose: Boolean
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMdpValidation.FormCloseQuery(Sender: TObject; Var CanClose: Boolean);
Begin

    CanClose := Self.bAllowQuit;

End;

End.
