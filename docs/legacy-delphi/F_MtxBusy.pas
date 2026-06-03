Unit F_MtxBusy;

Interface

Uses
    Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
    Dialogs, ExtCtrls, StdCtrls, JvExStdCtrls, JvButton, JvCtrls;

Type
    TfrmMutexBusy = Class(TForm)
        lblInfo: TLabel;
        imgWarning: TImage;
        btnCancel: TJvImgBtn;
        Procedure FormCreate(Sender: TObject);
        Procedure btnCancelClick(Sender: TObject);
    Private
        FGiveUp: Boolean;

    Public
        Property bGiveUp: Boolean Read FGiveUp;

    End;

Var
    frmMutexBusy: TfrmMutexBusy;

Implementation

{$R *.dfm}

{*-------------------------------------------------------------------------------
  Procedure: TfrmMutexBusy.FormCreate
  @Author:   cb le 07/09/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMutexBusy.FormCreate(Sender: TObject);
Begin

    Self.FGiveUp := False;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmMutexBusy.btnCancelClick
  @Author:   cb le 07/09/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmMutexBusy.btnCancelClick(Sender: TObject);
Begin

    FGiveUp := True;

End;

End.
