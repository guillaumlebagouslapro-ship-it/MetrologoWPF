Unit F_SaisieValFres;

Interface

Uses
    Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
    Dialogs, StdCtrls, JvExStdCtrls, JvButton, JvCtrls, AdvEdit;

Type
    TfrmSaisieValFreq = Class(TForm)
        lbl1: TLabel;
        edtValFreq: TAdvEdit;
        btnOk: TJvImgBtn;
    procedure FormShow(Sender: TObject);
    procedure btnOkClick(Sender: TObject);

    Private
        FFreqLue: Double;
        Procedure setFFreqLue(Const Value: Double);

    Public
        Property dFreqLue: Double Read FFreqLue Write setFFreqLue;

    End;

Var
    frmSaisieValFreq: TfrmSaisieValFreq;

Implementation

{$R *.dfm}

{ TfrmSaisieValFreq }

{*-------------------------------------------------------------------------------
  Procedure: TfrmSaisieValFreq.setFFreqLue
  @Author:   cb le 22/09/2009
  @Param:    const Value: Double
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmSaisieValFreq.setFFreqLue(Const Value: Double);
Begin

    FFreqLue := Value;
    edtValFreq.Text := FloatToStr(Value);

End;


{*-------------------------------------------------------------------------------
  Procedure: TfrmSaisieValFreq.FormShow
  @Author:   cb le 19/10/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
procedure TfrmSaisieValFreq.FormShow(Sender: TObject);
begin

    BringWindowToTop(Self.Handle);
    edtValFreq.SetFocus;

end;



{*-------------------------------------------------------------------------------
  Procedure: TfrmSaisieValFreq.btnOkClick
  @Author:   cb le 19/10/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
procedure TfrmSaisieValFreq.btnOkClick(Sender: TObject);
begin

    Self.FFreqLue := StrToFloat(edtValFreq.Text);

end;

End.
