Unit F_ParamsIncert;

Interface

Uses
    Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
    Dialogs, StdCtrls, AdvEdit, JvExStdCtrls, JvButton, JvCtrls, StrUtils,
  U_OutilsString;

Type
    TfrmParamsIncert = Class(TForm)
        edtResolution: TAdvEdit;
        edtIncertSupp: TAdvEdit;
        btnOk: TJvImgBtn;
    procedure btnOkClick(Sender: TObject);
    procedure edtIncertSuppKeyPress(Sender: TObject; var Key: Char);
    procedure FormShow(Sender: TObject);
    Private
        FResol: Double;
        FIncertSupp: Double;
        Procedure setFresol(Const Value: Double);
        Procedure setFIncertSupp(Const Value: Double);

    Public
        Property dResol: Double Read FResol Write setFresol;
        Property dIncertSupp: Double Read FIncertSupp Write setFIncertSupp;

    End;

Var
    frmParamsIncert: TfrmParamsIncert;

Implementation

{$R *.dfm}

{ TfrmParamsIncert }

{*-------------------------------------------------------------------------------
  Procedure: TfrmParamsIncert.setFIncertSupp
  @Author:   cb le 22/09/2009
  @Param:    const Value: Double
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmParamsIncert.setFIncertSupp(Const Value: Double);
Begin

    FIncertSupp := Value;
    edtIncertSupp.Text := FloatToStr(Value);

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmParamsIncert.setFresol
  @Author:   cb le 22/09/2009
  @Param:    const Value: Double
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmParamsIncert.setFresol(Const Value: Double);
Begin

    FResol := Value;
    edtResolution.Text := FloatToStr(Value);

End;


{*-------------------------------------------------------------------------------
  Procedure: TfrmParamsIncert.btnOkClick
  @Author:   cb le 19/10/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
procedure TfrmParamsIncert.btnOkClick(Sender: TObject);
begin

    Self.FResol := StrToFloat(edtResolution.Text);
    Self.FIncertSupp := StrToFloat(edtIncertSupp.Text);

end;

procedure TfrmParamsIncert.edtIncertSuppKeyPress(Sender: TObject; var Key: Char);
var
    slVal: string;
    dlVal: Double;

begin

    slVal := '';
    with edtIncertSupp do begin
        if SelStart > 0 then
            slVal := LeftStr(Text, SelStart);
        slVal := slVal + Key;
        if SelStart + SelLength < length(Text) then
            slVal := slVal + Mid(Text, SelStart + SelLength + 1);

    end;

    if not TryStrToFloat(slVal, dlVal) then
        Key := #0;

end;


{*-------------------------------------------------------------------------------
  Procedure: TfrmParamsIncert.FormShow
  @Author:   cb le 09/03/2011
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
procedure TfrmParamsIncert.FormShow(Sender: TObject);
begin

    BringWindowToTop(Self.Handle);
    edtResolution.SetFocus;

end;

End.
