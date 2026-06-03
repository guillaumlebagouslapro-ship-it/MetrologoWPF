Unit F_VoiesMux;

Interface

Uses
    Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
    Dialogs, StdCtrls, ExtCtrls, JvExStdCtrls, JvButton, JvCtrls,
    JvExExtCtrls, JvRadioGroup, U_DeclarationsMETROLOGO;

Type
    TfrmVoieMux = Class(TForm)
        btnOk: TJvImgBtn;
        btnCancel: TJvImgBtn;
        lbl1: TLabel;
        grp1: TGroupBox;
        rgpVoies: TJvRadioGroup;
        Procedure btnOkClick(Sender: TObject);
    Private
        FMesure: TMesure;
        Procedure ConfigMesure(Const Value: TMesure);

    Public
        Property oMesure: TMesure Read FMesure Write ConfigMesure;
    End;

Var
    frmVoieMux: TfrmVoieMux;

Implementation

{$R *.dfm}

{ TfrmVoieMux }

{*-------------------------------------------------------------------------------
  Procedure: TfrmVoieMux.ConfigMesure
  @Author:   cb le 28/08/2009
  @Param:    const Value: TMesure
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmVoieMux.ConfigMesure(Const Value: TMesure);
Begin

    FMesure := Value;
    rgpVoies.ItemIndex := FMesure.VoieMux;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmVoieMux.btnOkClick
  @Author:   cb le 28/08/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmVoieMux.btnOkClick(Sender: TObject);
Begin

    FMesure.VoieMux := rgpVoies.ItemIndex;

End;

End.
