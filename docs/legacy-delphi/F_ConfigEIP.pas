Unit F_ConfigEIP;

Interface

Uses
    Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
    Dialogs, StdCtrls, JvExStdCtrls, JvButton, JvCtrls, ExtCtrls,
    JvExExtCtrls, JvRadioGroup, U_DeclarationsMETROLOGO;

Type
    TfrmConfigEIP = Class(TForm)
        chkInitManu: TCheckBox;
        grp1: TGroupBox;
        rgpGammesF: TJvRadioGroup;
        btnOk: TJvImgBtn;
        btnCancel: TJvImgBtn;
        Procedure btnOkClick(Sender: TObject);
    Private
        FMesure: TMesure;
        Procedure ConfigMesure(Const Value: TMesure);

    Public
        Property oMesure: TMesure Read FMesure Write ConfigMesure;
    End;

Var
    frmConfigEIP: TfrmConfigEIP;

Implementation

{$R *.dfm}

Const
    AS_INPUT: Array[0..2] Of String = ('B1', 'B2', 'B3');

{ TfrmConfigEIP }

{*-------------------------------------------------------------------------------
  Procedure: TfrmConfigEIP.ConfigMesure
  @Author:   cb le 31/08/2009
  @Param:    const Value: TMesure
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmConfigEIP.ConfigMesure(Const Value: TMesure);
Begin

    FMesure := Value;

    chkInitManu.Enabled := (FMesure.TypeMesure = etFrequence);
    chkInitManu.Checked := FMesure.bInitManu;

    rgpGammesF.ItemIndex := Ord(FMesure.InputEIP);

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmConfigEIP.btnOkClick
  @Author:   cb le 31/08/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmConfigEIP.btnOkClick(Sender: TObject);
Var
    olApp: TAppareilIEEE;

Begin

    With FMesure Do Begin
        bInitManu := chkInitManu.Checked;

        InputEIP := enInputEIP(rgpGammesF.ItemIndex);
    End;

    olApp := TAppareilIEEE(ListeFrequencemetres.Items[ord(eaEIP)]);
    olApp.sConfEntree := AS_INPUT[rgpGammesF.ItemIndex];

End;

End.
