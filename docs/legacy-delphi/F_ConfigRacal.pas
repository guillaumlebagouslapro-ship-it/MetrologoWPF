Unit F_ConfigRacal;

Interface

Uses
    Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
    Dialogs, StdCtrls, JvExStdCtrls, JvButton, JvCtrls, ExtCtrls,
    JvExExtCtrls, JvRadioGroup, U_DeclarationsMETROLOGO;

Type
    TfrmConfigRacal = Class(TForm)
        chkInitManu: TCheckBox;
        grp1: TGroupBox;
        rgpGammesF: TJvRadioGroup;
        grpCadreCouplage: TGroupBox;
        rgpCouplage: TJvRadioGroup;
        btnOk: TJvImgBtn;
        btnCancel: TJvImgBtn;
        Procedure rgpGammesFClick(Sender: TObject);
        Procedure btnOkClick(Sender: TObject);

    Private
        FMesure: TMesure;
        Procedure ConfigMesure(Const Value: TMesure);

    Public
        Property oMesure: TMesure Read FMesure Write ConfigMesure;

    End;

Var
    frmConfigRacal: TfrmConfigRacal;

Implementation

{$R *.dfm}

Const
    AS_INPUT: Array[0..2] Of String = ('FN2 AZ1', 'FN2 AZ0', 'FN1');
    AS_IMPEDANCE: Array[0..1] Of String = ('AA1', 'AA0');
    S_SEP: String = ' ';

{ TfrmConfigRacal }

{*-------------------------------------------------------------------------------
  Procedure: TfrmConfigRacal.ConfigMesure
  @Author:   cb le 28/08/2009
  @Param:    const Value: TMesure
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmConfigRacal.ConfigMesure(Const Value: TMesure);
Begin

    FMesure := Value;

//    chkInitManu.Enabled := (FMesure.TypeMesure = etFrequence);
    chkInitManu.Enabled := (FMesure.TypeMesure in [etFrequence, etStabilite]);
    chkInitManu.Checked := FMesure.bInitManu;

    rgpGammesF.ItemIndex := Ord(FMesure.InputRacal);
    rgpCouplage.ItemIndex := Ord(FMesure.CouplingRacal);

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmConfigRacal.rgpGammesFClick
  @Author:   cb le 28/08/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmConfigRacal.rgpGammesFClick(Sender: TObject);
Begin

    grpCadreCouplage.Visible := Not (rgpGammesF.ItemIndex = Ord(eiInputC));

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmConfigRacal.btnOkClick
  @Author:   cb le 28/08/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmConfigRacal.btnOkClick(Sender: TObject);
Var
    olApp: TAppareilIEEE;

Begin

    With FMesure Do Begin
        bInitManu := chkInitManu.Checked;

        InputRacal := enInputRacal(rgpGammesF.ItemIndex);
        CouplingRacal := enInputCoupling(rgpCouplage.ItemIndex);
    End;

    olApp := TAppareilIEEE(ListeFrequencemetres.Items[ord(eaRacal)]);
    olApp.sConfEntree := AS_INPUT[rgpGammesF.ItemIndex];
    If (grpCadreCouplage.Visible) And (rgpCouplage.ItemIndex <> -1) Then
        olApp.sConfEntree := olApp.sConfEntree + S_SEP + AS_IMPEDANCE[rgpCouplage.ItemIndex];

End;

End.
