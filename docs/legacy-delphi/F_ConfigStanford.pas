Unit F_ConfigStanford;

Interface

Uses
    Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
    Dialogs, StdCtrls, JvExStdCtrls, JvButton, JvCtrls, ExtCtrls,
    JvExExtCtrls, JvRadioGroup, U_DeclarationsMETROLOGO;

Type
    TfrmConfigStanford = Class(TForm)
        chkInitManu: TCheckBox;
        btnOk: TJvImgBtn;
        btnCancel: TJvImgBtn;
        grp1: TGroupBox;
        rgpGammesF: TJvRadioGroup;
        grpCadreCouplage: TGroupBox;
        rgpCouplage: TJvRadioGroup;
        Procedure btnOkClick(Sender: TObject);
        Procedure rgpGammesFClick(Sender: TObject);
    Private
        FMesure: TMesure;
        Procedure ConfigMesure(Const Value: TMesure);

    Public
        Property oMesure: TMesure Read FMesure Write ConfigMesure;

    End;

Var
    frmConfigStanford: TfrmConfigStanford;

Implementation

{$R *.dfm}

Const
    AS_INPUT: Array[0..2] Of String = ('term1,0', 'term1,1', 'term1,2');
    AS_IMPEDANCE: Array[0..1] Of String = ('tcpl1,1', 'tcpl1,0');
    S_SEP: String = ';';

{ TfrmConfigStanford }

{*-------------------------------------------------------------------------------
  Procedure: TfrmConfigStanford.ConfigMesure
  @Author:   cb le 28/08/2009
  @Param:    const Value: TMesure
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmConfigStanford.ConfigMesure(Const Value: TMesure);
Begin

    FMesure := Value;

    chkInitManu.Enabled := (FMesure.TypeMesure = etFrequence);
    chkInitManu.Checked := FMesure.bInitManu;

    rgpGammesF.ItemIndex := Ord(FMesure.InputStanford);
    rgpCouplage.ItemIndex := Ord(FMesure.CouplingStanford);

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmConfigStanford.btnOkClick
  @Author:   cb le 28/08/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmConfigStanford.btnOkClick(Sender: TObject);
Var
    olApp: TAppareilIEEE;

Begin

    With FMesure Do
    Begin
        bInitManu := chkInitManu.Checked;

        InputStanford := enInputStanford(rgpGammesF.ItemIndex);
        CouplingStanford := enInputCoupling(rgpCouplage.ItemIndex);

    End;

    olApp := TAppareilIEEE(ListeFrequencemetres.Items[ord(eaStanford)]);
    olApp.sConfEntree := AS_INPUT[rgpGammesF.ItemIndex];
    If (grpCadreCouplage.Visible) And (rgpCouplage.ItemIndex <> -1) Then
        olApp.sConfEntree := olApp.sConfEntree + S_SEP + AS_IMPEDANCE[rgpCouplage.ItemIndex];

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmConfigStanford.rgpGammesFClick
  @Author:   cb le 28/08/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmConfigStanford.rgpGammesFClick(Sender: TObject);
Begin

    grpCadreCouplage.Visible := Not (rgpGammesF.ItemIndex = Ord(eiStanfordUHF));

End;

End.
