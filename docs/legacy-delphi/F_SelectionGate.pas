Unit F_SelectionGate;

Interface

Uses
    Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
    Dialogs, StdCtrls, AdvCombo, JvExStdCtrls, JvButton, JvCtrls, ExtCtrls,
    JvExExtCtrls, JvRadioGroup, U_DeclarationsMETROLOGO;

Type
    TfrmSelectionGate = Class(TForm)
        cbxValGate: TAdvComboBox;
        grpProcedures: TGroupBox;
        rgpProcAutos: TJvRadioGroup;
        btnOk: TJvImgBtn;
        btnCancel: TJvImgBtn;
        Procedure rgpProcAutosClick(Sender: TObject);
        Procedure cbxValGateChange(Sender: TObject);
        Procedure btnCancelClick(Sender: TObject);
        Procedure btnOkClick(Sender: TObject);
    Private
        FMesure: TMesure;
        Procedure ConfigMesure(Const Value: TMesure);

    Public
        Property oMesure: TMesure Read FMesure Write ConfigMesure;

    End;

Const
    S_PROC1: String = 'Procédure Auto 1 : 10ms -> 100s';                        // Procédure stab 10ms à 100s pour fréquencemètres autres que EIP545
    S_PROC2: String = 'Procédure Auto 2 : 10ms -> 10s';                         // Procédure stab 10ms à 10s pour fréquencemètres autres que EIP545
    S_PROC3: String = 'Procédure Auto : 10ms, 100ms et 1s';                     // Procédure stab 10ms à 1s pour fréquencemètre EIP545

Var
    frmSelectionGate: TfrmSelectionGate;

Implementation

{$R *.dfm}

{ TfrmSelectionGate }

{*-------------------------------------------------------------------------------
  Procedure: TfrmSelectionGate.ConfigMesure
  @Author:   cb le 02/09/2009
  @Param:    const Value: TMesure
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmSelectionGate.ConfigMesure(Const Value: TMesure);
Begin

    Self.FMesure := Value;

    With Self.FMesure Do
    Begin

        If Frequencemetre = eaEIP Then
            rgpProcAutos.Items.Add(S_PROC3)
        Else
        Begin
            rgpProcAutos.Items.Add(S_PROC1);
            rgpProcAutos.Items.Add(S_PROC2);
        End;

        cbxValGate.Items.Assign(TAppareilIEEE(ListeFrequencemetres.Items[ord(Frequencemetre)]).lstLibellesGates);

        While True Do
        Begin
            If TypeMesure = etStabilite Then
            Begin
                Self.Caption := 'Mesure de Stabilité';
                grpProcedures.Visible := True;
                If Frequencemetre = eaEIP Then
                    rgpProcAutos.ItemIndex := 0
                Else
                    rgpProcAutos.ItemIndex := 1;
                Break;
            End;

            Self.Caption := 'Mesure de Fréquence';
            grpProcedures.Visible := False;

            If (TypeMesure <> etFreqAvantInterv) Or (Frequencemetre = eaEIP) Then
                cbxValGate.ItemIndex := cbxValGate.Items.IndexOf(AT_VAL_GATES[I_GATE1S].sLibGate)
            Else
                cbxValGate.ItemIndex := cbxValGate.Items.IndexOf(AT_VAL_GATES[I_GATE10S].sLibGate);

            Break;
        End;

    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmSelectionGate.rgpProcAutosClick
  @Author:   cb le 02/09/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmSelectionGate.rgpProcAutosClick(Sender: TObject);
Begin

    cbxValGate.ItemIndex := -1;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmSelectionGate.cbxValGateChange
  @Author:   cb le 02/09/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmSelectionGate.cbxValGateChange(Sender: TObject);
Begin

    If cbxValGate.ItemIndex >= 0 Then rgpProcAutos.ItemIndex := -1;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmSelectionGate.btnCancelClick
  @Author:   cb le 10/09/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmSelectionGate.btnCancelClick(Sender: TObject);
Begin

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmSelectionGate.btnOkClick
  @Author:   cb le 02/09/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmSelectionGate.btnOkClick(Sender: TObject);
Begin

    With Self.FMesure Do
    Begin
        If cbxValGate.ItemIndex >= 0 Then
            IndexGate := cbxValGate.ItemIndex
        Else
        Begin
            If Frequencemetre = eaEIP Then
                IndexGate := I_IDX_PROCEDURE_EIP
            Else
            Begin
                If rgpProcAutos.ItemIndex = 0 Then
                    IndexGate := I_IDX_PROCEDURE_100S
                Else
                    IndexGate := I_IDX_PROCEDURE_10S;
            End;
        End;
    End;
End;

End.
