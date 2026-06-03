Unit F_GetInfosDerive;

Interface

Uses
    Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
    Dialogs, StdCtrls, JvExStdCtrls, JvButton, JvCtrls, AdvEdit, AdvCombo,
    U_DeclarationsMETROLOGO;

Type
    TfrmGetInfosDerive = Class(TForm)
        cbxValGate: TAdvComboBox;
        edtNbMesures: TAdvEdit;
        btnOk: TJvImgBtn;
        btnCancel: TJvImgBtn;
        edtNbCycles: TAdvEdit;
        cbxIntervalCycles: TAdvComboBox;
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
    frmGetInfosDerive: TfrmGetInfosDerive;

Implementation

{$R *.dfm}

{ TfrmGetInfosDerive }

{*-------------------------------------------------------------------------------
  Procedure: TfrmGetInfosDerive.ConfigMesure
  @Author:   cb le 03/09/2009
  @Param:    const Value: TMesure
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmGetInfosDerive.ConfigMesure(Const Value: TMesure);
Begin

    FMesure := Value;
    FbAllowQuit := False;

    With FMesure Do Begin
        cbxValGate.Items.Assign(TAppareilIEEE(ListeFrequencemetres.Items[ord(Frequencemetre)]).lstLibellesGates);
        cbxValGate.ItemIndex := cbxValGate.Items.IndexOf(AT_VAL_GATES[I_GATE10S].sLibGate);

        edtNbMesures.Text := IntToStr(NbMesDerive);
        If ModeMetrologo = emValidation Then Begin
            cbxIntervalCycles.ItemIndex := I_IDX_PERIODICITE_DERIVE_VALIDATION;
            edtNbCycles.Text := IntToStr(I_NBCYCLESDERIVE_VALIDATION);
        End
        Else Begin
            cbxIntervalCycles.ItemIndex := IndexIntervalDerive;
            edtNbCycles.Text := IntToStr(NbCyclesDerive);
        End;
    End;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmGetInfosDerive.btnCancelClick
  @Author:   cb le 03/09/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmGetInfosDerive.btnCancelClick(Sender: TObject);
Begin

    FbAllowQuit := True;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmGetInfosDerive.btnOkClick
  @Author:   cb le 03/09/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmGetInfosDerive.btnOkClick(Sender: TObject);
Var
    blErreur, blDispErr: Boolean;
    ilValnbMes, ilValNbCycles: Integer;
    slVal: String;

Begin

    blErreur := True;
    blDispErr := False;

    While True Do
    Begin
        If cbxValGate.ItemIndex < 0 Then
        Begin
            cbxValGate.SetFocus;
            Application.MessageBox('Il faut sélectionner un temps de mesure !', 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
            Break;
        End;

        If cbxIntervalCycles.ItemIndex < 0 Then
        Begin
            cbxIntervalCycles.SetFocus;
            Application.MessageBox('Il faut sélectionner un intervalle de cycle !', 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
            Break;
        End;

        blDispErr := True;

        edtNbCycles.SetFocus;                                                   // Teste Nb cycles
        slVal := edtNbCycles.Text;
        If Not TryStrToInt(slVal, ilValNbCycles) Then break;
        If ilValNbCycles < 1 Then Break;

        edtNbMesures.SetFocus;                                                  // Teste Nb mesures
        slVal := edtNbMesures.Text;
        If Not TryStrToInt(slVal, ilValnbMes) Then break;
        If ilValNbMes < 1 Then Break;

        blErreur := False;
        Break;
    End;

    If blErreur Then
    Begin
        If blDispErr Then
            Application.MessageBox(PChar(Format('Nombre incorrect : %s', [slVal])), 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
    End
    Else
    Begin
        With FMesure Do
        Begin
//            IndexGateDerive := cbxValGate.ItemIndex;
            IndexGate := cbxValGate.ItemIndex;
            IndexIntervalDerive := cbxIntervalCycles.ItemIndex;
            NbMesDerive := ilValnbMes;
            NbCyclesDerive := ilValNbCycles;
        End;
    End;

    FbAllowQuit := Not blErreur;

End;

End.
