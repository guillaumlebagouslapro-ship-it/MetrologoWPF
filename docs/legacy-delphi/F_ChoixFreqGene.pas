Unit F_ChoixFreqGene;

Interface

Uses
    Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
    Dialogs, StdCtrls, JvExStdCtrls, JvButton, JvCtrls;

Type
    TfrmChoixFreqGene = Class(TForm)
        btnOk: TJvImgBtn;
        btnCancel: TJvImgBtn;
        grpChoix: TGroupBox;
        optFreq: TRadioButton;
        optGene: TRadioButton;
        Procedure optFreqClick(Sender: TObject);
        Procedure optGeneClick(Sender: TObject);
    Private
        FChoixFreq: Boolean;
        Procedure SetChoixFreq(Const Value: Boolean);

    Public
        Property bChoixFreq: Boolean Read FChoixFreq Write SetChoixFreq;

    End;

Var
    frmChoixFreqGene: TfrmChoixFreqGene;

Implementation

{$R *.dfm}

{ TForm2 }

{*-------------------------------------------------------------------------------
  Procedure: TForm2.SetChoixFreq
  @Author:   cb le 25/08/2009
  @Param:    const Value: Boolean
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmChoixFreqGene.SetChoixFreq(Const Value: Boolean);
Begin

    FChoixFreq := Value;
    If Value Then
        optFreq.Checked := True
    Else
        optGene.Checked := True;

End;

{*-------------------------------------------------------------------------------
  Procedure: TForm2.optFreqClick
  @Author:   cb le 25/08/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmChoixFreqGene.optFreqClick(Sender: TObject);
Begin

    FChoixFreq := True;

End;

{*-------------------------------------------------------------------------------
  Procedure: TForm2.optGeneClick
  @Author:   cb le 25/08/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmChoixFreqGene.optGeneClick(Sender: TObject);
Begin

    FChoixFreq := False;

End;

End.

