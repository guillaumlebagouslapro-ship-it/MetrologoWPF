Unit F_DateDebutGraphe;

Interface

Uses
    Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
    Dialogs, StdCtrls, ComCtrls, JvExComCtrls, JvDateTimePicker,
    JvExStdCtrls, JvButton, JvCtrls;

Type
    TfrmGetDatesGrapheDeriveRubidium = Class(TForm)
        dtpDateDebut: TJvDateTimePicker;
        lblChoixdate: TLabel;
        btnOk: TJvImgBtn;
        btnCancel: TJvImgBtn;
        dtpDateFin: TJvDateTimePicker;
        lbl2: TLabel;
        Procedure FormCreate(Sender: TObject);
        Procedure dtpDateFinChange(Sender: TObject);
        Procedure dtpDateDebutChange(Sender: TObject);
    procedure btnOkClick(Sender: TObject);
    Private
        FDateDebut: Integer;
        FDateFin: Integer;

    Public
        Property DateDebut: Integer Read FDateDebut Write FDateDebut;
        Property DateFin: Integer Read FDateFin Write FDateFin;
    End;

Var
    frmGetDatesGrapheDeriveRubidium: TfrmGetDatesGrapheDeriveRubidium;

Implementation

Uses DateUtils;

{$R *.dfm}

{*-------------------------------------------------------------------------------
  Procedure: TfrmGetDateDebutGrapheDeriveRubidium.FormCreate
  @Author:   cb le 13/10/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmGetDatesGrapheDeriveRubidium.FormCreate(Sender: TObject);
Begin

    dtpDateFin.DateTime := Now;
    dtpDateDebut.DateTime := IncWeek(dtpDateFin.DateTime, -3);
End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmGetDateDebutGrapheDeriveRubidium.dtpDateFinChange
  @Author:   cb le 13/10/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmGetDatesGrapheDeriveRubidium.dtpDateFinChange(Sender: TObject);
Var
    dlLimite: TDateTime;

Begin

    dlLimite := IncWeek(dtpDateDebut.DateTime, -3);
    If dtpDateDebut.DateTime <= dlLimite Then Exit;
    dtpDateDebut.DateTime := dlLimite;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmGetDateDebutGrapheDeriveRubidium.dtpDateDebutChange
  @Author:   cb le 13/10/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmGetDatesGrapheDeriveRubidium.dtpDateDebutChange(Sender: TObject);
Var
    dlLimite: TDateTime;

Begin

    dlLimite := IncWeek(dtpDateDebut.DateTime, 3);

    If dtpDateFin.DateTime >= dlLimite Then Exit;
    dtpDateFin.DateTime := dlLimite;

End;


{*-------------------------------------------------------------------------------
  Procedure: TfrmGetDateDebutGrapheDeriveRubidium.btnOkClick
  @Author:   cb le 13/10/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
procedure TfrmGetDatesGrapheDeriveRubidium.btnOkClick(Sender: TObject);
begin

    FDateFin := Trunc(DateTimeToModifiedJulianDate(dtpDateFin.DateTime));
    FDateDebut := Trunc(DateTimeToModifiedJulianDate(dtpDateDebut.DateTime));

end;

End.

