Unit F_SaisieMarcheHebdo;

Interface

Uses
    Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
    Dialogs, StdCtrls, JvExStdCtrls, JvButton, JvCtrls, ActnList, ComCtrls,
    JvExComCtrls, JvDateTimePicker, ExtCtrls, AdvPanel, E2MPanel, AdvEdit,
    DateUtils;

Type
    TfrmSaisieMarcheHebdo = Class(TForm)
        btnOk: TJvImgBtn;
        btnCancel: TJvImgBtn;
    edtMarche: TAdvEdit;
        pnlDates: TE2MPanel;
        dtpDate: TJvDateTimePicker;
        lblDateMarche: TLabel;
        edtJulian: TAdvEdit;
        actConDates: TActionList;
        actConvert: TAction;
        edtNumSemaine: TAdvEdit;
        lblRubActif: TLabel;
        Procedure FormCreate(Sender: TObject);
        Procedure btnOkClick(Sender: TObject);
        Procedure btnCancelClick(Sender: TObject);
        Procedure actConvertExecute(Sender: TObject);
        Procedure FormCloseQuery(Sender: TObject; Var CanClose: Boolean);
    procedure FormShow(Sender: TObject);
        
    Private
        bAllowQuit: Boolean;
        FValeurMarche: Double;
        FJulianDate: Integer;

    Public
        property ValeurMarche: Double read FValeurMarche write FValeurMarche;
        property JulianDate: Integer read FJulianDate write FJulianDate;

    End;

Var
    frmSaisieMarcheHebdo: TfrmSaisieMarcheHebdo;

Implementation

{$R *.dfm}

{*-------------------------------------------------------------------------------
  Procedure: TfrmSaisieMarcheHebdo.FormCreate
  @Author:   cb le 13/10/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmSaisieMarcheHebdo.FormCreate(Sender: TObject);
Begin

    Self.bAllowQuit := True;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmSaisieMarcheHebdo.btnOkClick
  @Author:   cb le 13/10/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmSaisieMarcheHebdo.btnOkClick(Sender: TObject);
var
    dlVal: Double;

Begin

    Self.bAllowQuit := False;
    if TryStrToFloat(edtMarche.Text, dlVal) then begin

        if Application.MessageBox(PChar(Format('Confirmez-vous que la valeur de la marche est %s µs ?', [edtMarche.Text])), 'Confirmation', MB_YESNO + MB_ICONQUESTION + MB_TOPMOST) = IDNO then
        begin
            edtMarche.SetFocus;
            Exit;
        end;

        Self.FValeurMarche := dlVal * 1e-6;
        Self.FJulianDate := StrToInt(edtJulian.Text);
        Self.bAllowQuit := True;
    end
    else begin
        Application.MessageBox(PChar(Format('Valeur %s incorrecte !', [edtMarche.Text])), 'Attention', MB_OK + MB_ICONWARNING + MB_TOPMOST);
        edtMarche.SetFocus;
    end;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmSaisieMarcheHebdo.btnCancelClick
  @Author:   cb le 13/10/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmSaisieMarcheHebdo.btnCancelClick(Sender: TObject);
Begin

    Self.bAllowQuit := True;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmSaisieMarcheHebdo.actConvertExecute
  @Author:   cb le 13/10/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmSaisieMarcheHebdo.actConvertExecute(Sender: TObject);
Begin

    edtJulian.Text := Format('%d', [Trunc(DateTimeToModifiedJulianDate(dtpDate.DateTime))]);
    edtNumSemaine.Text := Format('%d', [WeekOf(dtpDate.DateTime)]);

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmSaisieMarcheHebdo.FormCloseQuery
  @Author:   cb le 13/10/2009
  @Param:    Sender: TObject; var CanClose: Boolean
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmSaisieMarcheHebdo.FormCloseQuery(Sender: TObject; Var CanClose: Boolean);
Begin

    CanClose := Self.bAllowQuit;

End;


{*-------------------------------------------------------------------------------
  Procedure: TfrmSaisieMarcheHebdo.FormShow
  @Author:   cb le 13/10/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
procedure TfrmSaisieMarcheHebdo.FormShow(Sender: TObject);
begin

    dtpDate.DateTime := Now;
    actConvert.Execute;

end;

End.

