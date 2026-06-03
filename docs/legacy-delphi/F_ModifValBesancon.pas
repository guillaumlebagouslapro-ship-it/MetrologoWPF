Unit F_ModifValBesancon;

Interface

Uses
    Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
    Dialogs, StdCtrls, AdvEdit, AdvEdBtn, ComCtrls, JvExComCtrls,
    JvDateTimePicker, HTMLabel, JvExStdCtrls, JvButton, JvCtrls, Mask, ADODB;

Type
    TfrmModifValBesancon = Class(TForm)
        lblInfos: THTMLabel;
        dtpDate: TJvDateTimePicker;
        edtDateJulienne: TAdvEditBtn;
        lblDate: TLabel;
        lblDateJul: TLabel;
        btnOk: TJvImgBtn;
        btnCancel: TJvImgBtn;
        edtValJour: TAdvEdit;
        Procedure dtpDateChange(Sender: TObject);
        procedure edtDateJulienneClickBtn(Sender: TObject);
        procedure edtDateJulienneExit(Sender: TObject);
        procedure btnOkClick(Sender: TObject);
        procedure FormClose(Sender: TObject; var Action: TCloseAction);

    Private
        FDateToChange: TDateTime;
        FDateJulienne: Integer;
        bModifByUser: Boolean;
        FValJour: double;
        FCnx: TADOConnection;

        Procedure SetDateToChange(Const Value: TDateTime);
        procedure setDateJulienne(const Value: Integer);

    Public
        constructor Create(AOwner: TComponent; opCnx: TADOConnection); reintroduce;
        Property DateJulienne: Integer Read FDateJulienne Write setDateJulienne;
        Property DateToChange: TDateTime Read FDateToChange Write SetDateToChange;
        property ValJour: double read FValJour write FValJour;
        property Cnx: TADOConnection read FCnx write FCnx;
        
    End;

Var
    frmModifValBesancon: TfrmModifValBesancon;

Implementation

Uses
    DateUtils;

{$R *.dfm}

Procedure TfrmModifValBesancon.dtpDateChange(Sender: TObject);
Var
    ilDateJulienne: Integer;

Begin

    bModifByUser := True;
    DateToChange := dtpDate.Date;
    bModifByUser := False;

End;


Procedure TfrmModifValBesancon.SetDateToChange(Const Value: TDateTime);
var
    olQry: TADOQuery;

Begin

    FDateToChange := Value;
    if not bModifByUser then            // Lors du FormCreate
        dtpDate.Date := FDateToChange;
    DateJulienne := Trunc(DateTimeToModifiedJulianDate(FDateToChange));

    olQry := TADOQuery.Create(Self);
    With olQry do begin
        edtValJour.Text := '';
        Connection := Cnx;
        SQL.Clear;
        SQL.Add(Format('select DAT_VALEUR from T_METROLOGO_DATESRUBIS where DAT_ID=%d and RUB_ACTIF=1', [FDateJulienne]));
        Open();
        if Recordcount > 0 then
            edtValJour.Text := FloatToStr(FieldByName('DAT_VALEUR').AsFloat);
        Close;
        Free;
    end;

End;

constructor TfrmModifValBesancon.Create(AOwner: TComponent; opCnx: TADOConnection);
begin

   inherited Create(AOwner);
   Cnx := opCnx;

   btnOk.ModalResult := mrNone;
   bModifByUser := False;
   DateToChange := Now;

end;


procedure TfrmModifValBesancon.setDateJulienne(const Value: Integer);
begin

    FDateJulienne := Value;
    edtDateJulienne.Text := IntToStr(FDateJulienne);

end;


procedure TfrmModifValBesancon.edtDateJulienneClickBtn(Sender: TObject);
begin

    DateToChange := ModifiedJulianDateToDateTime(StrToFloat(edtDateJulienne.Text));

end;


procedure TfrmModifValBesancon.edtDateJulienneExit(Sender: TObject);
begin

    edtDateJulienneClickBtn(Sender);

end;


procedure TfrmModifValBesancon.btnOkClick(Sender: TObject);
begin

    if not TryStrToFloat(edtValJour.Text, FValJour) then begin
        Application.MessageBox('Valeur incorrecte !', 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
        edtValJour.SetFocus;
        Exit;
    end;

    if abs(FValJour) > 1e-9 then begin
        Application.MessageBox(PChar(Format('La valeur saisie est trop élevée (%8.4g) !',
           [fValJour])), 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
        edtValJour.SetFocus;
        Exit;
    end;

    if (DateJulienne < 50000) or (DateJulienne > 80000) then begin          // Teste si date julienne est une date ''acceptable''
        Application.MessageBox('Date julienne incorrecte !', 'Erreur', MB_OK + MB_ICONSTOP + MB_TOPMOST);
        edtDateJulienne.SetFocus;
        Exit;
    end;

    ModalResult := mrOk;

end;


procedure TfrmModifValBesancon.FormClose(Sender: TObject; var Action: TCloseAction);
begin

    Action := caHide;

end;

End.

