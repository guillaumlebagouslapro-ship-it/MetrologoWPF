Unit F_ChoixRubidium;

Interface

Uses
    Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
    Dialogs, DB, ADODB, StdCtrls, ExtCtrls, JvExStdCtrls, JvButton, JvCtrls,
    U_DeclarationsMETROLOGO, Math;

Type
    TfrmChoixRubidium = Class(TForm)
        rgbRubidiums: TRadioGroup;
        rgbRaccord: TRadioGroup;
        dstRubidiums: TADODataSet;
        btnOk: TJvImgBtn;
        btnCancel: TJvImgBtn;
        Procedure FormShow(Sender: TObject);
        Procedure btnOkClick(Sender: TObject);
        Procedure FormCreate(Sender: TObject);
        Procedure btnCancelClick(Sender: TObject);
    Private
        bAllowQuit: Boolean;
        FAvecGPS: Boolean;
        FIDRubidium: Integer;
        FNomRubidium: String;

    Public
        atRubidiums: Array Of TRubidiums;
        Property iSelRubIDRubidium: Integer Read FIDRubidium Write FIDRubidium;
        Property bSelRubAvecGPS: Boolean Read FAvecGPS Write FAvecGPS;
        property NomRubidium: String read FNomRubidium write FNomRubidium;

    End;

Var
    frmChoixRubidium: TfrmChoixRubidium;

Implementation

{$R *.dfm}

{*-------------------------------------------------------------------------------
  Procedure: TfrmChoixRubidium.FormShow
  @Author:   cb le 08/10/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmChoixRubidium.FormShow(Sender: TObject);
Var
    ilCnt, ilIndexRubi, ilIndexGPS: Integer;

Begin

    ilIndexRubi := -1;
    ilIndexGPS := -1;
    With dstRubidiums Do Begin
        Open;
        If RecordCount > 0 Then Begin
            SetLength(Self.atRubidiums, RecordCount);
            For ilCnt := 0 To RecordCount - 1 Do Begin
                atRubidiums[ilCnt].ID := Fieldbyname(S_RUB_ID).asInteger;
                atRubidiums[ilCnt].Design := Fieldbyname(S_RUB_DESIGNATION).AsString;
                rgbRubidiums.Items.Add(atRubidiums[ilCnt].Design);
                If FieldByName(S_RUB_ACTIF).AsBoolean Then Begin
                    ilIndexRubi := ilCnt;
                    ilIndexGPS := ifthen(FieldByName(S_RUB_AVECGPS).asboolean, 1, 0);
                End;
                Next;
            End;
        End;
        Close;
    End;

    rgbRubidiums.ItemIndex := ilIndexRubi;
    rgbRaccord.ItemIndex := ilIndexGPS;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmChoixRubidium.btnOkClick
  @Author:   cb le 08/10/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmChoixRubidium.btnOkClick(Sender: TObject);
Begin

    Self.bAllowQuit := False;

    If (rgbRubidiums.ItemIndex = -1) Or (rgbRaccord.ItemIndex = -1) Then Begin
        Application.MessageBox('Il faut impérativement spécifier le Rubidium et le mode de raccordement !',
            'Attention', MB_OK + MB_ICONWARNING + MB_TOPMOST);
        Exit;
    End;

    FNomRubidium := atRubidiums[rgbRubidiums.itemindex].Design;
    FIDRubidium := atRubidiums[rgbRubidiums.itemindex].ID;
    FAvecGPS := (rgbRaccord.ItemIndex > 0);

    Self.bAllowQuit := True;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmChoixRubidium.FormCreate
  @Author:   cb le 08/10/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmChoixRubidium.FormCreate(Sender: TObject);
Begin

    Self.bAllowQuit := True;

End;

{*-------------------------------------------------------------------------------
  Procedure: TfrmChoixRubidium.btnCancelClick
  @Author:   cb le 13/10/2009
  @Param:    Sender: TObject
  @Result:   None
-------------------------------------------------------------------------------}
Procedure TfrmChoixRubidium.btnCancelClick(Sender: TObject);
Begin

    Self.bAllowQuit := True;

End;

End.

