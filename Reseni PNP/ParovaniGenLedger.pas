unit ParovaniGenLedger;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Grids, AdvObj,
  BaseGrid, AdvGrid, AdvUtil;

type
  TfmSparovaniVDeniku = class(TForm)
    asgSparovaniVDeniku: TAdvStringGrid;
    btnNactiData: TButton;
    btnProvedSparovani: TButton;
    editKodUctu: TEdit;
    chbPouzeNulovyRozdil: TCheckBox;
    chbShodneProtiucty: TCheckBox;
    editLimit: TEdit;
    lblLimit: TLabel;
    Label1: TLabel;
    procedure asgSparovaniVDenikuClickCell(Sender: TObject; ARow,
      ACol: Integer);
    procedure asgSparovaniVDenikuCanSort(Sender: TObject; ACol: Integer;
      var DoSort: Boolean);
    procedure asgSparovaniVDenikuGetAlignment(Sender: TObject; ARow,
      ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure btnNactiDataClick(Sender: TObject);
    procedure btnProvedSparovaniClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure asgSparovaniVDenikuCanEditCell(Sender: TObject; ARow,
      ACol: Integer; var CanEdit: Boolean);
  private
    { Private declarations }
  public
    procedure nactiData;
    procedure provedSparovani;
  end;

var
  fmSparovaniVDeniku: TfmSparovaniVDeniku;
  asgSparovaniAllRowsChecked: boolean;

implementation

{$R *.dfm}

uses
  DesUtils;



procedure TfmSparovaniVDeniku.nactiData;
var
  SQLStr, accountId: string;
  radek, sloupec: integer;
begin
// nalezení zákazníkù s pøeplatky na 325 z Abry

  with DesU.qrAbra do begin
    SQL.Text := 'SELECT Id FROM Accounts'
              + ' WHERE Code = ''' + trim(editKodUctu.Text) + ''''
              + ' AND Hidden = ''N'' ';
    Open;
    if not Eof then begin
      accountId := FieldByName('Id').AsString;
    end;
    Close;

    asgSparovaniVDeniku.Cells[1, 1] :=  accountId;

  end;


  SQLStr := 'SELECT * FROM'
  + ' ('
  + ' SELECT F.Name, G1.Firm_ID as MD_Firm_ID, G1.Amount as MD_castka, G1.Text as MD_text, G1.ACCGROUP_ID as MD_AccGroupID,'
  + '   G1.ACCDATE$DATE as MD_datum, CAccounts.Code as MD_D_Code, G1.CREDITACCOUNT_ID as MD_CreditAccount_ID, G1.ID as MD_ID'
  + ' FROM GENERALLEDGER G1'
  + '   JOIN Firms F ON G1.Firm_ID = F.Id '
  + '   JOIN Accounts CAccounts ON G1.CREDITACCOUNT_ID = CAccounts.Id '
  + ' WHERE G1.DebitAccount_ID = ''' + accountId + ''''
  + '   AND NOT EXISTS (SELECT G2.ID FROM GENERALLEDGER G2 WHERE G2.AccGroup_ID = G1.AccGroup_ID AND G2.ID <> G1.ID)'
  + ' ) as Samostatny_MD'

  + ' JOIN'

  + ' ('
  + ' SELECT G1.Firm_ID as D_Firm_ID, G1.Amount as D_castka, G1.Text as D_text, G1.ACCGROUP_ID as D_AccGroupID,'
  + '   G1.ACCDATE$DATE as D_datum, DAccounts.Code as D_MD_Code, G1.DEBITACCOUNT_ID as D_DebitAccount_ID, G1.ID as D_ID'
  + ' FROM GENERALLEDGER G1'
  + '   JOIN Accounts DAccounts ON G1.DEBITACCOUNT_ID = DAccounts.Id '


  + ' WHERE G1.CreditAccount_ID = ''' + accountId + ''''
  + '   AND NOT EXISTS (SELECT G2.ID FROM GENERALLEDGER G2 WHERE G2.AccGroup_ID = G1.AccGroup_ID AND G2.ID <> G1.ID)'
  + ' ) as Samostatny_D'

  + ' ON MD_Firm_ID = D_Firm_ID WHERE MD_ID <> D_ID';

  if chbShodneProtiucty.Checked then
    SQLStr := SQLStr +' AND MD_CreditAccount_ID = D_DebitAccount_ID';

  if chbPouzeNulovyRozdil.Checked then
    SQLStr := SQLStr +' AND (MD_castka - D_castka) = 0';



  with DesU.qrAbra, asgSparovaniVDeniku do begin
    DesU.dbAbra.Reconnect;
    Screen.Cursor := crHourGlass;
    ClearNormalCells;
    RowCount := 2;
    //CheckFalse := '0';
    //CheckTrue := '1';
    asgSparovaniAllRowsChecked := true;

    radek := 0;
    SQL.Text := SQLStr;
    Open;
    while not EOF do begin
      Inc(radek);
      if radek > StrToInt(editLimit.Text) then Break;

      RowCount := radek + 1;
      AddCheckBox(0, radek, True, True);
      Cells[1, radek] := FieldByName('Name').AsString;
      Cells[2, radek] := DateToStr(FieldByName('MD_datum').AsFloat);
      Cells[3, radek] := FieldByName('MD_text').AsString;
      Floats[4, radek] := FieldByName('MD_castka').AsCurrency;
      Cells[5, radek] := FieldByName('MD_D_Code').AsString;

      Cells[6, radek] := DateToStr(FieldByName('D_datum').AsFloat);
      Cells[7, radek] := FieldByName('D_text').AsString;
      Floats[8, radek] := FieldByName('D_castka').AsCurrency;
      Cells[9, radek] := FieldByName('D_MD_Code').AsString;

      Floats[10, radek] := FieldByName('MD_castka').AsCurrency - FieldByName('D_castka').AsCurrency;
      Cells[11, radek] := FieldByName('MD_AccGroupID').AsString; //AccGroupID které budeme nastavovat
      Cells[12, radek] := FieldByName('D_ID').AsString; //id záznamu kterému budeme mìnit AccGroupID

      FontColors[11, radek] := $999999;
      FontColors[12, radek] := $999999;

      for sloupec := 2 to 5 do asgSparovaniVDeniku.Colors[sloupec, radek] := $FFF3E6;
      for sloupec := 6 to 9 do asgSparovaniVDeniku.Colors[sloupec, radek] := clCream;

      Application.ProcessMessages;
      Next;
    end;
    Close;

    Screen.Cursor := crDefault;

  end;
end;


procedure TfmSparovaniVDeniku.provedSparovani;
var
  SQLStr: AnsiString;
  radek: integer;
  chbstate: boolean;
begin

  with DesU.qrAbra, asgSparovaniVDeniku do begin

    for radek := FixedRows to RowCount - 1 do begin
      GetCheckBoxState(0, radek, chbstate);
      if chbstate then
      begin
        Cells[0, radek] := '...';
        RemoveCheckBox(0, radek);
        try
          SQL.Text := 'UPDATE GeneralLedger SET ACCGROUP_ID = ''' + Cells[11, radek] + ''''
                    + ' WHERE Id = ''' + Cells[12, radek] + '''';
          ExecSQL;
          Close;

          Cells[0, radek] := 'ok';
        except
          on E: Exception do
          Cells[0, radek] := 'fail';
        end;
        Application.ProcessMessages;
      end;
    end;

    DesU.dbAbra.Reconnect;

    Close;

  end;
end;

{*********************** akce Input elementù **********************************}

procedure TfmSparovaniVDeniku.FormShow(Sender: TObject);
begin
   nactiData;
end;

procedure TfmSparovaniVDeniku.btnNactiDataClick(Sender: TObject);
begin
  nactiData;
end;

procedure TfmSparovaniVDeniku.btnProvedSparovaniClick(Sender: TObject);
begin
  provedSparovani;
end;

procedure TfmSparovaniVDeniku.asgSparovaniVDenikuCanEditCell(Sender: TObject;
  ARow, ACol: Integer; var CanEdit: Boolean);
begin
  CanEdit := true;
end;

procedure TfmSparovaniVDeniku.asgSparovaniVDenikuCanSort(Sender: TObject;
  ACol: Integer; var DoSort: Boolean);
begin
  DoSort := ACol <> 0;
end;

procedure TfmSparovaniVDeniku.asgSparovaniVDenikuClickCell(Sender: TObject;
  ARow, ACol: Integer);
var
  radek: integer;
begin
  asgSparovaniAllRowsChecked := not asgSparovaniAllRowsChecked;
  if (ARow = 0) and (ACol = 0) then
    for radek := 1 to asgSparovaniVDeniku.RowCount-1 do
      asgSparovaniVDeniku.SetCheckBoxState(ACol, radek, asgSparovaniAllRowsChecked);
end;

procedure TfmSparovaniVDeniku.asgSparovaniVDenikuGetAlignment(Sender: TObject;
  ARow, ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin

  case ACol of
    0: HAlign := taCenter  ;
    2,4..6,8..10: HAlign := taRightJustify;
  end;

end;



end.
