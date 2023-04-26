unit Customers;
// 12.5.2017 zobrazí nìkteré údaje z tabulek "customers" a "contracts" podle zadaných kritérií

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, StdCtrls, Forms, Math, IniFiles, Dialogs, Grids, BaseGrid, AdvGrid, AdvObj,
  DB, ZAbstractRODataset, ZAbstractDataset, ZDataset, ZAbstractConnection, ZConnection, VypisyMain, AdvUtil;

type
  TfmCustomers = class(TForm)
    lbJmeno: TLabel;
    edJmeno: TEdit;
    lbPrijmeni: TLabel;
    edPrijmeni: TEdit;
    lbVS: TLabel;
    edVS: TEdit;
    btNajdi: TButton;
    asgCustomers: TAdvStringGrid;
    chbJenSNezaplacenym: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure btNajdiClick(Sender: TObject);
    procedure asgCustomersGetAlignment(Sender: TObject; ARow, ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure edJmenoKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure edPrijmeniKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure edVSKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
  end;

var
  fmCustomers: TfmCustomers;

implementation
uses
  AbraEntities, DesUtils, Superobject;
{$R *.dfm}

procedure TfmCustomers.FormShow(Sender: TObject);
begin
  with fmMain.asgMain do
    if (Cells[2, Row] <> '') then edVS.Text := Cells[2, Row];
  edVS.SetFocus;
end;

procedure TfmCustomers.edJmenoKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Key = 13) then btNajdiClick(Self);
end;

procedure TfmCustomers.edPrijmeniKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Key = 13) then btNajdiClick(Self);
end;

procedure TfmCustomers.edVSKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Key = 13) then btNajdiClick(Self);
end;

procedure TfmCustomers.asgCustomersGetAlignment(Sender: TObject; ARow, ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
  HAlign := taCenter;
  if (ACol = 0) and (ARow > 0) then HAlign := taLeftJustify;
end;

procedure TfmCustomers.btNajdiClick(Sender: TObject);
// jednoduché hledání z tabulek customers a contracts v databázi aplikace
var
  Zakaznik,
  PosledniZakaznik,
  SQLStr: AnsiString;
  SirkaOkna,
  Radek: integer;
  nezaplacenyDoklad : TDoklad;
begin
  SQLStr := 'SELECT DISTINCT Cu.Abra_code, Cu.Variable_symbol, Cu.Title, Cu.First_name, Cu.Surname, Cu.Business_name,'
  + ' C.Number, C.State, C.Invoice, C.Invoice_from, C.Canceled_at'
  + ' FROM customers Cu, contracts C'
  + ' WHERE Cu.Id = C.Customer_Id'
  + ' AND Cu.First_name LIKE ''' + edJmeno.Text + ''''
  + ' AND (Cu.Business_name LIKE ''' + edPrijmeni.Text + ''''
  + ' OR Cu.Surname LIKE ''' + edPrijmeni.Text + ''')'
  + ' AND (Cu.Variable_symbol LIKE ''' + edVS.Text + ''''
  + ' OR C.Number LIKE ''' + edVS.Text + ''')'
  + ' ORDER BY Business_name, Surname, First_name';
  SQLStr := StringReplace(SQLStr, '*', '%', [rfReplaceAll]);      // aby fungovala i *

  with DesU.qrZakos, asgCustomers do try
    Screen.Cursor := crSQLWait;
    ClearNormalCells;
    RowCount := 2;
    Radek := 0;
    PosledniZakaznik := '';
    SQL.Text := SQLStr;
    Open;
    while not EOF do begin

      DesU.qrAbra.SQL.Text :=
            'SELECT ID, DocumentType, DocDate$Date FROM ('
          + 'SELECT ID, ''03'' as DocumentType, DocDate$Date FROM ISSUEDINVOICES'
          + ' WHERE VarSymbol = ''' + FieldByName('Number').AsString + ''''
          + ' AND (LOCALAMOUNT - LOCALPAIDAMOUNT - LOCALCREDITAMOUNT + LOCALPAIDCREDITAMOUNT) <> 0 '
          + 'UNION '
          + 'SELECT ID, ''10'' as DocumentType, DocDate$Date FROM ISSUEDDINVOICES'
          + ' WHERE VarSymbol = ''' + FieldByName('Number').AsString + ''''
          + ' AND (LOCALAMOUNT - LOCALPAIDAMOUNT) <> 0'
          + ') ORDER BY DocDate$Date';

      DesU.qrAbra.Open; //najdu IDèka nezaplacených fa a ZL

      if not chbJenSNezaplacenym.Checked OR not DesU.qrAbra.Eof then begin
        Inc(Radek);
        RowCount := Radek + 1;
        Zakaznik := FieldByName('Surname').AsString;
        if (FieldByName('First_name').AsString <> '') then Zakaznik := FieldByName('First_name').AsString + ' ' + Zakaznik;
        if (FieldByName('Title').AsString <> '') then Zakaznik := FieldByName('Title').AsString + ' ' + Zakaznik;
        if (FieldByName('Business_name').AsString <> '') then Zakaznik := FieldByName('Business_name').AsString + ', ' + Zakaznik;
        if (Zakaznik <> PosledniZakaznik) then begin      // pro stejného zákazníka se vypisují jen údaje o smlouvì
          Cells[0, Radek] := Zakaznik;
          Cells[1, Radek] := FieldByName('Abra_code').AsString;
          Cells[2, Radek] := FieldByName('Variable_symbol').AsString;
          PosledniZakaznik := Zakaznik;
        end;
        Cells[3, Radek] := FieldByName('Number').AsString;
        Cells[4, Radek] := FieldByName('State').AsString;
        if (FieldByName('Invoice').AsInteger = 1) then Cells[5, Radek] := 'ano'
        else Cells[5, Radek] := 'ne';
        if (Cells[5, Radek] = 'ano') and (FieldByName('Invoice_from').AsString <> '') then
          Cells[5, Radek] := Cells[5, Radek] + ', od ' + FieldByName('Invoice_from').AsString;
        if (Cells[5, Radek] = 'ne') and (FieldByName('Canceled_at').AsString <> '') then
          Cells[5, Radek] := Cells[5, Radek] + ', do ' + FieldByName('Canceled_at').AsString;

        while not DesU.qrAbra.Eof do begin
          Inc(Radek);
          RowCount := Radek + 1;
          nezaplacenyDoklad := TDoklad.create(DesU.qrAbra.FieldByName('ID').AsString, DesU.qrAbra.FieldByName('DocumentType').AsString); //vytvoøím objekty nezaplacených fa a ZL
          Cells[6, Radek] := nezaplacenyDoklad.cisloDokladu;
          Cells[7, Radek] := DateToStr(nezaplacenyDoklad.datumDokladu);
          Cells[8, Radek] := format('%m', [nezaplacenyDoklad.CastkaNezaplaceno]);
          DesU.qrAbra.Next;
        end;

        Row := Radek;
        Application.ProcessMessages;
      end;

      DesU.qrAbra.Close;
      Next;
    end;
  finally
    Row := 1;
    AutoSize := True;
    SirkaOkna := 0;
    for Radek := 0 to ColCount-1 do SirkaOkna := SirkaOkna + ColWidths[Radek];
    fmCustomers.ClientWidth := Max(825, SirkaOkna + 4);
    //fmCustomers.Left := Max(0, Round((Screen.Width - fmCustomers.Width) / 2));
    fmCustomers.ClientHeight := Min(RowCount * 18 + 46, Screen.Height - 30);
    if (fmCustomers.ClientHeight = Screen.Height - 30) then fmCustomers.ClientWidth := fmCustomers.ClientWidth + 26;
    //fmCustomers.Top := Max(0, Round((Screen.Height - fmCustomers.Height) / 2));
    Screen.Cursor := crDefault;
  end;
end;

end.
