unit NZLtisk;

interface

uses
  Windows, Messages, Forms, Controls, Classes, Dialogs, SysUtils, DateUtils, Printers,
  DesUtils;

type
  TdmTisk = class(TDataModule)
  public
    function FO3Tisk(InvoiceId, Fr3FileName: string) : TDesResult;
    procedure TiskniFO3;
  end;

var
  dmTisk: TdmTisk;

implementation

{$R *.dfm}

uses DesInvoices, NZLmain, NZLcommon;

// ------------------------------------------------------------------------------------------------

procedure TdmTisk.TiskniFO3;
// použije data z asgMain
var
  Radek: integer;
  Faktura : TDesInvoice;
  VysledekPrevedeni : TDesResult;
begin
  with fmMain do try
    fmMain.Zprava('Tisk FO3');
    with asgMain do begin
      fmMain.Zprava(Format('Poèet FO3 k tisku: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));
      Screen.Cursor := crHourGlass;
      apnTisk.Visible := False;
      lbPozor1.Visible := False;
      apbProgress.Position := 0;
      apbProgress.Visible := True;
// hlavní smyèka
      for Radek := 1 to RowCount-1 do begin
        Row := Radek;
        apbProgress.Position := Round(100 * Radek / RowCount-1);
        Application.ProcessMessages;
        if Prerusit then begin
          Prerusit := False;
          apbProgress.Position := 0;
          apbProgress.Visible := False;
          lbPozor1.Visible := True;
          btVytvorit.Enabled := True;
          asgMain.Visible := True;
          lbxLog.Visible := False;
          Break;
        end;

        if Ints[0, Radek] = 1 then begin  // pokud zaškrtnuto, pøevádíme fa do PDF


          //VysledekPrevedeni := FO3Tisk(asgMain.Cells[10, Radek], 'FO3doPDF.fr3');

          Faktura := TDesInvoice.create(asgMain.Cells[10, Radek]);
          Faktura.reportAry['CisloZL'] := asgMain.Cells[1, Radek];
          Faktura.reportAry['CastkaZalohy'] := asgMain.Floats[3, Radek];

          VysledekPrevedeni := Faktura.printByFr3('FO3doPDF-104.fr3');

          fmMain.Zprava(Format('%s (%s): %s', [asgMain.Cells[6, Radek], asgMain.Cells[7, Radek], VysledekPrevedeni.Messg]));

          if VysledekPrevedeni.isOk then begin
            asgMain.Ints[0, Radek] := 0;
          end
          else
          begin
            if Dialogs.MessageDlg( 'Chyba pøi tisku: '
              + VysledekPrevedeni.Messg + sLineBreak + 'Pokraèovat?',
              mtConfirmation, [mbYes, mbNo], 0 ) = mrNo then Prerusit := True;
          end;
        end;
      end;  // for
    end;
// konec hlavní smyèky
  finally
    Printer.PrinterIndex := -1;
    apbProgress.Position := 0;
    apbProgress.Visible := False;
    lbPozor1.Visible := True;
    apnTisk.Visible := True;
    asgMain.Visible := False;
    lbxLog.Visible := True;
    Screen.Cursor := crDefault;                                             // default
    fmMain.Zprava('Tisk FO3 ukonèen');
  end;
end;

// ------------------------------------------------------------------------------------------------

function TdmTisk.FO3Tisk(InvoiceId, Fr3FileName: string) : TDesResult;
var
  Faktura : TDesInvoice;

begin
  Faktura := TDesInvoice.create(InvoiceId);
  Result := Faktura.printByFr3(Fr3FileName);
end;

{
procedure TdmTisk.FO3Tisk(Radek: integer);
var
  AbraKod,
  SQLStr: AnsiString;
  Zaplaceno: double;
begin
  with fmMain, asgMain do begin
    with qrAbra do begin
// údaje z faktury do privátních promìnných
      Close;
      SQLStr := 'SELECT F.Code, VarSymbol, Street, City, PostCode, OrgIdentNumber, VATIdentNumber, LocalAmount, LocalPaidAmount,'
       + ' DocDate$DATE, VATDate$DATE, DueDate$DATE'
      + ' FROM Firms F, Addresses A, IssuedDInvoices II'
      + ' WHERE F.ID = II.Firm_ID'
      + ' AND A.ID = F.ResidenceAddress_ID'
      + ' AND II.ID = ' + Ap + Cells[10, Radek] + Ap;
      SQL.Text := SQLStr;
      Open;
      FO_Id := Cells[10, Radek];
      Cislo := Cells[7, Radek];
      CastkaZalohy := Floats[3, Radek];
      CisloZL := Cells[1, Radek];
      AbraKod := FieldByName('Code').AsString;
      VS := FieldByName('VarSymbol').AsString;
      OJmeno := Cells[6, Radek];
      OUlice := UTF8Decode(FieldByName('Street').AsString);
      OObec := UTF8Decode(FieldByName('PostCode').AsString + ' ' + FieldByName('City').AsString);
      OICO := FieldByName('OrgIdentNumber').AsString;
      ODIC := FieldByName('VATIdentNumber').AsString;
      DatumDokladu := FieldByName('DocDate$DATE').AsFloat;
      DatumPlneni := FieldByName('VATDate$DATE').AsFloat;
      DatumSplatnosti := FieldByName('DueDate$DATE').AsFloat;
      Vystaveni := FormatDateTime('dd.mm.yyyy', DatumDokladu);
      Plneni := FormatDateTime('dd.mm.yyyy', DatumPlneni);
      Splatnost := FormatDateTime('dd.mm.yyyy', DatumSplatnosti);
      Celkem := FieldByName('LocalAmount').AsFloat;
      Zaplaceno := FieldByName('LocalPaidAmount').AsFloat;
      Close;
// všechny Firm_Id pro Abrakód firmy
      SQLStr := 'SELECT * FROM DE$_CODE_TO_FIRM_ID (' + Ap + AbraKod + ApZ;
      SQL.Text := SQLStr;
      Open;
      if Fields[0].AsString = 'MULTIPLE' then begin
        fmMain.Zprava(Format('%s (%s): Více zákazníkù pro kód %s.', [OJmeno, VS, AbraKod]));
        Close;
        Exit;
      end;
      Saldo := 0;
  // a saldo pro všechny Firm_Id
      while not EOF do with qrAdresa do begin
        Close;
        SQLStr := 'SELECT SaldoPo + SaldoZLPo + Ucet325 FROM DE$_Firm_Totals (' + Ap + qrAbra.Fields[0].AsString + ApC + FloatToStr(DatumDokladu) + ')';
        SQL.Text := SQLStr;
        Open;
        Saldo := Saldo + Fields[0].AsFloat;
        qrAbra.Next;
      end; // while not EOF do with qrAdresa
    end;  // with qrAbra
    Celkem := Celkem - CastkaZalohy;
    Saldo := Saldo + Zaplaceno - CastkaZalohy;
    Zaplatit := Celkem - Saldo;
// text na fakturu
    if Saldo > 0 then Platek := 'pøeplatek'
    else if Saldo < 0 then Platek := 'nedoplatek'
    else Platek := ' ';
    if Zaplatit < 0 then Zaplatit := 0;
// údaje z tabulky Smlouvy do globálních promìnných
    with qrMain do begin
      Close;
      SQLStr := 'SELECT Postal_name, Postal_street, Postal_PSC, Postal_city FROM customers'
      + ' WHERE Variable_symbol = ' + Ap + VS + Ap;
      SQL.Text := SQLStr;
      Open;
      PJmeno := FieldByName('Postal_name').AsString;
      PUlice := FieldByName('Postal_street').AsString;
      PObec := FieldByName('Postal_PSC').AsString + ' ' + FieldByName('Postal_city').AsString;
      Close;
    end;  // with qrMain
// zasílací adresa
    if (PJmeno = '') or (PObec = '') then begin
      PJmeno := OJmeno;
      PUlice := OUlice;
      PObec := OObec;
    end;
    try
      frxReport.LoadFromFile(ExtractFilePath(ParamStr(0)) + 'FR3\FO3doPDF.fr3');
      frxReport.PrepareReport;
//      frxReport.ShowPreparedReport;
      frxReport.PrintOptions.ShowDialog := False;
      frxReport.Print;
      fmMain.Zprava(Format('%s (%s): %s byla odeslána na tiskárnu.', [OJmeno, VS, Cells[7, Radek]]));
      Ints[0, Radek] := 0;
    except on E: exception do
      begin
        fmMain.Zprava(Format('%s (%s): %s se nepodaøilo vytisknout.' + #13#10 + 'Chyba: %s',
         [OJmeno, VS, Cells[7, Radek], E.Message]));
        if Application.MessageBox(PChar('Chyba pøi tisku' + ^M + E.Message), 'Pokraèovat?',
         MB_YESNO + MB_ICONQUESTION) = IDNO then Prerusit := True;
      end;
    end;  // try
  end;  // with fmMain
end;
}

end.

