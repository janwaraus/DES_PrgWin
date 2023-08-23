unit FOx_tisk;

interface

uses
  Windows, Messages, Forms, Controls, Classes, Dialogs, SysUtils, DateUtils, Printers, FOx_main;

type
  TdmTisk = class(TDataModule)
    dlgTisk: TPrintDialog;
  private
    procedure FOxTisk(Radek: integer);
  public
    procedure TiskniFOx;
  end;

var
  dmTisk: TdmTisk;

implementation

{$R *.dfm}

uses FOx_common;

// ------------------------------------------------------------------------------------------------

procedure TdmTisk.TiskniFOx;
// použije data z asgMain
var
  Radek: integer;
begin
  with fmMain do try
    dmCommon.Zprava('Tisk dokladù');
    with asgMain do begin
      dmCommon.Zprava(Format('Poèet dokladù k tisku: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));
      if ColumnSum(0, 1, RowCount-1) >= 1 then
        if dlgTisk.Execute then
          frxReport.PrintOptions.Printer := Printer.Printers[Printer.PrinterIndex];
      Screen.Cursor := crHourGlass;
      apnTisk.Visible := False;
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
          btVytvorit.Enabled := True;
          asgMain.Visible := True;
          lbxLog.Visible := False;
          Break;
        end;
        if Cells[0, Radek] = '1' then FOxTisk(Radek);
      end;  // for
    end;
// konec hlavní smyèky
  finally
    Printer.PrinterIndex := -1;
    apbProgress.Position := 0;
    apbProgress.Visible := False;
    apnTisk.Visible := True;
    asgMain.Visible := False;
    lbxLog.Visible := True;
    Screen.Cursor := crDefault;                                             // default
    dmCommon.Zprava('Tisk dokladù ukonèen');
  end;
end;

// ------------------------------------------------------------------------------------------------

procedure TdmTisk.FOxTisk(Radek: integer);
var
  AbraKod,
  SQLStr: AnsiString;
  Zaplaceno: double;
begin
  with fmMain, asgMain do begin
    with qrAbra do begin
// údaje z faktury do privátních promìnných
      Close;
      SQLStr := 'SELECT F.Code, F.Name, Street, City, PostCode, OrgIdentNumber, VATIdentNumber, II.ID,'
       + ' D.Code || ''-'' || lpad(II.OrdNumber, 4, ''0'') || ''/'' || substring(P.Code from 3 for 2) AS Cislo,'
       + ' DocDate$DATE, DueDate$DATE, VATDate$DATE, LocalAmount, LocalPaidAmount'
      + ' FROM Firms F, Addresses A, IssuedInvoices II, DocQueues D, Periods P'
      + ' WHERE F.ID = II.Firm_ID'
      + ' AND A.ID = F.ResidenceAddress_ID'
      + ' AND D.ID = II.DocQueue_ID'
      + ' AND P.ID = II.Period_ID'
      + ' AND F.Hidden = ''N''' ;
      SQLStr := SQLStr + ' AND II.Period_ID = ' + Ap + Period_Id + Ap
      + ' AND II.OrdNumber = ' + Cells[2, Radek];
      if rbInternet.Checked then SQLStr := SQLStr + ' AND DocQueue_ID = ' + Ap + FO4Queue_Id + Ap
      else SQLStr := SQLStr + ' AND DocQueue_ID = ' + Ap + FO2Queue_Id + Ap;
      SQL.Text := SQLStr;
      Open;
      if RecordCount = 0 then begin
        dmCommon.Zprava(Format('Neexistuje doklad s èíslem %d nebo zákazník %s.', [Ints[2, Radek], Cells[4, Radek]]));
        Close;
        Exit;
      end;
      CisloFO := FieldByName('Cislo').AsString;
      if Trim(FieldByName('Code').AsString) = '' then begin
        dmCommon.Zprava(Format('Doklad %s: zákazník nemá kód Abry.', [CisloFO]));
        Close;
        Exit;
      end;
      AbraKod := FieldByName('Code').AsString;
      OJmeno := FieldByName('Name').AsString;
      OUlice := FieldByName('Street').AsString;
      OObec := FieldByName('PostCode').AsString + ' ' + FieldByName('City').AsString;
      OICO := FieldByName('OrgIdentNumber').AsString;
      ODIC := FieldByName('VATIdentNumber').AsString;
      FO_Id := FieldByName('ID').AsString;
      DatumDokladu := FieldByName('DocDate$DATE').AsFloat;
      DatumSplatnosti := FieldByName('DueDate$DATE').AsFloat;
      DatumPlneni := FieldByName('VATDate$DATE').AsFloat;
      Vystaveni := FormatDateTime('dd.mm.yyyy', DatumDokladu);
      Plneni := FormatDateTime('dd.mm.yyyy', DatumPlneni);
      Splatnost := FormatDateTime('dd.mm.yyyy', DatumSplatnosti);
      VS := Cells[1, Radek];
      Celkem := FieldByName('LocalAmount').AsFloat;
      Zaplaceno := FieldByName('LocalPaidAmount').AsFloat;
      Close;
//    Cislo := Format('FO4-%4.4d/%d', [Ints[2, Radek], YearOf(DatumDokladu)]);
// všechny Firm_Id pro Abrakód firmy
      SQLStr := 'SELECT * FROM DE$_CODE_TO_FIRM_ID (' + Ap + AbraKod + ApZ;
      SQL.Text := SQLStr;
      Open;
      if Fields[0].AsString = 'MULTIPLE' then begin
        dmCommon.Zprava(Format('%s (%s): Více zákazníkù pro kód %s.', [OJmeno, VS, AbraKod]));
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
    Saldo := Saldo + Zaplaceno;          // Saldo je po splatnosti (SaldoPo), je-li faktura už zaplacena, pøiète se platba
    Zaplatit := Celkem - Saldo;          // Celkem k úhradì = Celkem za fakt. období - Zùstatek minulých období(saldo)
// text na fakturu
    if Saldo > 0 then Platek := 'pøeplatek'
    else if Saldo < 0 then Platek := 'nedoplatek'
    else Platek := ' ';
    if Zaplatit < 0 then Zaplatit := 0;
// údaje z tabulky Smlouvy do globálních promìnných
    with qrMain do begin
      Close;
// u FOx je VS èíslo smlouvy
      SQLStr := 'SELECT Postal_name, Postal_street, Postal_PSC, Postal_city FROM customers Cu, contracts C'
      + ' WHERE Cu.Id = C.Customer_Id'
      + ' AND C.Number = ' + Ap + VS + Ap;
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
      frxReport.LoadFromFile(ExtractFilePath(ParamStr(0)) + 'FOxdoPDF.fr3');
      frxReport.PrepareReport;
//      frxReport.ShowPreparedReport;
      frxReport.PrintOptions.ShowDialog := False;
      frxReport.Print;
      dmCommon.Zprava(Format('%s (%s): %s byla odeslána na tiskárnu.', [OJmeno, VS, CisloFO]));
      Ints[0, Radek] := 0;
    except on E: exception do
      begin
        dmCommon.Zprava(Format('%s (%s): %s se nepodaøilo vytisknout.' + #13#10 + 'Chyba: %s',
         [OJmeno, VS, CisloFO, E.Message]));
        if Application.MessageBox(PChar('Chyba pøi tisku' + ^M + E.Message), 'Pokraèovat?',
         MB_YESNO + MB_ICONQUESTION) = IDNO then Prerusit := True;
      end;
    end;  // try
  end;  // with fmMain
end;

end.

