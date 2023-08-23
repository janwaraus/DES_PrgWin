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
// pou�ije data z asgMain
var
  Radek: integer;
begin
  with fmMain do try
    dmCommon.Zprava('Tisk doklad�');
    with asgMain do begin
      dmCommon.Zprava(Format('Po�et doklad� k tisku: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));
      if ColumnSum(0, 1, RowCount-1) >= 1 then
        if dlgTisk.Execute then
          frxReport.PrintOptions.Printer := Printer.Printers[Printer.PrinterIndex];
      Screen.Cursor := crHourGlass;
      apnTisk.Visible := False;
      apbProgress.Position := 0;
      apbProgress.Visible := True;
// hlavn� smy�ka
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
// konec hlavn� smy�ky
  finally
    Printer.PrinterIndex := -1;
    apbProgress.Position := 0;
    apbProgress.Visible := False;
    apnTisk.Visible := True;
    asgMain.Visible := False;
    lbxLog.Visible := True;
    Screen.Cursor := crDefault;                                             // default
    dmCommon.Zprava('Tisk doklad� ukon�en');
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
// �daje z faktury do priv�tn�ch prom�nn�ch
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
        dmCommon.Zprava(Format('Neexistuje doklad s ��slem %d nebo z�kazn�k %s.', [Ints[2, Radek], Cells[4, Radek]]));
        Close;
        Exit;
      end;
      CisloFO := FieldByName('Cislo').AsString;
      if Trim(FieldByName('Code').AsString) = '' then begin
        dmCommon.Zprava(Format('Doklad %s: z�kazn�k nem� k�d Abry.', [CisloFO]));
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
// v�echny Firm_Id pro Abrak�d firmy
      SQLStr := 'SELECT * FROM DE$_CODE_TO_FIRM_ID (' + Ap + AbraKod + ApZ;
      SQL.Text := SQLStr;
      Open;
      if Fields[0].AsString = 'MULTIPLE' then begin
        dmCommon.Zprava(Format('%s (%s): V�ce z�kazn�k� pro k�d %s.', [OJmeno, VS, AbraKod]));
        Close;
        Exit;
      end;
      Saldo := 0;
  // a saldo pro v�echny Firm_Id
      while not EOF do with qrAdresa do begin
        Close;
        SQLStr := 'SELECT SaldoPo + SaldoZLPo + Ucet325 FROM DE$_Firm_Totals (' + Ap + qrAbra.Fields[0].AsString + ApC + FloatToStr(DatumDokladu) + ')';
        SQL.Text := SQLStr;
        Open;
        Saldo := Saldo + Fields[0].AsFloat;
        qrAbra.Next;
      end; // while not EOF do with qrAdresa
    end;  // with qrAbra
    Saldo := Saldo + Zaplaceno;          // Saldo je po splatnosti (SaldoPo), je-li faktura u� zaplacena, p�i�te se platba
    Zaplatit := Celkem - Saldo;          // Celkem k �hrad� = Celkem za fakt. obdob� - Z�statek minul�ch obdob�(saldo)
// text na fakturu
    if Saldo > 0 then Platek := 'p�eplatek'
    else if Saldo < 0 then Platek := 'nedoplatek'
    else Platek := ' ';
    if Zaplatit < 0 then Zaplatit := 0;
// �daje z tabulky Smlouvy do glob�ln�ch prom�nn�ch
    with qrMain do begin
      Close;
// u FOx je VS ��slo smlouvy
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
// zas�lac� adresa
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
      dmCommon.Zprava(Format('%s (%s): %s byla odesl�na na tisk�rnu.', [OJmeno, VS, CisloFO]));
      Ints[0, Radek] := 0;
    except on E: exception do
      begin
        dmCommon.Zprava(Format('%s (%s): %s se nepoda�ilo vytisknout.' + #13#10 + 'Chyba: %s',
         [OJmeno, VS, CisloFO, E.Message]));
        if Application.MessageBox(PChar('Chyba p�i tisku' + ^M + E.Message), 'Pokra�ovat?',
         MB_YESNO + MB_ICONQUESTION) = IDNO then Prerusit := True;
      end;
    end;  // try
  end;  // with fmMain
end;

end.

