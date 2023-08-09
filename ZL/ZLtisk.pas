unit ZLtisk;

interface

uses
  Windows, Messages, Forms, Controls, Classes, Dialogs, SysUtils, DateUtils, Printers, DesUtils;

type
  TdmTisk = class(TDataModule)
    dlgTisk: TPrintDialog;
  private
    procedure ZLTisk(Radek: integer);                // vytiskne jeden ZL
  public
    procedure TiskniZL;
  end;

var
  dmTisk: TdmTisk;

implementation

{$R *.dfm}

uses DesInvoices, DesFastReports, AArray, ZLmain;

// ------------------------------------------------------------------------------------------------

procedure TdmTisk.TiskniZL;
// použije data z asgMain
var
  Radek: integer;
  Faktura : TDesInvoice;
  VysledekPrevedeni : TDesResult;
begin
  with fmMain do try
    fmMain.Zprava(Format('Tisk ZL od èísla %s do %s', [aedOd.Text, aedDo.Text]));
    with asgMain do begin
      fmMain.Zprava(Format('Poèet ZL k tisku: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));
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

        if Ints[0, Radek] = 1 then begin  // pokud zaškrtnuto, pøevádíme fa do PDF

          Faktura := TDesInvoice.create(asgMain.Cells[8, Radek], '10');
          VysledekPrevedeni := Faktura.printByFr3('ZLdoPDF-104.fr3');

          fmMain.Zprava(Format('%s (%s): %s', [asgMain.Cells[4, Radek], asgMain.Cells[1, Radek], VysledekPrevedeni.Messg]));

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
    apnTisk.Visible := True;
    asgMain.Visible := False;
    lbxLog.Visible := True;
    Screen.Cursor := crDefault;                                             // default
    fmMain.Zprava('Tisk ZL ukonèen');
  end;
end;

// ------------------------------------------------------------------------------------------------

procedure TdmTisk.ZLTisk(Radek: integer);
var
  AbraKod,
  SQLStr: AnsiString;
  Zaplaceno: double;
  i: integer;
begin
{
  with fmMain, fmMain.asgMain do begin
    with qrAbra do begin
// údaje z faktury do privátních promìnných
      Close;
      SQLStr := 'SELECT F.Code, F.Name, Street, City, PostCode, OrgIdentNumber, VATIdentNumber, II.ID, LocalAmount, LocalPaidAmount,'
       + ' DocDate$DATE, DueDate$DATE'
      + ' FROM Firms F, Addresses A, IssuedDInvoices II, Periods P'
      + ' WHERE F.ID = II.Firm_ID'
      + ' AND A.ID = F.ResidenceAddress_ID'
      + ' AND P.ID = II.Period_ID'
      + ' AND F.Hidden = ''N'''
      + ' AND F.Firm_ID IS NULL'                            // poslední, bez následovníka
      + ' AND P.Code = ' + Ap + Copy(Cells[2, Radek], Pos('/', Cells[2, Radek])+1, 4) + Ap
      + ' AND II.OrdNumber = ' + Copy(Cells[2, Radek], 0, Pos('/', Cells[2, Radek])-1)
      + ' AND II.DocQueue_ID = ' + Ap + AbraEnt.getDocQueue('Code=ZL1').ID
       + Ap;
      SQL.Text := SQLStr;
      Open;
      if RecordCount = 0 then begin
        fmMain.Zprava(Format('Neexistuje ZL %s nebo zákazník %s.', [Copy(Cells[2, Radek], 0, Pos('/', Cells[2, Radek])-1), Cells[4, Radek]]));
        Close;
        Exit;
      end;
      if Trim(FieldByName('Code').AsString) = '' then begin
        fmMain.Zprava(Format('ZL %s: zákazník nemá kód Abry.', [Copy(Cells[2, Radek], 0, Pos('/', Cells[2, Radek])-1)]));
        Close;
        Exit;
      end;
      AbraKod := FieldByName('Code').AsString;
      OJmeno := UTF8Decode(FieldByName('Name').AsString);
      OUlice := UTF8Decode(FieldByName('Street').AsString);
      OObec := UTF8Decode(FieldByName('PostCode').AsString + ' ' + FieldByName('City').AsString);
      OICO := FieldByName('OrgIdentNumber').AsString;
      ODIC := FieldByName('VATIdentNumber').AsString;
      ID := FieldByName('ID').AsString;
      DatumDokladu := FieldByName('DocDate$DATE').AsFloat;
      DatumSplatnosti := FieldByName('DueDate$DATE').AsFloat;
      Vystaveni := FormatDateTime('dd.mm.yyyy', DatumDokladu);
      Splatnost := FormatDateTime('dd.mm.yyyy', DatumSplatnosti);
      VS := Cells[1, Radek];
      Celkem := FieldByName('LocalAmount').AsFloat;
      Zaplaceno := FieldByName('LocalPaidAmount').AsFloat;
      Close;
      SS := Format('%6.6d%2.2d', [StrToInt(Copy(Cells[2, Radek], 0, Pos('/', Cells[2, Radek])-1)), aseRok.Value - 2000]);
      if Length(VS) = 8 then V := '00' + VS else V := VS;                      // na složenku
      S := Format('%8.8d%2.2d', [StrToInt(Copy(Cells[2, Radek], 0, Pos('/', Cells[2, Radek])-1)), aseRok.Value - 2000]);
      Cislo := Format('ZL1-%4.4d/%d', [StrToInt(Copy(Cells[2, Radek], 0, Pos('/', Cells[2, Radek])-1)), aseRok.Value]);
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
        SQLStr := 'SELECT SaldoPo + Ucet325 FROM DE$_Firm_Totals (' + Ap + qrAbra.Fields[0].AsString + ApC + FloatToStr(Date) + ')';
        SQL.Text := SQLStr;
        Open;
        Saldo := Saldo + Fields[0].AsFloat;
        qrAbra.Next;
      end; // while not EOF do with qrAdresa
    end;  // with qrAbra
// právì pøevádìný ZL mùže být pøed splatností
    if Date <= DatumSplatnosti then begin
      Saldo := Saldo + Zaplaceno;          // Saldo je po splatnosti (SaldoPo), je-li faktura už zaplacena, pøiète se platba
      Zaplatit := Celkem - Saldo;          // Celkem k úhradì = Celkem za fakt. období - Zùstatek minulých období(saldo)
// anebo je po splatnosti
    end else begin
      Zaplatit := -Saldo;
      Saldo := Saldo + Celkem;             // èástka faktury se odeète ze salda, aby tam nebyla dvakrát
    end;
// text na fakturu
    if Saldo > 0 then Platek := 'pøeplatek'
    else if Saldo < 0 then Platek := 'nedoplatek'
    else Platek := ' ';
    if Zaplatit < 0 then Zaplatit := 0;
    C := Format('%6.0f', [Zaplatit]);
    for i := 2 to 6 do
      if C[i] <> ' ' then begin
        C[i-1] := '~';
        Break;
      end;
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
     frxReport.LoadFromFile(ExtractFilePath(ParamStr(0)) + 'FR3\ZLdoPDF.fr3');
      frxReport.PrepareReport;
//      frxReport.ShowPreparedReport;
      frxReport.PrintOptions.ShowDialog := False;
      frxReport.Print;
      fmMain.Zprava(Format('%s (%s): ZL %s byl odeslán na tiskárnu.', [OJmeno, VS, Cislo]));
      Ints[0, Radek] := 0;
    except on E: exception do
      begin
        fmMain.Zprava(Format('%s (%s): ZL %s se nepodaøilo vytisknout.' + #13#10 + 'Chyba: %s',
         [OJmeno, VS, Cislo, E.Message]));
        if Application.MessageBox(PChar('Chyba pøi tisku' + ^M + E.Message), 'Pokraèovat?',
         MB_YESNO + MB_ICONQUESTION) = IDNO then Prerusit := True;
      end;
    end;  // try
  end;  // with fmMain
}
end;

end.

