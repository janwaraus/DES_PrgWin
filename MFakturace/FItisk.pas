unit FItisk;

interface

uses
  Windows, Messages, Forms, Controls, Classes, Dialogs, SysUtils, DateUtils, Printers, FImain;

type
  TdmTisk = class(TDataModule)
    dlgTisk: TPrintDialog;
  private
    procedure FakturaTisk(Radek: integer);                // vytiskne jednu fakturu
  public
    procedure TiskniFaktury;
  end;

var
  dmTisk: TdmTisk;

implementation

{$R *.dfm}

uses DesUtils, FIcommon;

// ------------------------------------------------------------------------------------------------

procedure TdmTisk.TiskniFaktury;
// pou�ije data z asgMain
var
  Posilani: AnsiString;
  Radek: integer;
begin
  with fmMain do try
{$IFDEF ABAK}
    Posilani := '';
{$ELSE}
    if rbBezSlozenky.Checked then Posilani := 'bez slo�enky';
    if rbSeSlozenkou.Checked then Posilani := 'se slo�enkou';
    if rbKuryr.Checked then Posilani := 'rozn�en�ch kur�rem';
{$ENDIF}
    if rbPodleSmlouvy.Checked then dmCommon.Zprava(Format('Tisk faktur %s od VS %s do %s', [Posilani, aedOd.Text, aedDo.Text]))
    else dmCommon.Zprava(Format('Tisk faktur %s od ��sla %s do %s', [Posilani, aedOd.Text, aedDo.Text]));
    with asgMain do begin
      dmCommon.Zprava(Format('Po�et faktur k tisku: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));
      if ColumnSum(0, 1, RowCount-1) >= 1 then
        if dlgTisk.Execute then
          //frxReport.PrintOptions.Printer := Printer.Printers[Printer.PrinterIndex];
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
        if Ints[0, Radek] = 1 then FakturaTisk(Radek);
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
    dmCommon.Zprava('Tisk faktur ukon�en');
  end;
end;

// ------------------------------------------------------------------------------------------------

procedure TdmTisk.FakturaTisk(Radek: integer);
var
  AbraKod,
  SQLStr: AnsiString;
  Zaplaceno: double;
  i: integer;
begin
  with fmMain, asgMain do begin
    with qrAbra do begin
// �daje z faktury do priv�tn�ch prom�nn�ch
      Close;
      SQLStr := 'SELECT Code, Name, Street, City, PostCode, OrgIdentNumber, VATIdentNumber, II.ID,'
       + ' DocDate$DATE, DueDate$DATE, VATDate$DATE, LocalAmount, LocalPaidAmount'
      + ' FROM Firms F, Addresses A, IssuedInvoices II'
      + ' WHERE F.ID = II.Firm_ID'
      + ' AND A.ID = F.ResidenceAddress_ID'
      + ' AND F.Hidden = ''N'''
      + ' AND F.Firm_ID IS NULL';                             // posledn�, bez n�sledovn�ka
      SQLStr := SQLStr + ' AND II.Period_ID = ' + Ap + globalAA['abraIiPeriod_Id'] + Ap
      + ' AND II.OrdNumber = ' + Cells[2, Radek]
      + ' AND II.DocQueue_ID = ';
{$IFDEF ABAK}
      if rbInternet.Checked then SQLStr := SQLStr + Ap + globalAA['abraIiDocQueue_Id'] + Ap
      else SQLStr := SQLStr + Ap + VDocQueue_Id + Ap;
{$ELSE}
      SQLStr := SQLStr + Ap + globalAA['abraIiDocQueue_Id'] + Ap;
{$ENDIF}
      SQL.Text := SQLStr;
      Open;
      if RecordCount = 0 then begin
        dmCommon.Zprava(Format('Neexistuje faktura %d nebo z�kazn�k %s.', [Ints[2, Radek], Cells[4, Radek]]));
        Close;
        Exit;
      end;
      if Trim(FieldByName('Code').AsString) = '' then begin
        dmCommon.Zprava(Format('Faktura %d: z�kazn�k nem� k�d Abry.', [Ints[2, Radek]]));
        Close;
        Exit;
      end;
      AbraKod := FieldByName('Code').AsString;
      OJmeno := FieldByName('Name').AsString;
      OUlice := FieldByName('Street').AsString;
      OObec := FieldByName('PostCode').AsString + ' ' + FieldByName('City').AsString;
      OICO := FieldByName('OrgIdentNumber').AsString;
      ODIC := FieldByName('VATIdentNumber').AsString;
      ID := FieldByName('ID').AsString;
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
      SS := Format('%6.6d%2.2d', [Ints[2, Radek], aseRok.Value - 2000]);
{$IFDEF ABAK}
      if rbInternet.Checked then FStr := 'FiI'
      else FStr := 'FtI';
{$ELSE}
      FStr := 'FO1';
      if Length(VS) = 8 then V := '00' + VS else V := VS;                      // na slo�enku
      S := Format('%8.8d%2.2d', [Ints[2, Radek], aseRok.Value - 2000]);
{$ENDIF}
      Cislo := Format('%s-%5.5d/%d', [FStr, Ints[2, Radek], aseRok.Value]);
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
{$IFDEF ABAK}
        if rbInternet.Checked then
          SQLStr := 'SELECT SaldoPo FROM AB$_Saldo_FiI (' + Ap + qrAbra.Fields[0].AsString + ApC + FloatToStr(DatumDokladu) + ')'
        else SQLStr := 'SELECT SaldoPo FROM AB$_Saldo_FtI (' + Ap + qrAbra.Fields[0].AsString + ApC + FloatToStr(DatumDokladu) + ')';
{$ELSE}
        SQLStr := 'SELECT SaldoPo + SaldoZLPo + Ucet325 FROM DE$_Firm_Totals (' + Ap + qrAbra.Fields[0].AsString + ApC + FloatToStr(DatumDokladu) + ')';
//        SQLStr := 'SELECT SaldoPo + Ucet325 FROM DE$_Firm_Totals (' + Ap + qrAbra.Fields[0].AsString + ApC + FloatToStr(Date) + ')';
{$ENDIF}
        SQL.Text := SQLStr;
        Open;
        Saldo := Saldo + Fields[0].AsFloat;
        qrAbra.Next;
      end; // while not EOF do with qrAdresa
    end;  // with qrAbra
// pr�v� p�ev�d�n� faktura m��e b�t p�ed splatnost�
//    if Date <= DatumSplatnosti then begin
      Saldo := Saldo + Zaplaceno;          // Saldo je po splatnosti (SaldoPo), je-li faktura u� zaplacena, p�i�te se platba
      Zaplatit := Celkem - Saldo;          // Celkem k �hrad� = Celkem za fakt. obdob� - Z�statek minul�ch obdob�(saldo)
// anebo je po splatnosti
{    end else begin
      Zaplatit := -Saldo;
      Saldo := Saldo + Celkem;             // ��stka faktury se ode�te ze salda, aby tam nebyla dvakr�t
    end;   }
// text na fakturu
    if Saldo > 0 then Platek := 'p�eplatek'
    else if Saldo < 0 then Platek := 'nedoplatek'
    else Platek := ' ';
    if Zaplatit < 0 then Zaplatit := 0;
{$IFNDEF ABAK}
    C := Format('%6.0f', [Zaplatit]);
    for i := 2 to 6 do
      if C[i] <> ' ' then begin
        C[i-1] := '~';
        Break;
      end;
{$ENDIF}
// �daje z tabulky Smlouvy do glob�ln�ch prom�nn�ch
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
// zas�lac� adresa
    if (PJmeno = '') or (PObec = '') then begin
      PJmeno := OJmeno;
      PUlice := OUlice;
      PObec := OObec;
    end;
    try

    {*hw* TODO
      if rbSeSlozenkou.Checked then
        frxReport.LoadFromFile(ExtractFilePath(ParamStr(0)) + 'FOseSlozenkou.fr3')
      else
        frxReport.LoadFromFile(ExtractFilePath(ParamStr(0)) + 'FOsPDP.fr3');

      frxReport.PrepareReport;
//      frxReport.ShowPreparedReport;
      frxReport.PrintOptions.ShowDialog := False;
      frxReport.Print;
      dmCommon.Zprava(Format('%s (%s): Faktura %s byla odesl�na na tisk�rnu.', [OJmeno, VS, Cislo]));
      Ints[0, Radek] := 0;
      }
    except on E: exception do
      begin
        dmCommon.Zprava(Format('%s (%s): Fakturu %s se nepoda�ilo vytisknout.' + #13#10 + 'Chyba: %s',
         [OJmeno, VS, Cislo, E.Message]));
        if Application.MessageBox(PChar('Chyba p�i tisku' + ^M + E.Message), 'Pokra�ovat?',
         MB_YESNO + MB_ICONQUESTION) = IDNO then Prerusit := True;
      end;
    end;  // try
  end;  // with fmMain
end;

end.

