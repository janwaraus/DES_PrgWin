// DES fakturuje najednou všechny smlouvy zákazníka, internet, VoIP i IPTV s datem vystavení i plnìní poslední den v mìsíci.
// ABAK fakturuje každou smlouvu zvláš, internetové smlouvy na zaèátku mìsíce (zálohovì), VoIPové ka konci
// Zákazníci nebo smlouvy k fakturaci jsou v asgMain, faktury se vytvoøí v cyklu pøes øádky.
// 4.12.14 ABAK: VoIPové smlouvy jednoho zákazníka se fakturují spoleènì.
// 9.4.15 ABAK: fakturování i smluv s pøíbìhem
// 23.1.17 Výkazy pro ÈTÚ vyžadují dìlení podle typu uživatele, technologie a rychlosti - do Abry byly pøidány 3 obchodní pøípady
// - VO (velkoobchod - DRC) , F a P (fyzická a právnická osoba) a 24 zakázek (W - WiFi, B - FFTB, A - FTTH AON, a P - FTTH PON,
// každá s šesti rùznými rychlostmi). Každému øádku faktury bude pøiøazen jeden obchodní pøípad a jedna zakázka.
// 27.1. vyèlenìna procedura FakturaAbra jen pro DES

unit FIfaktura;

interface

uses
  Windows, Messages, Dialogs, Classes, Forms, Controls, SysUtils, DateUtils, Variants, ComObj, Math, FImain;

type
  TdmFaktura = class(TDataModule)
  private
    isDRC : boolean;
    procedure FakturaAbra(Radek: integer);
  public
    procedure VytvorFaktury;
  end;

var
  dmFaktura: TdmFaktura;

implementation

{$R *.dfm}

uses DesUtils, FIcommon, FIlogin;

// ------------------------------------------------------------------------------------------------

procedure TdmFaktura.VytvorFaktury;
var
  Radek: integer;
begin
  with fmMain, fmMain.asgMain do try
// pøipojení k Abøe
    Screen.Cursor := crHourGlass;
    asgMain.Visible := False;
    lbxLog.Visible := True;

    {
    try
      dmCommon.Zprava('Pøipojení k Abøe ...');
      AbraOLE := CreateOLEObject('AbraOLE.Application');
      if not AbraOLE.Connect('@' + AbraConnection) then begin
        dmCommon.Zprava('Problém s Abrou (connect  "' + AbraConnection + '").');
        Screen.Cursor := crDefault;
        Exit;
      end;
      dmCommon.Zprava('OK');
      dmCommon.Zprava('Login ...');
      fmLogin.ShowModal;
      if not AbraOLE.Login(fmLogin.acbJmeno.Text, fmLogin.aedHeslo.Text) then begin
        dmCommon.Zprava('Problém s Abrou (jméno, heslo).');
        Screen.Cursor := crDefault;
        Exit;
      end;
      dmCommon.Zprava('OK');
    except on E: exception do
      begin
        Application.MessageBox(PChar('Problém s Abrou.' + ^M + E.Message), 'Abra', MB_ICONERROR + MB_OK);
        dmCommon.Zprava('Problém s Abrou.' + #13#10 + E.Message);
        btKonec.Caption := '&Konec';
        Screen.Cursor := crDefault;
        Exit;
      end;
    end;
    }

    dmCommon.Zprava(Format('Poèet faktur k vygenerování: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));
    apbProgress.Position := 0;
    apbProgress.Visible := True;
    asgMain.Visible := True;
    lbxLog.Visible := False;
// hlavní smyèka
    for Radek := 1 to RowCount-1 do begin
      Row := Radek;
      apbProgress.Position := Round(100 * Radek / RowCount-1);
      Application.ProcessMessages;
      if Prerusit then begin
        Prerusit := False;
        btVytvorit.Enabled := True;
        asgMain.Visible := True;
        lbxLog.Visible := False;
        Break;
      end;
      if Ints[0, Radek] = 1 then FakturaAbra(Radek);
    end;  // for
// konec hlavní smyèky
  finally

    apbProgress.Position := 0;
    apbProgress.Visible := False;
    asgMain.Visible := False;
    lbxLog.Visible := True;
    Screen.Cursor := crDefault;
    dmCommon.Zprava('Generování faktur ukonèeno');
  end;  //  with fmMain
end;

// ------------------------------------------------------------------------------------------------


procedure TdmFaktura.FakturaAbra(Radek: integer);
// pro mìsíc a rok zadaný v aseMesic a aseRok vytvoøí fakturu za pøipojení k internetu a za VoIP
// 27.1.17 celé pøehlednìji
var
  Firm_Id,
  FirmOffice_Id,
  BusOrder_Id,
  BusOrderCode,
  Speed,
  BusTransaction_Id,
  BusTransactionCode,
  IC,
  ID: string[10];
  FCena,
  CenaTarifu,
  Redukce,
  PausalVoIP,
  HovorneVoIP,
  DatumSpusteni,
  DatumUkonceni: double;
  FakturaVoIP,
  SmlouvaVoIP,
  ProvolatelnyPausal,
  PosledniFakturace,
  PrvniFakturace: boolean;
  Dotaz,
  CisloFaktury: integer;
  FObject,
  FData,
  FRow,
  FRowsCollection: variant;
  Zakaznik,
  Description,
  SQLStr: AnsiString;
begin
  with fmMain, qrSmlouva, asgMain do begin
// faktura se nebude vytváøet, je-li ve smlouvách také fakturovaný VoIP a VoIPy se nefakturují, nebo se VoIPy fakturují a smlouva není
    Close;
    SQLStr := 'SELECT COUNT(*) FROM ' + fiInvoiceView
    + ' WHERE (Tarif = ''EP-Home'' OR Tarif = ''EP-Profi'')'
    + ' AND VS = ' + Ap + Cells[1, Radek] + Ap;
    SQL.Text := SQLStr;
    Open;
    FakturaVoIP := Fields[0].AsInteger > 0;
    if (not cbSVoIP.Checked and FakturaVoIP) or (not cbBezVoIP.Checked and not FakturaVoIP) then begin
      Close;
      Exit;
    end;
// není víc zákazníkù pro jeden VS ? 31.1.2017
    Close;
    SQLStr := 'SELECT COUNT(*) FROM customers'
    + ' WHERE Variable_Symbol = ' + Ap + Cells[1, Radek] + Ap;
    SQL.Text := SQLStr;
    Open;
    if Fields[0].AsInteger > 1 then begin
      dmCommon.Zprava(Format('Variabilní symbol %s má více zákazníkù.', [Cells[1, Radek]]));
      Close;
      Exit;
    end;
// je pøenesená daòová povinnost ?  19.10.2016 - bude pro celou fakturu, ne pro jednotlivé smlouvy
    Close;
    SQLStr := 'SELECT COUNT(*) FROM contracts_tags'
    + ' WHERE Tag_Id = ' + IntToStr(DRCTag_Id)
    + ' AND Contract_Id IN (SELECT Id FROM contracts'
      + ' WHERE Invoice = 1'
      + ' AND Customer_Id = (SELECT Id FROM customers'
        + ' WHERE Variable_Symbol = ' + Ap + Cells[1, Radek] + ApZ + ')';
    SQL.Text := SQLStr;
    Open;
    isDRC := Fields[0].AsInteger > 0;
// vyhledání údajù o smlouvách
    Close;
    SQLStr := 'SELECT VS, AbraKod, Typ, Smlouva, AktivniOd, AktivniDo, FakturovatOd, Tarif, Posilani, Perioda, Text, Cena, DPH, Tarifni, CTU'
    + ' FROM ' + fiInvoiceView
    + ' WHERE VS = ' + Ap + Cells[1, Radek] + Ap
    + ' ORDER BY Smlouva, Tarifni DESC';
    SQL.Text := SQLStr;
    Open;
// pøi lecjaké chybì v databázi (napø. Tariff_Id je NULL) konec
    if RecordCount = 0 then begin
      dmCommon.Zprava(Format('Pro variabilní symbol %s není co fakturovat.', [FieldByName('VS').AsString]));
      Close;
      Exit;
    end;
    if FieldByName('AbraKod').AsString = '' then begin
      dmCommon.Zprava(Format('Smlouva %s: zákazník nemá kód Abry.', [FieldByName('Smlouva').AsString]));
      Close;
      Exit;
    end;
    with qrAbra do begin
// kontrola kódu firmy, pøi chybì konec
      Close;
      SQLStr := 'SELECT F.ID, F.Name, F.OrgIdentNumber, FO.Id FROM Firms F, FirmOffices FO'
      + ' WHERE Code = ' + Ap + qrSmlouva.FieldByName('AbraKod').AsString + Ap
      + ' AND F.Firm_ID IS NULL'         // bez pøedkù
      + ' AND F.Hidden = ''N'''
      + ' AND FO.Parent_Id = F.Id'
      + ' ORDER BY F.ID DESC';
      SQL.Text := SQLStr;
      Open;
      if RecordCount = 0 then begin
        dmCommon.Zprava(Format('Smlouva %s: zákazník s kódem %s není v adresáøi Abry.',
         [qrSmlouva.FieldByName('Smlouva').AsString, qrSmlouva.FieldByName('AbraKod').AsString]));
        Exit;
      end else begin
        Firm_Id := Fields[0].AsString;
        Zakaznik := Fields[1].AsString;
        IC := Trim(Fields[2].AsString);
        FirmOffice_Id := Fields[3].AsString;
      end;
// 24.1.2017 obchodní pøípady pro ÈTÚ - platí pro celou faktury
      if isDRC then BusTransactionCode := 'VO'                  // velkoobchod
      else if IC = '' then BusTransactionCode := 'F'         //  fyzická osoba
      else BusTransactionCode := 'P';                         //  právnická osoba
      with qrAbra do begin
        Close;
        SQLStr := 'SELECT Id FROM BusTransactions'
        + ' WHERE Code = ' + Ap + BusTransactionCode + Ap;
        SQL.Text := SQLStr;
        Open;
        BusTransaction_Id := Fields[0].AsString;
        Close;
      end;

// kontrola poslední faktury
      Close;
      SQLStr := 'SELECT OrdNumber, DocDate$DATE, VATDate$DATE, Amount FROM IssuedInvoices'
      + ' WHERE VarSymbol = ' + Ap + qrSmlouva.FieldByName('VS').AsString + Ap
      + ' AND VATDate$DATE >= ' + FloatToStr(Trunc(StartOfAMonth(aseRok.Value, aseMesic.Value)))
      + ' AND VATDate$DATE <= ' + FloatToStr(Trunc(EndOfAMonth(aseRok.Value, aseMesic.Value)));
      SQLStr := SQLStr + ' AND DocQueue_ID = ' + Ap + globalAA['abraIiDocQueue_Id'] + Ap;
      SQLStr := SQLStr + ' ORDER BY OrdNumber DESC';
      SQL.Text := SQLStr;
      Open;
      if isDebugMode then dmCommon.Zprava('Vyhledána data z Abry');
      if RecordCount > 0 then begin
        dmCommon.Zprava(Format('%s (%s): %d. faktura se stejným datem.',
         [Zakaznik, Cells[1, Radek], RecordCount + 1]));
        Dotaz := Application.MessageBox(PChar(Format('Pro zákazníka "%s" existuje faktura %s-%s s datem %s na èástku %s Kè. Má se vytvoøit další?',
         [Zakaznik, globalAA['invoiceDocQueueCode'], FieldByName('OrdNumber').AsString, DateToStr(FieldByName('DocDate$DATE').AsFloat), FieldByName('Amount').AsString])),
          'Pozor', MB_ICONQUESTION + MB_YESNOCANCEL + MB_DEFBUTTON1);
        if Dotaz = IDNO then begin
          dmCommon.Zprava('Ruèní zásah - faktura nevytvoøena.');
          Exit;
        end else if Dotaz = IDCANCEL then begin
          dmCommon.Zprava('Ruèní zásah - program ukonèen.');
          Prerusit := True;
          Exit;
        end else dmCommon.Zprava('Ruèní zásah - faktura se vytvoøí.');
      end;  // RecordCount > 0
    end;  // with qrAbra
    Description := Format('pøipojení %d/%d, ', [aseMesic.Value, aseRok.Value-2000]);
// vytvoøí se hlavièka faktury
    FObject:= AbraOLE.CreateObject('@IssuedInvoice');
    FData:= AbraOLE.CreateValues('@IssuedInvoice');
    FObject.PrefillValues(FData);
    FData.ValueByName('DocQueue_ID') := globalAA['abraIiDocQueue_Id'];
    FData.ValueByName('Period_ID') := globalAA['abraIiPeriod_Id'];
    FData.ValueByName('DocDate$DATE') := Floor(deDatumDokladu.Date);
    FData.ValueByName('AccDate$DATE') := Floor(deDatumDokladu.Date);
    FData.ValueByName('VATDate$DATE') := Floor(deDatumPlneni.Date);
    FData.ValueByName('CreatedBy_ID') := User_Id;
    FData.ValueByName('Firm_ID') := Firm_Id;
    FData.ValueByName('FirmOffice_ID') := FirmOffice_Id;
    FData.ValueByName('Address_ID') := MyAddress_Id;
    FData.ValueByName('BankAccount_ID') := MyAccount_Id;
    FData.ValueByName('ConstSymbol_ID') := '0000308000';
    FData.ValueByName('VarSymbol') := FieldByName('VS').AsString;
    FData.ValueByName('TransportationType_ID') := '1000000101';
    FData.ValueByName('DueTerm') := aedSplatnost.Text;
    FData.ValueByName('PaymentType_ID') := MyPayment_Id;
    FData.ValueByName('PricesWithVAT') := True;
    FData.ValueByName('VATFromAbovePrecision') := 6;
    FData.ValueByName('TotalRounding') := 259;                // zaokrouhlení na koruny dolù
    if isDRC then begin
      FData.ValueByName('IsReverseChargeDeclared') := True;
      FData.ValueByName('TotalRounding') := 0;
      FData.ValueByName('VATFromAbovePrecision') := 0;
    end;
// kolekce pro øádky faktury
    FRowsCollection := FData.Value[dmCommon.IndexByName(FData, 'Rows')];
// 1. øádek
    FRow:= AbraOLE.CreateValues('@IssuedInvoiceRow');
    FRow.ValueByName('Division_ID') := '1000000101';
    FRow.ValueByName('RowType') := '0';
    if FieldByName('Perioda').AsInteger = 1 then
      FRow.ValueByName('Text') := Format('Fakturujeme Vám za období od 1.%d.%d do %d.%d.%d',
       [aseMesic.Value, aseRok.Value, DayOfTheMonth(EndOfAMonth(aseRok.Value, aseMesic.Value)), aseMesic.Value, aseRok.Value]);
    FRowsCollection.Add(FRow);

// ============  smlouvy zákazníka - další øádky faktury se vytvoøí z qrSmlouvy
    while not EOF do begin  // qrSmlouva
      CenaTarifu := 0;
      PausalVoIP := 0;
      HovorneVoIP := 0;
// je-li datum aktivace menší než datum fakturace, vybere se prvni platba hotovì a fakturuje se pak celý mìsíc, jinak
// se platí jen èást mìsíce od data spuštìní
      DatumSpusteni := FieldByName('AktivniOd').AsDateTime;
      DatumUkonceni := FieldByName('AktivniDo').AsDateTime;
      Redukce := 1;
      PrvniFakturace := False;
      if DatumSpusteni >= FieldByName('FakturovatOd').AsDateTime then with qrAbra do begin
// už je nìjaká faktura ?
        Close;
        SQLStr := 'SELECT COUNT(*) FROM IssuedInvoices'
        + ' WHERE DocQueue_ID = ' + Ap + globalAA['abraIiDocQueue_Id'] + Ap
        + ' AND VarSymbol = ' + Ap + Cells[1, Radek] + Ap;
        SQL.Text := SQLStr;
        Open;
// zákazníkovi se ještì vùbec nefakturovalo
        PrvniFakturace := Fields[0].AsInteger = 0;
      end;  // with qrAbra
// ještì podle data aktivace (po pauze se fakturuje znovu)
      if not PrvniFakturace then
        PrvniFakturace := (MonthOf(DatumSpusteni) = MonthOf(deDatumDokladu.Date))
         and (YearOf(DatumSpusteni) = YearOf(deDatumDokladu.Date));
// redukce ceny pøipojení
      if PrvniFakturace then
        Redukce := ((YearOf(deDatumDokladu.Date) - YearOf(DatumSpusteni)) * 12          // rozdíl let * 12
          + MonthOf(deDatumDokladu.Date) - MonthOf(DatumSpusteni)                       // + rozdíl mìsícù + pomìrná èást 1. mìsíce
           + (DaysInMonth(DatumSpusteni) - DayOf(DatumSpusteni) + 1) / DaysInMonth(DatumSpusteni));
// poslední fakturace
      PosledniFakturace := (MonthOf(DatumUkonceni) = MonthOf(deDatumDokladu.Date))
        and (YearOf(DatumUkonceni) = YearOf(deDatumDokladu.Date));
      if PosledniFakturace then
        Redukce := DayOf(DatumUkonceni) / DaysInMonth(DatumUkonceni);             // pomìrná èást mìsíce
// datum spuštìní i ukonèení je ve stejném mìsíci
      if PrvniFakturace and PosledniFakturace then
        Redukce := (DayOf(DatumUkonceni) - DayOf(DatumSpusteni) + 1) / DaysInMonth(DatumUkonceni); // pomìrná èást mìsíce
// pro VoIP
      SmlouvaVoIP := Copy(qrSmlouva.FieldByName('Tarif').AsString, 1, 2) = 'EP';
      HovorneVoIP := 0;
      CenaTarifu := 0;
// paušál a hovorné
      if cbSVoIP.Checked and SmlouvaVoIP then with qrVoIP do begin
        Description := Description + 'VoIP, ';
        Close;
        SQLStr := 'SELECT SUM(Amount) FROM VoIP.Invoices_flat'
        + ' WHERE Num = ' + qrSmlouva.FieldByName('Smlouva').AsString
        + ' AND Year = ' + aseRok.Text
        + ' AND Month = ' + aseMesic.Text;
        SQL.Text := SQLStr;
        Open;
        HovorneVoIP := (100 + globalAA['abraVatRate'])/100 * Fields[0].AsFloat;
        Close;
      end;
// 24.1.2017 zakázky pro ÈTÚ - mohou být rùzné podle smlouvy
      with qrAbra do begin
        Close;
        BusOrderCode := qrSmlouva.FieldByName('CTU').AsString;
// kódy z tabulky contracts se musí pøedìlat
        Speed := Copy(BusOrderCode, Pos('_', BusOrderCode), 2);
        if Pos('WIFI', BusOrderCode) > 0 then BusOrderCode := 'W' + Speed
        else if Pos('FTTH', BusOrderCode) > 0 then BusOrderCode := 'A' + Speed
        else if Pos('FTTB', BusOrderCode) > 0 then BusOrderCode := 'B' + Speed
        else if Pos('PON', BusOrderCode) > 0 then BusOrderCode := 'P' + Speed
        else BusOrderCode := '1';
        SQLStr := 'SELECT Id FROM BusOrders'
        + ' WHERE Code = ' + Ap + BusOrderCode + Ap;
        SQL.Text := SQLStr;
        Open;
        BusOrder_Id := Fields[0].AsString;
        Close;
      end;
// cena za tarif
      if (FieldByName('Tarifni').AsInteger = 1) then begin
        CenaTarifu := qrSmlouva.FieldByName('Cena').AsFloat;
        if not SmlouvaVoIP then begin                    // pøipojení k Internetu
          FRow:= AbraOLE.CreateValues('@IssuedInvoiceRow');
          FRow.ValueByName('BusOrder_ID') := BusOrder_Id;
          if qrSmlouva.FieldByName('Typ').AsString = 'TvContract' then FRow.ValueByName('BusOrder_ID') := '1700000101';
          FRow.ValueByName('BusTransaction_ID') := BusTransaction_Id;
          FRow.ValueByName('Division_ID') := '1000000101';
          FRow.ValueByName('VATRate_ID') := globalAA['abraVatRate_Id'];
          FRow.ValueByName('VATRate') := IntToStr(globalAA['abraVatRate']);
          FRow.ValueByName('IncomeType_ID') := '2000000000';
          FRow.ValueByName('RowType') := '1';
          FRow.ValueByName('Text') :=
            Format('podle smlouvy  %s  službu  %s', [FieldByName('Smlouva').AsString, FieldByName('Text').AsString]);
          FRow.ValueByName('TotalPrice') := Format('%f', [CenaTarifu * Redukce]);
          FRowsCollection.Add(FRow);
          if isDRC then begin                                              // 19.10.2016
            FRow.ValueByName('VATIndex_ID') := globalAA['abraDrcVatIndex_Id'];
            FRow.ValueByName('DRCArticle_ID') := globalAA['abraDrcArticle_Id'];          // typ plnìní 21
            FRow.ValueByName('VATMode') := 1;
            FRow.ValueByName('TotalPrice') := Format('%f', [FieldByName('Cena').AsFloat * Redukce / 1.21]);
          end else FRow.ValueByName('VATIndex_ID') := globalAA['abraVatIndex_Id'];
        end else begin                                   // platby za VoIP
          if HovorneVoIP > 0 then begin                                      // hovorné
            FRow:= AbraOLE.CreateValues('@IssuedInvoiceRow');
            FRow.ValueByName('BusOrder_ID') := '1500000101';
            FRow.ValueByName('BusTransaction_ID') := BusTransaction_Id;
            FRow.ValueByName('Division_ID') := '1000000101';
            FRow.ValueByName('VATRate_ID') := globalAA['abraVatRate_Id'];
            FRow.ValueByName('VATIndex_ID') := globalAA['abraVatIndex_Id'];
            FRow.ValueByName('VATRate') := IntToStr(globalAA['abraVatRate']);
            FRow.ValueByName('IncomeType_ID') := '2000000000';
            FRow.ValueByName('RowType') := '1';
            FRow.ValueByName('Text') := 'hovorné VoIP';
            FRow.ValueByName('TotalPrice') := Format('%f', [HovorneVoIP]);
            FRowsCollection.Add(FRow);
            HovorneVoIP := 0;
          end;
          FRow:= AbraOLE.CreateValues('@IssuedInvoiceRow');             // paušál
          FRow.ValueByName('BusOrder_ID') := '2000000101';
          FRow.ValueByName('BusTransaction_ID') := BusTransaction_Id;
          FRow.ValueByName('Division_ID') := '1000000101';
          FRow.ValueByName('VATRate_ID') := globalAA['abraVatRate_Id'];
          FRow.ValueByName('VATIndex_ID') := globalAA['abraVatIndex_Id'];
          FRow.ValueByName('VATRate') := IntToStr(globalAA['abraVatRate']);
          FRow.ValueByName('IncomeType_ID') := '2000000000';
          FRow.ValueByName('RowType') := '1';
          FRow.ValueByName('Text') := Format('podle smlouvy  %s  mìsíèní platbu VoIP %s',
           [FieldByName('Smlouva').AsString, FieldByName('Text').AsString]);
          FRow.ValueByName('TotalPrice') := Format('%f', [CenaTarifu * Redukce]);
          FRowsCollection.Add(FRow);
        end;  // tarif VoIP
// nìco jiného než tarif
      end else begin
        FRow:= AbraOLE.CreateValues('@IssuedInvoiceRow');
        FRow.ValueByName('BusOrder_ID') := BusOrder_Id;
        if qrSmlouva.FieldByName('Typ').AsString = 'TvContract' then FRow.ValueByName('BusOrder_ID') := '1700000101';
        FRow.ValueByName('BusTransaction_ID') := BusTransaction_Id;
        FRow.ValueByName('Division_ID') := '1000000101';
        if Pos('auce', FieldByName('Text').AsString) > 0 then FRow.ValueByName('IncomeType_ID') := '1000000101'    // kauce
        else FRow.ValueByName('IncomeType_ID') := '2000000000';
        FRow.ValueByName('RowType') := '1';
        FRow.ValueByName('Text') := FieldByName('Text').AsString;
        FRow.ValueByName('TotalPrice') := Format('%f', [FieldByName('Cena').AsFloat * Redukce]);
        if FieldByName('DPH').AsString = '21%' then begin                   // 4.1.2013
          FRow.ValueByName('VATRate_ID') := globalAA['abraVatRate_Id'];
          if isDRC then begin                                              // 19.10.2016
            FRow.ValueByName('VATIndex_ID') := globalAA['abraDrcVatIndex_Id'];
            FRow.ValueByName('DRCArticle_ID') := globalAA['abraDrcArticle_Id'];          // typ plnìní 21
            FRow.ValueByName('VATMode') := 1;
            FRow.ValueByName('TotalPrice') := Format('%f', [FieldByName('Cena').AsFloat * Redukce / 1.21]);
          end else FRow.ValueByName('VATIndex_ID') := globalAA['abraVatIndex_Id'];
          FRow.ValueByName('VATRate') := IntToStr(globalAA['abraVatRate']);
        end else begin
          FRow.ValueByName('VATRate_ID') := '00000X0000';
          FRow.ValueByName('VATIndex_ID') := '7000000000';
          FRow.ValueByName('VATRate') := '0';
          if Pos('dar ', FieldByName('Text').AsString) > 0 then FRow.ValueByName('IncomeType_ID') := '3000000000';  // OS
        end;
        FRowsCollection.Add(FRow);
      end; // if tarif else
      Next;   //  qrSmlouva
    end;  //  while not qrSmlouva.EOF
// pøípadnì øádek s poštovným
    if (FieldByName('Posilani').AsString = 'Poštou')
     or (FieldByName('Posilani').AsString = 'Se složenkou') then begin    // pošta, složenka
      FRow:= AbraOLE.CreateValues('@IssuedInvoiceRow');
      FRow.ValueByName('BusOrder_ID') := '1000000101';
      FRow.ValueByName('BusTransaction_ID') := BusTransaction_Id;
      FRow.ValueByName('Division_ID') := '1000000101';
      FRow.ValueByName('VATRate_ID') := globalAA['abraVatRate_Id'];
      FRow.ValueByName('VATIndex_ID') := globalAA['abraVatIndex_Id'];
      FRow.ValueByName('VATRate') := IntToStr(globalAA['abraVatRate']);
      FRow.ValueByName('IncomeType_ID') := '2000000000';
      FRow.ValueByName('RowType') := '1';
      FRow.ValueByName('Text') := 'manipulaèní poplatek';
      FRow.ValueByName('TotalPrice') := '62';
      FRowsCollection.Add(FRow);
    end;
    Description := Description + Cells[1, Radek];        // VS
    FData.ValueByName('Description') := Description;
    if isDebugMode then dmCommon.Zprava('Data faktury pøipravena');
// vytvoøení faktury
    try
      ID := FObject.CreateNewFromValues(FData);
      FData := FObject.GetValues(ID);
      FCena := double(FData.Value[dmCommon.IndexByName(FData, 'Amount')]);
      dmCommon.Zprava(Format('%s (%s): Vytvoøena faktura %s.', [Zakaznik, Cells[1, Radek], string(FData.Value[dmCommon.IndexByName(FData, 'DisplayName')])]));
      Ints[0, Radek] := 0;
      Cells[2, Radek] := string(FData.Value[dmCommon.IndexByName(FData, 'OrdNumber')]);    // faktura
      Cells[3, Radek] := IntToStr(Round(FCena));                                                 // èástka
      Cells[4, Radek] := Zakaznik;                                                               // jméno
    except
      on E: Exception do begin
        dmCommon.Zprava(Format('%s (%s): Chyba v Abøe - %s', [Zakaznik, Cells[1, Radek], E.Message]));
        if Application.MessageBox(PChar(E.Message + ^M + 'Pokraèovat?'), 'Chyba',
         MB_YESNO + MB_ICONQUESTION) = IDNO then Prerusit := True;
      end;
    end;  // try
    Close;   // qrSmlouva
  end;  // with
end;  // procedury FakturaAbra


end.

