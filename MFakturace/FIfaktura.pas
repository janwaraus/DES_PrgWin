// DES fakturuje najednou všechny smlouvy zákazníka, internet, VoIP i IPTV s datem vystavení i plnìní poslední den v mìsíci.
// Zákazníci nebo smlouvy k fakturaci jsou v asgMain, faktury se vytvoøí v cyklu pøes øádky.
// 23.1.17 Výkazy pro ÈTÚ vyžadují dìlení podle typu uživatele, technologie a rychlosti - do Abry byly pøidány 3 obchodní pøípady
// - VO (velkoobchod - DRC) , F a P (fyzická a právnická osoba) a 24 zakázek (W - WiFi, B - FFTB, A - FTTH AON, a P - FTTH PON,
// každá s šesti rùznými rychlostmi). Každému øádku faktury bude pøiøazen jeden obchodní pøípad a jedna zakázka.

unit FIfaktura;

interface

uses
  Windows, Messages, Dialogs, Classes, Forms, Controls, SysUtils, DateUtils, Variants, ComObj, Math,
  Superobject, AArray,
  FImain;

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
    dmCommon.Zprava(Format('Poèet faktur k vygenerování: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));

    Screen.Cursor := crHourGlass;
    apbProgress.Position := 0;
    apbProgress.Visible := True;

    for Radek := 1 to RowCount-1 do begin
      Row := Radek;
      apbProgress.Position := Round(100 * Radek / RowCount-1);
      Application.ProcessMessages;
      if Prerusit then begin
        Prerusit := False;
        btVytvorit.Enabled := True;
        Break;
      end;
      if Ints[0, Radek] = 1 then FakturaAbra(Radek);
    end;

  finally
    apbProgress.Position := 0;
    apbProgress.Visible := False;
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
  //FObject,
  //FData,
  //FRow,
  //FRowsCollection: variant;
  Zakaznik,
  Description,
  SQLStr: AnsiString;

  boAA, boRowAA: TAArray;
  abraResponseJSON,
  newID: string;
  abraResponseSO : ISuperObject;
  abraWebApiResponse : TDesResult;

  cas01, cas02, cas03: double;

begin
  cas01 := Now;

  with fmMain, qrSmlouva, asgMain do begin
  // faktura se nebude vytváøet, je-li ve smlouvách také fakturovaný VoIP a není zašrktnuté vytváøení VoIP faktur,
  // nebo není zašrktlé vytváøení internet faktur a faktura nemé VoIP (tedy je internet)
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
        FirmOffice_Id := Fields[3].AsString; //
      end;
// 24.1.2017 obchodní pøípady pro ÈTÚ - platí pro celou faktury
      if isDRC then BusTransactionCode := 'VO'                  // velkoobchod (s DRC)
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

    boAA := TAArray.Create;

    //FObject:= AbraOLE.CreateObject('@IssuedInvoice');
    //FData:= AbraOLE.CreateValues('@IssuedInvoice');
    //FObject.PrefillValues(FData);
    boAA['DocQueue_ID'] := VarToStr(globalAA['abraIiDocQueue_Id']);
    boAA['Period_ID'] := VarToStr(globalAA['abraIiPeriod_Id']);
    boAA['DocDate$DATE'] := Floor(deDatumDokladu.Date);
    boAA['AccDate$DATE'] := Floor(deDatumDokladu.Date);
    boAA['VATDate$DATE'] := Floor(deDatumPlneni.Date);
    boAA['CreatedBy_ID'] := MyUser_Id; // TODO bylo jen User_Id ale to bylo prázdné, bez zalogování
    boAA['Firm_ID'] := Firm_Id;
    boAA['FirmOffice_ID'] := FirmOffice_Id;
    boAA['Address_ID'] := MyAddress_Id; // FA CONST
    boAA['BankAccount_ID'] := MyAccount_Id;
    boAA['ConstSymbol_ID'] := '0000308000';
    boAA['VarSymbol'] := FieldByName('VS').AsString;
    boAA['TransportationType_ID'] := '1000000101'; // FA CONST
    boAA['DueTerm'] := aedSplatnost.Text;
    boAA['PaymentType_ID'] := MyPayment_Id;  // FA CONST
    boAA['PricesWithVAT'] := True;
    boAA['VATFromAbovePrecision'] := 6;
    boAA['TotalRounding'] := 259;                // zaokrouhlení na koruny dolù
    if isDRC then begin
      boAA['IsReverseChargeDeclared'] := True;
      boAA['VATFromAbovePrecision'] := 0;
      boAA['TotalRounding'] := 0;
    end;
// kolekce pro øádky faktury
    //FRowsCollection := FData.Value[dmCommon.IndexByName(FData, 'Rows')];
// 1. øádek

    //FRow:= AbraOLE.CreateValues('@IssuedInvoiceRow');
    boRowAA := boAA.addRow();
    boRowAA['RowType'] := '0';
    boRowAA['Division_ID'] := '1000000101';
    if FieldByName('Perioda').AsInteger = 1 then
      boRowAA['Text'] := Format('Fakturujeme Vám za období od 1.%d.%d do %d.%d.%d',
       [aseMesic.Value, aseRok.Value, DayOfTheMonth(EndOfAMonth(aseRok.Value, aseMesic.Value)), aseMesic.Value, aseRok.Value]);

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

      if DatumSpusteni >= FieldByName('FakturovatOd').AsDateTime then
      with qrAbra do begin
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
      // paušál (? TODO) a hovorné

      if cbSVoIP.Checked and SmlouvaVoIP then with qrVoIP do begin // TODO je treba hlidat cbSVoIP.Checked?
        Description := Description + 'VoIP, ';
        HovorneVoIP := 3333; // TODO smazat
        {  TODO dodelat VoIP nepripojuje se
        Close;
        SQLStr := 'SELECT SUM(Amount) FROM VoIP.Invoices_flat'
        + ' WHERE Num = ' + qrSmlouva.FieldByName('Smlouva').AsString
        + ' AND Year = ' + aseRok.Text
        + ' AND Month = ' + aseMesic.Text;
        SQL.Text := SQLStr;
        Open;
        HovorneVoIP := (100 + globalAA['abraVatRate'])/100 * Fields[0].AsFloat;
        Close;
        }
      end;

      {  CTU, neni uz potreba
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
      }

// cena za tarif
      if (FieldByName('Tarifni').AsInteger = 1) then begin
        CenaTarifu := qrSmlouva.FieldByName('Cena').AsFloat;
        if not SmlouvaVoIP then begin                    // pøipojení k Internetu
          //FRow:= AbraOLE.CreateValues('@IssuedInvoiceRow');
          boRowAA := boAA.addRow();
          boRowAA['RowType'] := '1';
          // boRowAA['BusOrder_ID'] := BusOrder_Id; // CTU, neni uz potreba
          if qrSmlouva.FieldByName('Typ').AsString = 'TvContract' then
            boRowAA['BusOrder_ID'] := '1700000101';
          boRowAA['BusTransaction_ID'] := BusTransaction_Id;
          boRowAA['Division_ID'] := '1000000101';
          boRowAA['VATRate_ID'] := VarToStr(globalAA['abraVatRate_Id']);
          //boRowAA['VATRate'] := IntToStr(globalAA['abraVatRate']); // nesmi byt v requestu
          boRowAA['IncomeType_ID'] := '2000000000';
          boRowAA['Text'] :=
            Format('podle smlouvy  %s  službu  %s', [FieldByName('Smlouva').AsString, FieldByName('Text').AsString]);
          boRowAA['TotalPrice'] := Format('%f', [CenaTarifu * Redukce]);
          //FRowsCollection.Add(FRow);
          if isDRC then begin                                              // 19.10.2016
            boRowAA['VATIndex_ID'] := VarToStr(globalAA['abraDrcVatIndex_Id']);
            boRowAA['DRCArticle_ID'] := VarToStr(globalAA['abraDrcArticle_Id']);          // typ plnìní 21
            boRowAA['VATMode'] := 1;
            boRowAA['TotalPrice'] := Format('%f', [FieldByName('Cena').AsFloat * Redukce / 1.21]);
          end else boRowAA['VATIndex_ID'] := VarToStr(globalAA['abraVatIndex_Id']);
        end else begin                                   // platby za VoIP
          if HovorneVoIP > 0 then begin                                      // hovorné
            //FRow:= AbraOLE.CreateValues('@IssuedInvoiceRow');
            boRowAA := boAA.addRow();
            boRowAA['RowType'] := '1';
            boRowAA['BusOrder_ID'] := '1500000101';
            boRowAA['BusTransaction_ID'] := BusTransaction_Id;
            boRowAA['Division_ID'] := '1000000101';
            boRowAA['VATRate_ID'] := VarToStr(globalAA['abraVatRate_Id']);
            boRowAA['VATIndex_ID'] := VarToStr(globalAA['abraVatIndex_Id']);
            //boRowAA['VATRate'] := IntToStr(globalAA['abraVatRate']);
            boRowAA['IncomeType_ID'] := '2000000000';
            boRowAA['Text'] := 'hovorné VoIP';
            boRowAA['TotalPrice'] := Format('%f', [HovorneVoIP]);

            HovorneVoIP := 0;
          end;
          //FRow:= AbraOLE.CreateValues('@IssuedInvoiceRow');             // paušál
          boRowAA := boAA.addRow();
          boRowAA['RowType'] := '1';
          boRowAA['BusOrder_ID'] := '2000000101';
          boRowAA['BusTransaction_ID'] := BusTransaction_Id;
          boRowAA['Division_ID'] := '1000000101';
          boRowAA['VATRate_ID'] := VarToStr(globalAA['abraVatRate_Id']);
          boRowAA['VATIndex_ID'] := VarToStr(globalAA['abraVatIndex_Id']);
          //boRowAA['VATRate'] := IntToStr(globalAA['abraVatRate']);
          boRowAA['IncomeType_ID'] := '2000000000';
          boRowAA['Text'] := Format('podle smlouvy  %s  mìsíèní platbu VoIP %s',
           [FieldByName('Smlouva').AsString, FieldByName('Text').AsString]);
          boRowAA['TotalPrice'] := Format('%f', [CenaTarifu * Redukce]);

        end;  // tarif VoIP
// nìco jiného než tarif
      end else begin
        //FRow:= AbraOLE.CreateValues('@IssuedInvoiceRow');
        boRowAA := boAA.addRow();
        boRowAA['BusOrder_ID'] := BusOrder_Id;
        if qrSmlouva.FieldByName('Typ').AsString = 'TvContract' then
          boRowAA['BusOrder_ID'] := '1700000101';
        boRowAA['BusTransaction_ID'] := BusTransaction_Id;
        boRowAA['Division_ID'] := '1000000101';
        if Pos('auce', FieldByName('Text').AsString) > 0 then
          boRowAA['IncomeType_ID'] := '1000000101'    // kauce
        else
          boRowAA['IncomeType_ID'] := '2000000000';
        boRowAA['RowType'] := '1';
        boRowAA['Text'] := FieldByName('Text').AsString;
        boRowAA['TotalPrice'] := Format('%f', [FieldByName('Cena').AsFloat * Redukce]);
        if FieldByName('DPH').AsString = '21%' then begin                   // 4.1.2013
          boRowAA['VATRate_ID'] := VarToStr(globalAA['abraVatRate_Id']);
          if isDRC then begin                                              // 19.10.2016
            boRowAA['VATIndex_ID'] := VarToStr(globalAA['abraDrcVatIndex_Id']);
            boRowAA['DRCArticle_ID'] := VarToStr(globalAA['abraDrcArticle_Id']);          // typ plnìní 21
            boRowAA['VATMode'] := 1;
            boRowAA['TotalPrice'] := Format('%f', [FieldByName('Cena').AsFloat * Redukce / 1.21]);
          end else
            boRowAA['VATIndex_ID'] := VarToStr(globalAA['abraVatIndex_Id']);
          //boRowAA['VATRate'] := IntToStr(globalAA['abraVatRate']);
        end else begin
          boRowAA['VATRate_ID'] := '00000X0000';
          boRowAA['VATIndex_ID'] := '7000000000';
          //boRowAA['VATRate'] := '0';
          if Pos('dar ', FieldByName('Text').AsString) > 0 then
            boRowAA['IncomeType_ID'] := '3000000000';  // OS
        end;

      end; // if tarif else
      Next;   //  qrSmlouva
    end;  //  while not qrSmlouva.EOF
// pøípadnì øádek s poštovným
    if (FieldByName('Posilani').AsString = 'Poštou')
     or (FieldByName('Posilani').AsString = 'Se složenkou') then begin    // pošta, složenka
      //FRow:= AbraOLE.CreateValues('@IssuedInvoiceRow');
      boRowAA := boAA.addRow();
      boRowAA['BusOrder_ID'] := '1000000101';
      boRowAA['BusTransaction_ID'] := BusTransaction_Id;
      boRowAA['Division_ID'] := '1000000101';
      boRowAA['VATRate_ID'] := VarToStr(globalAA['abraVatRate_Id']);
      boRowAA['VATIndex_ID'] := VarToStr(globalAA['abraVatIndex_Id']);
      //boRowAA['VATRate'] := IntToStr(globalAA['abraVatRate']);
      boRowAA['IncomeType_ID'] := '2000000000';
      boRowAA['RowType'] := '1';
      boRowAA['Text'] := 'manipulaèní poplatek';
      boRowAA['TotalPrice'] := '62';
    end;
    Description := Description + Cells[1, Radek];        // v Cells[1, Radek] je VS
    boAA['Description'] := Description;
    if isDebugMode then dmCommon.Zprava('Data faktury pøipravena');

  cas02 := Now;
  dmCommon.Zprava(debugRozdilCasu(cas01, cas02, ' - èas pøípravy dat fa'));

// vytvoøení faktury
    try
      abraWebApiResponse := DesU.abraBoCreateWebApi(boAA, 'issuedinvoice');
      if abraWebApiResponse.isOk then begin
        abraResponseSO := SO(abraWebApiResponse.Messg);
        dmCommon.Zprava(Format('%s (%s): Vytvoøena faktura %s.', [Zakaznik, Cells[1, Radek], abraResponseSO.S['displayname']]));
        Ints[0, Radek] := 0;
        Cells[2, Radek] := abraResponseSO.S['ordnumber']; // faktura
        Cells[3, Radek] := abraResponseSO.S['amount'];                                                // èástka
        Cells[4, Radek] := Zakaznik;
      end else begin
        dmCommon.Zprava(Format('%s (%s): Chyba %s: %s', [Zakaznik, Cells[1, Radek], abraWebApiResponse.Code, abraWebApiResponse.Messg]));
        if Dialogs.MessageDlg( '(' + abraWebApiResponse.Code + ') '
           + abraWebApiResponse.Messg + sLineBreak + 'Pokraèovat?',
           mtConfirmation, [mbYes, mbNo], 0 ) = mrNo then Prerusit := True;
      end;
    except on E: exception do
      begin
        Application.MessageBox(PChar('Problem ' + ^M + E.Message), 'Vytvoøení fa');
      end;
    end;


    Close;   // qrSmlouva
  end;  // with

  cas03 := Now;
  dmCommon.Zprava(debugRozdilCasu(cas02, cas03, ' - èas zapsání fa do ABRA'));


end;  // procedury FakturaAbraAA



end.

