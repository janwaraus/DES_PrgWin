// DES fakturuje najednou v�echny smlouvy z�kazn�ka, internet, VoIP i IPTV s datem vystaven� i pln�n� posledn� den v m�s�ci.
// Z�kazn�ci nebo smlouvy k fakturaci jsou v asgMain, faktury se vytvo�� v cyklu p�es ��dky.
// 23.1.17 V�kazy pro �T� vy�aduj� d�len� podle typu u�ivatele, technologie a rychlosti - do Abry byly p�id�ny 3 obchodn� p��pady
// - VO (velkoobchod - DRC) , F a P (fyzick� a pr�vnick� osoba) a 24 zak�zek (W - WiFi, B - FFTB, A - FTTH AON, a P - FTTH PON,
// ka�d� s �esti r�zn�mi rychlostmi). Ka�d�mu ��dku faktury bude p�i�azen jeden obchodn� p��pad a jedna zak�zka.

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
    dmCommon.Zprava(Format('Po�et faktur k vygenerov�n�: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));

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
    dmCommon.Zprava('Generov�n� faktur ukon�eno');
  end;  //  with fmMain
end;

// ------------------------------------------------------------------------------------------------


procedure TdmFaktura.FakturaAbra(Radek: integer);
// pro m�s�c a rok zadan� v aseMesic a aseRok vytvo�� fakturu za p�ipojen� k internetu a za VoIP
// 27.1.17 cel� p�ehledn�ji
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
  // faktura se nebude vytv��et, je-li ve smlouv�ch tak� fakturovan� VoIP a nen� za�rktnut� vytv��en� VoIP faktur,
  // nebo nen� za�rktl� vytv��en� internet faktur a faktura nem� VoIP (tedy je internet)
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

// nen� v�c z�kazn�k� pro jeden VS ? 31.1.2017
    Close;
    SQLStr := 'SELECT COUNT(*) FROM customers'
    + ' WHERE Variable_Symbol = ' + Ap + Cells[1, Radek] + Ap;
    SQL.Text := SQLStr;
    Open;
    if Fields[0].AsInteger > 1 then begin
      dmCommon.Zprava(Format('Variabiln� symbol %s m� v�ce z�kazn�k�.', [Cells[1, Radek]]));
      Close;
      Exit;
    end;

// je p�enesen� da�ov� povinnost ?  19.10.2016 - bude pro celou fakturu, ne pro jednotliv� smlouvy
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

// vyhled�n� �daj� o smlouv�ch
    Close;
    SQLStr := 'SELECT VS, AbraKod, Typ, Smlouva, AktivniOd, AktivniDo, FakturovatOd, Tarif, Posilani, Perioda, Text, Cena, DPH, Tarifni, CTU'
    + ' FROM ' + fiInvoiceView
    + ' WHERE VS = ' + Ap + Cells[1, Radek] + Ap
    + ' ORDER BY Smlouva, Tarifni DESC';
    SQL.Text := SQLStr;
    Open;
// p�i lecjak� chyb� v datab�zi (nap�. Tariff_Id je NULL) konec
    if RecordCount = 0 then begin
      dmCommon.Zprava(Format('Pro variabiln� symbol %s nen� co fakturovat.', [FieldByName('VS').AsString]));
      Close;
      Exit;
    end;
    if FieldByName('AbraKod').AsString = '' then begin
      dmCommon.Zprava(Format('Smlouva %s: z�kazn�k nem� k�d Abry.', [FieldByName('Smlouva').AsString]));
      Close;
      Exit;
    end;

    with qrAbra do begin
// kontrola k�du firmy, p�i chyb� konec
      Close;
      SQLStr := 'SELECT F.ID, F.Name, F.OrgIdentNumber, FO.Id FROM Firms F, FirmOffices FO'
      + ' WHERE Code = ' + Ap + qrSmlouva.FieldByName('AbraKod').AsString + Ap
      + ' AND F.Firm_ID IS NULL'         // bez p�edk�
      + ' AND F.Hidden = ''N'''
      + ' AND FO.Parent_Id = F.Id'
      + ' ORDER BY F.ID DESC';
      SQL.Text := SQLStr;
      Open;
      if RecordCount = 0 then begin
        dmCommon.Zprava(Format('Smlouva %s: z�kazn�k s k�dem %s nen� v adres��i Abry.',
         [qrSmlouva.FieldByName('Smlouva').AsString, qrSmlouva.FieldByName('AbraKod').AsString]));
        Exit;
      end else begin
        Firm_Id := Fields[0].AsString;
        Zakaznik := Fields[1].AsString;
        IC := Trim(Fields[2].AsString);
        FirmOffice_Id := Fields[3].AsString; //
      end;
// 24.1.2017 obchodn� p��pady pro �T� - plat� pro celou faktury
      if isDRC then BusTransactionCode := 'VO'                  // velkoobchod (s DRC)
      else if IC = '' then BusTransactionCode := 'F'         //  fyzick� osoba
      else BusTransactionCode := 'P';                         //  pr�vnick� osoba
      with qrAbra do begin
        Close;
        SQLStr := 'SELECT Id FROM BusTransactions'
        + ' WHERE Code = ' + Ap + BusTransactionCode + Ap;
        SQL.Text := SQLStr;
        Open;
        BusTransaction_Id := Fields[0].AsString;
        Close;
      end;

// kontrola posledn� faktury
      Close;
      SQLStr := 'SELECT OrdNumber, DocDate$DATE, VATDate$DATE, Amount FROM IssuedInvoices'
      + ' WHERE VarSymbol = ' + Ap + qrSmlouva.FieldByName('VS').AsString + Ap
      + ' AND VATDate$DATE >= ' + FloatToStr(Trunc(StartOfAMonth(aseRok.Value, aseMesic.Value)))
      + ' AND VATDate$DATE <= ' + FloatToStr(Trunc(EndOfAMonth(aseRok.Value, aseMesic.Value)));
      SQLStr := SQLStr + ' AND DocQueue_ID = ' + Ap + globalAA['abraIiDocQueue_Id'] + Ap;
      SQLStr := SQLStr + ' ORDER BY OrdNumber DESC';
      SQL.Text := SQLStr;
      Open;
      if isDebugMode then dmCommon.Zprava('Vyhled�na data z Abry');
      if RecordCount > 0 then begin
        dmCommon.Zprava(Format('%s (%s): %d. faktura se stejn�m datem.',
         [Zakaznik, Cells[1, Radek], RecordCount + 1]));
        Dotaz := Application.MessageBox(PChar(Format('Pro z�kazn�ka "%s" existuje faktura %s-%s s datem %s na ��stku %s K�. M� se vytvo�it dal��?',
         [Zakaznik, globalAA['invoiceDocQueueCode'], FieldByName('OrdNumber').AsString, DateToStr(FieldByName('DocDate$DATE').AsFloat), FieldByName('Amount').AsString])),
          'Pozor', MB_ICONQUESTION + MB_YESNOCANCEL + MB_DEFBUTTON1);
        if Dotaz = IDNO then begin
          dmCommon.Zprava('Ru�n� z�sah - faktura nevytvo�ena.');
          Exit;
        end else if Dotaz = IDCANCEL then begin
          dmCommon.Zprava('Ru�n� z�sah - program ukon�en.');
          Prerusit := True;
          Exit;
        end else dmCommon.Zprava('Ru�n� z�sah - faktura se vytvo��.');
      end;  // RecordCount > 0
    end;  // with qrAbra

    Description := Format('p�ipojen� %d/%d, ', [aseMesic.Value, aseRok.Value-2000]);
// vytvo�� se hlavi�ka faktury

    boAA := TAArray.Create;

    //FObject:= AbraOLE.CreateObject('@IssuedInvoice');
    //FData:= AbraOLE.CreateValues('@IssuedInvoice');
    //FObject.PrefillValues(FData);
    boAA['DocQueue_ID'] := VarToStr(globalAA['abraIiDocQueue_Id']);
    boAA['Period_ID'] := VarToStr(globalAA['abraIiPeriod_Id']);
    boAA['DocDate$DATE'] := Floor(deDatumDokladu.Date);
    boAA['AccDate$DATE'] := Floor(deDatumDokladu.Date);
    boAA['VATDate$DATE'] := Floor(deDatumPlneni.Date);
    boAA['CreatedBy_ID'] := MyUser_Id; // TODO bylo jen User_Id ale to bylo pr�zdn�, bez zalogov�n�
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
    boAA['TotalRounding'] := 259;                // zaokrouhlen� na koruny dol�
    if isDRC then begin
      boAA['IsReverseChargeDeclared'] := True;
      boAA['VATFromAbovePrecision'] := 0;
      boAA['TotalRounding'] := 0;
    end;
// kolekce pro ��dky faktury
    //FRowsCollection := FData.Value[dmCommon.IndexByName(FData, 'Rows')];
// 1. ��dek

    //FRow:= AbraOLE.CreateValues('@IssuedInvoiceRow');
    boRowAA := boAA.addRow();
    boRowAA['RowType'] := '0';
    boRowAA['Division_ID'] := '1000000101';
    if FieldByName('Perioda').AsInteger = 1 then
      boRowAA['Text'] := Format('Fakturujeme V�m za obdob� od 1.%d.%d do %d.%d.%d',
       [aseMesic.Value, aseRok.Value, DayOfTheMonth(EndOfAMonth(aseRok.Value, aseMesic.Value)), aseMesic.Value, aseRok.Value]);

// ============  smlouvy z�kazn�ka - dal�� ��dky faktury se vytvo�� z qrSmlouvy
    while not EOF do begin  // qrSmlouva
      CenaTarifu := 0;
      PausalVoIP := 0;
      HovorneVoIP := 0;
// je-li datum aktivace men�� ne� datum fakturace, vybere se prvni platba hotov� a fakturuje se pak cel� m�s�c, jinak
// se plat� jen ��st m�s�ce od data spu�t�n�
      DatumSpusteni := FieldByName('AktivniOd').AsDateTime;
      DatumUkonceni := FieldByName('AktivniDo').AsDateTime;
      Redukce := 1;
      PrvniFakturace := False;

      if DatumSpusteni >= FieldByName('FakturovatOd').AsDateTime then
      with qrAbra do begin
        // u� je n�jak� faktura ?
        Close;
        SQLStr := 'SELECT COUNT(*) FROM IssuedInvoices'
        + ' WHERE DocQueue_ID = ' + Ap + globalAA['abraIiDocQueue_Id'] + Ap
        + ' AND VarSymbol = ' + Ap + Cells[1, Radek] + Ap;
        SQL.Text := SQLStr;
        Open;
// z�kazn�kovi se je�t� v�bec nefakturovalo
        PrvniFakturace := Fields[0].AsInteger = 0;
      end;  // with qrAbra

// je�t� podle data aktivace (po pauze se fakturuje znovu)
      if not PrvniFakturace then
        PrvniFakturace := (MonthOf(DatumSpusteni) = MonthOf(deDatumDokladu.Date))
         and (YearOf(DatumSpusteni) = YearOf(deDatumDokladu.Date));
// redukce ceny p�ipojen�
      if PrvniFakturace then
        Redukce := ((YearOf(deDatumDokladu.Date) - YearOf(DatumSpusteni)) * 12          // rozd�l let * 12
          + MonthOf(deDatumDokladu.Date) - MonthOf(DatumSpusteni)                       // + rozd�l m�s�c� + pom�rn� ��st 1. m�s�ce
           + (DaysInMonth(DatumSpusteni) - DayOf(DatumSpusteni) + 1) / DaysInMonth(DatumSpusteni));
// posledn� fakturace
      PosledniFakturace := (MonthOf(DatumUkonceni) = MonthOf(deDatumDokladu.Date))
        and (YearOf(DatumUkonceni) = YearOf(deDatumDokladu.Date));
      if PosledniFakturace then
        Redukce := DayOf(DatumUkonceni) / DaysInMonth(DatumUkonceni);             // pom�rn� ��st m�s�ce
// datum spu�t�n� i ukon�en� je ve stejn�m m�s�ci
      if PrvniFakturace and PosledniFakturace then
        Redukce := (DayOf(DatumUkonceni) - DayOf(DatumSpusteni) + 1) / DaysInMonth(DatumUkonceni); // pom�rn� ��st m�s�ce
// pro VoIP
      SmlouvaVoIP := Copy(qrSmlouva.FieldByName('Tarif').AsString, 1, 2) = 'EP';
      HovorneVoIP := 0;
      // pau��l (? TODO) a hovorn�

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
// 24.1.2017 zak�zky pro �T� - mohou b�t r�zn� podle smlouvy
      with qrAbra do begin
        Close;
        BusOrderCode := qrSmlouva.FieldByName('CTU').AsString;
// k�dy z tabulky contracts se mus� p�ed�lat
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
        if not SmlouvaVoIP then begin                    // p�ipojen� k Internetu
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
            Format('podle smlouvy  %s  slu�bu  %s', [FieldByName('Smlouva').AsString, FieldByName('Text').AsString]);
          boRowAA['TotalPrice'] := Format('%f', [CenaTarifu * Redukce]);
          //FRowsCollection.Add(FRow);
          if isDRC then begin                                              // 19.10.2016
            boRowAA['VATIndex_ID'] := VarToStr(globalAA['abraDrcVatIndex_Id']);
            boRowAA['DRCArticle_ID'] := VarToStr(globalAA['abraDrcArticle_Id']);          // typ pln�n� 21
            boRowAA['VATMode'] := 1;
            boRowAA['TotalPrice'] := Format('%f', [FieldByName('Cena').AsFloat * Redukce / 1.21]);
          end else boRowAA['VATIndex_ID'] := VarToStr(globalAA['abraVatIndex_Id']);
        end else begin                                   // platby za VoIP
          if HovorneVoIP > 0 then begin                                      // hovorn�
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
            boRowAA['Text'] := 'hovorn� VoIP';
            boRowAA['TotalPrice'] := Format('%f', [HovorneVoIP]);

            HovorneVoIP := 0;
          end;
          //FRow:= AbraOLE.CreateValues('@IssuedInvoiceRow');             // pau��l
          boRowAA := boAA.addRow();
          boRowAA['RowType'] := '1';
          boRowAA['BusOrder_ID'] := '2000000101';
          boRowAA['BusTransaction_ID'] := BusTransaction_Id;
          boRowAA['Division_ID'] := '1000000101';
          boRowAA['VATRate_ID'] := VarToStr(globalAA['abraVatRate_Id']);
          boRowAA['VATIndex_ID'] := VarToStr(globalAA['abraVatIndex_Id']);
          //boRowAA['VATRate'] := IntToStr(globalAA['abraVatRate']);
          boRowAA['IncomeType_ID'] := '2000000000';
          boRowAA['Text'] := Format('podle smlouvy  %s  m�s��n� platbu VoIP %s',
           [FieldByName('Smlouva').AsString, FieldByName('Text').AsString]);
          boRowAA['TotalPrice'] := Format('%f', [CenaTarifu * Redukce]);

        end;  // tarif VoIP
// n�co jin�ho ne� tarif
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
            boRowAA['DRCArticle_ID'] := VarToStr(globalAA['abraDrcArticle_Id']);          // typ pln�n� 21
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
// p��padn� ��dek s po�tovn�m
    if (FieldByName('Posilani').AsString = 'Po�tou')
     or (FieldByName('Posilani').AsString = 'Se slo�enkou') then begin    // po�ta, slo�enka
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
      boRowAA['Text'] := 'manipula�n� poplatek';
      boRowAA['TotalPrice'] := '62';
    end;
    Description := Description + Cells[1, Radek];        // v Cells[1, Radek] je VS
    boAA['Description'] := Description;
    if isDebugMode then dmCommon.Zprava('Data faktury p�ipravena');

  cas02 := Now;
  dmCommon.Zprava(debugRozdilCasu(cas01, cas02, ' - �as p��pravy dat fa'));

// vytvo�en� faktury
    try
      abraWebApiResponse := DesU.abraBoCreateWebApi(boAA, 'issuedinvoice');
      if abraWebApiResponse.isOk then begin
        abraResponseSO := SO(abraWebApiResponse.Messg);
        dmCommon.Zprava(Format('%s (%s): Vytvo�ena faktura %s.', [Zakaznik, Cells[1, Radek], abraResponseSO.S['displayname']]));
        Ints[0, Radek] := 0;
        Cells[2, Radek] := abraResponseSO.S['ordnumber']; // faktura
        Cells[3, Radek] := abraResponseSO.S['amount'];                                                // ��stka
        Cells[4, Radek] := Zakaznik;
      end else begin
        dmCommon.Zprava(Format('%s (%s): Chyba %s: %s', [Zakaznik, Cells[1, Radek], abraWebApiResponse.Code, abraWebApiResponse.Messg]));
        if Dialogs.MessageDlg( '(' + abraWebApiResponse.Code + ') '
           + abraWebApiResponse.Messg + sLineBreak + 'Pokra�ovat?',
           mtConfirmation, [mbYes, mbNo], 0 ) = mrNo then Prerusit := True;
      end;
    except on E: exception do
      begin
        Application.MessageBox(PChar('Problem ' + ^M + E.Message), 'Vytvo�en� fa');
      end;
    end;


    Close;   // qrSmlouva
  end;  // with

  cas03 := Now;
  dmCommon.Zprava(debugRozdilCasu(cas02, cas03, ' - �as zaps�n� fa do ABRA'));


end;  // procedury FakturaAbraAA



end.

