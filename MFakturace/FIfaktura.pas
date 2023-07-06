// DES fakturuje najednou všechny smlouvy zákazníka, internet, VoIP i IPTV s datem vystavení i plnìní poslední den v mìsíci.
// Zákazníci nebo smlouvy k fakturaci jsou v asgMain, faktury se vytvoøí v cyklu pøes øádky.
// 23.1.17 Výkazy pro ÈTÚ vyžadují dìlení podle typu uživatele, technologie a rychlosti - do Abry byly pøidány 3 obchodní pøípady
// - VO (velkoobchod - DRC) , F a P (fyzická a právnická osoba) a 24 zakázek (W - WiFi, B - FFTB, A - FTTH AON, a P - FTTH PON,
// každá s šesti rùznými rychlostmi). Každému øádku faktury bude pøiøazen jeden obchodní pøípad a jedna zakázka.

unit FIfaktura;

interface

uses
  Windows, Messages, Dialogs, Classes, Forms, Controls, SysUtils, DateUtils, StrUtils, Variants, ComObj, Math,
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

uses DesUtils, AbraEntities, DesInvoices;

// ------------------------------------------------------------------------------------------------

procedure TdmFaktura.VytvorFaktury;
var
  Radek: integer;
begin
  with fmMain, fmMain.asgMain do try
    fmMain.Zprava(Format('Poèet faktur k vygenerování: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));

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
    fmMain.Zprava('Generování faktur ukonèeno');
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
  // Speed,
  BusTransaction_Id,
  BusTransactionCode,
  OrgIdentNumber,
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

  FirmName,
  Description,
  CustomerVarSymbol,
  SQLStr: string;

  boRowAA: TAArray;
  abraResponseJSON,
  newID: string;
  abraResponseSO : ISuperObject;
  abraWebApiResponse : TDesResult;
  NewInvoice : TNewDesInvoiceAA;

  cas01, cas02, cas03: double;

begin
  cas01 := Now;

  with fmMain, qrSmlouva, asgMain do begin
    CustomerVarSymbol := Cells[1, Radek];

  // faktura se nebude vytváøet, je-li ve smlouvách také fakturovaný VoIP a není zašrktnuté vytváøení VoIP faktur,
  // nebo není zašrktlé vytváøení internet faktur a faktura nemá VoIP (tedy je internet)
  // HWTODO je tenhle check potreba?
  {
    Close;
    SQLStr := 'SELECT COUNT(*) FROM ' + fiInvoiceView
    + ' WHERE (Tarif = ''EP-Home'' OR Tarif = ''EP-Profi'')'
    + ' AND VS = ' + Ap + CustomerVarSymbol + Ap;
    SQL.Text := SQLStr;
    Open;
    FakturaVoIP := Fields[0].AsInteger > 0;
    if (not cbSVoIP.Checked and FakturaVoIP) or (not cbBezVoIP.Checked and not FakturaVoIP) then begin
      Close;
      Exit;
    end;
  }

    // není víc zákazníkù pro jeden VS ? 31.1.2017
    //HWTODO mìlo by se kontrolovat v Denní kontrole
    Close;
    SQL.Text := 'SELECT COUNT(*) FROM customers'
    + ' WHERE variable_symbol = ' + Ap + CustomerVarSymbol + Ap;
    Open;
    if Fields[0].AsInteger > 1 then begin
      fmMain.Zprava(Format('Variabilní symbol %s má více zákazníkù.', [CustomerVarSymbol]));
      Close;
      Exit;
    end;

    // je pøenesená daòová povinnost ?  19.10.2016 - bude pro celou fakturu, ne pro jednotlivé smlouvy
    Close;
    SQL.Text := 'SELECT COUNT(*) FROM contracts'
    + ' WHERE DRC = 1'
    + ' AND Customer_Id = (SELECT Id FROM customers'
      + ' WHERE Variable_Symbol = ' + Ap + CustomerVarSymbol + ApZ;
    Open;
    isDRC := Fields[0].AsInteger > 0;

    // vyhledání údajù o smlouvách
    Close;

    SQL.Text := 'CALL get_monthly_invoicing_by_vs('
        + Ap + FormatDateTime('yyyy-mm-dd', StartOfTheMonth(deDatumPlneni.Date)) + ApC
        + Ap + FormatDateTime('yyyy-mm-dd', deDatumPlneni.Date) + ApC
        + Ap + CustomerVarSymbol + ApZ;

    // pøejmenováno takto: AbraKod (cu_abra_code), Typ (co_type), Smlouva (co_number), AktivniOd (co_activated_at), AktivniDo (co_canceled_at),
    // FakturovatOd (co_invoice_from), Tarif (tariff_name), Posilani (cu_invoice_sending_method_name),
    // Perioda (bb_period), Text (bi_description), Cena (bi_price), DPH (bi_vat_name), Tarifni (bi_is_tariff), CTU (co_ctu_category)'

    Open;
    // pøi lecjaké chybì v databázi (napø. Tariff_Id je NULL) konec
    if RecordCount = 0 then begin
      fmMain.Zprava(Format('Pro variabilní symbol %s není co fakturovat.', [CustomerVarSymbol]));
      Close;
      Exit;
    end;
    if FieldByName('cu_abra_code').AsString = '' then begin
      fmMain.Zprava(Format('Smlouva %s: zákazník nemá kód Abry.', [FieldByName('co_number').AsString]));
      Close;
      Exit;
    end;

    { 1.9.2022 zakomentováno, nepotøebné
    if (FieldByName('co_ctu_category').AsString = '') and (FieldByName('co_type').AsString = 'InternetContract') then begin
      dmCommon.Zprava(Format('Smlouva %s: zákazník nemá kód pro ÈTÚ.', [FieldByName('co_number').AsString]));
      Close;
      Exit;
    end;
    }

    with qrAbra do begin
      // kontrola kódu firmy, pøi chybì konec, jinak naèteme
      Close;
      SQL.Text := 'SELECT F.ID as Firm_ID, F.Name as FirmName, F.OrgIdentNumber, FO.Id as FirmOffice_Id'
      + ' FROM Firms F, FirmOffices FO'
      + ' WHERE Code = ' + Ap + qrSmlouva.FieldByName('cu_abra_code').AsString + Ap
      + ' AND F.Firm_ID IS NULL'         // bez následovníkù
      + ' AND F.Hidden = ''N'''
      + ' AND FO.Parent_Id = F.Id'
      + ' ORDER BY F.ID DESC';
      Open;
      if RecordCount = 0 then begin
        fmMain.Zprava(Format('Smlouva %s: zákazník s kódem %s není v adresáøi Abry.',
         [qrSmlouva.FieldByName('co_number').AsString, qrSmlouva.FieldByName('cu_abra_code').AsString]));
        Exit;
      end else begin
        Firm_ID := FieldByName('Firm_ID').AsString;
        FirmName := FieldByName('FirmName').AsString;
        OrgIdentNumber := Trim(FieldByName('OrgIdentNumber').AsString);
        // FirmOffice_ID := FieldByName('FirmOffice_Id').AsString // nebudeme pouzivat, ABRA si sama do faktury zanese nejvhodnìjší FirmOffice
      end;

      // 24.1.2017 obchodní pøípady pro ÈTÚ - platí pro celou faktury
      if isDRC then BusTransactionCode := 'VO'                  // velkoobchod (s DRC)
      else if OrgIdentNumber = '' then BusTransactionCode := 'F'         //  fyzická osoba, nemá IÈ
      else BusTransactionCode := 'P';                         //  právnická osoba
      BusTransaction_Id := AbraEnt.getBusTransaction('Code=' + BusTransactionCode).ID;

// kontrola poslední faktury
      Close;
      SQLStr := 'SELECT OrdNumber, DocDate$DATE, VATDate$DATE, Amount FROM IssuedInvoices'
      + ' WHERE VarSymbol = ' + Ap + CustomerVarSymbol + Ap
      + ' AND VATDate$DATE >= ' + FloatToStr(Trunc(StartOfAMonth(aseRok.Value, aseMesic.Value)))
      + ' AND VATDate$DATE <= ' + FloatToStr(Trunc(EndOfAMonth(aseRok.Value, aseMesic.Value)));
      SQLStr := SQLStr + ' AND DocQueue_ID = ' + Ap + AbraEnt.getDocQueue('Code=FO1').ID + Ap;
      SQLStr := SQLStr + ' ORDER BY OrdNumber DESC';
      SQL.Text := SQLStr;
      Open;

      if RecordCount > 0 then begin
        fmMain.Zprava(Format('%s (%s): %d. faktura se stejným datem.',
         [FirmName, CustomerVarSymbol, RecordCount + 1]));
        Dotaz := Application.MessageBox(PChar(Format('Pro zákazníka "%s" existuje faktura %s-%s s datem %s na èástku %s Kè. Má se vytvoøit další?',
         [FirmName, main_invoiceDocQueueCode, FieldByName('OrdNumber').AsString, DateToStr(FieldByName('DocDate$DATE').AsFloat), FieldByName('Amount').AsString])),
          'Pozor', MB_ICONQUESTION + MB_YESNOCANCEL + MB_DEFBUTTON1);
        if Dotaz = IDNO then begin
          fmMain.Zprava('Ruèní zásah - faktura nevytvoøena.');
          Exit;
        end else if Dotaz = IDCANCEL then begin
          fmMain.Zprava('Ruèní zásah - program ukonèen.');
          Prerusit := True;
          Exit;
        end else fmMain.Zprava('Ruèní zásah - faktura se vytvoøí.');
      end;
    end;  // with qrAbra

    Description := Format('pøipojení %d/%d, ', [aseMesic.Value, aseRok.Value-2000]);


    // vytvoøí se objekt TNewDesInvoiceAA a pak zbytek hlavièky faktury
    NewInvoice := TNewDesInvoiceAA.create(Floor(deDatumDokladu.Date), CustomerVarSymbol);

    NewInvoice.AA['VATDate$DATE'] := Floor(deDatumPlneni.Date);
    NewInvoice.AA['DueTerm'] := aedSplatnost.Text;
    NewInvoice.AA['DocQueue_ID'] := AbraEnt.getDocQueue('Code=FO1').ID; //VarToStr(globalAA['abraIiDocQueue_Id'])
    NewInvoice.AA['Firm_ID'] := Firm_Id;
    // boAA['FirmOffice_ID'] := FirmOffice_Id; // ABRA si vytvoøí sama
    // boAA['CreatedBy_ID'] := MyUser_Id; // vytvoøí se podle uživatele, který se hlásí k ABRA WebApi
    if isDRC then begin
      NewInvoice.AA['IsReverseChargeDeclared'] := True;
      NewInvoice.AA['VATFromAbovePrecision'] := 0; // pro jistotu, default je stejnì 0
      NewInvoice.AA['TotalRounding'] := 0;
    end;

    // 1. øádek
    boRowAA := NewInvoice.createNew0Row(
      Format('Fakturujeme Vám za období od 1.%d.%d do %d.%d.%d',
       [aseMesic.Value, aseRok.Value, DayOfTheMonth(EndOfAMonth(aseRok.Value, aseMesic.Value)), aseMesic.Value, aseRok.Value])
    );

    // ============  smlouvy zákazníka - další øádky faktury se vytvoøí z qrSmlouvy
    while not EOF do begin  // qrSmlouva
      CenaTarifu := 0;
      PausalVoIP := 0;
      HovorneVoIP := 0;
      // je-li datum aktivace menší než datum fakturace, vybere se prvni platba hotovì a fakturuje se pak celý mìsíc, jinak
      // se platí jen èást mìsíce od data spuštìní
      DatumSpusteni := FieldByName('co_activated_at').AsDateTime;
      DatumUkonceni := FieldByName('co_canceled_at').AsDateTime;
      Redukce := 1;
      PrvniFakturace := False;

      if DatumSpusteni >= FieldByName('co_invoice_from').AsDateTime then
      with qrAbra do begin
        // už je nìjaká faktura ?
        Close;
        SQLStr := 'SELECT COUNT(*) FROM IssuedInvoices'
        + ' WHERE DocQueue_ID = ' + Ap + AbraEnt.getDocQueue('Code=FO1').ID + Ap
        + ' AND VarSymbol = ' + Ap + CustomerVarSymbol + Ap;
        SQL.Text := SQLStr;
        Open;
        // zákazníkovi se ještì vùbec nefakturovalo
        PrvniFakturace := Fields[0].AsInteger = 0;
      end;

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
      SmlouvaVoIP := Copy(qrSmlouva.FieldByName('tariff_name').AsString, 1, 2) = 'EP';
      HovorneVoIP := 0;

      if SmlouvaVoIP then begin
        if not ContainsText(Description, 'VoIP,') then // zabránìní vícenásobnému vložení VoIP do Description
          Description := Description + 'VoIP, ';

        if DesU.dbVoIP.Connected then with DesU.qrVoIP do begin
          SQLStr := 'SELECT SUM(Amount) FROM VoIP.Invoices_flat'
          + ' WHERE Num = ' + qrSmlouva.FieldByName('co_number').AsString
          + ' AND Year = ' + aseRok.Text
          + ' AND Month = ' + aseMesic.Text;
          SQL.Text := SQLStr;
          Open;
          HovorneVoIP := ((100 + main_zakladniSazbaDPH)/100) * Fields[0].AsFloat;
          Close;

        end else begin
          HovorneVoIP := 10;
          fmMain.Zprava('Není pøipojena DB VoIP, hovorné nastaveno na 10 Kè');
        end;

      end;

      {  CTU, neni uz potreba
      // 24.1.2017 zakázky pro ÈTÚ - mohou být rùzné podle smlouvy
      with qrAbra do begin
        Close;
        BusOrderCode := qrSmlouva.FieldByName('co_ctu_category').AsString;
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
      if (FieldByName('bi_is_tariff').AsInteger = 1) then begin
        CenaTarifu := qrSmlouva.FieldByName('bi_price').AsFloat;
        if not SmlouvaVoIP then begin

          // pøipojení k Internetu
          boRowAA := NewInvoice.createNew1Row( Format('podle smlouvy  %s  službu  %s',
              [FieldByName('co_number').AsString, FieldByName('bi_description').AsString]));
          // boRowAA['BusOrder_ID'] := BusOrder_Id; // CTU, neni uz potreba
          if qrSmlouva.FieldByName('co_type').AsString = 'TvContract' then
            boRowAA['BusOrder_ID'] := '1700000101';
          boRowAA['BusTransaction_ID'] := BusTransaction_Id;
          boRowAA['TotalPrice'] := Format('%f', [CenaTarifu * Redukce]);
          if isDRC then begin
            boRowAA['VATMode'] := 1;
            boRowAA['VATIndex_ID'] := AbraEnt.getVatIndex('Code=VýstR21').ID;
            boRowAA['DRCArticle_ID'] := AbraEnt.getDrcArticle('Code=21').ID;          // typ plnìní 21, nemá spojitost s DPH
            boRowAA['TotalPrice'] := Format('%f', [FieldByName('bi_price').AsFloat * Redukce / ((100 + main_zakladniSazbaDPH)/100)]);
          end;

        end else begin  // smlouva je VoIP

          if HovorneVoIP > 0 then begin

            // hovorné
            boRowAA := NewInvoice.createNew1Row('hovorné VoIP');
            boRowAA['BusOrder_ID'] := '1500000101';
            boRowAA['BusTransaction_ID'] := BusTransaction_Id;
            boRowAA['TotalPrice'] := Format('%f', [HovorneVoIP]);
            HovorneVoIP := 0;
          end;

          /// paušál
          boRowAA := NewInvoice.createNew1Row(Format('podle smlouvy  %s  mìsíèní platbu VoIP %s',
              [FieldByName('co_number').AsString, FieldByName('bi_description').AsString]));
          boRowAA['BusOrder_ID'] := '2000000101';
          boRowAA['BusTransaction_ID'] := BusTransaction_Id;
          boRowAA['TotalPrice'] := Format('%f', [CenaTarifu * Redukce]);

        end;  // tarif VoIP

      // nìco jiného než tarif
      end else begin

        boRowAA := NewInvoice.createNew1Row(FieldByName('bi_description').AsString);
        // boRowAA['BusOrder_ID'] := BusOrder_Id; // CTU, neni uz potreba
        if qrSmlouva.FieldByName('co_type').AsString = 'TvContract' then
          boRowAA['BusOrder_ID'] := '1700000101';
        boRowAA['BusTransaction_ID'] := BusTransaction_Id;
        if Pos('auce', FieldByName('bi_description').AsString) > 0 then
          boRowAA['IncomeType_ID'] := '1000000101';    // kauce
        boRowAA['TotalPrice'] := Format('%f', [FieldByName('bi_price').AsFloat * Redukce]);
        if FieldByName('bi_vat_name').AsString = '21%' then begin                  // 4.1.2013, DPH je 21%
          if isDRC then begin                                              // 19.10.2016
            boRowAA['VATMode'] := 1;
            boRowAA['VATIndex_ID'] := AbraEnt.getVatIndex('Code=VýstR21').ID;
            boRowAA['DRCArticle_ID'] := AbraEnt.getDrcArticle('Code=21').ID;          // typ plnìní 21, nemá spojitost s DPH
            boRowAA['TotalPrice'] := Format('%f', [FieldByName('bi_price').AsFloat * Redukce / ((100 + main_zakladniSazbaDPH)/100)]);
          end;
        end else begin // DPH je 0%
          boRowAA['VATRate_ID'] := AbraEnt.getVatIndex('Code=Mevd').VATRate_ID; // 00000X0000
          boRowAA['VATIndex_ID'] := AbraEnt.getVatIndex('Code=Mevd').ID; // 7000000000
          if Pos('dar ', FieldByName('bi_description').AsString) > 0 then
            boRowAA['IncomeType_ID'] := '3000000000';  // OS - Ostatní
        end;

      end; // if tarif else
      Next;   //  qrSmlouva
    end;  //  while not qrSmlouva.EOF
// pøípadnì øádek s poštovným
    if (FieldByName('cu_invoice_sending_method_name').AsString = 'Poštou')
     or (FieldByName('cu_invoice_sending_method_name').AsString = 'Se složenkou') then begin    // pošta, složenka

      boRowAA := NewInvoice.createNew1Row('manipulaèní poplatek');
      boRowAA['BusOrder_ID'] := '1000000101';  // internet Mníšek (Code=1)
      boRowAA['BusTransaction_ID'] := BusTransaction_Id;
      boRowAA['TotalPrice'] := '62';
    end;
    NewInvoice.AA['Description'] := Description + CustomerVarSymbol;

    if isDebugMode then begin
      cas02 := Now;
      fmMain.Zprava(debugRozdilCasu(cas01, cas02, ' - èas pøípravy dat fa'));
    end;


// vytvoøení faktury
    try
      abraWebApiResponse := DesU.abraBoCreateWebApi(NewInvoice.AA, 'issuedinvoice');
      if abraWebApiResponse.isOk then begin
        abraResponseSO := SO(abraWebApiResponse.Messg);
        fmMain.Zprava(Format('%s (%s): Vytvoøena faktura %s.', [FirmName, CustomerVarSymbol, abraResponseSO.S['displayname']]));
        Ints[0, Radek] := 0;
        Cells[2, Radek] := abraResponseSO.S['ordnumber']; // faktura
        Cells[3, Radek] := abraResponseSO.S['amount'];                                                // èástka
        Cells[4, Radek] := FirmName;
      end else begin
        fmMain.Zprava(Format('%s (%s): Chyba %s: %s', [FirmName, CustomerVarSymbol, abraWebApiResponse.Code, abraWebApiResponse.Messg]));
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

    if isDebugMode then begin
      cas03 := Now;
      fmMain.Zprava(debugRozdilCasu(cas02, cas03, ' - èas zapsání fa do ABRA'));
    end;

  end;  // with
end;

end.

