// DES fakturuje najednou v�echny smlouvy z�kazn�ka, internet, VoIP i IPTV s datem vystaven� i pln�n� posledn� den v m�s�ci.
// Z�kazn�ci nebo smlouvy k fakturaci jsou v asgMain, faktury se vytvo�� v cyklu p�es ��dky.
// 23.1.17 V�kazy pro �T� vy�aduj� d�len� podle typu u�ivatele, technologie a rychlosti - do Abry byly p�id�ny 3 obchodn� p��pady
// - VO (velkoobchod - DRC) , F a P (fyzick� a pr�vnick� osoba) a 24 zak�zek (W - WiFi, B - FFTB, A - FTTH AON, a P - FTTH PON,
// ka�d� s �esti r�zn�mi rychlostmi). Ka�d�mu ��dku faktury bude p�i�azen jeden obchodn� p��pad a jedna zak�zka.

unit FIfaktura;

interface

uses
  Windows, Messages, Dialogs, Classes, Forms, Controls, SysUtils, DateUtils, StrUtils, Variants, ComObj, Math,
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

uses DesUtils, AbraEntities, DesInvoices; // Superobject, AArray,

// ------------------------------------------------------------------------------------------------

procedure TdmFaktura.VytvorFaktury;
var
  Radek: integer;
begin
  with fmMain, fmMain.asgMain do try
    fmMain.Zprava(Format('Po�et faktur k vygenerov�n�: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));

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
    fmMain.Zprava('Generov�n� faktur ukon�eno');
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

  //boRowAA: TAArray;
  //abraResponseSO : ISuperObject;
  //abraWebApiResponse : TDesResult;
  //NewInvoice : TNewDesInvoiceAA;
  newII: TNewAbraBo;

  cas01, cas02, cas03: double;

begin
  cas01 := Now;

  with fmMain, DesU.qrZakos, asgMain do begin
    CustomerVarSymbol := Cells[1, Radek];

    // nen� v�c z�kazn�k� pro jeden VS ? 31.1.2017
    //HWTODO m�lo by se kontrolovat v Denn� kontrole
    Close;
    SQL.Text := 'SELECT COUNT(*) FROM customers'
    + ' WHERE variable_symbol = ' + Ap + CustomerVarSymbol + Ap;
    Open;
    if Fields[0].AsInteger > 1 then begin
      fmMain.Zprava(Format('Variabiln� symbol %s m� v�ce z�kazn�k�.', [CustomerVarSymbol]));
      Close;
      Exit;
    end;

    // je p�enesen� da�ov� povinnost ?  19.10.2016 - bude pro celou fakturu, ne pro jednotliv� smlouvy
    Close;
    SQL.Text := 'SELECT COUNT(*) FROM contracts'
    + ' WHERE drc = 1'
    + ' AND customer_id = (SELECT id FROM customers'
      + ' WHERE variable_symbol = ' + Ap + CustomerVarSymbol + ApZ;
    Open;
    isDRC := Fields[0].AsInteger > 0;

    // vyhled�n� �daj� o smlouv�ch
    Close;

    SQL.Text := 'CALL get_monthly_invoicing_by_vs('
        + Ap + FormatDateTime('yyyy-mm-dd', StartOfTheMonth(deDatumPlneni.Date)) + ApC
        + Ap + FormatDateTime('yyyy-mm-dd', deDatumPlneni.Date) + ApC
        + Ap + CustomerVarSymbol + ApZ;

    // p�ejmenov�no takto: AbraKod (cu_abra_code), Typ (co_type), Smlouva (co_number), AktivniOd (co_activated_at), AktivniDo (co_canceled_at),
    // FakturovatOd (co_invoice_from), Tarif (tariff_name), Posilani (cu_invoice_sending_method_name),
    // Perioda (bb_period), Text (bi_description), Cena (bi_price), DPH (bi_vat_name), Tarifni (bi_is_tariff), CTU (co_ctu_category)'

    Open;
    // p�i lecjak� chyb� v datab�zi (nap�. Tariff_Id je NULL) konec
    if RecordCount = 0 then begin
      fmMain.Zprava(Format('Pro variabiln� symbol %s nen� co fakturovat.', [CustomerVarSymbol]));
      Close;
      Exit;
    end;
    if FieldByName('cu_abra_code').AsString = '' then begin
      fmMain.Zprava(Format('Smlouva %s: z�kazn�k nem� k�d Abry.', [FieldByName('co_number').AsString]));
      Close;
      Exit;
    end;

    { 1.9.2022 zakomentov�no, nepot�ebn�
    if (FieldByName('co_ctu_category').AsString = '') and (FieldByName('co_type').AsString = 'InternetContract') then begin
      dmCommon.Zprava(Format('Smlouva %s: z�kazn�k nem� k�d pro �T�.', [FieldByName('co_number').AsString]));
      Close;
      Exit;
    end;
    }

    with DesU.qrAbra do begin
      // kontrola k�du firmy, p�i chyb� konec, jinak na�teme
      Close;
      SQL.Text := 'SELECT F.ID as Firm_ID, F.Name as FirmName, F.OrgIdentNumber'
      + ' FROM Firms F'
      + ' WHERE Code = ' + Ap + DesU.qrZakos.FieldByName('cu_abra_code').AsString + Ap
      + ' AND F.Firm_ID IS NULL'         // bez n�sledovn�k�
      + ' AND F.Hidden = ''N'''
      + ' ORDER BY F.ID DESC';
      Open;
      if RecordCount = 0 then begin
        fmMain.Zprava(Format('Smlouva %s: Z�kazn�k s k�dem %s nen� v adres��i Abry.',
         [DesU.qrZakos.FieldByName('co_number').AsString, DesU.qrZakos.FieldByName('cu_abra_code').AsString]));
        Exit;
      end
      else if RecordCount > 1 then begin
        fmMain.Zprava(Format('Smlouva %s: V Ab�e je v�ce z�kazn�k� s k�dem %s.',
          [DesU.qrZakos.FieldByName('co_number').AsString, DesU.qrZakos.FieldByName('cu_abra_code').AsString]));
        Exit;
      end
      else begin
        Firm_ID := FieldByName('Firm_ID').AsString;
        FirmName := FieldByName('FirmName').AsString;
        OrgIdentNumber := Trim(FieldByName('OrgIdentNumber').AsString);
      end;

      // 24.1.2017 obchodn� p��pady pro �T� - plat� pro celou faktury
      if isDRC then BusTransactionCode := 'VO'                  // velkoobchod (s DRC)
      else if OrgIdentNumber = '' then BusTransactionCode := 'F'         //  fyzick� osoba, nem� I�
      else BusTransactionCode := 'P';                         //  pr�vnick� osoba
      BusTransaction_Id := AbraEnt.getBusTransaction('Code=' + BusTransactionCode).ID;

      // kontrola posledn� faktury
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
        fmMain.Zprava(Format('%s (%s): %d. faktura se stejn�m datem.',
         [FirmName, CustomerVarSymbol, RecordCount + 1]));
        Dotaz := Application.MessageBox(PChar(Format('Pro z�kazn�ka "%s" existuje faktura %s-%s s datem %s na ��stku %s K�. M� se vytvo�it dal��?',
         [FirmName, 'FO1', FieldByName('OrdNumber').AsString, DateToStr(FieldByName('DocDate$DATE').AsFloat), FieldByName('Amount').AsString])),
          'Pozor', MB_ICONQUESTION + MB_YESNOCANCEL + MB_DEFBUTTON1);
        if Dotaz = IDNO then begin
          fmMain.Zprava('Ru�n� z�sah - faktura nevytvo�ena.');
          Exit;
        end else if Dotaz = IDCANCEL then begin
          fmMain.Zprava('Ru�n� z�sah - program ukon�en.');
          Prerusit := True;
          Exit;
        end else fmMain.Zprava('Ru�n� z�sah - faktura se vytvo��.');
      end;
    end;  // with qrAbra

    Description := Format('p�ipojen� %d/%d, ', [aseMesic.Value, aseRok.Value-2000]);


    // hlavi�ka faktury
    newII := TNewAbraBo.Create('issuedinvoice');
    newII.addInvoiceParams(Floor(deDatumDokladu.Date));
    newII.Item['Varsymbol'] := CustomerVarSymbol;


    newII.Item['VATDate$DATE'] := Floor(deDatumPlneni.Date);
    newII.Item['DueTerm'] := aedSplatnost.Text;
    newII.Item['DocQueue_ID'] := AbraEnt.getDocQueue('Code=FO1').ID; //VarToStr(globalAA['abraIiDocQueue_Id'])
    newII.Item['Firm_ID'] := Firm_Id;
    // boAA['FirmOffice_ID'] := FirmOffice_Id; // ABRA si vytvo�� sama
    // boAA['CreatedBy_ID'] := MyUser_Id; // vytvo�� se podle u�ivatele, kter� se hl�s� k ABRA WebApi
    if isDRC then begin
      newII.Item['IsReverseChargeDeclared'] := True;
      newII.Item['VATFromAbovePrecision'] := 0; // pro jistotu, default je stejn� 0
      newII.Item['TotalRounding'] := 0;
    end;

    // 1. ��dek
    newII.createNewInvoiceRow(0,
      Format('Fakturujeme V�m za obdob� od 1.%d.%d do %d.%d.%d',
       [aseMesic.Value, aseRok.Value, DayOfTheMonth(EndOfAMonth(aseRok.Value, aseMesic.Value)), aseMesic.Value, aseRok.Value])
    );

    // ============  smlouvy z�kazn�ka - dal�� ��dky faktury se vytvo�� z qrSmlouvy
    while not EOF do begin  // DesU.qrZakos
      CenaTarifu := 0;
      PausalVoIP := 0;
      HovorneVoIP := 0;
      // je-li datum aktivace men�� ne� datum fakturace, vybere se prvni platba hotov� a fakturuje se pak cel� m�s�c, jinak
      // se plat� jen ��st m�s�ce od data spu�t�n�
      DatumSpusteni := FieldByName('co_activated_at').AsDateTime;
      DatumUkonceni := FieldByName('co_canceled_at').AsDateTime;
      Redukce := 1;
      PrvniFakturace := False;

      if DatumSpusteni >= FieldByName('co_invoice_from').AsDateTime then
      with DesU.qrAbraOC do begin
        // u� je n�jak� faktura ?
        SQL.Text := 'SELECT COUNT(*) FROM IssuedInvoices'
        + ' WHERE DocQueue_ID = ' + Ap + AbraEnt.getDocQueue('Code=FO1').ID + Ap
        + ' AND VarSymbol = ' + Ap + CustomerVarSymbol + Ap;
        Open;
        // z�kazn�kovi se je�t� v�bec nefakturovalo
        PrvniFakturace := Fields[0].AsInteger = 0;
        Close;
      end;

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
      //SmlouvaVoIP := Copy(DesU.qrZakos.FieldByName('tariff_name').AsString, 1, 2) = 'EP'; // pot�eba vynechat kreditni VoIP (tariff_id = 2)
      SmlouvaVoIP := (DesU.qrZakos.FieldByName('co_tariff_id').AsInteger = 1) OR (DesU.qrZakos.FieldByName('co_tariff_id').AsInteger = 3);
      HovorneVoIP := 0;

      if SmlouvaVoIP then begin
        if not ContainsText(Description, 'VoIP,') then // zabr�n�n� v�cen�sobn�mu vlo�en� VoIP do Description
          Description := Description + 'VoIP, ';

        if dbNewVoIP.Connected then begin
          SQLStr := 'SELECT SUM(Amount) FROM VoIP.Invoices_flat'
          + ' WHERE Num = ' + DesU.qrZakos.FieldByName('co_number').AsString
          + ' AND Year = ' + aseRok.Text
          + ' AND Month = ' + aseMesic.Text;
          qrNewVoip.SQL.Text := SQLStr;
          qrNewVoip.Open;
          HovorneVoIP := DesU.VAT_MULTIPLIER * qrNewVoip.Fields[0].AsFloat;
          qrNewVoip.Close;

        end else begin
          //HovorneVoIP := 10;
          fmMain.Zprava('Nen� p�ipojena DB VoIP, faktura se nevytvo��');
          Prerusit := True;
          Exit;
        end;
      end;

      {  CTU, neni uz potreba
      // 24.1.2017 zak�zky pro �T� - mohou b�t r�zn� podle smlouvy
      with qrAbra do begin
        Close;
        BusOrderCode := DesU.qrZakos.FieldByName('co_ctu_category').AsString;
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
      if (FieldByName('bi_is_tariff').AsInteger = 1) then begin
        CenaTarifu := DesU.qrZakos.FieldByName('bi_price').AsFloat;
        if not SmlouvaVoIP then begin

          // p�ipojen� k Internetu
          newII.createNewInvoiceRow(1, Format('podle smlouvy  %s  slu�bu  %s',
              [FieldByName('co_number').AsString, FieldByName('bi_description').AsString]));

          // newII.rowItem['BusOrder_ID'] := BusOrder_Id; // CTU, neni uz potreba
          if DesU.qrZakos.FieldByName('co_type').AsString = 'TvContract' then
            newII.rowItem['BusOrder_ID'] := '1700000101';
          newII.rowItem['BusTransaction_ID'] := BusTransaction_Id;
          newII.rowItem['TotalPrice'] := Format('%f', [CenaTarifu * Redukce]);
          if isDRC then begin
            newII.rowItem['VATMode'] := 1;
            newII.rowItem['VATIndex_ID'] := AbraEnt.getVatIndex('Code=V�stR' + DesU.VAT_RATE).ID;
            newII.rowItem['DRCArticle_ID'] := AbraEnt.getDrcArticle('Code=21').ID;          // typ pln�n� 21, nem� spojitost s DPH
            newII.rowItem['TotalPrice'] := Format('%f', [FieldByName('bi_price').AsFloat * Redukce / DesU.VAT_MULTIPLIER ]);
          end;

        end else begin  // smlouva je VoIP

          if HovorneVoIP > 0 then begin

            // hovorn�
            newII.createNewInvoiceRow(1, 'hovorn� VoIP');
            newII.rowItem['BusOrder_ID'] := '1500000101';
            newII.rowItem['BusTransaction_ID'] := BusTransaction_Id;
            newII.rowItem['TotalPrice'] := Format('%f', [HovorneVoIP]);
            HovorneVoIP := 0;
          end;

          /// pau��l
          newII.createNewInvoiceRow(1, Format('podle smlouvy  %s  m�s��n� platbu VoIP %s',
              [FieldByName('co_number').AsString, FieldByName('bi_description').AsString]));
          newII.rowItem['BusOrder_ID'] := '2000000101';
          newII.rowItem['BusTransaction_ID'] := BusTransaction_Id;
          newII.rowItem['TotalPrice'] := Format('%f', [CenaTarifu * Redukce]);

        end;  // tarif VoIP

      // n�co jin�ho ne� tarif
      end else begin

        newII.createNewInvoiceRow(1, FieldByName('bi_description').AsString);
        // newII.rowItem['BusOrder_ID'] := BusOrder_Id; // CTU, neni uz potreba
        if DesU.qrZakos.FieldByName('co_type').AsString = 'TvContract' then
          newII.rowItem['BusOrder_ID'] := '1700000101';
        newII.rowItem['BusTransaction_ID'] := BusTransaction_Id;
        if Pos('auce', FieldByName('bi_description').AsString) > 0 then
          newII.rowItem['IncomeType_ID'] := '1000000101';    // kauce
        newII.rowItem['TotalPrice'] := Format('%f', [FieldByName('bi_price').AsFloat * Redukce]);
        if FieldByName('bi_vat_name').AsString = '21%' then begin                  // 4.1.2013, DPH je 21%
          if isDRC then begin                                              // 19.10.2016
            newII.rowItem['VATMode'] := 1;
            newII.rowItem['VATIndex_ID'] := AbraEnt.getVatIndex('Code=V�stR' + DesU.VAT_RATE).ID;
            newII.rowItem['DRCArticle_ID'] := AbraEnt.getDrcArticle('Code=21').ID;          // typ pln�n� 21, nem� spojitost s DPH
            newII.rowItem['TotalPrice'] := Format('%f', [FieldByName('bi_price').AsFloat * Redukce / DesU.VAT_MULTIPLIER]);
          end;
        end else begin // DPH je 0%
          newII.rowItem['VATRate_ID'] := AbraEnt.getVatIndex('Code=Mevd').VATRate_ID; // 00000X0000
          newII.rowItem['VATIndex_ID'] := AbraEnt.getVatIndex('Code=Mevd').ID; // 7000000000
          if Pos('dar ', FieldByName('bi_description').AsString) > 0 then
            newII.rowItem['IncomeType_ID'] := '3000000000';  // OS - Ostatn�
        end;

      end; // if tarif else
      Next;   //  DesU.qrZakos
    end;  //  while not DesU.qrZakos.EOF
// p��padn� ��dek s po�tovn�m
    if (FieldByName('cu_invoice_sending_method_name').AsString = 'Po�tou')
     or (FieldByName('cu_invoice_sending_method_name').AsString = 'Se slo�enkou') then begin    // po�ta, slo�enka

      newII.createNewInvoiceRow(1, 'manipula�n� poplatek');
      newII.rowItem['BusOrder_ID'] := '1000000101';  // internet Mn�ek (Code=1)
      newII.rowItem['BusTransaction_ID'] := BusTransaction_Id;
      newII.rowItem['TotalPrice'] := '62';
    end;
    newII.Item['Description'] := Description + CustomerVarSymbol;

    if isDebugMode then begin
      cas02 := Now;
      fmMain.Zprava(debugRozdilCasu(cas01, cas02, ' - �as p��pravy dat fa'));
    end;


// vytvo�en� faktury
    try
      //abraWebApiResponse := DesU.abraBoCreateWebApi(NewInvoice.AA, 'issuedinvoice');
      newII.writeToAbra;
      if newII.WriteResult.isOk then begin
        //abraResponseSO := SO(abraWebApiResponse.Messg);
        fmMain.Zprava(Format('%s (%s): Vytvo�ena faktura %s.', [FirmName, CustomerVarSymbol, newII.getCreatedBoItem('displayname')]));
        Ints[0, Radek] := 0;
        Cells[2, Radek] := newII.getCreatedBoItem('ordnumber'); // faktura
        Cells[3, Radek] := newII.getCreatedBoItem('amount');  // ��stka
        Cells[4, Radek] := FirmName;
      end else begin
        fmMain.Zprava(Format('%s (%s): Chyba %s: %s', [FirmName, CustomerVarSymbol, newII.WriteResult.Code, newII.WriteResult.Messg]));
        if Dialogs.MessageDlg( '(' + newII.WriteResult.Code + ') '
           + newII.WriteResult.Messg + sLineBreak + 'Pokra�ovat?',
           mtConfirmation, [mbYes, mbNo], 0 ) = mrNo then Prerusit := True;
      end;
    except on E: exception do
      begin
        Application.MessageBox(PChar('Problem ' + ^M + E.Message), 'Vytvo�en� fa');
      end;
    end;

    Close;   // DesU.qrZakos

    if isDebugMode then begin
      cas03 := Now;
      fmMain.Zprava(debugRozdilCasu(cas02, cas03, ' - �as zaps�n� fa do ABRA'));
    end;

  end;  // with
end;

end.

