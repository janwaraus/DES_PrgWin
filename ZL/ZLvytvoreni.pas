// od 19.12.2014 - Zálohové listy na pøipojení se generují v Abøe podle databáze iQuest. Podle Mìsíèní fakturace.
// 29.7.2019 slevy za 6 a 12 mìsícù neplatí ani pro tarify Lahovská
// 30.4.2020 BusOrder a BusTransaction do všech øádkù ZL
// 28.7.2024 další tarify bez slev, trochu upraveno

unit ZLvytvoreni;

interface

uses
  Windows, Messages, Dialogs, Classes, Forms, Controls, SysUtils, DateUtils, Variants, ComObj, Math,
  ZLmain;

type
  TdmVytvoreni = class(TDataModule)
  private
    procedure ZLAbra(Radek: integer);
  public
    procedure VytvorZL;
  end;

var
  dmVytvoreni: TdmVytvoreni;

implementation

{$R *.dfm}

uses DesUtils, AbraEntities, DesInvoices, Superobject, AArray;

// ------------------------------------------------------------------------------------------------

procedure TdmVytvoreni.VytvorZL;
var
  Radek: integer;
begin
  with fmMain, fmMain.asgMain do try
// pøipojení k Abøe
    Screen.Cursor := crHourGlass;
    asgMain.Visible := False;
    lbxLog.Visible := True;

    fmMain.Zprava(Format('Poèet ZL k vygenerování: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));
    asgMain.Visible := True;
    lbxLog.Visible := False;
    apbProgress.Position := 0;
    apbProgress.Visible := True;
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
      if Ints[0, Radek] = 1 then ZLAbra(Radek);
    end;  // for
// konec hlavní smyèky
  finally

    apbProgress.Position := 0;
    apbProgress.Visible := False;
    asgMain.Visible := False;
    lbxLog.Visible := True;
    Screen.Cursor := crDefault;
    fmMain.Zprava('Generování ZL ukonèeno');
  end;  //  with fmMain
end;

// ------------------------------------------------------------------------------------------------

procedure TdmVytvoreni.ZLAbra(Radek: integer);
// pro další období vytvoøí ZL na pøipojení k internetu

type
  TTarify = set of byte;     // 28.7.2024 seznam tarifù, které nemají slevu pøi placení na 6 nebo 12 mìsícù
// Protože nejnižší dotèený tarif má Id 202, bude se od Id tarifu odeèítat 200, aby i nejvyšší Id
// (dnes 330) byl ještì byte

var
  OrgIdentNumber,
  Firm_Id,
  BusOrder_Id,
  BusOrderCode,
  Speed,
  BusTransaction_Id,
  BusTransactionCode : string[10];
  CenaPripojeni : double;
  Dotaz : integer;
  //ZLObject,
  //ZLData,
  //RadkyZL,
  //RowsCollection: variant;
  Obdobi,
  FirmName,
  CustomerVarSymbol,
  SQLStr : string;

  //boRowAA : TAArray;
  //abraResponseSO : ISuperObject;
  //abraWebApiResponse : TDesResult;
  //NewInvoice : TNewDesInvoiceAA;
  newIDI: TNewAbraBo;
  testVytvoreni : boolean;

  Tarify_bez_slevy: TTarify;

begin
// 28.7.2024
  Tarify_bez_slevy := [
    2..24,             // Misauer
    25..34,            // Harna
    42,                // Lahovská
    45..47,            // Cukrák
    119, 123..126,     // DSL*
    127..130           // optical*
  ];
  testVytvoreni := fmMain.chbTestVytvoreni.Checked;

  with fmMain, DesU.qrZakos, asgMain do begin
    CustomerVarSymbol := Cells[1, Radek];

// není víc zákazníkù pro jeden VS ? 28.6.2018
    Close;
    SQL.Text := 'SELECT COUNT(*) FROM customers'
    + ' WHERE variable_symbol = ' + Ap + CustomerVarSymbol + Ap;
    Open;
    if Fields[0].AsInteger > 1 then begin
      fmMain.Zprava(Format('Variabilní symbol %s má více zákazníkù.', [CustomerVarSymbol]));
      Close;
      Exit;
    end;

    // vyhledání údajù o smlouvách
    Close;

    SQLStr := 'CALL get_deposit_invoicing_by_vs('
        + Ap + FormatDateTime('yyyy-mm-dd', deDatumDokladu.Date + 30) + ApC
        + Ap + FormatDateTime('yyyy-mm-dd', deDatumSplatnosti.Date) + ApC
        + Ap + CustomerVarSymbol + ApC;

    case argMesic.ItemIndex of
      0, 2: SQLStr := SQLStr + 'false,false)'; //první parametr øíká, zda naèítáme 6m periody, druhý, zda 12m periody
      1: SQLStr := SQLStr +  'true,false)'; // dtto
      3: SQLStr := SQLStr +  'true,true)'; // dtto
    end;

    SQL.Text := SQLStr;

    // pøejmenováno takto: AbraCode (cu_abra_code), Typ (co_type), Number (co_number), InvoiceSending (cu_invoice_sending_method_name), Period(bb_period),
    // BItDescription (bi_description), BItPrice (bi_price), BItTariff (bi_is_tariff), TarifId (co_tariff_id), CTU (co_ctu_category)'

    Open;
    // pøi lecjaké chybì v databázi (napø. Tariff_Id je NULL) konec
    if RecordCount = 0 then begin
      fmMain.Zprava(Format('Pro variabilní symbol %s není co generovat.', [CustomerVarSymbol]));
      Close;
      Exit;
    end;
    if FieldByName('cu_abra_code').AsString = '' then begin
      fmMain.Zprava(Format('Smlouva %s: zákazník nemá kód Abry.', [FieldByName('co_number').AsString]));
      Close;
      Exit;
    end;

    // 15.8.2022 {
    {if (FieldByName('CTU').AsString = '') and (FieldByName('co_type').AsString = 'InternetContract') then begin
      fmMain.Zprava(Format('Smlouva %s: zákazník nemá kód pro ÈTÚ.', [FieldByName('co_number').AsString]));
      Close;
      Exit;
    end;  }

    with DesU.qrAbra do begin
    // kontrola kódu firmy, pøi chybì konec
      Close;
      SQLStr := 'SELECT F.ID as Firm_ID, F.Name as FirmName, F.OrgIdentNumber FROM Firms F'
      + ' WHERE Code = ' + Ap + DesU.qrZakos.FieldByName('cu_abra_code').AsString  + Ap
      + ' AND F.Firm_ID IS NULL'         // bez následovníkù
      + ' AND F.Hidden = ''N'''
      + ' ORDER BY F.ID DESC';
      SQL.Text := SQLStr;
      Open;
      if RecordCount = 0 then begin
        fmMain.Zprava(Format('Smlouva %s: zákazník s kódem %s není v adresáøi Abry.',
         [DesU.qrZakos.FieldByName('co_number').AsString, DesU.qrZakos.FieldByName('cu_abra_code').AsString]));
        Exit;
      end
      else if RecordCount > 1 then begin
        fmMain.Zprava(Format('Smlouva %s: V Abøe je více zákazníkù s kódem %s.',
          [DesU.qrZakos.FieldByName('co_number').AsString, DesU.qrZakos.FieldByName('cu_abra_code').AsString]));
        Exit;
      end
      else begin
        Firm_Id := FieldByName('Firm_ID').AsString;
        FirmName := FieldByName('FirmName').AsString;
        OrgIdentNumber := Trim(FieldByName('OrgIdentNumber').AsString);
      end;

      // 28.6.2018 obchodní pøípady pro ÈTÚ - platí pro celou faktury
      if OrgIdentNumber = '' then BusTransactionCode := 'F'               //  fyzická osoba
      else BusTransactionCode := 'P';                         //  právnická osoba
      BusTransaction_Id := AbraEnt.getBusTransaction('Code=' + BusTransactionCode).ID;

      // kontrola posledního ZL

      if not testVytvoreni then begin

        Close;
        SQL.Text := 'SELECT OrdNumber, DocDate$DATE, Amount FROM IssuedDInvoices'
        + ' WHERE DocQueue_ID = ' + Ap + AbraEnt.getDocQueue('Code=ZL1').ID + Ap
        + ' AND VarSymbol = ' + Ap + CustomerVarSymbol + Ap
        + ' AND DocDate$DATE >= ' + FloatToStr(Trunc(VystavenoOd))
        + ' AND DocDate$DATE <= ' + FloatToStr(Trunc(VystavenoDo))
        + ' ORDER BY OrdNumber DESC';
        Open;
        if RecordCount > 0 then begin
          fmMain.Zprava(Format('%s, variabilní symbol %s: %d. zálohový list se stejným datem.',
          [FirmName, CustomerVarSymbol, RecordCount + 1]));
          Dotaz := Application.MessageBox(PChar(Format('Pro variabilní symbol %s existuje zálohový list ZL1-%s s datem %s na èástku %s Kè. Má se vytvoøit další?',
          [CustomerVarSymbol, FieldByName('OrdNumber').AsString, DateToStr(FieldByName('DocDate$DATE').AsFloat), FieldByName('Amount').AsString])),
          'Pozor', MB_ICONQUESTION + MB_YESNOCANCEL + MB_DEFBUTTON1);
          if Dotaz = IDNO then begin
            fmMain.Zprava('Ruèní zásah - zálohový list nevytvoøen.');
            Exit;
          end else if Dotaz = IDCANCEL then begin
            fmMain.Zprava('Ruèní zásah - vytváøení zálohových listù ukonèeno.');
            Prerusit := True;
            Exit;
          end else fmMain.Zprava('Ruèní zásah - zálohový list se vytvoøí.');
        end;  //   RecordCount > 0
      end;
    end;  // with qrAbra

    if testVytvoreni then Exit;

    fmMain.Zprava(Format('%s, variabilní symbol %s.', [FirmName, CustomerVarSymbol]));

    // období  23.9.17
    case FieldByName('bb_period').AsInteger of
      3: begin if argMesic.ItemIndex < 3 then
        Obdobi := Format('%d/%d - %d/%d',
         [3 * argMesic.ItemIndex + 4, aseRok.Value-2000, 3 * argMesic.ItemIndex + 6, aseRok.Value-2000])
      else
        Obdobi := Format('1/%d - 3/%d', [aseRok.Value+1-2000, aseRok.Value+1-2000])
      end;
      6: begin if argMesic.ItemIndex < 2 then
        Obdobi := Format('7/%d - 12/%d', [aseRok.Value-2000, aseRok.Value-2000])
      else
        Obdobi := Format('1/%d - 6/%d', [aseRok.Value+1-2000, aseRok.Value+1-2000])
      end;
      12: Obdobi := Format('1/%d - 12/%d', [aseRok.Value+1-2000, aseRok.Value+1-2000])
    end;


    // hlavièka faktury
    newIDI := TNewAbraBo.Create('issueddepositinvoice');
    newIDI.addInvoiceParams(Floor(deDatumDokladu.Date));
    newIDI.Item['Varsymbol'] := CustomerVarSymbol;
    newIDI.Item['DocQueue_ID'] := AbraEnt.getDocQueue('Code=ZL1').ID;
    newIDI.Item['Description'] := Format('Pøipojení %s, %s', [Obdobi, CustomerVarSymbol]);
    newIDI.Item['Firm_ID'] := Firm_Id;
    newIDI.Item['DueDate$DATE'] := Floor(deDatumSplatnosti.Date);

    // vytvoøí se objekt TNewDesInvoiceAA a pak zbytek hlavièky ZL
    //* NewInvoice := TNewDesInvoiceAA.create(Floor(deDatumDokladu.Date), CustomerVarSymbol, '10');

    //ZLObject := AbraOLE.CreateObject('@IssuedDepositInvoice');
    //ZLData := AbraOLE.CreateValues('@IssuedDepositInvoice');
    //ZLObject.PrefillValues(ZLData);
    //* NewInvoice.AA['DocQueue_ID'] := AbraEnt.getDocQueue('Code=ZL1').ID;
    //NewInvoice.AA['Period_ID'] := Period_Id; // *HW* automaticky z data dokladu
    //* NewInvoice.AA['Description'] := Format('Pøipojení %s, %s', [Obdobi, CustomerVarSymbol]);
    //* NewInvoice.AA['Firm_ID'] := Firm_Id;
    //* NewInvoice.AA['DueDate$DATE'] := Floor(deDatumSplatnosti.Date);        // 16.9.22 musí být až tady

    {  CTU, neni uz potreba
    // 28.6.2018 zakázky pro ÈTÚ - mohou být rùzné podle smlouvy
    with qrAbra do begin
      Close;
      BusOrderCode := FieldByName('co_ctu_category').AsString;
    // kódy z tabulky contracts se musí pøedìlat
      // 15.8.22
      BusOrder_Id := '';
      if BusOrderCode <> '' then begin
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
    end;
    }


    // 1. øádek s textem
    newIDI.createNewInvoiceRow(0, Format('Úètujeme Vám na období  %s', [Obdobi]));

    // další øádky faktury se vytvoøí z DesU.qrZakos
    while not EOF do begin
      CenaPripojeni := 0;
      // cena za pøipojení k Internetu (je tarif, tj. mùže to být i televize - v praxi se televize ale na ZL nedává, aby si zíkazníci mohli mìnit TV tarif dle potøeby)
      if (FieldByName('bi_is_tariff').AsInteger = 1) then
        CenaPripojeni := FieldByName('bi_price').AsFloat;
      // smlouva na pøipojení
      if CenaPripojeni <> 0 then begin
        newIDI.createNewInvoiceRow(4,
          Format('podle smlouvy  %s  službu  %s', [FieldByName('co_number').AsString, FieldByName('bi_description').AsString]), false);
        //boRowAA['BusOrder_ID'] := BusOrder_Id;
        newIDI.rowItem['BusTransaction_ID'] := BusTransaction_Id;
        newIDI.rowItem['TAmount'] := Format('%f', [FieldByName('bb_period').AsInteger * FieldByName('bi_price').AsFloat]);

        // platba na 6 mìsícù (sleva jen za internet) - neplatí pro tarify Misauer a Harna 30.3.20 a Cukrák
        // 29.7.2019 slevy neplatí ani pro tarif Lahovská
        // 28.7.2024 slevy neplatí ani pro tarify DSL* a optical*
        if (FieldByName('bb_period').AsInteger = 6) and (FieldByName('co_type').AsString = 'InternetContract')
        // Misauer 202 - 224, Harna 225 - 234, Lahovská 242, Cukrák 245 - 247, DSL* 319 a 323-326, optical* 327-330
//         and not (FieldByName('co_tariff_id').AsInteger in [202..234, 242, 245..247]) then begin   // Misauer 202 - 224, Harna 225 - 234, Lahovská 242, Cukrák 245 - 247
         and not (FieldByName('co_tariff_id').AsInteger - 200 in Tarify_bez_slevy) then begin        // 28.7.24
          //RadkyZL:= AbraOLE.CreateValues('@IssuedDepositInvoiceRow');
          newIDI.createNewInvoiceRow(4, 'slevu', false);
          //boRowAA['BusOrder_ID'] := BusOrder_Id;
          newIDI.rowItem['BusTransaction_ID'] := BusTransaction_Id;
          newIDI.rowItem['TAmount'] := Format('%f', [-CenaPripojeni]);
        end;

        // platba na 12 mìsícù (sleva jen za internet) - neplatí pro tarify Misauer a Harna 30.3.20 a Cukrák
        // 29.7.2019 slevy neplatí ani pro tarif Lahovská
        // 28.7.2024 slevy neplatí ani pro tarify DSL* a optical*
        if (FieldByName('bb_period').AsInteger = 12) and (FieldByName('co_type').AsString = 'InternetContract')
        // Misauer 202 - 224, Harna 225 - 234, Lahovská 242, Cukrák 245 - 247, DSL* 319 a 323-326, optical* 327-330
//         and not (FieldByName('co_tariff_id').AsInteger in [202..234, 242, 245..247]) then begin   // Misauer 202 - 224, Harna 225 - 234, Lahovská 242, Cukrák 245 - 247
         and not (FieldByName('co_tariff_id').AsInteger - 200 in Tarify_bez_slevy) then begin        // 28.7.24
          newIDI.createNewInvoiceRow(4, 'slevu', false);
          //boRowAA['BusOrder_ID'] := BusOrder_Id;
          newIDI.rowItem['BusTransaction_ID'] := BusTransaction_Id;
          newIDI.rowItem['TAmount'] := Format('%f', [-2*CenaPripojeni]);
        end;
        CenaPripojeni := 0; // HW asi zbyteèné tady
      end
      // nìco jiného než platba za pøipojení (CenaPripojeni = 0)
      else begin
        newIDI.createNewInvoiceRow(4, FieldByName('bi_description').AsString, false);
        //boRowAA['BusOrder_ID'] := BusOrder_Id;                          // 30.4.21
        newIDI.rowItem['BusTransaction_ID'] := BusTransaction_Id;              // 30.4.21
        newIDI.rowItem['TAmount'] := Format('%f', [FieldByName('bb_period').AsInteger * FieldByName('bi_price').AsFloat]);
        // platba na 6 mìsícù
        if (FieldByName('bb_period').AsInteger = 6) then
//          if (FieldByName('co_tariff_id').AsInteger in [202..234]) then
          if (FieldByName('co_tariff_id').AsInteger - 200 in Tarify_bez_slevy) then                    // 28.7.24 tarify bez slevy
            newIDI.rowItem['TAmount'] := Format('%f', [6 * FieldByName('bi_price').AsFloat])
          else newIDI.rowItem['TAmount'] := Format('%f', [5 * FieldByName('bi_price').AsFloat])  // ostatní
        // platba na 12 mìsícù
        else if FieldByName('bb_period').AsInteger = 12 then
          if (FieldByName('co_tariff_id').AsInteger - 200 in Tarify_bez_slevy) then                     // 28.7.24 tarify bez slevy
            newIDI.rowItem['TAmount'] := Format('%f', [12 * FieldByName('bi_price').AsFloat])
          else newIDI.rowItem['TAmount'] := Format('%f', [10 * FieldByName('bi_price').AsFloat])  // ostatní
        // ostatní
        else newIDI.rowItem['TAmount'] := Format('%f', [FieldByName('bb_period').AsInteger * FieldByName('bi_price').AsFloat]);
      end;
      Next;   //  DesU.qrZakos
    end;  //  while not DesU.qrZakos.EOF
    Close;   // DesU.qrZakos


    try
      //* abraWebApiResponse := DesU.abraBoCreate(NewInvoice.AA.AsJSon, 'issueddepositinvoice');
      newIDI.writeToAbra;
      if newIDI.WriteResult.isOk then begin
        fmMain.Zprava(Format('%s (%s): Vytvoøen zálohový list %s.', [FirmName, CustomerVarSymbol, newIDI.getCreatedBoItem('displayname')]));
        Ints[0, Radek] := 0;
        Cells[2, Radek] := newIDI.getCreatedBoItem('ordnumber'); // faktura
        Cells[3, Radek] := newIDI.getCreatedBoItem('amount');  // èástka
        Cells[4, Radek] := FirmName;
      end else begin
        fmMain.Zprava(Format('%s (%s): Chyba %s: %s', [FirmName, CustomerVarSymbol, newIDI.WriteResult.Code, newIDI.WriteResult.Messg]));
        if Dialogs.MessageDlg( '(' + newIDI.WriteResult.Code + ') '
           + newIDI.WriteResult.Messg + sLineBreak + 'Pokraèovat?',
           mtConfirmation, [mbYes, mbNo], 0 ) = mrNo then Prerusit := True;
      end;
    except on E: exception do
      begin
        Application.MessageBox(PChar('Problem ' + ^M + E.Message), 'Vytvoøení fa');
      end;
    end;


  end;  // with

end;  // procedury ZLAbra




end.

