// od 19.12.2014 - Zálohové listy na pøipojení se generují v Abøe podle databáze iQuest. Podle Mìsíèní fakturace.
// 29.7.2019 slevy za 6 a 12 mìsícù neplatí ani pro tarify Lahovská
// 30.4.2020 BusOrder a BusTransaction do všech øádkù ZL

unit ZLvytvoreni;

interface

uses
  Windows, Messages, Classes, Forms, Controls, SysUtils, DateUtils, Variants, ComObj, Math, ZLmain;

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

uses ZLcommon;

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
    try
      AbraOLE.Logout;
    except
    end;
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
var
  IC,
  BusOrder_Id,
  BusOrderCode,
  Speed,
  BusTransaction_Id,
  BusTransactionCode,
  FirmOffice_Id: string[10];
  CenaPripojeni: double;
  Dotaz,
  Cena: integer;
  ZLObject,
  ZLData,
  RadkyZL,
  RowsCollection: variant;
  CisloZL,
  Obdobi,
  Zakaznik,
  SQLStr: AnsiString;
begin
  with fmMain, qrSmlouva, asgMain do begin
// není víc zákazníkù pro jeden VS ? 28.6.2018
    Close;
    SQLStr := 'SELECT COUNT(*) FROM customers'
    + ' WHERE Variable_Symbol = ' + Ap + Cells[1, Radek] + Ap;
    SQL.Text := SQLStr;
    Open;
    if Fields[0].AsInteger > 1 then begin
      fmMain.Zprava(Format('Variabilní symbol %s má více zákazníkù.', [Cells[1, Radek]]));
      Close;
      Exit;
    end;
// vyhledání údajù o smlouvách
    Close;
    SQLStr := 'SELECT AbraCode, Typ, Number, InvoiceSending, Period, BItDescription, BItPrice, BItTariff, TarifId/co_tariff_id, CTU'
    + ' FROM ' + ZLView
    + ' WHERE VS = ' + Ap + Cells[1, Radek] + Ap
    + ' ORDER BY Number, BItTariff DESC';
    SQL.Text := SQLStr;
    Open;
// pøi lecjaké chybì v databázi (napø. Tariff_Id je NULL) konec
    if RecordCount = 0 then begin
      fmMain.Zprava(Format('Pro variabilní symbol %s není co generovat.', [Cells[1, Radek]]));
      Close;
      Exit;
    end;
    if FieldByName('AbraCode').AsString = '' then begin
      fmMain.Zprava(Format('Smlouva %s: zákazník nemá kód Abry.', [FieldByName('Number').AsString]));
      Close;
      Exit;
    end;
// 28.6.2018 opraveno 17.12.
// 15.8.2022 {
{    if (FieldByName('CTU').AsString = '') and (FieldByName('Typ').AsString = 'InternetContract') then begin
      fmMain.Zprava(Format('Smlouva %s: zákazník nemá kód pro ÈTÚ.', [FieldByName('Number').AsString]));
      Close;
      Exit;
    end;  }
    with qrAbra do begin
// kontrola kódu firmy, pøi chybì konec
      Close;
      SQLStr := 'SELECT F.ID, F.Name, F.OrgIdentNumber, FO.Id FROM Firms F, FirmOffices FO'
      + ' WHERE Code = ' + Ap + qrSmlouva.FieldByName('AbraCode').AsString + Ap
      + ' AND F.Firm_ID IS NULL'         // bez pøedkù
      + ' AND F.Hidden = ''N'''
      + ' AND FO.Parent_Id = F.Id'
      + ' ORDER BY F.ID DESC';
      SQL.Text := SQLStr;
      Open;
      if RecordCount = 0 then begin
        fmMain.Zprava(Format('Smlouva %s: zákazník s kódem %s není v adresáøi Abry.',
         [qrSmlouva.FieldByName('Number').AsString, qrSmlouva.FieldByName('AbraKod').AsString]));
        Exit;
      end else begin
        Firm_Id := Fields[0].AsString;
        Zakaznik := UTF8Decode(Fields[1].AsString);
        IC := Trim(Fields[2].AsString);
        FirmOffice_Id := Fields[3].AsString;
      end;
// 28.6.2018 obchodní pøípady pro ÈTÚ - platí pro celou faktury
      if IC = '' then BusTransactionCode := 'F'               //  fyzická osoba
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
// kontrola posledního ZL
      Close;
      SQLStr := 'SELECT OrdNumber, DocDate$DATE, Amount FROM IssuedDInvoices'
      + ' WHERE DocQueue_ID = ' + Ap + AbraEnt.getDocQueue('Code=ZL1').ID + Ap
      + ' AND VarSymbol = ' + Ap + Cells[1, Radek] + Ap
      + ' AND DocDate$DATE >= ' + FloatToStr(Trunc(VystavenoOd))
      + ' AND DocDate$DATE <= ' + FloatToStr(Trunc(VystavenoDo))
      + ' ORDER BY OrdNumber DESC';
      SQL.Text := SQLStr;
      Open;
      if RecordCount > 0 then begin
        fmMain.Zprava(Format('%s, variabilní symbol %s: %d. zálohový list se stejným datem.',
        [Zakaznik, Cells[1, Radek], RecordCount + 1]));
        Dotaz := Application.MessageBox(PChar(Format('Pro variabilní symbol %s existuje zálohový list ZL1-%s s datem %s na èástku %s Kè. Má se vytvoøit další?',
        [Cells[1, Radek], FieldByName('OrdNumber').AsString, DateToStr(FieldByName('DocDate$DATE').AsFloat), FieldByName('Amount').AsString])),
        'Pozor', MB_ICONQUESTION + MB_YESNOCANCEL + MB_DEFBUTTON1);
        if Dotaz = IDNO then begin
          fmMain.Zprava('Ruèní zásah - zálohový list nevytvoøen.');
          Exit;
        end else if Dotaz = IDCANCEL then begin
          fmMain.Zprava('Ruèní zásah - program ukonèen.');
          Prerusit := True;
          Exit;
        end else fmMain.Zprava('Ruèní zásah - zálohový list se vytvoøí.');
      end;  //   RecordCount > 0
    end;  // with qrAbra
    fmMain.Zprava(Format('%s, variabilní symbol %s.', [Zakaznik, Cells[1, Radek]]));
// období  23.9.17
    case qrSmlouva.FieldByName('Period').AsInteger of
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
// vytvoøí se hlavièka ZL
    ZLObject := AbraOLE.CreateObject('@IssuedDepositInvoice');
    ZLData := AbraOLE.CreateValues('@IssuedDepositInvoice');
    ZLObject.PrefillValues(ZLData);
    ZLData.ValueByName('DocQueue_ID') := AbraEnt.getDocQueue('Code=ZL1').ID;
    //ZLData.ValueByName('Period_ID') := Period_Id; // *HW* automaticky z data dokladu
    ZLData.ValueByName('DocDate$DATE') := Floor(DatumDokladu);
//    ZLData.ValueByName('DueDate$DATE') := Floor(DatumSplatnosti);
    ZLData.ValueByName('Description') := Format('Pøipojení %s, %s', [Obdobi, Cells[1, Radek]]);
    ZLData.ValueByName('Firm_ID') := Firm_Id;
    ZLData.ValueByName('ConstSymbol_ID') := '0000308000';
    ZLData.ValueByName('VarSymbol') := Cells[1, Radek];
    ZLData.ValueByName('DueDate$DATE') := Floor(DatumSplatnosti);        // 16.9.22 musí být až tady
// 28.6.2018 zakázky pro ÈTÚ - mohou být rùzné podle smlouvy
    with qrAbra do begin
      Close;
      BusOrderCode := qrSmlouva.FieldByName('CTU').AsString;
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
// kolekce pro øádky ZL
    RowsCollection := ZLData.Value[dmCommon.IndexByName(ZLData, 'Rows')];
// 1. øádek s textem
    RadkyZL:= AbraOLE.CreateValues('@IssuedDepositInvoiceRow');
    RadkyZL.ValueByName('BusOrder_ID') := BusOrder_Id;                     // 30.4.21
    RadkyZL.ValueByName('BusTransaction_ID') := BusTransaction_Id;         // 30.4.21
    RadkyZL.ValueByName('Division_ID') := '1000000101';
    RadkyZL.ValueByName('RowType') := '0';
    RadkyZL.ValueByName('Text') := Format('Úètujeme Vám na období  %s', [Obdobi]);
    RadkyZL.ValueByName('UnitRate') := 1;
    RowsCollection.Add(RadkyZL);
// další øádky faktury se vytvoøí z qrSmlouva
    while not EOF do begin
      CenaPripojeni := 0;
// cena za pøipojení k Internetu (je tarif, tj. mùže to být i televize)
      if (FieldByName('BItTariff').AsInteger = 1) then CenaPripojeni := qrSmlouva.FieldByName('BItPrice').AsFloat;
// smlouva na pøipojení
      if CenaPripojeni <> 0 then begin
        RadkyZL:= AbraOLE.CreateValues('@IssuedDepositInvoiceRow');
        RadkyZL.ValueByName('BusOrder_ID') := BusOrder_Id;
        RadkyZL.ValueByName('BusTransaction_ID') := BusTransaction_Id;
        RadkyZL.ValueByName('Division_ID') := '1000000101';
        RadkyZL.ValueByName('RowType') := '4';
        RadkyZL.ValueByName('Text') :=
          Format('podle smlouvy  %s  službu  %s', [FieldByName('Number').AsString, FieldByName('BItDescription').AsString]);
        RadkyZL.ValueByName('TAmount') := Format('%f', [qrSmlouva.FieldByName('Period').AsInteger * qrSmlouva.FieldByName('BItPrice').AsFloat]);
        RadkyZL.ValueByName('UnitRate') := 1;
        RowsCollection.Add(RadkyZL);
// platba na 6 mìsícù (sleva jen za internet) - neplatí pro tarify Misauer a Harna 30.3.20 a Cukrák
// 29.7.2019 slevy neplatí ani pro tarif Lahovská
        if (qrSmlouva.FieldByName('Period').AsInteger = 6) and (qrSmlouva.FieldByName('Typ').AsString = 'InternetContract')
         and not (qrSmlouva.FieldByName('co_tariff_id').AsInteger in [202..234, 242, 245..247]) then begin   // Misauer 202 - 224, Harna 225 - 234, Lahovská 242, Cukrák 245 - 247
          RadkyZL:= AbraOLE.CreateValues('@IssuedDepositInvoiceRow');
          RadkyZL.ValueByName('BusOrder_ID') := BusOrder_Id;
          RadkyZL.ValueByName('BusTransaction_ID') := BusTransaction_Id;
          RadkyZL.ValueByName('Division_ID') := '1000000101';
          RadkyZL.ValueByName('RowType') := '4';
          RadkyZL.ValueByName('Text') := 'slevu';
          RadkyZL.ValueByName('TAmount') := Format('%f', [-CenaPripojeni]);
          RadkyZL.ValueByName('UnitRate') := 1;
          RowsCollection.Add(RadkyZL);
        end;
// platba na 12 mìsícù (sleva jen za internet) - neplatí pro tarify Misauer a Harna 30.3.20 a Cukrák
// 29.7.2019 slevy neplatí ani pro tarif Lahovská
        if (qrSmlouva.FieldByName('Period').AsInteger = 12) and (qrSmlouva.FieldByName('Typ').AsString = 'InternetContract')
         and not (qrSmlouva.FieldByName('co_tariff_id').AsInteger in [202..234, 242, 245..247]) then begin   // Misauer 202 - 224, Harna 225 - 234, Lahovská 242, Cukrák 245 - 247
          RadkyZL:= AbraOLE.CreateValues('@IssuedDepositInvoiceRow');
          RadkyZL.ValueByName('BusOrder_ID') := BusOrder_Id;
          RadkyZL.ValueByName('BusTransaction_ID') := BusTransaction_Id;
          RadkyZL.ValueByName('Division_ID') := '1000000101';
          RadkyZL.ValueByName('RowType') := '4';
          RadkyZL.ValueByName('Text') := 'slevu';
          RadkyZL.ValueByName('TAmount') := Format('%f', [-2*CenaPripojeni]);
          RadkyZL.ValueByName('UnitRate') := 1;
          RowsCollection.Add(RadkyZL);
        end;
        CenaPripojeni := 0;
      end
// nìco jiného než platba za pøipojení (CenaPripojeni = 0)
      else begin
        RadkyZL:= AbraOLE.CreateValues('@IssuedDepositInvoiceRow');
        RadkyZL.ValueByName('BusOrder_ID') := BusOrder_Id;                          // 30.4.21
        RadkyZL.ValueByName('BusTransaction_ID') := BusTransaction_Id;              // 30.4.21
        RadkyZL.ValueByName('Division_ID') := '1000000101';
        RadkyZL.ValueByName('RowType') := '4';
        RadkyZL.ValueByName('Text') := FieldByName('BItDescription').AsString;
        RadkyZL.ValueByName('TAmount') := Format('%f', [qrSmlouva.FieldByName('Period').AsInteger * FieldByName('BItPrice').AsFloat]);
// platba na 6 mìsícù
        if (qrSmlouva.FieldByName('Period').AsInteger = 6) then
          if (qrSmlouva.FieldByName('co_tariff_id').AsInteger in [202..234]) then
            RadkyZL.ValueByName('TAmount') := Format('%f', [6 * FieldByName('BItPrice').AsFloat])   // tarify Misauer a Harna
          else RadkyZL.ValueByName('TAmount') := Format('%f', [5 * FieldByName('BItPrice').AsFloat])  // ostatní
// platba na 12 mìsícù
        else if qrSmlouva.FieldByName('Period').AsInteger = 12 then
          if (qrSmlouva.FieldByName('co_tariff_id').AsInteger in [202..234]) then
            RadkyZL.ValueByName('TAmount') := Format('%f', [12 * FieldByName('BItPrice').AsFloat])   // tarify Misauer a Harna
          else RadkyZL.ValueByName('TAmount') := Format('%f', [10 * FieldByName('BItPrice').AsFloat])  // ostatní
// ostatní
        else RadkyZL.ValueByName('TAmount') := Format('%f', [qrSmlouva.FieldByName('Period').AsInteger * FieldByName('BItPrice').AsFloat]);
        RadkyZL.ValueByName('UnitRate') := 1;
        RowsCollection.Add(RadkyZL);
      end;
      Next;   //  qrSmlouva
    end;  //  while not qrSmlouva.EOF
    Close;   // qrSmlouva
    try
      ID := ZLObject.CreateNewFromValues(ZLData);
      ZLData := ZLObject.GetValues(ID);
      Cena := Round(double(ZLData.Value[dmCommon.IndexByName(ZLData, 'Amount')]));
      CisloZL := string(ZLData.Value[dmCommon.IndexByName(ZLData, 'DisplayName')]);
      fmMain.Zprava(Format('%s (%s): Vytvoøen zálohový list %s.', [Zakaznik, Cells[1, Radek], CisloZL]));
      Ints[0, Radek] := 0;
      Cells[2, Radek] := string(ZLData.Value[dmCommon.IndexByName(ZLData, 'OrdNumber')]);        // ZL
      Cells[3, Radek] := IntToStr(Cena);                                                         // èástka
      Cells[4, Radek] := Zakaznik;                                                               // jméno
    except
      on E: Exception do begin
        fmMain.Zprava(Format('%s (%s): Chyba v Abøe - %s', [Zakaznik, Cells[1, Radek], E.Message]));
        if Application.MessageBox(PChar(E.Message + ^M + 'Pokraèovat?'), 'Chyba',
         MB_YESNO + MB_ICONQUESTION) = IDNO then Prerusit := True;
      end;
    end;
  end;  // with
end;  // procedury ZLAbra

end.

