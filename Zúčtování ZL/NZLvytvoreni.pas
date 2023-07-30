// od 19.12.2014 - Zálohové listy na pøipojení se generují v Abøe podle databáze iQuest. Podle Mìsíèní fakturace.

unit NZLvytvoreni;

interface

uses
  Windows, Messages, Classes, Forms, Controls, SysUtils, DateUtils, Variants, ComObj, Math, NZLmain;

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

uses NZLcommon;

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
      if not AbraOLE.Login('Supervisor', '') then begin
        dmCommon.Zprava('Problém s Abrou (login).');
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
    dmCommon.Zprava(Format('Poèet ZL k vygenerování: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));
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
    dmCommon.Zprava('Generování ZL ukonèeno');
  end;  //  with fmMain
end;

// ------------------------------------------------------------------------------------------------

procedure TdmVytvoreni.ZLAbra(Radek: integer);
// pro mìsíc a rok zadaný v aseMesic a aseRok vytvoøí fakturu za pøipojení k internetu a za VoIP
var
  FirmOffice_Id: string[10];
  CenaPripojeni: double;
  Dotaz,
  Cena: integer;
  ZLObject,
  ZLData,
  RadkyZL,
  RowsCollection: variant;
  CisloZL,
  Zakaznik,
  SQLStr: AnsiString;
begin
  with fmMain do begin
    with qrSmlouva, asgMain do begin
// vyhledání údajù o smlouvách
      Close;
      SQLStr := 'SELECT AbraCode, Typ, Number, InvoiceSending, Period, BItDescription, BItPrice, BItTariff'
//      + ' FROM ' + ZLView
      + ' WHERE VS = ' + Ap + Cells[1, Radek] + Ap
      + ' ORDER BY Number, BItTariff DESC';
      SQL.Text := SQLStr;
      Open;
// pøi lecjaké chybì v databázi (napø. Tariff_Id je NULL) konec
      if RecordCount = 0 then begin
        dmCommon.Zprava(Format('Pro variabilní symbol %s není co generovat.', [Cells[1, Radek]]));
        Close;
        Exit;
      end;
      if FieldByName('AbraCode').AsString = '' then begin
        dmCommon.Zprava(Format('Smlouva %s: zákazník nemá kód Abry.', [FieldByName('Smlouva').AsString]));
        Close;
        Exit;
      end;
      with qrAbra do begin
// kontrola kódu firmy, pøi chybì konec
        Close;
        SQLStr := 'SELECT F.ID, F.Name, FO.Id FROM Firms F, FirmOffices FO'
        + ' WHERE Code = ' + Ap + qrSmlouva.FieldByName('AbraCode').AsString + Ap
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
          Zakaznik := UTF8Decode(Fields[1].AsString);
          FirmOffice_Id := Fields[2].AsString;
        end;
// kontrola posledního ZL
        Close;
        SQLStr := 'SELECT OrdNumber, DocDate$DATE, Amount FROM IssuedDInvoices'
        + ' WHERE DocQueue_ID = ' + Ap + DocQueue_Id + Ap
        + ' AND VarSymbol = ' + Ap + Cells[1, Radek] + Ap
        + ' AND DocDate$DATE >= ' + FloatToStr(Trunc(VystavenoOd))
        + ' AND DocDate$DATE <= ' + FloatToStr(Trunc(VystavenoDo))
        + ' ORDER BY OrdNumber DESC';
        SQL.Text := SQLStr;
        Open;
        if RecordCount > 0 then begin
          dmCommon.Zprava(Format('%s, variabilní symbol %s: %d. zálohový list se stejným datem.',
          [Zakaznik, Cells[1, Radek], RecordCount + 1]));
          Dotaz := Application.MessageBox(PChar(Format('Pro variabilní symbol %s existuje zálohový list ZL1-%s s datem %s na èástku %s Kè. Má se vytvoøit další?',
          [Cells[1, Radek], FieldByName('OrdNumber').AsString, DateToStr(FieldByName('DocDate$DATE').AsFloat), FieldByName('Amount').AsString])),
          'Pozor', MB_ICONQUESTION + MB_YESNOCANCEL + MB_DEFBUTTON1);
          if Dotaz = IDNO then begin
            dmCommon.Zprava('Ruèní zásah - zálohový list nevytvoøen.');
            Exit;
          end else if Dotaz = IDCANCEL then begin
            dmCommon.Zprava('Ruèní zásah - program ukonèen.');
            Prerusit := True;
            Exit;
          end else dmCommon.Zprava('Ruèní zásah - zálohový list se vytvoøí.');
        end;  //   RecordCount > 0
      end;  // with qrAbra
      dmCommon.Zprava(Format('%s, variabilní symbol %s.', [Zakaznik, Cells[1, Radek]]));
// vytvoøí se hlavièka faktury
      ZLObject := AbraOLE.CreateObject('@IssuedDepositInvoice');
      ZLData := AbraOLE.CreateValues('@IssuedDepositInvoice');
      ZLObject.PrefillValues(ZLData);
      ZLData.ValueByName('DocQueue_ID') := DocQueue_Id;
      ZLData.ValueByName('Period_ID') := Period_Id;
      ZLData.ValueByName('DocDate$DATE') := DatumDokladu;
//      ZLData.ValueByName('DueDate$DATE') := DatumSplatnosti;
      ZLData.ValueByName('CreatedBy_ID') := '2200000101';
      ZLData.ValueByName('Description') := UTF8Encode('Pøipojení');
      ZLData.ValueByName('Firm_ID') := Firm_Id;
      ZLData.ValueByName('FirmOffice_ID') := FirmOffice_Id;
      ZLData.ValueByName('Address_ID') := MyAddress_Id;
      ZLData.ValueByName('BankAccount_ID') := MyAccount_Id;
      ZLData.ValueByName('ConstSymbol_ID') := '0000308000';
      ZLData.ValueByName('VarSymbol') := Cells[1, Radek];
//    ZLData.ValueByName('TransportationType_ID') :='1000000101';
      ZLData.ValueByName('PaymentType_ID') :='1000000101';
// kolekce pro øádky ZL
      RowsCollection := ZLData.Value[dmCommon.IndexByName(ZLData, 'Rows')];
// 1. øádek s textem
      RadkyZL:= AbraOLE.CreateValues('@IssuedDepositInvoiceRow');
      RadkyZL.ValueByName('Division_ID') := '1000000101';
      RadkyZL.ValueByName('RowType') := '0';
      RadkyZL.ValueByName('UnitRate') := 1;
      RowsCollection.Add(RadkyZL);
// další øádky faktury se vytvoøí z qrSmlouva
      while not EOF do begin
        CenaPripojeni := 0;
// cena za pøipojení k Internetu
        if (FieldByName('BItTariff').AsInteger = 1) then CenaPripojeni := qrSmlouva.FieldByName('BItPrice').AsFloat;
// smlouva na pøipojení
        if CenaPripojeni <> 0 then begin
          RadkyZL:= AbraOLE.CreateValues('@IssuedDepositInvoiceRow');
          RadkyZL.ValueByName('BusOrder_ID') := '1000000101';
          RadkyZL.ValueByName('Division_ID') := '1000000101';
          RadkyZL.ValueByName('RowType') := '4';
          RadkyZL.ValueByName('Text') :=
            Format('podle smlouvy  %s  službu  %s', [FieldByName('Number').AsString, FieldByName('BItDescription').AsString]);
          RadkyZL.ValueByName('TAmount') := Format('%f', [qrSmlouva.FieldByName('Period').AsInteger * qrSmlouva.FieldByName('BItPrice').AsFloat]);
          RadkyZL.ValueByName('UnitRate') := 1;
          RowsCollection.Add(RadkyZL);
// platba na 6 mìsícù
          if (qrSmlouva.FieldByName('Period').AsInteger = 6) and (qrSmlouva.FieldByName('Typ').AsString = 'InternetContract') then begin
            RadkyZL:= AbraOLE.CreateValues('@IssuedDepositInvoiceRow');
            RadkyZL.ValueByName('BusOrder_ID') := '1000000101';
            RadkyZL.ValueByName('Division_ID') := '1000000101';
            RadkyZL.ValueByName('RowType') := '4';
            RadkyZL.ValueByName('Text') := 'slevu';
            RadkyZL.ValueByName('TAmount') := Format('%f', [-CenaPripojeni]);
            RadkyZL.ValueByName('UnitRate') := 1;
            RowsCollection.Add(RadkyZL);
          end;
// platba na 12 mìsícù
          if (qrSmlouva.FieldByName('Period').AsInteger = 12) and (qrSmlouva.FieldByName('Typ').AsString = 'InternetContract') then begin
            RadkyZL:= AbraOLE.CreateValues('@IssuedDepositInvoiceRow');
            RadkyZL.ValueByName('BusOrder_ID') := '1000000101';
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
          RadkyZL.ValueByName('BusOrder_ID') := '1000000101';
          RadkyZL.ValueByName('Division_ID') := '1000000101';
          RadkyZL.ValueByName('RowType') := '4';
          RadkyZL.ValueByName('Text') := FieldByName('BItDescription').AsString;
          RadkyZL.ValueByName('TAmount') := Format('%f', [qrSmlouva.FieldByName('Period').AsInteger * FieldByName('BItPrice').AsFloat]);
// platba na 6 mìsícù
          if qrSmlouva.FieldByName('Period').AsInteger = 6 then
            RadkyZL.ValueByName('TAmount') := Format('%f', [5 * FieldByName('BItPrice').AsFloat])
// platba na 12 mìsícù
          else if qrSmlouva.FieldByName('Period').AsInteger = 12 then
            RadkyZL.ValueByName('TAmount') := Format('%f', [10 * FieldByName('BItPrice').AsFloat])
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
        dmCommon.Zprava(Format('%s (%s): Vytvoøen zálohový list %s.', [Zakaznik, Cells[1, Radek], CisloZL]));
        Ints[0, Radek] := 0;
        Cells[2, Radek] := string(ZLData.Value[dmCommon.IndexByName(ZLData, 'OrdNumber')]);        // ZL
        Cells[3, Radek] := IntToStr(Cena);                                                         // èástka
        Cells[4, Radek] := Zakaznik;                                                               // jméno
      except
        on E: Exception do begin
          dmCommon.Zprava(Format('%s (%s): Chyba v Abøe - %s', [Zakaznik, Cells[1, Radek], E.Message]));
          if Application.MessageBox(PChar(E.Message + ^M + 'Pokraèovat?'), 'Chyba',
           MB_YESNO + MB_ICONQUESTION) = IDNO then Prerusit := True;
        end;
      end;
    end;  // with qrSmlouva
  end;  // with fmMain
end;  // procedury ZLAbra

end.

