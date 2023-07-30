// od 19.12.2014 - Z�lohov� listy na p�ipojen� se generuj� v Ab�e podle datab�ze iQuest. Podle M�s��n� fakturace.

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
// p�ipojen� k Ab�e
    Screen.Cursor := crHourGlass;
    asgMain.Visible := False;
    lbxLog.Visible := True;
    try
      dmCommon.Zprava('P�ipojen� k Ab�e ...');
      AbraOLE := CreateOLEObject('AbraOLE.Application');
      if not AbraOLE.Connect('@' + AbraConnection) then begin
        dmCommon.Zprava('Probl�m s Abrou (connect  "' + AbraConnection + '").');
        Screen.Cursor := crDefault;
        Exit;
      end;
      dmCommon.Zprava('OK');
      dmCommon.Zprava('Login ...');
      if not AbraOLE.Login('Supervisor', '') then begin
        dmCommon.Zprava('Probl�m s Abrou (login).');
        Screen.Cursor := crDefault;
        Exit;
      end;
      dmCommon.Zprava('OK');
    except on E: exception do
      begin
        Application.MessageBox(PChar('Probl�m s Abrou.' + ^M + E.Message), 'Abra', MB_ICONERROR + MB_OK);
        dmCommon.Zprava('Probl�m s Abrou.' + #13#10 + E.Message);
        btKonec.Caption := '&Konec';
        Screen.Cursor := crDefault;
        Exit;
      end;
    end;
    dmCommon.Zprava(Format('Po�et ZL k vygenerov�n�: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));
    asgMain.Visible := True;
    lbxLog.Visible := False;
    apbProgress.Position := 0;
    apbProgress.Visible := True;
// hlavn� smy�ka
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
// konec hlavn� smy�ky
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
    dmCommon.Zprava('Generov�n� ZL ukon�eno');
  end;  //  with fmMain
end;

// ------------------------------------------------------------------------------------------------

procedure TdmVytvoreni.ZLAbra(Radek: integer);
// pro m�s�c a rok zadan� v aseMesic a aseRok vytvo�� fakturu za p�ipojen� k internetu a za VoIP
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
// vyhled�n� �daj� o smlouv�ch
      Close;
      SQLStr := 'SELECT AbraCode, Typ, Number, InvoiceSending, Period, BItDescription, BItPrice, BItTariff'
//      + ' FROM ' + ZLView
      + ' WHERE VS = ' + Ap + Cells[1, Radek] + Ap
      + ' ORDER BY Number, BItTariff DESC';
      SQL.Text := SQLStr;
      Open;
// p�i lecjak� chyb� v datab�zi (nap�. Tariff_Id je NULL) konec
      if RecordCount = 0 then begin
        dmCommon.Zprava(Format('Pro variabiln� symbol %s nen� co generovat.', [Cells[1, Radek]]));
        Close;
        Exit;
      end;
      if FieldByName('AbraCode').AsString = '' then begin
        dmCommon.Zprava(Format('Smlouva %s: z�kazn�k nem� k�d Abry.', [FieldByName('Smlouva').AsString]));
        Close;
        Exit;
      end;
      with qrAbra do begin
// kontrola k�du firmy, p�i chyb� konec
        Close;
        SQLStr := 'SELECT F.ID, F.Name, FO.Id FROM Firms F, FirmOffices FO'
        + ' WHERE Code = ' + Ap + qrSmlouva.FieldByName('AbraCode').AsString + Ap
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
          Zakaznik := UTF8Decode(Fields[1].AsString);
          FirmOffice_Id := Fields[2].AsString;
        end;
// kontrola posledn�ho ZL
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
          dmCommon.Zprava(Format('%s, variabiln� symbol %s: %d. z�lohov� list se stejn�m datem.',
          [Zakaznik, Cells[1, Radek], RecordCount + 1]));
          Dotaz := Application.MessageBox(PChar(Format('Pro variabiln� symbol %s existuje z�lohov� list ZL1-%s s datem %s na ��stku %s K�. M� se vytvo�it dal��?',
          [Cells[1, Radek], FieldByName('OrdNumber').AsString, DateToStr(FieldByName('DocDate$DATE').AsFloat), FieldByName('Amount').AsString])),
          'Pozor', MB_ICONQUESTION + MB_YESNOCANCEL + MB_DEFBUTTON1);
          if Dotaz = IDNO then begin
            dmCommon.Zprava('Ru�n� z�sah - z�lohov� list nevytvo�en.');
            Exit;
          end else if Dotaz = IDCANCEL then begin
            dmCommon.Zprava('Ru�n� z�sah - program ukon�en.');
            Prerusit := True;
            Exit;
          end else dmCommon.Zprava('Ru�n� z�sah - z�lohov� list se vytvo��.');
        end;  //   RecordCount > 0
      end;  // with qrAbra
      dmCommon.Zprava(Format('%s, variabiln� symbol %s.', [Zakaznik, Cells[1, Radek]]));
// vytvo�� se hlavi�ka faktury
      ZLObject := AbraOLE.CreateObject('@IssuedDepositInvoice');
      ZLData := AbraOLE.CreateValues('@IssuedDepositInvoice');
      ZLObject.PrefillValues(ZLData);
      ZLData.ValueByName('DocQueue_ID') := DocQueue_Id;
      ZLData.ValueByName('Period_ID') := Period_Id;
      ZLData.ValueByName('DocDate$DATE') := DatumDokladu;
//      ZLData.ValueByName('DueDate$DATE') := DatumSplatnosti;
      ZLData.ValueByName('CreatedBy_ID') := '2200000101';
      ZLData.ValueByName('Description') := UTF8Encode('P�ipojen�');
      ZLData.ValueByName('Firm_ID') := Firm_Id;
      ZLData.ValueByName('FirmOffice_ID') := FirmOffice_Id;
      ZLData.ValueByName('Address_ID') := MyAddress_Id;
      ZLData.ValueByName('BankAccount_ID') := MyAccount_Id;
      ZLData.ValueByName('ConstSymbol_ID') := '0000308000';
      ZLData.ValueByName('VarSymbol') := Cells[1, Radek];
//    ZLData.ValueByName('TransportationType_ID') :='1000000101';
      ZLData.ValueByName('PaymentType_ID') :='1000000101';
// kolekce pro ��dky ZL
      RowsCollection := ZLData.Value[dmCommon.IndexByName(ZLData, 'Rows')];
// 1. ��dek s textem
      RadkyZL:= AbraOLE.CreateValues('@IssuedDepositInvoiceRow');
      RadkyZL.ValueByName('Division_ID') := '1000000101';
      RadkyZL.ValueByName('RowType') := '0';
      RadkyZL.ValueByName('UnitRate') := 1;
      RowsCollection.Add(RadkyZL);
// dal�� ��dky faktury se vytvo�� z qrSmlouva
      while not EOF do begin
        CenaPripojeni := 0;
// cena za p�ipojen� k Internetu
        if (FieldByName('BItTariff').AsInteger = 1) then CenaPripojeni := qrSmlouva.FieldByName('BItPrice').AsFloat;
// smlouva na p�ipojen�
        if CenaPripojeni <> 0 then begin
          RadkyZL:= AbraOLE.CreateValues('@IssuedDepositInvoiceRow');
          RadkyZL.ValueByName('BusOrder_ID') := '1000000101';
          RadkyZL.ValueByName('Division_ID') := '1000000101';
          RadkyZL.ValueByName('RowType') := '4';
          RadkyZL.ValueByName('Text') :=
            Format('podle smlouvy  %s  slu�bu  %s', [FieldByName('Number').AsString, FieldByName('BItDescription').AsString]);
          RadkyZL.ValueByName('TAmount') := Format('%f', [qrSmlouva.FieldByName('Period').AsInteger * qrSmlouva.FieldByName('BItPrice').AsFloat]);
          RadkyZL.ValueByName('UnitRate') := 1;
          RowsCollection.Add(RadkyZL);
// platba na 6 m�s�c�
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
// platba na 12 m�s�c�
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
// n�co jin�ho ne� platba za p�ipojen� (CenaPripojeni = 0)
        else begin
          RadkyZL:= AbraOLE.CreateValues('@IssuedDepositInvoiceRow');
          RadkyZL.ValueByName('BusOrder_ID') := '1000000101';
          RadkyZL.ValueByName('Division_ID') := '1000000101';
          RadkyZL.ValueByName('RowType') := '4';
          RadkyZL.ValueByName('Text') := FieldByName('BItDescription').AsString;
          RadkyZL.ValueByName('TAmount') := Format('%f', [qrSmlouva.FieldByName('Period').AsInteger * FieldByName('BItPrice').AsFloat]);
// platba na 6 m�s�c�
          if qrSmlouva.FieldByName('Period').AsInteger = 6 then
            RadkyZL.ValueByName('TAmount') := Format('%f', [5 * FieldByName('BItPrice').AsFloat])
// platba na 12 m�s�c�
          else if qrSmlouva.FieldByName('Period').AsInteger = 12 then
            RadkyZL.ValueByName('TAmount') := Format('%f', [10 * FieldByName('BItPrice').AsFloat])
// ostatn�
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
        dmCommon.Zprava(Format('%s (%s): Vytvo�en z�lohov� list %s.', [Zakaznik, Cells[1, Radek], CisloZL]));
        Ints[0, Radek] := 0;
        Cells[2, Radek] := string(ZLData.Value[dmCommon.IndexByName(ZLData, 'OrdNumber')]);        // ZL
        Cells[3, Radek] := IntToStr(Cena);                                                         // ��stka
        Cells[4, Radek] := Zakaznik;                                                               // jm�no
      except
        on E: Exception do begin
          dmCommon.Zprava(Format('%s (%s): Chyba v Ab�e - %s', [Zakaznik, Cells[1, Radek], E.Message]));
          if Application.MessageBox(PChar(E.Message + ^M + 'Pokra�ovat?'), 'Chyba',
           MB_YESNO + MB_ICONQUESTION) = IDNO then Prerusit := True;
        end;
      end;
    end;  // with qrSmlouva
  end;  // with fmMain
end;  // procedury ZLAbra

end.

