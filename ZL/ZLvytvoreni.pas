// od 19.12.2014 - Z�lohov� listy na p�ipojen� se generuj� v Ab�e podle datab�ze iQuest. Podle M�s��n� fakturace.
// 29.7.2019 slevy za 6 a 12 m�s�c� neplat� ani pro tarify Lahovsk�
// 30.4.2020 BusOrder a BusTransaction do v�ech ��dk� ZL

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
// p�ipojen� k Ab�e
    Screen.Cursor := crHourGlass;
    asgMain.Visible := False;
    lbxLog.Visible := True;

    fmMain.Zprava(Format('Po�et ZL k vygenerov�n�: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));
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
    fmMain.Zprava('Generov�n� ZL ukon�eno');
  end;  //  with fmMain
end;

// ------------------------------------------------------------------------------------------------

procedure TdmVytvoreni.ZLAbra(Radek: integer);
// pro dal�� obdob� vytvo�� ZL na p�ipojen� k internetu
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
// nen� v�c z�kazn�k� pro jeden VS ? 28.6.2018
    Close;
    SQLStr := 'SELECT COUNT(*) FROM customers'
    + ' WHERE Variable_Symbol = ' + Ap + Cells[1, Radek] + Ap;
    SQL.Text := SQLStr;
    Open;
    if Fields[0].AsInteger > 1 then begin
      fmMain.Zprava(Format('Variabiln� symbol %s m� v�ce z�kazn�k�.', [Cells[1, Radek]]));
      Close;
      Exit;
    end;
// vyhled�n� �daj� o smlouv�ch
    Close;
    SQLStr := 'SELECT AbraCode, Typ, Number, InvoiceSending, Period, BItDescription, BItPrice, BItTariff, TarifId/co_tariff_id, CTU'
    + ' FROM ' + ZLView
    + ' WHERE VS = ' + Ap + Cells[1, Radek] + Ap
    + ' ORDER BY Number, BItTariff DESC';
    SQL.Text := SQLStr;
    Open;
// p�i lecjak� chyb� v datab�zi (nap�. Tariff_Id je NULL) konec
    if RecordCount = 0 then begin
      fmMain.Zprava(Format('Pro variabiln� symbol %s nen� co generovat.', [Cells[1, Radek]]));
      Close;
      Exit;
    end;
    if FieldByName('AbraCode').AsString = '' then begin
      fmMain.Zprava(Format('Smlouva %s: z�kazn�k nem� k�d Abry.', [FieldByName('Number').AsString]));
      Close;
      Exit;
    end;
// 28.6.2018 opraveno 17.12.
// 15.8.2022 {
{    if (FieldByName('CTU').AsString = '') and (FieldByName('Typ').AsString = 'InternetContract') then begin
      fmMain.Zprava(Format('Smlouva %s: z�kazn�k nem� k�d pro �T�.', [FieldByName('Number').AsString]));
      Close;
      Exit;
    end;  }
    with qrAbra do begin
// kontrola k�du firmy, p�i chyb� konec
      Close;
      SQLStr := 'SELECT F.ID, F.Name, F.OrgIdentNumber, FO.Id FROM Firms F, FirmOffices FO'
      + ' WHERE Code = ' + Ap + qrSmlouva.FieldByName('AbraCode').AsString + Ap
      + ' AND F.Firm_ID IS NULL'         // bez p�edk�
      + ' AND F.Hidden = ''N'''
      + ' AND FO.Parent_Id = F.Id'
      + ' ORDER BY F.ID DESC';
      SQL.Text := SQLStr;
      Open;
      if RecordCount = 0 then begin
        fmMain.Zprava(Format('Smlouva %s: z�kazn�k s k�dem %s nen� v adres��i Abry.',
         [qrSmlouva.FieldByName('Number').AsString, qrSmlouva.FieldByName('AbraKod').AsString]));
        Exit;
      end else begin
        Firm_Id := Fields[0].AsString;
        Zakaznik := UTF8Decode(Fields[1].AsString);
        IC := Trim(Fields[2].AsString);
        FirmOffice_Id := Fields[3].AsString;
      end;
// 28.6.2018 obchodn� p��pady pro �T� - plat� pro celou faktury
      if IC = '' then BusTransactionCode := 'F'               //  fyzick� osoba
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
// kontrola posledn�ho ZL
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
        fmMain.Zprava(Format('%s, variabiln� symbol %s: %d. z�lohov� list se stejn�m datem.',
        [Zakaznik, Cells[1, Radek], RecordCount + 1]));
        Dotaz := Application.MessageBox(PChar(Format('Pro variabiln� symbol %s existuje z�lohov� list ZL1-%s s datem %s na ��stku %s K�. M� se vytvo�it dal��?',
        [Cells[1, Radek], FieldByName('OrdNumber').AsString, DateToStr(FieldByName('DocDate$DATE').AsFloat), FieldByName('Amount').AsString])),
        'Pozor', MB_ICONQUESTION + MB_YESNOCANCEL + MB_DEFBUTTON1);
        if Dotaz = IDNO then begin
          fmMain.Zprava('Ru�n� z�sah - z�lohov� list nevytvo�en.');
          Exit;
        end else if Dotaz = IDCANCEL then begin
          fmMain.Zprava('Ru�n� z�sah - program ukon�en.');
          Prerusit := True;
          Exit;
        end else fmMain.Zprava('Ru�n� z�sah - z�lohov� list se vytvo��.');
      end;  //   RecordCount > 0
    end;  // with qrAbra
    fmMain.Zprava(Format('%s, variabiln� symbol %s.', [Zakaznik, Cells[1, Radek]]));
// obdob�  23.9.17
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
// vytvo�� se hlavi�ka ZL
    ZLObject := AbraOLE.CreateObject('@IssuedDepositInvoice');
    ZLData := AbraOLE.CreateValues('@IssuedDepositInvoice');
    ZLObject.PrefillValues(ZLData);
    ZLData.ValueByName('DocQueue_ID') := AbraEnt.getDocQueue('Code=ZL1').ID;
    //ZLData.ValueByName('Period_ID') := Period_Id; // *HW* automaticky z data dokladu
    ZLData.ValueByName('DocDate$DATE') := Floor(DatumDokladu);
//    ZLData.ValueByName('DueDate$DATE') := Floor(DatumSplatnosti);
    ZLData.ValueByName('Description') := Format('P�ipojen� %s, %s', [Obdobi, Cells[1, Radek]]);
    ZLData.ValueByName('Firm_ID') := Firm_Id;
    ZLData.ValueByName('ConstSymbol_ID') := '0000308000';
    ZLData.ValueByName('VarSymbol') := Cells[1, Radek];
    ZLData.ValueByName('DueDate$DATE') := Floor(DatumSplatnosti);        // 16.9.22 mus� b�t a� tady
// 28.6.2018 zak�zky pro �T� - mohou b�t r�zn� podle smlouvy
    with qrAbra do begin
      Close;
      BusOrderCode := qrSmlouva.FieldByName('CTU').AsString;
// k�dy z tabulky contracts se mus� p�ed�lat
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
// kolekce pro ��dky ZL
    RowsCollection := ZLData.Value[dmCommon.IndexByName(ZLData, 'Rows')];
// 1. ��dek s textem
    RadkyZL:= AbraOLE.CreateValues('@IssuedDepositInvoiceRow');
    RadkyZL.ValueByName('BusOrder_ID') := BusOrder_Id;                     // 30.4.21
    RadkyZL.ValueByName('BusTransaction_ID') := BusTransaction_Id;         // 30.4.21
    RadkyZL.ValueByName('Division_ID') := '1000000101';
    RadkyZL.ValueByName('RowType') := '0';
    RadkyZL.ValueByName('Text') := Format('��tujeme V�m na obdob�  %s', [Obdobi]);
    RadkyZL.ValueByName('UnitRate') := 1;
    RowsCollection.Add(RadkyZL);
// dal�� ��dky faktury se vytvo�� z qrSmlouva
    while not EOF do begin
      CenaPripojeni := 0;
// cena za p�ipojen� k Internetu (je tarif, tj. m��e to b�t i televize)
      if (FieldByName('BItTariff').AsInteger = 1) then CenaPripojeni := qrSmlouva.FieldByName('BItPrice').AsFloat;
// smlouva na p�ipojen�
      if CenaPripojeni <> 0 then begin
        RadkyZL:= AbraOLE.CreateValues('@IssuedDepositInvoiceRow');
        RadkyZL.ValueByName('BusOrder_ID') := BusOrder_Id;
        RadkyZL.ValueByName('BusTransaction_ID') := BusTransaction_Id;
        RadkyZL.ValueByName('Division_ID') := '1000000101';
        RadkyZL.ValueByName('RowType') := '4';
        RadkyZL.ValueByName('Text') :=
          Format('podle smlouvy  %s  slu�bu  %s', [FieldByName('Number').AsString, FieldByName('BItDescription').AsString]);
        RadkyZL.ValueByName('TAmount') := Format('%f', [qrSmlouva.FieldByName('Period').AsInteger * qrSmlouva.FieldByName('BItPrice').AsFloat]);
        RadkyZL.ValueByName('UnitRate') := 1;
        RowsCollection.Add(RadkyZL);
// platba na 6 m�s�c� (sleva jen za internet) - neplat� pro tarify Misauer a Harna 30.3.20 a Cukr�k
// 29.7.2019 slevy neplat� ani pro tarif Lahovsk�
        if (qrSmlouva.FieldByName('Period').AsInteger = 6) and (qrSmlouva.FieldByName('Typ').AsString = 'InternetContract')
         and not (qrSmlouva.FieldByName('co_tariff_id').AsInteger in [202..234, 242, 245..247]) then begin   // Misauer 202 - 224, Harna 225 - 234, Lahovsk� 242, Cukr�k 245 - 247
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
// platba na 12 m�s�c� (sleva jen za internet) - neplat� pro tarify Misauer a Harna 30.3.20 a Cukr�k
// 29.7.2019 slevy neplat� ani pro tarif Lahovsk�
        if (qrSmlouva.FieldByName('Period').AsInteger = 12) and (qrSmlouva.FieldByName('Typ').AsString = 'InternetContract')
         and not (qrSmlouva.FieldByName('co_tariff_id').AsInteger in [202..234, 242, 245..247]) then begin   // Misauer 202 - 224, Harna 225 - 234, Lahovsk� 242, Cukr�k 245 - 247
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
// n�co jin�ho ne� platba za p�ipojen� (CenaPripojeni = 0)
      else begin
        RadkyZL:= AbraOLE.CreateValues('@IssuedDepositInvoiceRow');
        RadkyZL.ValueByName('BusOrder_ID') := BusOrder_Id;                          // 30.4.21
        RadkyZL.ValueByName('BusTransaction_ID') := BusTransaction_Id;              // 30.4.21
        RadkyZL.ValueByName('Division_ID') := '1000000101';
        RadkyZL.ValueByName('RowType') := '4';
        RadkyZL.ValueByName('Text') := FieldByName('BItDescription').AsString;
        RadkyZL.ValueByName('TAmount') := Format('%f', [qrSmlouva.FieldByName('Period').AsInteger * FieldByName('BItPrice').AsFloat]);
// platba na 6 m�s�c�
        if (qrSmlouva.FieldByName('Period').AsInteger = 6) then
          if (qrSmlouva.FieldByName('co_tariff_id').AsInteger in [202..234]) then
            RadkyZL.ValueByName('TAmount') := Format('%f', [6 * FieldByName('BItPrice').AsFloat])   // tarify Misauer a Harna
          else RadkyZL.ValueByName('TAmount') := Format('%f', [5 * FieldByName('BItPrice').AsFloat])  // ostatn�
// platba na 12 m�s�c�
        else if qrSmlouva.FieldByName('Period').AsInteger = 12 then
          if (qrSmlouva.FieldByName('co_tariff_id').AsInteger in [202..234]) then
            RadkyZL.ValueByName('TAmount') := Format('%f', [12 * FieldByName('BItPrice').AsFloat])   // tarify Misauer a Harna
          else RadkyZL.ValueByName('TAmount') := Format('%f', [10 * FieldByName('BItPrice').AsFloat])  // ostatn�
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
      fmMain.Zprava(Format('%s (%s): Vytvo�en z�lohov� list %s.', [Zakaznik, Cells[1, Radek], CisloZL]));
      Ints[0, Radek] := 0;
      Cells[2, Radek] := string(ZLData.Value[dmCommon.IndexByName(ZLData, 'OrdNumber')]);        // ZL
      Cells[3, Radek] := IntToStr(Cena);                                                         // ��stka
      Cells[4, Radek] := Zakaznik;                                                               // jm�no
    except
      on E: Exception do begin
        fmMain.Zprava(Format('%s (%s): Chyba v Ab�e - %s', [Zakaznik, Cells[1, Radek], E.Message]));
        if Application.MessageBox(PChar(E.Message + ^M + 'Pokra�ovat?'), 'Chyba',
         MB_YESNO + MB_ICONQUESTION) = IDNO then Prerusit := True;
      end;
    end;
  end;  // with
end;  // procedury ZLAbra

end.

