// od 19.12.2014 - Z�lohov� listy na p�ipojen� se generuj� v Ab�e podle datab�ze iQuest. Podle M�s��n� fakturace.
// 29.7.2019 slevy za 6 a 12 m�s�c� neplat� ani pro tarify Lahovsk�
// 30.4.2020 BusOrder a BusTransaction do v�ech ��dk� ZL
// 28.7.2024 dal�� tarify bez slev, trochu upraveno

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

type
  TTarify = set of byte;     // 28.7.2024 seznam tarif�, kter� nemaj� slevu p�i placen� na 6 nebo 12 m�s�c�
// Proto�e nejni��� dot�en� tarif m� Id 202, bude se od Id tarifu ode��tat 200, aby i nejvy��� Id
// (dnes 330) byl je�t� byte

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
    42,                // Lahovsk�
    45..47,            // Cukr�k
    119, 123..126,     // DSL*
    127..130           // optical*
  ];
  testVytvoreni := fmMain.chbTestVytvoreni.Checked;

  with fmMain, DesU.qrZakos, asgMain do begin
    CustomerVarSymbol := Cells[1, Radek];

// nen� v�c z�kazn�k� pro jeden VS ? 28.6.2018
    Close;
    SQL.Text := 'SELECT COUNT(*) FROM customers'
    + ' WHERE variable_symbol = ' + Ap + CustomerVarSymbol + Ap;
    Open;
    if Fields[0].AsInteger > 1 then begin
      fmMain.Zprava(Format('Variabiln� symbol %s m� v�ce z�kazn�k�.', [CustomerVarSymbol]));
      Close;
      Exit;
    end;

    // vyhled�n� �daj� o smlouv�ch
    Close;

    SQLStr := 'CALL get_deposit_invoicing_by_vs('
        + Ap + FormatDateTime('yyyy-mm-dd', deDatumDokladu.Date + 30) + ApC
        + Ap + FormatDateTime('yyyy-mm-dd', deDatumSplatnosti.Date) + ApC
        + Ap + CustomerVarSymbol + ApC;

    case argMesic.ItemIndex of
      0, 2: SQLStr := SQLStr + 'false,false)'; //prvn� parametr ��k�, zda na��t�me 6m periody, druh�, zda 12m periody
      1: SQLStr := SQLStr +  'true,false)'; // dtto
      3: SQLStr := SQLStr +  'true,true)'; // dtto
    end;

    SQL.Text := SQLStr;

    // p�ejmenov�no takto: AbraCode (cu_abra_code), Typ (co_type), Number (co_number), InvoiceSending (cu_invoice_sending_method_name), Period(bb_period),
    // BItDescription (bi_description), BItPrice (bi_price), BItTariff (bi_is_tariff), TarifId (co_tariff_id), CTU (co_ctu_category)'

    Open;
    // p�i lecjak� chyb� v datab�zi (nap�. Tariff_Id je NULL) konec
    if RecordCount = 0 then begin
      fmMain.Zprava(Format('Pro variabiln� symbol %s nen� co generovat.', [CustomerVarSymbol]));
      Close;
      Exit;
    end;
    if FieldByName('cu_abra_code').AsString = '' then begin
      fmMain.Zprava(Format('Smlouva %s: z�kazn�k nem� k�d Abry.', [FieldByName('co_number').AsString]));
      Close;
      Exit;
    end;

    // 15.8.2022 {
    {if (FieldByName('CTU').AsString = '') and (FieldByName('co_type').AsString = 'InternetContract') then begin
      fmMain.Zprava(Format('Smlouva %s: z�kazn�k nem� k�d pro �T�.', [FieldByName('co_number').AsString]));
      Close;
      Exit;
    end;  }

    with DesU.qrAbra do begin
    // kontrola k�du firmy, p�i chyb� konec
      Close;
      SQLStr := 'SELECT F.ID as Firm_ID, F.Name as FirmName, F.OrgIdentNumber FROM Firms F'
      + ' WHERE Code = ' + Ap + DesU.qrZakos.FieldByName('cu_abra_code').AsString  + Ap
      + ' AND F.Firm_ID IS NULL'         // bez n�sledovn�k�
      + ' AND F.Hidden = ''N'''
      + ' ORDER BY F.ID DESC';
      SQL.Text := SQLStr;
      Open;
      if RecordCount = 0 then begin
        fmMain.Zprava(Format('Smlouva %s: z�kazn�k s k�dem %s nen� v adres��i Abry.',
         [DesU.qrZakos.FieldByName('co_number').AsString, DesU.qrZakos.FieldByName('cu_abra_code').AsString]));
        Exit;
      end
      else if RecordCount > 1 then begin
        fmMain.Zprava(Format('Smlouva %s: V Ab�e je v�ce z�kazn�k� s k�dem %s.',
          [DesU.qrZakos.FieldByName('co_number').AsString, DesU.qrZakos.FieldByName('cu_abra_code').AsString]));
        Exit;
      end
      else begin
        Firm_Id := FieldByName('Firm_ID').AsString;
        FirmName := FieldByName('FirmName').AsString;
        OrgIdentNumber := Trim(FieldByName('OrgIdentNumber').AsString);
      end;

      // 28.6.2018 obchodn� p��pady pro �T� - plat� pro celou faktury
      if OrgIdentNumber = '' then BusTransactionCode := 'F'               //  fyzick� osoba
      else BusTransactionCode := 'P';                         //  pr�vnick� osoba
      BusTransaction_Id := AbraEnt.getBusTransaction('Code=' + BusTransactionCode).ID;

      // kontrola posledn�ho ZL

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
          fmMain.Zprava(Format('%s, variabiln� symbol %s: %d. z�lohov� list se stejn�m datem.',
          [FirmName, CustomerVarSymbol, RecordCount + 1]));
          Dotaz := Application.MessageBox(PChar(Format('Pro variabiln� symbol %s existuje z�lohov� list ZL1-%s s datem %s na ��stku %s K�. M� se vytvo�it dal��?',
          [CustomerVarSymbol, FieldByName('OrdNumber').AsString, DateToStr(FieldByName('DocDate$DATE').AsFloat), FieldByName('Amount').AsString])),
          'Pozor', MB_ICONQUESTION + MB_YESNOCANCEL + MB_DEFBUTTON1);
          if Dotaz = IDNO then begin
            fmMain.Zprava('Ru�n� z�sah - z�lohov� list nevytvo�en.');
            Exit;
          end else if Dotaz = IDCANCEL then begin
            fmMain.Zprava('Ru�n� z�sah - vytv��en� z�lohov�ch list� ukon�eno.');
            Prerusit := True;
            Exit;
          end else fmMain.Zprava('Ru�n� z�sah - z�lohov� list se vytvo��.');
        end;  //   RecordCount > 0
      end;
    end;  // with qrAbra

    if testVytvoreni then Exit;

    fmMain.Zprava(Format('%s, variabiln� symbol %s.', [FirmName, CustomerVarSymbol]));

    // obdob�  23.9.17
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


    // hlavi�ka faktury
    newIDI := TNewAbraBo.Create('issueddepositinvoice');
    newIDI.addInvoiceParams(Floor(deDatumDokladu.Date));
    newIDI.Item['Varsymbol'] := CustomerVarSymbol;
    newIDI.Item['DocQueue_ID'] := AbraEnt.getDocQueue('Code=ZL1').ID;
    newIDI.Item['Description'] := Format('P�ipojen� %s, %s', [Obdobi, CustomerVarSymbol]);
    newIDI.Item['Firm_ID'] := Firm_Id;
    newIDI.Item['DueDate$DATE'] := Floor(deDatumSplatnosti.Date);

    // vytvo�� se objekt TNewDesInvoiceAA a pak zbytek hlavi�ky ZL
    //* NewInvoice := TNewDesInvoiceAA.create(Floor(deDatumDokladu.Date), CustomerVarSymbol, '10');

    //ZLObject := AbraOLE.CreateObject('@IssuedDepositInvoice');
    //ZLData := AbraOLE.CreateValues('@IssuedDepositInvoice');
    //ZLObject.PrefillValues(ZLData);
    //* NewInvoice.AA['DocQueue_ID'] := AbraEnt.getDocQueue('Code=ZL1').ID;
    //NewInvoice.AA['Period_ID'] := Period_Id; // *HW* automaticky z data dokladu
    //* NewInvoice.AA['Description'] := Format('P�ipojen� %s, %s', [Obdobi, CustomerVarSymbol]);
    //* NewInvoice.AA['Firm_ID'] := Firm_Id;
    //* NewInvoice.AA['DueDate$DATE'] := Floor(deDatumSplatnosti.Date);        // 16.9.22 mus� b�t a� tady

    {  CTU, neni uz potreba
    // 28.6.2018 zak�zky pro �T� - mohou b�t r�zn� podle smlouvy
    with qrAbra do begin
      Close;
      BusOrderCode := FieldByName('co_ctu_category').AsString;
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
    }


    // 1. ��dek s textem
    newIDI.createNewInvoiceRow(0, Format('��tujeme V�m na obdob�  %s', [Obdobi]));

    // dal�� ��dky faktury se vytvo�� z DesU.qrZakos
    while not EOF do begin
      CenaPripojeni := 0;
      // cena za p�ipojen� k Internetu (je tarif, tj. m��e to b�t i televize - v praxi se televize ale na ZL ned�v�, aby si z�kazn�ci mohli m�nit TV tarif dle pot�eby)
      if (FieldByName('bi_is_tariff').AsInteger = 1) then
        CenaPripojeni := FieldByName('bi_price').AsFloat;
      // smlouva na p�ipojen�
      if CenaPripojeni <> 0 then begin
        newIDI.createNewInvoiceRow(4,
          Format('podle smlouvy  %s  slu�bu  %s', [FieldByName('co_number').AsString, FieldByName('bi_description').AsString]), false);
        //boRowAA['BusOrder_ID'] := BusOrder_Id;
        newIDI.rowItem['BusTransaction_ID'] := BusTransaction_Id;
        newIDI.rowItem['TAmount'] := Format('%f', [FieldByName('bb_period').AsInteger * FieldByName('bi_price').AsFloat]);

        // platba na 6 m�s�c� (sleva jen za internet) - neplat� pro tarify Misauer a Harna 30.3.20 a Cukr�k
        // 29.7.2019 slevy neplat� ani pro tarif Lahovsk�
        // 28.7.2024 slevy neplat� ani pro tarify DSL* a optical*
        if (FieldByName('bb_period').AsInteger = 6) and (FieldByName('co_type').AsString = 'InternetContract')
        // Misauer 202 - 224, Harna 225 - 234, Lahovsk� 242, Cukr�k 245 - 247, DSL* 319 a 323-326, optical* 327-330
//         and not (FieldByName('co_tariff_id').AsInteger in [202..234, 242, 245..247]) then begin   // Misauer 202 - 224, Harna 225 - 234, Lahovsk� 242, Cukr�k 245 - 247
         and not (FieldByName('co_tariff_id').AsInteger - 200 in Tarify_bez_slevy) then begin        // 28.7.24
          //RadkyZL:= AbraOLE.CreateValues('@IssuedDepositInvoiceRow');
          newIDI.createNewInvoiceRow(4, 'slevu', false);
          //boRowAA['BusOrder_ID'] := BusOrder_Id;
          newIDI.rowItem['BusTransaction_ID'] := BusTransaction_Id;
          newIDI.rowItem['TAmount'] := Format('%f', [-CenaPripojeni]);
        end;

        // platba na 12 m�s�c� (sleva jen za internet) - neplat� pro tarify Misauer a Harna 30.3.20 a Cukr�k
        // 29.7.2019 slevy neplat� ani pro tarif Lahovsk�
        // 28.7.2024 slevy neplat� ani pro tarify DSL* a optical*
        if (FieldByName('bb_period').AsInteger = 12) and (FieldByName('co_type').AsString = 'InternetContract')
        // Misauer 202 - 224, Harna 225 - 234, Lahovsk� 242, Cukr�k 245 - 247, DSL* 319 a 323-326, optical* 327-330
//         and not (FieldByName('co_tariff_id').AsInteger in [202..234, 242, 245..247]) then begin   // Misauer 202 - 224, Harna 225 - 234, Lahovsk� 242, Cukr�k 245 - 247
         and not (FieldByName('co_tariff_id').AsInteger - 200 in Tarify_bez_slevy) then begin        // 28.7.24
          newIDI.createNewInvoiceRow(4, 'slevu', false);
          //boRowAA['BusOrder_ID'] := BusOrder_Id;
          newIDI.rowItem['BusTransaction_ID'] := BusTransaction_Id;
          newIDI.rowItem['TAmount'] := Format('%f', [-2*CenaPripojeni]);
        end;
        CenaPripojeni := 0; // HW asi zbyte�n� tady
      end
      // n�co jin�ho ne� platba za p�ipojen� (CenaPripojeni = 0)
      else begin
        newIDI.createNewInvoiceRow(4, FieldByName('bi_description').AsString, false);
        //boRowAA['BusOrder_ID'] := BusOrder_Id;                          // 30.4.21
        newIDI.rowItem['BusTransaction_ID'] := BusTransaction_Id;              // 30.4.21
        newIDI.rowItem['TAmount'] := Format('%f', [FieldByName('bb_period').AsInteger * FieldByName('bi_price').AsFloat]);
        // platba na 6 m�s�c�
        if (FieldByName('bb_period').AsInteger = 6) then
//          if (FieldByName('co_tariff_id').AsInteger in [202..234]) then
          if (FieldByName('co_tariff_id').AsInteger - 200 in Tarify_bez_slevy) then                    // 28.7.24 tarify bez slevy
            newIDI.rowItem['TAmount'] := Format('%f', [6 * FieldByName('bi_price').AsFloat])
          else newIDI.rowItem['TAmount'] := Format('%f', [5 * FieldByName('bi_price').AsFloat])  // ostatn�
        // platba na 12 m�s�c�
        else if FieldByName('bb_period').AsInteger = 12 then
          if (FieldByName('co_tariff_id').AsInteger - 200 in Tarify_bez_slevy) then                     // 28.7.24 tarify bez slevy
            newIDI.rowItem['TAmount'] := Format('%f', [12 * FieldByName('bi_price').AsFloat])
          else newIDI.rowItem['TAmount'] := Format('%f', [10 * FieldByName('bi_price').AsFloat])  // ostatn�
        // ostatn�
        else newIDI.rowItem['TAmount'] := Format('%f', [FieldByName('bb_period').AsInteger * FieldByName('bi_price').AsFloat]);
      end;
      Next;   //  DesU.qrZakos
    end;  //  while not DesU.qrZakos.EOF
    Close;   // DesU.qrZakos


    try
      //* abraWebApiResponse := DesU.abraBoCreate(NewInvoice.AA.AsJSon, 'issueddepositinvoice');
      newIDI.writeToAbra;
      if newIDI.WriteResult.isOk then begin
        fmMain.Zprava(Format('%s (%s): Vytvo�en z�lohov� list %s.', [FirmName, CustomerVarSymbol, newIDI.getCreatedBoItem('displayname')]));
        Ints[0, Radek] := 0;
        Cells[2, Radek] := newIDI.getCreatedBoItem('ordnumber'); // faktura
        Cells[3, Radek] := newIDI.getCreatedBoItem('amount');  // ��stka
        Cells[4, Radek] := FirmName;
      end else begin
        fmMain.Zprava(Format('%s (%s): Chyba %s: %s', [FirmName, CustomerVarSymbol, newIDI.WriteResult.Code, newIDI.WriteResult.Messg]));
        if Dialogs.MessageDlg( '(' + newIDI.WriteResult.Code + ') '
           + newIDI.WriteResult.Messg + sLineBreak + 'Pokra�ovat?',
           mtConfirmation, [mbYes, mbNo], 0 ) = mrNo then Prerusit := True;
      end;
    except on E: exception do
      begin
        Application.MessageBox(PChar('Problem ' + ^M + E.Message), 'Vytvo�en� fa');
      end;
    end;


  end;  // with

end;  // procedury ZLAbra




end.

