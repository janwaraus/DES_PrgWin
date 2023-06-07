unit FIcommon;
// 27.1.17 vyèlenìna procedura AktualizaceView jen pro DES

interface

uses
  Windows, SysUtils, Classes, Forms, Controls, DateUtils, Math, Registry, AdvGrid;

type
  TdmCommon = class(TDataModule)
  public
    procedure Zprava(TextZpravy: string);
    procedure AktualizaceView;
    procedure Plneni_asgMain;
  end;

const

  MyAddress_Id: string[10] = '7000000101';
  MyUser_Id: string[10] = '2200000101';          // automatická fakturace
  MyAccount_Id: string[10] = '1400000101';       // Fio
  MyPayment_Id: string[10] = '1000000101';       // typ platby: na bankovní úèet
  DRCTag_Id = 50;


var
  dmCommon: TdmCommon;

implementation

{$R *.dfm}

uses FIMain, FILogin, DesUtils, AbraEntities;

// ------------------------------------------------------------------------------------------------



procedure TdmCommon.Zprava(TextZpravy: string);
// do listboxu a logfile uloží èas a text zprávy
var
  TimeOut: integer;
begin
  with fmMain do begin
    TextZpravy := FormatDateTime('dd.mm.yy hh:nn:ss  ', Now) + TextZpravy;
    lbxLog.Items.Add(TextZpravy);
    lbxLog.ItemIndex := lbxLog.Count - 1;
    Application.ProcessMessages;
    DesUtils.appendToFile(globalAA['LogFileName'], TextZpravy);
  end;
end;


// ------------------------------------------------------------------------------------------------

procedure TdmCommon.AktualizaceView;
// aktualizuje view pro fakturaci databázi zákazníkù
// 27.1.17 celé pøehlednìji
var
  SQLStr: AnsiString;
begin
  with fmMain, DesU.qrZakos do begin
    Close;
// view s variabilními symboly smluv s EP-Home nebo EP-Profi
    SQLStr := 'CREATE OR REPLACE VIEW ' + fiVoipCustomersView
    + ' AS SELECT DISTINCT Variable_symbol FROM customers Cu, contracts C'
    + ' WHERE Cu.Id = C.Customer_Id'
    + ' AND (C.Tariff_Id = 1 OR C.Tariff_Id = 3)'
    + ' AND C.State = ''active'' '   // proè se tady hlídá "C.state = active" a v fiBillingView "C.invoice = 1"
    + ' AND Variable_symbol IS NOT NULL';
    SQL.Text := SQLStr;
    ExecSQL;
// pro testování
    SQLStr := 'CREATE OR REPLACE VIEW fiVoipCustomersView'
    + ' AS SELECT DISTINCT Variable_symbol FROM customers Cu, contracts C'
    + ' WHERE Cu.Id = C.Customer_Id'
    + ' AND (C.Tariff_Id = 1 OR C.Tariff_Id = 3)'
    + ' AND C.State = ''active'' '
    + ' AND Variable_symbol IS NOT NULL';
    //DesUtils.appendToFile(globalAA['LogFileName']+'.txt', SQLStr);
    SQL.Text := SQLStr;
    ExecSQL;
// aktuální data z billing_batches
    SQLStr := 'CREATE OR REPLACE VIEW ' + fiBBmaxView
    + ' AS SELECT Id, Contract_Id, From_date, Period FROM billing_batches B1'
    + ' WHERE From_date = (SELECT MAX(From_date) FROM billing_batches B2'
      + ' WHERE B2.From_date <= ' + Ap + FormatDateTime('yyyy-mm-dd', deDatumPlneni.Date) + Ap
      + ' AND B1.Contract_Id = B2.Contract_Id)';
    SQL.Text := SQLStr;
    ExecSQL;
// pro testování
    SQLStr := 'CREATE OR REPLACE VIEW fiBBmaxView'
    + ' AS SELECT Id, Contract_Id, From_date, Period FROM billing_batches B1'
    + ' WHERE From_date = (SELECT MAX(From_date) FROM billing_batches B2'
      + ' WHERE B2.From_date <= ' + Ap + FormatDateTime('yyyy-mm-dd', deDatumPlneni.Date) + Ap
      + ' AND B1.Contract_Id = B2.Contract_Id)';
    //DesUtils.appendToFile(globalAA['LogFileName']+'.txt', SQLStr);
    SQL.Text := SQLStr;
    ExecSQL;
// billing view k datu fakturace
    SQLStr := 'CREATE OR REPLACE VIEW ' + fiBillingView
    + ' AS SELECT C.Customer_Id, C.Number, C.Type, C.Tariff_Id, C.Activated_at, C.Canceled_at, C.Invoice_from, C.CTU_category,'
    + ' BB.Period, BI.Description, BI.Price, BI.VAT_Id, BI.Tariff'
    + ' FROM ' + fiBBmaxView + ' BB, billing_items BI, contracts C'
    + ' WHERE BB.Id = BI.Billing_batch_Id'
    + ' AND C.Id = BB.Contract_Id'
    + ' AND (C.Invoice = 1 OR (C.State = ''canceled'' AND C.Canceled_at IS NOT NULL'
      + ' AND C.Canceled_at >= ' + Ap + FormatDateTime('yyyy-mm-dd', StartOfTheMonth(deDatumPlneni.Date)) + Ap + '))'
    + ' AND (C.Invoice_from IS NULL OR C.Invoice_from <= ' + Ap + FormatDateTime('yyyy-mm-dd', deDatumPlneni.Date) + ApZ
    + ' AND C.Activated_at <= ' + Ap + FormatDateTime('yyyy-mm-dd', deDatumPlneni.Date) + Ap
    + ' AND BB.Period = 1';
    SQL.Text := SQLStr;
    ExecSQL;
// pro testování
    SQLStr := 'CREATE OR REPLACE VIEW fiBillingView'
    + ' AS SELECT C.Customer_Id, C.Number, C.Type, C.Tariff_Id, C.Activated_at, C.Canceled_at, C.Invoice_from, C.CTU_category,'
    + ' BB.Period, BI.Description, BI.Price, BI.VAT_Id, BI.Tariff'
    + ' FROM fiBBmaxView BB, billing_items BI, contracts C'
    + ' WHERE BB.Id = BI.Billing_batch_Id'
    + ' AND C.Id = BB.Contract_Id'
    + ' AND (C.Invoice = 1 OR (C.State = ''canceled'' AND C.Canceled_at IS NOT NULL'
      + ' AND C.Canceled_at >= ' + Ap + FormatDateTime('yyyy-mm-dd', StartOfTheMonth(deDatumPlneni.Date)) + Ap + '))'
    + ' AND (C.Invoice_from IS NULL OR C.Invoice_from <= ' + Ap + FormatDateTime('yyyy-mm-dd', deDatumPlneni.Date) + ApZ
    + ' AND C.Activated_at <= ' + Ap + FormatDateTime('yyyy-mm-dd', deDatumPlneni.Date) + Ap
    + ' AND BB.Period = 1';
    //DesUtils.appendToFile(globalAA['LogFileName']+'.txt', SQLStr);
    SQL.Text := SQLStr;
    ExecSQL;
// view k datu fakturace
    SQLStr := 'CREATE OR REPLACE VIEW ' + fiInvoiceView
    + ' (VS, Typ, Posilani, Mail, AbraKod, Smlouva, Tarif, AktivniOd, AktivniDo, FakturovatOd, Perioda, Text, Cena, DPH, Tarifni, Reklama, CTU)'
    + ' AS SELECT Variable_symbol, BV.Type, CB1.Name, Postal_mail, Abra_Code, Number, T.Name, Activated_at, Canceled_at, Invoice_from, Period,'
    + ' BV.Description, BV.Price, CB2.Name, Tariff, Disable_mailings, CTU_category'
    + ' FROM customers Cu'
    + ' JOIN ' + fiBillingView + ' BV ON Cu.Id = BV.Customer_Id'
    + ' LEFT JOIN codebooks CB1 ON Cu.Invoice_sending_method_Id = CB1.Id'
    + ' LEFT JOIN codebooks CB2 ON BV.VAT_Id = CB2.Id'
    + ' LEFT JOIN tariffs T ON BV.Tariff_Id = T.Id';
    SQL.Text := SQLStr;
    ExecSQL;
// pro testování
    SQLStr := 'CREATE OR REPLACE VIEW fiInvoiceView'
    + ' (VS, Typ, Posilani, Mail, AbraKod, Smlouva, Tarif, AktivniOd, AktivniDo, FakturovatOd, Perioda, Text, Cena, DPH, Tarifni, Reklama, CTU)'
    + ' AS SELECT Variable_symbol, BV.Type, CB1.Name, Postal_mail, Abra_Code, Number, T.Name, Activated_at, Canceled_at, Invoice_from, Period,'
    + ' BV.Description, BV.Price, CB2.Name, Tariff, Disable_mailings, CTU_category'
    + ' FROM customers Cu'
    + ' JOIN fiBillingView BV ON Cu.Id = BV.Customer_Id'
    + ' LEFT JOIN codebooks CB1 ON Cu.Invoice_sending_method_Id = CB1.Id'
    + ' LEFT JOIN codebooks CB2 ON BV.VAT_Id = CB2.Id'
    + ' LEFT JOIN tariffs T ON BV.Tariff_Id = T.Id';
    //DesUtils.appendToFile(globalAA['LogFileName']+'.txt', SQLStr);
    SQL.Text := SQLStr;
    ExecSQL;
  end;
end;

// ------------------------------------------------------------------------------------------------

procedure TdmCommon.Plneni_asgMain;
var
  Dotaz,
  Radek: integer;
  VarSymbol,
  Firm_ID: string[10];
  Zakaznik,
  FiltrovatPodle,
  ApNavic,
  SqlProcedureName,
  SQLStr: string;
begin
  with fmMain do try
    apnPrevod.Visible := False;
    apnTisk.Visible := False;
    apnMail.Visible := False;
    Screen.Cursor := crHourGlass;

    with asgMain do try
      ClearNormalCells;
      RowCount := 2;

      // ********************//
      // ***  Fakturace  *** //
      // ********************//
      if rbFakturace.Checked then begin          //výbìr zákazníkù/smluv k fakturaci
        // kontrola mìsíce a roku fakturace
        if (aseRok.Value * 12 + aseMesic.Value > YearOf(Date) * 12 + MonthOf(Date) + 1)     // vìtší než pøíští mìsíc, èi menší
         or (aseRok.Value * 12 + aseMesic.Value < YearOf(Date) * 12 + MonthOf(Date) - 1) then begin       // než minulý mìsíc
          SQLStr := Format('Opravdu fakturovat %d. mìsíc roku %d ?', [aseMesic.Value, aseRok.Value]);
          if Application.MessageBox(PChar(SQLStr), 'Pozor', MB_ICONQUESTION + MB_YESNO + MB_DEFBUTTON1) = IDNO then begin
            btVytvorit.Enabled := True;
            btKonec.Caption := '&Konec';
            Exit;
          end;
        end;

        dmCommon.Zprava(Format('Naètení zákazníkù k fakturaci na období %s.%s od VS %s do %s.', [aseMesic.Text, aseRok.Text, aedOd.Text, aedDo.Text]));

        if cbBezVoIP.Checked then dmCommon.Zprava('      - zákazníci bez VoIP');
        if cbSVoIP.Checked then dmCommon.Zprava('      - zákazníci s VoIP');



        Radek := 0;
        apbProgress.Position := 0;
        apbProgress.Visible := True;

        if cbBezVoIP.Checked and not cbSVoIP.Checked then
          SqlProcedureName := 'get_monthly_invoicing_cu_by_vsrange_nonvoip'
        else if cbSVoIP.Checked and not cbBezVoIP.Checked then
          SqlProcedureName := 'get_monthly_invoicing_cu_by_vsrange_voip'
        else
          SqlProcedureName := 'get_monthly_invoicing_cu_by_vsrange_all';

        DesU.qrZakos.SQL.Text := 'CALL ' + SqlProcedureName + '('
          + Ap + FormatDateTime('yyyy-mm-dd', StartOfTheMonth(deDatumPlneni.Date)) + ApC
          + Ap + FormatDateTime('yyyy-mm-dd', deDatumPlneni.Date) + ApC
          + Ap + aedOd.Text + ApC
          + Ap + aedDo.Text + ApZ;

        DesU.qrZakos.Open;
        while not DesU.qrZakos.EOF do begin
          VarSymbol := DesU.qrZakos.FieldByName('cu_variable_symbol').AsString;
          apbProgress.Position := Round(100 * DesU.qrZakos.RecNo / DesU.qrZakos.RecordCount);
          Application.ProcessMessages;

          if Prerusit then begin
            Prerusit := False;
            apbProgress.Position := 0;
            apbProgress.Visible := False;
            btVytvorit.Enabled := True;
            btKonec.Caption := '&Konec';
            Break;  // konec while a tedy skok skoro na konec procedury a konec naèítání
          end;

          with DesU.qrAbra do begin
            // ne Fakturace; HW jen kontroly
            if not rbFakturace.Checked then begin
              Close;
              DesU.dbAbra.Reconnect;
              // faktura(y) v Abøe v mìsíci aseMesic
              SQLStr := 'SELECT DISTINCT F.Id AS Firm_ID, F.Name FROM Firms F, IssuedInvoices II'
              + ' WHERE F.ID = II.Firm_ID'
              + ' AND F.Firm_ID IS NULL'
              + ' AND F.Hidden = ''N'''
              + ' AND II.VarSymbol = ' + Ap + VarSymbol + Ap;
              SQL.Text := SQLStr;
              Open;
              if RecordCount = 0 then begin       // žádná faktura v Abøe
                Zprava(Format('Zákazník s VS %s nemá ještì vystavenou žádnou fakturu.', [VarSymbol]));
                Dotaz := Application.MessageBox(PChar(Format('Zákazník s VS %s nemá ještì vystavenou žádnou fakturu. Je to v poøádku ?',
                [VarSymbol])), 'Pozor', MB_ICONQUESTION + MB_YESNO + MB_DEFBUTTON1);
                if Dotaz = IDYES then begin
                  Zprava('Ruèní zásah - faktura pøeskoèena.');
                  DesU.qrZakos.Next;
                  Continue;
                end else if Dotaz = IDNO then begin
                  Zprava('Ruèní zásah - generování ukonèeno.');
                  Screen.Cursor := crDefault;
                  Exit;
                end;
              end;
              Firm_ID := FieldByName('Firm_ID').AsString;
              Zakaznik := FieldByName('Name').AsString;
              Close;
              SQLStr := 'SELECT ID, OrdNumber, Amount, VATDate$DATE, DocDate$DATE FROM IssuedInvoices II'
              + ' WHERE Firm_ID = ' + Ap + Firm_ID + Ap
              + ' AND VATDate$DATE >= ' + FloatToStr(Trunc(StartOfAMonth(aseRok.Value, aseMesic.Value)))
              + ' AND VATDate$DATE <= ' + FloatToStr(Trunc(EndOfAMonth(aseRok.Value, aseMesic.Value)));

              SQLStr := SQLStr + ' AND DocQueue_ID = ' + Ap + AbraEnt.getDocQueue('Code=FO1').id + Ap;

              SQL.Text := SQLStr;
              Open;
              if RecordCount = 0 then begin       // faktura v Abøe neexistuje
                Zprava(Format('%s: Faktura na období %s.%s neexistuje.', [Zakaznik, aseMesic.Text, aseRok.Text]));
                Dotaz := Application.MessageBox(PChar(Format('%s: Faktura na období %s.%s neexistuje. Je to v poøádku ?',
                [Zakaznik, aseMesic.Text, aseRok.Text])), 'Pozor', MB_ICONQUESTION + MB_YESNO + MB_DEFBUTTON1);
                if Dotaz = IDYES then begin
                  Zprava('Ruèní zásah - faktura pøeskoèena.');
                  DesU.qrZakos.Next;
                  Continue;
                end else if Dotaz = IDNO then begin
                  Zprava('Ruèní zásah - generování ukonèeno.');
                  Screen.Cursor := crDefault;
                  Exit;
                end;
              end
              else if RecordCount > 1 then
              begin      // více faktur s jedním datem
                Zprava(Format('%s: Více faktur na období %s.%s.', [Zakaznik, aseMesic.Text, aseRok.Text]));
                Dotaz := Application.MessageBox(PChar(Format('%s: Více faktur na období %s.%s. Je to v poøádku ?',
                [Zakaznik, aseMesic.Text, aseRok.Text])), 'Pozor', MB_ICONQUESTION + MB_YESNOCANCEL + MB_DEFBUTTON1);
                if Dotaz = IDNO then begin
                  Zprava('Ruèní zásah - faktura pøeskoèena.');
                  DesU.qrZakos.Next;
                 Continue;
                end else if Dotaz = IDCANCEL then begin
                  Zprava('Ruèní zásah - generování ukonèeno.');
                  Screen.Cursor := crDefault;
                  Exit;
                end;
              end;  // if RecordCount
            end;  // if not rbFakturace.Checked

            // uložení do asgMain
            Inc(Radek);
            RowCount := Radek + 1;
            AddCheckBox(0, Radek, True, True);
            Ints[0, Radek] := 1;                                                   // fajfka
            Cells[1, Radek] := VarSymbol;                                          // VS
            if rbFakturace.Checked then begin
              Cells[2, Radek] := '';                                                 // faktura
              Cells[3, Radek] := '';                                                 // èástka
              Cells[4, Radek] := DesU.qrZakos.FieldByName('cu_abra_code').AsString;             // jméno
            end else begin
              Cells[2, Radek] := Format('%5.5d', [FieldByName('OrdNumber').AsInteger]);     // faktura
              Floats[3, Radek] := FieldByName('Amount').AsFloat;                    // èástka
              Cells[4, Radek] := Zakaznik;                                           // jméno
              Cells[7, Radek] := FieldByName('ID').AsString;                         // ID faktury
              Cells[8, Radek] := DateToStr(FieldByName('VATDate$DATE').AsFloat);
              Cells[9, Radek] := DateToStr(FieldByName('DocDate$DATE').AsFloat);
            end;  // if rbFakturace.Checked else...
            Cells[5, Radek] := DesU.qrZakos.FieldByName('cu_postal_mail').AsString;                // mail
            Ints[6, Radek] := DesU.qrZakos.FieldByName('cu_disable_mailings').AsInteger;             // reklama
            Application.ProcessMessages;
          end;  // with DesU.qrAbra
          DesU.qrZakos.Next;
        end;  // while not EOF
        DesU.qrZakos.Close;


      end
      else // if rbFakturace.Checked

      // *****************************//
      // ***  Pøevod, Tisk, Mail  *** //
      // *****************************//
      with DesU.qrAbra do
      begin
        if rbVyberPodleVS.Checked then begin
          FiltrovatPodle := 'VarSymbol';
          ApNavic := Ap;
          dmCommon.Zprava(Format('Naètení faktur s VS od %s do %s.', [aedOd.Text, aedDo.Text]));
        end;
        if rbVyberPodleFaktury.Checked then begin
          FiltrovatPodle := 'OrdNumber';
          ApNavic := '';
          dmCommon.Zprava(Format('Naètení faktur s èísly od %s do %s.', [aedOd.Text, aedDo.Text]));
        end;
// faktura(y) v Abøe v mìsíci aseMesic
        SQLStr := 'SELECT II.ID, F.Name, F.Code as Abrakod, II.OrdNumber, II.VarSymbol, II.Amount, II.VATDate$DATE, II.DocDate$DATE '
        + 'FROM IssuedInvoices II, Firms F'
        + ' WHERE II.Firm_ID = F.ID'
        + ' AND ' + FiltrovatPodle + ' >= ' + ApNavic + aedOd.Text + ApNavic
        + ' AND ' + FiltrovatPodle + ' <= ' + ApNavic + aedDo.Text + ApNavic
        + ' AND DocQueue_ID = ' + Ap + AbraEnt.getDocQueue('Code=FO1').id + Ap
        + ' AND VATDate$DATE >= ' + FloatToStr(Trunc(StartOfAMonth(aseRok.Value, aseMesic.Value)))  // aby se odfiltrovaly fa z jiných mìsícù, pokud by se nìjak dostaly do øady
        + ' AND VATDate$DATE <= ' + FloatToStr(Trunc(EndOfAMonth(aseRok.Value, aseMesic.Value))); // aby se odfiltrovaly fa z jiných mìsícù, pokud by se nìjak dostaly do øady

        if cbBezVoIP.Checked and not cbSVoIP.Checked then
          SQLStr := SQLStr + ' AND LOWER(II.Description) NOT LIKE ''%voip%''';
        if cbSVoIP.Checked and not cbBezVoIP.Checked then
          SQLStr := SQLStr + ' AND LOWER(II.Description) LIKE ''%voip%''';

        if rbMail.Checked then SQLStr := SQLStr + ' AND F.Firm_ID IS NULL';  // TODO provìøit

        SQL.Text := SQLStr;
        Open;
        Radek := 0;
        apbProgress.Position := 0;
        apbProgress.Visible := True;
        while not EOF do begin
          VarSymbol := FieldByName('VarSymbol').AsString;
          Zakaznik := FieldByName('Name').AsString;
          apbProgress.Position := Round(100 * RecNo / RecordCount);
          Application.ProcessMessages;
          if Prerusit then begin
            Prerusit := False;
            apbProgress.Position := 0;
            apbProgress.Visible := False;
            btVytvorit.Enabled := True;
            btKonec.Caption := '&Konec';
            Break;
          end;
          with DesU.qrZakos do begin
            Close;

            SQLStr := 'SELECT DISTINCT Postal_mail AS Mail, Disable_mailings AS Reklama FROM customers Cu'
            + ' WHERE Variable_symbol = ' + VarSymbol;

            if rbMail.Checked then SQLStr := SQLStr + ' AND Invoice_sending_method_id = 9'; // hodnota v codebooks "Mailem"
            if rbTisk.Checked then begin
              if rbBezSlozenky.Checked then SQLStr := SQLStr + ' AND Invoice_sending_method_id = 10' // hodnota v codebooks "Poštou"
              else if rbSeSlozenkou.Checked then SQLStr := SQLStr + ' AND Invoice_sending_method_id = 11' // hodnota v codebooks "Se složenkou"
              else if rbKuryr.Checked then SQLStr := SQLStr + ' AND Invoice_sending_method_id = 12'; // hodnota v codebooks "Kurýr"
            end;

            SQL.Text := SQLStr;
            Open;
            while not EOF do begin
              Inc(Radek);
              RowCount := Radek + 1;
              AddCheckBox(0, Radek, True, True);
              Ints[0, Radek] := 1;                                                   // fajfka tisk - mail
              Cells[1, Radek] := VarSymbol;                                          // smlouva
              Cells[2, Radek] := Format('%5.5d', [DesU.qrAbra.FieldByName('OrdNumber').AsInteger]);     // faktura
              Floats[3, Radek] := DesU.qrAbra.FieldByName('Amount').AsFloat;;             // èástka
              Cells[4, Radek] := Zakaznik;                                           // jméno
              Cells[5, Radek] := FieldByName('Mail').AsString;                       // mail
              Ints[6, Radek] := FieldByName('Reklama').AsInteger;                    // reklama
              Cells[7, Radek] := DesU.qrAbra.FieldByName('ID').AsString;             // ID faktury
              Cells[8, Radek] := DateToStr(DesU.qrAbra.FieldByName('VATDate$DATE').AsFloat);
              Cells[9, Radek] := DateToStr(DesU.qrAbra.FieldByName('DocDate$DATE').AsFloat);

              Next;
            end;  // while not EOF
          end; //with DesU.qrZakos
          Application.ProcessMessages;
          Next;
        end;  // while not EOF
        if isDebugMode then dmCommon.Zprava('Vyhledány fakturované smlouvy');
      end;  // if rbVyberPodleFaktury.Checked with DesU.qrAbra

      if not rbFakturace.Checked then begin // TODO zeptat se, co a proc se sortuje, proc fakturace ne a ostatni ano
        SortSettings.Column := 2;
        SortSettings.Full := True;
        //SortSettings.AutoFormat := False;
        SortSettings.Direction := sdAscending;
        QSort;
      end;

      dmCommon.Zprava(Format('Poèet faktur: %d', [RowCount-1]));
      if rbFakturace.Checked then btVytvorit.Caption := '&Vytvoøit'
      else if rbPrevod.Checked then btVytvorit.Caption := '&Pøevést'
      else if rbTisk.Checked then btVytvorit.Caption := '&Vytisknout'
      else if rbMail.Checked then btVytvorit.Caption := '&Odeslat';
    except on E: Exception do
      Zprava('Neošetøená chyba: ' + E.Message);
    end;  // with DesU.qrZakos
  finally
    DesU.qrZakos.Close;
    apbProgress.Position := 0;
    apbProgress.Visible := False;
    if rbPrevod.Checked then apnPrevod.Visible := True;
    if rbTisk.Checked then apnTisk.Visible := True;
    if rbMail.Checked then apnMail.Visible := True;
    Screen.Cursor := crDefault;
    btVytvorit.Enabled := True;
    btVytvorit.SetFocus;
  end;
end;

end.
