unit DesFrxUtils;

interface

uses
  Winapi.Windows, Winapi.ShellApi, Winapi.Messages,
  System.SysUtils, System.Variants, System.Classes, System.RegularExpressions,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  StrUtils,  // IOUtils, IniFiles, ComObj, //, Grids, AdvObj, StdCtrls,
  Data.DB, ZAbstractRODataset, ZAbstractDataset, ZDataset,
  ZAbstractConnection, ZConnection, AArray,

  frxClass, frxDBSet, frxDesgn, pCore2D, pBarcode2D, pQRCode,

  IdBaseComponent, IdComponent, IdTCPConnection,
  IdTCPClient, IdSMTP, IdHTTP, IdMessage, IdMessageClient, IdText, IdMessageParts,
  IdAntiFreezeBase, IdAntiFreeze, IdIOHandler,
  IdIOHandlerSocket, IdSSLOpenSSL, IdExplicitTLSClientServerBase, IdSMTPBase, IdAttachmentFile

  ;



type
  TDesFrxU = class(TForm)
    qrAbraDPH: TZQuery;
    qrAbraRadky: TZQuery;
    frxReport: TfrxReport;
    fdsDPH: TfrxDBDataset;
    fdsRadky: TfrxDBDataset;
    QRCode: TBarcode2D_QRCode;
    idMessage: TIdMessage;
    idSMTP: TIdSMTP;


    procedure frxReportGetValue(const ParName: string; var ParValue: Variant);
    procedure frxReportBeginDoc(Sender: TObject);


  public
    //pdfFileName, fr3FileName : string;
    reportData: TAArray;

    function fakturaNactiData(II_Id:  string) : string;

    function fakturaVytvorPfd(pdfFileName, fr3FileName : string) : string;
    function fakturaTisk(fr3FileName : string) : string;
    function faktura(action, fr3FileName : string) : string;

    function vytvorPdf(fr3FileName : string) : string;
    function tisk(fr3FileName : string) : string;

    function posliPdfEmailem(pdfFile, emailAddrStr, emailPredmet, emailZprava, emailOdesilatel : string) : string;

  end;

var
  DesFrxU: TDesFrxU;

implementation

uses DesUtils, AbraEntities, frxExportSynPDF;

{$R *.dfm}

function TDesFrxU.fakturaVytvorPfd(pdfFileName, fr3FileName : string) : string;
begin
  reportData['pdfFileName'] := pdfFileName;
  Result := faktura('vytvoreniPdf', fr3FileName);
end;

function TDesFrxU.fakturaTisk(fr3FileName : string) : string;
begin
  Result := faktura('tisk', fr3FileName);
end;

function TDesFrxU.faktura(action, fr3FileName : string) : string;
var
  ASymbolWidth,
  ASymbolHeight,
  AWidth,
  AHeight: integer;
begin

  // øádky faktury
  with qrAbraRadky do begin
    Close;
    SQL.Text := 'SELECT Text, TAmountWithoutVAT AS BezDane, VATRate AS Sazba, TAmount - TAmountWithoutVAT AS DPH, TAmount AS SDani FROM IssuedInvoices2'
    + ' WHERE Parent_ID = ' + Ap + reportData['ID'] + Ap                                       // ID faktury
    + ' AND NOT (Text = ''Zaokrouhlení'' AND TAmount = 0)'
    + ' ORDER BY PosIndex';
    Open;
  end;

  // rekapitulace
  with qrAbraDPH do begin
    Close;
    SQL.Text := 'SELECT VATRate AS Sazba, SUM(TAmountWithoutVAT) AS BezDane, SUM(TAmount - TAmountWithoutVAT) AS DPH, SUM(TAmount) AS SDani FROM IssuedInvoices2'
    + ' WHERE Parent_ID = ' + Ap + reportData['ID'] + Ap
    + ' AND VATIndex_ID IS NOT NULL'
    + ' GROUP BY Sazba';
    Open;
  end;

  if action = 'vytvoreniPdf' then begin
    Result := vytvorPdf(fr3FileName);
  end;

  if action = 'tisk' then begin
    Result := tisk(fr3FileName);
  end;

  qrAbraRadky.Close;
  qrAbraDPH.Close;
end;

function TDesFrxU.vytvorPdf(fr3FileName : string) : string;
var
    frxSynPDFExport: TfrxSynPDFExport;
    i : integer;
begin
  frxReport.LoadFromFile(DesU.PROGRAM_PATH + fr3FileName);
  frxReport.PrepareReport;

  // vytvoøená faktura se zpracuje do vlastního formuláøe a pøevede se do PDF
  // uložení pomocí Synopse
  frxSynPDFExport := TfrxSynPDFExport.Create(nil);
  with frxSynPDFExport do try
    FileName := reportData['pdfFileName'];
    Title := reportData['Title'];
    Author := reportData['Author'];
    EmbeddedFonts := False;
    Compressed := True;
    OpenAfterExport := False;
    ShowDialog := False;
    ShowProgress := False;
    PDFA := True; // important
    frxReport.Export(frxSynPDFExport);
  finally
    Free;
  end;

  // èekání na soubor - max. 5s
  for i := 1 to 50 do begin
    if FileExists( reportData['pdfFileName'] ) then Break;
    Sleep(100);
  end;

  Result := 'Vytvoøena fa ' + reportData['pdfFileName'];

end;

function TDesFrxU.tisk(fr3FileName : string) : string;
begin
  try
    frxReport.LoadFromFile(DesU.PROGRAM_PATH + fr3FileName);
    frxReport.PrepareReport;
    //frxReport.PrintOptions.ShowDialog := true;
    frxReport.PrintOptions.ShowDialog := false;
    frxReport.Print;
  except on E: exception do
    begin
      Result := Format('%s (%s): Fakturu %s se nepodaøilo vytisknout.' + #13#10 + 'Chyba: %s',
       [reportData['OJmeno'], reportData['VS'], reportData['Cislo'], E.Message]);
      Exit;
    end;
  end;  // try

  Result := 'Tisk OK';
end;


// ------------------------------------------------------------------------------------------------

procedure TDesFrxU.frxReportGetValue(const ParName: string; var ParValue: Variant);
// dosadí se promìné do formuláøe
begin

if ParName = 'Value = 0' then Exit; //pro jistotu, ve fr3 souboru toto bylo v highlight.condition

  try
    ParValue := self.reportData[ParName];
  except
    ShowMessage('Lehlo to na ' + ParName); //nefunguje mi
    //on E: Exception do
    //  ShowMessage(ParName + ' Chyba frxReportGetValue: '#13#10 + e.Message);
  end;

end;

procedure TDesFrxU.frxReportBeginDoc(Sender: TObject);
var
  ASymbolWidth,
  ASymbolHeight,
  AWidth,
  AHeight: integer;
begin

  // QR kód
  // nejdøíve ovìøení, že je reportData['sQrKodem'] typu boolean
  if varIsType(reportData['sQrKodem'], varBoolean) AND reportData['sQrKodem'] then begin
    QRCode.Barcode := Format('SPD*1.0*ACC:CZ6020100000002100098382*AM:%d*CC:CZK*DT:%s*X-VS:%s*X-SS:%s*MSG:QR PLATBA EUROSIGNAL',
     [Round(reportData['ZaplatitCislo']), FormatDateTime('yyyymmdd', reportData['DatumSplatnosti']), reportData['VS'], reportData['SS']]);
    QRCode.DrawToSize(AWidth, AHeight, ASymbolWidth, ASymbolHeight);
    with TfrxPictureView(frxReport.FindObject('pQR')).Picture.Bitmap do begin
      Width := AWidth;
      Height := AHeight;
      QRCode.DrawTo(Canvas, 0, 0);
    end;
  end;

end;


//

function TDesFrxU.fakturaNactiData(II_Id:  string) : string;
var
  slozenkaCastka,
  slozenkaVS,
  slozenkaSS,
  FrfFileName,
  OutFileName,
  OutDir,
  iDocQueue_Id,
  SQLStr: string;
  Celkem,
  Saldo,
  Zaplatit,
  Zaplaceno: double;
  i, iOrdNumber, iRok: integer;


begin

  with DesU.qrAbra do begin

    // údaje z faktury do reportData[] asociativního pole
    Close;
    SQL.Text := 'SELECT F.Code as AbraKod, F.Name, A.Street, A.City, A.PostCode, F.OrgIdentNumber, F.VATIdentNumber,'
      +' II.ID, II.OrdNumber, II.DocQueue_ID, II.VarSymbol, II.IsReverseChargeDeclared,'
      + ' II.DocDate$DATE, II.DueDate$DATE, II.VATDate$DATE, II.LocalAmount, II.LocalPaidAmount,'
      + ' Periods.Code as Rok'
      + ' FROM Firms F, Addresses A, IssuedInvoices II, Periods'
      + ' WHERE F.ID = II.Firm_ID'
      + ' AND A.ID = F.ResidenceAddress_ID'
      + ' AND Periods.ID = II.Period_ID'
      + ' AND F.Hidden = ''N'''
      + ' AND II.ID = ' + Ap + II_Id + Ap ;

    Open;

    if RecordCount = 0 then begin
      Result := Format('Neexistuje faktura s ID %d', [II_Id]);
      Close;
      Exit;
    end;

    iOrdNumber := FieldByName('OrdNumber').AsInteger;
    iRok := FieldByName('Rok').AsInteger;
    iDocQueue_Id := FieldByName('DocQueue_ID').AsString;

    if Trim(FieldByName('AbraKod').AsString) = '' then begin
      Result := Format('Faktura %d: zákazník nemá kód Abry.', [iOrdNumber]);
      Close;
      Exit;
    end;


    reportData := TAArray.Create;
    reportData['Title'] := 'Faktura za pøipojení k internetu';
    reportData['Author'] := 'Družstvo Eurosignal';

    reportData['AbraKod'] := FieldByName('AbraKod').AsString;
    reportData['OJmeno'] := FieldByName('Name').AsString;
    reportData['OUlice'] := FieldByName('Street').AsString;
    reportData['OObec'] := FieldByName('PostCode').AsString + ' ' + FieldByName('City').AsString;
    reportData['OICO'] := FieldByName('OrgIdentNumber').AsString;
    reportData['ODIC'] := FieldByName('VATIdentNumber').AsString;
    reportData['ID'] := FieldByName('ID').AsString;
    reportData['DatumDokladu'] := FieldByName('DocDate$DATE').AsFloat;
    reportData['DatumPlneni'] := FieldByName('VATDate$DATE').AsFloat;
    reportData['DatumSplatnosti'] := FieldByName('DueDate$DATE').AsFloat;
    reportData['VS'] := FieldByName('VarSymbol').AsString;;
    reportData['Celkem'] := FieldByName('LocalAmount').AsFloat;
    reportData['Zaplaceno'] := FieldByName('LocalPaidAmount').AsFloat;
    if FieldByName('IsReverseChargeDeclared').AsString = 'A' then
      reportData['DRCText'] := 'Podle §92a zákona è. 235/2004 Sb. o DPH daò odvede zákazník.  '
    else
      reportData['DRCText'] := ' ';

    Close;

    // všechny Firm_Id pro Abrakód firmy
    SQL.Text := 'SELECT * FROM DE$_Code_To_Firm_Id (' + Ap + reportData['AbraKod'] + ApZ;
    Open;
    if Fields[0].AsString = 'MULTIPLE' then begin
      Result := Format('%s (%s): Více zákazníkù pro kód %s.', [reportData['OJmeno'], reportData['VS'], reportData['AbraKod']]);
      Close;
      Exit;
    end;

    Saldo := 0;
    // a saldo pro všechny Firm_Id (saldo je záporné, pokud zákazník dluží)
    while not EOF do with DesU.qrAbra2 do begin
      Close;
      SQL.Text := 'SELECT SaldoPo + SaldoZLPo + Ucet325 FROM DE$_Firm_Totals (' + Ap + DesU.qrAbra.Fields[0].AsString + ApC + FloatToStr(reportData['DatumDokladu']) + ')';
      Open;
      Saldo := Saldo + Fields[0].AsFloat;
      DesU.qrAbra.Next;
    end; // while not EOF do with DesU.qrAbra2
  end;  // with DesU.qrAbra

  // právì pøevádìná faktura mùže být pøed splatností
  //    if Date <= DatumSplatnosti then begin
  Saldo := Saldo + reportData['Zaplaceno'];          // Saldo je po splatnosti (SaldoPo), je-li faktura už zaplacena, pøiète se platba
  Zaplatit := reportData['Celkem'] - Saldo;          // Celkem k úhradì = Celkem za fakt. období - Zùstatek minulých období(saldo)
  // anebo je po splatnosti
  {end else begin
    Zaplatit := -Saldo;
    Saldo := Saldo + Celkem;             // èástka faktury se odeète ze salda, aby tam nebyla dvakrát
  end;  }

  if Zaplatit < 0 then Zaplatit := 0;

  reportData['Saldo'] :=  Saldo;
  reportData['ZaplatitCislo'] := Zaplatit;
  reportData['Zaplatit'] := Format('%.2f Kè', [Zaplatit]);

  // text na fakturu
  if Saldo > 0 then reportData['Platek'] := 'pøeplatek'
  else if Saldo < 0 then reportData['Platek'] := 'nedoplatek'
  else reportData['Platek'] := ' ';

  reportData['Vystaveni'] := FormatDateTime('dd.mm.yyyy', reportData['DatumDokladu']);
  reportData['Plneni'] := FormatDateTime('dd.mm.yyyy', reportData['DatumPlneni']);
  reportData['Splatnost'] := FormatDateTime('dd.mm.yyyy', reportData['DatumSplatnosti']);

  reportData['SS'] := Format('%6.6d%2.2d', [iOrdNumber, iRok - 2000]);
  reportData['Cislo'] := Format('%s-%5.5d/%d', [DesU.getAbraDocqueueCodeById(iDocQueue_Id), iOrdNumber, iRok]);

  reportData['Resume'] := Format('Èástku %.0f,- Kè uhraïte, prosím, do %s na úèet 2100098382/2010 s variabilním symbolem %s.',
                                  [Zaplatit, reportData['Splatnost'], reportData['VS']]);

  // naètení údajù z tabulky Smlouvy
  with DesU.qrZakos do begin
    Close;
    SQLStr := 'SELECT Postal_name, Postal_street, Postal_PSC, Postal_city FROM customers'
    + ' WHERE Variable_symbol = ' + Ap + reportData['VS'] + Ap;
    SQL.Text := SQLStr;
    Open;
    reportData['PJmeno'] := FieldByName('Postal_name').AsString;
    reportData['PUlice'] := FieldByName('Postal_street').AsString;
    reportData['PObec'] := FieldByName('Postal_PSC').AsString + ' ' + FieldByName('Postal_city').AsString;
    Close;
  end;

  // zasílací adresa
  if (reportData['PJmeno'] = '') or (reportData['PObec'] = '') then begin
    reportData['PJmeno'] := reportData['OJmeno'];
    reportData['PUlice'] := reportData['OUlice'];
    reportData['PObec'] := reportData['OObec'];
  end;


  slozenkaCastka := Format('%6.0f', [Zaplatit]);
  //nahradíme vlnovkou poslední mezeru, tedy dáme vlnovku pøed první èíslici
  for i := 2 to 6 do
    if slozenkaCastka[i] <> ' ' then begin
      slozenkaCastka[i-1] := '~';
      Break;
    end;

  //pro starší osmimístná èísla smluv se pøidají dvì nuly na zaèátek
  if Length(reportData['VS']) = 8 then
    slozenkaVS := '00' + reportData['VS']
  else
    slozenkaVS := reportData['VS'];


  slozenkaSS := Format('%8.8d%2.2d', [iOrdNumber, iRok]);

  reportData['C1'] := slozenkaCastka[1];
  reportData['C2'] := slozenkaCastka[2];
  reportData['C3'] := slozenkaCastka[3];
  reportData['C4'] := slozenkaCastka[4];
  reportData['C5'] := slozenkaCastka[5];
  reportData['C6'] := slozenkaCastka[6];
  reportData['V01'] := slozenkaVS[1];
  reportData['V02'] := slozenkaVS[2];
  reportData['V03'] := slozenkaVS[3];
  reportData['V04'] := slozenkaVS[4];
  reportData['V05'] := slozenkaVS[5];
  reportData['V06'] := slozenkaVS[6];
  reportData['V07'] := slozenkaVS[7];
  reportData['V08'] := slozenkaVS[8];
  reportData['V09'] := slozenkaVS[9];
  reportData['V10'] := slozenkaVS[10];
  reportData['S01'] := slozenkaSS[1];
  reportData['S02'] := slozenkaSS[2];
  reportData['S03'] := slozenkaSS[3];
  reportData['S04'] := slozenkaSS[4];
  reportData['S05'] := slozenkaSS[5];
  reportData['S06'] := slozenkaSS[6];
  reportData['S07'] := slozenkaSS[7];
  reportData['S08'] := slozenkaSS[8];
  reportData['S09'] := slozenkaSS[9];
  reportData['S10'] := slozenkaSS[10];

  reportData['VS2'] := reportData['VS']; // asi by staèil jeden VS, ale nevadí
  reportData['SS2'] := slozenkaSS; // SS2 se liší od SS tak, že má 8 míst. SS má 6 míst (zleva jsou vždy pøidané nuly)

  reportData['Castka'] := slozenkaCastka + ',-';


end;

function TDesFrxU.posliPdfEmailem(pdfFile, emailAddrStr, emailPredmet, emailZprava, emailOdesilatel : string) : string;
begin
  idSMTP.Host :=  DesU.getIniValue('Mail', 'SMTPServer');
  idSMTP.Username := DesU.getIniValue('Mail', 'SMTPLogin');
  idSMTP.Password := DesU.getIniValue('Mail', 'SMTPPW');

  emailAddrStr := StringReplace(emailAddrStr, ',', ';', [rfReplaceAll]);    // èárky za støedníky

  with idMessage do begin
    Clear;
    From.Address := emailOdesilatel;
    ReceiptRecipient.Text := emailOdesilatel;

    // více mailových adres oddìlených støedníky se rozdìlí
    while Pos(';', emailAddrStr) > 0 do begin
      Recipients.Add.Address := Trim(Copy(emailAddrStr, 1, Pos(';', emailAddrStr)-1));
      emailAddrStr := Copy(emailAddrStr, Pos(';', emailAddrStr)+1, Length(emailAddrStr));
    end;
    Recipients.Add.Address := Trim(emailAddrStr);

    Subject := emailPredmet;

    with TIdText.Create(idMessage.MessageParts, nil) do begin
      Body.Text := emailZprava;
      ContentType := 'text/plain';
      Charset := 'utf-8';
    end;

    with TIdAttachmentFile.Create(IdMessage.MessageParts, pdfFile) do begin
      ContentType := 'application/pdf';
      FileName := extractfilename (pdfFile);
    end;

    { zatim vyhodim TODO
    // pøidá se pøíloha, je-li vybrána a zákazníkovi se posílá reklama
    if (Ints[6, Radek] = 0) and (fePriloha.FileName <> '') then
    with TIdAttachmentFile.Create(IdMessage.MessageParts, PDFFile) do begin
      ContentType := ''; //co je priloha za typ?
      FileName := fePriloha.FileName;
    end;

    ContentType := 'multipart/mixed';
    }

    { uz bylo vyhozeny
    with idSMTP do begin
      Port := 25;
      if Username = '' then AuthenticationType := atNone
      else AuthenticationType := atLogin;
    end;
    }

    if not idSMTP.Connected then idSMTP.Connect;
    idSMTP.Send(idMessage);

  end;

end;

end.

