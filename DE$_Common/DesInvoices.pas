unit DesInvoices;

interface

uses
  SysUtils, Variants, Classes, Controls, StrUtils,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, ZAbstractConnection, ZConnection,
  AbraEntities, DesUtils, DesFastReports, AArray;

type

  TDesInvoice = class
  private
    function prepareFastReportData : string;

  public
    ID : string[10];

    OrdNumber : integer;
    VS : string;
    IsReverseChargeDeclared : boolean;

    DocDate,
    DueDate,
    VATDate : double;

    Castka  : currency; // LocalAmount
    CastkaZaplaceno  : currency; // se zapo��t�n�m zaplacen� dobropis�
    CastkaZaplacenoZakaznikem  : currency; // se zapo��t�n�m zaplacen� dobropis�
    CastkaDobropisovano  : currency;
    CastkaNezaplaceno  : currency;

    CisloDokladu,                  // slo�en� ABRA "lidsk�" ��slo dokladu jak je na faktu�e
    CisloDokladuZkracene : string; // slo�en� ABRA "lidsk�" ��slo dokladu
    DefaultPdfName : string;

    // vlastnosti z DocQueues
    DocQueueID : string[10];
    DocQueueCode : string[10];
    DocumentType : string[2]; // 60 typ dokladu dobopis fa vydan�ch (DO),  61 je typ dokladu dobropis faktur p�ijat�ch (DD), 03 je faktura vydan�, 04 je faktura p�ijat�, 10 je ZL
    Year : string; // rok, do kter�ho doklad spad�, nap�. "2023"

    Firm : TAbraFirm;
    ResidenceAddress : TAbraAddress;

    constructor create(DocumentID : string; DocumentType : string = '03');
    function createPdf(Fr3FileName : string; OverwriteExistingPdf : boolean) : TDesResult;
  end;

  TNewDesInvoiceAA = class
  public
    AA : TAArray;
    constructor create(DocDate : double; VarSymbol : string);
    function createNew0Row(RowText : string) : TAArray;
    function createNew1Row(RowText : string) : TAArray;
  end;


implementation

constructor TDesInvoice.create(DocumentID : string; DocumentType : string = '03');
begin
  with DesU.qrAbraOC do begin
    // cteni z IssuedInvoices
    SQL.Text :=
        'SELECT ii.Firm_ID, ii.DocQueue_ID, ii.OrdNumber, ii.VarSymbol,'
      + ' ii.IsReverseChargeDeclared, ii.DocDate$DATE, ii.DueDate$DATE, ii.VATDate$DATE,'
      + ' ii.LocalAmount, ii.LocalPaidAmount, ii.LocalCreditAmount, ii.LocalPaidCreditAmount,'
      + ' d.Code as DocQueueCode, d.DocumentType, p.Code as YearCode,'
      + ' f.Name as FirmName, f.Code as AbraCode, f.OrgIdentNumber, f.VATIdentNumber,'
      + ' f.ResidenceAddress_ID, a.Street, a.City, a.PostCode '
      + 'FROM ISSUEDINVOICES ii'
      + ' JOIN DocQueues d ON ii.DocQueue_ID = d.ID'
      + ' JOIN Periods p ON ii.Period_ID = p.ID'
      + ' JOIN Firms f ON ii.Firm_ID = f.ID'
      + ' JOIN Addresses a ON f.ResidenceAddress_ID = a.ID'
      + ' WHERE ii.ID = ''' + DocumentID + '''';
    DesUtils.appendToFile(globalAA['LogFileName'], SQL.Text);

    Open;

    if not Eof then begin
      self.ID := DocumentID;

      self.OrdNumber := FieldByName('OrdNumber').AsInteger;
      self.VS := FieldByName('VarSymbol').AsString;
      IsReverseChargeDeclared := FieldByName('IsReverseChargeDeclared').AsString = 'A';

      self.DocDate := FieldByName('DocDate$Date').asFloat;
      self.DueDate := FieldByName('DueDate$Date').asFloat;
      self.VATDate := FieldByName('VATDate$Date').asFloat;

      self.Castka := FieldByName('LocalAmount').asCurrency;
      self.CastkaZaplaceno := FieldByName('LocalPaidAmount').asCurrency
                                    - FieldByName('LocalPaidCreditAmount').asCurrency;
      self.CastkaZaplacenoZakaznikem := FieldByName('LocalPaidAmount').asCurrency;
      self.CastkaDobropisovano := FieldByName('LocalCreditAmount').asCurrency;
      self.CastkaNezaplaceno := self.Castka - self.CastkaZaplaceno - self.CastkaDobropisovano;

      self.DocQueueID := FieldByName('DocQueue_ID').asString;
      self.DocQueueCode := FieldByName('DocQueueCode').asString;
      self.DocumentType := FieldByName('DocumentType').asString;
      self.Year := FieldByName('YearCode').asString;

      self.CisloDokladu := Format('%s-%d/%s', [self.DocQueueCode, self.OrdNumber, self.Year]); // s leading nulami by bylo jako '%s-%5.5d/%s'
      self.CisloDokladuZkracene := Format('%s-%d/%s', [self.DocQueueCode, self.OrdNumber, RightStr(self.Year, 2)]);

      self.Firm := TAbraFirm.create(
        FieldByName('Firm_ID').asString,
        FieldByName('FirmName').asString,
        FieldByName('AbraCode').asString,
        FieldByName('OrgIdentNumber').asString,
        FieldByName('VATIdentNumber').asString
      );

      self.ResidenceAddress := TAbraAddress.create(
        FieldByName('ResidenceAddress_ID').asString,
        FieldByName('Street').asString,
        FieldByName('City').asString,
        FieldByName('PostCode').asString
      );

    end;
    Close;
  end;
end;


function TDesInvoice.prepareFastReportData : string;
var
  slozenkaCastka,
  slozenkaVS,
  slozenkaSS,
  invoiceDocQueueId,
  invoiceDocQueueCode,
  SQLStr: string;
  Celkem,
  Saldo,
  Zaplatit,
  Zaplaceno: double;
  i, invoiceOrdNumber, invoiceRok: integer;
  reportAry: TAArray;


begin
  reportAry := TAArray.create;
  reportAry['Title'] := 'Faktura za p�ipojen� k internetu';
  reportAry['Author'] := 'Dru�stvo Eurosignal';

  reportAry['AbraKod'] := self.Firm.AbraCode;

  reportAry['OJmeno'] := self.Firm.Name;
  reportAry['OUlice'] := self.ResidenceAddress.Street;
  reportAry['OObec'] := self.ResidenceAddress.PostCode + ' ' + self.ResidenceAddress.City;
  reportAry['OICO'] := self.Firm.OrgIdentNumber;
  reportAry['ODIC'] := self.Firm.VATIdentNumber;

  reportAry['ID'] := self.ID;

  reportAry['DatumDokladu'] := self.DocDate;
  reportAry['DatumSplatnosti'] := self.DueDate;
  reportAry['DatumPlneni'] := self.VATDate;

  reportAry['VS'] := self.VS;
  reportAry['Celkem'] := self.Castka;
  reportAry['Zaplaceno'] := self.CastkaZaplacenoZakaznikem;
  if self.IsReverseChargeDeclared then
    reportAry['DRCText'] := 'Podle �92a z�kona �. 235/2004 Sb. o DPH da� odvede z�kazn�k.  '
  else
    reportAry['DRCText'] := ' ';

  Saldo := 0;

  with DesU.qrAbra do begin

    // v�echny Firm_Id pro Abrak�d firmy
    SQL.Text := 'SELECT * FROM DE$_Code_To_Firm_Id (' + Ap + self.Firm.AbraCode + ApZ;
    Open;

    { TODO tyhle kontroly d�t jinam
    if Fields[0].AsString = 'MULTIPLE' then begin
      Result := Format('%s (%s): V�ce z�kazn�k� pro k�d %s.', [self.Firm.Name, self.VS, self.Firm.AbraCode]);
      Close;
      Exit;  // tohle ale nesta��, je t�eba vysko�it z cel�ho p�evodu faktury, nebo to ud�lat n�jak jinak spr�vn�
             // by se m�lo zkontrolovat p�i na��t�n� faktur a v�bec nepou�t�t do p�evodu
    end;
    }

    // a saldo pro v�echny Firm_Id (saldo je z�porn�, pokud z�kazn�k dlu��)
    while not EOF do
    begin
      DesU.qrAbra2.SQL.Text := 'SELECT SaldoPo + SaldoZLPo + Ucet325 FROM DE$_Firm_Totals (' + Ap + DesU.qrAbra.Fields[0].AsString + ApC + FloatToStr(self.DocDate) + ')';
      DesU.qrAbra2.Open;
      Saldo := Saldo + DesU.qrAbra2.Fields[0].AsFloat;
      DesU.qrAbra2.Close;
      Next;
    end;
  end;  // with DesU.qrAbra

  // pr�v� p�ev�d�n� faktura m��e b�t p�ed splatnost�
  //    if Date <= DatumSplatnosti then begin
  Saldo := Saldo + self.CastkaZaplacenoZakaznikem;          // Saldo je po splatnosti (SaldoPo), je-li faktura u� zaplacena, p�i�te se platba
  Zaplatit := self.Castka - Saldo;          // Celkem k �hrad� = Celkem za fakt. obdob� - Z�statek minul�ch obdob�(saldo)
  // anebo je po splatnosti
  {end else begin
    Zaplatit := -Saldo;
    Saldo := Saldo + Celkem;             // ��stka faktury se ode�te ze salda, aby tam nebyla dvakr�t
  end;  }

  if Zaplatit < 0 then Zaplatit := 0;

  reportAry['Saldo'] :=  Saldo;
  reportAry['ZaplatitCislo'] := Zaplatit;
  reportAry['Zaplatit'] := Format('%.2f K�', [Zaplatit]);

  // text na fakturu
  if Saldo > 0 then reportAry['Platek'] := 'p�eplatek'
  else if Saldo < 0 then reportAry['Platek'] := 'nedoplatek'
  else reportAry['Platek'] := ' ';

  reportAry['Vystaveni'] := FormatDateTime('dd.mm.yyyy', reportAry['DatumDokladu']);
  reportAry['Plneni'] := FormatDateTime('dd.mm.yyyy', reportAry['DatumPlneni']);
  reportAry['Splatnost'] := FormatDateTime('dd.mm.yyyy', reportAry['DatumSplatnosti']);

  reportAry['SS'] := Format('%6.6d%s', [self.OrdNumber, RightStr(self.Year, 2)]);
  reportAry['Cislo'] := self.CisloDokladu;

  reportAry['Resume'] := Format('��stku %.0f,- K� uhra�te, pros�m, do %s na ��et 2100098382/2010 s variabiln�m symbolem %s.',
                                  [Zaplatit, reportAry['Splatnost'], reportAry['VS']]);

  // na�ten� �daj� z tabulky Smlouvy
  with DesU.qrZakos do begin
    Close;
    SQLStr := 'SELECT Postal_name, Postal_street, Postal_PSC, Postal_city FROM customers'
    + ' WHERE Variable_symbol = ' + Ap + reportAry['VS'] + Ap;
    SQL.Text := SQLStr;
    Open;
    reportAry['PJmeno'] := FieldByName('Postal_name').AsString;
    reportAry['PUlice'] := FieldByName('Postal_street').AsString;
    reportAry['PObec'] := FieldByName('Postal_PSC').AsString + ' ' + FieldByName('Postal_city').AsString;
    Close;
  end;

  // zas�lac� adresa
  if (reportAry['PJmeno'] = '') or (reportAry['PObec'] = '') then begin
    reportAry['PJmeno'] := reportAry['OJmeno'];
    reportAry['PUlice'] := reportAry['OUlice'];
    reportAry['PObec'] := reportAry['OObec'];
  end;

  // text pro QR k�d
  reportAry['QRText'] := Format('SPD*1.0*ACC:CZ6020100000002100098382*AM:%d*CC:CZK*DT:%s*X-VS:%s*X-SS:%s*MSG:QR PLATBA EUROSIGNAL',[
    Round(reportAry['ZaplatitCislo']),
    FormatDateTime('yyyymmdd', reportAry['DatumSplatnosti']),
    reportAry['VS'],
    reportAry['SS']]);

  slozenkaCastka := Format('%6.0f', [Zaplatit]);
  //nahrad�me vlnovkou posledn� mezeru, tedy d�me vlnovku p�ed prvn� ��slici
  for i := 2 to 6 do
    if slozenkaCastka[i] <> ' ' then begin
      slozenkaCastka[i-1] := '~';
      Break;
    end;

  //pro star�� osmim�stn� ��sla smluv se p�idaj� dv� nuly na za��tek
  if Length(self.VS) = 8 then
    slozenkaVS := '00' + self.VS
  else
    slozenkaVS := self.VS;

  slozenkaSS := Format('%8.8d%2.2d', [invoiceOrdNumber, invoiceRok]);

  reportAry['C1'] := slozenkaCastka[1];
  reportAry['C2'] := slozenkaCastka[2];
  reportAry['C3'] := slozenkaCastka[3];
  reportAry['C4'] := slozenkaCastka[4];
  reportAry['C5'] := slozenkaCastka[5];
  reportAry['C6'] := slozenkaCastka[6];
  reportAry['V01'] := slozenkaVS[1];
  reportAry['V02'] := slozenkaVS[2];
  reportAry['V03'] := slozenkaVS[3];
  reportAry['V04'] := slozenkaVS[4];
  reportAry['V05'] := slozenkaVS[5];
  reportAry['V06'] := slozenkaVS[6];
  reportAry['V07'] := slozenkaVS[7];
  reportAry['V08'] := slozenkaVS[8];
  reportAry['V09'] := slozenkaVS[9];
  reportAry['V10'] := slozenkaVS[10];
  reportAry['S01'] := slozenkaSS[1];
  reportAry['S02'] := slozenkaSS[2];
  reportAry['S03'] := slozenkaSS[3];
  reportAry['S04'] := slozenkaSS[4];
  reportAry['S05'] := slozenkaSS[5];
  reportAry['S06'] := slozenkaSS[6];
  reportAry['S07'] := slozenkaSS[7];
  reportAry['S08'] := slozenkaSS[8];
  reportAry['S09'] := slozenkaSS[9];
  reportAry['S10'] := slozenkaSS[10];

  reportAry['VS2'] := reportAry['VS']; // asi by sta�il jeden VS, ale nevad�
  reportAry['SS2'] := slozenkaSS; // SS2 se li�� od SS tak, �e m� 8 m�st. SS m� 6 m�st (zleva jsou v�dy p�idan� nuly)

  reportAry['Castka'] := slozenkaCastka + ',-';

  DesFastReport.reportData := reportAry;
end;


function TDesInvoice.createPdf(Fr3FileName : string; OverwriteExistingPdf : boolean) : TDesResult;
var
    PdfFileName,
    ExportDirName,
    FullPdfFileName : string;
    FVysledek : TDesResult;
begin
  DesFastReport.init('invoice', Fr3FileName); // nastaven� typu reportu a fr3 souboru

  PdfFileName := Format('zyz14-%s-%5.5d.pdf', [self.DocQueueCode, self.OrdNumber]);
  ExportDirName := Format('%s\i%s\%s\', [globalAA['PDFDir'], self.Year, FormatDateTime('mm', self.VATDate)]);
  DesFastReport.setExportDirName(ExportDirName);
  FullPdfFileName := ExportDirName + PdfFileName;

  self.prepareFastReportData;
  DesFastReport.prepareInvoiceDataSets(self.ID);
  Result := DesFastReport.createPdf(FullPdfFileName, OverwriteExistingPdf);
end;


constructor TNewDesInvoiceAA.create(DocDate : double; VarSymbol : string);
begin
  AA['DocDate'] := DocDate;
  AA['VarSymbol'] := VarSymbol;
end;

function TNewDesInvoiceAA.createNew0Row(RowText : string) : TAArray;
var
  RowAA: TAArray;
begin
  RowAA := AA.addRow();
  RowAA['Rowtype'] := 0;
  RowAA['Text'] := RowText;
  RowAA['Division_ID'] := AbraEnt.getDivisionId;
  Result := RowAA;
end;


function TNewDesInvoiceAA.createNew1Row(RowText : string) : TAArray;
var
  RowAA: TAArray;
begin
  RowAA := AA.addRow();
  RowAA['Rowtype'] := 1;
  RowAA['Text'] := RowText;
  RowAA['Division_ID'] := AbraEnt.getDivisionId;
  RowAA['Vatrate_ID'] := AbraEnt.getVatIndex('Code=V�st21').VATRate_ID;
  RowAA['VatIndex_ID'] := AbraEnt.getVatIndex('Code=V�st21').ID;
  RowAA['Incometype_ID'] := AbraEnt.getBusOrder('Code=SL').ID; // slu�by
  Result := RowAA;
end;

end.

