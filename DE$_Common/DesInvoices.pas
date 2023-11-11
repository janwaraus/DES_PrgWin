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
    isReverseChargeDeclared : boolean;

    DocDate,
    DueDate,
    VATDate : double;

    Castka  : currency; // LocalAmount
    CastkaZaplaceno  : currency; // se zapoËÌt·nÌm zaplacenÌ dobropis˘
    CastkaZaplacenoZakaznikem  : currency; // se zapoËÌt·nÌm zaplacenÌ dobropis˘
    CastkaDobropisovano  : currency;
    CastkaNezaplaceno  : currency;

    CisloDokladu,                  // sloûenÈ ABRA "lidskÈ" ËÌslo dokladu jak je na faktu¯e
    CisloDokladuZkracene : string; // sloûenÈ ABRA "lidskÈ" ËÌslo dokladu
    DefaultPdfName : string;

    // vlastnosti z DocQueues
    DocQueueID : string[10];
    DocQueueCode : string[10];
    DocumentType : string[2]; // 60 typ dokladu dobopis fa vydan˝ch (DO),  61 je typ dokladu dobropis faktur p¯ijat˝ch (DD), 03 je faktura vydan·, 04 je faktura p¯ijat·, 10 je ZL
    Year : string; // rok, do kterÈho doklad spad·, nap¯. "2023"

    Firm : TAbraFirm;
    ResidenceAddress : TAbraAddress;

    reportAry: TAArray;

    constructor create(DocumentID : string; DocumentType : string = '03');
    function getPdfFileName : string;
    function getExportSubDir : string;
    function getFullPdfFileName : string;


    function createPdfByFr3(Fr3FileName : string; OverwriteExistingPdf : boolean) : TDesResult;
    function printByFr3(Fr3FileName : string) : TDesResult;
  end;

  TNewDesInvoiceAA = class
  public
    AA : TAArray;
    constructor create(DocDate : double; VarSymbol : string; DocumentType : string = '03');
    function createNew0Row(RowText : string) : TAArray;
    function createNew1Row(RowText : string) : TAArray;
    function createNewRow_NoVat(RowType : integer; RowText : string) : TAArray;
  end;


implementation

constructor TDesInvoice.create(DocumentID : string; DocumentType : string = '03');
begin
  with DesU.qrAbraOC do begin

    if DocumentType = '10' then  //'10' majÌ z·lohovÈ listy (ZL)
      SQL.Text :=
          'SELECT ii.Firm_ID, ii.DocQueue_ID, ii.OrdNumber, ii.VarSymbol,'
        + ' ''N'' as IsReverseChargeDeclared, ii.DocDate$DATE, ii.DueDate$DATE, 0 as VATDate$DATE,'
        + ' ii.LocalAmount, ii.LocalPaidAmount, 0 as LocalCreditAmount, 0 as LocalPaidCreditAmount,'
        + ' d.Code as DocQueueCode, d.DocumentType, p.Code as YearCode,'
        + ' f.Name as FirmName, f.Code as AbraCode, f.OrgIdentNumber, f.VATIdentNumber,'
        + ' f.ResidenceAddress_ID, a.Street, a.City, a.PostCode '
        + 'FROM IssuedDInvoices ii'
        + ' JOIN DocQueues d ON ii.DocQueue_ID = d.ID'
        + ' JOIN Periods p ON ii.Period_ID = p.ID'
        + ' JOIN Firms f ON ii.Firm_ID = f.ID'
        + ' JOIN Addresses a ON f.ResidenceAddress_ID = a.ID'
        + ' WHERE ii.ID = ''' + DocumentID + ''''

    else // default '03', cteni z IssuedInvoices
      SQL.Text :=
          'SELECT ii.Firm_ID, ii.DocQueue_ID, ii.OrdNumber, ii.VarSymbol,'
        + ' ii.IsReverseChargeDeclared, ii.DocDate$DATE, ii.DueDate$DATE, ii.VATDate$DATE,'
        + ' ii.LocalAmount, ii.LocalPaidAmount, ii.LocalCreditAmount, ii.LocalPaidCreditAmount,'
        + ' d.Code as DocQueueCode, d.DocumentType, p.Code as YearCode,'
        + ' f.Name as FirmName, f.Code as AbraCode, f.OrgIdentNumber, f.VATIdentNumber,'
        + ' f.ResidenceAddress_ID, a.Street, a.City, a.PostCode '
        + 'FROM IssuedInvoices ii'
        + ' JOIN DocQueues d ON ii.DocQueue_ID = d.ID'
        + ' JOIN Periods p ON ii.Period_ID = p.ID'
        + ' JOIN Firms f ON ii.Firm_ID = f.ID'
        + ' JOIN Addresses a ON f.ResidenceAddress_ID = a.ID'
        + ' WHERE ii.ID = ''' + DocumentID + '''';

    Open;

    if not Eof then begin
      self.ID := DocumentID;

      self.OrdNumber := FieldByName('OrdNumber').AsInteger;
      self.VS := FieldByName('VarSymbol').AsString;
      self.isReverseChargeDeclared := FieldByName('IsReverseChargeDeclared').AsString = 'A';

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
  reportAry := TAArray.create;
end;


function TDesInvoice.getPdfFileName : string;
begin
  if self.DocQueueCode = 'FO1' then //FO1 majÌ pÏtimÌstnÈ ËÌslo dokladu, ostatnÌ Ëty¯mÌstnÈ
    Result := Format('%s-%5.5d.pdf', [self.DocQueueCode, self.OrdNumber]) // pro test 'new-%s-%5.5d.pdf'
  else
    Result := Format('%s-%4.4d.pdf', [self.DocQueueCode, self.OrdNumber]);
end;

function TDesInvoice.getExportSubDir : string;
begin
  Result := Format('%s\%s\', [self.Year, FormatDateTime('mm', self.DocDate)]);
end;

function TDesInvoice.getFullPdfFileName : string;
begin
  Result := DesU.PDF_PATH + self.getExportSubDir + self.getPdfFileName;
end;

function TDesInvoice.prepareFastReportData : string;
var
  slozenkaCastka,
  slozenkaVS,
  slozenkaSS: string;
  Celkem,
  Saldo,
  Zaplatit: double;
  i: integer;


begin

  if self.DocumentType = '10' then
    reportAry['Title'] := 'Z·lohov˝ list na p¯ipojenÌ k internetu'
  else
    reportAry['Title'] := 'Faktura za p¯ipojenÌ k internetu';

  reportAry['Author'] := 'Druûstvo Eurosignal';

  reportAry['AbraKod'] := self.Firm.AbraCode;

  reportAry['OJmeno'] := self.Firm.Name;
  reportAry['OUlice'] := self.ResidenceAddress.Street;
  reportAry['OObec'] := self.ResidenceAddress.PostCode + ' ' + self.ResidenceAddress.City;

  //reportAry['OICO'] := self.Firm.OrgIdentNumber;
  if Trim(self.Firm.OrgIdentNumber) <> '' then
    reportAry['OICO'] := 'I»: ' + Trim(self.Firm.OrgIdentNumber)
  else
    reportAry['OICO'] := ' ';

  //reportAry['ODIC'] := self.Firm.VATIdentNumber;
  if Trim(self.Firm.VATIdentNumber) <> '' then
    reportAry['ODIC'] := 'DI»: ' + Trim(self.Firm.VATIdentNumber)
  else
    reportAry['ODIC'] := ' ';

  reportAry['ID'] := self.ID;

  reportAry['DatumDokladu'] := self.DocDate;
  reportAry['DatumSplatnosti'] := self.DueDate;
  reportAry['DatumPlneni'] := self.VATDate;

  reportAry['VS'] := self.VS;
  reportAry['Celkem'] := self.Castka;
  reportAry['Zaplaceno'] := self.CastkaZaplacenoZakaznikem;
  if self.IsReverseChargeDeclared then
    reportAry['DRCText'] := 'Podle ß92a z·kona Ë. 235/2004 Sb. o DPH daÚ odvede z·kaznÌk. '
  else
    reportAry['DRCText'] := ' ';

  Saldo := 0;

  with DesU.qrAbraOC do begin

    // vöechny Firm_Id pro AbrakÛd firmy
    SQL.Text := 'SELECT * FROM DE$_Code_To_Firm_Id (' + Ap + self.Firm.AbraCode + ApZ;
    Open;

    // a saldo pro vöechny Firm_Id (saldo je z·pornÈ, pokud z·kaznÌk dluûÌ)
    while not EOF do
    begin

      if self.DocQueueCode = 'FO3' then begin
        DesU.qrAbraOC2.SQL.Text := 'SELECT SaldoPo + SaldoZL + Ucet325 FROM DE$_Firm_Totals ('
            + Ap + DesU.qrAbraOC.Fields[0].AsString + ApC + FloatToStr(Date) + ')'; //pro FO3
      end else begin
        DesU.qrAbraOC2.SQL.Text := 'SELECT SaldoPo + SaldoZLPo + Ucet325 FROM DE$_Firm_Totals ('
            + Ap + DesU.qrAbraOC.Fields[0].AsString + ApC + FloatToStr(self.DocDate) + ')'; //pro FO1
      end;

      DesU.qrAbraOC2.Open;
      Saldo := Saldo + DesU.qrAbraOC2.Fields[0].AsFloat;
      DesU.qrAbraOC2.Close;
      Next;
    end;
  end;  // with DesU.qrAbraOC


  if self.DocQueueCode = 'FO3' then begin
    Zaplatit := -Saldo; // FO3
  end else begin
    Saldo := Saldo + self.CastkaZaplacenoZakaznikem;          //FO1 - Saldo je po splatnosti (SaldoPo), je-li faktura uû zaplacena, p¯iËte se platba
    Zaplatit := -Saldo + self.Castka;          // FO1 - Celkem k ˙hradÏ = Celkem za fakt. obdobÌ - Z˘statek minul˝ch obdobÌ(saldo)
  end;


  { //pr·vÏ p¯ev·dÏn· faktura m˘ûe b˝t p¯ed splatnostÌ *HWTODO probrat a vyhodit, faktura je vûdy p¯ed splatnostÌ
   if Date <= DatumSplatnosti then begin
    Saldo := Saldo + self.CastkaZaplacenoZakaznikem;          //FO1 - Saldo je po splatnosti (SaldoPo), je-li faktura uû zaplacena, p¯iËte se platba
    Zaplatit := self.Castka - Saldo;          // FO1 - Celkem k ˙hradÏ = Celkem za fakt. obdobÌ - Z˘statek minul˝ch obdobÌ(saldo)

   //anebo je po splatnosti
  end else begin
    Zaplatit := -Saldo;
    Saldo := Saldo + Celkem;             // Ë·stka faktury se odeËte ze salda, aby tam nebyla dvakr·t
  end;
  }

  if Zaplatit < 0 then Zaplatit := 0;

  reportAry['Saldo'] := Saldo;
  reportAry['ZaplatitCislo'] := Zaplatit;
  //reportAry['Zaplatit'] := Format('%.2f KË', [Zaplatit]);
  reportAry['Zaplatit'] := Zaplatit;

  // text na fakturu
  if Saldo > 0 then reportAry['Platek'] := 'P¯eplatek'
  else if Saldo < 0 then reportAry['Platek'] := 'Nedoplatek'
  else reportAry['Platek'] := ' ';

  reportAry['Vystaveni'] := FormatDateTime('dd.mm.yyyy', reportAry['DatumDokladu']);
  reportAry['Plneni'] := FormatDateTime('dd.mm.yyyy', reportAry['DatumPlneni']);
  reportAry['Splatnost'] := FormatDateTime('dd.mm.yyyy', reportAry['DatumSplatnosti']);

  reportAry['SS'] := Format('%6.6d%s', [self.OrdNumber, RightStr(self.Year, 2)]);
  reportAry['Cislo'] := self.CisloDokladu;

  reportAry['Resume'] := Format('»·stku %.0f,- KË uhraÔte, prosÌm, do %s na ˙Ëet 2100098382/2010 s variabilnÌm symbolem %s.',
                                  [Zaplatit, reportAry['Splatnost'], reportAry['VS']]);

  // naËtenÌ ˙daj˘ z tabulky Smlouvy
  with DesU.qrZakos do begin
    Close;
    SQL.Text := 'SELECT Postal_name, Postal_street, Postal_PSC, Postal_city FROM customers'
              + ' WHERE Variable_symbol = ' + Ap + reportAry['VS'] + Ap;
    Open;
    reportAry['PJmeno'] := FieldByName('Postal_name').AsString;
    reportAry['PUlice'] := FieldByName('Postal_street').AsString;
    reportAry['PObec'] := FieldByName('Postal_PSC').AsString + ' ' + FieldByName('Postal_city').AsString;
    Close;
  end;

  // zasÌlacÌ adresa
  if (reportAry['PJmeno'] = '') or (reportAry['PObec'] = '') then begin
    reportAry['PJmeno'] := reportAry['OJmeno'];
    reportAry['PUlice'] := reportAry['OUlice'];
    reportAry['PObec'] := reportAry['OObec'];
  end;

  // text pro QR kÛd
  reportAry['QRText'] := Format('SPD*1.0*ACC:CZ6020100000002100098382*AM:%d*CC:CZK*RN:EUROSIGNAL*DT:%s*X-VS:%s*X-SS:%s*MSG:QR PLATBA INTERNET',[
    Round(reportAry['ZaplatitCislo']),
    FormatDateTime('yyyymmdd', reportAry['DatumSplatnosti']),
    reportAry['VS'],
    reportAry['SS']]);

  slozenkaCastka := Format('%6.0f', [Zaplatit]);
  //nahradÌme vlnovkou poslednÌ mezeru, tedy d·me vlnovku p¯ed prvnÌ ËÌslici
  for i := 2 to 6 do
    if slozenkaCastka[i] <> ' ' then begin
      slozenkaCastka[i-1] := '~';
      Break;
    end;

  slozenkaVS := self.VS.PadLeft(10, '0'); // p¯id·me nuly na zaË·tek pro VS s dÈlkou kratöÌ neû 10 znak˘

  slozenkaSS := Format('%8.8d%s', [self.OrdNumber, RightStr(self.Year, 2)]);

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

  reportAry['VS2'] := reportAry['VS']; // asi by staËil jeden VS, ale nevadÌ
  reportAry['SS2'] := reportAry['SS']; // SS2 se liöÌ od SS tak, ûe m· 8 mÌst. SS m· 6 mÌst (zleva jsou vûdy p¯idanÈ nuly)

  reportAry['Castka'] := slozenkaCastka + ',-';

  DesFastReport.reportData := reportAry;
end;


function TDesInvoice.createPdfByFr3(Fr3FileName : string; OverwriteExistingPdf : boolean) : TDesResult;
begin
  DesFastReport.init('invoice', Fr3FileName); // nastavenÌ typu reportu a fr3 souboru
  DesFastReport.setExportDirName(DesU.PDF_PATH + self.getExportSubDir);

  self.prepareFastReportData;
  DesFastReport.prepareInvoiceDataSets(self.ID, self.DocumentType);

  Result := DesFastReport.createPdf(self.getFullPdfFileName, OverwriteExistingPdf);
end;

function TDesInvoice.printByFr3(Fr3FileName : string) : TDesResult;
begin
  DesFastReport.init('invoice', Fr3FileName); // nastavenÌ typu reportu a fr3 souboru

  self.prepareFastReportData;
  DesFastReport.prepareInvoiceDataSets(self.ID, self.DocumentType);

  Result := DesFastReport.print();
  if Result.Code = 'ok' then
    Result.Messg := Format('Doklad %s byl odesl·n na tisk·rnu.', [self.CisloDokladu]);
end;


constructor TNewDesInvoiceAA.create(DocDate : double; VarSymbol : string; DocumentType : string = '03');
begin
  AA := TAArray.Create;
  AA['VarSymbol'] := VarSymbol;
  AA['DocDate$DATE'] := DocDate;
  AA['Period_ID'] := AbraEnt.getPeriod('Code=' + FormatDateTime('yyyy', DocDate)).ID;
  AA['Address_ID'] := '7000000101'; // FA CONST
  AA['BankAccount_ID'] := '1400000101'; // Fio
  AA['ConstSymbol_ID'] := '0000308000';
  AA['TransportationType_ID'] := '1000000101'; // FA CONST
  AA['PaymentType_ID'] := '1000000101'; // typ platby: na bankovnÌ ˙Ëet

  if DocumentType = '03' then begin
    AA['AccDate$DATE'] := DocDate;
    AA['PricesWithVAT'] := True;
    AA['VATFromAbovePrecision'] := 0; // 6 je nejvyööÌ p¯esnost, ABRA nabÌzÌ 0 - 6; v praxi jsou poloûky faktury i v DB stejnÏ na 2 desetinn· mÌsta, ale t¯eba je to takhle p¯esnÏjöÌ v˝poËet
    AA['TotalRounding'] := 259; // zaokrouhlenÌ na koruny dol˘, aù z·kaznÌky nedr·ûdÌme halÈ¯ov˝mi "p¯Ìplatky"
  end;
end;


function TNewDesInvoiceAA.createNew0Row(RowText : string) : TAArray;
var
  RowAA: TAArray;
begin
  RowAA := AA.addRow();
  RowAA['RowType'] := 0;
  RowAA['Text'] := RowText;
  RowAA['Division_ID'] := AbraEnt.getDivisionId;
  Result := RowAA;
end;


function TNewDesInvoiceAA.createNew1Row(RowText : string) : TAArray;
var
  RowAA: TAArray;
begin
  RowAA := AA.addRow();
  RowAA['RowType'] := 1;
  RowAA['Text'] := RowText;
  RowAA['Division_ID'] := AbraEnt.getDivisionId;
  RowAA['VATRate_ID'] := AbraEnt.getVatIndex('Code=V˝st' + DesU.VAT_RATE).VATRate_ID;
  RowAA['VATIndex_ID'] := AbraEnt.getVatIndex('Code=V˝st' + DesU.VAT_RATE).ID;
  RowAA['IncomeType_ID'] := AbraEnt.getIncomeType('Code=SL').ID; // sluûby
  Result := RowAA;
end;

function TNewDesInvoiceAA.createNewRow_NoVat(RowType : integer; RowText : string) : TAArray;
var
  RowAA: TAArray;
begin
  RowAA := AA.addRow();
  RowAA['RowType'] := RowType;
  RowAA['Text'] := RowText;
  RowAA['Division_ID'] := AbraEnt.getDivisionId;
  Result := RowAA;
end;

end.

