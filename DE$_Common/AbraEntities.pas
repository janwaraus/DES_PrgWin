unit AbraEntities;

interface

uses
  SysUtils, Variants, Classes, Controls, StrUtils, System.Generics.Collections,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, ZAbstractConnection, ZConnection,
  DesUtils, AArray, Superobject;

type

  {* t¯Ìda zajiöùujÌcÌ z·pis BO (= ABRA entit) do ABRY *}
  TNewAbraBo = class
  public
    BoName : string;
    Item : TAArray;
    RowItem : TAArray;
    WriteResult : TDesResult;
    constructor create(BoName : string);
    procedure createNewRow;
    procedure addInvoiceParams(DocDate : double);
    procedure createNewInvoiceRow(RowType : integer; RowText : string; isVatRow : boolean = true);
    procedure writeToAbra;
    function getCreatedBoItem(itemName: string) : string;
  end;


  {* t¯Ìdy pro pr·ci s jiû existujÌcÌmi ABRA entitami *}
  TDoklad = class
  public
    ID : string[10];
    docQueue_ID : string[10];
    firm_ID : string[10];
    firmName : string[100];
    datumDokladu  : double;
    datumSplatnosti  : double;
    documentType : string[2]; // 60 typ dokladu dobopis fa vydan˝ch (DO),  61 je typ dokladu dobropis faktur p¯ijat˝ch (DD), 03 je faktura vydan·, 04 je faktura p¯ijat·, 10 je ZL
    castka  : currency; // LocalAmount
    castkaZaplaceno  : currency;
    castkaDobropisovano  : currency;
    castkaNezaplaceno  : currency;
    cisloDokladu : string[20]; // sloûenÈ ABRA "lidskÈ" ËÌslo dokladu
    constructor create(Document_ID : string; Document_Type : string = '03');
  end;

  TAbraBankAccount = class
  public
    ID : string[10];
    Name : string[50];
    Number : string[42];
    AccountId : string[10];
    BankstatementDocqueueId : string[10];
    constructor create();
  published
    procedure loadByNumber(baNumber : string);
    function getPoradoveCisloMaxVypisu(pYear : string) : integer;
    function getExtPoradoveCisloMaxVypisu(pYear : string) : integer;
    function getDatumMaxVypisu(pYear : string) : double;
    function getPosledniDatumVypisu(pYear : string) : double;
    function getPocetVypisu(pYear : string) : integer;
    function getZustatek(pDatum : double) : double;
  end;

  TAbraPeriod = class
  public
    ID : string[10];
    Code : string[4];
    Name : string[10];
    Number : string[42];
    dateFrom, dateTo : double;
    constructor create(pYear : string); overload;
    constructor create(pDate : double); overload;
  end;

  TAbraDocQueue = class
  public
    ID : string[10];
    Code : string[10];
    DocumentType : string[2];
    Name : string;
    constructor create(pGetBy, pValue : string);
  end;

  TAbraVatIndex = class
  public
    ID : string[10];
    Code : string[10];
    Tariff : integer;
    VATRate_ID : string[10];
    constructor create(pCode : string);
  end;

  TAbraDrcArticle = class
  public
    ID : string[10];
    Code : string[20];
    Name : string;
    constructor create(iCode : string);
  end;

  TAbraBusOrder = class
  public
    ID : string[10];
    Code : string[20];
    Name : string;
    Parent_ID : string;
    constructor create(pGetBy, pValue : string);
  end;

  TAbraBusTransaction = class
  public
    ID : string[10];
    Code : string[20];
    Name : string;
    Parent_ID : string;
    constructor create(pGetBy, pValue : string);
  end;

  TAbraIncomeType = class
  public
    ID : string[10];
    Code : string[20];
    Name : string;
    constructor create(pGetBy, pValue : string);
  end;


  TAbraFirm = class
  public
    ID : string[10];
    Name,
    AbraCode,
    OrgIdentNumber,
    VATIdentNumber : string;
    constructor create(ID, Name, AbraCode, OrgIdentNumber, VATIdentNumber : string);
  end;

  TAbraAddress = class
  public
    ID : string[10];
    Street,
    City,
    PostCode : string;
    constructor create(ID, Street, City, PostCode : string);
  end;

  TAbraEnt = class
  public
    OA: TDictionary<string, TObject>;
    bvByPart, bvValuePart: string;
    constructor create;
    function getDivisionId : string;
    function getCurrencyId : string;
    function getPeriod(bvString : string) : TAbraPeriod;
    function getDocQueue(bvString : string) : TAbraDocQueue;
    function getVatIndex(bvString : string) : TAbraVatIndex;
    function getDrcArticle(bvString : string) : TAbraDrcArticle;
    function getBusOrder(bvString : string) : TAbraBusOrder;
    function getBusTransaction(bvString : string) : TAbraBusTransaction;
    function getIncomeType(bvString : string) : TAbraIncomeType;
  end;

var
  AbraEnt : TAbraEnt;

implementation



{************************************************************************}
{********************     t¯Ìda TNewAbraBo      *************************}
{************************************************************************}

constructor TNewAbraBo.create(BoName : string);
begin
  self.BoName := LowerCase(BoName);
  Item := TAArray.Create;
end;

procedure TNewAbraBo.addInvoiceParams(DocDate : double);
begin
  Item['DocDate$DATE'] := DocDate;
  Item['Period_ID'] := AbraEnt.getPeriod('Code=' + FormatDateTime('yyyy', DocDate)).ID;
  Item['Address_ID'] := '7000000101'; // FA CONST
  Item['BankAccount_ID'] := '1400000101'; // Fio
  Item['ConstSymbol_ID'] := '0000308000';
  Item['TransportationType_ID'] := '1000000101'; // FA CONST
  Item['PaymentType_ID'] := '1000000101'; // typ platby: na bankovnÌ ˙Ëet

  //if DocumentType = '03' then begin
  if BoName = 'issuedinvoice' then begin
    Item['AccDate$DATE'] := DocDate;
    Item['PricesWithVAT'] := True;
    Item['VATFromAbovePrecision'] := 0; // 6 je nejvyööÌ p¯esnost, ABRA nabÌzÌ 0 - 6; v praxi jsou poloûky faktury i v DB stejnÏ na 2 desetinn· mÌsta, ale t¯eba je to takhle p¯esnÏjöÌ v˝poËet
    Item['TotalRounding'] := 259; // zaokrouhlenÌ na koruny dol˘, aù z·kaznÌky nedr·ûdÌme halÈ¯ov˝mi "p¯Ìplatky"
  end;
end;

procedure TNewAbraBo.createNewRow;
begin
  self.rowItem := Item.addRow();
end;

procedure TNewAbraBo.createNewInvoiceRow(RowType : integer; RowText : string; isVatRow : boolean = true);
begin
  createNewRow;

  rowItem['RowType'] := RowType;
  rowItem['Text'] := RowText;
  rowItem['Division_ID'] := AbraEnt.getDivisionId;

  if (RowType > 0) AND isVatRow then begin
    rowItem['VATRate_ID'] := AbraEnt.getVatIndex('Code=V˝st' + DesU.VAT_RATE).VATRate_ID;
    rowItem['VATIndex_ID'] := AbraEnt.getVatIndex('Code=V˝st' + DesU.VAT_RATE).ID;
    rowItem['IncomeType_ID'] := AbraEnt.getIncomeType('Code=SL').ID; // sluûby
  end;

end;

procedure TNewAbraBo.writeToAbra;
begin
  WriteResult := DesU.abraBoCreate(self.Item.AsJSon, self.BoName);
end;

function TNewAbraBo.getCreatedBoItem(itemName: string) : string;
begin
  Result := '';
  if WriteResult.isOk then
    Result := SO(WriteResult.Messg).S[LowerCase(itemName)];
end;



{************************************************************************}
{********************     t¯Ìdy ABRA entit     **************************}
{************************************************************************}


{** class TDoklad **}

constructor TDoklad.create(Document_ID : string; Document_Type : string = '03');
begin
  with DesU.qrAbraOC do begin
    // cteni z IssuedInvoices
    SQL.Text :=
        'SELECT ii.ID, ii.DOCQUEUE_ID, ii.DOCDATE$DATE, ii.FIRM_ID, ii.DESCRIPTION, D.DOCUMENTTYPE,'
      + ' D.Code || ''-'' || II.OrdNumber || ''/'' || substring(P.Code from 3 for 2) as CisloDokladu,'
      + ' ii.LOCALAMOUNT, ii.LOCALPAIDAMOUNT, firms.Name as FirmName, ';

    if Document_Type = '10' then  //10 majÌ z·lohovÈ listy (ZL)
        SQL.Text := SQL.Text
        + ' 0 as LOCALCREDITAMOUNT, 0 as LOCALPAIDCREDITAMOUNT'
        + ' FROM ISSUEDDINVOICES ii'
    else if Document_Type = '61' then
        SQL.Text := SQL.Text
        + ' 0 as LOCALCREDITAMOUNT, 0 as LOCALPAIDCREDITAMOUNT'
        + ' FROM RECEIVEDCREDITNOTES ii'
    else
        SQL.Text := SQL.Text
        + ' ii.LOCALCREDITAMOUNT, ii.LOCALPAIDCREDITAMOUNT'
        + ' FROM ISSUEDINVOICES ii';

    SQL.Text := SQL.Text
      + ' JOIN Firms ON ii.Firm_ID = Firms.ID'
      + ' JOIN DocQueues D ON ii.DocQueue_ID = D.ID'
      + ' JOIN Periods P ON ii.Period_ID = P.ID'
      + ' WHERE ii.ID = ''' + Document_ID + '''';

    Open;

    if not Eof then begin
      self.ID := FieldByName('ID').AsString;
      self.Firm_ID := FieldByName('Firm_ID').AsString;
      self.FirmName := FieldByName('FirmName').AsString;
      self.DatumDokladu := FieldByName('DocDate$Date').asFloat;
      self.Castka := FieldByName('LocalAmount').Ascurrency;
      self.CastkaZaplaceno := FieldByName('LocalPaidAmount').Ascurrency
                                    - FieldByName('LocalPaidCreditAmount').Ascurrency;
      self.CastkaDobropisovano := FieldByName('LocalCreditAmount').Ascurrency;
      self.CastkaNezaplaceno := self.Castka - self.CastkaZaplaceno - self.CastkaDobropisovano;
      self.CisloDokladu := FieldByName('CisloDokladu').AsString;
      self.DocumentType := FieldByName('DocumentType').AsString;
    end;
    Close;
  end;
end;


{** class TAbraBankAccount **}

constructor TAbraBankAccount.create();
begin
  //self.qrAbra := DesU.qrAbraOC;
end;

procedure TAbraBankAccount.loadByNumber(baNumber : string);
begin
  with DesU.qrAbraOC do begin

    SQL.Text := 'SELECT ID, NAME, BANKACCOUNT, ACCOUNT_ID, BANKSTATEMENT_ID '
              + 'FROM BANKACCOUNTS '
              + 'WHERE BANKACCOUNT like ''' + baNumber  + '%'' '
              + 'AND HIDDEN = ''N'' ';

    Open;
    if not Eof then begin
      self.id := FieldByName('ID').AsString;
      self.name := FieldByName('NAME').AsString;
      self.number := FieldByName('BANKACCOUNT').AsString;
      self.accountId := FieldByName('ACCOUNT_ID').AsString;
      self.bankStatementDocqueueId := FieldByName('BANKSTATEMENT_ID').AsString; // v BankStatement_ID je ID ¯ady doklad˘
    end;
    Close;
  end;
end;

function TAbraBankAccount.getPoradoveCisloMaxVypisu(pYear : string) : integer;
begin
  with DesU.qrAbraOC do begin
    SQL.Text := 'SELECT MAX(bs.OrdNumber) as MaxOrdNumber '
              + ' FROM BANKSTATEMENTS bs, PERIODS p '
              + 'WHERE bs.DOCQUEUE_ID = ''' + self.bankStatementDocqueueId  + ''''
              + ' AND bs.PERIOD_ID = p.ID'
              + ' AND p.CODE = ''' + pYear  + '''';
    Open;
    if not Eof then
      Result := FieldByName('MaxOrdNumber').AsInteger
    else
      Result := 0;
    Close;
  end;
end;

function TAbraBankAccount.getExtPoradoveCisloMaxVypisu(pYear : string) : integer;
// najdu externÌ po¯adovÈ ËÌslo (EXTERNALNUMBER) v˝pisu s nejvyööÌm ËÌslem (ORDNUMBER) pro dan˝ bank. ˙Ëet a rok
begin
  with DesU.qrAbraOC do begin
    SQL.Text := 'SELECT bs1.OrdNumber as MaxOrdNumber, bs1.EXTERNALNUMBER as MaxExtOrdNumber'
              + ' FROM BANKSTATEMENTS bs1'
              + ' WHERE bs1.DOCQUEUE_ID = ''' + self.bankStatementDocqueueId  + ''''
              + ' AND bs1.PERIOD_ID = (SELECT ID FROM PERIODS p WHERE p.CODE = ''' + pYear  + ''')'
              + ' AND bs1.OrdNumber = (SELECT max(bs2.ORDNUMBER) FROM BANKSTATEMENTS bs2 WHERE bs1.DOCQUEUE_ID = bs2.DOCQUEUE_ID and bs1.PERIOD_ID = bs2.PERIOD_ID)';
    Open;
    if not Eof then
      Result := FieldByName('MaxExtOrdNumber').AsInteger
    else
      Result := 0;
    Close;
  end;
end;

function TAbraBankAccount.getDatumMaxVypisu(pYear : string) : double;
begin
  with DesU.qrAbraOC do begin
    SQL.Text := 'SELECT MAX(bs.DOCDATE$DATE) as PosledniDatum'
              + ' FROM BANKSTATEMENTS bs, PERIODS p '
              + 'WHERE bs.DOCQUEUE_ID = ''' + self.bankStatementDocqueueId  + ''''
              + ' AND bs.PERIOD_ID = p.ID'
              + ' AND p.CODE = ''' + pYear  + '''';

    Open;
    if not Eof then
      Result := FieldByName('PosledniDatum').AsFloat
    else
      Result := 0;
    Close;
  end;
end;

function TAbraBankAccount.getPosledniDatumVypisu(pYear : string) : double;
begin
  with DesU.qrAbraOC do begin
    SQL.Text := 'SELECT bs1.DOCDATE$DATE'
              + ' FROM BANKSTATEMENTS bs1'
              + ' WHERE bs1.DOCQUEUE_ID = ''' + self.bankStatementDocqueueId  + ''''
              + ' AND bs1.PERIOD_ID = (SELECT ID FROM PERIODS p WHERE p.CODE = ''' + pYear  + ''')'
              + ' AND bs1.OrdNumber = (SELECT max(bs2.ORDNUMBER) FROM BANKSTATEMENTS bs2 WHERE bs1.DOCQUEUE_ID = bs2.DOCQUEUE_ID and bs1.PERIOD_ID = bs2.PERIOD_ID)';

    Open;
    if not Eof then
      Result := FieldByName('DOCDATE$DATE').AsFloat
    else
      Result := 0;
    Close;
  end;
end;

function TAbraBankAccount.getPocetVypisu(pYear : string) : integer;
begin
  with DesU.qrAbraOC do begin
    SQL.Text := 'SELECT count(*) as PocetVypisu '
              + ' FROM BANKSTATEMENTS bs, PERIODS p '
              + 'WHERE bs.DOCQUEUE_ID = ''' + self.bankStatementDocqueueId  + ''''
              + ' AND bs.PERIOD_ID = p.ID'
              + ' AND p.CODE = ''' + pYear  + '''';
    Open;
    if not Eof then
      Result := FieldByName('PocetVypisu').AsInteger
    else
      Result := 0;
    Close;
  end;
end;

function TAbraBankAccount.getZustatek(pDatum : double) : double;
begin
 Result := DesU.getZustatekByAccountId(self.accountId, pDatum);
end;


{** class TAbraPeriod **}

constructor TAbraPeriod.create(pYear : string);
begin

  with DesU.qrAbraOC do begin

    SQL.Text := 'SELECT ID, Code, Name, DateFrom$DATE, DateTo$DATE'
              + ' FROM PERIODS'
              + ' WHERE Code = ''' + pYear  + '''';

    Open;
    if not Eof then begin
      self.id := FieldByName('ID').AsString;
      self.code := FieldByName('Code').AsString;
      self.name := FieldByName('Name').AsString;
      self.dateFrom := FieldByName('DateFrom$DATE').AsFloat;
      self.dateTo := FieldByName('DateTo$DATE').AsFloat;
    end;
    Close;
  end;
end;

constructor TAbraPeriod.create(pDate : double);
begin

  with DesU.qrAbraOC do begin

    SQL.Text := 'SELECT ID, Code, Name, DateFrom$DATE, DateTo$DATE '
              + ' FROM PERIODS'
              + ' WHERE DateFrom$DATE <= ' + FloatToStr(pDate)
              + ' AND DateTo$DATE > ' + FloatToStr(pDate);

    Open;
    if not Eof then begin
      self.id := FieldByName('ID').AsString;
      self.code := FieldByName('Code').AsString;
      self.name := FieldByName('Name').AsString;
      self.dateFrom := FieldByName('DateFrom$DATE').AsFloat;
      self.dateTo := FieldByName('DateTo$DATE').AsFloat;
    end;
    Close;
  end;
end;


{** class TAbraDocqueue **}

constructor TAbraDocQueue.create(pGetBy, pValue : string);
begin
  with DesU.qrAbraOC do begin
    SQL.Text := 'SELECT ID, Code, DocumentType, Name FROM DocQueues'
              + ' WHERE Hidden = ''N'' AND ' + pGetBy + ' = ''' + pValue + '''';
    Open;
    if not Eof then begin
      self.id := FieldByName('ID').AsString;
      self.code := FieldByName('Code').AsString;
      self.documentType := FieldByName('DocumentType').AsString;
      self.name := FieldByName('Name').AsString;
    end;
    Close;
  end;
end;


{** class TAbraVatIndex **}

constructor TAbraVatIndex.Create(pCode : string);
begin
  with DesU.qrAbraOC do begin
    SQL.Text := 'SELECT ID, Code, Tariff, VATRate_ID'
              + ' FROM VatIndexes'
              + ' WHERE Hidden = ''N'' AND Code = ''' + pCode  + '''';
    Open;
    if not Eof then begin
      self.ID := FieldByName('ID').AsString;
      self.Code := FieldByName('Code').AsString;
      self.Tariff := FieldByName('Tariff').AsInteger;
      self.VATRate_ID := FieldByName('VATRate_ID').AsString;
    end;
    Close;
  end;
end;


{** class TAbraDrcArticle **}

constructor TAbraDrcArticle.create(iCode : string);
begin
  with DesU.qrAbraOC do begin
    SQL.Text := 'SELECT ID, Code, Name'
              + ' FROM DrcArticles'
              + ' WHERE Hidden = ''N'' AND Code = ''' + iCode  + '''';
    Open;
    if not Eof then begin
      self.id := FieldByName('ID').AsString;
      self.code := FieldByName('Code').AsString;
      self.name := FieldByName('Name').AsString;
    end;
    Close;
  end;
end;


{** class TAbraBusOrder **}

constructor TAbraBusOrder.create(pGetBy, pValue : string);
begin
  with DesU.qrAbraOC do begin
    SQL.Text := 'SELECT ID, Code, Name, Parent_ID FROM BusOrders'
              + ' WHERE Closed = ''N'' AND ' + pGetBy + ' = ''' + pValue + '''';
    Open;
    if not Eof then begin
      self.ID := FieldByName('ID').AsString;
      self.Code := FieldByName('Code').AsString;
      self.Name := FieldByName('Name').AsString;
      self.Parent_ID := FieldByName('Parent_ID').AsString;
    end;
    Close;
  end;
end;


{** class TAbraBusTransaction **}

constructor TAbraBusTransaction.create(pGetBy, pValue : string);
begin
  with DesU.qrAbraOC do begin
    SQL.Text := 'SELECT ID, Code, Name, Parent_ID FROM BusTransactions'
              + ' WHERE Closed = ''N'' AND ' + pGetBy + ' = ''' + pValue + '''';
    Open;
    if not Eof then begin
      self.ID := FieldByName('ID').AsString;
      self.Code := FieldByName('Code').AsString;
      self.Name := FieldByName('Name').AsString;
      self.Parent_ID := FieldByName('Parent_ID').AsString;
    end;
    Close;
  end;
end;

{** class TAbraIncomeType **}

constructor TAbraIncomeType.create(pGetBy, pValue : string);
begin
  with DesU.qrAbraOC do begin
    SQL.Text := 'SELECT Id, Code, Name FROM IncomeTypes'
              + ' WHERE Hidden = ''N'' AND ' + pGetBy + ' = ''' + pValue + '''';
    Open;
    if not Eof then begin
      self.ID := FieldByName('ID').AsString;
      self.Code := FieldByName('Code').AsString;
      self.Name := FieldByName('Name').AsString;
    end;
    Close;
  end;
end;


{** class TAbraFirm **}

constructor TAbraFirm.create(ID, Name, AbraCode, OrgIdentNumber, VATIdentNumber : string);
begin
  self.ID := ID;
  self.Name := Name;
  self.AbraCode := Trim(AbraCode);
  self.OrgIdentNumber := OrgIdentNumber;
  self.VATIdentNumber := VATIdentNumber;
end;


{** class TAbraAddress **}

constructor TAbraAddress.create(ID, Street, City, PostCode : string);
begin
  self.ID := ID;
  self.Street := Street;
  self.City := City;
  self.PostCode := PostCode;
end;


{** class TAbraEnt **}

constructor TAbraEnt.Create;
begin
  self.OA:= TDictionary<string, TObject>.Create;
end;


function TAbraEnt.getDivisionId : string;
begin
  Result := '1000000101';
end;

function TAbraEnt.getCurrencyId : string;
begin
  Result := '0000CZK000'; //pouze CZK
end;




function TAbraEnt.getPeriod(bvString : string) : TAbraPeriod;
var
  outAbraEntity : TObject;
begin
  SplitStringInTwo(bvString, '=', bvByPart, bvValuePart);
  if OA.TryGetValue('Period_' + bvString, outAbraEntity) then
    Result := TAbraPeriod(outAbraEntity)
  else begin
    Result := TAbraPeriod.create(bvValuePart); // hled·me vûdy podle Code
    OA.Add('Period_' + bvString, Result);
  end;
end;


function TAbraEnt.getDocQueue(bvString : string) : TAbraDocQueue;
var
  outAbraEntity : TObject;
begin
  SplitStringInTwo(bvString, '=', bvByPart, bvValuePart);
  if OA.TryGetValue('DocQueue_' + bvString, outAbraEntity) then
    Result := TAbraDocQueue(outAbraEntity)
  else begin
    Result := TAbraDocQueue.create(bvByPart, bvValuePart);
    OA.Add('DocQueue_' + bvString, Result);
  end;
end;

function TAbraEnt.getVatIndex(bvString : string) : TAbraVatIndex;
var
  outAbraEntity : TObject;
begin
  SplitStringInTwo(bvString, '=', bvByPart, bvValuePart);
  if OA.TryGetValue('VatIndex_' + bvString, outAbraEntity) then
    Result := TAbraVatIndex(outAbraEntity)
  else begin
    Result := TAbraVatIndex.create(bvValuePart); // hled·me vûdy podle Code
    OA.Add('VatIndex_' + bvString, Result);
  end;
end;

function TAbraEnt.getDrcArticle(bvString : string) : TAbraDrcArticle;
var
  outAbraEntity : TObject;
begin
  SplitStringInTwo(bvString, '=', bvByPart, bvValuePart);
  if OA.TryGetValue('DrcArticle_' + bvString, outAbraEntity) then
    Result := TAbraDrcArticle(outAbraEntity)
  else begin
    Result := TAbraDrcArticle.create(bvValuePart); // hled·me vûdy podle Code
    OA.Add('DrcArticle_' + bvString, Result);
  end;
end;

function TAbraEnt.getBusOrder(bvString : string) : TAbraBusOrder;
var
  outAbraEntity : TObject;
begin
  SplitStringInTwo(bvString, '=', bvByPart, bvValuePart);
  if OA.TryGetValue('BusOrder_' + bvString, outAbraEntity) then
    Result := TAbraBusOrder(outAbraEntity)
  else begin
    Result := TAbraBusOrder.create(bvByPart, bvValuePart);
    OA.Add('BusOrder_' + bvString, Result);
  end;
end;

function TAbraEnt.getBusTransaction(bvString : string) : TAbraBusTransaction;
var
  outAbraEntity : TObject;
begin
  SplitStringInTwo(bvString, '=', bvByPart, bvValuePart);
  if OA.TryGetValue('BusTransaction_' + bvString, outAbraEntity) then
    Result := TAbraBusTransaction(outAbraEntity)
  else begin
    Result := TAbraBusTransaction.create(bvByPart, bvValuePart);
    OA.Add('BusTransaction_' + bvString, Result);
  end;
end;

function TAbraEnt.getIncomeType(bvString : string) : TAbraIncomeType;
var
  outAbraEntity : TObject;
begin
  SplitStringInTwo(bvString, '=', bvByPart, bvValuePart);
  if OA.TryGetValue('IncomeType_' + bvString, outAbraEntity) then
    Result := TAbraIncomeType(outAbraEntity)
  else begin
    Result := TAbraIncomeType.create(bvByPart, bvValuePart);
    OA.Add('IncomeType_' + bvString, Result);
  end;
end;



end.
