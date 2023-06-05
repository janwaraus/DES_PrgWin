unit AbraEntities;

interface

uses
  SysUtils, Variants, Classes, Controls, System.Generics.Collections,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, ZAbstractConnection, ZConnection;

type

  TDoklad = class
  public
    ID : string[10];
    docQueue_ID : string[10];
    firm_ID : string[10];
    firmName : string[100];
    datumDokladu  : double;
    datumSplatnosti  : double;
    documentType : string[2]; // 60 typ dokladu dobopis fa vydaných (DO),  61 je typ dokladu dobropis faktur pøijatých (DD), 03 je faktura vydaná, 04 je faktura pøijatá, 10 je ZL
    castka  : currency; // LocalAmount
    castkaZaplaceno  : currency;
    castkaDobropisovano  : currency;
    castkaNezaplaceno  : currency;
    cisloDokladu : string[20]; // složené ABRA "lidské" èíslo dokladu
    constructor create(Document_ID : string; Document_Type : string = '03');
  end;

  TAbraBankAccount = class
  public
    id : string[10];
    name : string[50];
    number : string[42];
    accountId : string[10];
    bankstatementDocqueueId : string[10];
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
    id : string[10];
    code : string[4];
    name : string[10];
    number : string[42];
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
    constructor Create(pCode : string);
  end;

  TAbraDrcArticle = class
  public
    id : string[10];
    code : string[20];
    name : string;
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
    VatIdentNumber : string;
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
    //VatIndex_Vyst21: TAbraVatIndex;
    //VatIndex_VystR21: TAbraVatIndex;
    //DrcArticle_21: TAbraDrcArticle;
    //DocQueue_FO1: TAbraDocQueue;
    constructor create;
    //procedure loadVatEntities;
    //procedure loadDocQueueByCode(pCode : string);
    function getDivisionId : string;
    function getCurrencyId : string;
    function getDocQueue(bvString : string) : TAbraDocQueue;
    function getVatIndex(bvString : string) : TAbraVatIndex;
    function getDrcArticle(bvString : string) : TAbraDrcArticle;
    function getBusOrder(bvString : string) : TAbraBusOrder;
    function getIncomeType(bvString : string) : TAbraIncomeType;
  end;

var
  AbraEnt : TAbraEnt;

implementation

uses
  DesUtils;

{** class TDoklad **}

constructor TDoklad.create(Document_ID : string; Document_Type : string = '03');
begin
  with DesU.qrAbraOC do begin
    // cteni z IssuedInvoices
    SQL.Text :=
        'SELECT ii.ID, ii.DOCQUEUE_ID, ii.DOCDATE$DATE, ii.FIRM_ID, ii.DESCRIPTION, D.DOCUMENTTYPE,'
      + ' D.Code || ''-'' || II.OrdNumber || ''/'' || substring(P.Code from 3 for 2) as CisloDokladu,'
      + ' ii.LOCALAMOUNT, ii.LOCALPAIDAMOUNT, firms.Name as FirmName, ';

    if Document_Type = '10' then  //10 mají zálohové listy (ZL)
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
      self.bankStatementDocqueueId := FieldByName('BANKSTATEMENT_ID').AsString;
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
// najdu externí poøadové èíslo (EXTERNALNUMBER) výpisu s nejvyšším èíslem (ORDNUMBER) pro daný bank. úèet a rok
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

    SQL.Text := 'SELECT ID, CODE, NAME, DATEFROM$DATE, DATETO$DATE'
              + ' FROM PERIODS'
              + ' WHERE CODE = ''' + pYear  + '''';

    Open;
    if not Eof then begin
      self.id := FieldByName('ID').AsString;
      self.code := FieldByName('CODE').AsString;
      self.name := FieldByName('NAME').AsString;
      self.dateFrom := FieldByName('DATEFROM$DATE').AsFloat;
      self.dateTo := FieldByName('DATETO$DATE').AsFloat;
    end;
    Close;
  end;
end;

constructor TAbraPeriod.create(pDate : double);
begin

  with DesU.qrAbraOC do begin

    SQL.Text := 'SELECT ID, CODE, NAME, DATEFROM$DATE, DATETO$DATE '
              + ' FROM PERIODS'
              + ' WHERE DATEFROM$DATE <= ' + FloatToStr(pDate)
              + ' AND DATETO$DATE > ' + FloatToStr(pDate);

    Open;
    if not Eof then begin
      self.id := FieldByName('ID').AsString;
      self.code := FieldByName('CODE').AsString;
      self.name := FieldByName('NAME').AsString;
      self.dateFrom := FieldByName('DATEFROM$DATE').AsFloat;
      self.dateTo := FieldByName('DATETO$DATE').AsFloat;
    end;
    Close;
  end;
end;


{** class TAbraDocqueue **}

constructor TAbraDocQueue.create(pGetBy, pValue : string);
begin
  with DesU.qrAbraOC do begin
    SQL.Text := 'SELECT Id, Code, DocumentType, Name FROM DocQueues'
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
    SQL.Text := 'SELECT Id, Code, Tariff, Vatrate_Id'
              + ' FROM VatIndexes'
              + ' WHERE Hidden = ''N'' AND Code = ''' + pCode  + '''';
    Open;
    if not Eof then begin
      self.ID := FieldByName('Id').AsString;
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
    SQL.Text := 'SELECT Id, Code, Name'
              + ' FROM DrcArticles'
              + ' WHERE Hidden = ''N'' AND Code = ''' + iCode  + '''';
    Open;
    if not Eof then begin
      self.id := FieldByName('Id').AsString;
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
    SQL.Text := 'SELECT Id, Code, DocumentType, Name FROM BusOrdres'
              + ' WHERE ' + pGetBy + ' = ''' + pValue + '''';
    Open;
    if not Eof then begin
      self.ID := FieldByName('ID').AsString;
      self.Code := FieldByName('Code').AsString;
      self.Name := FieldByName('Name').AsString;
      self.Parent_ID := FieldByName('Name').AsString;
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

{
procedure TAbraEnt.loadVatEntities;
begin
  VatIndex_Vyst21 := TAbraVatIndex.create('Výst21');
  VatIndex_VystR21 := TAbraVatIndex.create('VýstR21');
  DrcArticle_21 :=  TAbraDrcArticle.create('21');
end;

procedure TAbraEnt.loadDocQueueByCode(pCode : string);
begin
  if pCode = 'FO1' then
    DocQueue_FO1 := TAbraDocQueue.create(pCode,'Code')
  else if pCode = 'FO2' then
    DocQueue_FO2 := TAbraDocQueue.create(pCode,'Code')
  else if pCode = 'FO3' then
    DocQueue_FO3 := TAbraDocQueue.create(pCode,'Code')
  else if pCode = 'FO4' then
    DocQueue_FO4 := TAbraDocQueue.create(pCode,'Code');
end;
}

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
    Result := TAbraVatIndex.create(bvValuePart);
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
    Result := TAbraDrcArticle.create(bvValuePart);
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
