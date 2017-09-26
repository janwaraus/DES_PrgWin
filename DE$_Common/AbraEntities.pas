unit AbraEntities;

interface

uses
  SysUtils, Variants, Classes, Controls,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, ZAbstractConnection, ZConnection;  

type

  TDoklad = class
  private
    qrAbra: TZQuery;
  public
    ID : string[10];
    docQueue_ID : string[10];
    firm_ID : string[10];
    firmName : string[100];
    datumDokladu  : double;
    datumSplatnosti  : double;
    //AccDocQueue_ID : string[10];
    //FirmOffice_ID : string[10];
    //DocUUID : string[26];
    documentType : string[2];
    castka  : Currency;
    castkaZaplaceno  : Currency;
    castkaDobropisovano  : Currency;
    castkaNezaplaceno  : Currency;
    cisloDokladu : string[20];
  //published
    constructor create(qrAbra : TZQuery); overload;
    constructor create(Document_ID : string; Document_Type : string = '03'); overload;
  end;

  TAbraBankAccount = class
  private
    qrAbra: TZQuery;
  public
    id : string[10];
    name : string[50];
    number : string[42];
    accountId : string[10];
    bankstatementDocqueueId : string[10];
    constructor create();
  published
    procedure loadByNumber(baNumber : string);
    function getMaxPoradoveCisloVypisu(pYear : string) : integer;
    function getMaxExtPoradoveCisloVypisu(pYear : string) : integer;
    function getMaxDatumVypisu(pYear : string) : double;    
    function getPocetVypisu(pYear : string) : integer;
    function getZustatek(pDatum : double) : double;
  end;

  TAbraPeriod = class
  private
    qrAbra: TZQuery;
  public
    id : string[10];
    code : string[4];
    name : string[10];
    number : string[42];
    dateFrom, dateTo : double;
    constructor create(pYear : string); overload;
    constructor create(pDate : double); overload;
  end;

implementation

uses
  DesUtils;

{** class TDoklad **}

constructor TDoklad.create(qrAbra : TZQuery);
begin
 with qrAbra do begin //do qrAbra je na�ten� ��dek z DB
  self.ID := FieldByName('ID').AsString;
  self.Firm_ID := FieldByName('Firm_ID').AsString;
  self.FirmName := FieldByName('FirmName').AsString;
  self.DatumDokladu := FieldByName('DocDate$Date').asFloat;
  self.Castka := FieldByName('LocalAmount').AsCurrency;
  self.CastkaZaplaceno := FieldByName('LocalPaidAmount').AsCurrency
                                - FieldByName('LocalPaidCreditAmount').AsCurrency;
  self.CastkaDobropisovano := FieldByName('LocalCreditAmount').AsCurrency;
  self.CastkaNezaplaceno := self.Castka - self.CastkaZaplaceno - self.CastkaDobropisovano;
  self.CisloDokladu := FieldByName('CisloDokladu').AsString;
  self.DocumentType := FieldByName('DocumentType').AsString;
 end; 
end;

constructor TDoklad.create(Document_ID : string; Document_Type : string = '03');
begin
  self.qrAbra := DesU.qrAbra;
  with qrAbra do begin

    // cteni z IssuedInvoices
    SQL.Text :=
        'SELECT ii.ID, ii.DOCQUEUE_ID, ii.DOCDATE$DATE, ii.FIRM_ID, ii.DESCRIPTION, D.DOCUMENTTYPE,'
      + ' D.Code || ''-'' || II.OrdNumber || ''/'' || substring(P.Code from 3 for 2) as CisloDokladu,'
      + ' ii.LOCALAMOUNT, ii.LOCALPAIDAMOUNT, ii.LOCALCREDITAMOUNT, ii.LOCALPAIDCREDITAMOUNT,'
      + ' firms.Name as FirmName'
      + ' FROM ISSUEDINVOICES ii'

      + ' JOIN Firms ON ii.Firm_ID = Firms.ID'
      + ' JOIN DocQueues D ON ii.DocQueue_ID = D.ID'
      + ' JOIN Periods P ON ii.Period_ID = P.ID'
      + ' WHERE ii.ID = ''' + Document_ID + '''';

    Open;
    if not Eof then begin
      self.create(qrAbra);
    end;
    Close;
  end;
end;

{** class TAbraBankAccount **}

constructor TAbraBankAccount.create();
begin
  self.qrAbra := DesU.qrAbra;
end;

procedure TAbraBankAccount.loadByNumber(baNumber : string);
begin
  with DesU.qrAbra do begin

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

function TAbraBankAccount.getMaxPoradoveCisloVypisu(pYear : string) : integer;
begin
  with qrAbra do begin
    SQL.Text := 'SELECT MAX(bs.OrdNumber) as MaxOrdNumber '  //nem�lo by b�t max externalnumber?
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

function TAbraBankAccount.getMaxExtPoradoveCisloVypisu(pYear : string) : integer;
begin
  with qrAbra do begin
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

function TAbraBankAccount.getMaxDatumVypisu(pYear : string) : double;
begin
  with qrAbra do begin
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
  with qrAbra do begin
    SQL.Text := 'SELECT count(*) as PocetVypisu '
              + ' FROM BANKSTATEMENTS bs, PERIODS p '
              + 'WHERE bs.DOCQUEUE_ID = ''' + self.bankStatementDocqueueId  + ''''
              + ' AND bs.PERIOD_ID = p.ID'
              + ' AND p.CODE = ''' + pYear  + ''''
              + ' GROUP BY bs.DOCQUEUE_ID';
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

  with DesU.qrAbra do begin

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

  with DesU.qrAbra do begin

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

end.
