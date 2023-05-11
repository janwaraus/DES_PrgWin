unit DesInvoices;

interface

uses
  SysUtils, Variants, Classes, Controls, StrUtils,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, ZAbstractConnection, ZConnection,
  AbraEntities, DesUtils;

type

  TDesInvoice = class
  public
    ID : string[10];

    OrdNumber : integer;
    VS : string;
    IsReverseChargeDeclared : boolean;

    DocDate,
    DueDate,
    VATDate : double;
    //AccDocQueue_ID : string[10];
    //FirmOffice_ID : string[10];
    //DocUUID : string[26];

    Castka  : currency; // LocalAmount
    CastkaZaplaceno  : currency;
    CastkaDobropisovano  : currency;
    CastkaNezaplaceno  : currency;

    CisloDokladu,                  // složené ABRA "lidské" èíslo dokladu jak je na faktuøe
    CisloDokladuZkracene : string; // složené ABRA "lidské" èíslo dokladu
    DefaultPdfName : string;

    // vlastnosti z DocQueues
    DocQueueID : string[10];
    DocQueueCode : string[10];
    DocumentType : string[2]; // 60 typ dokladu dobopis fa vydaných (DO),  61 je typ dokladu dobropis faktur pøijatých (DD), 03 je faktura vydaná, 04 je faktura pøijatá, 10 je ZL
    Year : string; // rok, do kterého doklad spadá, napø. "2023"

    Firm : TAbraFirm;
    Address : TAbraAddress;



    constructor create(Document_ID : string; Document_Type : string = '03');
  end;



implementation

constructor TDesInvoice.create(Document_ID : string; Document_Type : string = '03');
begin
  with DesU.qrAbraOC do begin
    // cteni z IssuedInvoices
    SQL.Text :=
        'SELECT ii.Firm_ID, ii.DocQueue_ID, ii.OrdNumber, ii.VarSymbol,'
      + ' ii.IsReverseChargeDeclared, ii.DocDate$DATE, ii.DueDate$DATE, ii.VATDate$DATE,'
      + ' ii.LocalAmount, ii.LocalPaidAmount, ii.LocalCreditAmount, ii.LocalPaidCreditAmount,'
      + ' d.Code as DocQueueCode, d.DocumentType, p.Code as Year,'
      + ' f.Name as FirmName, f.Code as FirmAbraCode, f.OrgIdentNumber, f.VATIdentNumber'
      + ' f.ResidenceAddress_ID, a.Street, a.City, a.PostCode '
      + 'FROM ISSUEDINVOICES ii';
      + ' JOIN DocQueues d ON ii.DocQueue_ID = d.ID'
      + ' JOIN Periods p ON ii.Period_ID = p.ID'
      + ' JOIN Firms f ON ii.Firm_ID = f.ID'
      + ' JOIN Addresses a ON f.ResidenceAddress_ID = a.ID'
      + ' WHERE ii.ID = ''' + Document_ID + '''';

    Open;

    if not Eof then begin
      self.ID := Document_ID;

      self.OrdNumber := FieldByName('OrdNumber').AsInteger;
      self.VS := FieldByName('VarSymbol').AsString;
      IsReverseChargeDeclared := FieldByName('IsReverseChargeDeclared').AsString = 'A';

      self.DocDate := FieldByName('DocDate$Date').asFloat;
      self.DueDate := FieldByName('DueDate$Date').asFloat;
      self.VATDate := FieldByName('VATDate$Date').asFloat;

      self.Castka := FieldByName('LocalAmount').asCurrency;
      self.CastkaZaplaceno := FieldByName('LocalPaidAmount').asCurrency
                                    - FieldByName('LocalPaidCreditAmount').asCurrency;
      self.CastkaDobropisovano := FieldByName('LocalCreditAmount').asCurrency;
      self.CastkaNezaplaceno := self.Castka - self.CastkaZaplaceno - self.CastkaDobropisovano;

      self.DocQueueID := FieldByName('DocQueue_ID').asString;
      self.DocQueueCode := FieldByName('DocQueueCode').asString;
      self.DocumentType := FieldByName('DocumentType').asString;
      self.Year := FieldByName('Year').asString;

      self.CisloDokladu := Format('%s-%5.5d/%s.', [self.DocQueueCode, self.OrdNumber, self.Year]);
      self.CisloDokladuZkracene := Format('%s-%d/%s.', [self.DocQueueCode, self.OrdNumber, RightStr(self.Year, 2)]);
      // pdfFileName := Format('%s-%5.5d.pdf', [self.DocQueueCode, self.OrdNumber]); // nebudeme zde deklarovat

      self.Firm := TAbraFirm.create(
        FieldByName('Firm_ID').asString,
        FieldByName('FirmName').asString,
        FieldByName('FirmAbraCode').asString,
        FieldByName('OrgIdentNumber').asString,
        FieldByName('VATIdentNumber').asString
      );

      self.Address := TAbraAddress.create(
        FieldByName('ResidenceAddress_ID').asString,
        FieldByName('Street').asString,
        FieldByName('City').asString,
        FieldByName('PostCode').asString
      );

    end;
    Close;
  end;
end;

end.

