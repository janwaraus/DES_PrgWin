unit FIcommon;
// 27.1.17 vyèlenìna procedura AktualizaceView jen pro DES

interface

uses
  Windows, SysUtils, Classes, Forms, Controls, DateUtils, Math, Registry, AdvGrid;

type
  TdmCommon = class(TDataModule)
  public
    procedure AktualizaceView;
  end;

const

  MyAddress_Id: string[10] = '7000000101';
  MyUser_Id: string[10] = '2200000101';          // automatická fakturace
  MyAccount_Id: string[10] = '1400000101';       // Fio
  MyPayment_Id: string[10] = '1000000101';       // typ platby: na bankovní úèet


var
  dmCommon: TdmCommon;

implementation

{$R *.dfm}

uses FIMain, DesUtils, AbraEntities;




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





end.
