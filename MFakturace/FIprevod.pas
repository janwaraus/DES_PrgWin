unit FIprevod;
interface
uses
  Windows, Classes, Forms, Controls, SysUtils, Variants, DateUtils, Registry, Printers, Dialogs,
  DesUtils;

type
  TdmPrevod = class(TDataModule)
  public
    function fakturaPrevod(InvoiceId: string; OverwriteExistingPdf : boolean) : TDesResult;
    procedure PrevedFaktury;
  end;

var
  dmPrevod: TdmPrevod;

implementation

{$R *.dfm}

uses DesInvoices, DesFastReports, AArray, FImain, FIcommon;  //frxExportSynPDF, DesFrxUtils

// ------------------------------------------------------------------------------------------------

procedure TdmPrevod.PrevedFaktury;
var
  Radek,
  i: integer;
  Reg: TRegistry;
  VysledekPrevedeni : TDesResult;
begin
  Screen.Cursor := crHourGlass;

  with fmMain do try
    if rbVyberPodleVS.Checked then
      dmCommon.Zprava(Format('P�evod faktur do PDF od VS %s do %s', [aedOd.Text, aedDo.Text]))
      else dmCommon.Zprava(Format('P�evod faktur do PDF od ��sla %s do %s', [aedOd.Text, aedDo.Text]));
    with asgMain do begin
      dmCommon.Zprava(Format('Po�et faktur k p�evodu: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));
      Screen.Cursor := crHourGlass;
      apnPrevod.Visible := False;
      apbProgress.Position := 0;
      apbProgress.Visible := True;


      // hlavn� smy�ka
      for Radek := 1 to RowCount-1 do begin
        Row := Radek;
        apbProgress.Position := Round(100 * Radek / RowCount-1);
        Application.ProcessMessages;
        if Prerusit then begin
          Prerusit := False;
          apbProgress.Position := 0;
          apbProgress.Visible := False;
          btVytvorit.Enabled := True;
          Break;
        end;

        //if Ints[0, Radek] = 1 then FakturaPrevod_old(Radek)  // pokud za�krtnuto, p�ev�d�me fa do PDF
        if Ints[0, Radek] = 1 then begin  // pokud za�krtnuto, p�ev�d�me fa do PDF
          VysledekPrevedeni := fakturaPrevod(asgMain.Cells[7, Radek], not cbNeprepisovat.Checked);
          dmCommon.Zprava(Format('%s (%s): %s', [asgMain.Cells[4, Radek], asgMain.Cells[1, Radek], VysledekPrevedeni.Messg]));
          if VysledekPrevedeni.isOk then begin
            asgMain.Ints[0, Radek] := 0;
            asgMain.Row := Radek;
          end;
        end;


      end; // konec hlavn� smy�ky
    end;  // with asgMain
  finally
    Printer.PrinterIndex := -1;  // default
    apbProgress.Position := 0;
    apbProgress.Visible := False;
    apnPrevod.Visible := True;
    Screen.Cursor := crDefault;
    dmCommon.Zprava('P�evod faktur do PDF ukon�en');
  end;
end;
// ------------------------------------------------------------------------------------------------
function TdmPrevod.fakturaPrevod(InvoiceId: string; OverwriteExistingPdf : boolean) : TDesResult;
// podle faktury v Ab�e a stavu pohled�vek vytvo�� formul�� v PDF
var
  Faktura : TDesInvoice;

begin
  // DesFastReport.setInvoiceById(invoiceId);
  // Result := DesFastReport.createPdf(true);

  Faktura := TDesInvoice.create(InvoiceId);

  if Faktura.Firm.AbraCode = '' then begin
    Result := TDesResult.create('err', Format('Faktura %d: z�kazn�k nem� k�d Abry.', [Faktura.OrdNumber]));
    Exit;
  end;

  // faktura se vytvo�� v defaultn�m adres��i s defaultn�m jm�nem souboru.
  // TODO pokud by bylo pot�eba ten default zm�nit, lze to ud�lat p�edchoz�m p�i�azen�m jin�ch hodnot a hl�d�n�m, zda ji� maj� n�jakou hodnotu
  Result := Faktura.createPdf('FOsPDP.fr3', OverwriteExistingPdf);

end;

end.
