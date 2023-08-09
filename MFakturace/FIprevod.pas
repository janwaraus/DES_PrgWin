unit FIprevod;
interface
uses
  Windows, Classes, Forms, Controls, SysUtils, Variants, DateUtils, Registry, Printers, Dialogs,
  DesUtils;

type
  TdmPrevod = class(TDataModule)
  public
    function FakturaPrevod(InvoiceId: string; OverwriteExistingPdf : boolean) : TDesResult;
    procedure PrevedFaktury;
  end;

var
  dmPrevod: TdmPrevod;

implementation

{$R *.dfm}

uses DesInvoices, DesFastReports, AArray, FImain;  //frxExportSynPDF, DesFrxUtils

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
      fmMain.Zprava(Format('Pøevod faktur do PDF od VS %s do %s', [aedOd.Text, aedDo.Text]))
      else fmMain.Zprava(Format('Pøevod faktur do PDF od èísla %s do %s', [aedOd.Text, aedDo.Text]));
    with asgMain do begin
      fmMain.Zprava(Format('Poèet faktur k pøevodu: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));
      Screen.Cursor := crHourGlass;
      apnPrevod.Visible := False;
      apbProgress.Position := 0;
      apbProgress.Visible := True;

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

        if Ints[0, Radek] = 1 then begin  // pokud zaškrtnuto, pøevádíme fa do PDF
          VysledekPrevedeni := fakturaPrevod(asgMain.Cells[7, Radek], not cbNeprepisovat.Checked);
          fmMain.Zprava(Format('%s (%s): %s', [asgMain.Cells[4, Radek], asgMain.Cells[1, Radek], VysledekPrevedeni.Messg]));
          if VysledekPrevedeni.isOk then begin
            asgMain.Ints[0, Radek] := 0;
            asgMain.Row := Radek; // proè?
          end;
        end;


      end; // konec hlavní smyèky
    end;  // with asgMain
  finally
    apbProgress.Position := 0;
    apbProgress.Visible := False;
    apnPrevod.Visible := True;
    Screen.Cursor := crDefault;
    fmMain.Zprava('Pøevod faktur do PDF ukonèen');
  end;
end;
// ------------------------------------------------------------------------------------------------
function TdmPrevod.FakturaPrevod(InvoiceId: string; OverwriteExistingPdf : boolean) : TDesResult;
// podle faktury v Abøe a stavu pohledávek vytvoøí formuláø v PDF
var
  Faktura : TDesInvoice;

begin
  Faktura := TDesInvoice.create(InvoiceId);

  if Faktura.Firm.AbraCode = '' then begin  // HWTODO opravdu je potreba kontrolovat?
    Result := TDesResult.create('err', Format('Faktura %d: zákazník nemá kód Abry.', [Faktura.OrdNumber]));
    Exit;
  end;

  // faktura se vytvoøí v defaultním adresáøi s defaultním jménem souboru.
  // TODO pokud by bylo potøeba ten default zmìnit, lze to udìlat pøedchozím pøiøazením jiných hodnot a hlídáním, zda již mají nìjakou hodnotu
  Result := Faktura.createPdfByFr3('FOsPDP-104.fr3', OverwriteExistingPdf);

end;

end.
