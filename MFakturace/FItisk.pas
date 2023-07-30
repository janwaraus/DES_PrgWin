unit FItisk;

interface

uses
  Windows, Messages, Forms, Controls, Classes, Dialogs, SysUtils, DateUtils, Printers,
  DesUtils;

type
  TdmTisk = class(TDataModule)
    dlgTisk: TPrintDialog;
  public
    function FakturaTisk(InvoiceId, Fr3FileName: string) : TDesResult; // vytiskne jednu fakturu
    procedure TiskniFaktury;
  end;

var
  dmTisk: TdmTisk;

implementation

{$R *.dfm}

uses DesInvoices, DesFastReports, AArray, FImain;

// ------------------------------------------------------------------------------------------------

procedure TdmTisk.TiskniFaktury;
// použije data z asgMain
var
  Posilani: AnsiString;
  Radek: integer;
  VysledekPrevedeni : TDesResult;
begin
  with fmMain do try

    if rbBezSlozenky.Checked then Posilani := 'bez složenky';
    if rbSeSlozenkou.Checked then Posilani := 'se složenkou';
    if rbKuryr.Checked then Posilani := 'roznášených kurýrem';

    if rbVyberPodleVS.Checked then
      fmMain.Zprava(Format('Tisk faktur %s od VS %s do %s', [Posilani, aedOd.Text, aedDo.Text]))
    else
      fmMain.Zprava(Format('Tisk faktur %s od èísla %s do %s', [Posilani, aedOd.Text, aedDo.Text]));

    with asgMain do begin
      fmMain.Zprava(Format('Poèet faktur k tisku: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));
      Screen.Cursor := crHourGlass;
      apnTisk.Visible := False;
      apbProgress.Position := 0;
      apbProgress.Visible := True;
      DesFastReport.resetPrintCount;

      // hlavní smyèka
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

          if rbSeSlozenkou.Checked then
            VysledekPrevedeni := fakturaTisk(asgMain.Cells[7, Radek], 'FOseSlozenkou-104.fr3')
          else
            VysledekPrevedeni := fakturaTisk(asgMain.Cells[7, Radek], 'FOsPDP-104.fr3');

          fmMain.Zprava(Format('%s (%s): %s', [asgMain.Cells[4, Radek], asgMain.Cells[1, Radek], VysledekPrevedeni.Messg]));
          if VysledekPrevedeni.isOk then begin
            asgMain.Ints[0, Radek] := 0;
            asgMain.Row := Radek;
          end
          else
          begin
            if Dialogs.MessageDlg( 'Chyba pøi tisku: '
              + VysledekPrevedeni.Messg + sLineBreak + 'Pokraèovat?',
              mtConfirmation, [mbYes, mbNo], 0 ) = mrNo then Prerusit := True;
          end;
        end;

      end;  // for
    end;

  finally
    apbProgress.Position := 0;
    apbProgress.Visible := False;
    apnTisk.Visible := True;
    Screen.Cursor := crDefault;                                             // default
    fmMain.Zprava('Tisk faktur ukonèen');
  end;
end;

// ------------------------------------------------------------------------------------------------

function TdmTisk.fakturaTisk(InvoiceId, Fr3FileName: string) : TDesResult;
var
  Faktura : TDesInvoice;

begin
  Faktura := TDesInvoice.create(InvoiceId);
  Result := Faktura.printByFr3(Fr3FileName);
end;

end.

