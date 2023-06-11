unit FItisk;

interface

uses
  Windows, Messages, Forms, Controls, Classes, Dialogs, SysUtils, DateUtils, Printers, FImain;

type
  TdmTisk = class(TDataModule)
    dlgTisk: TPrintDialog;
  private
    procedure FakturaTisk(Radek: integer);                // vytiskne jednu fakturu
  public
    procedure TiskniFaktury;
  end;

var
  dmTisk: TdmTisk;

implementation

{$R *.dfm}

uses DesUtils, AArray, FIcommon;

// ------------------------------------------------------------------------------------------------

procedure TdmTisk.TiskniFaktury;
// pou�ije data z asgMain
var
  Posilani: AnsiString;
  Radek: integer;
begin
  with fmMain do try

    if rbBezSlozenky.Checked then Posilani := 'bez slo�enky';
    if rbSeSlozenkou.Checked then Posilani := 'se slo�enkou';
    if rbKuryr.Checked then Posilani := 'rozn�en�ch kur�rem';

    if rbVyberPodleVS.Checked then fmMain.Zprava(Format('Tisk faktur %s od VS %s do %s', [Posilani, aedOd.Text, aedDo.Text]))
    else fmMain.Zprava(Format('Tisk faktur %s od ��sla %s do %s', [Posilani, aedOd.Text, aedDo.Text]));
    with asgMain do begin
      fmMain.Zprava(Format('Po�et faktur k tisku: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));
      if ColumnSum(0, 1, RowCount-1) >= 1 then
        if dlgTisk.Execute then
          //frxReport.PrintOptions.Printer := Printer.Printers[Printer.PrinterIndex];
      Screen.Cursor := crHourGlass;
      apnTisk.Visible := False;
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
        if Ints[0, Radek] = 1 then FakturaTisk(Radek);
      end;  // for
    end;
// konec hlavn� smy�ky
  finally
    Printer.PrinterIndex := -1;
    apbProgress.Position := 0;
    apbProgress.Visible := False;
    apnTisk.Visible := True;
    Screen.Cursor := crDefault;                                             // default
    fmMain.Zprava('Tisk faktur ukon�en');
  end;
end;

// ------------------------------------------------------------------------------------------------

procedure TdmTisk.FakturaTisk(Radek: integer);
var

  FullPdfFileName,
  PdfDirName,
  desFrxUtilsResult: string;
  Mesic, i: integer;
  reportData: TAArray;

begin
  desFrxUtilsResult := '';

  with fmMain do begin
    { TODO na novy Report
    desFrxUtilsResult := DesFrxU.fakturaNactiData(asgMain.Cells[7, Radek]);
    fmMain.Zprava(desFrxUtilsResult);

    // !!! zde zavol�n� tisk !!!
    if rbSeSlozenkou.Checked then
      desFrxUtilsResult := DesFrxU.fakturaTisk('FOseSlozenkou.fr3')
    else begin
      DesFrxU.reportData['sQrKodem'] := true;
      desFrxUtilsResult := DesFrxU.fakturaTisk('FOsPDP.fr3');
    end;


    fmMain.Zprava(Format('%s (%s): Faktura %s byla odesl�na na tisk�rnu.', [DesFrxU.reportData['OJmeno'], DesFrxU.reportData['VS'], DesFrxU.reportData['Cislo']]));
    fmMain.Zprava(desFrxUtilsResult);
    asgMain.Ints[0, Radek] := 0;

    if desFrxUtilsResult <> 'Tisk OK' then
      if Application.MessageBox(PChar('Chyba p�i tisku'), 'Pokra�ovat?',
         MB_YESNO + MB_ICONQUESTION) = IDNO then Prerusit := True;
   }
  end;

end;

end.

