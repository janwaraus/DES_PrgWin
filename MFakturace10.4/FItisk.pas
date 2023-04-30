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

uses DesUtils, DesFrxUtils, AArray, FIcommon;

// ------------------------------------------------------------------------------------------------

procedure TdmTisk.TiskniFaktury;
// použije data z asgMain
var
  Posilani: AnsiString;
  Radek: integer;
begin
  with fmMain do try

    if rbBezSlozenky.Checked then Posilani := 'bez složenky';
    if rbSeSlozenkou.Checked then Posilani := 'se složenkou';
    if rbKuryr.Checked then Posilani := 'roznášených kurýrem';

    if rbPodleSmlouvy.Checked then dmCommon.Zprava(Format('Tisk faktur %s od VS %s do %s', [Posilani, aedOd.Text, aedDo.Text]))
    else dmCommon.Zprava(Format('Tisk faktur %s od èísla %s do %s', [Posilani, aedOd.Text, aedDo.Text]));
    with asgMain do begin
      dmCommon.Zprava(Format('Poèet faktur k tisku: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));
      if ColumnSum(0, 1, RowCount-1) >= 1 then
        if dlgTisk.Execute then
          //frxReport.PrintOptions.Printer := Printer.Printers[Printer.PrinterIndex];
      Screen.Cursor := crHourGlass;
      apnTisk.Visible := False;
      apbProgress.Position := 0;
      apbProgress.Visible := True;
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
          asgMain.Visible := True;
          lbxLog.Visible := False;
          Break;
        end;
        if Ints[0, Radek] = 1 then FakturaTisk(Radek);
      end;  // for
    end;
// konec hlavní smyèky
  finally
    Printer.PrinterIndex := -1;
    apbProgress.Position := 0;
    apbProgress.Visible := False;
    apnTisk.Visible := True;
    asgMain.Visible := False;
    lbxLog.Visible := True;
    Screen.Cursor := crDefault;                                             // default
    dmCommon.Zprava('Tisk faktur ukonèen');
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

    desFrxUtilsResult := DesFrxU.fakturaNactiData(asgMain.Cells[7, Radek]);
    dmCommon.Zprava(desFrxUtilsResult);

    // !!! zde zavolání tisk !!!
    if rbSeSlozenkou.Checked then
      desFrxUtilsResult := DesFrxU.fakturaTisk('FOseSlozenkou.fr3')
    else begin
      DesFrxU.reportData['sQrKodem'] := true;
      desFrxUtilsResult := DesFrxU.fakturaTisk('FOsPDP.fr3');
    end;


    dmCommon.Zprava(Format('%s (%s): Faktura %s byla odeslána na tiskárnu.', [DesFrxU.reportData['OJmeno'], DesFrxU.reportData['VS'], DesFrxU.reportData['Cislo']]));
    dmCommon.Zprava(desFrxUtilsResult);
    asgMain.Ints[0, Radek] := 0;

    if desFrxUtilsResult <> 'Tisk OK' then
      if Application.MessageBox(PChar('Chyba pøi tisku'), 'Pokraèovat?',
         MB_YESNO + MB_ICONQUESTION) = IDNO then Prerusit := True;

  end;

end;

end.

