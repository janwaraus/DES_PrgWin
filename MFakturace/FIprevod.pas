unit FIprevod;

interface

uses
  Windows, Classes, Forms, Controls, SysUtils, Variants, DateUtils, Registry, Printers, Dialogs;

type
  TdmPrevod = class(TDataModule)
  private
    procedure FakturaPrevod(Radek: integer);
  public
    procedure PrevedFaktury;
  end;


var
  dmPrevod: TdmPrevod;

implementation

{$R *.dfm}

uses DesUtils, DesFrxUtils, AArray, FImain, FIcommon;  //frxExportSynPDF;

// ------------------------------------------------------------------------------------------------

procedure TdmPrevod.PrevedFaktury;
var
  Radek,
  i: integer;
  Reg: TRegistry;
begin
  Screen.Cursor := crHourGlass;

  with fmMain do try

    if rbPodleSmlouvy.Checked then
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
          asgMain.Visible := True;
          lbxLog.Visible := False;
          Break;
        end;

        if Ints[0, Radek] = 1 then FakturaPrevod(Radek)  // pokud za�krtnuto, p�ev�d�me fa do PDF

      end; // konec hlavn� smy�ky
    end;  // with asgMain

  finally
    Printer.PrinterIndex := -1;  // default
    apbProgress.Position := 0;
    apbProgress.Visible := False;
    apnPrevod.Visible := True;
    asgMain.Visible := False;
    lbxLog.Visible := True;
    Screen.Cursor := crDefault;
    dmCommon.Zprava('P�evod faktur do PDF ukon�en');
  end;

end;

// ------------------------------------------------------------------------------------------------

procedure TdmPrevod.FakturaPrevod(Radek: integer);
// podle faktury v Ab�e a stavu pohled�vek vytvo�� formul�� v PDF
var

  FullPdfFileName,
  PdfDirName,
  //FStr,                              // prefix faktury
  desFrxUtilsResult: string;
  Celkem,
  Saldo,
  Zaplatit,
  Zaplaceno: double;
  Mesic, i: integer;
  datumDokladu : double;


begin
  desFrxUtilsResult := '';

  with fmMain do begin

    //desFrxUtilsResult := DesFrxU.fakturaNactiData(globalAA['abraIiDocQueue_Id'], Ints[2, Radek], aseRok.Value); //takhle to bylo
    desFrxUtilsResult := DesFrxU.fakturaNactiData(asgMain.Cells[7, Radek]);
    dmCommon.Zprava(desFrxUtilsResult);

    Mesic := MonthOf(DesFrxU.reportData['DatumPlneni']); //opravdu datum plneni, tedy VATDate$DATE

    // adres�� pro ukl�d�n� faktur v PDF nemus� existovat
    PdfDirName := Format('%s\%4d\%2.2d', [PDFDir, aseRok.Value, Mesic]);  // Mesic misto jednoducheho aseMesic.Value
    if not DirectoryExists(PdfDirName) then Forcedirectories(PdfDirName);

    // jm�no souboru s fakturou
    FullPdfFileName := PdfDirName + Format('\%s-%5.5d.pdf', [globalAA['invoiceDocQueueCode'], asgMain.Ints[2, Radek]]);

    // soubor u� existuje
    if FileExists(FullPdfFileName) AND cbNeprepisovat.Checked then begin
        dmCommon.Zprava(Format('%s (%s): Soubor %s u� existuje.', [asgMain.Cells[4, Radek], asgMain.Cells[1, Radek], FullPdfFileName]));
        Exit;
      end else
        DeleteFile(FullPdfFileName);


    // !!! zde zavol�n� vytvo�en� PDF !!!
    DesFrxU.reportData['sQrKodem'] := true;
    desFrxUtilsResult := DesFrxU.fakturaVytvorPfd(FullPdfFileName, 'FOsPDP.fr3');
    dmCommon.Zprava(desFrxUtilsResult);


    // hotovo
    if not FileExists(FullPdfFileName) then
      dmCommon.Zprava(Format('%s (%s): Nepoda�ilo se vytvo�it soubor %s.', [DesFrxU.reportData['OJmeno'], DesFrxU.reportData['VS'], FullPdfFileName]))
    else begin
      dmCommon.Zprava(Format('%s (%s): Vytvo�en soubor %s.', [DesFrxU.reportData['OJmeno'], DesFrxU.reportData['VS'], FullPdfFileName]));
      asgMain.Ints[0, Radek] := 0;
      asgMain.Row := Radek;
    end;
  end;  // with fmMain
end;


end.

