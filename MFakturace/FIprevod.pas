unit FIprevod;

interface

uses
  Windows, Classes, Forms, Controls, SysUtils, Variants, DateUtils, Registry, Printers, Dialogs;

type
  Ch32 = array [0..31] of char;
  TdmPrevod = class(TDataModule)
  private
{$IFNDEF ABAK}
    PWDbuf: Ch32;
    JePrint2PDF: boolean;
{$ENDIF}
    procedure FakturaPrevod(Radek: integer);
  public
    procedure PrevedFaktury;
  end;

{$IFNDEF ABAK}
const
  PWD: Ch32 =
    (char($44), char($F5), char($B3), char($27), char($D5), char($4D), char($01), char($1F),
     char($42), char($78), char($5E), char($DB), char($5B), char($4D), char($31), char($09),
     char($C6), char($DE), char($DA), char($F9), char($BD), char($69), char($CC), char($5A),
     char($64), char($4F), char($F6), char($12), char($A8), char($F8), char($3F), char($55));
{$ENDIF}

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

  OutFileName,
  OutDir,
  FStr,                              // prefix faktury
  desFrxUtilsResult: string;
  Celkem,
  Saldo,
  Zaplatit,
  Zaplaceno: double;
  Mesic, i: integer;
  datumDokladu : double;


begin
  desFrxUtilsResult := '';

  with fmMain, fmMain.asgMain do begin

    desFrxUtilsResult := DesFrxU.fakturaNactiData(globalAA['abraIiDocQueue_Id'], Ints[2, Radek], aseRok.Value);
    dmCommon.Zprava(desFrxUtilsResult);

    Mesic := MonthOf(DesFrxU.reportData['DatumPlneni']); //opravdu datum plneni, tedy VATDate$DATE
    // adres�� pro ukl�d�n� faktur v PDF nemus� existovat
    if not DirectoryExists(PDFDir) then CreateDir(PDFDir);
    OutDir := PDFDir + Format('\%4d', [aseRok.Value]);
    if not DirectoryExists(OutDir) then CreateDir(OutDir);
    OutDir := OutDir + Format('\%2.2d', [Mesic]);
    if not DirectoryExists(OutDir) then CreateDir(OutDir);

    // jm�no souboru s fakturou
    FStr := DesU.getAbraDocqueueCodeById(globalAA['abraIiDocQueue_Id']); //bylo natvrdo 'FO1'
    OutFileName := OutDir + Format('\%s-%5.5d.pdf', [FStr, Ints[2, Radek]]);
    // soubor u� existuje
    if FileExists(OutFileName) AND cbNeprepisovat.Checked then begin
        dmCommon.Zprava(Format('%s (%s): Soubor %s u� existuje.', [Cells[4, Radek], Cells[1, Radek], OutFileName]));
        Exit;
      end else
        DeleteFile(OutFileName);


    //DesFrxU.reportData['Cislo'] := Format('%s-%5.5d/%d', [FStr, Ints[2, Radek], aseRok.Value]); //TODO dat dovnitr - snad hotovo


    // !!! zde zavol�n� vytvo�en� PDF
    // !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    DesFrxU.reportData['sQrKodem'] := true;
    DesFrxU.fakturaVytvorPfd(OutFileName, 'FOsPDP.fr3');
    //DesFrxU.reportData['sQrKodem'] := false;
    //DesFrxU.fakturaTisk('FOseSlozenkou.fr3');

    // hotovo
    if not FileExists(OutFileName) then
      dmCommon.Zprava(Format('%s (%s): Nepoda�ilo se vytvo�it soubor %s.', [DesFrxU.reportData['OJmeno'], DesFrxU.reportData['VS'], OutFileName]))
    else begin
      dmCommon.Zprava(Format('%s (%s): Vytvo�en soubor %s.', [DesFrxU.reportData['OJmeno'], DesFrxU.reportData['VS'], OutFileName]));
      Ints[0, Radek] := 0;
      Row := Radek;
    end;
  end;  // with fmMain
end;


end.

