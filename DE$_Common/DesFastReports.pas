unit DesFastReports;

interface

uses
  Winapi.Windows, Winapi.ShellApi, Winapi.Messages,
  System.SysUtils, System.Variants, System.Classes, System.RegularExpressions,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  AArray, StrUtils, Data.DB,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, ZAbstractConnection, ZConnection,

  IdBaseComponent, IdComponent, IdTCPConnection,
  IdTCPClient, IdSMTP, IdHTTP, IdMessage, IdMessageClient, IdText, IdMessageParts,
  IdAntiFreezeBase, IdAntiFreeze, IdIOHandler,
  IdIOHandlerSocket, IdSSLOpenSSL, IdExplicitTLSClientServerBase, IdSMTPBase, IdAttachmentFile,

  frxClass, frxDBSet, frxDesgn, frxBarcode, frxBarcode2d,
  frxExportBaseDialog, frxExportPDF, frxExportPDFHelpers,

  DesUtils, AbraEntities;

type
  TDesFastReport = class(TForm)
    frxReport: TfrxReport;
    frxPDFExport1: TfrxPDFExport;
    frxBarCodeObject1: TfrxBarCodeObject;
    qrAbraDPH: TZQuery;
    qrAbraRadky: TZQuery;
    fdsRadky: TfrxDBDataset;
    fdsDPH: TfrxDBDataset;
    idMessage: TIdMessage;
    idSMTP: TIdSMTP;

    procedure frxReportGetValue(const ParName: string; var ParValue: Variant);
    procedure frxReportBeginDoc(Sender: TObject);

  private

  public
    reportData: TAArray;
    reportType,
    exportDirName,
    exportFileName,
    fr3FileName,
    invoiceId : string;

  published
    procedure init(pReportType, pFr3FileName : string);
    procedure setExportDirName(newExportDirName : string);
    procedure loadFr3File(iFr3FileName : string = '');
    procedure prepareInvoiceDataSets(invoiceId : string);
    function createPdf(FullPdfFileName : string; OverwriteExistingPdf : boolean) : TDesResult;
    function print(fr3FileName : string) : TDesResult;

  end;

var
  DesFastReport: TDesFastReport;

implementation

{$R *.dfm}

procedure TDesFastReport.init(pReportType, pFr3FileName : string);
begin
  { asi není potøeba oddalovat load FR3 souboru, tedy udìlá se hned takto nazaèátku
    pøípadnì je možné nejdøív nastavit filename self.fr3FileName := pFr3FileName;
    a pozdìji load pomocí self.loadFr3File(); }
  self.reportType := pReportType; // zatím se nepoužívá, máme jen faktury, ale do budoucna
  self.loadFr3File(pFr3FileName);
end;


procedure TDesFastReport.setExportDirName(newExportDirName : string);
begin
  self.exportDirName := newExportDirName;
  if not DirectoryExists(newExportDirName) then Forcedirectories(newExportDirName);
end;

{ Do reportu natáhne FR3 soubor, tím umožní práci s promìnnými atd
  pokud je jako vstupní parametr náyev souboru, je použit tento nový soubor.
  FR3 soubor mùže být umístìný buï pøímo v adresáøi s programem,
  nebo v ..\DE$_Common\resources\
}
procedure TDesFastReport.loadFr3File(iFr3FileName : string = '');
var
  fr3File : string;
begin
  if iFr3FileName = '' then
    iFr3FileName := self.fr3FileName
  else
    self.fr3FileName := iFr3FileName;

  fr3File := DesU.PROGRAM_PATH + self.fr3FileName;
  if not(FileExists(fr3File)) then
    fr3File := DesU.PROGRAM_PATH + '..\DE$_Common\resources\' + self.fr3FileName;

  if not(FileExists(fr3File)) then begin
    MessageDlg('Nenalezen FR3 soubor ' + fr3File, mtError, [mbOk], 0);
    // a mìla by se ideálnì vyhodit Exception a ta zpracovat na správném místì
  end
  else
    frxReport.LoadFromFile(fr3File);
end;


function TDesFastReport.createPdf(FullPdfFileName : string; OverwriteExistingPdf : boolean) : TDesResult;
var
  i : integer;
  barcode: TfrxBarcode2DView;
begin

  if FileExists(FullPdfFileName) AND (not OverwriteExistingPdf) then begin
    Result := TDesResult.create('err', Format('Soubor %s už existuje.', [FullPdfFileName]));
    Exit;
  end else
    DeleteFile(FullPdfFileName);

  frxReport.PrepareReport;

  with frxPDFExport1 do begin
    FileName := FullPdfFileName;
    Title := reportData['Title'];
    Author := reportData['Author'];
    {Set the PDF standard
     TPDFStandard = (psNone, psPDFA_1a, psPDFA_1b, psPDFA_2a, psPDFA_2b, psPDFA_3a, psPDFA_3b);
    It is required to add the frxExportPDFHelpers module to the uses list:
    uses frxExportPDFHelpers;}
    PDFStandard := psPDFA_1a; //PDF/A-1a satisfies all requirements in the specification
    PDFVersion := pv14;
    Compressed := True;
    EmbeddedFonts := False;
    {Disable export of objects with optimization for printing. With option enabled images will be high-quality but 9 times larger in volume}
    //frxPDFExport1.PrintOptimized := False;
    OpenAfterExport := False;
    ShowDialog := False;
    ShowProgress := False;
  end;

  frxReport.Export(frxPDFExport1);

  //frxReport.ShowReport; // zobrazí report v samostatném oknì, pro test hlavnì

  // èekání na soubor - max. 5s
  for i := 1 to 50 do begin
    if FileExists(FullPdfFileName) then Break;
    Sleep(100);
  end;

  // hotovo
  if not FileExists(FullPdfFileName) then
    Result := TDesResult.create('err', Format('Nepodaøilo se vytvoøit soubor %s.', [fullPdfFileName]))
  else begin
    Result := TDesResult.create('ok', Format('Vytvoøen soubor %s.', [fullPdfFileName]));
  end;

end;

function TDesFastReport.print(fr3FileName : string) : TDesResult;
begin
  try
    frxReport.LoadFromFile(DesU.PROGRAM_PATH + fr3FileName);
    frxReport.PrepareReport;
    //frxReport.PrintOptions.ShowDialog := true;
    frxReport.PrintOptions.ShowDialog := false;
    frxReport.Print;
  except on E: exception do
    begin
      Result := TDesResult.create('err', Format('%s (%s): Fakturu %s se nepodaøilo vytisknout. Chyba: %s',
       [reportData['OJmeno'], reportData['VS'], reportData['Cislo'], E.Message]));
      Exit;
    end;
  end;

  Result := TDesResult.create('ok', 'Tisk OK');
end;


// ------------------------------------------------------------------------------------------------


procedure TDesFastReport.frxReportGetValue(const ParName: string; var ParValue: Variant);
// dosadí se promìné do formuláøe
begin

if ParName = 'Value = 0' then Exit; //pro jistotu, ve fr3 souboru toto bylo v highlight.condition

  try
    ParValue := self.reportData[ParName];
  except
    ShowMessage('Lehlo to na ' + ParName); //nefunguje mi
    //on E: Exception do
    //  ShowMessage(ParName + ' Chyba frxReportGetValue: '#13#10 + e.Message);
  end;

end;


procedure TDesFastReport.frxReportBeginDoc(Sender: TObject);
var
    barcode: TfrxBarcode2DView;
begin
  //if varIsType(reportData['sQrKodem'], varBoolean) AND reportData['sQrKodem'] then begin   // takhle to mìl táta
  if reportData['QRText'] <> '' then begin // pokud  AArray nemá klíè hledané hodnoty, value je vždy '', to testujeme
   barcode := TfrxBarcode2DView(frxReport.FindObject('pQR'));
   barcode.Text := reportData['QRText'];
  end;
end;


procedure TDesFastReport.prepareInvoiceDataSets(invoiceId : string);
begin

  // øádky faktury
  with qrAbraRadky do begin
    Close;
    SQL.Text := 'SELECT Text, TAmountWithoutVAT AS BezDane, VATRate AS Sazba, TAmount - TAmountWithoutVAT AS DPH, TAmount AS SDani FROM IssuedInvoices2'
    + ' WHERE Parent_ID = ' + Ap + invoiceId + Ap
    + ' AND NOT (Text = ''Zaokrouhlení'' AND TAmount = 0)'
    + ' ORDER BY PosIndex';
    Open;
  end;

  // rekapitulace DPH
  with qrAbraDPH do begin
    Close;
    SQL.Text := 'SELECT VATRate AS Sazba, SUM(TAmountWithoutVAT) AS BezDane, SUM(TAmount - TAmountWithoutVAT) AS DPH, SUM(TAmount) AS SDani FROM IssuedInvoices2'
    + ' WHERE Parent_ID = ' + Ap + invoiceId + Ap
    + ' AND VATIndex_ID IS NOT NULL'
    + ' GROUP BY Sazba';
    Open;
  end;

  qrAbraRadky.Close;
  qrAbraDPH.Close;
end;

end.
