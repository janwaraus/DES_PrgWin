// 2.10.2014 program pro mìsíèní fakturaci sjednocený pro DES i ABAK. Základem jse dosavadní program MesicniFakturace.
// Verze pro ABAK má v Project/Options/Conditionals nastaveno ABAK
// Základní rozdíly mezi obìma verzemi:
// - DES fakturuje najednou všechny smlouvy zákazníka, internet, VoIP i IPTV s datem vystavení i plnìní poslední den v mìsíci.
//  Faktury pøevedené do PDF jsou elektronicky podepsané
// - ABAK fakturuje každou smlouvu zvláš, internetové smlouvy na zaèátku mìsíce (zálohovì), VoIPové ka konci. Datum vystavení
//  je aktuální datum, datum plnìní u internetových smluv aktuální datum, u VoIPových poslední den fakturovaného mìsíce
// 20.11. pro export do PDF/A (bez Print2PDF) použita Synopse
// 4.12. ABAK - VoIPové smlouvy jednoho zákazníka se fakturují spoleènì
// 19.10.2016 Faktury s pøenesenou daòovou povinností pro smlouvy s tagem DRC
// 23.1.2017 výkazy pro ÈTÚ vyžadují dìlení podle technologie a rychlosti - do Abry bylo pøidáno 36 zakázek a ve faktuøe bude jedna
// z nich pøiøazena každému øádku

unit FImain;

interface

uses
  Windows, Messages, Dialogs, SysUtils, Variants, Classes, Graphics, Controls, StdCtrls, ExtCtrls, Forms, Mask, ComObj, ComCtrls,
  AdvObj, AdvPanel, AdvEdit, AdvSpin, AdvDateTimePicker, AdvEdBtn, AdvFileNameEdit, AdvProgressBar, GradientLabel,
  Grids, BaseGrid, AdvGrid, //pCore2D, pBarcode2D,
  IniFiles, DateUtils, Math,
  DB, ZAbstractConnection, ZConnection, ZAbstractRODataset, ZAbstractDataset, ZDataset,

  frxClass, frxDBSet, frxDesgn, AdvUtil;

type
  TfmMain = class(TForm)
    qrMain: TZQuery;
    qrSmlouva: TZQuery;
    dbVoIP: TZConnection;
    qrVoIP: TZQuery;
    qrAbra: TZQuery;
    qrAdresa: TZQuery;
    apnVyberCinnosti: TAdvPanel;
    rbFakturace: TRadioButton;
    rbPrevod: TRadioButton;
    rbTisk: TRadioButton;
    rbMail: TRadioButton;
    glbFakturace: TGradientLabel;
    glbPrevod: TGradientLabel;
    glbTisk: TGradientLabel;
    glbMail: TGradientLabel;
    apnTop: TAdvPanel;
    apbProgress: TAdvProgressBar;
    apnMain: TAdvPanel;
    aseMesic: TAdvSpinEdit;
    aseRok: TAdvSpinEdit;
    lbFakturyZa: TLabel;
    cbBezVoIP: TCheckBox;
    cbSVoIP: TCheckBox;
    apnFakturyZa: TAdvPanel;
    rbInternet: TRadioButton;
    rbVoIP: TRadioButton;
    deDatumDokladu: TAdvDateTimePicker;
    deDatumPlneni: TAdvDateTimePicker;
    aedSplatnost: TAdvEdit;
    aedOd: TAdvEdit;
    aedDo: TAdvEdit;
    btVytvorit: TButton;
    btKonec: TButton;
    apnPrevod: TAdvPanel;
    cbNeprepisovat: TCheckBox;
    btOdeslat: TButton;
    btSablona: TButton;
    apnTisk: TAdvPanel;
    rbBezSlozenky: TRadioButton;
    rbSeSlozenkou: TRadioButton;
    rbKuryr: TRadioButton;
    apnMail: TAdvPanel;
    fePriloha: TAdvFileNameEdit;
    lbxLog: TListBox;
    asgMain: TAdvStringGrid;
    apnVyberPodle: TAdvPanel;
    lbVyber: TLabel;
    rbPodleFaktury: TRadioButton;
    rbPodleSmlouvy: TRadioButton;
    Button1: TButton;
    procedure FormShow(Sender: TObject);
    procedure glbFakturaceClick(Sender: TObject);
    procedure glbPrevodClick(Sender: TObject);
    procedure glbTiskClick(Sender: TObject);
    procedure glbMailClick(Sender: TObject);
    procedure aseMesicChange(Sender: TObject);
    procedure aseRokChange(Sender: TObject);
    procedure aedOdChange(Sender: TObject);
    procedure aedDoChange(Sender: TObject);
    procedure aedOdKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure aedOdExit(Sender: TObject);
    procedure aedDoKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure aedDoExit(Sender: TObject);
    procedure cbSVoIPClick(Sender: TObject);
    procedure cbBezVoIPClick(Sender: TObject);
    procedure rbInternetClick(Sender: TObject);
    procedure rbVoIPClick(Sender: TObject);
    procedure rbBezSlozenkyClick(Sender: TObject);
    procedure rbSeSlozenkouClick(Sender: TObject);
    procedure rbKuryrClick(Sender: TObject);
    procedure asgMainCanEditCell(Sender: TObject; ARow, ACol: Integer; var CanEdit: Boolean);
    procedure asgMainClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure asgMainCanSort(Sender: TObject; ACol: Integer; var DoSort: Boolean);
    procedure asgMainGetAlignment(Sender: TObject; ARow, ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asgMainDblClick(Sender: TObject);
    procedure lbxLogDblClick(Sender: TObject);
    procedure btVytvoritClick(Sender: TObject);
    procedure btOdeslatClick(Sender: TObject);
    procedure btSablonaClick(Sender: TObject);
    procedure btKonecClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure rbPodleSmlouvyClick(Sender: TObject);
    procedure rbPodleFakturyClick(Sender: TObject);
    procedure rbFakturaceClick(Sender: TObject);
    procedure rbPrevodClick(Sender: TObject);
    procedure rbTiskClick(Sender: TObject);
    procedure rbMailClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  public
    //F: TextFile;

    isDebugMode,
    Prerusit: boolean;


    AbraOLE: variant;

    User_Id: string[10]; //User_ID do ABRY


    fiVoipCustomersView,
    fiBBmaxView,
    fiBillingView,
    fiInvoiceView: ShortString;

  end;

var
  fmMain: TfmMain;

implementation

uses DesUtils, AbraEntities, FICommon, FIfaktura, FIPrevod, FITisk, FIMail,
AArray,
DesFastReports //pro test/demo jenom
;

{$R *.dfm}


procedure TfmMain.FormShow(Sender: TObject);
var
  FileHandle: integer;
  FIIni: TIniFile;
  LogDir,
  LogFileName,
  FIFileName: string;
  abraVatIndex : TAbraVatIndex;
  abraDrcArticle : TAbraDrcArticle;

begin
  //adresáø pro logy
  LogDir := DesU.PROGRAM_PATH + '\logy\Mìsíèní fakturace\';
  if not DirectoryExists(LogDir) then Forcedirectories(LogDir);
  globalAA['LogFileName'] := LogDir + FormatDateTime('yyyy.mm".log"', Date);
  DesUtils.appendToFile(globalAA['LogFileName'],''); //vloží prázdný øádek do logu
  dmCommon.Zprava('Start programu "Mìsíèní fakturace".');


  // jména pro viewjsou unikátní, aby program nebyl omezen na jednu instanci
  fiVoipCustomersView := FormatDateTime('VoIPyymmddhhnnss', Now);
  fiBBmaxView := FormatDateTime('BByymmddhhnnss', Now);
  fiBillingView := FormatDateTime('BVyymmddhhnnss', Now);
  fiInvoiceView := FormatDateTime('IVyymmddhhnnss', Now);

  // do 25. se oèekává fakturace za minulý mìsíc, pak už za aktuální
  if DayOf(Date) > 25 then begin
    aseMesic.Value := MonthOf(Date);
    aseRok.Value := YearOf(Date);
  end else begin
    aseMesic.Value := MonthOf(IncMonth(Date, -1));
    aseRok.Value := YearOf(IncMonth(Date, -1));
  end;


  // fajfky v asgMain
  asgMain.CheckFalse := '0';
  asgMain.CheckTrue := '1';
  apnFakturyZa.Visible := False;

  aseRokChange(nil);
  rbFakturaceClick(nil);

  //nastavení globálních promìnných pøi startu programu
  globalAA['PDFDir'] := DesU.getIniValue('Preferences', 'PDFDir');

  globalAA['invoiceDocQueueCode'] := 'FO1'; //kód øady faktur, tento program vystavuje pouze do této øady
  globalAA['abraIiDocQueue_Id'] := DesU.getAbraDocqueueId('FO1', '03');

  abraVatIndex := TAbraVatIndex.create('Výst21');
  globalAA['abraVatIndex_Id'] := abraVatIndex.id; //VATIndex_Id
  globalAA['abraVatRate_Id'] := abraVatIndex.vatrateId; //VATRate_Id
  globalAA['abraVatRate'] := abraVatIndex.tariff; //VATRate

  abraVatIndex := TAbraVatIndex.create('VýstR21');
  globalAA['abraDrcVatIndex_Id'] := abraVatIndex.id; // DRCVATIndex_Id

  abraDrcArticle := TAbraDrcArticle.create('21');
  globalAA['abraDrcArticle_Id'] := abraDrcArticle.id;

  // pro text, TODO smazat
  aedOd.Text := '10200555';
  aedDo.Text := '10205555';

end;


procedure TfmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  with DesU.qrZakos do try
    SQL.Text := 'DROP VIEW ' + fiVoipCustomersView;
    ExecSQL;
    SQL.Text := 'DROP VIEW ' + fiInvoiceView;
    ExecSQL;
    SQL.Text := 'DROP VIEW ' + fiBillingView;
    ExecSQL;
    SQL.Text := 'DROP VIEW ' + fiBBmaxView;
    ExecSQL;
  except
  end;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.btVytvoritClick(Sender: TObject);
// pro zadaný mìsíc se vytvoøí faktury v Abøe, nebo pøevedou do PDF, vytisknou èi rozešlou mailem
begin
  btVytvorit.Enabled := False;
  btKonec.Caption := '&Pøerušit';
  Prerusit := False;
  Application.ProcessMessages;
  try

    // *** Naplnìní asgMain ***
    if btVytvorit.Caption = '&Naèíst' then dmCommon.Plneni_asgMain
    else begin

      // *** Fakturace ***
      if rbFakturace.Checked then dmFaktura.VytvorFaktury;

      // *** Pøevod do PDF ***
      if rbPrevod.Checked then dmPrevod.PrevedFaktury;

      // *** Tisk ***
      if rbTisk.Checked then dmTisk.TiskniFaktury;

      // *** Posílání mailem ***
      if rbMail.Checked then dmMail.PosliFaktury;

    end;
  finally
    btKonec.Caption := '&Konec';
    btVytvorit.Enabled := True;
  end;
end;

procedure TfmMain.btKonecClick(Sender: TObject);
begin
  if btKonec.Caption = '&Pøerušit' then begin
    Prerusit := True;
    dmCommon.Zprava('Pøerušeno uživatelem.');
    btKonec.Caption := '&Konec';
  end else begin
    dmCommon.Zprava('Konec programu "Mìsíèní fakturace".');
    Close;
  end;
end;

procedure TfmMain.btSablonaClick(Sender: TObject);
begin
  // *hw* TODO frxReport.DesignReport(True, False);
end;

procedure TfmMain.btOdeslatClick(Sender: TObject);
// odeslání faktur pøevedených do PDF na vzdálený server
begin

  //WinExec(PChar(Format('WinSCP.com /command "option batch abort" "option confirm off" "open AbraPDF" "synchronize remote '
  // + '%s\%4d\%2.2d /home/abrapdf/%4d" "exit"', [PDFDir, aseRok.Value, aseMesic.Value, aseRok.Value])), SW_SHOWNORMAL);

  RunCMD (Format('WinSCP.com /command "option batch abort" "option confirm off" "open AbraPDF" "synchronize remote '
   + '%s\%4d\%2.2d /home/abrapdf/%4d" "exit"', [globalAA['PDFDir'], aseRok.Value, aseMesic.Value, aseRok.Value]), SW_SHOWNORMAL);

end;


// ------------------------------------------------------------------------------------------------


procedure TfmMain.rbFakturaceClick(Sender: TObject);
begin
  asgMain.Cells[0, 0] := 'fa';
  asgMain.ColWidths[6] := 0;
  glbFakturace.Color := clWhite;
  glbFakturace.ColorTo := clMenu;
  glbPrevod.Color := clSilver;
  glbPrevod.ColorTo := clGray;
  apnPrevod.Visible := False;
  glbTisk.Color := clSilver;
  glbTisk.ColorTo := clGray;
  apnTisk.Visible := False;
  glbMail.Color := clSilver;
  glbMail.ColorTo := clGray;
  apnMail.Visible := False;
  rbPodleSmlouvy.Checked := True;
  rbPodleFaktury.Enabled := False;
  apnVyberPodle.Visible := False;
  apbProgress.Visible := False;
  lbxLog.Visible := True;
  aseRokChange(Self);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.rbPrevodClick(Sender: TObject);
begin
  asgMain.Cells[0, 0] := 'PDF';
  asgMain.ColWidths[6] := 0;
  glbPrevod.Color := clWhite;
  glbPrevod.ColorTo := clMenu;
  apnPrevod.Visible := True;
  glbTisk.Color := clSilver;
  glbTisk.ColorTo := clGray;
  apnTisk.Visible := False;
  glbMail.Color := clSilver;
  glbMail.ColorTo := clGray;
  apnMail.Visible := False;
  glbFakturace.Color := clSilver;
  glbFakturace.ColorTo := clGray;
  rbPodleFaktury.Enabled := True;
  rbPodleFaktury.Checked := True;
  apnVyberPodle.Visible := True;
  aseRokChange(Self);


  {* HW testovaci nastaveni *}
  aseMesic.Value := 2;
  aedOd.Text := '5000';
  aedDo.Text := '5003';

end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.rbTiskClick(Sender: TObject);
begin
  asgMain.Cells[0, 0] := 'tisk';
  asgMain.ColWidths[6] := 60;
  glbTisk.Color := clWhite;
  glbTisk.ColorTo := clMenu;
  apnTisk.Visible := True;
  glbMail.Color := clSilver;
  glbMail.ColorTo := clGray;
  apnMail.Visible := False;
  glbFakturace.Color := clSilver;
  glbFakturace.ColorTo := clGray;
  glbPrevod.Color := clSilver;
  glbPrevod.ColorTo := clGray;
  apnPrevod.Visible := False;
  rbPodleFaktury.Enabled := True;
  rbPodleFaktury.Checked := True;
  apnVyberPodle.Visible := True;
  aseRokChange(Self);

  {* HW testovaci nastaveni *}
  aseMesic.Value := 2;
  aedOd.Text := '5000';
  aedDo.Text := '5003';


end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.rbMailClick(Sender: TObject);
begin
  asgMain.Cells[0, 0] := 'mail';
  asgMain.ColWidths[6] := 60;
  glbMail.Color := clWhite;
  glbMail.ColorTo := clMenu;
  apnMail.Visible := True;
  glbFakturace.Color := clSilver;
  glbFakturace.ColorTo := clGray;
  glbPrevod.Color := clSilver;
  glbPrevod.ColorTo := clGray;
  apnPrevod.Visible := False;
  glbTisk.Color := clSilver;
  glbTisk.ColorTo := clGray;
  apnTisk.Visible := False;
  rbPodleFaktury.Enabled := True;
  rbPodleFaktury.Checked := True;
  apnVyberPodle.Visible := True;
  aseRokChange(Self);
end;

// ------------------------------------------------------------------------------------------------


procedure TfmMain.aseMesicChange(Sender: TObject);
begin
  if (aseMesic.Value = 0) then begin
    aseMesic.Value := 12;
    aseRok.Value := aseRok.Value - 1;
  end;
  if (aseMesic.Value = 13) then begin
    aseMesic.Value := 1;
    aseRok.Value := aseRok.Value + 1;
  end;
  aseRokChange(nil);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.aseRokChange(Sender: TObject);
// pøi zmìnì (nejen) roku nastaví nové deDatumDokladu, deDatumPlneni, aedSplatnost, aedOd a aedDo
begin
  if (not DesU.dbZakos.Connected) or (not DesU.dbAbra.Connected) then begin
    ShowMessage('Nejsou pøipravené databáze (TfmMain.aseRokChange)');
    Exit;
  end;

  aedOd.Clear;
  aedDo.Clear;
  asgMain.ClearNormalCells;
  asgMain.RowCount := 2;
  btVytvorit.Caption := '&Naèíst';
  asgMain.Visible := True;
  lbxLog.Visible := False;


  globalAA['abraIiPeriod_Id'] := DesU.getAbraPeriodId(aseRok.Text); // tohle ale dát jinam, resp. vyhodit a øešit Period_ID až pøi tvorbì fa (jen tehdy je potøeba)

// datum fakturace i datum plnìní je poslední den v mìsíci
  deDatumDokladu.Date := EndOfAMonth(aseRok.Value, aseMesic.Value);
  deDatumPlneni.Date := deDatumDokladu.Date;
  aedSplatnost.Text := '10';

  // *** výbìr podle smlouvy
  if rbPodleSmlouvy.Checked then begin
    with DesU.qrZakos do try
      Screen.Cursor := crSQLWait;

      // view pro fakturaci
      dmCommon.AktualizaceView;

      // první a poslední èíslo smlouvy
      SQL.Text := 'SELECT MIN(VS), MAX(VS) FROM fiInvoiceView';
      Open;
      aedOd.Text := Fields[0].AsString;
      aedDo.Text := Fields[1].AsString;
      Close;
    finally
      Screen.Cursor := crDefault;
    end;
  end;

  // *** výbìr podle faktury
  if rbPodleFaktury.Checked then begin
    DesU.dbAbra.Reconnect;
    with DesU.qrAbra do begin
      // rozpìtí èísel FO1 v mìsíci
      SQL.Text := 'SELECT MIN(OrdNumber), MAX(OrdNumber) FROM IssuedInvoices'
      + ' WHERE VATDate$DATE >= ' + FloatToStr(Trunc(StartOfAMonth(aseRok.Value, aseMesic.Value)))
      + ' AND VATDate$DATE <= ' + FloatToStr(Trunc(EndOfAMonth(aseRok.Value, aseMesic.Value)))
      + ' AND DocQueue_ID = ' + Ap + globalAA['abraIiDocQueue_Id'] + Ap;
      Open;
      if RecordCount > 0 then begin
        aedOd.Text := Fields[0].AsString;
        aedDo.Text := Fields[1].AsString;
        Close;
      end else begin
        aedOd.Clear;
        aedDo.Clear;
      end;
    end;
  end;

end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.aedOdChange(Sender: TObject);
begin
  if btVytvorit.Caption <> '&Naèíst' then begin
    asgMain.ClearNormalCells;
    asgMain.RowCount := 2;
    btVytvorit.Caption := '&Naèíst';
  end;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.aedOdKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
// enter po zadání aedOd vyplní stejným èíslem aedDo
begin
  if Key = 13 then begin
    aedDo.Text := aedOd.Text;
    aedDo.SetFocus;
  end;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.aedOdExit(Sender: TObject);
// je-li aedOd vìtší než aedDo, opraví se
begin
//  if aedOd.IntValue > aedDo.IntValue then aedDo.Text := aedOd.Text;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.aedDoChange(Sender: TObject);
begin
  aedOdChange(nil);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.aedDoKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
// enter po zadání aedDo spustí pøevod
begin
  if Key = 13 then btVytvorit.SetFocus;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.aedDoExit(Sender: TObject);
// je-li aedOd menší než aedDo, opraví se
begin
//  if aedDo.IntValue < aedOd.IntValue then aedOd.Text := aedDo.Text;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.rbInternetClick(Sender: TObject);
// ABAK
begin
  aseRokChange(nil);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.rbVoIPClick(Sender: TObject);
begin
  aseRokChange(nil);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.cbBezVoIPClick(Sender: TObject);
// DES - musí být nìco vybráno
begin
  if not (cbBezVoIP.Checked or cbSVoIP.Checked) then cbSVoIP.Checked := True;
  aseRokChange(nil);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.cbSVoIPClick(Sender: TObject);
begin
  if not (cbBezVoIP.Checked or cbSVoIP.Checked) then cbBezVoIP.Checked := True;
  aseRokChange(nil);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.rbPodleSmlouvyClick(Sender: TObject);
begin
  aseRokChange(nil);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.rbPodleFakturyClick(Sender: TObject);
begin
  aseRokChange(nil);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.rbBezSlozenkyClick(Sender: TObject);
begin
  aedOdChange(nil);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.rbSeSlozenkouClick(Sender: TObject);
begin
  aedOdChange(nil);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.rbKuryrClick(Sender: TObject);
begin
  aedOdChange(nil);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.asgMainDblClick(Sender: TObject);
begin
  asgMain.Visible := False;
  lbxLog.Visible := True;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.asgMainCanEditCell(Sender: TObject; ARow, ACol: Integer; var CanEdit: Boolean);
begin
  CanEdit := (ARow > 0) and (ACol in [0, 5]);   // fajfky nebo email
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.asgMainClickCell(Sender: TObject; ARow, ACol: Integer);
// klik v prvním øádku podle sloupce buï oznaèuje/odznaèuje všchny checkboxy, nebo spouští tøídìní,
var
  Radek: integer;
begin
  with asgMain do if (ARow = 0) and (ACol in [0..1]) then
    if Ints[ACol, 1] = 1 then for Radek := 1 to RowCount-1 do Ints[ACol, Radek] := 0
    else if Ints[ACol, 1] = 0 then for Radek := 1 to RowCount-1 do Ints[ACol, Radek] := 1;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.asgMainCanSort(Sender: TObject; ACol: Integer; var DoSort: Boolean);
// klik v prvním øádku podle sloupce buï oznaèuje/odznaèuje všchny checkboxy, nebo spouští tøídìní,
begin
  DoSort := ACol > 0;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.asgMainGetAlignment(Sender: TObject; ARow, ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
  HAlign := taLeftJustify;
  if (ACol = 0) or (ARow = 0) then HAlign := Classes.taCenter
  else if (ACol in [1..3]) and (ARow > 0) then HAlign := taRightJustify;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.lbxLogDblClick(Sender: TObject);
begin
  asgMain.Visible := True;
  lbxLog.Visible := False;
end;


// ------------------------------------------------------------------------------------------------


procedure TfmMain.glbFakturaceClick(Sender: TObject);
begin
  rbFakturace.Checked := True;
  rbFakturaceClick(nil);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.glbPrevodClick(Sender: TObject);
begin
  rbPrevod.Checked := True;
  rbPrevodClick(nil);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.glbTiskClick(Sender: TObject);
begin
  rbTisk.Checked := True;
  rbTiskClick(nil);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.glbMailClick(Sender: TObject);
begin
  rbMail.Checked := True;
  rbMailClick(nil);
end;

// ------------------------------------------------------------------------------------------------


procedure TfmMain.Button1Click(Sender: TObject);
var
  source : string;
  i: integer;
  rData: TAArray;
  temps : string;
  vysledek : TDesResult;
begin

  DesFastReport.init('invoice', 'FOsPDP.fr3'); // nastavení typu reportu a fr3 souboru
  DesFastReport.setExportDirName(Format('%s\%4d\%2.2d\', [globalAA['PDFDir'], 2023, 15])); // 15 pro legraci

    vysledek := dmPrevod.fakturaPrevod('1K6P200101', true);  // pokud zaškrtnuto, pøevádíme fa do PDF
    //  '1K6P200101' '4K6P200101' '7K6P200101'
    dmCommon.Zprava(Format('%s', [vysledek.Messg]));

//ShowMessage('demooo_' +  BoolToStr(DesU.existujeVAbreDokladSPrazdnymVs(), true));
//ShowMessage('VRid_' +  DesU.getAbraVatrateId('Výst21'));
//ShowMessage('VRid_' +  DesU.getAbraVatindexId('Výst21'));

//Source := 'pakachar:ukulelee';
//Copy(Source, 8, MaxInt);

//ShowMessage('1_' +  Copy(Source, Pos(':', Source)+1, MaxInt) + '_');
//ShowMessage('1_' +  Copy(Source, 1, Pos(':', Source)-1) + '_');
//ShowMessage(inttostr(Pos('-', Source)) );


{
Zaplatit := 152.61;

    C := Format('%6.0f', [Zaplatit]);
    //nahradíme vlnovkou poslední mezeru, tedy dáme vlnovku pøed první èíslici
    for i := 2 to 6 do
      if C[i] <> ' ' then begin
        C[i-1] := '~';
        Break;
      end;

ShowMessage('-' + C + '-' );

ShowMessage (
  '- -'
  + Format('%8.8d%2.2d', [213456, 18]) + '- -'
  + Format('%10.10d', [91234567]) + '- -'
  + Format('%10.10d', [9123456789]) + '- -'
  + Format('%8.0d%2.0d', [213456, 213456]) + '- -'
  + Format('%3.3d%3.3d', [213456, 18]) + '- -'
  + Format('%3.0d%3.0d', [213456, 18]) + '- -'

  );

}

{
  rData := TAArray.Create;
  rData['Title'] := 'Faktura za pøipojení k internetu';
  rData['Author'] := 'Družstvo Eurosignal';
  ShowMessage (rData['Title'] +' - '+rData['kuku']+'-');
}

end;


end.
