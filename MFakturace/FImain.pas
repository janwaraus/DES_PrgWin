// 2.10.2014 program pro m�s��n� fakturaci sjednocen� pro DES i ABAK. Z�kladem jse dosavadn� program MesicniFakturace.
// Verze pro ABAK m� v Project/Options/Conditionals nastaveno ABAK
// Z�kladn� rozd�ly mezi ob�ma verzemi:
// - DES fakturuje najednou v�echny smlouvy z�kazn�ka, internet, VoIP i IPTV s datem vystaven� i pln�n� posledn� den v m�s�ci.
//  Faktury p�eveden� do PDF jsou elektronicky podepsan�
// - ABAK fakturuje ka�dou smlouvu zvl᚝, internetov� smlouvy na za��tku m�s�ce (z�lohov�), VoIPov� ka konci. Datum vystaven�
//  je aktu�ln� datum, datum pln�n� u internetov�ch smluv aktu�ln� datum, u VoIPov�ch posledn� den fakturovan�ho m�s�ce
// 20.11. pro export do PDF/A (bez Print2PDF) pou�ita Synopse
// 4.12. ABAK - VoIPov� smlouvy jednoho z�kazn�ka se fakturuj� spole�n�
// 19.10.2016 Faktury s p�enesenou da�ovou povinnost� pro smlouvy s tagem DRC
// 23.1.2017 v�kazy pro �T� vy�aduj� d�len� podle technologie a rychlosti - do Abry bylo p�id�no 36 zak�zek a ve faktu�e bude jedna
// z nich p�i�azena ka�d�mu ��dku

unit FImain;

interface

uses
  Windows, Messages, Dialogs, SysUtils, Variants, Classes, Graphics, Controls, StdCtrls, ExtCtrls, Forms, Mask, ComObj, ComCtrls,
  AdvObj, AdvPanel, AdvEdit, AdvSpin, AdvDateTimePicker, AdvEdBtn, AdvFileNameEdit, AdvProgressBar, GradientLabel,
  Grids, BaseGrid, AdvGrid, pCore2D, pBarcode2D, IniFiles, DateUtils, Math,
  DB, ZAbstractConnection, ZConnection, ZAbstractRODataset, ZAbstractDataset, ZDataset,

  frxClass, frxDBSet, frxDesgn;

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
    F: TextFile;
    DRC,
    Check,
    Prerusit: boolean;
    //Mesic,
    VATRate: integer;
    Celkem,
    Zaplaceno,
    Saldo,
    Zaplatit,
    DatumDokladu,
    DatumPlneni,
    DatumSplatnosti: double;

    //pak smazat

    AbraOLE: variant;

    C, V, S: string[10];                           // pole ��sel na slo�enku
    User_Id: string[10]; //User_ID do ABRY

    ID,
    Firm_Id,
    //Period_Id,
    //IiDocQueue_Id,
    //VDocQueue_Id,
    DRCArticle_Id,
    DRCVATIndex_Id,
    VATRate_Id,
    VATIndex_Id: string[10];
    FStr,                              // prefix faktury TODO smazat
    PDFDir,
    fiVoipCustomersView,
    fiBBmaxView,
    fiBillingView,
    fiInvoiceView: ShortString;
    Cislo,
    VS,
    SS,
    PJmeno,
    PUlice,
    PObec,
    OJmeno,
    OUlice,
    OObec,
    OICO,
    ODIC,
    Vystaveni,
    Plneni,
    Splatnost,
    Platek: AnsiString;
  end;

var
  fmMain: TfmMain;

implementation

uses DesUtils, AbraEntities, FICommon, FIfaktura, FIPrevod, FITisk, FIMail,
AArray
;

{$R *.dfm}


procedure TfmMain.FormShow(Sender: TObject);
var
  FileHandle: integer;
  FIIni: TIniFile;
  LogDir,
  LogFileName,
  FIFileName: AnsiString;
  abraVatIndex : TAbraVatIndex;
  abraDrcArticle : TAbraDrcArticle;

begin
  //adres�� pro logy
  LogDir := DesU.PROGRAM_PATH + '\logy\M�s��n� fakturace\';
  if not DirectoryExists(LogDir) then Forcedirectories(LogDir);

// vytvo�en� logfile, pokud neexistuje - 5.11. do jm�na p�id�no datum - 9.4.15 jen rok a m�s�c
// 2.1.  LogFileName := ExtractFilePath(ParamStr(0)) + FormatDateTime('"Fakturace "dd.mm.yyyy".log"', Date);
// 8.4.  LogFileName := LogDir + FormatDateTime('\yyyy.mm.dd".log"', Date);
  LogFileName := LogDir + FormatDateTime('yyyy.mm".log"', Date);
  if not FileExists(LogFileName) then begin
    FileHandle := FileCreate(LogFileName);
    FileClose(FileHandle);
  end;
  AssignFile(F, LogFileName);
  Append(F);
  Writeln(F);
  CloseFile(F);
  dmCommon.Zprava('Start programu "M�s��n� fakturace".');


  // jm�na pro viewjsou unik�tn�, aby program nebyl omezen na jednu instanci
  fiVoipCustomersView := FormatDateTime('VoIPyymmddhhnnss', Now);
  fiBBmaxView := FormatDateTime('BByymmddhhnnss', Now);
  fiBillingView := FormatDateTime('BVyymmddhhnnss', Now);
  fiInvoiceView := FormatDateTime('IVyymmddhhnnss', Now);

  // do 25. se o�ek�v� fakturace za minul� m�s�c, pak u� za aktu�ln�
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


  //nastaven� glob�ln�ch prom�nn�ch p�i startu programu
  PDFDir := DesU.getIniValue('Preferences', 'PDFDir');

  globalAA['abraIiDocQueue_Id'] := DesU.getAbraDocqueueId('FO1', '03');

  abraVatIndex := TAbraVatIndex.create('V�st21');
  globalAA['abraVatIndex_Id'] := abraVatIndex.id;
  globalAA['abraVatRate_Id'] := abraVatIndex.vatrateId;
  globalAA['abraVatRate'] := abraVatIndex.tariff;

  abraVatIndex := TAbraVatIndex.create('V�stR21');
  globalAA['abraDrcVatIndex_Id'] := abraVatIndex.id;

  abraDrcArticle := TAbraDrcArticle.create('21');
  globalAA['abraDrcArticle_Id'] := abraDrcArticle.id;

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
// p�i zm�n� (nejen) roku nastav� nov� deDatumDokladu, deDatumPlneni, aedSplatnost, aedOd a aedDo
begin
  if (not DesU.dbZakos.Connected) or (not DesU.dbAbra.Connected) then begin
    ShowMessage('Nejsou p�ipraven� datab�ze (TfmMain.aseRokChange)');
    Exit;
  end;

  aedOd.Clear;
  aedDo.Clear;
  asgMain.ClearNormalCells;
  asgMain.RowCount := 2;
  btVytvorit.Caption := '&Na��st';
  asgMain.Visible := True;
  lbxLog.Visible := False;


  globalAA['abraIiPeriod_Id'] := DesU.getAbraPeriodId(aseRok.Text); // tohle ale d�t jinam

// datum fakturace i datum pln�n� je posledn� den v m�s�ci
  deDatumDokladu.Date := EndOfAMonth(aseRok.Value, aseMesic.Value);
  deDatumPlneni.Date := deDatumDokladu.Date;
  aedSplatnost.Text := '10';

  // *** v�b�r podle smlouvy
  if rbPodleSmlouvy.Checked then begin
    with DesU.qrZakos do try
      Screen.Cursor := crSQLWait;

      // view pro fakturaci
      dmCommon.AktualizaceView;

      // prvn� a posledn� ��slo smlouvy
      SQL.Text := 'SELECT MIN(VS), MAX(VS) FROM fiInvoiceView';
      Open;
      aedOd.Text := Fields[0].AsString;
      aedDo.Text := Fields[1].AsString;
      Close;
    finally
      Screen.Cursor := crDefault;
    end;
  end;

  // *** v�b�r podle faktury
  if rbPodleFaktury.Checked then begin
    DesU.dbAbra.Reconnect;
    with DesU.qrAbra do begin
      // rozp�t� ��sel FO1 v m�s�ci
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
  if btVytvorit.Caption <> '&Na��st' then begin
    asgMain.ClearNormalCells;
    asgMain.RowCount := 2;
    btVytvorit.Caption := '&Na��st';
  end;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.aedOdKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
// enter po zad�n� aedOd vypln� stejn�m ��slem aedDo
begin
  if Key = 13 then begin
    aedDo.Text := aedOd.Text;
    aedDo.SetFocus;
  end;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.aedOdExit(Sender: TObject);
// je-li aedOd v�t�� ne� aedDo, oprav� se
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
// enter po zad�n� aedDo spust� p�evod
begin
  if Key = 13 then btVytvorit.SetFocus;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.aedDoExit(Sender: TObject);
// je-li aedOd men�� ne� aedDo, oprav� se
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
// DES - mus� b�t n�co vybr�no
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
// klik v prvn�m ��dku podle sloupce bu� ozna�uje/odzna�uje v�chny checkboxy, nebo spou�t� t��d�n�,
var
  Radek: integer;
begin
  with asgMain do if (ARow = 0) and (ACol in [0..1]) then
    if Ints[ACol, 1] = 1 then for Radek := 1 to RowCount-1 do Ints[ACol, Radek] := 0
    else if Ints[ACol, 1] = 0 then for Radek := 1 to RowCount-1 do Ints[ACol, Radek] := 1;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.asgMainCanSort(Sender: TObject; ACol: Integer; var DoSort: Boolean);
// klik v prvn�m ��dku podle sloupce bu� ozna�uje/odzna�uje v�chny checkboxy, nebo spou�t� t��d�n�,
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

procedure TfmMain.btVytvoritClick(Sender: TObject);
// pro zadan� m�s�c se vytvo�� faktury v Ab�e, nebo p�evedou do PDF, vytisknou �i roze�lou mailem
begin
  btVytvorit.Enabled := False;
  btKonec.Caption := '&P�eru�it';
  Prerusit := False;
  Application.ProcessMessages;
  try

    // *** Napln�n� asgMain ***
    if btVytvorit.Caption = '&Na��st' then dmCommon.Plneni_asgMain
    else begin

      // *** Fakturace ***
      if rbFakturace.Checked then dmFaktura.VytvorFaktury;

      // *** P�evod do PDF ***
      if rbPrevod.Checked then dmPrevod.PrevedFaktury;

      // *** Tisk ***
      if rbTisk.Checked then dmTisk.TiskniFaktury;

      // *** Pos�l�n� mailem ***
      if rbMail.Checked then dmMail.PosliFaktury;

    end;
  finally
    btKonec.Caption := '&Konec';
    btVytvorit.Enabled := True;
  end;
end;

procedure TfmMain.btKonecClick(Sender: TObject);
begin
  if btKonec.Caption = '&P�eru�it' then begin
    Prerusit := True;
    dmCommon.Zprava('P�eru�eno u�ivatelem.');
    btKonec.Caption := '&Konec';
  end else begin
    dmCommon.Zprava('Konec programu "M�s��n� fakturace".');
    Close;
  end;
end;

procedure TfmMain.btSablonaClick(Sender: TObject);
begin
  // *hw* TODO frxReport.DesignReport(True, False);
end;

procedure TfmMain.btOdeslatClick(Sender: TObject);
// odesl�n� faktur p�eveden�ch do PDF na vzd�len� server
begin

  //WinExec(PChar(Format('WinSCP.com /command "option batch abort" "option confirm off" "open AbraPDF" "synchronize remote '
  // + '%s\%4d\%2.2d /home/abrapdf/%4d" "exit"', [PDFDir, aseRok.Value, aseMesic.Value, aseRok.Value])), SW_SHOWNORMAL);

  RunCMD (Format('WinSCP.com /command "option batch abort" "option confirm off" "open AbraPDF" "synchronize remote '
   + '%s\%4d\%2.2d /home/abrapdf/%4d" "exit"', [PDFDir, aseRok.Value, aseMesic.Value, aseRok.Value]), SW_SHOWNORMAL);

end;


// ------------------------------------------------------------------------------------------------

procedure TfmMain.Button1Click(Sender: TObject);
var
  source : string;
  i: integer;
  rData: TAArray;
begin

//ShowMessage('demooo_' +  BoolToStr(DesU.existujeVAbreDokladSPrazdnymVs(), true));
//ShowMessage('VRid_' +  DesU.getAbraVatrateId('V�st21'));
//ShowMessage('VRid_' +  DesU.getAbraVatindexId('V�st21'));

Source := 'pakachar:ukulelee';

Copy(Source, 8, MaxInt);

//ShowMessage('1_' +  Copy(Source, Pos(':', Source)+1, MaxInt) + '_');
//ShowMessage('1_' +  Copy(Source, 1, Pos(':', Source)-1) + '_');
//ShowMessage(inttostr(Pos('-', Source)) );


{
Zaplatit := 152.61;

    C := Format('%6.0f', [Zaplatit]);
    //nahrad�me vlnovkou posledn� mezeru, tedy d�me vlnovku p�ed prvn� ��slici
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
  rData := TAArray.Create;
  rData['Title'] := 'Faktura za p�ipojen� k internetu';
  rData['Author'] := 'Dru�stvo Eurosignal';
  ShowMessage ('-'+rData['kuku']+'-');

end;


end.
