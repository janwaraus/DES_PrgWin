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
  Grids, BaseGrid, AdvGrid, //pCore2D, pBarcode2D,
  IniFiles, DateUtils, Math,
  DB, ZAbstractConnection, ZConnection, ZAbstractRODataset, ZAbstractDataset, ZDataset,

  frxClass, frxDBSet, frxDesgn, AdvUtil;

type
  TfmMain = class(TForm)
    qrMain: TZQuery;
    qrSmlouva: TZQuery;
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
    rbVyberPodleFaktury: TRadioButton;
    rbVyberPodleVS: TRadioButton;
    btnTest: TButton;
    lblPocetKNacteni: TLabel;
    lblPocetOdChanges: TLabel;
    chbFakturyKNacteni: TCheckBox;
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
    procedure btVytvoritClick(Sender: TObject);
    procedure btOdeslatClick(Sender: TObject);
    procedure btSablonaClick(Sender: TObject);
    procedure btKonecClick(Sender: TObject);
    procedure rbVyberPodleVSClick(Sender: TObject);
    procedure rbVyberPodleFakturyClick(Sender: TObject);
    procedure rbFakturaceClick(Sender: TObject);
    procedure rbPrevodClick(Sender: TObject);
    procedure rbTiskClick(Sender: TObject);
    procedure rbMailClick(Sender: TObject);
    procedure btnTestClick(Sender: TObject);
    procedure chbFakturyKNacteniClick(Sender: TObject);

  private
    LogDir : string;
    function getSqlTextFakturace() : string;
    function getSqlTextPrevodMailTisk(FiltrovatPodle : string; ApNavic : boolean) : string;

  public
    isDebugMode,
    Prerusit: boolean;
    pocetOdChanges : integer;
    aedOdTextOld, aedDoTextOld : string;
    main_invoiceDocQueueCode : string;
    main_zakladniSazbaDPH : integer;
    procedure Zprava(TextZpravy: string);
    procedure PlneniAsgMain;

  end;



var
  fmMain: TfmMain;

implementation

uses DesUtils, AbraEntities, FIfaktura, FIPrevod, FITisk, FIMail,
AArray, DesFastReports;

{$R *.dfm}

// ------------------------------------------------------------------------------------------------
// funkce prvk� formul��e fmMain
// ------------------------------------------------------------------------------------------------


procedure TfmMain.FormShow(Sender: TObject);
var
  abraVatIndex : TAbraVatIndex;
  abraDrcArticle : TAbraDrcArticle;
  dbVoipConnectResult : TDesResult;

begin

  main_invoiceDocQueueCode := 'FO1'; //k�d �ady faktur, tento program vystavuje pouze do t�to �ady
  main_zakladniSazbaDPH  := AbraEnt.getVatIndex('Code=V�st21').Tariff;

  //adres�� pro logy
  LogDir := DesU.PROGRAM_PATH + 'Logy\M�s��n� fakturace\';
  if not DirectoryExists(LogDir) then Forcedirectories(LogDir);
  //globalAA['LogFileName'] := LogDir + FormatDateTime('yyyy.mm".log"', Date);

  DesUtils.appendToFile(LogDir + FormatDateTime('yyyy.mm".log"', Date), ''); //vlo�� pr�zdn� ��dek do logu
  Zprava('Start programu "M�s��n� fakturace".');

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

  rbFakturaceClick(nil);

  {
  if cbSVoIP.Enabled then begin
    dbVoipConnectResult := DesU.connectDbVoip;
    Zprava(dbVoipConnectResult.Messg);
  end;
  }

  // pro test, TODO smazat
  aedOd.Text := '10200555';
  aedDo.Text := '10205555';

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
    if btVytvorit.Caption = '&Na��st' then fmMain.PlneniAsgMain
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
    Zprava('P�eru�eno u�ivatelem.');
    btKonec.Caption := '&Konec';
  end else begin
    Zprava('Konec programu "M�s��n� fakturace".');
    Close;
  end;
end;

procedure TfmMain.btSablonaClick(Sender: TObject);
begin
  DesFastReport.frxReport.DesignReport(True, False);
end;

procedure TfmMain.btOdeslatClick(Sender: TObject);
// odesl�n� faktur p�eveden�ch do PDF na vzd�len� server
begin

  //WinExec(PChar(Format('WinSCP.com /command "option batch abort" "option confirm off" "open AbraPDF" "synchronize remote '
  // + '%s\%4d\%2.2d /home/abrapdf/%4d" "exit"', [PDFDir, aseRok.Value, aseMesic.Value, aseRok.Value])), SW_SHOWNORMAL);

  RunCMD (Format('WinSCP.com /command "option batch abort" "option confirm off" "open AbraPDF" "synchronize remote '
   + '%s%4d\%2.2d /home/abrapdf/%4d" "exit"', [DesU.PDF_PATH, aseRok.Value, aseMesic.Value, aseRok.Value]), SW_SHOWNORMAL);

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
  rbVyberPodleVS.Checked := True;
  rbVyberPodleFaktury.Enabled := False;
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
  rbVyberPodleFaktury.Enabled := True;
  rbVyberPodleFaktury.Checked := True;
  apnVyberPodle.Visible := True;
  aseRokChange(Self);


  {* HW testovaci nastaveni TODO odstranit*}
  aseMesic.Value := 2;
  aedOd.Text := '5000';
  aedDo.Text := '5005';

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
  rbVyberPodleFaktury.Enabled := True;
  rbVyberPodleFaktury.Checked := True;
  apnVyberPodle.Visible := True;
  aseRokChange(Self);

  {* HW testovaci nastaveni TODO odstranit *}
  aseMesic.Value := 2;
  aedOd.Text := '6000';
  aedDo.Text := '6250';


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
  rbVyberPodleFaktury.Enabled := True;
  rbVyberPodleFaktury.Checked := True;
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
// p�i zm�n� (nejen) roku nastav� nov� deDatumDokladu, deDatumPlneni, aedSplatnost, aedOd a aedDo
begin
  aedOd.Clear; // pro automatick� refresh tla��tka Na��st (a po�tu faktur k na�ten�)
  //aedDo.Clear;
  asgMain.ClearNormalCells;
  asgMain.RowCount := 2;
  btVytvorit.Caption := '&Na��st';

// datum fakturace i datum pln�n� je posledn� den v m�s�ci
  deDatumDokladu.Date := EndOfAMonth(aseRok.Value, aseMesic.Value);
  deDatumPlneni.Date := deDatumDokladu.Date;
  aedSplatnost.Text := '10';


  if rbFakturace.Checked then begin
    // *** v�b�r pro Fakturaci, pouze podle VS
    with DesU.qrZakos do try
      Screen.Cursor := crSQLWait;

      // prvn� a posledn� VS
      SQL.Text := 'CALL get_monthly_invoicing_minmaxvs('
        + Ap + FormatDateTime('yyyy-mm-dd', StartOfTheMonth(deDatumPlneni.Date)) + ApC
        + Ap + FormatDateTime('yyyy-mm-dd', deDatumPlneni.Date) + ApZ;

      Open;
      aedOd.Text := FieldByName('min_vs').AsString;
      aedDo.Text := FieldByName('max_vs').AsString;
      Close;
    finally
      Screen.Cursor := crDefault;
    end;
  end else begin
    // *** v�b�r pro P�evod, Tisk, Mail
    DesU.dbAbra.Reconnect;
    with DesU.qrAbra do begin
      // rozp�t� ��sel FO1 v m�s�ci
      if rbVyberPodleVS.Checked then
        SQL.Text := 'SELECT MIN(VarSymbol), MAX(VarSymbol) ';

      if rbVyberPodleFaktury.Checked then
        SQL.Text := 'SELECT MIN(OrdNumber), MAX(OrdNumber) ';

      SQL.Text := SQL.Text + 'FROM IssuedInvoices'
      + ' WHERE DocQueue_ID = ' + Ap + AbraEnt.getDocQueue('Code=FO1').ID + Ap
      + ' AND VATDate$DATE >= ' + FloatToStr(Trunc(StartOfAMonth(aseRok.Value, aseMesic.Value)))
      + ' AND VATDate$DATE <= ' + FloatToStr(Trunc(EndOfAMonth(aseRok.Value, aseMesic.Value)));

      Open;
      if RecordCount > 0 then begin
        aedOd.Text := Fields[0].AsString;
        aedDo.Text := Fields[1].AsString;
      end else begin
        aedOd.Clear;
        aedDo.Clear;
      end;
      Close;
    end;
  end;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.aedOdChange(Sender: TObject);
var
  pocetFakturKNacteni : integer;
begin
  Inc(pocetOdChanges);

  if btVytvorit.Caption <> '&Na��st' then begin
    asgMain.ClearNormalCells;
    asgMain.RowCount := 2;
    btVytvorit.Caption := '&Na��st';
  end;


  if chbFakturyKNacteni.Checked then begin

    if rbFakturace.Checked then begin
      DesU.qrZakosOC.SQL.Text := getSqlTextFakturace();
      DesU.qrZakosOC.Open;
      pocetFakturKNacteni := DesU.qrZakosOC.RecordCount;
      DesU.qrZakosOC.Close;
    end else begin
      if rbVyberPodleVS.Checked then
        DesU.qrAbraOC.SQL.Text := getSqlTextPrevodMailTisk('VarSymbol', true);
      if rbVyberPodleFaktury.Checked then
        DesU.qrAbraOC.SQL.Text := getSqlTextPrevodMailTisk('OrdNumber', false);
      DesU.qrAbraOC.Open;
      pocetFakturKNacteni := DesU.qrAbraOC.RecordCount;
      DesU.qrAbraOC.Close;
    end;

    lblPocetKNacteni.Caption := IntToStr(pocetFakturKNacteni);

  end else
    lblPocetKNacteni.Caption := 'nezji��ov�no';


  if isDebugMode then begin
    lblPocetOdChanges.Caption := 'Po�et aedOdChange event�: ' + IntToStr(pocetOdChanges);
    lblPocetOdChanges.Visible := True;
    aedOdTextOld := aedOd.Text;
    aedDoTextOld := aedDo.Text;
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

procedure TfmMain.chbFakturyKNacteniClick(Sender: TObject);
begin
  aedOdChange(nil);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.rbVyberPodleVSClick(Sender: TObject);
begin
  aseRokChange(nil);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.rbVyberPodleFakturyClick(Sender: TObject);
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
// funkce nespojen� p��mo s konkr�tn�m prvkem formul��e
// ------------------------------------------------------------------------------------------------


procedure TfmMain.Zprava(TextZpravy: string);
// do listboxu a logfile ulo�� �as a text zpr�vy
begin
  TextZpravy := FormatDateTime('dd.mm.yy hh:nn:ss  ', Now) + TextZpravy;
  lbxLog.Items.Add(TextZpravy);
  lbxLog.ItemIndex := lbxLog.Count - 1;
  Application.ProcessMessages;
  DesUtils.appendToFile(LogDir + FormatDateTime('yyyy.mm".log"', Date), TextZpravy);
end;

// ------------------------------------------------------------------------------------------------

function TfmMain.getSqlTextFakturace() : string;
var
  SqlProcedureName : string;
begin
  if cbBezVoIP.Checked and not cbSVoIP.Checked then
    SqlProcedureName := 'get_monthly_invoicing_cu_by_vsrange_nonvoip'
  else if cbSVoIP.Checked and not cbBezVoIP.Checked then
    SqlProcedureName := 'get_monthly_invoicing_cu_by_vsrange_voip'
  else
    SqlProcedureName := 'get_monthly_invoicing_cu_by_vsrange_all';

  Result := 'CALL ' + SqlProcedureName + '('
    + Ap + FormatDateTime('yyyy-mm-dd', StartOfTheMonth(deDatumPlneni.Date)) + ApC
    + Ap + FormatDateTime('yyyy-mm-dd', deDatumPlneni.Date) + ApC
    + Ap + aedOd.Text + ApC
    + Ap + aedDo.Text + ApZ;
end;


function TfmMain.getSqlTextPrevodMailTisk(FiltrovatPodle : string; ApNavic : boolean) : string;
var
  ApStr : string;
begin
  if ApNavic then
    ApStr := Ap
  else
    ApStr := '';

  // faktura(y) v Ab�e v m�s�ci aseMesic
  Result := 'SELECT II.ID, F.Name as FirmName, F.Code as Abrakod, II.OrdNumber, II.VarSymbol, II.Amount '
  + 'FROM IssuedInvoices II, Firms F'
  + ' WHERE II.Firm_ID = F.ID'
  + ' AND ' + FiltrovatPodle + ' >= ' + ApStr + aedOd.Text + ApStr
  + ' AND ' + FiltrovatPodle + ' <= ' + ApStr + aedDo.Text + ApStr
  + ' AND DocQueue_ID = ' + Ap + AbraEnt.getDocQueue('Code=FO1').id + Ap
  + ' AND VATDate$DATE >= ' + FloatToStr(Trunc(StartOfAMonth(aseRok.Value, aseMesic.Value)))  // aby se odfiltrovaly fa z jin�ch m�s�c�, pokud by se n�jak dostaly do �ady
  + ' AND VATDate$DATE <= ' + FloatToStr(Trunc(EndOfAMonth(aseRok.Value, aseMesic.Value))); // aby se odfiltrovaly fa z jin�ch m�s�c�, pokud by se n�jak dostaly do �ady

  if cbBezVoIP.Checked and not cbSVoIP.Checked then
    Result := Result + ' AND LOWER(II.Description) NOT LIKE ''%voip%''';
  if cbSVoIP.Checked and not cbBezVoIP.Checked then
    Result := Result + ' AND LOWER(II.Description) LIKE ''%voip%''';

  if rbMail.Checked then Result := Result + ' AND F.Firm_ID IS NULL';  // TODO prov��it
  Result := Result + ' ORDER BY VarSymbol';
end;

procedure TfmMain.PlneniAsgMain;
var
  Radek: integer;
  VarSymbol : string[10];
  SQLStr: string;
begin
  try
    apnPrevod.Visible := False;
    apnTisk.Visible := False;
    apnMail.Visible := False;
    Screen.Cursor := crHourGlass;

    with asgMain do try
      ClearNormalCells;
      RowCount := 2;

      // ********************//
      // ***  Fakturace  *** //
      // ********************//
      if rbFakturace.Checked then begin          //v�b�r z�kazn�k�/smluv k fakturaci
        // kontrola m�s�ce a roku fakturace
        if (aseRok.Value * 12 + aseMesic.Value > YearOf(Date) * 12 + MonthOf(Date) + 1)     // v�t�� ne� p��t� m�s�c, �i men��
         or (aseRok.Value * 12 + aseMesic.Value < YearOf(Date) * 12 + MonthOf(Date) - 1) then begin       // ne� minul� m�s�c
          SQLStr := Format('Opravdu fakturovat %d. m�s�c roku %d ?', [aseMesic.Value, aseRok.Value]);
          if Application.MessageBox(PChar(SQLStr), 'Pozor', MB_ICONQUESTION + MB_YESNO + MB_DEFBUTTON1) = IDNO then begin
            btVytvorit.Enabled := True;
            btKonec.Caption := '&Konec';
            Exit;
          end;
        end;

        Zprava(Format('Na�ten� z�kazn�k� k fakturaci na obdob� %s.%s od VS %s do %s.', [aseMesic.Text, aseRok.Text, aedOd.Text, aedDo.Text]));

        if cbBezVoIP.Checked then Zprava('      - z�kazn�ci bez VoIP');
        if cbSVoIP.Checked then Zprava('      - z�kazn�ci s VoIP');

        Radek := 0;
        apbProgress.Position := 0;
        apbProgress.Visible := True;

        DesU.qrZakos.SQL.Text := getSqlTextFakturace();

        DesU.qrZakos.Open;
        while not DesU.qrZakos.EOF do begin
          VarSymbol := DesU.qrZakos.FieldByName('cu_variable_symbol').AsString;
          apbProgress.Position := Round(100 * DesU.qrZakos.RecNo / DesU.qrZakos.RecordCount);
          Application.ProcessMessages;

          if Prerusit then begin
            Prerusit := False;
            apbProgress.Position := 0;
            apbProgress.Visible := False;
            btVytvorit.Enabled := True;
            btKonec.Caption := '&Konec';
            Break;  // konec while a tedy skok skoro na konec procedury a konec na��t�n�
          end;

          with DesU.qrAbra do begin
            // ulo�en� do asgMain
            Inc(Radek);
            RowCount := Radek + 1;
            AddCheckBox(0, Radek, True, True);
            Ints[0, Radek] := 1;                                                   // fajfka
            Cells[1, Radek] := VarSymbol;                                          // VS
            Cells[2, Radek] := '';                                                 // faktura
            Cells[3, Radek] := '';                                                 // ��stka
            Cells[4, Radek] := DesU.qrZakos.FieldByName('cu_abra_code').AsString;             // jm�no
            Cells[5, Radek] := DesU.qrZakos.FieldByName('cu_postal_mail').AsString;                // mail
            Ints[6, Radek] := DesU.qrZakos.FieldByName('cu_disable_mailings').AsInteger;             // reklama
            Application.ProcessMessages;
          end;  // with DesU.qrAbra
          DesU.qrZakos.Next;
        end;  // while not EOF
        DesU.qrZakos.Close;

      end
      else // if rbFakturace.Checked

      // *****************************//
      // ***  P�evod, Tisk, Mail  *** //
      // *****************************//
      with DesU.qrAbra do
      begin
        Radek := 0;
        apbProgress.Position := 0;
        apbProgress.Visible := True;

        if rbVyberPodleVS.Checked then begin
          SQL.Text := getSqlTextPrevodMailTisk('VarSymbol', true);
          Zprava(Format('Na�ten� faktur s VS od %s do %s.', [aedOd.Text, aedDo.Text]));
        end;
        if rbVyberPodleFaktury.Checked then begin
          SQL.Text := getSqlTextPrevodMailTisk('OrdNumber', false);
          Zprava(Format('Na�ten� faktur s ��sly od %s do %s.', [aedOd.Text, aedDo.Text]));
        end;

        Open;
        while not EOF do begin
          VarSymbol := FieldByName('VarSymbol').AsString;
          apbProgress.Position := Round(100 * RecNo / RecordCount);
          Application.ProcessMessages;
          if Prerusit then begin
            Prerusit := False;
            apbProgress.Position := 0;
            apbProgress.Visible := False;
            btVytvorit.Enabled := True;
            btKonec.Caption := '&Konec';
            Break;
          end;
          with DesU.qrZakos do begin
            Close;

            SQLStr := 'SELECT postal_mail, disable_mailings FROM customers'
            + ' WHERE variable_symbol = ' + VarSymbol;

            if rbMail.Checked then SQLStr := SQLStr + ' AND invoice_sending_method_id = 9'; // hodnota v codebooks "Mailem"
            if rbTisk.Checked then begin
              if rbBezSlozenky.Checked then SQLStr := SQLStr + ' AND invoice_sending_method_id = 10' // hodnota v codebooks "Po�tou"
              else if rbSeSlozenkou.Checked then SQLStr := SQLStr + ' AND invoice_sending_method_id = 11' // hodnota v codebooks "Se slo�enkou"
              else if rbKuryr.Checked then SQLStr := SQLStr + ' AND invoice_sending_method_id = 12'; // hodnota v codebooks "Kur�r"
            end;

            SQL.Text := SQLStr;
            Open;
            if RecordCount > 0 then begin
              // ulo�en� do asgMain
              Inc(Radek);
              RowCount := Radek + 1;
              AddCheckBox(0, Radek, True, True);
              Ints[0, Radek] := 1;                                                   // fajfka tisk - mail
              Cells[1, Radek] := VarSymbol;                                          // smlouva
              Cells[2, Radek] := Format('%5.5d', [DesU.qrAbra.FieldByName('OrdNumber').AsInteger]);     // faktura
              Floats[3, Radek] := DesU.qrAbra.FieldByName('Amount').AsFloat;;             // ��stka
              Cells[4, Radek] := DesU.qrAbra.FieldByName('FirmName').AsString;                        // jm�no
              Cells[5, Radek] := FieldByName('postal_mail').AsString;                       // mail
              Ints[6, Radek] := FieldByName('disable_mailings').AsInteger;                    // reklama
              Cells[7, Radek] := DesU.qrAbra.FieldByName('ID').AsString;             // ID faktury
            end;
          end; //with DesU.qrZakos
          Application.ProcessMessages;
          Next;
        end;  // while not EOF
      end;  // with DesU.qrAbra

      { vyreseno ORDER BY v selectu
      if not rbFakturace.Checked then begin // TODO zeptat se, co a proc se sortuje, proc fakturace ne a ostatni ano
        SortSettings.Column := 2;
        SortSettings.Full := True;
        //SortSettings.AutoFormat := False;
        SortSettings.Direction := sdAscending;
        QSort;
      end;
      }

      Zprava(Format('Po�et faktur: %d', [RowCount-1]));
      if rbFakturace.Checked then btVytvorit.Caption := '&Vytvo�it'
      else if rbPrevod.Checked then btVytvorit.Caption := '&P�ev�st'
      else if rbTisk.Checked then btVytvorit.Caption := '&Vytisknout'
      else if rbMail.Checked then btVytvorit.Caption := '&Odeslat';
    except on E: Exception do
      Zprava('Neo�et�en� chyba: ' + E.Message);
    end;  // with DesU.qrZakos
  finally
    DesU.qrZakos.Close;
    apbProgress.Position := 0;
    apbProgress.Visible := False;
    if rbPrevod.Checked then apnPrevod.Visible := True;
    if rbTisk.Checked then apnTisk.Visible := True;
    if rbMail.Checked then apnMail.Visible := True;
    Screen.Cursor := crDefault;
    btVytvorit.Enabled := True;
    btVytvorit.SetFocus;
  end;
end;



// ------------------------------------------------------------------------------------------------
// testovac� funkce
// ------------------------------------------------------------------------------------------------

procedure TfmMain.btnTestClick(Sender: TObject);
var
  source : string;
  i: integer;
  rData: TAArray;
  temps : string;
  vysledek : TDesResult;
begin

  //vysledek := dmPrevod.fakturaPrevod('1K6P200101', true);  // pokud za�krtnuto, p�ev�d�me fa do PDF
  //  '1K6P200101' '4K6P200101' '7K6P200101'
  //Zprava(Format('%s', [vysledek.Messg]));


  vysledek := dmTisk.FakturaTisk('6IOQ200101', 'FOseSlozenkou.fr3');    // 7X8P200101
  Zprava(Format('%s', [vysledek.Messg]));


{


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

end;


end.
