// 24.8.2016 zahájení - program pro automatické zúètování zaplacených ZL fakturou
// 23.5.2017 automatické zaplacení vygenerované faktury (zúètování zálohy) - Honza
// 22.6. pøenášení zakázek a obchodních pøípadù do FO ze ZL
// 3.1.2019 rozdìlení Period_Id na Period_Id_ZL a Period_Id_FO - øešení problému rùzných rokù
// 21.8.2022 Abra 22.1 pøešla na UTF8, nutné úpravy
// 5.3.2023  Maily s TLS

unit NZLmain;

interface

uses
  Windows, Messages, Dialogs, SysUtils, Variants, Classes, Graphics, Controls, StdCtrls, ExtCtrls, Forms, Mask, ComObj, ComCtrls,
  AdvObj, AdvPanel, AdvEdit, AdvSpin, AdvDateTimePicker, AdvEdBtn, AdvFileNameEdit, AdvProgressBar, GradientLabel,
  AdvOfficeButtons, AdvGroupBox, Grids, BaseGrid, AdvGrid, DateUtils, Math, Registry,
  DB, ZAbstractConnection, ZConnection, ZAbstractRODataset, ZAbstractDataset, ZDataset,
  frxClass, frxDBSet, frxDesgn, AdvCombo, AdvUtil;

type
  TfmMain = class(TForm)
    apnTop: TAdvPanel;
    apnVyberCinnosti: TAdvPanel;
    glbVytvoreni: TGradientLabel;
    glbPrevod: TGradientLabel;
    glbTisk: TGradientLabel;
    glbMail: TGradientLabel;
    arbVytvoreni: TAdvOfficeRadioButton;
    arbPrevod: TAdvOfficeRadioButton;
    arbTisk: TAdvOfficeRadioButton;
    arbMail: TAdvOfficeRadioButton;
    apbProgress: TAdvProgressBar;
    apnMain: TAdvPanel;
    acbRada: TAdvComboBox;
    aseRok: TAdvSpinEdit;
    deDatumDokladu: TAdvDateTimePicker;
    cbCast: TCheckBox;
    btVytvorit: TButton;
    btKonec: TButton;
    apnPrevod: TAdvPanel;
    cbNeprepisovat: TCheckBox;
    btOdeslat: TButton;
    btSablona: TButton;
    apnTisk: TAdvPanel;
    apnMail: TAdvPanel;
    fePriloha: TAdvFileNameEdit;
    lbxLog: TListBox;
    asgMain: TAdvStringGrid;
    lbPozor1: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure glbVytvoreniClick(Sender: TObject);
    procedure glbPrevodClick(Sender: TObject);
    procedure glbTiskClick(Sender: TObject);
    procedure glbMailClick(Sender: TObject);
    procedure arbVytvoreniClick(Sender: TObject);
    procedure arbPrevodClick(Sender: TObject);
    procedure arbTiskClick(Sender: TObject);
    procedure arbMailClick(Sender: TObject);
    procedure aseRokChange(Sender: TObject);
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
    procedure deDatumDokladuChange(Sender: TObject);


  private
    LogDir : string;

  public
    F: TextFile;
    Prerusit: boolean;
    CastkaZalohy,
    Celkem,
    Saldo,
    DatumDokladu,
    DatumPlneni,
    DatumSplatnosti,
    Zaplatit: double;
    AbraOLE: variant;
    Firm_Id,
    Period_Id_ZL,
    Period_Id_FO,
    FO_Id: string[10];
    PDFDir,
    AbraConnection: ShortString;
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
    CisloZL,
    Platek: AnsiString;
    procedure Zprava(TextZpravy: string);
  end;

var
  fmMain: TfmMain;

implementation

uses NZLCommon, NZLfaktura, NZLtisk, NZLprevod, NZLmail, DesUtils, DesFastReports;

{$R *.dfm}

// ------------------------------------------------------------------------------------------------
// funkce prvkù formuláøe fmMain
// ------------------------------------------------------------------------------------------------

procedure TfmMain.FormCreate(Sender: TObject);
begin
  Prerusit := False;
end;

procedure TfmMain.FormShow(Sender: TObject);

begin
  LogDir := DesU.PROGRAM_PATH +  'Logy\Nezúètované ZL\';
  if not DirectoryExists(LogDir) then Forcedirectories(LogDir);

  DesUtils.appendToFile(LogDir + FormatDateTime('yyyy.mm".log"', Date), ''); //vloží prázdný øádek do logu
  fmMain.Zprava('Start programu "Nezúètované zálohové listy".');

  Prerusit := True;                   // pøíznak startu

  acbRada.Clear;
  acbRada.Items.Add('%');
  with DesU.qrAbra do begin
    SQL.Text := 'SELECT Code FROM DocQueues'                 // øady ZL
    + ' WHERE DocumentType = ''10'''
    + ' AND Hidden = ''N'''
    + ' ORDER BY Code';
    Open;
    while not EOF do begin
      acbRada.Items.Add(FieldByName('Code').AsString);
      Next;
    end;
    Close;
  end;
  acbRada.ItemIndex := 0;
  fmMain.Zprava('OK');
  aseRokChange(nil);
  arbVytvoreniClick(nil);

// fajfky v asgMain
  with asgMain do begin
    CheckFalse := '0';
    CheckTrue := '1';
    ColWidths[0] := 28;
    ColWidths[1] := 64;
    ColWidths[2] := 74;
    ColWidths[3] := 64;
    ColWidths[4] := 64;
    ColWidths[5] := 64;
    ColWidths[6] := 170;
    ColWidths[7] := 64;
    //ColWidths[8] := 0;
    //ColWidths[9] := 0;
    //ColWidths[10] := 0;
  end;
// pøedvyplnìní formuláøe
  aseRok.Value := YearOf(Date);                            // aktuální rok
  deDatumDokladu.Date := Trunc(Date);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.FormActivate(Sender: TObject);
begin
  Prerusit := False;
end;


// ------------------------------------------------------------------------------------------------

procedure TfmMain.arbVytvoreniClick(Sender: TObject);
begin
  asgMain.Cells[0, 0] := 'FO';
  glbVytvoreni.Color := clWhite;
  glbVytvoreni.ColorTo := clMenu;
  glbPrevod.Color := clSilver;
  glbPrevod.ColorTo := clGray;
  apnPrevod.Visible := False;
  glbTisk.Color := clSilver;
  glbTisk.ColorTo := clGray;
  apnTisk.Visible := False;
  glbMail.Color := clSilver;
  glbMail.ColorTo := clGray;
  apnMail.Visible := False;
  apbProgress.Visible := False;
  lbPozor1.Visible := True;
  acbRada.Visible := True;
  aseRok.Visible := True;
  cbCast.Visible := True;
  lbxLog.Visible := True;
  aseRokChange(Self);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.arbPrevodClick(Sender: TObject);
begin
  asgMain.Cells[0, 0] := 'PDF';
  glbPrevod.Color := clWhite;
  glbPrevod.ColorTo := clMenu;
  apnPrevod.Visible := True;
  glbVytvoreni.Color := clSilver;
  glbVytvoreni.ColorTo := clGray;
  glbTisk.Color := clSilver;
  glbTisk.ColorTo := clGray;
  apnTisk.Visible := False;
  glbMail.Color := clSilver;
  glbMail.ColorTo := clGray;
  apnMail.Visible := False;
  acbRada.Visible := False;
  aseRok.Visible := False;
  cbCast.Visible := False;
  aseRokChange(Self);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.arbTiskClick(Sender: TObject);
begin
  asgMain.Cells[0, 0] := 'tisk';
  glbTisk.Color := clWhite;
  glbTisk.ColorTo := clMenu;
  apnTisk.Visible := True;
  glbMail.Color := clSilver;
  glbMail.ColorTo := clGray;
  apnMail.Visible := False;
  glbVytvoreni.Color := clSilver;
  glbVytvoreni.ColorTo := clGray;
  glbPrevod.Color := clSilver;
  glbPrevod.ColorTo := clGray;
  apnPrevod.Visible := False;
  acbRada.Visible := False;
  aseRok.Visible := False;
  cbCast.Visible := False;
  aseRokChange(Self);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.arbMailClick(Sender: TObject);
begin
  asgMain.Cells[0, 0] := 'mail';
  glbMail.Color := clWhite;
  glbMail.ColorTo := clMenu;
  apnMail.Visible := True;
  glbVytvoreni.Color := clSilver;
  glbVytvoreni.ColorTo := clGray;
  glbPrevod.Color := clSilver;
  glbPrevod.ColorTo := clGray;
  apnPrevod.Visible := False;
  glbTisk.Color := clSilver;
  glbTisk.ColorTo := clGray;
  apnTisk.Visible := False;
  acbRada.Visible := False;
  aseRok.Visible := False;
  cbCast.Visible := False;
  aseRokChange(Self);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.glbVytvoreniClick(Sender: TObject);
begin
  arbVytvoreni.Checked := True;
  arbVytvoreniClick(nil);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.glbPrevodClick(Sender: TObject);
begin
  arbPrevod.Checked := True;
  arbPrevodClick(nil);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.glbTiskClick(Sender: TObject);
begin
  arbTisk.Checked := True;
  arbTiskClick(nil);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.glbMailClick(Sender: TObject);
begin
  arbMail.Checked := True;
  arbMailClick(nil);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.aseRokChange(Sender: TObject);
// pøi zmìnì roku nastaví nové deDatumDokladu
begin
  asgMain.ClearNormalCells;
  asgMain.RowCount := 2;
  btVytvorit.Caption := '&Naèíst';
  asgMain.Visible := True;
  lbxLog.Visible := False;
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
  CanEdit := (ARow > 0);   // fajfky
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.asgMainClickCell(Sender: TObject; ARow, ACol: Integer);
// klik v prvním øádku podle sloupce buï oznaèuje/odznaèuje všchny checkboxy, nebo spouští tøídìní,
var
  Radek: integer;
begin
  with asgMain do if (ARow = 0) and (ACol = 0) then
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
  else if (ACol in [2..5]) and (ARow > 0) then HAlign := taRightJustify;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.lbxLogDblClick(Sender: TObject);
begin
  asgMain.Visible := True;
  lbxLog.Visible := False;
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
      if arbVytvoreni.Checked then dmVytvoreni.VytvorFO;
// *** Pøevod do PDF ***
      if arbPrevod.Checked then dmPrevod.PrevedFO3;
// *** Tisk ***
      if arbTisk.Checked then dmTisk.TiskniFO3;
// *** Posílání mailem ***
      if arbMail.Checked then dmMail.PosliFO3;
    end;
  finally
    btKonec.Caption := '&Konec';
    btVytvorit.Enabled := True;
  end;
end;

procedure TfmMain.deDatumDokladuChange(Sender: TObject);
begin
  self.aseRokChange(nil);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.btSablonaClick(Sender: TObject);
begin
  DesFastReport.frxReport.DesignReport(True, False);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.btOdeslatClick(Sender: TObject);
// odeslání faktur pøevedených do PDF na vzdálený server
begin
  //WinExec(PChar(Format('WinSCP.com /command "option batch abort" "option confirm off" "open AbraPDF" "synchronize remote '
  // + PDFDir + '\%4d\%2.2d /home/abrapdf/%4d" "exit"',
  //  [YearOf(deDatumDokladu.Date), MonthOf(deDatumDokladu.Date), YearOf(deDatumDokladu.Date)])), SW_SHOWNORMAL);

  DesU.syncAbraPdfToRemoteServer(YearOf(deDatumDokladu.Date), MonthOf(deDatumDokladu.Date));

end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.btKonecClick(Sender: TObject);
begin
  if btKonec.Caption = '&Pøerušit' then begin
    Prerusit := True;
    fmMain.Zprava('Pøerušeno uživatelem.');
    btKonec.Caption := '&Konec';
  end else begin
    fmMain.Zprava('Konec programu "Zúètování ZL".');
    Close;
  end;
end;



// ------------------------------------------------------------------------------------------------
// funkce nespojené pøímo s konkrétním prvkem formuláøe
// ------------------------------------------------------------------------------------------------


procedure TfmMain.Zprava(TextZpravy: string);
// do listboxu a logfile uloží èas a text zprávy
begin
  TextZpravy := FormatDateTime('dd.mm.yy hh:nn:ss  ', Now) + TextZpravy;
  with fmMain do begin
    lbxLog.Items.Add(TextZpravy);
    lbxLog.ItemIndex := lbxLog.Count - 1;
    Application.ProcessMessages;
    DesUtils.appendToFile(LogDir + FormatDateTime('yyyy.mm".log"', Date),
      Format('(%s - %s) ', [Trim(getWindowsCompName), Trim(getWindowsUserName)]) + TextZpravy);
  end;
end;

end.
