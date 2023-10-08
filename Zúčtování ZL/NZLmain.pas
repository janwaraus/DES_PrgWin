// 24.8.2016 zahájení - program pro automatické zúètování zaplacených ZL fakturou
// 23.5.2017 automatické zaplacení vygenerované faktury (zúètování zálohy) - Honza
// 22.6. pøenášení zakázek a obchodních pøípadù do FO ze ZL
// 3.1.2019 rozdìlení Period_Id na Period_Id_ZL a Period_Id_FO - øešení problému rùzných rokù
// 21.8.2022 Abra 22.1 pøešla na UTF8, nutné úpravy
// 5.3.2023  Maily s TLS
// 10.7.2023 pøepsáno pro Delphi 10.x

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
    arbVytvoreni: TAdvOfficeRadioButton;
    arbPrevod: TAdvOfficeRadioButton;
    apbProgress: TAdvProgressBar;
    apnVytvoreni: TAdvPanel;
    acbRada: TAdvComboBox;
    aseRok: TAdvSpinEdit;
    deDatumDokladu: TAdvDateTimePicker;
    cbCast: TCheckBox;
    btVytvoritFO: TButton;
    btNacistZL: TButton;
    apnPrevod: TAdvPanel;
    cbNeprepisovat: TCheckBox;
    btOdeslatNaServer: TButton;
    btSablona: TButton;
    lbxLog: TListBox;
    asgMain: TAdvStringGrid;
    lbPozor1: TLabel;
    deDatumFOxOd: TAdvDateTimePicker;
    deDatumFOxDo: TAdvDateTimePicker;
    chbFO3: TCheckBox;
    chbFO2: TCheckBox;
    chbFO4: TCheckBox;
    btKonec: TButton;
    btNacistFOx: TButton;
    btVytvoritPDF: TButton;
    apnVytvoritPDF: TAdvPanel;
    apnMailTisk: TAdvPanel;
    btOdeslatMailem: TButton;
    btTisk: TButton;
    cbJenNezpracovaneFOx: TCheckBox;
    apnVytvorFOx: TAdvPanel;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure glbVytvoreniClick(Sender: TObject);
    procedure glbPrevodClick(Sender: TObject);
    procedure arbVytvoreniClick(Sender: TObject);
    procedure arbPrevodClick(Sender: TObject);
    procedure asgMainCanEditCell(Sender: TObject; ARow, ACol: Integer; var CanEdit: Boolean);
    procedure asgMainClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure asgMainCanSort(Sender: TObject; ACol: Integer; var DoSort: Boolean);
    procedure asgMainGetAlignment(Sender: TObject; ARow, ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asgMainDblClick(Sender: TObject);
    procedure lbxLogDblClick(Sender: TObject);
    procedure btVytvoritFOClick(Sender: TObject);
    procedure btOdeslatNaServerClick(Sender: TObject);
    procedure btSablonaClick(Sender: TObject);
    procedure btNacistZLClick(Sender: TObject);
    procedure resetAsgMain;
    procedure btKonecClick(Sender: TObject);
    procedure uiActionsBeforeProcessing;
    procedure uiActionsAfterProcessing;
    procedure btNacistFOxClick(Sender: TObject);
    procedure btVytvoritPDFClick(Sender: TObject);
    procedure btOdeslatMailemClick(Sender: TObject);
    procedure btTiskClick(Sender: TObject);


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
  DesU.programInit('Nezúètované ZL');
  DesUtils.appendToFile(DesU.LOG_FILENAME, ''); //vloží prázdný øádek do logu
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
  resetAsgMain;
  //arbVytvoreniClick(nil); // produkcne
  //arbPrevodClick(nil); // pro test

// fajfky v asgMain
  with asgMain do begin
    CheckFalse := '0';
    CheckTrue := '1';
    ColWidths[0] := 28;
    ColWidths[1] := 90;
    ColWidths[2] := 28;
    ColWidths[3] := 64;
    ColWidths[4] := 74;
    ColWidths[5] := 64;
    ColWidths[6] := 64;
    ColWidths[7] := 64;
    ColWidths[8] := 170;
    ColWidths[9] := 64;
    ColWidths[10] := 64;
    ColWidths[11] := 64;
    //ColWidths[12] := 0;
    //ColWidths[13] := 0;
    //ColWidths[14] := 0;
  end;
// pøedvyplnìní formuláøe
  aseRok.Value := YearOf(Date);                            // aktuální rok
  //deDatumDokladu.Date := Date;
  //deDatumDokladu.Date := ISO8601ToDate('2023-04-15'); //9.7.2023
  deDatumDokladu.Date := StrToDateTime('9.7.2023'); //9.7.2023  // 45116
  deDatumDokladu.Date := 45015;
  //fmMain.Zprava(FloatToStr(deDatumDokladu.Date));
  //fmMain.Zprava(TimeToStr(deDatumDokladu.Date));
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
  apnVytvoreni.Visible := True;
  apnPrevod.Visible := False;

  apbProgress.Visible := False;
  lbPozor1.Visible := True;

  lbxLog.Visible := True;
  resetAsgMain;
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
  apnVytvoreni.Visible := False;
  apnPrevod.Visible := True;

  resetAsgMain;
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

procedure TfmMain.resetAsgMain;
begin
  asgMain.ClearNormalCells;
  asgMain.RowCount := 2;
  asgMain.Visible := True;
  //lbxLog.Visible := False;
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
  // CanEdit := (ARow > 0);   // fajfky
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.asgMainClickCell(Sender: TObject; ARow, ACol: Integer);
// klik v prvním øádku podle sloupce buï oznaèuje/odznaèuje všchny checkboxy, nebo spouští tøídìní,
var
  Radek: integer;
begin
  with asgMain do if (ARow = 0) and ((ACol in [0, 2])) then
    if Ints[ACol, 1] = 1 then for Radek := 1 to RowCount-1 do Ints[ACol, Radek] := 0
    else if Ints[ACol, 1] = 0 then for Radek := 1 to RowCount-1 do Ints[ACol, Radek] := 1;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.asgMainCanSort(Sender: TObject; ACol: Integer; var DoSort: Boolean);
// klik v prvním øádku podle sloupce buï oznaèuje/odznaèuje všchny checkboxy, nebo spouští tøídìní,
begin
  DoSort := True;
  if (ACol in [0, 2]) then DoSort := False;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.asgMainGetAlignment(Sender: TObject; ARow, ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
  HAlign := taLeftJustify;
  if (ACol = 0) or (ACol = 2) or (ARow = 0) then HAlign := Classes.taCenter
  else if (ACol in [4..7]) and (ARow > 0) then HAlign := taRightJustify;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.lbxLogDblClick(Sender: TObject);
begin
  asgMain.Visible := True;
  lbxLog.Visible := False;
end;


// ------------------------------------------------------------------------------------------------

procedure TfmMain.btNacistZLClick(Sender: TObject);
begin
  fmMain.uiActionsBeforeProcessing;
  try
    // *** Naplnìní asgMain - naèíst ZL***
    dmCommon.Plneni_asgMain;
  finally
    fmMain.uiActionsAfterProcessing;
  end;
end;


procedure TfmMain.btVytvoritFOClick(Sender: TObject);
// pro zadaný mìsíc se vytvoøí faktury v Abøe, nebo pøevedou do PDF, vytisknou èi rozešlou mailem
begin
  fmMain.uiActionsBeforeProcessing;
  try
    // *** Fakturace ***
    dmVytvoreni.VytvorFO;
    // *** Pøevod do PDF ***
      //if arbPrevod.Checked then dmPrevod.PrevedFO3;


  finally
    fmMain.uiActionsAfterProcessing;
  end;
end;


procedure TfmMain.btVytvoritPDFClick(Sender: TObject);
begin
  fmMain.uiActionsBeforeProcessing;
  try
    // *** Pøevod do PDF ***
    dmPrevod.PrevedFOx;
  finally
    fmMain.uiActionsAfterProcessing;
  end;
end;

procedure TfmMain.btOdeslatMailemClick(Sender: TObject);
begin
  fmMain.uiActionsBeforeProcessing;
  try
    // *** Posílání mailem ***
    dmMail.PosliFOx;
  finally
    fmMain.uiActionsAfterProcessing;
  end;
end;


procedure TfmMain.btTiskClick(Sender: TObject);
begin
  fmMain.uiActionsBeforeProcessing;
  try
    // *** Tisk ***
    dmTisk.TiskniFOx;
  finally
    fmMain.uiActionsAfterProcessing;
  end;
end;


procedure TfmMain.btNacistFOxClick(Sender: TObject);
begin
  fmMain.uiActionsBeforeProcessing;
  try
    // *** Naplnìní asgMain - naèíst FOx***
    dmCommon.Plneni_asgMain;
  finally
    fmMain.uiActionsAfterProcessing;
  end;
end;



procedure TfmMain.uiActionsBeforeProcessing;
begin
  btNacistZL.Enabled := False;
  btVytvoritFO.Enabled := False;
  btNacistFOx.Enabled := False;
  btVytvoritPDF.Enabled := False;
  btOdeslatMailem.Enabled := False;
  btTisk.Enabled := False;

  apbProgress.Position := 0;
  apbProgress.Visible := True;

  asgMain.Visible := True;
  //lbxLog.Visible := False;

  Screen.Cursor := crHourGlass;
  btKonec.Caption := '&Pøerušit';
  Prerusit := False;
  Application.ProcessMessages;
end;

procedure TfmMain.uiActionsAfterProcessing;
begin
  btNacistZL.Enabled := True;
  btVytvoritFO.Enabled := True;
  btNacistFOx.Enabled := True;
  btVytvoritPDF.Enabled := True;
  btOdeslatMailem.Enabled := True;
  btTisk.Enabled := True;

  apbProgress.Position := 0;
  apbProgress.Visible := False;

  Screen.Cursor := crDefault;
  btKonec.Caption := '&Konec';
end;


// ------------------------------------------------------------------------------------------------

procedure TfmMain.btSablonaClick(Sender: TObject);
begin
  DesFastReport.frxReport.DesignReport(True, False);
end;


// ------------------------------------------------------------------------------------------------


procedure TfmMain.btOdeslatNaServerClick(Sender: TObject);
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
  lbxLog.Items.Add(DesU.zalogujZpravu(TextZpravy));
  lbxLog.ItemIndex := lbxLog.Count - 1;
  Application.ProcessMessages;
end;

end.
