// od 19.12.2014 - Zálohové listy na pøipojení se generují v Abøe podle databáze iQuest. Podle Mìsíèní fakturace.
// 27.11.17 pro tarify sítí Misauer a Harna nejsou slevy
// 28.6.18 pøidány zakázky a obchodní pøípady pro statistiky ÈTÚ
// 16.9. pro MySQL set autocommit = 1
// 30.9. zrušeny testovací view
// 21.8.2022 úpravy kvùli Abøe 22.1 s UTF8
// 5.3.23 Maily s TLS

unit ZLmain;

interface

uses
  Windows, Messages, Dialogs, SysUtils, Variants, Classes, Graphics, Controls, StdCtrls, ExtCtrls, Forms, Mask, ComObj, ComCtrls,
  AdvObj, AdvPanel, AdvEdit, AdvSpin, AdvDateTimePicker, AdvEdBtn, AdvFileNameEdit, AdvProgressBar, GradientLabel,
  AdvOfficeButtons, AdvGroupBox, Grids, BaseGrid, AdvGrid,DateUtils, Math, Registry,
  DB, ZAbstractConnection, ZConnection, ZAbstractRODataset, ZAbstractDataset, ZDataset,
  frxClass, frxDBSet, frxDesgn, AdvUtil;

type
  TfmMain = class(TForm)
    dbMain: TZConnection;
    qrMain: TZQuery;
    qrSmlouva: TZQuery;
    dbAbra: TZConnection;
    qrAbra: TZQuery;
    qrAdresa: TZQuery;
    qrRadky: TZQuery;
    frxReport: TfrxReport;
    frxDesigner: TfrxDesigner;
    fdsRadky: TfrxDBDataset;
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
    argMesic: TAdvOfficeRadioGroup;
    aseRok: TAdvSpinEdit;
    deDatumDokladu: TAdvDateTimePicker;
    deDatumSplatnosti: TAdvDateTimePicker;
    aedOd: TAdvEdit;
    aedDo: TAdvEdit;
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
    qrTmp: TZQuery;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure dbMainAfterConnect(Sender: TObject);
    procedure dbAbraAfterConnect(Sender: TObject);
    procedure glbVytvoreniClick(Sender: TObject);
    procedure glbPrevodClick(Sender: TObject);
    procedure glbTiskClick(Sender: TObject);
    procedure glbMailClick(Sender: TObject);
    procedure arbVytvoreniClick(Sender: TObject);
    procedure arbPrevodClick(Sender: TObject);
    procedure arbTiskClick(Sender: TObject);
    procedure arbMailClick(Sender: TObject);
    procedure argMesicClick(Sender: TObject);
    procedure aseRokChange(Sender: TObject);
    procedure deDatumDokladuChange(Sender: TObject);
    procedure deDatumSplatnostiChange(Sender: TObject);
    procedure aedOdChange(Sender: TObject);
    procedure aedDoChange(Sender: TObject);
    procedure aedOdKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure aedOdExit(Sender: TObject);
    procedure aedDoKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure aedDoExit(Sender: TObject);
    procedure asgMainCanEditCell(Sender: TObject; ARow, ACol: Integer; var CanEdit: Boolean);
    procedure asgMainClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure asgMainCanSort(Sender: TObject; ACol: Integer; var DoSort: Boolean);
    procedure asgMainGetAlignment(Sender: TObject; ARow, ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asgMainDblClick(Sender: TObject);
    procedure lbxLogDblClick(Sender: TObject);
    procedure frxReportBeginDoc(Sender: TObject);
    procedure frxReportEndDoc(Sender: TObject);
    procedure frxReportGetValue(const ParName: string; var ParValue: Variant);
    procedure btVytvoritClick(Sender: TObject);
    procedure btOdeslatClick(Sender: TObject);
    procedure btSablonaClick(Sender: TObject);
    procedure btKonecClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);

  private
    LogDir : string;

  public
    Prerusit: boolean;

    //ID,
    //Period_Idsmazat,
    //DocQueue_Id: string[10];
    //PDFDir,
    BBmax,
    BillingView,
    ZLView: ShortString;


    procedure Zprava(TextZpravy: string);
  end;

var
  fmMain: TfmMain;

implementation

uses ZLCommon, ZLvytvoreni, ZLprevod, ZLtisk, ZLmail, DesUtils, DesFastReports;

{$R *.dfm}



// ------------------------------------------------------------------------------------------------

procedure TfmMain.FormCreate(Sender: TObject);
begin
  Prerusit := False;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.FormShow(Sender: TObject);
// inicializace, získání pøipojovacích informací pro databáze, pøipojení
var
  Month: integer;
begin
// jména pro view jsou unikátní, aby program nebyl omezen na jednu instanci
  BBmax := FormatDateTime('BByymmddhhnnss', Now);
  BillingView := FormatDateTime('BVyymmddhhnnss', Now);
  ZLView := FormatDateTime('"ZV"yymmddhhnnss', Now);
// jméno FI.ini

  LogDir := DesU.PROGRAM_PATH + 'Logy\Zálohové listy\';
  if not DirectoryExists(LogDir) then Forcedirectories(LogDir);

  DesUtils.appendToFile(LogDir + FormatDateTime('yyyy.mm".log"', Date), ''); //vloží prázdný øádek do logu
  fmMain.Zprava('Start programu "Zálohové listy".');

  Prerusit := True; // pøíznak startu

  // fajfky v asgMain
  asgMain.CheckFalse := '0';
  asgMain.CheckTrue := '1';

  // pøedvyplnìní formuláøe
  aseRok.Value := YearOf(Date);                            // aktuální rok
  Month := MonthOf(Date);
  if Month = 1 then begin
    Month := 13;
    aseRok.Value := aseRok.Value - 1;
  end;
  argMesic.ItemIndex := Floor((Month+1)/3) - 1;             // tolerance -2 / +1 mìsíc
  deDatumDokladu.Date := Trunc(StrToDate(Format('15.%d.%d', [3*(argMesic.ItemIndex + 1), aseRok.Value])));
  deDatumSplatnosti.Date := Trunc(EndOfTheMonth(deDatumDokladu.Date));
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.FormActivate(Sender: TObject);
begin
  if not Prerusit then Exit;             // to by mìlo ohlídat volání jen pøi startu

  aseRokChange(nil);
  arbVytvoreniClick(nil);
end;



// ------------------------------------------------------------------------------------------------

procedure TfmMain.arbVytvoreniClick(Sender: TObject);
begin
  asgMain.Cells[0, 0] := 'ZL1';
  asgMain.ColWidths[6] := 0;
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
  aedOd.LabelCaption := 'od VS';
  aedDo.LabelCaption := 'do VS';
  apbProgress.Visible := False;
  lbxLog.Visible := True;
  aseRokChange(Self);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.arbPrevodClick(Sender: TObject);
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
  glbVytvoreni.Color := clSilver;
  glbVytvoreni.ColorTo := clGray;
  aedOd.LabelCaption := 'od ZL1';
  aedDo.LabelCaption := 'do ZL1';
  aseRokChange(Self);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.arbTiskClick(Sender: TObject);
begin
  asgMain.Cells[0, 0] := 'tisk';
  asgMain.ColWidths[6] := 60;
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
  aedOd.LabelCaption := 'od ZL1';
  aedDo.LabelCaption := 'do ZL1';
  aseRokChange(Self);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.arbMailClick(Sender: TObject);
begin
  asgMain.Cells[0, 0] := 'mail';
  asgMain.ColWidths[6] := 60;
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
  aedOd.LabelCaption := 'od ZL1';
  aedDo.LabelCaption := 'do ZL1';
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

procedure TfmMain.argMesicClick(Sender: TObject);
begin
  aseRokChange(nil);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.aseRokChange(Sender: TObject);
// pøi zmìnì (nejen) roku nastaví nové deDatumDokladu, deDatumPlneni, aedSplatnost, aedOd a aedDo
var
SQLStr: AnsiString;
begin
  if not dbMain.Connected or not dbAbra.Connected then Exit;
  aedOd.Clear;
  aedDo.Clear;
  asgMain.ClearNormalCells;
  asgMain.RowCount := 2;
  btVytvorit.Caption := '&Naèíst';
  asgMain.Visible := True;
  lbxLog.Visible := False;
  {
  // Id období
  with qrAbra do begin
    Close;
    SQL.Text := 'SELECT Id FROM Periods WHERE Code = ' + Ap + aseRok.Text + Ap;
    Open;
    Period_Idsmazat := FieldByName('Id').AsString;
    Close;
  end;  // with qrAbra
  }
// datum dokladu je 15.3., 15.6., 15.9. a 15.12., datum splatnosti poslední den v témže mìsíci
  deDatumDokladu.Date := Trunc(StrToDate(Format('15.%d.%d', [3*(argMesic.ItemIndex+1), aseRok.Value])));
  deDatumSplatnosti.Date := Trunc(EndOfTheMonth(deDatumDokladu.Date));
// VystavenoOd pro hledání vystavených ZL je 1.3., 1.6., 1.9., 1.12.,
  VystavenoOd := StrToDate(Format('1.%d.%d', [3*(argMesic.ItemIndex + 1 ), aseRok.Value]));
// VystavenoDo pro hledání vystavených ZL je 30.5., 31.8., 30.11. a 28/29.2. v dalším roce,
  VystavenoDo := EndOfTheMonth(IncMonth(VystavenoOd, 2));
// *** výbìr ZL podle smlouvy
  if arbVytvoreni.Checked then with qrMain do try
    Screen.Cursor := crSQLWait;
// view pro generování
    dmCommon.AktualizaceView;
// první a poslední èíslo smlouvy
    SQL.Text := 'SELECT MIN(VS), MAX(VS) FROM ' + ZLView;
    Open;
    aedOd.Text := Fields[0].AsString;
    aedDo.Text := Fields[1].AsString;
    Close;
    finally
      Screen.Cursor := crDefault;
    end  // arbVytvoreni.Checked
// *** výbìr podle ZL (pøevod, tisk, mail)
  else begin
    dbAbra.Reconnect;
    with qrAbra do begin
// rozpìtí èísel ZL v období VystavenoOd až VystavenoDo
      SQLStr := 'SELECT MIN(OrdNumber) FROM IssuedDInvoices, Periods'
      + ' WHERE DocDate$DATE >= ' + FloatToStr(Trunc(VystavenoOd))
      + ' AND Periods.ID = Period_ID'
      + ' AND Periods.Code = ' + Ap + FloatToStr(YearOf(VystavenoOd)) + Ap
      + ' AND DocQueue_ID = ' + Ap + AbraEnt.getDocQueue('Code=ZL1').ID + Ap;
      Close;
      SQL.Text := SQLStr;
      Open;
      if Fields[0].AsInteger > 0 then aedOd.Text := Fields[0].AsString
      else begin                  // v roce VystavenoOd nebyl žádný ZL - to je asi jen pøi zahájení øady ZL1
        SQLStr := 'SELECT MIN(OrdNumber) FROM IssuedDInvoices, Periods'
        + ' WHERE DocDate$DATE >= ' + FloatToStr(Trunc(VystavenoOd))
        + ' AND Periods.ID = Period_ID'
        + ' AND Periods.Code = ' + Ap + FloatToStr(YearOf(VystavenoDo)) + Ap
        + ' AND DocQueue_ID = ' + Ap + AbraEnt.getDocQueue('Code=ZL1').ID + Ap;
        Close;
        SQL.Text := SQLStr;
        Open;
        aedOd.Text := Fields[0].AsString;
      end;
      SQLStr := 'SELECT MAX(OrdNumber) FROM IssuedDInvoices, Periods'
      + ' WHERE DocDate$DATE <= ' + FloatToStr(Trunc(VystavenoDo))
      + ' AND Periods.ID = Period_ID'
      + ' AND Periods.Code = ' + Ap + FloatToStr(YearOf(VystavenoDo)) + Ap
      + ' AND DocQueue_ID = ' + Ap + AbraEnt.getDocQueue('Code=ZL1').ID + Ap;
      Close;
      SQL.Text := SQLStr;
      Open;
      if Fields[0].AsInteger > 0 then begin
        aedDo.Text := Fields[0].AsString;
// vystaveno do mùže být v dalším roce, pak se pøidá rok do aedDo.Text
        if  YearOf(VystavenoDo) > YearOf(VystavenoOd) then aedDo.Text := aedDo.Text + '/' + FloatToStr(YearOf(VystavenoDo))
// když není ZL, zkusí se ještì YearOf(VystavenoOd)
      end else begin
        SQLStr := 'SELECT MAX(OrdNumber) FROM IssuedDInvoices, Periods'
        + ' WHERE DocDate$DATE <= ' + FloatToStr(Trunc(VystavenoDo))
        + ' AND Periods.ID = Period_ID'
        + ' AND Periods.Code = ' + Ap + FloatToStr(YearOf(VystavenoOd)) + Ap
        + ' AND DocQueue_ID = ' + Ap + AbraEnt.getDocQueue('Code=ZL1').ID + Ap;
        Close;
        SQL.Text := SQLStr;
        Open;
        aedDo.Text := Fields[0].AsString;
      end;
      Close;
      if aedOd.Text = '0' then
    end;  // with qrAbra
  end;  // if ... else arbVytvoreni.Checked
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.deDatumDokladuChange(Sender: TObject);
begin
  DatumDokladu := deDatumDokladu.Date;
//  16.9.22  deDatumSplatnosti.Date := Trunc(EndOfTheMonth(deDatumDokladu.Date));
  DatumSplatnosti := deDatumSplatnosti.Date;

  {
  // Id období
  with qrAbra do begin
    Close;
    SQL.Text := 'SELECT Id FROM Periods WHERE Code = ' + Ap + FormatDateTime('yyyy', DatumDokladu) + Ap;
    Open;
    Period_Idsmazat := FieldByName('Id').AsString;
    Close;
  end;  // with qrAbra
  }
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.deDatumSplatnostiChange(Sender: TObject);
begin
  DatumSplatnosti := deDatumSplatnosti.Date;
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
// 15.12.2020 oprava - ne VS
  if (aedOd.LabelCaption = 'od ZL1') and (aedOd.IntValue > aedDo.IntValue) then aedDo.Text := aedOd.Text;
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
// je-li aedDo menší než aedOd, opraví se
begin
// 15.12.2020 oprava - ne VS
  if (aedOd.LabelCaption = 'od ZL1') and (aedDo.IntValue < aedOd.IntValue) then aedOd.Text := aedDo.Text;
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
      if arbVytvoreni.Checked then dmVytvoreni.VytvorZL;
// *** Pøevod do PDF ***
      if arbPrevod.Checked then dmPrevod.PrevedZL;
// *** Tisk ***
      if arbTisk.Checked then dmTisk.TiskniZL;
// *** Posílání mailem ***
      if arbMail.Checked then dmMail.PosliZL;
    end;
  finally
    btKonec.Caption := '&Konec';
    btVytvorit.Enabled := True;
  end;
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
  // + PDFDir + '\%4d\%2.2d /home/abrapdf/%4d" "exit"', [aseRok.Value, 3*(argMesic.ItemIndex+1), aseRok.Value])), SW_SHOWNORMAL);

  DesU.syncAbraPdfToRemoteServer(aseRok.Value, 3*(argMesic.ItemIndex+1));
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.btKonecClick(Sender: TObject);
begin
  if btKonec.Caption = '&Pøerušit' then begin
    Prerusit := True;
    fmMain.Zprava('Pøerušeno uživatelem.');
    btKonec.Caption := '&Konec';
  end else begin
    fmMain.Zprava('Konec programu "Zálohové listy".');
    Close;
  end;
end;

procedure TfmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  with qrMain do try
    SQL.Text := 'DROP VIEW ' + ZLView;
    ExecSQL;
    SQL.Text := 'DROP VIEW ' + BillingView;
    ExecSQL;
    SQL.Text := 'DROP VIEW ' + BBmax;
    ExecSQL;
  except
  end;
end;

// ------------------------------------------------------------------------------------------------
// funkce nespojené pøímo s konkrétním prvkem formuláøe
// ------------------------------------------------------------------------------------------------


procedure TfmMain.Zprava(TextZpravy: string);
// do listboxu a logfile uloží èas a text zprávy
begin
  TextZpravy := FormatDateTime('dd.mm.yy hh:nn:ss  ', Now) + TextZpravy;
  lbxLog.Items.Add(TextZpravy);
  lbxLog.ItemIndex := lbxLog.Count - 1;
  Application.ProcessMessages;
  DesUtils.appendToFile(LogDir + FormatDateTime('yyyy.mm".log"', Date),
    Format('(%s - %s) ', [Trim(getWindowsCompName), Trim(getWindowsUserName)]) + TextZpravy);
end;

// ------------------------------------------------------------------------------------------------



end.
