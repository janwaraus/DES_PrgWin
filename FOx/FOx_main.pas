// 30.6.2017 faktury na kredit VoIP jsou te� automaticky vytv��eny v programu "V�pisy". Tady se mouhou p�ev�st do PDF, vytisknout
// a odeslat mailem.
// 19.10. tot� pro kreditn� faktury za p�ipojen�

unit FOx_main;

interface

uses
  Windows, Messages, Dialogs, SysUtils, Variants, Classes, Graphics, Controls, StdCtrls, ExtCtrls, Forms, Mask, ComObj, ComCtrls,
  AdvObj, AdvPanel, AdvEdit, AdvSpin, AdvDateTimePicker, AdvEdBtn, AdvFileNameEdit, AdvProgressBar, GradientLabel, AdvCombo,
  AdvOfficeButtons, AdvGroupBox, Grids, BaseGrid, AdvGrid, IniFiles, DateUtils, Math, Registry,
  DB, ZAbstractConnection, ZConnection, ZAbstractRODataset, ZAbstractDataset, ZDataset, frxClass, frxDBSet, frxDesgn;

type
  TfmMain = class(TForm)
    dbMain: TZConnection;
    qrMain: TZQuery;
    qrSmlouva: TZQuery;
    dbAbra: TZConnection;
    qrAbra: TZQuery;
    qrAdresa: TZQuery;
    qrRadky: TZQuery;
    qrDPH: TZQuery;
    frxReport: TfrxReport;
    frxDesigner: TfrxDesigner;
    fdsRadky: TfrxDBDataset;
    fdsDPH: TfrxDBDataset;
    apnTop: TAdvPanel;
    apnMain: TAdvPanel;
    aseMesic: TAdvSpinEdit;
    aseRok: TAdvSpinEdit;
    aedOd: TAdvEdit;
    aedDo: TAdvEdit;
    lbFakturyZa: TLabel;
    rbInternet: TRadioButton;
    rbVoIP: TRadioButton;
    btVytvorit: TButton;
    btKonec: TButton;
    apnVyberCinnosti: TAdvPanel;
    glbPrevod: TGradientLabel;
    glbTisk: TGradientLabel;
    glbMail: TGradientLabel;
    arbPrevod: TAdvOfficeRadioButton;
    arbTisk: TAdvOfficeRadioButton;
    arbMail: TAdvOfficeRadioButton;
    apbProgress: TAdvProgressBar;
    apnPrevod: TAdvPanel;
    cbNeprepisovat: TCheckBox;
    btOdeslat: TButton;
    btSablona: TButton;
    apnTisk: TAdvPanel;
    apnMail: TAdvPanel;
    fePriloha: TAdvFileNameEdit;
    lbxLog: TListBox;
    asgMain: TAdvStringGrid;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure dbMainAfterConnect(Sender: TObject);
    procedure dbAbraAfterConnect(Sender: TObject);
    procedure glbPrevodClick(Sender: TObject);
    procedure glbTiskClick(Sender: TObject);
    procedure glbMailClick(Sender: TObject);
    procedure arbPrevodClick(Sender: TObject);
    procedure arbTiskClick(Sender: TObject);
    procedure arbMailClick(Sender: TObject);
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
    procedure aseMesicChange(Sender: TObject);
    procedure aseRokChange(Sender: TObject);
    procedure aedOdChange(Sender: TObject);
    procedure aedOdKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure aedDoChange(Sender: TObject);
    procedure aedDoKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure rbInternetClick(Sender: TObject);
    procedure rbVoIPClick(Sender: TObject);
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
    Period_Id,
    FO2Queue_Id,
    FO4Queue_Id,
    FO_Id: string[10];
    PDFDir,
    AbraConnection: ShortString;
    CisloFO,
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
  end;

var
  fmMain: TfmMain;

implementation

uses FOx_Common, FOx_tisk, FOx_prevod, FOx_mail;

{$R *.dfm}

procedure TfmMain.FormCreate(Sender: TObject);
begin
  Prerusit := False;
end;

procedure TfmMain.FormShow(Sender: TObject);
// inicializace, z�sk�n� p�ipojovac�ch informac� pro datab�ze, p�ipojen�
var
  FileHandle: integer;
  FIIni: TIniFile;
  LogDir,
  LogFileName,
  FIFileName: AnsiString;
begin
// jm�no FI.ini
  FIFileName := ExtractFilePath(ParamStr(0)) + 'FIDES.ini';
  LogDir := ExtractFilePath(ParamStr(0)) + '\FOx - logy';
  if not DirectoryExists(LogDir) then CreateDir(LogDir);
  LogFileName := LogDir + FormatDateTime('\yyyy.mm.dd".log"', Date);
  if not FileExists(LogFileName) then begin
    FileHandle := FileCreate(LogFileName);
    FileClose(FileHandle);
  end;
  AssignFile(F, LogFileName);
  Append(F);
  Writeln(F);
  CloseFile(F);
  dmCommon.Zprava('Start programu "FOx".');
  if FileExists(FIFileName) then begin                     // existuje FI.ini ?
    FIIni := TIniFile.Create(FIFileName);
// p�ihla�ovac� �daje z FI.ini
    with FIIni do try
      dbAbra.HostName := ReadString('Preferences', 'AbraHN', '');
      dbAbra.Database := ReadString('Preferences', 'AbraDB', '');
      dbAbra.User := ReadString('Preferences', 'AbraUN', '');
      dbAbra.Password := ReadString('Preferences', 'AbraPW', '');
      dbMain.HostName := ReadString('Preferences', 'ZakHN', '');
      dbMain.Database := ReadString('Preferences', 'ZakDB', '');
      dbMain.User := ReadString('Preferences', 'ZakUN', '');
      dbMain.Password := ReadString('Preferences', 'ZakPW', '');
      dmMail.idSMTP.Host := ReadString('Mail', 'SMTPServer', 'mail.eurosignal.cz');
      dmMail.idSMTP.Username := ReadString('Mail', 'SMTPLogin', '');
      dmMail.idSMTP.Password := ReadString('Mail', 'SMTPPW', '');
      AbraConnection := ReadString('Preferences', 'AbraConn', '');
      PDFDir := ReadString('Preferences', 'PDFDir', '');
    finally
      FIIni.Free;
    end;
  end else begin
    Application.MessageBox(PChar(Format('Neexistuje soubor %s, program ukon�en.', [FIFileName])), PChar(FIFileName),
     MB_ICONERROR + MB_OK);
    dmCommon.Zprava(Format('Neexistuje soubor %s, program ukon�en.', [FIFileName]));
    Application.Terminate;
    fmMain.Close;
    Exit;
  end;
  Prerusit := True;                   // p��znak startu
// fajfky v asgMain
  with asgMain do begin
    CheckFalse := '0';
    CheckTrue := '1';
  end;
  aseMesic.Value := MonthOf(Date);
  aseRok.Value := YearOf(Date);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.FormActivate(Sender: TObject);
// aby bylo vid�t, �e se n�co d�je, p�ipojuj� se datab�ze a� tady, kdy u� je vid�t fmMain, proto�e OnActivate nastane i p�i p�epnut�
// z jin�ho okna, kontroluj� se existuj�c� p�ipojen�
begin
  if not Prerusit then Exit;             // to by m�lo ohl�dat vol�n� jen p�i startu
// p�ipojen� datab�z�
  if not dbAbra.Connected then try
    dmCommon.Zprava('P�ipojen� datab�ze Abry ...');
    dbAbra.Connect;
  except on E: exception do
    begin
      Application.MessageBox(PChar('Ned� se p�ipojit k datab�zi Abry, program ukon�en.' + ^M + E.Message), 'Abra', MB_ICONERROR + MB_OK);
      dmCommon.Zprava('Ned� se p�ipojit k datab�zi Abry, program ukon�en. ' + #13#10 + 'Chyba: ' + E.Message);
      Application.Terminate;
      fmMain.Close;
    end;
  end;
  if not dbMain.Connected then try
    dmCommon.Zprava('P�ipojen� datab�ze smluv ...');
    dbMain.Connect;
  except on E: exception do
    begin
      Application.MessageBox(PChar('Ned� se p�ipojit k datab�zi smluv, nep�jde faktury tisknout a rozes�lat mailem.' + ^M + E.Message), 'Abra', MB_ICONERROR + MB_OK);
      dmCommon.Zprava('Ned� se p�ipojit k datab�zi smluv, nep�jde faktury tisknout a rozes�lat mailem. ' + #13#10 + 'Chyba: ' + E.Message);
      arbTisk.Enabled := False;
      glbTisk.Enabled := False;
      arbMail.Enabled := False;
      glbMail.Enabled := False;
    end;
  end;
  Prerusit := False;
  aseRokChange(nil);
  arbPrevodClick(nil);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.dbMainAfterConnect(Sender: TObject);
begin
  with qrMain do begin                 // MySQL datab�ze z�kazn�k�
// p�eklad z UTF-8
    SQL.Text := 'SET CHARACTER SET cp1250';
    ExecSQL;
  end;
  dmCommon.Zprava('OK');
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.dbAbraAfterConnect(Sender: TObject);
begin
  with qrAbra do begin
    Close;
// Id doklad� FO
    SQL.Text := 'SELECT Id FROM DocQueues WHERE Code = ''FO2'' AND DocumentType = ''03''';
    Open;
    FO2Queue_Id := FieldByName('Id').AsString;
    Close;
    SQL.Text := 'SELECT Id FROM DocQueues WHERE Code = ''FO4'' AND DocumentType = ''03''';
    Open;
    FO4Queue_Id := FieldByName('Id').AsString;
    Close;
  end;
  dmCommon.Zprava('OK');
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.arbPrevodClick(Sender: TObject);
begin
  asgMain.Cells[0, 0] := 'PDF';
  btVytvorit.Caption := '&Na��st';
  glbPrevod.Color := clWhite;
  glbPrevod.ColorTo := clMenu;
  apnPrevod.Visible := True;
  glbTisk.Color := clSilver;
  glbTisk.ColorTo := clGray;
  apnTisk.Visible := False;
  glbMail.Color := clSilver;
  glbMail.ColorTo := clGray;
  apnMail.Visible := False;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.arbTiskClick(Sender: TObject);
begin
  asgMain.Cells[0, 0] := 'tisk';
  btVytvorit.Caption := '&Na��st';
  glbTisk.Color := clWhite;
  glbTisk.ColorTo := clMenu;
  apnTisk.Visible := True;
  glbMail.Color := clSilver;
  glbMail.ColorTo := clGray;
  apnMail.Visible := False;
  glbPrevod.Color := clSilver;
  glbPrevod.ColorTo := clGray;
  apnPrevod.Visible := False;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.arbMailClick(Sender: TObject);
begin
  asgMain.Cells[0, 0] := 'mail';
  btVytvorit.Caption := '&Na��st';
  glbMail.Color := clWhite;
  glbMail.ColorTo := clMenu;
  apnMail.Visible := True;
  glbPrevod.Color := clSilver;
  glbPrevod.ColorTo := clGray;
  apnPrevod.Visible := False;
  glbTisk.Color := clSilver;
  glbTisk.ColorTo := clGray;
  apnTisk.Visible := False;
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
begin
  if not dbMain.Connected or not dbAbra.Connected then Exit;
  aedOd.Clear;
  aedDo.Clear;
  asgMain.ClearNormalCells;
  asgMain.RowCount := 2;
  btVytvorit.Caption := '&Na��st';
  asgMain.Visible := True;
  lbxLog.Visible := False;
  dbAbra.Reconnect;
// Id obdob�
  with qrAbra do begin
    Close;
    SQL.Text := 'SELECT Id FROM Periods WHERE Code = ' + Ap + aseRok.Text + Ap;
    Open;
    Period_Id := FieldByName('Id').AsString;
    Close;
// rozp�t� ��sel FO v m�s�ci
    SQL.Text := 'SELECT MIN(OrdNumber), MAX(OrdNumber) FROM IssuedInvoices'
    + ' WHERE VATDate$DATE >= ' + FloatToStr(Trunc(StartOfAMonth(aseRok.Value, aseMesic.Value)))
    + ' AND VATDate$DATE <= ' + FloatToStr(Trunc(EndOfAMonth(aseRok.Value, aseMesic.Value)));
    if rbInternet.Checked then SQL.Text := SQL.Text + ' AND DocQueue_ID = ' + Ap + FO4Queue_Id + Ap
    else SQL.Text := SQL.Text + ' AND DocQueue_ID = ' + Ap + FO2Queue_Id + Ap;
    Open;
    if RecordCount > 0 then begin
      aedOd.Text := Fields[0].AsString;
      aedDo.Text := Fields[1].AsString;
      Close;
    end else begin
      aedOd.Clear;
      aedDo.Clear;
    end;
  end;  // with qrAbra
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

procedure TfmMain.rbInternetClick(Sender: TObject);
begin
  aseRokChange(nil);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.rbVoIPClick(Sender: TObject);
begin
  aseRokChange(nil);
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
// klik v prvn�m ��dku podle sloupce bu� ozna�uje/odzna�uje v�chny checkboxy, nebo spou�t� t��d�n�,
var
  Radek: integer;
begin
  with asgMain do if (ARow = 0) and (ACol = 0) then
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
  HAlign := Classes.taCenter;
  if (ACol > 0) and (ARow > 0) then
    if (ACol = 3) then HAlign := taRightJustify
    else if (ACol in [4..5]) then HAlign := taLeftJustify;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.lbxLogDblClick(Sender: TObject);
begin
  asgMain.Visible := True;
  lbxLog.Visible := False;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.frxReportBeginDoc(Sender: TObject);
var
  SQLStr: AnsiString;
begin
// ��dky faktury
  with qrRadky do begin
    Close;
    SQLStr := 'SELECT Text, TAmountWithoutVAT AS BezDane, VATRate AS Sazba, TAmount - TAmountWithoutVAT AS DPH, TAmount AS SDani'
    + ' FROM IssuedInvoices2'
    + ' WHERE Parent_ID = ' + Ap + FO_Id + Ap
    + ' AND NOT (Text = ''Zaokrouhlen�'' AND TAmount = 0)'
    + ' ORDER BY PosIndex';
    SQL.Text := SQLStr;
    Open;
  end;
// rekapitulace
  with qrDPH do begin
    Close;
    SQLStr := 'SELECT VATRate AS Sazba, SUM(TAmountWithoutVAT) AS BezDane, SUM(TAmount - TAmountWithoutVAT) AS DPH, SUM(TAmount) AS SDani FROM IssuedInvoices2'
    + ' WHERE Parent_ID = ' + Ap + FO_Id + Ap
    + ' AND VATIndex_ID IS NOT NULL'
    + ' GROUP BY Sazba';
    SQL.Text := SQLStr;
    Open;
  end;
end;
// ------------------------------------------------------------------------------------------------

procedure TfmMain.frxReportGetValue(const ParName: string; var ParValue: Variant);
// dosad� se prom�n� do formul��e
begin
  if ParName = 'Cislo' then ParValue := CisloFO
  else if ParName = 'VS' then ParValue := VS
  else if ParName = 'PJmeno' then ParValue := PJmeno
  else if ParName = 'PUlice' then ParValue := PUlice
  else if ParName = 'PObec' then ParValue := PObec
  else if ParName = 'OJmeno' then ParValue := OJmeno
  else if ParName = 'OUlice' then ParValue := OUlice
  else if ParName = 'OObec' then ParValue := OObec
  else if ParName = 'OICO' then begin
    if Trim(OICO) <> '' then ParValue := 'I�: ' + OICO else ParValue := ' '
  end else if ParName = 'ODIC' then begin
    if Trim(ODIC) <> '' then ParValue := 'DI�: ' + ODIC else ParValue := ' '
  end else if ParName = 'Vystaveni' then ParValue := Vystaveni
  else if ParName = 'Plneni' then ParValue := Plneni
  else if ParName = 'Splatnost' then ParValue := Splatnost
  else if ParName = 'CisloZL' then ParValue := CisloZL
  else if ParName = 'CastkaZalohy' then ParValue := CastkaZalohy
  else if ParName = 'Platek' then ParValue := Platek
  else if ParName = 'Celkem' then ParValue := Celkem
  else if ParName = 'Saldo' then ParValue := Abs(Saldo)
  else if ParName = 'Zaplatit' then ParValue := Format('%.2f K�', [Zaplatit])
  else if ParName = 'Resume' then ParValue := Format('��stku %.0f,- K� uhra�te, pros�m, do %s na ��et 2100098382/2010 s variabiln�m symbolem %s.',
   [Zaplatit, Splatnost, VS]);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.frxReportEndDoc(Sender: TObject);
begin
  qrRadky.Close;
  qrDPH.Close;
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
//      if arbVytvoreni.Checked then dmVytvoreni.VytvorFO;
// *** P�evod do PDF ***
      if arbPrevod.Checked then dmPrevod.PrevedFOx;
// *** Tisk ***
      if arbTisk.Checked then dmTisk.TiskniFOx;
// *** Pos�l�n� mailem ***
      if arbMail.Checked then dmMail.PosliFOx;
      btVytvorit.Caption := '&Na��st';
    end;
  finally
    btKonec.Caption := '&Konec';
    btVytvorit.Enabled := True;
  end;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.btSablonaClick(Sender: TObject);
begin
  frxReport.DesignReport(True, False);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.btOdeslatClick(Sender: TObject);
// odesl�n� faktur p�eveden�ch do PDF na vzd�len� server
begin
  WinExec(PAnsiChar(Format('WinSCP.com /command "option batch abort" "option confirm off" "open AbraPDF" "synchronize remote '
   + PDFDir + '\%4d\%2.2d /home/abrapdf/%4d" "exit"',
    [aseRok.Value, aseMesic.Value, aseRok.Value])), SW_SHOWNORMAL);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.btKonecClick(Sender: TObject);
begin
  if btKonec.Caption = '&P�eru�it' then begin
    Prerusit := True;
    dmCommon.Zprava('P�eru�eno u�ivatelem.');
    btKonec.Caption := '&Konec';
  end else begin
    dmCommon.Zprava('Konec programu "FOx".');
    Close;
  end;
end;

end.
