// 31.10.2008 dal�� verze v�pisu nezaplacen�ch faktur
// 5.1.2009 do tabulky p�id�ny sloupce K�d a Smlouva, ve kter�ch se vypisuje k�d firmy z Abry a pomoc� n�j
// ��slo smlouvy z datab�ze DES (tabulka Smlouvy)
// 18.11.09 rozes�l�n� varovn�ch mail�
// 30.3.10 rozes�l�n� mail� pomoc� idSMTP z Indy - umo��uje p�ihl�en� k serveru
// 12.4.11 zobrazen� v�ech smluv z�kazn�ka, potla�en� v�b�ru kreditn�ch smluv na VoIP
// 22.4. v�b�r textu mailu, tisk dopis� pro u�ivatele bez mailu
// 28.4. kontrola dlu�n� ��stky s 311-325
// 19.7. hromadn� omezov�n� a odpojov�n�
// 31.8. ozna�ov�n� nov�ch z�kazn�k� (s fakturac� od 1.7.11) s mo�nost� jejich v�b�ru a rozd�ln�ho �ikanov�n�
// 21.9. zm�na po�ad� a zobrazen� sloupc�
// 24.11.13 zru�eno rozli�ov�n� star�ch a nov�ch z�kazn�k�, potvrzen� pro omezov�n� a odpojov�n�, ukl�d�n� zpr�vy do
// tabulky Notes
// 23.1.14 logov�n� mail� a omezen�ch nebo odpojen�ch z�kazn�k�
// 26.6. omezov�n� a odpojov�n� pomoc� API od iQuestu
// 8.4.15 sjednoceno pro ABAK
// 17.12. ozna�ov�n� z�kazn�k� p�ipojen�ch v jin� s�ti

unit NF;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Forms, Controls, StdCtrls, ExtCtrls, ComCtrls, ComObj, Math,
  Mask, AdvCombo, AdvEdit, Dialogs, Grids, BaseGrid, AdvGrid, DB, IBDatabase, IBCustomDataSet, IBQuery, DateUtils, IniFiles,
  rxToolEdit, ZAbstractRODataset, ZAbstractDataset, ZDataset, ZConnection, IdBaseComponent, IdComponent, IdTCPConnection,
  IdTCPClient, IdMessageClient, IdSMTP, IdMessage, IdHTTP,
  IdAntiFreezeBase, IdAntiFreeze, ZAbstractConnection, AdvObj, IdIOHandler,
  IdIOHandlerSocket, IdSSLOpenSSL, IdExplicitTLSClientServerBase, IdSMTPBase;

type
  TfmMain = class(TForm)
    dbMain: TZConnection;
    qrMain: TZQuery;
    qrRows: TZQuery;
    dbAbra: TZConnection;
    qrAbra: TZQuery;
    qrAbra2: TZQuery;
    qrAbra3: TZQuery;
    idMessage: TIdMessage;
    idSMTP: TIdSMTP;
    dlgExport: TSaveDialog;
    pnMain: TPanel;
    lbDo: TLabel;
    lbOd: TLabel;
    deDatumOd: TDateEdit;
    deDatumDo: TDateEdit;
    acbRada: TAdvComboBox;
    aedPocet: TAdvEdit;
    aedPocetDo: TAdvEdit;
    acbDruhSmlouvy: TAdvComboBox;
    cbCast: TCheckBox;
    btVyber: TButton;
    btExport: TButton;
    btMail: TButton;
    btOmezit: TButton;
    btOdpojit: TButton;
    btKonec: TButton;
    asgPohledavky: TAdvStringGrid;
    pnBottom: TPanel;
    mmMail: TMemo;
    IdAntiFreeze1: TIdAntiFreeze;
    idHTTP: TIdHTTP;
    rgText: TRadioGroup;
    idSSLHandler: TIdSSLIOHandlerSocket;
    procedure FormShow(Sender: TObject);
    procedure dbAbraAfterConnect(Sender: TObject);
    procedure dbMainAfterConnect(Sender: TObject);
    procedure asgPohledavkyGetAlignment(Sender: TObject; ARow, ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asgPohledavkyGetFormat(Sender: TObject; ACol: Integer; var AStyle: TSortStyle; var aPrefix, aSuffix: string);
    procedure asgPohledavkyDblClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure asgPohledavkyCanSort(Sender: TObject; ACol: Integer; var DoSort: Boolean);
    procedure asgPohledavkyClickSort(Sender: TObject; ACol: Integer);
    procedure asgPohledavkyClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure asgPohledavkyCanEditCell(Sender: TObject; ARow, ACol: Integer; var CanEdit: Boolean);
    procedure btVyberClick(Sender: TObject);
    procedure btExportClick(Sender: TObject);
    procedure btMailClick(Sender: TObject);
    procedure btOmezitClick(Sender: TObject);
    procedure btOdpojitClick(Sender: TObject);
    procedure btKonecClick(Sender: TObject);
    procedure rgTextClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  public
    F: TextFile;
    Radek: integer;
    MailText: AnsiString;
  end;

const
  Ap = chr(39);
  ApC = Ap + ',';
  ApZ = Ap + ')';

var
  fmMain: TfmMain;

implementation

uses NF2D;

{$R *.dfm}

procedure TfmMain.FormShow(Sender: TObject);
var
  FileHandle: integer;
  FIIni: TIniFile;
  LogDir,
  LogFileName,
  FIFileName: AnsiString;
begin
  deDatumDo.Date := EndOfTheMonth(IncMonth(Now, -1));
  deDatumOd.Date := StartOfTheMonth(IncYear(Now, -1));
{$IFNDEF ABAK}
  rgText.Top := 350;
  rgText.Height := 73;
  rgText.Items.Add('Text 3');
{$ENDIF}
// 30.12.14 adres�� pro logy
  LogDir := ExtractFilePath(ParamStr(0)) + '\Nezaplacen� FO - logy';
  if not DirectoryExists(LogDir) then CreateDir(LogDir);
// vytvo�en� logfile, pokud neexistuje - 20.11.14 do jm�na p�id�no datum - 8.4.15 jen rok a m�s�c
  LogFileName := LogDir + FormatDateTime('\yyyy.mm".log"', Date);
  if not FileExists(LogFileName) then begin
    FileHandle := FileCreate(LogFileName);
    FileClose(FileHandle);
  end;
  AssignFile(F, LogFileName);
// jm�no FI.ini
{$IFDEF ABAK}
  FIFileName := ExtractFilePath(ParamStr(0)) + 'FIABAK.ini';
{$ELSE}
  FIFileName := ExtractFilePath(ParamStr(0)) + 'FIDES.ini';
{$ENDIF}
  if FileExists(FIFileName) then begin                     // existuje FI.ini ?
    FIIni := TIniFile.Create(FIFileName);
    with FIIni do try
      dbAbra.HostName := ReadString('Preferences', 'AbraHN', '');
      dbAbra.Database := ReadString('Preferences', 'AbraDB', '');
      dbAbra.User := ReadString('Preferences', 'AbraUN', '');
      dbAbra.Password := ReadString('Preferences', 'AbraPW', '');
      dbMain.HostName := ReadString('Preferences', 'ZakHN', '');
      dbMain.Database := ReadString('Preferences', 'ZakDB', '');
      dbMain.User := ReadString('Preferences', 'ZakUN', '');
      dbMain.Password := ReadString('Preferences', 'ZakPW', '');
      idSMTP.Host := ReadString('Mail', 'SMTPServer', 'mail.eurosignal.cz');
      idSMTP.Username := ReadString('Mail', 'SMTPLogin', '');
      idSMTP.Password := ReadString('Mail', 'SMTPPW', '');
    finally
      FIIni.Free;
    end;
  end else begin
    Application.MessageBox('Neexistuje soubor FI.ini, program ukon�en', 'FI.ini', MB_OK + MB_ICONERROR);
    Application.Terminate;
  end;
  try
    dbAbra.Connect;
  except on E: exception do
    begin
      Application.MessageBox(PChar('Ned� se p�ipojit k datab�zi Abry, program ukon�en.' + ^M + E.Message), 'Abra', MB_ICONERROR + MB_OK);
      Application.Terminate;
    end;
  end;
  try
    dbMain.Connect;
  except on E: exception do
    begin
      Application.MessageBox(PChar('Ned� se p�ipojit k datab�zi smluv, program ukon�en.' + ^M + E.Message), 'mySQL', MB_ICONERROR + MB_OK);
      Application.Terminate;
    end;
  end;
  with asgPohledavky do begin
    Clear;
    Cells[0, 0] := ' z�kazn�k';
    Cells[1, 0] := ' k�d';
    Cells[2, 0] := 'po�et FO';
    Cells[3, 0] := 'pohled�vky';
    Cells[4, 0] := '311-325';
    Cells[5, 0] := 'druh';
    Cells[6, 0] := 'smlouva';
    ColWidths[7] := 18;                // checkmark
    Cells[8, 0] := ' mail';
    Cells[9, 0] := 'telefon';
    ColWidths[10] := 0;                 // Cu.Id
    ColWidths[11] := 0;                 // C.Id
    CheckFalse := '0';
    CheckTrue := '1';
  end;
  aedPocet.Text := '1';
  aedPocetDo.Text := '12';
{$IFDEF ABAK}
  MailText := 'V�en� pane, v�en� pan�,' + #13#10
  + 'dovolujeme si V�s upozornit, �e Va�e faktura za internetov� p�ipojen� prost�ednictv�m s�t� �jezd.net za m�s�c '
  + FormatDateTime('mmmm', deDatumDo.Date) + ' nebyla dosud uhrazena. Prov��te, pros�m, tuto skute�nost a v p��pad� nejasnost� '
  + 'n�s laskav� kontaktujte.';
{$ELSE}
  MailText := Format('V�en� pane, v�en� pan�,' + #13#10
  + 'dovolujeme si V�s upozornit, �e je %d dn� po splatnosti pravideln� m�s��n� faktury za p�ipojen� k internetu '
  + 'a st�le od V�s postr�d�me jej� �hradu.' + #13#10
  + 'I kdy� pen�le za %d dn� zpo�d�n� platby ve v��i %.1f%% je v tuto chv�li zanedbateln� a jist� ho na V�s nebudeme vym�hat, '
  + 'pot�ilo by n�s, kdybyste dlu�nou ��stku co nejd��ve uhradili.', [DayOf(Date) - 10, DayOf(Date) - 10, (DayOf(Date) - 10) * 0.3]);
{$ENDIF}
  mmMail.Text := MailText;
end;

procedure TfmMain.dbAbraAfterConnect(Sender: TObject);
var
  SQLStr: ShortString;
begin
  with qrAbra do begin
    SQLStr := 'SELECT Code FROM DocQueues'                 // �ady faktur
    + ' WHERE DocumentType = ''03'''
    + ' AND Hidden = ''N'''
    + ' ORDER BY Code';
    SQL.Text := SQLStr;
    acbRada.Clear;
    acbRada.Items.Add('%');
    Open;
    while not EOF do begin
      acbRada.Items.Add(FieldByName('Code').AsString);
      Next;
    end;
  end;
  acbRada.ItemIndex := 0;
end;

procedure TfmMain.dbMainAfterConnect(Sender: TObject);
var
  SQLStr: ShortString;
begin
  with qrMain do begin
    SQL.Text := 'SET CHARACTER SET cp1250';                // p�eklad z UTF-8
    ExecSQL;
    SQLStr := 'SELECT DISTINCT State FROM contracts'       // stav smlouvy
    + ' ORDER BY State';
    SQL.Text := SQLStr;
    acbDruhSmlouvy.Clear;
    acbDruhSmlouvy.Items.Add('%');
    Open;
    while not EOF do begin
      acbDruhSmlouvy.Items.Add(FieldByName('State').AsString);
      Next;
    end;
  end;
  acbDruhSmlouvy.ItemIndex := 0;
end;

procedure TfmMain.asgPohledavkyGetAlignment(Sender: TObject; ARow, ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
  if ARow = 0 then HAlign := taLeftJustify
  else if (ACol in [2..4]) then HAlign := taRightJustify;
end;

procedure TfmMain.asgPohledavkyGetFormat(Sender: TObject; ACol: Integer; var AStyle: TSortStyle; var aPrefix, aSuffix: String);
begin
  if ACol in [2..4] then AStyle := ssNumeric
  else if ACol in [0, 1, 5, 6, 8, 9] then AStyle := ssAlphabetic;
end;

procedure TfmMain.asgPohledavkyCanSort(Sender: TObject; ACol: Integer; var DoSort: Boolean);
begin
  with asgPohledavky do
    if ACol = 7 then DoSort := False
    else if RowCount > 2 then RemoveRows(RowCount-1, 1);
end;

procedure TfmMain.asgPohledavkyCanEditCell(Sender: TObject; ARow, ACol: Integer; var CanEdit: Boolean);
begin
  CanEdit := ACol in [6..9];
end;

procedure TfmMain.asgPohledavkyClickSort(Sender: TObject; ACol: Integer);
begin
  with asgPohledavky do
    if (ACol <> 7) and (RowCount > 2) then begin
      RowCount := RowCount + 1;
      Cells[0, RowCount-1] := 'Celkem';
      Cells[2, RowCount-1] := Format('%d', [Trunc(ColumnSum(2, 1, RowCount-2))]);
      Floats[3, RowCount-1] := ColumnSum(3, 1, RowCount-2);
    end;
end;

procedure TfmMain.asgPohledavkyClickCell(Sender: TObject; ARow, ACol: Integer);
var
  Radek: integer;
begin
  if (ARow = 0) and (ACol = 7) then with asgPohledavky do
    if ColumnSum(7, 1, RowCount-2) = 0 then for Radek := 1 to RowCount-2 do Ints[7, Radek] := 1
    else for Radek := 1 to RowCount-2 do Ints[7, Radek] := 0;
end;

procedure TfmMain.asgPohledavkyDblClickCell(Sender: TObject; ARow, ACol: Integer);
begin
  Radek := ARow;
  fmDetail.ShowModal;
end;

procedure TfmMain.rgTextClick(Sender: TObject);
begin
  case rgText.ItemIndex of
{$IFDEF ABAK}
   0: MailText := 'V�en� pane, v�en� pan�,' + #13#10
      + 'dovolujeme si V�s upozornit, �e Va�e faktura za internetov� p�ipojen� prost�ednictv�m s�t� �jezd.net za m�s�c '
      + FormatDateTime('mmmm', deDatumDo.Date) + ' nebyla dosud uhrazena. Prov��te, pros�m tuto skute�nost a v p��pad� nejasnost� n�s laskav� kontaktujte.';
   1: MailText := 'V�en� pane, v�en� pan�,' + #13#10
      + 'dovolujeme si V�s upozornit, �e V� dluh za p�ipojen� k internetu dos�hl dvou m�s��n�ch plateb, nebo tuto ��stku ji� p�es�hl. '
      + 'Nebudou-li dlu�n� faktury ve v��i %%%,- K� uhrazeny do 5ti pracovn�ch dn�, budeme nuceni V�m linku do�asn� pozastavit. '
      + 'Poplatek za znovup�ipojen� �in� 300,- K� bez DPH.';
{$ELSE}
   0: MailText := Format('V�en� pane, v�en� pan�,' + #13#10
      + 'dovolujeme si V�s upozornit, �e je %d dn� po splatnosti pravideln� m�s��n� faktury za p�ipojen� k internetu '
      + 'a st�le od V�s postr�d�me jej� �hradu.' + #13#10
      + 'I kdy� pen�le za %d dn� zpo�d�n� platby ve v��i %.1f%% je v tuto chv�li zanedbateln� a jist� ho na V�s nebudeme vym�hat, '
      + 'pot�ilo by n�s, kdybyste dlu�nou ��stku co nejd��ve uhradili.', [DayOf(Date) - 10, DayOf(Date) - 10, (DayOf(Date) - 10) * 0.3]);
   1: MailText := 'V�en� pane, v�en� pan�,' + #13#10
      + 'upozor�ujeme V�s, �e V� dluh za p�ipojen� k internetu dos�hl dvou m�s��n�ch plateb, nebo tuto ��stku ji� p�es�hl. '
      + 'V brzk� dob� proto m��ete o�ek�vat sankce v podob� sn�en� rychlosti p�ipojen�.' + #13#10
      + 'Bli��� informace m��ete v pracovn� dny (9-16 h.) z�skat na ��sle 227 031 807, nebo kdykoli na sv�m z�kaznick�m ��tu na www.eurosignal.cz';
   2: MailText := 'V�en� pane, v�en� pan�,' + #13#10
      + 'upozor�ujeme V�s, �e V� dluh za p�ipojen� k internetu dos�hl t�� m�s��n�ch plateb, nebo tuto ��stku ji� p�es�hl. '
      + 'V nejbli��� dob� V�m proto bude p�eru�eno p�ipojen� k internetu.' + #13#10
      + 'Dal�� informace m��ete z�skat v pracovn� dny (9-16 h.) na ��sle 227 031 807, nebo kdykoli na sv�m z�kaznick�m ��tu na www.eurosignal.cz';
{$ENDIF}
  end;
  mmMail.Text := MailText;
end;

procedure TfmMain.btVyberClick(Sender: TObject);
var
  SQLStr: AnsiString;
  fmWidth,
  Radek: integer;
begin
  Screen.Cursor := crHourGlass;
  with qrAbra, asgPohledavky do try
    ClearNormalCells;
    SQLStr := 'SELECT F.Name AS Zakaznik, F.Code AS Kod, '
    + ' SUM(II.LocalAmount - II.LocalCreditAmount - (II.LocalPaidAmount - II.LocalPaidCreditAmount)) AS Castka, COUNT(*)'
    + ' FROM IssuedInvoices II'
    + ' INNER JOIN DocQueues DQ ON II.DocQueue_ID = DQ.Id'               // �ada doklad�
    + ' INNER JOIN Firms F ON II.Firm_ID = F.Id'                         // z�kazn�ci
    + ' WHERE II.DocDate$DATE <= ' + IntToStr(Trunc(deDatumDo.Date))
    + ' AND II.DocDate$DATE >= ' + IntToStr(Trunc(deDatumOd.Date));
    if cbCast.Checked then SQLStr := SQLStr
    + ' AND II.LocalAmount - II.LocalCreditAmount - (II.LocalPaidAmount - II.LocalPaidCreditAmount) > 0'
    else SQLStr := SQLStr
    + ' AND II.LocalAmount - II.LocalCreditAmount > 0'
    + ' AND II.LocalPaidAmount - II.LocalPaidCreditAmount = 0';
    if acbRada.Text <> '%' then SQLStr := SQLStr
    + ' AND DQ.Code = ' + Ap + acbRada.Text + Ap;
    SQLStr := SQLStr
    + ' GROUP BY F.Name, F.Code';
    Close;
    SQL.Text := SQLStr;
    Open;
    Radek := 0;
    while not EOF do begin
      if (FieldByName('Zakaznik').AsString <> '') and (FieldByName('Count').AsInteger >= StrToInt(aedPocet.Text))
       and (FieldByName('Count').AsInteger <= StrToInt(aedPocetDo.Text)) then begin
        Inc(Radek);
        RowCount := Radek + 1;
        AddCheckBox(7, Radek, True, True);
        Ints[7, Radek] := 0;
        Cells[0, Radek] := FieldByName('Zakaznik').AsString;
        Cells[1, Radek] := FieldByName('Kod').AsString;
        Ints[2, Radek] := FieldByName('Count').AsInteger;
        Floats[3, Radek] := FieldByName('Castka').AsFloat;
// kontrola dluhu proti 311-325
        with qrAbra2 do begin
// 24.11.13 v�echny Firm_Id pro Abrak�d firmy
          SQLStr := 'SELECT * FROM DE$_CODE_TO_FIRM_ID (' + Ap + Cells[1, Radek] + ApZ;
          SQL.Text := SQLStr;
          Open;
          Floats[4, Radek] := 0;
// a saldo pro v�echny Firm_Id
          while not EOF do with qrAbra3 do begin
            Close;
            SQLStr := 'SELECT Ucet311 + Ucet325 FROM DE$_Firm_Totals (' + Ap + qrAbra2.Fields[0].AsString + ApC + FloatToStr(Date) + ')';
            SQL.Text := SQLStr;
            Open;
            Floats[4, Radek] := Floats[4, Radek] - Fields[0].AsFloat;
            qrAbra2.Next;
          end; // while not EOF do with qrAbra3
        end;  // with qrAbra2
// v�echny smlouvy pro jeden abrak�d
        if dbMain.Connected then with qrMain do begin
          Close;
          SQLStr := 'SELECT DISTINCT Cu.Id AS CuId, C.Id AS CId, Postal_mail, Phone, vip, Number, State, Tag_Id'
          + ' FROM contracts C'
          + ' JOIN customers Cu ON Cu.Id = C.Customer_Id'
          + ' LEFT JOIN contracts_tags CT ON CT.Contract_Id = C.Id'      // 17.12.15
{$IFNDEF ABAK}
//          + ' AND Tariff_Id <> 2'                          // 12.4.11 ne EP-Basic
{$ENDIF}
          + ' WHERE Abra_Code = ' + Ap + Cells[1, Radek] + Ap
          + ' AND Activated_at <= ' + Ap + FormatDateTime('yyyy-mm-dd', deDatumDo.Date) + Ap;
          if acbDruhSmlouvy.Text <> '%' then SQLStr := SQLStr + ' AND State = ' + Ap + acbDruhSmlouvy.Text + Ap;
          SQL.Text := SQLStr;
          Open;
          Cells[5, Radek] := FieldByName('State').AsString;
          Cells[6, Radek] := FieldByName('Number').AsString;
// 17.12.15 p�ipojen� v jin� s�ti
          if FieldByName('Tag_Id').AsInteger in [20, 21, 25, 26, 27, 30] then Colors[6, Radek] := clRed;
          Cells[8, Radek] := FieldByName('Postal_mail').AsString;
          Cells[9, Radek] := FieldByName('Phone').AsString;
          Cells[10, Radek] := FieldByName('CuId').AsString;
          Cells[11, Radek] := FieldByName('CId').AsString;
          if Pos('@', FieldByName('Postal_mail').AsString) > 0 then Ints[7, Radek] := 1;
// 28.4.11 barvi�ky
          if Floats[3, Radek] - Floats[4, Radek] <> 0 then Colors[3, Radek] := clRed;   // dluh se li�� od 311-325
          if FieldByName('vip').AsInteger > 0 then Colors[0, Radek] := clSilver;              // vip
          if not EOF then Next;                            // 12.4.11 m��e b�t v�c smluv
          while not EOF do begin
            Inc(Radek);                                    // 22.9.11 a ka�d� m� sv�j ��dek
            RowCount := Radek + 1;
//            Cells[2, Radek] := Cells[2, Radek-1];
            Cells[5, Radek] := FieldByName('State').AsString;
            Cells[6, Radek] := FieldByName('Number').AsString;
            if FieldByName('Tag_Id').AsInteger in [20, 21, 25, 26, 27, 30] then Colors[6, Radek] := clRed;
            AddCheckBox(7, Radek, True, True);
            Ints[7, Radek] := 0;
            Cells[10, Radek] := FieldByName('CuId').AsString;
            Cells[11, Radek] := FieldByName('CId').AsString;
            Next;
          end;
          if RecordCount = 0 then begin
            ClearRows(Radek, 1);
            Dec(Radek);
          end;
          Application.ProcessMessages;
        end;
      end;
      Next;
    end;
// �prava zobrazen�
    AutoSize := True;
    ColWidths[7] := 18;
    ColWidths[10] := 0;                 // Cu.Id
    ColWidths[11] := 0;                 // C.Id
    if RowCount > 2 then begin
      Inc(Radek);
      RowCount := Radek + 1;
      Cells[0, RowCount-1] := 'Celkem';
      Cells[2, RowCount-1] := Format('%d', [Trunc(ColumnSum(2, 1, RowCount-2))]);
      Floats[3, RowCount-1] := ColumnSum(3, 1, RowCount-2);
    end;
    fmWidth := 0;
    for Radek := 0 to ColCount-1 do fmWidth := fmWidth +  ColWidths[Radek];
    fmMain.Width := fmWidth + 120;
    fmMain.Height := Min(800, RowCount * 19 + 164);
  finally
    qrMain.Close;
    qrAbra.Close;
//    ShowMessage(TimeToStr(Time - t));
    Screen.Cursor := crDefault;
  end;
end;

procedure TfmMain.btExportClick(Sender: TObject);
begin
{$IFDEF ABAK}
  with dlgExport do begin
    DefaultExt := '.csv';
    Filter := 'csv|*.csv';
    asgPohledavky.QuoteEmptyCells := True;
    if Execute then asgPohledavky.SaveToCSV(dlgExport.FileName);
  end;  
{$ELSE}
  if dlgExport.Execute then asgPohledavky.SaveToXLS(dlgExport.FileName);
{$ENDIF}
end;

procedure TfmMain.btMailClick(Sender: TObject);
var
  RadekDo,
  Radek,
  CommId: integer;
  MailStr,
  SQLStr: AnsiString;
begin
  Screen.Cursor := crHourGlass;
  with asgPohledavky, idMessage do begin
    if RowCount > 2 then RadekDo := RowCount - 2 else RadekDo := 1;
    Append(F);
    Writeln (F, FormatDateTime(#13#10 + 'dd.mm.yy hh:nn  ', Now) + 'Odesl�n� zpr�vy: ' + #13#10 + mmMail.Text + #13#10);
    CloseFile(F);
    for Radek := 1 to RadekDo do
      if Ints[7, Radek] = 1 then begin
        Clear;
        ContentType := 'text/plain';
        Charset := 'Windows-1250';
{$IFDEF ABAK}
        From.Address := 'abak@abak.cz';
        CCList.Add.Address := 'abak@abak.cz';
{$ELSE}
        From.Address := 'kontrola@eurosignal.cz';
{$ENDIF}
        MailStr := Cells[8, Radek];
        MailStr := StringReplace(MailStr, ',', ';', [rfReplaceAll]);    // ��rky za st�edn�ky
        while Pos(';', MailStr) > 0 do begin
          Recipients.Add.Address := Trim(Copy(MailStr, 1, Pos(';', MailStr)-1));
          MailStr := Copy(MailStr, Pos(';', MailStr)+1, Length(MailStr));
        end;
        Recipients.Add.Address := Trim(MailStr);
        Subject := 'Kontrola plateb smlouvy ' + Cells[6, Radek];
// 16.11.16        Body.Text := StringReplace(mmMail.Text, '%%%', IntToStr(Round(Floats[3, Radek])), [rfIgnoreCase]);
        Body.Text := StringReplace(mmMail.Text, '%%%', IntToStr(Round(Floats[4, Radek])), [rfIgnoreCase]);
        Body.Add(' ');
        Body.Add('S pozdravem');
        Body.Add(' ');
{$IFDEF ABAK}
        Body.Add('L.Kotal�kov�');
        Body.Add('Tel. 246030670-1');
{$ELSE}
        Body.Add('V� Eurosignal');
        Body.Add(' ');
        Body.Add('Na tuto zpr�vu neodpov�dejte, byla generov�na automaticky.');
{$ENDIF}
        with idSMTP do begin
          Port := 25;
          if Username = '' then AuthenticationType := atNone
          else AuthenticationType := atLogin;
        end;
        try
          if not idSMTP.Connected then idSMTP.Connect;
          idSMTP.Send(idMessage);
        except on E: exception do
          ShowMessage('Mail se nepoda�ilo odeslat: ' + E.Message);
        end;
        if (Colors[7, Radek] <> clSilver) then Colors[7, Radek] := clSilver
        else Colors[7, Radek] := clWhite;
        with qrMain do try                               // 15.3.2011
          Close;
          SQL.Text := 'SELECT MAX(Id) FROM communications';
          Open;
          CommId := Fields[0].AsInteger + 1;
          Close;
          SQLStr := 'INSERT INTO communications ('
          + ' Id,'
          + ' Customer_id,'
          + ' User_id,'
          + ' Communication_type_id,'
          + ' Content,'
          + ' Created_at,'
          + ' Updated_at) VALUES ('
          + IntToStr(CommId) + ', '
          + Cells[10, Radek] + ', '
          + '1, '                                        // admin
          + '2, '                                        // mail
          + Ap + mmMail.Text + ApC
          + Ap + FormatDateTime('yyyy-mm-dd hh:nn:ss', Now) + ApC
          + Ap + FormatDateTime('yyyy-mm-dd hh:nn:ss', Now) + ApZ;
          SQL.Text := SQLStr;
          ExecSQL;
          System.Append(F);
          Writeln (F, Format('%s (%s)  -  %s', [Cells[0, Radek], Cells[6, Radek], Cells[8, Radek]]));
          CloseFile(F);
        except on E: exception do
          ShowMessage('Mail se nepoda�ilo ulo�it do tabulky communications: ' + E.Message);
        end;
      end;
  end;
  if idSMTP.Connected then idSMTP.Disconnect;
  Screen.Cursor := crDefault;
end;

procedure TfmMain.btOmezitClick(Sender: TObject);
var
  RadekDo,
  Radek,
  NotesId,
  SinId: integer;
  URL,
  HTTPMessage,
  SQLStr: AnsiString;
begin
  Screen.Cursor := crHourGlass;
  with asgPohledavky, qrMain do try
    if RowCount > 2 then RadekDo := RowCount - 2 else RadekDo := 1;
    Radek := Trunc(ColumnSum (7, 1, RadekDo));             // po�et vybran�ch ��dk�
// 24.11.13 potvrzen�
    if Application.MessageBox(PChar(Format('Opravdu omezit %d z�kazn�k�?', [Radek])),
     'Pozor', MB_ICONQUESTION + MB_YESNO + MB_DEFBUTTON1) = IDNO then Exit;
    System.Append(F);
    Writeln (F, FormatDateTime(#13#10 + 'dd.mm.yy hh:nn  ', Now) + 'Omezen� rychlosti:' + #13#10);
    CloseFile(F);
// 26.6.14
    idHTTP.Request.Clear;
    idHTTP.Request.BasicAuthentication := True;
{$IFDEF ABAK}
    idHTTP.Request.Username := 'nfo';
    idHTTP.Request.Password := '9RG16Nu3';
    idHTTP.IOHandler := idSSLHandler;
    URL := 'https://iquest.ujd';
{$ELSE}
    idHTTP.Request.Username := 'NFO';
    idHTTP.Request.Password := 'NFO2014';
    URL := 'http://aplikace.eurosignal.cz';
{$ENDIF}
    for Radek := 1 to RadekDo do
      if Ints[7, Radek] = 1 then try
        HTTPMessage := idHTTP.Get(Format('%s/api/contracts/change_state?number=%s&state=restricted', [URL, Cells[6, Radek]]));
        if HTTPMessage <> 'OK' then
          if Application.MessageBox(PChar(Format('Omezen� se nepoda�ilo ulo�it do datab�ze: %s. Pokra�ovat?', [HTTPMessage])),
           'Pozor', MB_ICONQUESTION + MB_YESNO + MB_DEFBUTTON1) = IDNO then Exit
          else Continue;
        if (Colors[7, Radek] <> clSilver) then Colors[7, Radek] := clSilver
        else Colors[7, Radek] := clWhite;
        System.Append(F);
        Writeln (F, Format('%s (%s)', [Cells[0, Radek], Cells[6, Radek]]));
        CloseFile(F);
{$IFNDEF ABAK}
        try                                                  // 7.10.11
          SQL.Text := 'SELECT MAX(Id) FROM sinners';
          Open;
          SinId := Fields[0].AsInteger + 1;
          SQLStr := 'SELECT Invoice_debt, Account_debt, Enforcement_state FROM sinners'
          + ' WHERE Customer_id = ' + Cells[10, Radek]
          + ' AND Id = (SELECT MAX(Id) FROM sinners'
            + ' WHERE Customer_id = ' + Cells[10, Radek] + ')';
          Close;
          SQL.Text := SQLStr;
          Open;
          SQLStr := 'INSERT INTO sinners ('
          + ' Id,'
          + ' Customer_id,'
          + ' Contract_id,'
          + ' Date,'
          + ' Contract_state,'
          + ' Due_invoice_no,'
          + ' Invoice_debt,'
          + ' Account_debt,'
          + ' Enforcement_state,'
          + ' Event,'
          + ' Comment) VALUES ('
          + IntToStr(SinId) + ', '
          + Cells[10, Radek] + ', '
          + Cells[11, Radek] + ', '
          + Ap + FormatDateTime('yyyy-mm-dd hh:nn:ss', Now) + ApC
          + Ap + 'restricted' + ApC                         // stav smlouvy
          + Cells[2, Radek] + ', '                         // po�et faktur
          + Ap + FieldByName('Invoice_debt').AsString + ApC
          + Ap + FieldByName('Account_debt').AsString + ApC
          + Ap + FieldByName('Enforcement_state').AsString + ApC              // stav vym�h�n�
          + Ap + 'omezen�' + ApC
          + Ap + 'program NezaplaceneFaktury' + ApZ;
          Close;
          SQL.Text := SQLStr;
          ExecSQL;
        except on E: exception do
          ShowMessage('Omezen� se nepoda�ilo ulo�it do sinners: ' + E.Message);
        end;
{$ENDIF}
      except on E: exception do
        if Application.MessageBox(PChar(E.Message + #13#10 + 'Omezen� se nepoda�ilo ulo�it do datab�ze. Pokra�ovat?'),
         'Pozor', MB_ICONQUESTION + MB_YESNO + MB_DEFBUTTON1) = IDNO then Exit
        else Continue;
      end;
  finally
    Close;
    Screen.Cursor := crDefault;
  end;
end;

procedure TfmMain.btOdpojitClick(Sender: TObject);
var
  RadekDo,
  Radek,
  NotesId,
  SinId: integer;
  URL,
  HTTPMessage,
  SQLStr: AnsiString;
begin
  Screen.Cursor := crHourGlass;
  with asgPohledavky, qrMain do try
    if RowCount > 2 then RadekDo := RowCount - 2 else RadekDo := 1;
    Radek := Trunc(ColumnSum (7, 1, RadekDo));             // po�et vybran�ch ��dk�
// 24.11.13 potvrzen�
    if Application.MessageBox(PChar(Format('Opravdu odpojit %d z�kazn�k�?', [Radek])),
     'Pozor', MB_ICONQUESTION + MB_YESNO + MB_DEFBUTTON1) = IDNO then Exit;
    System.Append(F);
    Writeln (F, FormatDateTime(#13#10 + 'dd.mm.yy hh:nn  ', Now) + 'Odpojen�:' + #13#10);
    CloseFile(F);
// 26.6.14
    idHTTP.Request.Clear;
    idHTTP.Request.BasicAuthentication := True;
//    idHTTP.Request.UserAgent := 'Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.0; SLCC1)';
{$IFDEF ABAK}
// pro ABAK na test 13.9.2016
    idHTTP.Request.Username := 'nfo';
    idHTTP.Request.Password := '9RG16Nu3';
    idHTTP.IOHandler := idSSLHandler;
//    URL := 'https://iquest.ujd';
    URL := 'http://portal.ujezd.net';
{$ELSE}
    idHTTP.Request.Username := 'NFO';
    idHTTP.Request.Password := 'NFO2014';
    URL := 'http://aplikace.eurosignal.cz';
{$ENDIF}
    for Radek := 1 to RadekDo do
      if Ints[7, Radek] = 1 then try
        HTTPMessage := idHTTP.Get(Format('%s/api/contracts/change_state?number=%s&state=disconnected', [URL, Cells[6, Radek]]));
        if HTTPMessage <> 'OK' then
          if Application.MessageBox(PChar(Format('Odpojen� se nepoda�ilo ulo�it do datab�ze: %s. Pokra�ovat?', [HTTPMessage])),
           'Pozor', MB_ICONQUESTION + MB_YESNO + MB_DEFBUTTON1) = IDNO then Exit
          else Continue;
        if (Colors[7, Radek] <> clSilver) then Colors[7, Radek] := clSilver
        else Colors[7, Radek] := clWhite;
        System.Append(F);
        Writeln (F, Format('%s (%s)', [Cells[0, Radek], Cells[6, Radek]]));
        CloseFile(F);
{$IFNDEF ABAK}
        try                                                  // 7.10.11
          SQL.Text := 'SELECT MAX(Id) FROM sinners';
          Open;
          SinId := Fields[0].AsInteger + 1;
          SQLStr := 'SELECT Invoice_debt, Account_debt, Enforcement_state FROM sinners'
          + ' WHERE Customer_id = ' + Cells[10, Radek]
          + ' AND Id = (SELECT MAX(Id) FROM sinners'
            + ' WHERE Customer_id = ' + Cells[10, Radek] + ')';
          Close;
          SQL.Text := SQLStr;
          Open;
          SQLStr := 'INSERT INTO sinners ('
          + ' Id,'
          + ' Customer_id,'
          + ' Contract_id,'
          + ' Date,'
          + ' Contract_state,'
          + ' Due_invoice_no,'
          + ' Invoice_debt,'
          + ' Account_debt,'
          + ' Enforcement_state,'
          + ' Event,'
          + ' Comment) VALUES ('
          + IntToStr(SinId) + ', '
          + Cells[10, Radek] + ', '
          + Cells[11, Radek] + ', '
          + Ap + FormatDateTime('yyyy-mm-dd hh:nn:ss', Now) + ApC
          + Ap + 'disconnected' + ApC                         // stav smlouvy
          + Cells[2, Radek] + ', '                         // po�et faktur
          + Ap + FieldByName('Invoice_debt').AsString + ApC
          + Ap + FieldByName('Account_debt').AsString + ApC
          + Ap + FieldByName('Enforcement_state').AsString + ApC              // stav vym�h�n�
          + Ap + 'odpojen�' + ApC
          + Ap + 'program NezaplaceneFaktury' + ApZ;
          SQL.Text := SQLStr;
          ExecSQL;
        except on E: exception do
          ShowMessage('Odpojen� se nepoda�ilo ulo�it do sinners: ' + E.Message);
        end;
{$ENDIF}
      except on E: exception do
        if Application.MessageBox(PChar(E.Message + #13#10 + 'Odpojen� se nepoda�ilo ulo�it do datab�ze. Pokra�ovat?'),
         'Pozor', MB_ICONQUESTION + MB_YESNO + MB_DEFBUTTON1) = IDNO then Exit
        else Continue;
      end;
  finally
    Close;
    Screen.Cursor := crDefault;
  end;
end;

procedure TfmMain.btKonecClick(Sender: TObject);
begin
  Close;
end;

procedure TfmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  qrAbra.Close;
end;

end.
