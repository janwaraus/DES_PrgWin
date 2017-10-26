// 31.10.2008 dal�� verze v�pisu nezaplacen�ch faktur
// 5.1.2009 do tabulky p�id�ny sloupce K�d a Smlouva, ve kter�ch se vypisuje k�d firmy z Abry a pomoc� n�j
// ��slo smlouvy z datab�ze DES (tabulka Smlouvy)
// 18.11.09 rozes�l�n� varovn�ch mail�
// 30.3.10 rozes�l�n� mail� pomoc� idSMTP z Indy - umo��uje p�ihl�en� k serveru
// 12.4.11 zobrazen� v�ech smluv z�kazn�ka, potla�en� v�b�ru kreditn�ch smluv na VoIP
// 22.4.11 v�b�r textu mailu, tisk dopis� pro u�ivatele bez mailu
// 28.4.11 kontrola dlu�n� ��stky s 311-325
// 19.7.11 hromadn� omezov�n� a odpojov�n�
// 20.1.14 potvrzen� pro omezov�n� a odpojov�n�, ukl�d�n� zpr�vy do tabulky Notes
// 29.1.14 logov�n� mail� a omezen�ch nebo odpojen�ch z�kazn�k�
// �30.12. omezov�n� a odpojov�n� pomoc� API od iQuestu

unit NZL;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Forms, Controls, StdCtrls, ExtCtrls, ComCtrls, ComObj,
  Mask, AdvCombo, AdvEdit, Dialogs, Grids, BaseGrid, AdvGrid, DB, DateUtils,
  //IBDatabase, IBCustomDataSet, IBQuery, IniFiles,
  //rxToolEdit,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, ZConnection, IdBaseComponent, IdComponent, IdTCPConnection,
  IdTCPClient, IdSMTP, IdHTTP, IdMessage, IdMessageClient, IdText, IdMessageParts,
  IdAntiFreezeBase, IdAntiFreeze, ZAbstractConnection, AdvObj, IdIOHandler,
  IdIOHandlerSocket, IdSSLOpenSSL, IdExplicitTLSClientServerBase, IdSMTPBase;

type
  TfmZL = class(TForm)
    //DesU.qrZakos: TZQuery;
    //qrRows: TZQuery;
    //dbAbra: TZConnection;
    //qrAbra: TZQuery;
    //qrAbra2: TZQuery;
    //qrAbra3: TZQuery;
    idMessage: TIdMessage;
    idSMTP: TIdSMTP;
    dlgExport: TSaveDialog;
    pnMain: TPanel;
    lbDo: TLabel;
    lbOd: TLabel;
    acbRada: TAdvComboBox;
    acbDruhSmlouvy: TAdvComboBox;
    cbCast: TCheckBox;
    btVyber: TButton;
    btExport: TButton;
    btMail: TButton;
    rgText: TRadioGroup;
    btOdpojit: TButton;
    btKonec: TButton;
    asgPohledavky: TAdvStringGrid;
    pnBottom: TPanel;
    mmMail: TMemo;
    idHTTP: TIdHTTP;
    IdAntiFreeze1: TIdAntiFreeze;
    deDatumOd: TDateTimePicker;
    deDatumDo: TDateTimePicker;
    btSMS: TButton;
    procedure FormShow(Sender: TObject);
    procedure asgPohledavkyGetAlignment(Sender: TObject; ARow, ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asgPohledavkyGetFormat(Sender: TObject; ACol: Integer; var AStyle: TSortStyle; var aPrefix, aSuffix: string);
    procedure asgPohledavkyDblClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure asgPohledavkyCanSort(Sender: TObject; ACol: Integer; var DoSort: Boolean);
    procedure asgPohledavkyClickSort(Sender: TObject; ACol: Integer);
    procedure asgPohledavkyCanEditCell(Sender: TObject; ARow, ACol: Integer; var CanEdit: Boolean);
    procedure btVyberClick(Sender: TObject);
    procedure btExportClick(Sender: TObject);
    procedure btMailClick(Sender: TObject);
    procedure btOdpojitClick(Sender: TObject);
    procedure btKonecClick(Sender: TObject);
    procedure rgTextClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure asgPohledavkyClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure btSMSClick(Sender: TObject);
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
  fmZL: TfmZL;

implementation

uses DesUtils, NZL2D;

{$R *.dfm}

procedure TfmZL.FormShow(Sender: TObject);
var
  FileHandle: integer;
  LogDir,
  LogFileName: AnsiString;
begin
  deDatumDo.Date := EndOfTheMonth(IncMonth(Now, -1));
  deDatumOd.Date := StartOfTheMonth(IncYear(Now, -1));
// 30.12.14 adres�� pro logy
  LogDir := DesU.PROGRAM_PATH + '\logy\Nezaplacen� ZL\';
  if not DirectoryExists(LogDir) then Forcedirectories(LogDir);
// vytvo�en� logfile, pokud neexistuje - 27.7.16 ve jm�n� jen rok a m�s�c
  LogFileName := LogDir + FormatDateTime('\yyyy.mm".log"', Date);
  if not FileExists(LogFileName) then begin
    FileHandle := FileCreate(LogFileName);
    FileClose(FileHandle);
  end;
  AssignFile(F, LogFileName);

  with asgPohledavky do begin
    ClearNormalCells;
    {
    Cells[0, 0] := ' z�kazn�k';
    Cells[1, 0] := 'pohled�vky';
    Cells[2, 0] := 'po�et ZL';
    Cells[4, 0] := 'druh';
    Cells[5, 0] := 'smlouva';
    Cells[8, 0] := ' mail';
    }
    ColWidths[3] := 0;                 // Firms.ID
    ColWidths[6] := 0;                 // Firms.Code
    ColWidths[7] := 18;                // checkmark
    ColWidths[9] := 0;                 // Cu.Id
    CheckFalse := '0';
    CheckTrue := '1';
  end;
  MailText := 'V�en� pane, v�en� pan�,' + sLineBreak
  + 'dovolujeme si V�s upozornit, �e je XXX dn� po splatnosti z�lohy na p�ipojen� k internetu a st�le od V�s postr�d�me jej� �plnou �hradu.'
  + ' Dlu�n� ��stka �in� YYY K�.' + sLineBreak
  + 'Pot�ilo by n�s, kdybyste ��stku z�lohy co nejd��ve uhradili a my V�m nemuseli omezovat poskytovan� slu�by.';
  mmMail.Text := MailText;

  acbRada.Clear;
  acbRada.Items.Add('%');
  with DesU.qrAbra do begin
    SQL.Text := 'SELECT Code FROM DocQueues'                 // �ady faktur
    + ' WHERE DocumentType = ''10'''
    + ' AND Hidden = ''N'''
    + ' ORDER BY Code';
    Open;
    while not EOF do begin
      acbRada.Items.Add(FieldByName('Code').AsString);
      Next;
    end;
  end;
  acbRada.ItemIndex := 0;

  acbDruhSmlouvy.Clear;
  acbDruhSmlouvy.Items.Add('%');
  with DesU.qrZakos do begin
    //SQL.Text := 'SET CHARACTER SET cp1250';                // p�eklad z UTF-8
    //ExecSQL;
    SQL.Text := 'SELECT DISTINCT State FROM contracts'       // stav smlouvy
    + ' ORDER BY State';
    Open;
    while not EOF do begin
      acbDruhSmlouvy.Items.Add(FieldByName('State').AsString);
      Next;
    end;
  end;
  acbDruhSmlouvy.ItemIndex := 0;
end;



procedure TfmZL.rgTextClick(Sender: TObject);
begin
  case rgText.ItemIndex of
   0: MailText := 'V�en� pane, v�en� pan�,' + sLineBreak
      + 'dovolujeme si V�s upozornit, �e je XXX dn� po splatnosti z�lohy na p�ipojen� k internetu a st�le od V�s postr�d�me jej� �plnou �hradu.'
      + ' Dlu�n� ��stka �in� YYY K�.' + sLineBreak
      + 'Pot�ilo by n�s, kdybyste ��stku z�lohy co nejd��ve uhradili a my V�m nemuseli omezovat poskytovan� slu�by.';
   1: MailText := 'V�en� pane, v�en� pan�,' + sLineBreak
      + 'upozor�ujeme V�s, �e je XXX dn� po splatnosti z�lohy na p�ipojen� k internetu a st�le od V�s postr�d�me jej� �plnou �hradu.'
      + ' Dlu�n� ��stka �in� YYY K�.' + sLineBreak
      + 'V brzk� dob� proto m��ete o�ek�vat sn�en� rychlosti p�ipojen� na 64 kbps.' + sLineBreak
      + 'Bli��� informace m��ete v pracovn� dny (9-16 h.) z�skat na ��sle 227 031 807.';
   2: MailText := 'V�en� pane, v�en� pan�,' + sLineBreak
      + 'upozor�ujeme V�s, �e je XXX dn� po splatnosti z�lohy na p�ipojen� k internetu a st�le od V�s postr�d�me jej� �plnou �hradu.'
      + ' Dlu�n� ��stka �in� YYY K�.' + sLineBreak
      + 'V nejbli��� dob� V�m proto bude Va�e p�ipojen� k internetu p�eru�eno.' + sLineBreak
      + 'Dal�� informace m��ete z�skat v pracovn� dny (9-16 h.) na ��sle 227 031 807.';
  end;
  mmMail.Text := MailText;
end;

procedure TfmZL.btVyberClick(Sender: TObject);
var
  SQLStr: AnsiString;
  fmWidth,
  Radek: integer;
begin
  Screen.Cursor := crHourGlass;
  with DesU.qrAbra, asgPohledavky do try
    ClearNormalCells;
    SQLStr := 'SELECT MIN(IDI.DueDate$DATE) AS Datum, F.Name AS Zakaznik, F.Id AS Id, F.Code AS Kod,'
    + ' SUM(IDI.LocalAmount - IDI.LocalPaidAmount) AS Castka, COUNT(*)'
    + ' FROM IssuedDInvoices IDI'
    + ' INNER JOIN DocQueues DQ ON IDI.DocQueue_ID = DQ.Id'               // �ada doklad�
    + ' INNER JOIN Firms F ON IDI.Firm_ID = F.Id'                         // z�kazn�ci
    + ' WHERE IDI.DueDate$DATE < ' + IntToStr(Trunc(deDatumDo.Date)+1)
    + ' AND IDI.DueDate$DATE >= ' + IntToStr(Trunc(deDatumOd.Date));
    if cbCast.Checked then SQLStr := SQLStr
    + ' AND IDI.LocalAmount - IDI.LocalPaidAmount > 0'
    else SQLStr := SQLStr
    + ' AND IDI.LocalAmount > 0'
    + ' AND IDI.LocalPaidAmount = 0';
    if acbRada.Text <> '%' then SQLStr := SQLStr
    + ' AND DQ.Code = ' + Ap + acbRada.Text + Ap;
    SQLStr := SQLStr
    + ' GROUP BY F.Name, F.Code, F.Id';
    Close;

    SQL.Text := SQLStr;
    Open;
    Radek := 0;
    while not EOF do begin
      if FieldByName('Zakaznik').AsString <> '' then begin
        Inc(Radek);
        RowCount := Radek + 1;
        AddCheckBox(7, Radek, True, True);
        Ints[7, Radek] := 0;
        Cells[0, Radek] := FieldByName('Zakaznik').AsString;
        Floats[1, Radek] := FieldByName('Castka').AsFloat;
        Ints[2, Radek] := FieldByName('Count').AsInteger;
        Cells[3, Radek] := FieldByName('Id').AsString;
        Cells[6, Radek] := FieldByName('Kod').AsString;
        Ints[11, Radek] := Trunc(Date) - FieldByName('Datum').AsInteger;
        with DesU.qrZakos do begin
          Close;
          SQLStr := 'SELECT DISTINCT Cu.Id AS CuId, C.Id AS CId, Postal_mail, Phone, vip, Number, State, Tag_Id'
          + ' FROM contracts C'
          + ' JOIN customers Cu ON Cu.Id = C.Customer_Id'
          + ' LEFT JOIN contracts_tags CT ON CT.Contract_Id = C.Id'      // 17.12.15
          + ' WHERE Tariff_Id <> 2'                          // 12.4.11 ne EP-Basic
          + ' AND Abra_Code = ' + Ap + Cells[6, Radek] + Ap
          + ' AND Activated_at <= ' + Ap + FormatDateTime('yyyy-mm-dd', deDatumDo.Date) + Ap;
          if acbDruhSmlouvy.Text <> '%' then SQLStr := SQLStr + ' AND State = ' + Ap + acbDruhSmlouvy.Text + Ap;
          SQL.Text := SQLStr;
          Open;
          Cells[4, Radek] := FieldByName('State').AsString;
          Cells[5, Radek] := FieldByName('Number').AsString;
// 29.7.16 p�ipojen� v jin� s�ti, vip
          if FieldByName('Tag_Id').AsInteger in [20, 21, 25, 26, 27, 30] then Colors[5, Radek] := clRed;
          if FieldByName('vip').AsInteger > 0 then Colors[0, Radek] := clSilver;
          if (Pos('active', FieldByName('State').AsString) > 0) and (Pos('@', FieldByName('Postal_mail').AsString) > 0) then begin
            Ints[7, Radek] := 1;
            Cells[8, Radek] := FieldByName('Postal_mail').AsString;
          end; // !!!podivat se

          Cells[12, Radek] := FieldByName('Phone').AsString;
          Cells[13, Radek] := destilujMobilCislo(FieldByName('Phone').AsString);
          Cells[9, Radek] := FieldByName('CuId').AsString;
          Cells[10, Radek] := FieldByName('CId').AsString;
          if not EOF then Next;                            // 29.7.16 m��e b�t v�c smluv
          while not EOF do begin
            if (Cells[5, Radek] <> FieldByName('Number').AsString) then begin    // 18.1.17 u smlouvy m��e b�t v�c tag�
              Inc(Radek);                                  // 29.7.16 ka�d� smlouva m� sv�j ��dek
              RowCount := Radek + 1;
              Cells[4, Radek] := FieldByName('State').AsString;
              Cells[5, Radek] := FieldByName('Number').AsString;
              if FieldByName('Tag_Id').AsInteger in [20, 21, 25, 26, 27, 30] then Colors[5, Radek] := clRed;    // je v jin� s�ti
              AddCheckBox(7, Radek, True, True);
              Ints[7, Radek] := 0;
            end else if FieldByName('Tag_Id').AsInteger in [20, 21, 25, 26, 27, 30] then Colors[5, Radek-1] := clRed;
            Next;
          end;
          if RecordCount = 0 then Dec(Radek);
          Application.ProcessMessages;
        end;
      end;
      Next;
    end;
// �prava zobrazen�
    AutoSize := True;
    ColWidths[3] := 0;                 // Firms.ID
    ColWidths[6] := 0;                 // Firms.Code
    ColWidths[7] := 18;
    ColWidths[9] := 0;                 // customers.Id
    ColWidths[10] := 0;                // contracts.Id
    if RowCount > 2 then begin
      Inc(Radek);
      RowCount := Radek + 1;
      Cells[0, RowCount-1] := 'Celkem';
      Floats[1, RowCount-1] := ColumnSum(1, 1, RowCount-2);
      Cells[2, RowCount-1] := Format('%d', [Trunc(ColumnSum(2, 1, RowCount-2))]);
    end;
    fmWidth := 0;
    for Radek := 0 to ColCount-1 do fmWidth := fmWidth +  ColWidths[Radek];
    fmZL.Width := fmWidth + 120;
  finally
    DesU.qrZakos.Close;
    DesU.qrAbra.Close;
//    ShowMessage(TimeToStr(Time - t));
    Screen.Cursor := crDefault;
  end;
end;

procedure TfmZL.btExportClick(Sender: TObject);
begin
  if dlgExport.Execute then asgPohledavky.SaveToXLS(dlgExport.FileName);
end;

procedure TfmZL.btMailClick(Sender: TObject);
var
  RadekDo,
  Radek,
  CommId: integer;
  Zprava,
  MailStr,
  SQLStr: AnsiString;
begin
  Screen.Cursor := crHourGlass;
  idSMTP.Host :=  DesU.getIniValue('Mail', 'SMTPServer');
  idSMTP.Username := DesU.getIniValue('Mail', 'SMTPLogin');
  idSMTP.Password := DesU.getIniValue('Mail', 'SMTPPW');

  Append(F);
  Writeln (F, FormatDateTime(sLineBreak + 'dd.mm.yy hh:nn  ', Now) + 'Odesl�n� zpr�vy: ' + sLineBreak + mmMail.Text + sLineBreak);
  CloseFile(F);

  with asgPohledavky, idMessage do begin
    if RowCount > 2 then RadekDo := RowCount - 2 else RadekDo := 1;

    Radek := Trunc(ColumnSum (7, 1, RadekDo));             // po�et vybran�ch ��dk�
    if Application.MessageBox(PChar(Format('Opravdu poslat %d e-mail�?', [Radek])),
     'Pozor', MB_ICONQUESTION + MB_YESNO + MB_DEFBUTTON1) = IDNO then Exit;

    for Radek := 1 to RadekDo do
      if Ints[7, Radek] = 1 then begin
        Clear;

        From.Address := 'kontrola@eurosignal.cz';

        MailStr := Cells[8, Radek];
        MailStr := StringReplace(MailStr, ',', ';', [rfReplaceAll]);    // ��rky za st�edn�ky
        while Pos(';', MailStr) > 0 do begin
          Recipients.Add.Address := Trim(Copy(MailStr, 1, Pos(';', MailStr)-1));
          MailStr := Copy(MailStr, Pos(';', MailStr)+1, Length(MailStr));
        end;
        Recipients.Add.Address := Trim(MailStr);

        Subject := 'Kontrola plateb smlouvy ' + Cells[5, Radek];

        with TIdText.Create(idMessage.MessageParts, nil) do begin
          Zprava := StringReplace(mmMail.Text, 'XXX', Cells[11, Radek], []);
          Zprava := StringReplace(Zprava, 'YYY', Cells[1, Radek], []);
          Body.Text := Zprava
           + sLineBreak + sLineBreak
           +'S pozdravem'
           + sLineBreak + sLineBreak
           + 'V� Eurosignal'
           + sLineBreak + sLineBreak
           +'Na tuto zpr�vu neodpov�dejte, byla generov�na automaticky.';

          ContentType := 'text/plain';
          Charset := 'utf-8';
        end;

        ContentType := 'multipart/mixed';

        {
        with idSMTP do begin
          Port := 25;
          if Username = '' then AuthenticationType := atNone
          else AuthenticationType := atLogin;
        end;
        }
        try
          if not idSMTP.Connected then idSMTP.Connect;
          idSMTP.Send(idMessage);
        except on E: exception do
          ShowMessage('Mail se nepoda�ilo odeslat: ' + E.Message);
        end;

        if (Colors[7, Radek] <> clSilver) then Colors[7, Radek] := clSilver
        else Colors[7, Radek] := clWhite;
        with DesU.qrZakos do try                               // 15.3.2011
          Close;
          SQL.Text := 'SELECT MAX(Id) FROM communications';
          Open;
          CommId := Fields[0].AsInteger + 1;
          Close;
          SQLStr := 'INSERT INTO communications ('
          + ' id,'
          + ' customer_id,'
          + ' user_id,'
          + ' communication_type_id,'
          + ' content,'
          + ' created_at,'
          + ' updated_at) VALUES ('
          + IntToStr(CommId) + ', '
          + Cells[9, Radek] + ', '
          + '1, '                                        // admin
          + '2, '                                        // mail
          + Ap + Zprava + ApC
          + Ap + FormatDateTime('yyyy-mm-dd hh:nn:ss', Now) + ApC
          + Ap + FormatDateTime('yyyy-mm-dd hh:nn:ss', Now) + ApZ;
          SQL.Text := SQLStr;
          ExecSQL;
          System.Append(F);
          Writeln (F, Format('%s (%s) email %s', [Cells[0, Radek], Cells[5, Radek], Cells[8, Radek]]));
          CloseFile(F);
        except on E: exception do
          ShowMessage('Mail se nepoda�ilo ulo�it do tabulky communications: ' + E.Message);
        end;
      end;
  end;
  if idSMTP.Connected then idSMTP.Disconnect;
  Screen.Cursor := crDefault;
end;


procedure TfmZL.btSMSClick(Sender: TObject);
var
  RadekDo,
  Radek,
  CommId: integer;
  smsText, callResult,
  SQLStr: string;
begin
  Screen.Cursor := crHourGlass;


  with asgPohledavky do begin
    if RowCount > 2 then RadekDo := RowCount - 2 else RadekDo := 1;

    Radek := Trunc(ColumnSum (7, 1, RadekDo));             // po�et vybran�ch ��dk�
    if Application.MessageBox(PChar(Format('Opravdu poslat %d SMS?', [Radek])),
     'Pozor', MB_ICONQUESTION + MB_YESNO + MB_DEFBUTTON1) = IDNO then Exit;

    Append(F);
    Writeln (F, FormatDateTime(sLineBreak + 'dd.mm.yy hh:nn  ', Now) + 'Odesl�n� SMS zpr�vy: ' + sLineBreak + mmMail.Text + sLineBreak);
    CloseFile(F);
    for Radek := 1 to RadekDo do
      if Ints[7, Radek] = 1 then begin

        smsText := StringReplace(mmMail.Text, 'XXX', Cells[11, Radek], []);
        smsText := StringReplace(smsText, 'YYY', Cells[1, Radek], []);

        callResult := DesU.sendGodsSms(Cells[13, Radek], smsText);

        if (Colors[7, Radek] <> clSilver) then Colors[7, Radek] := clSilver
        else Colors[7, Radek] := clWhite;
        with DesU.qrZakos do try                               // 15.3.2011
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
          + '23, '                                        // SMS
          + Ap + mmMail.Text + ApC
          + Ap + FormatDateTime('yyyy-mm-dd hh:nn:ss', Now) + ApC
          + Ap + FormatDateTime('yyyy-mm-dd hh:nn:ss', Now) + ApZ;
          SQL.Text := SQLStr;
          ExecSQL;
          System.Append(F);
          Writeln (F, Format('%s (%s)  -  %s  -  %s', [Cells[0, Radek], Cells[5, Radek], Cells[13, Radek], callResult]));
          CloseFile(F);
        except on E: exception do
          ShowMessage('SMS se nepoda�ilo ulo�it do tabulky communications: ' + E.Message);
        end;
      end;
  end;
  Screen.Cursor := crDefault;
end;


procedure TfmZL.btOdpojitClick(Sender: TObject);
var
  RadekDo,
  Radek,
  NotesId,
  SinId: integer;
  SQLStr: AnsiString;
begin
  Screen.Cursor := crHourGlass;
  with asgPohledavky, DesU.qrZakos do try
    if RowCount > 2 then RadekDo := RowCount - 2 else RadekDo := 1;
    Radek := Trunc(ColumnSum (7, 1, RadekDo));             // po�et vybran�ch ��dk�
    if Application.MessageBox(PChar(Format('Opravdu odpojit %d z�kazn�k�?', [Radek])),
     'Pozor', MB_ICONQUESTION + MB_YESNO + MB_DEFBUTTON1) = IDNO then Exit;
    System.Append(F);
    Writeln (F);
    Writeln (F, FormatDateTime('dd.mm.yy hh:nn  ', Now) + 'Odpojen�:');
    CloseFile(F);
// 30.12.14
    idHTTP.Request.Clear;
    idHTTP.Request.BasicAuthentication := True;
    idHTTP.Request.Username := 'NFO';
    idHTTP.Request.Password := 'NFO2014';
    for Radek := 1 to RadekDo do
      if Ints[7, Radek] = 1 then try
        if idHTTP.Get(Format('http://aplikace.eurosignal.cz/api/contracts/change_state?number=%s&state=disconnected', [Cells[5, Radek]])) <> 'OK' then
          if Application.MessageBox(PChar('Odpojen� se nepoda�ilo ulo�it do datab�ze. Pokra�ovat?'),
           'Pozor', MB_ICONQUESTION + MB_YESNO + MB_DEFBUTTON1) = IDNO then Exit
          else Continue;
{
    for Radek := 1 to RadekDo do
      if Ints[7, Radek] = 1 then try
        Close;
        SQL.Text := 'SELECT MAX(Id) FROM notes';
        Open;
        NotesId := Fields[0].AsInteger + 1;
        Close;
        SQLStr := 'INSERT INTO notes ('
        + ' Id,'
        + ' ParenTable_id,'
        + ' ParenTable_type,'
        + ' Note,'
        + ' User_id,'
        + ' Created_at,'
        + ' Updated_at) VALUES ('
        + IntToStr(NotesId) + ', '
        + Cells[9, Radek] + ', '
        + Ap + 'Customer' + ApC
        + Ap + Format('Smlouva %s - odpojeno programem Nezaplacen� ZL.', [Cells[5, Radek]]) + ApC
        + '2, '                                            // automat
        + Ap + FormatDateTime('yyyy-mm-dd hh:nn:ss', Now) + ApC
        + Ap + FormatDateTime('yyyy-mm-dd hh:nn:ss', Now) + ApZ;
        SQL.Text := SQLStr;
        ExecSQL;
        Inc(NotesId);
        SQLStr := 'INSERT INTO notes ('
        + ' Id,'
        + ' ParenTable_id,'
        + ' ParenTable_type,'
        + ' Note,'
        + ' User_id,'
        + ' Created_at,'
        + ' Updated_at) VALUES ('
        + IntToStr(NotesId) + ', '
        + Cells[10, Radek] + ', '
        + Ap + 'Contract' + ApC
        + Ap + 'Odpojeno programem Nezaplacen� ZL.' + ApC
        + '2, '                                            // automat
        + Ap + FormatDateTime('yyyy-mm-dd hh:nn:ss', Now) + ApC
        + Ap + FormatDateTime('yyyy-mm-dd hh:nn:ss', Now) + ApZ;
        SQL.Text := SQLStr;
        ExecSQL;
        SQLStr := 'UPDATE contracts SET'
        + ' State = ''disconnected'','
        + ' Invoice = 0'
        + ' WHERE Id = ' + Cells[10, Radek];
        SQL.Text := SQLStr;
        ExecSQL;   }
        if (Colors[7, Radek] <> clSilver) then Colors[7, Radek] := clSilver
        else Colors[7, Radek] := clWhite;
        System.Append(F);
        Writeln (F, Format('%s (%s)', [Cells[0, Radek], Cells[5, Radek]]));
        CloseFile(F);
        try                                                  // 7.10.11
          SQL.Text := 'SELECT MAX(Id) FROM sinners';
          Open;
          SinId := Fields[0].AsInteger + 1;
          SQLStr := 'SELECT Invoice_debt, Account_debt, Enforcement_state FROM sinners'
          + ' WHERE Customer_id = ' + Cells[9, Radek]
          + ' AND Id = (SELECT MAX(Id) FROM sinners'
            + ' WHERE Customer_id = ' + Cells[9, Radek] + ')';
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
          + Cells[9, Radek] + ', '
          + Cells[10, Radek] + ', '
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
      except on E: exception do
        if Application.MessageBox(PChar(E.Message + sLineBreak + 'Odpojen� se nepoda�ilo ulo�it do datab�ze. Pokra�ovat?'),
         'Pozor', MB_ICONQUESTION + MB_YESNO + MB_DEFBUTTON1) = IDNO then Exit
        else Continue;
      end;
  finally
    Close;
    Screen.Cursor := crDefault;
  end;
end;






{*********************** akce Input element� **********************************}

procedure TfmZL.asgPohledavkyGetAlignment(Sender: TObject; ARow, ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
  if ARow = 0 then HAlign := taLeftJustify
  else if (ACol in [1..4]) then HAlign := taRightJustify;
end;

procedure TfmZL.asgPohledavkyGetFormat(Sender: TObject; ACol: Integer; var AStyle: TSortStyle; var aPrefix, aSuffix: String);
begin
  if ACol in [1..2] then AStyle := ssNumeric
  else if ACol in [0, 3..5] then AStyle := ssAlphabetic;
end;

procedure TfmZL.asgPohledavkyCanSort(Sender: TObject; ACol: Integer; var DoSort: Boolean);
begin
  with asgPohledavky do
    if ACol = 7 then DoSort := False
    else if RowCount > 2 then RemoveRows(RowCount-1, 1);
end;

procedure TfmZL.asgPohledavkyCanEditCell(Sender: TObject; ARow, ACol: Integer; var CanEdit: Boolean);
begin
  CanEdit := (ACol in [7, 8]) or (ACol = 13);
end;

procedure TfmZL.asgPohledavkyClickSort(Sender: TObject; ACol: Integer);
begin
  with asgPohledavky do
    if (ACol <> 7) and (RowCount > 2) then begin
      RowCount := RowCount + 1;
      Cells[0, RowCount-1] := 'Celkem';
      Floats[1, RowCount-1] := ColumnSum(1, 1, RowCount-2);
      Cells[2, RowCount-1] := Format('%d', [Trunc(ColumnSum(2, 1, RowCount-2))]);
    end;
end;

procedure TfmZL.asgPohledavkyClickCell(Sender: TObject; ARow, ACol: Integer);
var
  Radek: integer;
begin
  if (ARow = 0) and (ACol = 7) then with asgPohledavky do
    if ColumnSum(7, 1, RowCount-2) = 0 then for Radek := 1 to RowCount-2 do Ints[7, Radek] := 1
    else for Radek := 1 to RowCount-2 do Ints[7, Radek] := 0;
end;

procedure TfmZL.asgPohledavkyDblClickCell(Sender: TObject; ARow, ACol: Integer);
begin
  Radek := ARow;
  fmZLDetail.ShowModal;
end;

procedure TfmZL.btKonecClick(Sender: TObject);
begin
  Close;
end;

procedure TfmZL.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  DesU.qrAbra.Close;
end;

end.
