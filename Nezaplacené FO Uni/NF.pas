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

// 2017-09: odstran�n� omezov�n�, odstran�n� ABAK, p�id�n� zas�l�n� SMS

unit NF;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Forms, Controls, StdCtrls, ExtCtrls, ComCtrls, ComObj, Math,
  Mask, AdvCombo, AdvEdit, Dialogs, Grids, BaseGrid, AdvGrid, DB,
  //IBDatabase, IBCustomDataSet, IBQuery,
  DateUtils, IniFiles,
  //rxToolEdit,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, ZConnection, IdBaseComponent, IdComponent, IdTCPConnection,
  IdTCPClient, IdSMTP, IdHTTP, IdMessage, IdMessageClient, IdText, IdMessageParts,
  IdAntiFreezeBase, IdAntiFreeze, ZAbstractConnection, AdvObj, IdIOHandler,
  IdIOHandlerSocket, IdSSLOpenSSL, IdExplicitTLSClientServerBase, IdSMTPBase,
  AdvDateTimePicker;

type
  TfmMain = class(TForm)
    idMessage: TIdMessage;
    idSMTP: TIdSMTP;
    dlgExport: TSaveDialog;
    pnBottom: TPanel;
    mmMail: TMemo;
    IdAntiFreeze1: TIdAntiFreeze;
    idHTTP: TIdHTTP;
    pnMain: TPanel;
    lbDo: TLabel;
    lbOd: TLabel;
    btKonec: TButton;
    acbRada: TAdvComboBox;
    btVyber: TButton;
    btExport: TButton;
    aedPocetOd: TAdvEdit;
    btMail: TButton;
    acbDruhSmlouvy: TAdvComboBox;
    cbCast: TCheckBox;
    aedPocetDo: TAdvEdit;
    btOdpojit: TButton;
    rgText: TRadioGroup;
    btSMS: TButton;
    deDatumOd: TDateTimePicker;
    deDatumDo: TDateTimePicker;
    asgPohledavky: TAdvStringGrid;
    procedure FormShow(Sender: TObject);
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
    procedure btOdpojitClick(Sender: TObject);
    procedure btKonecClick(Sender: TObject);
    procedure rgTextClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure vyberNestandardCisla();
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
  fmMain: TfmMain;

implementation

uses DesUtils, NF2D;

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

// 30.12.14 adres�� pro logy
  LogDir := DesU.PROGRAM_PATH + '\Nezaplacen� FO - logy';
  if not DirectoryExists(LogDir) then CreateDir(LogDir);
// vytvo�en� logfile, pokud neexistuje - 20.11.14 do jm�na p�id�no datum - 8.4.15 jen rok a m�s�c
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
    Cells[1, 0] := ' k�d';
    Cells[2, 0] := 'po�et FO';
    Cells[3, 0] := 'pohled�vky';
    Cells[4, 0] := '311-325';
    Cells[5, 0] := 'druh';
    Cells[6, 0] := 'smlouva';
    Cells[8, 0] := ' mail';
    Cells[9, 0] := 'telefon';
    Cells[12, 0] := 'mobil SMS';
    }
    ColWidths[7] := 18;                // checkmark
    ColWidths[10] := 0;                 // Cu.Id
    ColWidths[11] := 0;                 // C.Id
    CheckFalse := '0';
    CheckTrue := '1';
  end;
  //aedPocetOd.Text := '1';
  //aedPocetDo.Text := '12';

  MailText := Format('V�en� pane, v�en� pan�,' + sLineBreak
  + 'dovolujeme si V�s upozornit, �e je %d dn� po splatnosti pravideln� m�s��n� faktury za p�ipojen� k internetu '
  + 'a st�le od V�s postr�d�me jej� �hradu.' + sLineBreak
  + 'I kdy� pen�le za %d dn� zpo�d�n� platby ve v��i %.1f%% je v tuto chv�li zanedbateln� a jist� ho na V�s nebudeme vym�hat, '
  + 'pot�ilo by n�s, kdybyste dlu�nou ��stku co nejd��ve uhradili.', [DayOf(Date) - 10, DayOf(Date) - 10, (DayOf(Date) - 10) * 0.3]);

  mmMail.Text := MailText;

  acbRada.Clear;
  acbRada.Items.Add('%');
  with DesU.qrAbra do begin
    SQL.Text := 'SELECT Code FROM DocQueues'                 // �ady faktur
    + ' WHERE DocumentType = ''03'''
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
    //SQL.Text := 'SET CHARACTER SET cp1250';                // p�eklad z UTF-8   //*HW asi neni potreba
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




procedure TfmMain.rgTextClick(Sender: TObject);
begin
  case rgText.ItemIndex of
   0: MailText := Format('V�en� pane, v�en� pan�,' + sLineBreak
      + 'dovolujeme si V�s upozornit, �e je %d dn� po splatnosti pravideln� m�s��n� faktury za p�ipojen� k internetu '
      + 'a st�le od V�s postr�d�me jej� �hradu.' + sLineBreak
      + 'I kdy� pen�le za %d dn� zpo�d�n� platby ve v��i %.1f%% je v tuto chv�li zanedbateln� a jist� ho na V�s nebudeme vym�hat, '
      + 'pot�ilo by n�s, kdybyste dlu�nou ��stku co nejd��ve uhradili.', [DayOf(Date) - 10, DayOf(Date) - 10, (DayOf(Date) - 10) * 0.3]);
   1: MailText := 'V�en� pane, v�en� pan�,' + sLineBreak
      + 'upozor�ujeme V�s, �e V� dluh za p�ipojen� k internetu dos�hl dvou m�s��n�ch plateb, nebo tuto ��stku ji� p�es�hl. '
      + 'V brzk� dob� proto m��ete o�ek�vat sankce v podob� sn�en� rychlosti p�ipojen�.' + sLineBreak
      + 'Bli��� informace m��ete v pracovn� dny (9-16 h.) z�skat na ��sle 227 031 807, nebo kdykoli na sv�m z�kaznick�m ��tu na www.eurosignal.cz';
   2: MailText := 'V�en� pane, v�en� pan�,' + sLineBreak
      + 'upozor�ujeme V�s, �e V� dluh za p�ipojen� k internetu dos�hl t�� m�s��n�ch plateb, nebo tuto ��stku ji� p�es�hl. '
      + 'V nejbli��� dob� V�m proto bude p�eru�eno p�ipojen� k internetu.' + sLineBreak
      + 'Dal�� informace m��ete z�skat v pracovn� dny (9-16 h.) na ��sle 227 031 807, nebo kdykoli na sv�m z�kaznick�m ��tu na www.eurosignal.cz';
  end;
  mmMail.Text := MailText;
end;

procedure TfmMain.vyberNestandardCisla();
var
  SQLStr: string;
  fmWidth,
  Radek: integer;
begin
  Screen.Cursor := crHourGlass;
  with DesU.qrZakos, asgPohledavky do try
    Radek := 0;
    ClearNormalCells;

  //if dbMain.Connected then with qrMain do begin
    Close;
    SQLStr := 'SELECT DISTINCT Cu.Id AS CuId, C.Id AS CId, Postal_mail, Phone, vip, Number, State, Tag_Id'
    + ' FROM contracts C'
    + ' JOIN customers Cu ON Cu.Id = C.Customer_Id'
    + ' LEFT JOIN contracts_tags CT ON CT.Contract_Id = C.Id'      // 17.12.15


    //+ ' WHERE  phone <> '''' and phone not like ''7%'' and phone not like ''6%'' ';
    + ' WHERE  email like ''sef%'' ';
    //+ ' WHERE  phone like ''%+420%''';

    if acbDruhSmlouvy.Text <> '%' then SQLStr := SQLStr + ' AND State = ' + Ap + acbDruhSmlouvy.Text + Ap;
    SQLStr := SQLStr + ' LIMIT 800 ';

    SQL.Text := SQLStr;
    Open;

// 28.4.11 barvi�ky

    while not EOF do begin
      Inc(Radek);                                    // 22.9.11 a ka�d� m� sv�j ��dek
      RowCount := Radek + 1;
      Cells[5, Radek] := FieldByName('State').AsString;
      Cells[6, Radek] := FieldByName('Number').AsString;
  // 17.12.15 p�ipojen� v jin� s�ti
      if FieldByName('Tag_Id').AsInteger in [20, 21, 25, 26, 27, 30] then Colors[6, Radek] := clRed;
      Cells[8, Radek] := FieldByName('Postal_mail').AsString;
      Cells[9, Radek] := FieldByName('Phone').AsString;
      Cells[10, Radek] := FieldByName('CuId').AsString;
      Cells[11, Radek] := FieldByName('CId').AsString;
      if Pos('@', FieldByName('Postal_mail').AsString) > 0 then Ints[7, Radek] := 1;


      AddCheckBox(7, Radek, True, True);
      Ints[7, Radek] := 0;

      Cells[3, Radek] := destilujTelCislo(FieldByName('Phone').AsString);

      Next;
    end;

    Application.ProcessMessages;


// �prava zobrazen�
    AutoSize := True;
    ColWidths[7] := 18;
    ColWidths[10] := 0;                 // Cu.Id
    ColWidths[11] := 0;                 // C.Id

    fmWidth := 0;
    for Radek := 0 to ColCount-1 do fmWidth := fmWidth +  ColWidths[Radek];
    fmMain.Width := fmWidth + 120;
    fmMain.Height := Min(800, RowCount * 19 + 164);
  finally
//    ShowMessage(TimeToStr(Time - t));
    Screen.Cursor := crDefault;
  end;
end;


procedure TfmMain.btVyberClick(Sender: TObject);
var
  SQLStr: AnsiString;
  fmWidth,
  Radek: integer;
begin
  Screen.Cursor := crHourGlass;
  with DesU.qrAbra, asgPohledavky do try
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
      if (FieldByName('Zakaznik').AsString <> '') and (FieldByName('Count').AsInteger >= StrToInt(aedPocetOd.Text))
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

        with DesU.qrAbra2 do begin
// 24.11.13 v�echny Firm_Id pro Abrak�d firmy
          SQLStr := 'SELECT * FROM DE$_CODE_TO_FIRM_ID (' + Ap + Cells[1, Radek] + ApZ;
          SQL.Text := SQLStr;
          Open;
          Floats[4, Radek] := 0;
// a saldo pro v�echny Firm_Id

          while not EOF do with DesU.qrAbra3 do begin
            DesU.qrAbra3.Close;
            SQLStr := 'SELECT Ucet311 + Ucet325 FROM DE$_Firm_Totals (' + Ap + DesU.qrAbra2.Fields[0].AsString + ApC + FloatToStr(Date) + ')';
            DesU.qrAbra3.SQL.Text := SQLStr;
            DesU.qrAbra3.Open;
            Floats[4, Radek] := Floats[4, Radek] - Fields[0].AsFloat;
            DesU.qrAbra2.Next;
          end; // while not EOF do with qrAbra3

        end;  // with qrAbra2

// v�echny smlouvy pro jeden abrak�d

        with DesU.qrZakos do begin
          Close;
          SQLStr := 'SELECT DISTINCT Cu.Id AS CuId, C.Id AS CId, Postal_mail, Phone, vip, Number, State, Tag_Id'
          + ' FROM contracts C'
          + ' JOIN customers Cu ON Cu.Id = C.Customer_Id'
          + ' LEFT JOIN contracts_tags CT ON CT.Contract_Id = C.Id'      // 17.12.15

//          + ' AND Tariff_Id <> 2'                          // 12.4.11 ne EP-Basic

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
          Cells[12, Radek] := destilujMobilCislo(FieldByName('Phone').AsString);
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
//    ShowMessage(TimeToStr(Time - t));
    Screen.Cursor := crDefault;
  end;
end;


procedure TfmMain.btExportClick(Sender: TObject);
begin

  with dlgExport do begin
    DefaultExt := '.csv';
    Filter := 'csv|*.csv';
    asgPohledavky.QuoteEmptyCells := True;
    if Execute then asgPohledavky.SaveToCSV(dlgExport.FileName);
  end;

  // if dlgExport.Execute then asgPohledavky.SaveToXLS(dlgExport.FileName); //takto to bylo pro ABAK, mozna se muze nekdy hodit
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
  idSMTP.Host :=  DesU.getIniValue('Mail', 'SMTPServer');
  idSMTP.Username := DesU.getIniValue('Mail', 'SMTPLogin');
  idSMTP.Password := DesU.getIniValue('Mail', 'SMTPPW');

  with asgPohledavky, idMessage do begin
    if RowCount > 2 then RadekDo := RowCount - 2 else RadekDo := 1;

    Radek := Trunc(ColumnSum (7, 1, RadekDo));             // po�et vybran�ch ��dk�
    if Application.MessageBox(PChar(Format('Opravdu poslat %d e-mail�?', [Radek])),
     'Pozor', MB_ICONQUESTION + MB_YESNO + MB_DEFBUTTON1) = IDNO then Exit;

    Append(F);
    Writeln (F, FormatDateTime(sLineBreak + 'dd.mm.yy hh:nn  ', Now) + 'Odesl�n� zpr�vy: ' + sLineBreak + mmMail.Text + sLineBreak);
    CloseFile(F);
    for Radek := 1 to RadekDo do
      if Ints[7, Radek] = 1 then begin
        Clear;

        From.Address := 'kontrola@eurosignal.cz';
        //CCList.Add.Address := 'a@a.cz';

        MailStr := Cells[8, Radek];
        MailStr := StringReplace(MailStr, ',', ';', [rfReplaceAll]);    // ��rky za st�edn�ky
        while Pos(';', MailStr) > 0 do begin
          Recipients.Add.Address := Trim(Copy(MailStr, 1, Pos(';', MailStr)-1));
          MailStr := Copy(MailStr, Pos(';', MailStr)+1, Length(MailStr));
        end;
        Recipients.Add.Address := Trim(MailStr);

        Subject := 'Kontrola �lu�ou�k�ch plateb smlouvy ' + Cells[6, Radek];

        with TIdText.Create(idMessage.MessageParts, nil) do begin
          Body.Text := StringReplace(mmMail.Text, '%%%', IntToStr(Round(Floats[4, Radek])), [rfIgnoreCase])
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
        Body.Text := StringReplace(mmMail.Text, '%%%', IntToStr(Round(Floats[4, Radek])), [rfIgnoreCase]);
        Body.Add(' ');
        Body.Add('S pozdravem');
        Body.Add(' ');
        Body.Add('V� Eurosignal');
        Body.Add(' ');
        Body.Add('Na tuto zpr�vu neodpov�dejte, byla generov�na automaticky.');
        }

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


procedure TfmMain.btSMSClick(Sender: TObject);
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

        smsText := StringReplace(mmMail.Text, '%%%', IntToStr(Round(Floats[4, Radek])), [rfIgnoreCase]);
        callResult := DesU.sendGodsSms(Cells[12, Radek], smsText);

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
          Writeln (F, Format('%s (%s)  -  %s  -  %s', [Cells[0, Radek], Cells[6, Radek], Cells[12, Radek], callResult]));
          CloseFile(F);
        except on E: exception do
          ShowMessage('SMS se nepoda�ilo ulo�it do tabulky communications: ' + E.Message);
        end;
      end;
  end;

  Screen.Cursor := crDefault;
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
  with asgPohledavky, DesU.qrZakos do try
    if RowCount > 2 then RadekDo := RowCount - 2 else RadekDo := 1;
    Radek := Trunc(ColumnSum (7, 1, RadekDo));             // po�et vybran�ch ��dk�
// 24.11.13 potvrzen�
    if Application.MessageBox(PChar(Format('Opravdu odpojit %d z�kazn�k�?', [Radek])),
     'Pozor', MB_ICONQUESTION + MB_YESNO + MB_DEFBUTTON1) = IDNO then Exit;
    System.Append(F);
    Writeln (F, FormatDateTime(sLineBreak + 'dd.mm.yy hh:nn  ', Now) + 'Odpojen�:' + sLineBreak);
    CloseFile(F);
// 26.6.14
    idHTTP.Request.Clear;
    idHTTP.Request.BasicAuthentication := True;
//    idHTTP.Request.UserAgent := 'Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.0; SLCC1)';

    idHTTP.Request.Username := 'NFO';
    idHTTP.Request.Password := 'NFO2014';
    URL := 'http://aplikace.eurosignal.cz';

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

procedure TfmMain.asgPohledavkyDblClickCell(Sender: TObject; ARow, ACol: Integer);
begin
  Radek := ARow;
  fmDetail.ShowModal;
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
  CanEdit := (ACol in [6..9]) or (ACol = 12);
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

procedure TfmMain.btKonecClick(Sender: TObject);
begin
  Close;
end;

procedure TfmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  DesU.qrAbra.Close;
end;

end.
