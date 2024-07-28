// 31.10.2008 další verze výpisu nezaplacených faktur
// 5.1.2009 do tabulky pøidány sloupce Kód a Smlouva, ve kterých se vypisuje kód firmy z Abry a pomocí nìj
// èíslo smlouvy z databáze DES (tabulka Smlouvy)
// 18.11.09 rozesílání varovných mailù
// 30.3.10 rozesílání mailù pomocí idSMTP z Indy - umožòuje pøihlášení k serveru
// 12.4.11 zobrazení všech smluv zákazníka, potlaèení výbìru kreditních smluv na VoIP
// 22.4.11 výbìr textu mailu, tisk dopisù pro uživatele bez mailu
// 28.4.11 kontrola dlužné èástky s 311-325
// 19.7.11 hromadné omezování a odpojování
// 20.1.14 potvrzení pro omezování a odpojování, ukládání zprávy do tabulky Notes
// 29.1.14 logování mailù a omezených nebo odpojených zákazníkù
// 30.12. omezování a odpojování pomocí API od iQuestu
// 25.4.2018 https místo http
// 29.8.2019 úpravy vzhledu
// 18.2.2021 úprava posílání mailù (Indy10), vráceno omezování
// 18.1.2022 úprava pro netypické smlouvy (nemají Tariff_Id)
// 4.3.2023 maily s TLS
// 28.3.2024 oprava posílání mailù

unit NZL;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Forms, Controls, StdCtrls, ExtCtrls, ComCtrls, ComObj, Mask, AdvCombo,
  AdvEdit, AdvObj, Dialogs, Grids, BaseGrid, AdvGrid, DB, DateUtils, IniFiles, rxToolEdit, Math,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, ZAbstractConnection, ZConnection,
  IdText, IdSMTP, IdSMTPBase, IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdMessageClient, IdMessage, IdMessageParts, IdIOHandler,
  IdIOHandlerSocket, IdIOHandlerStack, IdSSL, IdSSLOpenSSL, IdExplicitTLSClientServerBase, IdHTTP, IdAntiFreezeBase, IdAntiFreeze,
  AdvUtil, IdAuthentication;

type
  TfmZL = class(TForm)
    dlgExport: TSaveDialog;
    pnBottom: TPanel;
    mmMail: TMemo;
    idHTTP: TIdHTTP;
    IdAntiFreeze: TIdAntiFreeze;
    idMessage: TIdMessage;
    idSMTP: TIdSMTP;
    idSSLHandler: TIdSSLIOHandlerSocketOpenSSL;
    pnMain: TPanel;
    lbDo: TLabel;
    lbOd: TLabel;
    deDatumOd: TDateEdit;
    deDatumDo: TDateEdit;
    btKonec: TButton;
    asgPohledavky: TAdvStringGrid;
    acbRada: TAdvComboBox;
    btVyber: TButton;
    btExport: TButton;
    btMail: TButton;
    acbDruhSmlouvy: TAdvComboBox;
    cbCast: TCheckBox;
    btOdpojit: TButton;
    rgText: TRadioGroup;
    btSMS: TButton;
    btOmezit: TButton;
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
    procedure btOmezitClick(Sender: TObject);
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
  FIIni: TIniFile;
  LogDir,
  LogFileName: AnsiString;
begin
  deDatumDo.Date := EndOfTheMonth(IncMonth(Now, -1));
  deDatumOd.Date := StartOfTheMonth(IncYear(Now, -1));
// 30.12.14 adresáø pro logy
  LogDir := ExtractFilePath(ParamStr(0)) + '\Logy\Nezaplacené ZL';
  if not DirectoryExists(LogDir) then CreateDir(LogDir);
// vytvoøení logfile, pokud neexistuje - 27.7.16 ve jménì jen rok a mìsíc
  LogFileName := LogDir + FormatDateTime('\yyyy.mm".log"', Date);
  if not FileExists(LogFileName) then begin
    FileHandle := FileCreate(LogFileName);
    FileClose(FileHandle);
  end;
  AssignFile(F, LogFileName);

  with asgPohledavky do begin
    Clear;
    Cells[0, 0] := ' zákazník';
    Cells[1, 0] := 'poèet ZL';
    Cells[2, 0] := 'dluh ZL';
    Cells[3, 0] := 'celkem';
    ColWidths[4] := 0;                  // Firms.ID
    Cells[5, 0] := 'stav';
    Cells[6, 0] := 'smlouva';
    ColWidths[7] := 0;                  // Firms.Code
    ColWidths[8] := 18;                 // checkmark
    Cells[9, 0] := ' mail';
    ColWidths[10] := 0;                 // Cu.Id
    ColWidths[11] := 0;                 // C.Id
    Cells[12, 0] := 'dny';
    Cells[13, 0] := 'mobil SMS';
    CheckFalse := '0';
    CheckTrue := '1';
  end;
  MailText := 'Vážený pane, vážená paní,' + #13#10
  + 'dovolujeme si Vás upozornit, že je XXX dní po splatnosti zálohy na pøipojení k internetu a stále od Vás postrádáme její úplnou úhradu.'
  + ' Dlužná èástka èiní YYY Kè.' + #13#10
  + 'Potìšilo by nás, kdybyste èástku zálohy co nejdøíve uhradili a my Vám nemuseli omezovat poskytované služby.';
  mmMail.Text := MailText;

  with DesU.qrAbra do begin
    SQL.Text := 'SELECT Code FROM DocQueues'                 // øady faktur
    + ' WHERE DocumentType = ''10'''
    + ' AND Hidden = ''N'''
    + ' ORDER BY Code';
    acbRada.Clear;
    acbRada.Items.Add('%');
    Open;
    while not EOF do begin
      acbRada.Items.Add(FieldByName('Code').AsString);
      Next;
    end;
  end;
  acbRada.ItemIndex := 2;

  with DesU.qrZakos do begin
    SQL.Text := 'SELECT DISTINCT State FROM contracts'       // stav smlouvy
    + ' ORDER BY State';
    acbDruhSmlouvy.Clear;
    acbDruhSmlouvy.Items.Add('%');
    Open;
    while not EOF do begin
      acbDruhSmlouvy.Items.Add(FieldByName('State').AsString);
      Next;
    end;
  end;
  acbDruhSmlouvy.ItemIndex := 1;
end;

procedure TfmZL.asgPohledavkyGetAlignment(Sender: TObject; ARow, ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
  if ARow = 0 then HAlign := taLeftJustify
  else if (ACol in [1, 6]) then HAlign := taRightJustify
  else if (ACol in [2..5, 12, 13]) then HAlign := taCenter;
end;

procedure TfmZL.asgPohledavkyGetFormat(Sender: TObject; ACol: Integer; var AStyle: TSortStyle; var aPrefix, aSuffix: String);
begin
  if ACol in [1..3] then AStyle := ssNumeric
  else if ACol in [0, 4..6] then AStyle := ssAlphabetic;
end;

procedure TfmZL.asgPohledavkyCanSort(Sender: TObject; ACol: Integer; var DoSort: Boolean);
begin
  with asgPohledavky do
    if ACol = 8 then DoSort := False
    else if RowCount > 2 then RemoveRows(RowCount-1, 1);
end;

procedure TfmZL.asgPohledavkyCanEditCell(Sender: TObject; ARow, ACol: Integer; var CanEdit: Boolean);
begin
  CanEdit := ACol in [6, 8, 9, 13];
end;

procedure TfmZL.asgPohledavkyClickSort(Sender: TObject; ACol: Integer);
begin
  with asgPohledavky do
    if (ACol <> 8) and (RowCount > 2) then begin
      RowCount := RowCount + 1;
      Cells[0, RowCount-1] := 'Celkem';
      Cells[1, RowCount-1] := Format('%d', [Trunc(ColumnSum(1, 1, RowCount-2))]);
      Floats[2, RowCount-1] := ColumnSum(2, 1, RowCount-2);
      Floats[3, RowCount-1] := ColumnSum(3, 1, RowCount-2);
    end;
end;

procedure TfmZL.asgPohledavkyClickCell(Sender: TObject; ARow, ACol: Integer);
var
  Radek: integer;
begin
  if (ARow = 0) and (ACol = 8) then with asgPohledavky do
    if ColumnSum(8, 1, RowCount-2) = 0 then for Radek := 1 to RowCount-2 do Ints[8, Radek] := 1
    else for Radek := 1 to RowCount-2 do Ints[8, Radek] := 0;
end;

procedure TfmZL.asgPohledavkyDblClickCell(Sender: TObject; ARow, ACol: Integer);
begin
  Radek := ARow;
  fmZLDetail.ShowModal;
end;

procedure TfmZL.rgTextClick(Sender: TObject);
begin
  case rgText.ItemIndex of
   0: MailText := 'Vážený pane, vážená paní,' + #13#10
      + 'dovolujeme si Vás upozornit, že je XXX dní po splatnosti zálohy na pøipojení k internetu a stále od Vás postrádáme její úplnou úhradu.'
      + ' Dlužná èástka èiní YYY Kè.' + #13#10
      + 'Potìšilo by nás, kdybyste èástku zálohy co nejdøíve uhradili a my Vám nemuseli omezovat poskytované služby.';
   1: MailText := 'Vážený pane, vážená paní,' + #13#10
      + 'upozoròujeme Vás, že je XXX dní po splatnosti zálohy na pøipojení k internetu a stále od Vás postrádáme její úplnou úhradu.'
      + ' Dlužná èástka èiní YYY Kè.' + #13#10
      + 'V nejbližší dobì Vám proto bude pøipojení k internetu pøerušeno.' + #13#10
      + 'Potøebujete-li ke svým platbám bližší informace, mùžete je v pracovní dny (9-16 h.) získat na èísle 227 031 807,'
      + ' nebo kdykoli na svém zákaznickém úètu na www.eurosignal.cz.';
   2: MailText := 'Je nám líto, ale Vaše závazky zùstávají neuhrazené, proto Vám bude pøipojení k internetu pøerušeno.'
      + ' Pro obnovení mùžete volat Eurosignal, tel. 227031807.';
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
    + ' INNER JOIN DocQueues DQ ON IDI.DocQueue_ID = DQ.Id'               // øada dokladù
    + ' INNER JOIN Firms F ON IDI.Firm_ID = F.Id'                         // zákazníci
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
        AddCheckBox(8, Radek, True, True);
        Ints[8, Radek] := 0;
        Cells[0, Radek] := FieldByName('Zakaznik').AsString;
        Floats[2, Radek] := FieldByName('Castka').AsFloat;
        Ints[1, Radek] := FieldByName('Count').AsInteger;
        Cells[4, Radek] := FieldByName('Id').AsString;
        Cells[7, Radek] := FieldByName('Kod').AsString;
        Ints[12, Radek] := Trunc(Date) - FieldByName('Datum').AsInteger;
        Application.ProcessMessages;
        Row := Radek;
        if DesU.dbZakos.Connected then with DesU.qrZakos do begin
          Close;
          SQLStr := 'SELECT DISTINCT Cu.Id AS CuId, C.Id AS CId, Postal_mail, Phone, vip, Number, State, Tag_Id'
          + ' FROM contracts C'
          + ' JOIN customers Cu ON Cu.Id = C.Customer_Id'
          + ' LEFT JOIN contracts_tags CT ON CT.Contract_Id = C.Id'      // 17.12.15
//          + ' WHERE Tariff_Id <> 2'                          // 12.4.11 ne EP-Basic
          + ' WHERE (Tariff_Id <> 2 OR Tariff_Id IS NULL)'                 // 18.1.22
          + ' AND Abra_Code = ' + Ap + Cells[7, Radek] + Ap;
//          + ' AND Activated_at <= ' + Ap + FormatDateTime('yyyy-mm-dd', deDatumDo.Date) + Ap;    // 18.1.22
          if acbDruhSmlouvy.Text <> '%' then SQLStr := SQLStr + ' AND State = ' + Ap + acbDruhSmlouvy.Text + Ap;
          SQL.Text := SQLStr;
          Open;
          if DesU.qrZakos.RecordCount > 0 then begin               // 6.9.22
            Cells[5, Radek] := FieldByName('State').AsString;
            Cells[6, Radek] := FieldByName('Number').AsString;
// 29.7.16 pøipojení v jiné síti, vip
            if FieldByName('Tag_Id').AsInteger in [20, 21, 25, 26, 27, 30] then Colors[6, Radek] := clRed;
            if FieldByName('vip').AsInteger > 0 then Colors[0, Radek] := clSilver;
// 27.2.2024  if (Pos('active', FieldByName('State').AsString) > 0) and (Pos('@', FieldByName('Postal_mail').AsString) > 0) then begin
            if ((Pos('active', FieldByName('State').AsString) > 0) or (Pos('restricted', FieldByName('State').AsString) > 0))
             and (Pos('@', FieldByName('Postal_mail').AsString) > 0) then begin
              Ints[8, Radek] := 1;
              Cells[9, Radek] := FieldByName('Postal_mail').AsString;
            end else Cells[9, Radek] := FieldByName('Phone').AsString;
            Cells[10, Radek] := FieldByName('CuId').AsString;
            Cells[11, Radek] := FieldByName('CId').AsString;
// 26.10.18 pro SMS
            Cells[13, Radek] := destilujMobilCislo(FieldByName('Phone').AsString);
            if not EOF then Next;                            // 29.7.16 mùže být víc smluv
            while not EOF do begin
              if (Cells[6, Radek] <> FieldByName('Number').AsString) then begin    // 18.1.17 u smlouvy mùže být víc tagù
                Inc(Radek);                                  // 29.7.16 každá smlouva má svùj øádek
                RowCount := Radek + 1;
                Cells[5, Radek] := FieldByName('State').AsString;
                Cells[6, Radek] := FieldByName('Number').AsString;
                if FieldByName('Tag_Id').AsInteger in [20, 21, 25, 26, 27, 30] then Colors[6, Radek] := clRed;    // je v jiné síti
                AddCheckBox(8, Radek, True, True);
                Ints[8, Radek] := 0;
              end else if FieldByName('Tag_Id').AsInteger in [20, 21, 25, 26, 27, 30] then Colors[6, Radek] := clRed;
              Next;
            end;
          end else begin  // if qrMain.RecordCount > 0       // 6.9.22
//          if qrMain.RecordCount = 0 then begin
            asgPohledavky.ClearRows(Radek, 1);
            Dec(Radek);
            RowCount := Radek + 1;
          end;  // if qrMain.RecordCount < 0
        end;  // with qrMain
      end;  // if FieldByName('Zakaznik').AsString <> ''
      Next;
    end;

// 1.9.2022 i ostatní pohledávky
    for Radek := 1 to RowCount-1 do begin
      DesU.qrAbra.Close;
// všechny Firm_Id pro Abrakód firmy
      SQLStr := 'SELECT * FROM DE$_CODE_TO_FIRM_ID (' + Ap + Cells[7, Radek] + ApZ;
      SQL.Text := SQLStr;
      Open;
      Floats[3, Radek] := Floats[2, Radek];
// a saldo pro všechny Firm_Id
      while not EOF do with DesU.qrAbra2 do begin
        Close;                                                           // Firm_ID
        SQLStr := 'SELECT Ucet311 + Ucet325 FROM DE$_Firm_Totals (' + Ap + DesU.qrAbra.Fields[0].AsString + ApC + FloatToStr(Date) + ')';
        SQL.Text := SQLStr;
        Open;
        Floats[3, Radek] := Floats[3, Radek] - Fields[0].AsFloat;
        DesU.qrAbra.Next;
      end; // while not EOF
    end;

// úprava zobrazení
    AutoSize := True;
    ColWidths[4] := 0;                 // Firms.ID
    ColWidths[7] := 0;                 // Firms.Code
    ColWidths[8] := 18;
    ColWidths[10] := 0;                 // customers.Id
    ColWidths[11] := 0;                // contracts.Id
    if RowCount > 2 then begin
      Inc(Radek);
      RowCount := Radek + 1;
      Cells[0, RowCount-1] := 'Celkem';
      Cells[1, RowCount-1] := Format('%d', [Trunc(ColumnSum(1, 1, RowCount-2))]);
      Floats[2, RowCount-1] := ColumnSum(2, 1, RowCount-2);
      Floats[3, RowCount-1] := ColumnSum(3, 1, RowCount-2);
    end;
    fmWidth := 0;
    for Radek := 0 to ColCount-1 do fmWidth := fmWidth +  ColWidths[Radek];
    fmZL.ClientWidth := fmWidth + 122;
    fmZL.ClientHeight := Min(RowCount * 18 + 120, Screen.Height - 100);
    fmZL.Top := Max(0, Round((Screen.Height - fmZL.Height) / 2));

  finally
    DesU.qrZakos.Close;
    DesU.qrAbra.Close;
    Screen.Cursor := crDefault;
  end;
end;

procedure TfmZL.btExportClick(Sender: TObject);
begin
  if dlgExport.Execute then asgPohledavky.SaveToXLS(dlgExport.FileName);
end;

procedure TfmZL.btMailClick(Sender: TObject);
var
  FIIni: TIniFile;
  RadekDo,
  Radek,
  CommId: integer;
  Zprava,
  MailStr,
  SQLStr: string;
begin
  Screen.Cursor := crHourGlass;
// 4.3.2023 pøesunutí pøihlašovacíc údajù sem umožní mìnit jejich nastavení bez restartování programu
  idSMTP.Host :=  DesU.getIniValue('Mail', 'SMTPServer');
  idSMTP.HeloName := idSMTP.Host;
  idSMTP.Username := DesU.getIniValue('Mail', 'SMTPLogin');
  idSMTP.Password := DesU.getIniValue('Mail', 'SMTPPW');

  with asgPohledavky, idMessage do try
    if RowCount > 2 then RadekDo := RowCount - 2 else RadekDo := 1;

    Radek := Trunc(ColumnSum (8, 1, RadekDo));             // poèet vybraných øádkù
    if Application.MessageBox(PChar(Format('Opravdu poslat %d e-mailù?', [Radek])),
     'Pozor', MB_ICONQUESTION + MB_YESNO + MB_DEFBUTTON1) = IDNO then Exit;
    Append(F);
    Writeln (F, FormatDateTime(sLineBreak + 'dd.mm.yy hh:nn  ', Now) + 'Odeslání zprávy: ' + sLineBreak + mmMail.Text + sLineBreak);
    CloseFile(F);

    for Radek := 1 to RadekDo do
      if Ints[8, Radek] = 1 then begin
        Zprava := StringReplace(mmMail.Text, 'XXX', Cells[12, Radek], []);
        Zprava := StringReplace(Zprava, 'YYY', Cells[3, Radek], []);
        Clear;
        ContentType := 'multipart/mixed';
//        Charset := 'utf-8';
//        From.Address := 'kontrola@eurosignal.cz';
        From.Address := 'uctarna@eurosignal.cz';
        MailStr := Cells[9, Radek];
        MailStr := StringReplace(MailStr, ',', ';', [rfReplaceAll]);    // èárky za støedníky
        while Pos(';', MailStr) > 0 do begin
          Recipients.Add.Address := Trim(Copy(MailStr, 1, Pos(';', MailStr)-1));
          MailStr := Copy(MailStr, Pos(';', MailStr)+1, Length(MailStr));
        end;
        Recipients.Add.Address := Trim(MailStr);
        Subject := 'Kontrola plateb smlouvy ' + Cells[6, Radek];
        with TIdText.Create(idMessage.MessageParts, nil) do begin
          Body.Text := UTF8Encode(Zprava + #13#10#10
          + 'S pozdravem' + #13#10#10
          + 'Váš Eurosignal');
          ContentType := 'text/plain';
          Charset := 'utf-8';
        end;
        try
          if not idSMTP.Connected then idSMTP.Connect;
          idSMTP.Send(idMessage);
        except on E: exception do
          ShowMessage('Mail se nepodaøilo odeslat: ' + E.Message);
        end;
        if (Colors[8, Radek] <> clSilver) then Colors[8, Radek] := clSilver
        else Colors[8, Radek] := clWhite;

// 28.3.24  Zprava := UlozKomunikaci('2', Cells[10, Radek], MailStr);          // v DesU
// 28.7.24       Zprava := UlozKomunikaci('2', Cells[10, Radek], Zprava);          // v DesU
        if not DesU.UlozKomunikaci(2, Ints[10, Radek], Zprava).isOk then
//        if Zprava <> 'ok' then
          ShowMessage('Mail se nepodaøilo uložit do tabulky "communications"')
        else begin
          System.Append(F);
          Writeln (F, Format('%s (%s) email %s', [Cells[0, Radek], Cells[6, Radek], Cells[9, Radek]]));
          Writeln (F, FormatDateTime(#13#10 + 'dd.mm.yy hh:nn  ', Now) + 'Odeslání zprávy: ' + #13#10 + Zprava + #13#10);
          CloseFile(F);
        end;
      end;

  finally
    if idSMTP.Connected then idSMTP.Disconnect;
    Screen.Cursor := crDefault;
  end;
end;

procedure TfmZL.btSMSClick(Sender: TObject);
// 9.8.2022 poskytovatelem služby je SMSbrána
var
  RadekDo,
  Radek,
  CommId: integer;
  smsText,
  callResult,
  Zprava,
  SQLStr: string;
begin
  Screen.Cursor := crHourGlass;

  with asgPohledavky do try
    if RowCount > 2 then RadekDo := RowCount - 2 else RadekDo := 1;
    Radek := Trunc(ColumnSum (8, 1, RadekDo));             // poèet vybraných øádkù
    if Application.MessageBox(PChar(Format('Opravdu poslat %d SMS?', [Radek])),
     'Pozor', MB_ICONQUESTION + MB_YESNO + MB_DEFBUTTON1) = IDNO then begin
      Screen.Cursor := crDefault;
      Exit;
    end;
    Append(F);
    Writeln (F, FormatDateTime(#13#10 + 'dd.mm.yy hh:nn  ', Now) + 'Odeslání SMS zprávy: ' + #13#10 + mmMail.Text + #13#10);
    CloseFile(F);
    for Radek := 1 to RadekDo do
//      if Ints[8, Radek] = 1 then begin
      if (Ints[8, Radek] = 1) and (Cells[13, Radek] <> '') then begin    // 2.12.18 musí být èíslo pro SMS
//        smsText := StringReplace(mmMail.Text, 'xxx', IntToStr(Round(Floats[4, Radek])), [rfIgnoreCase]);
        smsText := StringReplace(mmMail.Text, 'xxx', IntToStr(Round(Floats[3, Radek])), [rfIgnoreCase]);   // 29.4.24 je to jinak než v NF
        callResult := DesU.sendSms(Cells[13, Radek], StringReplace(smsText, ' ', '+', [rfReplaceAll]));
        if (Colors[8, Radek] <> clSilver) then Colors[8, Radek] := clSilver
        else Colors[8, Radek] := clWhite;
// 28.7.2024
        if not DesU.UlozKomunikaci(23, Ints[10, Radek], smsText).isOk then
//        if Zprava <> '' then
          ShowMessage('SMS se nepodaøilo uložit do tabulky "communications"')
        else begin
          System.Append(F);
          Writeln (F, Format('%s (%s) - %s - %s', [Cells[0, Radek], Cells[7, Radek], Cells[13, Radek], callResult]));
          CloseFile(F);
        end;
      end;
    Append(F);
    Writeln (F, #13#10);
    CloseFile(F);
{
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
          Writeln (F, Format('%s (%s) - %s - %s', [Cells[0, Radek], Cells[7, Radek], Cells[13, Radek], callResult]));
          CloseFile(F);
        except on E: exception do
          ShowMessage('SMS se nepodaøilo uložit do tabulky communications: ' + E.Message);
        end;
}
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfmZL.btOmezitClick(Sender: TObject);
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
    Radek := Trunc(ColumnSum (8, 1, RadekDo));             // poèet vybraných øádkù
    if Application.MessageBox(PChar(Format('Opravdu omezit %d zákazníkù?', [Radek])),
     'Pozor', MB_ICONQUESTION + MB_YESNO + MB_DEFBUTTON1) = IDNO then Exit;
    System.Append(F);
    Writeln (F);
    Writeln (F, FormatDateTime('dd.mm.yy hh:nn  ', Now) + 'Omezení rychlosti:');
    CloseFile(F);
// 30.12.14
    idHTTP.IOHandler := idSSLHandler;         // 25.4.2019
    idHTTP.Request.Clear;
    idHTTP.Request.BasicAuthentication := True;
    idHTTP.Request.Username := 'NFO';
    idHTTP.Request.Password := 'NFO2014';
    for Radek := 1 to RadekDo do
      if Ints[8, Radek] = 1 then try
        if idHTTP.Get(Format('https://aplikace.eurosignal.cz/api/contracts/change_state?number=%s&state=restricted', [Cells[6, Radek]])) <> 'OK' then
          if Application.MessageBox(PChar('Omezení se nepodaøilo uložit do databáze. Pokraèovat?'),
           'Pozor', MB_ICONQUESTION + MB_YESNO + MB_DEFBUTTON1) = IDNO then Exit
          else Continue;
        if (Colors[8, Radek] <> clSilver) then Colors[8, Radek] := clSilver
        else Colors[8, Radek] := clWhite;
        System.Append(F);
        Writeln (F, Format('%s (%s)', [Cells[0, Radek], Cells[6, Radek]]));
        CloseFile(F);
        try
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
          + Cells[1, Radek] + ', '                         // poèet ZL
          + Ap + FieldByName('Invoice_debt').AsString + ApC
          + Ap + FieldByName('Account_debt').AsString + ApC
          + Ap + FieldByName('Enforcement_state').AsString + ApC              // stav vymáhání
          + Ap + 'omezení' + ApC
          + Ap + 'program NezaplaceneFaktury' + ApZ;
          Close;
          SQL.Text := SQLStr;
          ExecSQL;
        except on E: exception do
          ShowMessage('Omezení se nepodaøilo uložit do sinners: ' + E.Message);
        end;
      except on E: exception do
        if Application.MessageBox(PChar(E.Message + #13#10 + 'Omezení se nepodaøilo uložit do databáze. Pokraèovat?'),
         'Pozor', MB_ICONQUESTION + MB_YESNO + MB_DEFBUTTON1) = IDNO then Exit
        else Continue;
      end;
  finally
    Close;
    Screen.Cursor := crDefault;
  end;
end;

procedure TfmZL.btOdpojitClick(Sender: TObject);
var
  RadekDo,
  Radek,
  NotesId,
  SinId: integer;
  Invoice_debt,
  Account_debt,
  Enforcement_state,
  SQLStr: AnsiString;
begin
  Screen.Cursor := crHourGlass;
  with asgPohledavky, DesU.qrZakos do try
    if RowCount > 2 then RadekDo := RowCount - 2 else RadekDo := 1;
    Radek := Trunc(ColumnSum (8, 1, RadekDo));             // poèet vybraných øádkù
    if Application.MessageBox(PChar(Format('Opravdu odpojit %d zákazníkù?', [Radek])),
     'Pozor', MB_ICONQUESTION + MB_YESNO + MB_DEFBUTTON1) = IDNO then Exit;
    System.Append(F);
    Writeln (F);
    Writeln (F, FormatDateTime('dd.mm.yy hh:nn  ', Now) + 'Odpojení:');
    CloseFile(F);
// 30.12.14
    idHTTP.IOHandler := idSSLHandler;         // 25.4.2019
    idHTTP.Request.Clear;
    idHTTP.Request.BasicAuthentication := True;
    idHTTP.Request.Username := 'NFO';
    idHTTP.Request.Password := 'NFO2014';
    for Radek := 1 to RadekDo do
      if Ints[8, Radek] = 1 then try
        if idHTTP.Get(Format('https://aplikace.eurosignal.cz/api/contracts/change_state?number=%s&state=disconnected', [Cells[6, Radek]])) <> 'OK' then
          if Application.MessageBox(PChar('Odpojení se nepodaøilo uložit do databáze. Pokraèovat?'),
           'Pozor', MB_ICONQUESTION + MB_YESNO + MB_DEFBUTTON1) = IDNO then Exit
          else Continue;
{
    for Radek := 1 to RadekDo do
      if Ints[8, Radek] = 1 then try
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
        + Cells[10, Radek] + ', '
        + Ap + 'Customer' + ApC
        + Ap + Format('Smlouva %s - odpojeno programem Nezaplacené ZL.', [Cells[6, Radek]]) + ApC
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
        + Cells[11, Radek] + ', '
        + Ap + 'Contract' + ApC
        + Ap + 'Odpojeno programem Nezaplacené ZL.' + ApC
        + '2, '                                            // automat
        + Ap + FormatDateTime('yyyy-mm-dd hh:nn:ss', Now) + ApC
        + Ap + FormatDateTime('yyyy-mm-dd hh:nn:ss', Now) + ApZ;
        SQL.Text := SQLStr;
        ExecSQL;
        SQLStr := 'UPDATE contracts SET'
        + ' State = ''disconnected'','
        + ' Invoice = 0'
        + ' WHERE Id = ' + Cells[11, Radek];
        SQL.Text := SQLStr;
        ExecSQL;   }
        if (Colors[8, Radek] <> clSilver) then Colors[8, Radek] := clSilver
        else Colors[8, Radek] := clWhite;
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
          if (RecordCount = 0) then begin
            Invoice_debt := '0';
            Account_debt := '0';
            Enforcement_state := '';
          end else begin
            Invoice_debt := FieldByName('Invoice_debt').AsString;
            Account_debt := FieldByName('Account_debt').AsString;
            Enforcement_state := FieldByName('Enforcement_state').AsString;
          end;
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
          + Cells[1, Radek] + ', '                         // poèet faktur
          + Ap + Invoice_debt + ApC
          + Ap + Account_debt + ApC
          + Ap + Enforcement_state + ApC              // stav vymáhání
          + Ap + 'odpojení' + ApC
          + Ap + 'program NezaplaceneFaktury' + ApZ;
          SQL.Text := SQLStr;
          ExecSQL;
        except on E: exception do
          ShowMessage('Odpojení se nepodaøilo uložit do sinners: ' + E.Message);
        end;
      except on E: exception do
        if Application.MessageBox(PChar(E.Message + #13#10 + 'Odpojení se nepodaøilo uložit do databáze. Pokraèovat?'),
         'Pozor', MB_ICONQUESTION + MB_YESNO + MB_DEFBUTTON1) = IDNO then Exit
        else Continue;
      end;
  finally
    Close;
    Screen.Cursor := crDefault;
  end;
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
