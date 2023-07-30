unit ZLmail;

interface

uses
  Forms, Controls, SysUtils, Classes, DateUtils, Math, IdText, IdBaseComponent, IdComponent, IdTCPConnection,
  IdTCPClient, IdMessageClient, IdSMTPBase, IdSMTP, IdMessage, IdExplicitTLSClientServerBase, IdAttachmentFile,
  ZLmain, IdIOHandler, IdIOHandlerSocket, IdIOHandlerStack, IdSSL,
  IdSSLOpenSSL;

type
  TdmMail = class(TDataModule)
    idMessage: TIdMessage;
    idSMTP: TIdSMTP;
    idSSLHandler: TIdSSLIOHandlerSocketOpenSSL;
  private
    procedure ZLMail(Radek: integer);                // po�le jeden ZL
  public
    procedure PosliZL;
  end;

var
  dmMail: TdmMail;

implementation

{$R *.dfm}

uses ZLcommon;

// ------------------------------------------------------------------------------------------------

procedure TdmMail.PosliZL;
// pou�ije data z asgMain
var
  Radek: integer;
begin
  with fmMain do try
    fmMain.Zprava(Format('Rozes�l�n� ZL od ��sla %s do %s', [aedOd.Text, aedDo.Text]));
    with asgMain do begin
      fmMain.Zprava(Format('Po�et ZL k rozesl�n�: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));
      Screen.Cursor := crHourGlass;
      apnMail.Visible := False;
      apbProgress.Position := 0;
      apbProgress.Visible := True;
// hlavn� smy�ka
      for Radek := 1 to RowCount-1 do begin
        Row := Radek;
        apbProgress.Position := Round(100 * Radek / RowCount-1);
        Application.ProcessMessages;
        if Prerusit then begin
          Prerusit := False;
          apbProgress.Position := 0;
          apbProgress.Visible := False;
          btVytvorit.Enabled := True;
          asgMain.Visible := True;
          lbxLog.Visible := False;
          Break;
        end;
// 18.12.18        if Ints[0, Radek] = 1 then ZLMail(Radek);
        if (Ints[0, Radek] = 1) and (Cells[5, Radek] <> '') then ZLMail(Radek);
      end;  // for
    end;
// konec hlavn� smy�ky
  finally
    if idSMTP.Connected then idSMTP.Disconnect;
    apbProgress.Position := 0;
    apbProgress.Visible := False;
    apnMail.Visible := True;
    asgMain.Visible := False;
    lbxLog.Visible := True;
    Screen.Cursor := crDefault;
    fmMain.Zprava('Rozes�l�n� ZL ukon�eno');
  end;
end;

// ------------------------------------------------------------------------------------------------

procedure TdmMail.ZLMail(Radek: integer);
var
  MailStr,
  CisloZL,
  PDFFile: AnsiString;
begin
  with fmMain, fmMain.asgMain do begin
    CisloZL := Copy(Cells[2, Radek], 0, Pos('/', Cells[2, Radek])-1);
// mus� existovat PDF soubor se z�lohou
    PDFFile := Format('%s\%4d\%2.2d\ZL1-%4.4d.pdf',
     [PDFDir, YearOf(Floats[7, Radek]), MonthOf(Floats[7, Radek]), StrToInt(CisloZL)]);
    if not FileExists(PDFFile) then begin
      fmMain.Zprava(Format('%s (%s): Soubor %s neexistuje. P�esko�eno.', [Cells[4, Radek], Cells[1, Radek], PDFFile]));
      Exit;
    end;
// alespo� n�jak� kontrola mailov� adresy
    if Pos('@', Cells[5, Radek]) = 0 then begin
      fmMain.Zprava(Format('%s (%s): Neplatn� mailov� adresa "%s". P�esko�eno.', [Cells[4, Radek], Cells[1, Radek], Cells[5, Radek]]));
      Exit;
    end;
    MailStr := Cells[5, Radek];
    MailStr := StringReplace(MailStr, ',', ';', [rfReplaceAll]);    // ��rky za st�edn�ky
    with idMessage do begin
      Clear;
//      ContentType := 'text/plain';
// v�ce mailov�ch adres odd�len�ch st�edn�ky se rozd�l�
//      Charset := 'Windows-1250';
      while Pos(';', MailStr) > 0 do begin
        Recipients.Add.Address := Trim(Copy(MailStr, 1, Pos(';', MailStr)-1));
        MailStr := Copy(MailStr, Pos(';', MailStr)+1, Length(MailStr));
      end;
      Recipients.Add.Address := Trim(MailStr);
      From.Address := 'uctarna@eurosignal.cz';
//      ReceiptRecipient.Text := 'uctarna@eurosignal.cz';
      Subject := Format('Dru�stvo EUROSIGNAL, z�lohov� list na internetovou konektivitu ZL1-%4.4d/%d', [StrToInt(CisloZL), aseRok.Value]);

      with TIdText.Create(idMessage.MessageParts, nil) do begin
        Body.Text := UTF8Encode(Format('Z�lohov� list ZL1-%4.4d/%d na p�ipojen� k internetu v dal��m obdob� je v p�ilo�en�m PDF dokumentu.',
         [StrToInt(CisloZL), aseRok.Value]) + #13#10#10#10
         + 'P�ejeme p�kn� den' + #13#10#10
         + 'V� Eurosignal');
        ContentType := 'text/plain';
        Charset := 'utf-8';
      end;

      ContentType := 'multipart/mixed';

{      Body.Add(Format('Z�lohov� list ZL1-%4.4d/%d na p�ipojen� k internetu je v p�ilo�en�m PDF dokumentu.'
      + ' Posledn� verze programu Adobe Reader, kter�m m��ete fakturu zobrazit'
      + ' i vytisknout, je zdarma ke sta�en� na http://get.adobe.com/reader/otherversions/.',
       [StrToInt(Copy(Cells[2, Radek], 0, Pos('/', Cells[2, Radek])-1)), aseRok.Value]));
      Body.Add(' ');
      Body.Add('Pokud dostanete tuto zpr�vu bez p��lohy, napi�te n�m, pros�m, my se to pokus�me napravit.');
      Body.Add(' ');
      Body.Add(' ');
      Body.Add('P�ejeme p�kn� den');
      Body.Add(' ');
      Body.Add('Dru�stvo Eurosignal');
      TIdAttachment.Create(MessageParts, PDFFile);
// p�id� se p��loha, je-li vybr�na a z�kazn�kovi se pos�l� reklama
      if (Ints[6, Radek] = 0) and (fePriloha.FileName <> '') then TIdAttachment.Create(MessageParts, fePriloha.FileName);
      with idSMTP do begin
        Port := 25;
        if Username = '' then AuthenticationType := atNone
        else AuthenticationType := atLogin;
      end;  }

      TIdAttachmentFile.Create(MessageParts, PDFFile);
// p�id� se p��loha, je-li vybr�na a z�kazn�kovi se pos�l� reklama
      if (Ints[6, Radek] = 0) and (fePriloha.FileName <> '') then TIdAttachmentFile.Create(MessageParts, fePriloha.FileName);

      try
        if not idSMTP.Connected then idSMTP.Connect;
        idSMTP.Send(idMessage);
        fmMain.Zprava(Format('%s (%s): Soubor %s byl odesl�n na adresu %s.',
         [Cells[4, Radek], Cells[1, Radek], PDFFile, Cells[5, Radek]]));
        Ints[0, Radek] := 0;
      except on E: exception do
        fmMain.Zprava(Format('%s (%s): Soubor %s se nepoda�ilo odeslat na adresu %s.' + ^M + 'Chyba: %s',
         [Cells[4, Radek], Cells[1, Radek], PDFFile, Cells[5, Radek], E.Message]));
      end;
      Application.ProcessMessages;
    end;
  end;  // with fmMain
end;  // procedury ZLMail

end.
