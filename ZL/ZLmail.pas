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
    procedure ZLMail(Radek: integer);                // pošle jeden ZL
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
// použije data z asgMain
var
  Radek: integer;
begin
  with fmMain do try
    fmMain.Zprava(Format('Rozesílání ZL od èísla %s do %s', [aedOd.Text, aedDo.Text]));
    with asgMain do begin
      fmMain.Zprava(Format('Poèet ZL k rozeslání: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));
      Screen.Cursor := crHourGlass;
      apnMail.Visible := False;
      apbProgress.Position := 0;
      apbProgress.Visible := True;
// hlavní smyèka
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
// konec hlavní smyèky
  finally
    if idSMTP.Connected then idSMTP.Disconnect;
    apbProgress.Position := 0;
    apbProgress.Visible := False;
    apnMail.Visible := True;
    asgMain.Visible := False;
    lbxLog.Visible := True;
    Screen.Cursor := crDefault;
    fmMain.Zprava('Rozesílání ZL ukonèeno');
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
// musí existovat PDF soubor se zálohou
    PDFFile := Format('%s\%4d\%2.2d\ZL1-%4.4d.pdf',
     [PDFDir, YearOf(Floats[7, Radek]), MonthOf(Floats[7, Radek]), StrToInt(CisloZL)]);
    if not FileExists(PDFFile) then begin
      fmMain.Zprava(Format('%s (%s): Soubor %s neexistuje. Pøeskoèeno.', [Cells[4, Radek], Cells[1, Radek], PDFFile]));
      Exit;
    end;
// alespoò nìjaká kontrola mailové adresy
    if Pos('@', Cells[5, Radek]) = 0 then begin
      fmMain.Zprava(Format('%s (%s): Neplatná mailová adresa "%s". Pøeskoèeno.', [Cells[4, Radek], Cells[1, Radek], Cells[5, Radek]]));
      Exit;
    end;
    MailStr := Cells[5, Radek];
    MailStr := StringReplace(MailStr, ',', ';', [rfReplaceAll]);    // èárky za støedníky
    with idMessage do begin
      Clear;
//      ContentType := 'text/plain';
// více mailových adres oddìlených støedníky se rozdìlí
//      Charset := 'Windows-1250';
      while Pos(';', MailStr) > 0 do begin
        Recipients.Add.Address := Trim(Copy(MailStr, 1, Pos(';', MailStr)-1));
        MailStr := Copy(MailStr, Pos(';', MailStr)+1, Length(MailStr));
      end;
      Recipients.Add.Address := Trim(MailStr);
      From.Address := 'uctarna@eurosignal.cz';
//      ReceiptRecipient.Text := 'uctarna@eurosignal.cz';
      Subject := Format('Družstvo EUROSIGNAL, zálohový list na internetovou konektivitu ZL1-%4.4d/%d', [StrToInt(CisloZL), aseRok.Value]);

      with TIdText.Create(idMessage.MessageParts, nil) do begin
        Body.Text := UTF8Encode(Format('Zálohový list ZL1-%4.4d/%d na pøipojení k internetu v dalším období je v pøiloženém PDF dokumentu.',
         [StrToInt(CisloZL), aseRok.Value]) + #13#10#10#10
         + 'Pøejeme pìkný den' + #13#10#10
         + 'Váš Eurosignal');
        ContentType := 'text/plain';
        Charset := 'utf-8';
      end;

      ContentType := 'multipart/mixed';

{      Body.Add(Format('Zálohový list ZL1-%4.4d/%d na pøipojení k internetu je v pøiloženém PDF dokumentu.'
      + ' Poslední verze programu Adobe Reader, kterým mùžete fakturu zobrazit'
      + ' i vytisknout, je zdarma ke stažení na http://get.adobe.com/reader/otherversions/.',
       [StrToInt(Copy(Cells[2, Radek], 0, Pos('/', Cells[2, Radek])-1)), aseRok.Value]));
      Body.Add(' ');
      Body.Add('Pokud dostanete tuto zprávu bez pøílohy, napište nám, prosím, my se to pokusíme napravit.');
      Body.Add(' ');
      Body.Add(' ');
      Body.Add('Pøejeme pìkný den');
      Body.Add(' ');
      Body.Add('Družstvo Eurosignal');
      TIdAttachment.Create(MessageParts, PDFFile);
// pøidá se pøíloha, je-li vybrána a zákazníkovi se posílá reklama
      if (Ints[6, Radek] = 0) and (fePriloha.FileName <> '') then TIdAttachment.Create(MessageParts, fePriloha.FileName);
      with idSMTP do begin
        Port := 25;
        if Username = '' then AuthenticationType := atNone
        else AuthenticationType := atLogin;
      end;  }

      TIdAttachmentFile.Create(MessageParts, PDFFile);
// pøidá se pøíloha, je-li vybrána a zákazníkovi se posílá reklama
      if (Ints[6, Radek] = 0) and (fePriloha.FileName <> '') then TIdAttachmentFile.Create(MessageParts, fePriloha.FileName);

      try
        if not idSMTP.Connected then idSMTP.Connect;
        idSMTP.Send(idMessage);
        fmMain.Zprava(Format('%s (%s): Soubor %s byl odeslán na adresu %s.',
         [Cells[4, Radek], Cells[1, Radek], PDFFile, Cells[5, Radek]]));
        Ints[0, Radek] := 0;
      except on E: exception do
        fmMain.Zprava(Format('%s (%s): Soubor %s se nepodaøilo odeslat na adresu %s.' + ^M + 'Chyba: %s',
         [Cells[4, Radek], Cells[1, Radek], PDFFile, Cells[5, Radek], E.Message]));
      end;
      Application.ProcessMessages;
    end;
  end;  // with fmMain
end;  // procedury ZLMail

end.
