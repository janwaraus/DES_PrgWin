unit FOx_mail;

interface

uses
  Forms, Controls, SysUtils, Classes, DateUtils, Math, FOx_main,
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdMessage, IdMessageClient, IdMessageParts, IdText,
  IdSMTPBase, IdSMTP, IdAttachmentFile, IdAntiFreezeBase, IdAntiFreeze, IdExplicitTLSClientServerBase;

type
  TdmMail = class(TDataModule)
    idMessage: TIdMessage;
    idSMTP: TIdSMTP;
    IdAntiFreeze: TIdAntiFreeze;
  private
    procedure FOxMail(Radek: integer);                // po�le jeden doklad
  public
    procedure PosliFOx;
  end;

var
  dmMail: TdmMail;

implementation

{$R *.dfm}

uses FOx_common;

// ------------------------------------------------------------------------------------------------

procedure TdmMail.PosliFOx;
// pou�ije data z asgMain
var
  Radek: integer;
begin
  with fmMain do try
    dmCommon.Zprava('Rozes�l�n� doklad�.');
    with asgMain do begin
      dmCommon.Zprava(Format('Po�et doklad� k rozesl�n�: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));
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
        if Ints[0, Radek] = 1 then FOxMail(Radek);
      end;  // for
    end;
// konec hlavn� smy�ky
  finally
    apbProgress.Position := 0;
    apbProgress.Visible := False;
    apnMail.Visible := True;
    asgMain.Visible := False;
    lbxLog.Visible := True;
    if idSMTP.Connected then idSMTP.Disconnect;
    Screen.Cursor := crDefault;
    dmCommon.Zprava('Rozes�l�n� doklad� ukon�eno');
  end;
end;

// ------------------------------------------------------------------------------------------------

procedure TdmMail.FOxMail(Radek: integer);
var
  OutDir,
  MailStr,
  PDFFile: AnsiString;
begin
  with fmMain, asgMain do begin
    OutDir := Format('%s\%4d\%2.2d', [PDFDir, aseRok.Value, aseMesic.Value]);
    if rbInternet.Checked then CisloFO := Format('FO4-%4.4d', [Ints[2, Radek]])
    else CisloFO := Format('FO2-%4.4d', [Ints[2, Radek]]);
// mus� existovat PDF soubor s fakturou
    PDFFile := Format('%s\%s.pdf', [OutDir, CisloFO]);
    if not FileExists(PDFFile) then begin
      dmCommon.Zprava(Format('Soubor %s neexistuje. P�esko�eno.', [PDFFile]));
      Exit;
    end;
// alespo� n�jak� kontrola mailov� adresy
    if Pos('@', Cells[5, Radek]) = 0 then begin
      dmCommon.Zprava(Format('Neplatn� mailov� adresa "%s". P�esko�eno.', [Cells[6, Radek]]));
      Exit;
    end;
    MailStr := Cells[5, Radek];
    MailStr := StringReplace(MailStr, ',', ';', [rfReplaceAll]);    // ��rky za st�edn�ky
    with idMessage do begin
      Clear;
//      ContentType := 'text/plain';
//      Charset := 'Windows-1250';
// v�ce mailov�ch adres odd�len�ch st�edn�ky se rozd�l�
      while Pos(';', MailStr) > 0 do begin
        Recipients.Add.Address := Trim(Copy(MailStr, 1, Pos(';', MailStr)-1));
        MailStr := Copy(MailStr, Pos(';', MailStr)+1, Length(MailStr));
      end;
      Recipients.Add.Address := Trim(MailStr);
      From.Address := 'uctarna@eurosignal.cz';
      ReceiptRecipient.Text := 'uctarna@eurosignal.cz';
      Subject := Format('Dru�stvo EUROSIGNAL, da�ov� doklad k zaplacen�mu kreditu %s/%d', [CisloFO, aseRok.Value]);
      with TIdText.Create(idMessage.MessageParts, nil) do begin
        Body.Text := Format('Da�ov� doklad %s/%d k zaplacen�mu kreditu je v p�ilo�en�m PDF dokumentu.'
         + ' Posledn� verze programu Adobe Reader, kter�m m��ete PDF dokumenty zobrazit i vytisknout,'
         + ' je zdarma ke sta�en� na http://get.adobe.com/reader/otherversions/.', [CisloFO, aseRok.Value])
         + sLineBreak + sLineBreak
         +'S pozdravem'
         + sLineBreak + sLineBreak
         + 'Dru�stvo Eurosignal';
        ContentType := 'text/plain';
        Charset := 'utf-8';
      end;
      ContentType := 'multipart/mixed';
{
      Body.Add(Format('Da�ov� doklad %s/%d k zaplacen�mu kreditu je v p�ilo�en�m PDF dokumentu.'
      + ' Posledn� verze programu Adobe Reader, kter�m m��ete PDF dokumenty zobrazit i vytisknout,'
      + ' je zdarma ke sta�en� na http://get.adobe.com/reader/otherversions/.', [CisloFO, aseRok.Value]));
      Body.Add(' ');
      Body.Add('Pokud dostanete tuto zpr�vu bez p��lohy, napi�te n�m, pros�m, my se to pokus�me napravit.');
      Body.Add(' ');
      Body.Add(' ');
      Body.Add('P�ejeme p�kn� den');
      Body.Add(' ');
      Body.Add('Dru�stvo Eurosignal');
}
      TIdAttachmentFile.Create(MessageParts, PDFFile);
// p�id� se p��loha, je-li vybr�na a z�kazn�kovi se pos�l� reklama
      if (Ints[6, Radek] = 0) and (fePriloha.FileName <> '') then TIdAttachmentFile.Create(MessageParts, fePriloha.FileName);
      with idSMTP do begin
        Port := 25;
        if Username = '' then AuthType := satNone
        else AuthType := satDefault;
      end;
      try
        if not idSMTP.Connected then idSMTP.Connect;
        idSMTP.Send(idMessage);
        dmCommon.Zprava(Format('%s (%s): Soubor %s byl odesl�n na adresu %s.',
         [Cells[4, Radek], Cells[1, Radek], PDFFile, Cells[5, Radek]]));
        Ints[0, Radek] := 0;
      except on E: exception do
        dmCommon.Zprava(Format('%s (%s): Soubor %s se nepoda�ilo odeslat na adresu %s.' + #13#10 + 'Chyba: %s',
         [Cells[4, Radek], Cells[1, Radek], PDFFile, Cells[5, Radek], E.Message]));
      end;
      Application.ProcessMessages;
    end;
  end;  // with fmMain
end;  // procedury ZLMail

end.
