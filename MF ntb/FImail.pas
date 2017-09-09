unit FImail;

interface

uses
  Forms, Controls, SysUtils, Classes, IdComponent, IdTCPConnection, IdTCPClient, IdMessageClient, IdSMTP,
  IdBaseComponent, IdMessage, FImain, IdExplicitTLSClientServerBase, IdSMTPBase;

type
  TdmMail = class(TDataModule)
    idMessage: TIdMessage;
    idSMTP: TIdSMTP;
  private
    procedure FakturaMail(Radek: integer);                // po�le jednu fakturu
  public
    procedure PosliFaktury;
  end;

var
  dmMail: TdmMail;

implementation

{$R *.dfm}

uses FIcommon;

// ------------------------------------------------------------------------------------------------

procedure TdmMail.PosliFaktury;
// pou�ije data z asgMain
var
  Radek: integer;
begin
  with fmMain do try
    if rbPodleSmlouvy.Checked then
      dmCommon.Zprava(Format('Rozes�l�n� faktur od VS %s do %s', [aedOd.Text, aedDo.Text]))
    else dmCommon.Zprava(Format('Rozes�l�n� faktur od ��sla %s do %s', [aedOd.Text, aedDo.Text]));
    with asgMain do begin
      dmCommon.Zprava(Format('Po�et faktur k rozesl�n�: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));
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
        if Ints[0, Radek] = 1 then FakturaMail(Radek);
      end;  // for
    end;
// konec hlavn� smy�ky
  finally
    apbProgress.Position := 0;
    apbProgress.Visible := False;
    apnMail.Visible := True;
    asgMain.Visible := False;
    lbxLog.Visible := True;
    Screen.Cursor := crDefault;
    dmCommon.Zprava('Rozes�l�n� faktur ukon�eno');
  end;
end;

// ------------------------------------------------------------------------------------------------

procedure TdmMail.FakturaMail(Radek: integer);
var
  MailStr,
  PDFFile: AnsiString;
begin
  with fmMain, fmMain.asgMain do begin
    FStr := 'FO1';

    // mus� existovat PDF soubor s fakturou
    PDFFile := Format('%s\%4d\%2.2d\%s-%5.5d.pdf', [PDFDir, aseRok.Value, aseMesic.Value, FStr, Ints[2, Radek]]);
    if not FileExists(PDFFile) then begin
      dmCommon.Zprava(Format('%s (%s): Soubor %s neexistuje. P�esko�eno.', [Cells[4, Radek], Cells[1, Radek], PDFFile]));
      Exit;
    end;

    // alespo� n�jak� kontrola mailov� adresy
    if Pos('@', Cells[5, Radek]) = 0 then begin
      dmCommon.Zprava(Format('%s (%s): Neplatn� mailov� adresa "%s". P�esko�eno.', [Cells[4, Radek], Cells[1, Radek], Cells[5, Radek]]));
      Exit;
    end;

    MailStr := Cells[5, Radek];
    MailStr := StringReplace(MailStr, ',', ';', [rfReplaceAll]);    // ��rky za st�edn�ky
    with idMessage do begin
      Clear;
      ContentType := 'text/plain';
      Charset := 'Windows-1250';
      // v�ce mailov�ch adres odd�len�ch st�edn�ky se rozd�l�
      while Pos(';', MailStr) > 0 do begin
        Recipients.Add.Address := Trim(Copy(MailStr, 1, Pos(';', MailStr)-1));
        MailStr := Copy(MailStr, Pos(';', MailStr)+1, Length(MailStr));
      end;
      Recipients.Add.Address := Trim(MailStr);
      From.Address := 'uctarna@eurosignal.cz';
      ReceiptRecipient.Text := 'uctarna@eurosignal.cz';
      Subject := Format('Dru�stvo EUROSIGNAL, faktura za internet FO1-%5.5d/%d', [Ints[2, Radek], aseRok.Value]);
      Body.Add(Format('Faktura FO1-%5.5d/%d za p�ipojen� k internetu je v p�ilo�en�m PDF dokumentu.'
      + ' Posledn� verze programu Adobe Reader, kter�m m��ete PDF dokumenty zobrazit i vytisknout,'
      + ' je zdarma ke sta�en� na http://get.adobe.com/reader/otherversions/.', [Ints[2, Radek], aseRok.Value]));

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
end;  // procedury FakturaMail

end.
