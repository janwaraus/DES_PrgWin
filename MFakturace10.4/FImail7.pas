unit FImail;

interface

uses
  Forms, Controls, SysUtils, Classes, IdComponent, IdTCPConnection, IdTCPClient, IdMessageClient, IdSMTP,
  IdBaseComponent, IdMessage, FImain;

type
  TdmMail = class(TDataModule)
    idMessage: TIdMessage;
    idSMTP: TIdSMTP;
  private
    procedure FakturaMail(Radek: integer);                // pošle jednu fakturu
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
// pouije data z asgMain
var
  Radek: integer;
begin
  with fmMain do try
    if rbPodleSmlouvy.Checked then
      dmCommon.Zprava(Format('Rozesílání faktur od VS %s do %s', [aedOd.Text, aedDo.Text]))
    else dmCommon.Zprava(Format('Rozesílání faktur od èísla %s do %s', [aedOd.Text, aedDo.Text]));
    with asgMain do begin
      dmCommon.Zprava(Format('Poèet faktur k rozeslání: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));
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
        if Ints[0, Radek] = 1 then FakturaMail(Radek);
      end;  // for
    end;
// konec hlavní smyèky
  finally
    apbProgress.Position := 0;
    apbProgress.Visible := False;
    apnMail.Visible := True;
    asgMain.Visible := False;
    lbxLog.Visible := True;
    Screen.Cursor := crDefault;
    dmCommon.Zprava('Rozesílání faktur ukonèeno');
  end;
end;

// ------------------------------------------------------------------------------------------------

procedure TdmMail.FakturaMail(Radek: integer);
var
  MailStr,
  PDFFile: AnsiString;
begin
  with fmMain, fmMain.asgMain do begin
{$IFDEF ABAK}
    if rbInternet.Checked then FStr := 'FiI'
    else FStr := 'FtI';
{$ELSE}
    FStr := 'FO1';
{$ENDIF}
// musí existovat PDF soubor s fakturou
    PDFFile := Format('%s\%4d\%2.2d\%s-%5.5d.pdf', [PDFDir, aseRok.Value, aseMesic.Value, FStr, Ints[2, Radek]]);
    if not FileExists(PDFFile) then begin
      dmCommon.Zprava(Format('%s (%s): Soubor %s neexistuje. Pøeskoèeno.', [Cells[4, Radek], Cells[1, Radek], PDFFile]));
      Exit;
    end;
// alespoò nìjaká kontrola mailové adresy
    if Pos('@', Cells[5, Radek]) = 0 then begin
      dmCommon.Zprava(Format('%s (%s): Neplatná mailová adresa "%s". Pøeskoèeno.', [Cells[4, Radek], Cells[1, Radek], Cells[5, Radek]]));
      Exit;
    end;
    MailStr := Cells[5, Radek];
    MailStr := StringReplace(MailStr, ',', ';', [rfReplaceAll]);    // èárky za støedníky
    with idMessage do begin
      Clear;
      ContentType := 'text/plain';
      Charset := 'Windows-1250';
// více mailovıch adres oddìlenıch støedníky se rozdìlí
      while Pos(';', MailStr) > 0 do begin
        Recipients.Add.Address := Trim(Copy(MailStr, 1, Pos(';', MailStr)-1));
        MailStr := Copy(MailStr, Pos(';', MailStr)+1, Length(MailStr));
      end;
      Recipients.Add.Address := Trim(MailStr);
{$IFDEF ABAK}
      From.Address := 'abak@abak.cz';
      CCList.Add.Address := 'abak@abak.cz';
      ReceiptRecipient.Text := 'abak@abak.cz';
      if rbInternet.Checked then Subject := Format('ABAK spol. s r.o., faktura za internet FiI-%5.5d/%d', [Ints[2, Radek], aseRok.Value])
      else Subject := Format('ABAK spol. s r.o., faktura za VoIP FtI-%5.5d/%d', [Ints[2, Radek], aseRok.Value]);
//      if rbInternet.Checked then Body.Add(Format('Faktura FiI-%5.5d/%d za pripojeni k internetu je v prilozeném PDF dokumentu.'
//      + ' Posledni verze programu Adobe Reader, kterım muzete fakturu zobrazit i vytisknout,'
//      + ' je zdarma ke stazeni na http://get.adobe.com/reader/otherversions/.', [Ints[2, Radek], aseRok.Value]))
      if rbInternet.Checked then begin
        Body.Add('Váení pøátelé,');
        Body.Add(' ');
        Body.Add('v pøíloze naleznete fakturu za sluby elektronickıch komunikací od spoleènosti ABAK spol. s r.o., provozující sí újezd.net.');
        Body.Add(' ');
        Body.Add('V zájmu lepší kvality a dostupnosti slueb jsme pøipravili dotazník, ve kterıch se Vás, jako našich váenıch klientù ptáme'
        + ' na hodnocení našich slueb. Prosíme Vás o 5 minut Vašeho èasu a o vyplnìní dotazníku uloeného v zabezpeèeném cloudu spoleènosti'
        + ' Google zde: https://goo.gl/forms/WoLeXfaYzZLfTb7V2');
      end else Body.Add(Format('Faktura FtI-%5.5d/%d za sluzbu VoIP je v prilozeném PDF dokumentu.'
      + ' Posledni verze programu Adobe Reader, kterım mùete PDF dokumenty zobrazit i vytisknout,'
      + ' je zdarma ke stazeni na http://get.adobe.com/reader/otherversions/.', [Ints[2, Radek], aseRok.Value]));
{$ELSE}
      From.Address := 'uctarna@eurosignal.cz';
      ReceiptRecipient.Text := 'uctarna@eurosignal.cz';
      Subject := Format('Drustvo EUROSIGNAL, faktura za internet FO1-%5.5d/%d', [Ints[2, Radek], aseRok.Value]);
      Body.Add(Format('Faktura FO1-%5.5d/%d za pøipojení k internetu je v pøiloeném PDF dokumentu.'
      + ' Poslední verze programu Adobe Reader, kterım mùete PDF dokumenty zobrazit i vytisknout,'
      + ' je zdarma ke staení na http://get.adobe.com/reader/otherversions/.', [Ints[2, Radek], aseRok.Value]));
{$ENDIF}
      Body.Add(' ');
      Body.Add('Pokud dostanete tuto zprávu bez pøílohy, napište nám, prosím, my se to pokusíme napravit.');
      Body.Add(' ');
      Body.Add(' ');
      Body.Add('Pøejeme pìknı den');
      Body.Add(' ');
{$IFDEF ABAK}
      Body.Add('Váš újezd.net');
{$ELSE}
      Body.Add('Drustvo Eurosignal');
{$ENDIF}
      TIdAttachment.Create(MessageParts, PDFFile);
// pøidá se pøíloha, je-li vybrána a zákazníkovi se posílá reklama
      if (Ints[6, Radek] = 0) and (fePriloha.FileName <> '') then TIdAttachment.Create(MessageParts, fePriloha.FileName);
      with idSMTP do begin
        Port := 25;
        if Username = '' then AuthenticationType := atNone
        else AuthenticationType := atLogin;
      end;
      try
        if not idSMTP.Connected then idSMTP.Connect;
        idSMTP.Send(idMessage);
        dmCommon.Zprava(Format('%s (%s): Soubor %s byl odeslán na adresu %s.',
         [Cells[4, Radek], Cells[1, Radek], PDFFile, Cells[5, Radek]]));
        Ints[0, Radek] := 0;
      except on E: exception do
        dmCommon.Zprava(Format('%s (%s): Soubor %s se nepodaøilo odeslat na adresu %s.' + #13#10 + 'Chyba: %s',
         [Cells[4, Radek], Cells[1, Radek], PDFFile, Cells[5, Radek], E.Message]));
      end;
      Application.ProcessMessages;
    end;
  end;  // with fmMain
end;  // procedury FakturaMail

end.
