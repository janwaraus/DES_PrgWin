unit FImail;

interface

uses
  Forms, Controls, SysUtils, Classes,

  IdBaseComponent, IdComponent, IdTCPConnection,
  IdTCPClient, IdSMTP, IdHTTP, IdMessage, IdMessageClient, IdText, IdMessageParts,
  IdAntiFreezeBase, IdAntiFreeze, ZAbstractConnection, IdIOHandler,
  IdIOHandlerSocket, IdSSLOpenSSL, IdExplicitTLSClientServerBase, IdSMTPBase, IdAttachmentFile
  ;

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

uses DesUtils, FImain;

{$R *.dfm}


// ------------------------------------------------------------------------------------------------

procedure TdmMail.PosliFaktury;
// pou�ije data z asgMain
var
  Radek: integer;
begin

  { preneseno do DesUtils
  idSMTP.Host :=  DesU.getIniValue('Mail', 'SMTPServer');
  idSMTP.Username := DesU.getIniValue('Mail', 'SMTPLogin');
  idSMTP.Password := DesU.getIniValue('Mail', 'SMTPPW');
  }


  with fmMain do try

    if rbVyberPodleVS.Checked then
      fmMain.Zprava(Format('Rozes�l�n� faktur od VS %s do %s', [aedOd.Text, aedDo.Text]))
    else fmMain.Zprava(Format('Rozes�l�n� faktur od ��sla %s do %s', [aedOd.Text, aedDo.Text]));

    with asgMain do begin
      fmMain.Zprava(Format('Po�et faktur k rozesl�n�: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));
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
          Break;
        end;
        if Ints[0, Radek] = 1 then FakturaMail(Radek);
      end;  // konec hlavn� smy�ky
    end;

  finally
    apbProgress.Position := 0;
    apbProgress.Visible := False;
    apnMail.Visible := True;
    Screen.Cursor := crDefault;
    fmMain.Zprava('Rozes�l�n� faktur ukon�eno');
  end;
end;

// ------------------------------------------------------------------------------------------------

procedure TdmMail.FakturaMail(Radek: integer);
var
  emailAddrStr,
  emailPredmet,
  emailZprava,
  emailOdesilatel,
  FullPdfFileName,
  ExtraPrilohaFileName: string;
  VysledekZaslani : TDesResult;
begin
  with fmMain, fmMain.asgMain do begin
    // mus� existovat PDF soubor s fakturou
    FullPdfFileName := Format('%s\%4d\%2.2d\%s-%5.5d.pdf', [globalAA['PDFDir'], aseRok.Value, aseMesic.Value, globalAA['invoiceDocQueueCode'], Ints[2, Radek]]);
    // PDFFileName := Format('%s-%5.5d.pdf', [globalAA['invoiceDocQueueCode'], Ints[2, Radek]]); // neni potreba doufam
    if not FileExists(FullPdfFileName) then begin
      fmMain.Zprava(Format('%s (%s): Soubor %s neexistuje. P�esko�eno.', [Cells[4, Radek], Cells[1, Radek], FullPdfFileName]));
      Exit;
    end;

    emailAddrStr := Cells[5, Radek];
    // alespo� n�jak� kontrola mailov� adresy
    if Pos('@', emailAddrStr) = 0 then begin
      fmMain.Zprava(Format('%s (%s): Neplatn� mailov� adresa "%s". P�esko�eno.', [Cells[4, Radek], Cells[1, Radek], Cells[5, Radek]]));
      Exit;
    end;

    emailOdesilatel := 'uctarna@eurosignal.cz';
    emailPredmet := Format('Dru�stvo EUROSIGNAL, faktura za internet FO1-%5.5d/%d', [Ints[2, Radek], aseRok.Value]);

    emailZprava := Format('Faktura FO1-%5.5d/%d za p�ipojen� k internetu je v p�ilo�en�m PDF dokumentu.'
      // + ' Posledn� verze programu Adobe Reader, kter�m m��ete PDF dokumenty zobrazit i vytisknout,'
      // + ' je zdarma ke sta�en� na http://get.adobe.com/reader/otherversions/.'
      , [Ints[2, Radek], aseRok.Value])
      + sLineBreak + sLineBreak
      // + 'Pokud dostanete tuto zpr�vu bez p��lohy, napi�te n�m, pros�m, my se to pokus�me napravit.' + sLineBreak + sLineBreak
      + 'P�ejeme p�kn� den'
      + sLineBreak + sLineBreak
      + 'V� Eurosignal';

    if (Ints[6, Radek] = 0) and (fePriloha.FileName <> '') then
      ExtraPrilohaFileName := fePriloha.FileName
    else
      ExtraPrilohaFileName := '';

    VysledekZaslani := DesU.posliPdfEmailem(FullPdfFileName, emailAddrStr, emailPredmet, emailZprava, emailOdesilatel, ExtraPrilohaFileName);
    fmMain.Zprava(Format('%s (%s): %s', [Cells[4, Radek], Cells[1, Radek], VysledekZaslani.Messg]));
    Ints[0, Radek] := 0;
    Application.ProcessMessages;

  end;  // with fmMain
end;  // procedury FakturaMail

end.
