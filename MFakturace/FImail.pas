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

uses DesUtils, DesFrxUtils, FIcommon, FImain;

{$R *.dfm}


// ------------------------------------------------------------------------------------------------

procedure TdmMail.PosliFaktury;
// pou�ije data z asgMain
var
  Radek: integer;
begin

  idSMTP.Host :=  DesU.getIniValue('Mail', 'SMTPServer');
  idSMTP.Username := DesU.getIniValue('Mail', 'SMTPLogin');
  idSMTP.Password := DesU.getIniValue('Mail', 'SMTPPW');


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
      end;  // konec hlavn� smy�ky
    end;

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
  emailAddrStr,
  emailPredmet,
  emailZprava,
  emailOdesilatel,
  pdfFile: string;
begin
  with fmMain, fmMain.asgMain do begin
    FStr := 'FO1';
// mus� existovat PDF soubor s fakturou
    PDFFile := Format('%s\%4d\%2.2d\%s-%5.5d.pdf', [PDFDir, aseRok.Value, aseMesic.Value, FStr, Ints[2, Radek]]);
    //PDFFileName := Format('%s-%5.5d.pdf', [FStr, Ints[2, Radek]]); // neni potreba doufam
    if not FileExists(PDFFile) then begin
      dmCommon.Zprava(Format('%s (%s): Soubor %s neexistuje. P�esko�eno.', [Cells[4, Radek], Cells[1, Radek], PDFFile]));
      Exit;
    end;
// alespo� n�jak� kontrola mailov� adresy
    if Pos('@', Cells[5, Radek]) = 0 then begin
      dmCommon.Zprava(Format('%s (%s): Neplatn� mailov� adresa "%s". P�esko�eno.', [Cells[4, Radek], Cells[1, Radek], Cells[5, Radek]]));
      Exit;
    end;


    emailOdesilatel := 'uctarna@eurosignal.cz';
    emailPredmet := Format('Dru�stvo EUROSIGNAL, faktura za internet FO1-%5.5d/%d', [Ints[2, Radek], aseRok.Value]);

    emailZprava := Format('Faktura FO1-%5.5d/%d za p�ipojen� k internetu je v p�ilo�en�m PDF dokumentu.'
      + ' Posledn� verze programu Adobe Reader, kter�m m��ete PDF dokumenty zobrazit i vytisknout,'
      + ' je zdarma ke sta�en� na http://get.adobe.com/reader/otherversions/.', [Ints[2, Radek], aseRok.Value])
      + sLineBreak + sLineBreak
      +'Pokud dostanete tuto zpr�vu bez p��lohy, napi�te n�m, pros�m, my se to pokus�me napravit.'
      + sLineBreak + sLineBreak
      + 'P�ejeme p�kn� den'
      + sLineBreak + sLineBreak
      +'Dru�stvo Eurosignal';

    try
      // !!! samotn� posl�n� mailu
      DesFrxU.posliPdfEmailem(pdfFile, emailAddrStr, emailPredmet, emailZprava, emailOdesilatel);

      dmCommon.Zprava(Format('%s (%s): Soubor %s byl odesl�n na adresu %s.',
       [Cells[4, Radek], Cells[1, Radek], PDFFile, Cells[5, Radek]]));
      Ints[0, Radek] := 0;
    except on E: exception do
      dmCommon.Zprava(Format('%s (%s): Soubor %s se nepoda�ilo odeslat na adresu %s.' + #13#10 + 'Chyba: %s',
       [Cells[4, Radek], Cells[1, Radek], PDFFile, Cells[5, Radek], E.Message]));
    end;
    Application.ProcessMessages;

  end;  // with fmMain
end;  // procedury FakturaMail

end.
