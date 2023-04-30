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
    procedure FakturaMail(Radek: integer);                // pošle jednu fakturu
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
// použije data z asgMain
var
  Radek: integer;
begin

  idSMTP.Host :=  DesU.getIniValue('Mail', 'SMTPServer');
  idSMTP.Username := DesU.getIniValue('Mail', 'SMTPLogin');
  idSMTP.Password := DesU.getIniValue('Mail', 'SMTPPW');


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
      end;  // konec hlavní smyèky
    end;

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
  emailAddrStr,
  emailPredmet,
  emailZprava,
  emailOdesilatel,
  FullPdfFileName: string;
begin
  with fmMain, fmMain.asgMain do begin
// musí existovat PDF soubor s fakturou
    FullPdfFileName := Format('%s\%4d\%2.2d\%s-%5.5d.pdf', [globalAA['PDFDir'], aseRok.Value, aseMesic.Value, globalAA['invoiceDocQueueCode'], Ints[2, Radek]]);
    //PDFFileName := Format('%s-%5.5d.pdf', [globalAA['invoiceDocQueueCode'], Ints[2, Radek]]); // neni potreba doufam
    if not FileExists(FullPdfFileName) then begin
      dmCommon.Zprava(Format('%s (%s): Soubor %s neexistuje. Pøeskoèeno.', [Cells[4, Radek], Cells[1, Radek], FullPdfFileName]));
      Exit;
    end;
// alespoò nìjaká kontrola mailové adresy
    if Pos('@', Cells[5, Radek]) = 0 then begin
      dmCommon.Zprava(Format('%s (%s): Neplatná mailová adresa "%s". Pøeskoèeno.', [Cells[4, Radek], Cells[1, Radek], Cells[5, Radek]]));
      Exit;
    end;

    emailAddrStr := Cells[5, Radek];
    emailOdesilatel := 'uctarna@eurosignal.cz';
    emailPredmet := Format('Družstvo EUROSIGNAL, faktura za internet FO1-%5.5d/%d', [Ints[2, Radek], aseRok.Value]);

    emailZprava := Format('Faktura FO1-%5.5d/%d za pøipojení k internetu je v pøiloženém PDF dokumentu.'
      + ' Poslední verze programu Adobe Reader, kterým mùžete PDF dokumenty zobrazit i vytisknout,'
      + ' je zdarma ke stažení na http://get.adobe.com/reader/otherversions/.', [Ints[2, Radek], aseRok.Value])
      + sLineBreak + sLineBreak
      +'Pokud dostanete tuto zprávu bez pøílohy, napište nám, prosím, my se to pokusíme napravit.'
      + sLineBreak + sLineBreak
      + 'Pøejeme pìkný den'
      + sLineBreak + sLineBreak
      +'Družstvo Eurosignal';

    try
      // !!! samotné poslání mailu
      DesFrxU.posliPdfEmailem(FullPdfFileName, emailAddrStr, emailPredmet, emailZprava, emailOdesilatel);

      dmCommon.Zprava(Format('%s (%s): Soubor %s byl odeslán na adresu %s.',
       [Cells[4, Radek], Cells[1, Radek], FullPdfFileName, Cells[5, Radek]]));
      Ints[0, Radek] := 0;
    except on E: exception do
      dmCommon.Zprava(Format('%s (%s): Soubor %s se nepodaøilo odeslat na adresu %s.' + #13#10 + 'Chyba: %s',
       [Cells[4, Radek], Cells[1, Radek], FullPdfFileName, Cells[5, Radek], E.Message]));
    end;
    Application.ProcessMessages;

  end;  // with fmMain
end;  // procedury FakturaMail

end.
