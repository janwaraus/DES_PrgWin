unit FImail;

interface

uses
  Forms, Controls, SysUtils, Classes, Dialogs,

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
    procedure FakturaMail(Radek: integer); // po�le jednu fakturu
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
var
  Radek: integer;
begin
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
      end;
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
    FullPdfFileName := Format('%s%4d\%2.2d\%s-%5.5d.pdf', [DesU.PDF_PATH, aseRok.Value, aseMesic.Value, 'FO1', Ints[2, Radek]]);

    emailAddrStr := Cells[5, Radek];

    emailOdesilatel := 'uctarna@eurosignal.cz';
    emailPredmet := Format('Dru�stvo EUROSIGNAL, faktura za internet FO1-%d/%d', [Ints[2, Radek], aseRok.Value]);

    emailZprava := Format('Faktura FO1-%d/%d za p�ipojen� k internetu je v p�ilo�en�m PDF dokumentu.'
      , [Ints[2, Radek], aseRok.Value])
      + sLineBreak + sLineBreak
      + 'P�ejeme p�kn� den'
      + sLineBreak + sLineBreak
      + 'V� Eurosignal';

    if (Ints[6, Radek] = 0) and (fePriloha.FileName <> '') then
      ExtraPrilohaFileName := fePriloha.FileName
    else
      ExtraPrilohaFileName := '';

    VysledekZaslani := DesU.posliPdfEmailem(FullPdfFileName, emailAddrStr, emailPredmet, emailZprava, emailOdesilatel, ExtraPrilohaFileName);
    fmMain.Zprava(Format('%s (%s): %s', [Cells[4, Radek], Cells[1, Radek], VysledekZaslani.Messg]));
    if VysledekZaslani.isOk then begin
      Ints[0, Radek] := 0;
    end;

  end;
end;

end.
