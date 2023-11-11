unit NZLmail;

interface

uses
  Forms, Controls, SysUtils, Classes, DateUtils, Math;

type
  TdmMail = class(TDataModule)
  private
    procedure FOxMail(Radek: integer); // pošle jednu FA
  public
    procedure PosliFOx;
  end;

var
  dmMail: TdmMail;

implementation

{$R *.dfm}

uses DesUtils, DesInvoices, NZLmain, NZLcommon;

// ------------------------------------------------------------------------------------------------

procedure TdmMail.PosliFOx;

var
  Radek: integer;
begin
  with fmMain do try
    fmMain.Zprava('Rozesílání FO3.');
    with asgMain do begin
      fmMain.Zprava(Format('Poèet faktur k rozeslání: %d', [Trunc(ColumnSum(2, 1, RowCount-1))]));

      for Radek := 1 to RowCount-1 do begin
        Row := Radek;
        apbProgress.Position := Round(100 * Radek / RowCount-1);
        Application.ProcessMessages;
        if Prerusit then Break;

        if Ints[2, Radek] = 1 then FOxMail(Radek);
      end;
    end;

  finally
    //asgMain.Visible := False;
    lbxLog.Visible := True;
    fmMain.Zprava('Odeslání faktur ukonèeno');
  end;
end;

// ------------------------------------------------------------------------------------------------

procedure TdmMail.FOxMail(Radek: integer);
var
  emailAddrStr,
  emailPredmet,
  emailZprava,
  emailOdesilatel,
  FullPdfFileName: string;
  VysledekZaslani : TDesResult;
  Od, Po: integer;
  Faktura : TDesInvoice;
begin
  with fmMain, fmMain.asgMain do begin

    Faktura := TDesInvoice.create(Cells[14, Radek]);
    FullPdfFileName := Faktura.getFullPdfFileName;

    emailAddrStr := Cells[12, Radek];
    emailOdesilatel := 'uctarna@eurosignal.cz';

    if Faktura.DocQueueCode = 'FO3' then begin

      emailPredmet := Format('Družstvo EUROSIGNAL, doklad k zaplacenému zálohovému listu %s', [Cells[3, Radek]]);
      emailZprava := Format('Zúètování zálohového listu %s je ve faktuøe %s v pøiloženém PDF dokumentu.',
        [Cells[3, Radek], Cells[9, Radek]])

    end else
    if Faktura.DocQueueCode = 'FO2' then begin

      emailPredmet := Format('Družstvo EUROSIGNAL, doklad k zaplacenému kreditu na VoIP %s', [Cells[9, Radek]]);
      emailZprava := Format('Daòový doklad %s k zaplacenému kreditu na VoIP je v pøiloženém PDF dokumentu.',
        [Cells[9, Radek]])

    end else
    if Faktura.DocQueueCode = 'FO4' then begin

      emailPredmet := Format('Družstvo EUROSIGNAL, doklad k zaplacenému kreditu na internet %s', [Cells[9, Radek]]);
      emailZprava := Format('Daòový doklad %s k zaplacenému kreditu na internet je v pøiloženém PDF dokumentu.',
        [Cells[9, Radek]])

    end else begin
      fmMain.Zprava(Format('%s (%s): Pro øadu %s není definovaná podoba e-mailu', [Cells[8, Radek], Cells[9, Radek], Faktura.DocQueueCode]));
      Exit;
    end;

    emailZprava := emailZprava
      + sLineBreak + sLineBreak
      + 'Pøejeme pìkný den'
      + sLineBreak + sLineBreak
      + 'Váš Eurosignal';


    VysledekZaslani := DesU.posliPdfEmailem(FullPdfFileName, emailAddrStr, emailPredmet, emailZprava, emailOdesilatel);

    //fmMain.Zprava(Format('%s (%s): %s', [Cells[8, Radek], Cells[9, Radek], 'test poslano']));
    fmMain.Zprava(Format('%s (%s): %s', [Cells[8, Radek], Cells[9, Radek], VysledekZaslani.Messg]));
    if VysledekZaslani.isOk then begin
      Ints[2, Radek] := 0;
      DesU.ulozKomunikaci(2, DesU.getCustomerIdByCoNumber(Faktura.VS), emailZprava);
    end;

  end;
end;

end.
