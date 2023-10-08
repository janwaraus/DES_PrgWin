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

uses DesUtils, NZLmain, NZLcommon;

// ------------------------------------------------------------------------------------------------

procedure TdmMail.PosliFOx;

var
  Radek: integer;
begin
  with fmMain do try
    fmMain.Zprava('Rozesílání FO3.');
    with asgMain do begin
      fmMain.Zprava(Format('Poèet faktur k rozeslání: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));

      for Radek := 1 to RowCount-1 do begin
        Row := Radek;
        apbProgress.Position := Round(100 * Radek / RowCount-1);
        Application.ProcessMessages;
        if Prerusit then Break;

        if Ints[2, Radek] = 1 then FOxMail(Radek);
      end;
    end;

  finally
    asgMain.Visible := False;
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
begin
  with fmMain, fmMain.asgMain do begin

    Od := Pos('-', Cells[7, Radek]) + 1;
    Po := Pos('/', Cells[7, Radek]);
    FullPdfFileName := Format('%s%4d\%2.2d\FO3-%4.4d.pdf',
     [DesU.PDF_PATH, YearOf(deDatumDokladu.Date), MonthOf(deDatumDokladu.Date), StrToInt(Copy(Cells[7, Radek],  Od, Po-Od))]);

    emailAddrStr := Cells[6, Radek];

    emailOdesilatel := 'uctarna@eurosignal.cz';
    emailPredmet := Format('Družstvo EUROSIGNAL, doklad k zaplacenému zálohovému listu %s', [Cells[1, Radek]]);

    emailZprava := Format('Zúètování zálohového listu %s je ve faktuøe %s v pøiloženém PDF dokumentu.'
      , [Cells[1, Radek], Cells[7, Radek]])
      + sLineBreak + sLineBreak
      + 'Pøejeme pìkný den'
      + sLineBreak + sLineBreak
      + 'Váš Eurosignal';


    VysledekZaslani := DesU.posliPdfEmailem(FullPdfFileName, emailAddrStr, emailPredmet, emailZprava, emailOdesilatel);
    fmMain.Zprava(Format('%s (%s): %s', [Cells[4, Radek], Cells[1, Radek], VysledekZaslani.Messg]));
    if VysledekZaslani.isOk then begin
      Ints[0, Radek] := 0;
    end;

  end;
end;

end.
