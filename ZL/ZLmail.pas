unit ZLmail;

interface

uses
  Forms, Controls, SysUtils, Classes, DateUtils, Math;

type
  TdmMail = class(TDataModule)
  private
    procedure ZLMail(Radek: integer);                // po�le jeden ZL
  public
    procedure PosliZL;
  end;

var
  dmMail: TdmMail;

implementation

{$R *.dfm}

uses DesUtils, ZLmain;

// ------------------------------------------------------------------------------------------------

procedure TdmMail.PosliZL;
// pou�ije data z asgMain
var
  Radek: integer;
begin
  with fmMain do try
    fmMain.Zprava(Format('Rozes�l�n� ZL od ��sla %s do %s', [aedOd.Text, aedDo.Text]));
    with asgMain do begin
      fmMain.Zprava(Format('Po�et ZL k rozesl�n�: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));
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
// 18.12.18        if Ints[0, Radek] = 1 then ZLMail(Radek);
        if (Ints[0, Radek] = 1) and (Cells[5, Radek] <> '') then ZLMail(Radek);
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
    fmMain.Zprava('Rozes�l�n� ZL ukon�eno');
  end;
end;

// ------------------------------------------------------------------------------------------------

procedure TdmMail.ZLMail(Radek: integer);
var
  MailStr,
  CisloZL,
  PDFFile: AnsiString;

  emailAddrStr,
  emailPredmet,
  emailZprava,
  emailOdesilatel,
  FullPdfFileName: string;
  VysledekZaslani : TDesResult;
begin
  with fmMain, fmMain.asgMain do begin
    CisloZL := Copy(Cells[2, Radek], 0, Pos('/', Cells[2, Radek])-1);
    FullPdfFileName := Format('%s%4d\%2.2d\ZL1-%4.4d.pdf',
     [DesU.PDF_PATH, YearOf(Floats[7, Radek]), MonthOf(Floats[7, Radek]), StrToInt(CisloZL)]);

    emailAddrStr := Cells[5, Radek];

    emailOdesilatel := 'uctarna@eurosignal.cz';
    emailPredmet := Format('Dru�stvo EUROSIGNAL, z�lohov� list na internetovou konektivitu ZL1-%4.4d/%d', [StrToInt(CisloZL), aseRok.Value]);

    emailZprava := Format('Z�lohov� list ZL1-%4.4d/%d na p�ipojen� k internetu v dal��m obdob� je v p�ilo�en�m PDF dokumentu.',
         [StrToInt(CisloZL), aseRok.Value])
      + sLineBreak + sLineBreak
      + 'P�ejeme p�kn� den'
      + sLineBreak + sLineBreak
      + 'V� Eurosignal';


    VysledekZaslani := DesU.posliPdfEmailem(FullPdfFileName, emailAddrStr, emailPredmet, emailZprava, emailOdesilatel);
    fmMain.Zprava(Format('%s (%s): %s', [Cells[4, Radek], Cells[1, Radek], VysledekZaslani.Messg]));
    if VysledekZaslani.isOk then begin
      Ints[0, Radek] := 0;
    end;

  end;
end;

end.
