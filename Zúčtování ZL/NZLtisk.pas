unit NZLtisk;

interface

uses
  Windows, Messages, Forms, Controls, Classes, Dialogs, SysUtils, DateUtils, Printers,
  DesUtils;

type
  TdmTisk = class(TDataModule)
  public
    procedure TiskniFO3;
  end;

var
  dmTisk: TdmTisk;

implementation

{$R *.dfm}

uses DesInvoices, NZLmain, NZLcommon;

// ------------------------------------------------------------------------------------------------

procedure TdmTisk.TiskniFO3;
// použije data z asgMain
var
  Radek: integer;
  Faktura : TDesInvoice;
  VysledekPrevedeni : TDesResult;
begin
  with fmMain do try
    fmMain.Zprava('Tisk FO3');
    with asgMain do begin
      fmMain.Zprava(Format('Poèet FO3 k tisku: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));
      Screen.Cursor := crHourGlass;
      apnTisk.Visible := False;
      lbPozor1.Visible := False;
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
          lbPozor1.Visible := True;
          btVytvorit.Enabled := True;
          asgMain.Visible := True;
          lbxLog.Visible := False;
          Break;
        end;

        if Ints[0, Radek] = 1 then begin  // pokud zaškrtnuto, pøevádíme fa do PDF

          Faktura := TDesInvoice.create(asgMain.Cells[10, Radek]);
          Faktura.reportAry['CisloZL'] := asgMain.Cells[1, Radek];
          Faktura.reportAry['CastkaZalohy'] := asgMain.Floats[3, Radek];

          VysledekPrevedeni := Faktura.printByFr3('FO3doPDF-104.fr3');

          fmMain.Zprava(Format('%s (%s): %s', [asgMain.Cells[6, Radek], asgMain.Cells[7, Radek], VysledekPrevedeni.Messg]));

          if VysledekPrevedeni.isOk then begin
            asgMain.Ints[0, Radek] := 0;
          end
          else
          begin
            if Dialogs.MessageDlg( 'Chyba pøi tisku: '
              + VysledekPrevedeni.Messg + sLineBreak + 'Pokraèovat?',
              mtConfirmation, [mbYes, mbNo], 0 ) = mrNo then Prerusit := True;
          end;
        end;
      end;  // for
    end;
// konec hlavní smyèky
  finally
    Printer.PrinterIndex := -1;
    apbProgress.Position := 0;
    apbProgress.Visible := False;
    lbPozor1.Visible := True;
    apnTisk.Visible := True;
    asgMain.Visible := False;
    lbxLog.Visible := True;
    Screen.Cursor := crDefault;                                             // default
    fmMain.Zprava('Tisk FO3 ukonèen');
  end;
end;

end.

