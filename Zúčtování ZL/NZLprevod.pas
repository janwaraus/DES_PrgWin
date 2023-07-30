unit NZLprevod;

interface

uses
  Windows, Classes, Forms, Controls, SysUtils, Variants, DateUtils, Registry, Printers, Dialogs,
  DesUtils;

type
  TdmPrevod = class(TDataModule)
  public
    procedure PrevedFO3;
  end;

var
  dmPrevod: TdmPrevod;

implementation

{$R *.dfm}

uses DesInvoices, NZLmain, NZLcommon;

// ------------------------------------------------------------------------------------------------

procedure TdmPrevod.PrevedFO3;
var
  Radek: integer;
  Faktura : TDesInvoice;
  VysledekPrevedeni : TDesResult;
begin
  Screen.Cursor := crHourGlass;
  with fmMain do try
    fmMain.Zprava('Pøevod zúètovacích faktur do PDF');
    with asgMain do begin
      fmMain.Zprava(Format('Poèet faktur k pøevodu: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));
      Screen.Cursor := crHourGlass;
      apnPrevod.Visible := False;
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

          VysledekPrevedeni := Faktura.createPdfByFr3('FO3doPDF-104.fr3', not cbNeprepisovat.Checked);

          fmMain.Zprava(Format('%s (%s): %s', [asgMain.Cells[6, Radek], asgMain.Cells[7, Radek], VysledekPrevedeni.Messg]));
          if VysledekPrevedeni.isOk then begin
            asgMain.Ints[0, Radek] := 0;
          end;
        end;

      end;  // for
    end;  // with asgMain
// konec hlavní smyèky
  finally
    apbProgress.Position := 0;
    apbProgress.Visible := False;
    lbPozor1.Visible := True;
    apnPrevod.Visible := True;
    asgMain.Visible := False;
    lbxLog.Visible := True;
    Screen.Cursor := crDefault;
    fmMain.Zprava('Pøevod do PDF ukonèen');
  end;
end;

end.

