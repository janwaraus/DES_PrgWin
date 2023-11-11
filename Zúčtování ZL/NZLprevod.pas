unit NZLprevod;

interface

uses
  Windows, Classes, Forms, Controls, SysUtils, Variants, DateUtils, Registry, Printers, Dialogs,
  DesUtils;

type
  TdmPrevod = class(TDataModule)
  public
    procedure PrevedFOx;
  end;

var
  dmPrevod: TdmPrevod;

implementation

{$R *.dfm}

uses DesInvoices, NZLmain, NZLcommon;

// ------------------------------------------------------------------------------------------------

procedure TdmPrevod.PrevedFOx;
var
  Radek: integer;
  Faktura : TDesInvoice;
  VysledekPrevedeni : TDesResult;
  Fr3Filename : string;

begin
  with fmMain, fmMain.asgMain do try

    fmMain.Zprava('Pøevod zúètovacích faktur do PDF');
    fmMain.Zprava(Format('Poèet faktur k pøevodu: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));

    Prerusit := False;
    apbProgress.Position := 0;
    apbProgress.Visible := True;

    for Radek := 1 to RowCount-1 do begin
      Row := Radek;
      apbProgress.Position := Round(100 * Radek / RowCount-1);
      Application.ProcessMessages;
      if Prerusit then begin
        Break;
      end;

      if Ints[0, Radek] = 1 then begin  // pokud zaškrtnuto, pøevádíme fa do PDF

        Faktura := TDesInvoice.create(asgMain.Cells[14, Radek]);

        if Faktura.DocQueueCode = 'FO3' then begin
          Faktura.reportAry['CisloZL'] := asgMain.Cells[3, Radek];
          Faktura.reportAry['CastkaZalohy'] := asgMain.Floats[5, Radek];
          Fr3Filename := 'FO3doPDF-104.fr3';

        end else begin
          Fr3Filename := 'FOkreditDoPDF-104.fr3';
        end;


        VysledekPrevedeni := Faktura.createPdfByFr3(Fr3Filename, not cbNeprepisovat.Checked);

        fmMain.Zprava(Format('%s (%s): %s', [asgMain.Cells[8, Radek], asgMain.Cells[9, Radek], VysledekPrevedeni.Messg]));
        if VysledekPrevedeni.isOk then begin
          asgMain.Ints[0, Radek] := 0;
          FontColors[1, Radek] := $000000;
          asgMain.Ints[2, Radek] := 1;
          ReadOnly[2, Radek] := False;
        end;
      end;

    end;

  finally
    //asgMain.Visible := False;
    lbxLog.Visible := True;
    fmMain.Zprava('Pøevod do PDF ukonèen');
  end;
end;

end.

