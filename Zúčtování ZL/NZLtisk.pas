unit NZLtisk;

interface

uses
  Windows, Messages, Forms, Controls, Classes, Dialogs, SysUtils, DateUtils, Printers,
  DesUtils;

type
  TdmTisk = class(TDataModule)
  public
    procedure TiskniFOx;
  end;

var
  dmTisk: TdmTisk;

implementation

{$R *.dfm}

uses DesInvoices, NZLmain, NZLcommon;

// ------------------------------------------------------------------------------------------------

procedure TdmTisk.TiskniFOx;
var
  Radek: integer;
  Faktura : TDesInvoice;
  VysledekPrevedeni : TDesResult;
  Fr3Filename : string;

begin
  with fmMain do try
    fmMain.Zprava('Tisk FO3');
    with asgMain do begin
      fmMain.Zprava(Format('Poèet faktur k tisku: %d', [Trunc(ColumnSum(2, 1, RowCount-1))]));

      for Radek := 1 to RowCount-1 do begin
        Row := Radek;
        apbProgress.Position := Round(100 * Radek / RowCount-1);
        Application.ProcessMessages;
        if Prerusit then Break;

        if Ints[2, Radek] = 1 then begin  // pokud zaškrtnuto, pøevádíme fa do PDF

          Faktura := TDesInvoice.create(asgMain.Cells[14, Radek]);

          if Faktura.DocQueueCode = 'FO3' then begin
            Faktura.reportAry['CisloZL'] := asgMain.Cells[3, Radek];
            Faktura.reportAry['CastkaZalohy'] := asgMain.Floats[5, Radek];
            Fr3Filename := 'FO3doPDF-104.fr3';

          end else begin
            Fr3Filename := 'FOkreditDoPDF-104.fr3';
          end;

          VysledekPrevedeni := Faktura.printByFr3(Fr3Filename);

          fmMain.Zprava(Format('%s (%s): %s', [asgMain.Cells[8, Radek], asgMain.Cells[9, Radek], VysledekPrevedeni.Messg]));

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
      end;
    end;

  finally
    asgMain.Visible := False;
    lbxLog.Visible := True;
    fmMain.Zprava('Tisk FO3 ukonèen');
  end;
end;

end.

