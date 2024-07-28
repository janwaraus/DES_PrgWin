unit NZL2D;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms, Dialogs, Grids, BaseGrid, AdvGrid, NZL, AdvObj,
  AdvUtil;

type
  TfmZLDetail = class(TForm)
    asgDetail: TAdvStringGrid;
    procedure FormShow(Sender: TObject);
    procedure asgDetailGetAlignment(Sender: TObject; ARow, ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
  end;

var
  fmZLDetail: TfmZLDetail;

implementation

uses DesUtils;

{$R *.dfm}

procedure TfmZLDetail.FormShow(Sender: TObject);
var
  SQLStr: string;
  fmWidth,
  R: integer;
begin
  Screen.Cursor := crHourGlass;
  with DesU.qrAbra, asgDetail do begin
    SQLStr := 'SELECT IDI.DueDate$DATE AS Datum, DQ.Code AS Rada, IDI.OrdNumber AS Cislo, P.Code AS Rok, IDI.Description AS Text,'
    + ' IDI.LocalAmount AS Vystaveno, IDI.LocalPaidAmount AS Zaplaceno'
    + ' FROM IssuedDInvoices IDI'
    + ' INNER JOIN DocQueues DQ ON IDI.DocQueue_ID = DQ.Id'               // øada dokladù
    + ' INNER JOIN Periods P ON IDI.Period_ID = P.Id'                     // rok
    + ' WHERE IDI.DocDate$DATE <= ' + IntToStr(Trunc(StrToDate(fmZL.deDatumDo.Text)))
    + ' AND IDI.DocDate$DATE >= ' + IntToStr(Trunc(StrToDate(fmZL.deDatumOd.Text)))
    + ' AND IDI.Firm_ID = (SELECT ID FROM Firms WHERE Id = ' + Ap + fmZL.asgPohledavky.Cells[4, fmZL.Radek] + ApZ
    + ' AND IDI.LocalAmount - IDI.LocalPaidAmount > 0';
  //  + ' ORDER BY IDI.DocDate$DATE';
    SQLStr := SQLStr + ' UNION ALL'
    + ' SELECT II.DueDate$DATE AS Datum, DQ.Code AS Rada, II.OrdNumber AS Cislo, P.Code AS Rok, II.Description AS Text,'
    + ' II.LocalAmount AS Vystaveno, II.LocalPaidAmount AS Zaplaceno'
    + ' FROM IssuedInvoices II'
    + ' INNER JOIN DocQueues DQ ON II.DocQueue_ID = DQ.Id'               // øada dokladù
    + ' INNER JOIN Periods P ON II.Period_ID = P.Id'                     // rok
    + ' WHERE II.DocDate$DATE <= ' + IntToStr(Trunc(StrToDate(fmZL.deDatumDo.Text)))
    + ' AND II.DocDate$DATE >= ' + IntToStr(Trunc(StrToDate(fmZL.deDatumOd.Text)))
    + ' AND II.Firm_ID = (SELECT ID FROM Firms WHERE Id = ' + Ap + fmZL.asgPohledavky.Cells[4, fmZL.Radek] + ApZ
    + ' AND II.LocalAmount - II.LocalPaidAmount > 0';
 //   + ' ORDER BY II.DocDate$DATE';
    Close;
    SQL.Text := SQLStr;
    Open;
    R := 0;                                                                                        
    while not EOF do begin
      Inc(R);
      RowCount := R + 1;
      FixedRows := 1;
      Cells[0, R] := FormatDateTime('dd.mm.yyyy', FieldByName('Datum').AsFloat);
      Cells[1, R] := FieldByName('Rada').AsString + '-' + FieldByName('Cislo').AsString
       + '/' + Copy(FieldByName('Rok').AsString, 3, 2);
      Floats[2, R] := FieldByName('Vystaveno').AsFloat;
      Floats[3, R] := FieldByName('Zaplaceno').AsFloat;
      Cells[4, R] := FieldByName('Text').AsString;
      Next;
    end;
  end;
  with asgDetail do begin
    AutoSize := True;
    Caption := fmZL.asgPohledavky.Cells[0, fmZL.Radek];
    fmWidth := 0;
    for R := 0 to ColCount-1 do fmWidth := fmWidth +  ColWidths[R];
    fmZLDetail.ClientWidth := fmWidth;
    fmZLDetail.ClientHeight := RowCount * 18;
  end;
  Screen.Cursor := crDefault;
end;

procedure TfmZLDetail.asgDetailGetAlignment(Sender: TObject; ARow, ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
  if ARow = 0 then HAlign := taLeftJustify
  else if (ACol in [2..4]) then HAlign := taRightJustify;
end;

end.
