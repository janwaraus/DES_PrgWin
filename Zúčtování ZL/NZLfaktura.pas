unit NZLfaktura;

interface

uses
  Windows, Messages, Classes, Forms, Controls, SysUtils, DateUtils, Variants, Math, Dialogs;

type
  TdmVytvoreni = class(TDataModule)
  private
    procedure FOAbra(Radek: integer);
  public
    procedure VytvorFO;
  end;

var
  dmVytvoreni: TdmVytvoreni;

implementation

{$R *.dfm}

uses DesUtils, AbraEntities, DesInvoices, NZLmain, NZLcommon; //Superobject, AArray,

// ------------------------------------------------------------------------------------------------

procedure TdmVytvoreni.VytvorFO;
var
  Radek: integer;

begin
  with fmMain, fmMain.asgMain do try

    fmMain.Zprava(Format('Poèet FO k vygenerování: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));

    for Radek := 1 to RowCount-1 do begin
      Row := Radek;
      apbProgress.Position := Round(100 * Radek / RowCount-1);
      Application.ProcessMessages;
      if Prerusit then Break;

      if Ints[2, Radek] = 1 then FOAbra(Radek);
    end;

  finally
    //asgMain.Visible := False;
    lbxLog.Visible := True;
    fmMain.Zprava('Generování faktur ukonèeno');
  end;
end;

// ------------------------------------------------------------------------------------------------

procedure TdmVytvoreni.FOAbra(Radek: integer);
// vytvoøí fakturu na zúètování èástky z vybraného ZL
var

  SQLStr :string;

  //rowAA,
  //iduAA: TAArray;
  abraResponseJSON,
  newID: string;
  //abraResponseSO : ISuperObject;
  //abraWebApiResponse : TDesResult;
  NewInvoice : TNewDesInvoiceAA;
  newIDU, newII: TNewAbraBo;

begin

  with fmMain, asgMain, DesU.qrAbra do begin
    Close;
    SQLStr := 'SELECT F.Id AS FId, VarSymbol, IDI2.Text, IDI2.RowType, IDI2.LocalTAmount,'
    + ' IDI2.BusOrder_Id, IDI2.BusTransaction_Id'
    + ' FROM IssuedDInvoices IDI'
    + ' INNER JOIN IssuedDInvoices2 IDI2 ON IDI.ID = IDI2.Parent_Id'      // øádky ZL dokladu
    + ' INNER JOIN Firms F ON IDI.Firm_ID = F.Id'
    + ' WHERE IDI.Id = ' + Ap + Cells[13, Radek] + Ap
    + ' ORDER BY IDI2.PosIndex';
    SQL.Text := SQLStr;
    Open;
    fmMain.Zprava(Format('%s, %s.', [Cells[3, Radek], Cells[8, Radek]]));


    // hlavièka faktury
    newII := TNewAbraBo.Create('issuedinvoice');
    newII.addInvoiceParams(Floor(deDatumDokladu.Date));
    newII.Item['Varsymbol'] := FieldByName('VarSymbol').AsString;

    //NewInvoice := TNewDesInvoiceAA.create(Floor(deDatumDokladu.Date), FieldByName('VarSymbol').AsString);
    if Pos('ZL1', Cells[3, Radek]) > 0 then      // ZL1 se zúètuje pøes FO3, ZL pøes FO
      newII.Item['DocQueue_ID'] := AbraEnt.getDocQueue('Code=FO3').ID
    else
      newII.Item['DocQueue_ID'] := AbraEnt.getDocQueue('Code=FO').ID;

    newII.Item['DueDate$DATE'] := Floor(deDatumDokladu.Date) + 10;
    newII.Item['VATDate$DATE'] := Floor(StrToDateTime(Cells[6, Radek]));
    newII.Item['Description'] := 'zúètování ' + Cells[3, Radek];
    newII.Item['Firm_ID'] := FieldByName('FId').AsString;
    newII.Item['VATFromAbovePrecision'] := 6;

    // øádky faktury, z qrAbra
    while not EOF do begin
      if FieldByName('RowType').AsInteger = 0 then begin
        newII.createNewInvoiceRow(0, FieldByName('Text').AsString);
        //rowAA := NewInvoice.createNew0Row(FieldByName('Text').AsString);
      end
      else
      begin
        newII.createNewInvoiceRow(1, FieldByName('Text').AsString);
        //rowAA := NewInvoice.createNew1Row(FieldByName('Text').AsString);
        if FieldByName('RowType').AsInteger <> 4 then
          newII.rowItem['RowType'] := FieldByName('RowType').AsString;
        newII.rowItem['TotalPrice'] := FieldByName('LocalTAmount').AsString;
        newII.rowItem['BusOrder_ID'] := FieldByName('BusOrder_Id').AsString;
        newII.rowItem['BusTransaction_ID'] := FieldByName('BusTransaction_Id').AsString;
      end;
      Next;
    end;  //  while not qrAbra.EOF


    newII.writeToAbra;
    //abraWebApiResponse := DesU.abraBoCreateWebApi(NewInvoice.AA, 'issuedinvoice');
    if newII.WriteResult.isOk then begin
      fmMain.Zprava(Format('%s: Vytvoøena faktura %s.', [Cells[6, Radek], newII.getCreatedBoItem('displayname')]));
      Ints[2, Radek] := 0;
      Cells[9, Radek] := newII.getCreatedBoItem('displayname'); // faktura
      Cells[11, Radek] := newII.getCreatedBoItem('localamount'); // faktura
    end
    else
    begin
      fmMain.Zprava(Format('%s: Chyba pøi vytváøení faktury - %s: %s',
        [Cells[8, Radek], newII.WriteResult.Code, newII.WriteResult.Messg]));
      if Dialogs.MessageDlg( '(' + newII.WriteResult.Code + ') '
         + newII.WriteResult.Messg + sLineBreak + 'Pokraèovat?',
         mtConfirmation, [mbYes, mbNo], 0 ) = mrNo then Prerusit := True;
      Exit;
    end;


    newIDU := TNewAbraBo.Create('issueddepositusage');
    newIDU.Item['DepositDocument_ID'] := Cells[13, Radek];
    newIDU.Item['Amount'] := newII.getCreatedBoItem('amount');
    newIDU.Item['PDocumentType'] := '03';
    newIDU.Item['PDocument_ID'] := newII.getCreatedBoItem('id');
    newIDU.Item['PaymentDate$DATE'] := Floor(StrToDate(Cells[6, Radek]));
    newIDU.Item['AccDate$DATE'] := Floor(deDatumDokladu.Date);
    newIDU.writeToAbra;

    if newIDU.WriteResult.isOk then begin
      fmMain.Zprava(Format('%s: Zúètována záloha %s', [Cells[8, Radek], Cells[3, Radek]]));
    end else begin
      fmMain.Zprava(Format('%s: Chyba pøi zúètování zálohy %s - %s: %s',
        [Cells[8, Radek], Cells[3, Radek], newIDU.WriteResult.Code, newIDU.WriteResult.Messg]));
      if Dialogs.MessageDlg( '(' + newIDU.WriteResult.Code + ') '
         + newIDU.WriteResult.Messg + sLineBreak + 'Pokraèovat?',
         mtConfirmation, [mbYes, mbNo], 0 ) = mrNo then Prerusit := True;
    end;

    {
    abraWebApiResponse := DesU.abraBoCreateWebApi(NewInvoice.AA, 'issuedinvoice');
    if abraWebApiResponse.isOk then begin
      abraResponseSO := SO(abraWebApiResponse.Messg);
      fmMain.Zprava(Format('%s: Vytvoøena faktura %s.', [Cells[6, Radek], abraResponseSO.S['displayname']]));
      Ints[2, Radek] := 0;
      Cells[9, Radek] := abraResponseSO.S['displayname']; // faktura
    end
    else
    begin
      fmMain.Zprava(Format('%s: Chyba pøi vytváøení faktury - %s: %s',
        [Cells[8, Radek], abraWebApiResponse.Code, abraWebApiResponse.Messg]));
      if Dialogs.MessageDlg( '(' + abraWebApiResponse.Code + ') '
         + abraWebApiResponse.Messg + sLineBreak + 'Pokraèovat?',
         mtConfirmation, [mbYes, mbNo], 0 ) = mrNo then Prerusit := True;
      Exit;
    end;

    iduAA := TAArray.Create;
    iduAA['DepositDocument_ID'] := Cells[13, Radek];
    iduAA['Amount'] := abraResponseSO.S['amount'];
    iduAA['PDocumentType'] := '03';
    iduAA['PDocument_ID'] := abraResponseSO.S['id'];
    iduAA['PaymentDate$DATE'] := Floor(StrToDate(Cells[6, Radek]));
    iduAA['AccDate$DATE'] := Floor(deDatumDokladu.Date);

    abraWebApiResponse := DesU.abraBoCreateWebApi(iduAA, 'issueddepositusage');
    if abraWebApiResponse.isOk then begin
      fmMain.Zprava(Format('%s: Zúètována záloha %s', [Cells[8, Radek], Cells[3, Radek]]));
    end else begin
      fmMain.Zprava(Format('%s: Chyba pøi zúètování zálohy %s - %s: %s',
        [Cells[8, Radek], Cells[3, Radek], abraWebApiResponse.Code, abraWebApiResponse.Messg]));
      if Dialogs.MessageDlg( '(' + abraWebApiResponse.Code + ') '
         + abraWebApiResponse.Messg + sLineBreak + 'Pokraèovat?',
         mtConfirmation, [mbYes, mbNo], 0 ) = mrNo then Prerusit := True;
    end;
    }
    DesU.qrAbra.Close;
    // AutoSize := True;
    // Colwidths[8] := 0;
  end;


end;

end.

