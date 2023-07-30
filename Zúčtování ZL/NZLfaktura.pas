// od 2.9.2016 - k nezúètovaným ZL se vystaví faktura (na celou èástku)
// 7.10. pro pøipojení ZL import manager z Abry
// 20.5.2017 automatické zaplacení vygenerované faktury (zúètování zálohy)

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

uses DesUtils, AbraEntities, DesInvoices, Superobject, AArray, NZLmain, NZLcommon;

// ------------------------------------------------------------------------------------------------

procedure TdmVytvoreni.VytvorFO;
var
  Radek: integer;
begin
  with fmMain, fmMain.asgMain do try
// pøipojení k Abøe
    Screen.Cursor := crHourGlass;
    asgMain.Visible := False;
    lbxLog.Visible := True;

    fmMain.Zprava(Format('Poèet FO k vygenerování: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));
    asgMain.Visible := True;
    lbxLog.Visible := False;
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
        btVytvorit.Enabled := True;
        apbProgress.Position := 0;
        apbProgress.Visible := False;
        lbPozor1.Visible := True;
        asgMain.Visible := True;
        lbxLog.Visible := False;
        Break;
      end;
      if Ints[0, Radek] = 1 then FOAbra(Radek);
    end;  // for
// konec hlavní smyèky
  finally
    apbProgress.Position := 0;
    apbProgress.Visible := False;
    lbPozor1.Visible := True;
    asgMain.Visible := False;
    lbxLog.Visible := True;
    Screen.Cursor := crDefault;
    fmMain.Zprava('Generování faktur ukonèeno');
  end;  //  with fmMain
end;

// ------------------------------------------------------------------------------------------------

procedure TdmVytvoreni.FOAbra(Radek: integer);
// vytvoøí fakturu na zúètování èástky z vybraného ZL
var

  SQLStr :string;

  rowAA,
  iduAA: TAArray;
  abraResponseJSON,
  newID: string;
  abraResponseSO : ISuperObject;
  abraWebApiResponse : TDesResult;
  NewInvoice : TNewDesInvoiceAA;


begin



  with fmMain, asgMain, DesU.qrAbra do begin
    Close;
    SQLStr := 'SELECT F.Id AS FId, F.Code AS Abrakod, VarSymbol, IDI.FirmOffice_Id AS FOId, IDI2.Text, IDI2.RowType, IDI2.LocalTAmount,'
    + ' IDI2.BusOrder_Id, IDI2.BusTransaction_Id'
    + ' FROM IssuedDInvoices IDI'
    + ' INNER JOIN IssuedDInvoices2 IDI2 ON IDI.ID = IDI2.Parent_Id'      // øádky dokladu
    + ' INNER JOIN Firms F ON IDI.Firm_ID = F.Id'
    + ' WHERE IDI.Id = ' + Ap + Cells[8, Radek] + Ap
    + ' ORDER BY IDI2.PosIndex';
    SQL.Text := SQLStr;
    Open;
    fmMain.Zprava(Format('%s, %s.', [Cells[1, Radek], Cells[6, Radek]]));

    // hlavièka faktury
    NewInvoice := TNewDesInvoiceAA.create(Floor(deDatumDokladu.Date), FieldByName('VarSymbol').AsString);
    if Pos('ZL1', Cells[1, Radek]) > 0 then      // ZL1 se zúètuje pøes FO3, ZL pøes FO
      NewInvoice.AA['DocQueue_ID'] := AbraEnt.getDocQueue('Code=FO3').ID
    else
      NewInvoice.AA['DocQueue_ID'] := AbraEnt.getDocQueue('Code=FO').ID;
    //FOData.ValueByName('Period_ID') := AbraEnt.getPeriod('Code=' + FormatDateTime('yyyy', deDatumDokladu.Date)).ID;
    //FOData.ValueByName('DocDate$DATE') := Floor(deDatumDokladu.Date);
    NewInvoice.AA['DueDate$DATE'] := Floor(deDatumDokladu.Date) + 10;
    NewInvoice.AA['VATDate$DATE'] := Floor(StrToDateTime(Cells[4, Radek]));
    NewInvoice.AA['Description'] := 'zúètování ' + Cells[1, Radek];
    NewInvoice.AA['Firm_ID'] := FieldByName('FId').AsString;
    NewInvoice.AA['VATFromAbovePrecision'] := 6;

    // øádky faktury, z qrAbra
    while not EOF do begin
      if FieldByName('RowType').AsInteger = 0 then begin
        rowAA := NewInvoice.createNew0Row(FieldByName('Text').AsString);
      end
      else
      begin
        rowAA := NewInvoice.createNew1Row(FieldByName('Text').AsString);
        if FieldByName('RowType').AsInteger <> 4 then
          rowAA['RowType'] := FieldByName('RowType').AsString;
        rowAA['TotalPrice'] := FieldByName('LocalTAmount').AsString;
        rowAA['BusOrder_ID'] := FieldByName('BusOrder_Id').AsString;
        rowAA['BusTransaction_ID'] := FieldByName('BusTransaction_Id').AsString;
      end;
      Next;
    end;  //  while not qrAbra.EOF


    abraWebApiResponse := DesU.abraBoCreateWebApi(NewInvoice.AA, 'issuedinvoice');
    if abraWebApiResponse.isOk then begin
      abraResponseSO := SO(abraWebApiResponse.Messg);
      fmMain.Zprava(Format('%s: Vytvoøena faktura %s.', [Cells[6, Radek], abraResponseSO.S['displayname']]));
      Ints[0, Radek] := 0;
      Cells[7, Radek] := abraResponseSO.S['displayname']; // faktura
    end
    else
    begin
      fmMain.Zprava(Format('%s: Chyba pøi vytváøení faktury - %s: %s',
        [Cells[6, Radek], abraWebApiResponse.Code, abraWebApiResponse.Messg]));
      if Dialogs.MessageDlg( '(' + abraWebApiResponse.Code + ') '
         + abraWebApiResponse.Messg + sLineBreak + 'Pokraèovat?',
         mtConfirmation, [mbYes, mbNo], 0 ) = mrNo then Prerusit := True;
      Exit;
    end;

    iduAA := TAArray.Create;
    iduAA['DepositDocument_ID'] := Cells[8, Radek];
    iduAA['Amount'] := abraResponseSO.S['amount'];
    iduAA['PDocumentType'] := '03';
    iduAA['PDocument_ID'] := abraResponseSO.S['id'];
    iduAA['PaymentDate$DATE'] := Floor(StrToDate(Cells[4, Radek]));
    iduAA['AccDate$DATE'] := Floor(deDatumDokladu.Date);

    abraWebApiResponse := DesU.abraBoCreateWebApi(iduAA, 'issueddepositusage');
    if abraWebApiResponse.isOk then begin
      fmMain.Zprava(Format('%s: Zúètována záloha %s', [Cells[6, Radek], Cells[1, Radek]]));
    end else begin
      fmMain.Zprava(Format('%s: Chyba pøi zúètování zálohy %s - %s: %s',
        [Cells[6, Radek], Cells[1, Radek], abraWebApiResponse.Code, abraWebApiResponse.Messg]));
      if Dialogs.MessageDlg( '(' + abraWebApiResponse.Code + ') '
         + abraWebApiResponse.Messg + sLineBreak + 'Pokraèovat?',
         mtConfirmation, [mbYes, mbNo], 0 ) = mrNo then Prerusit := True;
    end;

    DesU.qrAbra.Close;
    AutoSize := True;
    // Colwidths[8] := 0;
  end;


end;

end.

