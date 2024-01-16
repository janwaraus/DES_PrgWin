unit NZLcommon;

interface

uses
  Windows, SysUtils, Classes, Forms, Controls, DateUtils, Math, Registry, NZLMain;

type
  TdmCommon = class(TDataModule)
  public
    procedure Plneni_asgMain;
  end;


var
  dmCommon: TdmCommon;

implementation

{$R *.dfm}

uses AdvGrid, DesUtils, AbraEntities, DesInvoices;



// ------------------------------------------------------------------------------------------------

procedure TdmCommon.Plneni_asgMain;
var
  Radek,
  IIOrdNo,
  customerId: integer;
  II_ID,
  IDI_ID,
  postalMail,
  SQLStr: string;
  pdfSendingDateTime : double;
  pdfFileExists,
  postalMailMissing : boolean;
  Faktura : TDesInvoice;
  Zalohac : TDesInvoice;

begin
  //asgMain.Visible := True;
  //lbxLog.Visible := False;

  DesU.dbAbra.Reconnect;
  with fmMain, DesU.qrAbra, asgMain do try
    ClearNormalCells;
    RowCount := 2;
    Close;

    // *****************************//
    // ***     Zúètování ZL     *** //
    // *****************************//
    if arbVytvoreni.Checked then begin
      fmMain.Zprava('Naètení nezúètovaných ZL');
      Cells[0, 0] := 'FO';
      SQLStr := 'SELECT DISTINCT DQ.Code ||''-''|| Ordnumber ||''/''|| P.Code AS Doklad,'
      + ' LocalAmount, LocalPaidAmount, LocalUsedAmount, VarSymbol,'
      + ' F.Name, F.Code, IDI.Id AS IDId, F.Id AS FId'
      + ' FROM IssuedDInvoices IDI, DocQueues DQ, Periods P, Firms F'
      + ' WHERE DocQueue_Id = DQ.Id'
      + ' AND Period_Id = P.Id'
      + ' AND IDI.Firm_Id = F.Id'
      + ' AND F.Firm_ID IS NULL'         // bez následovníkù
      + ' AND F.Hidden = ''N'''
      + ' AND LocalUsedAmount <> LocalPaidAmount'; // aby se vybraly jen nezúètované
      if cbCast.Checked then
        SQLStr := SQLStr + ' AND LocalAmount >= LocalPaidAmount' // aby se neobjevil pøeplacený ZL
      else
        SQLStr := SQLStr + ' AND LocalAmount = LocalPaidAmount';
      if acbRada.Text <> '%' then
        SQLStr := SQLStr + ' AND DQ.Code = ' + Ap + acbRada.Text + Ap;

      SQLStr := SQLStr + ' AND Period_Id = ' + Ap + AbraEnt.getPeriod('Code=' + aseRok.Text).ID + Ap
      + ' ORDER BY DocDate$Date, OrdNumber';
      SQL.Text := SQLStr;
      Open;
      Radek := 0;
      while not EOF do begin
        apbProgress.Position := Round(100 * RecNo / RecordCount);
        Application.ProcessMessages;
        if Prerusit then Break;

        Inc(Radek);
        RowCount := Radek + 1;
        AddCheckBox(2, Radek, True, True);
        Ints[2, Radek] := 1;                                                   // fajfka
        // ReadOnly[0, Radek] := True; // test disabluje checkbox
        Cells[3, Radek] := FieldByName('Doklad').AsString;
        Floats[4, Radek] := FieldByName('LocalAmount').AsFloat;
        Floats[5, Radek] := FieldByName('LocalPaidAmount').AsFloat;
        Floats[7, Radek] := FieldByName('LocalUsedAmount').AsFloat;
        Cells[8, Radek] := FieldByName('Name').AsString;
        Cells[9, Radek] := '';
        Cells[10, Radek] := FieldByName('VarSymbol').AsString;
        Cells[13, Radek] := FieldByName('IDId').AsString;
        //Cells[14, Radek] := FieldByName('FId').AsString;
        //Cells[14, Radek] := '';

        DesU.qrAbraOC.SQL.Text := 'SELECT DocDate$Date FROM LastPaymentForDocument(''10'',' + Ap + Cells[13, Radek] + ApZ;
        DesU.qrAbraOC.Open;
        Cells[6, Radek] := FormatDateTime('dd.mm.yyyy', DesU.qrAbraOC.Fields[0].AsFloat);
        DesU.qrAbraOC.Close;

        Application.ProcessMessages;
        Next;
      end;  // while not EOF
      Cells[2, 0] := '';
      AutoSize := True;
      ColWidths[0] := 0;
      ColWidths[1] := 0;
      ColWidths[12] := 0;
      ColWidths[13] := 0;
      ColWidths[14] := 0;
      ColWidths[15] := 0;
      ColWidths[16] := 0;

    end;   // if arbVytvoreni

    // *****************************//
    // ***  Pøevod, Tisk, Mail  *** //
    // *****************************//
    if not arbVytvoreni.Checked then begin
      fmMain.Zprava('Naètení vytvoøených faktur.');
      Close;
      DesU.dbAbra.Reconnect;

      SQLStr := 'SELECT II.ID as II_ID, IDI.ID AS IDI_ID '
      + ' FROM IssuedInvoices II '
      + ' LEFT JOIN IssuedDepositUsages IDU ON IDU.PDocument_ID = II.ID'
      + ' LEFT JOIN IssuedDInvoices IDI ON IDI.ID = IDU.DepositDocument_ID'
      + ' WHERE II.DocDate$Date >= ' + FloatToStr(Floor(deDatumFOxOd.Date))
      + '   AND II.DocDate$Date <= ' + FloatToStr(Floor(deDatumFOxDo.Date))
      + ' AND II.DocQueue_ID IN (';

      if chbFO2.Checked then
        SQLStr := SQLStr + Ap + AbraEnt.getDocQueue('Code=FO2').ID + ApC;

      if chbFO3.Checked then
        SQLStr := SQLStr + Ap + AbraEnt.getDocQueue('Code=FO3').ID + ApC;

      if chbFO4.Checked then
        SQLStr := SQLStr + Ap + AbraEnt.getDocQueue('Code=FO4').ID + ApC;

      SQLStr := SQLStr + ' ''1'') ';  //aby to fungovalo, když se nevybere žádná FOx øada

      if TryStrToInt(editOrdNo.Text, IIOrdNo) then
        SQLStr := SQLStr + ' AND II.OrdNumber = ' + editOrdNo.Text;

      SQLStr := SQLStr + ' ORDER BY II.DocQueue_ID, II.OrdNumber';

      SQL.Text := SQLStr;
      Open;
      Radek := 0;

      while not EOF do begin

        if Prerusit then Break;

        II_ID := FieldByName('II_ID').AsString;
        IDI_ID := FieldByName('IDI_ID').AsString;

        Faktura := TDesInvoice.create(II_ID);

        DesU.qrZakos.SQL.Text := 'SELECT cu.id, cu.postal_mail FROM customers cu, contracts co'
          + ' WHERE cu.id = co.customer_id'
          + ' AND co.number = ' + Ap + Faktura.VS + Ap;
        DesU.qrZakos.Open;
        postalMail := DesU.qrZakos.FieldByName('postal_mail').AsString;
        customerId := DesU.qrZakos.FieldByName('id').AsInteger;
        DesU.qrZakos.Close;

        pdfSendingDateTime := DesU.pdfToCustomerSendingDateTime(customerId, Faktura.CisloDokladu);
        pdfFileExists := FileExists(Faktura.getFullPdfFileName);
        postalMailMissing := Pos('@', postalMail) <= 0;  // když není vyplnìný email

        if (pdfFileExists and cbJenNezpracovaneFOx.Checked
          and ((pdfSendingDateTime > 0) or (postalMailMissing) )
          ) then begin
          Next;
          Continue;
        end;


        apbProgress.Position := Round(100 * RecNo / RecordCount);
        Application.ProcessMessages;

        Inc(Radek);
        RowCount := Radek + 1;
        AddCheckBox(0, Radek, True, True);
        AddCheckBox(2, Radek, True, True);


        if pdfFileExists then begin
          Ints[0, Radek] := 0; // odškrtnout 1. checkbox
          if (pdfSendingDateTime > 0) or (postalMailMissing)
          then Ints[2, Radek] := 0; // odškrtnout 2. checkbox, už bylo v minulosti posláno
        end
        else begin
          Ints[2, Radek] := 0; // odškrtnout 2. checkbox
          ReadOnly[2, Radek] := True; // disablovat 2. checkbox
          FontColors[1, Radek] := $808080
        end;

        Cells[1, Radek] := Faktura.getExportSubDir + Faktura.getPdfFileName;

        //asgMain.FontColors[1, i+1] := clRed; $800000

        if IDI_ID <> '' then begin
          Zalohac := TDesInvoice.create(IDI_ID, '10');

          Cells[3, Radek] := Zalohac.CisloDokladu;
          Floats[4, Radek] := Zalohac.Castka;
          Floats[5, Radek] := Zalohac.CastkaZaplaceno;
        end;
        //Cells[6, Radek] := inttostr(customerId);
        //Cells[7, Radek] := FormatDateTime('dd.mm.yyyy', pdfSendingDateTime); // floattostr(pdfSendingDateTime) + ' - ' +
        Cells[8, Radek] := Faktura.Firm.Name;
        Cells[9, Radek] := Faktura.CisloDokladu;
        Cells[10, Radek] := Faktura.VS;
        Floats[11, Radek] := Faktura.Castka;
        Cells[13, Radek] := IDI_ID;
        Cells[14, Radek] := II_ID;

        //naètení e-mail adresy a id zákazníka
        Cells[12, 0] := 'mail';
        Cells[12, Radek] := postalMail;
        if pdfSendingDateTime > 0 then
          Cells[15, Radek] := FormatDateTime('dd.mm.yyyy', pdfSendingDateTime) + ' (custId ' + inttostr(customerId) + ')'
        else
          Cells[15, Radek] := '';
        Cells[16, Radek] := FormatDateTime('dd.mm.yyyy', Faktura.DocDate);


        Application.ProcessMessages;
        Next;
      end;  // while not EOF
      Close;
      AutoSize := True;
      ColWidths[6] := 0;
      ColWidths[7] := 0;
      ColWidths[13] := 0;
      ColWidths[14] := 0;
      ColWidths[15] := 0;


      if not chbFO3.Checked then begin
        ColWidths[3] := 0;
        ColWidths[4] := 0;
        ColWidths[5] := 0;
      end;


      //Font.Color := RGB(150, 99, 99);
      //ColWidths[13] := 0;
      //ColWidths[14] := 0;
    end;  //  if not arbVytvoreni.Checked
    fmMain.Zprava(Format('Poèet faktur: %d', [Trunc(RowCount-1)]));
    {
    if ColumnSum(0, 1, RowCount-1) > 0 then
      if arbVytvoreni.Checked then btVytvorit.Caption := '&Vytvoøit'
      else if arbPrevod.Checked then btVytvorit.Caption := '&Pøevést';
    }
  except on E: Exception do
    Zprava('Neošetøená chyba: ' + E.Message);
  end;  // with qrAbra

end;

end.

