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

uses AdvGrid, DesUtils, AbraEntities;



// ------------------------------------------------------------------------------------------------

procedure TdmCommon.Plneni_asgMain;
var
  Radek: integer;
  SQLStr: string;
begin
  with fmMain do try
    asgMain.Visible := True;
    lbxLog.Visible := False;
    apnPrevod.Visible := False;
    apnTisk.Visible := False;
    apnMail.Visible := False;
    Screen.Cursor := crHourGlass;
    with DesU.qrAbra, asgMain do try
      ClearNormalCells;
      Cells[6, 0] := 'jméno';
      RowCount := 2;
      Close;
// tabulka pro generování FO
      if arbVytvoreni.Checked then begin
        fmMain.Zprava('Naètení nezúètovaných ZL');
        Cells[0, 0] := 'FO';
        SQLStr := 'SELECT DISTINCT DQ.Code ||''-''|| Ordnumber ||''/''|| P.Code AS Doklad,'
        + ' LocalAmount, LocalPaidAmount, LocalUsedAmount, F.Name, F.Code, IDI.Id AS IDId, F.Id AS FId'
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
          SQLStr := SQLStr + ' AND DQ.Code = ' + Ap + acbRada.Text + Ap
        + ' AND Period_Id = ' + Ap + AbraEnt.getPeriod('Code=' + aseRok.Text).ID + Ap
        + ' ORDER BY DocDate$Date, OrdNumber';
        SQL.Text := SQLStr;
        Open;
        Radek := 0;
        lbPozor1.Visible := False;
        apbProgress.Position := 0;
        apbProgress.Visible := True;
        while not EOF do begin
          apbProgress.Position := Round(100 * RecNo / RecordCount);
          Application.ProcessMessages;
          if Prerusit then begin
            Prerusit := False;
            apbProgress.Position := 0;
            apbProgress.Visible := False;
            lbPozor1.Visible := True;
            btVytvorit.Enabled := True;
            btKonec.Caption := '&Konec';
            Break;
          end;
          Inc(Radek);
          RowCount := Radek + 1;
          AddCheckBox(0, Radek, True, True);
          Ints[0, Radek] := 1;                                                   // fajfka
          Cells[1, Radek] := FieldByName('Doklad').AsString;
          Floats[2, Radek] := FieldByName('LocalAmount').AsFloat;
          Floats[3, Radek] := FieldByName('LocalPaidAmount').AsFloat;
          Floats[5, Radek] := FieldByName('LocalUsedAmount').AsFloat;
          Cells[6, Radek] := FieldByName('Name').AsString;
          Cells[7, Radek] := '';
          Cells[8, Radek] := FieldByName('IDId').AsString;
          Cells[9, Radek] := FieldByName('FId').AsString;
          Cells[10, Radek] := '';
          with DesU.qrAbraOC do begin        // procedura z Abry
            SQLStr := 'SELECT DocDate$Date FROM LastPaymentForDocument(''10'',' + Ap + Cells[8, Radek] + ApZ;
            SQL.Text := SQLStr;
            Open;
            Cells[4, Radek] := FormatDateTime('dd.mm.yyyy', Fields[0].AsFloat);
            Close;
          end;
          Application.ProcessMessages;
          Next;
        end;  // while not EOF
        AutoSize := True;
        //ColWidths[8] := 0;
        //ColWidths[9] := 0;
        //ColWidths[10] := 0;
        lbPozor1.Visible := True;
      end;   // if arbVytvoreni
// ne generování FO (tj. pøevod, tisk, mail), faktury FO3 se hledají podle data vytvoøení z IssuedDepositUsages
      if not arbVytvoreni.Checked then begin
        fmMain.Zprava('Naètení vytvoøených faktur.');
        Close;
        Desu.dbAbra.Reconnect;
        SQLStr := 'SELECT DISTINCT DQ.Code ||''-''|| IDI.Ordnumber ||''/''|| P1.Code AS ZL,'
        + ' IDI.LocalAmount, IDI.LocalPaidAmount, IDI.LocalUsedAmount, F.Name, F.Code AS Abrakod, IDI.Id AS IDId, II.Id AS IId, F.Id AS FId,'
        + ' ''FO3-''|| II.Ordnumber ||''/''|| P2.Code AS FO'
        + ' FROM IssuedDInvoices IDI, IssuedInvoices II, DocQueues DQ, Periods P1, Periods P2, Firms F, IssuedDepositUsages IDU'
        + ' WHERE II.DocDate$Date = ' + FloatToStr(Floor(deDatumDokladu.Date))
        + ' AND II.DocQueue_ID = ' + Ap + AbraEnt.getDocQueue('Code=FO3').ID + Ap
//        + ' AND EXISTS (SELECT DepositDocument_ID FROM IssuedDepositUsages WHERE PDocument_ID = II.Id)'
        + ' AND II.Period_Id = P2.Id'
        + ' AND II.Id = IDU.PDocument_ID'                  // 31.8.2021
//        + ' AND IDI.Id = (SELECT DepositDocument_ID FROM IssuedDepositUsages WHERE PDocument_ID = II.Id)'
        + ' AND IDI.Id = IDU.DepositDocument_ID'           // 31.8.2021
        + ' AND IDI.Firm_Id = F.Id'
        + ' AND IDI.DocQueue_Id = DQ.Id'
        + ' AND IDI.Period_Id = P1.Id'
        + ' ORDER BY II.OrdNumber';
        SQL.Text := SQLStr;
        Open;
        Radek := 0;
        apbProgress.Position := 0;
        apbProgress.Visible := True;
        while not EOF do begin
          apbProgress.Position := Round(100 * RecNo / RecordCount);
          Application.ProcessMessages;
          if Prerusit then begin
            Prerusit := False;
            apbProgress.Position := 0;
            apbProgress.Visible := False;
            btVytvorit.Enabled := True;
            btKonec.Caption := '&Konec';
            Break;
          end;
          Inc(Radek);
          RowCount := Radek + 1;
          AddCheckBox(0, Radek, True, True);
          Cells[1, Radek] := FieldByName('ZL').AsString;
          Floats[2, Radek] := FieldByName('LocalAmount').AsFloat;
          Floats[3, Radek] := FieldByName('LocalPaidAmount').AsFloat;
          Floats[5, Radek] := FieldByName('LocalUsedAmount').AsFloat;
          Cells[6, Radek] := FieldByName('Name').AsString;
          Cells[7, Radek] := FieldByName('FO').AsString;
          Cells[8, Radek] := FieldByName('IDId').AsString;
          Cells[9, Radek] := FieldByName('FId').AsString;
          Cells[10, Radek] := FieldByName('IId').AsString;
          if arbMail.Checked then with DesU.qrZakos do begin
            Cells[6, 0] := 'mail';
            SQLStr := 'SELECT DISTINCT Postal_mail AS Mail FROM customers'
            + ' WHERE Abra_code = ' + Ap + DesU.qrAbra.FieldByName('Abrakod').AsString + Ap;
            SQL.Text := SQLStr;
            Open;
            Cells[6, Radek] := FieldByName('Mail').AsString;
            Close;
          end;  //  with qrMain
          Application.ProcessMessages;
          Next;
        end;  // while not EOF
        Close;
        AutoSize := True;
        //ColWidths[4] := 0;
        //ColWidths[8] := 0;
        //ColWidths[9] := 0;
        //ColWidths[10] := 0;
      end;  //  if not arbVytvoreni.Checked
      fmMain.Zprava(Format('Poèet faktur: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));
      if ColumnSum(0, 1, RowCount-1) > 0 then
        if arbVytvoreni.Checked then btVytvorit.Caption := '&Vytvoøit'
        else if arbPrevod.Checked then btVytvorit.Caption := '&Pøevést'
        else if arbTisk.Checked then btVytvorit.Caption := '&Vytisknout'
        else if arbMail.Checked then btVytvorit.Caption := '&Odeslat';
    except on E: Exception do
      Zprava('Neošetøená chyba: ' + E.Message);
    end;  // with qrAbra
  finally
    apbProgress.Position := 0;
    apbProgress.Visible := False;
    if arbPrevod.Checked then apnPrevod.Visible := True;
    if arbTisk.Checked then apnTisk.Visible := True;
    if arbMail.Checked then apnMail.Visible := True;
    Screen.Cursor := crDefault;
    btVytvorit.Enabled := True;
    btVytvorit.SetFocus;
  end;
end;

end.

