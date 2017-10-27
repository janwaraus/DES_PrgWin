unit FOx_prevod;

interface

uses
  Windows, Classes, Forms, Controls, SysUtils, Variants, DateUtils, Registry, Printers, Dialogs, FOx_main;

type
  TdmPrevod = class(TDataModule)
  private
    procedure FOxPrevod(Radek: integer);
  public
    procedure PrevedFOx;
  end;

var
  dmPrevod: TdmPrevod;

implementation

{$R *.dfm}

uses FOx_common, frxExportSynPDF;

// ------------------------------------------------------------------------------------------------

procedure TdmPrevod.PrevedFOx;
var
  Radek: integer;
begin
  Screen.Cursor := crHourGlass;
  with fmMain do try
    dmCommon.Zprava('P�evod do PDF');
    with asgMain do begin
      dmCommon.Zprava(Format('Po�et doklad� k p�evodu: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));
      Screen.Cursor := crHourGlass;
      apnPrevod.Visible := False;
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
        if Ints[0, Radek] = 1 then FOxPrevod(Radek);
      end;  // for
    end;  // with asgMain
// konec hlavn� smy�ky
  finally
    apbProgress.Position := 0;
    apbProgress.Visible := False;
    apnPrevod.Visible := True;
    asgMain.Visible := False;
    lbxLog.Visible := True;
    Screen.Cursor := crDefault;
    dmCommon.Zprava('P�evod do PDF ukon�en');
  end;
end;

// ------------------------------------------------------------------------------------------------

procedure TdmPrevod.FOxPrevod(Radek: integer);
// podle z�lohov�ho listu v Ab�e vytvo�� formul�� v PDF
var
  OutFileName,
  OutDir,
  AbraKod,
  SQLStr: AnsiString;
  i: integer;
  Zaplaceno: double;
  frxSynPDFExport: TfrxSynPDFExport;
begin
  with fmMain, asgMain do begin
    with qrAbra do begin
// �daje z FOx do glob�ln�ch prom�nn�ch
      Close;
      SQLStr := 'SELECT F.Code, F.Name, Street, City, PostCode, OrgIdentNumber, VATIdentNumber, II.ID,'
       + ' D.Code || ''-'' || lpad(II.OrdNumber, 4, ''0'') || ''/'' || substring(P.Code from 3 for 2) AS Cislo,'
       + ' DocDate$DATE, DueDate$DATE, VATDate$DATE, LocalAmount, LocalPaidAmount'
      + ' FROM Firms F, Addresses A, IssuedInvoices II, DocQueues D, Periods P'
      + ' WHERE F.ID = II.Firm_ID'
      + ' AND A.ID = F.ResidenceAddress_ID'
      + ' AND D.ID = II.DocQueue_ID'
      + ' AND P.ID = II.Period_ID'
      + ' AND F.Hidden = ''N''' ;
      SQLStr := SQLStr + ' AND II.Period_ID = ' + Ap + Period_Id + Ap
      + ' AND II.OrdNumber = ' + Cells[2, Radek];
      if rbInternet.Checked then SQLStr := SQLStr + ' AND DocQueue_ID = ' + Ap + FO4Queue_Id + Ap
      else SQLStr := SQLStr + ' AND DocQueue_ID = ' + Ap + FO2Queue_Id + Ap;
      SQL.Text := SQLStr;
      Open;
      if RecordCount = 0 then begin
        dmCommon.Zprava(Format('Neexistuje doklad s ��slem %d nebo z�kazn�k %s.', [Ints[2, Radek], Cells[4, Radek]]));
        Close;
        Exit;
      end;
      CisloFO := FieldByName('Cislo').AsString;
      if Trim(FieldByName('Code').AsString) = '' then begin
        dmCommon.Zprava(Format('Doklad %s: z�kazn�k nem� k�d Abry.', [CisloFO]));
        Close;
        Exit;
      end;
      AbraKod := FieldByName('Code').AsString;
      OJmeno := FieldByName('Name').AsString;
      OUlice := FieldByName('Street').AsString;
      OObec := FieldByName('PostCode').AsString + ' ' + FieldByName('City').AsString;
      OICO := FieldByName('OrgIdentNumber').AsString;
      ODIC := FieldByName('VATIdentNumber').AsString;
      FO_Id := FieldByName('ID').AsString;
      DatumDokladu := FieldByName('DocDate$DATE').AsFloat;
      DatumSplatnosti := FieldByName('DueDate$DATE').AsFloat;
      DatumPlneni := FieldByName('VATDate$DATE').AsFloat;
      Vystaveni := FormatDateTime('dd.mm.yyyy', DatumDokladu);
      Plneni := FormatDateTime('dd.mm.yyyy', DatumPlneni);
      Splatnost := FormatDateTime('dd.mm.yyyy', DatumSplatnosti);
      VS := Cells[1, Radek];
      Celkem := FieldByName('LocalAmount').AsFloat;
      Zaplaceno := FieldByName('LocalPaidAmount').AsFloat;
      Close;
//      Cislo := Format('FO4-%4.4d/%d', [Ints[2, Radek], YearOf(DatumDokladu)]);
// v�echny Firm_Id pro Abrak�d firmy
      SQLStr := 'SELECT * FROM DE$_Code_To_Firm_Id (' + Ap + AbraKod + ApZ;
      SQL.Text := SQLStr;
      Open;
      if Fields[0].AsString = 'MULTIPLE' then begin
        dmCommon.Zprava(Format('%s (%s): V�ce z�kazn�k� pro k�d %s.', [OJmeno, VS, AbraKod]));
        Close;
        Exit;
      end;
      Saldo := 0;
// a saldo pro v�echny Firm_Id (saldo je z�porn�, pokud z�kazn�k dlu��)
      while not EOF do with qrAdresa do begin
        Close;
        SQLStr := 'SELECT SaldoPo + SaldoZLPo + Ucet325 FROM DE$_Firm_Totals (' + Ap + qrAbra.Fields[0].AsString + ApC + FloatToStr(DatumDokladu) + ')';
        SQL.Text := SQLStr;
        Open;
        Saldo := Saldo + Fields[0].AsFloat;
        qrAbra.Next;
      end; // while not EOF do with qrAdresa
    end;  // with qrAbra
    Saldo := Saldo + Zaplaceno;          // Saldo je po splatnosti (SaldoPo), je-li faktura u� zaplacena, p�i�te se platba
    Zaplatit := Celkem - Saldo;          // Celkem k �hrad� = Celkem za fakt. obdob� - Z�statek minul�ch obdob�(saldo)
// text na fakturu
    if Saldo > 0 then Platek := 'p�eplatek'
    else if Saldo < 0 then Platek := 'nedoplatek'
    else Platek := ' ';
    if Zaplatit < 0 then Zaplatit := 0;
// �daje z tabulky Smlouvy do glob�ln�ch prom�nn�ch
    with qrMain do begin
      Close;
// u FOx je VS ��slo smlouvy
      SQLStr := 'SELECT Postal_name, Postal_street, Postal_PSC, Postal_city FROM customers Cu, contracts C'
      + ' WHERE Cu.Id = C.Customer_Id'
      + ' AND C.Number = ' + Ap + VS + Ap;
      SQL.Text := SQLStr;
      Open;
      PJmeno := FieldByName('Postal_name').AsString;
      PUlice := FieldByName('Postal_street').AsString;
      PObec := FieldByName('Postal_PSC').AsString + ' ' + FieldByName('Postal_city').AsString;
      Close;
    end;  // with qrMain
// zas�lac� adresa
    if (PJmeno = '') or (PObec = '') then begin
      PJmeno := OJmeno;
      PUlice := OUlice;
      PObec := OObec;
    end;
// adres�� pro ukl�d�n� faktur v PDF nemus� existovat
    if not DirectoryExists(PDFDir) then CreateDir(PDFDir);           // PDFDir je v FI.ini
    OutDir := PDFDir + Format('\%4d', [aseRok.Value]);
    if not DirectoryExists(OutDir) then CreateDir(OutDir);
    OutDir := OutDir + Format('\%2.2d', [aseMesic.Value]);
    if not DirectoryExists(OutDir) then CreateDir(OutDir);
// jm�no souboru s fakturou
//    OutFileName := OutDir + Copy(CisloFO, 1, 8) + '.pdf';            // bez roku
    OutFileName := Format('%s\%s.pdf', [OutDir, Copy(CisloFO, 1, 8)]);            // bez roku
// soubor u� existuje
    if FileExists(OutFileName) then
      if cbNeprepisovat.Checked then begin
        dmCommon.Zprava(Format('%s (%s): Soubor %s u� existuje.', [OJmeno, VS, OutFileName]));
        Exit;
      end else DeleteFile(OutFileName);
// vytvo�en� faktura se zpracuje do vlastn�ho formul��e a p�evede se do PDF
      frxReport.LoadFromFile(ExtractFilePath(ParamStr(0)) + 'FOxdoPDF.fr3');
      frxReport.PrepareReport;
//  frxReport.ShowPreparedReport;
// ulo�en�
      frxSynPDFExport := TfrxSynPDFExport.Create(nil);
      with frxSynPDFExport do try
        FileName := OutFileName;
        Title := 'Da�ov� doklad za kredit';
        Author := 'Dru�stvo Eurosignal';
        EmbeddedFonts := False;
        Compressed := True;
        OpenAfterExport := False;
        ShowDialog := False;
        ShowProgress := False;
        PDFA := True; // important
        frxReport.Export(frxSynPDFExport);
      finally
        Free;
      end;
// �ek�n� na soubor - max. 5s
    for i := 1 to 50 do begin
      if FileExists(OutFileName) then Break;
      Sleep(100);
    end;
// hotovo
    if not FileExists(OutFileName) then
      dmCommon.Zprava(Format('%s (%s): Nepoda�ilo se vytvo�it soubor %s.', [OJmeno, VS, OutFileName]))
    else begin
      dmCommon.Zprava(Format('%s (%s): Vytvo�en soubor %s.', [OJmeno, VS, OutFileName]));
      Ints[0, Radek] := 0;
      Row := Radek;
    end;
  end;  // with fmMain
end;

end.

