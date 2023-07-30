unit ZLprevod;

interface

uses
  Windows, Classes, Forms, Controls, SysUtils, Variants, DateUtils, Registry, Printers, Dialogs, ZLmain;

type
  TdmPrevod = class(TDataModule)
  private
    procedure ZLPrevod(Radek: integer);
  public
    procedure PrevedZL;
  end;

var
  dmPrevod: TdmPrevod;

implementation

{$R *.dfm}

uses ZLcommon, frxExportSynPDF;

// ------------------------------------------------------------------------------------------------

procedure TdmPrevod.PrevedZL;
var
  Radek: integer;
begin
  Screen.Cursor := crHourGlass;
// k p�evodu se pou�ije FastReport se SynPDF
  with fmMain, asgMain do try
    fmMain.Zprava(Format('P�evod ZL do PDF od ��sla %s do %s', [aedOd.Text, aedDo.Text]));
    fmMain.Zprava(Format('Po�et ZL k p�evodu: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));
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
      if Ints[0, Radek] = 1 then ZLPrevod(Radek);
    end;  // for
// konec hlavn� smy�ky
  finally
    Printer.PrinterIndex := -1;                  // default
    apbProgress.Position := 0;
    apbProgress.Visible := False;
    apnPrevod.Visible := True;
    asgMain.Visible := False;
    lbxLog.Visible := True;
    Screen.Cursor := crDefault;
    fmMain.Zprava('P�evod ZL do PDF ukon�en');
  end;
end;

// ------------------------------------------------------------------------------------------------

procedure TdmPrevod.ZLPrevod(Radek: integer);
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
// �daje z ZL do glob�ln�ch prom�nn�ch
      Close;
      SQLStr := 'SELECT F.Code, F.Name, Street, City, PostCode, OrgIdentNumber, VATIdentNumber, II.ID, LocalAmount, LocalPaidAmount,'
       + ' DocDate$DATE, DueDate$DATE'
      + ' FROM Firms F, Addresses A, IssuedDInvoices II, Periods P'
      + ' WHERE F.ID = II.Firm_ID'
      + ' AND A.ID = F.ResidenceAddress_ID'
      + ' AND P.ID = II.Period_ID'
      + ' AND F.Hidden = ''N'''
      + ' AND F.Firm_ID IS NULL'                            // posledn�, bez n�sledovn�ka
      + ' AND P.Code = ' + Ap + Copy(Cells[2, Radek], Pos('/', Cells[2, Radek])+1, 4) + Ap
      + ' AND II.OrdNumber = ' + Copy(Cells[2, Radek], 0, Pos('/', Cells[2, Radek])-1)
      + ' AND II.DocQueue_ID = ' + Ap + AbraEnt.getDocQueue('Code=ZL1').ID + Ap;
      SQL.Text := SQLStr;
      Open;
      if RecordCount = 0 then begin
        fmMain.Zprava(Format('Neexistuje ZL1-%s nebo z�kazn�k %s.', [Copy(Cells[2, Radek], 0, Pos('/', Cells[2, Radek])-1), Cells[4, Radek]]));
        Close;
        Exit;
      end;
      if Trim(FieldByName('Code').AsString) = '' then begin
        fmMain.Zprava(Format('ZL1-%s: z�kazn�k nem� k�d Abry.', [Copy(Cells[2, Radek], 0, Pos('/', Cells[2, Radek])-1)]));
        Close;
        Exit;
      end;
      AbraKod := FieldByName('Code').AsString;
      OJmeno := UTF8Decode(FieldByName('Name').AsString);
      OUlice := UTF8Decode(FieldByName('Street').AsString);
      OObec := UTF8Decode(FieldByName('PostCode').AsString + ' ' + FieldByName('City').AsString);
      OICO := FieldByName('OrgIdentNumber').AsString;
      ODIC := FieldByName('VATIdentNumber').AsString;
      ID := FieldByName('ID').AsString;
      DatumDokladu := FieldByName('DocDate$DATE').AsFloat;
      DatumSplatnosti := FieldByName('DueDate$DATE').AsFloat;
      Vystaveni := FormatDateTime('dd.mm.yyyy', DatumDokladu);
      Splatnost := FormatDateTime('dd.mm.yyyy', DatumSplatnosti);
      VS := Cells[1, Radek];
      Celkem := FieldByName('LocalAmount').AsFloat;
      Zaplaceno := FieldByName('LocalPaidAmount').AsFloat;
      Close;
      SS := Format('%6.6d%2.2d', [StrToInt(Copy(Cells[2, Radek], 0, Pos('/', Cells[2, Radek])-1)), aseRok.Value - 2000]);
      Cislo := Format('ZL1-%4.4d/%d', [StrToInt(Copy(Cells[2, Radek], 0, Pos('/', Cells[2, Radek])-1)), aseRok.Value]);
// v�echny Firm_Id pro Abrak�d firmy
      SQLStr := 'SELECT * FROM DE$_Code_To_Firm_Id (' + Ap + AbraKod + ApZ;
      SQL.Text := SQLStr;
      Open;
      if Fields[0].AsString = 'MULTIPLE' then begin
        fmMain.Zprava(Format('%s (%s): V�ce z�kazn�k� pro k�d %s.', [OJmeno, VS, AbraKod]));
        Close;
        Exit;
      end;
      Saldo := 0;
// a saldo pro v�echny Firm_Id (saldo je z�porn�, pokud z�kazn�k dlu��)
      while not EOF do with qrAdresa do begin
        Close;
        SQLStr := 'SELECT SaldoPo + SaldoZLPo + Ucet325 FROM DE$_Firm_Totals (' + Ap + qrAbra.Fields[0].AsString + ApC + FloatToStr(Date) + ')';
        SQL.Text := SQLStr;
        Open;
        Saldo := Saldo + Fields[0].AsFloat;
        qrAbra.Next;
      end; // while not EOF do with qrAdresa
    end;  // with qrAbra
// pr�v� p�ev�d�n� ZL m��e b�t p�ed splatnost�
    if Date <= DatumSplatnosti then begin
      Saldo := Saldo_PredSpl + Zaplaceno;          // Saldo je po splatnosti (SaldoPo), je-li faktura u� zaplacena, p�i�te se platba
      Zaplatit := Celkem - Saldo;          // Celkem k �hrad� = Celkem za fakt. obdob� - Z�statek minul�ch obdob�(saldo)
// anebo je po splatnosti
    end else begin
      Zaplatit := -Saldo;
      Saldo := Saldo + Celkem;             // ��stka faktury se ode�te ze salda, aby tam nebyla dvakr�t

{
	  Saldo_PoSpl = Saldo_PredSpl + Zaplaceno - Celkem (po u� zapo��t� tuto fa)
	  Saldo_PredSpl = Saldo_PoSpl - Zaplaceno + Celkem (to je �prava p�edchoz�ho vzorce)

	  kdy� je "Saldo = Saldo_PoSpl + Celkem" tak se to rovn� "Saldo_PredSpl + Zaplaceno - Celkem + Celkem = Saldo_PredSpl + Zaplaceno" a sed� to
	  
	  A v�po�tov�, jak slo�it Saldo kdy� m�m z DB Saldo_PoSpl:
	  1) Saldo = Saldo_PredSpl + Zaplaceno = Saldo_PoSpl - Zaplaceno + Celkem + Zaplaceno = Saldo_PoSpl + Celkem (!!!)
	  
	  A v�po�tov�, jak slo�it Zaplatit kdy� m�m z DB Saldo_PoSpl:
	  Zaplatit := Celkem - Saldo; a tedy
	  1) Zaplatit := Celkem - Saldo_PredSpl - Zaplaceno = Celkem - (Saldo_PoSpl - Zaplaceno + Celkem) - Zaplaceno = - Saldo_PoSpl (!!!)
	  2) Zaplatit := Celkem - (Saldo_PoSpl + Celkem ) = - Saldo_PoSpl (!!!)
}

    end;
// text na ZL
    if Saldo > 0 then Platek := 'p�eplatek'
    else if Saldo < 0 then Platek := 'nedoplatek'
    else Platek := ' ';
    if Zaplatit < 0 then Zaplatit := 0;
// �daje z tabulky Smlouvy do glob�ln�ch prom�nn�ch
    with qrMain do begin
      Close;
      SQLStr := 'SELECT Postal_name, Postal_street, Postal_PSC, Postal_city FROM customers'
      + ' WHERE Variable_symbol = ' + Ap + VS + Ap;
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
    OutDir := PDFDir + Format('\%4d', [YearOf(DatumDokladu)]);
    if not DirectoryExists(OutDir) then CreateDir(OutDir);
    OutDir := OutDir + Format('\%2.2d', [MonthOf(DatumDokladu)]);
    if not DirectoryExists(OutDir) then CreateDir(OutDir);
// jm�no souboru s fakturou
    OutFileName := OutDir + Format('\ZL1-%4.4d.pdf', [StrToInt(Copy(Cells[2, Radek], 0, Pos('/', Cells[2, Radek])-1))]);
// soubor u� existuje
    if FileExists(OutFileName) then
      if cbNeprepisovat.Checked then begin
        fmMain.Zprava(Format('%s (%s): Soubor %s u� existuje.', [Cells[4, Radek], Cells[1, Radek], OutFileName]));
        Exit;
      end else DeleteFile(OutFileName);
// vytvo�en� faktura se zpracuje do vlastn�ho formul��e a p�evede se do PDF
// ulo�en� pomoc� Synopse
    frxReport.LoadFromFile(ExtractFilePath(ParamStr(0)) + 'FR3\ZLdoPDF.fr3');
    frxReport.PrepareReport;
//    frxReport.ShowPreparedReport;
// ulo�en�
    frxSynPDFExport := TfrxSynPDFExport.Create(nil);
    with frxSynPDFExport do try
      FileName := OutFileName;
      Title := 'Z�lohov� list na p�ipojen� k internetu';
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
      fmMain.Zprava(Format('%s (%s): Nepoda�ilo se vytvo�it soubor %s.', [OJmeno, VS, OutFileName]))
    else begin
      fmMain.Zprava(Format('%s (%s): Vytvo�en soubor %s.', [OJmeno, VS, OutFileName]));
      Ints[0, Radek] := 0;
      Row := Radek;
    end;
  end;  // with fmMain
end;

end.

