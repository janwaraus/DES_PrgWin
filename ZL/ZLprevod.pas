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
// k pøevodu se použije FastReport se SynPDF
  with fmMain, asgMain do try
    fmMain.Zprava(Format('Pøevod ZL do PDF od èísla %s do %s', [aedOd.Text, aedDo.Text]));
    fmMain.Zprava(Format('Poèet ZL k pøevodu: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));
    Screen.Cursor := crHourGlass;
    apnPrevod.Visible := False;
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
        btVytvorit.Enabled := True;
        asgMain.Visible := True;
        lbxLog.Visible := False;
        Break;
      end;
      if Ints[0, Radek] = 1 then ZLPrevod(Radek);
    end;  // for
// konec hlavní smyèky
  finally
    Printer.PrinterIndex := -1;                  // default
    apbProgress.Position := 0;
    apbProgress.Visible := False;
    apnPrevod.Visible := True;
    asgMain.Visible := False;
    lbxLog.Visible := True;
    Screen.Cursor := crDefault;
    fmMain.Zprava('Pøevod ZL do PDF ukonèen');
  end;
end;

// ------------------------------------------------------------------------------------------------

procedure TdmPrevod.ZLPrevod(Radek: integer);
// podle zálohového listu v Abøe vytvoøí formuláø v PDF
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
// údaje z ZL do globálních promìnných
      Close;
      SQLStr := 'SELECT F.Code, F.Name, Street, City, PostCode, OrgIdentNumber, VATIdentNumber, II.ID, LocalAmount, LocalPaidAmount,'
       + ' DocDate$DATE, DueDate$DATE'
      + ' FROM Firms F, Addresses A, IssuedDInvoices II, Periods P'
      + ' WHERE F.ID = II.Firm_ID'
      + ' AND A.ID = F.ResidenceAddress_ID'
      + ' AND P.ID = II.Period_ID'
      + ' AND F.Hidden = ''N'''
      + ' AND F.Firm_ID IS NULL'                            // poslední, bez následovníka
      + ' AND P.Code = ' + Ap + Copy(Cells[2, Radek], Pos('/', Cells[2, Radek])+1, 4) + Ap
      + ' AND II.OrdNumber = ' + Copy(Cells[2, Radek], 0, Pos('/', Cells[2, Radek])-1)
      + ' AND II.DocQueue_ID = ' + Ap + AbraEnt.getDocQueue('Code=ZL1').ID + Ap;
      SQL.Text := SQLStr;
      Open;
      if RecordCount = 0 then begin
        fmMain.Zprava(Format('Neexistuje ZL1-%s nebo zákazník %s.', [Copy(Cells[2, Radek], 0, Pos('/', Cells[2, Radek])-1), Cells[4, Radek]]));
        Close;
        Exit;
      end;
      if Trim(FieldByName('Code').AsString) = '' then begin
        fmMain.Zprava(Format('ZL1-%s: zákazník nemá kód Abry.', [Copy(Cells[2, Radek], 0, Pos('/', Cells[2, Radek])-1)]));
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
// všechny Firm_Id pro Abrakód firmy
      SQLStr := 'SELECT * FROM DE$_Code_To_Firm_Id (' + Ap + AbraKod + ApZ;
      SQL.Text := SQLStr;
      Open;
      if Fields[0].AsString = 'MULTIPLE' then begin
        fmMain.Zprava(Format('%s (%s): Více zákazníkù pro kód %s.', [OJmeno, VS, AbraKod]));
        Close;
        Exit;
      end;
      Saldo := 0;
// a saldo pro všechny Firm_Id (saldo je záporné, pokud zákazník dluží)
      while not EOF do with qrAdresa do begin
        Close;
        SQLStr := 'SELECT SaldoPo + SaldoZLPo + Ucet325 FROM DE$_Firm_Totals (' + Ap + qrAbra.Fields[0].AsString + ApC + FloatToStr(Date) + ')';
        SQL.Text := SQLStr;
        Open;
        Saldo := Saldo + Fields[0].AsFloat;
        qrAbra.Next;
      end; // while not EOF do with qrAdresa
    end;  // with qrAbra
// právì pøevádìný ZL mùže být pøed splatností
    if Date <= DatumSplatnosti then begin
      Saldo := Saldo_PredSpl + Zaplaceno;          // Saldo je po splatnosti (SaldoPo), je-li faktura už zaplacena, pøiète se platba
      Zaplatit := Celkem - Saldo;          // Celkem k úhradì = Celkem za fakt. období - Zùstatek minulých období(saldo)
// anebo je po splatnosti
    end else begin
      Zaplatit := -Saldo;
      Saldo := Saldo + Celkem;             // èástka faktury se odeète ze salda, aby tam nebyla dvakrát

{
	  Saldo_PoSpl = Saldo_PredSpl + Zaplaceno - Celkem (po už zapoèítá tuto fa)
	  Saldo_PredSpl = Saldo_PoSpl - Zaplaceno + Celkem (to je úprava pøedchozího vzorce)

	  když je "Saldo = Saldo_PoSpl + Celkem" tak se to rovná "Saldo_PredSpl + Zaplaceno - Celkem + Celkem = Saldo_PredSpl + Zaplaceno" a sedí to
	  
	  A výpoètovì, jak složit Saldo když mám z DB Saldo_PoSpl:
	  1) Saldo = Saldo_PredSpl + Zaplaceno = Saldo_PoSpl - Zaplaceno + Celkem + Zaplaceno = Saldo_PoSpl + Celkem (!!!)
	  
	  A výpoètovì, jak složit Zaplatit když mám z DB Saldo_PoSpl:
	  Zaplatit := Celkem - Saldo; a tedy
	  1) Zaplatit := Celkem - Saldo_PredSpl - Zaplaceno = Celkem - (Saldo_PoSpl - Zaplaceno + Celkem) - Zaplaceno = - Saldo_PoSpl (!!!)
	  2) Zaplatit := Celkem - (Saldo_PoSpl + Celkem ) = - Saldo_PoSpl (!!!)
}

    end;
// text na ZL
    if Saldo > 0 then Platek := 'pøeplatek'
    else if Saldo < 0 then Platek := 'nedoplatek'
    else Platek := ' ';
    if Zaplatit < 0 then Zaplatit := 0;
// údaje z tabulky Smlouvy do globálních promìnných
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
// zasílací adresa
    if (PJmeno = '') or (PObec = '') then begin
      PJmeno := OJmeno;
      PUlice := OUlice;
      PObec := OObec;
    end;
// adresáø pro ukládání faktur v PDF nemusí existovat
    if not DirectoryExists(PDFDir) then CreateDir(PDFDir);           // PDFDir je v FI.ini
    OutDir := PDFDir + Format('\%4d', [YearOf(DatumDokladu)]);
    if not DirectoryExists(OutDir) then CreateDir(OutDir);
    OutDir := OutDir + Format('\%2.2d', [MonthOf(DatumDokladu)]);
    if not DirectoryExists(OutDir) then CreateDir(OutDir);
// jméno souboru s fakturou
    OutFileName := OutDir + Format('\ZL1-%4.4d.pdf', [StrToInt(Copy(Cells[2, Radek], 0, Pos('/', Cells[2, Radek])-1))]);
// soubor už existuje
    if FileExists(OutFileName) then
      if cbNeprepisovat.Checked then begin
        fmMain.Zprava(Format('%s (%s): Soubor %s už existuje.', [Cells[4, Radek], Cells[1, Radek], OutFileName]));
        Exit;
      end else DeleteFile(OutFileName);
// vytvoøená faktura se zpracuje do vlastního formuláøe a pøevede se do PDF
// uložení pomocí Synopse
    frxReport.LoadFromFile(ExtractFilePath(ParamStr(0)) + 'FR3\ZLdoPDF.fr3');
    frxReport.PrepareReport;
//    frxReport.ShowPreparedReport;
// uložení
    frxSynPDFExport := TfrxSynPDFExport.Create(nil);
    with frxSynPDFExport do try
      FileName := OutFileName;
      Title := 'Zálohový list na pøipojení k internetu';
      Author := 'Družstvo Eurosignal';
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
// èekání na soubor - max. 5s
    for i := 1 to 50 do begin
      if FileExists(OutFileName) then Break;
      Sleep(100);
    end;
// hotovo
    if not FileExists(OutFileName) then
      fmMain.Zprava(Format('%s (%s): Nepodaøilo se vytvoøit soubor %s.', [OJmeno, VS, OutFileName]))
    else begin
      fmMain.Zprava(Format('%s (%s): Vytvoøen soubor %s.', [OJmeno, VS, OutFileName]));
      Ints[0, Radek] := 0;
      Row := Radek;
    end;
  end;  // with fmMain
end;

end.

