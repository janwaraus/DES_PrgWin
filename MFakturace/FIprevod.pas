unit FIprevod;

interface

uses
  Windows, Classes, Forms, Controls, SysUtils, Variants, DateUtils, Registry, Printers, Dialogs, FImain;

type
  Ch32 = array [0..31] of char;
  TdmPrevod = class(TDataModule)
  private
{$IFNDEF ABAK}
    PWDbuf: Ch32;
    JePrint2PDF: boolean;
{$ENDIF}
    procedure FakturaPrevod(Radek: integer);
  public
    procedure PrevedFaktury;
    procedure demoPrevod();
  end;

{$IFNDEF ABAK}
const
  PWD: Ch32 =
    (char($44), char($F5), char($B3), char($27), char($D5), char($4D), char($01), char($1F),
     char($42), char($78), char($5E), char($DB), char($5B), char($4D), char($31), char($09),
     char($C6), char($DE), char($DA), char($F9), char($BD), char($69), char($CC), char($5A),
     char($64), char($4F), char($F6), char($12), char($A8), char($F8), char($3F), char($55));
{$ENDIF}

var
  dmPrevod: TdmPrevod;

implementation

{$R *.dfm}

uses FIcommon, frxExportSynPDF;

// ------------------------------------------------------------------------------------------------

procedure TdmPrevod.PrevedFaktury;
var
  Radek,
  i: integer;
  Reg: TRegistry;
begin
  Screen.Cursor := crHourGlass;
// k p�evodu se pou�ije bu� Print2PDF (pro DES), nebo FastReport se SynPDF
{$IFNDEF ABAK}
  JePrint2PDF := False;
//  JePrint2PDF := not VarIsNull(FindWindow('Print2PDF_window_class', ''));
// je�t� mus� b�t tisk�rna Print2PDF
  if JePrint2PDF then for i := 0 to Printer.Printers.Count do                  // cyklus pro v�echny naistalovan� tisk�rny
    if i = Printer.Printers.Count then JePrint2PDF := False                    // nena�la se tisk�rna Print2PDF, cyklus p�ejel konec seznamu
    else if Printer.Printers[i] = 'Print2PDF' then begin
      Printer.PrinterIndex := i;                                               // Print2PDF se na�la, tak se vybere
      Break;                                                                   // a m��eme ven z cyklu
    end;
{$ENDIF}
  with fmMain do try
{$IFNDEF ABAK}
    if JePrint2PDF then begin
      if rbPodleSmlouvy.Checked then dmCommon.Zprava(Format('P�evod faktur do PDF od VS %s do %s pomoc� Print2PDF', [aedOd.Text, aedDo.Text]))
      else dmCommon.Zprava(Format('P�evod faktur do PDF od ��sla %s do %s pomoc� Print2PDF', [aedOd.Text, aedDo.Text]));
// nastaven� pro Print2PDF v registru
      Reg := TRegistry.Create;
      try
        Reg.RootKey := HKEY_CURRENT_USER;
//  p��padn� vytvo�en� kl��e
        if not Reg.KeyExists('\Software\Software602\Print2PDF\PDF\SDK') then
          Reg.CreateKey('\Software\Software602\Print2PDF\PDF\SDK');
// nastaven� voleb pro ukl�d�n� a digit�ln� podpis
        if Reg.OpenKey('\Software\Software602\Print2PDF\PDF\SDK', True) then begin
          Reg.WriteInteger('Save', 1);
          Reg.WriteInteger('ShowDialog', 0);
          Reg.WriteInteger('OutputFormatPDF', 0);
          Reg.WriteInteger('Save', 1);
          Reg.WriteInteger('UseDefaultName', 1);
          Reg.WriteInteger('OpenAcrobat', 0);
          Reg.WriteString('PDF_Title', Format('Faktura za p�ipojen� k internetu v %d. m�s�ci %d', [aseMesic.Value, aseRok.Value]));
          Reg.WriteString('PDF_Author', 'Dru�stvo Eurosignal');
          Reg.WriteInteger('DownsizeImages', 0);
          Reg.WriteInteger('FontEmbed', 2);
          Reg.WriteInteger('PDF_Security', 0);
          Reg.WriteInteger('PDF_Signature', 1);
          Reg.WriteString('SignedPDF_Signer_Name', 'Dru�stvo Eurosignal');
          PWDbuf := PWD;
          Reg.WriteBinaryData('SignedPDF_Password', PWDbuf, 32);
          Reg.WriteInteger('SignedPDF_ShowImage', 0);
          Reg.WriteString('SignedPDF_Info_Location', 'Roh��ova 23, Praha 3');
          Reg.WriteString('SignedPDF_Info_Reason_00', 'Potvrzujeme spr�vnost a �plnost t�to faktury');
          Reg.WriteInteger('Send', 0);
          Reg.WriteInteger('UseStamp', 0);
          Reg.WriteInteger('UseWatermark', 0);
        end else dmCommon.Zprava('Probl�m p�i nastavov�n� registru pro Print2PDF (OpenKey \Software\Software602\Print2PDF\PDF\SDK).');
      finally
        Reg.CloseKey;
        Reg.Free;
      end;  // try
    end else  // if JePrint2PDF
{$ENDIF}
    if rbPodleSmlouvy.Checked then
      dmCommon.Zprava(Format('P�evod faktur do PDF od VS %s do %s', [aedOd.Text, aedDo.Text]))
      else dmCommon.Zprava(Format('P�evod faktur do PDF od ��sla %s do %s', [aedOd.Text, aedDo.Text]));
    with asgMain do begin
      dmCommon.Zprava(Format('Po�et faktur k p�evodu: %d', [Trunc(ColumnSum(0, 1, RowCount-1))]));
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
        if Ints[0, Radek] = 1 then FakturaPrevod(Radek)
      end;  // for
    end;  // with asgMain
// konec hlavn� smy�ky
  finally
    Printer.PrinterIndex := -1;                  // default
    apbProgress.Position := 0;
    apbProgress.Visible := False;
    apnPrevod.Visible := True;
    asgMain.Visible := False;
    lbxLog.Visible := True;
    Screen.Cursor := crDefault;
    dmCommon.Zprava('P�evod faktur do PDF ukon�en');
  end;
end;

// ------------------------------------------------------------------------------------------------

procedure TdmPrevod.FakturaPrevod(Radek: integer);
// podle faktury v Ab�e a stavu pohled�vek vytvo�� formul�� v PDF
var
  FrfFileName,
  OutFileName,
  OutDir,
  AbraKod,
  SQLStr: AnsiString;
  Zaplaceno: double;
  i: integer;
  Reg: TRegistry;
  frxSynPDFExport: TfrxSynPDFExport;
begin
  with fmMain, fmMain.asgMain do begin
    with qrAbra do begin
// �daje z faktury do glob�ln�ch prom�nn�ch
      Close;
      SQLStr := 'SELECT Code, Name, Street, City, PostCode, OrgIdentNumber, VATIdentNumber, II.ID, II.IsReverseChargeDeclared,'
       + ' DocDate$DATE, DueDate$DATE, VATDate$DATE, LocalAmount, LocalPaidAmount'
      + ' FROM Firms F, Addresses A, IssuedInvoices II'
      + ' WHERE F.ID = II.Firm_ID'
      + ' AND A.ID = F.ResidenceAddress_ID'
      + ' AND F.Hidden = ''N''' ;
//      + ' AND F.Firm_ID IS NULL';                             // posledn�, bez n�sledovn�ka
      SQLStr := SQLStr + ' AND II.Period_ID = ' + Ap + Period_Id + Ap
      + ' AND II.OrdNumber = ' + Cells[2, Radek]
      + ' AND II.DocQueue_ID = ';



      SQLStr := SQLStr + Ap + IDocQueue_Id + Ap;

      SQL.Text := SQLStr;
      Open;

      if RecordCount = 0 then begin
        dmCommon.Zprava(Format('Neexistuje faktura %d nebo z�kazn�k %s.', [Ints[2, Radek], Cells[4, Radek]]));
        Close;
        Exit;
      end;
      if Trim(FieldByName('Code').AsString) = '' then begin
        dmCommon.Zprava(Format('Faktura %d: z�kazn�k nem� k�d Abry.', [Ints[2, Radek]]));
        Close;
        Exit;
      end;
      AbraKod := FieldByName('Code').AsString;
      OJmeno := FieldByName('Name').AsString;
      OUlice := FieldByName('Street').AsString;
      OObec := FieldByName('PostCode').AsString + ' ' + FieldByName('City').AsString;
      OICO := FieldByName('OrgIdentNumber').AsString;
      ODIC := FieldByName('VATIdentNumber').AsString;
      ID := FieldByName('ID').AsString;
      DRC := FieldByName('IsReverseChargeDeclared').AsString = 'A';
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
      SS := Format('%6.6d%2.2d', [Ints[2, Radek], aseRok.Value - 2000]);
      Mesic := MonthOf(DatumDokladu);
      FStr := 'FO1';
      Cislo := Format('%s-%5.5d/%d', [FStr, Ints[2, Radek], aseRok.Value]);

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
        SQL.Text := 'SELECT SaldoPo + SaldoZLPo + Ucet325 FROM DE$_Firm_Totals (' + Ap + qrAbra.Fields[0].AsString + ApC + FloatToStr(DatumDokladu) + ')';
        Open;
        Saldo := Saldo + Fields[0].AsFloat;
        qrAbra.Next;
      end; // while not EOF do with qrAdresa
    end;  // with qrAbra

    // pr�v� p�ev�d�n� faktura m��e b�t p�ed splatnost�
    //    if Date <= DatumSplatnosti then begin
    Saldo := Saldo + Zaplaceno;          // Saldo je po splatnosti (SaldoPo), je-li faktura u� zaplacena, p�i�te se platba
    Zaplatit := Celkem - Saldo;          // Celkem k �hrad� = Celkem za fakt. obdob� - Z�statek minul�ch obdob�(saldo)
    // anebo je po splatnosti
    {end else begin
      Zaplatit := -Saldo;
      Saldo := Saldo + Celkem;             // ��stka faktury se ode�te ze salda, aby tam nebyla dvakr�t
    end;  }

    // text na fakturu
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
    OutDir := PDFDir + Format('\%4d', [aseRok.Value]);
    if not DirectoryExists(OutDir) then CreateDir(OutDir);
    OutDir := OutDir + Format('\%2.2d', [Mesic]);
    if not DirectoryExists(OutDir) then CreateDir(OutDir);

    // jm�no souboru s fakturou
    OutFileName := OutDir + Format('\%s-%5.5d.pdf', [FStr, Ints[2, Radek]]);
    // soubor u� existuje
    if FileExists(OutFileName) then
      if cbNeprepisovat.Checked then begin
        dmCommon.Zprava(Format('%s (%s): Soubor %s u� existuje.', [Cells[4, Radek], Cells[1, Radek], OutFileName]));
        Exit;
      end else
        DeleteFile(OutFileName);
    // vytvo�en� faktura se zpracuje do vlastn�ho formul��e a p�evede se do PDF

    if JePrint2PDF then begin        // ulo�en� pomoc� Print2PDF
      // nastav� se jm�no a cesta
      Reg := TRegistry.Create;
      try
        Reg.RootKey := HKEY_CURRENT_USER;
        if Reg.OpenKey('\Software\Software602\Print2PDF\PDF', True) then
          Reg.WriteInteger('SDK_UseKey', 1)                // Print2PDF pou�ije nastaven� v ..\PDF\SDK
        else dmCommon.Zprava(VS + ': probl�m p�i nastavov�n� registru pro Print2PDF (OpenKey \Software\Software602\Print2PDF\PDF).');
        Reg.CloseKey;
        if Reg.OpenKey('\Software\Software602\Print2PDF\PDF\SDK', True) then begin
          Reg.WriteString('DefaultName', Format('%s-%5.5d.pdf', [FStr, Ints[2, Radek]]));
          Reg.WriteString('SavePath', OutDir + '\');
        end else dmCommon.Zprava(VS + ': probl�m p�i nastavov�n� registru pro Print2PDF (OpenKey \Software\Software602\Print2PDF\PDF\SDK.');
        Reg.CloseKey;
      finally
        Reg.CloseKey;
        Reg.Free;
      end;
//      frPrevod.LoadFromFile(ExtractFilePath(ParamStr(0)) + 'FOdoPDF.frf');
//      frPrevod.PrepareReport;
//      frPrevod.ShowPreparedReport;
//      frPrevod.PrintPreparedReport('1', 1);
    end else begin                                         // ulo�en� pomoc� Synopse
      frxReport.LoadFromFile(ExtractFilePath(ParamStr(0)) + 'FOsPDP.fr3');

      frxReport.PrepareReport;
//  frxReport.ShowPreparedReport;
// ulo�en�
//  frxPDFExport.FileName := OutFileName;
//  frxReport.Export(frxPDFExport);
      frxSynPDFExport := TfrxSynPDFExport.Create(nil);
      with frxSynPDFExport do try
        FileName := OutFileName;
        Title := 'Faktura za p�ipojen� k internetu';

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

    end;      // if JePrint2PDF else

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


procedure TdmPrevod.demoPrevod();
// podle faktury v Ab�e a stavu pohled�vek vytvo�� formul�� v PDF
var
  FrfFileName,
  OutFileName,
  OutDir,
  AbraKod,
  SQLStr: AnsiString;
  Zaplaceno: double;
  Radek, i: integer;
  Reg: TRegistry;
  frxSynPDFExport: TfrxSynPDFExport;
begin
  Radek := 1;

  with fmMain, fmMain.asgMain do begin
    with qrAbra do begin

      AbraKod := 'ABRAKKK';
      OJmeno := 'Pepa';
      OUlice := 'Kratka';
      OObec := '55523 Kulickov';
      OICO := 'xxx';
      ODIC := 'xxx';
      ID := 'xxx';
      DRC := false;
      DatumDokladu := 43000;
      DatumSplatnosti := 43000;
      DatumPlneni := 43000;
      Vystaveni := FormatDateTime('dd.mm.yyyy', 43000);
      Plneni := FormatDateTime('dd.mm.yyyy', 43000);
      Splatnost := FormatDateTime('dd.mm.yyyy', 43000);
      VS := '123123123';
      Celkem := 998;
      Zaplaceno := 10;

      SS := '555';
      Mesic := MonthOf(DatumDokladu);
      FStr := 'FO1';
      Cislo := Format('%s-%5.5d/%d', ['yy', 5678, 2015]);



      Saldo := 0;

    end;  // with qrAbra


    Saldo := Saldo + Zaplaceno;          // Saldo je po splatnosti (SaldoPo), je-li faktura u� zaplacena, p�i�te se platba
    Zaplatit := Celkem - Saldo;          // Celkem k �hrad� = Celkem za fakt. obdob� - Z�statek minul�ch obdob�(saldo)


    // text na fakturu
    if Saldo > 0 then Platek := 'p�eplatek'
    else if Saldo < 0 then Platek := 'nedoplatek'
    else Platek := ' ';
    if Zaplatit < 0 then Zaplatit := 0;

    // �daje z tabulky Smlouvy do glob�ln�ch prom�nn�ch

      PJmeno := 'PJmeno';
      PUlice := 'PUlice';
      PObec := 'PObec';


    // zas�lac� adresa
    if (PJmeno = '') or (PObec = '') then begin
      PJmeno := OJmeno;
      PUlice := OUlice;
      PObec := OObec;
    end;

    // adres�� pro ukl�d�n� faktur v PDF nemus� existovat
    if not DirectoryExists(PDFDir) then CreateDir(PDFDir);           // PDFDir je v FI.ini
    OutDir := PDFDir + Format('\%4d', [2017]);
    if not DirectoryExists(OutDir) then CreateDir(OutDir);
    OutDir := OutDir + Format('\%2.2d', [Mesic]);
    if not DirectoryExists(OutDir) then CreateDir(OutDir);

    // jm�no souboru s fakturou
    OutFileName := OutDir + Format('\%s-%5.5d.pdf', ['yy', 5678]);
    // soubor u� existuje
    if FileExists(OutFileName) then

        DeleteFile(OutFileName);
    // vytvo�en� faktura se zpracuje do vlastn�ho formul��e a p�evede se do PDF

    if JePrint2PDF then begin        // ulo�en� pomoc� Print2PDF

    end else begin                                         // ulo�en� pomoc� Synopse
      frxReport.LoadFromFile(ExtractFilePath(ParamStr(0)) + 'fr3\FOsPDP.fr3');

      frxReport.PrepareReport;
//  frxReport.ShowPreparedReport;
// ulo�en�
//  frxPDFExport.FileName := OutFileName;
//  frxReport.Export(frxPDFExport);
      frxSynPDFExport := TfrxSynPDFExport.Create(nil);
      with frxSynPDFExport do try
        FileName := OutFileName;
        Title := 'Faktura za p�ipojen� k internetu';

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

    end;      // if JePrint2PDF else

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

