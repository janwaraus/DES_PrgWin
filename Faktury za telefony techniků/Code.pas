unit Code;

interface

uses
  Windows, SysUtils, DateUtils, Classes, Forms, Types, Controls, Variants, Math, Dialogs, ComObj, StrUtils, rxFileUtil,
  ShellAPI, RegularExpressions, MainFT;

type
  TdmCode = class(TDataModule)
  private
    iMesic,                // mìsíc a rok vyúètování
    iRok: integer;
  public
    Doklad: ShortString;
    MailAddr,
    MailStr: AnsiString;
    PoslatMail: boolean;
    procedure PrevodPDFdoTXT(PDFFileName: string);
    procedure AnalyzaVyuctovani;
    procedure FakturaTechnika(aTechnik: string);
    procedure ConnectAbra;
    procedure UpdateTechniciXLS(aRadek: integer);
  end;

const
  Ap = chr(39);
  ApC = Ap + ',';
  ApZ = Ap + ')';
  FO_Id: string[10] = 'H000000101';
  PP_Id: string[10] = '4000000101';
  MyAddress_Id: string[10] = '7000000101';
  MyAccount_Id: string[10] = '1400000101';       // Fio
  User_Id: string[10] = '2200000101';            // automatická fakturace
  Payment_Id: string[10] = '1000000101';         // typ platby: na bankovní úèet

var
  dmCode: TdmCode;

implementation

uses DES_common;

{$R *.dfm}

// ------------------------------------------------------------------------------------------------

procedure TdmCode.PrevodPDFdoTXT(PDFFileName: string);
// Procedura pøevede PDF soubor s vyúètováním tel. hovorù od T-Mobile do textu. Pro pøevod použije program "blb2txt",
// z projektu Balabolka (http://balabolka.site/btext.htm). Program musí být umístìný v adresáøi ExePath (kde je hlavní
// program) a volá se pomocí ShellExecute. Parametrem je jméno PDF souboru od T-Mobile a výstup je Mobily.tmp v adresáøi
// ExePath

begin
  with fmMainFT do begin
// PDF soubor se pøevede do textu
    ChDir(ExePath);
    if ShellExecute(Handle, 'open', PChar('blb2txt.exe'), PChar('-f ' + PDFFileName + ' -out Mobily.tmp'), nil,
     SW_SHOWNORMAL) < 33 then begin            // 0 až 32 jsou chybové stavy ShellExecute
      Application.MessageBox(PChar('Chyba pøi pøevodu ' + odPDF.FileName), 'PDF', MB_ICONERROR + MB_OK);
      Exit;
    end;
  end;
end;

// ------------------------------------------------------------------------------------------------

procedure TdmCode.AnalyzaVyuctovani;
// Procedura nejdøív ve vyúètování najde rok a mìsíc, podle roku naplní asgTelefony. Projede vyúètování a k èíslùm
// najde techniky a útratu.

var
  Technik,
  sTmp: string;
  TxtFile: TextFile;
  iTlfRadek,
  iRadek: integer;
  reRegExpr: TRegEx;

begin
// otevøe se soubor s úètovanými èástkami
  AssignFile(TxtFile, fmMainFT.ExePath + 'Mobily.tmp');
  Reset(TxtFile);
  sTmp := '';
// úètované období je na zaèátku souboru
  while not (reRegExpr.IsMatch(sTmp, 'za období') or EOF(TxtFile)) do ReadLn(TxtFile, sTmp);
  if EOF(TxtFile) then begin
    Application.MessageBox(PChar('V souboru není úètované období'), 'Chyba souboru', MB_ICONERROR + MB_OK);
    Exit;
  end;

// období se dá do labelu
  fmMainFT.lbObdobi.Caption := Copy(sTmp, Pos('za období', sTmp)+10, Length(sTmp)-10);
// a uloží se mìsíc a rok úètování (sTmp je tøeba: za období 20.3. - 19.4.2019)
  sTmp := Copy(sTmp, Pos('- ', sTmp)+2, Length(sTmp));     // zbyde 19.4.2019
  sTmp := Copy(sTmp, Pos('.', sTmp)+1, Length(sTmp));      // zbyde 4.2019
  iMesic := StrToInt(Copy(sTmp, 1, Pos('.', sTmp)-1));
  iRok := StrToInt(Copy(sTmp, Length(sTmp)-1, 2));         // rok je na dvì místa

// s znalostí roku se otevøe soubor s telefonními èísly technikù a uloží do asgTelefony (asgTelefony.Visible := False,
// sloupec 0 je tel. èíslo, 1 nezajímavý vlastník, 2 zajímavý plátce)
  with fmMainFT.asgTelefony do begin
    ClearNormalCells;
    RowCount := 2;
    LoadFromXLS(Format('"J:\Eurosignal\telefony%2.2d.xls"', [iRok]));
    ColCount := 3;
    ColWidths[1] := 0;
    for iTlfRadek := 1 to RowCount-1 do                        // èísla bez mezer
      Cells[0, iTlfRadek] := StringReplace(Cells[0, iTlfRadek], ' ', '', [rfReplaceAll]);
  end;

// v pokraèování souboru s úètovanými èástkami se zajímavé øádky uloží do asgMain (1 je technik, 2 èíslo, 3 èástka)
  iRadek := 0;
  with fmMainFT.asgMain do begin
    ClearNormalCells;
    RowCount := 2;
    while not EOF(TxtFile) do begin
      Readln(TxtFile, sTmp);
// když je na zaèátku øádku devìt èíslic, je to úètované telefonní èíslo (sloupec 2)
      if reRegExpr.IsMatch(sTmp, '^\d{9}') then begin
        Inc(iRadek);
        RowCount := iRadek + 1;
        Cells[2, iRadek] := Copy(sTmp, 1, 9);
// v asgTelefony se k èíslu najde technik
        with fmMainFT.asgTelefony do
          for iTlfRadek := 1 to RowCount-1 do
            if (Cells[0, iTlfRadek] = fmMainFT.asgMain.Cells[2, iRadek]) then begin
              fmMainFT.asgMain.Cells[1, iRadek] := Cells[2, iTlfRadek];      // technik se zkopíruje do sloupce 1
              Break;
            end;  // if
      end else
// øádek s úètovanou èástkou má na konci Kè a pøijde do sloupce 3
      if reRegExpr.IsMatch(sTmp, 'Kè$') then begin
        Inc(iRadek);
        RowCount := iRadek + 1;
        Cells[3, iRadek] := StringReplace(sTmp, 'Kè', '', [rfReplaceAll]);
      end;
    end;  // while not EOF(TxtFile)
// konec importu ze souboru
    CloseFile(TxtFile);
  end;  // with asgMain

end;

// ------------------------------------------------------------------------------------------------

procedure TdmCode.FakturaTechnika(aTechnik: string);
// vytvoøí fakturu pro technika s kódem aTechnik
var
  Radek: integer;
  Period_Id,
  Firm_Id,
  FirmOffice_Id,
  ID: string[10];
  SQLStr: string;
  FaktObject,
  FaktData,
  RadkyFak,
  RadekFak: variant;
  PPObject,
  PPData,
  RadkyPP,
  RadekPP: variant;
  DatumDokladu: double;

begin
// jméno a adresa technika, perioda
  with fmMainFT.qrAbra do begin
    Close;
    SQLStr := 'SELECT F.Id AS FId, FO.Id AS FOId FROM Firms F, FirmOffices FO'
    + ' WHERE F.Code = ' + Ap + aTechnik + Ap
    + ' AND F.Firm_ID IS NULL'
    + ' AND F.Hidden = ''N'''
    + ' AND FO.Parent_Id = F.Id';
    SQL.Text := SQLStr;
    Open;
    if RecordCount = 0 then begin
      ShowMessage(Format('Technik %s není v adresáøi Abry.', [aTechnik]));
      Close;
      Exit;
    end;
    Firm_Id := FieldByName('FId').AsString;
    FirmOffice_Id := FieldByName('FOId').AsString;
    Close;
    SQL.Text := 'SELECT Id FROM Periods WHERE Code = ' + Ap + Copy(fmMainFT.lbObdobi.Caption, Length(fmMainFT.lbObdobi.Caption)-3, 4) + Ap;
    Open;
    Period_Id := FieldByName('Id').AsString;
    Close;
  end;  // with qrAbra
  DatumDokladu := Floor(EndOfTheMonth(StrToDate(Copy(fmMainFT.lbObdobi.Caption, Pos('- ', fmMainFT.lbObdobi.Caption)+2, Length(fmMainFT.lbObdobi.Caption)))));
// hlavièka FO
  FaktObject:= fmMainFT.AbraOLE.CreateObject('@IssuedInvoice');
  FaktData:= fmMainFT.AbraOLE.CreateValues('@IssuedInvoice');
  FaktObject.PrefillValues(FaktData);
  FaktData.ValueByName('DocQueue_ID') := FO_Id;
  FaktData.ValueByName('Period_ID') := Period_Id;
  FaktData.ValueByName('DocDate$DATE') := DatumDokladu;
  FaktData.ValueByName('AccDate$DATE') := DatumDokladu;
  FaktData.ValueByName('VATDate$DATE') := DatumDokladu;
  FaktData.ValueByName('VATAdmitDate$DATE') := DatumDokladu;
  FaktData.ValueByName('CreatedBy_ID') := User_Id;
  FaktData.ValueByName('Firm_ID') := Firm_Id;
  FaktData.ValueByName('FirmOffice_ID') := FirmOffice_Id;
  FaktData.ValueByName('Address_ID') := MyAddress_Id;
  FaktData.ValueByName('BankAccount_ID') := MyAccount_Id;
  FaktData.ValueByName('ConstSymbol_ID') := '0000308000';
//    FaktData.ValueByName('VarSymbol') := VS;
  FaktData.ValueByName('TransportationType_ID') := '1000000101';
  FaktData.ValueByName('PaymentType_ID') := Payment_Id;
  FaktData.ValueByName('DueTerm') := '10';
  FaktData.ValueByName('PricesWithVAT') := True;
  FaktData.ValueByName('VATFromAbovePrecision') := 0;
  FaktData.ValueByName('TotalRounding') := 259;                // zaokrouhlení na koruny dolù
  FaktData.ValueByName('Description') := 'telefon';
// kolekce pro øádky faktury
  RadkyFak := FaktData.Value[dmDES_common.IndexByName(FaktData, 'Rows')];
// 1. øádek prázdný
  RadekFak:= fmMainFT.AbraOLE.CreateValues('@IssuedInvoiceRow');
  RadekFak.ValueByName('Division_ID') := '1000000101';
  RadekFak.ValueByName('RowType') := '0';
  RadkyFak.Add(RadekFak);
// další øádky faktury z asgMain
  Radek := 1;
  with fmMainFT.asgMain do begin
// 2. øádek - 1. technika v asgMain
    while (Cells[1, Radek] <> aTechnik) and (Radek < RowCount-1) do Inc(Radek);
// všechny øádky jednoho technika
    while (Cells[1, Radek] = aTechnik) do begin
      RadekFak:= fmMainFT.AbraOLE.CreateValues('@IssuedInvoiceRow');
      RadekFak.ValueByName('Division_ID') := '1000000101';
      RadekFak.ValueByName('Text') := Cells[2, Radek];
      RadekFak.ValueByName('IncomeType_ID') := '2000000000';
      RadekFak.ValueByName('RowType') := '1';
      if (Cells[2, Radek] = 'platební služby') then begin
        RadekFak.ValueByName('VATRate_ID') := '00000X0000';
        RadekFak.ValueByName('VATIndex_ID') := '7000000000';
        RadekFak.ValueByName('VATRate') := '0';
        RadekFak.ValueByName('TotalPrice') := Cells[3, Radek];
      end else begin
        RadekFak.ValueByName('VATRate_ID') := '02100X0000';
        RadekFak.ValueByName('VATIndex_ID') := '6521000000';
        RadekFak.ValueByName('VATRate') := '21';
        RadekFak.ValueByName('TotalPrice') := FloatToStr(Floats[3, Radek] * 1.21);
      end;
      RadkyFak.Add(RadekFak);
      Inc(Radek);
    end;
// vytvoøení FO
    try
      ID := FaktObject.CreateNewFromValues(FaktData);
      FaktData := FaktObject.GetValues(ID);
    except
      on E: Exception do begin
        ShowMessage('Chyba pøi vytváøení faktury.' + ^M + E.Message);
        Exit;
      end;
    end;  // try
// znovu poèáteèní øádek a uložení výsledku
    Radek := 1;
    while (Cells[1, Radek] <> aTechnik) and (Radek < RowCount-1) do Inc(Radek);
    Cells[3, Radek] := string(FaktData.Value[dmDES_common.IndexByName(FaktData, 'Amount')]);
    Cells[4, Radek] := string(FaktData.Value[dmDES_common.IndexByName(FaktData, 'DisplayName')]);
    Cells[5, Radek] := string(FaktData.Value[dmDES_common.IndexByName(FaktData, 'ID')]);
// zaplacení faktury v hotovosti
// hlavièka PP
    PPObject:= fmMainFT.AbraOLE.CreateObject('@CashReceived');
    PPData:= fmMainFT.AbraOLE.CreateValues('@CashReceived');
    PPObject.PrefillValues(PPData);
    PPData.ValueByName('DocQueue_ID') := PP_Id;
    PPData.ValueByName('Period_ID') := Period_Id;
    PPData.ValueByName('DocDate$DATE') := Date;
    PPData.ValueByName('AccDate$DATE') := Date;
    PPData.ValueByName('CreatedBy_ID') := User_Id;
    PPData.ValueByName('Firm_ID') := Firm_Id;
    PPData.ValueByName('FirmOffice_ID') := FirmOffice_Id;
//    PPData.ValueByName('Address_ID') := MyAddress_Id;
    PPData.ValueByName('VATDocument') := False;
    PPData.ValueByName('Currency_ID'):= '0000CZK000';
    PPData.ValueByName('Description') := Format('Zaplacení %s', [Cells[4, Radek]]);
    PPData.ValueByName('PDocumentType'):= '03';
    PPData.ValueByName('PDocument_ID'):= Cells[5, Radek];
    RadkyPP := PPData.Value[dmDES_common.IndexByName(PPData, 'Rows')];
// prázdný øádek
    RadekPP := fmMainFT.AbraOLE.CreateValues('@CashReceivedRow');
    RadekPP.ValueByName('Division_ID') := '1000000101';
    RadekPP.ValueByName('RowType') := '0';
    RadkyPP.Add(RadekPP);
// øádek s èástkou
    RadekPP:= fmMainFT.AbraOLE.CreateValues('@CashReceivedRow');
    RadekPP.ValueByName('Division_ID') := '1000000101';
//    RadekPP.ValueByName('Text') := Format('Zaplacení %s', [Cells[4, Radek]]);
//    RadekPP.ValueByName('IncomeType_ID') := '2000000000';
    RadekPP.ValueByName('RowType') := '1';
    RadekPP.ValueByName('TotalPrice') := Cells[3, Radek];
    RadkyPP.Add(RadekPP);
// vytvoøení PP
    if RadkyPP.Count > 0 then try
      ID := PPObject.CreateNewFromValues(PPData);
      PPData := PPObject.GetValues(ID);
      Cells[5, Radek] := string(PPData.Value[dmDES_common.IndexByName(PPData, 'DisplayName')]);    // pøepíše ID faktury
// vymazání zpracovaných plateb
    while (Cells[1, Radek+1] = aTechnik) and (Radek < RowCount-1) do begin
      Cells[3, Radek+1] := '';
      Inc(Radek);
    end;
    except
      on E: Exception do begin
        ShowMessage('Chyba pøi vytváøení dokladu.' + ^M + E.Message);
        Exit;
      end;
    end;  // try
  end;  // with fmMainFT ...

end;

// ------------------------------------------------------------------------------------------------

procedure TdmCode.ConnectAbra;
// pøipojení k Abøe, pokud je potøeba
begin
{  Screen.Cursor := crHourGlass;
  with fmMain do try
    if VarIsEmpty(AbraOLE) then begin
      Screen.Cursor := crHourGlass;
      dmCode.Zprava('Pøipojení k Abøe ...');
      try
        AbraOLE := CreateOLEObject('AbraOLE.Application');
        if not AbraOLE.Connect('@' + AbraConnection) then begin
          dmCode.Zprava('Problém s Abrou (connect  "' + AbraConnection + '").');
          Screen.Cursor := crDefault;
          Exit;
        end;
        if not AbraOLE.Login(fmLogin.acbJmeno.Text, fmLogin.aedHeslo.Text) then begin
          dmCode.Zprava('Problém s Abrou (jméno, heslo).');
          Screen.Cursor := crDefault;
          Exit;
        end;
        dmCode.Zprava('OK');
      except on E: exception do
        begin
          Application.MessageBox(PChar('Problém s Abrou.' + ^M + E.Message), 'Abra', MB_ICONERROR + MB_OK);
          dmCode.Zprava('Problém s Abrou.' + #13#10 + E.Message);
          Screen.Cursor := crDefault;
          Exit;
        end;
      end;
      Screen.Cursor := crDefault;
    end;  // if VarIsEmpty(AbraOLE)
  finally
    Screen.Cursor := crDefault;
  end;  // with fmMain
}end;

// ------------------------------------------------------------------------------------------------

procedure TdmCode.UpdateTechniciXLS(aRadek: integer);
// peníze z PP do technici.xls
// kód v Excelu dìlá problémy, proto Open Office
{ OpenOffice start at cell (0,0) while Excel at (1,1)        }
{ Also, Excel uses (row, col) and OpenOffice uses (col, row) }
const
  ooHAlignStd = 0; //    com.sun.star.table.CellHoriJustify.STANDARD
  ooHAlignLeft = 1; //   com.sun.star.table.CellHoriJustify.LEFT
  ooHAlignCenter = 2; // com.sun.star.table.CellHoriJustify.CENTER
  ooHAlignRight = 3; //  com.sun.star.table.CellHoriJustify.RIGHT
var
  ServiceManager,
  Desktop,
  LoadParams,
  Dokument,
  Sheet: variant;
  Hotovost: double;
  XLSName: string;
  Radek: integer;

  function SetParams (Name: string; Data: variant): variant;
// zadání parametrù pøi otevírání souboru (?)
  var
    Reflection: variant;
  begin
    Reflection := ServiceManager.CreateInstance('com.sun.star.reflection.CoreReflection');
    Reflection.forName('com.sun.star.beans.PropertyValue').createObject(result);
    result.Name := Name;
    result.Value := Data;
  end;

begin
{  Screen.Cursor := crHourGlass;
  with fmMain, asgMain do try
// hotovost z asgItems
    Hotovost := 0;
    for Radek := 1 to asgItems.RowCount-1 do
      if asgItems.Ints[4, Radek] = 1 then Hotovost := Hotovost + asgItems.Floats[2, Radek];
// v asgMain musí existovat PP a èástka musí být záporná
    if (Pos('PP', Cells[5, aRadek] )= 0) or (Hotovost >= 0) then Exit;
// OO ještì nebyl spuštìn
    if VarIsEmpty(ServiceManager) then try
// zkusí se, jestli nebìží
      ServiceManager := GetActiveOleObject('com.sun.star.ServiceManager');
    except try
// jinak se spustí
      ServiceManager := CreateOleObject('com.sun.star.ServiceManager');
      except
        on E: Exception do begin
          cbClear.Checked := False;
          Zprava('Nelze spustit OpenOffice' + ^M + E.Message);
          Exit;
        end;
      end;  // try
    end;  // try
    if (VarIsEmpty(ServiceManager) or VarIsNull(ServiceManager)) then begin
      cbClear.Checked := False;
      Zprava('Nelze spustit OpenOffice.');
      Exit;
    end;
// povedlo se
    if VarIsEmpty(Desktop) then Desktop := ServiceManager.CreateInstance('com.sun.star.frame.Desktop');
// soubor technikù
    XLSName := FormatDateTime('"W:/"yy"/technici"yy".xls"', Date);
// (popis parametrù v OO MediaDescriptor)
    LoadParams := VarArrayCreate([0, 0], varVariant);
// listy nebudou vidìt
    LoadParams[0] := SetParams('Hidden', True);
// soubor se otevøe
    try
      Dokument := Desktop.LoadComponentFromURL('file:///' + XLSName, '_default', 0, LoadParams);
    except
      on E: Exception do begin
        cbClear.Checked := False;
        Zprava('Nepodaøilo se otevøít ' + XLSName + ^M + E.Message);
        if not (VarIsEmpty(ServiceManager) or VarIsNull(ServiceManager)) then ServiceManager := Unassigned;
        Exit;
      end;
    end;
    if (VarIsEmpty(Dokument) or VarIsNull(Dokument)) then begin
      cbClear.Checked := False;
      Zprava('Nepodaøilo se otevøít ' + XLSName);
      if not (VarIsEmpty(ServiceManager) or VarIsNull(ServiceManager)) then ServiceManager := Unassigned;
      Exit;
    end;
// aktivní list (podle technika)
    try
      Sheet := Dokument.getSheets.getByName(Cells[7, aRadek]);
    except
      on E: Exception do begin
        cbClear.Checked := False;
        Zprava('Nepodaøilo se otevøít list ' + Cells[7, aRadek] + ^M + E.Message);
        if not (VarIsEmpty(Desktop) or VarIsNull(Desktop)) then Desktop := Unassigned;
        if not (VarIsEmpty(ServiceManager) or VarIsNull(ServiceManager)) then ServiceManager := Unassigned;
        Exit;
      end;
    end;
    if (VarIsEmpty(Sheet) or VarIsNull(Sheet)) then begin
      cbClear.Checked := False;
      Zprava('Nepodaøilo se otevøít list ' + Cells[7, aRadek]);
      if not (VarIsEmpty(Desktop) or VarIsNull(Desktop)) then Desktop := Unassigned;
      if not (VarIsEmpty(ServiceManager) or VarIsNull(ServiceManager)) then ServiceManager := Unassigned;
      Exit;
    end;
    Dokument.getCurrentController.setActiveSheet(Sheet);
// bude se hledat první prázdný øádek
    Radek := 3;
    while (Sheet.getCellByPosition(2, Radek).getFormula <> '') do Inc(Radek);      // sloupec s èástkami
// vloží se hodnoty
    Sheet.getCellByPosition(0, Radek).String := FormatDateTime('d.m.', Date);
    Sheet.getCellByPosition(0, Radek).HoriJustify := ooHAlignLeft;
    Sheet.getCellByPosition(1, Radek).Value := Hotovost;
    Sheet.getCellByPosition(1, Radek).HoriJustify := ooHAlignRight;;
    Sheet.getCellByPosition(2, Radek).String := Trim(Cells[5, aRadek]) +  ' ' + Cells[2, aRadek];  // PP + VS
    Sheet.getCellByPosition(2, Radek).HoriJustify := ooHAlignLeft;
    try
      Dokument.Store;
      Zprava(Format('Èástka %f pøidána na list %s.', [Hotovost, Cells[7, aRadek]]));
    except
      Application.MessageBox(PChar('Nedá se uložit ' +  XLSName + '. Není nìkde otevøený ?' ), 'Open Office Calc', MB_ICONERROR + MB_OK);
      Zprava('Nedá se uložit ' +  XLSName);
      cbClear.Checked := False;
    end;
    Dokument.Dispose;
  finally
// Clean up
    Dokument := Null;
    Sheet := Null;
    if not (VarIsEmpty(Desktop) or VarIsNull(Desktop)) then Desktop := Unassigned;
    if not (VarIsEmpty(ServiceManager) or VarIsNull(ServiceManager)) then ServiceManager := Unassigned;
    Screen.Cursor := crDefault;
  end;  // with fmMain ...
}end;

// ------------------------------------------------------------------------------------------------

end.

