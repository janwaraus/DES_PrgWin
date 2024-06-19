// 1.6.2024 automatizace placen� faktur technik�. N�kolika technik�m se vystavuj� faktury nej�ast�ji za televizi, kter� tito
// neplat�, a faktury se hrad� hotov� s vystaven�m p��jmov�ho dokladu a n�sledn�m ode�ten�m v souboru W:\techniciYY.xls.
// Program cel� postup automatizuje.

unit FO;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, System.DateUtils, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.StdCtrls, AdvUtil, Vcl.Grids, AdvObj, BaseGrid, AdvGrid,
  IniFiles, ComObj, Math, ZAbstractConnection, ZConnection, Data.DB, ZAbstractRODataset, ZAbstractDataset, ZDataset;

type
  TfmMain = class(TForm)
    pnTop: TPanel;
    btNacist: TButton;
    btZaplatit: TButton;
    asgMain: TAdvStringGrid;
    qrAbra: TZQuery;
    cnAbra: TZConnection;
    cbxRok: TComboBox;
    procedure FormShow(Sender: TObject);
    procedure btNacistClick(Sender: TObject);
    procedure btZaplatitClick(Sender: TObject);
  private
    TechniciXLS,
    TechDir: string;
    AbraOLE,
    Sheets,
    DocTech: variant;
  end;

const
  Ap = chr(39);
  ApC = Ap + ',';
  ApZ = Ap + ')';

var
  fmMain: TfmMain;

implementation

uses DES_common, DES_OO_common;

{$R *.dfm}

// ------------------------------------------------------------------------------------------------

procedure TfmMain.FormShow(Sender: TObject);

var
  DESIni: TIniFile;
  ExePath: string;

begin
  ExePath := ExtractFilePath(ParamStr(0));
// parametry z AbraDESProgramy.ini
  if not FileExists((ExePath + 'AbraDESProgramy.ini'))
    and not(FileExists(ExePath + '..\DE$_Common\AbraDESProgramy.ini')) then
  begin
    Application.MessageBox(PChar('Nenalezen soubor AbraDESProgramy.ini, program ukon�en'), 'AbraDESProgramy.ini', MB_OK + MB_ICONERROR);
    Application.Terminate;
  end;
// DESIni
  if FileExists(ExePath + 'AbraDESProgramy.ini') then
    DESIni := TIniFile.Create(ExePath + 'AbraDESProgramy.ini')
  else
    DESIni := TIniFile.Create(ExePath + '..\DE$_Common\AbraDESProgramy.ini');

  with DESIni do try
    cnAbra.Database := ReadString('Preferences', 'AbraDB', '');
    cnAbra.User := ReadString('Preferences', 'AbraUN', '');
    cnAbra.Password := ReadString('Preferences', 'AbraPW', '');
    TechDir := ReadString('Preferences', 'Dirs', 'W:\');
  finally
    DESIni.Free;
  end;

  try
    cnAbra.Connect;
  except on E: exception do
    begin
      Application.MessageBox(PChar('Ned� se p�ipojit k datab�zi Abry, program ukon�en.' + ^M + E.Message), 'Abra', MB_ICONERROR + MB_OK);
      Application.Terminate;
    end;
  end;

// aktu�ln� a minul� rok do cbxRok
  cbxRok.Items.Add(IntToStr(YearOf(Date)));
  cbxRok.Items.Add(IntToStr(YearOf(Date)-1));
  cbxRok.ItemIndex := 1;

// fajfky
  asgMain.CheckFalse := '0';
  asgMain.CheckTrue := '1';

  btZaplatit.Enabled := False;

  end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.btNacistClick(Sender: TObject);
// na�te seznam technik� z techniciYY.xls a z Abry seznam jejich nezaplacen�ch faktur
var
  FileHandle : integer;
  Sheet: variant;
  iRadek,
  i: integer;
  SQLStr: string;

begin
  Screen.Cursor := crHourGlass;
// existuje techniciYY.xls ?
  TechniciXLS := Format(TechDir + '%s\technici%s.xls', [Copy(cbxRok.Text, 3, 2), Copy(cbxRok.Text, 3, 2)]);
  if FileExists(TechniciXLS) then begin
// d� se otev��t ?
    FileHandle := System.SysUtils.FileOpen(TechniciXLS, fmOpenReadWrite);
    if FileHandle > 0 then begin
      btZaplatit.Visible := True;
      FileClose(FileHandle);
    end else begin
      Application.MessageBox(PChar('Soubor ' + TechniciXLS + ' je asi otev�en�.'), PChar(TechniciXLS), MB_OK + MB_ICONWARNING);
      Screen.Cursor := crDefault;
      Exit;
    end;
  end else begin
    Application.MessageBox(PChar('Soubor ' + TechniciXLS + ' neexistuje.'), PChar(TechniciXLS), MB_OK + MB_ICONWARNING);
    Screen.Cursor := crDefault;
    Exit;
  end;

// otev�en� Open Office
  with dmDES_OO_Common do try
    ServiceManager := OpenOO;
//    if IsNullEmpty(ServiceManager) then Exit;
// povedlo se, desktop
    if IsNullEmpty(Desktop) then Desktop := OpenDesktop;
    if IsNullEmpty(Desktop) then begin
      Exit;
    end;
// parametry otev�en� - nic nebude vid�t?
    LoadParams := VarArrayCreate([0, 0], varVariant);
    LoadParams[0] := SetParams('Hidden', True);
//    LoadParams[0] := SetParams('Hidden', False);
// otev�e se soubor
    TechDir := StringReplace(ExcludeTrailingPathDelimiter(TechDir), '\', '/', [rfReplaceAll]);
    TechniciXLS := Format(TechDir + '/%s/technici%s.xls', [Copy(cbxRok.Text, 3, 2), Copy(cbxRok.Text, 3, 2)]);
    DocTech := OpenDoc(TechniciXLS);
    if IsNullEmpty(DocTech) then begin
      Exit;
    end;
// v�echny listy
    Sheets := DocTech.getSheets;
    if IsNullEmpty(Sheets) then begin
      Exit;
    end;

// pln�n� asgMain
    iRadek := 0;
    asgMain.ClearNormalCells;
    for i := 1 to Sheets.GetCount-1 do begin
      Sheet := Sheets.GetByIndex(i);

// nezaplacen� faktury technika
      SQLStr := 'SELECT II.ID, DQ.Code ||''-''|| II.OrdNumber ||''/''|| P.Code AS Doklad,'
      + ' II.LocalAmount - II.LocalCreditAmount - (II.LocalPaidAmount - II.LocalPaidCreditAmount) AS Castka'
      + ' FROM IssuedInvoices II'
      + ' INNER JOIN DocQueues DQ ON II.DocQueue_ID = DQ.Id'               // �ada doklad�
      + ' INNER JOIN Periods P ON II.Period_ID = P.Id'                     // rok
      + ' WHERE II.Firm_ID = '
      +   '(SELECT ID FROM Firms WHERE Code = ' + Ap + Sheet.GetName + Ap
        + ' AND Firm_ID IS NULL'
        + ' AND Hidden = ''N'')'
      + ' AND II.LocalAmount - II.LocalCreditAmount - (II.LocalPaidAmount - II.LocalPaidCreditAmount) > 0'
      + ' ORDER BY II.DocDate$DATE';

      with qrAbra, asgMain do begin
        Close;
        SQL.Text := SQLStr;
        Open;
        if RecordCount > 0 then
          while not EOF do begin
            Inc(iRadek);
            RowCount := iRadek + 1;
            AddCheckBox(0, iRadek, True, True);
//            Ints[0, Radek] := 1;
            Cells[1, iRadek] := Sheet.GetName;
            Cells[2, iRadek] := FieldByName('Doklad').AsString;
            Floats[3, iRadek] := FieldByName('Castka').AsFloat;
//            Cells[2, iRadek] := Format('%s-%s/%s', [FieldByName('Rada').AsString, FieldByName('Cislo').AsString, FieldByName('Rok').AsString]);
//            Floats[3, iRadek] := FieldByName('Vystaveno').AsFloat - FieldByName('Zaplaceno').AsFloat;// - FieldByName('Dobropis').AsFloat;
            Cells[4, iRadek] := FieldByName('ID').AsString;
            Next;
          end;
        if iRadek > 0 then  btZaplatit.Visible := True;
      end;
    end;

  except on E: exception do
    begin
      Screen.Cursor := crDefault;
      DocTech.Dispose;
      Desktop.Terminate;
      Desktop := Unassigned;
      ServiceManager := Unassigned
    end;
  end;

  Screen.Cursor := crDefault;
  btZaplatit.Enabled := True;

end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.btZaplatitClick(Sender: TObject);
// podle dat ve fmMain zaplat� hotov� faktury
var
  iRadek,
  iRadekXLS: integer;
  Period_Id,
  Firm_Id,
  FirmOffice_Id,
  ID: string[10];
  SQLStr: string;
  PPObject,
  PPData,
  RadkyPP,
  RadekPP: variant;
  DatumDokladu: double;
  Sheets,
  Sheet: variant;

const
  ooHAlignStd = 0; //    com.sun.star.table.CellHoriJustify.STANDARD
  ooHAlignLeft = 1; //   com.sun.star.table.CellHoriJustify.LEFT
  ooHAlignCenter = 2; // com.sun.star.table.CellHoriJustify.CENTER
  ooHAlignRight = 3; //  com.sun.star.table.CellHoriJustify.RIGHT

begin
  Screen.Cursor := crHourGlass;
  btNacist.Enabled := False;
  btZaplatit.Enabled := False;
  Application.ProcessMessages;
// p�ipojen� Abry (pokud je t�eba)
  if VarIsEmpty(AbraOLE) then try
    AbraOLE := CreateOLEObject('AbraOLE.Application');
    if not AbraOLE.Connect('@DES') then begin
      ShowMessage('Probl�m s Abrou (connect DES).');
      Screen.Cursor := crDefault;
      Exit;
    end;
    if not AbraOLE.Login('Supervisor', '') then begin
      ShowMessage('Probl�m s Abrou (Login Supervisor).');
      Screen.Cursor := crDefault;
      Exit;
    end;
  except on E: exception do
    begin
      Application.MessageBox(PChar('Probl�m s Abrou.' + #13#10  + E.Message), 'Abra', MB_ICONERROR + MB_OK);
      Screen.Cursor := crDefault;
      Exit;
    end;
  end;  // if VarIsEmpty(AbraOLE)

// pro v�echny vybran� ��dky v asgMain se katura zaplat�
  for iRadek := 1 to asgMain.RowCount-1 do try
    if (asgMain.Ints[0, iRadek] = 1) then begin

// data technika pro Abru
      with qrAbra do begin
        Close;
        SQLStr := 'SELECT F.Id AS FId, FO.Id AS FOId FROM Firms F, FirmOffices FO'
        + ' WHERE F.Code = ' + Ap + asgMain.Cells[1, iRadek] + Ap
        + ' AND F.Firm_ID IS NULL'
        + ' AND F.Hidden = ''N'''
        + ' AND FO.Parent_Id = F.Id';
        SQL.Text := SQLStr;
        Open;
        if RecordCount = 0 then begin
          ShowMessage(Format('Technik %s nen� v adres��i Abry.', [asgMain.Cells[1, iRadek]]));
          Close;
          Exit;
        end;
        Firm_Id := FieldByName('FId').AsString;
        FirmOffice_Id := FieldByName('FOId').AsString;
        Close;
        SQL.Text := 'SELECT Id FROM Periods WHERE Code = ' + Ap + cbxRok.Text + Ap;
        Open;
        Period_Id := FieldByName('Id').AsString;
        Close;
      end;  // with qrAbra

// zaplacen� faktury hotov� (vytvo�en� PP)
      DatumDokladu := Floor(Date);
// hlavi�ka PP
      PPObject:= AbraOLE.CreateObject('@CashReceived');
      PPData:= AbraOLE.CreateValues('@CashReceived');
      PPObject.PrefillValues(PPData);
      PPData.ValueByName('DocQueue_ID') := '4000000101';               // PP
      PPData.ValueByName('Period_ID') := Period_Id;
      PPData.ValueByName('DocDate$DATE') := DatumDokladu;
      PPData.ValueByName('AccDate$DATE') := DatumDokladu;
      PPData.ValueByName('CreatedBy_ID') := '2200000101';              // automatick� fakturace
      PPData.ValueByName('Firm_ID') := Firm_Id;
      PPData.ValueByName('FirmOffice_ID') := FirmOffice_Id;
      PPData.ValueByName('VATDocument') := False;
      PPData.ValueByName('Currency_ID'):= '0000CZK000';
      PPData.ValueByName('Description') := Format('Zaplacen� %s', [asgMain.Cells[2, iRadek]]);
      PPData.ValueByName('PDocumentType'):= '03';
      PPData.ValueByName('PDocument_ID'):= asgMain.Cells[4, iRadek];
      RadkyPP := PPData.Value[dmDES_common.IndexByName(PPData, 'Rows')];
// pr�zdn� ��dek
      RadekPP := AbraOLE.CreateValues('@CashReceivedRow');
      RadekPP.ValueByName('Division_ID') := '1000000101';
      RadekPP.ValueByName('RowType') := '0';
      RadkyPP.Add(RadekPP);
// ��dek s ��stkou
      RadekPP:= AbraOLE.CreateValues('@CashReceivedRow');
      RadekPP.ValueByName('Division_ID') := '1000000101';
      RadekPP.ValueByName('RowType') := '1';
      RadekPP.ValueByName('TotalPrice') := asgMain.Cells[3, iRadek];
      RadkyPP.Add(RadekPP);
// vytvo�en� PP
      if RadkyPP.Count > 0 then try
        ID := PPObject.CreateNewFromValues(PPData);
        PPData := PPObject.GetValues(ID);
        asgMain.Cells[4, iRadek] := string(PPData.Value[dmDES_common.IndexByName(PPData, 'DisplayName')]);    // p�ep�e ID faktury

      except
        on E: Exception do begin
          ShowMessage('Chyba p�i vytv��en� dokladu.' + #13#10 + E.Message);
          Exit;
        end;
      end;  // try

// ulo�en� do techniciYY.xls (je otev�en� v OO)
      Sheets := DocTech.getSheets;
      Sheet := Sheets.getByName(asgMain.Cells[1, iRadek]);
// bude se hledat prvn� pr�zdn� ��dek
      iRadekXLS := 3;
      while (Sheet.getCellByPosition(2, iRadekXLS).getFormula <> '') do Inc(iRadekXLS);      // sloupec s ��stkami
// vlo�� se hodnoty
      Sheet.getCellByPosition(0, iRadekXLS).String := FormatDateTime('d.m.', (Date));
      Sheet.getCellByPosition(1, iRadekXLS).String := '-' + asgMain.Cells[3, iRadek];
      Sheet.getCellByPosition(1, iRadekXLS).HoriJustify := ooHAlignRight;;
      Sheet.getCellByPosition(2, iRadekXLS).String := asgMain.Cells[4, iRadek] + ' (televize)';
      Sheet.getCellByPosition(2, iRadekXLS).HoriJustify := ooHAlignLeft;

    end;   //  if (asgMain.Ints[0, iRadek] = 1)

  except on E: Exception  do
    begin
      Screen.Cursor := crDefault;
      DocTech.Dispose;
      dmDES_OO_Common.Desktop.Terminate;
      dmDES_OO_Common.Desktop := Unassigned;
      dmDES_OO_Common.ServiceManager := Unassigned;
    end;

  end;  // for ... try

  btNacist.Enabled := True;
  Screen.Cursor := crDefault;
  DocTech.Store;
  DocTech.Dispose;
  dmDES_OO_Common.Desktop.Terminate;
  dmDES_OO_Common.Desktop := Unassigned;
  dmDES_OO_Common.ServiceManager := Unassigned;

  ShowMessage('OK');

end;

// ------------------------------------------------------------------------------------------------

end.
