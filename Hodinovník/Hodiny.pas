// 25.8.2013 Pro jednotlivé techniky vypisuje mìsíèní souèet hodin a schválených hodin, uložených
// v databázi BEZDRAT.
// 16.4.2017 hodiny z MySQL pomocí Schejbalova API
// 10.8. nové API, upraveno
// 12.12.18 vložena hodinová sazba technika, posun doprava
// 13.12. ze simple_billings pøidány tržby od zákazníkù a stav skladu technika
// 3.2.2019 souèty hodin technikù z tabulky "reported_hours" - není
// 25.4. https místo http
// 9.5. pøi pomalém natahování SSL knuhovny druhý pokus
// 16.5. import hodinových sazeb z tabulky "income_rates"
// 24.8. import odmìn (bonus) do pøidaného sloupce v .xls
// 18.10. úprava výbìru technikù do cbxJmeno
// 20.10. import pohybù hotovosti technikù ze souboru techniciyy.xls do tabulky wallet_transactions
// 4.2021 návrat ke koøenùm - API je pøíliš pomalé
// 19.4.2021 Pro jednotlivé techniky vypisuje mìsíèní souèty schválených a ostatních hodin, uložených v aplikaci v tabulkách
// "technicians_tasks" a "reported_hours".
// 6.12. opraven výpoèet neschválených hodin a kilometrù
// 11.1.2022 opraveny tržby od zákazníkù
// 5.1.2024 podle toho, co je vybráno v cbxJmeno, se pøevádí do nového roku všechny zùstatky, nebo jen vybraný technik
// 8.1. kontrola, zda nejsou otevøené soubory techniciYY.xls


unit Hodiny;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms, DateUtils, ComObj, IniFiles,
  Dialogs, Grids, BaseGrid, AdvGrid, DB, ZAbstractRODataset, ZAbstractDataset, ZDataset,
  ZConnection, StdCtrls, ExtCtrls, ZAbstractConnection, AdvObj, AdvCombo, AdvUtil;

type
  TfmMain = class(TForm)
    cnMain: TZConnection;
    qrMain: TZReadOnlyQuery;
    cnAbra: TZConnection;
    qrAbra: TZQuery;
    pnTop: TPanel;
    cbxRok: TComboBox;
    cbxJmeno: TComboBox;
    asgMain: TAdvStringGrid;
    btDoXLS: TButton;
    btPrevod: TButton;
    asgKasa: TAdvStringGrid;
    procedure FormShow(Sender: TObject);
    procedure cnMainAfterConnect(Sender: TObject);
    procedure cbxRokChange(Sender: TObject);
    procedure cbxJmenoChange(Sender: TObject);
    procedure asgMainGetAlignment(Sender: TObject; ARow, ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asgMainGetCellColor(Sender: TObject; ARow, ACol: Integer; AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
    procedure btDoXLSClick(Sender: TObject);
    procedure btPrevodClick(Sender: TObject);
  private
    TechDir: string;
    Document: variant;
  end;

const
  Ap = chr(39);
  ApC = Ap + ',';
  ApZ = Ap + ')';

var
  fmMain: TfmMain;

implementation

{$R *.dfm}

uses DES_common, DES_OO_Common;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.FormShow(Sender: TObject);
var
  DESIni: TIniFile;
  ExePath: string;
begin
  ExePath := ExtractFilePath(ParamStr(0));
// AbraDESProgramy.ini ?
  if not(FileExists(ExePath + 'AbraDESProgramy.ini'))
    and not(FileExists(ExePath + '..\DE$_Common\AbraDESProgramy.ini')) then
  begin
    Application.MessageBox(PChar('Nenalezen soubor AbraDESProgramy.ini, program ukonèen'),
      'AbraDESProgramy.ini', MB_OK + MB_ICONERROR);
    Application.Terminate;
  end;
// DESIni
  if FileExists(ExePath + 'AbraDESProgramy.ini') then
    DESIni := TIniFile.Create(ExePath + 'AbraDESProgramy.ini')
  else
    DESIni := TIniFile.Create(ExePath + '..\DE$_Common\AbraDESProgramy.ini');
// parametry z AbraDESProgramy.ini
  with DESIni do try
    cnMain.HostName := ReadString('Preferences', 'ZakHN', '');
    cnMain.Database := ReadString('Preferences', 'ZakDB', '');
    cnMain.User := ReadString('Preferences', 'ZakUN', '');
    cnMain.Password := ReadString('Preferences', 'ZakPW', '');
    cnAbra.Database := ReadString('Preferences', 'AbraDB', '');
    cnAbra.User := ReadString('Preferences', 'AbraUN', '');
    cnAbra.Password := ReadString('Preferences', 'AbraPW', '');
    TechDir := ReadString('Preferences', 'Dirs', 'W:\');
  finally
    DESIni.Free;
  end;
  try
    cnMain.Connect;
  except on E: exception do
    begin
      Application.MessageBox(PChar('Nedá se pøipojit k databázi aplikace, program ukonèen.' + ^M + E.Message), 'mySQL', MB_ICONERROR + MB_OK);
      Application.Terminate;
    end;
  end;
  try
    cnAbra.Connect;
  except on E: exception do
    begin
      Application.MessageBox(PChar('Nedá se pøipojit k databázi Abry, program ukonèen.' + ^M + E.Message), 'Abra', MB_ICONERROR + MB_OK);
      Application.Terminate;
    end;
  end;
  asgMain.MergeCells(1, 0, 2, 1);
  asgMain.MergeCells(3, 0, 2, 1);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.cnMainAfterConnect(Sender: TObject);
// naplnìní cbxRok a cbxJmeno
begin
  with qrMain do begin
    SQL.Text := 'SELECT DISTINCT EXTRACT(year FROM start_at) AS Rok FROM technicians_tasks'
    + ' WHERE EXTRACT(year FROM start_at) < 2100';
    Open;
    while not EOF do begin
      cbxRok.Items.Add(FieldByName('Rok').AsString);
      Next;
    end;
    Close;
    cbxRok.Text := FormatDateTime('yyyy', Date);
// 18.1.2024
    cbxRokChange(nil);
{
    SQL.Text := 'SELECT DISTINCT TT.user_id, CONCAT(U.issue_shortcut, '' / '', U.first_name, '' '', U.last_name) AS Jmeno'
    + ' FROM technicians_tasks TT, users U'
    + ' WHERE TT.user_id = U.Id'
    + ' AND EXTRACT(year FROM TT.start_at) = ' + cbxRok.Text             // jen ti, co nìco odpracovali
    + ' AND U.issue_shortcut <> ''AT'''
    + ' AND U.issue_shortcut <> ''Adm''';
    Open;
    while not EOF do begin
      if (FieldByName('Jmeno').AsString <> '') then cbxJmeno.Items.Add(FieldByName('Jmeno').AsString);
      Next;
    end;
    Close;
    cbxJmeno.ItemIndex := -1;
}
    btDoXLS.Caption := Format('%s do XLS', [cbxRok.Text]);
    btPrevod.Caption := Format('%d do %s', [StrToInt(cbxRok.Text)-1, cbxRok.Text]);
  end;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.cbxRokChange(Sender: TObject);
var
  TechFile: string;
  FileHandle : integer;
begin
  cbxJmeno.Clear;
  cbxJmeno.Text := 'Jméno';
  with qrMain do begin
    SQL.Text := 'SELECT DISTINCT TT.user_id, CONCAT(U.issue_shortcut, '' / '', U.first_name, '' '', U.last_name) AS Jmeno'
    + ' FROM technicians_tasks TT, users U'
    + ' WHERE  TT.user_id = U.Id'
    + ' AND EXTRACT(year FROM TT.start_at) = ' + cbxRok.Text             // jen ti, co nìco odpracovali
    + ' AND U.issue_shortcut <> ''AT'''
    + ' AND U.issue_shortcut <> ''Adm''';
    Open;
    while not EOF do begin
      if (FieldByName('Jmeno').AsString <> '') then cbxJmeno.Items.Add(FieldByName('Jmeno').AsString);
      Next;
    end;
    Close;
  end;
  cbxJmeno.ItemIndex := -1;
  asgMain.ClearNormalCells;

// 8.1.
  if cbxRok.Text = '' then Exit;
// automatický pøevod do technici.xls jen pro poslední dva roky
  if (StrToInt(cbxRok.Text) >= YearOf(Date)-1) then  begin
    btDoXLS.Visible := False;
// existuje techniciYY.xls ?
    TechFile := Format(TechDir + '%s\technici%s.xls', [Copy(cbxRok.Text, 3, 2), Copy(cbxRok.Text, 3, 2)]);
    if FileExists(TechFile) then begin
// dá se otevøít ?
      FileHandle := SysUtils.FileOpen(TechFile, fmOpenReadWrite);
      if FileHandle > 0 then begin
        btDoXLS.Visible := True;
        FileClose(FileHandle);
      end else Application.MessageBox(PChar('Soubor ' + TechFile + ' je asi otevøený.'), PChar(TechFile), MB_OK + MB_ICONWARNING);
    end else Application.MessageBox(PChar('Soubor ' + TechFile + ' neexistuje.'), PChar(TechFile), MB_OK + MB_ICONWARNING);

    btPrevod.Visible := False;
// existuje techniciYY-1.xls ?
    TechFile := Format(TechDir + '%s\technici%s.xls', [FormatDateTime('yy', IncYear(Date, -1)), FormatDateTime('yy', IncYear(Date, -1))]);
    if FileExists(TechFile) then begin
// dá se otevøít ?
      FileHandle := SysUtils.FileOpen(TechFile, fmOpenReadWrite);
      if FileHandle > 0 then begin
        btPrevod.Visible := True;
        FileClose(FileHandle);
      end else Application.MessageBox(PChar('Soubor ' + TechFile + ' je asi otevøený.'), PChar(TechFile), MB_OK + MB_ICONWARNING);
    end else Application.MessageBox(PChar('Soubor ' + TechFile + ' neexistuje.'), PChar(TechFile), MB_OK + MB_ICONWARNING);
  end else Exit;

  btDoXLS.Caption := Format('%s do XLS', [cbxRok.Text]);
  btPrevod.Caption := Format('%d do %s', [StrToInt(cbxRok.Text)-1, cbxRok.Text]);

end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.cbxJmenoChange(Sender: TObject);
var
  sTech,
  sTechId,
  SQLStr: string;
  sYM: string[6];
  Mesic,
  Sloupec: integer;
begin
  sTech := Copy(cbxJmeno.Text, 1, Pos(' /', cbxJmeno.Text)-1);
  with qrMain, asgMain do begin
    SQLStr := 'SELECT Id FROM users'
    + ' WHERE issue_shortcut = ''' + Copy(cbxJmeno.Text, 1, Pos(' /', cbxJmeno.Text)) + '''';
    SQL.Text := SQLStr;
    Open;
    sTechId := Fields[0].AsString;
    Close;
    ClearNormalCells;
    SQLStr := 'SELECT SUM(approved_working_hours + approved_on_road_hours + approved_with_driver_hours) as Hodiny,'
    + ' SUM(approved_on_road_distance) as KM, SUM(approved_bonus) as Bonus, EXTRACT(month FROM start_at) AS Mesic'
    + ' FROM eurosignal_production.reported_hours RH, eurosignal_production.users U, eurosignal_production.technicians_tasks TT'
    + ' WHERE TT.user_id = U.id'
    + ' AND TT.id = RH.technicians_task_id'
    + ' AND U.issue_shortcut = ''' + sTech + ''''
    + ' AND TT.hours_state = ''approved'''
    + ' AND EXTRACT(year FROM start_at) = ' + cbxRok.Text
    + ' GROUP BY Mesic'
    + ' ORDER BY Mesic';
    SQL.Text := SQLStr;
    Open;
    while not EOF do begin
      Floats[1, FieldByName('Mesic').AsInteger] := FieldByName('Hodiny').AsFloat;
      Floats[2, FieldByName('Mesic').AsInteger] := FieldByName('Hodiny').AsFloat;
// 6.12.21      Floats[3, FieldByName('Mesic').AsInteger] := FieldByName('KM').AsFloat;
      Floats[4, FieldByName('Mesic').AsInteger] := FieldByName('KM').AsFloat;    // schválené km
      Floats[5, FieldByName('Mesic').AsInteger] := FieldByName('Bonus').AsFloat;
      Next;
    end;
    Close;
{ 8.12.21
    SQLStr := 'SELECT SUM(approved_working_hours + approved_on_road_hours + approved_with_driver_hours) as Hodiny,'
    + ' SUM(approved_on_road_distance) as KM, EXTRACT(month FROM start_at) AS Mesic'
    + ' FROM eurosignal_production.reported_hours RH, eurosignal_production.users U, eurosignal_production.technicians_tasks TT'
    + ' WHERE TT.user_id = U.id'
    + ' AND TT.id = RH.technicians_task_id'
    + ' AND U.issue_shortcut = ''' + sTech + ''''
    + ' AND TT.hours_state <> ''approved'''
    + ' AND EXTRACT(year FROM start_at) = ' + cbxRok.Text
    + ' GROUP BY Mesic'
    + ' ORDER BY Mesic';
    }
// 8.12.21 neschválené hodiny
    SQLStr := 'SELECT SUM(working_hours + on_road_hours + with_driver_hours) as Hodiny,'
    + ' EXTRACT(month FROM start_at) AS Mesic'
    + ' FROM eurosignal_production.reported_hours RH, eurosignal_production.users U, eurosignal_production.technicians_tasks TT'
    + ' WHERE TT.user_id = U.id'
    + ' AND U.issue_shortcut = ''' + sTech + ''''
    + ' AND TT.id = RH.technicians_task_id'
    + ' AND TT.hours_state IN (''for_amendment'', ''waiting'')'
    + ' AND RH.state IN (''new'', ''for_amendment'')'
    + ' AND EXTRACT(year FROM start_at) = ' + cbxRok.Text
    + ' GROUP BY Mesic'
    + ' ORDER BY Mesic';
    SQL.Text := SQLStr;
    Open;
    while not EOF do begin
      if (FieldByName('Hodiny').AsString <> '') then
        Floats[1, FieldByName('Mesic').AsInteger] := FieldByName('Hodiny').AsFloat + Floats[2, FieldByName('Mesic').AsInteger];
//      if (FieldByName('KM').AsString <> '') then
//        Floats[3, FieldByName('Mesic').AsInteger] := FieldByName('KM').AsFloat + Floats[4, FieldByName('Mesic').AsInteger];
      Next;
    end;
    Close;

// 8.12.21 km celkem
    SQLStr := 'SELECT SUM(on_road_distance) as KM, EXTRACT(month FROM start_at) AS Mesic'
    + ' FROM eurosignal_production.reported_hours RH, eurosignal_production.users U, eurosignal_production.technicians_tasks TT'
    + ' WHERE TT.user_id = U.id'
    + ' AND TT.id = RH.technicians_task_id'
    + ' AND U.issue_shortcut = ''' + sTech + ''''
    + ' AND RH.state <> ''rejected'''
    + ' AND EXTRACT(year FROM start_at) = ' + cbxRok.Text
    + ' GROUP BY Mesic'
    + ' ORDER BY Mesic';
    SQL.Text := SQLStr;
    Open;
    while not EOF do begin
      Floats[3, FieldByName('Mesic').AsInteger] := FieldByName('KM').AsFloat;
      Next;
    end;
    Close;
    for Sloupec := 1 to 5 do Floats[Sloupec, 13] := ColumnSum(Sloupec, 1, 12);
// hodinové sazby
    for Mesic := 1 to 12 do begin
      sYM := Format('%s%2.2d', [cbxRok.Text, Mesic]);
      SQLStr := 'SELECT per_hour, per_km'
      + ' FROM income_rates'
      + ' WHERE user_id = ' + sTechId
      + ' AND EXTRACT(year_month FROM since) = (SELECT MAX(EXTRACT(year_month FROM IR1.since)) FROM income_rates IR1'
      + '  WHERE EXTRACT(year_month FROM IR1.since) <= ' + sYM
      + '  AND IR1.user_id = ' + sTechId + ')';
      Close;
      SQL.Text := SQLStr;
      Open;
      Cells[6, Mesic] := FieldByName('per_hour').AsString;
      Cells[7, Mesic] := FieldByName('per_km').AsString;
      Close;
      Application.ProcessMessages;
    end;  // for
  end;  // with
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.asgMainGetAlignment(Sender: TObject; ARow, ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
  if (ARow > 0) and (ACol in [1, 3]) then HAlign := taRightJustify
  else if (ARow > 0) and (ACol in [2, 4]) then HAlign := taLeftJustify
  else HAlign := taCenter;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.asgMainGetCellColor(Sender: TObject; ARow, ACol: Integer; AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
begin
  if (ARow = 0) or (ACol in [0, 1, 3, 5..7]) then Exit
  else with asgMain do
    if Cells[ACol, ARow] <> '' then
      if Cells[ACol, ARow] <> Cells[ACol-1, ARow] then AFont.Color := clRed;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.btDoXLSClick(Sender: TObject);
// hodiny a km do technici.xls
{ OpenOffice start at cell (0,0) and uses (col, row) }
const
  ooHAlignStd = 0; //    com.sun.star.table.CellHoriJustify.STANDARD
  ooHAlignLeft = 1; //   com.sun.star.table.CellHoriJustify.LEFT
  ooHAlignCenter = 2; // com.sun.star.table.CellHoriJustify.CENTER
  ooHAlignRight = 3; //  com.sun.star.table.CellHoriJustify.RIGHT
var
  Sheets,
  Sheet: variant;
  XLSName,
  TechFile,
  Technik,
  TechId,
  sDatum,
  SQLStr: string;
  sYM: string[6];
  ixJmeno,
  RadekasgKasa,
  Sloupec,
  Radek: integer;

begin
  Screen.Cursor := crHourGlass;
  asgKasa.ClearNormalCells;
  asgKasa.RowCount := 2;
  asgKasa.ColCount := 4;
  RadekasgKasa := 0;

  // 20.10.19 spuštìní Open Office a otevøení souboru s využitím dmDES_OO_Common
  with dmDES_OO_Common do begin
    ServiceManager := OpenOO;
    if IsNullEmpty(ServiceManager) then Exit;
// povedlo se, desktop
    if IsNullEmpty(Desktop) then Desktop := OpenDesktop;
    if IsNullEmpty(Desktop) then begin
      ServiceManager := Unassigned;
      Exit;
    end;
// parametry otevøení - nic nebude vidìt
    LoadParams := VarArrayCreate([0, 0], varVariant);
    LoadParams[0] := SetParams('Hidden', True);
// otevøe se soubor
    TechDir := StringReplace(ExcludeTrailingPathDelimiter(TechDir), '\', '/', [rfReplaceAll]);
    TechFile := Format(TechDir + '/%s/technici%s.xls', [Copy(cbxRok.Text, 3, 2), Copy(cbxRok.Text, 3, 2)]);
    Document := OpenDoc(TechFile);
    if IsNullEmpty(TechFile) then begin
      Desktop.Terminate;
      Desktop := Unassigned;
      ServiceManager := Unassigned;
      Exit;
    end;

    try
// všechny listy 26.5.16
      Sheets := Document.getSheets;
    except
      on E: Exception do ShowMessage('Nepodaøilo se otevøít listy ' + ^M + E.Message);
    end;
    if IsNullEmpty(Sheets) then begin
      Document.Dispose;
      Document.Unassigned;
      Desktop.Terminate;
      Desktop := Unassigned;
      ServiceManager := Unassigned;
      Exit;
    end;
  end;  // with dmDES_OO_Common

// cyklus pro všechny techniky v cbxJmeno
  try
    for ixJmeno := 0 to cbxJmeno.Items.Count-1 do begin
      Technik := Copy(cbxJmeno.Items[ixJmeno], 1, Pos(' /', cbxJmeno.Items[ixJmeno])-1);
//      if (Technik = 'AT') or (Technik = 'PJ') or (Technik = 'MPi') then Continue;
//      if (Technik = 'JuK') then Continue;         // 13.7.19 - nemá hodiny v aplikaci
// 26.5.16 když není v Excelu, tak nic
      if not Sheets.hasByName(Technik) then Continue;
      try
        Sheet := Sheets.getByName(Technik);
      except
        on E: Exception do begin
          ShowMessage('Nepodaøilo se otevøít list ' + Technik + ^M + E.Message);
          Continue;
        end;
      end;
      if dmDES_OO_Common.IsNullEmpty(Sheet) then begin
        ShowMessage('Nepodaøilo se otevøít list ' + Technik);
        Continue;
      end;
// list technika se nastaví jako aktivní
      Document.getCurrentController.setActiveSheet(Sheet);
// pomocí cbxJmenoChange se vyplní asgMain, z nìj se bude ukládat do Excelu
      cbxJmeno.ItemIndex := ixJmeno;
      cbxJmenoChange(nil);

// vloží se hodnoty pro jednotlivé mìsíce
      for Radek := 1 to 12 do with asgMain do begin
// vloží se hodnoty
        if Cells[2, Radek] <> '' then begin     // jsou nìjaké hodiny
// 12.12.18 pøidán sloupec s hodinovou sazbou, posun o jednu dopraba
          Sheet.getCellByPosition(7, Radek).Value := Floats[2, Radek];           // hodiny
          Sheet.getCellByPosition(7, Radek).HoriJustify := ooHAlignRight;
          if Floats[1, Radek] > Floats[2, Radek] then
            Sheet.getCellByPosition(7, Radek).CharColor := $A0A0A0               // šedivì
          else Sheet.getCellByPosition(7, Radek).CharColor := clBlack;
        end else Sheet.getCellByPosition(7, Radek).String := '';
        if Cells[4, Radek] <> '' then begin
          Sheet.getCellByPosition(8, Radek).Value := Floats[4, Radek];           // kilometry
          Sheet.getCellByPosition(8, Radek).HoriJustify := ooHAlignRight;
          if Floats[3, Radek] > Floats[4, Radek] then
            Sheet.getCellByPosition(8, Radek).CharColor := $A0A0A0
          else Sheet.getCellByPosition(8, Radek).CharColor := clBlack;
        end else Sheet.getCellByPosition(8, Radek).String := '';
// 24.8.19
        if Cells[5, Radek] <> '' then begin
          Sheet.getCellByPosition(10, Radek).Value := Floats[5, Radek];           // bonusy
          Sheet.getCellByPosition(10, Radek).HoriJustify := ooHAlignRight;
        end;
        Sheet.getCellByPosition(6, Radek).Value := Floats[6, Radek];             // hodinové sazby
        Sheet.getCellByPosition(6, Radek).HoriJustify := ooHAlignRight;
        Sheet.getCellByPosition(6, Radek).CharColor := clBlack;
      end;  // for Radek ... do with asgMain

// 13.12.18 tržby od zákazníkù ze simple_billings
// 10.11.2020 pro sichr se všechno nejdøív vymaže
//      for Radek := 3 to 5 do Sheet.getCellByPosition(15, Radek).string := '';
//      for Radek := 16 to 28 do
      for Sloupec := 3 to 5 do
        for Radek := 15 to 28 do Sheet.getCellByPosition(Sloupec, Radek).string := '';
      with qrMain do begin
        Close;
        SQLStr := 'SELECT SUM(SB.price) AS Trzba, SB.device FROM simple_billings SB, users U'     // vydìlal vùbec nìco ?
        + ' WHERE SB.user_id = U.id'
        + ' AND U.issue_shortcut = ''' + Technik + ''''
        + ' AND SB.cash = 0'
        + ' AND SB.description NOT LIKE ''Placeno%'''
//        + ' AND EXTRACT(year FROM SB.created_at) = ' + FormatDateTime('yyyy', Date)   11.1.2022
        + ' AND EXTRACT(year FROM SB.created_at) = ' + cbxRok.Text
        + ' GROUP BY device';
        SQL.Text := SQLStr;
        Open;
        if RecordCount > 0 then begin       // když vydìlal, pøipraví se tabulka a uloží souèty
          Sheet.getCellByPosition(3, 15).string := 'tržby';
          Sheet.getCellByPosition(4, 15).string := 'práce';
          Sheet.getCellByPosition(5, 15).string := 'mat.';
          Sheet.getCellByPosition(3, 16).string := 'I.';
          Sheet.getCellByPosition(3, 17).string := 'II.';
          Sheet.getCellByPosition(3, 18).string := 'III.';
          Sheet.getCellByPosition(3, 19).string := 'IV.';
          Sheet.getCellByPosition(3, 20).string := 'V.';
          Sheet.getCellByPosition(3, 21).string := 'VI.';
          Sheet.getCellByPosition(3, 22).string := 'VII.';
          Sheet.getCellByPosition(3, 23).string := 'VIII.';
          Sheet.getCellByPosition(3, 24).string := 'IX.';
          Sheet.getCellByPosition(3, 25).string := 'X.';
          Sheet.getCellByPosition(3, 26).string := 'XI.';
          Sheet.getCellByPosition(3, 27).string := 'XII.';
          Sheet.getCellByPosition(3, 28).string := 'Celkem';
          for Sloupec := 3 to 5 do Sheet.getCellByPosition(Radek, 15).HoriJustify := ooHAlignCenter;
          for Radek := 16 to 28 do Sheet.getCellByPosition(3, Radek).HoriJustify := ooHAlignCenter;
          while not EOF do begin
            if FieldByName('device').AsInteger = 0 then Sheet.getCellByPosition(4, 28).Value := FieldByName('Trzba').AsFloat
            else Sheet.getCellByPosition(5, 28).Value := FieldByName('Trzba').AsFloat;
            Next;
          end;
          Close;
          SQLStr := 'SELECT EXTRACT(month FROM SB.created_at) AS Radek, SUM(SB.price) AS Trzba, SB.device'    // tržby po mìsících
          + ' FROM simple_billings SB, users U'
          + ' WHERE SB.user_id = U.id'
          + ' AND U.issue_shortcut = ''' + Technik + ''''
          + ' AND SB.cash = 0'
          + ' AND SB.description NOT LIKE ''Placeno%'''
//          + ' AND EXTRACT(year FROM SB.created_at) = ' + FormatDateTime('yyyy', Date)    11.1.2022
          + ' AND EXTRACT(year FROM SB.created_at) = ' + cbxRok.Text
          + ' GROUP BY EXTRACT(month FROM SB.created_at), device'
          + ' ORDER BY EXTRACT(month FROM SB.created_at)';
          SQL.Text := SQLStr;
          Open;
          while not EOF do begin
            if FieldByName('Trzba').AsFloat > 0 then
              if FieldByName('device').AsInteger = 0 then
                Sheet.getCellByPosition(4, FieldByName('Radek').AsInteger + 15).Value := FieldByName('Trzba').AsFloat
              else Sheet.getCellByPosition(5, FieldByName('Radek').AsInteger + 15).Value := FieldByName('Trzba').AsFloat;
            Next;
          end;
        end;  // if
        Close;
      end;  // with qrMain

// cena zboží u technika
      with qrAbra do begin
        Close;
        SQLStr := 'SELECT SUM(PurchasePrice * Quantity) AS Cena FROM StoreSubCards'
        + ' WHERE Store_Id = (SELECT Id FROM Stores WHERE Code = ''' + Technik + ''')';
        SQL.Text := SQLStr;
        Open;
        Sheet.getCellByPosition(2, 1).String := Format('sklad: %f', [FieldByName('Cena').AsFloat]) + FormatDateTime('  d.m.yy', Date);
        Sheet.getCellByPosition(2, 1).HoriJustify := ooHAlignCenter;
        Sheet.getCellByPosition(2, 1).CharColor := clBlue;     // kvùli jinému formátu barev to bude èervená
        Close;
      end;  // with qrAbra

// 20.10.19 kasa technika do asgKasa  6.2.21 upraveno
      with  qrMain, asgKasa do begin
        SQLStr := 'SELECT Id FROM users'
        + ' WHERE issue_shortcut = ''' + Copy(cbxJmeno.Text, 1, Pos(' /', cbxJmeno.Text)) + '''';
        SQL.Text := SQLStr;
        Open;
        TechId := Fields[0].AsString;          // pro ukládání do wallet_transactions
        Close;
// poèáteèní stav kasy se ukládá jen pro rok 2010 (první uložený rok) - problém se souèty v aplikaci
        if cbxRok.Text = '2010' then begin
          Inc(RadekasgKasa);
          RowCount := RadekasgKasa + 1;
          Cells[0, RadekasgKasa] := FormatDateTime('dd.mm.yyyy', StrToDateTime(Sheet.GetCellByPosition(0, 2).string + cbxRok.Text));
          Cells[1, RadekasgKasa] := TechId;
          Floats[2, RadekasgKasa] := Sheet.GetCellByPosition(1, 2).value;
          if (Cells[2, RadekasgKasa] = '') then Floats[2, RadekasgKasa] := 0;
          Cells[3, RadekasgKasa] := 'poèáteèní stav';
          sDatum := Cells[0, RadekasgKasa];
        end else  // není 2010
          sDatum := FormatDateTime('dd.mm.yyyy', StrToDateTime(Sheet.GetCellByPosition(0, 2).string + cbxRok.Text));
// další øádky až do konce
        Radek := 2;
        while (Radek < dmDES_OO_Common.LastRow(Sheet)+1) do begin
          Inc(Radek);
          Inc(RadekasgKasa);
          RowCount := RadekasgKasa + 1;
          if Sheet.GetCellByPosition(0, Radek).string <> '' then              // je-li vyplnìné datum, znovu se uloží
            sDatum := FormatDateTime('dd.mm.yyyy', StrToDateTime(Sheet.GetCellByPosition(0, Radek).string + cbxRok.Text));
          Cells[0, RadekasgKasa] := sDatum;
          Cells[1, RadekasgKasa] := TechId;
          if Sheet.GetCellByPosition(1, Radek).string <> '' then              // je-li vyplnìná èástka, uloží se
            Floats[2, RadekasgKasa] := Sheet.GetCellByPosition(1, Radek).value;
          if Sheet.GetCellByPosition(2, Radek).string <> '' then              // je-li vyplnìný text, uloží se
            Cells[3, RadekasgKasa] := Sheet.GetCellByPosition(2, Radek).string;
          if (Cells[2, RadekasgKasa] = '') and (Cells[3, RadekasgKasa] = '') then begin
            Dec(RadekasgKasa);
            RowCount := RadekasgKasa + 1;
          end;
        end;  // while
      end;  // with qrMain, asgKasa

// konec práce s jedním technikem
    end;  // for ixJmeno

// uložení dat z asgKasa do wallet_transactions - všichni technici naráz
    with  qrMain, asgKasa do begin
      SQL.Text := 'START TRANSACTION';
      ExecSQL;
      SQL.Text := 'DELETE FROM wallet_transactions'
      + ' WHERE YEAR(paid_at) = ' + Ap + cbxRok.Text + Ap;
      ExecSQL;

      for RadekasgKasa := 1 to RowCount-1 do begin
        if (Cells[2, RadekasgKasa] = '') and (Cells[3, RadekasgKasa] = '') then Continue;    // nevyplnìný øádek se neuloží
        Row := RadekasgKasa;
        SQLStr := 'INSERT INTO wallet_transactions ('
        + ' user_id,'
        + ' amount,'
        + ' description,'
        + ' paid_at,'
        + ' is_valid,'
        + ' created_at,'
        + ' updated_at'
        + ') VALUES ('
        + Cells[1, RadekasgKasa] + ','
        + StringReplace(Cells[2, RadekasgKasa], ',', '.', []) + ','
        + Ap + Cells[3, RadekasgKasa] + ApC
        + Ap + FormatDateTime('yyyy-mm-dd', StrToDate(Cells[0, RadekasgKasa])) + ApC
        + '1,'
        + Ap + FormatDateTime('yyyy-mm-dd hh:MM:ss', Now) + ApC
        + Ap + FormatDateTime('yyyy-mm-dd hh:MM:ss', Now) + ApZ;
        SQL.Text := SQLStr;
        try
          ExecSQL;
        except
          on E: Exception do begin
            ShowMessage('Chyba pøi ukládání' + ^M + E.Message);
            SQL.Text := 'ROLLBACK';
            ExecSQL;
            Exit;
          end;
        end;
        Application.ProcessMessages;
      end;
      SQL.Text := 'COMMIT';
      ExecSQL;
    end;  // with qrMain, asgKasa

    try
      Document.Store;
      cbxJmeno.ItemIndex := -1;
      asgMain.ClearNormalCells;
      ShowMessage('Hotovo.');
    except
      Application.MessageBox(PChar('Nedá se uložit ' +  XLSName + '. Není nìkde otevøený ?' ), 'Open Office Calc', MB_ICONERROR + MB_OK);
    end;

  finally
    cbxJmeno.ItemIndex := -1;
    asgMain.ClearNormalCells;
// Clean up
    Document.Dispose;
    dmDES_OO_Common.Desktop.Terminate;
    dmDES_OO_Common.Desktop := Unassigned;
    dmDES_OO_Common.ServiceManager := Unassigned;
    Screen.Cursor := crDefault;
  end;  // try
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.btPrevodClick(Sender: TObject);
// pøevod zùstatkù do  dalšího roku
// 5.1.2024 podle toho, co je vybráno v cbxJmeno, se pøevádí všechny zùstatky, nebo jen vybraný technik
const
  ooHAlignStd = 0; //    com.sun.star.table.CellHoriJustify.STANDARD
  ooHAlignLeft = 1; //   com.sun.star.table.CellHoriJustify.LEFT
  ooHAlignCenter = 2; // com.sun.star.table.CellHoriJustify.CENTER
  ooHAlignRight = 3; //  com.sun.star.table.CellHoriJustify.RIGHT
var
  ServiceManager,
  Desktop,
  LoadParams,
  DokumentOld,
  DokumentNew,
  SheetOld,
  SheetNew: variant;
  Technik,
  XLSNameOld,
  XLSNameNew: string;
  iJmeno,
  iRok: integer;

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
  Screen.Cursor := crHourGlass;
  with asgMain do try
// OO ještì nebyl spuštìn
    if VarIsEmpty(ServiceManager) then try
// zkusí se, jestli nebìží
      ServiceManager := GetActiveOleObject('com.sun.star.ServiceManager');
    except try
// jinak se spustí
      ServiceManager := CreateOleObject('com.sun.star.ServiceManager');
      except
        on E: Exception do begin
          ShowMessage('Nelze spustit OpenOffice' + ^M + E.Message);
          Exit;
        end;
      end;  // try
    end;  // try
    if (VarIsEmpty(ServiceManager) or VarIsNull(ServiceManager)) then begin
      ShowMessage('Nelze spustit OpenOffice.');
      Exit;
    end;
// povedlo se
    if VarIsEmpty(Desktop) then Desktop := ServiceManager.CreateInstance('com.sun.star.frame.Desktop');
// soubory technikù
    TechDir := StringReplace(ExcludeTrailingPathDelimiter(TechDir), '\', '/', [rfReplaceAll]);
    iRok := StrToInt(Copy(cbxRok.Text, 3, 2));
    XLSNameNew := Format(TechDir + '/%d/technici%d.xls', [iRok, iRok]);
    XLSNameOld := Format(TechDir + '/%d/technici%d.xls', [iRok-1, iRok-1]);
// (popis parametrù v OO MediaDescriptor)
    LoadParams := VarArrayCreate([0, 0], varVariant);
// listy nebudou vidìt
    LoadParams[0] := SetParams('Hidden', True);
// soubor se otevøe
    try
      DokumentOld := Desktop.LoadComponentFromURL('file:///' + XLSNameOld, '_default', 0, LoadParams);
      DokumentNew := Desktop.LoadComponentFromURL('file:///' + XLSNameNew, '_default', 0, LoadParams);
    except
      on E: Exception do begin
        ShowMessage('Nepodaøilo se otevøít Technici.xls' + ^M + E.Message);
        if not (VarIsEmpty(ServiceManager) or VarIsNull(ServiceManager)) then ServiceManager := Unassigned;
        Exit;
      end;
    end;
// 5.1.24 celý cyklus probìhne pouze pokud není v cbxJmeno vybraný technik, tj. cbxJmeno.Text = 'Jméno', jinak
// se aktualizuje pouze list vybraného technika
// aktivní list (podle technika)
//    for iJmeno := 1 to DokumentNew.GetSheets.GetCount-1 do begin
    for iJmeno := 1 to DokumentOld.GetSheets.GetCount-1 do begin
      try
        SheetOld := DokumentOld.getSheets.GetByIndex(iJmeno);
        Technik := SheetOld.GetName;
      except
        on E: Exception do begin
          if (cbxJmeno.Text = 'Jméno') then
           ShowMessage('Nepodaøilo se otevøít list ' + Technik + ' v ' + XLSNameOld + ^M + E.Message);
          Continue;
        end;
      end;
      try
        SheetNew := DokumentNew.getSheets.getByName(Technik);
      except
        on E: Exception do begin
          if (cbxJmeno.Text = 'Jméno') then
           ShowMessage('Nepodaøilo se otevøít list ' + Technik + ' v ' + XLSNameNew + ^M + E.Message);
          Continue;
        end;
      end;
// 5.1.24
      if (cbxJmeno.Text = 'Jméno') or (Copy(cbxJmeno.Text, 1, Pos(' /', cbxJmeno.Text)-1) = Technik) then begin
// zùstatek v pokladnì
        SheetNew.getCellByPosition(1, 2).Value := SheetOld.getCellByPosition(1, 0).Value;
// nevyfakturovaný výdìlek
        SheetNew.getCellByPosition(9, 0).string := Format('rùzné %s', [SheetOld.getCellByPosition(4, 0).Value]);
      end;
    end;  // for iJmeno
    try
      DokumentNew.Store;
      ShowMessage('Hotovo.');
    except
      Application.MessageBox(PChar('Nedá se uložit ' +  XLSNameNew + '. Není nìkde otevøený ?' ), 'Open Office Calc', MB_ICONERROR + MB_OK);
    end;
    DokumentNew.Dispose;
    DokumentOld.Dispose;
  finally
    DokumentOld := Unassigned;
    SheetOld := Unassigned;
    DokumentNew := Unassigned;
    SheetNew := Unassigned;
    if not (VarIsEmpty(Desktop) or VarIsNull(Desktop)) then Desktop := Unassigned;
    if not (VarIsEmpty(ServiceManager) or VarIsNull(ServiceManager)) then ServiceManager := Unassigned;
    Screen.Cursor := crDefault;
  end;  // with fmMain ...

end;

end.
