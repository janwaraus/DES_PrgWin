// 23.10.2019 Program k automatické tvorbì faktur za hovorné technikù. Úèet od T-Mobilu pøichází ve formátu PDF, nejdøíve se
// ruènì pøevede do textu v UTF8 free programem FDF2TEXT firmy LotApps.
// Zpracování pøevedeného textu s využitím regulárních výrazù do dat pro fakturaci a vlastní tvoba faktur se provádí zde, zbytek v Code.
// 5.1.2020 Zmìna programu pro pøevod PDF do textu, použit program "blb2txt" z projektu Balabolka (http://balabolka.site/btext.htm),
// který se spouští z pøíkazového øádku. S tím souvisí i úprava regulárních výrazù.

unit MainFT;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms, ComObj, ComCtrls,
  Dialogs, Grids, AdvObj, BaseGrid, AdvGrid, StdCtrls, ExtCtrls, Math, DateUtils, IniFiles, RegularExpressions,
  ZAbstractConnection, ZConnection, DB, ZAbstractRODataset, ZAbstractDataset, ZDataset;

type
  TfmMainFT = class(TForm)
    pnTop: TPanel;
    btNacist: TButton;
    btFO: TButton;
    asgMain: TAdvStringGrid;
    odPDF: TOpenDialog;
    asgTelefony: TAdvStringGrid;
    lbObdobi: TLabel;
    qrAbra: TZQuery;
    cnAbra: TZConnection;
    procedure btNacistClick(Sender: TObject);
    procedure btFOClick(Sender: TObject);
    procedure asgMainGetFormat(Sender: TObject; ACol: Integer; var AStyle: TSortStyle; var aPrefix, aSuffix: string);
    procedure FormShow(Sender: TObject);
    procedure asgMainGetAlignment(Sender: TObject; ARow, ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
  private
    TechDir: string;
  public
    ExePath: string;
    AbraOLE: variant;
  end;

const
  Ap = chr(39);
  ApC = Ap + ',';
  ApZ = Ap + ')';
  FO_Id: string[10] = 'H000000101';
  MyAddress_Id: string[10] = '7000000101';
  MyAccount_Id: string[10] = '1400000101';       // Fio
  User_Id: string[10] = '2200000101';            // automatická fakturace
  Payment_Id: string[10] = '1000000101';         // typ platby: na bankovní úèet

var
  fmMainFT: TfmMainFT;

implementation

uses DES_common, Code;

{$R *.dfm}

// ------------------------------------------------------------------------------------------------

procedure TfmMainFT.FormShow(Sender: TObject);
var
  DESIni: TIniFile;
begin
  ExePath := ExtractFilePath(ParamStr(0));
// AbraDESProgramy.ini ?
  if not(FileExists(ExePath + 'AbraDESProgramy.ini'))
    and not(FileExists(ExePath + '..\DE$_Common\AbraDESProgramy.ini')) then
  begin
    Application.MessageBox(PChar('Nenalezen soubor AbraDESProgramy.ini, program ukonèen'), 'AbraDESProgramy.ini', MB_OK + MB_ICONERROR);
    Application.Terminate;
  end;
// DESIni
  if FileExists(ExePath + 'AbraDESProgramy.ini') then
    DESIni := TIniFile.Create(ExePath + 'AbraDESProgramy.ini')
  else
    DESIni := TIniFile.Create(ExePath + '..\DE$_Common\AbraDESProgramy.ini');
// parametry z AbraDESProgramy.ini
  with DESIni do try
    cnAbra.Database := ReadString('Preferences', 'AbraDB', '');
    cnAbra.User := ReadString('Preferences', 'AbraUN', '');
    cnAbra.Password := ReadString('Preferences', 'AbraPW', '');
    TechDir := ReadString('Preferences', 'Dirs', 'W:\');
    TechDir := StringReplace(ExcludeTrailingPathDelimiter(TechDir), '\', '/', [rfReplaceAll]);
  finally
    DESIni.Free;
  end;
  try
    cnAbra.Connect;
  except on E: exception do
    begin
      Application.MessageBox(PChar('Nedá se pøipojit k databázi Abry, program ukonèen.' + ^M + E.Message), 'Abra', MB_ICONERROR + MB_OK);
      Application.Terminate;
    end;
  end;
  asgMain.CheckFalse := '0';
  asgMain.CheckTrue := '1';
// naètení souborù
//  btNacistClick(Self);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMainFT.btNacistClick(Sender: TObject);
var
//  FileRow: RawByteString;
  Technik,
  TmpStr: string;
  TxtFile: TextFile;
  i,
  Radek: integer;
  RegExpr: TRegEx;
begin
// otevøe se soubor s úètovanými èástkami od T-Mobile v PDF
  if not odPDF.Execute then Exit;
// pøevod z PDF do textu
  dmCode.PrevodPDFdoTXT(odPDF.FileName);
// útrata za jednotlivá èísla
  dmCode.AnalyzaVyuctovani;
{// naète se soubor s telefonními èísly technikù (asgTelefony.Visible := False)
  with asgTelefony do begin
    ClearNormalCells;
    RowCount := 2;
    LoadFromXLS(FormatDateTime('"J:\Eurosignal\telefony"yy".xls"', Date));
    ColCount := 3;
    ColWidths[1] := 0;
    for i := 1 to RowCount-1 do                            // èísla bez mezer
      Cells[0, i] := StringReplace(Cells[0, i], ' ', '', [rfReplaceAll]);
  end;
// otevøe se soubor s úètovanými èástkami od T-Mobile
  AssignFile(TxtFile, ExePath + 'Mobily.tmp');
  Reset(TxtFile);
  Radek := 0;
// a uloží do asgMain
  with asgMain do begin
    while not EOF(TxtFile) do begin
      Inc(Radek);
      RowCount := Radek + 1;
//      Readln(TxtFile, FileRow);                            // FileRow musí být RawByteString, jinak to v XE nebude fungovat
//      TmpStr := UTF8Decode(FileRow);                       // soubor je v UTF-8
      Readln(TxtFile, TmpStr);
      Cells[1, Radek] := TmpStr;
// období úètování do labelu
      if RegExpr.IsMatch(TmpStr, 'za období') then
        lbObdobi.Caption := Copy(TmpStr, Pos('za období', TmpStr)+10, Pos(' K ', TmpStr) - Pos('za obdob', TmpStr)-10);
// když je na zaèátku øádku devìt èíslic, je to úètované telefonní èíslo
      if RegExpr.IsMatch(TmpStr, '^\d{9}{'{) then begin
        Cells[2, Radek] := Copy(TmpStr, 1, 9);
// v asgTelefony se k èíslu najde technik
{        with asgTelefony do
          for i := 1 to RowCount-1 do
            if (Cells[0, i] = asgMain.Cells[2, Radek]) then begin
              asgMain.Cells[1, Radek] := Cells[2, i];      // technik se zkopíruje
              Break;
            end;  // if
// když na zaèátku øádku není devìt èíslic ubere se poèitadlo a øádek se pøepíše dalším záznamem
      end else Dec(Radek);
// øádek s úètovanou èástkou má na konci Kè a pøijde do sloupce 2
//      if RegExpr.IsMatch(TmpStr, 'Kè$') then
//        Cells[3, Radek] := StringReplace(TmpStr, 'Kè', '', [rfReplaceAll]);
      if RegExpr.IsMatch(TmpStr, 'Kc$') then
        Cells[3, Radek] := StringReplace(TmpStr, 'Kc', '', [rfReplaceAll]);
    end;  // while not EOF(TxtFile)
// konec importu ze souboru
    CloseFile(TxtFile);
// vyhodí se øádky s èísly, která se nefakturují
    Radek := 0;
    while Radek < RowCount-1 do begin
      Inc(Radek);
      if Cells[1, Radek] = '' then begin
        RemoveRows(Radek, 1);
        Dec(Radek);
      end;
    end;  // while
// jsou-li ve sloupci 2 dvì èástky, druhá je za platební služby a dá se na samostatný øádek
    Radek := 0;
    while Radek < RowCount-1 do begin
      Inc(Radek);
{     if RegExpr.IsMatch(Cells[3, Radek], '\d+,\d{2}{.+\d+,\d{2}{'
{ then begin     // dvì èástky dddd,dd dd,dd
        Inc(Radek);
        asgMain.InsertRows(Radek, 1);                      // platební služby na nový øádek
        Cells[1, Radek] := Cells[1, Radek-1];              // tentýž technik
        Cells[2, Radek] := 'platební služby';
        Cells[3, Radek] := Copy(Cells[3, Radek-1], Pos(' ', Cells[3, Radek-1])+2, Length(Cells[3, Radek-1]));
        Cells[3, Radek-1] := Copy(Cells[3, Radek-1], 1, Pos(' ', Cells[3, Radek-1])-1);
      end;  // if dvì èástky
    end;  // while
// asgMain se setøídí podle technika
    SortSettings.Column := 1;
    QSort;
// fajfky
    Technik := '';
    for Radek := 1 to RowCount-1 do begin
      if Cells[1, Radek] <> Technik then begin
        Technik := Cells[1, Radek];
        AddCheckBox(0, Radek, False, True);                 // pøidá se možnost fajfky
      end;
    end;
  end;  // with asgMain  }
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMainFT.asgMainGetAlignment(Sender: TObject; ARow, ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
  if (ACol = 1) then HAlign := taCenter;
  if (ACol = 3) then HAlign := Classes.taRightJustify;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMainFT.asgMainGetFormat(Sender: TObject; ACol: Integer; var AStyle: TSortStyle; var aPrefix, aSuffix: string);
// kvùli správnému tøídìní
begin
  if (ACol = 2) then AStyle := ssAlphanumeric;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMainFT.btFOClick(Sender: TObject);
// vytvoøení faktur
var
  Period_Id: string[10];
  Radek: integer;
  DatumFaktury: double;
  Technik,
  SQLStr: string;

begin
  Screen.Cursor := crHourGlass;
// pøipojení k Abøe
  AbraOLE := CreateOLEObject('AbraOLE.Application');
  if not AbraOLE.Connect('@DES') then begin
    ShowMessage('Problém s Abrou (connect DES).');
    Screen.Cursor := crDefault;
    Exit;
  end;
  if not AbraOLE.Login('Supervisor', '') then begin
    ShowMessage('Problém s Abrou (Login Supervisor).');
    Screen.Cursor := crDefault;
    Exit;
  end;
// hlavní práce
  with asgMain do
    for Radek := 1 to RowCount-1 do
      if Ints[0, Radek] = 1 then
        dmCode.FakturaTechnika(Cells[1, Radek]);

  if not (VarIsEmpty(AbraOLE) or VarIsNull(AbraOLE) or VarIsClear(AbraOLE)) then try
    AbraOLE.Logout;
  except
  end;
  Screen.Cursor := crDefault;
end;

end.
