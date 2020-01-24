// 23.10.2019 Program k automatick� tvorb� faktur za hovorn� technik�. ��et od T-Mobilu p�ich�z� ve form�tu PDF, nejd��ve se
// ru�n� p�evede do textu v UTF8 free programem FDF2TEXT firmy LotApps.
// Zpracov�n� p�eveden�ho textu s vyu�it�m regul�rn�ch v�raz� do dat pro fakturaci a vlastn� tvoba faktur se prov�d� zde, zbytek v Code.
// 5.1.2020 Zm�na programu pro p�evod PDF do textu, pou�it program "blb2txt" z projektu Balabolka (http://balabolka.site/btext.htm),
// kter� se spou�t� z p��kazov�ho ��dku. S t�m souvis� i �prava regul�rn�ch v�raz�.

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
  User_Id: string[10] = '2200000101';            // automatick� fakturace
  Payment_Id: string[10] = '1000000101';         // typ platby: na bankovn� ��et

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
    Application.MessageBox(PChar('Nenalezen soubor AbraDESProgramy.ini, program ukon�en'), 'AbraDESProgramy.ini', MB_OK + MB_ICONERROR);
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
      Application.MessageBox(PChar('Ned� se p�ipojit k datab�zi Abry, program ukon�en.' + ^M + E.Message), 'Abra', MB_ICONERROR + MB_OK);
      Application.Terminate;
    end;
  end;
  asgMain.CheckFalse := '0';
  asgMain.CheckTrue := '1';
// na�ten� soubor�
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
// otev�e se soubor s ��tovan�mi ��stkami od T-Mobile v PDF
  if not odPDF.Execute then Exit;
// p�evod z PDF do textu
  dmCode.PrevodPDFdoTXT(odPDF.FileName);
// �trata za jednotliv� ��sla
  dmCode.AnalyzaVyuctovani;
{// na�te se soubor s telefonn�mi ��sly technik� (asgTelefony.Visible := False)
  with asgTelefony do begin
    ClearNormalCells;
    RowCount := 2;
    LoadFromXLS(FormatDateTime('"J:\Eurosignal\telefony"yy".xls"', Date));
    ColCount := 3;
    ColWidths[1] := 0;
    for i := 1 to RowCount-1 do                            // ��sla bez mezer
      Cells[0, i] := StringReplace(Cells[0, i], ' ', '', [rfReplaceAll]);
  end;
// otev�e se soubor s ��tovan�mi ��stkami od T-Mobile
  AssignFile(TxtFile, ExePath + 'Mobily.tmp');
  Reset(TxtFile);
  Radek := 0;
// a ulo�� do asgMain
  with asgMain do begin
    while not EOF(TxtFile) do begin
      Inc(Radek);
      RowCount := Radek + 1;
//      Readln(TxtFile, FileRow);                            // FileRow mus� b�t RawByteString, jinak to v XE nebude fungovat
//      TmpStr := UTF8Decode(FileRow);                       // soubor je v UTF-8
      Readln(TxtFile, TmpStr);
      Cells[1, Radek] := TmpStr;
// obdob� ��tov�n� do labelu
      if RegExpr.IsMatch(TmpStr, 'za obdob�') then
        lbObdobi.Caption := Copy(TmpStr, Pos('za obdob�', TmpStr)+10, Pos(' K ', TmpStr) - Pos('za obdob', TmpStr)-10);
// kdy� je na za��tku ��dku dev�t ��slic, je to ��tovan� telefonn� ��slo
      if RegExpr.IsMatch(TmpStr, '^\d{9}{'{) then begin
        Cells[2, Radek] := Copy(TmpStr, 1, 9);
// v asgTelefony se k ��slu najde technik
{        with asgTelefony do
          for i := 1 to RowCount-1 do
            if (Cells[0, i] = asgMain.Cells[2, Radek]) then begin
              asgMain.Cells[1, Radek] := Cells[2, i];      // technik se zkop�ruje
              Break;
            end;  // if
// kdy� na za��tku ��dku nen� dev�t ��slic ubere se po�itadlo a ��dek se p�ep�e dal��m z�znamem
      end else Dec(Radek);
// ��dek s ��tovanou ��stkou m� na konci K� a p�ijde do sloupce 2
//      if RegExpr.IsMatch(TmpStr, 'K�$') then
//        Cells[3, Radek] := StringReplace(TmpStr, 'K�', '', [rfReplaceAll]);
      if RegExpr.IsMatch(TmpStr, 'Kc$') then
        Cells[3, Radek] := StringReplace(TmpStr, 'Kc', '', [rfReplaceAll]);
    end;  // while not EOF(TxtFile)
// konec importu ze souboru
    CloseFile(TxtFile);
// vyhod� se ��dky s ��sly, kter� se nefakturuj�
    Radek := 0;
    while Radek < RowCount-1 do begin
      Inc(Radek);
      if Cells[1, Radek] = '' then begin
        RemoveRows(Radek, 1);
        Dec(Radek);
      end;
    end;  // while
// jsou-li ve sloupci 2 dv� ��stky, druh� je za platebn� slu�by a d� se na samostatn� ��dek
    Radek := 0;
    while Radek < RowCount-1 do begin
      Inc(Radek);
{     if RegExpr.IsMatch(Cells[3, Radek], '\d+,\d{2}{.+\d+,\d{2}{'
{ then begin     // dv� ��stky dddd,dd dd,dd
        Inc(Radek);
        asgMain.InsertRows(Radek, 1);                      // platebn� slu�by na nov� ��dek
        Cells[1, Radek] := Cells[1, Radek-1];              // tent�� technik
        Cells[2, Radek] := 'platebn� slu�by';
        Cells[3, Radek] := Copy(Cells[3, Radek-1], Pos(' ', Cells[3, Radek-1])+2, Length(Cells[3, Radek-1]));
        Cells[3, Radek-1] := Copy(Cells[3, Radek-1], 1, Pos(' ', Cells[3, Radek-1])-1);
      end;  // if dv� ��stky
    end;  // while
// asgMain se set��d� podle technika
    SortSettings.Column := 1;
    QSort;
// fajfky
    Technik := '';
    for Radek := 1 to RowCount-1 do begin
      if Cells[1, Radek] <> Technik then begin
        Technik := Cells[1, Radek];
        AddCheckBox(0, Radek, False, True);                 // p�id� se mo�nost fajfky
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
// kv�li spr�vn�mu t��d�n�
begin
  if (ACol = 2) then AStyle := ssAlphanumeric;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMainFT.btFOClick(Sender: TObject);
// vytvo�en� faktur
var
  Period_Id: string[10];
  Radek: integer;
  DatumFaktury: double;
  Technik,
  SQLStr: string;

begin
  Screen.Cursor := crHourGlass;
// p�ipojen� k Ab�e
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
// hlavn� pr�ce
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
