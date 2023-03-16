unit DES_common;

interface

uses
  Windows, SysUtils, Classes, ShellAPI;

type
  TdmDES_common = class(TDataModule)
  private
    { Private declarations }
  public
    function UserName: string;                                                   // pøívìtivìjší GetUserName
    function CompName: string;                                                   // pøívìtivìjší GetComputerName
    function DiskDrives: string;                                                 // seznam lokálních a mapovaných diskù
    function IndexByName(DataObject: variant; Name: string): integer;            // do Abry: náhrada za nefunkèní DataObject.ValuByName(Name)
    function InvoiceToServer(aDir: string; aMonth, aYear: integer): integer;     // kopíruje soubory z adresáøe aDir\aYear\aMonth do /home/abrapdf/aYear
  end;

var
  dmDES_common: TdmDES_common;

implementation

{%CLASSGROUP 'Vcl.Controls.TControl'}

{$R *.dfm}

// ------------------------------------------------------------------------------------------------

function TdmDES_common.UserName: string;
// pøívìtivìjší GetUserName
var
  dwSize : DWord;
begin
  SetLength(Result, 32);
  dwSize := 31;
  GetUserName(PChar(Result), dwSize);
  SetLength(Result, dwSize-1);
end;

// ------------------------------------------------------------------------------------------------

function TdmDES_common.CompName: string;
// pøívìtivìjší GetComputerName
var
  dwSize : DWord;
begin
  SetLength(Result, 32);
  dwSize := 31;
  GetComputerName(PChar(Result), dwSize);
  SetLength(Result, dwSize);
end;

// ------------------------------------------------------------------------------------------------

function TdmDES_common.DiskDrives: string;
// vypíše seznam logických diskù (písmen)
var
  lwRes : LongWord;
  arDrives : array[0..128] of char;
  pDrive : PChar;
begin
  Result := '';
  lwRes := GetLogicalDriveStrings(SizeOf(arDrives), arDrives);
  if lwRes = 0 then Exit;
  pDrive := arDrives;
  while pDrive^ <> #0 do
  begin
    Result := Result + pDrive + ' ';
    Inc(pDrive, 4);
  end;
end;

// ------------------------------------------------------------------------------------------------

function TdmDES_common.IndexByName(DataObject: variant; Name: string): integer;
// do Abry: náhrada za nefunkèní DataObject.ValuByName(Name)
var
  i: integer;
begin
  Result := -1;
  i := 0;
  while i < DataObject.Count do begin
    if DataObject.Names[i] = Name then begin
      Result := i;
      Break;
    end;
    Inc(i);
  end;
end;

// ------------------------------------------------------------------------------------------------

function TdmDES_common.InvoiceToServer(aDir: string; aMonth, aYear: integer): integer;
// kopíruje soubory z adresáøe aDir\aYear\aMonth pomocí WinSCP na server aplikace.eurosignal.cz do /home/abrapdf/aYear. WinSCP má konfiguraci
// uloženou ve WinSCP.ini, chybové stavy ShellExecute jsou 0 až 32
var
  Handle: HWND;
begin
  Result := ShellExecute(Handle, 'open', PChar('WinSCP.com'),
   PChar(Format('/command "option batch abort" "option confirm off" "open AbraPDF" "synchronize remote %s\%4d\%2.2d /home/abrapdf/%4d" "exit"',
   [aDir, aYear, aMonth, aYear])), nil, SW_SHOWNORMAL);
end;

// ------------------------------------------------------------------------------------------------

end.
