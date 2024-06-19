unit AbraFunkce_demo;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.StrUtils,
  System.Classes, Vcl.Graphics, System.Generics.Collections,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Data.DB,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, FireDAC.Stan.Intf,
  ComObj,
  FireDAC.Stan.Option, FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf,
  FireDAC.Stan.Def, FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys,
  FireDAC.Phys.PG, FireDAC.Phys.PGDef, FireDAC.VCLUI.Wait, FireDAC.Stan.Param,
  FireDAC.DatS, FireDAC.DApt.Intf, FireDAC.DApt, FireDAC.Comp.DataSet,
  FireDAC.Comp.Client, ZAbstractConnection, ZConnection, AdvUtil, Vcl.Grids,
  AdvObj, BaseGrid, AdvGrid, AdvSprd, tmsAdvGridExcel, AdvGridWorkbook;

type
  TForm1 = class(TForm)
    Memo1: TMemo;
    btnDoIt1: TButton;
    edit1: TEdit;
    lbl1: TLabel;
    Button1: TButton;
    Button2: TButton;
    dbNewVoip: TZConnection;
    qrNewVoip: TZQuery;
    AdvSpreadGrid1: TAdvSpreadGrid;
    AdvStringGrid1: TAdvStringGrid;
    AdvGridExcelIO1: TAdvGridExcelIO;
    AdvGridWorkbook1: TAdvGridWorkbook;
    AdvGridExcelIO2: TAdvGridExcelIO;
    ListBox1: TListBox;
    procedure btnDoIt1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure CreateExcelFile;
    procedure CreateExcelFileUsingTAdvSpreadGrid;
    procedure ExportGridDataToExcel;
    procedure CopyFileToNetworkDrive(const SourceFile, DestIP, DestShare, DestFileName: string);
    procedure LoadExcelFileAndListSheets(const FileName: string);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;


implementation

{$R *.dfm}

uses DesUtils, AbraEntities, Superobject, AArray;


procedure TForm1.Button1Click(Sender: TObject);
var
  maxCisloVypisu : integer;
  accountName : string;
  abraBankaccount : TAbraBankaccount;
  doklad1 : TDoklad;
begin
  try
    // Example usage
    CopyFileToNetworkDrive('C:\data\VypisX.txt', '192.168.0.7', 'public', 'Vypis_test.txt');
    Memo1.Lines.Add('File copied successfully.');
  except
    on E: Exception do
      Memo1.Lines.Add('An error occurred: ' + E.Message);
  end;

  // DesU.opravRadekVypisuPomociPDocument_ID('2UF2000101', '9M6E000101', 'KQ0U000101', '03');
  // Memo1.Lines.Add('Radek opraveeen:');                .
  //DesU.dbAbra.Reconnect;
  {
  abraBankAccount := TAbraBankaccount.create();
  abraBankaccount.loadByNumber('171336270/0300');
  Memo1.Lines.Add('Poèet výpisù: ' + IntToStr(abraBankaccount.getPocetVypisu('2023')));
  Memo1.Lines.Add('Jméno (s UTF8Decode): ' + UTF8Decode(abraBankaccount.name));
  Memo1.Lines.Add('Jméno (bez úpravy): ' + abraBankaccount.name);
  // test
  DesU.dbAbra.Reconnect;

  doklad1 := TDoklad.create(trim(edit1.Text));

  Memo1.Lines.Add(doklad1.cisloDokladu);
  Memo1.Lines.Add(format('Nezaplaceno: %m', [doklad1.castkaNezaplaceno]));
  Memo1.Lines.Add(format('Pøeplatek: %m', [-doklad1.castkaNezaplaceno]));
  }
end;

procedure TForm1.Button2Click(Sender: TObject);
var
  OriginalString, LeftPart, RightPart: string;
begin
Memo1.Lines.Clear;

  Memo1.Lines.Append('Left part: ' + LeftPart);


end;

procedure TForm1.FormShow(Sender: TObject);
begin
  edit1.Text := 'KQ0U000101';
end;


procedure TForm1.btnDoIt1Click(Sender: TObject);
var
  text: string;

begin
  Memo1.Lines.Add('Nacitam xls');
  LoadExcelFileAndListSheets('technici24.xls');

end;


procedure TForm1.LoadExcelFileAndListSheets(const FileName: string);
var
  i: integer;
begin
{
  SheetNames := TStringList.Create;
  try
    AdvGridExcelIO1.LoadSheetName ('technici24.xls', SheetNames);

    // Display sheet names in a list box or any other control
    ListBox1.Items.Assign(SheetNames);
  finally
    SheetNames.Free;
  end;
  }



  try

    AdvGridExcelIO1.XLSImport(FileName);

    //AdvGridExcelIO1.LoadSheetNames();

     for i := 0 to AdvGridExcelIO1.SheetNamesCount - 1 do
  begin
    ListBox1.Items.Add(AdvGridExcelIO1.SheetNames[i]);
  end;


  finally

  end;
end;

procedure TForm1.CreateExcelFile;
var
  ExcelApp, Workbook, Worksheet: OleVariant;
  FileName: string;
begin
  // Create an instance of the Excel application
  ExcelApp := CreateOleObject('Excel.Application');
  try
    // Add a new workbook
    Workbook := ExcelApp.Workbooks.Add;
    // Access the first worksheet
    Worksheet := Workbook.Worksheets[1];

    // Write data to the cells
    Worksheet.Cells[1, 1] := 'Name';
    Worksheet.Cells[1, 2] := 'Age';
    Worksheet.Cells[2, 1] := 'Alice';
    Worksheet.Cells[2, 2] := 25;
    Worksheet.Cells[3, 1] := 'Bob';
    Worksheet.Cells[3, 2] := 30;

    // Define a filename
    FileName := ExtractFilePath(ParamStr(0)) + 'TestExcel.xlsx';
    // Save the workbook to a file
    Workbook.SaveAs(FileName);

    // Optionally display the Excel application to view the result
    ExcelApp.Visible := True;
  finally
    // Cleanup: Close the workbook without saving further changes
    Workbook.Close(False);
    // Quit Excel application
    ExcelApp.Quit;
  end;
end;


procedure TForm1.CreateExcelFileUsingTAdvSpreadGrid;
var
  //G/id: TAdvStringGrid;
  //ExcelIO: TAdvGridExcelIO;
  a: integer;
begin
  // Create an instance of the grid
  //Grid := TAdvStringGrid.Create(Self);
  try
    //Grid.Parent := Self;  // Optional: for visual representation in a Form
    AdvStringGrid1.ColCount := 3;   // Set the number of columns
    AdvStringGrid1.RowCount := 3;   // Set the number of rows

    // Populate the grid with some data
    AdvStringGrid1.Cells[0, 0] := 'Name';
    AdvStringGrid1.Cells[1, 0] := 'Age';
    AdvStringGrid1.Cells[2, 0] := 'Country';

    AdvStringGrid1.Cells[0, 1] := 'Alice';
    AdvStringGrid1.Cells[1, 1] := '30';
    AdvStringGrid1.Cells[2, 1] := 'USA';

    AdvStringGrid1.Cells[0, 2] := 'Bob';
    AdvStringGrid1.Cells[1, 2] := '25';
    AdvStringGrid1.Cells[2, 2] := 'UK';

    // Create an Excel IO handler for TAdvStringGrid
    //ExcelIO := TAdvGridExcelIO.Create(Self);
    try
      AdvGridExcelIO1.XLSExport(ExtractFilePath(ParamStr(0)) + 'GridData.xlsx', 'sesit123');

    //AdvGridExcelIO1.SheetNames;
      //AdvGridExcelIO1.xls
    finally
      //ExcelIO.Free;
    end;
  finally
    //Grid.Free;
  end;
end;


procedure TForm1.ExportGridDataToExcel;
begin
  // Populate the grid with some data
  AdvStringGrid1.ColCount := 3;   // Set number of columns
  AdvStringGrid1.RowCount := 3;   // Set number of rows

  // Setting headers
  AdvStringGrid1.Cells[0, 0] := 'Name';
  AdvStringGrid1.Cells[1, 0] := 'Age';
  AdvStringGrid1.Cells[2, 0] := 'Country';

  // Adding data
  AdvStringGrid1.Cells[0, 1] := 'Alice';
  AdvStringGrid1.Cells[1, 1] := '30';
  AdvStringGrid1.Cells[2, 1] := 'USA';

  AdvStringGrid1.Cells[0, 2] := 'Bob';
  AdvStringGrid1.Cells[1, 2] := '25';
  AdvStringGrid1.Cells[2, 2] := 'UK';

  // Export to Excel
  try
    AdvGridExcelIO1.XLSExport('C:\Data\GridDataa.xlsx');
  except
    on E: Exception do
      ShowMessage('Failed to export to Excel: ' + E.Message);
  end;
end;




{ *** DB VoIP spojení ***


 //text := DesU.ulozKomunikaci('2', '111', 'Ppøíliš žlutý mail''s book!');

 //dbNewVoIP.HostName := '10.0.250.12';
 dbNewVoIP.Connect;

 qrNewVoip.SQL.Text := 'SHOW server_version';

 qrNewVoip.SQL.Text := 'SELECT Amount FROM VoIP.Invoices_flat WHERE Num = 20069018';
 qrNewVoip.Open;

  while not qrNewVoip.Eof do
  begin
    // Access the result values using field names or indexes
    //ShowMessage(FDQuery1.Fields[0].AsString);
    Memo1.Lines.Add(qrNewVoip.Fields[0].AsString);
    qrNewVoip.Next;
  end;

 Memo1.Lines.Append('hh');

---

FDConnection1 := TFDConnection.Create(Self);
  FDConnection1.DriverName := 'PG';
  FDConnection1.Params.Add('Server=127.0.0.1');
  FDConnection1.Params.Add('Port=5432');
  FDConnection1.Params.Add('Database=dvdrental');
  FDConnection1.Params.Add('User_Name=postgres');
  FDConnection1.Params.Add('Password=masterkey');

  FDQuery1 := TFDQuery.Create(Self);
  FDQuery1.Connection := FDConnection1;
  FDQuery1.SQL.Text := 'SHOW server_version';
  FDQuery1.SQL.Text := 'SELECT title FROM public.film where film_id < 12';
  FDQuery1.Open;
  while not FDQuery1.Eof do
  begin
    // Access the result values using field names or indexes
    //ShowMessage(FDQuery1.Fields[0].AsString);
    Memo1.Lines.Add(FDQuery1.Fields[0].AsString);
    FDQuery1.Next;
  end;
  FDQuery1.Close; // Close the dataset



 Memo1.Lines.Add('Start');
 DesU.connectDbVoip;
  Memo1.Lines.Add('Connected do DB VoIP');

  with DesU.qrVoip do begin
    //SQL.Text := 'SELECT amount FROM voip.invoices_flat';
    SQL.Text := 'SHOW server_version;';
    Open;
    if not Eof then begin
      Memo1.Lines.Add(Fields[0].AsString);
      //Memo1.Lines.Add(floattostr(Fields[0].AsFloat));
      // +  FormatDateTime('dddd d of mmmm yyyy', FieldByName('From_date').AsDateTime));
      //Next;
    end;
    Close;
  end;

  /////

  FDConnection1.DriverName := 'PG';
  FDConnection1.Params.Add('Server=10.0.2.4');
  FDConnection1.Params.Add('Port=5432');
  FDConnection1.Params.Add('Database=cdasterisk');
  FDConnection1.Params.Add('User_Name=mnisekvoip');
  FDConnection1.Params.Add('Password=ssmplfgGHTF');

  FDQuery1 := TFDQuery.Create(Self);
  FDQuery1.Connection := FDConnection1;
  FDQuery1.SQL.Text := 'SHOW server_version';
  //FDQuery1.SQL.Text := 'SELECT title FROM public.film where film_id < 12';
  FDQuery1.Open;
  while not FDQuery1.Eof do
  begin
    // Access the result values using field names or indexes
    //ShowMessage(FDQuery1.Fields[0].AsString);
    Memo1.Lines.Add('PG version: ' + FDQuery1.Fields[0].AsString);
    FDQuery1.Next;
  end;
  FDQuery1.Close; // Close the dataset

}

procedure TForm1.CopyFileToNetworkDrive(const SourceFile, DestIP, DestShare, DestFileName: string);
var
  DestPath: string;
  Success: Boolean;
begin
  // Construct the destination path using IP address and share name
  DestPath := Format('\\%s\%s\%s', [DestIP, DestShare, DestFileName]);

  // Use CopyFile to copy the file
  Success := CopyFile(PChar(SourceFile), PChar(DestPath), False);

  // Check if the copy was successful
  if not Success then
    raise Exception.CreateFmt('Failed to copy file to %s. Error code: %d', [DestPath, GetLastError]);
end;



end.
