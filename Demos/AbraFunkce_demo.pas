unit AbraFunkce_demo;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.StrUtils,
  System.Classes, Vcl.Graphics, System.Generics.Collections,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Data.DB,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, FireDAC.Stan.Intf,
  FireDAC.Stan.Option, FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf,
  FireDAC.Stan.Def, FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys,
  FireDAC.Phys.PG, FireDAC.Phys.PGDef, FireDAC.VCLUI.Wait, FireDAC.Stan.Param,
  FireDAC.DatS, FireDAC.DApt.Intf, FireDAC.DApt, FireDAC.Comp.DataSet,
  FireDAC.Comp.Client;

type
  TForm1 = class(TForm)
    Memo1: TMemo;
    btnDoIt1: TButton;
    edit1: TEdit;
    lbl1: TLabel;
    Button1: TButton;
    Button2: TButton;
    procedure btnDoIt1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);

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


  // DesU.opravRadekVypisuPomociPDocument_ID('2UF2000101', '9M6E000101', 'KQ0U000101', '03');
  // Memo1.Lines.Add('Radek opraveeen:');
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


text := DesU.ulozKomunikaci('2', '111', 'Ppøíliš žlutý mail''s book!');
 Memo1.Lines.Append(text);

end;







{ *** DB VoIP spojení ***

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



end.
