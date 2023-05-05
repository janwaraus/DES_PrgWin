unit AbraFunkce_demo;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls;

type
  TForm1 = class(TForm)
    Memo1: TMemo;
    btnDoIt1: TButton;
    edit1: TEdit;
    lbl1: TLabel;
    Button1: TButton;
    procedure btnDoIt1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Button1Click(Sender: TObject);

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
begin
  // DesU.opravRadekVypisuPomociPDocument_ID('2UF2000101', '9M6E000101', 'KQ0U000101', '03');
  // Memo1.Lines.Add('Radek opraveeen:');
  //DesU.dbAbra.Reconnect;
  abraBankAccount := TAbraBankaccount.create();
  abraBankaccount.loadByNumber('171336270/0300');
  Memo1.Lines.Add('Poèet výpisù: ' + IntToStr(abraBankaccount.getPocetVypisu('2023')));
  Memo1.Lines.Add('Jméno (s UTF8Decode): ' + UTF8Decode(abraBankaccount.name));
  Memo1.Lines.Add('Jméno (bez úpravy): ' + abraBankaccount.name);
  // test
end;

procedure TForm1.FormShow(Sender: TObject);
begin
  edit1.Text := 'KQ0U000101';
end;


procedure TForm1.btnDoIt1Click(Sender: TObject);
var
  text: string;
  doklad1 : TDoklad;
begin
  DesU.dbAbra.Reconnect;

  doklad1 := TDoklad.create(trim(edit1.Text));

  Memo1.Lines.Add(doklad1.cisloDokladu);
  Memo1.Lines.Add(format('Nezaplaceno: %m', [doklad1.castkaNezaplaceno]));
  Memo1.Lines.Add(format('Pøeplatek: %m', [-doklad1.castkaNezaplaceno]));

  {
  text := 'pøepl. | 20102980 | 20102980';
  text := 'pøepasdasdasd80';
  Memo1.Lines.Add( trim (copy(text, LastDelimiter('|', text) + 1, length(text)) ) );
  }



end;




end.
