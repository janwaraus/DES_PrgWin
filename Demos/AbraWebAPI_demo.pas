unit AbraWebAPI_demo;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls,
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient,
  IdSMTP, IdMessage, IdMessageParts, IdAttachment, IdEMailAddress, IdAttachmentFile, IdText,
  IdExplicitTLSClientServerBase, IdMessageClient, IdSMTPBase;

type
  TForm1 = class(TForm)
    btnGet: TButton;
    Memo1: TMemo;
    Memo2: TMemo;
    btnPost: TButton;
    btnPut: TButton;
    editUrl: TEdit;
    lblUrl: TLabel;
    lblReqData: TLabel;
    Memo3: TMemo;
    btnCreateByAA: TButton;
    btnUpdateByAa: TButton;
    btnSendEmail: TButton;
    IdSMTP1: TIdSMTP;
    IdMessage1: TIdMessage;
    procedure btnGetClick(Sender: TObject);
    procedure btnPostClick(Sender: TObject);
    procedure btnPutClick(Sender: TObject);
    procedure btnCreateByAAClick(Sender: TObject);
    procedure btnUpdateByAaClick(Sender: TObject);
    procedure btnSendEmailClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;


implementation

{$R *.dfm}

uses DesUtils, Superobject, AArray;



////////////////////////////////////
// *** SuperObject (SO) demos *** //

procedure TForm1.btnGetClick(Sender: TObject);
var
  sResponse: string;
  MyObject: ISuperObject;
  mySuperArray: TSuperAvlEntry;
  item: ISuperObject;
begin
  //editUrl.Text := 'periods?where=code+gt+2015';
  editUrl.Text := 'bankstatements/36E2000101';

  sResponse := DesU.abraBoGet(editUrl.Text);

  memo2.Lines.Add(sResponse);

  MyObject := SO(sResponse);
  memo2.Lines.Add (MyObject.AsJSon(True));
  memo2.Lines.Add (MyObject.AsArray[0].AsString);
  memo2.Lines.Add (MyObject.AsArray[0].AsObject.S['name']);

  for item in MyObject do
    memo2.Lines.Add (item.AsObject.S['name']);

  //hodnota := MyObject.S['last'];
  //hodnota := SO(Mydata).S['last'];  //zkracena verze
end;

procedure TForm1.btnPostClick(Sender: TObject);
var
  sResponse: string;
  Json: string;
begin

editUrl.Text := 'issuedinvoice';

Json := '{'+
  '"docqueue_id": "L000000101",'+
  '"period_id": "1L20000101",'+
  '"docdate$date": "2017-06-04",'+
  '"firm_id": "2SZ1000101",'+
  '"description": "pppp�ipojen�, 2018020612",'+
  '"accdate$date": "2017-05-30",'+
  '"varsymbol": "2014020777",'+
  '"priceswithvat": true,'+

  '"rows": ['+
  '  {'+
  '    "totalprice": 3881.5,'+
  '    "division_id": "1000000101",'+
  '    "busorder_id": "9D00000101",'+
  '    "bustransaction_id": "3L00000101",'+
  '    "text": "podle smlouvy  2019020777  slu�bu  5AA-Optimal",'+
  '    "vatrate_id": "02100X0000",'+
  '    "rowtype": 1,'+
  '    "incometype_id": "2000000000",'+
  '  }'+
  ']'+
  '}';

  memo1.Lines.Add(Json);
  sResponse := DesU.abraBoCreate_So(SO(Json), editUrl.Text);
  memo2.Lines.Add (SO(sResponse).AsJSon(True));
  memo2.Lines.Add(sResponse);
end;


procedure TForm1.btnPutClick(Sender: TObject);
var
  sResponse: string;
  Json: string;
  JsonSO: ISuperObject;
begin

  //editUrl.Text := 'bankstatements/36E2000101/rows/5BRD000101';

  JsonSO := SO;
  JsonSO.S['varsymbol'] := '1116378';
  Json := JsonSO.AsJSon(true);
  memo1.Lines.Add(Json);
  sResponse := DesU.abraBoUpdate_So(SO(Json), 'bankstatement', '36E2000101', 'row', '5BRD000101');
  memo2.Lines.Add (SO(sResponse).AsJSon(true));
end;


//////////////////////////
// *** AArray demos *** //

procedure TForm1.btnUpdateByAaClick(Sender: TObject);
var
  boAA: TAArray;
  sResponse, newid: String;
begin

  boAA := TAArray.Create;
  boAA['varsymbol'] := '77989';
  sResponse := DesU.abraBoUpdate(boAA, 'bankstatement', '36E2000101', 'row', '5BRD000101'); //pou�ije PUT (pokud je metoda WebApi)

  memo2.Lines.Add (boAA.AsJSon());
  memo2.Lines.Add ('-------');
  // memo2.Lines.Add (sResponse);
  //memo2.Lines.Add ('-------');
  memo2.Lines.Add (SO(sResponse).AsJSon(true));
end;

procedure TForm1.btnCreateByAAClick(Sender: TObject);
var
  FruitColors: TAArray;
  boAA, boRowAA: TAArray;
  TestString, newid: String;
  i: integer;
begin

  boAA := TAArray.Create;

  boAA['docqueue_id'] := DesU.getAbraDocqueueId('FO1', '03');
  boAA['period_id'] := DesU.getAbraPeriodId('2017');
  boAA['docdate$date'] := '2017-06-28';
  boAA['accdate$date'] := '2017-06-30';
  boAA['firm_id'] := '2SZ1000101';
  boAA['varsymbol'] := '911123335';
  boAA['description'] := 'pppp�ipojen� rychlou�k� b��';
  boAA['priceswithvat'] := true;

  boRowAA := boAA.addRow();
  boRowAA['rowtype'] := 0;
  boRowAA['text'] := 'prvni radek fakturujeme';
  boRowAA['division_id'] := '1000000101';

  boRowAA := boAA.addRow();
  boRowAA['rowtype'] := 1;
  boRowAA['text'] := 'Za na�e sv�l� Slu�by';
  boRowAA['totalprice'] := 246;
  boRowAA['vatrate_id'] := DesU.getAbraVatrateId('V�st21');;
  boRowAA['incometype_id'] := DesU.getAbraIncometypeId('SL');
  boRowAA['division_id'] := '1000000101';

  newid := DesU.abraBoCreate(boAA, 'issuedinvoice'); //pou�ije POST (pokud je metoda WebApi)
  //newid := DesU.abraBoCreateOLE(boAA, 'issuedinvoice');
  //newid := DesU.abraBoCreateWebApi(boAA, 'issuedinvoice');

  memo2.Lines.Add(newid);
  {
  FruitColors := TAArray.Create;
  FruitColors['Apple'] := 'Green';
  FruitColors['Peach'] := 'bbla'; //TRUE;
  FruitColors['Some Fruit'] := 'ctrnact'; //14;

  memo2.Lines.Add('Position=' + inttostr(FruitColors.Position) + ' Count=' + inttostr(FruitColors.Count) );

  for i := 0 to FruitColors.Count - 1 do
    memo2.Lines.Add(FruitColors.Values[i]);
  }
  {
  if FruitColors['Peach'] then
    memo2.Lines.Add('ttttttttttt')
  else
      memo2.Lines.Add('ffffffffff');
  while FruitColors.Foreach do
  begin
    TestString := FruitColors.Value[FruitColors.Position];
    memo2.Lines.Add(FruitColors[FruitColors.Position]);
  end;
  }

end;




procedure TForm1.btnSendEmailClick(Sender: TObject);
begin
   Memo3.Clear;
   Memo3.Text := '<b>Takhle</b> chod� m�j �lu�ou�k� ��� mail��ek';

   //setup SMTP
   IdSMTP1.Host := 'localhost';
   IdSMTP1.Port := 25;

   //setup mail message
   IdMessage1.From.Address := 'joeiii@honza.cz';
   IdMessage1.Recipients.EMailAddresses := 'bazen@bulharsko.sko';

   IdMessage1.Subject := '�lu�ou�k� ��� mail��ek a5';
   //IdMessage1.Body.Text := 'Takhle chod� m�j testovac� e-mail��ek'; //nemel by se pouzit, radsi jako part

   { tahle je to Indy9 asi
   if FileExists('c:\develop\delphi.txt') then
    TIdAttachment.Create(MailMessage.MessageParts, 'c:\develop\delphi.txt') ;
   }

    //tohle by melo
    //AnsiToUtf8
    with TIdText.Create(IdMessage1.MessageParts, nil) do begin
      Body.Text := Memo3.Text;
      //Body.Text := AnsiToUtf8(Memo3.Text);
      //ContentType := 'text/html';
      ContentType := 'text/plain';
      //Charset := 'Windows-1250';
      Charset := 'utf-8';
    end;


    with TIdAttachmentFile.Create(IdMessage1.MessageParts, 'c:\develop\p.jpg') do begin
      //ContentID := '12345';
      ContentType := 'image/jpeg';
      FileName := 'p.jpg';
    end;

    with TIdAttachmentFile.Create(IdMessage1.MessageParts, 'c:\develop\sella.jpg') do begin
      //ContentID := '56789';
      ContentType := 'image/jpeg';
      FileName := 'sella.jpg';
    end;

    //IdMessage1.ContentType := 'multipart/related; type="text/html"';
    IdMessage1.ContentType := 'multipart/mixed';

   //send mail
   try
     try
       IdSMTP1.Connect() ;
       IdSMTP1.Send(IdMessage1) ;
     except on E:Exception do
       Memo2.Lines.Insert(0, 'ERROR: ' + E.Message) ;
     end;
   finally
     if IdSMTP1.Connected then IdSMTP1.Disconnect;
   end;

end;


end.
