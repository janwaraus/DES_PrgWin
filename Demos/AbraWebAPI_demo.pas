unit AbraWebAPI_demo;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls,
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient,
  IdSMTP, IdMessage, IdMessageParts, IdAttachment, IdEMailAddress, IdAttachmentFile, IdText,
  IdExplicitTLSClientServerBase, IdMessageClient, IdSMTPBase,
  System.NetEncoding, System.RegularExpressions, System.Character, System.JSON
  ;

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
    btnSendSms: TButton;
    Button1: TButton;
    procedure btnGetClick(Sender: TObject);
    procedure btnPostClick(Sender: TObject);
    procedure btnPutClick(Sender: TObject);
    procedure btnCreateByAAClick(Sender: TObject);
    procedure btnUpdateByAaClick(Sender: TObject);
    procedure btnSendEmailClick(Sender: TObject);
    procedure btnSendSmsClick(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;


function DecodeUnicodeEscapeSequences(const input: string): string;


var
  Form1: TForm1;


implementation

{$R *.dfm}

uses DesUtils, Superobject, AArray, AbraEntities;



////////////////////////////////////
// *** SuperObject (SO) demos *** //

function DecodeUnicodeEscapeSequences(const input: string): string;
var
  regex: TRegEx;
  match: TMatch;
  startIndex: Integer;
  codePoint: Integer;
  replaceChar: Char;
begin
  regex := TRegEx.Create('\\u([0-9a-fA-F]{4})');
  match := regex.Match(input);
  startIndex := 1;
  Result := '';

  while match.Success do
  begin
    Result := Result + Copy(input, startIndex, match.Index - startIndex);

    codePoint := StrToInt('$' + match.Groups[1].Value);
    replaceChar := WideChar(codePoint);
    Result := Result + replaceChar;

    startIndex := match.Index + match.Length;
    match := match.NextMatch;
  end;

  Result := Result + Copy(input, startIndex, Length(input));
end;

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
  abraResponse : TDesResult;
  newID: string;
  abraResponseSO : ISuperObject;
begin

editUrl.Text := 'issuedinvoice';

Json := '{'+
  '"docqueue_id": "L000000101",'+
  '"period_id": "2390000101",'+
  '"docdate$date": "2023-06-03",'+
  '"firm_id": "2SZ1000101",'+
  '"description": "pppp¯ipojenÌ, 2018020612",'+
  '"varsymbol": "2014020777",'+
  '"priceswithvat": true,'+

  '"rows": ['+
  '  {'+
  '    "totalprice": 3881,'+
  '    "division_id": "1000000101",'+
  '    "busorder_id": "9D00000101",'+
  '    "bustransaction_id": "3L00000101",'+
  '    "text": "pppodle smlouvy 2019020777 sluûbu  5AA-Optimal",'+
  '    "vatrate_id": "02100X0000",'+
  '    "rowtype": 1,'+
  '    "incometype_id": "2000000000",'+
  '  }'+
  ']'+
  '}';

  memo1.Lines.Add(Json);

  abraResponse := DesU.abraBoCreate_SoWebApi(SO(Json), 'issuedinvoice');
  if abraResponse.isOk then begin
    memo2.Lines.Add(abraResponse.Messg);
    abraResponseSO := SO(abraResponse.Messg);
    newId := abraResponseSO.S['id'];
    memo2.Lines.Add('------x------');
    memo2.Lines.Add(abraResponseSO.AsJSon(true));
    memo2.Lines.Add(DecodeUnicodeEscapeSequences(abraResponseSO.AsJSon(true)));
    memo2.Lines.Add(abraResponseSO.S['tradetypedescription']);
    memo2.Lines.Add(abraResponseSO.S['description']);
    memo2.Lines.Add(DecodeUnicodeEscapeSequences(abraResponseSO.S['tradetypedescription']));
    memo2.Lines.Add(DecodeUnicodeEscapeSequences(abraResponseSO.S['description']));
  end else begin
    memo2.Lines.Add(abraResponse.Code);
    memo2.Lines.Add(abraResponse.Messg);
    memo2.Lines.Add(UTF8Encode(abraResponse.Messg));

  end;




  {
  sResponse := DesU.abraBoCreate_So(SO(Json), editUrl.Text);
  memo2.Lines.Add (SO(sResponse).AsJSon(True));
  memo2.Lines.Add(sResponse);
  }
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
  sResponse := DesU.abraBoUpdate_SoWebApi(SO(Json), 'bankstatement', '36E2000101', 'row', '5BRD000101');
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
  sResponse := DesU.abraBoUpdate(boAA, 'bankstatement', '36E2000101', 'row', '5BRD000101'); //pouûije PUT (pokud je metoda WebApi)

  memo2.Lines.Add (boAA.AsJSon());
  memo2.Lines.Add ('-------');
  // memo2.Lines.Add (sResponse);
  //memo2.Lines.Add ('-------');
  memo2.Lines.Add (SO(sResponse).AsJSon(true));
end;

procedure TForm1.Button1Click(Sender: TObject);

var
  apiResponse: string;
  translatedString: string;
  CodePoint : integer;
begin
  apiResponse := 'V\u00E1m za obdob\u00ED slu\u017Ebu  5G'; // Vaöe odpovÏÔ z API
  memo2.Lines.Add(apiResponse); // ZobrazÌ "V·m"

  // DekÛdov·nÌ ¯etÏzce
  translatedString := UTF8ToString(apiResponse);

  memo2.Lines.Add(translatedString); // ZobrazÌ "V·m"


  translatedString := TNetEncoding.URL.Decode(apiResponse);
  memo2.Lines.Add(translatedString); // ZobrazÌ "V·m"


  translatedString := DecodeUnicodeEscapeSequences(apiResponse);
  memo2.Lines.Add(translatedString); // ZobrazÌ "V·m"

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
  boAA['description'] := 'pppp¯ipojenÌ rychlouËkÈ bÏûÌ';
  boAA['priceswithvat'] := true;

  boRowAA := boAA.addRow();
  boRowAA['rowtype'] := 0;
  boRowAA['text'] := 'prvni radek fakturujeme';
  boRowAA['division_id'] := '1000000101';

  boRowAA := boAA.addRow();
  boRowAA['rowtype'] := 1;
  boRowAA['text'] := 'Za naöe svÏlÈ Sluûby';
  boRowAA['totalprice'] := 246;
  boRowAA['vatrate_id'] := DesU.getAbraVatrateId('V˝st21');
  boRowAA['incometype_id'] := AbraEnt.getIncomeType('Code=SL').ID;
  boRowAA['division_id'] := '1000000101';

  newid := DesU.abraBoCreate(boAA, 'issuedinvoice'); //pouûije POST (pokud je metoda WebApi)
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
   Memo3.Text := '<b>Takhle</b> chodÌ m˘j éluùouËk˝ ûË¯ mailÌËek';

   //setup SMTP
   IdSMTP1.Host := 'localhost';
   IdSMTP1.Port := 25;

   //setup mail message
   IdMessage1.From.Address := 'joeiii@honza.cz';
   IdMessage1.Recipients.EMailAddresses := 'bazen@bulharsko.sko';

   IdMessage1.Subject := 'éluùouËk˝ ûË¯ mailÌËek a5';
   //IdMessage1.Body.Text := 'Takhle chodÌ m˘j testovacÌ e-mailÌËek'; //nemel by se pouzit, radsi jako part

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


procedure TForm1.btnSendSmsClick(Sender: TObject);
var
  callResult : string;
begin
  callResult := DesU.sendGodsSms('603260797', 'auto prg vol·nÌ GODS br·na');
  Memo3.Lines.Add(callResult);
end;

end.
