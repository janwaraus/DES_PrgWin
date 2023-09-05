unit AbraWebAPI_demo;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls,
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient,
  IdSMTP, IdMessage, IdMessageParts, IdAttachment, IdEMailAddress, IdAttachmentFile, IdText,
  IdExplicitTLSClientServerBase, IdMessageClient, IdSMTPBase,
  IdHTTP,
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
    function String2Hex(const Buffer: AnsiString): string;
  end;


function DecodeUnicodeEscapeSequences(const input: string): string;
procedure createBoByAA_v1;
procedure createBoByAA_v2;
procedure createBoByAA_v3;


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
  abraResponseSO: ISuperObject;
  mySuperArray: TSuperAvlEntry;
  item: ISuperObject;
  idHTTP: TIdHTTP;
  endpoint : string;
  UrlContent: TStringStream; // You can also use TMemoryStream here
begin
  //editUrl.Text := 'periods?where=code+gt+2015';
  editUrl.Text := 'bankstatements/36E2000101/rows/ZHRD000101';
  sResponse := DesU.abraBoGet(editUrl.Text); // misto toho dame cely kod sem


  memo2.Lines.Clear;
  memo2.Lines.Add(sResponse);

  abraResponseSO := SO(sResponse);

  memo2.Lines.Add(abraResponseSO.S['text']);
  memo2.Lines.Add('- UTF8Decode resp- -');
  memo2.Lines.Add(UTF8Decode(abraResponseSO.S['text']));
  memo2.Lines.Add('-HEX-');
  memo2.Lines.Add(String2Hex(abraResponseSO.S['text']));

  {
  memo2.Lines.Add('- abraResponseSO.AsJSon(true) -');
  memo2.Lines.Add (abraResponseSO.AsJSon(True));
  memo2.Lines.Add('- UTF8Decode(abraResponseSO.AsJSon(true)) -');
  memo2.Lines.Add(UTF8Decode(abraResponseSO.AsJSon(true)));
  }

  {
  memo2.Lines.Add (MyObject.AsArray[0].AsString);
  memo2.Lines.Add (MyObject.AsArray[0].AsObject.S['name']);

  for item in MyObject do
    memo2.Lines.Add (item.AsObject.S['name']);
  }

  //hodnota := MyObject.S['last'];
  //hodnota := SO(Mydata).S['last'];  //zkracena verze
end;

procedure TForm1.btnPostClick(Sender: TObject);
var
  sResponse: string;
  Json: string;
  abraResponse : TDesResult;
  newID: string;
  jsonSO: ISuperObject;
  abraResponseSO : ISuperObject;
begin

editUrl.Text := 'issuedinvoice';

Json := '{'+
  '"docqueue_id": "L000000101",'+
  '"period_id": "2390000101",'+
  '"docdate$date": "2023-06-03",'+
  '"firm_id": "2SZ1000101",'+
  '"description": "p¯Ìliö ûluùouËk˝ k˘Ú, 123456",'+
  '"varsymbol": "2014020777",'+
  '"priceswithvat": true,'+

  '"rows": ['+
  '  {'+
  '    "totalprice": 10001,'+
  '    "division_id": "1000000101",'+
  '    "busorder_id": "9D00000101",'+
  '    "bustransaction_id": "3L00000101",'+
  '    "text": "PÿÕLIä éLUçOU»K› KŸ“, 7890",'+
  '    "vatrate_id": "02100X0000",'+
  '    "rowtype": 1,'+
  '    "incometype_id": "2000000000",'+
  '  }'+
  ']'+
  '}';

  memo1.Lines.Add(Json);

  abraResponse := DesU.abraBoCreate(SO(Json).AsJSon(true), 'issuedinvoice');
  if abraResponse.isOk then begin
    memo2.Lines.Add(abraResponse.Messg);
    memo2.Lines.Add('- UTF8Decode resp- -');
    memo2.Lines.Add(UTF8Decode(abraResponse.Messg));

    abraResponseSO := SO(abraResponse.Messg);
    newId := abraResponseSO.S['id'];

    memo2.Lines.Add('- abraResponseSO.AsJSon(true) -');
    memo2.Lines.Add(abraResponseSO.AsJSon(true));
    memo2.Lines.Add('- UTF8Decode(abraResponseSO.AsJSon(true)) -');
    memo2.Lines.Add(UTF8Decode(abraResponseSO.AsJSon(true)));
    memo2.Lines.Add('- DecodeUnicodeEscapeSequences -');
    memo2.Lines.Add(DecodeUnicodeEscapeSequences(abraResponseSO.AsJSon(true)));
    memo2.Lines.Add('------------');
    memo2.Lines.Add(abraResponseSO.S['tradetypedescription']);
    memo2.Lines.Add(abraResponseSO.S['description']);
    memo2.Lines.Add(DecodeUnicodeEscapeSequences(abraResponseSO.S['tradetypedescription']));
    memo2.Lines.Add(UTF8Encode(abraResponseSO.S['tradetypedescription']));
    memo2.Lines.Add(UTF8Decode(abraResponseSO.S['tradetypedescription']));
  end else begin
    memo2.Lines.Add(abraResponse.Code);
    memo2.Lines.Add('-----Messg-------');
    memo2.Lines.Add(abraResponse.Messg);
    memo2.Lines.Add('-UTF8Decode-----------');
    memo2.Lines.Add(UTF8Decode(abraResponse.Messg));
    memo2.Lines.Add('-UTF8Encode-----------');
    memo2.Lines.Add(UTF8Encode(abraResponse.Messg));
    memo2.Lines.Add('-HEX-');
    memo2.Lines.Add(String2Hex(abraResponse.Messg));

  end;




  {
  sResponse := DesU.abraBoCreate_So(SO(Json), editUrl.Text);
  memo2.Lines.Add (SO(sResponse).AsJSon(True));
  memo2.Lines.Add(sResponse);
  }
end;

function TForm1.String2Hex(const Buffer: AnsiString): string;
begin
  SetLength(Result, Length(Buffer) * 2);
  BinToHex(@Buffer[1], PWideChar(@Result[1]), Length(Buffer));
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
  sResponse := DesU.abraBoUpdate(JsonSO.AsJSon(true), 'bankstatement', '36E2000101', 'row', '5BRD000101');
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
  sResponse := DesU.abraBoUpdate(boAA.AsJSon, 'bankstatement', '36E2000101', 'row', '5BRD000101'); //pouûije PUT

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
begin
  createBoByAA_v3;
end;


procedure createBoByAA_v1;
var
  FruitColors: TAArray;
  boAA, boRowAA: TAArray;
  TestString, newid: string;
  i: integer;
  abraWebApiResponse : TDesResult;
begin

  boAA := TAArray.Create;

  boAA['docqueue_id'] := DesU.getAbraDocqueueId('FO1', '03');
  boAA['period_id'] := DesU.getAbraPeriodId('2023');
  boAA['docDate$date'] := '2023-06-03';
  boAA['Firm_ID'] := '2SZ1000101';
  boAA['VarSymbol'] := '911123335';
  boAA['description'] := 'pppp¯ipojenÌ rychlouËkÈ bÏûÌ';
  boAA['priceSWithvat'] := true;
  boAA['varsymbol'] := '911123336';

  boRowAA := boAA.addRow();
  boRowAA['rowtype'] := 0;
  boRowAA['TEXT'] := 'prvni radek fakturujeme';
  boRowAA['division_id'] := '1000000101';

  boRowAA := boAA.addRow();
  boRowAA['rowtype'] := 1;
  boRowAA['text'] := 'Za naöe svÏlÈ Sluûby';
  boRowAA['TOTALPrice'] := 246;
  boRowAA['vatrate_id'] := DesU.getAbraVatrateId('V˝st21');
  boRowAA['incometype_id'] := AbraEnt.getIncomeType('Code=SL').ID;
  boRowAA['division_id'] := '1000000101';

  abraWebApiResponse := DesU.abraBoCreate(boAA.AsJSon, 'issuedinvoice');
  newid := SO(abraWebApiResponse.Messg).S['id'];

  Form1.memo2.Lines.Add(newid);
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

procedure createBoByAA_v2;
var
  newBo: TNewAbraBo;
  TestString, newid: string;
begin

  newBo := TNewAbraBo.Create('issuedinvoice');

  newBo.Item['docqueue_id'] := DesU.getAbraDocqueueId('FO1', '03');
  newBo.Item['period_id'] := DesU.getAbraPeriodId('2023');
  newBo.Item['docDate$date'] := '2023-06-03';
  newBo.Item['Firm_ID'] := '2SZ1000101';
  newBo.Item['VarSymbol'] := '911123332';
  newBo.Item['description'] := 'pppp¯ipojenÌ rychlouËkÈ bÏûÌ';
  newBo.Item['priceSWithvat'] := true;
  newBo.Item['varsymbol'] := '911123335';

  newBo.createNewRow;
  newBo.rowItem['rowtype'] := 0;
  newBo.rowItem['TEXT'] := 'PPPrvni ¯Ëû·dek fakturujeme';
  newBo.rowItem['division_id'] := '1000000101';

  newBo.createNewRow;
  newBo.rowItem['rowtype'] := 1;
  newBo.rowItem['text'] := 'ZZa naöe svÏlÈ Sluûby';
  newBo.rowItem['TOTALPrice'] := 246;
  newBo.rowItem['vatrate_id'] := DesU.getAbraVatrateId('V˝st21');
  newBo.rowItem['incometype_id'] := AbraEnt.getIncomeType('Code=SL').ID;
  newBo.rowItem['division_id'] := '1000000101';
  newBo.rowItem['division_id'] := '1000000101';

  newBo.writeToAbra;


  if newBo.WriteResult.isOk then begin
    newid := SO(newBo.WriteResult.Messg).S['id'];
    Form1.memo2.Lines.Add(newid);
    Form1.memo2.Lines.Add(newBo.getCreatedBoItem('id'));
    Form1.memo2.Lines.Add(SO(newBo.WriteResult.Messg).S['displayname']);
    Form1.memo2.Lines.Add(newBo.getCreatedBoItem('displayname'));
  end;


end;

procedure createBoByAA_v3;
var
  newBo: TNewAbraBo;
  TestString, newid: string;
begin

  newBo := TNewAbraBo.Create('issuedinvoice');
  newBo.addInvoiceParams(45130);

  newBo.Item['docqueue_id'] := DesU.getAbraDocqueueId('FO1', '03');
  newBo.Item['Firm_ID'] := '2SZ1000101';
  newBo.Item['VarSymbol'] := '911123335';
  newBo.Item['description'] := 'NovÈe p¯ipojenÌ rychlouËkÈ bÏûÌ';

  newBo.createNewInvoiceRow(0, 'PPPrvni ¯Ëû·dek fakturujeme');

  newBo.createNewInvoiceRow(1, 'ZZa naöe svÏlÈ Sluûby');
  newBo.rowItem['TOTALPrice'] := 648;

  newBo.writeToAbra;


  if newBo.WriteResult.isOk then begin
    newid := SO(newBo.WriteResult.Messg).S['id'];
    Form1.memo2.Lines.Add(newid);
    Form1.memo2.Lines.Add(newBo.getCreatedBoItem('id'));
    Form1.memo2.Lines.Add(SO(newBo.WriteResult.Messg).S['displayname']);
    Form1.memo2.Lines.Add(newBo.getCreatedBoItem('displayname'));
  end;


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
