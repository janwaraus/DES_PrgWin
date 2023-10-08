unit DesUtils;

interface

uses
  Winapi.Windows, Winapi.ShellApi, Winapi.Messages,
  System.SysUtils, System.Variants, System.Classes, System.RegularExpressions,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Math, StrUtils, IOUtils, IniFiles, ComObj,

  IdIOHandler, IdIOHandlerSocket, IdIOHandlerStack, IdSSL, IdSSLOpenSSL, IdBaseComponent, IdComponent,
  IdTCPConnection, IdTCPClient, IdHTTP, IdSMTP, IdExplicitTLSClientServerBase, IdSMTPBase,
  IdText, IdMessage, IdMessageParts, IdMessageClient, IdAttachmentFile,
  Data.DB, ZAbstractRODataset, ZAbstractDataset, ZDataset, ZAbstractConnection, ZConnection,
  Superobject, AArray;


type
  TDesResult = class
  public
    Code,
    Messg : string;
    constructor create(Code, Messg : string);
    function isOk : boolean;
    function isErr : boolean;
  end;

  TDesU = class(TForm)
    dbAbra: TZConnection;
    qrAbra: TZQuery;
    qrAbra2: TZQuery;
    qrAbra3: TZQuery;
    qrAbraOC: TZQuery;
    qrAbraOC2: TZQuery;
    dbZakos: TZConnection;
    qrZakos: TZQuery;
    qrZakosOC: TZQuery;
    idMessage: TIdMessage;
    idSMTP: TIdSMTP;
    IdSSLIOHandler: TIdSSLIOHandlerSocketOpenSSL;
    IdHTTPAbra: TIdHTTP;

    procedure FormCreate(Sender: TObject);
    procedure desUtilsInit(createOptions : string = '');
    procedure programInit(programName : string);
    function getIniValue(iniGroup, iniItem : string) : string;
    function zalogujZpravu(TextZpravy: string): string;

    function odstranNulyZCislaUctu(cisloU : string) : string;
    function prevedCisloUctuNaText(cisloU : string) : string;
    function opravRadekVypisuPomociPDocument_ID(Vypis_ID, RadekVypisu_ID, PDocument_ID, PDocumentType : string) : string;
    procedure opravRadekVypisuPomociVS(Vypis_ID, RadekVypisu_ID, VS : string);
    function vytvorFaZaInternetKredit(VS : string; castka : currency; datum : double) : string;
    function vytvorFaZaVoipKredit(VS : string; castka : currency; datum : double) : string;
    function zrusPenizeNaCeste(VS : string) : string;
    function pdfToCustomerSendingDateTime(customerId : integer; pdfFileName : string) : double;
    function ulozKomunikaci(Typ, Customer_id, Zprava: string): TDesResult;   // typ 2 je mail, typ 23 SMS
    function posliPdfEmailem(FullPdfFileName, emailAddrStr, emailPredmet, emailZprava, emailOdesilatel : string; ExtraPrilohaFileName: string = '') : TDesResult;
    procedure syncAbraPdfToRemoteServer(year, month : integer);


    public
      PROGRAM_NAME,
      PROGRAM_PATH,
      LOG_PATH,
      LOG_FILENAME,
      GPC_PATH,
      PDF_PATH,
      VAT_RATE,
      abraDefaultCommMethod,
      abraConnName,
      abraUserUN,
      abraUserPW,
      abraWebApiUrl : string;
      appMode: integer;
      AbraOLE: variant;
      VAT_MULTIPLIER: double;
      adpIniFile: TIniFile;

      //abraValues : TAArray;




      { ABRA WebAPI funkce }
      function abraBoGet(abraBoName : string; sId: string = '') : string;

      function abraBoCreate(jsonString, abraBoName : string) : TDesResult;
      function abraBoCreateRow(jsonString, abraBoName, parent_id : string) : TDesResult;
      function abraBoUpdate(jsonString, abraBoName, abraBoId : string; abraBoChildName: string = ''; abraBoChildId: string = '') : string;

      procedure logJson(jsonToLog, headerToLog : string);

      { ABRA zapisující funkce pomocí SuperObjectu, používáno }
      function abraBoCreateWebApi_v0(boAA: TAArray; abraBoName : string) : string; // to be deleted

      { ABRA OLE funkce, mìly by se pøestat používat, mohou být však užiteèné pro nìjaké testování napø. v budoucnu }
      function getAbraOLE() : variant;
      procedure abraOLELogout();
      function getOleObjDataDisplay(abraOleObj_Data : variant) : ansistring;
      function abraBoCreateOLE(boAA: TAArray; abraBoName : string) : string;
      function abraBoUpdateOLE(boAA: TAArray; abraBoName, abraBoId : string; abraBoChildName: string = ''; abraBoChildId: string = '') : string;


      { ABRA rùzné funkce }
      function getAbraPeriodId(pYear : string) : string; overload;
      function getAbraPeriodId(pDate : double) : string; overload;
      function getAbraDocqueueId(code, documentType : string) : string;
      function getAbraDocqueueCodeById(id : string) : string;
      function getAbraVatrateId(code : string) : string;
      function getAbraVatindexId(code : string) : string;

      function getAbracodeByVs(vs : string) : string;
      function getAbracodeByContractNumber(cnumber : string) : string;
      function getFirmIdByCode(code : string) : string;
      function getZustatekByAccountId (accountId : string; datum : double) : double;
      function existujeVAbreDokladSPrazdnymVs() : boolean;

      function isVoipKreditContract(cnumber : string) : boolean;
      function isCreditContract(cnumber : string) : boolean;

      function sendGodsSms(telCislo, smsText : string) : string;
      function sendSms(telCislo, smsText : string) : string;


  end;


function odstranDiakritiku(textDiakritika : string) : string;
function removeLeadingZeros(const Value: string): string;
function LeftPad(value:integer; length:integer=8; pad:char='0'): string; overload;
function LeftPad(value: string; length:integer=8; pad:char='0'): string; overload;
function Str6digitsToDate(datum : string) : double;
function IndexByName(DataObject: variant; Name: ShortString): integer;
function pocetRadkuTxtSouboru(SName: string): integer;
function RemoveSpaces(const s: string): string;
function KeepOnlyNumbers(const S: string): string;
function SplitStringInTwo(const InputString, Delimiter: string; out LeftPart, RightPart: string): Boolean;
function destilujTelCislo(telCislo: string): string;
function destilujMobilCislo(telCislo: string): string;
function FindInFolder(sFolder, sFile: string; bUseSubfolders: Boolean): string;
procedure writeToFile(pFileName, pContent : string);
procedure appendToFile(pFileName, pContent : string);
function LoadFileToStr(const FileName: TFileName): ansistring;
function FloatToStrFD (pFloat : extended) : string;
function RandString(const stringsize: integer): string;
function debugRozdilCasu(cas01, cas02 : double; textZpravy : string) : string;
procedure RunCMD(cmdLine: string; WindowMode: integer);
function getWindowsUserName : string;
function getWindowsCompName : string;


const
  Ap = chr(39);
  ApC = Ap + ',';
  ApZ = Ap + ')';


var
  DesU : TDesU;


implementation

{$R *.dfm}

uses AbraEntities;

{****************************************************************************}
{**********************     ABRA common functions     ***********************}
{****************************************************************************}


procedure TDesU.FormCreate(Sender: TObject);
begin
  desUtilsInit();
end;

procedure TDesU.desUtilsInit(createOptions : string);

begin

  VAT_RATE := '21'; // globání nastavení sazby DPH!
  VAT_MULTIPLIER := (100 + StrToInt(VAT_RATE))/100;

  PROGRAM_PATH := ExtractFilePath(ParamStr(0));

  if not(FileExists(PROGRAM_PATH + 'abraDesProgramy.ini'))
    AND not(FileExists(PROGRAM_PATH + '..\DE$_Common\abraDesProgramy.ini')) then
  begin
    Application.MessageBox(PChar('Nenalezen soubor abraDesProgramy.ini, program ukonèen'),
      'abraDesProgramy.ini', MB_OK + MB_ICONERROR);
    Application.Terminate;
  end;

  if FileExists(PROGRAM_PATH + 'abraDesProgramy.ini') then
    adpIniFile := TIniFile.Create(PROGRAM_PATH + 'abraDesProgramy.ini')
  else
    adpIniFile := TIniFile.Create(PROGRAM_PATH + '..\DE$_Common\abraDesProgramy.ini');

  with adpIniFile do try
    appMode := StrToInt(ReadString('Preferences', 'AppMode', '1'));
    abraDefaultCommMethod := ReadString('Preferences', 'AbraDefaultCommMethod', 'OLE');
    abraConnName := ReadString('Preferences', 'abraConnName', '');
    abraUserUN := ReadString('Preferences', 'AbraUserUN', '');
    abraUserPW := ReadString('Preferences', 'AbraUserPW', '');
    abraWebApiUrl := ReadString('Preferences', 'AbraWebApiUrl', '');
    IdHTTPAbra.Request.Username := abraUserUN;
    IdHTTPAbra.Request.Password := abraUserPW;
    IdHTTPAbra.ReadTimeout := Round (600 * 1000); // ReadTimeout je v milisekundách

    GPC_PATH := IncludeTrailingPathDelimiter(ReadString('Preferences', 'GpcPath', ''));
    PDF_PATH := IncludeTrailingPathDelimiter(ReadString('Preferences', 'PDFDir', ''));

    dbAbra.HostName := ReadString('Preferences', 'AbraHN', '');
    dbAbra.Database := ReadString('Preferences', 'AbraDB', '');
    dbAbra.User := ReadString('Preferences', 'AbraUN', '');
    dbAbra.Password := ReadString('Preferences', 'AbraPW', '');

    dbZakos.HostName := ReadString('Preferences', 'ZakHN', '');
    dbZakos.Database := ReadString('Preferences', 'ZakDB', '');
    dbZakos.User := ReadString('Preferences', 'ZakUN', '');
    dbZakos.Password := ReadString('Preferences', 'ZakPW', '');

    idSMTP.Host :=  ReadString('Mail', 'SMTPServer', '');
    idSMTP.HeloName := idSMTP.Host; // táta to tak má, ale HeloName by mìlo být jméno klienta (volajícího), tedy tohoto programu. ale asi je úplnì jedno, co tam je
    idSMTP.Username := ReadString('Mail', 'SMTPLogin', '');
    idSMTP.Password := ReadString('Mail', 'SMTPPW', '');
  finally
    adpIniFile.Free;
  end;

  if not dbAbra.Connected then try
    dbAbra.Connect;
  except on E: exception do
    begin
      Application.MessageBox(PChar('Nedá se pøipojit k databázi Abry, program ukonèen.' + ^M + E.Message), 'DesU DB Abra', MB_ICONERROR + MB_OK);
      Application.Terminate;
    end;
  end;

  if (not dbZakos.Connected) AND (dbZakos.Database <> '') then try begin
    dbZakos.Connect;
    //qrZakos.SQL.Text := 'SET CHARACTER SET utf8'; // pøesunuto do vlastností TZConnection
    //qrZakos.SQL.Text := 'SET AUTOCOMMIT = 1';       // pøesunuto do vlastností TZConnection
   // qrZakos.ExecSQL;
  end;
  except on E: exception do
    begin
      Application.MessageBox(PChar('Nedá se pøipojit k databázi smluv, program ukonèen.' + ^M + E.Message), 'DesU DB smluv', MB_ICONERROR + MB_OK);
      Application.Terminate;
    end;
  end;

  AbraEnt := TAbraEnt.Create;

end;

procedure TDesU.programInit(programName : string);
begin
  PROGRAM_NAME := programName;
  LOG_PATH := PROGRAM_PATH + 'Logy\' + PROGRAM_NAME + '\';
  if not DirectoryExists(LOG_PATH) then Forcedirectories(LOG_PATH);
  LOG_FILENAME := LOG_PATH + FormatDateTime('yyyy.mm".log"', Date);
end;

function TDesU.getIniValue(iniGroup, iniItem : string) : string;
begin
  if FileExists(PROGRAM_PATH + 'abraDesProgramy.ini') then
    adpIniFile := TIniFile.Create(PROGRAM_PATH + 'abraDesProgramy.ini')
  else
    adpIniFile := TIniFile.Create(PROGRAM_PATH + '..\DE$_Common\abraDesProgramy.ini');
  try
    Result :=  adpIniFile.ReadString(iniGroup, iniItem, '');
  finally
    adpIniFile.Free;
  end;
end;

function TDesU.zalogujZpravu(TextZpravy: string): string;
begin
  Result := FormatDateTime('dd.mm.yy hh:nn:ss  ', Now) + TextZpravy;
  DesUtils.appendToFile(LOG_FILENAME,
    Format('(%s - %s) ', [Trim(getWindowsCompName), Trim(getWindowsUserName)]) + Result);
end;



{*** ABRA WebApi functions ***}

function TDesU.abraBoGet(abraBoName : string; sId: string = '') : string;
var
  endpoint : string;
  responseContent: TStringStream;
begin
  endpoint := abraWebApiUrl + abraBoName;
  if sId <> '' then
    endpoint := endpoint + '/' + sId;

  responseContent := TStringStream.Create('', TEncoding.UTF8);

  try
    IdHTTPAbra.Get(endpoint, responseContent);
    Result := responseContent.DataString;
  except
    on E: Exception do
      ShowMessage('Error on request: '#13#10 + e.Message);
  end;

end;


function TDesU.abraBoCreate(jsonString, abraBoName : string) : TDesResult;
var
  sourceJsonStream,
  responseContent: TStringStream;
  newAbraBo : string;
  responseCode : integer;
  exceptionSO : ISuperObject;
begin
  sourceJsonStream := TStringStream.Create(jsonString, TEncoding.UTF8);
  responseContent := TStringStream.Create('', TEncoding.UTF8);

  try
    try begin
      IdHTTPAbra.Post(abraWebApiUrl + abraBoName + 's', sourceJsonStream, responseContent);
      responseCode := IdHTTPAbra.ResponseCode;

      if responseCode < 400 then begin // HTTP responsy pod 400 nejsou errorové
        newAbraBo := responseContent.DataString;
        self.logJson(jsonString + sLineBreak + sLineBreak + 'response:' + sLineBreak + newAbraBo,
          'abraBoCreate - WebApi - ' + abraWebApiUrl + abraBoName + 's');
        Result := TDesResult.create('ok', newAbraBo);
      end
      else
      begin
        self.logJson(jsonString + sLineBreak + sLineBreak + 'error response HTTP '
          + IntToStr(responseCode) + sLineBreak + responseContent.DataString,
          'abraBoCreate - WebApi - ' + abraWebApiUrl + abraBoName + 's');
        exceptionSO := SO(responseContent.DataString);
        Result := TDesResult.create('err_http' + IntToStr(IdHTTPAbra.ResponseCode),
          exceptionSO.S['title']
          + ': ' + exceptionSO.S['description']
        );
      end;
    end;
    except
      on E: EIdHTTPProtocolException do begin
        { * takto bylo døíve, ale kvùli nefunènosti èeských znakù v E.Message pøesunuto mimo except blok *

          exceptionSO := SO(E.errorMessage);
          Result := TDesResult.create('err_http' + IntToStr(IdHTTPAbra.ResponseCode),
            exceptionSO.S['title']
            + ': ' + exceptionSO.S['description']
          );
          self.logJson(E.errorMessage, 'abraBoCreateWebApi_SO response - ' + abraWebApiUrl + abraBoName + 's');

        }
      end;
      on E: Exception do begin
        ShowMessage('Error on request: '#13#10 + E.Message);
      end;
    end;
  finally
    sourceJsonStream.Free;
    responseContent.Free;
  end;
end;


//ABRA BO create row
function TDesU.abraBoCreateRow(jsonString, abraBoName, parent_id : string) : TDesResult;
var
  sourceJsonStream,
  responseContent: TStringStream;
  responseCode : integer;
  exceptionSO : ISuperObject;
  sstreamJson: TStringStream;
  jsonRequest, newAbraBo : string;
begin

  jsonRequest := '{"rows":[' + jsonString + ']}';

  self.logJson(jsonRequest, 'abraBoCreateRow - WebApi - ' + abraWebApiUrl + abraBoName + 's/' + parent_id);

  sourceJsonStream := TStringStream.Create(jsonRequest, TEncoding.UTF8);
  responseContent := TStringStream.Create('', TEncoding.UTF8);

  sstreamJson := TStringStream.Create(jsonRequest, TEncoding.UTF8);

  try
    try begin
      IdHTTPAbra.Put(abraWebApiUrl + abraBoName + 's/' + parent_id, sourceJsonStream, responseContent);
      responseCode := IdHTTPAbra.ResponseCode;

      if responseCode < 400 then begin // HTTP responsy pod 400 nejsou errorové
        newAbraBo := responseContent.DataString;
        self.logJson(jsonString + sLineBreak + sLineBreak + 'response:' + sLineBreak + newAbraBo,
          'abraBoCreateRow - WebApi - ' + abraWebApiUrl + abraBoName + 's');
        Result := TDesResult.create('ok', newAbraBo);
      end
      else
      begin
        self.logJson(jsonString + sLineBreak + sLineBreak + 'error response HTTP '
          + IntToStr(responseCode) + sLineBreak + responseContent.DataString,
          'abraBoCreateRow - WebApi - ' + abraWebApiUrl + abraBoName + 's');
        exceptionSO := SO(responseContent.DataString);
        Result := TDesResult.create('err_http' + IntToStr(IdHTTPAbra.ResponseCode),
          exceptionSO.S['title']
          + ': ' + exceptionSO.S['description']
        );
      end;

    end;
    except
      on E: Exception do begin
        ShowMessage('Error on request: '#13#10 + e.Message);
      end;
    end;
  finally
    sourceJsonStream.Free;
    responseContent.Free;
  end;
end;




//ABRA BO update
function TDesU.abraBoUpdate(jsonString, abraBoName, abraBoId, abraBoChildName, abraBoChildId : string) : string;
var
  idHTTP: TIdHTTP;
  sstreamJson: TStringStream;
  endpoint : string;

begin
  // http://localhost/DES/issuedinvoices/8L6U000101/rows/5A3K100101
  endpoint := abraWebApiUrl + abraBoName + 's/' + abraBoId;
  if abraBoChildName <> '' then
    endpoint := endpoint + '/' + abraBoChildName + 's/' + abraBoChildId;

  self.logJson(jsonString, 'abraBoUpdate - WebApi - ' + endpoint);

  sstreamJson := TStringStream.Create(jsonString, TEncoding.UTF8);

  try
    Result := idHTTPAbra.Put(endpoint, sstreamJson);
  finally
    sstreamJson.Free;
    idHTTP.Free;
  end;
end;



procedure TDesU.logJson(jsonToLog, headerToLog : string);
begin
  if appMode >= 3 then begin
    if not DirectoryExists(PROGRAM_PATH + '\log\json\') then
      Forcedirectories(PROGRAM_PATH + '\log\json\');
    writeToFile(PROGRAM_PATH + '\log\json\' + formatdatetime('yymmdd-hhnnss', Now) + '_' + RandString(3) + '.txt', headerToLog + sLineBreak + jsonToLog);
  end;
end;


{
// ABRA BO create
function TDesU.abraBoCreate(boAA: TAArray; abraBoName : string) : string;
begin
  if AnsiLowerCase(abraDefaultCommMethod) = 'webapi' then
    Result := self.abraBoCreateWebApi_v0(boAA, abraBoName)
  else begin
    Result := self.abraBoCreateOLE(boAA, abraBoName);
    self.abraOLELogout;
  end;
end;
}

function TDesU.abraBoCreateWebApi_v0(boAA: TAArray; abraBoName : string) : string;
var
  idHTTP: TIdHTTP;
  sstreamJson: TStringStream;
  newAbraBo : string;
begin

  self.logJson(boAA.AsJSon(), 'abraBoCreateWebApi_AA_v0 - ' + abraWebApiUrl + abraBoName + 's');

  sstreamJson := TStringStream.Create(boAA.AsJSon(), TEncoding.UTF8);
  //idHTTP := newAbraIdHttp(900, true);

  try
    try begin
      newAbraBo := idHTTP.Post(abraWebApiUrl + abraBoName + 's', sstreamJson);
      Result := SO(newAbraBo).S['id'];
    end;
    except
      on E: Exception do begin
        ShowMessage('Error on request: '#13#10 + e.Message);
      end;
    end;
  finally
    sstreamJson.Free;
    idHTTP.Free;
  end;
end;




{*** ABRA OLE functions ***}

function TDesU.getAbraOLE() : variant;
begin
  Result := null;
  if VarIsEmpty(AbraOLE) then try
    AbraOLE := CreateOLEObject('AbraOLE.Application');
    if not AbraOLE.Connect('@' + abraConnName) then begin
      ShowMessage('Problém s Abrou (connect ' + abraConnName + ').');
      Exit;
    end;
    //Zprava('Pøipojeno k Abøe (connect DES).');
    if not AbraOLE.Login(abraUserUN, abraUserPW) then begin
      ShowMessage('Problém s Abrou (login ' + abraUserUN +').');
      Exit;
    end;
    //Zprava('Pøihlášeno k Abøe (login Supervisor).');
  except on E: exception do
    begin
      Application.MessageBox(PChar('Problém s Abrou.' + ^M + E.Message), 'Abra', MB_ICONERROR + MB_OK);
      //Zprava('Problém s Abrou - ' + E.Message);
      Exit;
    end;
  end;
  Result := AbraOLE;
end;

procedure TDesU.abraOLELogout();
begin
  if VarIsEmpty(AbraOLE) then Exit;
  try
    self.AbraOLE.Logout;
  except
  end;
  self.AbraOLE := Unassigned;
end;

function TDesU.getOleObjDataDisplay(abraOleObj_Data : variant) : ansistring;
var
  j : integer;
begin
  Result := '';
  for j := 0 to abraOleObj_Data.Count - 1 do begin
    Result := Result + inttostr(j) + 'r ' + abraOleObj_Data.Names[j] + ': ' + vartostr(abraOleObj_Data.Value[j]) + sLineBreak;
  end;
end;

function TDesU.abraBoCreateOLE(boAA: TAArray; abraBoName : string) : string;
var
  i, j : integer;
  BO_Object,
  BO_Data,
  BORow_Object,
  BORow_Data,
  BO_Data_Coll,
  NewID : variant;

  iBoRowAA : TAArray;

begin

  self.logJson(boAA.AsJSon(), 'abraBoCreateOLE_AA - abraBoName=' + abraBoName);
  AbraOLE := getAbraOLE();
  BO_Object:= AbraOLE.CreateObject('@'+abraBoName);
  BO_Data:= AbraOLE.CreateValues('@'+abraBoName);
  BO_Object.PrefillValues(BO_Data);

  for i := 0 to boAA.Count - 1 do
    BO_Data.ValueByName(boAA.Keys[i]) := boAA.Values[i];

  if boAA.RowList.Count > 0 then
  begin
    BORow_Object := AbraOLE.CreateObject('@'+abraBoName+'Row');
    BO_Data_Coll := BO_Data.Value[IndexByName(BO_Data, 'Rows')];

    for i := 0 to boAA.RowList.Count - 1 do
    begin
      BORow_Data := AbraOLE.CreateValues('@'+abraBoName+'Row');
      BORow_Object.PrefillValues(BORow_Data);

      iBoRowAA := TAArray(boAA.RowList[i]);
      for j := 0 to iBoRowAA.Count - 1 do
        BORow_Data.ValueByName(iBoRowAA.Keys[j]) := iBoRowAA.Values[j];

      BO_Data_Coll.Add(BORow_Data);
    end;
  end;

  try begin
    NewID := BO_Object.CreateNewFromValues(BO_Data); //NewID je ID Abry
    Result := NewID;
  end;
  except on E: exception do
    begin
      Application.MessageBox(PChar('Problem ' + ^M + E.Message), 'AbraOLE');
      Result := 'Chyba pøi zakládání BO';
    end;
  end;

end;

function TDesU.abraBoUpdateOLE(boAA: TAArray; abraBoName, abraBoId, abraBoChildName, abraBoChildId : string) : string;
var
  i: integer;
  BO_Object,
  BO_Data,
  BORow_Object,
  BORow_Data,
  BO_Data_Coll : variant;

begin
  abraBoName := abraBoName + abraBoChildName;

  if abraBoChildName <> '' then
    abraBoId := abraBoChildId;  //budeme pracovat s ID childa

  self.logJson(boAA.AsJSon(), 'abraBoUpdateOLE_AA - abraBoName=' + abraBoName + ' abraBoId=' + abraBoId);

  AbraOLE := getAbraOLE();
  BO_Object := AbraOLE.CreateObject('@' + abraBoName);
  BO_Data := AbraOLE.CreateValues('@' + abraBoName);

  BO_Data := BO_Object.GetValues(abraBoId);

  for i := 0 to boAA.Count - 1 do
    BO_Data.ValueByName(boAA.Keys[i]) := boAA.Values[i];

  BO_Object.UpdateValues(abraBoId, BO_Data);

end;



{*** ABRA data manipulating functions ***}

function TDesU.odstranNulyZCislaUctu(cisloU : string) : string;
begin
  Result := removeLeadingZeros(cisloU);
  if Result = '/0000' then Result := ''; //toto je u PayU, jako èíslo úètu protistrany uvádí samé nuly - 0000000000000000/0000
end;


function TDesU.prevedCisloUctuNaText(cisloU : string) : string;
begin
  Result := cisloU;
  if cisloU = '2100098382/2010' then Result := 'DES Fio bìžný';
  if cisloU = '2602372070/2010' then Result := 'DES Fio spoøící';
  if cisloU = '2800098383/2010' then Result := 'DES Fiokonto';
  if cisloU = '171336270/0300' then Result := 'DES ÈSOB';
  if cisloU = '2107333410/2700' then Result := 'PayU';
  if cisloU = '160987123/0300' then Result := 'Èeská pošta';
end;


procedure TDesU.opravRadekVypisuPomociVS(Vypis_ID, RadekVypisu_ID, VS : string);
var
  boAA: TAArray;
  sResponse: string;
begin
  boAA := TAArray.Create;
  boAA['VarSymbol'] := ''; //odstranit VS aby se Abra chytla pøi pøiøazení
  sResponse := self.abraBoUpdate(boAA.AsJSon, 'bankstatement', Vypis_ID, 'row', RadekVypisu_ID);

  boAA := TAArray.Create;
  boAA['VarSymbol'] := VS;
  sResponse := self.abraBoUpdate(boAA.AsJSon, 'bankstatement', Vypis_ID, 'row', RadekVypisu_ID);
end;


function TDesU.opravRadekVypisuPomociPDocument_ID(Vypis_ID, RadekVypisu_ID, PDocument_ID, PDocumentType : string) : string;
var
  boAA: TAArray;
  faktura : TDoklad;
  sResponse, text: string;
  preplatek : currency;
  abraWebApiResponse : TDesResult;
begin
  //pozor, funguje jen pro faktury, tedy PDocumentType "03"

  dbAbra.Reconnect;
  with qrAbraOC do begin

      // naètu dluh na faktuøe. když není kladný, konèíme
      SQL.Text := 'SELECT'
      + ' (ii.LOCALAMOUNT - ii.LOCALPAIDAMOUNT - ii.LOCALCREDITAMOUNT + ii.LOCALPAIDCREDITAMOUNT) as Dluh'
      + ' from ISSUEDINVOICES ii'
      + ' WHERE ii.ID = ''' + PDocument_ID  + '''';
      Open;
      if not Eof then begin
        if FieldByName('Dluh').AsCurrency <= 0 then begin
          Result := 'no_unpaid_amount';
          Close;
          Exit;
        end;
      end;
      Close;
  end;

  //spárovat fakturu Vypis_ID s øádkem výpisu RadekVypisu_ID
  boAA := TAArray.Create;
  boAA['PDocumentType'] := PDocumentType;
  boAA['PDocument_ID'] := PDocument_ID;
  sResponse := self.abraBoUpdate(boAA.AsJSon, 'bankstatement', Vypis_ID, 'row', RadekVypisu_ID);
  Result := 'ok';


  {* podívám se, jestli není po spárování pøeplacená (ii.LOCALAMOUNT - ii.LOCALPAIDAMOUNT - ii.LOCALCREDITAMOUNT + ii.LOCALPAIDCREDITAMOUNT)
   *  bude záporné. Pokud je pøeplacená,
   *    1) vložíme nový øádek s pøeplatkem - vyplníme Firm_ID, ale nevyplòujeme VS
   *    2) amount spárovaného øádku RadekVypisu_ID ponížíme o velikost pøeplatku
   }

  dbAbra.Reconnect;
  faktura := TDoklad.create(PDocument_ID, PDocumentType);

  if faktura.castkaNezaplaceno < 0 then //faktura je pøeplacená
  begin
    preplatek := - faktura.castkaNezaplaceno;

    with qrAbraOC do begin

      // naètu èástku z øádky výpisu
      SQL.Text := 'SELECT amount, text,'
        + ' bankaccount, specsymbol, docdate$date, accdate$date, firm_id'
        + ' FROM BANKSTATEMENTS2 WHERE id = ''' + RadekVypisu_ID + '''';
      Open;
      if Eof then begin
        Exit;
      end;

      //øádku z výpisu upravím -> snížím èástku o pøeplatek
      boAA := TAArray.Create;
      boAA['Amount'] := FieldByName('amount').AsCurrency - preplatek;
      text := FieldByName('text').AsString;
      boAA['Text'] := trim (copy(text, LastDelimiter('|', text) + 1, length(text)) ); //zkopíruju obsah za posledním svislítkem
      sResponse := self.abraBoUpdate(boAA.AsJSon, 'bankstatement', Vypis_ID, 'row', RadekVypisu_ID);

      //vyrobím nový øádek výpisu s èástkou pøeplatku
      boAA := TAArray.Create;
      boAA['amount'] := preplatek;
      boAA['text'] := FieldByName('text').AsString;
      boAA['bankaccount'] := FieldByName('bankaccount').AsString;
      boAA['specsymbol'] := FieldByName('specsymbol').AsString;
      boAA['docdate$date'] := FieldByName('docdate$date').AsFloat;
      boAA['accdate$date'] := FieldByName('accdate$date').AsFloat;
      boAA['firm_id'] := FieldByName('firm_id').AsString;
      boAA['credit'] := true;
      boAA['division_id'] := AbraEnt.getDivisionId;
      boAA['currency_id'] := AbraEnt.getCurrencyId;
      abraWebApiResponse := self.abraBoCreateRow(boAA.AsJSon, 'bankstatement', Vypis_ID);

      Close;
    end;
    Result := 'new_bsline_added';
  end;

end;

function TDesU.vytvorFaZaInternetKredit(VS : string; castka : currency; datum : double) : string;
var
  i: integer;
  firmAbraCode: String;
  newBo: TNewAbraBo;
begin
  Result := '';

  firmAbraCode := self.getAbracodeByContractNumber(VS);
  if firmAbraCode = '' then Exit;

  newBo := TNewAbraBo.Create('issuedinvoice');
  newBo.addInvoiceParams(datum);
  newBo.Item['DocQueue_ID'] := AbraEnt.getDocQueue('Code=FO4').ID;
  newBo.Item['VatDate$DATE'] := datum;
  newBo.Item['Firm_ID'] := self.getFirmIdByCode(firmAbraCode);
  newBo.Item['Varsymbol'] := VS;
  newBo.Item['Description'] := 'Kredit Internet';

  // 1. øádek
  newBo.createNewInvoiceRow(0, ' ');

 //2. øádek
  newBo.createNewInvoiceRow(1, 'Kredit Internet');
  newBo.rowItem['Totalprice'] := castka;
  //newBo.rowItem['BusOrder_Id'] := AbraEnt.getBusOrder('Name=kredit Internet').ID; // Busorder "kredit Internet" neexistuje


  try begin
    newBo.writeToAbra;
    if newBo.WriteResult.isOk then begin
      Result := newBo.getCreatedBoItem('id');
    end;
  end;
  except on E: exception do
    begin
      Application.MessageBox(PChar('Problem ' + ^M + E.Message), 'Vytvoøení fa');
      Result := 'Chyba pøi vytváøení faktury';
    end;
  end;


{ 24.3.2023 bylo zakomentováno, tabulka DE$_EuroFree se nepoužívá od cca 02-2023 
// 26.10.2019 datum vytvoøení faktury se uloží do tabulky DE$_EuroFree v databázi Abry
  with qrAbra do begin
    SQL.Text := 'UPDATE DE$_EuroFree SET'
    + ' AccDate = ''' + FormatDateTime('dd.mm.yyyy hh:nn:ss.zzz', Date) + ''''
    + ' WHERE Firm_Id = ''' + self.getFirmIdByCode(firmAbraCode) + ''''
    + ' AND AccDate IS NULL'
    + ' AND PayUDate = (SELECT MAX(PayUDate) FROM DE$_EuroFree'
     + ' WHERE Firm_Id = ''' + self.getFirmIdByCode(firmAbraCode) + ''''
     + ' AND AccDate IS NULL)';
    ExecSQL;
  end;
}
end;

function TDesU.vytvorFaZaVoipKredit(VS : string; castka : currency; datum : double) : string;
var
  firmAbraCode: string;
  newBo: TNewAbraBo;
begin
  Result := '';

  firmAbraCode := self.getAbracodeByContractNumber(VS);
  if firmAbraCode = '' then Exit;


  newBo := TNewAbraBo.Create('issuedinvoice');
  newBo.addInvoiceParams(datum);
  newBo.Item['DocQueue_ID'] := AbraEnt.getDocQueue('Code=FO2').ID;
  newBo.Item['VatDate$DATE'] := datum;
  newBo.Item['Firm_ID'] := self.getFirmIdByCode(firmAbraCode);
  newBo.Item['Varsymbol'] := VS;
  newBo.Item['Description'] := 'Kredit VoIP';


  // 1. øádek
  newBo.createNewInvoiceRow(0, ' ');

 //2. øádek
  newBo.createNewInvoiceRow(1, 'Kredit VoIP');
  newBo.rowItem['Totalprice'] := castka;
  newBo.rowItem['BusOrder_Id'] := AbraEnt.getBusOrder('Name=kredit VoIP').ID; // self.getAbraBusorderId('kredit VoIP'); // '6400000101' je 'kredit VoIP' v DB


  try begin
    newBo.writeToAbra;
    if newBo.WriteResult.isOk then begin
      Result := newBo.getCreatedBoItem('id');
    end;
  end;
  except on E: exception do
    begin
      Application.MessageBox(PChar('Problem ' + ^M + E.Message), 'Vytvoøení fa');
      Result := 'Chyba pøi vytváøení faktury';
    end;
  end;

end;


function TDesU.getAbraPeriodId(pYear : string) : string;
begin
  Result := AbraEnt.getPeriod('Code=' + pYear).ID
end;

function TDesU.getAbraPeriodId(pDate : double) : string;
var
    abraPeriod : TAbraPeriod;
begin
  abraPeriod := TAbraPeriod.create(pDate);
  Result := abraPeriod.ID;
end;


function TDesU.getAbraDocqueueId(code, documentType : string) : string;
begin
  with DesU.qrAbraOC do begin
    SQL.Text := 'SELECT Id FROM DocQueues'
              + ' WHERE Hidden = ''N'' AND Code = ''' + code  + ''' AND DocumentType = ''' + documentType + '''';
    Open;
    if not Eof then begin
      Result := FieldByName('Id').AsString;
    end;
    Close;
  end;
end;

function TDesU.getAbraDocqueueCodeById(id : string) : string;
begin
  with DesU.qrAbraOC do begin
    SQL.Text := 'SELECT Code FROM DocQueues'
              + ' WHERE Hidden = ''N'' AND Id = ''' + id  + '''';
    Open;
    if not Eof then begin
      Result := FieldByName('Code').AsString;
    end;
    Close;
  end;
end;

function TDesU.getAbraVatrateId(code : string) : string;
var
    abraVatIndex : TAbraVatIndex;
begin
  abraVatIndex := TAbraVatIndex.create(code);
  Result := abraVatIndex.VATRate_Id;
end;


function TDesU.getAbraVatindexId(code : string) : string;
var
    abraVatIndex : TAbraVatIndex;
begin
  abraVatIndex := TAbraVatIndex.create(code);
  Result := abraVatIndex.id;
end;



function TDesU.getAbracodeByVs(vs : string) : string;
begin
  with DesU.qrZakosOC do begin
    SQL.Text := 'SELECT abra_code FROM customers'
              + ' WHERE variable_symbol = ''' + vs + '''';
    Open;
    if not Eof then begin
      Result := FieldByName('abra_code').AsString;
    end;
    Close;
  end;
end;

function TDesU.getAbracodeByContractNumber(cnumber : string) : string;
begin
  with DesU.qrZakosOC do begin
    SQL.Text := 'SELECT cu.abra_code FROM customers cu, contracts co '
              + ' WHERE co.number = ''' + cnumber + ''''
              + ' AND cu.id = co.customer_id';
    Open;
    if not Eof then begin
      Result := FieldByName('abra_code').AsString;
    end;
    Close;
  end;
end;


function TDesU.getFirmIdByCode(code : string) : string;
begin
  with DesU.qrAbraOC do begin
    SQL.Text := 'SELECT Id FROM Firms'
              + ' WHERE Code = ''' + code + '''';
    Open;
    if not Eof then begin
      Result := FieldByName('Id').AsString;
    end;
    Close;
  end;
end;


function TDesU.getZustatekByAccountId (accountId : string; datum : double) : double;
var
sqlstring : string;
//currentAbraPeriod,
previousAbraPeriod, currentAbraPeriod : TAbraPeriod;
pocatecniZustatek, aktualniZustatek: double;

begin

  with DesU.qrAbraOC do begin

    currentAbraPeriod := TAbraPeriod.create(datum);

    SQL.Text := 'SELECT DEBITBEGINNING as pocatecniZustatek,'
      + '(DEBITBEGINNING - CREDITBEGINNING + DEBITBEGINNIGTURNOVER - CREDITBEGINNIGTURNOVER + DEBITTURNOVER - CREDITTURNOVER) as aktualniZustatek'
      + ' FROM ACCOUNTCALCULUS'
      + ' (''%'',''' + accountId + ''',null,null,'
      + FloatToStrFD(currentAbraPeriod.dateFrom) + ',' + FloatToStrFD(datum) + ','
      + ' '''',''0'','''',''0'','''',''0'','''',''0'',null,null,null,0)';
    Open;
    if not Eof then begin
      pocatecniZustatek := FieldByName('pocatecniZustatek').AsCurrency;
      aktualniZustatek := FieldByName('aktualniZustatek').AsCurrency;
    end;
    Close;

    if pocatecniZustatek = 0 then begin
      // Pokud je pocatecniZustatek nulový, znamená to nesprávnou hodnotu kvùli neuzavøenému pøedchozímu roku. Proto vypoèítáme koneèný zùstatek pøedchozího roku a ten pøièteme k aktuálnímu zùstatku.

      previousAbraPeriod := TAbraPeriod.create(datum - 365);

      SQL.Text := 'SELECT '
        + '(DEBITBEGINNING - CREDITBEGINNING + DEBITBEGINNIGTURNOVER - CREDITBEGINNIGTURNOVER + DEBITTURNOVER - CREDITTURNOVER) as konecnyZustatekPredchozihoRoku'
        + ' FROM ACCOUNTCALCULUS'
        + ' (''%'',''' + accountId + ''',null,null,'
        + FloatToStrFD(previousAbraPeriod.dateFrom) + ',' + FloatToStrFD(previousAbraPeriod.dateTo) + ','
        + ' '''',''0'','''',''0'','''',''0'','''',''0'',null,null,null,0)';
      Open;
      Result := aktualniZustatek + FieldByName('konecnyZustatekPredchozihoRoku').AsCurrency;
      Close;

    end else begin
      Result := aktualniZustatek;
    end;
  end;
end;


function TDesU.existujeVAbreDokladSPrazdnymVs() : boolean;
var
  messagestr: string;
begin
  Result := false;
  messagestr := '';
  dbAbra.Reconnect;

  //nutné rozdìlit do dvou èástí: pøetypování (CAST) padá, pokud existuje VS jako prázný øetìzec
  with DesU.qrAbraOC do begin

    // existuje fa s prázným VS?
//    SQL.Text := 'SELECT Id, DE$_CISLO_DOKLADU(DOCQUEUE_ID, ORDNUMBER, PERIOD_ID) as DOKLAD '  //takhle je to pomocí stored function, ale ABRA neumí stored function zálohovat
    SQL.Text := 'SELECT Id, (SELECT DOKLAD FROM DE$_CISLO_DOKLADU(DOCQUEUE_ID, ORDNUMBER, PERIOD_ID) as DOKLAD) '
              + ' FROM IssuedInvoices WHERE VarSymbol = ''''';
    Open;
    if not Eof then begin
      Result := true;
      ShowMessage('POZOR! V Abøe existuje faktura ' + FieldByName('DOKLAD').AsString + ' s prázným VS. Je potøeba ji opravit pøed pokraèováním práce s programem.');
      Close;
      Exit;
    end;
    Close;

    // existuje zálohový list s prázným VS?
    SQL.Text := 'SELECT Id, (SELECT DOKLAD FROM DE$_CISLO_DOKLADU(DOCQUEUE_ID, ORDNUMBER, PERIOD_ID) as DOKLAD) '
              + ' FROM IssuedDInvoices WHERE VarSymbol = ''''';
    Open;
    if not Eof then begin
      Result := true;
      ShowMessage('POZOR! V Abøe existuje zál. list ' + FieldByName('DOKLAD').AsString + ' s prázným VS. Je potøeba ho opravit pøed pokraèováním práce s programem.');
      Close;
      Exit;
    end;
    Close;

    // existuje fa s VS, který obahuje pouze nuly?
    SQL.Text := 'SELECT ID, VARSYMBOL, DOKLAD'
          + ' FROM (SELECT ID, VARSYMBOL, CAST(VARSYMBOL AS INTEGER) as VS_INT, (SELECT DOKLAD FROM DE$_CISLO_DOKLADU(DOCQUEUE_ID, ORDNUMBER, PERIOD_ID) as DOKLAD)'
          + ' FROM IssuedInvoices where VARSYMBOL < ''1'')'
          + ' WHERE VS_INT = 0';
    Open;
    while not Eof do begin
      Result := true;
      messagestr := messagestr + 'POZOR! V Abøe existuje faktura ' + FieldByName('DOKLAD').AsString + ' s VS "' + FieldByName('VARSYMBOL').AsString + '"' + sLineBreak;
      Next;
    end;

    // existuje zálohový list s VS, který obahuje pouze nuly?
    SQL.Text := 'SELECT ID, VARSYMBOL, DOKLAD'
          + ' FROM (SELECT ID, VARSYMBOL, CAST(VARSYMBOL AS INTEGER) as VS_INT, (SELECT DOKLAD FROM DE$_CISLO_DOKLADU(DOCQUEUE_ID, ORDNUMBER, PERIOD_ID) as DOKLAD)'
          + ' FROM IssuedDInvoices where VARSYMBOL < ''1'')'
          + ' WHERE VS_INT = 0';
    Open;
    while not Eof do begin
      Result := true;
      messagestr := messagestr + 'POZOR! V Abøe existuje zál. list ' + FieldByName('DOKLAD').AsString + ' s VS "' + FieldByName('VARSYMBOL').AsString + '"' + sLineBreak;
      Next;
    end;

    if messagestr <> '' then ShowMessage(messagestr);

    Close;
  end;
end;



{ *** DB Zakos (Aplikace) manipulating functions *** }

function TDesU.isVoipKreditContract(cnumber : string) : boolean;
begin
  with DesU.qrZakosOC do begin
    SQL.Text := 'SELECT co.id FROM contracts co  '
              + ' WHERE co.number = ''' + cnumber + ''''
             // + ' AND co.credit = 1' kredit // není nutné hlídat
              + ' AND co.tariff_id = 2'; //je to kreditni Voip
    Open;
    if not Eof then
      Result := true
    else
      Result := false;
    Close;
  end;
end;

function TDesU.isCreditContract(cnumber : string) : boolean;
begin
  with DesU.qrZakosOC do begin
    SQL.Text := 'SELECT co.id FROM contracts co  '
              + ' WHERE co.number = ''' + cnumber + ''''
              + ' AND co.credit = 1';
    Open;
    if not Eof then
      Result := true
    else
      Result := false;
    Close;
  end;
end;


function TDesU.zrusPenizeNaCeste(VS : string) : string;
var
  i: integer;
  boAA, boRowAA: TAArray;
  newId, firmAbraCode: String;
begin
  with DesU.qrZakosOC do begin
    SQL.Text := 'UPDATE customers SET pay_u_payment = 0 where variable_symbol = ''' + VS + '''';
    ExecSQL;
    Close;
  end;

end;


function TDesU.pdfToCustomerSendingDateTime(customerId : integer; pdfFileName : string) : double;
begin
  with DesU.qrZakosOC do begin
    SQL.Text := 'SELECT created_at FROM communications '
              + ' WHERE customer_id = ' + IntToStr(customerId)
              + ' AND content LIKE ''%' + pdfFileName + '%''';
    Open;
    if not Eof then begin
      Result := FieldByName('created_at').AsDateTime;
    end else
      Result := 0;
    Close;
  end;
end;


function TDesU.ulozKomunikaci(Typ, Customer_id, Zprava: string) : TDesResult;   // typ 2 je mail, typ 23 SMS
// ukládá záznam o odeslané zprávì zákazníkovi do tabulky "communications" v databázi aplikace
var
  CommId: integer;
  SQLStr: string;
begin
  with DesU.qrZakosOC do try
    Close;
    SQLStr := 'INSERT INTO communications ('
    + ' customer_id,'
    + ' user_id,'
    + ' communication_type_id,'
    + ' content,'
    + ' created_at,'
    + ' updated_at) VALUES ('
    + Customer_id + ', '
    + '1, '                                        // admin
    + Typ + ','
    + QuotedStr(Zprava) + ','
    + Ap + FormatDateTime('yyyy-mm-dd hh:nn:ss', Now) + ApC
    + Ap + FormatDateTime('yyyy-mm-dd hh:nn:ss', Now) + ApZ;
    SQL.Text := SQLStr;
    ExecSQL;
    Close;
    Result := TDesResult.create('ok', '');
  except on E: exception do
    Result := TDesResult.create('err', E.Message);
  end;

end;


{ *** PDF file's functions *** }

function TDesU.posliPdfEmailem(FullPdfFileName, emailAddrStr, emailPredmet, emailZprava, emailOdesilatel : string; ExtraPrilohaFileName: string = '') : TDesResult;
begin

  if not FileExists(FullPdfFileName) then begin
    Result := TDesResult.create('err', Format('Soubor %s neexistuje. Pøeskoèeno.', [FullPdfFileName]));
    Exit;
  end;

  // alespoò nìjaká kontrola mailové adresy
  if Pos('@', emailAddrStr) = 0 then begin
    Result := TDesResult.create('err', Format('Neplatná mailová adresa "%s". Pøeskoèeno.', [emailAddrStr]));
    Exit;
  end;


  emailAddrStr := StringReplace(emailAddrStr, ',', ';', [rfReplaceAll]);    // èárky za støedníky

  with idMessage do begin
    Clear;
    From.Address := emailOdesilatel;
    // ReceiptRecipient.Text := emailOdesilatel;

    // více mailových adres oddìlených støedníky se rozdìlí
    while Pos(';', emailAddrStr) > 0 do begin
      Recipients.Add.Address := Trim(Copy(emailAddrStr, 1, Pos(';', emailAddrStr)-1));
      emailAddrStr := Copy(emailAddrStr, Pos(';', emailAddrStr)+1, Length(emailAddrStr));
    end;
    Recipients.Add.Address := Trim(emailAddrStr);

    Subject := emailPredmet;
    ContentType := 'multipart/mixed';

    with TIdText.Create(idMessage.MessageParts, nil) do begin
      Body.Text := emailZprava;
      ContentType := 'text/plain';
      Charset := 'utf-8';
    end;

    with TIdAttachmentFile.Create(IdMessage.MessageParts, FullPdfFileName) do begin
      ContentType := 'application/pdf';
      FileName := ExtractFileName(FullPdfFileName);
    end;

    // pøidá se extra pøíloha, je-li vybrána
    if ExtraPrilohaFileName <> '' then
    with TIdAttachmentFile.Create(IdMessage.MessageParts, ExtraPrilohaFileName) do begin
      ContentType := 'application/pdf'; //pøípadná extra pøíloha má být dle zadání vždy jen pdf
      FileName := ExtraPrilohaFileName;
    end;

    try
      if not idSMTP.Connected then idSMTP.Connect;
      idSMTP.Send(idMessage);  // ! samotné poslání mailu
      Result := TDesResult.create('ok', Format('Soubor %s byl odeslán na adresu %s.', [FullPdfFileName, emailAddrStr]));
    except on E: exception do
      Result := TDesResult.create('err', Format('Soubor %s se nepodaøilo odeslat na adresu %s.' + #13#10 + 'Chyba: %s',
       [FullPdfFileName, emailAddrStr, E.Message]));
    end;

  end;

end;

procedure TDesU.syncAbraPdfToRemoteServer(year, month : integer);
// odeslání faktur pøevedených do PDF na vzdálený server
begin
  RunCMD (Format('WinSCP.com /command "option batch abort" "option confirm off" "open AbraPDF" "synchronize remote '
   + '%s%4d\%2.2d /home/abrapdf/%4d" "exit"', [DesU.PDF_PATH, year, month, year]), SW_SHOWNORMAL);
end;


{********************   GODS ************}
function TDesU.sendGodsSms(telCislo, smsText : string) : string;
var
  idHTTP: TIdHTTP;
  sstreamJson: TStringStream;
  callJson, postCallResult, godsSmsUrl: string;
begin
  Result := '';
  if (godsSmsUrl = '') then Result := 'chybi SMS url';
  if (telCislo = '') then Result := 'chybi tel. cislo';
  if (smsText = '') then Result := 'chybi SMS text';

  idHTTP := TidHTTP.Create;

  godsSmsUrl := getIniValue('Preferences', 'GodsSmsUrl');
  idHTTP.Request.BasicAuthentication := True;
  idHTTP.Request.Username := getIniValue('Preferences', 'GodsSmsUN');
  idHTTP.Request.Password := getIniValue('Preferences', 'GodsSmsPW');
  idHTTP.ReadTimeout := Round (10 * 1000); // ReadTimeout je v milisekundách
  idHTTP.Request.ContentType := 'application/json';
  idHTTP.Request.CharSet := 'utf-8';

   callJson := '{"from":"Eurosignal", "msisdn":"420' + telCislo + '",'
        + '"message":"' + smsText + '"}';

  sstreamJson := TStringStream.Create(callJson, TEncoding.UTF8);

  try
    try
      if (godsSmsUrl <> '') AND (telCislo <> '') AND (smsText <> '') then

      begin
        Result := idHTTP.Post(godsSmsUrl, sstreamJson);


    //  takhle vypadá vzorový response z GODS SMS brány
    //  {
    //    "status": "success",
    //    "data": {"message-id": "c3d8ebd4-879e-412e-8e16-3d4a2335d967"  }
    //  }

      end;
    except
      on E: Exception do begin
        ShowMessage('Error on request: '#13#10 + e.Message);
      end;
    end;
  finally
    sstreamJson.Free;
    idHTTP.Free;
  end;

  if appMode >= 1 then begin
    if not DirectoryExists(PROGRAM_PATH + '\logy\') then
      Forcedirectories(PROGRAM_PATH + '\logy\');
    appendToFile(PROGRAM_PATH + '\logy\SMS_odeslane.txt', formatdatetime('dd.mm.yyyy hh:nn:ss', Now) + ' - ' + telCislo + ' - '  + smsText + ' - '  + Result);
  end;

end;

{********************   SMSbrána ************}
function TDesU.sendSms(telCislo, smsText: string): string;
var
  idHTTP: TIdHTTP;
  SmsUrl,
  SmsUN,
  SmsPW,
  sXMLResult,
  sHTTPtext: string;
  iErrOd,
  iErrDo,
  iResult: integer;
begin
  Result := '';

  idHTTP := TidHTTP.Create;
  idHTTP.IOHandler := TIdSSLIOHandlerSocketOpenSSL.Create;
  SmsUrl := getIniValue('Preferences', 'SmsUrl');
  SmsUN := getIniValue('Preferences', 'SmsUN');
  SmsPW := getIniValue('Preferences', 'SmsPW');
//  idHTTP.Request.BasicAuthentication := True;
//  idHTTP.Request.Username := SmsUN;
//  idHTTP.Request.Password := SmsPW;
  idHTTP.ReadTimeout := Round (10 * 1000); // ReadTimeout je v milisekundách
  idHTTP.Request.ContentType := 'text/plain';
  idHTTP.Request.CharSet := 'ASCII';

  sHTTPtext := SmsUrl + '?login=' + SmsUN + '&password=' + SmsPW + '&action=send_sms&delivery_report=0&number=' + telCislo + '&message=' + odstranDiakritiku(smsText);

  try
    try
      if (SmsUrl <> '') AND (telCislo <> '') AND (smsText <> '') then sXMLResult := idHTTP.Get(sHTTPtext);
    except
      on E: Exception do begin
        ShowMessage('Error on request: '#13#10 + E.Message);
        Exit;
      end;
    end;

// vrácené XML má mezi <err> a </err> èíslo (-1 až 12), 0 je OK
    iErrOd := Pos('<err>', sXMLResult) + 5;
    iErrDo := Pos('</err>', sXMLResult);
    iResult := StrToInt(Copy(sXMLResult, iErrOd, iErrDo-iErrOd));
    case iResult of
      0: Result := 'OK';
      2: Result := 'Bad username';
      3: Result := 'Bad password';
      9: Result := 'Low credit';
      10: Result := 'Bad phone number';
      11: Result := 'No message';
      12: Result := 'Too long message ';
    else
      Result := 'Error number ' + IntToStr(iResult);
    end;

  finally
    idHTTP.Free;
    idHTTP.IOHandler.Free;
  end;

end;



{***************************************************************************}
{********************     General helper functions     *********************}
{***************************************************************************}


{** class TDesResult **}

constructor TDesResult.create(Code, Messg : string);
begin
  self.Code := Code;
  self.Messg := Messg;
end;

function TDesResult.isOk : boolean;
begin
  Result := LowerCase(LeftStr(self.Code, 2)) = 'ok';
end;

function TDesResult.isErr : boolean;
begin
  Result := LowerCase(LeftStr(self.Code, 3)) = 'err';
end;


{** other functions **}

// pøedaný string vrátí bez háèku a èárek
function odstranDiakritiku(textDiakritika : string) : string;
var
  s: string;
  b: TBytes;
begin
  s := textDiakritika;
  b := TEncoding.Unicode.GetBytes(s);
  b := TEncoding.Convert(TEncoding.Unicode, TEncoding.ASCII, b);
  Result := TEncoding.ASCII.GetString(b);
end;

// odstraní ze stringu nuly na zaèátku
function removeLeadingZeros(const Value: string): string;
// ostraní nuly z øetìzce zleva. pokud je øetìzec složený pouze z nul, vrátí prázdný øetìzec.
var
  i: Integer;
begin
  for i := 1 to Length(Value) do
    if Value[i]<>'0' then
    begin
      Result := Copy(Value, i, MaxInt);
      exit;
    end;
  Result := '';
end;

//zaplní øetìzec nulama zleva až do celkové délky lenght
function LeftPad(value:integer; length:integer=8; pad:char='0'): string; overload;
begin
   result := RightStr(StringOfChar(pad,length) + IntToStr(value), length );
end;

function LeftPad(value: string; length:integer=8; pad:char='0'): string; overload;
begin
   result := RightStr(StringOfChar(pad,length) + value, length );
end;

function Str6digitsToDate(datum : string) : double;
begin
  Result := StrToDate(copy(datum, 1, 2) + '.' + copy(datum, 3, 2) + '.20' + copy(datum, 5, 2));
end;

function IndexByName(DataObject: variant; Name: ShortString): integer;
// náhrada za nefunkèní DataObject.ValuByName(Name)
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

function pocetRadkuTxtSouboru(SName: string): integer;
var
  oSL : TStringlist;
begin
  oSL := TStringlist.Create;
  oSL.LoadFromFile(SName);
  Result := oSL.Count;
  oSL.Free;
end;

function RemoveSpaces(const s: string): string;
var
  len, p: integer;
  pc: PChar;
const
  WhiteSpace = [#0, #9, #10, #13, #32, #160];

begin
  len := Length(s);
  SetLength(Result, len);

  pc := @s[1];
  p := 0;
  while len > 0 do begin
    if not (pc^ in WhiteSpace) then begin
      inc(p);
      Result[p] := pc^;
    end;
  inc(pc);
  dec(len);
  end;

  SetLength(Result, p);
end;

function KeepOnlyNumbers(const S: string): string;
var
  i: Integer;
begin
  Result := '';
  for i := 1 to Length(S) do
    if CharInSet(S[i], ['0'..'9']) then
      Result := Result + S[i];
end;

function SplitStringInTwo(const InputString, Delimiter: string; out LeftPart, RightPart: string): Boolean;
var
  DelimiterPos: Integer;
begin
  DelimiterPos := Pos(Delimiter, InputString);
  Result := (DelimiterPos > 0);

  if Result then
  begin
    LeftPart := Copy(InputString, 1, DelimiterPos - 1);
    RightPart := Copy(InputString, DelimiterPos + Length(Delimiter), Length(InputString) - DelimiterPos - Length(Delimiter) + 1);
  end
  else
  begin
    LeftPart := '';
    RightPart := '';
  end;
end;

function destilujTelCislo(telCislo: string): string;
var
  regexpr : TRegEx;
  match   : TMatch;
begin
  Result := '';

  telCislo := stringreplace(telCislo, '+420', '', [rfReplaceAll, rfIgnoreCase]);
  telCislo := RemoveSpaces(telCislo); // nekdy jsou cisla psana jako 3 skupiny po 3 znacich

  regexpr := TRegEx.Create('\d{9}',[roIgnoreCase,roMultiline]); //hledam devitimistne cislo
  match := regexpr.Match(telCislo);

  if match.Success then
      Result := match.Value;

end;

function destilujMobilCislo(telCislo: string): string;
var
  regexpr : TRegEx;
  match   : TMatch;
begin
  Result := '';

  telCislo := stringreplace(telCislo, '+420', '', [rfReplaceAll, rfIgnoreCase]);
  telCislo := RemoveSpaces(telCislo); // nekdy jsou cisla psana jako 3 skupiny po 3 znacich

  regexpr := TRegEx.Create('[6-7]{1}\d{8}',[roIgnoreCase,roMultiline]); //hledam devitimistne cislo zacinajici na 6 nebo 7
  match := regexpr.Match(telCislo);

  if match.Success then
      Result := match.Value;

end;


function FindInFolder(sFolder, sFile: string; bUseSubfolders: Boolean): string;
var
  sr: TSearchRec;
  i: Integer;
  sDatFile: String;
begin
  Result := '';
  sFolder := IncludeTrailingPathDelimiter(sFolder);
  if System.SysUtils.FindFirst(sFolder + sFile, faAnyFile - faDirectory, sr) = 0 then
  begin
    Result := sFolder + sr.Name;
    System.SysUtils.FindClose(sr);
    Exit;
  end;

  //not found .... search in subfolders
  if bUseSubfolders then
  begin
    //find first subfolder
    if System.SysUtils.FindFirst(sFolder + '*.*', faDirectory, sr) = 0 then
    begin
      try
        repeat
          if ((sr.Attr and faDirectory) <> 0) and (sr.Name <> '.') and (sr.Name <> '..') then //is real folder?
          begin
            //recursive call!
            //Result := FindInFolder(sFolder + sr.Name, sFile, bUseSubfolders); // plná rekurze
            Result := FindInFolder(sFolder + sr.Name, sFile, false); // rekurze jen do 1. úrovnì

            if Length(Result) > 0 then Break; //found it ... escape
          end;
        until System.SysUtils.FindNext(sr) <> 0;  //...next subfolder
      finally
        System.SysUtils.FindClose(sr);
      end;
    end;
  end;
end;

procedure writeToFile(pFileName, pContent : string);
var
    OutputFile : TextFile;
begin
  AssignFile(OutputFile, pFileName);
  ReWrite(OutputFile);
  WriteLn(OutputFile, pContent);
  CloseFile(OutputFile);
end;

procedure appendToFile(pFileName, pContent : string);
var
    F : TextFile;
begin
  AssignFile(F, pFileName);
  try
    if FileExists(pFileName) = false then
      Rewrite(F)
    else
    begin
      Append(F);
    end;
    Writeln(F, pContent);
  finally
    CloseFile(F);
  end;
end;

function LoadFileToStr(const FileName: TFileName): ansistring;
var
  FileStream : TFileStream;
begin
  FileStream:= TFileStream.Create(FileName, fmOpenRead or fmShareDenyWrite);
    try
     if FileStream.Size>0 then
     begin
      SetLength(Result, FileStream.Size);
      FileStream.Read(Pointer(Result)^, FileStream.Size);
     end;
    finally
     FileStream.Free;
    end;
end;

function RandString(const stringsize: integer): string;
var
  i: integer;
  const ss: string = 'abcdefghijklmnopqrstuvwxyz';
begin
  for i:=1 to stringsize do
      Result:=Result + ss[random(length(ss)) + 1];
end;

// FloatToStr s nahrazením èárek teèkami
function FloatToStrFD (pFloat : extended) : string;
begin
  Result := AnsiReplaceStr(FloatToStr(pFloat), ',', '.');
end;

function debugRozdilCasu(cas01, cas02 : double; textZpravy : string) : string;
// do DebugString vypíše èasový rozdíl a zprávu
begin
  cas01 := cas01 * 24 * 3600;
  cas02 := cas02 * 24 * 3600;
  Result := 'Trvání: ' + FloatToStr(RoundTo(cas02 - cas01, -2)) + ' s, ' + textZpravy;
  if DesU.appMode >= 3 then OutputDebugString(PChar(Result));
end;

procedure RunCMD(cmdLine: string; WindowMode: integer);

  procedure ShowWindowsErrorMessage(r: integer);
  var
    sMsg: string;
  begin
    sMsg := SysErrorMessage(r);
    if (sMsg = '') and (r = ERROR_BAD_EXE_FORMAT) then
      sMsg := SysErrorMessage(ERROR_BAD_FORMAT);
    MessageDlg(sMsg, mtError, [mbOK], 0);
  end;

var
  si: TStartupInfo;
  pi: TProcessInformation;
  sei: TShellExecuteInfo;
  err: Integer;
begin
  // We need a function which does following:
  // 1. Replace the Environment strings, e.g. %SystemRoot%  --> ExpandEnvStr
  // 2. Runs EXE files with parameters (e.g. "cmd.exe /?")  --> WinExec
  // 3. Runs EXE files without path (e.g. "calc.exe")       --> WinExec
  // 4. Runs EXE files without extension (e.g. "calc")      --> WinExec
  // 5. Runs non-EXE files (e.g. "Letter.doc")              --> ShellExecute
  // 6. Commands with white spaces (e.g. "C:\Program Files\xyz.exe") must be enclosed in quotes.

  //cmdLine := ExpandEnvStr(cmdLine); // *hw* nefunguje mi, nepotrebujeme
  {$IFDEF UNICODE}
  UniqueString(cmdLine);
  {$ENDIF}

  ZeroMemory(@si, sizeof(si));
  si.cb := sizeof(si);
  si.dwFlags := STARTF_USESHOWWINDOW;
  si.wShowWindow := WindowMode;

  if CreateProcess(nil, PChar(cmdLine), nil, nil, False, 0, nil, nil, si, pi) then
  begin
    CloseHandle(pi.hThread);
    CloseHandle(pi.hProcess);
    Exit;
  end;

  err := GetLastError;
  if (err = ERROR_BAD_EXE_FORMAT) or
     (err = ERROR_BAD_FORMAT) then
  begin
    ZeroMemory(@sei, sizeof(sei));
    sei.cbSize := sizeof(sei);
    sei.lpFile := PChar(cmdLine);
    sei.nShow := WindowMode;

    if ShellExecuteEx(@sei) then Exit;
    err := GetLastError;
  end;

  ShowWindowsErrorMessage(err);
end;

function getWindowsUserName : string;
// pøívìtivìjší GetUserName
var
  dwSize : DWord;
begin
  SetLength(Result, 32);
  dwSize := 31;
  GetUserName(PChar(Result), dwSize);
  SetLength(Result, dwSize);
end;

// ------------------------------------------------------------------------------------------------

function getWindowsCompName : string;
// pøívìtivìjší GetComputerName
var
  dwSize : DWord;
begin
  SetLength(Result, 32);
  dwSize := 31;
  GetComputerName(PChar(Result), dwSize);
  SetLength(Result, dwSize);
end;


end.

