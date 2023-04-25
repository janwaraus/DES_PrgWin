unit uTVypis;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, IniFiles, Forms,
  Dialogs, StdCtrls, Grids, AdvObj, BaseGrid, AdvGrid, StrUtils, RegularExpressions,
  DB, ComObj, AdvEdit, DateUtils, Math, ExtCtrls,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, ZAbstractConnection, ZConnection,
  uTPlatbaZVypisu, AbraEntities;

type
  TArrayOf2Int = array[0..1] of integer;

  TVypis = class
  private
    qrAbra: TZQuery;
  public
    Platby : TList;
    abraBankaccount : TAbraBankaccount;
    poradoveCislo : integer;
    cisloUctuVlastni : string[16]; // èíslo úètu bez nul na zaèátku a bez kódu banky
    datum  : double; //datum vypisu se urci jako datum poslední platby, v metodì init()
    datumZHlavicky  : double; //nemusí být správnì, pro PayU napø. není
    zustatekPocatecni, zustatekKoncovy : currency;
    obratDebet, obratKredit  : currency;
    maxExistujiciPoradoveCislo : integer;
    maxExistujiciExtPoradoveCislo : integer;
    searchIndex : integer;

    constructor Create(gpcLine : string);
  published
    procedure init(gpcFilename : string = '');
    procedure setridit();    
    procedure nactiMaxExistujiciPoradoveCislo();
    function isNavazujeNaRadu() : boolean;
    function isPayuVypis() : boolean;
    function prictiCastkuPokudDvojitaPlatba(pPlatbaZVypisu : TPlatbaZVypisu) : integer;
    function isPayuProvize(pPlatbaZVypisu : TPlatbaZVypisu) : boolean;
    function hledej(needle : string) : TArrayOf2Int;
  end;

implementation

uses
  DesUtils;

constructor TVypis.Create(gpcLine : string);
begin
  self.qrAbra := DesU.qrAbra;
  self.Platby := TList.Create;
  self.AbraBankAccount := TAbraBankaccount.create;

  self.cisloUctuVlastni := removeLeadingZeros(copy(gpcLine, 4, 16));
  self.zustatekPocatecni := StrToInt(copy(gpcLine, 46, 14)) / 100; //pøípadnì TODO znaménko
  self.zustatekKoncovy := StrToInt(copy(gpcLine, 61, 14)) / 100; //pøípadnì TODO znaménko
  self.obratDebet := StrToInt(copy(gpcLine, 76, 14)) / 100;
  self.obratKredit := StrToInt(copy(gpcLine, 91, 14)) / 100;
  self.poradoveCislo := StrToInt(copy(gpcLine, 106, 3));
  self.datumZHlavicky := Str6digitsToDate(copy(gpcLine, 109, 6));
end;


function TVypis.isNavazujeNaRadu() : boolean;
begin
  if (self.poradoveCislo - self.maxExistujiciPoradoveCislo = 1)
    OR (self.poradoveCislo = 0) then //PayU nikdy nenavazuje, výpis má vždy èíslo 0
    Result := true
  else
    Result := false;
end;

function TVypis.isPayuVypis() : boolean;
begin
  if self.cisloUctuVlastni = '2389210008000000' then
    Result := true
  else
    Result := false;
end;


procedure TVypis.nactiMaxExistujiciPoradoveCislo();
begin
  with qrAbra do begin
    SQL.Text := 'SELECT bs1.OrdNumber as MaxOrdNumber, bs1.EXTERNALNUMBER as MaxExtOrdNumber'
              + ' FROM BANKSTATEMENTS bs1'
              + ' WHERE bs1.DOCQUEUE_ID = ''' + self.AbraBankAccount.bankStatementDocqueueId  + ''''
              + ' AND bs1.PERIOD_ID = (SELECT ID FROM PERIODS p WHERE p.DATEFROM$DATE <= ' + FloatToStr(self.datum)
              + ' AND p.DATETO$DATE > ' + FloatToStr(self.datum) + ')'
              + ' AND bs1.OrdNumber = (SELECT max(bs2.ORDNUMBER) FROM BANKSTATEMENTS bs2 WHERE bs1.DOCQUEUE_ID = bs2.DOCQUEUE_ID and bs1.PERIOD_ID = bs2.PERIOD_ID)';
    Open;
    if not Eof then begin
      self.maxExistujiciPoradoveCislo := FieldByName('MaxOrdNumber').AsInteger;
      self.maxExistujiciExtPoradoveCislo := FieldByName('MaxExtOrdNumber').AsInteger;
    end
    else begin
      self.maxExistujiciPoradoveCislo := 0;
      self.maxExistujiciExtPoradoveCislo := 0;
    end;
    Close;
  end;
end;


procedure TVypis.init(gpcFilename : string = '');
var
  i : integer;
  iPlatba, payuProvizePP : TPlatbaZVypisu;
  scitatPayUProvize : boolean;
  payuProvize : currency;
  datumVypisu: string;
  //Match: TMatch;
  Matches : TMatchCollection;
begin

  self.abraBankaccount.loadByNumber(self.cisloUctuVlastni);
  self.datum := TPlatbaZVypisu(self.Platby[self.Platby.Count - 1]).datum; //datum vypisu se urci jako datum poslední platby
  self.nactiMaxExistujiciPoradoveCislo();

  if self.isPayuVypis() then
  begin
    // urèíme datum výpisu z názvu souboru
    datumVypisu := TRegEx.Match(gpcFilename, '\d{4}\D\d{2}\D\d{2}\.gpc').Value; //nejprve vezmu poslední datum pøed .gpc
    Matches := TRegEx.Matches(datumVypisu, '(\d+)'); // pak z nìj vezmu èíslice. v jednom regexp jsem to napsat nedokázal :)
    datumVypisu := Matches[2].Value + '.' + Matches[1].Value + '.' + Matches[0].Value;

    if not (self.datum = StrToDate(datumVypisu)) then begin
      MessageDlg('Datum poslední platby ve výpisu je ' + DateToStr(self.datum)+ sLineBreak + 'Datum výpisu odvozené z názvu GPC souboru je však ' + DateToStr(StrToDate(datumVypisu)), mtInformation, [mbOk], 0);
      self.datum := StrToDate(datumVypisu);
    end;

    // vyøešíme seètení PayU provizí
    payuProvize := 0;
    {if Dialogs.MessageDlg('Seèíst PayU provize?', mtConfirmation, [mbYes, mbNo], 0, mbYes) = mrYes then}
    // pro každou platbu se podíváme, zda je PayU provize
    for i := self.Platby.Count - 1 downto 0 do
    begin
      iPlatba := TPlatbaZVypisu(self.Platby[i]);

      // pokud je PayU provize, pøièteme èástku a platbu odstraníme
      if self.isPayuProvize(iPlatba) then
      begin
        payuProvize := payuProvize + iPlatba.castka;
        self.Platby.Delete(i);
      end;
      if (AnsiContainsStr(iPlatba.nazevProtistrany, 'ubscription fee')) then
        iPlatba.nazevProtistrany := formatdatetime('myy', iPlatba.datum) + ' suma provize'; // jedno za mìsíc si PayU strhává 199 Kè, oznaèí se také jako "MYY suma provize", aby se to v ABRA dalo jednoduše seèíst
    end;

    if payuProvize > 0 then
    begin
      payuProvizePP := TPlatbaZVypisu.Create(-payuProvize);
      payuProvizePP.datum := self.datum;
      payuProvizePP.nazevProtistrany := formatdatetime('myy', payuProvizePP.datum) + ' suma provize';
      self.Platby.Add(payuProvizePP);
    end;
  end;

end;

function TVypis.isPayuProvize(pPlatbaZVypisu : TPlatbaZVypisu) : boolean;
{ Výpis urèí, zda je zkoumaná platba (pPlatbaZVypisu) PayU provizí,
  a to tak, že se snaží nalézt platbu, která
   - je pøíslušným kreditem ke zkoumané vstupní platbì (pPlatbaZVypisu)
   - pøièemž zkoumaná platba musí být debetem v rozmezí 1-7% pøíslušného kreditu (uvažujeme PauY provizi v rozmezí 1-7%, v roce 2023 je 3%)
    - pøípadnì menší než 10 Kè (PayU u malých plateb zvedá procento poplatku)
}
var
  i : integer;
  iPlatba : TPlatbaZVypisu;
begin
  Result := false;
  for i := 0 to self.Platby.Count - 1 do
  begin
    iPlatba := TPlatbaZVypisu(self.Platby[i]);
    if (pPlatbaZVypisu.kodUctovani = '1') //je to èistý debet, není to storno kreditu
      AND (iPlatba.VS = pPlatbaZVypisu.VS) AND (iPlatba.kodUctovani = '2') // existuje jiná platba se stejným VS  a ta je kredit
      AND (pPlatbaZVypisu.castka > iPlatba.castka * 0.01) AND (pPlatbaZVypisu.castka < Max(iPlatba.castka * 0.07, 10))  // èástka je vìtší než 1% kreditu a menší než 7% kreditu (nebo menší než 10 Kè). pokud není, nejedná se o provizi
      // AND self.cisloUctuVlastni = '2389210008000000' // toto není nutné, funkce se volá jen na PayU výpisy. pro jistotu by se dalo dát, ale zase je to hardcode èísla úètu PayU...      
    then begin
      Result := true;
      exit;
    end;
  end;
end;

procedure TVypis.setridit();
var
  i : integer;
  iPlatba : TPlatbaZVypisu;

begin

  for i := self.Platby.Count - 1 downto 0 do
  begin
    iPlatba := TPlatbaZVypisu(self.Platby[i]);
    if (iPlatba.problemLevel = 5) AND (iPlatba.VS = iPlatba.VSOrig) then begin
      self.Platby.Delete(i);
      self.Platby.Add(iPlatba);
    end;
  end;

  for i := self.Platby.Count - 1 downto 0 do
  begin
    iPlatba := TPlatbaZVypisu(self.Platby[i]);
    if (iPlatba.problemLevel = 2) AND (iPlatba.VS = iPlatba.VSOrig) then begin
      self.Platby.Delete(i);
      self.Platby.Add(iPlatba);
    end;
  end;

  for i := self.Platby.Count - 1 downto 0 do
  begin
    iPlatba := TPlatbaZVypisu(self.Platby[i]);
    if (iPlatba.problemLevel = 3) AND (iPlatba.VS = iPlatba.VSOrig) then begin
      self.Platby.Delete(i);
      self.Platby.Add(iPlatba);
    end;
  end;

  for i := self.Platby.Count - 1 downto 0 do
  begin
    iPlatba := TPlatbaZVypisu(self.Platby[i]);
    if (iPlatba.problemLevel = 1) AND (iPlatba.VS = iPlatba.VSOrig) then begin
      self.Platby.Delete(i);
      self.Platby.Add(iPlatba);
    end;
  end;

  for i := self.Platby.Count - 1 downto 0 do
  begin
    iPlatba := TPlatbaZVypisu(self.Platby[i]);
    if (iPlatba.problemLevel = 0) AND (iPlatba.VS = iPlatba.VSOrig) then begin
      self.Platby.Delete(i);
      self.Platby.Add(iPlatba);
    end;
  end;

  // debety dozadu
  for i := self.Platby.Count - 1 downto 0 do
  begin
    iPlatba := TPlatbaZVypisu(self.Platby[i]);
    if iPlatba.debet then begin
      self.Platby.Delete(i);
      self.Platby.Add(iPlatba);
    end;
  end;

end;

function TVypis.prictiCastkuPokudDvojitaPlatba(pPlatbaZVypisu : TPlatbaZVypisu) : integer;
// Pokud pøijde 2 a více plateb ze stejného úètu se stejným VS ve stejný den, seèteme je. Dále k nim pøistupujeme jako k jedné platbì
// návratová hodnota je poøadové èíslo (identifikace) již existující platby
var
  i : integer;
  iPlatba : TPlatbaZVypisu;

begin
  Result := -1;
  for i := 0 to self.Platby.Count - 1 do
  begin
    iPlatba := TPlatbaZVypisu(self.Platby[i]);
    if ( (iPlatba.cisloUctu <> '160987123/0300') // 'Èeská pošta, nemá cenu naèítat historii a navíc je to pomalé
     and (iPlatba.cisloUctu <> '') ) // èíslo úètu protistrany nulové, má to tak PayU
    then
    if (iPlatba.VS = pPlatbaZVypisu.VS) AND (iPlatba.cisloUctu = pPlatbaZVypisu.cisloUctu)
      AND (iPlatba.kredit = true) AND (iPlatba.znamyPripad = false)
      AND (pPlatbaZVypisu.kredit = true) AND (pPlatbaZVypisu.znamyPripad = false)
    then begin
      iPlatba.castka := iPlatba.castka + pPlatbaZVypisu.castka;
      Result := i;
      exit;
    end;
  end;
end;

function TVypis.hledej(needle : string) : TArrayOf2Int;

  procedure IncSearchIndex(var sIndex: Integer; pCount: Integer); inline;
  begin
    Inc(sIndex);
    if sIndex = pCount then sIndex := 0;
  end;

var
  searchedItems : integer;
  iPlatba : TPlatbaZVypisu;

begin
  searchedItems := 0;

  for searchedItems := 0 to self.Platby.Count - 1 do
  begin
    iPlatba := TPlatbaZVypisu(self.Platby[self.searchIndex]);

    if AnsiContainsStr(iPlatba.VS, needle) then begin
      Result[0] := self.searchIndex;
      Result[1] := 2;
      IncSearchIndex(self.searchIndex, self.Platby.Count);
      Exit;
    end;

    if AnsiContainsStr(iPlatba.cisloUctuKZobrazeni, needle) then begin
      Result[0] := self.searchIndex;
      Result[1] := 4;
      IncSearchIndex(self.searchIndex, self.Platby.Count);
      Exit;
    end;


    if AnsiContainsStr(AnsiLowerCase(iPlatba.nazevProtistrany), AnsiLowerCase(needle)) then begin
      Result[0] := self.searchIndex;
      Result[1] := 5;
      IncSearchIndex(self.searchIndex, self.Platby.Count);
      Exit;
    end;

  IncSearchIndex(self.searchIndex, self.Platby.Count);
  end;

  Result[0] := 0;
  Result[1] := 0;

end;




end.
