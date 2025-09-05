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
    cisloUctuVlastni : string[16]; // ��slo ��tu bez nul na za��tku a bez k�du banky
    datum  : double; //datum vypisu se urci jako datum posledn� platby, v metod� init()
    datumZHlavicky  : double; //nemus� b�t spr�vn�, pro PayU nap�. nen�
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
  self.zustatekPocatecni := StrToInt(copy(gpcLine, 46, 14)) / 100; //p��padn� TODO znam�nko
  self.zustatekKoncovy := StrToInt(copy(gpcLine, 61, 14)) / 100; //p��padn� TODO znam�nko
  self.obratDebet := StrToInt(copy(gpcLine, 76, 14)) / 100;
  self.obratKredit := StrToInt(copy(gpcLine, 91, 14)) / 100;
  self.poradoveCislo := StrToInt(copy(gpcLine, 106, 3));
  self.datumZHlavicky := Str6digitsToDate(copy(gpcLine, 109, 6));
end;


function TVypis.isNavazujeNaRadu() : boolean;
begin
  if (self.poradoveCislo - self.maxExistujiciPoradoveCislo = 1)
    OR (self.poradoveCislo = 0) then //PayU nikdy nenavazuje, v�pis m� v�dy ��slo 0
    Result := true
  else
    Result := false;
end;

function TVypis.isPayuVypis() : boolean;
begin
  //if self.cisloUctuVlastni = '2389210008000000' then
  if self.cisloUctuVlastni = '4003292157000000' then
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
  self.datum := TPlatbaZVypisu(self.Platby[self.Platby.Count - 1]).datum; //datum vypisu se urci jako datum posledn� platby
  self.nactiMaxExistujiciPoradoveCislo();

  if self.isPayuVypis() then
  begin
    // ur��me datum v�pisu z n�zvu souboru
    datumVypisu := TRegEx.Match(gpcFilename, '\d{4}\D\d{2}\D\d{2}\.gpc').Value; //nejprve vezmu posledn� datum p�ed .gpc
    Matches := TRegEx.Matches(datumVypisu, '(\d+)'); // pak z n�j vezmu ��slice. v jednom regexp jsem to napsat nedok�zal :)
    datumVypisu := Matches[2].Value + '.' + Matches[1].Value + '.' + Matches[0].Value;

    if not (self.datum = StrToDate(datumVypisu)) then begin
      MessageDlg('Datum posledn� platby ve v�pisu je ' + DateToStr(self.datum)+ sLineBreak + 'Datum v�pisu odvozen� z n�zvu GPC souboru je v�ak ' + DateToStr(StrToDate(datumVypisu)), mtInformation, [mbOk], 0);
      self.datum := StrToDate(datumVypisu);
    end;

    // vy�e��me se�ten� PayU proviz�
    payuProvize := 0;
    {if Dialogs.MessageDlg('Se��st PayU provize?', mtConfirmation, [mbYes, mbNo], 0, mbYes) = mrYes then}
    // pro ka�dou platbu se pod�v�me, zda je PayU provize
    for i := self.Platby.Count - 1 downto 0 do
    begin
      iPlatba := TPlatbaZVypisu(self.Platby[i]);

      // pokud je PayU provize, p�i�teme ��stku a platbu odstran�me
      if self.isPayuProvize(iPlatba) then
      begin
        payuProvize := payuProvize + iPlatba.castka;
        self.Platby.Delete(i);
      end;
      if (AnsiContainsStr(iPlatba.nazevProtistrany, 'ubscription fee')) then
        iPlatba.nazevProtistrany := formatdatetime('myy', iPlatba.datum) + ' suma provize'; // jedno za m�s�c si PayU strh�v� 199 K�, ozna�� se tak� jako "MYY suma provize", aby se to v ABRA dalo jednodu�e se��st
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
{ V�pis ur��, zda je zkouman� platba (pPlatbaZVypisu) PayU proviz�,
  a to tak, �e se sna�� nal�zt platbu, kter�
   - je p��slu�n�m kreditem ke zkouman� vstupn� platb� (pPlatbaZVypisu)
   - p�i�em� zkouman� platba mus� b�t debetem v rozmez� 1-7% p��slu�n�ho kreditu (uva�ujeme PauY provizi v rozmez� 1-7%, v roce 2023 je 3%)
    - p��padn� men�� ne� 10 K� (PayU u mal�ch plateb zved� procento poplatku)
}
var
  i : integer;
  iPlatba : TPlatbaZVypisu;
begin
  Result := false;
  for i := 0 to self.Platby.Count - 1 do
  begin
    iPlatba := TPlatbaZVypisu(self.Platby[i]);
    if (pPlatbaZVypisu.kodUctovani = '1') //je to �ist� debet, nen� to storno kreditu
      AND (iPlatba.VS = pPlatbaZVypisu.VS) AND (iPlatba.kodUctovani = '2') // existuje jin� platba se stejn�m VS  a ta je kredit
      AND (pPlatbaZVypisu.castka > iPlatba.castka * 0.01) AND (pPlatbaZVypisu.castka < Max(iPlatba.castka * 0.07, 10))  // ��stka je v�t�� ne� 1% kreditu a men�� ne� 7% kreditu (nebo men�� ne� 10 K�). pokud nen�, nejedn� se o provizi
      AND self.isPayuVypis() // toto nen� nutn�, funkce se vol� jen na PayU v�pisy. pro jistotu...
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
// Pokud p�ijde 2 a v�ce plateb ze stejn�ho ��tu se stejn�m VS ve stejn� den, se�teme je. D�le k nim p�istupujeme jako k jedn� platb�
// n�vratov� hodnota je po�adov� ��slo (identifikace) ji� existuj�c� platby
var
  i : integer;
  iPlatba : TPlatbaZVypisu;

begin
  Result := -1;
  for i := 0 to self.Platby.Count - 1 do
  begin
    iPlatba := TPlatbaZVypisu(self.Platby[i]);
    if ( (iPlatba.cisloUctu <> '160987123/0300') // '�esk� po�ta, nem� cenu na��tat historii a nav�c je to pomal�
     and (iPlatba.cisloUctu <> '') ) // ��slo ��tu protistrany nulov�, m� to tak PayU
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
