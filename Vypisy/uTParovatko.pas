unit uTParovatko;

interface

uses
  SysUtils, Variants, Classes, Controls, StrUtils,
  Windows, Messages, Dialogs, Forms,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, ZAbstractConnection, ZConnection,  
  uTVypis, uTPlatbaZVypisu, AbraEntities;


type

  TPlatbaDokladPar = class
  public
    Platba : TPlatbaZVypisu;
    Doklad : TDoklad;
    Doklad_ID : string[10];
    CastkaPouzita : currency;
    Popis : string;
    vazbaNaDoklad : boolean;
    pdparTyp : string; // speci·lnÌ typy - "VoipKredit" nebo "InternetKredit"
  end;

  TParovatko = class
  public
    Vypis: TVypis;
    qrAbra: TZQuery;
    AbraOLE: variant;
    listPlatbaDokladPar : TList; // list vöech p·r˘ platba-doklad (objekty z TPlatbaDokladPar)
    constructor create(Vypis: TVypis);
  published
    procedure sparujPlatbu(Platba : TPlatbaZVypisu);
    procedure odparujPlatbu(currPlatba : TPlatbaZVypisu);
    procedure vytvorPDPar(Platba : TPlatbaZVypisu; Doklad : TDoklad;
                Castka: currency; popis : string; vazbaNaDoklad : boolean; pdparTyp : string = '');
    function zapisDoAbry() : string;
    function getUzSparovano(Doklad_ID : string) : currency;
    function getPDParyAsText() : string;
    function getPDParyPlatbyAsText(currPlatba : TPlatbaZVypisu) : string;
    function getPDPar(currPlatba : TPlatbaZVypisu; currDoklad_ID: string) : TPlatbaDokladPar;
    function getFirmIdKdyzNeniDoklad(currPlatba : TPlatbaZVypisu) : string;
  end;




implementation

uses
  DesUtils, AArray, Superobject;


constructor TParovatko.create(Vypis: TVypis);
begin
  self.Vypis := Vypis;

  self.qrAbra := DesU.qrAbra;
  self.listPlatbaDokladPar := TList.Create();
end;


procedure TParovatko.sparujPlatbu(Platba : TPlatbaZVypisu);

  procedure zpracujPreplatek(Platba : TPlatbaZVypisu; Doklad : TDoklad; Castka: currency);
  var
    vsPatriKreditniVoipSmlouve, vsPatriKreditniSmlouve : boolean;
  begin
    vsPatriKreditniVoipSmlouve := false;
    vsPatriKreditniSmlouve := false;
    Platba.potrebaPotvrzeniUzivatelem := false;

    if ((copy(Platba.VS, 5, 1) = '9') AND (copy(Platba.VS, 1, 2) = '20')) OR (Platba.SS = '8888') then
    { kreditnÌ smlouva m· na 5. mÌstÏ ËÌslici 9, tedy VS je 20xx9xxxxx; nebo m· SS 8888
      pokud je VS ve VoIP form·tu nebo je SS 8888 tak se podÌv·m jestli je v dbZakos contract oznacen jako VoipContract
      nedÌv·m se tedy pro vöechny z·znamy kv˘li öet¯enÌ Ëasu a moûn˝m chyb·m v dbZakos }
      vsPatriKreditniVoipSmlouve := DesU.isVoipKreditContract(Platba.VS);

    if not vsPatriKreditniVoipSmlouve then
      vsPatriKreditniSmlouve := DesU.isCreditContract(Platba.VS);

    if vsPatriKreditniVoipSmlouve OR vsPatriKreditniSmlouve then
      Platba.potrebaPotvrzeniUzivatelem := true; // v 8. sloupci StringGridu s platbami se nakonec objevÌ checkbox

    if Platba.potrebaPotvrzeniUzivatelem AND Platba.jePotvrzenoUzivatelem then
    begin
      if vsPatriKreditniVoipSmlouve then
      begin
        vytvorPDPar(Platba, nil, Castka, 'kredit VoIP |', false, 'VoipKredit');
        Platba.zprava := 'VoIP kredit ' + FloatToStr(Castka) + ' KË';
        Platba.problemLevel := 3;
      end
      else if vsPatriKreditniSmlouve then
      begin
        vytvorPDPar(Platba, nil, Castka, 'kredit Internet |' , false, 'InternetKredit');
        Platba.zprava := 'Internet kredit ' + FloatToStr(Castka) + ' KË';
        Platba.problemLevel := 3;
      end;
    end
    else
    begin
      //if Platba.getProcentoPredchozichPlatebZeStejnehoUctu() > 0.5 then begin //zmÏna z nadpoloviËnÌ vÏtöiny na vÌce jak 3 platby
      if Platba.getPocetPredchozichPlatebZeStejnehoUctu() >= 3 then begin
        vytvorPDPar(Platba, Doklad, Castka, 'p¯epl. | ' + Platba.VS + ' |' , false);
        Platba.zprava := 'p¯epl. ' + FloatToStr(Castka) + ' KË';
        Platba.problemLevel := 1;
      end else begin
        vytvorPDPar(Platba, Doklad, Castka, 'p¯epl. | ' + Platba.VS + ' |', false);
        Platba.zprava := 'nezn·m˝ p¯epl. ' + FloatToStr(Castka) + ' KË';
        Platba.problemLevel := 5;
      end;
    end;

  end;

var
  i : integer;
  nezaplaceneDoklady : TList;
  iDoklad : TDoklad;
  zbyvaCastka,
  kNaparovani : currency;
begin

  iDoklad := nil;
  Platba.rozdeleniPlatby := 0;
  Platba.castecnaUhrada := 0;

  self.odparujPlatbu(Platba); // vûdy smaûeme vöechny poloûky listPlatbaDokladPar, kde figuruje vstupnÌ Platba (a zaËneme pak p·rovat nanovo)

  if Platba.DokladyList.Count > 0 then
    iDoklad := TDoklad(Platba.DokladyList[0]); //pokud je alespon 1 doklad, priradime si ho pro debety a kredity bez nezaplacenych dokladu

  if Platba.debet then
  begin //platba je debet (mÌnusov·)
      Platba.zprava := 'debet';
      Platba.problemLevel := 0;
      vytvorPDPar(Platba, iDoklad, Platba.Castka, '', false);
  end   // end platba je debet

  else

  begin //platba je kredit (plusov·)

    // vyrobÌm si list jen nezaplacen˝ch doklad˘
    nezaplaceneDoklady := TList.Create;
    for i := 0 to Platba.DokladyList.Count - 1 do
      if TDoklad(Platba.DokladyList[i]).CastkaNezaplaceno <> 0 then
        nezaplaceneDoklady.Add(Platba.DokladyList[i]);

    // pokud nenÌ û·dn˝ nezaplacen˝
    if (nezaplaceneDoklady.Count = 0) then
    begin
      if Platba.znamyPripad then
      begin
        vytvorPDPar(Platba, iDoklad, Platba.Castka, '', false);
        Platba.zprava := 'zn·m˝ kredit';
        Platba.problemLevel := 0;
      end else begin
        zpracujPreplatek(Platba, iDoklad, Platba.Castka);
      end;
      Exit;
    end;

    zbyvaCastka := Platba.Castka;

    for i := nezaplaceneDoklady.Count - 1 downto 0 do
    // begin existujÌ nezaplacenÈ doklady
    begin
      iDoklad := TDoklad(nezaplaceneDoklady[i]);
      kNaparovani := iDoklad.CastkaNezaplaceno - getUzSparovano(iDoklad.ID); // getUzSparovano se podÌv· na vöechny pdP·ry a zjistÌ, zda uû je iDoklad hrazzen˝ a jakou Ë·stkou

      if (kNaparovani <> 0) then
      begin
        if (kNaparovani = zbyvaCastka) then
        begin
          vytvorPDPar(Platba, iDoklad, zbyvaCastka, '', true); //p¯esnÏ
          Platba.zprava := 'p¯esnÏ';
          if Platba.rozdeleniPlatby > 0 then
            Platba.problemLevel := 0 //bylo 1
          else
            Platba.problemLevel := 0;
          zbyvaCastka := 0;
          Break;
        end;

        if (kNaparovani > zbyvaCastka) AND not(iDoklad.DocumentType = '10') then //Ë·steËn· ˙hrada, doklad nesmÌ b˝t z·lohov˝m listem
        begin
          vytvorPDPar(Platba, iDoklad, zbyvaCastka, 'Ë·st. ' + floattostr(zbyvaCastka) + ' z ' + floattostr(kNaparovani) + ' KË |', true);
          Platba.zprava := 'Ë·steËn· ˙hrada';
          Platba.castecnaUhrada := 1;
          Platba.problemLevel := 1;
          zbyvaCastka := 0;
          Break;
        end;

        if (kNaparovani < zbyvaCastka) then
        begin
          //vytvorPDPar(Platba, iDoklad, kNaparovani, 'p¯i¯azeno ' + floattostr(kNaparovani), true); // dÏlenÌ
          vytvorPDPar(Platba, iDoklad, kNaparovani, '', true); // dÏlenÌ
          zbyvaCastka := zbyvaCastka - kNaparovani;
          Inc(Platba.rozdeleniPlatby);
        end;
      end;
    end;
    // end existujÌ nezaplacenÈ doklady

    if (zbyvaCastka > 0) then
    begin
      zpracujPreplatek(Platba, iDoklad, zbyvaCastka);
    end;

    if (Platba.getPocetPredchozichPlatebZeStejnehoUctu() = 0)
      AND (Platba.PredchoziPlatbyVsList.Count > 3) then
    begin
      Platba.zprava := 'nezn·m˝ ˙Ëet - ' + Platba.zprava;
      Platba.problemLevel := 2;
    end;

  end;
  // end platba je kredit

end;


procedure TParovatko.odparujPlatbu(currPlatba : TPlatbaZVypisu);
// smaûu vöechny poloûky listPlatbaDokladPar, kde Platba je vtupnÌ currPlatba
var
  i : integer;
  iPDPar : TPlatbaDokladPar;
begin

  for i := listPlatbaDokladPar.Count - 1 downto 0 do
  begin
    iPDPar := TPlatbaDokladPar(listPlatbaDokladPar[i]);
    if iPDPar.Platba = currPlatba then
      listPlatbaDokladPar.Delete(i);
  end;
end;


procedure TParovatko.vytvorPDPar(Platba : TPlatbaZVypisu; Doklad : TDoklad;
            Castka: currency; popis : string; vazbaNaDoklad : boolean; pdparTyp : string = '');
var
  iPDPar : TPlatbaDokladPar;
begin
  iPDPar := TPlatbaDokladPar.Create();
  // spr·vnÏ by bylo vytvo¯it po¯·dn˝ konstruktor a nep¯i¯azovat hodnoty do iPDPar v tÈto procedu¯e
  iPDPar.Platba := Platba;
  iPDPar.Doklad := Doklad;
  if assigned(iPDPar.Doklad) then
    iPDPar.Doklad_ID := iPDPar.Doklad.ID
  else
    iPDPar.Doklad_ID := '';
  iPDPar.CastkaPouzita := Castka;
  iPDPar.Popis := Popis;
  iPDPar.vazbaNaDoklad := vazbaNaDoklad;
  iPDPar.pdparTyp := pdparTyp;
  self.listPlatbaDokladPar.Add(iPDPar);
end;


function TParovatko.zapisDoAbry() : string;
var
  i, j : integer;
  iPDPar : TPlatbaDokladPar;
  boAA, boRowAA: TAArray;
  newID, faId : string;
begin

  if (listPlatbaDokladPar.Count = 0) then Exit;

  Result := 'Z·pis pomocÌ AArray metoda ' +  DesU.abraDefaultCommMethod + ' v˝pisu pro ˙Ëet ' + self.Vypis.abraBankaccount.name + '.';


  boAA := TAArray.Create;
  boAA['DocQueue_ID'] := self.Vypis.abraBankaccount.bankStatementDocqueueId;
  boAA['Period_ID'] := DesU.getAbraPeriodId(self.Vypis.Datum);
  boAA['BankAccount_ID'] := self.Vypis.abraBankaccount.id;
  boAA['ExternalNumber'] := self.Vypis.PoradoveCislo;
  boAA['DocDate$DATE'] := self.Vypis.Datum;
  //boAA['CreatedAt$DATE'] := IntToStr(Trunc(Date));


  for i := 0 to listPlatbaDokladPar.Count - 1 do
  begin
    iPDPar := TPlatbaDokladPar(listPlatbaDokladPar[i]);

    boRowAA := boAA.addRow();
    boRowAA['Amount'] := iPDPar.CastkaPouzita;
    //boRowAA['Credit'] := IfThen(iPDPar.Platba.Kredit,'1','0'); //pro WebApi nefungovalo dob¯e
    boRowAA['Credit'] := iPDPar.Platba.Kredit;
    boRowAA['BankAccount'] := iPDPar.Platba.cisloUctuSNulami;
    boRowAA['Text'] := Trim(iPDPar.popis + ' ' + iPDPar.Platba.nazevProtistrany);
    boRowAA['SpecSymbol'] := iPDPar.Platba.SS;
    boRowAA['DocDate$DATE'] := iPDPar.Platba.Datum;
    boRowAA['AccDate$DATE'] := iPDPar.Platba.Datum;
    boRowAA['Division_ID'] := AbraEnt.getDivisionId;
    boRowAA['Currency_ID'] := AbraEnt.getCurrencyId;


    if Assigned(iPDPar.Doklad) then
      if iPDPar.vazbaNaDoklad then //Doklad vyplnime jen pokud chceme vazbu (vazbaNaDoklad je true). Doklad m·me naËten˝ i v situaci kdy vazbu nechceme - kv˘li Firm_ID
      begin
        boRowAA['PDocumentType'] := iPDPar.Doklad.DocumentType; // nap¯. 03, 10, 61
        boRowAA['PDocument_ID'] := iPDPar.Doklad.ID;
      end
      else
      begin
        boRowAA['Firm_ID'] := iPDPar.Doklad.Firm_ID;
      end
    else //nenÌ Assigned(iPDPar.Doklad)
      if not(iPDPar.pdparTyp = 'VoipKredit') AND not(iPDPar.pdparTyp = 'InternetKredit') then  //tyto podmÌnky nejsou nutnÈ, Abra by Firm_ID p¯ebila hodnotou z fa, nicmÈnÏ s VoipKredit a InternetKredit pracujeme v algoritmu nÌûe a p·rujeme na faktury, tedy Firm_ID nebudeme zasÌlat
        boRowAA['Firm_ID'] := getFirmIdKdyzNeniDoklad(iPDPar.Platba); //kdyû m· text platby na zaË·tku Abra kÛd, najde se firma. Jinak se d· '3Y90000101' (DES), aby tam nebyl default "drobn˝ n·kup"

    if iPDPar.pdparTyp = 'VoipKredit' then
    begin
      faId := DesU.vytvorFaZaVoipKredit(iPDPar.Platba.VS, iPDPar.CastkaPouzita, iPDPar.Platba.Datum);
      if faId = '' then
        //pokud nenajdeme podle VS firmu, zapÌöeme VS - nemÏlo by se st·t, ale pro jistotu
        boRowAA['VarSymbol'] := iPDPar.Platba.VS
      else begin
        //byla vytvo¯ena fa a tu teÔ p¯ipojÌme. VS pak abra automaticky doplnÌ
        boRowAA['PDocumentType'] := '03'; // je to vûdy faktura
        boRowAA['PDocument_ID'] := faId;
      end;
    end;

    if iPDPar.pdparTyp = 'InternetKredit' then
    begin
      faId := DesU.vytvorFaZaInternetKredit(iPDPar.Platba.VS, iPDPar.CastkaPouzita, iPDPar.Platba.Datum);
      if faId = '' then
        //pokud nenajdeme podle VS firmu, zapÌöeme VS - nemÏlo by se st·t, ale pro jistotu
        boRowAA['VarSymbol'] := iPDPar.Platba.VS
      else begin
        //byla vytvo¯ena fa a tu teÔ p¯ipojÌme. VS pak abra automaticky doplnÌ
        boRowAA['PDocumentType'] := '03'; // je to vûdy faktura
        boRowAA['PDocument_ID'] := faId;
      end;
    end;


    if iPDPar.Platba.Debet then
      boRowAA['VarSymbol'] := iPDPar.Platba.VS; //pro debety aby vûdy z˘stal VS

    if (iPDPar.Platba.cisloUctuVlastni = '2389210008000000') AND iPDPar.Platba.kredit then begin //PayU platba, ruöÌme penÌze na cestÏ
      DesU.zrusPenizeNaCeste(iPDPar.Platba.VS);
    end;

  end;

  try begin
    newId := DesU.abraBoCreate(boAA, 'bankstatement');
    Result := Result + ' »Ìslo novÈho v˝pisu je ' + newID;
    // DesU.abraOLELogout; // 29.3. p¯esunuto do DesU.abraBoCreateOLE
  end;
  except on E: exception do
    begin
      Application.MessageBox(PChar('Problem ' + ^M + E.Message), 'Vytv·¯enÌ v˝pisu');
      Result := 'Chyba p¯i vytv·¯enÌ v˝pisu';
    end;
  end;

end;


function TParovatko.getUzSparovano(Doklad_ID : string) : currency;
var
  i : integer;
  iPDPar : TPlatbaDokladPar;
begin
  Result := 0;

  if listPlatbaDokladPar.Count > 0 then
    for i := 0 to listPlatbaDokladPar.Count - 1 do
    begin
      iPDPar := TPlatbaDokladPar(listPlatbaDokladPar[i]);
      if Assigned(iPDPar.Doklad) AND (iPDPar.vazbaNaDoklad) then
        if (iPDPar.Doklad.ID = Doklad_ID)  then
          Result := Result + iPDPar.CastkaPouzita;
    end;
end;


function TParovatko.getPDParyAsText() : string;
var
  i : integer;
  iPDPar : TPlatbaDokladPar;
begin
  Result := '';

  if listPlatbaDokladPar.Count = 0 then exit;

  for i := 0 to listPlatbaDokladPar.Count - 1 do
  begin
    iPDPar := TPlatbaDokladPar(listPlatbaDokladPar[i]);
    Result := Result + 'VS: ' + iPDPar.Platba.VS + ' | ';
    if iPDPar.vazbaNaDoklad AND Assigned(iPDPar.Doklad) then
      Result := Result + 'Na doklad ' + iPDPar.Doklad.CisloDokladu + ' nap·rov·no ' + FloatToStr(iPDPar.CastkaPouzita) + ' KË | ';
    Result := Result + iPDPar.Popis + sLineBreak;
  end;
end;

function TParovatko.getPDParyPlatbyAsText(currPlatba : TPlatbaZVypisu) : string;
var
  i : integer;
  iPDPar : TPlatbaDokladPar;
begin
  Result := '';
  if listPlatbaDokladPar.Count = 0 then exit;

  for i := 0 to listPlatbaDokladPar.Count - 1 do
  begin
    iPDPar := TPlatbaDokladPar(listPlatbaDokladPar[i]);
    if iPDPar.Platba = currPlatba then begin
      Result := Result + 'VS: ' + iPDPar.Platba.VS + ' | ';
      if iPDPar.vazbaNaDoklad AND Assigned(iPDPar.Doklad) then
        Result := Result + 'Na doklad ' + iPDPar.Doklad.CisloDokladu + ' nap·rov·no ' + FloatToStr(iPDPar.CastkaPouzita) + ' KË | ';
      Result := Result + iPDPar.Popis;
      Result := Result.Trim.TrimRight(['|']) + sLineBreak;
    end;
  end;
end;

function TParovatko.getPDPar(currPlatba : TPlatbaZVypisu; currDoklad_ID: string) : TPlatbaDokladPar;
var
  i : integer;
  iPDPar : TPlatbaDokladPar;
begin
  Result := nil;
  if listPlatbaDokladPar.Count = 0 then exit;

  for i := 0 to listPlatbaDokladPar.Count - 1 do
  begin
    iPDPar := TPlatbaDokladPar(listPlatbaDokladPar[i]);
    if (iPDPar.Platba = currPlatba) and (iPDPar.Doklad_ID = currDoklad_ID) and
    (iPDPar.vazbaNaDoklad) then
      Result := iPDPar;
  end;
end;

function TParovatko.getFirmIdKdyzNeniDoklad(currPlatba : TPlatbaZVypisu) : string;
var
  platbaText, prvni2pismena : string;
begin
  platbaText := currPlatba.nazevProtistrany;
  prvni2pismena := copy(platbaText, 1, 2);
  if (prvni2pismena = 'T0') or (prvni2pismena = 'M0') or (prvni2pismena = 'N0') then
  begin
    Result := DesU.getFirmIdByCode(Copy(platbaText, 1, Pos(' ', platbaText + ' ') - 1));
  end
  else
    Result := '3Y90000101'; //aù je firma DES;
end;


end.
