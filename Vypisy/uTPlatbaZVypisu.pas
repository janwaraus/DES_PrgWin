unit uTPlatbaZVypisu;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, IniFiles, Forms,
  Dialogs, StdCtrls, Grids, AdvObj, BaseGrid, AdvGrid, StrUtils,
  DB, ComObj, AdvEdit, DateUtils, Math, ExtCtrls,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, ZAbstractConnection, ZConnection;

{
  SysUtils, Classes, DB, StrUtils, ZAbstractRODataset, ZAbstractDataset, ZDataset,
  ZAbstractConnection, ZConnection, Dialogs, DesUtils; }


type

  TPlatbaZVypisu = class
  private
    qrAbra: TZQuery;
  public
    typZaznamu: string[3];
    cisloUctuVlastni: string[16];
    cisloUctu: string[21]; //èíslo úètu i s kódem banky za lomítkem, bez úvodních nul
    cisloUctuSNulami: string[21]; //formát vèetnì nul na zaèátku, celé èíslo úètu i s kódem banky za lomítkem
    cisloUctuKZobrazeni: string[21];
    // cisloDokladu: string[13]; // 7.4.2023 není potøeba
    castka: currency;
    kodUctovani: string[1];
    VS: string[10];
    VSOrig: string[10];
    KS: string[4];
    SS: string[10];
    valuta: string[6];
    nazevProtistrany: string[255];
    kodMeny: string[4];
    Datum: double;
    kredit, debet: boolean;
    znamyPripad: boolean;
    potrebaPotvrzeniUzivatelem: boolean;
    jePotvrzenoUzivatelem: boolean;
    zprava : string;
    vsechnyDoklady: boolean;
    pocetNacitanychPP: integer;

    problemLevel: integer;
    rozdeleniPlatby: integer;
    castecnaUhrada: integer;

    PredchoziPlatbyUcetList : TList; // list pro objekty TPredchoziPlatba
    PredchoziPlatbyVsList : TList; // list pro objekty TPredchoziPlatba
    DokladyList : TList; // list pro objekty TDoklad
    
    constructor create(castka : currency); overload;
    constructor create(gpcLine : string); overload;

    procedure loadPredchoziPlatbyPodleVS(); overload;
    procedure loadPredchoziPlatbyPodleUctu(); overload;

  published
    procedure init(pocetPredchozichPlateb : integer);
    procedure loadPredchoziPlatby(pocetPlateb : integer);
    procedure loadPredchoziPlatbyPodleUctu(pocetPlateb : integer); overload;
    procedure loadPredchoziPlatbyPodleVS(pocetPlateb : integer); overload;
    procedure loadDokladyPodleVS();
    function automatickyOpravVS() : string;
    function getVSzMinulostiByBankAccount() : string;
    function getPocetPredchozichPlatebNaStejnyVS() : integer;
    function getProcentoPredchozichPlatebNaStejnyVS() : single;
    function getPocetPredchozichPlatebZeStejnehoUctu() : integer;
    function getProcentoPredchozichPlatebZeStejnehoUctu() : single;
    procedure setZnamyPripad(popis : string);
  end;


  TPredchoziPlatba = class
  public
    VS : string[10];
    Firm_ID  : string[10];
    Castka  : Currency;
    cisloUctuSNulami : string[21];
    cisloUctuKZobrazeni : string[21];
    Datum  : double;
    FirmName : string;
    constructor create(qrAbra : TZQuery);
  end;

  
implementation

uses
  AbraEntities, DesUtils;


constructor TPlatbaZVypisu.create(castka : currency);
begin
  self.qrAbra := DesU.qrAbra;
  self.castka := abs(castka);
  if castka >= 0 then self.kredit := true else self.kredit := false;
  self.debet := not self.kredit;

  self.PredchoziPlatbyUcetList := TList.Create;
  self.PredchoziPlatbyVsList := TList.Create;
  self.DokladyList := TList.Create;
end;

constructor TPlatbaZVypisu.create(gpcLine : string);
begin
  self.qrAbra := DesU.qrAbra;

  self.typZaznamu := copy(gpcLine, 1, 3);
  self.cisloUctuVlastni := removeLeadingZeros(copy(gpcLine, 4, 16));
  self.cisloUctuSNulami := copy(gpcLine, 20, 16) + '/' + copy(gpcLine, 74, 4); //formát vèetnì nul na zaèátku
  //self.cisloDokladu := copy(gpcLine, 36, 13); // 7.4.2023 není potøeba, je to "identifikátor transakce"
  self.castka := StrToInt(removeLeadingZeros(copy(gpcLine, 49, 12))) / 100;
  self.kodUctovani := copy(gpcLine, 61, 1); // 1 debet (odchozí položka), 2 kredit (pøíchozí položka), 4 storno položky debet, 5 storno položky kreditní
  self.VS := RemoveSpaces(removeLeadingZeros(copy(gpcLine, 62, 10)));
  self.KS := copy(gpcLine, 78, 4);
  self.SS := removeLeadingZeros(copy(gpcLine, 82, 10));
  //self.valuta := copy(gpcLine, 92, 6);
  self.nazevProtistrany := Trim(copy(gpcLine, 98, 20)); // název protistrany NEBO slovní popis položky
  //self.kodMeny := copy(gpcLine, 119, 4);
  self.Datum := Str6digitsToDate(copy(gpcLine, 123, 6));

  //self.isInternetKredit := false; // nebude vlastost, protoze platba muze byt rozdelena a cast bude kredit a cast ne
  //self.isVoipKredit := false; //dtto
  self.znamyPripad := false;
  self.potrebaPotvrzeniUzivatelem := false;
  self.jePotvrzenoUzivatelem := true;
  self.pocetNacitanychPP := 5;
  self.vsechnyDoklady := false;
  self.VSOrig := self.VS;

  if (self.kodUctovani = '1') OR (self.kodUctovani = '5') then self.kredit := false; //1 je debet, 5 je storno kreditu
  if (self.kodUctovani = '2') OR (self.kodUctovani = '4') then self.kredit := true; //2 je kredit, 4 je storno debetu
  self.debet := not self.kredit;

  { pøepsáno chování VoIP kreditu, bude se brát až z toho co zbyde po napárování na faktury
  //pokud je VS ve VoIP formátu nebo je SS 8888 tak se podívám jestli je v dbZakos contract oznacen jako VoipContract
  //nedívám se tedy pro všechny záznamy kvùli šetøení èasu a možným chybám v dbZakos
  if ((copy(self.VS, 5, 1) = '9') AND (copy(self.VS, 1, 2) = '20')) OR (self.SS = '8888') then
    //self.isVoipKredit := true; //bylo, ale navíc kontrola na typ
    self.isVoipKredit := DesU.isVoipContract(self.VS);
  }

  //self.cisloUctu := removeLeadingZeros(self.cisloUctuSNulami);
  self.cisloUctu := DesU.odstranNulyZCislaUctu(self.cisloUctuSNulami);

  if kredit AND (self.cisloUctu = '2100098382/2010') then setZnamyPripad('z Fio BÚ');
  if kredit AND (self.cisloUctu = '2602372070/2010') then setZnamyPripad('z Fio Sp.Ú');
  if kredit AND (self.cisloUctu = '2800098383/2010') then setZnamyPripad('z Fioonto');
  if kredit AND (self.cisloUctu = '171336270/0300') then setZnamyPripad('z ÈSOB');
  if kredit AND (self.cisloUctu = '2107333410/2700') then setZnamyPripad('z PayU');
  if debet AND (self.cisloUctu = '2100098382/2010') then setZnamyPripad('na Fio BÚ');
  if debet AND (self.cisloUctu = '2602372070/2010') then setZnamyPripad('na Fio Sp.Ú');
  if debet AND (self.cisloUctu = '2800098383/2010') then setZnamyPripad('na Fiokonto');
  if debet AND (self.cisloUctu = '171336270/0300') then setZnamyPripad('na ÈSOB');
  if debet AND (self.cisloUctuVlastni = '2389210008000000') AND (AnsiContainsStr(nazevProtistrany, 'illing')) then setZnamyPripad('z PayU na BÚ'); //položka PayU výpisu, která znamená výbìr z PayU, obsahuje slovo "Billing"

  self.cisloUctuKZobrazeni := DesU.prevedCisloUctuNaText(self.cisloUctu);

  self.PredchoziPlatbyUcetList := TList.Create;
  self.PredchoziPlatbyVsList := TList.Create;
  self.DokladyList := TList.Create;

end;


procedure TPlatbaZVypisu.init(pocetPredchozichPlateb : integer);
begin
  loadPredchoziPlatby(pocetPredchozichPlateb);
  loadDokladyPodleVS();
end;


procedure TPlatbaZVypisu.loadPredchoziPlatby(pocetPlateb : integer);
begin
  loadPredchoziPlatbyPodleUctu(pocetPlateb);
  loadPredchoziPlatbyPodleVS(pocetPlateb);
end;

procedure TPlatbaZVypisu.loadPredchoziPlatbyPodleUctu(pocetPlateb : integer);
begin
  self.pocetNacitanychPP := pocetPlateb;
  self.PredchoziPlatbyUcetList := TList.Create;

  // posledních N plateb ze stejného èísla úètu
  if ( (self.cisloUctu <> '160987123/0300') // 'Èeská pošta, nemá cenu naèítat historii a navíc je to pomalé
   and (self.cisloUctu <> '') ) // èíslo úètu protistrany nulové, má to tak PayU
  then
  with qrAbra do begin
    SQL.Text := 'SELECT FIRST ' + IntToStr(self.pocetNacitanychPP) + ' bs.VarSymbol, bs.Firm_ID, bs.Amount, '
              + 'bs.Credit, bs.BankAccount, bs.DocDate$Date, firms.Name as FirmName '
              + 'FROM BankStatements2 bs '
              + 'JOIN Firms ON bs.Firm_ID = Firms.Id '
              + 'WHERE bs.BankAccount = ''' + self.cisloUctuSNulami  + ''' '
              // + 'AND bs.BankStatementRow_ID is null ' // toto vyluèuje v Abøe rozdìlené platby, ovšem zaèalo to zpomalovat query a tìch rozdìlených není moc, takže to nevadí vynechat
              + 'ORDER BY DocDate$Date DESC';
    Open;

    while not Eof do begin
      self.PredchoziPlatbyUcetList.Add(TPredchoziPlatba.create(qrAbra));
      Next;
    end;
    Close;
  end;
end;

procedure TPlatbaZVypisu.loadPredchoziPlatbyPodleUctu();
begin
  loadPredchoziPlatbyPodleUctu(self.pocetNacitanychPP);
end;

procedure TPlatbaZVypisu.loadPredchoziPlatbyPodleVS(pocetPlateb : integer);
begin
  self.pocetNacitanychPP := pocetPlateb;
  self.PredchoziPlatbyVsList := TList.Create;

  if StrToIntDef(self.VS, 0) = 0 then Exit; //pro platby bez (platného èíselného) VS konèíme


  // posledních N plateb na stejný VS, VS musí být platný, tedy neprázdný a èíselný
  with qrAbra do begin
    SQL.Text := 'SELECT FIRST ' + IntToStr(self.pocetNacitanychPP) + ' bs.VarSymbol, bs.Firm_ID, bs.Amount, '
              + 'bs.Credit, bs.BankAccount, bs.DocDate$Date, firms.Name as FirmName '
              + 'FROM BankStatements2 bs '
              + 'JOIN Firms ON bs.Firm_ID = Firms.Id '
              + 'WHERE bs.VarSymbol = ''' + self.VS  + ''' '
              // + 'AND bs.BankStatementRow_ID is null ' // toto vyluèuje v Abøe rozdìlené platby, ovšem zaèalo to zpomalovat query a tìch rozdìlených není moc, takže to nevadí vynechat
              + 'ORDER BY DocDate$Date DESC';
    Open;
    while not Eof do begin
      self.PredchoziPlatbyVsList.Add(TPredchoziPlatba.create(qrAbra));
      Next;
    end;
    Close;
  end;
end;

procedure TPlatbaZVypisu.loadPredchoziPlatbyPodleVS();
begin
  loadPredchoziPlatbyPodleVS(self.pocetNacitanychPP);
end;


procedure TPlatbaZVypisu.loadDokladyPodleVS();
var
  SQLiiSelect, SQLiiJenNezaplacene, SQLiiOrder,
  SQLStr : string;
begin
  self.DokladyList := TList.Create;

  if StrToIntDef(self.VS, 0) = 0 then Exit; //pro platby bez (platného èíselného) VS konèíme

  with qrAbra do begin

    // cteni z IssuedInvoices (faktury)

    SQLiiSelect := 'SELECT ii.ID FROM ISSUEDINVOICES ii'
                 + ' WHERE ii.VarSymbol = ''' + self.VS  + '''';

    SQLiiJenNezaplacene :=  ' AND (ii.LOCALAMOUNT - ii.LOCALPAIDAMOUNT - ii.LOCALCREDITAMOUNT + ii.LOCALPAIDCREDITAMOUNT) <> 0';
    SQLiiOrder := ' order by ii.DocDate$Date DESC';

    if self.vsechnyDoklady then
      SQL.Text := SQLiiSelect +  SQLiiOrder
    else
      SQL.Text := SQLiiSelect + SQLiiJenNezaplacene + SQLiiOrder;

    Open;
    while not Eof do begin
      self.DokladyList.Add(TDoklad.create(FieldByName('ID').AsString, '03')); // 03 je faktura vydaná
      Next;
    end;
    Close;

    // cteni z IssuedDInvoices (zalohove listy)
    SQLiiSelect := 'SELECT ii.ID FROM ISSUEDDINVOICES ii'
                 + ' WHERE ii.VarSymbol = ''' + self.VS  + '''';

    SQLiiJenNezaplacene :=  ' AND (ii.LOCALAMOUNT - ii.LOCALPAIDAMOUNT) <> 0';
    SQLiiOrder := ' order by ii.DocDate$Date DESC';

    if self.vsechnyDoklady then
      SQL.Text := SQLiiSelect +  SQLiiOrder
    else
      SQL.Text := SQLiiSelect + SQLiiJenNezaplacene + SQLiiOrder;

    Open;
    while not Eof do begin
      self.DokladyList.Add(TDoklad.create(FieldByName('ID').AsString, '10')); // 10 je vydaný zálohový list
      Next;
    end;
    Close;


    // cteni z ReceivedCreditNotes (Dobropis faktury pøijaté / došlé)
    SQLiiSelect := 'SELECT rcn.ID FROM RECEIVEDCREDITNOTES rcn'
                 + ' WHERE rcn.VarSymbol = ''' + self.VS  + '''';

    SQLiiJenNezaplacene :=  ' AND (rcn.LOCALAMOUNT - rcn.LOCALPAIDAMOUNT) <> 0';
    SQLiiOrder := ' order by rcn.DocDate$Date DESC';

    if self.vsechnyDoklady then
      SQL.Text := SQLiiSelect +  SQLiiOrder
    else
      SQL.Text := SQLiiSelect + SQLiiJenNezaplacene + SQLiiOrder;

    Open;
    while not Eof do begin
      self.DokladyList.Add(TDoklad.create(FieldByName('ID').AsString, '61')); // 61 je typ dokladu dobropis faktur pøijatých (DD)
      Next;
    end;
    Close;


    // když se nenajde nezaplacená faktura, zálohový list nebo dobropis, natáhnu 1 zaplacený abych mohl pøiøadit firmu
    if self.DokladyList.Count = 0 then begin

      SQL.Text := 'SELECT FIRST 1 ii.ID FROM ISSUEDINVOICES ii'
                   + ' WHERE ii.VarSymbol = ''' + self.VS  + ''''
                   + ' order by ii.DocDate$Date DESC';

      Open;
      if not Eof then begin
        self.DokladyList.Add(TDoklad.create(FieldByName('ID').AsString, '03'));
      end;
      Close;

    end;
  end;
end;

function TPlatbaZVypisu.automatickyOpravVS() : string;
var
  pouzivanyVSvMinulosti : string;
begin
  Result := ''; // nepoužíváme

  if self.problemLevel = 5 then
  begin
    pouzivanyVSvMinulosti := getVSzMinulostiByBankAccount();
    if (pouzivanyVSvMinulosti <> '') AND (pouzivanyVSvMinulosti <> self.VS) then begin
      self.VS := pouzivanyVSvMinulosti; // pùvodní VS zùsttává v atributu VSOrig
      loadPredchoziPlatbyPodleVS();
      loadDokladyPodleVS;
    end;
  end;
end;


function TPlatbaZVypisu.getVSzMinulostiByBankAccount() : string;
// Vezmu posledních 7 plateb z è. úètu a dívám se, jestli jsou mezi nimi alespoò 4 ze stejného VS. Když ano, vrátím tento VS, jinak vracím prázdný øetìzec.
begin
  Result := '';

  if ( (self.cisloUctu <> '160987123/0300') // 'Èeská pošta, nemá cenu naèítat
   and (self.cisloUctu <> '') ) // èíslo úètu protistrany nulové, má to tak PayU
  then
  with qrAbra do begin
    SQL.Text := 'SELECT varsymbol, count(*) as pocet FROM'
              + ' (SELECT FIRST 7 VarSymbol, Firm_ID FROM BankStatements2'
              + ' WHERE BankAccount = ''' + self.cisloUctuSNulami  + ''''
              // + ' AND BankStatementRow_ID is null' // // toto vyluèuje v Abøe rozdìlené platby, ovšem zaèalo to zpomalovat query a tìch rozdìlených není moc, takže to nevadí vynechat
              + ' ORDER BY DocDate$Date DESC)'

              + ' GROUP BY VARSYMBOL'
              + ' order by pocet desc';

    Open;
    if not Eof then begin
      if StrToInt(FieldByName('Pocet').AsString) > 4 then
          Result := FieldByName('Varsymbol').AsString;
    end;
    Close;
  end;
end;


function TPlatbaZVypisu.getPocetPredchozichPlatebNaStejnyVS() : integer;
var
  i : integer;
begin
  Result := 0;
  if self.PredchoziPlatbyUcetList.Count > 0 then
  begin
    for i := 0 to PredchoziPlatbyUcetList.Count - 1 do
      if TPredchoziPlatba(self.PredchoziPlatbyUcetList[i]).VS = self.VS then Inc(Result);
  end;
end;


function TPlatbaZVypisu.getProcentoPredchozichPlatebNaStejnyVS() : single;
begin
  if getPocetPredchozichPlatebNaStejnyVS() < 3 then
    Result := 0
  else
    Result := getPocetPredchozichPlatebNaStejnyVS() / PredchoziPlatbyUcetList.Count;
end;


function TPlatbaZVypisu.getPocetPredchozichPlatebZeStejnehoUctu() : integer;
var
  i : integer;
begin
  Result := 0;
  if self.PredchoziPlatbyVsList.Count > 0 then
  begin
    for i := 0 to PredchoziPlatbyVsList.Count - 1 do
      if TPredchoziPlatba(self.PredchoziPlatbyVsList[i]).cisloUctuSNulami = self.cisloUctuSNulami then Inc(Result);
  end;
end;

function TPlatbaZVypisu.getProcentoPredchozichPlatebZeStejnehoUctu() : single;
begin
  if getPocetPredchozichPlatebZeStejnehoUctu() < 3 then
    Result := 0
  else
    Result := getPocetPredchozichPlatebZeStejnehoUctu() / PredchoziPlatbyVsList.Count;
end;

procedure TPlatbaZVypisu.setZnamyPripad(popis : string);
begin
  self.nazevProtistrany := popis;
  self.znamyPripad := true;
end;


{** class TPredchoziPlatba **}

constructor TPredchoziPlatba.create(qrAbra : TZQuery);
begin
 with qrAbra do begin
  self.VS := RemoveSpaces(FieldByName('VarSymbol').AsString);
  self.Firm_ID := FieldByName('Firm_ID').AsString;
  self.castka := FieldByName('Amount').AsCurrency;
  self.cisloUctuSNulami := FieldByName('BankAccount').AsString;
  self.Datum := FieldByName('DocDate$Date').asFloat;
  self.FirmName := FieldByName('FirmName').AsString;

  self.cisloUctuKZobrazeni := DesU.prevedCisloUctuNaText(DesU.odstranNulyZCislaUctu(cisloUctuSNulami));

  if (FieldByName('Credit').AsString = 'N') then
    self.Castka := - self.Castka;

 end;
end;


end.
