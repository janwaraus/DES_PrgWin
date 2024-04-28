unit uTBankUcet;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, IniFiles, Forms,
  Dialogs, StdCtrls, Grids, AdvObj, BaseGrid, AdvGrid, StrUtils, RegularExpressions,
  DB, ComObj, AdvEdit, DateUtils, Math, ExtCtrls,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, ZAbstractConnection, ZConnection,
  AbraEntities, DesUtils;

type

  TBankUcet = class

  public
    abraBankaccount : TAbraBankaccount;


    pocetVypisu,
    poradoveCisloMaxVypisu,
    extPoradoveCisloMaxVypisu,
    aktualniCisloVypisu : integer;

    cisloUctu,
    cisloUctuBezLomitka,
    nazevUctu,
    rokVypisu,
    infoSouhrnOVypisech,
    gpcSoubor,
    pdfSoubor,
    hledanyGpcSoubor : string;

    constructor create(cisloUctu, rokVypisu : string);

  published
    function stahniVypis(cisloVypisuKeStazeni : integer; formaVypisu : string = 'gpc') : TDesResult;
    function stahniAktualnePotrebnyVypis() : TDesResult;

  end;

implementation


constructor TBankUcet.create(cisloUctu, rokVypisu : string);
begin
  self.cisloUctu := cisloUctu;
  self.cisloUctuBezLomitka := Copy(cisloUctu, 1, Pos('/', cisloUctu) - 1);

  self.rokVypisu := rokVypisu;

  abraBankAccount := TAbraBankaccount.create();
  abraBankaccount.loadByNumber(cisloUctu);
  nazevUctu := abraBankaccount.Name;

  pocetVypisu := abraBankaccount.getPocetVypisu(rokVypisu);
  poradoveCisloMaxVypisu := abraBankaccount.getPoradoveCisloMaxVypisu(rokVypisu);
  extPoradoveCisloMaxVypisu := abraBankaccount.getExtPoradoveCisloMaxVypisu(rokVypisu);

  aktualniCisloVypisu := poradoveCisloMaxVypisu + 1;

  if self.cisloUctu = '2100098382/2010' then begin
    hledanyGpcSoubor := 'Vypis_z_uctu-2100098382_' + rokVypisu + '*-' + IntToStr(aktualniCisloVypisu) + '.gpc';
  end;
  if self.cisloUctu = '2602372070/2010' then begin
    hledanyGpcSoubor := 'Vypis_z_uctu-2602372070_' + rokVypisu + '*-' + IntToStr(aktualniCisloVypisu) + '.gpc';
  end;
  if self.cisloUctu = '2800098383/2010' then begin
    hledanyGpcSoubor := 'Vypis_z_uctu-2800098383_' + rokVypisu + '*-' + IntToStr(aktualniCisloVypisu) + '.gpc';
  end;

  gpcSoubor := FindInFolder(DesU.GPC_PATH, hledanyGpcSoubor, true);

  infoSouhrnOVypisech := format('%d výpisù v roce %s. Max. èíslo výpisu %d, externí èíslo %d, datum posledního výpisu %s', [  //TODO pro PayU vyhodit ext èíslo
    pocetVypisu, rokVypisu, poradoveCisloMaxVypisu, extPoradoveCisloMaxVypisu,
    DateToStr(abraBankaccount.getDatumMaxVypisu(rokVypisu))
    ]);

end;


function TBankUcet.stahniVypis(cisloVypisuKeStazeni : integer; formaVypisu : string = 'gpc') : TDesResult;
var
  wgGpcResponse,
  wgPdfResponse : TDesResult;
  wgGpcResponseContent,
  fioEndpoint,
  gpcDateStart,
  gpcDateEnd,
  gpcSubdir : string;
begin
  try
    // Fio
    if self.cisloUctu = '2100098382/2010' then begin
      fioEndpoint := 'https://www.fio.cz/ib_api/rest/by-id/' + DesU.getIniValue('BankingTokens', 'TokenFio') + '/' + rokVypisu + '/';
      gpcSubdir := 'Fio';
    end;

    // Fio spoøící
    if self.cisloUctu = '2602372070/2010' then begin
      fioEndpoint := 'https://www.fio.cz/ib_api/rest/by-id/' + DesU.getIniValue('BankingTokens', 'TokenFioSporici') + '/' + rokVypisu + '/';
      gpcSubdir := 'FioSporici';
    end;

    // Fiokonto
    if self.cisloUctu = '2800098383/2010' then begin
      fioEndpoint := 'https://www.fio.cz/ib_api/rest/by-id/' + DesU.getIniValue('BankingTokens', 'TokenFiokonto') + '/' + rokVypisu + '/';
      gpcSubdir := 'FioKonto';
    end;

    // gpcSubdir := '000'; //TODO odstranit po testu

    wgGpcResponse := DesU.webHttpsGet(fioEndpoint + IntToStr(cisloVypisuKeStazeni) + '/transactions.gpc');

    if wgGpcResponse.isOk then begin
      wgGpcResponseContent := wgGpcResponse.Messg;
      gpcDateStart := '20' + Copy(wgGpcResponseContent, 44, 2) + Copy(wgGpcResponseContent, 42, 2) + Copy(wgGpcResponseContent, 40, 2);
      gpcDateEnd := '20' + Copy(wgGpcResponseContent, 113, 2) + Copy(wgGpcResponseContent, 111, 2) + Copy(wgGpcResponseContent, 109, 2);
      gpcSoubor := Format('%s%s\Vypis_z_uctu-%s_%s-%s_cislo-%s.gpc', [DesU.GPC_PATH, gpcSubdir, self.cisloUctuBezLomitka, gpcDateStart, gpcDateEnd, IntToStr(cisloVypisuKeStazeni)]);
      writeToFile(gpcSoubor, wgGpcResponseContent);
      Result := TDesResult.create('ok', Format('Výpis è. %s úètu %s byl stažen do souboru %s', [IntToStr(cisloVypisuKeStazeni), nazevUctu, gpcSoubor]));

      sleep (2000);
      wgPdfResponse := DesU.webHttpsGet(fioEndpoint + IntToStr(cisloVypisuKeStazeni) + '/transactions.pdf');
      if wgPdfResponse.isOk then begin
        pdfSoubor := Format('%s%s\Vypis_z_uctu-%s_%s-%s_cislo-%s.pdf', [DesU.GPC_PATH, gpcSubdir, self.cisloUctuBezLomitka, gpcDateStart, gpcDateEnd, IntToStr(cisloVypisuKeStazeni)]);
        writeToFile(pdfSoubor, wgPdfResponse.Messg);
        Result.Messg := Result.Messg + sLineBreak + Format('Stažen také soubor %s', [pdfSoubor]);
      end else begin
        Result.Messg := Result.Messg + sLineBreak + Format('Pdf soubor nestažen, chyba: %s %s', [wgPdfResponse.Code, wgPdfResponse.Messg]);
      end;

    end else begin
      if Pos('<errorCode>21</errorCode>', wgGpcResponse.Messg) > 0 then
        wgGpcResponse.Messg := 'výpis neexistuje';
      Result := wgGpcResponse;
    end;

  except on E: exception do
    Result := TDesResult.create('err', Format('Chyba stahování výpisu è. %s úètu %s.' + sLineBreak + 'Chyba: %s',
      [IntToStr(cisloVypisuKeStazeni), nazevUctu, E.Message]));
  end;

end;


function TBankUcet.stahniAktualnePotrebnyVypis() : TDesResult;
begin
  Result := self.stahniVypis(self.aktualniCisloVypisu);
  //Result := self.stahniVypis(self.aktualniCisloVypisu + 1);
  //Result := self.stahniVypis(self.aktualniCisloVypisu + 2);
end;






end.
