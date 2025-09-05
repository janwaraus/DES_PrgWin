unit VypisyMain;
interface
uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, IniFiles, Forms,
  Dialogs, StdCtrls, Grids, AdvObj, BaseGrid, AdvGrid, StrUtils, IOUtils,
  DB, ComObj, AdvEdit, DateUtils, Math, ExtCtrls,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, ZAbstractConnection, ZConnection,
  uTVypis, uTPlatbaZVypisu, uTParovatko, uTBankUcet, AdvUtil;
type
  TfmMain = class(TForm)
    btnNacti: TButton;
    btnZapisDoAbry: TButton;
    Memo1: TMemo;
    asgMain: TAdvStringGrid;
    NactiGpcDialog: TOpenDialog;
    asgPredchoziPlatby: TAdvStringGrid;
    asgPredchoziPlatbyVs: TAdvStringGrid;
    btnSparujPlatby: TButton;
    editPocetPredchPlateb: TEdit;
    btnReconnect: TButton;
    chbZobrazitBezproblemove: TCheckBox;
    lblHlavicka: TLabel;
    chbZobrazitDebety: TCheckBox;
    chbZobrazitStandardni: TCheckBox;
    lblPrechoziPlatbySVs: TLabel;
    lblPrechoziPlatbyZUctu: TLabel;
    Memo2: TMemo;
    btnShowPrirazeniPnpForm: TButton;
    btnVypisFio: TButton;
    lblVypisFioGpc: TLabel;
    lblVypisFioInfo: TLabel;
    btnVypisFioSporici: TButton;
    btnVypisCsob: TButton;
    btnVypisPayU: TButton;
    lblVypisFioSporiciGpc: TLabel;
    lblVypisFioSporiciInfo: TLabel;
    lblVypisCsobInfo: TLabel;
    lblVypisCsobGpc: TLabel;
    btnZavritVypis: TButton;
    btnCustomers: TButton;
    btnHledej: TButton;
    editHledej: TEdit;
    lblVypisPayuGpc: TLabel;
    lblVypisPayuInfo: TLabel;
    pnBottom: TPanel;
    lblNalezeneDoklady: TLabel;
    asgNalezeneDoklady: TAdvStringGrid;
    chbVsechnyDoklady: TCheckBox;
    lblVypisFiokontoGpc: TLabel;
    lblVypisFiokontoInfo: TLabel;
    btnVypisFiokonto: TButton;
    lblVypisOverview: TLabel;
    lblPomocPrg: TLabel;
    lblZobrazit: TLabel;
    lblHlavickaVpravo: TLabel;
    btnStahniVypisy: TButton;
    lblVypisRbInfo: TLabel;
    lblVypisRbGpc: TLabel;
    btnVypisRb: TButton;
    procedure btnNactiClick(Sender: TObject);
    procedure btnZapisDoAbryClick(Sender: TObject);
    procedure asgMainGetAlignment(Sender: TObject; ARow, ACol: Integer;
              var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asgMainClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure chbVsechnyDokladyClick(Sender: TObject);
    procedure btnSparujPlatbyClick(Sender: TObject);
    procedure asgMainCellsChanged(Sender: TObject; R: TRect);
    procedure asgNalezeneDokladyGetAlignment(Sender: TObject; ARow,
      ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asgPredchoziPlatbyGetAlignment(Sender: TObject; ARow,
      ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asgPredchoziPlatbyVsGetAlignment(Sender: TObject; ARow,
      ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure btnReconnectClick(Sender: TObject);
    procedure btnHledejClick(Sender: TObject);
    procedure asgPredchoziPlatbyButtonClick(Sender: TObject; ACol,
      ARow: Integer);
    procedure chbZobrazitBezproblemoveClick(Sender: TObject);
    procedure chbZobrazitDebetyClick(Sender: TObject);
    procedure asgMainCanEditCell(Sender: TObject; ARow, ACol: Integer;
      var CanEdit: Boolean);
    procedure asgMainKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure chbZobrazitStandardniClick(Sender: TObject);
    procedure asgMainGetEditorType(Sender: TObject; ACol, ARow: Integer;
      var AEditor: TEditorType);
    procedure asgMainGetCellColor(Sender: TObject; ARow, ACol: Integer;
      AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
    procedure btnShowPrirazeniPnpFormClick(Sender: TObject);
    procedure asgMainButtonClick(Sender: TObject; ACol, ARow: Integer);
    procedure btnVypisFioClick(Sender: TObject);
    procedure btnVypisFioSporiciClick(Sender: TObject);
    procedure btnVypisFiokontoClick(Sender: TObject);
    procedure btnVypisCsobClick(Sender: TObject);
    procedure btnVypisRbClick(Sender: TObject);
    procedure btnZavritVypisClick(Sender: TObject);
    procedure btnCustomersClick(Sender: TObject);
    procedure asgMainCheckBoxClick(Sender: TObject; ACol, ARow: Integer;
      State: Boolean);
    procedure btnVypisPayUClick(Sender: TObject);
    procedure asgMainCellValidate(Sender: TObject; ACol, ARow: Integer;
      var Value: string; var Valid: Boolean);
    procedure btnStahniVypisyClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);

  public
    procedure nactiGpc(GpcFilename : string);
    procedure stahniVypisy;
    procedure vyplnNacitaciButtony;
    procedure vyplnPrichoziPlatby;
    procedure vyplnPredchoziPlatby;
    procedure vyplnDoklady;
    procedure vyplnVysledekParovaniPP(i : integer);
    procedure sparujPrichoziPlatbu(i : integer);
    procedure sparujVsechnyPrichoziPlatby;
    procedure urciCurrPlatbaZVypisu;
    procedure filtrujZobrazeniPlateb;
    procedure provedAkcePoZmeneVS;
    procedure Zprava(TextZpravy : string);
    procedure zpravaRozdilCasu(cas01, cas02 : double; TextZpravy : string);
  end;

var
  fmMain : TfmMain;
  Vypis : TVypis;
  currPlatbaZVypisu : TPlatbaZVypisu;
  Parovatko : TParovatko;
  bankUcetFio,
  bankUcetFioSporici,
  bankUcetFiokonto,
  bankUcetCsob,
  bankUcetPayU : TBankUcet;
  prvniZobrazeni : boolean;

implementation
uses
  AbraEntities, DesUtils, Superobject, Customers, PrirazeniPNP;
{$R *.dfm}

procedure TfmMain.FormActivate(Sender: TObject);
begin
  if prvniZobrazeni then begin
    //Memo1.Lines.Add('Pøipojení k databázím...');
    //DesU.desUtilsInit;
    Memo1.Lines.Add('Naètení dat...');
    vyplnNacitaciButtony;

    if DesU.appMode >= 3 then
    begin
      btnReconnect.Visible := true;
      btnSparujPlatby.Visible := true;
      memo2.Visible := true;
    end;
    //fmPrirazeniPnp.Show;  //pøi programování kvùli zrychlení práce (DEVEL)
    Memo1.Lines.Add('Kontrola dokladù s prázným VS...');
    DesU.existujeVAbreDokladSPrazdnymVs(); //TODO zaktivovat
    { *** stažení výpisù po startu programu, zakomentováno 06-2024 ***
    Memo1.Lines.Add('Stažení Fio výpisù...');
    stahniVypisy;
    }
    prvniZobrazeni := false;
  end;
end;

procedure TfmMain.FormShow(Sender: TObject);
begin
  //DesU.desUtilsInit('');
  //asgMain.CheckFalse := '0';
  //asgMain.CheckTrue := '1';
  lblHlavicka.Caption := '';
  lblVypisOverview.Caption := '';
  lblHlavickaVpravo.Caption := '';
  Memo1.Lines.Add('Start programu...');
  prvniZobrazeni := true;

end;

procedure TfmMain.vyplnNacitaciButtony;
var
  rokVypisu : string;

  maxCisloVypisu, i1, i2, i3 : integer;
  posledniDatum : double;
  nalezenyGpcSoubor, hledanyGpcSoubor,
  hledanePayuDatumVypisu : string;
  abraBankaccount : TAbraBankaccount;
begin
  rokVypisu := IntToStr(SysUtils.CurrentYear);

  //Fio
  bankUcetFio :=  TBankUcet.create('2100098382/2010', rokVypisu);

  lblVypisFioInfo.caption := bankUcetFio.infoSouhrnOVypisech;

  if bankUcetFio.gpcSoubor = '' then begin //nenašel se
    lblVypisFioGpc.caption := bankUcetFio.hledanyGpcSoubor + ' nenalezen';
    btnVypisFio.Enabled := false;
  end else begin
    lblVypisFioGpc.caption := bankUcetFio.gpcSoubor;
    btnVypisFio.Enabled := true;
  end;

  if (bankUcetFio.pocetVypisu = bankUcetFio.poradoveCisloMaxVypisu) and
    (bankUcetFio.pocetVypisu = bankUcetFio.extPoradoveCisloMaxVypisu) then
    lblVypisFioInfo.Font.Color := $008000
  else
    lblVypisFioInfo.Font.Color := $800000;


  /// Fio spoøicí
  bankUcetFioSporici :=  TBankUcet.create('2602372070/2010', rokVypisu);

  lblVypisFioSporiciInfo.caption := bankUcetFioSporici.infoSouhrnOVypisech;

  if bankUcetFioSporici.gpcSoubor = '' then begin //nenašel se
    lblVypisFioSporiciGpc.caption := bankUcetFioSporici.hledanyGpcSoubor + ' nenalezen';
    btnVypisFioSporici.Enabled := false;
  end else begin
    lblVypisFioSporiciGpc.caption := bankUcetFioSporici.gpcSoubor;
    btnVypisFioSporici.Enabled := true;
  end;

  if (bankUcetFioSporici.pocetVypisu = bankUcetFioSporici.poradoveCisloMaxVypisu) and
    (bankUcetFioSporici.pocetVypisu = bankUcetFioSporici.extPoradoveCisloMaxVypisu) then
    lblVypisFioSporiciInfo.Font.Color := $008000
  else
    lblVypisFioSporiciInfo.Font.Color := $800000;


  /// Fiokonto
  bankUcetFiokonto :=  TBankUcet.create('2800098383/2010', rokVypisu);

  lblVypisFiokontoInfo.caption := bankUcetFiokonto.infoSouhrnOVypisech;

  if bankUcetFiokonto.gpcSoubor = '' then begin //nenašel se
    lblVypisFiokontoGpc.caption := bankUcetFiokonto.hledanyGpcSoubor + ' nenalezen';
    btnVypisFiokonto.Enabled := false;
  end else begin
    lblVypisFiokontoGpc.caption := bankUcetFiokonto.gpcSoubor;
    btnVypisFiokonto.Enabled := true;
  end;

  if (bankUcetFiokonto.pocetVypisu = bankUcetFiokonto.poradoveCisloMaxVypisu) and
    (bankUcetFiokonto.pocetVypisu = bankUcetFiokonto.extPoradoveCisloMaxVypisu) then
    lblVypisFiokontoInfo.Font.Color := $008000
  else
    lblVypisFiokontoInfo.Font.Color := $800000;

  {
  /// Fio
  abraBankaccount.loadByNumber('2100098382/2010');
  maxCisloVypisuFio := abraBankaccount.getPoradoveCisloMaxVypisu(rokVypisu);
  hledanyGpcSoubor := 'Vypis_z_uctu-2100098382_' + rokVypisu + '*-' + IntToStr(maxCisloVypisuFio + 1) + '.gpc';
  nalezenyGpcSoubor := FindInFolder(DesU.GPC_PATH, hledanyGpcSoubor, true);
  if nalezenyGpcSoubor = '' then begin //nenašel se
    lblVypisFioGpc.caption := hledanyGpcSoubor + ' nenalezen';
    //btnVypisFio.Enabled := false;
    btnVypisFio.Caption := 'Stáhnout Fio výpis';
  end else begin
    lblVypisFioGpc.caption := nalezenyGpcSoubor;
    btnVypisFio.Enabled := true;
    btnVypisFio.Caption := 'Fio výpis';
  end;
    i1 := abraBankaccount.getPocetVypisu(rokVypisu);
    i2 := abraBankaccount.getPoradoveCisloMaxVypisu(rokVypisu);
    i3 := abraBankaccount.getExtPoradoveCisloMaxVypisu(rokVypisu);
  lblVypisFioInfo.Caption := format('%d výpisù v roce %s. Max. èíslo výpisu %d, externí èíslo %d, datum posledního výpisu %s', [
    i1, rokVypisu, i2, i3,
    DateToStr(abraBankaccount.getDatumMaxVypisu(rokVypisu))
    ]);
  if (i1 = i2) and (i1 = i3) then
    lblVypisFioInfo.Font.Color := $008000
  else
    lblVypisFioInfo.Font.Color := $800000;



  /// Fio Spoøicí
  abraBankaccount.loadByNumber('2602372070/2010');
  maxCisloVypisu := abraBankaccount.getPoradoveCisloMaxVypisu(rokVypisu);
  hledanyGpcSoubor := 'Vypis_z_uctu-2602372070_' + rokVypisu + '*-' + IntToStr(maxCisloVypisu + 1) + '.gpc';
  nalezenyGpcSoubor := FindInFolder(DesU.GPC_PATH, hledanyGpcSoubor, true);
  if nalezenyGpcSoubor = '' then begin //nenašel se
    lblVypisFioSporiciGpc.caption := hledanyGpcSoubor + ' nenalezen';
    btnVypisFioSporici.Enabled := false;
  end else begin
    lblVypisFioSporiciGpc.caption := nalezenyGpcSoubor;
    btnVypisFioSporici.Enabled := true;
  end;
    i1 := abraBankaccount.getPocetVypisu(rokVypisu);
    i2 := abraBankaccount.getPoradoveCisloMaxVypisu(rokVypisu);
    i3 := abraBankaccount.getExtPoradoveCisloMaxVypisu(rokVypisu);
  lblVypisFioSporiciInfo.Caption := format('%d výpisù v roce %s. Max. èíslo výpisu %d, externí èíslo %d, datum posledního výpisu %s', [
    i1, rokVypisu, i2, i3,
    DateToStr(abraBankaccount.getDatumMaxVypisu(rokVypisu))
    ]);
  if (i1 = i2) and (i1 = i3) then
    lblVypisFioSporiciInfo.Font.Color := $008000
  else
    lblVypisFioSporiciInfo.Font.Color := $800000;

  /// Fiokonto
  abraBankaccount.loadByNumber('2800098383/2010');
  maxCisloVypisu := abraBankaccount.getPoradoveCisloMaxVypisu(rokVypisu);
  hledanyGpcSoubor := 'Vypis_z_uctu-2800098383_' + rokVypisu + '*-' + IntToStr(maxCisloVypisu + 1) + '.gpc';
  nalezenyGpcSoubor := FindInFolder(DesU.GPC_PATH, hledanyGpcSoubor, true);
  if nalezenyGpcSoubor = '' then begin //nenašel se
    lblVypisFiokontoGpc.caption := hledanyGpcSoubor + ' nenalezen';
    btnVypisFiokonto.Enabled := false;
  end else begin
    lblVypisFiokontoGpc.caption := nalezenyGpcSoubor;
    btnVypisFiokonto.Enabled := true;
  end;
    i1 := abraBankaccount.getPocetVypisu(rokVypisu);
    i2 := abraBankaccount.getPoradoveCisloMaxVypisu(rokVypisu);
    i3 := abraBankaccount.getExtPoradoveCisloMaxVypisu(rokVypisu);
  lblVypisFiokontoInfo.Caption := format('%d výpisù v roce %s. Max. èíslo výpisu %d, externí èíslo %d, datum posledního výpisu %s', [
    i1, rokVypisu, i2, i3,
    DateToStr(abraBankaccount.getDatumMaxVypisu(rokVypisu))
    ]);
  if (i1 = i2) and (i1 = i3) then
    lblVypisFiokontoInfo.Font.Color := $008000
  else
    lblVypisFiokontoInfo.Font.Color := $800000;
  }


  /// ÈSOB
  abraBankAccount := TAbraBankaccount.create();
  abraBankaccount.loadByNumber('171336270/0300');
  maxCisloVypisu := abraBankaccount.getPoradoveCisloMaxVypisu(rokVypisu);
  //hledanyGpcSoubor := 'BB117641_171336270_' + rokVypisu + '*_' + IntToStr(maxCisloVypisu + 1) + '.gpc'; //takhle to bylo do zmìny v létì 2018
  hledanyGpcSoubor := '171336270_' + rokVypisu + '*_' + IntToStr(maxCisloVypisu + 1) + '_CZ.gpc';
  nalezenyGpcSoubor := FindInFolder(DesU.GPC_PATH, hledanyGpcSoubor, true);
  if nalezenyGpcSoubor = '' then begin //nenašel se
    lblVypisCsobGpc.caption := hledanyGpcSoubor + ' nenalezen';
    btnVypisCsob.Enabled := false;
  end else begin
    lblVypisCsobGpc.caption := nalezenyGpcSoubor;
    btnVypisCsob.Enabled := true;
  end;
    i1 := abraBankaccount.getPocetVypisu(rokVypisu);
    i2 := abraBankaccount.getPoradoveCisloMaxVypisu(rokVypisu);
    i3 := abraBankaccount.getExtPoradoveCisloMaxVypisu(rokVypisu);
  lblVypisCsobInfo.Caption := format('%d výpisù v roce %s. Max. èíslo výpisu %d, externí èíslo %d, datum posledního výpisu %s', [
    i1, rokVypisu, i2, i3,
    DateToStr(abraBankaccount.getDatumMaxVypisu(rokVypisu))
    ]);
  if (i1 = i2) and (i1 = i3) then
    lblVypisCsobInfo.Font.Color := $008000
  else
    lblVypisCsobInfo.Font.Color := $0000A0;


  /// RB
  abraBankAccount := TAbraBankaccount.create();
  abraBankaccount.loadByNumber('2921537004/5500');
  maxCisloVypisu := abraBankaccount.getPoradoveCisloMaxVypisu(rokVypisu);
  hledanyGpcSoubor := 'Vypis_2921537004_CZK_' + rokVypisu + '*' + IntToStr(maxCisloVypisu + 1) + '.gpc';
  nalezenyGpcSoubor := FindInFolder(DesU.GPC_PATH, hledanyGpcSoubor, true);
  if nalezenyGpcSoubor = '' then begin //nenašel se
    lblVypisRbGpc.caption := hledanyGpcSoubor + ' nenalezen';
    btnVypisRb.Enabled := false;
  end else begin
    lblVypisRbGpc.caption := nalezenyGpcSoubor;
    btnVypisRb.Enabled := true;
  end;
    i1 := abraBankaccount.getPocetVypisu(rokVypisu);
    i2 := abraBankaccount.getPoradoveCisloMaxVypisu(rokVypisu);
    i3 := abraBankaccount.getExtPoradoveCisloMaxVypisu(rokVypisu);
  lblVypisRbInfo.Caption := format('%d výpisù v roce %s. Max. èíslo výpisu %d, externí èíslo %d, datum posledního výpisu %s', [
    i1, rokVypisu, i2, i3,
    DateToStr(abraBankaccount.getDatumMaxVypisu(rokVypisu))
    ]);
  if (i1 = i2) and (i1 = i3) then
    lblVypisRbInfo.Font.Color := $008000
  else
    lblVypisRbInfo.Font.Color := $0000A0;

  //Pay U
  abraBankAccount := TAbraBankaccount.create();
  //abraBankaccount.loadByNumber('2389210008000000/0300'); //zmìnìno 08-2025
  abraBankaccount.loadByNumber('4003292157000000/0300');
  posledniDatum := abraBankaccount.getPosledniDatumVypisu(rokVypisu);
  hledanePayuDatumVypisu := FormatDateTime('yyyy-mm-dd', posledniDatum + 1);
  hledanyGpcSoubor := 'vypis_Eurosignal_' + hledanePayuDatumVypisu + '_*.gpc';
  nalezenyGpcSoubor := FindInFolder(DesU.GPC_PATH, hledanyGpcSoubor, true);
  if nalezenyGpcSoubor = '' then begin //nenašel se
    lblVypisPayuGpc.caption := hledanyGpcSoubor + ' nenalezen';
    btnVypisPayu.Enabled := false;
  end else begin
    lblVypisPayuGpc.caption := nalezenyGpcSoubor;
    btnVypisPayu.Enabled := true;
  end;
    i1 := abraBankaccount.getPocetVypisu(rokVypisu);
    i2 := abraBankaccount.getPoradoveCisloMaxVypisu(rokVypisu);
    // PayU nemá externí èíslo výpisu, proto nenèítáme externí èíslo posledního výpisu
  lblVypisPayuInfo.Caption := format('%d výpisù v roce %s. Max. èíslo výpisu %d, datum posledního výpisu %s', [
    i1, rokVypisu, i2,
    DateToStr(abraBankaccount.getDatumMaxVypisu(rokVypisu))
    ]);
  if (i1 = i2) then
    lblVypisPayuInfo.Font.Color := $008000
  else
    lblVypisPayuInfo.Font.Color := $800000;

end;


procedure TfmMain.btnStahniVypisyClick(Sender: TObject);
begin
  stahniVypisy;
end;

procedure TfmMain.stahniVypisy;
var
  wgResponse,
  cisloVypisu,
  gpcDate,
  gpcFileName : string;
  vysledekStazeni : TDesResult;
begin

  // Fio
  if bankUcetFio.gpcSoubor = '' then begin
    vysledekStazeni := bankUcetFio.stahniAktualnePotrebnyVypis;
    if vysledekStazeni.isOk then begin
      lblVypisFioGpc.caption := bankUcetFio.gpcSoubor;
      btnVypisFio.Enabled := true;
      Memo1.Lines.Add(vysledekStazeni.Messg);
    end else begin
      lblVypisFioGpc.caption := 'Výpis è. ' + IntToStr(bankUcetFio.aktualniCisloVypisu)  + ' se nepodaøilo stáhnout: ' + vysledekStazeni.Messg;
      btnVypisFio.Enabled := false;
      Memo1.Lines.Add('Chyba pøi stahování výpisu Fio: ' + vysledekStazeni.Code + '; ' + vysledekStazeni.Messg);
    end;
  end;

  {
  vysledekStazeni := bankUcetFio.stahniVypis(bankUcetFio.aktualniCisloVypisu + 1);
  if vysledekStazeni.isOk then begin
    Memo1.Lines.Add(vysledekStazeni.Messg);
  end else begin
    Memo1.Lines.Add('Chyba pøi stahování výpisu Fio: ' + vysledekStazeni.Code + '; ' + vysledekStazeni.Messg);
  end;
  vysledekStazeni := bankUcetFio.stahniVypis(bankUcetFio.aktualniCisloVypisu + 2);
  if vysledekStazeni.isOk then begin
    Memo1.Lines.Add(vysledekStazeni.Messg);
  end else begin
    Memo1.Lines.Add('Chyba pøi stahování výpisu Fio: ' + vysledekStazeni.Code + '; ' + vysledekStazeni.Messg);
  end;
  }

  // Fio spoøící
  if bankUcetFioSporici.gpcSoubor = '' then begin
    vysledekStazeni := bankUcetFioSporici.stahniAktualnePotrebnyVypis;
    if vysledekStazeni.isOk then begin
      lblVypisFioSporiciGpc.caption := bankUcetFioSporici.gpcSoubor;
      btnVypisFioSporici.Enabled := true;
      Memo1.Lines.Add(vysledekStazeni.Messg);
    end else begin
      lblVypisFioSporiciGpc.caption := 'Výpis è. ' + IntToStr(bankUcetFioSporici.aktualniCisloVypisu)  + ' se nepodaøilo stáhnout: ' + vysledekStazeni.Messg;
      btnVypisFioSporici.Enabled := false;
      Memo1.Lines.Add('Chyba pøi stahování výpisu Fio spoøící: ' + vysledekStazeni.Code + '; ' + vysledekStazeni.Messg);
    end;
  end;

  {
  vysledekStazeni := bankUcetFioSporici.stahniVypis(bankUcetFioSporici.aktualniCisloVypisu + 1);
  if vysledekStazeni.isOk then begin
    Memo1.Lines.Add(vysledekStazeni.Messg);
  end else begin
    Memo1.Lines.Add('Chyba pøi stahování výpisu Fio: ' + vysledekStazeni.Code + '; ' + vysledekStazeni.Messg);
  end;
  vysledekStazeni := bankUcetFioSporici.stahniVypis(bankUcetFioSporici.aktualniCisloVypisu + 2);
  if vysledekStazeni.isOk then begin
    Memo1.Lines.Add(vysledekStazeni.Messg);
  end else begin
    Memo1.Lines.Add('Chyba pøi stahování výpisu Fio: ' + vysledekStazeni.Code + '; ' + vysledekStazeni.Messg);
  end;
  }

  //Fiokonto
  if bankUcetFiokonto.gpcSoubor = '' then begin
    vysledekStazeni := bankUcetFiokonto.stahniAktualnePotrebnyVypis;
    if vysledekStazeni.isOk then begin
      lblVypisFiokontoGpc.caption := bankUcetFiokonto.gpcSoubor;
      btnVypisFiokonto.Enabled := true;
      Memo1.Lines.Add(vysledekStazeni.Messg);
    end else begin
      lblVypisFiokontoGpc.caption := 'Výpis è. ' + IntToStr(bankUcetFiokonto.aktualniCisloVypisu)  + ' se nepodaøilo stáhnout: ' + vysledekStazeni.Messg;
      btnVypisFiokonto.Enabled := false;
      Memo1.Lines.Add('Chyba pøi stahování výpisu Fiokonto: ' + vysledekStazeni.Code + '; ' + vysledekStazeni.Messg);
    end;
  end;

end;

procedure TfmMain.nactiGpc(GpcFilename : string);
var
  GpcInputFile : TextFile;
  GpcFileLine : string;
  iPlatbaZVypisu : TPlatbaZVypisu;
  i, pocetPlatebGpc, kontrolaDvojitaPlatba: integer;
  ucetniZustatek : currency;
  casStart, casPolozkaStart, cas02, cas03: double;
begin
  try
    casStart := Now;
    DesU.dbAbra.Reconnect;
    AssignFile(GpcInputFile, GpcFilename);
    Reset(GpcInputFile);

    Screen.Cursor := crHourGlass;
    btnNacti.Enabled := false;
    asgMain.Visible := true;
    asgMain.ClearNormalCells;
    asgMain.RowCount := 2;
    asgPredchoziPlatby.ClearNormalCells;
    asgPredchoziPlatbyVs.ClearNormalCells;
    asgNalezeneDoklady.ClearNormalCells;
    lblHlavicka.Font.Color := $660000;
    lblHlavickaVpravo.Font.Color := $666666;
    Application.ProcessMessages;

    // spoèítáme poèet plateb v GPC, abnychom mohli sledovat progres naèítání
    pocetPlatebGpc := 0;
    while not Eof(GpcInputFile) do
    begin
      ReadLn(GpcInputFile, GpcFileLine);
      if copy(GpcFileLine, 1, 3) = '075' then
        Inc(pocetPlatebGpc);
    end;
    CloseFile(GpcInputFile);
    Reset(GpcInputFile);

    Vypis := nil;
    i := 0;
    while not Eof(GpcInputFile) do
    begin
      casPolozkaStart := Now;
      lblHlavicka.Caption := '... naèítání ' + IntToStr(i) + '. z ' + IntToStr(pocetPlatebGpc);
      Application.ProcessMessages;
      ReadLn(GpcInputFile, GpcFileLine);

      // zpracuju první øádek GPC, musí to být hlavièka
      if i = 0 then
      begin
        Inc(i);
        if copy(GpcFileLine, 1, 3) = '074' then begin
          Vypis := TVypis.Create(GpcFileLine);
          Parovatko := TParovatko.Create(Vypis);
        end else begin
          MessageDlg('Neplatný GPC soubor, 1. øádek není hlavièka', mtInformation, [mbOk], 0);
          Break;
        end;
      end;

      // zpracuju øádek výpisu s platbou
      if copy(GpcFileLine, 1, 3) = '075' then //radek vypisu zacina 075
      begin
        Inc(i);
        // vytvoøím nový objekt PlatbaZVypisu
        iPlatbaZVypisu := TPlatbaZVypisu.Create(GpcFileLine);

        // zkontroluju, zda není vícenásobná platba, pokud ano,
        kontrolaDvojitaPlatba := Vypis.prictiCastkuPokudDvojitaPlatba(iPlatbaZVypisu);
        if kontrolaDvojitaPlatba > -1 then begin
          Memo1.Lines.Add('Dvojnásobná (vícenásobná) platba z úètu è. ' + iPlatbaZVypisu.cisloUctuKZobrazeni + ' s VS ' + iPlatbaZVypisu.VS);
          Parovatko.sparujPlatbu(Vypis.Platby[kontrolaDvojitaPlatba]);

        end else begin
          cas02 := Now;
          iPlatbaZVypisu.init(StrToInt(editPocetPredchPlateb.text));
          Parovatko.sparujPlatbu(iPlatbaZVypisu);
          iPlatbaZVypisu.automatickyOpravVS();
          Vypis.Platby.Add(iPlatbaZVypisu);
        end;

      end;
      zpravaRozdilCasu(casPolozkaStart, Now, 'zpracování položky výpisu');
    end;
    if assigned(Vypis) then
      if (Vypis.Platby.Count > 0) then
      begin
        Vypis.init(GpcFilename);
        Vypis.setridit();
        sparujVsechnyPrichoziPlatby;
        vyplnPrichoziPlatby;
        filtrujZobrazeniPlateb;
        ucetniZustatek := Vypis.abraBankaccount.getZustatek(Vypis.datum);  //zadáváme datum výpisu (= datum poslední platby), dostaneme poèáteèní stav bank. úètu k tomuto datu
        lblHlavicka.Caption := Vypis.abraBankaccount.name // + ', ' + Vypis.abraBankaccount.number
                        + ', è.' + IntToStr(Vypis.poradoveCislo) + ' (max è. ' + IntToStr(Vypis.maxExistujiciPoradoveCislo) + '). Plateb: '
                        + IntToStr(Vypis.Platby.Count);
        if Vypis.zustatekPocatecni <> ucetniZustatek then
        begin
          lblHlavickaVpravo.Caption := 'K ' + DateToStr(Vypis.datum) + ' bank. zùst: ' + format('%m', [Vypis.zustatekPocatecni])
                        + ' (úè. zùst: ' + format('%m', [ucetniZustatek]) + ')';
          lblHlavickaVpravo.Font.Color := $0000FF;
          Dialogs.MessageDlg('Pozor, poèáteèní zùstatek výpisu se neshoduje s úèetním zùstatkem v ABRA (zùstatky k ' + DateToStr(Vypis.datum) + ')', mtInformation, [mbOK], 0);
        end else
          lblHlavickaVpravo.Caption := 'K ' + DateToStr(Vypis.datum) + ' bank. zùst = úè. zùst: ' + format('%m', [ucetniZustatek]);
        lblVypisOverview.Caption := 'Výpis k úètu ' + Vypis.cisloUctuVlastni + ' (' + Vypis.abraBankaccount.name + ')'
                        + sLineBreak + 'Výpis k ' + DateToStr(Vypis.datum) + ', poøadové èíslo ' + IntToStr(Vypis.poradoveCislo)
                        + sLineBreak + 'Poèáteèní zùstatek: ' + format('%m', [Vypis.zustatekPocatecni])
                        + sLineBreak + 'Koncový zùstatek: ' + format('%m', [Vypis.zustatekKoncovy])
                        + sLineBreak + 'Kreditní obrat ' + format('%m', [Vypis.obratKredit]) + ', debetní obrat ' + format('%m', [Vypis.obratDebet]);
        if Vypis.isPayuVypis() then begin
          //Vypis.isPayuVypis();
          lblVypisOverview.Caption := lblVypisOverview.Caption + sLineBreak + '(PayU výpis neuvádí poøadové èíslo)';
        end;

        if not Vypis.isNavazujeNaradu() then
          Dialogs.MessageDlg('Doklad è. '+ IntToStr(Vypis.poradoveCislo) + ' nenavazuje na øadu!', mtInformation, [mbOK], 0);
        asgMainClick(nil);
      end;
  finally
    btnNacti.Enabled := true;
    btnHledej.Enabled := true;
    btnZapisDoAbry.Enabled := true;
    btnZavritVypis.Enabled := true;
    lblZobrazit.Enabled := true;
    chbZobrazitBezproblemove.Enabled := true;
    chbZobrazitStandardni.Enabled := true;
    chbZobrazitDebety.Enabled := true;
    lblPrechoziPlatbyZUctu.Enabled := true;
    lblPrechoziPlatbySVs.Enabled := true;
    lblNalezeneDoklady.Enabled := true;
    chbVsechnyDoklady.Enabled := true;
    chbVsechnyDoklady.Checked := false;

    CloseFile(GpcInputFile);
    Screen.Cursor := crDefault;
  end;
end;

procedure TfmMain.vyplnPrichoziPlatby;
var
  i : integer;
  iPlatbaZVypisu : TPlatbaZVypisu;
begin
  with asgMain do
  begin
    Enabled := true;
    ControlLook.NoDisabledButtonLook := true;
    ClearNormalCells;
    RowCount := Vypis.Platby.Count + 1;
    Row := 1;
    for i := 0 to Vypis.Platby.Count - 1 do
    begin
      RemoveButton(0, i+1);
      iPlatbaZVypisu := TPlatbaZVypisu(Vypis.Platby[i]);
      if iPlatbaZVypisu.VS <> iPlatbaZVypisu.VSOrig then
        AddButton(0, i+1, 76, 16, iPlatbaZVypisu.VSOrig, haCenter, vaCenter);
      if (iPlatbaZVypisu.kredit) then
        Cells[1, i+1] := format('%m', [iPlatbaZVypisu.castka])
      else
        Cells[1, i+1] := format('%m', [-iPlatbaZVypisu.castka]);
      if iPlatbaZVypisu.debet then asgMain.FontColors[1, i+1] := clRed;
      Cells[2, i+1] := iPlatbaZVypisu.VS;
      Cells[3, i+1] := iPlatbaZVypisu.SS;
      Cells[4, i+1] := iPlatbaZVypisu.cisloUctuKZobrazeni;
      //Cells[5, i+1] := Format('%8.2f', [iPlatbaZVypisu.getProcentoPredchozichPlatebNaStejnyVS]) + Format('%8.2f', [iPlatbaZVypisu.getProcentoPredchozichPlatebZeStejnehoUctu]) + iPlatbaZVypisu.nazevKlienta;
      Cells[5, i+1] := iPlatbaZVypisu.nazevProtistrany;
      Cells[6, i+1] := DateToStr(iPlatbaZVypisu.Datum);
      vyplnVysledekParovaniPP(i);
    end;
  end;
end;

procedure TfmMain.vyplnVysledekParovaniPP(i : integer);
var
  iPlatbaZVypisu : TPlatbaZVypisu;
begin
  iPlatbaZVypisu := TPlatbaZVypisu(Vypis.Platby[i]);
  case iPlatbaZVypisu.problemLevel of
    0: asgMain.Colors[2, i+1] := $AAFFAA;
    1: asgMain.Colors[2, i+1] := $CDFAFF;
    2: asgMain.Colors[2, i+1] := $60A4F4;
    3: asgMain.Colors[2, i+1] := $FFFACD;
    5: asgMain.Colors[2, i+1] := $BBBBFF;
  end;
  if iPlatbaZVypisu.rozdeleniPlatby > 0 then
    asgMain.Cells[8, i+1] := IntToStr (iPlatbaZVypisu.rozdeleniPlatby) + ' dìlení, ' + iPlatbaZVypisu.zprava
  else
    asgMain.Cells[8, i+1] := iPlatbaZVypisu.zprava;
  asgMain.RemoveCheckBox(7, i+1);
  if iPlatbaZVypisu.potrebaPotvrzeniUzivatelem then
    asgMain.AddCheckBox(7, i+1, iPlatbaZVypisu.jePotvrzenoUzivatelem, false);
end;

procedure TfmMain.filtrujZobrazeniPlateb;
var
  i : integer;
  iPlatbaZVypisu : TPlatbaZVypisu;
  zobrazitRadek : boolean;
begin
  for i := 0 to Vypis.Platby.Count - 1 do
  begin
    iPlatbaZVypisu := TPlatbaZVypisu(Vypis.Platby[i]);
    zobrazitRadek := false;
    if iPlatbaZVypisu.problemLevel > 1 then
      zobrazitRadek := true;
    if chbZobrazitBezproblemove.Checked AND (iPlatbaZVypisu.problemLevel = 0) then
      zobrazitRadek := true;
    if chbZobrazitStandardni.Checked AND (iPlatbaZVypisu.problemLevel = 1) then
      zobrazitRadek := true;
    if iPlatbaZVypisu.debet then
      if chbZobrazitDebety.Checked then
        zobrazitRadek := true
      else
        zobrazitRadek := false;
    if zobrazitRadek then
      asgMain.RowHeights[i+1] := asgMain.DefaultRowHeight
    else
      asgMain.RowHeights[i+1] := 0;
  end;
end;

procedure TfmMain.sparujVsechnyPrichoziPlatby;
var
  i : integer;
begin
  asgMain.RowCount := Vypis.Platby.Count + 1;
  Parovatko := TParovatko.create(Vypis);
  for i := 0 to Vypis.Platby.Count - 1 do
    sparujPrichoziPlatbu(i);
end;

procedure TfmMain.sparujPrichoziPlatbu(i : integer);
var
  iPlatbaZVypisu : TPlatbaZVypisu;
begin
  iPlatbaZVypisu := TPlatbaZVypisu(Vypis.Platby[i]);
  Parovatko.sparujPlatbu(iPlatbaZVypisu);
  vyplnVysledekParovaniPP(i);
end;

procedure TfmMain.vyplnPredchoziPlatby;
var
  i : integer;
  iPredchoziPlatba : TPredchoziPlatba;
begin
  // pøedchozí platby z úètu
  with asgPredchoziPlatby do begin
    Enabled := true;
    ClearNormalCells;
    
    if currPlatbaZVypisu.cisloUctuKZobrazeni = '' then
      lblPrechoziPlatbyZUctu.Caption := 'Prázdný úèet protistrany, pøedchozí platby nenaèítáme'
    else if currPlatbaZVypisu.cisloUctuSNulami = '0000000160987123/0300' then
      lblPrechoziPlatbyZUctu.Caption := 'Pro Èeskou poštu pøedchozí platby nenaèítáme'
    else if currPlatbaZVypisu.PredchoziPlatbyUcetList.Count = 0 then
      lblPrechoziPlatbyZUctu.Caption := 'Žádné pøedchozí platby z úètu ' + currPlatbaZVypisu.cisloUctuKZobrazeni
    else
      lblPrechoziPlatbyZUctu.Caption := 'Pøedchozí platby z úètu ' + currPlatbaZVypisu.cisloUctuKZobrazeni;

    if (currPlatbaZVypisu.cisloUctuKZobrazeni <> '') AND (currPlatbaZVypisu.PredchoziPlatbyUcetList.Count > 0) then
    begin
      RowCount := currPlatbaZVypisu.PredchoziPlatbyUcetList.Count + 1;
      for i := 0 to RowCount - 2 do begin
        iPredchoziPlatba := TPredchoziPlatba(currPlatbaZVypisu.PredchoziPlatbyUcetList[i]);
        if iPredchoziPlatba.VS <> currPlatbaZVypisu.VS then
          AddButton(0,i+1,25,18,'<--',haCenter,vaCenter);
        Cells[1, i+1] := iPredchoziPlatba.VS;
        Cells[2, i+1] := format('%m', [iPredchoziPlatba.Castka]);
        if iPredchoziPlatba.Castka < 0 then asgPredchoziPlatby.FontColors[2, i+1] := clRed;
        Cells[3, i+1] := DateToStr(iPredchoziPlatba.Datum);
        Cells[4, i+1] := iPredchoziPlatba.FirmName;
      end;
    end else
       RowCount := 2;
  end;

  // pøechozí platby s VS
  with asgPredchoziPlatbyVs do begin
    Enabled := true;
    ClearNormalCells;

    if currPlatbaZVypisu.VS = '' then
      lblPrechoziPlatbySVs.Caption := 'Prázdný VS, pøedchozí platby nenaèítáme'
    else
      lblPrechoziPlatbySVs.Caption := 'Pøedchozí platby s VS ' + currPlatbaZVypisu.VS;

    if (currPlatbaZVypisu.VS <> '') AND (currPlatbaZVypisu.PredchoziPlatbyVsList.Count > 0) then
    begin
      RowCount := currPlatbaZVypisu.PredchoziPlatbyVsList.Count + 1;
      for i := 0 to RowCount - 2 do begin
        iPredchoziPlatba := TPredchoziPlatba(currPlatbaZVypisu.PredchoziPlatbyVsList[i]);
        Cells[0, i+1] := iPredchoziPlatba.cisloUctuKZobrazeni;
        Cells[1, i+1] := format('%m', [iPredchoziPlatba.Castka]);
        if iPredchoziPlatba.Castka < 0 then asgPredchoziPlatbyVs.FontColors[1, i+1] := clRed;
        Cells[2, i+1] := DateToStr(iPredchoziPlatba.Datum);
        Cells[3, i+1] := iPredchoziPlatba.FirmName;
      end;
    end else
      RowCount := 2;
  end;
end;

procedure TfmMain.vyplnDoklady;
var
  iDoklad : TDoklad;
  iPDPar : TPlatbaDokladPar;
  i : integer;
begin

  with asgNalezeneDoklady do begin
    Enabled := true;
    ClearNormalCells;
    if currPlatbaZVypisu.DokladyList.Count > 0 then
    begin
      RowCount := currPlatbaZVypisu.DokladyList.Count + 1;
      for i := 0 to RowCount - 2 do begin
        iDoklad := TDoklad(currPlatbaZVypisu.DokladyList[i]);
        Cells[0, i+1] := iDoklad.CisloDokladu;
        Cells[1, i+1] := DateToStr(iDoklad.DatumDokladu);
        Cells[2, i+1] := iDoklad.FirmName;
        Cells[3, i+1] := format('%m', [iDoklad.Castka]);
        Cells[4, i+1] := format('%m', [iDoklad.CastkaZaplaceno]);
        Cells[5, i+1] := format('%m', [iDoklad.CastkaDobropisovano]);
        Cells[6, i+1] := format('%m', [iDoklad.CastkaNezaplaceno]);
        iPDPar := Parovatko.getPDPar(currPlatbaZVypisu, iDoklad.ID);
        if Assigned(iPDPar) then begin
          Cells[7, i+1] := iPDPar.Popis.Trim.TrimRight(['|']);
          if iPDPar.CastkaPouzita = iDoklad.CastkaNezaplaceno then
          begin
            Colors[6, i+1] := $AAFFAA;
            Cells[7, i+1] := 'zaplaceno celé';
          end
          else
            Colors[6, i+1] := $CDFAFF;
        end;
        if iDoklad.CastkaNezaplaceno = 0 then Colors[6, i+1] := $BBBBFF;
      end;
      chbVsechnyDoklady.Checked := currPlatbaZVypisu.vsechnyDoklady;
      if chbVsechnyDoklady.Checked then
        lblNalezeneDoklady.Caption := 'Všechny doklady s VS ' +  currPlatbaZVypisu.VS
      else
      begin
        iDoklad := TDoklad(currPlatbaZVypisu.DokladyList[0]);
        if (currPlatbaZVypisu.DokladyList.Count = 1) AND (iDoklad.castkaNezaplaceno = 0) then
            lblNalezeneDoklady.Caption := 'Zaplacený doklad s VS ' +  currPlatbaZVypisu.VS
          else
            lblNalezeneDoklady.Caption := 'Nezaplacené doklady s VS ' +  currPlatbaZVypisu.VS
      end;
    end else begin
      RowCount := 2;
      if currPlatbaZVypisu.VS <> '' then
        lblNalezeneDoklady.Caption := 'Žádné vystavené doklady s VS ' +  currPlatbaZVypisu.VS
      else
        lblNalezeneDoklady.Caption := 'Prázdný VS, žádné vystavené doklady' +  currPlatbaZVypisu.VS
    end;
  end;
end;

procedure TfmMain.urciCurrPlatbaZVypisu;
begin
  if assigned(Vypis) then
    if assigned(Vypis.Platby[asgMain.row - 1]) then
      currPlatbaZVypisu := TPlatbaZVypisu(Vypis.Platby[asgMain.row - 1]);
end;

procedure TfmMain.btnZapisDoAbryClick(Sender: TObject);
var
  vysledekZapisu  : TDesResult;
  casStart, dobaZapisu: double;
begin
  if not Vypis.isNavazujeNaRadu() then
    if Dialogs.MessageDlg('Èíslo dokladu ' + IntToStr(Vypis.poradoveCislo)
        + ' nenavazuje na existující øadu. Opravdu zapsat do Abry?',
        mtConfirmation, [mbYes, mbNo], 0 ) = mrNo then Exit;
  Screen.Cursor := crHourGlass;
  btnZapisDoAbry.Enabled := False;
  casStart := Now;
  try
    sparujVsechnyPrichoziPlatby;
    vysledekZapisu := Parovatko.zapisDoAbry();
  finally
    Screen.Cursor := crDefault;
  end;

  dobaZapisu := (Now - casStart) * 24 * 3600;
  Memo1.Lines.Add('Doba trvání: ' + floattostr(RoundTo(dobaZapisu, -2))
              + ' s (' + floattostr(RoundTo(dobaZapisu / 60, -2)) + ' min)');


  if vysledekZapisu.isOk then begin
    MessageDlg('Zápis do Abry úspìšnì dokonèen', mtInformation, [mbOk], 0);

  end else begin
    Memo1.Lines.Add('Chyba zápisu ' + vysledekZapisu.Code + ': ' + vysledekZapisu.Messg);
    MessageDlg('Chyba zápisu ' + vysledekZapisu.Code + ': ' + vysledekZapisu.Messg, mtInformation, [mbOk], 0);

  end;

  Memo1.Lines.Add(vysledekZapisu.Messg);
  DesU.dbAbra.Reconnect;
  vyplnNacitaciButtony;
end;

procedure TfmMain.provedAkcePoZmeneVS;
begin
  asgMain.Cells[2, asgMain.row] := currPlatbaZVypisu.VS;
  asgMain.RemoveButton(0, asgMain.row);
  if currPlatbaZVypisu.VS <> currPlatbaZVypisu.VSOrig then
    asgMain.AddButton(0, asgMain.row, 76, 16, currPlatbaZVypisu.VSOrig, haCenter, vaCenter);
  currPlatbaZVypisu.loadPredchoziPlatbyPodleVS(StrToInt(editPocetPredchPlateb.text));
  currPlatbaZVypisu.loadDokladyPodleVS();
  sparujVsechnyPrichoziPlatby;
  vyplnPredchoziPlatby;
  vyplnDoklady;
  Memo2.Clear;
  Memo2.Lines.Add(Parovatko.getPDParyPlatbyAsText(currPlatbaZVypisu));
end;

{*********************** akce Input elementù **********************************}
procedure TfmMain.asgMainClick(Sender: TObject);
begin
  urciCurrPlatbaZVypisu;
  vyplnPredchoziPlatby;
  vyplnDoklady;
  Memo2.Clear;
  Memo2.Lines.Add(Parovatko.getPDParyPlatbyAsText(currPlatbaZVypisu));
end;

procedure TfmMain.asgMainCellsChanged(Sender: TObject; R: TRect);
begin
  { //pøesunuto do TfmMain.asgMainCellValidate
  if asgMain.col = 2 then //zmìna VS
  begin
     //asgMain.Colors[asgMain.col, asgMain.row] := clMoneyGreen;
     currPlatbaZVypisu.VS := asgMain.Cells[2, asgMain.row]; //do pøíslušného objektu platby zapíšu zmìnìný VS
     provedAkcePoZmeneVS;
  end;
  }

  if asgMain.col = 5 then //zmìna textu (názvu klienta)
  begin
     //asgMain.Colors[asgMain.col, asgMain.row] := clMoneyGreen;
     currPlatbaZVypisu.nazevProtistrany := asgMain.Cells[5, asgMain.row]; //do pøíslušného objektu platby zapíšu zmìnìný text
  end;
end;

procedure TfmMain.asgMainCellValidate(Sender: TObject; ACol, ARow: Integer;
  var Value: string; var Valid: Boolean);
var
 len: integer;
begin
  if asgMain.col = 2 then //zmìna VS
  begin
    Value := KeepOnlyNumbers(Value);
    currPlatbaZVypisu.VS := Value; //do pøíslušného objektu platby zapíšu zmìnìný VS
    provedAkcePoZmeneVS;
  end;
end;

procedure TfmMain.asgMainCheckBoxClick(Sender: TObject; ACol, ARow: Integer;
  State: Boolean);
begin
  asgMain.row := ARow;
  urciCurrPlatbaZVypisu;
  currPlatbaZVypisu.jePotvrzenoUzivatelem := State;
  //Memo1.Lines.Add(BoolToStr(State,true));

end;

procedure TfmMain.asgPredchoziPlatbyButtonClick(Sender: TObject; ACol,
  ARow: Integer);
begin
  urciCurrPlatbaZVypisu;
  currPlatbaZVypisu.VS := TPredchoziPlatba(currPlatbaZVypisu.PredchoziPlatbyUcetList[ARow - 1]).VS;
  provedAkcePoZmeneVS;
end;
procedure TfmMain.asgMainKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
 //showmessage('Stisknuto: ' + IntToStr(Key));
  if Key = 27 then
  begin
    currPlatbaZVypisu.VS := currPlatbaZVypisu.VSOrig;
    provedAkcePoZmeneVS;
  end;
end;

procedure TfmMain.chbVsechnyDokladyClick(Sender: TObject);
begin
  if assigned(currPlatbaZVypisu) then begin
    currPlatbaZVypisu.vsechnyDoklady := chbVsechnyDoklady.Checked;
    currPlatbaZVypisu.loadDokladyPodleVS();
    vyplnDoklady;
  end
  else
    chbVsechnyDoklady.Checked := false;

end;

procedure TfmMain.btnSparujPlatbyClick(Sender: TObject);
begin
  sparujVsechnyPrichoziPlatby;
end;

procedure TfmMain.Zprava(TextZpravy: string);
// do listboxu a logfile uloží èas a text zprávy
begin
  Memo1.Lines.Add(FormatDateTime('dd.mm.yy hh:nn:ss  ', Now) + TextZpravy);
  {lbxLog.ItemIndex := lbxLog.Count - 1;
  Application.ProcessMessages;
  Append(F);
  Writeln (F, FormatDateTime('dd.mm.yy hh:nn  ', Now) + TextZpravy);
  CloseFile(F);  }
  Application.ProcessMessages;
end;
procedure TfmMain.zpravaRozdilCasu(cas01, cas02 : double; textZpravy : string);
// do listboxu vypíše èasový rozdíl a zprávu
begin
  debugRozdilCasu(cas01, cas02, textZpravy);
  cas01 := cas01 * 24 * 3600;
  cas02 := cas02 * 24 * 3600;
  //Memo1.Lines.Add('Èas1: ' + floattostr(RoundTo(cas01, -2)));
  //Memo1.Lines.Add('Èas2: ' + floattostr(RoundTo(cas02, -2)));
  Memo2.Lines.Add('Trvání: ' + floattostr(RoundTo(cas02 - cas01, -2))
              + ' s, ' + textZpravy);

end;

procedure TfmMain.asgMainGetAlignment(Sender: TObject; ARow, ACol: Integer;
  var HAlign: TAlignment; var VAlign: TVAlignment);
begin
  if (ARow = 0) then HAlign := taCenter
  else case ACol of
    0,6: HAlign := taCenter;
    1..4: HAlign := taRightJustify;
    //4: HAlign := taLeftJustify;
  end;
end;
procedure TfmMain.asgNalezeneDokladyGetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
  if (ARow = 0) then HAlign := taCenter
  else case ACol of
    //0,6: HAlign := taCenter;
    1,3..6: HAlign := taRightJustify;
    //4: HAlign := taLeftJustify;
  end;
end;
procedure TfmMain.asgPredchoziPlatbyGetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
  case ACol of
    1..3: HAlign := taRightJustify;
  end;
end;
procedure TfmMain.asgPredchoziPlatbyVsGetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
  case ACol of
    0..2: HAlign := taRightJustify;
  end;
end;
procedure TfmMain.btnReconnectClick(Sender: TObject);
var
  jsonstring,
  newIssuedInvoice : string;
begin
  //Memo2.Lines.Add(DesU.vytvorFaZaVoipKredit('795532', 2561, 42914));
  {
  jsonstring := LoadFileToStr(DesU.PROGRAM_PATH + '!jsonin.txt');
  Memo2.Lines.Add(SO(jsonstring).AsJSon(true, true));
  newIssuedInvoice := DesU.abraBoCreateOLE('issuedinvoice',  SO(jsonstring));
  //Memo2.Lines.Add(SO(newIssuedInvoice).S['id']);
  Memo2.Lines.Add(newIssuedInvoice);
  }
  DesU.dbAbra.Reconnect;
  vyplnNacitaciButtony;
  //
  //Memo2.Lines.Add('FirmId: ' + DesU.getFirmIdByCode(DesU.getAbracodeByContractNumber('20179001')));   iPlatbaZVypisu.nazevKlienta
  //DesU.getFirmIdByCode();
end;
procedure TfmMain.btnHledejClick(Sender: TObject);
var
  hledejResult : TArrayOf2Int;
begin
  hledejResult := Vypis.hledej(Trim(editHledej.Text));
  asgMain.row := hledejResult[0] + 1;
  asgMain.col := hledejResult[1];
end;

procedure TfmMain.chbZobrazitBezproblemoveClick(Sender: TObject);
begin
  filtrujZobrazeniPlateb;
end;
procedure TfmMain.chbZobrazitDebetyClick(Sender: TObject);
begin
  filtrujZobrazeniPlateb;
end;
procedure TfmMain.chbZobrazitStandardniClick(Sender: TObject);
begin
  filtrujZobrazeniPlateb;
end;


procedure TfmMain.asgMainCanEditCell(Sender: TObject; ARow, ACol: Integer;
  var CanEdit: Boolean);
begin
  case ACol of
    0..1: CanEdit := false;
    3..4: CanEdit := false;
    6: CanEdit := false;
    //8: CanEdit := false;
  end;
end;

procedure TfmMain.asgMainGetEditorType(Sender: TObject; ACol,
  ARow: Integer; var AEditor: TEditorType);
begin
{
  case ACol of
    1..2: AEditor := edRichEdit;
  end;
}
end;
procedure TfmMain.asgMainGetCellColor(Sender: TObject; ARow, ACol: Integer;
  AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
begin
  if (ARow > 0) then
  case ACol of
    1..2: AFont.Style := [];
  end;
end;


procedure TfmMain.btnShowPrirazeniPnpFormClick(Sender: TObject);
begin
  fmPrirazeniPnp.Show;
end;
procedure TfmMain.asgMainButtonClick(Sender: TObject; ACol, ARow: Integer);
begin
  asgMain.row := ARow;
  urciCurrPlatbaZVypisu;
  currPlatbaZVypisu.VS := currPlatbaZVypisu.VSOrig;
  provedAkcePoZmeneVS;
end;
procedure TfmMain.btnNactiClick(Sender: TObject);
begin
  //if DesU.existujeVAbreDokladSPrazdnymVs() then exit;
  // naètení GPC na základì dialogu
  NactiGpcDialog.InitialDir := DesU.GPC_PATH; //'J:\Eurosignal\HB\';
  NactiGpcDialog.Filter := 'Bankovní výpisy (*.gpc)|*.gpc';
	if NactiGpcDialog.Execute then
    nactiGpc(NactiGpcDialog.Filename);
end;

procedure TfmMain.btnVypisFioClick(Sender: TObject);
begin
  nactiGpc(lblVypisFioGpc.caption);
end;
procedure TfmMain.btnVypisFioSporiciClick(Sender: TObject);
begin
  nactiGpc(lblVypisFioSporiciGpc.caption);
end;
procedure TfmMain.btnVypisFiokontoClick(Sender: TObject);
begin
  nactiGpc(lblVypisFiokontoGpc.caption);
end;
procedure TfmMain.btnVypisPayUClick(Sender: TObject);
begin
  nactiGpc(lblVypisPayuGpc.caption);
end;

procedure TfmMain.btnVypisCsobClick(Sender: TObject);
begin
  nactiGpc(lblVypisCsobGpc.caption);
end;

procedure TfmMain.btnVypisRbClick(Sender: TObject);
begin
  nactiGpc(lblVypisRbGpc.caption);
end;


procedure TfmMain.btnZavritVypisClick(Sender: TObject);
begin
  DesU.dbAbra.Reconnect;
  asgMain.ClearNormalCells;
  asgMain.Visible := false;
  asgPredchoziPlatby.ClearNormalCells;
  asgPredchoziPlatbyVs.ClearNormalCells;
  asgNalezeneDoklady.ClearNormalCells;
  lblHlavicka.Caption := '';
  lblHlavickaVpravo.Caption := '';
  lblVypisOverview.Caption := '';
  btnHledej.Enabled := false;
  btnZapisDoAbry.Enabled := false;
  btnZavritVypis.Enabled := false;
  lblZobrazit.Enabled := false;
  chbZobrazitBezproblemove.Enabled := false;
  chbZobrazitStandardni.Enabled := false;
  chbZobrazitDebety.Enabled := false;
  lblPrechoziPlatbyZUctu.Caption := 'Pøedchozí platby z úètu';
  lblPrechoziPlatbyZUctu.Enabled := false;
  lblPrechoziPlatbySVs.Caption := 'Pøedchozí platby s VS';
  lblPrechoziPlatbySVs.Enabled := false;
  lblNalezeneDoklady.Caption := 'Doklady podle VS';
  lblNalezeneDoklady.Enabled := false;
  chbVsechnyDoklady.Enabled := false;
  chbVsechnyDoklady.Checked := false;
  Memo1.Clear;
  Memo2.Clear;
  vyplnNacitaciButtony;
end;


procedure TfmMain.btnCustomersClick(Sender: TObject);
begin
  fmCustomers.Show;
end;


end.
