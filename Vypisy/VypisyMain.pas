unit VypisyMain;
interface
uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, IniFiles, Forms,
  Dialogs, StdCtrls, Grids, AdvObj, BaseGrid, AdvGrid, StrUtils,
  DB, ComObj, AdvEdit, DateUtils, Math, ExtCtrls,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, ZAbstractConnection, ZConnection,
  uTVypis, uTPlatbaZVypisu, uTParovatko, AdvUtil;
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
    procedure btnNactiClick(Sender: TObject);
    procedure btnZapisDoAbryClick(Sender: TObject);
    procedure asgMainGetAlignment(Sender: TObject; ARow, ACol: Integer;
              var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
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
    procedure btnZavritVypisClick(Sender: TObject);
    procedure btnCustomersClick(Sender: TObject);
    procedure asgMainCheckBoxClick(Sender: TObject; ACol, ARow: Integer;
      State: Boolean);
    procedure btnVypisPayUClick(Sender: TObject);

  public
    procedure nactiGpc(GpcFilename : string);
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
implementation
uses
  AbraEntities, DesUtils, Superobject, Customers, PrirazeniPNP;
{$R *.dfm}
procedure TfmMain.FormShow(Sender: TObject);
begin
  //DesU.desUtilsInit('');
  //asgMain.CheckFalse := '0';
  //asgMain.CheckTrue := '1';
  vyplnNacitaciButtony;
  if DesU.appMode >= 3 then
  begin
    btnReconnect.Visible := true;
    btnSparujPlatby.Visible := true;
  end;
  //fmPrirazeniPnp.Show;  //pøi programování kvùli zrychlení práce (DEVEL)
  DesU.existujeVAbreDokladSPrazdnymVs();
end;

procedure TfmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if assigned (Vypis) then
    if assigned (Vypis.Platby) then
      Vypis.Platby.Free;
end;
procedure TfmMain.vyplnNacitaciButtony;
var
  maxCisloVypisu : integer;
  posledniDatum : double;
  fRok, nalezenyGpcSoubor, hledanyGpcSoubor,
  hledanePayuDatumVypisu : string;
  abraBankaccount : TAbraBankaccount;
begin
  fRok := IntToStr(SysUtils.CurrentYear);
  abraBankAccount := TAbraBankaccount.create();

  //Fio
  abraBankaccount.loadByNumber('2100098382/2010');
  maxCisloVypisu := abraBankaccount.getPoradoveCisloMaxVypisu(fRok);
  hledanyGpcSoubor := 'Vypis_z_uctu-2100098382_' + fRok + '*-' + IntToStr(maxCisloVypisu + 1) + '.gpc';
  nalezenyGpcSoubor := FindInFolder(DesU.GPC_PATH, hledanyGpcSoubor, true);
  if nalezenyGpcSoubor = '' then begin //nenašel se
    lblVypisFioGpc.caption := hledanyGpcSoubor + ' nenalezen';
    btnVypisFio.Enabled := false;
  end else begin
    lblVypisFioGpc.caption := nalezenyGpcSoubor;
    btnVypisFio.Enabled := true;
  end;
  lblVypisFioInfo.Caption := format('Poèet výpisù: %d, max. èíslo výpisu: %d, externí èíslo: %d, datum %s', [
    abraBankaccount.getPocetVypisu(fRok),
    abraBankaccount.getPoradoveCisloMaxVypisu(fRok),
    abraBankaccount.getExtPoradoveCisloMaxVypisu(fRok),
    DateToStr(abraBankaccount.getDatumMaxVypisu(fRok))
    ]);

  /// Fio Spoøicí
  abraBankaccount.loadByNumber('2602372070/2010');
  maxCisloVypisu := abraBankaccount.getPoradoveCisloMaxVypisu(fRok);
  hledanyGpcSoubor := 'Vypis_z_uctu-2602372070_' + fRok + '*-' + IntToStr(maxCisloVypisu + 1) + '.gpc';
  nalezenyGpcSoubor := FindInFolder(DesU.GPC_PATH, hledanyGpcSoubor, true);
  if nalezenyGpcSoubor = '' then begin //nenašel se
    lblVypisFioSporiciGpc.caption := hledanyGpcSoubor + ' nenalezen';
    btnVypisFioSporici.Enabled := false;
  end else begin
    lblVypisFioSporiciGpc.caption := nalezenyGpcSoubor;
    btnVypisFioSporici.Enabled := true;
  end;
  lblVypisFioSporiciInfo.Caption := format('Poèet výpisù: %d, max. èíslo výpisu: %d, externí èíslo: %d, datum %s', [
    abraBankaccount.getPocetVypisu(fRok),
    abraBankaccount.getPoradoveCisloMaxVypisu(fRok),
    abraBankaccount.getExtPoradoveCisloMaxVypisu(fRok),
    DateToStr(abraBankaccount.getDatumMaxVypisu(fRok))
    ]);

  /// Fiokonto
  abraBankaccount.loadByNumber('2800098383/2010');
  maxCisloVypisu := abraBankaccount.getPoradoveCisloMaxVypisu(fRok);
  hledanyGpcSoubor := 'Vypis_z_uctu-2800098383_' + fRok + '*-' + IntToStr(maxCisloVypisu + 1) + '.gpc';
  nalezenyGpcSoubor := FindInFolder(DesU.GPC_PATH, hledanyGpcSoubor, true);
  if nalezenyGpcSoubor = '' then begin //nenašel se
    lblVypisFiokontoGpc.caption := hledanyGpcSoubor + ' nenalezen';
    btnVypisFiokonto.Enabled := false;
  end else begin
    lblVypisFiokontoGpc.caption := nalezenyGpcSoubor;
    btnVypisFiokonto.Enabled := true;
  end;
  lblVypisFiokontoInfo.Caption := format('Poèet výpisù: %d, max. èíslo výpisu: %d, externí èíslo: %d, datum %s', [
    abraBankaccount.getPocetVypisu(fRok),
    abraBankaccount.getPoradoveCisloMaxVypisu(fRok),
    abraBankaccount.getExtPoradoveCisloMaxVypisu(fRok),
    DateToStr(abraBankaccount.getDatumMaxVypisu(fRok))
    ]);

  /// ÈSOB
  abraBankaccount.loadByNumber('171336270/0300');
  maxCisloVypisu := abraBankaccount.getPoradoveCisloMaxVypisu(fRok);
  //hledanyGpcSoubor := 'BB117641_171336270_' + fRok + '*_' + IntToStr(maxCisloVypisu + 1) + '.gpc'; //takhle to bylo do zmìny v létì 2018
  hledanyGpcSoubor := '171336270_' + fRok + '*_' + IntToStr(maxCisloVypisu + 1) + '_CZ.gpc';
  nalezenyGpcSoubor := FindInFolder(DesU.GPC_PATH, hledanyGpcSoubor, true);
  if nalezenyGpcSoubor = '' then begin //nenašel se
    lblVypisCsobGpc.caption := hledanyGpcSoubor + ' nenalezen';
    btnVypisCsob.Enabled := false;
  end else begin
    lblVypisCsobGpc.caption := nalezenyGpcSoubor;
    btnVypisCsob.Enabled := true;
  end;
  lblVypisCsobInfo.Caption := format('Poèet výpisù: %d, max. èíslo výpisu: %d, externí èíslo: %d, datum %s', [
    abraBankaccount.getPocetVypisu(fRok),
    abraBankaccount.getPoradoveCisloMaxVypisu(fRok),
    abraBankaccount.getExtPoradoveCisloMaxVypisu(fRok),
    DateToStr(abraBankaccount.getDatumMaxVypisu(fRok))
    ]);
  //Pay U
  abraBankaccount.loadByNumber('2389210008000000/0300');
  posledniDatum := abraBankaccount.getPosledniDatumVypisu(fRok);
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
  lblVypisPayuInfo.Caption := format('Datum posledního výpisu %s', [
    FormatDateTime('dd.mm.yyyy', posledniDatum)
    ]);
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
    asgMain.Visible := true;
    asgMain.ClearNormalCells;
    asgMain.RowCount := 2;
    asgPredchoziPlatby.ClearNormalCells;
    asgPredchoziPlatbyVs.ClearNormalCells;
    asgNalezeneDoklady.ClearNormalCells;
    btnNacti.Enabled := false;
    lblHlavicka.Font.Color := $000000;
    Application.ProcessMessages;
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
      if i = 0 then //první øádek musí být hlavièka výpisu
      begin
        Inc(i);
        if copy(GpcFileLine, 1, 3) = '074' then begin
          Vypis := TVypis.Create(GpcFileLine);
          Parovatko := TParovatko.create(Vypis);
        end else begin
          MessageDlg('Neplatný GPC soubor, 1. øádek není hlavièka', mtInformation, [mbOk], 0);
          Break;
        end;
      end;
      if copy(GpcFileLine, 1, 3) = '075' then //radek vypisu zacina 075
      begin
        Inc(i);
        iPlatbaZVypisu := TPlatbaZVypisu.Create(GpcFileLine);
        kontrolaDvojitaPlatba := Vypis.prictiCastkuPokudDvojitaPlatba(iPlatbaZVypisu);
        if kontrolaDvojitaPlatba > -1 then begin
          Memo1.Lines.Add('Dvojnásobná platba:  VS '+ iPlatbaZVypisu.VS + ' na cisle uctu ' + iPlatbaZVypisu.cisloUctuKZobrazeni);
          Parovatko.odparujPlatbu(Vypis.Platby[kontrolaDvojitaPlatba]);
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
        Vypis.init();
        Vypis.setridit();
        sparujVsechnyPrichoziPlatby;
        vyplnPrichoziPlatby;
        filtrujZobrazeniPlateb;
        ucetniZustatek := Vypis.abraBankaccount.getZustatek(Vypis.datum);  //zadáváme poèáteèní datum výpisu, dostaneme poèáteèní stav bank. úètu k tomuto datu
        lblHlavicka.Caption := Vypis.abraBankaccount.name // + ', ' + Vypis.abraBankaccount.number
                        + ', è.' + IntToStr(Vypis.poradoveCislo) + ' (max è. ' + IntToStr(Vypis.maxExistujiciPoradoveCislo) + '). Plateb: '
                        + IntToStr(Vypis.Platby.Count)
                        + ' Bank. zùst: ' + FloatToStr(Vypis.zustatekStary)
                        + ' Úè. zùst: ' + FloatToStr(ucetniZustatek);
        if Vypis.zustatekStary <> ucetniZustatek then
          lblHlavicka.Font.Color := $0000FF;
        if not Vypis.isNavazujeNaradu() then
          Dialogs.MessageDlg('Doklad è. '+ IntToStr(Vypis.poradoveCislo) + ' nenavazuje na øadu!', mtInformation, [mbOK], 0);
        //currPlatbaZVypisu := TPlatbaZVypisu(Vypis.Platby[0]); //mùže být ale nemìlo by být potøeba
        asgMainClick(nil);
      end;
  finally
    btnNacti.Enabled := true;
    btnZapisDoAbry.Enabled := true;
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
      if iPlatbaZVypisu.VS <> iPlatbaZVypisu.VS_orig then
        AddButton(0, i+1, 76, 16, iPlatbaZVypisu.VS_orig, haCenter, vaCenter);
      if (iPlatbaZVypisu.kredit) then
        Cells[1, i+1] := format('%m', [iPlatbaZVypisu.castka])
      else
        Cells[1, i+1] := format('%m', [-iPlatbaZVypisu.castka]);
      if iPlatbaZVypisu.debet then asgMain.FontColors[1, i+1] := clRed;
      Cells[2, i+1] := iPlatbaZVypisu.VS;
      Cells[3, i+1] := iPlatbaZVypisu.SS;
      Cells[4, i+1] := iPlatbaZVypisu.cisloUctuKZobrazeni;
      //Cells[5, i+1] := Format('%8.2f', [iPlatbaZVypisu.getProcentoPredchozichPlatebNaStejnyVS]) + Format('%8.2f', [iPlatbaZVypisu.getProcentoPredchozichPlatebZeStejnehoUctu]) + iPlatbaZVypisu.nazevKlienta;
      Cells[5, i+1] := iPlatbaZVypisu.nazevKlienta;
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
  with asgPredchoziPlatby do begin
    Enabled := true;
    ClearNormalCells;
    lblPrechoziPlatbyZUctu.Caption := 'Pøedchozí platby z úètu '
        + currPlatbaZVypisu.cisloUctuKZobrazeni;
    if currPlatbaZVypisu.PredchoziPlatbyList.Count > 0 then
    begin
      RowCount := currPlatbaZVypisu.PredchoziPlatbyList.Count + 1;
      for i := 0 to RowCount - 2 do begin
        iPredchoziPlatba := TPredchoziPlatba(currPlatbaZVypisu.PredchoziPlatbyList[i]);
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
  with asgPredchoziPlatbyVs do begin
    Enabled := true;
    ClearNormalCells;
    lblPrechoziPlatbySVs.Caption := 'Pøedchozí platby s VS ' + currPlatbaZVypisu.VS;
    if currPlatbaZVypisu.PredchoziPlatbyVsList.Count > 0 then
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
  //currPlatbaZVypisu.loadDokladyPodleVS(); //bylo v minulosti
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
        Cells[7, i+1] := iDoklad.ID;
        iPDPar := Parovatko.getPDPar(currPlatbaZVypisu, iDoklad.ID);
        if Assigned(iPDPar) then begin
          Cells[8, i+1] := iPDPar.Popis; // + floattostr(iPDPar.CastkaPouzita);
          if iPDPar.CastkaPouzita = iDoklad.CastkaNezaplaceno then
            Colors[6, i+1] := $AAFFAA
          else
            Colors[6, i+1] := $CDFAFF;
        end;
        if iDoklad.CastkaNezaplaceno = 0 then Colors[6, i+1] := $BBBBFF;
      end;
      chbVsechnyDoklady.Checked := currPlatbaZVypisu.vsechnyDoklady;
      if chbVsechnyDoklady.Checked then
        lblNalezeneDoklady.Caption := 'Doklady s VS ' +  currPlatbaZVypisu.VS
      else
        lblNalezeneDoklady.Caption := 'Doklady s VS ' +  currPlatbaZVypisu.VS;
    end else begin
      RowCount := 2;
      lblNalezeneDoklady.Caption := 'Žádné vystavené doklady s VS ' +  currPlatbaZVypisu.VS;
    end;
  end;
end;

procedure TfmMain.urciCurrPlatbaZVypisu();
begin
  if assigned(Vypis) then
    if assigned(Vypis.Platby[asgMain.row - 1]) then
      currPlatbaZVypisu := TPlatbaZVypisu(Vypis.Platby[asgMain.row - 1]);
end;

procedure TfmMain.btnZapisDoAbryClick(Sender: TObject);
var
  vysledek  : string;
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
    vysledek := Parovatko.zapisDoAbry();
  finally
    Screen.Cursor := crDefault;
  end;
  Memo1.Lines.Add(vysledek);
  dobaZapisu := (Now - casStart) * 24 * 3600;
  Memo1.Lines.Add('Doba trvání: ' + floattostr(RoundTo(dobaZapisu, -2))
              + ' s (' + floattostr(RoundTo(dobaZapisu / 60, -2)) + ' min)');
  DesU.dbAbra.Reconnect;
  MessageDlg('Zápis do Abry dokonèen', mtInformation, [mbOk], 0);
  vyplnNacitaciButtony;
end;

procedure TfmMain.provedAkcePoZmeneVS;
begin
  asgMain.Cells[2, asgMain.row] := currPlatbaZVypisu.VS;
  asgMain.RemoveButton(0, asgMain.row);
  if currPlatbaZVypisu.VS <> currPlatbaZVypisu.VS_orig then
    asgMain.AddButton(0, asgMain.row, 76, 16, currPlatbaZVypisu.VS_orig, haCenter, vaCenter);
  currPlatbaZVypisu.loadPredchoziPlatbyPodleVS(StrToInt(editPocetPredchPlateb.text));
  vyplnPredchoziPlatby;
  currPlatbaZVypisu.loadDokladyPodleVS();
  vyplnDoklady;
  sparujVsechnyPrichoziPlatby;
  //sparujPrichoziPlatbu(asgMain.row - 1);
  Memo2.Clear;
  Memo2.Lines.Add(Parovatko.getPDParyPlatbyAsText(currPlatbaZVypisu));
end;

{*********************** akce Input elementù **********************************}
procedure TfmMain.asgMainClick(Sender: TObject);
begin
  urciCurrPlatbaZVypisu();
  vyplnPredchoziPlatby;
  vyplnDoklady;
  Memo2.Clear;
  Memo2.Lines.Add(Parovatko.getPDParyPlatbyAsText(currPlatbaZVypisu));
end;

procedure TfmMain.asgMainCellsChanged(Sender: TObject; R: TRect);
begin
  if asgMain.col = 2 then //zmìna VS
  begin
     //asgMain.Colors[asgMain.col, asgMain.row] := clMoneyGreen;
     currPlatbaZVypisu.VS := asgMain.Cells[2, asgMain.row]; //do pøíslušného objektu platby zapíšu zmìnìný VS
     provedAkcePoZmeneVS;
  end;
  if asgMain.col = 5 then //zmìna textu (názvu klienta)
  begin
     //asgMain.Colors[asgMain.col, asgMain.row] := clMoneyGreen;
     currPlatbaZVypisu.nazevKlienta := asgMain.Cells[5, asgMain.row]; //do pøíslušného objektu platby zapíšu zmìnìný text
  end;
end;
procedure TfmMain.asgMainCheckBoxClick(Sender: TObject; ACol, ARow: Integer;
  State: Boolean);
begin
  asgMain.row := ARow;
  urciCurrPlatbaZVypisu();
  currPlatbaZVypisu.jePotvrzenoUzivatelem := State;
  //Memo1.Lines.Add(BoolToStr(State,true));

end;

procedure TfmMain.asgPredchoziPlatbyButtonClick(Sender: TObject; ACol,
  ARow: Integer);
begin
  urciCurrPlatbaZVypisu();
  currPlatbaZVypisu.VS := TPredchoziPlatba(currPlatbaZVypisu.PredchoziPlatbyList[ARow - 1]).VS;
  provedAkcePoZmeneVS;
end;
procedure TfmMain.asgMainKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
 //showmessage('Stisknuto: ' + IntToStr(Key));
  if Key = 27 then
  begin
    currPlatbaZVypisu.VS := currPlatbaZVypisu.VS_orig;
    provedAkcePoZmeneVS;
  end;
end;

procedure TfmMain.chbVsechnyDokladyClick(Sender: TObject);
begin
  currPlatbaZVypisu.vsechnyDoklady := chbVsechnyDoklady.Checked;
  currPlatbaZVypisu.loadDokladyPodleVS();
  vyplnDoklady;
end;

procedure TfmMain.btnSparujPlatbyClick(Sender: TObject);
begin
  sparujVsechnyPrichoziPlatby;
end;

procedure TfmMain.Zprava(TextZpravy: string);
// do listboxu a logfile uloží èas a text zprávy
begin
  Memo1.Lines.Add(FormatDateTime('dd.mm.yy hh:nn  ', Now) + TextZpravy);
  {lbxLog.ItemIndex := lbxLog.Count - 1;
  Application.ProcessMessages;
  Append(F);
  Writeln (F, FormatDateTime('dd.mm.yy hh:nn  ', Now) + TextZpravy);
  CloseFile(F);  }
end;
procedure TfmMain.zpravaRozdilCasu(cas01, cas02 : double; textZpravy : string);
// do listboxu vypíše èasový rozdíl a zprávu
begin
  debugRozdilCasu(cas01, cas02, textZpravy);
  cas01 := cas01 * 24 * 3600;
  cas02 := cas02 * 24 * 3600;
  //Memo1.Lines.Add('Èas1: ' + floattostr(RoundTo(cas01, -2)));
  //Memo1.Lines.Add('Èas2: ' + floattostr(RoundTo(cas02, -2)));
  Memo1.Lines.Add('Trvání: ' + floattostr(RoundTo(cas02 - cas01, -2))
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
  urciCurrPlatbaZVypisu();
  currPlatbaZVypisu.VS := currPlatbaZVypisu.VS_orig;
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
procedure TfmMain.btnZavritVypisClick(Sender: TObject);
begin
  asgMain.Visible := false;
  asgMain.ClearNormalCells;
  asgPredchoziPlatby.ClearNormalCells;
  asgPredchoziPlatbyVs.ClearNormalCells;
  asgNalezeneDoklady.ClearNormalCells;
  lblHlavicka.Caption := '';
  DesU.dbAbra.Reconnect;
end;

procedure TfmMain.btnCustomersClick(Sender: TObject);
begin
  fmCustomers.Show;
end;

end.
