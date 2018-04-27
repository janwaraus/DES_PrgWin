// 27.12.10 Program pro jednor�zovou fakturaci z datab�ze iQuestu. Vyp�e ��dosti o fakturaci, po v�b�ru zkontroluje existenci z�kazn�ka
// v Ab�e a vytvo�� fakturu.
// 7.1.11 Pokud je instalace kompletn� uhrazena v hotovosti, vyrob� se PP
// 12.12.14 �pravy pro pou�it� v ABAku
// 7.15 XE5, dal�� �pravy
// 10.15 p�estav�no, procedury v Code


unit F1;

interface

uses
{
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, StdCtrls, ExtCtrls, ComCtrls, Variants,
  DateUtils, Math, Grids, BaseGrid, AdvGrid, AdvEdit, AdvObj, AdvDateTimePicker, IniFiles,
  DB, DBTables, ZAbstractRODataset, ZAbstractDataset, ZDataset, ZConnection, ZAbstractConnection;
}
  Windows, Messages, Dialogs, SysUtils, Variants, Classes, Graphics, Controls, StdCtrls, ExtCtrls, Forms, Mask, ComObj, ComCtrls,
  AdvObj, AdvPanel, AdvEdit, AdvSpin, AdvDateTimePicker, AdvEdBtn, AdvFileNameEdit, AdvProgressBar, GradientLabel,
  Grids, BaseGrid, AdvGrid, pCore2D, pBarcode2D, IniFiles, DateUtils, Math,
  DB, ZAbstractConnection, ZConnection, ZAbstractRODataset, ZAbstractDataset, ZDataset;

type
  TfmMain = class(TForm)
    qrAbra: TZQuery;
    qrMain: TZQuery;
    qrItems: TZQuery;
    pnTop: TPanel;
    asgMain: TAdvStringGrid;
    asgItems: TAdvStringGrid;
    adeDatumDokladu: TAdvDateTimePicker;
    adeDatumPlneni: TAdvDateTimePicker;
    aedDoklad: TAdvEdit;
    aedDL: TAdvEdit;
    aedTechnik: TAdvEdit;
    aedHotovost: TAdvEdit;
    cbImport: TCheckBox;
    cbOprava: TCheckBox;
    cbDoklad: TCheckBox;
    cbDL: TCheckBox;
    cbXLS: TCheckBox;
    cbClear: TCheckBox;
    btStart: TButton;
    lbxLog: TListBox;
    btReload: TButton;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure asgMainCheckBoxClick(Sender: TObject; ACol, ARow: Integer; State: Boolean);
    procedure btStartClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure asgMainCanClickCell(Sender: TObject; ARow, ACol: Integer; var Allow: Boolean);
    procedure asgMainCanEditCell(Sender: TObject; ARow, ACol: Integer; var CanEdit: Boolean);
    procedure btReloadClick(Sender: TObject);
    procedure asgItemsEditingDone(Sender: TObject);
  private
    Poprve: boolean;
  public
    F: TextFile;
    AbraConnection: string;
    AbraOLE: variant;
    PP_Id,
    FO_Id,
    DL_Id,
    PR_Id,
    User_Id: string[10];
  end;

var
  fmMain: TfmMain;

const
  Ap = char(39);
  ApC = char(39) + ',';
  ApZ = char(39) + ')';

  MyAddress_Id: string[10] = '7000000101';
  MyAccount_Id: string[10] = '1400000101';       // Fio
  MyPayment_Id: string[10] = '1000000101';       // typ platby: na bankovn� ��et


implementation

uses DesUtils, DesFrxUtils, AbraEntities, AArray,
  Login, Code;

{$R *.DFM}

// ------------------------------------------------------------------------------------------------

procedure TfmMain.FormCreate(Sender: TObject);
begin
  Poprve := True;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.FormShow(Sender: TObject);
// inicializace, z�sk�n� p�ipojovac�ch informac� pro datab�ze
var
  FileHandle: integer;
  FIIni: TIniFile;
  LogDir,
  LogFileName,
  FIFileName: ShortString;
begin
  adeDatumDokladu.Date := Date;
  adeDatumPlneni.Date := Date;
// adres�� pro logy
  LogDir := ExtractFilePath(ParamStr(0)) + '\Jednor�zov� fakturace - logy';
  if not DirectoryExists(LogDir) then CreateDir(LogDir);
  LogFileName := LogDir + FormatDateTime('\yyyy.mm".log"', Date);
// vytvo�en� logfile, pokud neexistuje
  if not FileExists(LogFileName) then begin
    FileHandle := FileCreate(LogFileName);
    FileClose(FileHandle);
  end;
  AssignFile(F, LogFileName);
  Append(F);
  Writeln(F);
  CloseFile(F);
  dmCode.Zprava('Start programu "Jednor�zov� fakturace"');
  Screen.Cursor := crHourGlass;
  Application.ProcessMessages;

  asgMain.CheckFalse := '0';
  asgMain.CheckTrue := '1';
  asgItems.CheckFalse := '0';
  asgItems.CheckTrue := '1';

  with qrAbra do begin

    SQL.Text := 'SELECT Id FROM DocQueues WHERE Code = ''FO'' AND DocumentType = ''03'' AND Hidden = ''N'' ';

    Open;
    FO_Id := FieldByName('Id').AsString;
    Close;
    SQL.Text := 'SELECT Id FROM DocQueues WHERE Code = ''PP'' AND DocumentType = ''05'' AND Hidden = ''N'' ';
    Open;
    PP_Id := FieldByName('Id').AsString;
    Close;
    SQL.Text := 'SELECT Id FROM DocQueues WHERE Code = ''DL'' AND DocumentType = ''21'' AND Hidden = ''N'' ';
    Open;
    DL_Id := FieldByName('Id').AsString;
    Close;
    SQL.Text := 'SELECT Id FROM DocQueues WHERE Code = ''PR'' AND DocumentType = ''20'' AND Hidden = ''N'' ';
    Open;
    PR_Id := FieldByName('Id').AsString;
    Close;
  end;
  dmCode.Zprava('OK');

end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.FormActivate(Sender: TObject);
// aby bylo vid�t, �e se n�co d�je, p�ipojuj� se datab�ze a� tady, kdy u� je vid�t fmMain, proto�e OnActivate nastane i p�i p�epnut�
// z jin�ho okna, pou��v� se p��znak Poprv�
begin
  if not Poprve then Exit;             // procedura jen p�i startu programu
  Poprve := False;


// na�te z�kazn�ky ze v�ech nevy��zen�ch ��dost� v tabulce simple_billings
  if dmCode.FillasgMain > 0 then with asgMain do begin               // je alespo� jeden ��dek v asgMain
// vybere se prvn� ��dek asgMain, pokud v n�m n�co je
    Row := 1;
    Ints[0, 1] := 1;
    dmCode.FillasgItems(1);
// kdy� nic, tak vy�istit asgItems a upravit velikost
  end else with asgItems do begin
    asgMain.RowCount := 2;
    ClearNormalCells;
    RowCount := 2;
    pnTop.Height := pnTop.Constraints.MinHeight;
    fmMain.ClientHeight := fmMain.Constraints.MinHeight + lbxLog.Items.Count * 13;
    btStart.Enabled := False;
  end;  // if
  Screen.Cursor := crDefault;
end;



// ------------------------------------------------------------------------------------------------

procedure TfmMain.asgMainCanClickCell(Sender: TObject; ARow, ACol: Integer; var Allow: Boolean);
begin
  Allow := (ACol = 0) or (ACol = 2) or (ARow = 0);
  if (ACol = 2) then begin
    asgMain.Ints[0, ARow] := 1;
    asgMainCheckBoxClick(nil, 0, ARow, True);
  end;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.asgMainCanEditCell(Sender: TObject; ARow, ACol: Integer; var CanEdit: Boolean);
begin
  CanEdit := (ACol = 0) or (ACol = 2);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.asgMainCheckBoxClick(Sender: TObject; ACol, ARow: Integer; State: Boolean);
// hl�d�, aby byl ozna�en� jen jeden ��dek v asgMain
var
  Radek: integer;
begin
// pro vybran� ��dek se vypln� asgItems, ostatn� se odzna��
  if State then with asgMain do begin
    for Radek := 1 to RowCount-1 do
      if Radek <> ARow then Ints[0, Radek] := 0;
    Row := ARow;
    dmCode.FillasgItems(ARow);
    dmCode.Zprava(Cells[3, ARow]);     // jm�no z�kazn�ka
// jinak se vy�ist�
  end else with asgItems do begin
    ClearNormalCells;
    RowCount := 2;
    pnTop.Height := pnTop.Constraints.MinHeight;
    fmMain.ClientHeight := fmMain.Constraints.MinHeight + lbxLog.Items.Count * 13;
  end;
  btStart.Enabled := State;            // kdy� nen� nic vybran�ho, nejde btStart
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.asgItemsEditingDone(Sender: TObject);
// po oprav� se znovu spust� kontrola druh� p��jmu a vyhled�v�n� zbo��
begin
  if asgItems.ColCount = 2 then Exit;
  dmCode.FillItemsZRowsPR;
  dmCode.FillItemsZRowsDL;
// kontrola vypln�n� asgItems
  with asgItems do
    if (Ints[0, Row] = 1) and (Ints[4, Row] = 0) then          // ��dek se m� zpracovat
      if (Cells[3, Row] = '') then begin                       // druh nen� vypln�n
        Application.MessageBox(PChar(Cells[1, Row] + ' - druh p��jmu nen� vypln�n.'), 'Druh p��jmu', MB_ICONERROR + MB_OK);
        Ints[0, Row] := 0;
      end else                                                       // druh je vypln�n
        if (UpperCase(Cells[3, Row][1]) = 'S') or (UpperCase(Cells[3, Row][1]) = 'K')
         or (UpperCase(Cells[3, Row][1]) = 'Z') then
          Cells[3, Row] := UpperCase(char(Cells[3, Row][1]))
        else

          Application.MessageBox(PChar(Cells[1, Row] + ' - druh p��jmu je S, K nebo Z.'), 'Druh p��jmu', MB_ICONERROR + MB_OK);

end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.btStartClick(Sender: TObject);
// vybran� z�kazn�k z asgMain vlo�� do adres��e Abry (pokud tam nen�), vytvo�� se mu faktura a DL
var
  Radek: integer;
begin
// najde se vybran� ��dek v asgMain
  Radek := 1;
  with asgMain do begin
    while (Ints[0, Radek] = 0) and (Radek < RowCount) do Inc(Radek);
    if (Radek = RowCount) then Exit;
  end;  // with asgMain
  with asgItems, dmCode do try
// pro jistotu
    btStart.Enabled := False;
    btReload.Enabled := False;
// vytvo�en� firmy v Ab�e
    if cbImport.Checked then begin
      ConnectAbra;
      MakeAddress(Radek);
      DESU.dbAbra.Reconnect;
      FillasgMain;
      asgMain.Ints[0, Radek] := 1;
      cbImport.Checked := False;
      if not cbOprava.Checked then begin
         FillasgItems(Radek);
         Exit;
      end;
    end;
// oprava z�kazn�ka v datab�zi
    if cbOprava.Checked then begin
      RepairRecord(Radek);
      FillasgItems(Radek);
      cbOprava.Checked := False;
      Exit;
    end;
// vyskladn�n� �eleza
    if cbDL.Checked then begin
      ConnectAbra;
      MakeDL(Radek);
      MakePR(Radek);
      cbDL.Checked := False;
    end;
// vytvo�en� PP nebo FO
    if cbDoklad.Checked then begin
      ConnectAbra;
// cel� zaplaceno hotov�, bude PP
      if aedHotovost.Text = '0,00' then MakePP(Radek)
// nen� uhrazeno v�echno, vystavuje se faktura
      else MakeFO(Radek);
      cbDoklad.Checked := False;
    end;
// hotovost do technici.xls
    if cbXLS.Checked then begin
      UpdateTechniciXLS(Radek);
      cbXLS.Checked := False;
    end;
// vymaz�n� ��dost� o fakturaci
    if cbClear.Checked then begin
      ClearRequest(Radek);
      FillasgMain;
// v asgMain je�t� n�co zb�v�
      if (asgMain.Cells[2, 1] <> '') then begin      // VS
        Radek := 1;
        asgMain.Ints[0, Radek] := 1;
        FillasgItems(Radek);
      end;
      cbClear.Checked := False;
    end;
  finally
    Screen.Cursor := crDefault;
    btStart.Enabled := True;
    btReload.Enabled := True;
  end;  //  with asgMain
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.btReloadClick(Sender: TObject);
// znovu na�te simple_billings
begin
  DesU.dbZakos.Reconnect;
  DESU.dbAbra.Reconnect;
  dmCode.FillasgMain;
  if dmCode.FillasgMain > 0 then with asgMain do begin               // je alespo� jeden ��dek v asgMain
// vybere se prvn� ��dek asgMain, pokud v n�m n�co je
    if RowCount > 1 then begin
      Ints[0, 1] := 1;
      dmCode.FillasgItems(1);
// kdy� nic, tak vy�istit asgItems a upravit velikost
    end else with asgItems do begin
      ClearNormalCells;
      RowCount := 2;
      fmMain.ClientHeight := fmMain.Constraints.MinHeight + lbxLog.Height - lbxLog.Constraints.MinHeight;
      pnTop.Height := pnTop.Constraints.MinHeight;
      btStart.Enabled := False;
    end;  // if
  end;  // with
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if not VarIsEmpty(AbraOLE) then try
    AbraOLE.Logout;
  except
  end;
  dmCode.Zprava('Konec programu "Jednor�zov� fakturace".');
end;

end.
