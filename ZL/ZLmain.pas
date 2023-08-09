// od 19.12.2014 - Z�lohov� listy na p�ipojen� se generuj� v Ab�e podle datab�ze iQuest. Podle M�s��n� fakturace.
// 27.11.17 pro tarify s�t� Misauer a Harna nejsou slevy
// 28.6.18 p�id�ny zak�zky a obchodn� p��pady pro statistiky �T�
// 16.9. pro MySQL set autocommit = 1
// 30.9. zru�eny testovac� view
// 21.8.2022 �pravy kv�li Ab�e 22.1 s UTF8
// 5.3.23 Maily s TLS

unit ZLmain;

interface

uses
  Windows, Messages, Dialogs, SysUtils, Variants, Classes, Graphics, Controls, StdCtrls, ExtCtrls, Forms, Mask, ComObj, ComCtrls,
  AdvObj, AdvPanel, AdvEdit, AdvSpin, AdvDateTimePicker, AdvEdBtn, AdvFileNameEdit, AdvProgressBar, GradientLabel,
  AdvOfficeButtons, AdvGroupBox, Grids, BaseGrid, AdvGrid,DateUtils, Math, Registry,
  DB, ZAbstractConnection, ZConnection, ZAbstractRODataset, ZAbstractDataset, ZDataset,
  frxClass, frxDBSet, frxDesgn, AdvUtil;

type
  TfmMain = class(TForm)
    apnTop: TAdvPanel;
    apnVyberCinnosti: TAdvPanel;
    glbVytvoreni: TGradientLabel;
    glbPrevod: TGradientLabel;
    glbTisk: TGradientLabel;
    glbMail: TGradientLabel;
    arbVytvoreni: TAdvOfficeRadioButton;
    arbPrevod: TAdvOfficeRadioButton;
    arbTisk: TAdvOfficeRadioButton;
    arbMail: TAdvOfficeRadioButton;
    apbProgress: TAdvProgressBar;
    apnMain: TAdvPanel;
    argMesic: TAdvOfficeRadioGroup;
    aseRok: TAdvSpinEdit;
    deDatumDokladu: TAdvDateTimePicker;
    deDatumSplatnosti: TAdvDateTimePicker;
    aedOd: TAdvEdit;
    aedDo: TAdvEdit;
    btVytvorit: TButton;
    btKonec: TButton;
    apnPrevod: TAdvPanel;
    cbNeprepisovat: TCheckBox;
    btOdeslat: TButton;
    btSablona: TButton;
    apnTisk: TAdvPanel;
    apnMail: TAdvPanel;
    fePriloha: TAdvFileNameEdit;
    lbxLog: TListBox;
    asgMain: TAdvStringGrid;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure glbVytvoreniClick(Sender: TObject);
    procedure glbPrevodClick(Sender: TObject);
    procedure glbTiskClick(Sender: TObject);
    procedure glbMailClick(Sender: TObject);
    procedure arbVytvoreniClick(Sender: TObject);
    procedure arbPrevodClick(Sender: TObject);
    procedure arbTiskClick(Sender: TObject);
    procedure arbMailClick(Sender: TObject);
    procedure argMesicClick(Sender: TObject);
    procedure aseRokChange(Sender: TObject);
    procedure deDatumDokladuChange(Sender: TObject);
    procedure deDatumSplatnostiChange(Sender: TObject);
    procedure aedOdChange(Sender: TObject);
    procedure aedDoChange(Sender: TObject);
    procedure aedOdKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure aedOdExit(Sender: TObject);
    procedure aedDoKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure aedDoExit(Sender: TObject);
    procedure asgMainCanEditCell(Sender: TObject; ARow, ACol: Integer; var CanEdit: Boolean);
    procedure asgMainClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure asgMainCanSort(Sender: TObject; ACol: Integer; var DoSort: Boolean);
    procedure asgMainGetAlignment(Sender: TObject; ARow, ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asgMainDblClick(Sender: TObject);
    procedure lbxLogDblClick(Sender: TObject);
    procedure btVytvoritClick(Sender: TObject);
    procedure btOdeslatClick(Sender: TObject);
    procedure btSablonaClick(Sender: TObject);
    procedure btKonecClick(Sender: TObject);

  private
    LogDir : string;

  public
    Prerusit: boolean;
    VystavenoOd,
    VystavenoDo: double;

    //ID,
    //Period_Idsmazat,
    //DocQueue_Id: string[10];
    //PDFDir,
    //BBmax,
    //BillingView,
    //ZLView: ShortString;


    procedure Zprava(TextZpravy: string);
    procedure PlneniAsgMain;
  end;

var
  fmMain: TfmMain;

implementation

uses ZLvytvoreni, ZLprevod, ZLtisk, ZLmail, DesUtils, AbraEntities, DesFastReports;

{$R *.dfm}



// ------------------------------------------------------------------------------------------------

procedure TfmMain.FormCreate(Sender: TObject);
begin
  Prerusit := False;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.FormShow(Sender: TObject);
// inicializace, z�sk�n� p�ipojovac�ch informac� pro datab�ze, p�ipojen�
var
  Month: integer;
begin
// jm�na pro view jsou unik�tn�, aby program nebyl omezen na jednu instanci
  //BBmax := FormatDateTime('BByymmddhhnnss', Now);
  //BillingView := FormatDateTime('BVyymmddhhnnss', Now);
  //ZLView := FormatDateTime('"ZV"yymmddhhnnss', Now);
// jm�no FI.ini
  DesU.programInit('Z�lohov� listy');
  DesUtils.appendToFile(DesU.LOG_FILENAME, ''); //vlo�� pr�zdn� ��dek do logu
  fmMain.Zprava('Start programu "Z�lohov� listy".');

  Prerusit := True; // p��znak startu

  // fajfky v asgMain
  asgMain.CheckFalse := '0';
  asgMain.CheckTrue := '1';

  // p�edvypln�n� formul��e
  aseRok.Value := YearOf(Date); // aktu�ln� rok

  // 2.-4. m�s�c = index 0; 5.-7. m�s�c = index 1; 8.-10. m�s�c = index 2; 11.-1. m�s�c = index 3;
  Month := MonthOf(Date);
  if Month = 1 then begin
    Month := 13;
    aseRok.Value := aseRok.Value - 1;
  end;
  argMesic.ItemIndex := Floor((Month+1)/3) - 1;             // tolerance -2 / +1 m�s�c

  //deDatumDokladu.Date := Trunc(StrToDate(Format('15.%d.%d', [3*(argMesic.ItemIndex + 1), aseRok.Value])));
  //deDatumSplatnosti.Date := Trunc(EndOfTheMonth(deDatumDokladu.Date));
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.FormActivate(Sender: TObject);
begin
  if not Prerusit then Exit;             // to by m�lo ohl�dat vol�n� jen p�i startu

  aseRokChange(nil);
  arbVytvoreniClick(nil);
end;



// ------------------------------------------------------------------------------------------------

procedure TfmMain.arbVytvoreniClick(Sender: TObject);
begin
  asgMain.Cells[0, 0] := 'ZL1';
  asgMain.ColWidths[6] := 0;
  glbVytvoreni.Color := clWhite;
  glbVytvoreni.ColorTo := clMenu;
  glbPrevod.Color := clSilver;
  glbPrevod.ColorTo := clGray;
  apnPrevod.Visible := False;
  glbTisk.Color := clSilver;
  glbTisk.ColorTo := clGray;
  apnTisk.Visible := False;
  glbMail.Color := clSilver;
  glbMail.ColorTo := clGray;
  apnMail.Visible := False;
  aedOd.LabelCaption := 'od VS';
  aedDo.LabelCaption := 'do VS';
  apbProgress.Visible := False;
  lbxLog.Visible := True;
  aseRokChange(Self);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.arbPrevodClick(Sender: TObject);
begin
  asgMain.Cells[0, 0] := 'PDF';
  asgMain.ColWidths[6] := 0;
  glbPrevod.Color := clWhite;
  glbPrevod.ColorTo := clMenu;
  apnPrevod.Visible := True;
  glbTisk.Color := clSilver;
  glbTisk.ColorTo := clGray;
  apnTisk.Visible := False;
  glbMail.Color := clSilver;
  glbMail.ColorTo := clGray;
  apnMail.Visible := False;
  glbVytvoreni.Color := clSilver;
  glbVytvoreni.ColorTo := clGray;
  aedOd.LabelCaption := 'od ZL1';
  aedDo.LabelCaption := 'do ZL1';
  aseRokChange(Self);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.arbTiskClick(Sender: TObject);
begin
  asgMain.Cells[0, 0] := 'tisk';
  asgMain.ColWidths[6] := 60;
  glbTisk.Color := clWhite;
  glbTisk.ColorTo := clMenu;
  apnTisk.Visible := True;
  glbMail.Color := clSilver;
  glbMail.ColorTo := clGray;
  apnMail.Visible := False;
  glbVytvoreni.Color := clSilver;
  glbVytvoreni.ColorTo := clGray;
  glbPrevod.Color := clSilver;
  glbPrevod.ColorTo := clGray;
  apnPrevod.Visible := False;
  aedOd.LabelCaption := 'od ZL1';
  aedDo.LabelCaption := 'do ZL1';
  aseRokChange(Self);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.arbMailClick(Sender: TObject);
begin
  asgMain.Cells[0, 0] := 'mail';
  asgMain.ColWidths[6] := 60;
  glbMail.Color := clWhite;
  glbMail.ColorTo := clMenu;
  apnMail.Visible := True;
  glbVytvoreni.Color := clSilver;
  glbVytvoreni.ColorTo := clGray;
  glbPrevod.Color := clSilver;
  glbPrevod.ColorTo := clGray;
  apnPrevod.Visible := False;
  glbTisk.Color := clSilver;
  glbTisk.ColorTo := clGray;
  apnTisk.Visible := False;
  aedOd.LabelCaption := 'od ZL1';
  aedDo.LabelCaption := 'do ZL1';
  aseRokChange(Self);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.glbVytvoreniClick(Sender: TObject);
begin
  arbVytvoreni.Checked := True;
  arbVytvoreniClick(nil);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.glbPrevodClick(Sender: TObject);
begin
  arbPrevod.Checked := True;
  arbPrevodClick(nil);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.glbTiskClick(Sender: TObject);
begin
  arbTisk.Checked := True;
  arbTiskClick(nil);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.glbMailClick(Sender: TObject);
begin
  arbMail.Checked := True;
  arbMailClick(nil);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.argMesicClick(Sender: TObject);
begin
  aseRokChange(nil);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.aseRokChange(Sender: TObject);
// p�i zm�n� (nejen) roku nastav� nov� deDatumDokladu, deDatumPlneni, aedSplatnost, aedOd a aedDo
var
SQLStr: AnsiString;
begin
  aedOd.Clear;
  aedDo.Clear;
  asgMain.ClearNormalCells;
  asgMain.RowCount := 2;
  btVytvorit.Caption := '&Na��st';
  asgMain.Visible := True;
  lbxLog.Visible := False;
  {
  // Id obdob�
  with qrAbra do begin
    Close;
    SQL.Text := 'SELECT Id FROM Periods WHERE Code = ' + Ap + aseRok.Text + Ap;
    Open;
    Period_Idsmazat := FieldByName('Id').AsString;
    Close;
  end;  // with qrAbra
  }
  // datum dokladu je 15.3., 15.6., 15.9. a 15.12., datum splatnosti posledn� den v t�m�e m�s�ci
  deDatumDokladu.Date := Trunc(StrToDate(Format('15.%d.%d', [3*(argMesic.ItemIndex+1), aseRok.Value])));
  deDatumSplatnosti.Date := Trunc(EndOfTheMonth(deDatumDokladu.Date));
  // VystavenoOd pro hled�n� vystaven�ch ZL je 1.3., 1.6., 1.9., 1.12.,
  VystavenoOd := StrToDate(Format('1.%d.%d', [3*(argMesic.ItemIndex + 1 ), aseRok.Value]));
  // VystavenoDo pro hled�n� vystaven�ch ZL je 30.5., 31.8., 30.11. a 28/29.2. v dal��m roce,
  VystavenoDo := EndOfTheMonth(IncMonth(VystavenoOd, 2));

  // *** v�b�r ZL podle VS
  if arbVytvoreni.Checked then begin
    with DesU.qrZakos do try
      Screen.Cursor := crSQLWait;
      // prvn� a posledn� VS
      SQL.Text := 'CALL get_deposit_invoicing_minmaxvs('
        + Ap + FormatDateTime('yyyy-mm-dd', deDatumDokladu.Date + 30) + ApC
        + Ap + FormatDateTime('yyyy-mm-dd', deDatumSplatnosti.Date) + ApZ;
      Open;
      aedOd.Text := FieldByName('min_vs').AsString;
      aedDo.Text := FieldByName('max_vs').AsString;
      Close;
    finally
      Screen.Cursor := crDefault;
    end;
  end
  // *** v�b�r podle ZL (p�evod, tisk, mail)
  else begin
    DesU.dbAbra.Reconnect;
    with DesU.qrAbraOC do begin
      // rozp�t� ��sel ZL v obdob� VystavenoOd a� VystavenoDo
      SQLStr := 'SELECT MIN(OrdNumber) FROM IssuedDInvoices, Periods'
      + ' WHERE DocDate$DATE >= ' + FloatToStr(Trunc(VystavenoOd))
      + ' AND Periods.ID = Period_ID'
      + ' AND Periods.Code = ' + Ap + FloatToStr(YearOf(VystavenoOd)) + Ap
      + ' AND DocQueue_ID = ' + Ap + AbraEnt.getDocQueue('Code=ZL1').ID + Ap;
      Close;
      SQL.Text := SQLStr;
      Open;
      if Fields[0].AsInteger > 0 then aedOd.Text := Fields[0].AsString
      else begin                  // v roce VystavenoOd nebyl ��dn� ZL - to je asi jen p�i zah�jen� �ady ZL1
        SQLStr := 'SELECT MIN(OrdNumber) FROM IssuedDInvoices, Periods'
        + ' WHERE DocDate$DATE >= ' + FloatToStr(Trunc(VystavenoOd))
        + ' AND Periods.ID = Period_ID'
        + ' AND Periods.Code = ' + Ap + FloatToStr(YearOf(VystavenoDo)) + Ap
        + ' AND DocQueue_ID = ' + Ap + AbraEnt.getDocQueue('Code=ZL1').ID + Ap;
        Close;
        SQL.Text := SQLStr;
        Open;
        aedOd.Text := Fields[0].AsString;
      end;

      SQLStr := 'SELECT MAX(OrdNumber) FROM IssuedDInvoices, Periods'
      + ' WHERE DocDate$DATE <= ' + FloatToStr(Trunc(VystavenoDo))
      + ' AND Periods.ID = Period_ID'
      + ' AND Periods.Code = ' + Ap + FloatToStr(YearOf(VystavenoDo)) + Ap
      + ' AND DocQueue_ID = ' + Ap + AbraEnt.getDocQueue('Code=ZL1').ID + Ap;
      Close;
      SQL.Text := SQLStr;
      Open;
      if Fields[0].AsInteger > 0 then begin
        aedDo.Text := Fields[0].AsString;
        // vystaveno do m��e b�t v dal��m roce, pak se p�id� rok do aedDo.Text
        if  YearOf(VystavenoDo) > YearOf(VystavenoOd) then aedDo.Text := aedDo.Text + '/' + FloatToStr(YearOf(VystavenoDo))
      // kdy� nen� ZL, zkus� se je�t� YearOf(VystavenoOd)
      end
      else begin   // kdy� nen� ZL, zkus� se je�t� YearOf(VystavenoOd)
        SQLStr := 'SELECT MAX(OrdNumber) FROM IssuedDInvoices, Periods'
        + ' WHERE DocDate$DATE <= ' + FloatToStr(Trunc(VystavenoDo))
        + ' AND Periods.ID = Period_ID'
        + ' AND Periods.Code = ' + Ap + FloatToStr(YearOf(VystavenoOd)) + Ap
        + ' AND DocQueue_ID = ' + Ap + AbraEnt.getDocQueue('Code=ZL1').ID + Ap;
        Close;
        SQL.Text := SQLStr;
        Open;
        aedDo.Text := Fields[0].AsString;
      end;
      Close;
    end;  // with qrAbra
  end;  // if ... else arbVytvoreni.Checked
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.deDatumDokladuChange(Sender: TObject);
begin
  //DatumDokladu := deDatumDokladu.Date;
//  16.9.22  deDatumSplatnosti.Date := Trunc(EndOfTheMonth(deDatumDokladu.Date));
  //DatumSplatnosti := deDatumSplatnosti.Date;

  {
  // Id obdob�
  with qrAbra do begin
    Close;
    SQL.Text := 'SELECT Id FROM Periods WHERE Code = ' + Ap + FormatDateTime('yyyy', DatumDokladu) + Ap;
    Open;
    Period_Idsmazat := FieldByName('Id').AsString;
    Close;
  end;  // with qrAbra
  }
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.deDatumSplatnostiChange(Sender: TObject);
begin
  //DatumSplatnosti := deDatumSplatnosti.Date;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.aedOdChange(Sender: TObject);
begin
  if btVytvorit.Caption <> '&Na��st' then begin
    asgMain.ClearNormalCells;
    asgMain.RowCount := 2;
    btVytvorit.Caption := '&Na��st';
  end;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.aedOdKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
// enter po zad�n� aedOd vypln� stejn�m ��slem aedDo
begin
  if Key = 13 then begin
    aedDo.Text := aedOd.Text;
    aedDo.SetFocus;
  end;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.aedOdExit(Sender: TObject);
// je-li aedOd v�t�� ne� aedDo, oprav� se
begin
// 15.12.2020 oprava - ne VS
  if (aedOd.LabelCaption = 'od ZL1') and (aedOd.IntValue > aedDo.IntValue) then aedDo.Text := aedOd.Text;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.aedDoChange(Sender: TObject);
begin
  aedOdChange(nil);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.aedDoKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
// enter po zad�n� aedDo spust� p�evod
begin
  if Key = 13 then btVytvorit.SetFocus;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.aedDoExit(Sender: TObject);
// je-li aedDo men�� ne� aedOd, oprav� se
begin
// 15.12.2020 oprava - ne VS
  if (aedOd.LabelCaption = 'od ZL1') and (aedDo.IntValue < aedOd.IntValue) then aedOd.Text := aedDo.Text;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.asgMainDblClick(Sender: TObject);
begin
  asgMain.Visible := False;
  lbxLog.Visible := True;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.asgMainCanEditCell(Sender: TObject; ARow, ACol: Integer; var CanEdit: Boolean);
begin
  CanEdit := (ARow > 0) and (ACol in [0, 5]);   // fajfky nebo email
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.asgMainClickCell(Sender: TObject; ARow, ACol: Integer);
// klik v prvn�m ��dku podle sloupce bu� ozna�uje/odzna�uje v�chny checkboxy, nebo spou�t� t��d�n�,
var
  Radek: integer;
begin
  with asgMain do if (ARow = 0) and (ACol in [0..1]) then
    if Ints[ACol, 1] = 1 then for Radek := 1 to RowCount-1 do Ints[ACol, Radek] := 0
    else if Ints[ACol, 1] = 0 then for Radek := 1 to RowCount-1 do Ints[ACol, Radek] := 1;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.asgMainCanSort(Sender: TObject; ACol: Integer; var DoSort: Boolean);
// klik v prvn�m ��dku podle sloupce bu� ozna�uje/odzna�uje v�chny checkboxy, nebo spou�t� t��d�n�,
begin
  DoSort := ACol > 0;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.asgMainGetAlignment(Sender: TObject; ARow, ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
  HAlign := taLeftJustify;
  if (ACol = 0) or (ARow = 0) then HAlign := Classes.taCenter
  else if (ACol in [1..3]) and (ARow > 0) then HAlign := taRightJustify;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.lbxLogDblClick(Sender: TObject);
begin
  asgMain.Visible := True;
  lbxLog.Visible := False;
end;


// ------------------------------------------------------------------------------------------------

procedure TfmMain.btVytvoritClick(Sender: TObject);
begin
  btVytvorit.Enabled := False;
  btKonec.Caption := '&P�eru�it';
  Prerusit := False;
  Application.ProcessMessages;
  try
// *** Napln�n� asgMain ***
    if btVytvorit.Caption = '&Na��st' then PlneniAsgMain
    else begin
// *** Fakturace ***
      if arbVytvoreni.Checked then dmVytvoreni.VytvorZL;
// *** P�evod do PDF ***
      if arbPrevod.Checked then dmPrevod.PrevedZL;
// *** Tisk ***
      if arbTisk.Checked then dmTisk.TiskniZL;
// *** Pos�l�n� mailem ***
      if arbMail.Checked then dmMail.PosliZL;
    end;
  finally
    btKonec.Caption := '&Konec';
    btVytvorit.Enabled := True;
  end;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.btSablonaClick(Sender: TObject);
begin
  DesFastReport.frxReport.DesignReport(True, False);
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.btOdeslatClick(Sender: TObject);
// odesl�n� faktur p�eveden�ch do PDF na vzd�len� server
begin
  //WinExec(PChar(Format('WinSCP.com /command "option batch abort" "option confirm off" "open AbraPDF" "synchronize remote '
  // + PDFDir + '\%4d\%2.2d /home/abrapdf/%4d" "exit"', [aseRok.Value, 3*(argMesic.ItemIndex+1), aseRok.Value])), SW_SHOWNORMAL);

  DesU.syncAbraPdfToRemoteServer(aseRok.Value, 3*(argMesic.ItemIndex+1));
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.btKonecClick(Sender: TObject);
begin
  if btKonec.Caption = '&P�eru�it' then begin
    Prerusit := True;
    fmMain.Zprava('P�eru�eno u�ivatelem.');
    btKonec.Caption := '&Konec';
  end else begin
    fmMain.Zprava('Konec programu "Z�lohov� listy".');
    Close;
  end;
end;


// ------------------------------------------------------------------------------------------------
// funkce nespojen� p��mo s konkr�tn�m prvkem formul��e
// ------------------------------------------------------------------------------------------------


procedure TfmMain.Zprava(TextZpravy: string);
// do listboxu a logfile ulo�� �as a text zpr�vy
begin
  lbxLog.Items.Add(DesU.zalogujZpravu(TextZpravy));
  lbxLog.ItemIndex := lbxLog.Count - 1;
  Application.ProcessMessages;
end;

// ------------------------------------------------------------------------------------------------

procedure TfmMain.PlneniAsgMain;
var
  Radek: integer;
  VarSymbol,
  SQLStr: string;
begin
  with fmMain do try
    asgMain.Visible := True;
    lbxLog.Visible := False;
    apnPrevod.Visible := False;
    apnTisk.Visible := False;
    apnMail.Visible := False;
    Screen.Cursor := crHourGlass;
    with asgMain do try
      ClearNormalCells;
      RowCount := 2;

      // ************************//
      // ***  Generov�n� ZL  *** //
      // ************************//
      if arbVytvoreni.Checked then begin
        fmMain.Zprava(Format('Na�ten� z�kazn�k� od VS %s do %s.', [aedOd.Text, aedDo.Text]));

        SQLStr := 'CALL get_deposit_invoicing_cu_by_vsrange('
          + Ap + FormatDateTime('yyyy-mm-dd', deDatumDokladu.Date + 30) + ApC
          + Ap + FormatDateTime('yyyy-mm-dd', deDatumSplatnosti.Date) + ApC
          + Ap + aedOd.Text + ApC
          + Ap + aedDo.Text + ApC;


        case argMesic.ItemIndex of
          0, 2: SQLStr := SQLStr + 'false,false)'; //prvn� parametr ��k�, zda na��t�me 6m periody, druh�, zda 12m periody
          1: SQLStr := SQLStr +  'true,false)'; // dtto
          3: SQLStr := SQLStr +  'true,true)'; // dtto
        end;

        DesU.qrZakos.SQL.Text := SQLStr;
        DesU.qrZakos.Open;

        Radek := 0;
        apbProgress.Position := 0;
        apbProgress.Visible := True;
        while not DesU.qrZakos.EOF do begin
          apbProgress.Position := Round(100 * DesU.qrZakos.RecNo / DesU.qrZakos.RecordCount);
          Application.ProcessMessages;
          if Prerusit then begin
            Prerusit := False;
            apbProgress.Position := 0;
            apbProgress.Visible := False;
            btVytvorit.Enabled := True;
            btKonec.Caption := '&Konec';
            Break;
          end;
          Inc(Radek);
          RowCount := Radek + 1;
          AddCheckBox(0, Radek, True, True);
          Ints[0, Radek] := 1;                                                   // fajfka
          Cells[1, Radek] := DesU.qrZakos.FieldByName('cu_variable_symbol').AsString;    // VS
          Cells[2, Radek] := '';                                                 // ZL
          Cells[3, Radek] := '';                                                 // ��stka
          Cells[4, Radek] := DesU.qrZakos.FieldByName('cu_abra_code').AsString;            // jm�no
          Cells[5, Radek] := DesU.qrZakos.FieldByName('cu_postal_mail').AsString;                // mail
          Ints[6, Radek] := DesU.qrZakos.FieldByName('cu_disable_mailings').AsInteger;             // reklama
          Application.ProcessMessages;
          DesU.qrZakos.Next;
        end;
        DesU.qrZakos.Close;
      end
      else

      // *****************************//
      // ***  P�evod, Tisk, Mail  *** //
      // *****************************//
      with DesU.qrAbra do begin              // if arbVytvoreni.Checked
        fmMain.Zprava(Format('Na�ten� ZL od %s do %s.', [aedOd.Text, aedDo.Text]));
        Close;
        DesU.dbAbra.Reconnect;
        SQLStr := 'SELECT II.ID, F.Name as FirmName, II.OrdNumber, P.Code, II.VarSymbol, II.Amount, II.DocDate$DATE '
        +'FROM Firms F, IssuedDInvoices II, Periods P'
        + ' WHERE DocQueue_ID = ' + Ap + AbraEnt.getDocQueue('Code=ZL1').ID + Ap
        + ' AND F.ID = II.Firm_ID'
        + ' AND P.ID = II.Period_ID'
        + ' AND (P.Code = ' + Ap + FloatToStr(YearOf(VystavenoOd)) + Ap
         + ' OR P.Code = ' + Ap + FloatToStr(YearOf(VystavenoDo)) + ApZ;


        if Pos('/', aedDo.Text) = 0 then // ZL jsou z jednoho roku
        SQLStr := SQLStr
          + ' AND OrdNumber >= ' + aedOd.Text
          + ' AND OrdNumber <= ' + aedDo.Text

        else // v aedDo jsou ZL u� z dal��ho roku
        SQLStr := SQLStr + ' AND (OrdNumber <= (SELECT MAX(OrdNumber) FROM IssuedDInvoices, Periods'
            + ' WHERE DocDate$DATE >= ' + FloatToStr(Trunc(VystavenoOd))
            + ' AND Periods.ID = Period_ID'
            + ' AND Periods.Code = ' + Ap + FloatToStr(YearOf(VystavenoOd)) + Ap
            + ' AND DocQueue_ID = ' + Ap + AbraEnt.getDocQueue('Code=ZL1').ID + ApZ
          + ' OR OrdNumber <= ' + Copy(aedDo.Text, 0, Pos('/', aedDo.Text)-1) + ')';
        SQLStr := SQLStr + ' AND (DocDate$DATE >= ' + FloatToStr(Trunc(VystavenoOd))
        + ' AND DocDate$DATE <= ' + FloatToStr(Trunc(VystavenoDo))
        + ' OR DocDate$Date = ' + FloatToStr(Floor(deDatumDokladu.Date)) + ')'       // 12.5.17 - aby se dal p�ev�st i doklad mimo
        + ' ORDER BY OrdNumber';                                                     // ��dn� obdob�
        SQL.Text := SQLStr;
        Open;
        Radek := 0;
        apbProgress.Position := 0;
        apbProgress.Visible := True;
        while not EOF do begin
          VarSymbol := FieldByName('VarSymbol').AsString;
          apbProgress.Position := Round(100 * RecNo / RecordCount);
          Application.ProcessMessages;

          if Prerusit then begin
            Prerusit := False;
            apbProgress.Position := 0;
            apbProgress.Visible := False;
            btVytvorit.Enabled := True;
            btKonec.Caption := '&Konec';
            Break;
          end;

          with DesU.qrZakos do begin
            Close;
            SQL.Text := 'SELECT DISTINCT postal_mail, disable_mailings FROM customers'
            + ' WHERE variable_symbol = ' + VarSymbol;
            Open;
            if RecordCount > 0 then begin
              Inc(Radek);
              RowCount := Radek + 1;
              AddCheckBox(0, Radek, True, True);
              Ints[0, Radek] := 1;                                                   // fajfka tisk - mail
              Cells[1, Radek] := VarSymbol;                                          // VS
              Cells[2, Radek] := Format('%4.4d/%d',                                  // ZL
               [DesU.qrAbra.FieldByName('OrdNumber').AsInteger, DesU.qrAbra.FieldByName('Code').AsInteger]);
              Floats[3, Radek] := DesU.qrAbra.FieldByName('Amount').AsFloat;              // ��stka
              Cells[4, Radek] := DesU.qrAbra.FieldByName('FirmName').AsString;;                        // jm�no firmy/z�kazn�ka
              Cells[5, Radek] := FieldByName('postal_mail').AsString;                       // mail
              Ints[6, Radek] := FieldByName('disable_mailings').AsInteger;                    // reklama
              Cells[7, Radek] := DesU.qrAbra.FieldByName('DocDate$DATE').AsString;        // datum ZL
              Cells[8, Radek] := DesU.qrAbra.FieldByName('ID').AsString;             // ID faktury
              Application.ProcessMessages;
            end;
            Close;
          end;  //  with DesU.qrZakos
          Next;
        end;  // while not EOF
      end;  //  .. else with qrAbra
      fmMain.Zprava(Format('Po�et ZL: %d', [RowCount-1]));
      if arbVytvoreni.Checked then btVytvorit.Caption := '&Vytvo�it'
      else if arbPrevod.Checked then btVytvorit.Caption := '&P�ev�st'
      else if arbTisk.Checked then btVytvorit.Caption := '&Vytisknout'
      else if arbMail.Checked then btVytvorit.Caption := '&Odeslat';
    except on E: Exception do
      Zprava('Neo�et�en� chyba: ' + E.Message);
    end;  // with qrMain
  finally
    apbProgress.Position := 0;
    apbProgress.Visible := False;
    if arbPrevod.Checked then apnPrevod.Visible := True;
    if arbTisk.Checked then apnTisk.Visible := True;
    if arbMail.Checked then apnMail.Visible := True;
    Screen.Cursor := crDefault;
    btVytvorit.Enabled := True;
    btVytvorit.SetFocus;
  end;
end;

end.
