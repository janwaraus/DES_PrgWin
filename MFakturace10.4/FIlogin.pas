unit FIlogin;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, StdCtrls, AdvEdit, AdvCombo, FImain;

type
  TfmLogin = class(TForm)
    acbJmeno: TAdvComboBox;
    aedHeslo: TAdvEdit;
    btOK: TButton;
    procedure FormShow(Sender: TObject);
    procedure acbJmenoSelect(Sender: TObject);
    procedure aedHesloKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure btOKClick(Sender: TObject);
  end;

var
  fmLogin: TfmLogin;

implementation

{$R *.dfm}

uses DesUtils, FIcommon;

procedure TfmLogin.FormShow(Sender: TObject);
var
  SQLStr: AnsiString;
begin
  acbJmeno.Clear;
  aedHeslo.Clear;
  with fmMain, DesU.qrAbra do begin
    Close;
    SQL.Text := 'SELECT LoginName FROM SecurityUsers'
    + ' WHERE Locked = ''N'' '
    + ' ORDER BY LoginName';
    Open;
    while not EOF do begin
      acbJmeno.Items.Append(Fields[0].AsString);
      Next;
    end;
    Close;
  end;
end;

procedure TfmLogin.acbJmenoSelect(Sender: TObject);
var
  SQLStr: AnsiString;
begin
  with fmMain, DesU.qrAbra do begin
    Close;
    SQL.Text := 'SELECT ID FROM SecurityUsers'
    + ' WHERE LoginName = ''' + acbJmeno.Text + '''';
    Open;
    fmMain.User_Id := Fields[0].AsString;
  end;
  aedHeslo.SetFocus;
end;

procedure TfmLogin.aedHesloKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = 13 then fmLogin.Close;
end;

procedure TfmLogin.btOKClick(Sender: TObject);
begin
  fmLogin.Close;
end;

end.
