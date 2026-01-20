program IFIAbra;

uses
  Forms,
  F1 in 'F1.pas' {fmMain},
  Code in 'Code.pas' {dmCode: TDataModule},
  Login in 'Login.pas' {fmLogin},
  AArray in '..\DE$_Common\AArray.pas',
  AbraEntities in '..\DE$_Common\AbraEntities.pas',
  DesUtils in '..\DE$_Common\DesUtils.pas' {DesU},
  DesFastReports in '..\DE$_Common\DesFastReports.pas' {DesFastReport};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfmMain, fmMain);
  Application.CreateForm(TdmCode, dmCode);
  Application.CreateForm(TfmLogin, fmLogin);
  Application.CreateForm(TDesU, DesU);
  Application.CreateForm(TDesFastReport, DesFastReport);
  Application.Run;
end.
