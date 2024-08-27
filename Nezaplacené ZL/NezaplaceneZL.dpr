program NezaplaceneZL;

uses
  Forms,
  NZL in 'NZL.pas' {fmZL},
  NZL2D in 'NZL2D.pas' {fmZLDetail},
  AArray in '..\DE$_Common\AArray.pas',
  AbraEntities in '..\DE$_Common\AbraEntities.pas',
  DesUtils in '..\DE$_Common\DesUtils.pas' {DesU},
  Superdate in '..\DE$_Common\Superdate.pas',
  Superobject in '..\DE$_Common\Superobject.pas',
  Supertimezone in '..\DE$_Common\Supertimezone.pas',
  Supertypes in '..\DE$_Common\Supertypes.pas',
  Vcl.Themes,
  Vcl.Styles;

{$R *.res}

begin
  Application.Initialize;
  TStyleManager.TrySetStyle('Iceberg Classico');
  Application.CreateForm(TfmZL, fmZL);
  Application.CreateForm(TfmZLDetail, fmZLDetail);
  Application.CreateForm(TDesU, DesU);
  Application.Run;
end.
