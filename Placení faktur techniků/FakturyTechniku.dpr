program FakturyTechniku;

uses
  Vcl.Forms,
  FO in 'FO.pas' {fmMain},
  DES_OO_common in '..\DE$_Common\DES_OO_common.pas' {dmDES_OO_Common: TDataModule},
  DES_common in '..\DE$_Common\DES_common.pas' {dmDES_common: TDataModule},
  Vcl.Themes,
  Vcl.Styles;

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  TStyleManager.TrySetStyle('Iceberg Classico');
  Application.CreateForm(TfmMain, fmMain);
  Application.CreateForm(TdmDES_OO_Common, dmDES_OO_Common);
  Application.CreateForm(TdmDES_common, dmDES_common);
  Application.Run;
end.
