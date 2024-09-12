program Hodinovnik;

uses
  Forms,
  Hodiny in 'Hodiny.pas' {fmMain},
  Vcl.Themes,
  Vcl.Styles,
  DES_common in '..\DE$_Common\DES_common.pas' {dmDES_common: TDataModule},
  DES_OO_common in '..\DE$_Common\DES_OO_common.pas' {dmDES_OO_Common: TDataModule};

{$R *.res}

begin
  Application.Initialize;
  TStyleManager.TrySetStyle('Silver');
  Application.CreateForm(TfmMain, fmMain);
  Application.CreateForm(TdmDES_common, dmDES_common);
  Application.CreateForm(TdmDES_OO_Common, dmDES_OO_Common);
  Application.Run;
end.
