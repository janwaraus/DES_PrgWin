program FOx;

uses
  Forms,
  FOx_mail in 'FOx_mail.pas' {dmMail: TDataModule},
  FOx_main in 'FOx_main.pas' {fmMain},
  FOx_common in 'FOx_common.pas' {dmCommon: TDataModule},
  FOx_tisk in 'FOx_tisk.pas' {dmTisk: TDataModule},
  FOx_prevod in 'FOx_prevod.pas' {dmPrevod: TDataModule},
  AbraEntities in '..\DE$_Common\AbraEntities.pas',
  DesFastReports in '..\DE$_Common\DesFastReports.pas' {DesFastReport},
  DesInvoices in '..\DE$_Common\DesInvoices.pas',
  DesUtils in '..\DE$_Common\DesUtils.pas' {DesU};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TdmMail, dmMail);
  Application.CreateForm(TfmMain, fmMain);
  Application.CreateForm(TdmCommon, dmCommon);
  Application.CreateForm(TdmTisk, dmTisk);
  Application.CreateForm(TdmPrevod, dmPrevod);
  Application.CreateForm(TDesFastReport, DesFastReport);
  Application.CreateForm(TDesU, DesU);
  Application.Run;
end.
