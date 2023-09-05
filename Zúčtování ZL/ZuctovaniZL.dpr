program ZuctovaniZL;

uses
  Forms,
  NZLmain in 'NZLmain.pas' {fmMain},
  NZLcommon in 'NZLcommon.pas' {dmCommon: TDataModule},
  NZLfaktura in 'NZLfaktura.pas' {dmVytvoreni: TDataModule},
  NZLprevod in 'NZLprevod.pas' {dmPrevod: TDataModule},
  NZLmail in 'NZLmail.pas' {dmMail: TDataModule},
  NZLtisk in 'NZLtisk.pas' {dmTisk: TDataModule},
  AbraEntities in '..\DE$_Common\AbraEntities.pas',
  DesFastReports in '..\DE$_Common\DesFastReports.pas' {DesFastReport},
  DesInvoices in '..\DE$_Common\DesInvoices.pas',
  DesUtils in '..\DE$_Common\DesUtils.pas' {DesU},
  AArray in '..\DE$_Common\AArray.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'Zúètování ZL';
  Application.CreateForm(TfmMain, fmMain);
  Application.CreateForm(TdmCommon, dmCommon);
  Application.CreateForm(TdmVytvoreni, dmVytvoreni);
  Application.CreateForm(TdmPrevod, dmPrevod);
  Application.CreateForm(TdmMail, dmMail);
  Application.CreateForm(TdmTisk, dmTisk);
  Application.CreateForm(TDesFastReport, DesFastReport);
  Application.CreateForm(TDesU, DesU);
  Application.Run;
end.
