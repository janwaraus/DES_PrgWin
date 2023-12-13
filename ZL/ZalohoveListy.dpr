program ZalohoveListy;

uses
  Forms,
  ZLmain in 'ZLmain.pas' {fmMain},
  ZLvytvoreni in 'ZLvytvoreni.pas' {dmVytvoreni: TDataModule},
  ZLprevod in 'ZLprevod.pas' {dmPrevod: TDataModule},
  ZLtisk in 'ZLtisk.pas' {dmTisk: TDataModule},
  ZLmail in 'ZLmail.pas' {dmMail: TDataModule},
  AbraEntities in '..\DE$_Common\AbraEntities.pas',
  DesFastReports in '..\DE$_Common\DesFastReports.pas' {DesFastReport},
  DesInvoices in '..\DE$_Common\DesInvoices.pas',
  DesUtils in '..\DE$_Common\DesUtils.pas' {DesU},
  AArray in '..\DE$_Common\AArray.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfmMain, fmMain);
  Application.CreateForm(TdmVytvoreni, dmVytvoreni);
  Application.CreateForm(TdmPrevod, dmPrevod);
  Application.CreateForm(TdmTisk, dmTisk);
  Application.CreateForm(TdmMail, dmMail);
  Application.CreateForm(TDesFastReport, DesFastReport);
  Application.CreateForm(TDesU, DesU);
  Application.Run;
end.
