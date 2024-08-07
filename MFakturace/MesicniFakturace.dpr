program MesicniFakturace;
uses
  Forms,
  AbraEntities in '..\DE$_Common\AbraEntities.pas',
  DesUtils in '..\DE$_Common\DesUtils.pas' {DesU},
  FIfaktura in 'FIfaktura.pas' {dmFaktura: TDataModule},
  FImail in 'FImail.pas' {dmMail: TDataModule},
  FIprevod in 'FIprevod.pas' {dmPrevod: TDataModule},
  FItisk in 'FItisk.pas' {dmTisk: TDataModule},
  FImain in 'FImain.pas' {fmMain},
  DesInvoices in '..\DE$_Common\DesInvoices.pas',
  DesFastReports in '..\DE$_Common\DesFastReports.pas' {DesFastReport},
  AArray in '..\DE$_Common\AArray.pas';

{$R *.res}
begin
  Application.Initialize;
  Application.Title := 'M�s��n� fakturace s �T�';
  Application.CreateForm(TfmMain, fmMain);
  Application.CreateForm(TDesU, DesU);
  Application.CreateForm(TdmFaktura, dmFaktura);
  Application.CreateForm(TdmMail, dmMail);
  Application.CreateForm(TdmPrevod, dmPrevod);
  Application.CreateForm(TdmTisk, dmTisk);
  Application.CreateForm(TDesFastReport, DesFastReport);
  Application.Run;
end.
