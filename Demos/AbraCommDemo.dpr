program AbraCommDemo;

uses
  Vcl.Forms,
  AbraWebAPI_demo in 'AbraWebAPI_demo.pas' {Form1},
  DesUtils in '..\DE$_Common\DesUtils.pas' {DesU},
  AbraEntities in '..\DE$_Common\AbraEntities.pas',
  AArray in '..\DE$_Common\AArray.pas',
  DesInvoices in '..\DE$_Common\DesInvoices.pas',
  DesFastReports in '..\DE$_Common\DesFastReports.pas' {DesFastReport};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TDesU, DesU);
  Application.CreateForm(TDesFastReport, DesFastReport);
  Application.Run;
end.
