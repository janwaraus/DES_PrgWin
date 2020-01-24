program Faktury_za_telefony;

uses
  Forms,
  MainFT in 'MainFT.pas' {fmMainFT},
  DES_common in '..\DE$_Common\DES_common.pas' {dmDES_common: TDataModule},
  Code in 'Code.pas' {dmCode: TDataModule};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TfmMainFT, fmMainFT);
  Application.CreateForm(TdmDES_common, dmDES_common);
  Application.CreateForm(TdmCode, dmCode);
  Application.Run;
end.
