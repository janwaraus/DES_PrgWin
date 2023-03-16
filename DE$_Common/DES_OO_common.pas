unit DES_OO_common;

interface

uses
  SysUtils, Classes, Dialogs, Variants, ComObj;

type
  TdmDES_OO_Common = class(TDataModule)
  private
    { Private declarations }
  public
    ServiceManager,
    Desktop,
    LoadParams: variant;
// funkce pro práci s OpenOffice (LibreOffice)
    function IsNullEmpty(aVariant: variant): boolean;          // test prázdné promìnné
    function SetParams (Name: string; Data: variant): variant; // zadání parametrù pøi otevírání souboru
    function OpenOO: variant;                                  // spuštìní OpenOffice
    function OpenDesktop: variant;                             // otevøení desktopu
    function OpenDoc(aFile: string): variant;                  // otevøení souboru
    function LastRow(aSheet: variant): integer;                // poslední použitý øádek v aktivním listu spreadsheetu

  end;

var
  dmDES_OO_Common: TdmDES_OO_Common;


implementation

{%CLASSGROUP 'Vcl.Controls.TControl'}

{$R *.dfm}

// ------------------------------------------------------------------------------------------------

function TdmDES_OO_Common.IsNullEmpty(aVariant: variant): boolean;
begin
  Result:= VarIsEmpty(aVariant) or VarIsNull(aVariant) or VarIsClear(aVariant);
end;

// ------------------------------------------------------------------------------------------------

function TdmDES_OO_Common.SetParams (Name: string; Data: variant): variant;
// zadání parametrù pøi otevírání souboru
var
  Reflection: variant;
begin
  Reflection := ServiceManager.CreateInstance('com.sun.star.reflection.CoreReflection');
  Reflection.forName('com.sun.star.beans.PropertyValue').createObject(result);
  Result.Name := Name;
  Result.Value := Data;
end;

// ------------------------------------------------------------------------------------------------

function TdmDES_OO_Common.OpenOO: variant;
// Pøipojení k OpenOffice
begin
  Result := null;
  try
// zkusí se, jestli nebìží
    Result := GetActiveOleObject('com.sun.star.ServiceManager');
  except try
// jinak se spustí
    Result := CreateOleObject('com.sun.star.ServiceManager');
    except
      on E: Exception do ShowMessage('Nelze spustit OpenOffice' + ^M + E.Message);
    end;  // try
  end;  // try
end;

// ------------------------------------------------------------------------------------------------

function TdmDES_OO_Common.OpenDesktop: variant;
// Otevøení desktopu
begin
  Result := null;
  try
    Result := ServiceManager.CreateInstance('com.sun.star.frame.Desktop');
  except
    on E: Exception do ShowMessage('Nelze otevøít desktop' + ^M + E.Message);
  end;  // try
end;

// ------------------------------------------------------------------------------------------------

function TdmDES_OO_Common.OpenDoc(aFile: string): variant;
// otevøe soubor
begin
  Result := null;
  try
    Result := Desktop.LoadComponentFromURL('file:///' + aFile, '_default', 0, LoadParams);
  except
    on E: Exception do ShowMessage(Format('Nepodaøilo se otevøít %s. %s%s', [aFile, ^M, E.Message]));
  end;
end;

// ------------------------------------------------------------------------------------------------

function TdmDES_OO_Common.LastRow(aSheet: variant): integer;
// poslední øádek v listu
var
  vCursor: variant;
begin
  Result := 0;
  if not IsNullEmpty(aSheet) then begin
    vCursor := aSheet.CreateCursor;
    vCursor.GotoEndOfUsedArea(False);
    Result := vCursor.RangeAddress.EndRow;
    vCursor := Unassigned;
  end;
end;

// ------------------------------------------------------------------------------------------------

end.
