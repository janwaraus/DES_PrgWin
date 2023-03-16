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
// funkce pro pr�ci s OpenOffice (LibreOffice)
    function IsNullEmpty(aVariant: variant): boolean;          // test pr�zdn� prom�nn�
    function SetParams (Name: string; Data: variant): variant; // zad�n� parametr� p�i otev�r�n� souboru
    function OpenOO: variant;                                  // spu�t�n� OpenOffice
    function OpenDesktop: variant;                             // otev�en� desktopu
    function OpenDoc(aFile: string): variant;                  // otev�en� souboru
    function LastRow(aSheet: variant): integer;                // posledn� pou�it� ��dek v aktivn�m listu spreadsheetu

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
// zad�n� parametr� p�i otev�r�n� souboru
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
// P�ipojen� k OpenOffice
begin
  Result := null;
  try
// zkus� se, jestli neb��
    Result := GetActiveOleObject('com.sun.star.ServiceManager');
  except try
// jinak se spust�
    Result := CreateOleObject('com.sun.star.ServiceManager');
    except
      on E: Exception do ShowMessage('Nelze spustit OpenOffice' + ^M + E.Message);
    end;  // try
  end;  // try
end;

// ------------------------------------------------------------------------------------------------

function TdmDES_OO_Common.OpenDesktop: variant;
// Otev�en� desktopu
begin
  Result := null;
  try
    Result := ServiceManager.CreateInstance('com.sun.star.frame.Desktop');
  except
    on E: Exception do ShowMessage('Nelze otev��t desktop' + ^M + E.Message);
  end;  // try
end;

// ------------------------------------------------------------------------------------------------

function TdmDES_OO_Common.OpenDoc(aFile: string): variant;
// otev�e soubor
begin
  Result := null;
  try
    Result := Desktop.LoadComponentFromURL('file:///' + aFile, '_default', 0, LoadParams);
  except
    on E: Exception do ShowMessage(Format('Nepoda�ilo se otev��t %s. %s%s', [aFile, ^M, E.Message]));
  end;
end;

// ------------------------------------------------------------------------------------------------

function TdmDES_OO_Common.LastRow(aSheet: variant): integer;
// posledn� ��dek v listu
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
