unit AArray;

interface

uses
  Classes, SysUtils, StrUtils, Variants;

type TAArray = class
private
  procedure PutV (Key, Value: Variant);
  function GetV (Key: Variant): Variant;
  function JsonEscapeString(const AValue: string): string;
public
  Values: array of Variant;
  Keys: array of string[255];
  Position: Integer;
  RowList : TList;

  constructor Create;
  property Items [Index : Variant]: Variant read GetV write PutV; default;
  function IndexByKey (Key: Variant): Integer;
  function Count: Integer;
  procedure Reset;
  function CurrKey: Variant;
  function CurrValue: Variant;
  function Foreach: Boolean;
  procedure Sort;

  function AsJSon(indent: boolean = false): string;
  function printKVPairAsJson(index : integer): string;
  function addRow : TAArray;

end;

implementation

constructor TAArray.Create;
 begin
   SetLength(Keys, 0);
   SetLength(Values, 0);
   self.Position := 0;
   self.RowList := TList.Create;
 end;

function TAArray.Count: Integer;
 begin
   Result := Length (Self.Keys)
 end;

 procedure TAArray.PutV (Key, Value: Variant);
 var Cursor: Integer;
 begin
   Cursor := Self.IndexByKey(Key);
   if Cursor = -1 then
   begin
     SetLength (Self.Keys, Length(Self.Keys)+1);
     SetLength (Self.Values, Length(Self.Values)+1);
     Cursor := Length(Self.Keys)-1;
   end;
   Self.Keys[Cursor] := Key;
   Self.Values[Cursor] := Value;
 end;

 function TAArray.GetV (Key: Variant): Variant;
 var Cursor: Integer;
 begin
   Cursor := Self.IndexByKey (Key);
   if Cursor = -1 then
   begin
     Result := '';
   end
   else
   begin
     Result := Self.Values[Cursor];
   end;
 end;

 function TAArray.IndexByKey (Key: Variant): Integer;
 var Current,Records: Integer;
 begin
   Result := -1;
   Records := Length (Self.Keys);
   for Current := 0 to Records - 1 do
   begin
     if Self.Keys[Current] = Key then
     begin
       Result := Current;
       break;
     end;
   end;
 end;

 function TAArray.CurrKey: Variant;
 begin
   Result := Self.Keys[Self.Position];
 end;

 function TAArray.CurrValue: Variant;
 begin
   Result := Self.Values[Self.Position];
 end;

 procedure TAArray.Reset;
 begin
   Self.Position := 0;
 end;

 function TAArray.Foreach: Boolean;
 begin
   if Self.Position < Self.Count then
     Result := True
   else
     Result := False;
   Self.Position := Self.Position + 1;
 end;

 procedure TAArray.Sort;
 var Current, Records: Integer;
 begin
   Records := Length (Self.Keys);
   for Current := 0 to Records-1 do
   begin

   end;
 end;

function TAArray.JsonEscapeString(const AValue: string): string;

  procedure AddChars(const AChars: string; var Dest: string; var AIndex: Integer); inline;
  begin
    System.Insert(AChars, Dest, AIndex);
    System.Delete(Dest, AIndex + 2, 1);
    Inc(AIndex, 2);
  end;

  procedure AddUnicodeChars(const AChars: string; var Dest: string; var AIndex: Integer); inline;
  begin
    System.Insert(AChars, Dest, AIndex);
    System.Delete(Dest, AIndex + 6, 1);
    Inc(AIndex, 6);
  end;

var
  i, ix: Integer;
  AChar: Char;
begin
  Result := AValue;
  ix := 1;
  for i := 1 to System.Length(AValue) do
  begin
    AChar :=  AValue[i];
    case AChar of
      '/', '\', '"':
      begin
        System.Insert('\', Result, ix);
        Inc(ix, 2);
      end;
      #8:  //backspace \b
      begin
        AddChars('\b', Result, ix);
      end;
      #9:
      begin
        AddChars('\t', Result, ix);
      end;
      #10:
      begin
        AddChars('\n', Result, ix);
      end;
      #12:
      begin
        AddChars('\f', Result, ix);
      end;
      #13:
      begin
        AddChars('\r', Result, ix);
      end;
      #0 .. #7, #11, #14 .. #31:
      begin
        AddUnicodeChars('\u' + IntToHex(Word(AChar), 4), Result, ix);
      end
      else
      begin
        if Word(AChar) > 127 then
        begin
          AddUnicodeChars('\u' + IntToHex(Word(AChar), 4), Result, ix);
        end
        else
        begin
          Inc(ix);
        end;
      end;
    end;
  end;
end;

function TAArray.printKVPairAsJson(index : integer): string;
var
  basicType  : Integer;
begin

  basicType := VarType(self.Values[index]);

  // Set a string to match the type
  case basicType of

  //case VarType(self.Values[index]) of
    varString, varUString, varStrArg:
      Result := ' "' + self.Keys[index] +  '" : "' + JsonEscapeString(VarToStr(self.Values[index])) + '",' + sLineBreak;
  else
      Result := ' "' + self.Keys[index] +  '" : ' + LowerCase(AnsiReplaceStr(VarToStr(self.Values[index]), ',', '.')) + ',' + sLineBreak;
  end;

end;

function TAArray.AsJSon(indent: boolean = false): string;
var
  i, j: integer;
  iBoRowAA : TAArray;
  js : string;
begin

  js := '{ ' + sLineBreak;

  for i := 0 to self.Count - 1 do
    js := js + self.printKVPairAsJson(i);

  if self.RowList.Count > 0 then
  begin
    js := js + ' "rows": [' + sLineBreak + '  {' + sLineBreak;

    for i := 0 to self.RowList.Count - 1 do
    begin
      iBoRowAA := TAArray(self.RowList[i]);
      for j := 0 to iBoRowAA.Count - 1 do
        js := js + '  ' + iBoRowAA.printKVPairAsJson(j);
      if i < self.RowList.Count - 1 then
        js := js + '  },{' + sLineBreak;
    end;
    js := js + '  }]' + sLineBreak
  end;
  Result := js + '}';
end;

function TAArray.addRow(): TAArray;
var
  newAA : TAArray;
begin
  newAA := TAArray.Create;
  self.RowList.Add(newAA); //pøidáme nové AArray do listu
  Result := newAA; //vrátíme ukazatel, aby se s novým AArray dalo hned pracovat
end;


end.
