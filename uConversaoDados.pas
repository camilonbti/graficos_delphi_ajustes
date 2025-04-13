unit uConversaoDados;

interface

uses
  System.SysUtils, System.Classes, Data.DB, System.JSON;

type
  TConversaoDados = class
  public
    class function ConvertStringToUTF8(const AValue: string): string;
    class function DataSetToJSON(const ADataSet: TDataSet;
                                const ACampoRotulo, ACampoValor: string): string;
    class function FieldValueToString(const Field: TField): string;
  end;

implementation

class function TConversaoDados.ConvertStringToUTF8(const AValue: string): string;
var
  RBS: RawByteString;
begin
  RBS := UTF8Encode(AValue);
  SetCodePage(RBS, 0, False);
  Result := UnicodeString(RBS);
end;

class function TConversaoDados.DataSetToJSON(const ADataSet: TDataSet;
                                          const ACampoRotulo, ACampoValor: string): string;
var
  JObject: TJSONObject;
  JArray: TJSONArray;
  Bookmark: TBookmark;
begin
  if not Assigned(ADataSet) or ADataSet.IsEmpty then
    Exit('[]');

  JArray := TJSONArray.Create;
  try
    Bookmark := ADataSet.Bookmark;
    ADataSet.DisableControls;
    try
      ADataSet.First;
      while not ADataSet.Eof do
      begin
        JObject := TJSONObject.Create;
        JObject.AddPair(ACampoRotulo, ADataSet.FieldByName(ACampoRotulo).AsString);
        JObject.AddPair(ACampoValor, TJSONNumber.Create(ADataSet.FieldByName(ACampoValor).AsFloat));
        JArray.AddElement(JObject);
        ADataSet.Next;
      end;

      Result := JArray.ToString;
    finally
      ADataSet.Bookmark := Bookmark;
      ADataSet.EnableControls;
    end;
  finally
    JArray.Free;
  end;
end;

class function TConversaoDados.FieldValueToString(const Field: TField): string;
begin
  if Field = nil then
    Exit('null');

  case Field.DataType of
    ftString, ftWideString, ftMemo, ftWideMemo, ftFixedChar, ftFixedWideChar:
      Result := QuotedStr(Field.AsString);
    ftInteger, ftSmallint, ftWord, ftLargeint, ftLongWord, ftShortint, ftByte:
      Result := Field.AsString;
    ftFloat, ftCurrency, ftBCD, ftFMTBcd, ftSingle:
      Result := FloatToStr(Field.AsFloat);
    ftDate, ftTime, ftDateTime, ftTimeStamp:
      Result := QuotedStr(FormatDateTime('yyyy-mm-dd', Field.AsDateTime));
    ftBoolean:
      Result := LowerCase(BoolToStr(Field.AsBoolean, True));
    else
      Result := QuotedStr(Field.AsString);
  end;
end;

end.
