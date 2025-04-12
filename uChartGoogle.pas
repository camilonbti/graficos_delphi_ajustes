unit uChartGoogle;

interface

uses
  SysUtils,
  StrUtils,
  Classes,
  Registry,
  Data.DB,
  SHDocVw,
  Winapi.ActiveX,
  Winapi.Windows;

type
  TChartTipo      = (ctPizza, ctBarra);
  TTipoOrientacao = (toVertical, toHorizontal);
  TArrayString    = array of string;
  TArrayInteger   = array of integer;

  TChartGoogle = class
  private
    FHtml, FDiv, FDataTable                            : TStringList;
    FNome, FTitulo, FCallBack                          : string;
    FTituloVertical, FTituloHorizontal                 : string;
    FBarraTipoVH, FLabels, FFieldName                  : string;
    FClassGrafico, FFields, FColumns                   : TArrayString;
    FIndexField                                        : TArrayInteger;
    FHeader, FOpcoes, FLineColumns, FFieldValues       : string;
    FSeparador, FFixColmun                             : string;
    FTipoGrafico                                       : TChartTipo;
    FAltura, FLargura, FColumnVirtuais, FFixColumnField: Int64;
    FDataSet                                           : TDataSet;
    FWebBrowser                                        : TWebBrowser;
    procedure DefineIEVersion();
    function Padrao(): string;
    function Labels(): String;
    function ConvertString(AValue: String): String;
    procedure HtmlBrowserGenerated(const HTMLCode: string);
    function GetColumns(): String;
    function GetFieldValues(): String;
  public
    constructor Create(Nome: string; Altura, Largura: Int64);
    destructor Destroy; override;

    property Html: TStringList read FHtml write FHtml;

    function Clear(): TChartGoogle;
    function Browser(AValue: TWebBrowser): TChartGoogle;
    function Grafico(TipoGrafico: TChartTipo; BarraTipo: TTipoOrientacao = toVertical): TChartGoogle;
    function Titulo(ATitulo: string; Vertical: string = ''; Horizontal: string = ''): TChartGoogle;
    function Header(ClassGrafico: TArrayString): TChartGoogle;
    function ArrayToDataTable(): TChartGoogle;
    function DataTable(): TChartGoogle;
    function Opcoes(Opcoes: string): TChartGoogle;
    function &Div(Altura, Largura: Int64): TChartGoogle;
    function Callback(ACallBack: string; IndexField: TArrayInteger; Separador: string = ' - '): TChartGoogle;
    function DataSet(ADataSet: TDataSet; AFieldName: string): TChartGoogle;
    function Columns(AColumns: TArrayString; AColumnVirtuais: integer = 0): TChartGoogle;
    function Fields(AFields: TArrayString; AFixColmun: string = ''; IndexField: integer = 0): TChartGoogle;
    procedure Generated();
    procedure NavigateIn(lHtml: string);
  end;

implementation

{ TChartGoogle }

function TChartGoogle.&Div(Altura, Largura: Int64): TChartGoogle;
begin
  Result := Self;
  FDiv.Add( '<div id= "' + FNome + '" style="height: 100%; overflow: auto;"></div>  ');

//
//  FDiv.Add(' <div id="' + FNome + '" style="' + //
//      'width: ' + Largura.ToString + //
//      '; height:' + Altura.ToString + //
//      '"></div> ');
end;

function TChartGoogle.Browser(AValue: TWebBrowser): TChartGoogle;
begin
  Result      := Self;
  FWebBrowser := AValue;
end;

function TChartGoogle.Callback(ACallBack: string; IndexField: TArrayInteger; Separador: string = ' - ')
  : TChartGoogle;
begin
  Result := Self;
  if not Trim(ACallBack).IsEmpty and (Length(IndexField) > 0) then
  begin
    FCallBack   := ACallBack;
    FIndexField := IndexField;
    FSeparador  := Separador;
  end;
end;

constructor TChartGoogle.Create(Nome: string; Altura, Largura: Int64);
begin
  inherited Create;
  FNome         := LowerCase(StringReplace(Nome, ' ', '_', []));
  FClassGrafico := ['corechart'];
  FAltura       := Altura;
  FLargura      := Largura;
  FHtml         := TStringList.Create;
  FDiv          := TStringList.Create;
  FDataTable    := TStringList.Create;

  if FAltura = 0 then
    FAltura := 400;
  if FLargura = 0 then
    FLargura := 300;

  DefineIEVersion;
end;

destructor TChartGoogle.Destroy;
begin
  if Assigned(FHtml) then
    FreeAndNil(FHtml);
  if Assigned(FDiv) then
    FreeAndNil(FDiv);
  if Assigned(FDataTable) then
    FreeAndNil(FDataTable);

  inherited;
end;

function TChartGoogle.DataSet(ADataSet: TDataSet; AFieldName: string): TChartGoogle;
begin
  Result     := Self;
  FDataSet   := ADataSet;
  FFieldName := AFieldName;
end;

function TChartGoogle.Columns(AColumns: TArrayString; AColumnVirtuais: integer = 0): TChartGoogle;
var
  Idx: integer;
const
  ColVirtual = ' {type: ''string'', role: ''xxx''}';
begin
  Result   := Self;
  FColumns := [];
  if Length(AColumns) > 0 then
  begin
    FColumns        := AColumns;
    FColumnVirtuais := AColumnVirtuais;

    if FColumnVirtuais > 0 then
    begin
      while AColumnVirtuais <> 0 do
      begin
        Idx := Length(FColumns);
        SetLength(FColumns, Idx + 1);
        FColumns[Idx] := ColVirtual;
        Dec(AColumnVirtuais);
      end;
    end;
  end;
end;

function TChartGoogle.Fields(AFields: TArrayString; AFixColmun: string = ''; IndexField: integer = 0)
  : TChartGoogle;
begin
  Result  := Self;
  FFields := [];
  if Length(AFields) > 0 then
    FFields := AFields;

  if not Trim(AFixColmun).IsEmpty then
  begin
    if not(IndexField > 0) then
      raise Exception.Create('IndexField da função Fields() deve ser maior que 0');

    FFixColmun      := AFixColmun;
    FFixColumnField := IndexField;
  end;
end;

function TChartGoogle.GetColumns(): String;
var
  J  : integer;
  Aux: String;
begin
  FLineColumns := '[';
  Aux          := ''', ''';
  FLineColumns := FLineColumns + '''';

  for J := Low(FColumns) to High(FColumns) do
  begin
    if J = High(FColumns) then
      Aux := '''';

    if FColumns[J].Contains('{') then
    begin
      FLineColumns := Copy(FLineColumns, 0, Length(FLineColumns) - 2);
      FLineColumns := FLineColumns + FColumns[J] + IfThen(J = High(FColumns), '', ', ''');
    end
    else
      FLineColumns := FLineColumns + FColumns[J] + Aux;
  end;
  FLineColumns := FLineColumns + ']';

  Result := FLineColumns;
end;

function TChartGoogle.GetFieldValues(): String;
var
  Idx, J: integer;
  Aux   : String;
begin
  FFieldValues := '[';
  Aux          := ''', ''';
  FFieldValues := FFieldValues + '''';

  FDataSet.DisableControls;
  FDataSet.First;
  for Idx := 0 to Pred(FDataSet.RecordCount) do
  begin
    for J := Low(FFields) to High(FFields) do
    begin
      if J = High(FFields) then
        Aux := '''';

      if (J = FFixColumnField) and (FFixColumnField > 0) then
        FFieldValues := FFieldValues + FFixColmun + Aux
      else
        case FDataSet.FieldByName(FFields[J]).DataType of
          ftSingle, ftFloat, ftCurrency, ftBCD, ftFMTBCD, ftExtended:
            begin
              FFieldValues := Copy(FFieldValues, 0, Length(FFieldValues) - 2);
              FFieldValues := FFieldValues + //
                StringReplace(FDataSet.FieldByName(FFields[J]).AsString, ',', '.', [rfReplaceAll]) + //
                IfThen(J = High(FFields), '', ', ''');
            end;
        else
          FFieldValues := FFieldValues + FDataSet.FieldByName(FFields[J]).AsString + Aux;
        end;
    end;
    FFieldValues := FFieldValues + '],[''';
    Aux          := ''', ''';

    FDataSet.Next;
  end;
  FFieldValues := FFieldValues + ']';
  FDataSet.EnableControls;

  FFieldValues := Copy(FFieldValues, 0, Length(FFieldValues) - 4);
  Result       := FFieldValues;
end;

function TChartGoogle.Labels: String;
var
  Idx: integer;
  Aux: String;
begin
  FLabels := '[';
  Aux     := '", "';
  FLabels := FLabels + '"';

  FDataSet.DisableControls;
  FDataSet.First;
  for Idx := 0 to Pred(FDataSet.RecordCount) do
  begin
    if Idx = Pred(FDataSet.RecordCount) then
      Aux := '"';

    // Field que será a descrição a ser exibida no grafico
    FLabels := FLabels + FDataSet.FieldByName(FFieldName).AsString + Aux;

    FDataSet.Next;
  end;
  FDataSet.EnableControls;
  FLabels := FLabels + ']';

  Result := FLabels;
end;

function TChartGoogle.ConvertString(AValue: String): String;
var
  rbs: RawByteString;
begin
  rbs := UTF8Encode(AValue);
  SetCodePage(rbs, 0, False);
  Result := UnicodeString(rbs);
end;

procedure TChartGoogle.Generated();
var
  lHtml: string;
begin
  if not Assigned(FWebBrowser) then
    raise Exception.Create('TWebBrowser não definido.');

  if Assigned(FDataSet) and not Trim(FFieldName).IsEmpty then
    Labels;

  lHtml              := Padrao();
  FWebBrowser.Silent := True;
  HtmlBrowserGenerated(ConvertString(lHtml));
end;

procedure TChartGoogle.NavigateIn(lHtml: string);
begin
  if not Assigned(FWebBrowser) then
    raise Exception.Create('TWebBrowser não definido.');

  FWebBrowser.Silent := True;
  HtmlBrowserGenerated(ConvertString(lHtml));
end;

procedure TChartGoogle.HtmlBrowserGenerated(const HTMLCode: string);
var
  sl: TStringList;
  ms: TMemoryStream;
begin
  if NOT Assigned(FWebBrowser.Document) then
    FWebBrowser.Navigate('about:blank');

  sl := TStringList.Create;
  try
    ms := TMemoryStream.Create;
    try
      sl.Text := HTMLCode;
      sl.SaveToStream(ms);
      ms.Seek(0, 0);
      (FWebBrowser.Document as IPersistStreamInit).Load(TStreamAdapter.Create(ms));
    finally
      ms.Free;
    end;
  finally
    sl.Free;
  end;
end;

procedure TChartGoogle.DefineIEVersion();
const
  REG_KEY = 'Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION';
var
  Reg    : TRegistry;
  AppName: String;
begin
  AppName := ExtractFileName(ExtractFileName(ParamStr(0)));
  Reg     := nil;
  try
    Reg         := TRegistry.Create();
    Reg.RootKey := HKEY_CURRENT_USER;
    if Reg.OpenKey(REG_KEY, True) then
    begin
      Reg.WriteInteger(AppName, 11000);
      Reg.CloseKey;
    end;
  except
  end;
  if (Assigned(Reg)) then
    FreeAndNil(Reg);
end;

function TChartGoogle.Grafico(TipoGrafico: TChartTipo; BarraTipo: TTipoOrientacao = toVertical): TChartGoogle;
begin
  Result       := Self;
  FTipoGrafico := TipoGrafico;

  case BarraTipo of
    toVertical:
      FBarraTipoVH := 'vertical';
    toHorizontal:
      FBarraTipoVH := 'horizontal';
  end;
end;

function TChartGoogle.Titulo(ATitulo: string; Vertical: string = ''; Horizontal: string = ''): TChartGoogle;
begin
  Result            := Self;
  FTitulo           := ConvertString(ATitulo);
  FTituloVertical   := Vertical;
  FTituloHorizontal := Horizontal;
end;

function TChartGoogle.Clear(): TChartGoogle;
begin
  Result := Self;
  FHtml.Clear;
  FDiv.Clear;
  FDataTable.Clear;
  FLineColumns := '';
  FFieldValues := '';
end;

function TChartGoogle.Header(ClassGrafico: TArrayString): TChartGoogle;
var
  AClass: string;

  function GetClassFormatArray(): String;
  var
    Idx       : integer;
    Aux, Value: String;
  begin
    Aux   := ''', ''';
    Value := Value + '''';

    for Idx := Low(ClassGrafico) to High(ClassGrafico) do
    begin
      if Idx = High(ClassGrafico) then
        Aux := '''';

      Value := Value + ClassGrafico[Idx] + Aux;
    end;
    Result := Value;
  end;

begin
  FClassGrafico := ClassGrafico;
  AClass        := GetClassFormatArray();
  Result        := Self;
  FHeader       := //
    '<html lang="pt-br"> ' + sLineBreak +
    '    <head> ' + sLineBreak +
    '        <meta http-equiv="X-UA-Compatible" content="IE=Edge,chrome=1" /> ' + sLineBreak +
    '        <meta charset="UTF-8"/> ' + sLineBreak +
    '	       <head> ' + sLineBreak +
    '	           <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.2.1/jquery.min.js"></script> ' + sLineBreak +
    '	           <script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script> ' + sLineBreak +
    '	           <script type="text/javascript"> ' + sLineBreak +
    '	               google.charts.load(''current'', {''packages'':[' + AClass + ']}); ' + sLineBreak +
    '	               google.charts.setOnLoadCallback(drawChart); ' + sLineBreak +
    '	               function drawChart() { ' + sLineBreak ;
end;

function TChartGoogle.Opcoes(Opcoes: string): TChartGoogle;
begin
  Result  := Self;
  FOpcoes := Opcoes;
end;

function TChartGoogle.ArrayToDataTable(): TChartGoogle;
begin
  Result := Self;

  if not Assigned(FDataSet) then
    raise Exception.Create('É necessário informar o dataset na classe.');

  GetColumns();
  GetFieldValues();

  FDataTable.Text := //
    '         var data = google.visualization.arrayToDataTable([                     ' + sLineBreak + //
  // FHtml.Add(FLabels);
    '           ' + FLineColumns + ',                                                ' + sLineBreak + //
    '           ' + FFieldValues + ',                                                ' + sLineBreak + //
  // '           [''Year'', ''Sales'', ''Expenses'', ''Profit''],                     ' + sLineBreak +  //
  // '           [''2014'', 1000, 400, 200],                                          ' + sLineBreak + //
  // '           [''2015'', 1170, 460, 250],                                          ' + sLineBreak + //
  // '           [''2016'', 660, 1120, 300],                                          ' + sLineBreak +  //
  // '           [''2017'', 1030, 540, 350]                                           ' + sLineBreak + //
    '         ]);                                                                    ' + sLineBreak + //
    '	                                                                               ' + sLineBreak;
end;

function TChartGoogle.DataTable(): TChartGoogle;
var
  Idx: integer;
  Aux: string;
begin
  Result := Self;

  if not Assigned(FDataSet) then
    raise Exception.Create('É necessário informar o dataset na classe.');

  FDataTable.Text := //
    '      var data = new google.visualization.DataTable();   ' + sLineBreak;

  FDataSet.DisableControls;
  FDataSet.First;
  for Idx := 0 to Pred(FDataSet.RecordCount) do
  begin
    if Idx = Pred(FDataSet.RecordCount) then
      Aux := '"';

    if Idx = 0 then
    begin

    end;

    // FDataTable.Add(
    FFieldValues := FFieldValues + FDataSet.FieldByName(FFieldName).AsString + Aux; // FFields[J]

    FDataSet.Next;
  end;
  FFieldValues := FFieldValues + ']';
  FDataSet.EnableControls;

  // '      data.addColumn(''string'', ''Pizza'');             ' + sLineBreak + //
  // '       data.addRows([                                    ' + sLineBreak + //
  // '        [''Pepperoni'', 33],                             ' + sLineBreak + //
  // '        [''Hawaiian'', 26],                              ' + sLineBreak + //
  // '        [''Mushroom'', 22],                              ' + sLineBreak + //
  // '        [''Sausage'', 10], // Below limit.               ' + sLineBreak + //
  // '        [''Anchovies'', 9] // Below limit.               ' + sLineBreak + //
  // '      ]);                                                ' + sLineBreak + //
  // '                                                         ' + sLineBreak;
end;

function TChartGoogle.Padrao(): string;
var
  S           : integer;
  IdxCall, Aux: string;
const
  CallLine = 'data.getValue(selectedItem.row, %s) %s';
begin
  IdxCall := '';
  Header(FClassGrafico);

  if Trim(FDiv.Text).IsEmpty then
    &Div(FAltura, FLargura);

  FHtml.Text := FHeader + FDataTable.Text +
  //
    IfThen(Trim(FOpcoes).IsEmpty, //
      '	        var options = {''title'':''' + FTitulo + ''',                                 ' + sLineBreak +
      '	                       ''width'':' + FLargura.ToString + ',                           ' + sLineBreak +
      '	                       ''height'':' + FAltura.ToString + ',                           ' + sLineBreak +
      IfThen(FTipoGrafico = ctBarra, '''bars: ''' + FBarraTipoVH + '''', '') + sLineBreak +
      '                       };                                                              ' + sLineBreak, //
    FOpcoes);
  //

  if not Trim(FCallBack).IsEmpty then
  begin
    FHtml.Add(' function selectHandler() {                                                  ' + sLineBreak +
        '	          var selectedItem = chart.getSelection()[0];                             ' + sLineBreak +
        '	          if (selectedItem) {                                                     ' + sLineBreak);
    //

    IdxCall := '	            var selected = ';
    for S   := Low(FIndexField) to High(FIndexField) do
    begin
      Aux := '+' + QuotedStr(FSeparador) + '+';
      if S = High(FIndexField) then
        Aux := '';

      IdxCall := IdxCall + Format(CallLine, [IntToStr(FIndexField[S]), Aux]);
    end;
    //
    FHtml.Add(IdxCall + sLineBreak +
        '	            alert(''The user selected '' + selected);                             ' + sLineBreak +
        '	          }                                                                       ' + sLineBreak +
        '	        }                                                                         ' + sLineBreak +
        '	                                                                                  ' + sLineBreak +
        '	        google.visualization.events.addListener(chart, ''select'', selectHandler);' + sLineBreak +
        '	                                                                                  ' + sLineBreak);
  end;
  //
  FHtml.Add('  Grafico1.draw(dados1, options); ' + sLineBreak +
      '            $(window).resize(function(){ ' + sLineBreak +
      '            drawGrafico1() ' + sLineBreak +
      '            }); ' + sLineBreak +
      '	       </script> ' + sLineBreak +
      '	   </head> ' + sLineBreak +
      '	   <body>  ' + sLineBreak);
  //

  FHtml.Add(FDiv.Text);
  //

  FHtml.Add('    </body>                                                                       ' + sLineBreak +
      '	</html>                                                                             ');
  //

  Result := FHtml.Text;
end;

end.
