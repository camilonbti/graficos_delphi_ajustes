unit uGraficos;

interface

uses SHDocVw, System.Generics.Collections, System.Classes, system.SysUtils, uGraficoGoogle,
  // Added
  ComCtrls, OleCtrls, MSHTML;

type
  TGraficos = class
  private
    FWebBrowser: TWebBrowser;
    FOnBeforeNavigate: TWebBrowserBeforeNavigate2;
    FGraficosGoogle: TObjectList<TGraficoGoogle>;

    procedure SetWebBrowser(const Value: TWebBrowser);
    procedure ExecuteScript(Value: string);

    function GetHTML : string;
    function GetHTMLGraficos : string;
    function Cabecalho : string;

    procedure DoBeforeNavigate(ASender: TObject; const pDisp: IDispatch;
                                                           const URL: OleVariant;
                                                           const Flags: OleVariant;
                                                           const TargetFrameName: OleVariant;
                                                           const PostData: OleVariant;
                                                           const Headers: OleVariant;
                                                           var Cancel: WordBool);
    procedure DefineIEVersion;
    function RetornaFuncoesRandomColor(): string;
    function ConvertString(AValue: String): String;
  public
    procedure AfterConstruction; override;
    procedure BeforeDestruction; override;

    procedure AddGrafico(AGrafico : TGraficoGoogle);
    procedure Atualizar();
    procedure Gerar;
  published
    property WebBrowser : TWebBrowser read FWebBrowser write SetWebBrowser;
    property HTML : string read GetHTML;
  end;

function Space(Tamanho: Integer): string;

implementation

uses
  Winapi.ActiveX, System.StrUtils, Win.Registry, Winapi.Windows;

function Space(Tamanho: Integer): string;
begin
  Result := StringOfChar(' ', Tamanho);
end;

  { TGraficos }

procedure TGraficos.AddGrafico(AGrafico: TGraficoGoogle);
begin
  FGraficosGoogle.Add(AGrafico);
end;

procedure TGraficos.AfterConstruction;
begin
  inherited;

  DefineIEVersion;

  FGraficosGoogle := TObjectList<TGraficoGoogle>.Create(True);
  FOnBeforeNavigate := nil;
end;

// Added
function TGraficos.ConvertString(AValue: String): String;
var
  rbs: RawByteString;
begin
  rbs := UTF8Encode(AValue);
  SetCodePage(rbs, 0, False);
  Result := UnicodeString(rbs);
end;

procedure TGraficos.Gerar;
var
  sl: TStringList;
  ms: TMemoryStream;
begin
  if not Assigned(FWebBrowser) then
    raise Exception.Create('TWebBrowser não definido.');

  if NOT Assigned(FWebBrowser.Document) then
    FWebBrowser.Navigate('about:blank');

  sl := TStringList.Create;
  try
    ms := TMemoryStream.Create;
    try          // Added
      sl.Text := ConvertString(GetHTML); //GetHTML;
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

procedure TGraficos.BeforeDestruction;
begin
  FGraficosGoogle.Free;
end;

function TGraficos.Cabecalho: string;
var
  pacotes : TStringList;
  aux : string;
  Grafico : TGraficoGoogle;
  I: Integer;
begin
  pacotes := TStringList.Create;
  try
    for Grafico in FGraficosGoogle do
    begin
      aux := QuotedStr(Grafico.Package);
      if pacotes.IndexOf(aux) < 0 then
        pacotes.Add(aux);
    end;
    pacotes.Delimiter := ',';
    pacotes.StrictDelimiter := true;
    aux := pacotes.DelimitedText;
  finally
    pacotes.Free;
  end;

  Result  := //
    '<html lang="pt-br"> ' + sLineBreak +
    '    <head> '  + sLineBreak +
    '        <meta http-equiv="X-UA-Compatible" content="IE=Edge,chrome=1" /> ' + sLineBreak +
    '        <meta charset="UTF-8"/> ' + sLineBreak +
    '        <head> ' + sLineBreak +
    '            <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.2.1/jquery.min.js"></script> '+ sLineBreak +
    '            <script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script> '+ sLineBreak +
    '            <script type="text/javascript"> '+ sLineBreak + sLineBreak;

  // Added
  Result := Result + '            let ';
  //Retornar os dados em variável antes da function
  for I := 1 to FGraficosGoogle.Count do
  begin
    Result := Result + FGraficosGoogle[Pred(i)].GetUpdateData(i) + sLineBreak + sLineBreak;
  end;

  Result := Result + //
    '                google.charts.load(''current'', {''packages'':[' + Aux + ']}); '+ sLineBreak +
    '                google.charts.setOnLoadCallback(drawChart); '+ sLineBreak +
    '                function drawChart() { '+ sLineBreak;

  //Informa funcoes para cada grafico carregar
  for I := 1 to FGraficosGoogle.Count do
    Result := Result + '                    drawGrafico' + IntToStr(i) + '();'+ sLineBreak;

  Result := Result +
    '                } ' + sLineBreak;

  // Added
  // Funções de randomColor
  for Grafico in FGraficosGoogle do
  begin
    if Grafico.FRandomColor then
    begin
      Result := Result + sLineBreak + sLineBreak + RetornaFuncoesRandomColor() +sLineBreak + sLineBreak;
      break;
    end;
  end;

    //Cria as funcoes de cada grafico
  for I := 1 to FGraficosGoogle.Count do
  begin
    Result := Result + '                function drawGrafico' + IntToStr(i) + '() {' + sLineBreak +
    '                    '+ FGraficosGoogle[Pred(i)].GetFuncaoDraw('Grafico' + intTostr(i), i) + '}' + sLineBreak;
  end;

  // Added
  Result := Result + sLineBreak;
  for I := 1 to FGraficosGoogle.Count do
  begin
    Result := Result + '                $(window).resize(function(){ ' + sLineBreak +
                       '	                  drawGrafico'+ IntToStr(i) +'(); ' + sLineBreak +
                       '		            }); ' + sLineBreak ;
  end;

  Result := Result +
    '            </script>  ' +
    '        </head>  ';

end;

procedure TGraficos.DefineIEVersion;
const
  REG_KEY = 'Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION';
var
  Reg: TRegistry;
  AppName: String;
begin
  AppName := ExtractFileName(ExtractFileName(ParamStr(0)));
  Reg := nil;
  try
    Reg := TRegistry.Create();
    Reg.RootKey := HKEY_CURRENT_USER;
    if Reg.OpenKey(REG_KEY, True) then
    begin
//      if Versao = 0 then
//        Reg.DeleteValue(AppName)
//      else
        Reg.WriteInteger(AppName, 11000);
      Reg.CloseKey;
    end;
  except
  end;
  if (Assigned(Reg)) then
    FreeAndNil(Reg);
end;

procedure TGraficos.DoBeforeNavigate(ASender: TObject; const pDisp: IDispatch; const URL, Flags, TargetFrameName, PostData, Headers: OleVariant; var Cancel: WordBool);
var
  Target : String;
  Aux, Method: string;
  idx : integer;
  Params : TStringList;
begin
  Target := URL;
  if UpperCase(Target).StartsWith('ACTIONCALLBACKJS') then
  begin
    Method := Copy(Target, Pos(':', Target) + 1, Length(Target));
    Method := Copy(Method, 1, Pos('(', Method) - 1);
    idx := StrToInt(ReplaceStr(Method, 'Retorno', ''));
    Params := TStringList.Create;
    try
      Aux := Copy(Target, Pos('(', Target) + 1, Length(Target));
      Aux := Copy(Aux, 1, LastDelimiter(')', Aux)-1);
      Params.CommaText := Aux;
      if Assigned(FGraficosGoogle[Pred(idx)].OnCallBack) then
        FGraficosGoogle[Pred(idx)].OnCallBack(Params);
      Cancel := True;
    finally
      Params.Free;
    end;
  end;

  if Assigned(FOnBeforeNavigate) then
    FOnBeforeNavigate(ASender, pDisp, URL, Flags, TargetFrameName, PostData, Headers, Cancel);
end;

function TGraficos.GetHTML: string;
begin
  Result := Cabecalho;
  Result := Result + sLineBreak +
    '        <body> ' + sLineBreak ;

  Result := Result + GetHTMLGraficos + sLineBreak;

  Result := Result + sLineBreak +
    '        </body> ' + sLineBreak +
    '        </head> ' + sLineBreak + //Added
    ' </html> ' ;
end;

function TGraficos.GetHTMLGraficos: string;
var
  i : Integer;
begin
  Result := '';

  for i := 1 to FGraficosGoogle.Count do
    Result := Result + sLineBreak + //
    '            <div id="Grafico' + intTostr(i) + '" style="center: 100%; width: 95%"></div>';
end;

// Added
procedure TGraficos.Atualizar();
var
  lData: string;
begin
  lData := '';

  //Retornar os dados em variável
  for var I: Integer := 1 to FGraficosGoogle.Count do
  begin
    lData := lData + FGraficosGoogle[Pred(i)].GetUpdateData(i) + sLineBreak + sLineBreak + //
    'drawGrafico'+ IntToStr(i) +'(); ';
  end;

  ExecuteScript(lData);
end;

procedure TGraficos.ExecuteScript(Value: string);
var
  Doc : IHTMLDocument2;
  HTMLWindow: IHTMLWindow2;
begin
  Doc := FWebBrowser.Document as IHTMLDocument2;
  if Assigned(Doc) then
  begin
    HTMLWindow := Doc.parentWindow;
    if Assigned(HTMLWindow) then
    begin
      try
        HTMLWindow.execScript(Value, 'JavaScript');
      except
        on E: Exception do
          raise Exception.Create('Erro ao Executar Script ');
      end;
    end;
  end;
end;

procedure TGraficos.SetWebBrowser(const Value: TWebBrowser);
begin
  if Assigned(FWebBrowser) then
    FWebBrowser.OnBeforeNavigate2 := FOnBeforeNavigate;

  FWebBrowser                   := Value;
  FWebBrowser.Silent            := True;
  FOnBeforeNavigate             := FWebBrowser.OnBeforeNavigate2;
  FWebBrowser.OnBeforeNavigate2 := DoBeforeNavigate;
end;

function TGraficos.RetornaFuncoesRandomColor(): string;
begin
  Result := //
      '                                                                                   ' + sLineBreak + //
      '	function mt_rand (min, max) {                                                     ' + sLineBreak + //
      '		var argc = arguments.length                                                     ' + sLineBreak + //
      '		if (argc === 0) {                                                               ' + sLineBreak + //
      '			min = 0                                                                       ' + sLineBreak + //
      '			max = 2147483647                                                              ' + sLineBreak + //
      '		} else if (argc === 1) {                                                        ' + sLineBreak + //
      '			throw new Error(''Warning: mt_rand() expects exactly 2 parameters, 1 given'') ' + sLineBreak + //
      '		}                                                                               ' + sLineBreak + //
      '		return Math.floor(Math.random() * (max - min + 1)) + min                        ' + sLineBreak + //
      '	}                                                                                 ' + sLineBreak + //
      '                                                                                   ' + sLineBreak + //
      '   function dechex (d) {                                                           ' + sLineBreak + //
      '	   var hex = Number(d).toString(16)                                               ' + sLineBreak + //
      '	   hex = ''000000''.substr(0, 6 - hex.length) + hex                               ' + sLineBreak + //
      '	   return hex                                                                     ' + sLineBreak + //
      '   }                                                                               ' + sLineBreak + //
      '                                                                                   ' + sLineBreak + //
      '   function str_pad (str) {                                                        ' + sLineBreak + //
      '	   str += ''''                                                                    ' + sLineBreak + //
      '	   while (str.length < 3) {                                                       ' + sLineBreak + //
      '		  str = str + str                                                               ' + sLineBreak + //
      '	   }                                                                              ' + sLineBreak + //
      '	   return str                                                                     ' + sLineBreak + //
      '   }                                                                               ' + sLineBreak + //
      '                                                                                   ' + sLineBreak + //
      '   function random_color_part () {                                                 ' + sLineBreak + //
      '	   // return str_pad( dechex( mt_rand( 0, 255 ) ), 2, ''0'', 1);                  ' + sLineBreak + //
      '	   return mt_rand(0, 255)                                                         ' + sLineBreak + //
      '   }                                                                               ' + sLineBreak + //
      '                                                                                   ' + sLineBreak + //
      '   function random_color () {                                                      ' + sLineBreak + //
      '	   return ''rgb('' + random_color_part() + '','' + random_color_part() + '','' + random_color_part() + '')'' '
      + sLineBreak + //
      '   }                                                                               ' + sLineBreak + //
      '                                                                                   ' + sLineBreak +
      sLineBreak;
end;

end.



