unit uComunicacaoJS;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Classes,
  Vcl.OleCtrls, SHDocVw, MSHTML, System.StrUtils;

type
  TComunicacaoJS = class
  private
    const
      CALLBACK_PREFIX = 'ActionCallBackJS:';
  public
    function IsCallbackURL(const URL: string): Boolean;
    function ExtrairParametrosCallback(const URL: string): TStringList;
    procedure ExecutarScript(WebBrowser: TWebBrowser; const Script: string);
  end;

implementation

function TComunicacaoJS.IsCallbackURL(const URL: string): Boolean;
begin
  Result := StartsText(CALLBACK_PREFIX, URL);
end;

function TComunicacaoJS.ExtrairParametrosCallback(const URL: string): TStringList;
var
  Method, Params, Aux: string;
  OpenParenIndex, CloseParenIndex: Integer;
begin
  Result := TStringList.Create;

  if not IsCallbackURL(URL) then
    Exit;

  // Remove o prefixo para obter o método e parâmetros
  Method := Copy(URL, Length(CALLBACK_PREFIX) + 1, Length(URL));

  // Encontra os índices dos parênteses
  OpenParenIndex := Pos('(', Method);

  if OpenParenIndex > 0 then
  begin
    // Obtém o nome do método
    Aux := Copy(Method, 1, OpenParenIndex - 1);
    Result.Add(Aux);

    // Encontra o fechamento do parêntese, considerando possíveis JSON aninhados
    CloseParenIndex := LastDelimiter(')', Method);

    if CloseParenIndex > OpenParenIndex then
    begin
      // Obtém a string de parâmetros
      Params := Copy(Method, OpenParenIndex + 1, CloseParenIndex - OpenParenIndex - 1);

      // Processa os parâmetros - aqui temos que lidar com JSON
      // Para simplificar, vamos apenas dividir por vírgula, mas em um caso real
      // precisaria de um parser JSON adequado
      Params := StringReplace(Params, '[', '', [rfReplaceAll]);
      Params := StringReplace(Params, ']', '', [rfReplaceAll]);
      Params := StringReplace(Params, '"', '', [rfReplaceAll]);

      // Adiciona os parâmetros à lista
      Result.CommaText := Params;
    end;
  end;
end;

procedure TComunicacaoJS.ExecutarScript(WebBrowser: TWebBrowser; const Script: string);
var
  Doc: IHTMLDocument2;
  Window: IHTMLWindow2;
begin
  if not Assigned(WebBrowser) then
    Exit;

  if not Assigned(WebBrowser.Document) then
    Exit;

  try
    Doc := WebBrowser.Document as IHTMLDocument2;

    if Assigned(Doc) then
    begin
      Window := Doc.parentWindow;

      if Assigned(Window) then
        Window.execScript(Script, 'JavaScript');
    end;
  except
    on E: Exception do
      raise Exception.Create('Erro ao executar script: ' + E.Message);
  end;
end;

end.
