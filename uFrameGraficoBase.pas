unit uFrameGraficoBase;

interface

uses
  uComunicacaoJS,
  uConversaoDados,
  uConfigBrowser,
  uFuncoes,

  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Vcl.OleCtrls, SHDocVw, Data.DB, MSHTML, ActiveX, Registry;

type
  TCallbackGraficoEvent = procedure(Sender: TObject; const Params: TStringList) of object;

  TFrameGraficoBase = class(TFrame)
    WebBrowser: TWebBrowser;
    procedure FrameResize(Sender: TObject);
    procedure WebBrowserBeforeNavigate2(ASender: TObject; const pDisp: IDispatch;
                              const URL, Flags, TargetFrameName,
                              PostData, Headers: OleVariant;
                              var Cancel: WordBool);
    procedure WebBrowserDocumentComplete(ASender: TObject; const pDisp: IDispatch;
                              const URL: OleVariant);
  private
    FComunicacaoJS: TComunicacaoJS;
    FDataSet: TDataSet;
    FOnCallback: TCallbackGraficoEvent;
    FOldBeforeNavigate: TWebBrowserBeforeNavigate2;
    FHTMLCarregado: Boolean;

    procedure DoBeforeNavigate(ASender: TObject; const pDisp: IDispatch;
                              const URL, Flags, TargetFrameName,
                              PostData, Headers: OleVariant;
                              var Cancel: WordBool);
    procedure ConfigurarBrowser;
    procedure DiagnosticarProblemas;
  protected
    procedure CarregarHTML(const HTML: string);
    procedure ExecutarScript(const Script: string);
    function GerarHTMLBase: string; virtual; abstract;
    procedure ProcessarCallback(const Params: TStringList); virtual;
    procedure Loaded; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    procedure AtualizarGrafico; virtual; abstract;
    procedure CarregarDados(ADataSet: TDataSet); virtual;
    procedure LimparGrafico; virtual;
    function VerificarBrowser: Boolean;

    property DataSet: TDataSet read FDataSet write FDataSet;
    property OnCallback: TCallbackGraficoEvent read FOnCallback write FOnCallback;
  end;

implementation

{$R *.dfm}

constructor TFrameGraficoBase.Create(AOwner: TComponent);
begin
  inherited;

  {$IFDEF DEBUG}
  if TForm(AOwner).Owner is TForm then
    TextoMemo_Form('TFrameGraficoBase.Create - Início');
  {$ENDIF}

  FComunicacaoJS := TComunicacaoJS.Create;
  FHTMLCarregado := False;

  // Configuração do navegador será feita no Loaded
end;

destructor TFrameGraficoBase.Destroy;
begin
  {$IFDEF DEBUG}
  if Owner is TForm then
    TextoMemo_Form('TFrameGraficoBase.Destroy - Início');
  {$ENDIF}

  if Assigned(WebBrowser) then
    WebBrowser.OnBeforeNavigate2 := FOldBeforeNavigate;

  FComunicacaoJS.Free;
  inherited;
end;

procedure TFrameGraficoBase.Loaded;
begin
  inherited;

  {$IFDEF DEBUG}
  if Owner is TForm then
    TextoMemo_Form('TFrameGraficoBase.Loaded - Início');
  {$ENDIF}

  ConfigurarBrowser;

  // Inicializa o WebBrowser
  if Assigned(WebBrowser) then
  begin
    WebBrowser.Navigate('about:blank');

    {$IFDEF DEBUG}
    if Owner is TForm then
      TextoMemo_Form('TFrameGraficoBase.Loaded - WebBrowser inicializado');
    {$ENDIF}
  end
  else
  begin
    {$IFDEF DEBUG}
    if Owner is TForm then
      TextoMemo_Form('TFrameGraficoBase.Loaded - WebBrowser não disponível');
    {$ENDIF}
  end;
end;

procedure TFrameGraficoBase.ConfigurarBrowser;
begin
  {$IFDEF DEBUG}
  if Owner is TForm then
    TextoMemo_Form('TFrameGraficoBase.ConfigurarBrowser - Início');
  {$ENDIF}

  TConfigBrowser.ConfigurarIEVersion;

  if Assigned(WebBrowser) then
  begin
    WebBrowser.Silent := True;
    FOldBeforeNavigate := WebBrowser.OnBeforeNavigate2;
    WebBrowser.OnBeforeNavigate2 := DoBeforeNavigate;

    {$IFDEF DEBUG}
    if Owner is TForm then
      TextoMemo_Form('TFrameGraficoBase.ConfigurarBrowser - WebBrowser configurado');
    {$ENDIF}
  end
  else
  begin
    {$IFDEF DEBUG}
    if Owner is TForm then
      TextoMemo_Form('TFrameGraficoBase.ConfigurarBrowser - Erro: WebBrowser não está atribuído');
    {$ENDIF}
  end;
end;

function TFrameGraficoBase.VerificarBrowser: Boolean;
begin
  Result := Assigned(WebBrowser) and Assigned(WebBrowser.Document);

  {$IFDEF DEBUG}
  if Owner is TForm then
  begin
    if Result then
      TextoMemo_Form('VerificarBrowser - WebBrowser OK')
    else
      TextoMemo_Form('VerificarBrowser - WebBrowser NÃO está pronto');
  end;
  {$ENDIF}
end;

procedure TFrameGraficoBase.DoBeforeNavigate(ASender: TObject; const pDisp: IDispatch;
                              const URL, Flags, TargetFrameName,
                              PostData, Headers: OleVariant;
                              var Cancel: WordBool);
var
  Params: TStringList;
begin
  {$IFDEF DEBUG}
  if Owner is TForm then
    TextoMemo_Form('DoBeforeNavigate - URL: ' + URL);
  {$ENDIF}

  if FComunicacaoJS.IsCallbackURL(URL) then
  begin
    Cancel := True;

    {$IFDEF DEBUG}
    if Owner is TForm then
      TextoMemo_Form('DoBeforeNavigate - Detectado callback URL');
    {$ENDIF}

    Params := FComunicacaoJS.ExtrairParametrosCallback(URL);
    try
      ProcessarCallback(Params);
    finally
      Params.Free;
    end;
  end;

  if Assigned(FOldBeforeNavigate) then
    FOldBeforeNavigate(ASender, pDisp, URL, Flags, TargetFrameName, PostData, Headers, Cancel);
end;

procedure TFrameGraficoBase.WebBrowserBeforeNavigate2(ASender: TObject; const pDisp: IDispatch;
                          const URL, Flags, TargetFrameName,
                          PostData, Headers: OleVariant;
                          var Cancel: WordBool);
begin
  // Este evento será substituído no construtor por DoBeforeNavigate
end;

procedure TFrameGraficoBase.WebBrowserDocumentComplete(ASender: TObject; const pDisp: IDispatch;
                              const URL: OleVariant);
begin
  {$IFDEF DEBUG}
  if Owner is TForm then
    TextoMemo_Form('WebBrowserDocumentComplete - URL: ' + URL);
  {$ENDIF}

  // Marcar como carregado se for o documento principal
  if (pDisp = WebBrowser.Application) then
  begin
    FHTMLCarregado := True;

    {$IFDEF DEBUG}
    if Owner is TForm then
      TextoMemo_Form('WebBrowserDocumentComplete - Documento principal carregado');
    {$ENDIF}
  end;
end;

procedure TFrameGraficoBase.CarregarHTML(const HTML: string);
var
  TempFile: string;
  HTMLFile: TStringList;
begin
  {$IFDEF DEBUG}
  if Owner is TForm then
    TextoMemo_Form('CarregarHTML - Início');
  {$ENDIF}

  if not Assigned(WebBrowser) then
  begin
    {$IFDEF DEBUG}
    if Owner is TForm then
      TextoMemo_Form('CarregarHTML - WebBrowser não está inicializado');
    {$ENDIF}
    Exit;
  end;

  // Salvar em arquivo temporário
  TempFile := ExtractFilePath(ParamStr(0)) + 'chart_' + FormatDateTime('hhnnsszzz', Now) + '.html';

  {$IFDEF DEBUG}
  if Owner is TForm then
    TextoMemo_Form('CarregarHTML - Salvando em: ' + TempFile);
  {$ENDIF}

  HTMLFile := TStringList.Create;
  try
    HTMLFile.Text := TConversaoDados.ConvertStringToUTF8(HTML);

    // Verificar se o diretório existe
    if not DirectoryExists(ExtractFilePath(TempFile)) then
      ForceDirectories(ExtractFilePath(TempFile));

    HTMLFile.SaveToFile(TempFile);

    // Navegar para o arquivo com URL válida
    WebBrowser.Navigate('file:///' + StringReplace(TempFile, '\', '/', [rfReplaceAll]));

    // Aguardar carregamento
    while WebBrowser.ReadyState <> READYSTATE_COMPLETE do
      Application.ProcessMessages;

    FHTMLCarregado := True;

    {$IFDEF DEBUG}
    if Owner is TForm then
      TextoMemo_Form('CarregarHTML - HTML carregado com sucesso');
    {$ENDIF}
  finally
    HTMLFile.Free;
  end;
end;

procedure TFrameGraficoBase.ExecutarScript(const Script: string);
begin
  {$IFDEF DEBUG}
  if Owner is TForm then
  begin
    TextoMemo_Form('ExecutarScript - Início');
    TextoMemo_Form('ExecutarScript - Script: ' + Copy(Script, 1, 100) + '...');
  end;
  {$ENDIF}

  if not VerificarBrowser then
  begin
    {$IFDEF DEBUG}
    if Owner is TForm then
      TextoMemo_Form('ExecutarScript - WebBrowser não está pronto');
    {$ENDIF}
    Exit;
  end;

  try
    FComunicacaoJS.ExecutarScript(WebBrowser, Script);

    {$IFDEF DEBUG}
    if Owner is TForm then
      TextoMemo_Form('ExecutarScript - Script executado com sucesso');
    {$ENDIF}
  except
    on E: Exception do
    begin
      {$IFDEF DEBUG}
      if Owner is TForm then
        TextoMemo_Form('ExecutarScript - Erro ao executar script: ' + E.Message);
      {$ENDIF}
    end;
  end;
end;

procedure TFrameGraficoBase.FrameResize(Sender: TObject);
begin
  {$IFDEF DEBUG}
  if Owner is TForm then
    TextoMemo_Form('FrameResize - Novas dimensões: ' + IntToStr(Width) + 'x' + IntToStr(Height));
  {$ENDIF}

  // Atualizar dimensões do gráfico quando o frame for redimensionado
  if VerificarBrowser and FHTMLCarregado then
  begin
    {$IFDEF DEBUG}
    if Owner is TForm then
      TextoMemo_Form('FrameResize - Atualizando gráfico');
    {$ENDIF}

    AtualizarGrafico;
  end;
end;

procedure TFrameGraficoBase.CarregarDados(ADataSet: TDataSet);
var
  HTMLBase: string;
  TempFile: string;
  HTMLFile: TStringList;
begin
  {$IFDEF DEBUG}
  TextoMemo_Form('CarregarDados - Início');
  if Assigned(ADataSet) then
    TextoMemo_Form('CarregarDados - Dataset: ' + IntToStr(ADataSet.RecordCount) + ' registros')
  else
    TextoMemo_Form('CarregarDados - Dataset não atribuído');
  {$ENDIF}

  FDataSet := ADataSet;

  if FDataSet <> nil then
  begin
    // Gerar o HTML
    HTMLBase := GerarHTMLBase;

    {$IFDEF DEBUG}
    TextoMemo_Form('CarregarDados - HTML gerado, tamanho: ' + IntToStr(Length(HTMLBase)));
    TextoMemo_Form('CarregarDados - HTML completo:');
    TextoMemo_Form('--- INÍCIO HTML ---');
    TextoMemo_Form(HTMLBase);
    TextoMemo_Form('--- FIM HTML ---');
    {$ENDIF}

    // Salvar em arquivo
    TempFile := ExtractFilePath(ParamStr(0)) + 'chart_' + FormatDateTime('hhnnsszzz', Now) + '.html';

    {$IFDEF DEBUG}
    TextoMemo_Form('CarregarDados - HTML será salvo em: ' + TempFile);
    {$ENDIF}

    HTMLFile := TStringList.Create;
    try
      HTMLFile.Text := TConversaoDados.ConvertStringToUTF8(HTMLBase);

      // Verificar se o diretório existe
      if not DirectoryExists(ExtractFilePath(TempFile)) then
        ForceDirectories(ExtractFilePath(TempFile));

      HTMLFile.SaveToFile(TempFile);

      {$IFDEF DEBUG}
      TextoMemo_Form('CarregarDados - HTML salvo com sucesso em: ' + TempFile);
      {$ENDIF}

      // Carregar no navegador
      WebBrowser.Navigate('file:///' + StringReplace(TempFile, '\', '/', [rfReplaceAll]));

      // Aguardar carregamento
      while WebBrowser.ReadyState <> READYSTATE_COMPLETE do
        Application.ProcessMessages;

      FHTMLCarregado := True;

      {$IFDEF DEBUG}
      TextoMemo_Form('CarregarDados - HTML carregado via Navigate');
      {$ENDIF}

      // Verificar se o gráfico foi carregado corretamente após um tempo
      TThread.CreateAnonymousThread(procedure
      begin
        Sleep(1000); // Aguardar 1 segundo

        {$IFDEF DEBUG}
        TThread.Synchronize(nil, procedure
        begin
          TextoMemo_Form('CarregarDados - Thread de verificação iniciada após 1000ms');
        end);
        {$ENDIF}

        TThread.Synchronize(nil, procedure
        begin
          try
            // Verificar se o gráfico está visível executando um script
            var scriptVerificacao := 'if (typeof window.myChart !== "undefined") { ' +
                                    'location.assign("ActionCallBackJS:GraficoCarregado(ok)"); } ' +
                                    'else { ' +
                                    'location.assign("ActionCallBackJS:GraficoCarregado(falha)"); }';

            {$IFDEF DEBUG}
            TextoMemo_Form('CarregarDados - Script de verificação: ' + scriptVerificacao);
            {$ENDIF}

            ExecutarScript(scriptVerificacao);
          except
            on E: Exception do
            begin
              {$IFDEF DEBUG}
              TextoMemo_Form('CarregarDados - Erro ao verificar gráfico: ' + E.Message);
              {$ENDIF}
            end;
          end;
        end);
      end).Start;
    finally
      HTMLFile.Free;
    end;
  end
  else
  begin
    {$IFDEF DEBUG}
    TextoMemo_Form('CarregarDados - Nenhuma ação realizada (DataSet nulo)');
    {$ENDIF}
  end;
end;
procedure TFrameGraficoBase.LimparGrafico;
begin
  {$IFDEF DEBUG}
  if Owner is TForm then
    TextoMemo_Form('LimparGrafico - Início');
  {$ENDIF}

  FHTMLCarregado := False;
  CarregarHTML('<html><body><div style="color:#666;font-family:Arial;text-align:center;margin-top:20px;">Sem dados para exibir</div></body></html>');
end;

procedure TFrameGraficoBase.ProcessarCallback(const Params: TStringList);
begin
  {$IFDEF DEBUG}
  if Owner is TForm then
  begin
    TextoMemo_Form('ProcessarCallback - Início');

    if Assigned(Params) then
    begin
      TextoMemo_Form('ProcessarCallback - Número de parâmetros: ' + IntToStr(Params.Count));

      for var i := 0 to Params.Count - 1 do
        TextoMemo_Form('ProcessarCallback - Param[' + IntToStr(i) + ']: ' + Params[i]);
    end;
  end;
  {$ENDIF}

  if Assigned(FOnCallback) then
  begin
    {$IFDEF DEBUG}
    if Owner is TForm then
      TextoMemo_Form('ProcessarCallback - Chamando handler de callback');
    {$ENDIF}

    FOnCallback(Self, Params);
  end
  else
  begin
    {$IFDEF DEBUG}
    if Owner is TForm then
      TextoMemo_Form('ProcessarCallback - Nenhum handler de callback atribuído');
    {$ENDIF}
  end;
end;

procedure TFrameGraficoBase.DiagnosticarProblemas;
begin
  if not Assigned(WebBrowser) then
  begin
    {$IFDEF DEBUG}
    if Owner is TForm then
      TextoMemo_Form('DiagnosticarProblemas - WebBrowser não atribuído');
    {$ENDIF}
    Exit;
  end;

  try
    // Verificar configuração do IE
    var Reg := TRegistry.Create;
    try
      Reg.RootKey := HKEY_CURRENT_USER;
      if Reg.OpenKeyReadOnly('Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION') then
      begin
        var AppName := ExtractFileName(ParamStr(0));
        if Reg.ValueExists(AppName) then
        begin
          var Valor := Reg.ReadInteger(AppName);
          {$IFDEF DEBUG}
          if Owner is TForm then
            TextoMemo_Form('DiagnosticarProblemas - Registro IE: ' + IntToStr(Valor));
          {$ENDIF}
        end
        else
        begin
          {$IFDEF DEBUG}
          if Owner is TForm then
            TextoMemo_Form('DiagnosticarProblemas - Registro não configurado para ' + AppName);
          {$ENDIF}
        end;
      end;
    finally
      Reg.Free;
    end;

    // Verifica se o WebBrowser está inicializado
    {$IFDEF DEBUG}
    if Owner is TForm then
    begin
      TextoMemo_Form('DiagnosticarProblemas - ReadyState: ' + IntToStr(Integer(WebBrowser.ReadyState)));
      TextoMemo_Form('DiagnosticarProblemas - Silent: ' + BoolToStr(WebBrowser.Silent, True));
    end;
    {$ENDIF}

    // Tenta carregar uma página HTML simples
    WebBrowser.Navigate('about:blank');
    while WebBrowser.ReadyState <> READYSTATE_COMPLETE do
      Application.ProcessMessages;

    if Assigned(WebBrowser.Document) then
    begin
      var Doc := WebBrowser.Document as IHTMLDocument2;
      var V: OleVariant;
      V := '<html><body><h1>Teste de Diagnóstico</h1></body></html>';

      try
        Doc.open('', EmptyParam, EmptyParam, EmptyParam);
        Doc.write(PSafeArray(TVarData(V).VArray));
        Doc.close;

        {$IFDEF DEBUG}
        if Owner is TForm then
          TextoMemo_Form('DiagnosticarProblemas - HTML básico carregado com sucesso');
        {$ENDIF}
      except
        on E: Exception do
        begin
          {$IFDEF DEBUG}
          if Owner is TForm then
            TextoMemo_Form('DiagnosticarProblemas - Erro ao carregar HTML básico: ' + E.Message);
          {$ENDIF}
        end;
      end;
    end;
  except
    on E: Exception do
    begin
      {$IFDEF DEBUG}
      if Owner is TForm then
        TextoMemo_Form('DiagnosticarProblemas - Erro: ' + E.Message);
      {$ENDIF}
    end;
  end;
end;

end.
