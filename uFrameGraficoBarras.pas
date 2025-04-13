unit uFrameGraficoBarras;

interface

uses
  uFuncoes,
  uFrameGraficoBase,
  uGraficoChartJS,
  uConversaoDados,

  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Data.DB, System.JSON, Vcl.OleCtrls, SHDocVw;

type
  TFrameGraficoBarras = class(TFrameGraficoBase)
  private
    FGraficoChartJS: TGraficoBarrasChartJS;
    FCampoRotulo: string;
    FCampoValor: string;
    FTitulo: string;
    FHorizontal: Boolean;

    procedure SetHorizontal(const Value: Boolean);
  protected
    function GerarHTMLBase: string; override;
    procedure ProcessarCallback(const Params: TStringList); override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    procedure AtualizarGrafico; override;
    procedure ConfigurarGrafico(const ACampoRotulo, ACampoValor, ATitulo: string);

    property CampoRotulo: string read FCampoRotulo write FCampoRotulo;
    property CampoValor: string read FCampoValor write FCampoValor;
    property Titulo: string read FTitulo write FTitulo;
    property Horizontal: Boolean read FHorizontal write SetHorizontal;
  end;

implementation

{$R *.dfm}

constructor TFrameGraficoBarras.Create(AOwner: TComponent);
begin
  inherited;
  FGraficoChartJS := TGraficoBarrasChartJS.Create;
  FHorizontal := True; // Barras horizontais por padrão
end;

destructor TFrameGraficoBarras.Destroy;
begin
  FGraficoChartJS.Free;
  inherited;
end;

procedure TFrameGraficoBarras.SetHorizontal(const Value: Boolean);
begin
  if FHorizontal <> Value then
  begin
    FHorizontal := Value;
    if Assigned(DataSet) then
      AtualizarGrafico;
  end;
end;

function TFrameGraficoBarras.GerarHTMLBase: string;
var
  HTML: string;
begin
  {$IFDEF DEBUG}
  TextoMemo_Form('GerarHTMLBase - Início - Dataset Assigned: ' + BoolToStr(Assigned(DataSet), True));
  if Assigned(DataSet) then
    TextoMemo_Form('GerarHTMLBase - Dataset possui ' + IntToStr(DataSet.RecordCount) + ' registros');
  {$ENDIF}

  FGraficoChartJS.Titulo := FTitulo;
  FGraficoChartJS.CampoRotulo := FCampoRotulo;
  FGraficoChartJS.CampoValor := FCampoValor;
  FGraficoChartJS.Horizontal := FHorizontal;

  try
    HTML := FGraficoChartJS.GerarHTML(DataSet);
    {$IFDEF DEBUG}
    TextoMemo_Form('GerarHTMLBase - HTML gerado com sucesso - Tamanho: ' + IntToStr(Length(HTML)));
    TextoMemo_Form('GerarHTMLBase - Primeiros 200 caracteres: ' + Copy(HTML, 1, 200));
    {$ENDIF}
    Result := HTML;
  except
    on E: Exception do
    begin
      {$IFDEF DEBUG}
      TextoMemo_Form('GerarHTMLBase - Erro ao gerar HTML: ' + E.Message);
      {$ENDIF}
      Result := '<html><body>Erro ao gerar gráfico: ' + E.Message + '</body></html>';
    end;
  end;
end;

procedure TFrameGraficoBarras.ConfigurarGrafico(const ACampoRotulo, ACampoValor, ATitulo: string);
begin
  FCampoRotulo := ACampoRotulo;
  FCampoValor := ACampoValor;
  FTitulo := ATitulo;

  if Assigned(DataSet) then
    AtualizarGrafico;
end;

procedure TFrameGraficoBarras.AtualizarGrafico;
var
  Script: string;
begin
  {$IFDEF DEBUG}
  TextoMemo_Form('AtualizarGrafico - Início');
  TextoMemo_Form('AtualizarGrafico - Dataset Assigned: ' + BoolToStr(Assigned(DataSet), True));
  TextoMemo_Form('AtualizarGrafico - WebBrowser Assigned: ' + BoolToStr(Assigned(WebBrowser), True));
  if Assigned(WebBrowser) then
    TextoMemo_Form('AtualizarGrafico - WebBrowser.Document Assigned: ' + BoolToStr(Assigned(WebBrowser.Document), True));
  {$ENDIF}

  if not Assigned(DataSet) or not Assigned(WebBrowser) or not Assigned(WebBrowser.Document) then
  begin
    {$IFDEF DEBUG}
    TextoMemo_Form('AtualizarGrafico - Saindo devido a componentes não atribuídos');
    {$ENDIF}
    Exit;
  end;

  try
    Script := FGraficoChartJS.GerarScriptAtualizacao(DataSet);
    {$IFDEF DEBUG}
    TextoMemo_Form('AtualizarGrafico - Script gerado: ' + Copy(Script, 1, 100) + '...');
    {$ENDIF}
    ExecutarScript(Script);
  except
    on E: Exception do
    begin
      {$IFDEF DEBUG}
      TextoMemo_Form('AtualizarGrafico - Erro ao atualizar gráfico: ' + E.Message);
      {$ENDIF}
    end;
  end;
end;

procedure TFrameGraficoBarras.ProcessarCallback(const Params: TStringList);
begin
  // Processar parâmetros específicos de barras antes de chamar o evento principal
  inherited;
end;

end.
