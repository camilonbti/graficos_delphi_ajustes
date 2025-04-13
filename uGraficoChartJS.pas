unit uGraficoChartJS;

interface

uses
  uFuncoes,
  uConversaoDados,
  uTemplatesChartJS,

  Vcl.Forms, System.SysUtils, System.Classes, Data.DB, System.JSON;

type
  TGraficoChartJS = class abstract
  private
    FTitulo: string;
    FCampoRotulo: string;
    FCampoValor: string;
    FCampoID: string; // Novo campo para identificação (filtro)
    FIdentificador: string;
    FCores: TStringList;
    FRandomCores: Boolean;

  protected
    function ObterTemplate: string; virtual; abstract;
    function GerarDadosJSON(const ADataSet: TDataSet): string; virtual;
    function GerarConfigJSON: string; virtual;
    function GerarScriptCallback: string; virtual;
    function GerarCores(Quantidade: Integer): string;

  public
    constructor Create;
    destructor Destroy; override;

    function GerarHTML(const ADataSet: TDataSet): string; virtual;
    function GerarScriptAtualizacao(const ADataSet: TDataSet): string; virtual;

    property Titulo: string read FTitulo write FTitulo;
    property CampoRotulo: string read FCampoRotulo write FCampoRotulo;
    property CampoValor: string read FCampoValor write FCampoValor;
    property CampoID: string read FCampoID write FCampoID; // Nova propriedade
    property Identificador: string read FIdentificador write FIdentificador;
    property RandomCores: Boolean read FRandomCores write FRandomCores;
  end;

  TGraficoBarrasChartJS = class(TGraficoChartJS)
  private
    FHorizontal: Boolean;
  protected
    function ObterTemplate: string; override;
    function GerarConfigJSON: string; override;
  public
    constructor Create;

    property Horizontal: Boolean read FHorizontal write FHorizontal;
  end;

implementation

constructor TGraficoChartJS.Create;
begin
  FIdentificador := 'grafico' + FormatDateTime('hhnnsszzz', Now);
  FCores := TStringList.Create;
  FRandomCores := True;
  FCampoID := ''; // Inicializa o campo ID

  // Cores padrão
  FCores.Add('rgba(255, 99, 132, 0.8)');
  FCores.Add('rgba(54, 162, 235, 0.8)');
  FCores.Add('rgba(255, 206, 86, 0.8)');
  FCores.Add('rgba(75, 192, 192, 0.8)');
  FCores.Add('rgba(153, 102, 255, 0.8)');
end;

destructor TGraficoChartJS.Destroy;
begin
  FCores.Free;
  inherited;
end;

function TGraficoChartJS.GerarDadosJSON(const ADataSet: TDataSet): string;
var
  Rotulos, Valores, IDValores: TStringList;
  Bookmark: TBookmark;
  FloatValue: Double;
  FloatStr: string;
  SB: TStringBuilder;
  I: Integer;
begin
  {$IFDEF DEBUG}
  TextoMemo_Form('GerarDadosJSON - Início');
  if Assigned(ADataSet) then
    TextoMemo_Form('GerarDadosJSON - Dataset possui ' + IntToStr(ADataSet.RecordCount) + ' registros')
  else
    TextoMemo_Form('GerarDadosJSON - Dataset não está atribuído');
  {$ENDIF}

  if not Assigned(ADataSet) or ADataSet.IsEmpty then
  begin
    {$IFDEF DEBUG}
    TextoMemo_Form('GerarDadosJSON - Dataset vazio ou não atribuído, retornando JSON vazio');
    {$ENDIF}
    Exit('{ labels: [], datasets: [{ data: [], idValues: [] }] }');
  end;

  Rotulos := TStringList.Create;
  Valores := TStringList.Create;
  IDValores := TStringList.Create;
  SB := TStringBuilder.Create;

  try
    {$IFDEF DEBUG}
    TextoMemo_Form('GerarDadosJSON - Verificando campos');
    TextoMemo_Form('GerarDadosJSON - Campo Rótulo: ' + FCampoRotulo);
    TextoMemo_Form('GerarDadosJSON - Campo Valor: ' + FCampoValor);
    TextoMemo_Form('GerarDadosJSON - Campo ID: ' + FCampoID);

    // Verificar se os campos existem no dataset
    if ADataSet.FindField(FCampoRotulo) = nil then
      TextoMemo_Form('GerarDadosJSON - ERRO: Campo rótulo "' + FCampoRotulo + '" não encontrado no dataset')
    else
      TextoMemo_Form('GerarDadosJSON - Campo rótulo "' + FCampoRotulo + '" encontrado');

    if ADataSet.FindField(FCampoValor) = nil then
      TextoMemo_Form('GerarDadosJSON - ERRO: Campo valor "' + FCampoValor + '" não encontrado no dataset')
    else
      TextoMemo_Form('GerarDadosJSON - Campo valor "' + FCampoValor + '" encontrado');

    if (Trim(FCampoID) <> '') then
    begin
      if ADataSet.FindField(FCampoID) = nil then
        TextoMemo_Form('GerarDadosJSON - AVISO: Campo ID "' + FCampoID + '" não encontrado no dataset')
      else
        TextoMemo_Form('GerarDadosJSON - Campo ID "' + FCampoID + '" encontrado');
    end;
    {$ENDIF}

    Bookmark := ADataSet.Bookmark;
    ADataSet.DisableControls;

    try
      ADataSet.First;
      var recordCount := 0;

      {$IFDEF DEBUG}
      TextoMemo_Form('GerarDadosJSON - Iniciando loop de extração de dados');
      {$ENDIF}

      while not ADataSet.Eof do
      begin
        Inc(recordCount);

        // Registrar alguns registros para debug
        {$IFDEF DEBUG}
        if recordCount <= 5 then
        begin
          TextoMemo_Form('GerarDadosJSON - Registro ' + IntToStr(recordCount) + ':');
          TextoMemo_Form('  - Rótulo: ' + ADataSet.FieldByName(FCampoRotulo).AsString);
          TextoMemo_Form('  - Valor: ' + FloatToStr(ADataSet.FieldByName(FCampoValor).AsFloat));
          if (Trim(FCampoID) <> '') and (ADataSet.FindField(FCampoID) <> nil) then
            TextoMemo_Form('  - ID: ' + ADataSet.FieldByName(FCampoID).AsString);
        end;
        {$ENDIF}

        // Adicionar rótulos sem aspas extras - serão adicionadas depois na construção JSON
        Rotulos.Add(ADataSet.FieldByName(FCampoRotulo).AsString);

        // Corrigir formato numérico para JavaScript (ponto decimal)
        FloatValue := ADataSet.FieldByName(FCampoValor).AsFloat;
        FloatStr := FloatToStr(FloatValue);
        FloatStr := StringReplace(FloatStr, ',', '.', [rfReplaceAll]);
        Valores.Add(FloatStr);

        // Adicionar valores de ID se o campo foi especificado
        if (Trim(FCampoID) <> '') and (ADataSet.FindField(FCampoID) <> nil) then
          IDValores.Add(ADataSet.FieldByName(FCampoID).AsString)
        else
          IDValores.Add('');

        ADataSet.Next;
      end;

      {$IFDEF DEBUG}
      TextoMemo_Form('GerarDadosJSON - Dados extraídos. Total de registros processados: ' + IntToStr(recordCount));
      TextoMemo_Form('GerarDadosJSON - Total de rótulos: ' + IntToStr(Rotulos.Count));
      TextoMemo_Form('GerarDadosJSON - Total de valores: ' + IntToStr(Valores.Count));
      TextoMemo_Form('GerarDadosJSON - Total de IDs: ' + IntToStr(IDValores.Count));
      {$ENDIF}

      // Construção manual do JSON para evitar problemas com aspas
      SB.Capacity := (Rotulos.Count * 50) + 200; // Pré-alocar memória suficiente
      SB.Append('{ labels: [');

      // Adicionar rótulos corretamente entre aspas
      for I := 0 to Rotulos.Count - 1 do
      begin
        SB.Append('"').Append(Rotulos[I]).Append('"');
        if I < Rotulos.Count - 1 then
          SB.Append(', ');
      end;

      SB.Append('], datasets: [{ data: [');

      // Adicionar valores numéricos (sem aspas)
      for I := 0 to Valores.Count - 1 do
      begin
        SB.Append(Valores[I]);
        if I < Valores.Count - 1 then
          SB.Append(', ');
      end;

      SB.Append('], idValues: [');

      // Adicionar IDs corretamente entre aspas
      for I := 0 to IDValores.Count - 1 do
      begin
        SB.Append('"').Append(IDValores[I]).Append('"');
        if I < IDValores.Count - 1 then
          SB.Append(', ');
      end;

      SB.Append('], backgroundColor: ').Append(GerarCores(Valores.Count));
      SB.Append(', borderWidth: 1 }] }');

      Result := SB.ToString;

      {$IFDEF DEBUG}
      TextoMemo_Form('GerarDadosJSON - JSON gerado com sucesso');
      TextoMemo_Form('GerarDadosJSON - Tamanho do JSON: ' + IntToStr(Length(Result)) + ' caracteres');
      TextoMemo_Form('GerarDadosJSON - Primeiros 200 caracteres: ' + Copy(Result, 1, 200));
      {$ENDIF}
    finally
      ADataSet.Bookmark := Bookmark;
      ADataSet.EnableControls;
    end;
  finally
    Rotulos.Free;
    Valores.Free;
    IDValores.Free;
    SB.Free;
  end;
end;

function TGraficoChartJS.GerarConfigJSON: string;
begin
  Result := '{ ' +
            '  responsive: true, ' +
            '  maintainAspectRatio: false, ' +
            '  title: { ' +
            '    display: true, ' +
            '    text: ' + QuotedStr(FTitulo) + ' ' +
            '  }, ' +
            '  legend: { ' +
            '    display: false ' +
            '  } ' +
            '}';
end;

function TGraficoChartJS.GerarScriptCallback: string;
begin
  Result :=
    'document.getElementById("' + FIdentificador + '").onclick = function(evt) { ' +
    '  var activePoints = myChart.getElementsAtEvent(evt); ' +
    '  if (activePoints.length > 0) { ' +
    '    var clickedElementIndex = activePoints[0]._index; ' +
    '    var label = myChart.data.labels[clickedElementIndex]; ' +
    '    var value = myChart.data.datasets[0].data[clickedElementIndex]; ' +
    '    var idValue = myChart.data.datasets[0].idValues[clickedElementIndex]; ' +
    '    var url = "ActionCallBackJS:Retorno("+JSON.stringify([label, value, "' + FCampoID + '", idValue])+")"; ' +
    '    location.assign(url); ' +
    '  } ' +
    '};';
end;

function TGraficoChartJS.GerarCores(Quantidade: Integer): string;
const
  CoresPadraoSistema: array[0..9] of string = (
    'rgba(255, 99, 132, 0.6)',   // Vermelho
    'rgba(54, 162, 235, 0.6)',   // Azul
    'rgba(255, 206, 86, 0.6)',   // Amarelo
    'rgba(75, 192, 192, 0.6)',   // Verde
    'rgba(153, 102, 255, 0.6)',  // Roxo
    'rgba(255, 159, 64, 0.6)',   // Laranja
    'rgba(199, 199, 199, 0.6)',  // Cinza
    'rgba(83, 102, 255, 0.6)',   // Azul escuro
    'rgba(255, 99, 71, 0.6)',    // Tomate
    'rgba(0, 128, 128, 0.6)'     // Verde-azulado
  );
var
  I: Integer;
  SB: TStringBuilder;
begin
  SB := TStringBuilder.Create;
  try
    SB.Append('[');
    for I := 0 to Quantidade - 1 do
    begin
      SB.Append('"').Append(CoresPadraoSistema[I mod Length(CoresPadraoSistema)]).Append('"');
      if I < Quantidade - 1 then
        SB.Append(', ');
    end;
    SB.Append(']');
    Result := SB.ToString;
  finally
    SB.Free;
  end;
end;

function TGraficoChartJS.GerarHTML(const ADataSet: TDataSet): string;
var
  Template: string;
  DadosJSON, ConfigJSON: string;
begin
  {$IFDEF DEBUG}
  TextoMemo_Form('GerarHTML - Início da geração do HTML');
  if Assigned(ADataSet) then
    TextoMemo_Form('GerarHTML - Dataset possui ' + IntToStr(ADataSet.RecordCount) + ' registros')
  else
    TextoMemo_Form('GerarHTML - Dataset não está atribuído');
  {$ENDIF}

  Template := ObterTemplate;

  try
    {$IFDEF DEBUG}
    TextoMemo_Form('GerarHTML - Template base obtido, tamanho: ' + IntToStr(Length(Template)));
    TextoMemo_Form('GerarHTML - Gerando dados JSON...');
    {$ENDIF}

    DadosJSON := GerarDadosJSON(ADataSet);

    {$IFDEF DEBUG}
    TextoMemo_Form('GerarHTML - Dados JSON gerados, tamanho: ' + IntToStr(Length(DadosJSON)));
    TextoMemo_Form('GerarHTML - Gerando configuração JSON...');
    {$ENDIF}

    ConfigJSON := GerarConfigJSON;

    {$IFDEF DEBUG}
    TextoMemo_Form('GerarHTML - Configuração JSON gerada, tamanho: ' + IntToStr(Length(ConfigJSON)));
    TextoMemo_Form('GerarHTML - Substituindo variáveis no template...');
    {$ENDIF}

    Template := StringReplace(Template, '{{GRAFICO_ID}}', FIdentificador, [rfReplaceAll]);
    Template := StringReplace(Template, '{{DADOS_JSON}}', DadosJSON, [rfReplaceAll]);
    Template := StringReplace(Template, '{{CONFIG_JSON}}', ConfigJSON, [rfReplaceAll]);
    Template := StringReplace(Template, '{{SCRIPT_CALLBACK}}', GerarScriptCallback, [rfReplaceAll]);

    {$IFDEF DEBUG}
    TextoMemo_Form('GerarHTML - Template completo pronto com ' + IntToStr(Length(Template)) + ' caracteres');
    {$ENDIF}

    Result := Template;
  except
    on E: Exception do
    begin
      {$IFDEF DEBUG}
      TextoMemo_Form('GerarHTML - ERRO: ' + E.Message);
      TextoMemo_Form('GerarHTML - Local do erro: ' + E.StackTrace);
      {$ENDIF}
      Result := '<!DOCTYPE html><html><body>Erro ao gerar HTML: ' + E.Message + '</body></html>';
    end;
  end;
end;

function TGraficoChartJS.GerarScriptAtualizacao(const ADataSet: TDataSet): string;
begin
  Result := 'myChart.data = ' + GerarDadosJSON(ADataSet) + '; ' +
            'myChart.update();';
end;

{ TGraficoBarrasChartJS }

constructor TGraficoBarrasChartJS.Create;
begin
  inherited Create;
  FHorizontal := True;
end;

function TGraficoBarrasChartJS.ObterTemplate: string;
begin
  Result := TTemplatesChartJS.TemplateBarras;
end;

function TGraficoBarrasChartJS.GerarConfigJSON: string;
var
  Config: string;
begin
  Config := '{' +
            '  responsive: true,' +
            '  maintainAspectRatio: false,' +
            '  title: {' +
            '    display: true,' +
            '    text: ' + QuotedStr(FTitulo) +
            '  },' +
            '  scales: {' +
            '    y: {' +
            '      beginAtZero: true' +
            '    },' +
            '    x: {' +
            '      ticks: {' +
            '        maxRotation: 45,' +
            '        minRotation: 45' +
            '      }' +
            '    }' +
            '  },' +
            '  plugins: {' +
            '    legend: {' +
            '      display: false' +
            '    },' +
            '    tooltip: {' +
            '      callbacks: {' +
            '        label: function(context) {' +
            '          let label = context.dataset.label || "";' +
            '          if (label) {' +
            '            label += ": ";' +
            '          }' +
            '          label += new Intl.NumberFormat("pt-BR", { ' +
            '            style: "currency", ' +
            '            currency: "BRL" ' +
            '          }).format(context.raw);' +
            '          return label;' +
            '        }' +
            '      }' +
            '    }' +
            '  }';

  // Adicionar a opção indexAxis para barras horizontais
  if FHorizontal then
    Config := Config + ', indexAxis: "y"';

  Config := Config + '}';

  Result := Config;
end;

end.
