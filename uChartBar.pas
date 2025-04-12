unit uChartBar;

interface

uses
  uChartGoogle,
  //
  SysUtils,
  StrUtils,
  Classes,
  Registry,
  Data.DB,
  SHDocVw,
  Winapi.ActiveX,
  Winapi.Windows;

type
  TChartBar = class(TChartGoogle)
  private
    FOptions         : TStringList;
    FOrientacao      : TTipoOrientacao;
    FSubtitulo, FNome: string;
  public
    constructor Create(Nome: string; Altura, Largura: Int64; Orientacao: TTipoOrientacao);
    destructor Destroy; override;

    property Options: TStringList read FOptions write FOptions;
    function Titulos(ATitulo, Subtitulo: string; Vertical: string = ''; Horizontal: string = ''): TChartBar;
  end;

var
  ChartBar: TChartBar;

implementation

{ TChartBar }

constructor TChartBar.Create(Nome: string; Altura, Largura: Int64; Orientacao: TTipoOrientacao);
begin
  inherited Create(Nome, Altura, Largura);

  FOptions := TStringList.Create;
  FNome    := LowerCase(StringReplace(Nome, ' ', '_', []));
  Self.Grafico(ctBarra, Orientacao);
  FOrientacao := Orientacao;
end;

destructor TChartBar.Destroy;
begin
  if Assigned(FOptions) then
    FreeAndNil(FOptions);

  inherited;
end;

function TChartBar.Titulos(ATitulo, Subtitulo: string; Vertical: string = ''; Horizontal: string = ''): TChartBar;
var
  lDesc: string;
begin
  Result     := Self;
  FSubtitulo := Subtitulo;
  Result.Titulo(ATitulo, Vertical, Horizontal);

  if not Trim(FSubtitulo).IsEmpty then
  begin
    case FOrientacao of
      toVertical:
        lDesc := 'vertical';
      toHorizontal:
        lDesc := 'horizontal';
    end;

    Options.Text := //
      '      var options = {                                                  ' + sLineBreak + //
      '        chart: {                                                       ' + sLineBreak + //
      '          title: ''' + ATitulo + ''',                                  ' + sLineBreak + //
      '          subtitle: ''' + FSubtitulo + ''',                            ' + sLineBreak + //
      '        },                                                             ' + sLineBreak + //
      '        bars: ''' + lDesc + ''' // Required for Material Bar Charts.   ' + sLineBreak + //
      '      };                                                               ' + sLineBreak + //
      '                                                                       ' + sLineBreak + //
      '    var chart = new google.charts.Bar(document.getElementById(''' + FNome + ''')); ' + sLineBreak;
  end;

  Result.Header(['corechart', 'bar']).Opcoes(Options.Text);
end;

end.
