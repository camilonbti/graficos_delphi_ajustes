unit uChartPie;

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
  TChartPie = class(TChartGoogle)
  private
    FOptions: TStringList;
    FNome   : string;
  public
    constructor Create(Nome: string; Altura, Largura: Int64);
    destructor Destroy; override;

    property Options: TStringList read FOptions write FOptions;
    function Titulos(ATitulo: string): TChartPie;
  end;

var
  ChartPie: TChartPie;

implementation

{ TChartPie }

constructor TChartPie.Create(Nome: string; Altura, Largura: Int64);
begin
  inherited Create(Nome, Altura, Largura);

  FNome    := LowerCase(StringReplace(Nome, ' ', '_', []));
  FOptions := TStringList.Create;
  Self.Grafico(ctPizza);
end;

destructor TChartPie.Destroy;
begin
  if Assigned(FOptions) then
    FreeAndNil(FOptions);

  inherited;
end;

function TChartPie.Titulos(ATitulo: string): TChartPie;
begin
  Result := Self;
  Result.Titulo(ATitulo, '', '');

  Options.Text := //
    '      var options = { ' + sLineBreak + //
    '          title: ''' + ATitulo + ''' ' + sLineBreak + //
    '      }; ' + sLineBreak + //
    '  var chart = new google.visualization.PieChart(document.getElementById(''' + FNome + ''')); ' + sLineBreak;
  Result.Opcoes(Options.Text);
end;

end.
