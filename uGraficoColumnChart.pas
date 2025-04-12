unit uGraficoColumnChart;

interface

uses
  uGraficoGoogle,
  System.StrUtils;

type
  TGraficoColumnChart = class(TGraficoGoogle)
  private
    FTitle: string;
  public
    constructor Create(ARandomColor: Boolean = True); overload;
    destructor Destroy; override;

    function package: string; override;

    function GetFuncaoDraw(AElementID: string; idx: integer): string; override;
    function GetUpdateData(Idx: Integer): string; override;

    procedure AddTitle(ATitle: string); override;
  end;

implementation

uses
  SysUtils,
  uGraficos;

{ TGraficoColumnChart }

constructor TGraficoColumnChart.Create(ARandomColor: Boolean = True);
begin
  inherited Create;
  FRandomColor := ARandomColor;
end;

destructor TGraficoColumnChart.Destroy;
begin
  inherited;
end;

procedure TGraficoColumnChart.AddTitle(ATitle: string);
begin
  inherited;
  FTitle := ATitle;
end;

function TGraficoColumnChart.GetFuncaoDraw(AElementID: string; idx: integer): string;
begin
  FObjDados := 'dados' + IntToStr(idx);

  Result := sLineBreak + //
    ' var ' + FObjDados + ' = new google.visualization.arrayToDataTable(chartData'+ IntToStr(idx) +'); ' + sLineBreak +

    Space(16) + ' const options = { title: "' + FTitle + '",' + sLineBreak + //
    Space(20) + ' height: "100%",    ' + sLineBreak + //
    Space(20) + ' theme: "material", ' + sLineBreak + //
    Space(20) + ' animation: {       ' + sLineBreak + //
    Space(20) + '    startup: true,  ' + sLineBreak + //
    Space(20) + '    duration: 800,  ' + sLineBreak + //
    Space(20) + '    easing: "out",  ' + sLineBreak + //
    Space(20) + ' },                 ' + sLineBreak + //
    Space(20) + ' chartArea:{        ' + sLineBreak + //
    Space(20) + '    left:70,        ' + sLineBreak + //
    Space(20) + '    right:70,       ' + sLineBreak + //
    Space(20) + '    bottom:80,      ' + sLineBreak + //
    Space(20) + '    top:30          ' + sLineBreak + //
    Space(20) + '  }                 ' + sLineBreak + //
    Space(16) + '};                  ' + sLineBreak + sLineBreak + //

    Space(16) + AElementID + ' = new google.visualization.ColumnChart(document.getElementById(''' + AElementID +
    ''')); ' + sLineBreak + sLineBreak;

  Result := Result + Space(16) + RetornaCallBack(AElementID, idx);
  Result := Result + Space(16) + sLineBreak + AElementID + '.draw(' + FObjDados + ', options); ' + sLineBreak + sLineBreak;
end;

function TGraficoColumnChart.GetUpdateData(Idx: Integer): string;
begin
  Result := 'chartData'+ IntToStr(Idx) +' = [ ' + sLineBreak + //
    '            [ ' +  RetornaDadosDataTable() + '];' + sLineBreak;
end;

function TGraficoColumnChart.package: string;
begin
  Result := 'corechart';
end;

end.
