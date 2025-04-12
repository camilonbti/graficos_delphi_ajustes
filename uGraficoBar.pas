unit uGraficoBar;

interface

uses
  uGraficoGoogle, System.StrUtils;

type
  TGraficoBar = class(TGraficoGoogle)
  private
    FTitle: string;
    FHorizontal: Boolean;
  public
    constructor Create(Horizontal: Boolean = False); overload;
    destructor Destroy; override;

    function package : string; override;

    function GetFuncaoDraw(AElementID : string; idx : integer) : string; override;

    procedure AddTitle(ATitle: string); override;
  end;

implementation

uses SysUtils, uGraficos;

{ TGraficoBar }

constructor TGraficoBar.Create(Horizontal: Boolean = False);
begin
  inherited Create;
  FHorizontal := Horizontal;
end;

destructor TGraficoBar.Destroy;
begin
  inherited;
end;

procedure TGraficoBar.AddTitle(ATitle: string);
begin
  inherited;
  FTitle := ATitle;
end;

function TGraficoBar.GetFuncaoDraw(AElementID: string; idx : integer): string;
begin
  FObjDados := 'dados' + IntToStr(idx);

  Result := //
       RetornaDados + //
       Space(16) + 'var options = { title: "'+ FTitle +'", responsive: true, '+  //
       ifThen(FHorizontal, 'bars: horizontal, ', '') + sLineBreak + //
       Space(20) + ' chartArea:{      '+ sLineBreak + //
       Space(20) + '    left:5,       '+ sLineBreak + //
       Space(20) + '    right:5,      '+ sLineBreak + //
       Space(20) + '    bottom:3,     '+ sLineBreak + //
       Space(20) + '    top:30,       '+ sLineBreak + //
       Space(20) + '    width:"100%", '+ sLineBreak + //
       Space(20) + '    height:"100%" '+ sLineBreak + //
       Space(20) + '  }               '+ sLineBreak + //
       Space(16) + '};                '+ sLineBreak + sLineBreak + //

       Space(16) + 'var ' + AElementID + ' = new google.charts.Bar(document.getElementById(''' + AElementID + ''')); '+sLineBreak +
       Space(16) + 'var ' + AElementID + '.draw(' + FObjDados + ', google.charts.Bar.convertOptions(options));  ' + sLineBreak + sLineBreak;

  Result := Result + RetornaCallBack(AElementID, idx);

{
	   var view = new google.visualization.DataView(data);
	   var chart = new google.visualization.ColumnChart(document.getElementById("columnchart_values"));
	   chart.draw(view, options);
}
end;

function TGraficoBar.package: string;
begin
  Result := 'bar';
end;

end.
