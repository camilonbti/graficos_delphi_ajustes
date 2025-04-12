unit uGraficoPie;

interface

uses
  uGraficoGoogle, System.StrUtils;

type
  TGraficoPie = class(TGraficoGoogle)
  private
    FTitle: string;
    FPieDonat: Boolean;
  public
    constructor Create(PieDonat: Boolean = False); overload;
    destructor Destroy; override;

    function package : string; override;

    function GetFuncaoDraw(AElementID : string; idx : integer) : string; override;

    procedure AddTitle(ATitle: string); override; //Added
  end;

implementation

uses SysUtils, uGraficos;

{ TGraficoPie }

constructor TGraficoPie.Create(PieDonat: Boolean = False);
begin
  inherited Create;
  FPieDonat := PieDonat;
end;

destructor TGraficoPie.Destroy;
begin
  inherited;
end;

procedure TGraficoPie.AddTitle(ATitle: string);
begin
  inherited;
  FTitle := ATitle;
end;

function TGraficoPie.GetFuncaoDraw(AElementID: string; idx : integer): string;
begin
  FObjDados := 'dados' + IntToStr(idx);

  Result :=
       //Added
       RetornaDadosDataTable + //
       Space(16) + ' var options = { title: "' + FTitle + '",' + sLineBreak + //
       Space(20) + ' height: "100%",    ' + sLineBreak + //
       Space(20) + ' theme: "material", ' + sLineBreak + //
       ifThen(FPieDonat, 'pieHole: 0.4,', '') + sLineBreak + //
       Space(20) + ' chartArea:{      '+ sLineBreak + //
       Space(20) + '    left:5,       '+ sLineBreak + //
       Space(20) + '    right:5,      '+ sLineBreak + //
       Space(20) + '    bottom:3,     '+ sLineBreak + //
       Space(20) + '    top:30        '+ sLineBreak + //
       Space(20) + '  }               '+ sLineBreak + //
       Space(16) + '};                '+ sLineBreak + sLineBreak + //

       Space(16) + AElementID + ' = new google.visualization.PieChart(document.getElementById(''' + AElementID + ''')); '+sLineBreak + sLineBreak;

  Result := Result + RetornaCallBack(AElementID, idx);
  Result := Result + AElementID + '.draw(' + FObjDados + ', options); ';
end;

function TGraficoPie.package: string;
begin
  Result := 'corechart';
end;

end.
