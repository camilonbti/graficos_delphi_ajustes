unit uGraficoGoogle;

interface

uses System.Classes, System.Variants;

type
  TGoogleCallBack = procedure (Params : TStringList) of object;

  TTipoCampoVirtual = (tcString, tcNumerico, tcDate); //Added

  TGraficoGoogle = class abstract
  private
    FOnCallBack: TGoogleCallBack;
    FFieldCallBack: string;
    FDescricaoColunaX: string;
    FDescricaoColunaY: string;

    FColunaVirtual : array of string;
    FColunaVirtualTipo: array of TTipoCampoVirtual; //Added
    FDados : array of array of Variant;

    procedure SetOnCallBack(const Value: TGoogleCallBack);
    procedure SetDescricaoColunaX(const Value: string);
    procedure SetDescricaoColunaY(const Value: string);
  protected
    FObjDados : string;

    function RetornaCallBack(AElementID : string; idx: integer) : string;
    function RetornaDados : string;
    function RetornaDadosDataTable(): string; //Added
  public
    FRandomColor: Boolean; //Added
    procedure AfterConstruction; override;

    function package : string; virtual; abstract;

    function GetFuncaoDraw(AElementID : string; idx : integer) : string; virtual; abstract;
    function GetUpdateData(Idx: Integer): string; virtual; abstract;

    procedure AddMargem(ALeft, ARight, ATop, ABottom: Integer); virtual; abstract; //Added
    procedure AddTitle(ATitle: string); virtual; abstract; //Added

    procedure AddDadosColunaVirtual(ARole : string; ATipo: TTipoCampoVirtual = TTipoCampoVirtual.tcString);
    function Reset(): TGraficoGoogle;
    procedure AddDados(ADados : array of Variant);
    procedure LimpaDados;

    property OnCallBack : TGoogleCallBack read FOnCallBack write SetOnCallBack;
    property FieldCallBack : string read FFieldCallBack write FFieldCallBack;
    property DescricaoColunaX : string read FDescricaoColunaX write SetDescricaoColunaX;
    property DescricaoColunaY : string read FDescricaoColunaY write SetDescricaoColunaY;
{
    property Altura;
    property Largura;
    }
  end;

implementation

uses System.SysUtils, System.StrUtils;

{ TGraficoGoogle }

// Added
function TGraficoGoogle.Reset(): TGraficoGoogle;
begin
  Result := Self;
  SetLength(FDados, 0);
end;

procedure TGraficoGoogle.AddDados(ADados: array of Variant);
var
  Row, Cols, i : integer;
begin
  Cols := Length(FColunaVirtual) + 5;

  if Length(ADados) <> Cols then
    raise Exception.Create('Número de colunas inválidas!');

  Row := Length(FDados);
  SetLength(FDados, Row + 1);


  SetLength(FDados[Row], Cols);

  for i := 0 to Pred(Cols) do
    FDados[Row, i] := ADados[i];
end;

procedure TGraficoGoogle.AddDadosColunaVirtual(ARole: string; ATipo: TTipoCampoVirtual = TTipoCampoVirtual.tcString);
begin
  if Length(FDados) > 0 then
    raise Exception.Create('Dados já preenchidos, não pode incluir coluna virtual!');

  SetLength(FColunaVirtual, Length(FColunaVirtual) + 1);
  FColunaVirtual[Pred(Length(FColunaVirtual))] := ARole;

  //Added
  SetLength(FColunaVirtualTipo, Length(FColunaVirtualTipo) + 1);
  FColunaVirtualTipo[Pred(Length(FColunaVirtualTipo))] := ATipo;
end;

procedure TGraficoGoogle.AfterConstruction;
begin
  inherited;

  FOnCallBack := nil;
end;

procedure TGraficoGoogle.LimpaDados;
begin
  SetLength(FDados, 0);
  SetLength(FColunaVirtual, 0); //Added
  SetLength(FColunaVirtualTipo, 0); //Added
end;

function TGraficoGoogle.RetornaCallBack(AElementID : string; idx: integer): string;
begin
  Result := '';

  if Assigned(OnCallBack) then
  begin
    Result :=
      '                google.visualization.events.addListener(' + AElementID + ', ''select'', selectHandler' + AElementID + '); ' + sLineBreak + sLineBreak +

      '                function selectHandler' + AElementID + '() {' + sLineBreak +
      '                   var item = ' + AElementID + '.getSelection()[0]; '+ sLineBreak +
      '                   var retorno = [];                          '+ sLineBreak +
      '                   for(i=0;i<'+FObjDados+'.getNumberOfColumns();i++) '+ sLineBreak +
      '                     retorno.push('+FObjDados+'.getValue(item.row, i));'+ sLineBreak + sLineBreak +

      '                     var url = "ActionCallBackJS:Retorno' + IntToStr(idx) + '("+JSON.stringify(retorno)+")"; '+ sLineBreak +
      '                     location.assign(url); '+ sLineBreak +
      '                } ';
  end;
end;

function TGraficoGoogle.RetornaDados: string;
var
  I: Integer;
  j: Integer;
  aux : Variant;
begin
  Result :=
       '                    '+ FObjDados + ' = new google.visualization.DataTable(); ' +  sLineBreak + //
       '                    '+ FObjDados + '.addColumn(''string'', "' + FDescricaoColunaX + '"); ' +  sLineBreak + //
       '                    '+ FObjDados + '.addColumn(''number'', "' + FDescricaoColunaY + '"); '+ sLineBreak ; //

  for I := 0 to High(FColunaVirtual) do
  begin
    case FColunaVirtualTipo[I] of
      tcString:
        Result := Result +
       '                    '+ FObjDados + '.addColumn({type: "string", role:"' + FColunaVirtual[i] + '"});' + sLineBreak ; //

      //Added
      tcNumerico:
        Result := Result + '                    '+ FObjDados + '.addColumn({type: "number", role:"' + FColunaVirtual[i] + '"});' + sLineBreak ; //
      tcDate:
        Result := Result + '                    '+ FObjDados + '.addColumn({type: "date", role:"' + FColunaVirtual[i] + '"});' + sLineBreak ; //
    end;
  end;

  Result := Result + '                    ' + FObjDados + '.addRows([ ' + sLineBreak ; //
  for I := 0 to High(FDados) do
  begin
   Result := Result + ifthen(i > 0, ',', '') + sLineBreak +  '                        [';

   for j := 0 to High(FDados[i]) do
   begin
     aux := FDados[i, j];

     //Added
       Result := Result + varToStr(aux) + ifthen(j < High(FDados[i]), ',', '') ; //

//     Result := Result + ifthen(VarIsStr(aux), '"' + varToStr(aux) + '"', varToStr(aux)) +
//                        ifthen(j < High(FDados[i]), ',', '');
   end;

   Result := Result + ']' + sLineBreak
  end;

  Result := Result +
     '                ]); '+ sLineBreak;
end;

function TGraficoGoogle.RetornaDadosDataTable(): string;
var
  I: Integer;
  j: Integer;
  aux : Variant;
  lField, dado: string;
begin
  lField := '';
  Result := '';

  for I := 0 to High(FColunaVirtual) do
  begin
    Result := Result + ' "'+ FColunaVirtual[i] +'", ';
    lField := lField + ' {role:" '+ FColunaVirtual[i] +'"}, ';
  end;

  Result := Result + sLineBreak + //
    lField + //
    ifThen(not Trim(FFieldCallBack).isEmpty, ' {role:" '+ FFieldCallBack +'"}, ', '') +
    ' {role: "annotation" }, ' + //
    ifThen(FRandomColor, ' {role: "style" } ', '') + //
    ' ], ' + sLineBreak + sLineBreak;

  for I := 0 to High(FDados) do
  begin
    dado := '';

    for j := 0 to High(FDados[i]) do
    begin
      dado := dado + varToStr(FDados[i, j]) + ifthen(j < High(FDados[i]), ', ', '') ;
    end;

    Result := Result + '['+ dado +'], ' + sLineBreak;
  end;

  Result := Result + sLineBreak;
end;

procedure TGraficoGoogle.SetDescricaoColunaX(const Value: string);
begin
  FDescricaoColunaX := Value;
end;

procedure TGraficoGoogle.SetDescricaoColunaY(const Value: string);
begin
  FDescricaoColunaY := Value;
end;

procedure TGraficoGoogle.SetOnCallBack(const Value: TGoogleCallBack);
begin
  FOnCallBack := Value;
end;

end.
