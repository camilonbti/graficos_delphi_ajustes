unit uFiltroGrafico;

interface

uses
  System.SysUtils, System.Classes, Data.DB, System.StrUtils,
  System.Generics.Collections;

type
  TFiltroChangeEvent = procedure(Sender: TObject) of object;

  TFiltroItem = class
  private
    FCampo: string;
    FOperador: string;
    FValor: string;
    FValor2: string;
    FDescricao: string;
  public
    property Campo: string read FCampo write FCampo;
    property Operador: string read FOperador write FOperador;
    property Valor: string read FValor write FValor;
    property Valor2: string read FValor2 write FValor2;
    property Descricao: string read FDescricao write FDescricao;
  end;

  TFiltroGrafico = class
  private
    FFiltros: TObjectList<TFiltroItem>;
    FOnChange: TFiltroChangeEvent;

    procedure DoChange;
  public
    constructor Create;
    destructor Destroy; override;

    procedure AdicionarFiltro(const ACampo, AOperador, AValor, ADescricao: string); overload;
    procedure AdicionarFiltro(const ACampo, AOperador, AValor, AValor2, ADescricao: string); overload;
    function RemoverFiltro(const ACampo: string): Boolean;
    procedure LimparFiltros;
    function ObterFiltroSQL: string;

    property OnChange: TFiltroChangeEvent read FOnChange write FOnChange;
  end;

implementation

constructor TFiltroGrafico.Create;
begin
  FFiltros := TObjectList<TFiltroItem>.Create(True);
end;

destructor TFiltroGrafico.Destroy;
begin
  FFiltros.Free;
  inherited;
end;

procedure TFiltroGrafico.DoChange;
begin
  if Assigned(FOnChange) then
    FOnChange(Self);
end;

procedure TFiltroGrafico.AdicionarFiltro(const ACampo, AOperador, AValor, ADescricao: string);
var
  Item: TFiltroItem;
begin
  // Remover filtro existente com o mesmo campo
  RemoverFiltro(ACampo);

  Item := TFiltroItem.Create;
  Item.Campo := ACampo;
  Item.Operador := AOperador;
  Item.Valor := AValor;
  Item.Descricao := ADescricao;

  FFiltros.Add(Item);
  DoChange;
end;

procedure TFiltroGrafico.AdicionarFiltro(const ACampo, AOperador, AValor, AValor2, ADescricao: string);
var
  Item: TFiltroItem;
begin
  // Remover filtro existente com o mesmo campo
  RemoverFiltro(ACampo);

  Item := TFiltroItem.Create;
  Item.Campo := ACampo;
  Item.Operador := AOperador;
  Item.Valor := AValor;
  Item.Valor2 := AValor2;
  Item.Descricao := ADescricao;

  FFiltros.Add(Item);
  DoChange;
end;

function TFiltroGrafico.RemoverFiltro(const ACampo: string): Boolean;
var
  I: Integer;
begin
  Result := False;

  for I := FFiltros.Count - 1 downto 0 do
  begin
    if SameText(FFiltros[I].Campo, ACampo) then
    begin
      FFiltros.Delete(I);
      Result := True;
    end;
  end;

  if Result then
    DoChange;
end;

procedure TFiltroGrafico.LimparFiltros;
begin
  if FFiltros.Count > 0 then
  begin
    FFiltros.Clear;
    DoChange;
  end;
end;

function TFiltroGrafico.ObterFiltroSQL: string;
var
  I: Integer;
  Filtro: TFiltroItem;
  ValorStr, Valor2Str: string;
begin
  Result := '';

  for I := 0 to FFiltros.Count - 1 do
  begin
    Filtro := FFiltros[I];

    // Formatar valores conforme o operador
    if UpperCase(Filtro.Operador) = 'LIKE' then
      ValorStr := QuotedStr(Filtro.Valor + '%')
    else
      ValorStr := QuotedStr(Filtro.Valor);

    if Filtro.Valor2 <> '' then
      Valor2Str := QuotedStr(Filtro.Valor2);

    // Construir cláusula SQL
    if UpperCase(Filtro.Operador) = 'BETWEEN' then
      Result := Result + ' AND ' + Filtro.Campo + ' BETWEEN ' + ValorStr + ' AND ' + Valor2Str
    else
      Result := Result + ' AND ' + Filtro.Campo + ' ' + Filtro.Operador + ' ' + ValorStr;
  end;
end;

end.
