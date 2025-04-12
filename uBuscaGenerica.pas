unit uBuscaGenerica;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Controls, Vcl.Forms, StdCtrls, Vcl.ComCtrls,  cxGraphics, cxControls, cxLookAndFeels,
  cxLookAndFeelPainters, cxContainer, cxEdit, dxSkinsCore,  dxSkinsDefaultPainters, cxCheckBox,
  System.StrUtils;

const
  WM_KILLCONTROL   = WM_USER + 1;
  CONST_PREFIX_CHK = 'chf_';
type
  TLastChanger = (lcAddFilter, lcDelFilter);

  TChangerType = (ctAuto, ctManual);
  TNotifyChanger = procedure(Sender: TObject; ChangerType: TChangerType) of object;

  TFilter = class
    tabela   : string;
    campo    : string;
    campo2   : string;
    valor    : string;
    valor2   : string;
    caption  : string;
    operador : string;
    operador2: string;
  end;

  TBuscaGenerica = class(TWinControl)
  private
    fFiltros           : TStringList;
    fGroupCheckBox     : TToolBar;
    fPnlFiltroPesquisa : TScrollBox;
    fOnChangeFilter    : TNotifyChanger;
    FLastChange        : TLastChanger;

    procedure KillControl(var message: TMessage); message WM_KILLCONTROL;

    function getItemByName(name : String) : Integer;
    function getFiltroObjectByName(compName: String): TFilter;

    procedure removeFilter(name: string);
    procedure chkClick(Sender: TObject);
    procedure ChangeFilter(AType: TChangerType);
    procedure setPnlFiltroPesquisa(const Value: TScrollBox);

  public
    function getSql(CurrentSql: string = ''): String;

    function clearFilter(ArrName: array of string): Boolean; overload;
    function clearFilter(AName: string = ''): Boolean; overload;

    procedure addFilter( ptabela, pcampo, poperador, pvalor, pcampo2, poperador2, pvalor2, pdecode, pcaption : String ); overload;
    procedure addFilter( ptabela, pcampo, poperador, pvalor, poperador2, pvalor2, pdecode, pcaption : String ); overload;
    procedure addFilter( ptabela, pcampo, poperador, pvalor, pvalor2, pdecode, pcaption : String ); overload;
    procedure addFilter( ptabela, pcampo, poperador, pvalor, pvalor2, pcaption : String ); overload;
    procedure addFilter( pcampo, poperador, pvalor, pcaption : String); overload;

    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  published
    property PnlFiltroPesquisa : TScrollBox read fPnlFiltroPesquisa write setPnlFiltroPesquisa;
    property LastChange: TLastChanger read FLastChange write FLastChange;
    property OnChangeFilter : TNotifyChanger read fOnChangeFilter write fOnChangeFilter;
  end;

implementation

function TBuscaGenerica.clearFilter(ArrName: array of string): Boolean;
begin
  if Length(ArrName) <= 0 then
    exit;

  for var I: Integer := Low(ArrName) to High(ArrName) do
    clearFilter(ArrName[I]);
end;

function TBuscaGenerica.clearFilter(AName: string = ''): Boolean;
var
  I: Integer;
  J: Boolean;
  CheckBox: TcxCheckBox;
begin
  Result := False;

  for I := fGroupCheckBox.ComponentCount - 1 downto 0 do
  begin
    if fGroupCheckBox.Components[i] is TcxCheckBox then
    begin
      CheckBox := fGroupCheckBox.Components[i] as TcxCheckBox;

      J := True;  // Inicialmente, assumimos que o componente é um TcxCheckBox

      if not Trim(AName).IsEmpty then
        J := J and (UpperCase(CheckBox.Name) = 'CHF_' + UpperCase(AName));

      if J then
      begin
        LastChange := TLastChanger.lcDelFilter;
        CheckBox.Checked := False;

        // Verificar se o componente é válido antes de chamar a função removeFilter
        if Assigned(CheckBox) then
        begin
          removeFilter(CheckBox.Name);
          CheckBox.OnClick := nil;
          fGroupCheckBox.Width := fGroupCheckBox.Width - 15 + Length(CheckBox.Caption) * 8;
          //PostMessage(Self.Handle, WM_KILLCONTROL, 0, integer(CheckBox));
        end;

        Result := True;
        if not Trim(AName).IsEmpty then
          Break;
      end;
    end;
  end;

//var
//  I: Integer;
//  J: Boolean;
//begin
//  Result := False;
//
//  for I := 0 to fGroupCheckBox.ComponentCount -1 do
//  begin
//    J := fGroupCheckBox.Components[i] is TcxCheckBox;
//
//    if not Trim(AName).IsEmpty then
//      J := J and (UpperCase((fGroupCheckBox.Components[i] as TcxCheckBox).Name) = 'CHKFILTRO_' + UpperCase(AName));
//
//    if J then
//    begin
//      LastChange := TLastChanger.lcDelFilter;
//      (fGroupCheckBox.Components[i] as TcxCheckBox).Checked := False;
//
//      removeFilter((fGroupCheckBox.Components[i] as TcxCheckBox).Name);
//      (fGroupCheckBox.Components[i] as TcxCheckBox).OnClick := nil;
//
//      fGroupCheckBox.Width := fGroupCheckBox.Width - //
//                              15 + Length((fGroupCheckBox.Components[i] as TcxCheckBox).Caption) * 8;
//
//      //PostMessage(Self.Handle, WM_KILLCONTROL, 0, integer((fGroupCheckBox.Components[i] as TcxCheckBox)));
//
//      Result := True;
//      if not Trim(AName).IsEmpty then
//        break;
//    end;
//  end;
end;

procedure TBuscaGenerica.addFilter( ptabela, pcampo, poperador, pvalor, pcampo2, poperador2, pvalor2, pdecode, pcaption : String );
var
  nitens   : Integer;
  filtro   : TFilter;
  compName : string;
  item     : TcxCheckBox;
  chngtp   : TChangerType;
begin
  compName := CONST_PREFIX_CHK + pcampo;

  filtro := getFiltroObjectByName(compName);
  if Not(Assigned(filtro)) then
  begin
    filtro := TFilter.Create;
    fFiltros.AddObject(pcampo, filtro);
  end;

  filtro.tabela    := ptabela;
  filtro.campo     := pcampo;
  filtro.campo2    := pcampo2;
  filtro.valor     := pvalor;
  filtro.valor2    := pvalor2;
  filtro.caption   := pcaption;
  filtro.operador  := poperador;
  filtro.operador2 := poperador2;

  // Cria o checkbox
  item := TcxCheckBox(fGroupCheckBox.FindComponent(compName));
  if not(Assigned(item)) then
  begin
    item :=  TcxCheckBox.Create(fGroupCheckBox);
    chngtp := TChangerType.ctAuto;
  end else
  begin
    fGroupCheckBox.Width := fGroupCheckBox.Width - item.Width;
    chngtp := TChangerType.ctManual;
  end;

  item.Parent      := fGroupCheckBox;
  item.ParentColor := False;
  item.Transparent := True;
  item.Visible     := True;

  item.Caption          := pcaption + ': ' + //
                           IfThen(not Trim(pdecode).isEmpty,
                                  pdecode,
                                  pvalor + IfThen(not Trim(pvalor2).isEmpty, ' à ' + pvalor2, ''));
  item.Checked          := True;
  item.OnClick          := chkClick;
  item.Name             := compName;
  item.Width            := 15 + (Length(item.Caption) * 8);
  item.Style.Font.Name  := 'Courier New';

  fGroupCheckBox.Width := fGroupCheckBox.Width + item.Width;

  LastChange := TLastChanger.lcAddFilter;
  ChangeFilter(chngtp);
end;

procedure TBuscaGenerica.addFilter( ptabela, pcampo, poperador, pvalor, poperador2, pvalor2, pdecode, pcaption : String );
begin
  addFilter(ptabela, pcampo, poperador, pvalor, '', '', pvalor2, '', pcaption);
end;

procedure TBuscaGenerica.addFilter( ptabela, pcampo, poperador, pvalor, pvalor2, pdecode, pcaption : String );
begin
  addFilter(ptabela, pcampo, poperador, pvalor, '', pvalor2, '', pcaption);
end;

procedure TBuscaGenerica.addFilter( ptabela, pcampo, poperador, pvalor, pvalor2, pcaption : String );
begin
  addFilter(ptabela, pcampo, poperador, pvalor, pvalor2, '', pcaption);
end;

procedure TBuscaGenerica.addFilter( pcampo, poperador, pvalor, pcaption : String );
begin
  addFilter('', pcampo, poperador, pvalor, '', pcaption);
end;

procedure TBuscaGenerica.removeFilter(name : string);
var
  idx : Integer;
begin
  idx := getItemByName(name);
  if (idx >= 0) then
    fFiltros.Delete(idx);
end;

procedure TBuscaGenerica.chkClick(Sender: TObject);
var
  CheckBox : TcxCheckBox;
begin
  // Verificar se o Sender é um objeto válido
  if not Assigned(Sender) then
    Exit;  // Sair da função se não for

  // Verificar se o Sender é um TcxCheckBox válido
  if not (Sender is TcxCheckBox) then
    Exit;

  CheckBox := TcxCheckBox(Sender);
  if not(CheckBox.Checked) then
  begin
    removeFilter(CheckBox.name);
    CheckBox.OnClick := nil;

    fGroupCheckBox.Width := fGroupCheckBox.Width - CheckBox.Width;
    LastChange           := TLastChanger.lcDelFilter;
    ChangeFilter(TChangerType.ctManual);

    PostMessage(Self.Handle, WM_KILLCONTROL, 0, integer(Sender));
  end;



//  CheckBox := TcxCheckBox(Sender);
//  if not(CheckBox.Checked) then
//  begin
//    removeFilter(CheckBox.name);
//    CheckBox.OnClick := nil;
//
//    fGroupCheckBox.Width := fGroupCheckBox.Width - CheckBox.Width;
//    LastChange           := TLastChanger.lcDelFilter;
//    ChangeFilter(TChangerType.ctManual);
//
//    PostMessage(Self.Handle, WM_KILLCONTROL, 0, integer(Sender));
//  end;
end;

function TBuscaGenerica.getItemByName(name: String): Integer;
var
  campo : string;
begin
  campo := StringReplace(name, CONST_PREFIX_CHK, '', [rfReplaceAll]);
  Result := fFiltros.IndexOf(campo);
end;

procedure TBuscaGenerica.KillControl(var message: TMessage);
var
  control: TControl;
begin
  control := TObject(message.LParam) as TControl;
  FreeAndNil(control);
end;

function TBuscaGenerica.getFiltroObjectByName(compName : String) : TFilter;
var
  idx : Integer;
begin
  idx := getItemByName(compName);
  if (idx < 0 ) then
  begin
    Result := nil;
    Exit;
  end;

  Result := TFilter(fFiltros.Objects[idx]);
end;

destructor TBuscaGenerica.Destroy;
begin
  fFiltros.Free;
  inherited;
end;

constructor TBuscaGenerica.Create(AOwner: TComponent);
begin
  inherited;
  self.Parent := TForm(AOwner);
  fFiltros := TStringList.Create;

  fGroupCheckBox := TToolBar.Create(self);
end;

function TBuscaGenerica.getSql(CurrentSql: string = '') : String;
var
  i      : integer;
  filter : tfilter;
  sql, vValor, vValor2 : string;
begin
  for i := 0 to fFiltros.Count -1 do
  begin
    filter := TFilter( fFiltros.Objects[i] );

    vValor  := filter.valor;
    vValor2 := filter.valor2;

//    if UpperCase(filter.operador) = 'LIKE' then
//      vValor := '%'+vValor+'%';
//
//    if (UpperCase(filter.operador) = 'LIKE') and (UpperCase(filter.operador2) = 'OR') then
//    begin
//      vValor2 := '%'+vValor2+'%';
//      vValor2 := QuotedStr(vValor2);
//    end;

    if UpperCase(filter.operador) = 'LIKE' then
      vValor := vValor+'%';

    if (UpperCase(filter.operador) = 'LIKE') and (UpperCase(filter.operador2) = 'OR') then
    begin
      vValor2 := vValor2+'%';
      vValor2 := QuotedStr(vValor2);
    end;

    vValor := QuotedStr(vValor);

    if not Trim(filter.tabela).isEmpty and not (UpperCase(filter.campo).Contains(filter.tabela +'.')) then
      filter.campo := UpperCase(filter.tabela) +'.'+ filter.campo;

    if UpperCase(filter.operador) = 'BETWEEN' then
    begin
      vValor2 := QuotedStr(vValor2);

      sql := sql  + #13#10 +' and ' + filter.campo + ' ' + filter.operador + ' ' + vValor + ' and ' + vValor2;
    end else
    if UpperCase(filter.operador2) = 'OR' then
    begin
      // se houver operador2 OR, verificar o campo2 para preencher
      if not Trim(filter.campo2).isEmpty then
        sql := sql  + #13#10 +' and ((' + filter.campo +' '+ filter.operador +' '+
                                vValor + ') '+ filter.operador2 +' (' + filter.campo2 +' '+ filter.operador +' '+ vValor2 + ')) '
      else
        sql := sql  + #13#10 +' and ((' + filter.campo + ' ' + filter.operador + ' ' +
                                vValor + ') '+ filter.operador2 +' (' + filter.campo + ' ' + vValor2 + ')) '
    end else
      sql := sql  + #13#10 +' and ' + filter.campo + ' ' + filter.operador + ' ' + vValor + ' ';
  end;

  Result := sql;
end;


procedure TBuscaGenerica.ChangeFilter(AType: TChangerType);
begin
  if Assigned(fOnChangeFilter) then
    fOnChangeFilter(self, AType);
end;

procedure TBuscaGenerica.setPnlFiltroPesquisa(const Value: TScrollBox);
begin
  fPnlFiltroPesquisa := Value;

  fPnlFiltroPesquisa.Height   := 45;
  with fGroupCheckBox do
  begin
    Parent       := fPnlFiltroPesquisa;
    EdgeBorders  := [];
    Align        := alNone;
    Flat         := True;
    Width        := 0;
    Top          := 0;
    Left         := 8;
  end;

  fGroupCheckBox.Height := fPnlFiltroPesquisa.Height;
end;

end.
