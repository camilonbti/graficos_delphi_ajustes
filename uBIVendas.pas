unit uBIVendas;

interface

uses
  uBuscaGenerica, uFormGrafico,
  //
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, DateUtils,
  System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB,
  Datasnap.DBClient, Vcl.Grids, Vcl.DBGrids, Vcl.StdCtrls, Vcl.ComCtrls, StrUtils,
  Vcl.ExtCtrls, FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  FireDAC.Comp.DataSet, FireDAC.Comp.Client, FireDAC.UI.Intf, FireDAC.Stan.Def,
  FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys, FireDAC.Phys.SQLite,
  FireDAC.Phys.SQLiteDef, FireDAC.Stan.ExprFuncs, FireDAC.VCLUI.Wait,
  FireDAC.Phys.SQLiteVDataSet, FireDAC.DApt, RxCtrls, dxGDIPlusClasses,
  cxGraphics, cxControls, cxLookAndFeels, cxLookAndFeelPainters, cxContainer,
  cxEdit, cxCheckBox, FireDAC.Phys.SQLiteWrapper.Stat, Vcl.OleCtrls, dxflchrt,
  SHDocVw, dxSkinsCore, dxSkinsDefaultPainters,
  Vcl.Menus, System.Actions, Vcl.ActnList;

type
  TfBIVendas = class(TForm)
    dsVendasGeral: TDataSource;
    cdsVendasGeral: TClientDataSet;
    FDMTVendas: TFDMemTable;
    FDConnection1: TFDConnection;
    FDLocalSQL1: TFDLocalSQL;
    FDPhysSQLiteDriverLink1: TFDPhysSQLiteDriverLink;
    PageControl1: TPageControl;
    pgVendasGrid: TTabSheet;
    ScrollBox1: TScrollBox;
    pgVendasCharts: TTabSheet;
    pnlgridRegiao: TPanel;
    gridRegiao: TDBGrid;
    Panel1: TPanel;
    RxLabel7: TRxLabel;
    pnlgridLoja: TPanel;
    gridLoja: TDBGrid;
    Panel16: TPanel;
    RxLabel11: TRxLabel;
    pnlCidades: TPanel;
    pnlgridCidadeClientes: TPanel;
    Panel18: TPanel;
    RxLabel20: TRxLabel;
    gridCidadeClientes: TDBGrid;
    Panel42: TPanel;
    Panel43: TPanel;
    RxLabel21: TRxLabel;
    gridCidadeLojas: TDBGrid;
    pnlgridgrupo: TPanel;
    Panel45: TPanel;
    RxLabel22: TRxLabel;
    gridGrupo: TDBGrid;
    Panel6: TPanel;
    Panel8: TPanel;
    RxLabel52: TRxLabel;
    pnlfornecedor: TPanel;
    Panel49: TPanel;
    RxLabel24: TRxLabel;
    gridFornecedor: TDBGrid;
    pnlMarca: TPanel;
    Panel51: TPanel;
    RxLabel25: TRxLabel;
    gridMarca: TDBGrid;
    pnlvendedor: TPanel;
    Panel53: TPanel;
    RxLabel26: TRxLabel;
    gridVendedor: TDBGrid;
    pnlPublicidade: TPanel;
    Panel56: TPanel;
    Panel55: TPanel;
    RxLabel27: TRxLabel;
    gridPublicidade: TDBGrid;
    Panel57: TPanel;
    Panel58: TPanel;
    RxLabel28: TRxLabel;
    gridDiaSemana: TDBGrid;
    Panel61: TPanel;
    RxLabel30: TRxLabel;
    gridABC: TDBGrid;
    Panel59: TPanel;
    Panel60: TPanel;
    RxLabel29: TRxLabel;
    gridHora: TDBGrid;
    pnlprodutos: TPanel;
    Panel63: TPanel;
    RxLabel31: TRxLabel;
    gridProdutos: TDBGrid;
    pnlConta: TPanel;
    Panel66: TPanel;
    Panel65: TPanel;
    RxLabel32: TRxLabel;
    gridFinanConta: TDBGrid;
    Panel67: TPanel;
    Panel68: TPanel;
    RxLabel33: TRxLabel;
    gridFinanHistorico: TDBGrid;
    Panel72: TPanel;
    Panel73: TPanel;
    RxLabel35: TRxLabel;
    gridFinanPlanoVenda: TDBGrid;
    gridSubGrupo: TDBGrid;
    pnlTabelaPreco: TPanel;
    gridFinanTabelaPreco: TDBGrid;
    Panel71: TPanel;
    RxLabel34: TRxLabel;
    Image1: TImage;
    Image2: TImage;
    Image3: TImage;
    Image4: TImage;
    Image5: TImage;
    Image6: TImage;
    Image7: TImage;
    Image8: TImage;
    Image9: TImage;
    Image10: TImage;
    Image11: TImage;
    Image12: TImage;
    Image13: TImage;
    Image14: TImage;
    Image15: TImage;
    Image16: TImage;
    Image17: TImage;
    Image18: TImage;
    Timer1: TTimer;
    ActionMenu: TActionList;
    actRegional: TAction;
    actLoja: TAction;
    actCidadeCli: TAction;
    actCidadeLj: TAction;
    actGrupo: TAction;
    actSubgrupo: TAction;
    actVendedor: TAction;
    actMarca: TAction;
    actFornecedor: TAction;
    actProduto: TAction;
    PopupMenu: TPopupMenu;
    Regional1: TMenuItem;
    Loja1: TMenuItem;
    CidadeCli1: TMenuItem;
    CidadeLj1: TMenuItem;
    Grupo1: TMenuItem;
    Subgrupo1: TMenuItem;
    Vendedor1: TMenuItem;
    Marca1: TMenuItem;
    Fornecedor1: TMenuItem;
    Produto1: TMenuItem;
    FScrollBox: TScrollBox;
    pnlTopoGeral: TPanel;
    pnlTopoDetail2: TPanel;
    Panel25: TPanel;
    lblTOTAL_QTD_ID_VENDEDOR: TRxLabel;
    Panel26: TPanel;
    RxLabel12: TRxLabel;
    Panel27: TPanel;
    lblTOTAL_QTD_ID_FORNECEDOR: TRxLabel;
    Panel28: TPanel;
    RxLabel13: TRxLabel;
    Panel29: TPanel;
    lblTOTAL_QTD_ID_GRUPO: TRxLabel;
    Panel30: TPanel;
    RxLabel14: TRxLabel;
    Panel31: TPanel;
    lblTOTAL_QTD_ID_MARCA: TRxLabel;
    Panel32: TPanel;
    RxLabel15: TRxLabel;
    Panel33: TPanel;
    lblTOTAL_QTD_ID_CLIENTE: TRxLabel;
    Panel34: TPanel;
    RxLabel16: TRxLabel;
    Panel21: TPanel;
    lblTOTAL_QTD_ID_CIDADE_LOJA: TRxLabel;
    Panel22: TPanel;
    RxLabel9: TRxLabel;
    Panel23: TPanel;
    lblTOTAL_QTD_ID_CIDADE_CLIENTE: TRxLabel;
    Panel24: TPanel;
    RxLabel10: TRxLabel;
    Panel10: TPanel;
    lblTOTAL_QTD_PRODUTOS_QUANTIDADE: TRxLabel;
    Panel11: TPanel;
    RxLabel4: TRxLabel;
    Panel19: TPanel;
    lblTOTAL_QTD_PRODUTOS_DISTINTOS: TRxLabel;
    Panel20: TPanel;
    RxLabel8: TRxLabel;
    pnlPesquisa: TPanel;
    Panel14: TPanel;
    RxLabel6: TRxLabel;
    ScrollBox: TScrollBox;
    btnMenu: TButton;
    pnlTopoDetail1: TPanel;
    Panel2: TPanel;
    btnPesquisar: TImage;
    Panel3: TPanel;
    RxLabel2: TRxLabel;
    chkAtualizarAutomatico: TcxCheckBox;
    dpkDataIni: TDateTimePicker;
    dpkDataFim: TDateTimePicker;
    Panel4: TPanel;
    lblTOTAL_VENDAS: TRxLabel;
    Panel5: TPanel;
    lblCaptionVendas: TRxLabel;
    Panel79: TPanel;
    lblTOTAL_CMV: TRxLabel;
    Panel80: TPanel;
    RxLabel39: TRxLabel;
    Panel81: TPanel;
    lblLUCRO_BRUTO_RS: TRxLabel;
    Panel82: TPanel;
    RxLabel41: TRxLabel;
    Panel83: TPanel;
    lblLUCRO_BRUTO_PCT: TRxLabel;
    Panel84: TPanel;
    RxLabel43: TRxLabel;
    Panel85: TPanel;
    lblTOTAL_VALOR_DESPESA_ADM_CX: TRxLabel;
    Panel86: TPanel;
    RxLabel45: TRxLabel;
    Panel87: TPanel;
    lblTOTAL_VALOR_DESPESA_ADM_CX_PCT: TRxLabel;
    Panel88: TPanel;
    RxLabel47: TRxLabel;
    Panel37: TPanel;
    lblVALOR_LUCRO_LIQUIDO: TRxLabel;
    Panel38: TPanel;
    RxLabel18: TRxLabel;
    Panel39: TPanel;
    lblVALOR_LUCRO_LIQUIDO_PCT: TRxLabel;
    Panel40: TPanel;
    RxLabel19: TRxLabel;
    Panel78: TPanel;
    pnlMedVendaVendedor: TPanel;
    lblVENDAS_MED_VENDEDOR: TRxLabel;
    Panel36: TPanel;
    RxLabel17: TRxLabel;
    pnlTKMedio: TPanel;
    lblVENDAS_MED_PEDIDO: TRxLabel;
    Panel13: TPanel;
    RxLabel5: TRxLabel;
    pnlTKMedioProduto: TPanel;
    lblVENDAS_MED_PRODUTO: TRxLabel;
    Panel9: TPanel;
    RxLabel1: TRxLabel;
    pnlpctacrescimo: TPanel;
    lblTOTAL_ACRESCIMO_PCT: TRxLabel;
    Panel77: TPanel;
    RxLabel38: TRxLabel;
    pnldescontopct: TPanel;
    lblTOTAL_DESCONTO_PCT: TRxLabel;
    Panel75: TPanel;
    RxLabel37: TRxLabel;
    pnlAcrescimo: TPanel;
    lblTOTAL_Acrescimo: TRxLabel;
    Panel90: TPanel;
    RxLabel49: TRxLabel;
    pnlDesconto: TPanel;
    lblTOTAL_DESCONTO: TRxLabel;
    Panel92: TPanel;
    RxLabel51: TRxLabel;
    pnlQTDPedidos: TPanel;
    lblTOTAL_QTD_PEDIDOS: TRxLabel;
    Panel7: TPanel;
    RxLabel3: TRxLabel;
    procedure cdsVendasGeralBeforeOpen(DataSet: TDataSet);
    procedure gridRegiaoDblClick(Sender: TObject);
    procedure gridProdutosDblClick(Sender: TObject);
    procedure gridLojaDblClick(Sender: TObject);
    procedure btnPesquisarClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure gridCidadeClientesDblClick(Sender: TObject);
    procedure gridCidadeLojasDblClick(Sender: TObject);
    procedure gridGrupoDblClick(Sender: TObject);
    procedure gridSubGrupoDblClick(Sender: TObject);
    procedure gridFornecedorDblClick(Sender: TObject);
    procedure gridMarcaDblClick(Sender: TObject);
    procedure gridVendedorDblClick(Sender: TObject);
    procedure gridPublicidadeDblClick(Sender: TObject);
    procedure gridDiaSemanaDblClick(Sender: TObject);
    procedure gridHoraDblClick(Sender: TObject);
    procedure gridFinanContaDblClick(Sender: TObject);
    procedure gridFinanHistoricoDblClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure gridABCDblClick(Sender: TObject);
    procedure gridFinanPlanoVendaDblClick(Sender: TObject);
    procedure gridFinanTabelaPrecoDblClick(Sender: TObject);
    procedure gridRegiaoTitleClick(Column: TColumn);
    procedure gridLojaTitleClick(Column: TColumn);
    procedure FormMouseWheelDown(Sender: TObject; Shift: TShiftState; MousePos: TPoint; var Handled: Boolean);
    procedure FormMouseWheelUp(Sender: TObject; Shift: TShiftState; MousePos: TPoint; var Handled: Boolean);
    procedure gridCidadeClientesTitleClick(Column: TColumn);
    procedure gridCidadeLojasTitleClick(Column: TColumn);
    procedure gridGrupoTitleClick(Column: TColumn);
    procedure gridSubGrupoTitleClick(Column: TColumn);
    procedure gridFornecedorTitleClick(Column: TColumn);
    procedure gridMarcaTitleClick(Column: TColumn);
    procedure gridVendedorTitleClick(Column: TColumn);
    procedure gridPublicidadeTitleClick(Column: TColumn);
    procedure gridDiaSemanaTitleClick(Column: TColumn);
    procedure gridABCTitleClick(Column: TColumn);
    procedure gridHoraTitleClick(Column: TColumn);
    procedure gridProdutosTitleClick(Column: TColumn);
    procedure gridFinanContaTitleClick(Column: TColumn);
    procedure gridFinanHistoricoTitleClick(Column: TColumn);
    procedure gridFinanPlanoVendaTitleClick(Column: TColumn);
    procedure gridFinanTabelaPrecoTitleClick(Column: TColumn);
    procedure Image1Click(Sender: TObject);
    procedure Image2Click(Sender: TObject);
    procedure Image3Click(Sender: TObject);
    procedure Image4Click(Sender: TObject);
    procedure Image5Click(Sender: TObject);
    procedure Image6Click(Sender: TObject);
    procedure Image7Click(Sender: TObject);
    procedure Image8Click(Sender: TObject);
    procedure Image9Click(Sender: TObject);
    procedure Image10Click(Sender: TObject);
    procedure Image11Click(Sender: TObject);
    procedure Image13Click(Sender: TObject);
    procedure Image12Click(Sender: TObject);
    procedure Image14Click(Sender: TObject);
    procedure Image15Click(Sender: TObject);
    procedure Image16Click(Sender: TObject);
    procedure Image17Click(Sender: TObject);
    procedure Image18Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure dpkDataIniChange(Sender: TObject);
    procedure dpkDataFimChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure actExecute(Sender: TObject);

    procedure BuscaChangeFilter(Sender: TObject; ChangerType: TChangerType);
    procedure PageControl1Change(Sender: TObject);
  private

    vWhereBI            : string;

    TOTAL_VENDAS, //
    TK_MEDIO_PRODUTO, //
    TK_MED_PEDIDO       : double;

    fBuscaGenerica      : TBuscaGenerica;

    QryTotais           : TFDQuery;
    dsTotais            : TDataSource;

    QryRegiao           : TFDQuery;
    dsRegiao            : TDataSource;

    QryLoja             : TFDQuery;
    dsLoja              : TDataSource;

    QryCidadeClientes   : TFDQuery;
    dsCidadeClientes    : TDataSource;

    dsCidadeLoja        : TDataSource;
    QryCidadeLoja       : TFDQuery;

    QrySubGrupo         : TFDQuery;
    dsSubGrupo          : TDataSource;

    QryGrupo            : TFDQuery;
    dsGrupo             : TDataSource;

    QryFornecedor       : TFDQuery;
    dsFornecedor        : TDataSource;

    QryMarca            : TFDQuery;
    dsMarca             : TDataSource;

    QryVendedor         : TFDQuery;
    dsVendedor          : TDataSource;

    QryPublicidade      : TFDQuery;
    dsPublicidade       : TDataSource;

    QryDiaSemana        : TFDQuery;
    dsDiaSemana         : TDataSource;

    QryHora             : TFDQuery;
    dsHora              : TDataSource;

    QryABC              : TFDQuery;
    dsABC               : TDataSource;

    QryProdutos         : TFDQuery;
    dsProdutos          : TDataSource;

    qryFinanPlanoVenda  : TFDQuery;
    dsFinanPlanoVenda   : TDataSource;

    qryFinanTabelaPreco : TFDQuery;
    dsFinanTabelaPreco  : TDataSource;

    qryFinanConta       : TFDQuery;
    dsFinanConta        : TDataSource;

    qryFinanHistorico   : TFDQuery;
    dsFinanHistorico    : TDataSource;

    procedure CriarQuerys;
    procedure DestruirQuerys;
    procedure ConsVendasGeral;
    procedure FiltroBI;
    procedure FecharDataSet;
    procedure CarregaLabels;
    procedure LimpaFiltros;
    procedure ConfiguraGrid;
    procedure OrdenaGrid(Column: TColumn);
    procedure ControlsCDS(vStatus: Boolean);
    procedure ChecaDataAtlzAuto;

    procedure CallBackGrafico(Params : TStringList);

    procedure DefineIEVersion;
  public
    procedure CallBackRegiao(value: string);
    procedure CallBackLoja(value: string);

  end;

var
  fBIVendas: TfBIVendas;


implementation

uses
  Modulo, uFuncoes, u_Funcoes, Registry, uVariaveis_Globais;

{$R *.dfm}
{ Tfteste }

procedure TfBIVendas.btnPesquisarClick(Sender: TObject);
begin
  ConsVendasGeral;
end;

procedure TfBIVendas.cdsVendasGeralBeforeOpen(DataSet: TDataSet);
begin
  with TClientDataSet(DataSet) do
  begin
    if Trim(ProviderName) = '' then
    begin
      RemoteServer := dm.SocketConnection;
      ProviderName := dm.SocketConnection.AppServer.Criar_DSP;
    end;
  end;
end;

procedure TfBIVendas.CriarQuerys;
begin
  QryTotais                       := TFDQuery.Create(nil);
  QryTotais.Connection            := FDConnection1;
  dsTotais                        := TDataSource.Create(nil);
  dsTotais.DataSet                := QryTotais;

  QryRegiao                       := TFDQuery.Create(nil);
  QryRegiao.Connection            := FDConnection1;
  dsRegiao                        := TDataSource.Create(nil);
  dsRegiao.DataSet                := QryRegiao;
  gridRegiao.DataSource           := dsRegiao;

  QryLoja                         := TFDQuery.Create(nil);
  QryLoja.Connection              := FDConnection1;
  dsLoja                          := TDataSource.Create(nil);
  dsLoja.DataSet                  := QryLoja;
  gridLoja.DataSource             := dsLoja;

  QryCidadeClientes               := TFDQuery.Create(nil);
  QryCidadeClientes.Connection    := FDConnection1;
  dsCidadeClientes                := TDataSource.Create(nil);
  dsCidadeClientes.DataSet        := QryCidadeClientes;
  gridCidadeClientes.DataSource   := dsCidadeClientes;

  QryCidadeLoja                   := TFDQuery.Create(nil);
  QryCidadeLoja.Connection        := FDConnection1;
  dsCidadeLoja                    := TDataSource.Create(nil);
  dsCidadeLoja.DataSet            := QryCidadeLoja;
  gridCidadeLojas.DataSource      := dsCidadeLoja;

  QryGrupo                        := TFDQuery.Create(nil);
  QryGrupo.Connection             := FDConnection1;
  dsGrupo                         := TDataSource.Create(nil);
  dsGrupo.DataSet                 := QryGrupo;
  gridGrupo.DataSource            := dsGrupo;

  QrySubGrupo                     := TFDQuery.Create(nil);
  QrySubGrupo.Connection          := FDConnection1;
  dsSubGrupo                      := TDataSource.Create(nil);
  dsSubGrupo.DataSet              := QrySubGrupo;
  gridSubGrupo.DataSource         := dsSubGrupo;

  QryFornecedor                   := TFDQuery.Create(nil);
  QryFornecedor.Connection        := FDConnection1;
  dsFornecedor                    := TDataSource.Create(nil);
  dsFornecedor.DataSet            := QryFornecedor;
  gridFornecedor.DataSource       := dsFornecedor;

  QryMarca                        := TFDQuery.Create(nil);
  QryMarca.Connection             := FDConnection1;
  dsMarca                         := TDataSource.Create(nil);
  dsMarca.DataSet                 := QryMarca;
  gridMarca.DataSource            := dsMarca;

  QryVendedor                     := TFDQuery.Create(nil);
  QryVendedor.Connection          := FDConnection1;
  dsVendedor                      := TDataSource.Create(nil);
  dsVendedor.DataSet              := QryVendedor;
  gridVendedor.DataSource         := dsVendedor;

  QryPublicidade                  := TFDQuery.Create(nil);
  QryPublicidade.Connection       := FDConnection1;
  dsPublicidade                   := TDataSource.Create(nil);
  dsPublicidade.DataSet           := QryPublicidade;
  gridPublicidade.DataSource      := dsPublicidade;

  QryDiaSemana                    := TFDQuery.Create(nil);
  QryDiaSemana.Connection         := FDConnection1;
  dsDiaSemana                     := TDataSource.Create(nil);
  dsDiaSemana.DataSet             := QryDiaSemana;
  gridDiaSemana.DataSource        := dsDiaSemana;

  QryHora                         := TFDQuery.Create(nil);
  QryHora.Connection              := FDConnection1;
  dsHora                          := TDataSource.Create(nil);
  dsHora.DataSet                  := QryHora;
  gridHora.DataSource             := dsHora;

  QryABC                          := TFDQuery.Create(nil);
  QryABC.Connection               := FDConnection1;
  dsABC                           := TDataSource.Create(nil);
  dsABC.DataSet                   := QryABC;
  gridABC.DataSource              := dsABC;

  QryProdutos                     := TFDQuery.Create(nil);
  QryProdutos.Connection          := FDConnection1;
  dsProdutos                      := TDataSource.Create(nil);
  dsProdutos.DataSet              := QryProdutos;
  gridProdutos.DataSource         := dsProdutos;

  qryFinanPlanoVenda              := TFDQuery.Create(nil);
  qryFinanPlanoVenda.Connection   := FDConnection1;
  dsFinanPlanoVenda               := TDataSource.Create(nil);
  dsFinanPlanoVenda.DataSet       := qryFinanPlanoVenda;
  gridFinanPlanoVenda.DataSource  := dsFinanPlanoVenda;

  qryFinanTabelaPreco             := TFDQuery.Create(nil);
  qryFinanTabelaPreco.Connection  := FDConnection1;
  dsFinanTabelaPreco              := TDataSource.Create(nil);
  dsFinanTabelaPreco.DataSet      := qryFinanTabelaPreco;
  gridFinanTabelaPreco.DataSource := dsFinanTabelaPreco;

  qryFinanConta                   := TFDQuery.Create(nil);
  qryFinanConta.Connection        := FDConnection1;
  dsFinanConta                    := TDataSource.Create(nil);
  dsFinanConta.DataSet            := qryFinanConta;
  gridFinanConta.DataSource       := dsFinanConta;

  qryFinanHistorico               := TFDQuery.Create(nil);
  qryFinanHistorico.Connection    := FDConnection1;
  dsFinanHistorico                := TDataSource.Create(nil);
  dsFinanHistorico.DataSet        := qryFinanHistorico;
  gridFinanHistorico.DataSource   := dsFinanHistorico;
end;

procedure TfBIVendas.DestruirQuerys;
begin
  FreeAndNil(QryTotais);
  FreeAndNil(dsTotais);

  FreeAndNil(QryRegiao);
  FreeAndNil(dsRegiao);

  FreeAndNil(QryLoja);
  FreeAndNil(dsLoja);

  FreeAndNil(QryCidadeClientes);
  FreeAndNil(dsCidadeClientes);

  FreeAndNil(QryCidadeLoja);
  FreeAndNil(dsCidadeLoja);

  FreeAndNil(QryGrupo);
  FreeAndNil(dsGrupo);

  FreeAndNil(QrySubGrupo);
  FreeAndNil(dsSubGrupo);

  FreeAndNil(QryFornecedor);
  FreeAndNil(dsFornecedor);

  FreeAndNil(QryMarca);
  FreeAndNil(dsMarca);

  FreeAndNil(QryVendedor);
  FreeAndNil(dsVendedor);

  FreeAndNil(QryPublicidade);
  FreeAndNil(dsPublicidade);

  FreeAndNil(QryDiaSemana);
  FreeAndNil(dsDiaSemana);

  FreeAndNil(QryHora);
  FreeAndNil(dsHora);

  FreeAndNil(QryABC);
  FreeAndNil(dsABC);

  FreeAndNil(QryProdutos);
  FreeAndNil(dsProdutos);

  FreeAndNil(qryFinanPlanoVenda);
  FreeAndNil(dsFinanPlanoVenda);

  FreeAndNil(qryFinanTabelaPreco);
  FreeAndNil(dsFinanTabelaPreco);

  FreeAndNil(qryFinanConta);
  FreeAndNil(dsFinanConta);

  FreeAndNil(qryFinanHistorico);
  FreeAndNil(dsFinanHistorico);
end;

procedure TfBIVendas.ConsVendasGeral;
begin
  try
    Aguarde(self, True);

    ControlsCDS(False);
    FecharDataSet;
    LimpaFiltros;

    Application.ProcessMessages;

    FDConnection1.DriverName := 'SQLite';
    FDConnection1.Connected  := True;

    FDLocalSQL1.Connection   := FDConnection1;
    FDLocalSQL1.Active       := False;
    FDLocalSQL1.DataSets.Clear;

    if FDMTVendas.Active then
      FDMTVendas.EmptyDataSet;

    FDLocalSQL1.DataSets.Add(FDMTVendas);
    cdsVendasGeral.Close;
    cdsVendasGeral.CommandText := ' SELECT  ' + //
      ' * ' + //
      ' FROM ' + //
      ' NP_VENDAS_BI( ' + //
      QuotedStr(FormatDateTime('DD.MM.YYYY', dpkDataIni.Date)) + ' , ' + //
      QuotedStr(FormatDateTime('DD.MM.YYYY', dpkDataFim.Date)) + ') VBI ' + //
      ' JOIN NP_VENDAS_BI_ABC( ' + //
      QuotedStr(FormatDateTime('DD.MM.YYYY', dpkDataIni.Date)) + ' , ' + //
      QuotedStr(FormatDateTime('DD.MM.YYYY', dpkDataFim.Date)) + //
      ' ) T ON T.ID_PRODUTO = VBI.ID_PRODUTO ';
    cdsVendasGeral.Open;
//     mensagem(nil, cdsVendasGeral.CommandText);

    if (cdsVendasGeral.RecordCount > 0) then
    begin
      FDMTVendas.CopyDataSet(cdsVendasGeral, [coStructure, coRestart, coAppend]);
      FDLocalSQL1.Active := True;
      FiltroBI;
    end;

    Aguarde(self, False);

  except
    on Exc: Exception do
    begin
      Aguarde(self, False);
      Mensagem(self, 'Erro. ' + #13#10 + Exc.Message);
    end;
  end;
end;

procedure TfBIVendas.LimpaFiltros;
begin
  fBuscaGenerica.clearFilter([ //
    'ID_REGIAO',  //
    'ID_LOJA',  //
    'ID_CIDADE_CLIENTE',  //
    'ID_CIDADE_LOJA',  //
    'ID_GRUPO',  //
    'ID_SUBGRUPO',  //
    'ID_FORNECEDOR',  //
    'ID_MARCA',  //
    'ID_VENDEDOR',  //
    'ID_PUBLICIDADE',  //
    'ID_DIA_SEMANA',  //
    'ID_HORA',  //
    'CURVA',  //
    'ID_PRODUTO',  //
    'ID_CONTA',  //
    'ID_HISTORICO',  //
    'ID_PLANO',  //
    'TABELA_PRECO' //
    ]);
end;

procedure TfBIVendas.ConfiguraGrid;
begin

  AddFieldGrid(gridRegiao,            ['Cód',                 'Região',               'Qt. Prod.',              'Qt. Ped.',             'Total de Vendas',      '% Vend.',            'CMV',              'Lucro Bruto',          '% Lucro Bruto',            'Desp. ADM',                '% Med. Desp.',         'Lucro Liq.',                 '% Lucro Liq.',                 'Acréscimo',            '% Acréscimo',              'Desconto',         '% Desconto'], //
                                      ['ID_REGIAO',           'REGIAO',               'TOTAL_QTD_PRODUTOS',     'TOTAL_QTD_PEDIDOS',    'TOTAL_VENDAS',         'VENDA_PCT',          'TOTAL_CMV',        'LUCRO_BRUTO_RS',       'LUCRO_BRUTO_PCT',          'VALOR_DESPESA_ADM_CX',     'PCT_DESPESA_ADM_CX',   'VALOR_LUCRO_LIQUIDO',        'VALOR_LUCRO_LIQUIDO_PCT',      'TOTAL_ACRESCIMO',      'TOTAL_ACRESCIMO_PCT',      'TOTAL_DESCONTO',   'TOTAL_DESCONTO_PCT'],    //
                                      [57,                    250,                    55,                       55,                     85,                     82,                   82,                 82,                     82,                         82,                         82,                     82,                           80,                             80,                     80,                         80,                 80]);

  AddFieldGrid(gridLoja,              ['Cód',                 'Loja',                 'Qt. Prod.',              'Qt. Ped.',             'Total de Vendas',      '% Vend.',            'CMV',              'Lucro Bruto',          '% Lucro Bruto',            'Desp. ADM',                '% Med. Desp.',         'Lucro Liq.',                 '% Lucro Liq.',                 'Acréscimo',            '% Acréscimo',              'Desconto',         '% Desconto'], //
                                      ['ID_LOJA',             'LOJA',                 'TOTAL_QTD_PRODUTOS',     'TOTAL_QTD_PEDIDOS',    'TOTAL_VENDAS',         'VENDA_PCT',          'TOTAL_CMV',        'LUCRO_BRUTO_RS',       'LUCRO_BRUTO_PCT',          'VALOR_DESPESA_ADM_CX',     'PCT_DESPESA_ADM_CX',   'VALOR_LUCRO_LIQUIDO',        'VALOR_LUCRO_LIQUIDO_PCT',      'TOTAL_ACRESCIMO',      'TOTAL_ACRESCIMO_PCT',      'TOTAL_DESCONTO',   'TOTAL_DESCONTO_PCT'],    //
                                      [57,                    250,                    55,                       55,                     85,                     82,                   82,                 82,                     82,                         82,                         82,                     82,                           80,                             80,                     80,                         80,                 80]);

  AddFieldGrid(gridCidadeClientes,    ['Cód',                 'Cidade do Cliente',    'Total',                  '% Vend.',              'Lucro Liq.',           '% Lucro Liq.'],    //
                                      ['ID_CIDADE_CLIENTE',   'CIDADE_CLIENTE',       'TOTAL_VENDAS',           'VENDA_PCT',            'VALOR_LUCRO_LIQUIDO',  'VALOR_LUCRO_LIQUIDO_PCT'], //
                                      [65,                    210,                    90,                       55,                     85,                     55]);

  AddFieldGrid(gridCidadeLojas,       ['Cód',                 'Cidade da Loja',       'Total', '% Vend.',       'Lucro Liq.',           '% Lucro Liq.'], //
                                      ['ID_CIDADE_LOJA',      'CIDADE_LOJA',          'TOTAL_VENDAS',           'VENDA_PCT',            'VALOR_LUCRO_LIQUIDO',  'VALOR_LUCRO_LIQUIDO_PCT'], //
                                      [65,                    210,                    90,                       55,                     85,                     55]);

  AddFieldGrid(gridGrupo,             ['Cód',                 'Grupo',                'Qt. Prod.',              'Qt. Ped.',             'Total de Vendas',      '% Vend.',            'CMV',              'Lucro Bruto',          '% Lucro Bruto',            'Desp. ADM',                '% Med. Desp.',         'Lucro Liq.',                 '% Lucro Liq.',                 'Acréscimo',            '% Acréscimo',              'Desconto',         '% Desconto'], //
                                      ['ID_GRUPO',            'GRUPO',                'TOTAL_QTD_PRODUTOS',     'TOTAL_QTD_PEDIDOS',    'TOTAL_VENDAS',         'VENDA_PCT',          'TOTAL_CMV',        'LUCRO_BRUTO_RS',       'LUCRO_BRUTO_PCT',          'VALOR_DESPESA_ADM_CX',     'PCT_DESPESA_ADM_CX',   'VALOR_LUCRO_LIQUIDO',        'VALOR_LUCRO_LIQUIDO_PCT',      'TOTAL_ACRESCIMO',      'TOTAL_ACRESCIMO_PCT',      'TOTAL_DESCONTO',   'TOTAL_DESCONTO_PCT'],    //
                                      [57,                    250,                    55,                       55,                     85,                     82,                   82,                 82,                     82,                         82,                         82,                     82,                           80,                             80,                     80,                         80,                 80]);

  AddFieldGrid(gridSubGrupo,          ['Cód',                 'SubGrupo',             'Qt. Prod.',              'Qt. Ped.',             'Total de Vendas',      '% Vend.',            'CMV',              'Lucro Bruto',          '% Lucro Bruto',            'Desp. ADM',                '% Med. Desp.',         'Lucro Liq.',                 '% Lucro Liq.',                 'Acréscimo',            '% Acréscimo',              'Desconto',         '% Desconto'], //
                                      ['ID_SUBGRUPO',         'SUBGRUPO',             'TOTAL_QTD_PRODUTOS',     'TOTAL_QTD_PEDIDOS',    'TOTAL_VENDAS',         'VENDA_PCT',          'TOTAL_CMV',        'LUCRO_BRUTO_RS',       'LUCRO_BRUTO_PCT',          'VALOR_DESPESA_ADM_CX',     'PCT_DESPESA_ADM_CX',   'VALOR_LUCRO_LIQUIDO',        'VALOR_LUCRO_LIQUIDO_PCT',      'TOTAL_ACRESCIMO',      'TOTAL_ACRESCIMO_PCT',      'TOTAL_DESCONTO',   'TOTAL_DESCONTO_PCT'], //
                                      [57,                    250,                    55,                       55,                     85,                     82,                   82,                 82,                     82,                         82,                         82,                     82,                           80,                             80,                     80,                         80,                 80]);

  AddFieldGrid(gridFornecedor,        ['Cód',                 'Fornecedor',           'Qt. Prod.',              'Qt. Ped.',             'Total de Vendas',      '% Vend.',            'CMV',              'Lucro Bruto',          '% Lucro Bruto',            'Desp. ADM',                '% Med. Desp.',         'Lucro Liq.',                 '% Lucro Liq.',                 'Acréscimo',            '% Acréscimo',              'Desconto',         '% Desconto'], //
                                      ['ID_FORNECEDOR',       'FORNECEDOR',           'TOTAL_QTD_PRODUTOS',     'TOTAL_QTD_PEDIDOS',    'TOTAL_VENDAS',         'VENDA_PCT',          'TOTAL_CMV',        'LUCRO_BRUTO_RS',       'LUCRO_BRUTO_PCT',          'VALOR_DESPESA_ADM_CX',     'PCT_DESPESA_ADM_CX',   'VALOR_LUCRO_LIQUIDO',        'VALOR_LUCRO_LIQUIDO_PCT',      'TOTAL_ACRESCIMO',      'TOTAL_ACRESCIMO_PCT',      'TOTAL_DESCONTO',   'TOTAL_DESCONTO_PCT'], //
                                      [57,                    250,                    55,                       55,                     85,                     82,                   82,                 82,                     82,                         82,                         82,                     82,                           80,                             80,                     80,                         80,                 80]);

  AddFieldGrid(gridMarca,             ['Cód',                 'Marca',                'Qt. Prod.',              'Qt. Ped.',             'Total de Vendas',      '% Vend.',            'CMV',              'Lucro Bruto',          '% Lucro Bruto',            'Desp. ADM',                '% Med. Desp.',         'Lucro Liq.',                 '% Lucro Liq.',                 'Acréscimo',            '% Acréscimo',              'Desconto',         '% Desconto'], //
                                      ['ID_MARCA',            'MARCA',                'TOTAL_QTD_PRODUTOS',     'TOTAL_QTD_PEDIDOS',    'TOTAL_VENDAS',         'VENDA_PCT',          'TOTAL_CMV',        'LUCRO_BRUTO_RS',       'LUCRO_BRUTO_PCT',          'VALOR_DESPESA_ADM_CX',     'PCT_DESPESA_ADM_CX',   'VALOR_LUCRO_LIQUIDO',        'VALOR_LUCRO_LIQUIDO_PCT',      'TOTAL_ACRESCIMO',      'TOTAL_ACRESCIMO_PCT',      'TOTAL_DESCONTO',   'TOTAL_DESCONTO_PCT'],    //
                                      [57,                    250,                    55,                       55,                     85,                     82,                   82,                 82,                     82,                         82,                         82,                     82,                           80,                             80,                     80,                         80,                 80]);

  AddFieldGrid(gridVendedor,          ['Cód',                 'Vendedor',             'Qt. Prod.',              'Qt. Ped.',             'Total de Vendas',      '% Vend.',            'CMV',              'Lucro Bruto',          '% Lucro Bruto',            'Desp. ADM',                '% Med. Desp.',         'Lucro Liq.',                 '% Lucro Liq.',                 'Acréscimo',            '% Acréscimo',              'Desconto',         '% Desconto'], //
                                      ['ID_VENDEDOR',         'VENDEDOR',             'TOTAL_QTD_PRODUTOS',     'TOTAL_QTD_PEDIDOS',    'TOTAL_VENDAS',         'VENDA_PCT',          'TOTAL_CMV',        'LUCRO_BRUTO_RS',       'LUCRO_BRUTO_PCT',          'VALOR_DESPESA_ADM_CX',     'PCT_DESPESA_ADM_CX',   'VALOR_LUCRO_LIQUIDO',        'VALOR_LUCRO_LIQUIDO_PCT',      'TOTAL_ACRESCIMO',      'TOTAL_ACRESCIMO_PCT',      'TOTAL_DESCONTO',   'TOTAL_DESCONTO_PCT'], //
                                      [57,                    250,                    55,                       55,                     85,                     82,                   82,                 82,                     82,                         82,                         82,                     82,                           80,                             80,                     80,                         80,                 80]);

  AddFieldGrid(gridPublicidade,       ['Cód',                 'Publicidade',          'Total de Vendas',        '% Vend.',              'Lucro Liq.',           '% Lucro Liq.'],    //
                                      ['ID_PUBLICIDADE',      'PUBLICIDADE',          'TOTAL_VENDAS',           'VENDA_PCT',            'VALOR_LUCRO_LIQUIDO',  'VALOR_LUCRO_LIQUIDO_PCT'], //
                                      [65,                    210,                    90,                       55,                     85,                     55]);

  AddFieldGrid(gridDiaSemana,         ['Dia',                 'Total de Vendas',      '% Vend.'], //
                                      ['DIA_SEMANA',          'TOTAL_VENDAS',         'VENDA_PCT'], //
                                      [135,                   100,                    55]);

  AddFieldGrid(gridABC,               ['Curva ABC',           'Total de Vendas',      '% Vend.',                '% Lucro Liq.'], //
                                      ['CURVA',               'TOTAL_VENDAS',         'VENDA_PCT',              'VALOR_LUCRO_LIQUIDO_PCT'], //
                                      [85,                    90,                     55,                       55]);

  AddFieldGrid(gridHora,              ['Horário',             'Total de Vendas',      '% Vend.'], //
                                      ['HORA',                'TOTAL_VENDAS',         'VENDA_PCT'], //
                                      [78,                    100,                    55]);

  AddFieldGrid(gridProdutos,          ['Cód',                 'Produto',              'Referência',             'Qtd.',                 'ABC',                  'Total de Vendas',    '% Vend.',          'Lucro Bruto',          '% Lucro Bruto',            'Desp. ADM',                '% Med. Desp.',         'Lucro Liq.',                 '% Lucro Liq.',                 'Acréscimo',            '% Acréscimo',              'Desconto',         '% Desconto'], //
                                      ['ID_PRODUTO',          'PRODUTO',              'REFERENCIA',             'TOTAL_QTD_PRODUTOS',   'CURVA',                'TOTAL_VENDAS',       'VENDA_PCT',        'LUCRO_BRUTO_RS',       'LUCRO_BRUTO_PCT',          'VALOR_DESPESA_ADM_CX',     'PCT_DESPESA_ADM_CX',   'VALOR_LUCRO_LIQUIDO',        'VALOR_LUCRO_LIQUIDO_PCT',      'TOTAL_ACRESCIMO',      'TOTAL_ACRESCIMO_PCT',      'TOTAL_DESCONTO',   'TOTAL_DESCONTO_PCT'],    //
                                      [70,                    280,                    155,                      50,                     30,                     85,                   55,                 70,                     50,                         70,                         80,                     80,                           80,                             80,                     80,                         80,                 80]);

  AddFieldGrid(gridFinanPlanoVenda,   ['Cód',                 'Plano de Venda',       'QP',                     'Total de Vendas',      '% Vend.'], //
                                      ['ID_PLANO',            'PLANO_VENDA',          'QP',                     'TOTAL_VENDAS',         'VENDA_PCT'], //
                                      [45,                    236,                    40,                       90, 55]);

  AddFieldGrid(gridFinanConta,        ['Cód',                 'Conta',                'Total de Vendas',        '% Vend.'], //
                                      ['ID_CONTA',            'CONTA',                'TOTAL_VENDAS',           'VENDA_PCT'], //
                                      [45,                    125,                    85,                       55]);

  AddFieldGrid(gridFinanHistorico,    ['Cód',                 'Histórico',            'Total de Vendas',        '% Vend.'], //
                                      ['ID_HISTORICO',        'HISTORICO',            'TOTAL_VENDAS',           'VENDA_PCT'], //
                                      [45,                    125,                    85,                       55]);

  AddFieldGrid(gridFinanTabelaPreco,  ['Coluna de Preço',     'Qt. Prod.',            'Qt. Ped.',               'Total de Vendas',      '% Vend.',              'CMV',                'Lucro Bruto',      '% Lucro Bruto',        'Desp. ADM',                '% Med. Desp.',             'Lucro Liq.',           '% Lucro Liq.',               'Acréscimo',                    '% Acréscimo',          'Desconto',                 '% Desconto'], //
                                      ['TABELA_PRECO',        'TOTAL_QTD_PRODUTOS',   'TOTAL_QTD_PEDIDOS',      'TOTAL_VENDAS',         'VENDA_PCT',            'TOTAL_CMV',          'LUCRO_BRUTO_RS',   'LUCRO_BRUTO_PCT',      'VALOR_DESPESA_ADM_CX',     'PCT_DESPESA_ADM_CX',       'VALOR_LUCRO_LIQUIDO',  'VALOR_LUCRO_LIQUIDO_PCT',    'TOTAL_ACRESCIMO',              'TOTAL_ACRESCIMO_PCT',  'TOTAL_DESCONTO',           'TOTAL_DESCONTO_PCT'],    //
                                      [120,                   55,                     55,                       85,                     82,                     82,                   82,                 82,                     82,                         82,                         82,                     80,                           80,                             80,                     80,                         80]);

end;

procedure TfBIVendas.OrdenaGrid(Column: TColumn);
begin
  with TFDQuery(TDataSource(Column.Grid.DataSource).DataSet) do
  begin
    IndexFieldNames               := EmptyStr;
    IndexFieldNames               := Column.FieldName;
    IndexDefs.AddIndexDef.Options := [ixDescending];
    IndexFieldNames               := Column.FieldName + ':D';
  end;
end;

procedure TfBIVendas.PageControl1Change(Sender: TObject);
begin
  btnMenu.Enabled := PageControl1.ActivePageIndex = 1;
end;

procedure TfBIVendas.FecharDataSet;
begin
  QryTotais.Close;
  QryRegiao.Close;
  QryLoja.Close;
  QryCidadeClientes.Close;
  QryCidadeLoja.Close;
  QryGrupo.Close;
  QrySubGrupo.Close;
  QryFornecedor.Close;
  QryMarca.Close;
  QryVendedor.Close;
  QryPublicidade.Close;
  QryDiaSemana.Close;
  QryHora.Close;
  QryProdutos.Close;
  QryABC.Close;

  qryFinanConta.Close;
  qryFinanHistorico.Close;
  qryFinanPlanoVenda.Close;
  qryFinanTabelaPreco.Close;

  CarregaLabels;
end;

procedure TfBIVendas.CarregaLabels;
begin
  if ((QryTotais.Active) and (QryTotais.RecordCount > 0)) then
  begin
    TOTAL_VENDAS                              := QryTotais.FieldByName('TOTAL_VENDAS').AsFloat;

    lblTOTAL_VENDAS.Caption                   := FormatFloat('#,##0.00',  QryTotais.FieldByName('TOTAL_VENDAS').AsFloat);
    lblTOTAL_QTD_PRODUTOS_QUANTIDADE.Caption  := Format('%.0n',          [QryTotais.FieldByName('TOTAL_QTD_PRODUTOS_QUANTIDADE').AsFloat]);
    lblTOTAL_QTD_PRODUTOS_DISTINTOS.Caption   := Format('%.0n',          [QryTotais.FieldByName('TOTAL_QTD_PRODUTOS_DISTINTOS').AsFloat]);
    lblTOTAL_QTD_PEDIDOS.Caption              := Format('%.0n',          [QryTotais.FieldByName('TOTAL_QTD_PEDIDOS').AsFloat]);
    lblVENDAS_MED_PRODUTO.Caption             := FormatFloat('#,##0.00',  QryTotais.FieldByName('VENDAS_MED_PRODUTO').AsFloat);
    lblVENDAS_MED_PEDIDO.Caption              := FormatFloat('#,##0.00',  QryTotais.FieldByName('VENDAS_MED_PEDIDO').AsFloat);
    lblVENDAS_MED_VENDEDOR.Caption            := FormatFloat('#,##0.00',  QryTotais.FieldByName('VENDAS_MED_VENDEDOR').AsFloat);
    lblTOTAL_QTD_ID_VENDEDOR.Caption          := Format('%.0n',          [QryTotais.FieldByName('TOTAL_QTD_ID_VENDEDOR').AsFloat]);
    lblTOTAL_QTD_ID_GRUPO.Caption             := Format('%.0n',          [QryTotais.FieldByName('TOTAL_QTD_ID_GRUPO').AsFloat]);
    lblTOTAL_QTD_ID_FORNECEDOR.Caption        := Format('%.0n',          [QryTotais.FieldByName('TOTAL_QTD_ID_FORNECEDOR').AsFloat]);
    lblTOTAL_QTD_ID_MARCA.Caption             := Format('%.0n',          [QryTotais.FieldByName('TOTAL_QTD_ID_MARCA').AsFloat]);
    lblTOTAL_QTD_ID_CLIENTE.Caption           := Format('%.0n',          [QryTotais.FieldByName('TOTAL_QTD_ID_CLIENTE').AsFloat]);
    lblTOTAL_QTD_ID_CIDADE_LOJA.Caption       := Format('%.0n',          [QryTotais.FieldByName('TOTAL_QTD_ID_CIDADE_LOJA').AsFloat]);
    lblTOTAL_QTD_ID_CIDADE_CLIENTE.Caption    := Format('%.0n',          [QryTotais.FieldByName('TOTAL_QTD_ID_CIDADE_CLIENTE').AsFloat]);
    lblVALOR_LUCRO_LIQUIDO.Caption            := FormatFloat('#,##0.00',  QryTotais.FieldByName('VALOR_LUCRO_LIQUIDO').AsFloat);
    lblVALOR_LUCRO_LIQUIDO_PCT.Caption        := FormatFloat('#,##0.00%', QryTotais.FieldByName('VALOR_LUCRO_LIQUIDO_PCT').AsFloat);
    lblTOTAL_DESCONTO.Caption                 := FormatFloat('#,##0.00',  QryTotais.FieldByName('TOTAL_DESCONTO').AsFloat);
    lblTOTAL_Acrescimo.Caption                := FormatFloat('#,##0.00',  QryTotais.FieldByName('TOTAL_ACRESCIMO').AsFloat);
    lblTOTAL_ACRESCIMO_PCT.Caption            := FormatFloat('#,##0.00%', QryTotais.FieldByName('TOTAL_ACRESCIMO_PCT').AsFloat);
    lblTOTAL_DESCONTO_PCT.Caption             := FormatFloat('#,##0.00%', QryTotais.FieldByName('TOTAL_DESCONTO_PCT').AsFloat);
    lblTOTAL_CMV.Caption                      := FormatFloat('#,##0.00',  QryTotais.FieldByName('TOTAL_CMV').AsFloat);
    lblLUCRO_BRUTO_RS.Caption                 := FormatFloat('#,##0.00',  QryTotais.FieldByName('LUCRO_BRUTO_RS').AsFloat);
    lblLUCRO_BRUTO_PCT.Caption                := FormatFloat('#,##0.00%', QryTotais.FieldByName('LUCRO_BRUTO_PCT').AsFloat);
    lblTOTAL_VALOR_DESPESA_ADM_CX.Caption     := FormatFloat('#,##0.00',  QryTotais.FieldByName('TOTAL_VALOR_DESPESA_ADM_CX').AsFloat);
    lblTOTAL_VALOR_DESPESA_ADM_CX_PCT.Caption := FormatFloat('#,##0.00',  QryTotais.FieldByName('TOTAL_VALOR_DESPESA_ADM_CX_PCT').AsFloat);

  end
  else
  begin
    lblTOTAL_VENDAS.Caption                   := '0,00';
    lblTOTAL_QTD_PRODUTOS_QUANTIDADE.Caption  := '0';
    lblTOTAL_QTD_PRODUTOS_DISTINTOS.Caption   := '0';
    lblTOTAL_QTD_PEDIDOS.Caption              := '0';
    lblVENDAS_MED_PRODUTO.Caption             := '0,00';
    lblVENDAS_MED_PEDIDO.Caption              := '0,00';
    lblVENDAS_MED_VENDEDOR.Caption            := '0,00';
    lblTOTAL_QTD_ID_VENDEDOR.Caption          := '0';
    lblTOTAL_QTD_ID_GRUPO.Caption             := '0';
    lblTOTAL_QTD_ID_FORNECEDOR.Caption        := '0';
    lblTOTAL_QTD_ID_MARCA.Caption             := '0';
    lblTOTAL_QTD_ID_CLIENTE.Caption           := '0';
    lblTOTAL_QTD_ID_CIDADE_LOJA.Caption       := '0';
    lblTOTAL_QTD_ID_CIDADE_CLIENTE.Caption    := '0';
    lblVALOR_LUCRO_LIQUIDO.Caption            := '0,00';
    lblVALOR_LUCRO_LIQUIDO_PCT.Caption        := '0,00%';
    lblTOTAL_DESCONTO.Caption                 := '0,00';
    lblTOTAL_Acrescimo.Caption                := '0,00';
    lblTOTAL_ACRESCIMO_PCT.Caption            := '0,00%';
    lblTOTAL_DESCONTO_PCT.Caption             := '0,00%';
    lblTOTAL_CMV.Caption                      := '0,00';
    lblLUCRO_BRUTO_RS.Caption                 := '0,00';
    lblLUCRO_BRUTO_PCT.Caption                := '0,00%';
    lblTOTAL_VALOR_DESPESA_ADM_CX.Caption     := '0,00';
    lblTOTAL_VALOR_DESPESA_ADM_CX_PCT.Caption := '0,00%';
  end;
  Application.ProcessMessages;
end;

procedure TfBIVendas.ControlsCDS(vStatus: Boolean);
begin
  dsRegiao.Enabled           := vStatus;
  dsLoja.Enabled             := vStatus;
  dsCidadeClientes.Enabled   := vStatus;
  dsCidadeLoja.Enabled       := vStatus;
  dsGrupo.Enabled            := vStatus;
  dsSubGrupo.Enabled         := vStatus;
  dsFornecedor.Enabled       := vStatus;
  dsMarca.Enabled            := vStatus;
  dsVendedor.Enabled         := vStatus;
  dsPublicidade.Enabled      := vStatus;
  dsDiaSemana.Enabled        := vStatus;
  dsHora.Enabled             := vStatus;
  dsProdutos.Enabled         := vStatus;
  dsABC.Enabled              := vStatus;

  dsFinanConta.Enabled       := vStatus;
  dsFinanHistorico.Enabled   := vStatus;
  dsFinanPlanoVenda.Enabled  := vStatus;
  dsFinanTabelaPreco.Enabled := vStatus;
end;

procedure TfBIVendas.gridRegiaoDblClick(Sender: TObject);
begin
  CallBackRegiao(QryRegiao.FieldByName('ID_REGIAO').AsString);
end;

procedure TfBIVendas.gridLojaDblClick(Sender: TObject);
begin
  if not fBuscaGenerica.clearFilter('ID_LOJA') then
  begin
    fBuscaGenerica.addFilter('', 'ID_LOJA', '=', //
      QryLoja.FieldByName('ID_LOJA').AsString, //
      '', //
      QryLoja.FieldByName('LOJA').AsString, //
      'Loja');
  end;

  vWhereBI := fBuscaGenerica.getSql(vWhereBI);
  FiltroBI;
end;

procedure TfBIVendas.gridCidadeClientesDblClick(Sender: TObject);
begin
  if not fBuscaGenerica.clearFilter('ID_CIDADE_CLIENTE') then
  begin
    fBuscaGenerica.addFilter('', 'ID_CIDADE_CLIENTE', '=', //
      QryCidadeClientes.FieldByName('ID_CIDADE_CLIENTE').AsString, //
      '', //
      QryCidadeClientes.FieldByName('CIDADE_CLIENTE').AsString, //
      'Cidade Cliente');
  end;

  vWhereBI := fBuscaGenerica.getSql(vWhereBI);
  FiltroBI;
end;

procedure TfBIVendas.gridCidadeLojasDblClick(Sender: TObject);
begin
  if not fBuscaGenerica.clearFilter('ID_CIDADE_LOJA') then
  begin
    fBuscaGenerica.addFilter('', 'ID_CIDADE_LOJA', '=', //
      QryCidadeLoja.FieldByName('ID_CIDADE_LOJA').AsString, //
      '', //
      QryCidadeLoja.FieldByName('CIDADE_LOJA').AsString, //
      'Cidade Loja');
  end;

  vWhereBI := fBuscaGenerica.getSql(vWhereBI);
  FiltroBI;
end;

procedure TfBIVendas.gridGrupoDblClick(Sender: TObject);
begin
  if not fBuscaGenerica.clearFilter('ID_GRUPO') then
  begin
    fBuscaGenerica.addFilter('', 'ID_GRUPO', '=', //
      QryGrupo.FieldByName('ID_GRUPO').AsString, //
      '', //
      QryGrupo.FieldByName('GRUPO').AsString, //
      'Grupo');
  end;

  vWhereBI := fBuscaGenerica.getSql(vWhereBI);
  FiltroBI;
end;

procedure TfBIVendas.gridSubGrupoDblClick(Sender: TObject);
begin
  if not fBuscaGenerica.clearFilter('ID_SUBGRUPO') then
  begin
    fBuscaGenerica.addFilter('', 'ID_SUBGRUPO', '=', //
      QrySubGrupo.FieldByName('ID_SUBGRUPO').AsString, //
      '', //
      QrySubGrupo.FieldByName('SUBGRUPO').AsString, //
      'SubGrupo');
  end;

  vWhereBI := fBuscaGenerica.getSql(vWhereBI);
  FiltroBI;
end;

procedure TfBIVendas.gridFornecedorDblClick(Sender: TObject);
begin
  if not fBuscaGenerica.clearFilter('ID_FORNECEDOR') then
  begin
    fBuscaGenerica.addFilter('', 'ID_FORNECEDOR', '=', //
      QryFornecedor.FieldByName('ID_FORNECEDOR').AsString, //
      '', //
      QryFornecedor.FieldByName('FORNECEDOR').AsString, //
      'Fornecedor');
  end;

  vWhereBI := fBuscaGenerica.getSql(vWhereBI);
  FiltroBI;
end;

procedure TfBIVendas.gridMarcaDblClick(Sender: TObject);
begin
  if not fBuscaGenerica.clearFilter('ID_MARCA') then
  begin
    fBuscaGenerica.addFilter('', 'ID_MARCA', '=', //
      QryMarca.FieldByName('ID_MARCA').AsString, //
      '', //
      QryMarca.FieldByName('MARCA').AsString, //
      'Marca');
  end;

  vWhereBI := fBuscaGenerica.getSql(vWhereBI);
  FiltroBI;
end;

procedure TfBIVendas.gridVendedorDblClick(Sender: TObject);
begin
  if not fBuscaGenerica.clearFilter('ID_VENDEDOR') then
  begin
    fBuscaGenerica.addFilter('', 'ID_VENDEDOR', '=', //
      QryVendedor.FieldByName('ID_VENDEDOR').AsString, //
      '', //
      QryVendedor.FieldByName('VENDEDOR').AsString, // top
      'Vendedor');
  end;

  vWhereBI := fBuscaGenerica.getSql(vWhereBI);
  FiltroBI;
end;

procedure TfBIVendas.gridPublicidadeDblClick(Sender: TObject);
begin
  if not fBuscaGenerica.clearFilter('ID_PUBLICIDADE') then
  begin
    fBuscaGenerica.addFilter('', 'ID_PUBLICIDADE', '=', //
      QryPublicidade.FieldByName('ID_PUBLICIDADE').AsString, //
      '', //
      QryPublicidade.FieldByName('PUBLICIDADE').AsString, // top
      'Publicidade');
  end;

  vWhereBI := fBuscaGenerica.getSql(vWhereBI);
  FiltroBI;
end;

procedure TfBIVendas.gridDiaSemanaDblClick(Sender: TObject);
begin
  if not fBuscaGenerica.clearFilter('ID_DIA_SEMANA') then
  begin
    fBuscaGenerica.addFilter('', 'ID_DIA_SEMANA', '=', //
      QryDiaSemana.FieldByName('ID_DIA_SEMANA').AsString, //
      '', //
      QryDiaSemana.FieldByName('DIA_SEMANA').AsString, // top
      'Dia da Semana');
  end;

  vWhereBI := fBuscaGenerica.getSql(vWhereBI);
  FiltroBI;
end;

procedure TfBIVendas.gridHoraDblClick(Sender: TObject);
begin
  if not fBuscaGenerica.clearFilter('ID_HORA') then
  begin
    fBuscaGenerica.addFilter('', 'ID_HORA', '=', //
      QryHora.FieldByName('ID_HORA').AsString, //
      '', //
      QryHora.FieldByName('HORA').AsString, // top
      'Horário');
  end;

  vWhereBI := fBuscaGenerica.getSql(vWhereBI);
  FiltroBI;
end;

procedure TfBIVendas.gridABCDblClick(Sender: TObject);
begin
  if not fBuscaGenerica.clearFilter('CURVA') then
  begin
    fBuscaGenerica.addFilter('', 'CURVA', '=', //
      QryABC.FieldByName('CURVA').AsString, //
      '', //
      QryABC.FieldByName('CURVA').AsString, //
      'Curva');
  end;

  vWhereBI := fBuscaGenerica.getSql(vWhereBI);
  FiltroBI;
end;

procedure TfBIVendas.gridProdutosDblClick(Sender: TObject);
begin
  if not fBuscaGenerica.clearFilter('ID_PRODUTO') then
  begin
    fBuscaGenerica.addFilter('', 'ID_PRODUTO', '=', //
      QryProdutos.FieldByName('ID_PRODUTO').AsString, //
      '', //
      QryProdutos.FieldByName('PRODUTO').AsString, //
      'Produto');
  end;

  vWhereBI := fBuscaGenerica.getSql(vWhereBI);
  FiltroBI;
end;

procedure TfBIVendas.gridFinanContaDblClick(Sender: TObject);
begin
  if not fBuscaGenerica.clearFilter('ID_CONTA') then
  begin
    fBuscaGenerica.addFilter('', 'ID_CONTA', '=', //
      qryFinanConta.FieldByName('ID_CONTA').AsString, //
      '', //
      qryFinanConta.FieldByName('CONTA').AsString, //
      'Conta');
  end;

  vWhereBI := fBuscaGenerica.getSql(vWhereBI);
  FiltroBI;
end;

procedure TfBIVendas.gridFinanHistoricoDblClick(Sender: TObject);
begin
  if not fBuscaGenerica.clearFilter('ID_HISTORICO') then
  begin
    fBuscaGenerica.addFilter('', 'ID_HISTORICO', '=', //
      qryFinanHistorico.FieldByName('ID_HISTORICO').AsString, //
      '', //
      qryFinanHistorico.FieldByName('HISTORICO').AsString, //
      'Histórico');
  end;

  vWhereBI := fBuscaGenerica.getSql(vWhereBI);
  FiltroBI;
end;

procedure TfBIVendas.gridFinanPlanoVendaDblClick(Sender: TObject);
begin
  if not fBuscaGenerica.clearFilter('ID_PLANO') then
  begin
    fBuscaGenerica.addFilter('', 'ID_PLANO', '=', //
      qryFinanPlanoVenda.FieldByName('ID_PLANO').AsString, //
      '', //
      qryFinanPlanoVenda.FieldByName('PLANO_VENDA').AsString, //
      'Plano');
  end;

  vWhereBI := fBuscaGenerica.getSql(vWhereBI);
  FiltroBI;
end;

procedure TfBIVendas.gridFinanTabelaPrecoDblClick(Sender: TObject);
begin
  if not fBuscaGenerica.clearFilter('TABELA_PRECO') then
  begin
    fBuscaGenerica.addFilter('', 'TABELA_PRECO', '=', //
      qryFinanTabelaPreco.FieldByName('TABELA_PRECO').AsString, //
      '', //
      qryFinanTabelaPreco.FieldByName('TABELA_PRECO').AsString, //
      'Tabela de Preço');
  end;

  vWhereBI := fBuscaGenerica.getSql(vWhereBI);
  FiltroBI;
end;

procedure TfBIVendas.FiltroBI;
var
  Flags: OLEVariant;
  I, J: Integer;
begin
  try
    Aguarde(self, True);
    ControlsCDS(False);

    QryTotais.Active := False;
    QryTotais.SQL.Text := ' SELECT   ' + //
                          '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS), 0) AS TOTAL_VENDAS, ' + //
                          '  COALESCE (SUM (TOTAL_CMV_PRODUTOS), 0) AS TOTAL_CMV, ' + //
                          '  COALESCE (SUM (TOTAL_ACRESCIMO), 0) AS TOTAL_ACRESCIMO, ' + //
                          '  COALESCE (SUM (TOTAL_DESCONTO), 0) AS TOTAL_DESCONTO, ' + //
                          '  COALESCE (SUM (VALOR_DESPESA_ADM_CX), 0) AS TOTAL_VALOR_DESPESA_ADM_CX, ' + //
                          '  COALESCE ((SUM (VALOR_DESPESA_ADM_CX)*100)/SUM (TOTAL_VENDAS_PRODUTOS) , 0) AS TOTAL_VALOR_DESPESA_ADM_CX_PCT, ' + //
                          '  COALESCE ((SUM (TOTAL_ACRESCIMO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_ACRESCIMO) , 0) AS TOTAL_ACRESCIMO_PCT, ' + //
                          '  COALESCE ((SUM (TOTAL_DESCONTO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_DESCONTO) , 0) AS TOTAL_DESCONTO_PCT, ' + //
                          '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS), 0) AS LUCRO_BRUTO_RS, ' + //
                          '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS LUCRO_BRUTO_PCT, ' + //
                          '  COALESCE (SUM (QUANTIDADE / NUMPAGS), 0) AS TOTAL_QTD_PRODUTOS_QUANTIDADE, ' + //
                          '  COALESCE (COUNT (DISTINCT (ID_PRODUTO)), 0) AS TOTAL_QTD_PRODUTOS_DISTINTOS, ' + //
                          '  COALESCE (COUNT (DISTINCT (ID_PEDIDO)), 0) AS TOTAL_QTD_PEDIDOS, ' + //
                          '  COALESCE (COUNT (DISTINCT (ID_CIDADE_LOJA)), 0) AS TOTAL_QTD_ID_CIDADE_LOJA, ' + //
                          '  COALESCE (COUNT (DISTINCT (ID_CIDADE_CLIENTE)), 0) AS TOTAL_QTD_ID_CIDADE_CLIENTE, ' + //
                          '  COALESCE (COUNT (DISTINCT (ID_VENDEDOR)), 0) AS TOTAL_QTD_ID_VENDEDOR, ' + //
                          '  COALESCE (COUNT (DISTINCT (ID_FORNECEDOR)), 0) AS TOTAL_QTD_ID_FORNECEDOR, ' + //
                          '  COALESCE (COUNT (DISTINCT (ID_GRUPO)), 0) AS TOTAL_QTD_ID_GRUPO, ' + //
                          '  COALESCE (COUNT (DISTINCT (ID_MARCA)), 0) AS TOTAL_QTD_ID_MARCA, ' + //
                          '  COALESCE (COUNT (DISTINCT (ID_CLIENTE)), 0) AS TOTAL_QTD_ID_CLIENTE, ' + //
                          '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) / COUNT (DISTINCT (ID_VENDEDOR)), 0) AS VENDAS_MED_VENDEDOR, ' + //
                          '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) / COUNT (DISTINCT (ID_PEDIDO)), 0) AS VENDAS_MED_PEDIDO, ' + //
                          '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) / SUM (QUANTIDADE / NUMPAGS), 0) AS VENDAS_MED_PRODUTO, ' + //
                          '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_LUCRO_LIQUIDO, ' + //
                          '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS VALOR_LUCRO_LIQUIDO_PCT ' + //
                          ' FROM FDMTVendas ' + //
                          ' WHERE 1 = 1 ' + //
      vWhereBI;
    QryTotais.Active := True;
    (QryTotais.FieldByName('TOTAL_VENDAS'                 ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryTotais.FieldByName('TOTAL_QTD_PRODUTOS_QUANTIDADE') as TFloatField).DisplayFormat := '#,##0.00';
    (QryTotais.FieldByName('VENDAS_MED_VENDEDOR'          ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryTotais.FieldByName('VENDAS_MED_PEDIDO'            ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryTotais.FieldByName('VENDAS_MED_PRODUTO'           ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryTotais.FieldByName('VALOR_LUCRO_LIQUIDO'          ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryTotais.FieldByName('VALOR_LUCRO_LIQUIDO_PCT'      ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryTotais.FieldByName('TOTAL_ACRESCIMO'              ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryTotais.FieldByName('TOTAL_ACRESCIMO_PCT'          ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryTotais.FieldByName('TOTAL_DESCONTO'               ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryTotais.FieldByName('TOTAL_DESCONTO_PCT'           ) as TFloatField).DisplayFormat := '#,##0.00';

    CarregaLabels;

    QryRegiao.Active := False;
    QryRegiao.SQL.Text := ' SELECT   ' + //
                          '  ID_REGIAO, ' + //
                          '  REGIAO, ' + //
                          '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS), 0) AS TOTAL_VENDAS, ' + //
                          '  COALESCE (SUM (TOTAL_ACRESCIMO), 0) AS TOTAL_ACRESCIMO, ' + //
                          '  COALESCE (SUM (TOTAL_DESCONTO), 0) AS TOTAL_DESCONTO, ' + //
                          '  COALESCE ((SUM (TOTAL_ACRESCIMO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_ACRESCIMO) , 0) AS TOTAL_ACRESCIMO_PCT, ' + //
                          '  COALESCE ((SUM (TOTAL_DESCONTO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_DESCONTO) , 0) AS TOTAL_DESCONTO_PCT, ' + //
                          '  COALESCE (COUNT ( DISTINCT(ID_PEDIDO) ), 0 ) AS TOTAL_QTD_PEDIDOS, ' + //
                          '  COALESCE (SUM (QUANTIDADE / NUMPAGS), 0) AS TOTAL_QTD_PRODUTOS, ' + //
                          '  COALESCE (((SUM (TOTAL_VENDAS_PRODUTOS)) * 100 / ' + uFuncoes.FB_Numeric(TOTAL_VENDAS, 15, 2) + '), 0) AS VENDA_PCT, ' + //
                          '  COALESCE (SUM (TOTAL_CMV_PRODUTOS), 0) AS TOTAL_CMV, ' + //
                          '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS), 0) AS LUCRO_BRUTO_RS, ' + //
                          '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS LUCRO_BRUTO_PCT, ' + //
                          '  COALESCE (SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_DESPESA_ADM_CX, ' + //
                          '  COALESCE (SUM (VALOR_DESPESA_ADM_CX) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS PCT_DESPESA_ADM_CX, ' + //
                          '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_LUCRO_LIQUIDO, ' + //
                          '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS VALOR_LUCRO_LIQUIDO_PCT ' + //
                          ' FROM FDMTVendas ' + //
                          ' WHERE 1 = 1 ' + //
                          vWhereBI + //
                          ' GROUP BY ' + //
                          '   ID_REGIAO, ' + //
                          '   REGIAO ' + //
                          ' ORDER BY TOTAL_VENDAS DESC ';

    QryRegiao.Active                                                                 := True;
    (QryRegiao.FieldByName('TOTAL_VENDAS'           ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryRegiao.FieldByName('VENDA_PCT'              ) as TFloatField).DisplayFormat := '#,##0.00%';
    (QryRegiao.FieldByName('TOTAL_CMV'              ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryRegiao.FieldByName('LUCRO_BRUTO_RS'         ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryRegiao.FieldByName('LUCRO_BRUTO_PCT'        ) as TFloatField).DisplayFormat := '#,##0.00%';
    (QryRegiao.FieldByName('VALOR_DESPESA_ADM_CX'   ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryRegiao.FieldByName('PCT_DESPESA_ADM_CX'     ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryRegiao.FieldByName('VALOR_LUCRO_LIQUIDO'    ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryRegiao.FieldByName('VALOR_LUCRO_LIQUIDO_PCT') as TFloatField).DisplayFormat := '#,##0.00%';
    (QryRegiao.FieldByName('TOTAL_QTD_PRODUTOS'     ) as TFloatField).DisplayFormat := '#0';
    (QryRegiao.FieldByName('TOTAL_ACRESCIMO'        ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryRegiao.FieldByName('TOTAL_ACRESCIMO_PCT'    ) as TFloatField).DisplayFormat := '#,##0.00%';
    (QryRegiao.FieldByName('TOTAL_DESCONTO'         ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryRegiao.FieldByName('TOTAL_DESCONTO_PCT'     ) as TFloatField).DisplayFormat := '#,##0.00%';

    QryLoja.Active := False;
    QryLoja.SQL.Text := ' SELECT   ' + //
                        '   ID_LOJA, ' + //
                        '   LOJA, ' + //
                        '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS), 0) AS TOTAL_VENDAS, ' + //
                        '  COALESCE (SUM (TOTAL_ACRESCIMO), 0) AS TOTAL_ACRESCIMO, ' + //
                        '  COALESCE (SUM (TOTAL_DESCONTO), 0) AS TOTAL_DESCONTO, ' + //
                        '  COALESCE ((SUM (TOTAL_ACRESCIMO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_ACRESCIMO) , 0) AS TOTAL_ACRESCIMO_PCT, ' + //
                        '  COALESCE ((SUM (TOTAL_DESCONTO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_DESCONTO) , 0) AS TOTAL_DESCONTO_PCT, ' + //
                        '  COALESCE( COUNT ( DISTINCT(ID_PEDIDO) ), 0 ) AS TOTAL_QTD_PEDIDOS, ' + //
                        '  COALESCE (SUM (QUANTIDADE / NUMPAGS), 0) AS TOTAL_QTD_PRODUTOS, ' + //
                        '  COALESCE (((SUM (TOTAL_VENDAS_PRODUTOS)) * 100 / ' + uFuncoes.FB_Numeric(TOTAL_VENDAS, 15, 2) + '), 0) AS VENDA_PCT, ' + //
                        '  COALESCE (SUM (TOTAL_CMV_PRODUTOS), 0) AS TOTAL_CMV, ' + //
                        '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS), 0) AS LUCRO_BRUTO_RS, ' + //
                        '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS LUCRO_BRUTO_PCT, ' + //
                        '  COALESCE (SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_DESPESA_ADM_CX, ' + //
                        '  COALESCE (SUM (VALOR_DESPESA_ADM_CX) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS PCT_DESPESA_ADM_CX, ' + //
                        '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_LUCRO_LIQUIDO, ' + //
                        '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS VALOR_LUCRO_LIQUIDO_PCT ' + //
                        ' FROM FDMTVendas ' + //
                        ' WHERE 1 = 1 ' + //
                        vWhereBI + //
                        ' GROUP BY ' +
                        '   ID_LOJA, ' + //
                        '   LOJA ' + //
                        ' ORDER BY TOTAL_VENDAS DESC ';

    QryLoja.Active := True;
    (QryLoja.FieldByName('TOTAL_VENDAS'           ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryLoja.FieldByName('VENDA_PCT'              ) as TFloatField).DisplayFormat := '#,##0.00%';
    (QryLoja.FieldByName('TOTAL_CMV'              ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryLoja.FieldByName('LUCRO_BRUTO_RS'         ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryLoja.FieldByName('LUCRO_BRUTO_PCT'        ) as TFloatField).DisplayFormat := '#,##0.00%';
    (QryLoja.FieldByName('VALOR_DESPESA_ADM_CX'   ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryLoja.FieldByName('PCT_DESPESA_ADM_CX'     ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryLoja.FieldByName('VALOR_LUCRO_LIQUIDO'    ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryLoja.FieldByName('VALOR_LUCRO_LIQUIDO_PCT') as TFloatField).DisplayFormat := '#,##0.00%';
    (QryLoja.FieldByName('TOTAL_QTD_PRODUTOS'     ) as TFloatField).DisplayFormat := '#0';
    (QryLoja.FieldByName('TOTAL_ACRESCIMO'        ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryLoja.FieldByName('TOTAL_ACRESCIMO_PCT'    ) as TFloatField).DisplayFormat := '#,##0.00%';
    (QryLoja.FieldByName('TOTAL_DESCONTO'         ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryLoja.FieldByName('TOTAL_DESCONTO_PCT'     ) as TFloatField).DisplayFormat := '#,##0.00%';

    QryCidadeClientes.Active := False;
    QryCidadeClientes.SQL.Text := ' SELECT   ' + //
                                  '  ID_CIDADE_CLIENTE, ' + //
                                  '  CIDADE_CLIENTE, ' + //
                                  '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS), 0) AS TOTAL_VENDAS, ' + //
                                  '  COALESCE (SUM (TOTAL_ACRESCIMO), 0) AS TOTAL_ACRESCIMO, ' + //
                                  '  COALESCE (SUM (TOTAL_DESCONTO), 0) AS TOTAL_DESCONTO, ' + //
                                  '  COALESCE ((SUM (TOTAL_ACRESCIMO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_ACRESCIMO) , 0) AS TOTAL_ACRESCIMO_PCT, ' + //
                                  '  COALESCE ((SUM (TOTAL_DESCONTO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_DESCONTO) , 0) AS TOTAL_DESCONTO_PCT, ' + // '  COALESCE( COUNT ( DISTINCT(ID_PEDIDO) ), 0 ) AS TOTAL_QTD_PEDIDOS, '+ //
                                  '  COALESCE (SUM (QUANTIDADE / NUMPAGS), 0) AS TOTAL_QTD_PRODUTOS, ' + //
                                  '  COALESCE (((SUM (TOTAL_VENDAS_PRODUTOS)) * 100 / ' + uFuncoes.FB_Numeric(TOTAL_VENDAS, 15, 2) + '), 0) AS VENDA_PCT, ' + //
                                  '  COALESCE (SUM (TOTAL_CMV_PRODUTOS), 0) AS TOTAL_CMV, ' + //
                                  '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS), 0) AS LUCRO_BRUTO_RS, ' + //
                                  '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS LUCRO_BRUTO_PCT, ' + //
                                  '  COALESCE (SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_DESPESA_ADM_CX, ' + //
                                  '  COALESCE (SUM (VALOR_DESPESA_ADM_CX) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS PCT_DESPESA_ADM_CX, ' + //
                                  '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_LUCRO_LIQUIDO, ' + //
                                  '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS VALOR_LUCRO_LIQUIDO_PCT ' + //
                                  ' FROM FDMTVendas ' + //
                                  ' WHERE 1 = 1 ' + //
                                  vWhereBI + //
                                  ' GROUP BY ' + //
                                  '   ID_CIDADE_CLIENTE, ' + //
                                  '   CIDADE_CLIENTE ' + //
                                  ' ORDER BY TOTAL_VENDAS DESC ';
    QryCidadeClientes.Active := True;
    (QryCidadeClientes.FieldByName('TOTAL_VENDAS'           ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryCidadeClientes.FieldByName('VENDA_PCT'              ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QryCidadeClientes.FieldByName('TOTAL_CMV'              ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryCidadeClientes.FieldByName('LUCRO_BRUTO_RS'         ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryCidadeClientes.FieldByName('LUCRO_BRUTO_PCT'        ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QryCidadeClientes.FieldByName('VALOR_DESPESA_ADM_CX'   ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryCidadeClientes.FieldByName('PCT_DESPESA_ADM_CX'     ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryCidadeClientes.FieldByName('VALOR_LUCRO_LIQUIDO'    ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryCidadeClientes.FieldByName('VALOR_LUCRO_LIQUIDO_PCT') as TFloatField).DisplayFormat :='#,##0.00%';
    (QryCidadeClientes.FieldByName('TOTAL_QTD_PRODUTOS'     ) as TFloatField).DisplayFormat :='#0';
    (QryCidadeClientes.FieldByName('TOTAL_ACRESCIMO'        ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryCidadeClientes.FieldByName('TOTAL_ACRESCIMO_PCT'    ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QryCidadeClientes.FieldByName('TOTAL_DESCONTO'         ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryCidadeClientes.FieldByName('TOTAL_DESCONTO_PCT'     ) as TFloatField).DisplayFormat :='#,##0.00%';

    QryCidadeLoja.Active := False;
    QryCidadeLoja.SQL.Text := ' SELECT   ' + //
                              '   ID_CIDADE_LOJA, ' + //
                              '   CIDADE_LOJA, ' + //
                              '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS), 0) AS TOTAL_VENDAS, ' + //
                              '  COALESCE (SUM (TOTAL_ACRESCIMO), 0) AS TOTAL_ACRESCIMO, ' + //
                              '  COALESCE (SUM (TOTAL_DESCONTO), 0) AS TOTAL_DESCONTO, ' + //
                              '  COALESCE ((SUM (TOTAL_ACRESCIMO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_ACRESCIMO) , 0) AS TOTAL_ACRESCIMO_PCT, ' + //
                              '  COALESCE ((SUM (TOTAL_DESCONTO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_DESCONTO) , 0) AS TOTAL_DESCONTO_PCT, ' + //
                              '  COALESCE (COUNT ( DISTINCT(ID_PEDIDO) ), 0 ) AS TOTAL_QTD_PEDIDOS, ' + //
                              '  COALESCE (SUM (QUANTIDADE / NUMPAGS), 0) AS TOTAL_QTD_PRODUTOS, ' + //
                              '  COALESCE (((SUM (TOTAL_VENDAS_PRODUTOS)) * 100 / ' + uFuncoes.FB_Numeric(TOTAL_VENDAS, 15, 2) + '), 0) AS VENDA_PCT, ' + //
                              '  COALESCE (SUM (TOTAL_CMV_PRODUTOS), 0) AS TOTAL_CMV, ' + //
                              '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS), 0) AS LUCRO_BRUTO_RS, ' + //
                              '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS LUCRO_BRUTO_PCT, ' + //
                              '  COALESCE (SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_DESPESA_ADM_CX, ' + //
                              '  COALESCE (SUM (VALOR_DESPESA_ADM_CX) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS PCT_DESPESA_ADM_CX, ' + //
                              '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_LUCRO_LIQUIDO, ' + //
                              '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS VALOR_LUCRO_LIQUIDO_PCT ' + //
                              ' FROM FDMTVendas ' + //
                              ' WHERE 1 = 1 ' + //
                              vWhereBI + //
                              ' GROUP BY ' + //
                              '   ID_CIDADE_LOJA, ' + //
                              '   CIDADE_LOJA ' + //
                              ' ORDER BY TOTAL_VENDAS DESC ';
    QryCidadeLoja.Active := True;
    (QryCidadeLoja.FieldByName('TOTAL_VENDAS'           ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryCidadeLoja.FieldByName('VENDA_PCT'              ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QryCidadeLoja.FieldByName('TOTAL_CMV'              ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryCidadeLoja.FieldByName('LUCRO_BRUTO_RS'         ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryCidadeLoja.FieldByName('LUCRO_BRUTO_PCT'        ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QryCidadeLoja.FieldByName('VALOR_DESPESA_ADM_CX'   ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryCidadeLoja.FieldByName('PCT_DESPESA_ADM_CX'     ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryCidadeLoja.FieldByName('VALOR_LUCRO_LIQUIDO'    ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryCidadeLoja.FieldByName('VALOR_LUCRO_LIQUIDO_PCT') as TFloatField).DisplayFormat :='#,##0.00%';
    (QryCidadeLoja.FieldByName('TOTAL_QTD_PRODUTOS'     ) as TFloatField).DisplayFormat :='#0';
    (QryCidadeLoja.FieldByName('TOTAL_ACRESCIMO'        ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryCidadeLoja.FieldByName('TOTAL_ACRESCIMO_PCT'    ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QryCidadeLoja.FieldByName('TOTAL_DESCONTO'         ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryCidadeLoja.FieldByName('TOTAL_DESCONTO_PCT'     ) as TFloatField).DisplayFormat :='#,##0.00%';

    QryGrupo.Active := False;
    QryGrupo.SQL.Text :=  ' SELECT   ' + //
                          '  ID_GRUPO, ' + //
                          '  GRUPO, ' + //
                          '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS), 0) AS TOTAL_VENDAS, ' + //
                          '  COALESCE (SUM (TOTAL_ACRESCIMO), 0) AS TOTAL_ACRESCIMO, ' + //
                          '  COALESCE (SUM (TOTAL_DESCONTO), 0) AS TOTAL_DESCONTO, ' + //
                          '  COALESCE ((SUM (TOTAL_ACRESCIMO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_ACRESCIMO) , 0) AS TOTAL_ACRESCIMO_PCT, ' + //
                          '  COALESCE ((SUM (TOTAL_DESCONTO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_DESCONTO) , 0) AS TOTAL_DESCONTO_PCT, ' + //
                          '  COALESCE (COUNT ( DISTINCT(ID_PEDIDO) ), 0 ) AS TOTAL_QTD_PEDIDOS, ' + //
                          '  COALESCE (SUM (QUANTIDADE / NUMPAGS), 0) AS TOTAL_QTD_PRODUTOS, ' + //
                          '  COALESCE (((SUM (TOTAL_VENDAS_PRODUTOS)) * 100 / ' + uFuncoes.FB_Numeric(TOTAL_VENDAS, 15, 2) + '), 0) AS VENDA_PCT, ' + //
                          '  COALESCE (SUM (TOTAL_CMV_PRODUTOS), 0) AS TOTAL_CMV, ' + //
                          '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS), 0) AS LUCRO_BRUTO_RS, ' + //
                          '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS LUCRO_BRUTO_PCT, ' + //
                          '  COALESCE (SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_DESPESA_ADM_CX, ' + //
                          '  COALESCE (SUM (VALOR_DESPESA_ADM_CX) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS PCT_DESPESA_ADM_CX, ' + //
                          '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_LUCRO_LIQUIDO, ' + //
                          '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS VALOR_LUCRO_LIQUIDO_PCT ' + //
                          ' FROM FDMTVendas ' + //
                          ' WHERE 1 = 1 ' + //
                          vWhereBI + //
                          ' GROUP BY ' + //
                          '  ID_GRUPO, ' + //
                          '  GRUPO ' + //
                          ' ORDER BY TOTAL_VENDAS DESC ';
    QryGrupo.Active := True;
    (QryGrupo.FieldByName('TOTAL_VENDAS'           ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryGrupo.FieldByName('VENDA_PCT'              ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QryGrupo.FieldByName('TOTAL_CMV'              ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryGrupo.FieldByName('LUCRO_BRUTO_RS'         ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryGrupo.FieldByName('LUCRO_BRUTO_PCT'        ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QryGrupo.FieldByName('VALOR_DESPESA_ADM_CX'   ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryGrupo.FieldByName('PCT_DESPESA_ADM_CX'     ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryGrupo.FieldByName('VALOR_LUCRO_LIQUIDO'    ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryGrupo.FieldByName('VALOR_LUCRO_LIQUIDO_PCT') as TFloatField).DisplayFormat :='#,##0.00%';
    (QryGrupo.FieldByName('TOTAL_QTD_PRODUTOS'     ) as TFloatField).DisplayFormat :='#0';
    (QryGrupo.FieldByName('TOTAL_ACRESCIMO'        ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryGrupo.FieldByName('TOTAL_ACRESCIMO_PCT'    ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QryGrupo.FieldByName('TOTAL_DESCONTO'         ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryGrupo.FieldByName('TOTAL_DESCONTO_PCT'     ) as TFloatField).DisplayFormat :='#,##0.00%';

    QrySubGrupo.Active := False;
    QrySubGrupo.SQL.Text := ' SELECT   ' + //
                            '  ID_SUBGRUPO, ' + //
                            '  SUBGRUPO, ' + //
                            '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS), 0) AS TOTAL_VENDAS, ' + //
                            '  COALESCE (SUM (TOTAL_ACRESCIMO), 0) AS TOTAL_ACRESCIMO, ' + //
                            '  COALESCE (SUM (TOTAL_DESCONTO), 0) AS TOTAL_DESCONTO, ' + //
                            '  COALESCE ((SUM (TOTAL_ACRESCIMO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_ACRESCIMO) , 0) AS TOTAL_ACRESCIMO_PCT, ' + //
                            '  COALESCE ((SUM (TOTAL_DESCONTO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_DESCONTO) , 0) AS TOTAL_DESCONTO_PCT, ' + //
                            '  COALESCE (COUNT ( DISTINCT(ID_PEDIDO) ), 0 ) AS TOTAL_QTD_PEDIDOS, ' + //
                            '  COALESCE (SUM (QUANTIDADE / NUMPAGS), 0) AS TOTAL_QTD_PRODUTOS, ' + //
                            '  COALESCE (((SUM (TOTAL_VENDAS_PRODUTOS)) * 100 / ' + uFuncoes.FB_Numeric(TOTAL_VENDAS, 15, 2) + '), 0) AS VENDA_PCT, ' + //
                            '  COALESCE (SUM (TOTAL_CMV_PRODUTOS), 0) AS TOTAL_CMV, ' + //
                            '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS), 0) AS LUCRO_BRUTO_RS, ' + //
                            '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS LUCRO_BRUTO_PCT, ' + //
                            '  COALESCE (SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_DESPESA_ADM_CX, ' + //
                            '  COALESCE (SUM (VALOR_DESPESA_ADM_CX) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS PCT_DESPESA_ADM_CX, ' + //
                            '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_LUCRO_LIQUIDO, ' + //
                            '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS VALOR_LUCRO_LIQUIDO_PCT ' + //
                            ' FROM FDMTVendas ' + //
                            ' WHERE 1 = 1 ' + //
                            vWhereBI + //
                            ' GROUP BY ' + //
                            '  ID_SUBGRUPO, ' + //
                            '  SUBGRUPO ' + //
                            ' ORDER BY TOTAL_VENDAS DESC ';
    QrySubGrupo.Active := True;
    (QrySubGrupo.FieldByName('TOTAL_VENDAS'           ) as TFloatField).DisplayFormat :='#,##0.00';
    (QrySubGrupo.FieldByName('VENDA_PCT'              ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QrySubGrupo.FieldByName('TOTAL_CMV'              ) as TFloatField).DisplayFormat :='#,##0.00';
    (QrySubGrupo.FieldByName('LUCRO_BRUTO_RS'         ) as TFloatField).DisplayFormat :='#,##0.00';
    (QrySubGrupo.FieldByName('LUCRO_BRUTO_PCT'        ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QrySubGrupo.FieldByName('VALOR_DESPESA_ADM_CX'   ) as TFloatField).DisplayFormat :='#,##0.00';
    (QrySubGrupo.FieldByName('PCT_DESPESA_ADM_CX'     ) as TFloatField).DisplayFormat :='#,##0.00';
    (QrySubGrupo.FieldByName('VALOR_LUCRO_LIQUIDO'    ) as TFloatField).DisplayFormat :='#,##0.00';
    (QrySubGrupo.FieldByName('VALOR_LUCRO_LIQUIDO_PCT') as TFloatField).DisplayFormat :='#,##0.00%';
    (QrySubGrupo.FieldByName('TOTAL_QTD_PRODUTOS'     ) as TFloatField).DisplayFormat :='#0';
    (QrySubGrupo.FieldByName('TOTAL_ACRESCIMO'        ) as TFloatField).DisplayFormat :='#,##0.00';
    (QrySubGrupo.FieldByName('TOTAL_ACRESCIMO_PCT'    ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QrySubGrupo.FieldByName('TOTAL_DESCONTO'         ) as TFloatField).DisplayFormat :='#,##0.00';
    (QrySubGrupo.FieldByName('TOTAL_DESCONTO_PCT'     ) as TFloatField).DisplayFormat :='#,##0.00%';

    QryFornecedor.Active := False;
    QryFornecedor.SQL.Text := ' SELECT   ' + //
                              '  ID_FORNECEDOR, ' + //
                              '  FORNECEDOR, ' + //
                              '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS), 0) AS TOTAL_VENDAS, ' + //
                              '  COALESCE (SUM (TOTAL_ACRESCIMO), 0) AS TOTAL_ACRESCIMO, ' + //
                              '  COALESCE (SUM (TOTAL_DESCONTO), 0) AS TOTAL_DESCONTO, ' + //
                              '  COALESCE ((SUM (TOTAL_ACRESCIMO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_ACRESCIMO) , 0) AS TOTAL_ACRESCIMO_PCT, ' + //
                              '  COALESCE ((SUM (TOTAL_DESCONTO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_DESCONTO) , 0) AS TOTAL_DESCONTO_PCT, ' + //
                              '  COALESCE (COUNT ( DISTINCT(ID_PEDIDO) ), 0 ) AS TOTAL_QTD_PEDIDOS, ' + //
                              '  COALESCE (SUM (QUANTIDADE / NUMPAGS), 0) AS TOTAL_QTD_PRODUTOS, ' + //
                              '  COALESCE (((SUM (TOTAL_VENDAS_PRODUTOS)) * 100 / ' + uFuncoes.FB_Numeric(TOTAL_VENDAS, 15, 2) + '), 0) AS VENDA_PCT, ' + //
                              '  COALESCE (SUM (TOTAL_CMV_PRODUTOS), 0) AS TOTAL_CMV, ' + //
                              '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS), 0) AS LUCRO_BRUTO_RS, ' + //
                              '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS LUCRO_BRUTO_PCT, ' + //
                              '  COALESCE (SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_DESPESA_ADM_CX, ' + //
                              '  COALESCE (SUM (VALOR_DESPESA_ADM_CX) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS PCT_DESPESA_ADM_CX, ' + //
                              '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_LUCRO_LIQUIDO, ' + //
                              '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS VALOR_LUCRO_LIQUIDO_PCT ' + //
                              ' FROM FDMTVendas ' + //
                              ' WHERE 1 = 1 ' + //
                              vWhereBI + //
                              ' GROUP BY ' + //
                              '  ID_FORNECEDOR, ' + //
                              '  FORNECEDOR ' + //
                              ' ORDER BY TOTAL_VENDAS DESC ';
    QryFornecedor.Active := True;
    (QryFornecedor.FieldByName('TOTAL_VENDAS'           ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryFornecedor.FieldByName('VENDA_PCT'              ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QryFornecedor.FieldByName('TOTAL_CMV'              ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryFornecedor.FieldByName('LUCRO_BRUTO_RS'         ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryFornecedor.FieldByName('LUCRO_BRUTO_PCT'        ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QryFornecedor.FieldByName('VALOR_DESPESA_ADM_CX'   ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryFornecedor.FieldByName('PCT_DESPESA_ADM_CX'     ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryFornecedor.FieldByName('VALOR_LUCRO_LIQUIDO'    ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryFornecedor.FieldByName('VALOR_LUCRO_LIQUIDO_PCT') as TFloatField).DisplayFormat :='#,##0.00%';
    (QryFornecedor.FieldByName('TOTAL_QTD_PRODUTOS'     ) as TFloatField).DisplayFormat :='#0';
    (QryFornecedor.FieldByName('TOTAL_ACRESCIMO'        ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryFornecedor.FieldByName('TOTAL_ACRESCIMO_PCT'    ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QryFornecedor.FieldByName('TOTAL_DESCONTO'         ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryFornecedor.FieldByName('TOTAL_DESCONTO_PCT'     ) as TFloatField).DisplayFormat :='#,##0.00%';

    QryMarca.Active := False;
    QryMarca.SQL.Text :=  ' SELECT   ' + //
                          '  ID_MARCA, ' + //
                          '  MARCA, ' + //
                          '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS), 0) AS TOTAL_VENDAS, ' + //
                          '  COALESCE (SUM (TOTAL_ACRESCIMO), 0) AS TOTAL_ACRESCIMO, ' + //
                          '  COALESCE (SUM (TOTAL_DESCONTO), 0) AS TOTAL_DESCONTO, ' + //
                          '  COALESCE ((SUM (TOTAL_ACRESCIMO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_ACRESCIMO) , 0) AS TOTAL_ACRESCIMO_PCT, ' + //
                          '  COALESCE ((SUM (TOTAL_DESCONTO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_DESCONTO) , 0) AS TOTAL_DESCONTO_PCT, ' + //
                          '  COALESCE (COUNT ( DISTINCT(ID_PEDIDO) ), 0 ) AS TOTAL_QTD_PEDIDOS, ' + //
                          '  COALESCE (SUM (QUANTIDADE / NUMPAGS), 0) AS TOTAL_QTD_PRODUTOS, ' + //
                          '  COALESCE (((SUM (TOTAL_VENDAS_PRODUTOS)) * 100 / ' + uFuncoes.FB_Numeric(TOTAL_VENDAS, 15, 2) + '), 0) AS VENDA_PCT, ' + //
                          '  COALESCE (SUM (TOTAL_CMV_PRODUTOS), 0) AS TOTAL_CMV, ' + //
                          '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS), 0) AS LUCRO_BRUTO_RS, ' + //
                          '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS LUCRO_BRUTO_PCT, ' + //
                          '  COALESCE (SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_DESPESA_ADM_CX, ' + //
                          '  COALESCE (SUM (VALOR_DESPESA_ADM_CX) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS PCT_DESPESA_ADM_CX, ' + //
                          '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_LUCRO_LIQUIDO, ' + //
                          '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS VALOR_LUCRO_LIQUIDO_PCT ' + //
                          ' FROM FDMTVendas ' + //
                          ' WHERE 1 = 1 ' + //
                          vWhereBI + //
                          ' GROUP BY ' + //
                          '   ID_MARCA, ' + //
                          '   MARCA ' + //
                          ' ORDER BY TOTAL_VENDAS DESC ';
    QryMarca.Active := True;
    (QryMarca.FieldByName('TOTAL_VENDAS'           ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryMarca.FieldByName('VENDA_PCT'              ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QryMarca.FieldByName('TOTAL_CMV'              ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryMarca.FieldByName('LUCRO_BRUTO_RS'         ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryMarca.FieldByName('LUCRO_BRUTO_PCT'        ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QryMarca.FieldByName('VALOR_DESPESA_ADM_CX'   ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryMarca.FieldByName('PCT_DESPESA_ADM_CX'     ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryMarca.FieldByName('VALOR_LUCRO_LIQUIDO'    ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryMarca.FieldByName('VALOR_LUCRO_LIQUIDO_PCT') as TFloatField).DisplayFormat :='#,##0.00%';
    (QryMarca.FieldByName('TOTAL_QTD_PRODUTOS'     ) as TFloatField).DisplayFormat :='#0';
    (QryMarca.FieldByName('TOTAL_ACRESCIMO'        ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryMarca.FieldByName('TOTAL_ACRESCIMO_PCT'    ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QryMarca.FieldByName('TOTAL_DESCONTO'         ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryMarca.FieldByName('TOTAL_DESCONTO_PCT'     ) as TFloatField).DisplayFormat :='#,##0.00%';

    QryVendedor.Active := False;
    QryVendedor.SQL.Text := ' SELECT   ' + //
                            '   ID_VENDEDOR,' + //
                            '   VENDEDOR, ' + //
                            '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS), 0) AS TOTAL_VENDAS, ' + //
                            '  COALESCE (SUM (TOTAL_ACRESCIMO), 0) AS TOTAL_ACRESCIMO, ' + //
                            '  COALESCE (SUM (TOTAL_DESCONTO), 0) AS TOTAL_DESCONTO, ' + //
                            '  COALESCE ((SUM (TOTAL_ACRESCIMO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_ACRESCIMO) , 0) AS TOTAL_ACRESCIMO_PCT, ' + //
                            '  COALESCE ((SUM (TOTAL_DESCONTO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_DESCONTO) , 0) AS TOTAL_DESCONTO_PCT, ' + //
                            '  COALESCE (COUNT ( DISTINCT(ID_PEDIDO) ), 0 ) AS TOTAL_QTD_PEDIDOS, ' + //
                            '  COALESCE (SUM (QUANTIDADE / NUMPAGS), 0) AS TOTAL_QTD_PRODUTOS, ' + //
                            '  COALESCE (((SUM (TOTAL_VENDAS_PRODUTOS)) * 100 / ' + uFuncoes.FB_Numeric(TOTAL_VENDAS, 15, 2) + '), 0) AS VENDA_PCT, ' + //
                            '  COALESCE (SUM (TOTAL_CMV_PRODUTOS), 0) AS TOTAL_CMV, ' + //
                            '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS), 0) AS LUCRO_BRUTO_RS, ' + //
                            '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS LUCRO_BRUTO_PCT, ' + //
                            '  COALESCE (SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_DESPESA_ADM_CX, ' + //
                            '  COALESCE (SUM (VALOR_DESPESA_ADM_CX) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS PCT_DESPESA_ADM_CX, ' + //
                            '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_LUCRO_LIQUIDO, ' + //
                            '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS VALOR_LUCRO_LIQUIDO_PCT ' + //
                            ' FROM FDMTVendas ' + //
                            ' WHERE 1 = 1 ' + //
                            vWhereBI + //
                            ' GROUP BY ' + //
                            '   ID_VENDEDOR,' + //
                            '   VENDEDOR ' + //
                            ' ORDER BY TOTAL_VENDAS DESC ';
    QryVendedor.Active := True;
    (QryVendedor.FieldByName('TOTAL_VENDAS'           ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryVendedor.FieldByName('VENDA_PCT'              ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QryVendedor.FieldByName('TOTAL_CMV'              ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryVendedor.FieldByName('LUCRO_BRUTO_RS'         ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryVendedor.FieldByName('LUCRO_BRUTO_PCT'        ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QryVendedor.FieldByName('VALOR_DESPESA_ADM_CX'   ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryVendedor.FieldByName('PCT_DESPESA_ADM_CX'     ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryVendedor.FieldByName('VALOR_LUCRO_LIQUIDO'    ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryVendedor.FieldByName('VALOR_LUCRO_LIQUIDO_PCT') as TFloatField).DisplayFormat :='#,##0.00%';
    (QryVendedor.FieldByName('TOTAL_QTD_PRODUTOS'     ) as TFloatField).DisplayFormat :='#0';
    (QryVendedor.FieldByName('TOTAL_ACRESCIMO'        ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryVendedor.FieldByName('TOTAL_ACRESCIMO_PCT'    ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QryVendedor.FieldByName('TOTAL_DESCONTO'         ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryVendedor.FieldByName('TOTAL_DESCONTO_PCT'     ) as TFloatField).DisplayFormat :='#,##0.00%';

    QryPublicidade.Active := False;
    QryPublicidade.SQL.Text :=' SELECT   ' + //
                              '  ID_PUBLICIDADE,' + //
                              '  PUBLICIDADE, ' + //
                              '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS), 0) AS TOTAL_VENDAS, ' + //
                              '  COALESCE (SUM (TOTAL_ACRESCIMO), 0) AS TOTAL_ACRESCIMO, ' + //
                              '  COALESCE (SUM (TOTAL_DESCONTO), 0) AS TOTAL_DESCONTO, ' + //
                              '  COALESCE ((SUM (TOTAL_ACRESCIMO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_ACRESCIMO) , 0) AS TOTAL_ACRESCIMO_PCT, ' + //
                              '  COALESCE ((SUM (TOTAL_DESCONTO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_DESCONTO) , 0) AS TOTAL_DESCONTO_PCT, ' + //
                              '  COALESCE (COUNT ( DISTINCT(ID_PEDIDO) ), 0 ) AS TOTAL_QTD_PEDIDOS, ' + //
                              '  COALESCE (SUM (QUANTIDADE / NUMPAGS), 0) AS TOTAL_QTD_PRODUTOS, ' + //
                              '  COALESCE (((SUM (TOTAL_VENDAS_PRODUTOS)) * 100 / ' + uFuncoes.FB_Numeric(TOTAL_VENDAS, 15, 2) + '), 0) AS VENDA_PCT, ' + //
                              '  COALESCE (SUM (TOTAL_CMV_PRODUTOS), 0) AS TOTAL_CMV, ' + //
                              '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS), 0) AS LUCRO_BRUTO_RS, ' + //
                              '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS LUCRO_BRUTO_PCT, ' + //
                              '  COALESCE (SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_DESPESA_ADM_CX, ' + //
                              '  COALESCE (SUM (VALOR_DESPESA_ADM_CX) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS PCT_DESPESA_ADM_CX, ' + //
                              '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_LUCRO_LIQUIDO, ' + //
                              '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS VALOR_LUCRO_LIQUIDO_PCT ' + //
                              ' FROM FDMTVendas ' + //
                              ' WHERE 1 = 1 ' + //
                              vWhereBI + //
                              ' GROUP BY ' + //
                              '   ID_PUBLICIDADE,' + //
                              '   PUBLICIDADE ' + //
                              ' ORDER BY TOTAL_VENDAS DESC ';
    QryPublicidade.Active := True;
    (QryPublicidade.FieldByName('TOTAL_VENDAS'           ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryPublicidade.FieldByName('VENDA_PCT'              ) as TFloatField).DisplayFormat := '#,##0.00%';
    (QryPublicidade.FieldByName('TOTAL_CMV'              ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryPublicidade.FieldByName('LUCRO_BRUTO_RS'         ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryPublicidade.FieldByName('LUCRO_BRUTO_PCT'        ) as TFloatField).DisplayFormat := '#,##0.00%';
    (QryPublicidade.FieldByName('VALOR_DESPESA_ADM_CX'   ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryPublicidade.FieldByName('PCT_DESPESA_ADM_CX'     ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryPublicidade.FieldByName('VALOR_LUCRO_LIQUIDO'    ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryPublicidade.FieldByName('VALOR_LUCRO_LIQUIDO_PCT') as TFloatField).DisplayFormat := '#,##0.00%';
    (QryPublicidade.FieldByName('TOTAL_QTD_PRODUTOS'     ) as TFloatField).DisplayFormat := '#0';
    (QryPublicidade.FieldByName('TOTAL_ACRESCIMO'        ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryPublicidade.FieldByName('TOTAL_ACRESCIMO_PCT'    ) as TFloatField).DisplayFormat := '#,##0.00%';
    (QryPublicidade.FieldByName('TOTAL_DESCONTO'         ) as TFloatField).DisplayFormat := '#,##0.00';
    (QryPublicidade.FieldByName('TOTAL_DESCONTO_PCT'     ) as TFloatField).DisplayFormat := '#,##0.00%';

    QryDiaSemana.Active := False;
    QryDiaSemana.SQL.Text :=' SELECT   ' + //
                            '  ID_DIA_SEMANA,' + //
                            '  DIA_SEMANA, ' + //
                            '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS), 0) AS TOTAL_VENDAS, ' + //
                            '  COALESCE (SUM (TOTAL_ACRESCIMO), 0) AS TOTAL_ACRESCIMO, ' + //
                            '  COALESCE (SUM (TOTAL_DESCONTO), 0) AS TOTAL_DESCONTO, ' + //
                            '  COALESCE ((SUM (TOTAL_ACRESCIMO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_ACRESCIMO) , 0) AS TOTAL_ACRESCIMO_PCT, ' + //
                            '  COALESCE ((SUM (TOTAL_DESCONTO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_DESCONTO) , 0) AS TOTAL_DESCONTO_PCT, ' + //
                            '  COALESCE (COUNT ( DISTINCT(ID_PEDIDO) ), 0 ) AS TOTAL_QTD_PEDIDOS, ' + //
                            '  COALESCE (SUM (QUANTIDADE / NUMPAGS), 0) AS TOTAL_QTD_PRODUTOS, ' + //
                            '  COALESCE (((SUM (TOTAL_VENDAS_PRODUTOS)) * 100 / ' + uFuncoes.FB_Numeric(TOTAL_VENDAS, 15, 2) + '), 0) AS VENDA_PCT, ' + //
                            '  COALESCE (SUM (TOTAL_CMV_PRODUTOS), 0) AS TOTAL_CMV, ' + //
                            '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS), 0) AS LUCRO_BRUTO_RS, ' + //
                            '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS LUCRO_BRUTO_PCT, ' + //
                            '  COALESCE (SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_DESPESA_ADM_CX, ' + //
                            '  COALESCE (SUM (VALOR_DESPESA_ADM_CX) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS PCT_DESPESA_ADM_CX, ' + //
                            '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_LUCRO_LIQUIDO, ' + //
                            '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS VALOR_LUCRO_LIQUIDO_PCT ' + //
                            ' FROM FDMTVendas ' + //
                            ' WHERE 1 = 1 ' + //
                            vWhereBI + //
                            ' GROUP BY ' + //
                            '   ID_DIA_SEMANA,' + //
                            '   DIA_SEMANA ' + //
                            ' ORDER BY ID_DIA_SEMANA ';
    QryDiaSemana.Active := True;
    (QryDiaSemana.FieldByName('TOTAL_VENDAS'           ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryDiaSemana.FieldByName('VENDA_PCT'              ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QryDiaSemana.FieldByName('TOTAL_CMV'              ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryDiaSemana.FieldByName('LUCRO_BRUTO_RS'         ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryDiaSemana.FieldByName('LUCRO_BRUTO_PCT'        ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QryDiaSemana.FieldByName('VALOR_DESPESA_ADM_CX'   ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryDiaSemana.FieldByName('PCT_DESPESA_ADM_CX'     ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryDiaSemana.FieldByName('VALOR_LUCRO_LIQUIDO'    ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryDiaSemana.FieldByName('VALOR_LUCRO_LIQUIDO_PCT') as TFloatField).DisplayFormat :='#,##0.00%';
    (QryDiaSemana.FieldByName('TOTAL_QTD_PRODUTOS'     ) as TFloatField).DisplayFormat :='#0';
    (QryDiaSemana.FieldByName('TOTAL_ACRESCIMO'        ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryDiaSemana.FieldByName('TOTAL_ACRESCIMO_PCT'    ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QryDiaSemana.FieldByName('TOTAL_DESCONTO'         ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryDiaSemana.FieldByName('TOTAL_DESCONTO_PCT'     ) as TFloatField).DisplayFormat :='#,##0.00%';

    QryHora.Active := False;
    QryHora.SQL.Text := ' SELECT   ' + //
                        '   ID_HORA, ' + //
                        '   HORA, ' + //
                        '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS), 0) AS TOTAL_VENDAS, ' + //
                        '  COALESCE (SUM (TOTAL_ACRESCIMO), 0) AS TOTAL_ACRESCIMO, ' + //
                        '  COALESCE (SUM (TOTAL_DESCONTO), 0) AS TOTAL_DESCONTO, ' + //
                        '  COALESCE ((SUM (TOTAL_ACRESCIMO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_ACRESCIMO) , 0) AS TOTAL_ACRESCIMO_PCT, ' + //
                        '  COALESCE ((SUM (TOTAL_DESCONTO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_DESCONTO) , 0) AS TOTAL_DESCONTO_PCT, ' + //
                        '  COALESCE (COUNT ( DISTINCT(ID_PEDIDO) ), 0 ) AS TOTAL_QTD_PEDIDOS, ' + //
                        '  COALESCE (SUM (QUANTIDADE / NUMPAGS), 0) AS TOTAL_QTD_PRODUTOS, ' + //
                        '  COALESCE (((SUM (TOTAL_VENDAS_PRODUTOS)) * 100 / ' + uFuncoes.FB_Numeric(TOTAL_VENDAS, 15, 2) + '), 0) AS VENDA_PCT, ' + //
                        '  COALESCE (SUM (TOTAL_CMV_PRODUTOS), 0) AS TOTAL_CMV, ' + //
                        '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS), 0) AS LUCRO_BRUTO_RS, ' + //
                        '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS LUCRO_BRUTO_PCT, ' + //
                        '  COALESCE (SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_DESPESA_ADM_CX, ' + //
                        '  COALESCE (SUM (VALOR_DESPESA_ADM_CX) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS PCT_DESPESA_ADM_CX, ' + //
                        '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_LUCRO_LIQUIDO, ' + //
                        '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS VALOR_LUCRO_LIQUIDO_PCT ' + //
                        ' FROM FDMTVendas ' + //
                        ' WHERE 1 = 1 ' + //
                        vWhereBI + //
                        ' GROUP BY ' + //
                        '   ID_HORA, ' + //
                        '   HORA ' + //
                        ' ORDER BY ID_HORA ';
    QryHora.Active := True;
    (QryHora.FieldByName('TOTAL_VENDAS'           ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryHora.FieldByName('VENDA_PCT'              ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QryHora.FieldByName('TOTAL_CMV'              ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryHora.FieldByName('LUCRO_BRUTO_RS'         ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryHora.FieldByName('LUCRO_BRUTO_PCT'        ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QryHora.FieldByName('VALOR_DESPESA_ADM_CX'   ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryHora.FieldByName('PCT_DESPESA_ADM_CX'     ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryHora.FieldByName('VALOR_LUCRO_LIQUIDO'    ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryHora.FieldByName('VALOR_LUCRO_LIQUIDO_PCT') as TFloatField).DisplayFormat :='#,##0.00%';
    (QryHora.FieldByName('TOTAL_QTD_PRODUTOS'     ) as TFloatField).DisplayFormat :='#0';
    (QryHora.FieldByName('TOTAL_ACRESCIMO'        ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryHora.FieldByName('TOTAL_ACRESCIMO_PCT'    ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QryHora.FieldByName('TOTAL_DESCONTO'         ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryHora.FieldByName('TOTAL_DESCONTO_PCT'     ) as TFloatField).DisplayFormat :='#,##0.00%';

    QryABC.Active := False;
    QryABC.SQL.Text :=' SELECT   ' + //
                      '   CURVA, ' + //
                      '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS), 0) AS TOTAL_VENDAS, ' + //
                      '  COALESCE (SUM (TOTAL_ACRESCIMO), 0) AS TOTAL_ACRESCIMO, ' + //
                      '  COALESCE (SUM (TOTAL_DESCONTO), 0) AS TOTAL_DESCONTO, ' + //
                      '  COALESCE ((SUM (TOTAL_ACRESCIMO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_ACRESCIMO) , 0) AS TOTAL_ACRESCIMO_PCT, ' + //
                      '  COALESCE ((SUM (TOTAL_DESCONTO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_DESCONTO) , 0) AS TOTAL_DESCONTO_PCT, ' + //
                      '  COALESCE (COUNT ( DISTINCT(ID_PEDIDO) ), 0 ) AS TOTAL_QTD_PEDIDOS, ' + //
                      '  COALESCE (SUM (QUANTIDADE / NUMPAGS), 0) AS TOTAL_QTD_PRODUTOS, ' + //
                      '  COALESCE (((SUM (TOTAL_VENDAS_PRODUTOS)) * 100 / ' + uFuncoes.FB_Numeric(TOTAL_VENDAS, 15, 2) + '), 0) AS VENDA_PCT, ' + //
                      '  COALESCE (SUM (TOTAL_CMV_PRODUTOS), 0) AS TOTAL_CMV, ' + //
                      '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS), 0) AS LUCRO_BRUTO_RS, ' + //
                      '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS LUCRO_BRUTO_PCT, ' + //
                      '  COALESCE (SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_DESPESA_ADM_CX, ' + //
                      '  COALESCE (SUM (VALOR_DESPESA_ADM_CX) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS PCT_DESPESA_ADM_CX, ' + //
                      '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_LUCRO_LIQUIDO, ' + //
                      '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS VALOR_LUCRO_LIQUIDO_PCT ' + //
                      ' FROM FDMTVendas ' + //
                      ' WHERE 1 = 1 ' + //
                      vWhereBI + //
                      ' GROUP BY ' + //
                      '   CURVA ' + //
                      ' ORDER BY CURVA ';
    QryABC.Active := True;
    (QryABC.FieldByName('TOTAL_VENDAS'           ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryABC.FieldByName('VENDA_PCT'              ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QryABC.FieldByName('TOTAL_CMV'              ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryABC.FieldByName('LUCRO_BRUTO_RS'         ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryABC.FieldByName('LUCRO_BRUTO_PCT'        ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QryABC.FieldByName('VALOR_DESPESA_ADM_CX'   ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryABC.FieldByName('PCT_DESPESA_ADM_CX'     ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryABC.FieldByName('VALOR_LUCRO_LIQUIDO'    ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryABC.FieldByName('VALOR_LUCRO_LIQUIDO_PCT') as TFloatField).DisplayFormat :='#,##0.00%';
    (QryABC.FieldByName('TOTAL_QTD_PRODUTOS'     ) as TFloatField).DisplayFormat :='#0';
    (QryABC.FieldByName('TOTAL_ACRESCIMO'        ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryABC.FieldByName('TOTAL_ACRESCIMO_PCT'    ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QryABC.FieldByName('TOTAL_DESCONTO'         ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryABC.FieldByName('TOTAL_DESCONTO_PCT'     ) as TFloatField).DisplayFormat :='#,##0.00%';

    QryProdutos.Active := False;
    QryProdutos.SQL.Text := ' SELECT   ' + //
                            '   ID_PRODUTO, ' + //
                            '   PRODUTO, ' + //
                            '   REFERENCIA, ' + //
                            '   CURVA, ' + //
                            '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS), 0) AS TOTAL_VENDAS, ' + //
                            '  COALESCE (SUM (TOTAL_ACRESCIMO), 0) AS TOTAL_ACRESCIMO, ' + //
                            '  COALESCE (SUM (TOTAL_DESCONTO), 0) AS TOTAL_DESCONTO, ' + //
                            '  COALESCE ((SUM (TOTAL_ACRESCIMO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_ACRESCIMO) , 0) AS TOTAL_ACRESCIMO_PCT, ' + //
                            '  COALESCE ((SUM (TOTAL_DESCONTO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_DESCONTO) , 0) AS TOTAL_DESCONTO_PCT, ' + //
                            '  COALESCE (COUNT ( DISTINCT(ID_PEDIDO) ), 0 ) AS TOTAL_QTD_PEDIDOS, ' + //
                            '  COALESCE (SUM (QUANTIDADE / NUMPAGS), 0) AS TOTAL_QTD_PRODUTOS, ' + //
                            '  COALESCE (((SUM (TOTAL_VENDAS_PRODUTOS)) * 100 / ' + uFuncoes.FB_Numeric(TOTAL_VENDAS, 15, 2) + '), 0) AS VENDA_PCT, ' + //
                            '  COALESCE (SUM (TOTAL_CMV_PRODUTOS), 0) AS TOTAL_CMV, ' + //
                            '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS), 0) AS LUCRO_BRUTO_RS, ' + //
                            '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS LUCRO_BRUTO_PCT, ' + //
                            '  COALESCE (SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_DESPESA_ADM_CX, ' + //
                            '  COALESCE (SUM (VALOR_DESPESA_ADM_CX) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS PCT_DESPESA_ADM_CX, ' + //
                            '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_LUCRO_LIQUIDO, ' + //
                            '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS VALOR_LUCRO_LIQUIDO_PCT ' + //
                            ' FROM FDMTVendas ' + //
                            ' WHERE 1 = 1 ' + //
                            vWhereBI + //
                            ' GROUP BY ' + //
                            '   ID_PRODUTO, ' + //
                            '   PRODUTO, ' + //
                            '   REFERENCIA, ' + //
                            '   CURVA ' + //
                            ' ORDER BY TOTAL_VENDAS DESC ';
    QryProdutos.Active := True;
    (QryProdutos.FieldByName('TOTAL_VENDAS'           ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryProdutos.FieldByName('VENDA_PCT'              ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QryProdutos.FieldByName('TOTAL_CMV'              ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryProdutos.FieldByName('LUCRO_BRUTO_RS'         ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryProdutos.FieldByName('LUCRO_BRUTO_PCT'        ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QryProdutos.FieldByName('VALOR_DESPESA_ADM_CX'   ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryProdutos.FieldByName('PCT_DESPESA_ADM_CX'     ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryProdutos.FieldByName('VALOR_LUCRO_LIQUIDO'    ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryProdutos.FieldByName('VALOR_LUCRO_LIQUIDO_PCT') as TFloatField).DisplayFormat :='#,##0.00%';
    (QryProdutos.FieldByName('TOTAL_QTD_PRODUTOS'     ) as TFloatField).DisplayFormat :='#0';
    (QryProdutos.FieldByName('TOTAL_ACRESCIMO'        ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryProdutos.FieldByName('TOTAL_ACRESCIMO_PCT'    ) as TFloatField).DisplayFormat :='#,##0.00%';
    (QryProdutos.FieldByName('TOTAL_DESCONTO'         ) as TFloatField).DisplayFormat :='#,##0.00';
    (QryProdutos.FieldByName('TOTAL_DESCONTO_PCT'     ) as TFloatField).DisplayFormat :='#,##0.00%';

    // financeiro
    qryFinanConta.Active := False;
    qryFinanConta.SQL.Text := ' SELECT   ' + //
                              '   ID_CONTA, ' + //
                              '   CONTA, ' + //
                              '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS), 0) AS TOTAL_VENDAS, ' + //
                              '  COALESCE (SUM (TOTAL_ACRESCIMO), 0) AS TOTAL_ACRESCIMO, ' + //
                              '  COALESCE (SUM (TOTAL_DESCONTO), 0) AS TOTAL_DESCONTO, ' + //
                              '  COALESCE ((SUM (TOTAL_ACRESCIMO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_ACRESCIMO) , 0) AS TOTAL_ACRESCIMO_PCT, ' + //
                              '  COALESCE ((SUM (TOTAL_DESCONTO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_DESCONTO) , 0) AS TOTAL_DESCONTO_PCT, ' + //
                              '  COALESCE (COUNT ( DISTINCT(ID_PEDIDO) ), 0 ) AS TOTAL_QTD_PEDIDOS, ' + //
                              '  COALESCE (SUM (QUANTIDADE / NUMPAGS), 0) AS TOTAL_QTD_PRODUTOS, ' + //
                              '  COALESCE (((SUM (TOTAL_VENDAS_PRODUTOS)) * 100 / ' + uFuncoes.FB_Numeric(TOTAL_VENDAS, 15, 2) + '), 0) AS VENDA_PCT, ' + //
                              '  COALESCE (SUM (TOTAL_CMV_PRODUTOS), 0) AS TOTAL_CMV, ' + //
                              '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS), 0) AS LUCRO_BRUTO_RS, ' + //
                              '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS LUCRO_BRUTO_PCT, ' + //
                              '  COALESCE (SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_DESPESA_ADM_CX, ' + //
                              '  COALESCE (SUM (VALOR_DESPESA_ADM_CX) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS PCT_DESPESA_ADM_CX, ' + //
                              '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_LUCRO_LIQUIDO, ' + //
                              '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS VALOR_LUCRO_LIQUIDO_PCT ' + //
                              ' FROM FDMTVendas ' + //
                              ' WHERE 1 = 1 ' + //
                              vWhereBI + //
                              ' GROUP BY ' + //
                              '   ID_CONTA, ' + //
                              '   CONTA ' + //
                              ' ORDER BY TOTAL_VENDAS DESC ';
    qryFinanConta.Active := True;
    (qryFinanConta.FieldByName('TOTAL_VENDAS'           ) as TFloatField).DisplayFormat :='#,##0.00';
    (qryFinanConta.FieldByName('VENDA_PCT'              ) as TFloatField).DisplayFormat :='#,##0.00%';
    (qryFinanConta.FieldByName('TOTAL_CMV'              ) as TFloatField).DisplayFormat :='#,##0.00';
    (qryFinanConta.FieldByName('LUCRO_BRUTO_RS'         ) as TFloatField).DisplayFormat :='#,##0.00';
    (qryFinanConta.FieldByName('LUCRO_BRUTO_PCT'        ) as TFloatField).DisplayFormat :='#,##0.00%';
    (qryFinanConta.FieldByName('VALOR_DESPESA_ADM_CX'   ) as TFloatField).DisplayFormat :='#,##0.00';
    (qryFinanConta.FieldByName('PCT_DESPESA_ADM_CX'     ) as TFloatField).DisplayFormat :='#,##0.00';
    (qryFinanConta.FieldByName('VALOR_LUCRO_LIQUIDO'    ) as TFloatField).DisplayFormat :='#,##0.00';
    (qryFinanConta.FieldByName('VALOR_LUCRO_LIQUIDO_PCT') as TFloatField).DisplayFormat :='#,##0.00%';
    (qryFinanConta.FieldByName('TOTAL_QTD_PRODUTOS'     ) as TFloatField).DisplayFormat :='#0';
    (qryFinanConta.FieldByName('TOTAL_ACRESCIMO'        ) as TFloatField).DisplayFormat :='#,##0.00';
    (qryFinanConta.FieldByName('TOTAL_ACRESCIMO_PCT'    ) as TFloatField).DisplayFormat :='#,##0.00%';
    (qryFinanConta.FieldByName('TOTAL_DESCONTO'         ) as TFloatField).DisplayFormat :='#,##0.00';
    (qryFinanConta.FieldByName('TOTAL_DESCONTO_PCT'     ) as TFloatField).DisplayFormat :='#,##0.00%';

    qryFinanHistorico.Active := False;
    qryFinanHistorico.SQL.Text := ' SELECT   ' + //
                                  '   ID_HISTORICO, ' + //
                                  '   HISTORICO, ' + //
                                  '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS), 0) AS TOTAL_VENDAS, ' + //
                                  '  COALESCE (SUM (TOTAL_ACRESCIMO), 0) AS TOTAL_ACRESCIMO, ' + //
                                  '  COALESCE (SUM (TOTAL_DESCONTO), 0) AS TOTAL_DESCONTO, ' + //
                                  '  COALESCE ((SUM (TOTAL_ACRESCIMO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_ACRESCIMO) , 0) AS TOTAL_ACRESCIMO_PCT, ' + //
                                  '  COALESCE ((SUM (TOTAL_DESCONTO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_DESCONTO) , 0) AS TOTAL_DESCONTO_PCT, ' + //
                                  '  COALESCE (COUNT ( DISTINCT(ID_PEDIDO) ), 0 ) AS TOTAL_QTD_PEDIDOS, ' + //
                                  '  COALESCE (SUM (QUANTIDADE / NUMPAGS), 0) AS TOTAL_QTD_PRODUTOS, ' + //
                                  '  COALESCE (((SUM (TOTAL_VENDAS_PRODUTOS)) * 100 / ' + uFuncoes.FB_Numeric(TOTAL_VENDAS, 15, 2) + '), 0) AS VENDA_PCT, ' + //
                                  '  COALESCE (SUM (TOTAL_CMV_PRODUTOS), 0) AS TOTAL_CMV, ' + //
                                  '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS), 0) AS LUCRO_BRUTO_RS, ' + //
                                  '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS LUCRO_BRUTO_PCT, ' + //
                                  '  COALESCE (SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_DESPESA_ADM_CX, ' + //
                                  '  COALESCE (SUM (VALOR_DESPESA_ADM_CX) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS PCT_DESPESA_ADM_CX, ' + //
                                  '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_LUCRO_LIQUIDO, ' + //
                                  '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS VALOR_LUCRO_LIQUIDO_PCT ' + //
                                  ' FROM FDMTVendas ' + //
                                  ' WHERE 1 = 1 ' + //
                                  vWhereBI + //
                                  ' GROUP BY ' + //
                                  '   ID_HISTORICO, ' + //
                                  '   HISTORICO ' + //
                                  ' ORDER BY TOTAL_VENDAS DESC ';
    qryFinanHistorico.Active := True;
    (qryFinanHistorico.FieldByName('TOTAL_VENDAS'           ) as TFloatField).DisplayFormat := '#,##0.00';
    (qryFinanHistorico.FieldByName('VENDA_PCT'              ) as TFloatField).DisplayFormat := '#,##0.00%';
    (qryFinanHistorico.FieldByName('TOTAL_CMV'              ) as TFloatField).DisplayFormat := '#,##0.00';
    (qryFinanHistorico.FieldByName('LUCRO_BRUTO_RS'         ) as TFloatField).DisplayFormat := '#,##0.00';
    (qryFinanHistorico.FieldByName('LUCRO_BRUTO_PCT'        ) as TFloatField).DisplayFormat := '#,##0.00%';
    (qryFinanHistorico.FieldByName('VALOR_DESPESA_ADM_CX'   ) as TFloatField).DisplayFormat := '#,##0.00';
    (qryFinanHistorico.FieldByName('PCT_DESPESA_ADM_CX'     ) as TFloatField).DisplayFormat := '#,##0.00';
    (qryFinanHistorico.FieldByName('VALOR_LUCRO_LIQUIDO'    ) as TFloatField).DisplayFormat := '#,##0.00';
    (qryFinanHistorico.FieldByName('VALOR_LUCRO_LIQUIDO_PCT') as TFloatField).DisplayFormat := '#,##0.00%';
    (qryFinanHistorico.FieldByName('TOTAL_QTD_PRODUTOS'     ) as TFloatField).DisplayFormat := '#0';
    (qryFinanHistorico.FieldByName('TOTAL_ACRESCIMO'        ) as TFloatField).DisplayFormat := '#,##0.00';
    (qryFinanHistorico.FieldByName('TOTAL_ACRESCIMO_PCT'    ) as TFloatField).DisplayFormat := '#,##0.00%';
    (qryFinanHistorico.FieldByName('TOTAL_DESCONTO'         ) as TFloatField).DisplayFormat := '#,##0.00';
    (qryFinanHistorico.FieldByName('TOTAL_DESCONTO_PCT'     ) as TFloatField).DisplayFormat := '#,##0.00%';

    qryFinanPlanoVenda.Active := False;
    qryFinanPlanoVenda.SQL.Text :=' SELECT   ' + //
                                  '   ID_PLANO, ' + //
                                  '   PLANO_VENDA, ' + //
                                  '   QP, ' + //
                                  '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS), 0) AS TOTAL_VENDAS, ' + //
                                  '  COALESCE (SUM (TOTAL_ACRESCIMO), 0) AS TOTAL_ACRESCIMO, ' + //
                                  '  COALESCE (SUM (TOTAL_DESCONTO), 0) AS TOTAL_DESCONTO, ' + //
                                  '  COALESCE ((SUM (TOTAL_ACRESCIMO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_ACRESCIMO) , 0) AS TOTAL_ACRESCIMO_PCT, ' + //
                                  '  COALESCE ((SUM (TOTAL_DESCONTO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_DESCONTO) , 0) AS TOTAL_DESCONTO_PCT, ' + //
                                  '  COALESCE (COUNT ( DISTINCT(ID_PEDIDO) ), 0 ) AS TOTAL_QTD_PEDIDOS, ' + //
                                  '  COALESCE (SUM (QUANTIDADE / NUMPAGS), 0) AS TOTAL_QTD_PRODUTOS, ' + //
                                  '  COALESCE (((SUM (TOTAL_VENDAS_PRODUTOS)) * 100 / ' + uFuncoes.FB_Numeric(TOTAL_VENDAS, 15, 2) + '), 0) AS VENDA_PCT, ' + //
                                  '  COALESCE (SUM (TOTAL_CMV_PRODUTOS), 0) AS TOTAL_CMV, ' + //
                                  '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS), 0) AS LUCRO_BRUTO_RS, ' + //
                                  '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS LUCRO_BRUTO_PCT, ' + //
                                  '  COALESCE (SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_DESPESA_ADM_CX, ' + //
                                  '  COALESCE (SUM (VALOR_DESPESA_ADM_CX) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS PCT_DESPESA_ADM_CX, ' + //
                                  '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_LUCRO_LIQUIDO, ' + //
                                  '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS VALOR_LUCRO_LIQUIDO_PCT ' + //
                                  ' FROM FDMTVendas ' + //
                                  ' WHERE 1 = 1 ' + //
                                  vWhereBI + //
                                  ' GROUP BY ' + //
                                  '   ID_PLANO, ' + //
                                  '   PLANO_VENDA, ' + //
                                  '   QP ' + //
                                  ' ORDER BY TOTAL_VENDAS DESC ';
    qryFinanPlanoVenda.Active := True;
    (qryFinanPlanoVenda.FieldByName('TOTAL_VENDAS'           ) as TFloatField).DisplayFormat := '#,##0.00';
    (qryFinanPlanoVenda.FieldByName('VENDA_PCT'              ) as TFloatField).DisplayFormat := '#,##0.00%';
    (qryFinanPlanoVenda.FieldByName('TOTAL_CMV'              ) as TFloatField).DisplayFormat := '#,##0.00';
    (qryFinanPlanoVenda.FieldByName('LUCRO_BRUTO_RS'         ) as TFloatField).DisplayFormat := '#,##0.00';
    (qryFinanPlanoVenda.FieldByName('LUCRO_BRUTO_PCT'        ) as TFloatField).DisplayFormat := '#,##0.00%';
    (qryFinanPlanoVenda.FieldByName('VALOR_DESPESA_ADM_CX'   ) as TFloatField).DisplayFormat := '#,##0.00';
    (qryFinanPlanoVenda.FieldByName('PCT_DESPESA_ADM_CX'     ) as TFloatField).DisplayFormat := '#,##0.00';
    (qryFinanPlanoVenda.FieldByName('VALOR_LUCRO_LIQUIDO'    ) as TFloatField).DisplayFormat := '#,##0.00';
    (qryFinanPlanoVenda.FieldByName('VALOR_LUCRO_LIQUIDO_PCT') as TFloatField).DisplayFormat := '#,##0.00%';
    (qryFinanPlanoVenda.FieldByName('TOTAL_QTD_PRODUTOS'     ) as TFloatField).DisplayFormat := '#0';
    (qryFinanPlanoVenda.FieldByName('TOTAL_ACRESCIMO'        ) as TFloatField).DisplayFormat := '#,##0.00';
    (qryFinanPlanoVenda.FieldByName('TOTAL_ACRESCIMO_PCT'    ) as TFloatField).DisplayFormat := '#,##0.00%';
    (qryFinanPlanoVenda.FieldByName('TOTAL_DESCONTO'         ) as TFloatField).DisplayFormat := '#,##0.00';
    (qryFinanPlanoVenda.FieldByName('TOTAL_DESCONTO_PCT'     ) as TFloatField).DisplayFormat := '#,##0.00%';

    qryFinanTabelaPreco.Active := False;
    qryFinanTabelaPreco.SQL.Text := ' SELECT   ' + //
                                    '  TABELA_PRECO, ' + //
                                    '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS), 0) AS TOTAL_VENDAS, ' + //
                                    '  COALESCE (SUM (TOTAL_ACRESCIMO), 0) AS TOTAL_ACRESCIMO, ' + //
                                    '  COALESCE (SUM (TOTAL_DESCONTO), 0) AS TOTAL_DESCONTO, ' + //
                                    '  COALESCE ((SUM (TOTAL_ACRESCIMO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_ACRESCIMO) , 0) AS TOTAL_ACRESCIMO_PCT, ' + //
                                    '  COALESCE ((SUM (TOTAL_DESCONTO) * 100 ) / SUM (TOTAL_VENDAS_PRODUTOS+TOTAL_DESCONTO) , 0) AS TOTAL_DESCONTO_PCT, ' + //
                                    '  COALESCE (COUNT ( DISTINCT(ID_PEDIDO) ), 0 ) AS TOTAL_QTD_PEDIDOS, ' + //
                                    '  COALESCE (SUM (QUANTIDADE / NUMPAGS), 0) AS TOTAL_QTD_PRODUTOS, ' + //
                                    '  COALESCE (((SUM (TOTAL_VENDAS_PRODUTOS)) * 100 / ' + uFuncoes.FB_Numeric(TOTAL_VENDAS, 15, 2) + '), 0) AS VENDA_PCT, ' + //
                                    '  COALESCE (SUM (TOTAL_CMV_PRODUTOS), 0) AS TOTAL_CMV, ' + //
                                    '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS), 0) AS LUCRO_BRUTO_RS, ' + //
                                    '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS LUCRO_BRUTO_PCT, ' + //
                                    '  COALESCE (SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_DESPESA_ADM_CX, ' + //
                                    '  COALESCE (SUM (VALOR_DESPESA_ADM_CX) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS PCT_DESPESA_ADM_CX, ' + //
                                    '  COALESCE (SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX), 0) AS VALOR_LUCRO_LIQUIDO, ' + //
                                    '  COALESCE ((SUM (TOTAL_VENDAS_PRODUTOS) - SUM (TOTAL_CMV_PRODUTOS) - SUM (VALOR_DESPESA_ADM_CX)) * 100 / SUM (TOTAL_VENDAS_PRODUTOS), 0) AS VALOR_LUCRO_LIQUIDO_PCT ' + //
                                    ' FROM FDMTVendas ' + //
                                    ' WHERE 1 = 1 ' + //
                                    vWhereBI + //
                                    ' GROUP BY ' + //
                                    '   TABELA_PRECO ' + //
                                    ' ORDER BY TOTAL_VENDAS DESC ';
    qryFinanTabelaPreco.Active := True;

    (qryFinanTabelaPreco.FieldByName('TOTAL_VENDAS'           ) as TFloatField).DisplayFormat := '#,##0.00';
    (qryFinanTabelaPreco.FieldByName('VENDA_PCT'              ) as TFloatField).DisplayFormat := '#,##0.00%';
    (qryFinanTabelaPreco.FieldByName('TOTAL_CMV'              ) as TFloatField).DisplayFormat := '#,##0.00';
    (qryFinanTabelaPreco.FieldByName('LUCRO_BRUTO_RS'         ) as TFloatField).DisplayFormat := '#,##0.00';
    (qryFinanTabelaPreco.FieldByName('LUCRO_BRUTO_PCT'        ) as TFloatField).DisplayFormat := '#,##0.00%';
    (qryFinanTabelaPreco.FieldByName('VALOR_DESPESA_ADM_CX'   ) as TFloatField).DisplayFormat := '#,##0.00';
    (qryFinanTabelaPreco.FieldByName('PCT_DESPESA_ADM_CX'     ) as TFloatField).DisplayFormat := '#,##0.00';
    (qryFinanTabelaPreco.FieldByName('VALOR_LUCRO_LIQUIDO'    ) as TFloatField).DisplayFormat := '#,##0.00';
    (qryFinanTabelaPreco.FieldByName('VALOR_LUCRO_LIQUIDO_PCT') as TFloatField).DisplayFormat := '#,##0.00%';
    (qryFinanTabelaPreco.FieldByName('TOTAL_QTD_PRODUTOS'     ) as TFloatField).DisplayFormat := '#0';
    (qryFinanTabelaPreco.FieldByName('TOTAL_ACRESCIMO'        ) as TFloatField).DisplayFormat := '#,##0.00';
    (qryFinanTabelaPreco.FieldByName('TOTAL_ACRESCIMO_PCT'    ) as TFloatField).DisplayFormat := '#,##0.00%';
    (qryFinanTabelaPreco.FieldByName('TOTAL_DESCONTO'         ) as TFloatField).DisplayFormat := '#,##0.00';
    (qryFinanTabelaPreco.FieldByName('TOTAL_DESCONTO_PCT'     ) as TFloatField).DisplayFormat := '#,##0.00%';

    ControlsCDS(True);

    for I := 0 to FScrollBox.ComponentCount -1 do
    begin
      if (FScrollBox.Components[I] is TFFormGrafico) then
      begin
        for J := 0 to Self.ComponentCount -1 do
        begin
          if (Self.Components[J] is TAction) and ((Self.Components[J] as TAction).Tag > 0) then
          begin
            if (Self.Components[J] as TAction).Tag = (FScrollBox.Components[I] as TFFormGrafico).Tag then
            begin
              (Self.Components[J] as TAction).Execute;
            end;
          end;
        end;
      end;
    end;

    Aguarde(self, False);
  except
    on Exc: Exception do
    begin
      Aguarde(self, False);
      Mensagem(self, 'Erro. ' + #13#10 + Exc.Message);
    end;
  end;
end;

procedure TfBIVendas.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  FreeAndNil(fBuscaGenerica);
  dm.UCIdle1.UserControl := dm.UserControl1;
  DestruirQuerys;
end;

procedure TfBIVendas.BuscaChangeFilter(Sender: TObject; ChangerType: TChangerType);
begin
  case TBuscaGenerica(Sender).LastChange of
    lcAddFilter: ;

    lcDelFilter:
      if ChangerType = TChangerType.ctManual then
      begin
        vWhereBI := fBuscaGenerica.getSql(vWhereBI);
        FiltroBI();
      end;
  end;
end;

procedure TfBIVendas.FormCreate(Sender: TObject);
begin
  DefineIEVersion;

  fBuscaGenerica                    := TBuscaGenerica.Create(self);
  fBuscaGenerica.PnlFiltroPesquisa  := ScrollBox;
  fBuscaGenerica.OnChangeFilter     := BuscaChangeFilter;
  ScrollBox1.VertScrollBar.Position := 0;

  ConfiguraGrid;
  CriarQuerys;

  dm.UCIdle1.UserControl := nil;
end;

procedure TfBIVendas.FormMouseWheelDown(Sender: TObject; Shift: TShiftState; MousePos: TPoint; var Handled: Boolean);
begin
  with ScrollBox1.VertScrollBar do
  begin
    if (Position <= (Range - Increment)) then
      Position := Position + Increment
    else
      Position := Range - Increment;
  end;
end;

procedure TfBIVendas.FormMouseWheelUp(Sender: TObject; Shift: TShiftState; MousePos: TPoint; var Handled: Boolean);
begin
  with ScrollBox1.VertScrollBar do
  begin
    if (Position >= Increment) then
      Position := Position - Increment
    else
      Position := 0;
  end;
end;

procedure TfBIVendas.FormShow(Sender: TObject);
begin
  dpkDataIni.Date := vdatahoje_dm;
  dpkDataFim.Date := vdatahoje_dm;

  PageControl1.ActivePageIndex := 0;
  btnMenu.Enabled := False;
end;

procedure TfBIVendas.CallBackRegiao(value: string);
begin
  if not fBuscaGenerica.clearFilter('ID_REGIAO') then
  begin
    fBuscaGenerica.addFilter('', 'ID_REGIAO', '=', //
      value, //
      '', //
      value, // top
      'Regional');

    vWhereBI := fBuscaGenerica.getSql(vWhereBI);
    FiltroBI;
  end;
end;

procedure TfBIVendas.CallBackGrafico(Params: TStringList);
var
  lData: TStringList;
  lTable, lField, lId: string;
begin
  if Assigned(Params) and not Trim(Params.Text).IsEmpty then
  begin
    lData      := Params;
    lData.Text := ReplaceStr(lData.Text, '[', '');
    lData.Text := ReplaceStr(lData.Text, ']', '');
    lData.Text := ReplaceStr(lData.Text, '%20', ' ');

    lTable := ReplaceStr(lData[2], '_', ' ');
    lField := lData[3];
    lId    := lData[4];

    // Remove o filtro caso clique sobre um grafico já filtrado
    // e add um filtro caso não tenha sido filtrado ainda
    if not fBuscaGenerica.clearFilter(lField) then
    begin
      fBuscaGenerica.addFilter(//
        lField, '=', lId, TitleCase(lTable));

      vWhereBI := fBuscaGenerica.getSql(vWhereBI);
      FiltroBI;
    end;
  end;
end;

procedure TfBIVendas.CallBackLoja(value: string);
begin
  if not fBuscaGenerica.clearFilter('ID_LOJA') then
  begin
    fBuscaGenerica.addFilter('', 'ID_LOJA', '=', //
      value, //
      '', //
      value, // top
      'Loja');

    vWhereBI := fBuscaGenerica.getSql(vWhereBI);
    FiltroBI;
  end;
end;

procedure TfBIVendas.gridVendedorTitleClick(Column: TColumn);
begin
  OrdenaGrid(Column);
end;

procedure TfBIVendas.gridFinanContaTitleClick(Column: TColumn);
begin
  OrdenaGrid(Column);
end;

procedure TfBIVendas.gridRegiaoTitleClick(Column: TColumn);
begin
  OrdenaGrid(Column);
end;

procedure TfBIVendas.gridLojaTitleClick(Column: TColumn);
begin
  OrdenaGrid(Column);
end;

procedure TfBIVendas.gridCidadeClientesTitleClick(Column: TColumn);
begin
  OrdenaGrid(Column);
end;

procedure TfBIVendas.gridCidadeLojasTitleClick(Column: TColumn);
begin
  OrdenaGrid(Column);
end;

procedure TfBIVendas.gridGrupoTitleClick(Column: TColumn);
begin
  OrdenaGrid(Column);
end;

procedure TfBIVendas.gridSubGrupoTitleClick(Column: TColumn);
begin
  OrdenaGrid(Column);
end;

procedure TfBIVendas.gridFornecedorTitleClick(Column: TColumn);
begin
  OrdenaGrid(Column);
end;

procedure TfBIVendas.gridMarcaTitleClick(Column: TColumn);
begin
  OrdenaGrid(Column);
end;

procedure TfBIVendas.gridPublicidadeTitleClick(Column: TColumn);
begin
  OrdenaGrid(Column);
end;

procedure TfBIVendas.gridDiaSemanaTitleClick(Column: TColumn);
begin
  OrdenaGrid(Column);
end;

procedure TfBIVendas.gridHoraTitleClick(Column: TColumn);
begin
  OrdenaGrid(Column);
end;

procedure TfBIVendas.gridABCTitleClick(Column: TColumn);
begin
  OrdenaGrid(Column);
end;

procedure TfBIVendas.gridProdutosTitleClick(Column: TColumn);
begin
  OrdenaGrid(Column);
end;

procedure TfBIVendas.gridFinanHistoricoTitleClick(Column: TColumn);
begin
  OrdenaGrid(Column);
end;

procedure TfBIVendas.gridFinanTabelaPrecoTitleClick(Column: TColumn);
begin
  OrdenaGrid(Column);
end;

procedure TfBIVendas.gridFinanPlanoVendaTitleClick(Column: TColumn);
begin
  OrdenaGrid(Column);
end;

procedure TfBIVendas.Image1Click(Sender: TObject);
begin
  try
    GerarExcelCds(self, QryRegiao, 'Vendas por Região');
  except
    on Exc: Exception do
    begin
      Aguarde(self, False);
      Mensagem(self, 'Erro. ' + #13 + Exc.Message);
    end;
  end;
end;

procedure TfBIVendas.Image2Click(Sender: TObject);
begin
  try
    GerarExcelCds(self, QryLoja, 'Vendas por Loja');
  except
    on Exc: Exception do
    begin
      Aguarde(self, False);
      Mensagem(self, 'Erro. ' + #13 + Exc.Message);
    end;
  end;
end;

procedure TfBIVendas.Image3Click(Sender: TObject);
begin
  try
    GerarExcelCds(self, QryCidadeLoja, 'Vendas por Cidades de Lojas');
  except
    on Exc: Exception do
    begin
      Aguarde(self, False);
      Mensagem(self, 'Erro. ' + #13 + Exc.Message);
    end;
  end;
end;

procedure TfBIVendas.Image4Click(Sender: TObject);
begin
  try
    GerarExcelCds(self, QryCidadeClientes, 'Vendas por Cidades de Cliente');
  except
    on Exc: Exception do
    begin
      Aguarde(self, False);
      Mensagem(self, 'Erro. ' + #13 + Exc.Message);
    end;
  end;
end;

procedure TfBIVendas.Image5Click(Sender: TObject);
begin
  try
    GerarExcelCds(self, QryGrupo, 'Vendas por Grupo');
  except
    on Exc: Exception do
    begin
      Aguarde(self, False);
      Mensagem(self, 'Erro. ' + #13 + Exc.Message);
    end;
  end;
end;

procedure TfBIVendas.Image6Click(Sender: TObject);
begin
  try
    GerarExcelCds(self, QrySubGrupo, 'Vendas por SubGrupo');
  except
    on Exc: Exception do
    begin
      Aguarde(self, False);
      Mensagem(self, 'Erro. ' + #13 + Exc.Message);
    end;
  end;
end;

procedure TfBIVendas.Image7Click(Sender: TObject);
begin
  try
    GerarExcelCds(self, QryFornecedor, 'Vendas por Fornecedor');
  except
    on Exc: Exception do
    begin
      Aguarde(self, False);
      Mensagem(self, 'Erro. ' + #13 + Exc.Message);
    end;
  end;
end;

procedure TfBIVendas.Image8Click(Sender: TObject);
begin
  try
    GerarExcelCds(self, QryMarca, 'Vendas por Marca');
  except
    on Exc: Exception do
    begin
      Aguarde(self, False);
      Mensagem(self, 'Erro. ' + #13 + Exc.Message);
    end;
  end;
end;

procedure TfBIVendas.Image9Click(Sender: TObject);
begin
  try
    GerarExcelCds(self, QryVendedor, 'Vendas por Vendedor');
  except
    on Exc: Exception do
    begin
      Aguarde(self, False);
      Mensagem(self, 'Erro. ' + #13 + Exc.Message);
    end;
  end;
end;

procedure TfBIVendas.Image10Click(Sender: TObject);
begin
  try
    GerarExcelCds(self, QryPublicidade, 'Vendas por Publicidade');
  except
    on Exc: Exception do
    begin
      Aguarde(self, False);
      Mensagem(self, 'Erro. ' + #13 + Exc.Message);
    end;
  end;
end;

procedure TfBIVendas.Image11Click(Sender: TObject);
begin
  try
    GerarExcelCds(self, QryDiaSemana, 'Vendas por Dia da Semana');
  except
    on Exc: Exception do
    begin
      Aguarde(self, False);
      Mensagem(self, 'Erro. ' + #13 + Exc.Message);
    end;
  end;
end;

procedure TfBIVendas.Image13Click(Sender: TObject);
begin
  try
    GerarExcelCds(self, QryABC, 'Vendas por Curva ABC');
  except
    on Exc: Exception do
    begin
      Aguarde(self, False);
      Mensagem(self, 'Erro. ' + #13 + Exc.Message);
    end;
  end;
end;

procedure TfBIVendas.Image12Click(Sender: TObject);
begin
  try
    GerarExcelCds(self, QryHora, 'Vendas por Horário');
  except
    on Exc: Exception do
    begin
      Aguarde(self, False);
      Mensagem(self, 'Erro. ' + #13 + Exc.Message);
    end;
  end;
end;

procedure TfBIVendas.Image14Click(Sender: TObject);
begin
  try
    GerarExcelCds(self, QryProdutos, 'Vendas por Produtos');
  except
    on Exc: Exception do
    begin
      Aguarde(self, False);
      Mensagem(self, 'Erro. ' + #13 + Exc.Message);
    end;
  end;
end;

procedure TfBIVendas.Image15Click(Sender: TObject);
begin
  try
    GerarExcelCds(self, qryFinanPlanoVenda, 'Vendas por Plano de Venda');
  except
    on Exc: Exception do
    begin
      Aguarde(self, False);
      Mensagem(self, 'Erro. ' + #13 + Exc.Message);
    end;
  end;
end;

procedure TfBIVendas.Image16Click(Sender: TObject);
begin
  try
    GerarExcelCds(self, qryFinanConta, 'Vendas por Contas');
  except
    on Exc: Exception do
    begin
      Aguarde(self, False);
      Mensagem(self, 'Erro. ' + #13 + Exc.Message);
    end;
  end;
end;

procedure TfBIVendas.Image17Click(Sender: TObject);
begin
  try
    GerarExcelCds(self, qryFinanHistorico, 'Vendas por Históricos');
  except
    on Exc: Exception do
    begin
      Aguarde(self, False);
      Mensagem(self, 'Erro. ' + #13 + Exc.Message);
    end;
  end;
end;

procedure TfBIVendas.Image18Click(Sender: TObject);
begin
  try
    GerarExcelCds(self, qryFinanTabelaPreco, 'Vendas por Coluna de Preço');
  except
    on Exc: Exception do
    begin
      Aguarde(self, False);
      Mensagem(self, 'Erro. ' + #13 + Exc.Message);
    end;
  end;
end;

procedure TfBIVendas.Timer1Timer(Sender: TObject);
begin
  try
    Timer1.Enabled := False;
    ChecaDataAtlzAuto;
    if (chkAtualizarAutomatico.Checked) then
      ConsVendasGeral;
    Timer1.Enabled := True;
  except
    on Exc: Exception do
    begin
      Aguarde(self, False);
      Timer1.Enabled := True;
      Mensagem(self, 'Erro. ' + #13 + Exc.Message);
    end;
  end
end;

procedure TfBIVendas.ChecaDataAtlzAuto;
begin
  if chkAtualizarAutomatico.Checked then
    if (CompareDate(vdatahoje_dm, dpkDataIni.Date) <> 0) or (CompareDate(vdatahoje_dm, dpkDataFim.Date) <> 0) then
    begin
      chkAtualizarAutomatico.Checked := False;
      Mensagem(self, 'Atualização automática dos dados está disponível apenas com a data do dia');
    end;
end;

procedure TfBIVendas.dpkDataIniChange(Sender: TObject);
begin
  ChecaDataAtlzAuto;
end;

procedure TfBIVendas.dpkDataFimChange(Sender: TObject);
begin
  ChecaDataAtlzAuto;
end;

procedure TfBIVendas.DefineIEVersion;
const
  REG_KEY = 'Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION';
var
  Reg    : TRegistry;
  AppName: String;
begin
  AppName := ExtractFileName(ExtractFileName(ParamStr(0)));
  Reg     := nil;
  try
    Reg         := TRegistry.Create();
    Reg.RootKey := HKEY_CURRENT_USER;
    if Reg.OpenKey(REG_KEY, True) then
    begin
      Reg.WriteInteger(AppName, 11000);
      Reg.CloseKey;
    end;
  except
  end;
  if (Assigned(Reg)) then
    FreeAndNil(Reg);
end;

procedure TfBIVendas.actExecute(Sender: TObject);
begin
  case TAction(Sender).Tag of
    1:
      begin
        GeraGrafico(TTipoGrafico.ColumnChart, fBIVendas.FScrollBox, QryRegiao, 1, 20, //
          'REGIAO', 'Gráfico Regional', 'REGIAO', 'ID_REGIAO', //
          CallBackGrafico);
      end;
    2:
      begin
        GeraGrafico(TTipoGrafico.ColumnChart, fBIVendas.FScrollBox, QryLoja, 2, 10, //
          'LOJA', 'Gráfico de Lojas', 'LOJA', 'ID_LOJA', //
          CallBackGrafico);
      end;
    3:
      begin
        GeraGrafico(TTipoGrafico.ColumnChart, FScrollBox, QryCidadeClientes, 3, 20, //
          'CIDADE_CLIENTE', 'Cidade de Clientes', 'CIDADE_CLIENTE', 'ID_CIDADE_CLIENTE', //
          CallBackGrafico);
      end;
    4:
      begin
        GeraGrafico(TTipoGrafico.ColumnChart, FScrollBox, QryCidadeLoja, 4, 15, //
          'CIDADE_LOJA', 'Cidade de Lojas', 'CIDADE_LOJA', 'ID_CIDADE_LOJA', //
          CallBackGrafico);
      end;
    5:
      begin
        GeraGrafico(TTipoGrafico.ColumnChart, FScrollBox, QryGrupo, 5, 15, //
          'GRUPO', 'Grupo', 'GRUPO', 'ID_GRUPO', //
          CallBackGrafico);
      end;
    6:
      begin
        GeraGrafico(TTipoGrafico.ColumnChart, FScrollBox, QrySubGrupo, 6, 15, //
          'SUBGRUPO', 'Subgrupo', 'SUBGRUPO', 'ID_SUBGRUPO', //
          CallBackGrafico);
      end;
    7:
      begin
        GeraGrafico(TTipoGrafico.ColumnChart, FScrollBox, QryVendedor, 7, 15, //
          'VENDEDOR', 'Vendedor', 'VENDEDOR', 'ID_VENDEDOR', //
          CallBackGrafico);
      end;
    8:
      begin
        GeraGrafico(TTipoGrafico.ColumnChart, FScrollBox, QryMarca, 8, 15, //
          'MARCA', 'Marca', 'MARCA', 'ID_MARCA', //
          CallBackGrafico);
      end;
    9:
      begin
        GeraGrafico(TTipoGrafico.ColumnChart, FScrollBox, QryFornecedor, 9, 15, //
          'FORNECEDOR', 'Fornecedor', 'FORNECEDOR', 'ID_FORNECEDOR', //
          CallBackGrafico);
      end;
    10:
      begin
        GeraGrafico(TTipoGrafico.ColumnChart, FScrollBox, QryProdutos, 10, 15, //
          'PRODUTO', 'Produtos', 'PRODUTO', 'ID_PRODUTO', //
          CallBackGrafico);
      end;
  end;
end;


end.


