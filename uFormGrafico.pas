unit uFormGrafico;

interface

uses
  uFormBase,
  uGraficos,
  uGraficoGoogle,
  uGraficoPie,
  uGraficoBar,
  uGraficoColumnChart,
  //
  Winapi.Windows,
  Winapi.Messages,
  System.SysUtils,
  System.Variants,
  System.Classes,
  Vcl.Graphics,
  Vcl.Controls,
  Vcl.Forms,
  Vcl.Dialogs,
  Vcl.OleCtrls,
  Data.DB,
  SHDocVw,
  ACBrBase,
  ACBrEnterTab,
  dxGDIPlusClasses,
  Vcl.ExtCtrls,
  Vcl.StdCtrls;

type
  TTipoGrafico   = (Pie, PieDonut, Bar, BarHorz, ColumnChart);
  TArrayVariants = Array of Variant;

  TFFormGrafico = class(TfBase)
    FBrowser: TWebBrowser;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
  private
    FTipoGrafico: TTipoGrafico;
  public
    FGraficos: TGraficos;
    Grafico  : TGraficoGoogle;

    constructor Create(AOwner: TComponent; TipoGrafico: TTipoGrafico; Index: Integer = 0); reintroduce;
    destructor Destroy; override;

    procedure Inicialize();

    procedure AddTitle(ACaption: string; AColunaX: string = ''; AColunaY: string = '');

    procedure Configure(Titulo, Tab, ColId: string; CallBack: TGoogleCallBack);
    procedure AddDados(ADados: TArrayVariants);

    procedure CriaGrafico(Titulo: string; DataSet: TDataSet; Col, Val, Tab, ColId: string;
      CallBack: TGoogleCallBack); overload;
    procedure CriaGrafico(Titulo, ColX, ColY: string; DataSet: TDataSet; Col, Val, Tab, ColId: string;
      CallBack: TGoogleCallBack); overload;

    property TipoGrafico: TTipoGrafico read FTipoGrafico write FTipoGrafico;
  end;

implementation

uses
  uFuncoes;

{$R *.dfm}

procedure TFFormGrafico.AddTitle(ACaption: string; AColunaX: string = ''; AColunaY: string = '');
begin
  Self.Caption := ACaption;

  Grafico.AddTitle(ACaption);
  Grafico.DescricaoColunaX := AColunaX;
  Grafico.DescricaoColunaY := AColunaY;
end;

procedure TFFormGrafico.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
  Action := caFree;
end;

procedure TFFormGrafico.FormShow(Sender: TObject);
begin
  inherited;

  Inicialize();
end;

procedure TFFormGrafico.Inicialize();
begin
  FGraficos            := TGraficos.Create;
  FGraficos.WebBrowser := FBrowser;

  case FTipoGrafico of
    Pie:
      begin
        Grafico := TGraficoPie.Create;
        FGraficos.AddGrafico(Grafico as TGraficoPie);
      end;
    PieDonut:
      begin
        Grafico := TGraficoPie.Create(True);
        FGraficos.AddGrafico(Grafico as TGraficoPie);
      end;
    Bar:
      begin
        Grafico := TGraficoBar.Create;
        FGraficos.AddGrafico(Grafico as TGraficoBar);
      end;
    BarHorz:
      begin
        Grafico := TGraficoBar.Create(True);
        FGraficos.AddGrafico(Grafico as TGraficoBar);
      end;
    ColumnChart:
      begin
        Grafico := TGraficoColumnChart.Create(True);
        FGraficos.AddGrafico(Grafico as TGraficoColumnChart);
      end;
  end;

  Grafico.FRandomColor := True;
end;

constructor TFFormGrafico.Create(AOwner: TComponent; TipoGrafico: TTipoGrafico; Index: Integer = 0);
begin
  inherited Create(AOwner);

  FTipoGrafico := TipoGrafico;
  Self.Tag     := Index;
end;

destructor TFFormGrafico.Destroy;
begin
  inherited;
end;

procedure TFFormGrafico.Configure(Titulo, Tab, ColId: string; CallBack: TGoogleCallBack);
begin
  AddTitle(Titulo);

  if Assigned(CallBack) then
    Grafico.OnCallBack := CallBack;

  Grafico.LimpaDados;

  Grafico.AddDadosColunaVirtual(Tab, TTipoCampoVirtual.tcString);
  Grafico.AddDadosColunaVirtual(ColId, TTipoCampoVirtual.tcNumerico);

  Grafico.FieldCallBack := ColId;
end;

procedure TFFormGrafico.AddDados(ADados: TArrayVariants);
var
  Row: Integer;
begin
  Row := Length(ADados) + 1;
  SetLength(ADados, Row);
  ADados[High(ADados)] := 'random_color()';
  Grafico.AddDados(ADados);
end;

procedure TFFormGrafico.CriaGrafico(Titulo: string; DataSet: TDataSet; Col, Val, Tab, ColId: string;
  CallBack: TGoogleCallBack);
begin
  CriaGrafico(Titulo, '', '', DataSet, Col, Val, Tab, ColId, CallBack);
end;

procedure TFFormGrafico.CriaGrafico(Titulo, ColX, ColY: string; DataSet: TDataSet; Col, Val, Tab, ColId: string;
  CallBack: TGoogleCallBack);
var
  ARecNo: Integer;
begin
  AddTitle(Titulo, ColX, ColY);

  if Assigned(CallBack) then
    Grafico.OnCallBack := CallBack;

  if not DataSet.isEmpty then
  begin
    Grafico.LimpaDados;

    Grafico.AddDadosColunaVirtual(Tab, TTipoCampoVirtual.tcString);
    Grafico.AddDadosColunaVirtual(ColId, TTipoCampoVirtual.tcNumerico);

    ARecNo := DataSet.RecNo;
    DataSet.DisableControls;
    DataSet.First;

    while not DataSet.Eof do
    begin
      Application.ProcessMessages;

      Grafico.AddDados([ //
          FB_Texto(DataSet.FieldByName(Col).AsString), //
          FB_Numeric(DataSet.FieldByName(Val).AsFloat, 15, 2), //
          FB_Texto(Tab), //
          DataSet.FieldByName(ColId).AsInteger, //
          'random_color()' //
          ]);

      DataSet.Next;
    end;

    DataSet.RecNo := ARecNo;
    DataSet.EnableControls;
  end;
  FGraficos.Gerar;
end;

end.
