unit uConfigBrowser;

interface

uses
  Winapi.Windows, System.SysUtils, Win.Registry;

type
  TConfigBrowser = class
  public
    class procedure ConfigurarIEVersion;
  end;

implementation

uses uFuncoes;

class procedure TConfigBrowser.ConfigurarIEVersion;
const
  REG_KEY = 'Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION';
var
  Reg: TRegistry;
  AppName: string;
begin
  AppName := ExtractFileName(ParamStr(0));
  Reg := nil;

  {$IFDEF DEBUG}
  // Use a classe que contém TextoMemo_Form
    TextoMemo_Form('ConfigurarIEVersion - AppName: ' + AppName);
  {$ENDIF}

  try
    Reg := TRegistry.Create;
    Reg.RootKey := HKEY_CURRENT_USER;

    if Reg.OpenKey(REG_KEY, True) then
    begin
      // Valor 11000 corresponde ao IE11 Edge mode
      Reg.WriteInteger(AppName, 11000);
      Reg.CloseKey;

      {$IFDEF DEBUG}
        TextoMemo_Form('ConfigurarIEVersion - Registro configurado com sucesso');
      {$ENDIF}
    end
    else
    begin
      {$IFDEF DEBUG}
        TextoMemo_Form('ConfigurarIEVersion - Não foi possível abrir a chave do registro');
      {$ENDIF}
    end;
  except
    on E: Exception do
    begin
      {$IFDEF DEBUG}
        TextoMemo_Form('ConfigurarIEVersion - Erro: ' + E.Message);
      {$ENDIF}
    end;
  end;

  if Assigned(Reg) then
    Reg.Free;
end;

end.
