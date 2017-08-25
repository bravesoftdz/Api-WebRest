unit formMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls, ActiveX, IniFiles,
  IdBaseComponent, IdComponent, IdCustomTCPServer, IdCustomHTTPServer,
  IdHTTPServer, Vcl.AppEvnts, Vcl.Buttons, IdContext, ServerUtils,
  dxGDIPlusClasses;

type
  TfrmMain = class(TForm)
    apl1: TApplicationEvents;
    TrayIcon1: TTrayIcon;
    txtInfoLabel: TStaticText;
    lblNome: TLabel;
    btnAtivar: TBitBtn;
    btnParar: TBitBtn;
    memoReq: TMemo;
    memoResp: TMemo;
    Label1: TLabel;
    Label2: TLabel;
    Image1: TImage;
    vHost: TEdit;
    vBanco: TEdit;
    vUser: TEdit;
    vPass: TEdit;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    BitBtn1: TBitBtn;
    Label7: TLabel;
    vPorta: TEdit;
    IdHTTPServer1: TIdHTTPServer;
    procedure TrayIcon1Click(Sender: TObject);
    procedure apl1Minimize(Sender: TObject);
    procedure btnAtivarClick(Sender: TObject);
    procedure btnPararClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure IdHTTPServer1CommandGet(AContext: TIdContext;
      ARequestInfo: TIdHTTPRequestInfo; AResponseInfo: TIdHTTPResponseInfo);
    procedure IdHTTPServer1CommandOther(AContext: TIdContext;
      ARequestInfo: TIdHTTPRequestInfo; AResponseInfo: TIdHTTPResponseInfo);
    procedure FormDestroy(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);

  private
    { Private declarations }
    ServerParams : TServerParams;
    procedure LoglastRequest (ARequestInfo: TIdHTTPRequestInfo);
    procedure LogLastResponse (AResponseInfo: TIdHTTPResponseInfo);
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;
  CriticalSection: TRTLCriticalSection;

implementation

{$R *.dfm}

Uses
     Systypes, UntServerMethods, UntConexao, UntFuncoes;

const
  cBanco = 'VW_RTPRIME';
  cHost = '192.168.2.30';


procedure TfrmMain.FormCreate(Sender: TObject);
var
  arq : TIniFile;
begin
  arq := TIniFile.Create(ExtractFilePath(Application.ExeName) + 'RestConf.ini');
  IdHTTPServer1.DefaultPort := arq.ReadInteger('CONFIG' , 'PORTA' , 1997);
  vPorta.Text := IntToStr( IdHTTPServer1.DefaultPort );
  CoInitialize(nil);

  publicConexao := TConexao.Create(cHost , cBanco);
  //publicConexao.Conexao.Free;
  publicConexao.User := arq.ReadString('CONFIG' , 'USUARIO' , 'suporte');
  publicConexao.Pass := arq.ReadString('CONFIG' , 'SENHA' , 'sfrtprime2017');
  publicConexao.DataBase := arq.ReadString('CONFIG' , 'BANCO' , 'VW_RTPRIME');
  publicConexao.Host := arq.ReadString('CONFIG' , 'SERVERNAME' , '192.168.2.30');
  publicConexao.Porta := arq.ReadInteger('CONFIG' , 'PORTA' , 1997);
  publicConexao.Conexao := ConnectDBEmpresa(publicConexao.Host , publicConexao.DataBase, 'eAudit - API REST' ,
                                            publicConexao.User, publicConexao.Pass);
  vPorta.Text := IntToStr(publicConexao.Porta);
  vHost.Text := publicConexao.Host;
  vBanco.Text := publicConexao.DataBase;
  vUser.Text := publicConexao.User;
  vPass.Text := publicConexao.Pass;

  btnAtivarClick(self);
  ServerParams := TServerParams.Create;
  ServerParams.HasAuthentication := true;
  ServerParams.UserName          := arq.ReadString('CONFIG' , 'USERAUTH' , '');
  ServerParams.Password          := arq.ReadString('CONFIG' , 'PASSAUTH' , '');

  arq.Free;
end;


procedure TfrmMain.FormDestroy(Sender: TObject);
begin
     publicConexao.Conexao.Free;
     publicConexao.Destroy;
     ServerParams.Free;
end;

procedure TfrmMain.TrayIcon1Click(Sender: TObject);
begin
     TrayIcon1.Visible := False;
     Show();
     WindowState := wsNormal;
     Application.BringToFront();
end;


procedure TfrmMain.apl1Minimize(Sender: TObject);
begin
     Self.Hide();
     Self.WindowState := wsMinimized;
     TrayIcon1.Visible := True;
     TrayIcon1.Animate := True;
     TrayIcon1.ShowBalloonHint;
end;

procedure TfrmMain.BitBtn1Click(Sender: TObject);
var
  arq : TIniFile;
begin
  arq := TIniFile.Create(ExtractFilePath(Application.ExeName) + 'RestConf.ini');
  btnParar.Click;
  publicConexao.Conexao.Free;
  publicConexao.User := vUser.Text;
  publicConexao.Pass := vPass.Text;
  publicConexao.DataBase := vBanco.Text;
  publicConexao.Host := vHost.Text;
  publicConexao.Porta := StrToInt(vPorta.Text);

  try
    publicConexao.Conexao := ConnectDBEmpresa(publicConexao.Host , publicConexao.DataBase, 'eAudit - API REST' ,
                                            publicConexao.User, publicConexao.Pass);
    arq.WriteString('CONFIG', 'SERVERNAME', vHost.Text);
    arq.WriteString('CONFIG' , 'PORTA', vPorta.Text);
    arq.WriteString('CONFIG' , 'USUARIO' , vUser.Text);
    arq.WriteString('CONFIG' , 'SENHA' , vPass.Text);
    arq.WriteString('CONFIG' , 'BANCO' , vBanco.Text);

  finally
    ARQ.Free;
  end;
  btnAtivar.Click;

  if IdHTTPServer1.DefaultPort <> publicConexao.Porta then
    ShowMessage('Reinicie o Programa para Alterar a Porta!');

end;

procedure TfrmMain.btnAtivarClick(Sender: TObject);
begin
     IdHTTPServer1.Active := True;
     txtInfoLabel.Caption := 'Aguardando requisições...';
end;


procedure TfrmMain.btnPararClick(Sender: TObject);
begin
     IdHTTPServer1.Active := False;
     txtInfoLabel.Caption := 'WebService parado.';
end;

procedure TfrmMain.LoglastRequest (ARequestInfo: TIdHTTPRequestInfo);
Begin
     EnterCriticalSection(CriticalSection);
     memoReq.Lines.Add(ARequestInfo.UserAgent + #13 + #10 +
                       ARequestInfo.RawHTTPCommand);
     LeaveCriticalSection(CriticalSection);
End;

procedure TfrmMain.LogLastResponse (AResponseInfo: TIdHTTPResponseInfo);
Begin
     EnterCriticalSection(CriticalSection);
     memoResp.Lines.Add(AResponseInfo.ContentText);
     LeaveCriticalSection(CriticalSection);
End;


procedure TfrmMain.IdHTTPServer1CommandGet(AContext: TIdContext;
  ARequestInfo: TIdHTTPRequestInfo; AResponseInfo: TIdHTTPResponseInfo);
Var
     Cmd           : String;
     Argumentos    : TArguments;
     ServerMethod1 : TServerMethods1;
     JSONStr       : string;
begin
     Cmd := ARequestInfo.RawHTTPCommand;
     If (ServerParams.HasAuthentication) then Begin
        if Not ((ARequestInfo.AuthUsername = ServerParams.Username) and
               (ARequestInfo.AuthPassword = ServerParams.Password))
         Then Begin
           AResponseInfo.AuthRealm := 'RestWebServer';
           AResponseInfo.WriteContent;
           Exit;
         End;
     End;
     if (UpperCase(Copy (Cmd, 1, 3)) = 'GET') OR
        (UpperCase(Copy (Cmd, 1, 4)) = 'POST') OR
        (UpperCase(Copy (Cmd, 1, 3)) = 'HEAD')
     then Begin
        Argumentos    := TServerUtils.ParseRESTURL (ARequestInfo.URI);
        ServerMethod1 := TServerMethods1.Create (nil);
        Try
           LoglastRequest (ARequestInfo);
           If UpperCase(Copy (Cmd, 1, 3)) = 'GET' Then
              JSONStr := ServerMethod1.CallGETServerMethod(Argumentos);
           If UpperCase(Copy (Cmd, 1, 4)) = 'POST' Then
              JSONStr := ServerMethod1.CallPOSTServerMethod(Argumentos);

           //memoResp.Lines.Clear;
           AResponseInfo.ContentText := JSONStr;
           LoglastResponse (AResponseInfo);
           AResponseInfo.WriteContent;
        Finally
           ServerMethod1.Free;
        End;
     end;
end;

procedure TfrmMain.IdHTTPServer1CommandOther(AContext: TIdContext;
  ARequestInfo: TIdHTTPRequestInfo; AResponseInfo: TIdHTTPResponseInfo);
Var
     Cmd           : String;
     Argumentos    : TArguments;
     ServerMethod1 : TServerMethods1;
     JSONStr       : string;
begin
     Cmd := ARequestInfo.RawHTTPCommand;
     If (ServerParams.HasAuthentication) then Begin
        if Not ((ARequestInfo.AuthUsername = ServerParams.Username) and
               (ARequestInfo.AuthPassword = ServerParams.Password))
         Then Begin
           AResponseInfo.AuthRealm := 'RestWebServer';
           AResponseInfo.WriteContent;
           Exit;
         End;
     End;
     if (UpperCase(Copy (Cmd, 1, 3)) = 'PUT') OR
        (UpperCase(Copy (Cmd, 1, 6)) = 'DELETE')
     then Begin
        Argumentos    := TServerUtils.ParseRESTURL (ARequestInfo.URI);
        ServerMethod1 := TServerMethods1.Create (nil);
        Try
           LoglastRequest (ARequestInfo);
           If UpperCase(Copy (Cmd, 1, 3)) = 'PUT' Then
              JSONStr := ServerMethod1.CallPUTServerMethod(Argumentos);
           If UpperCase(Copy (Cmd, 1, 6)) = 'DELETE' Then
              JSONStr := ServerMethod1.CallDELETEServerMethod(Argumentos);

           //memoResp.Lines.Clear;
           AResponseInfo.ContentText := JSONStr;
           LoglastResponse (AResponseInfo);
           AResponseInfo.WriteContent;
        Finally
           ServerMethod1.Free;
        End;
     end;
end;

initialization
  InitializeCriticalSection(CriticalSection);

finalization
  DeleteCriticalSection(CriticalSection);


end.
