program RestWebService;

uses
  Vcl.Forms,
  formMain in 'formMain.pas' {frmMain},
  Vcl.Themes,
  Vcl.Styles,
  ServerUtils in 'ServerUtils.pas',
  SysTypes in 'SysTypes.pas',
  UntServerMethods in 'UntServerMethods.pas',
  UntConexao in 'UntConexao.pas',
  UntFuncoes in 'UntFuncoes.pas';

{$R *.res}

begin
  //ReportMemoryLeaksOnShutdown := DebugHook <> 0;
  Application.Initialize;
  TStyleManager.TrySetStyle('Ruby Graphite');
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
