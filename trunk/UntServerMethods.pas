unit UntServerMethods;

interface

uses System.SysUtils, System.Classes, UntConexao, Data.Win.ADODB, ActiveX,
      System.JSON, Dialogs, ServerUtils, SysTypes;

type
{$METHODINFO ON}
  TServerMethods1 = class(TComponent)
  private
    { Private declarations }
    Function ReturnErro : String;

    function Consulta (empresa: string; nomeView : String; filtro : string; Banco : string = '') : String;

    function Insere (nomeView : String; Campos : string; Valores : string) : String;

    function Atualiza (nomeView : String; Campos : string; Filtro : string) : String;

    function Exclui (nomeView : String; Filtro : string) : String;
    function RemoveSpaces(pText: string): string;

  public
    { Public declarations }
    Constructor Create (aOwner : TComponent); Override;
    Destructor Destroy; Overload;
    function CallGETServerMethod (Argumentos : TArguments) : String;
    function CallPUTServerMethod (Argumentos : TArguments) : string;
    function CallDELETEServerMethod (Argumentos : TArguments) : string;
    function CallPOSTServerMethod (Argumentos : TArguments) : string;

  end;
{$METHODINFO OFF}

implementation

uses System.StrUtils, UntFuncoes;

Constructor TServerMethods1.Create (aOwner : TComponent);
Begin
     inherited Create (aOwner);

End;

Destructor TServerMethods1.Destroy;
begin
     //publicConexao.Conexao.Free;
     //publicConexao.Destroy;
     inherited;
End;

function TServerMethods1.RemoveSpaces(pText : string): string;
begin
  Result:= StringReplace(pText , '#@' , ' ' , [rfReplaceAll]);
end;

function TServerMethods1.Exclui(nomeView, Filtro: String): String;
Var
   Query : TADOQuery;
   JSONObject : TJSONObject;
   filtroWhere, sql : string;
Begin
  try
    filtroWhere := RemoveSpaces(filtro);
    sql := 'Delete From ' + nomeView + ' ' + filtroWhere;
    JSONObject := TJSONObject.Create;

    try
      Query := InstanciaQuery(publicConexao.Conexao);
      Query.SQL.Text := sql;
      Query.ExecSQL;
      JSONObject.AddPair('Resultado - ','Registro Deletados Com Sucesso!');
    except
      on e: exception do
      begin
        JSONObject.AddPair('Erro na instrução - ',E.Message);
      end;
    end;


    Result := JSONObject.ToString;

  Finally
    JSONObject.Free;
    Query.Free;
  end;

end;

function TServerMethods1.Insere(nomeView, Campos, Valores: string): String;
Var
   Query : TADOQuery;
   JSONObject : TJSONObject;
   filtroWhere, sql : string;
Begin
  try
    sql := 'Insert into ' + nomeView + '( ' + Campos + ' ) Values (' + Valores + ' )';
    JSONObject := TJSONObject.Create;

    try
      Query := InstanciaQuery(publicConexao.Conexao);
      Query.SQL.Text := sql;
      Query.ExecSQL;
      JSONObject.AddPair('Resultado - ','Registro Inserido Com Sucesso!');
    except
      on e: exception do
      begin
        JSONObject.AddPair('Erro na instrução - ',E.Message);
      end;
    end;


    Result := JSONObject.ToString;

  Finally
    JSONObject.Free;
    Query.Free;
  end;

end;

Function TServerMethods1.ReturnErro : String;
Var
     WSResult : TResultErro;
begin
     WSResult.STATUS   := -1;
     WSResult.MENSAGEM := 'Total de argumentos incorretos';
     Result := TServerUtils.Result2JSON(WSResult);
end;

function TServerMethods1.CallGETServerMethod (Argumentos : TArguments) : string;
begin
  if UpperCase(Argumentos[0]) = UpperCase('Consultar') then
    begin
      if Length(Argumentos) > 4 then
        Result := Consulta (Argumentos[1],Argumentos[2],Argumentos[3], Argumentos[4])
      else
        Result := Consulta (Argumentos[1],Argumentos[2],Argumentos[3], '');
    end
  else
  if UpperCase(Argumentos[0]) = UpperCase('Inserir') then
    Result := Insere(Argumentos[1], Argumentos[2],Argumentos[3])
  else
  if UpperCase(Argumentos[0]) = UpperCase('Atualizar') then
    Result := Atualiza (Argumentos[1], Argumentos[2],Argumentos[3])
  else
  if UpperCase(Argumentos[0]) = UpperCase('Excluir') then
    Result := Exclui (Argumentos[1], Argumentos[2]);
end;

function TServerMethods1.CallPOSTServerMethod (Argumentos : TArguments) : string;
begin
  if UpperCase(Argumentos[0]) = UpperCase('Atualizar') then
    Result := Atualiza (Argumentos[1], Argumentos[2],Argumentos[3]);
end;

function TServerMethods1.CallPUTServerMethod (Argumentos : TArguments) : string;
begin
  if UpperCase(Argumentos[0]) = UpperCase('Inserir') then
    Result := Insere(Argumentos[1], Argumentos[2],Argumentos[3]);
end;

function TServerMethods1.Atualiza(nomeView, Campos, Filtro: string): String;
Var
   Query : TADOQuery;
   JSONObject : TJSONObject;
   filtroWhere, sql : string;
Begin
  try
    filtroWhere := RemoveSpaces(filtro);
    sql := 'Update ' + nomeView + ' Set ' + Campos + ' ' + filtroWhere;
    JSONObject := TJSONObject.Create;

    try
      Query := InstanciaQuery(publicConexao.Conexao);
      Query.SQL.Text := sql;
      Query.ExecSQL;
      JSONObject.AddPair('Resultado - ','Registro Alterados Com Sucesso!');
    except
      on e: exception do
      begin
        JSONObject.AddPair('Erro na instrução - ',E.Message);
      end;
    end;


    Result := JSONObject.ToString;

  Finally
    JSONObject.Free;
    Query.Free;
  end;

end;

function TServerMethods1.CallDELETEServerMethod (Argumentos : TArguments) : string;
begin
  if UpperCase(Argumentos[0]) = UpperCase('Excluir') then
    Result := Exclui (Argumentos[1], Argumentos[2]);
end;

function TServerMethods1.Consulta (empresa: string; nomeView : String; filtro : string; Banco: string) : String;
Var
   Query : TADOQuery;
   List : TStringList;
   JSONObject : TJSONObject;
   ID, I : Integer;
   filtroWhere, sql : string;
   Aconnection : TADOConnection;
Begin
  try
    if Banco = '' then
      Aconnection := publicConexao.Conexao
    else
      Aconnection := ConnectDBEmpresa(publicConexao.Host, Banco, 'eAudit - API REST',
                                      publicConexao.User, publicConexao.Pass);
    filtroWhere := RemoveSpaces(filtro);
    sql := 'Select * from ' + nomeView + ' ' + filtroWhere;
    JSONObject := TJSONObject.Create;
    Result := TFuncoes.RetornaJSon(Aconnection , sql);

  Finally
    if Banco <> '' then
      Aconnection.Free;
  end;
end;

end.



