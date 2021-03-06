unit UntFuncoes;

interface

uses
  SysUtils, DB, DBClient, ADODB, Variants, VarUtils, Classes, StdCtrls, Controls,dxBarExtDBItems,
  Clipbrd, DBGrids, ComObj, Dialogs, StrUtils, VclTee.TeCanvas, System.DateUtils,
  cxData, cxDataStorage, cxNavigator, cxDBData, cxGridCustomTableView, cxGridTableView,
  cxGridDBTableView, cxGridLevel, cxClasses, cxGridCustomView, cxGrid, Windows,
  VCLTee.Chart, VCLTee.DBChart, VCLTee.TeEngine, ppReport, cxGridChartView, graphics, cxDBLookupComboBox, ExtCtrls,
  Forms, TypInfo, Math, Types, DBCtrls, PSApi, SynVirtualDataSet, Mormot, SynCommons,
  generics.collections, MormotVCL, Vcl.CheckLst, Tlhelp32, Contnrs, MidasLib, ActiveX;

type
  PrjBoolean = (Verdadeiro = 0, Falso = 1);

  TCell = array of array of String;

  TArrayOfString = array of String;

type
  TFuncoes = class
  public
    class function CriaQuery(pConnection: TADOConnection; AOwner: TComponent = nil): TADOQuery;
    class function CriaStoredProc(pConnection: TADOConnection): TADOStoredProc;
    class function RetornaValor(pConnection: TADOConnection; const pSQL: String): Variant;
    class function Func_CriaEClonaCDS(pCdsOrigem: TClientDataSet): TClientDataSet;
    class procedure Tira_ParentBackground_Labels(pParent: TComponent);
    class function RetornaValores(pConnection: TADOConnection; const pSql: String): Variant;
    class function RetornaDouble(pConnection: TADOConnection;
      const pSQL: String; pValorDef: Double = 0): Double;
    class function RetornaInteiro(pConnection: TADOConnection;
      const pSQL: String; pValorDef: Integer = 0): Integer;
    class function RetornaDocVariant(pConnection: TADOConnection; const pSql: String): Variant;
    class function RetornaJSon(pConnection: TADOConnection; const pSql: String): String;

    // Popula o record com o primeiro registro retornado pelo banco
    class function RetornaRecord(var Rec; TypeInfo: pointer; pConnection: TAdoConnection; const pSql: String): Boolean;

    class function RetornaArray(var pArray; TypeInfo: pointer; pConnection: TAdoConnection; const pSql: String): Boolean;

    // Popula o objeto com o primeiro registro retornado pelo banco
    // ATEN��O:: AS PROPRIEDADES DO OBJETO QUE V�O SER RETORNADAS TEM QUE ESTAR NA SE��O PUBLISHED!!!
    class function RetornaObjeto(var ObjectInstance; pConnection: TAdoConnection; const pSql: String): Boolean;

    // Popula um objectlist, cada registro retornado ser� um objeto na lista. TItemClass = Classe dos Itens do ObjectList
    // ATEN��O:: AS PROPRIEDADES DO OBJETO QUE DEVEM SER RETORNADAS TEM QUE ESTAR NA SE��O PUBLISHED!!!
    class function RetornaListaObjetos(var ObjectList: TObjectList; ItemClass: TClass; pConnection: TAdoConnection; const pSql: String): Boolean; overload; static;
    // Igual � anterior, mas recebe uma lista de objetos gen�rica por par�metro
    class function RetornaListaObjetos<T: class>(var ObjectList: TObjectList<T>; ItemClass: TClass; pConnection: TAdoConnection; const pSql: String): Boolean; overload; static;

    class function ExecutaProcedure(pConnection: TADOConnection;
      const pSQL: String; pValorDef: Integer = 0): Integer;

    class function Dias_Uteis(DtaI, DataF: TDateTime):Integer;
//    class procedure Proc_CopiarFields(pOrigem, pDestino: TClientDataSet);
  end;

procedure FillDataSetFromJSon(pDataSet: TDataSet; pJson: String);
function IntToBooleano(Valor: Integer): Boolean;
function RoundUp(Valor: Double; Digitos: Integer): Double;

// Cria uma tabela em mem�ria e configura autom�ticamente o cxColumn para buscar o valor do ListField que Corresponde ao KeyField
procedure CriaLookupField(pDatasetDestino: TDataSet; pConnection: TAdoConnection; pSql, pKeyField: String);
procedure FazLookupCxLookup(AOwner: TComponent; pCxComboBox: TcxDBLookupComboBox; pConnection: TAdoConnection; pSql, pKeyFieldNames, pListFieldNames: String);
procedure FazLookupCxColumn(AOwner: TComponent; pCol: TcxGridDBColumn; pConnection: TAdoConnection; pSql, pKeyFieldNames, pListFieldNames: String);
procedure FazLookupComboBox(pLookupComboBox: TDBLookupComboBox; pConnection: TAdoConnection; pSql: String);
procedure FazdxLookup(AOwner: TComponent; pCxComboBox: TdxBarLookupCombo; pConnection: TAdoConnection; pSql, pKeyFieldNames, pListFieldNames: String);

function CopiaECriaField(pFieldOrigem: TField; pDataSetDest: TDataSet): TField;
function VerficarSeAplicaticoEstarRodando(Nome:String) : Boolean;
function Get_Handle_Of_Application(partialTitle: string): HWND;
procedure CopiaValoresDataSet(Origem, Destino: TDataSet);

function VarToIntDef(pVar: Variant; pValorDef: Integer = 0): Integer;
function VarToFloatDef(pVar: Variant; pValorDef: Double = 0): Double;

function Func_DateTime_SqlServer(pData: TDateTime; withQuotes: Boolean = true): String;

procedure AbortMessage(Mensagem: String; const pTitulo: String; pFlags: Integer = 0);

function Func_Date_SqlServer(pData: TDateTime; pHoraIni: Boolean): String;
function Func_DateBetween_SqlServer(pDataIni, pDataFim: TDateTime): String;
function StrParaDate(fStr: String): TDate;
//procedure Proc_CxGrid_Para_Excel(pCxGrid: TcxGridDBTableView);
function ExcluiQuebra( pStr, pStrSubst: String): String;
procedure Func_DBGrid_Para_Excel(pGrid: TDBGrid);
procedure Proc_cxChart_Para_Bmp(pChart: TcxGridChartView; const pArquivo: String);

// Fun��es de convers�o de data
procedure GetTimeSettings(var pFormatSettings: TFormatSettings);
function TryHoraParaInt(const pHora: String; out Resultado: Integer): Boolean;
//function TryHoraParaDateTime(const pHora: String; out Resultado: TDateTime): Boolean;
function HoraParaDateTimeDef(const pHora: String; pDef: TDateTime = 0): TDateTime;
function DateTimeParaHoras(const pDate: TDateTime): String;
function IntParaHora(FInt: Integer): String;
function HoraParaInt(const AText: String): Integer;
function TempoTotal(fMin: Integer): String;

function GetDateTime(FDate : TDateTime; FHora : String) : Boolean;
function MontaDateTime(FDate : TDateTime; FHora : String) : TDateTime;

function TryStrCurToFloat(const pStr: String; var Valor: Double): Boolean;

function StringToStrList(const pDelimiter: Char; const pDelimitedText: String): TStringList;

function PrjMessage(const pText: String; const pTitulo: String; pFlags: Integer = 0): Integer;

procedure CopiaValorCampos(pSource, pDest: TDataSet; const pFieldsParaIgnorar: String = ''; pApenasExistentes: Boolean = false);

// Fun��es de gr�fico
procedure GfCopiarSerie(pOrigem, pDestino: TChartSeries);

function GetCellsFromExcelClipbrd: TCell;

function StringToVarArray(const aSeparator, aString: String): Variant;

function SplitString(const aSeparator, aString: String; aMax: Integer = 0): TArrayOfString;

procedure SavalReportParaPDF(rReport: TppReport; sNomeArquivo: String);

function CurIn(pValor: Currency; Valores: array of currency): Boolean;

function CurrToSqlStr(pValor: Currency): String;

function FloatToSqlStr(pValor: Extended): String;

function AreaIntToArea(pArea: Integer): Currency;

function AreaIntToStrArea(pArea: Integer): String;

function BlendColors(Color1, Color2: TColor; A: Byte): TColor;

procedure CarregaFormEmPanel(pForm: TForm; pPanel: TPanel);

function RetornaValoresCheckBox(Check:TCheckListBox) : string;

function InstanciaQuery(pConnection: TADOConnection): TADOQuery;

function ConnectDBEmpresa(pHost : string; pDataBase: string; pUser: string;
          pUserBD: string = 'studio_fiscal'; pPass : string = 'sf@2012!@#$'): TADOConnection;

procedure GetFieldsComboFlat(pFlat : TComboFlat; pIndex: integer = 10);

function ProcessoExiste(ExeFileName: string): boolean;

function FormatedFloat(pValor: Extended): String;

procedure ExpTXT(DataSet: TDataSet; Arq: string);

procedure ExpDOC(DataSet: TDataSet; Arq: string);

procedure ExpXLS(DataSet: TDataSet; Arq: string);

const
  cHorasDia = 8;

implementation

{ TFuncoes }

function IntToBooleano(Valor: Integer): Boolean;
begin
  if PrjBoolean(Valor) = Verdadeiro then
    Result:= True
  else
    Result:= False;
end;

function FormatedFloat(pValor: Extended): String;
var
  AFormatSettings: TFormatSettings;
begin
  AFormatSettings.Create;
  AFormatSettings.DecimalSeparator:= '.';
  Result:= FloatToStr(pValor, AFormatSettings);
end;

procedure FillDataSetFromJSon(pDataSet: TDataSet; pJson: String);
var
  I: Integer;
  FSource: TDataSet;
begin
  FSource:= JSONToDataSet(nil, pJson);
  try
    FSource.First;
    while not FSource.Eof do
    begin
      pDataSet.Append;
      CopiaValorCampos(FSource, pDataSet, '', True);
      pDataSet.Post;
      FSource.Next;
    end;

  finally
    FSource.Free;
  end;
end;

function RoundUp(Valor: Double; Digitos: Integer): Double;
var
  I: Integer;
  FMult: Integer;
  FValCalc: Extended;
  FValNormalizado: Extended;
begin
  FMult:= 1;
  for I := 0 to Digitos-1 do
    FMult:= FMult*10;

  FValCalc:= Valor * FMult;

  if CompareValue(FValCalc, Trunc(FValCalc), 0.09) = GreaterThanValue then
    FValCalc:= Trunc(FValCalc)+1;

//  FValNormalizado:= (Ceil((Valor * FMult) / FMult);
  Result:= RoundTo(FValCalc/FMult, Digitos*-1);
end;

procedure CarregaFormEmPanel(pForm: TForm; pPanel: TPanel);
begin
  pForm.FormStyle := fsNormal;
  pForm.Parent := pPanel;
//  pForm.WindowState := wsMaximized;
  pForm.Align:= alClient;
  pForm.BorderStyle := bsNone;
end;

function CopiaECriaField(pFieldOrigem: TField; pDataSetDest: TDataSet): TField;
var
  NewField: TField;
  FieldDef: TFieldDef;
begin
  FieldDef := pDataSetDest.FieldDefs.AddFieldDef;
  FieldDef.DataType := pFieldOrigem.DataType;
  FieldDef.Size := pFieldOrigem.Size;
  FieldDef.Name := pFieldOrigem.FieldName;

  NewField := FieldDef.CreateField(pDataSetDest);
  NewField.Visible := pFieldOrigem.Visible;
  NewField.DisplayLabel := pFieldOrigem.DisplayLabel;
  NewField.DisplayWidth := pFieldOrigem.DisplayWidth;
  NewField.EditMask := pFieldOrigem.EditMask;

 if IsPublishedProp(pFieldOrigem, 'currency')  then
   SetPropValue(NewField, 'currency', GetPropValue(pFieldOrigem, 'currency'));

  Result:= NewField;
end;

// Primeiro campo do Sql deve ser o C�digo (LookupKeyField), o segundo deve ser o nome (LookupResultField);
procedure CriaLookupField(pDatasetDestino: TDataSet; pConnection: TAdoConnection; pSql, pKeyField: String);
var
  FQry: TADOQuery;
  FFieldOrigem, FFieldDest: TField;
begin
  FQry:= TFuncoes.CriaQuery(pConnection, pDatasetDestino);
  try
    FQry.SQL.Text:= pSql;
    FQry.Open;

    FFieldOrigem:= FQry.Fields[0];
    if not Assigned(FFieldOrigem) then
      raise Exception.Create(Format('Erro! Campo n�o encontrado para fazer lookup. Fun��o FazLookupDataSet', []));

    FFieldDest:= CopiaECriaField(FQry.Fields[1], pDatasetDestino);
    FFieldDest.FieldKind:= fkLookup;
    FFieldDest.LookupDataSet:= FQry;
    FFieldDest.LookupKeyFields:= FQry.Fields[0].FieldName;
    FFieldDest.LookupResultField:= FQry.Fields[1].FieldName;
    FFieldDest.KeyFields:= pKeyField;
  except
    FQry.Free;
    raise;
  end;
end;

procedure FazLookupComboBox(pLookupComboBox: TDBLookupComboBox; pConnection: TAdoConnection; pSql: String);
var
  FQry: TADOQuery;
  FDataSource: TDataSource;
  FFieldOrigem, FFieldDest: TField;
begin
  FQry:= TFuncoes.CriaQuery(pConnection, pLookupComboBox);
  try
    FQry.SQL.Text:= pSql;
    FDataSource:= TDataSource.Create(FQry);
    FDataSource.DataSet:= FQry;

    FQry.Open;

    pLookupComboBox.ListSource:= FDataSource;
    pLookupComboBox.KeyField:= FQry.Fields[0].FieldName;
    pLookupComboBox.ListField:= FQry.Fields[1].FieldName;
  except
    FQry.Free;
    raise;
  end;
end;

procedure FazLookupCxColumn(AOwner: TComponent; pCol: TcxGridDBColumn; pConnection: TAdoConnection; pSql, pKeyFieldNames, pListFieldNames: String);
var
  FQry: TADOQuery;
  FDataSource: TDataSource;
  FProps: TcxLookupComboBoxProperties;
begin
  FQry:= TFuncoes.CriaQuery(pConnection, AOwner);
  try
    FQry.SQL.Text:= pSql;
    FDataSource:= TDataSource.Create(FQry);
    FDataSource.DataSet:= FQry;
    FQry.Open;

    pCol.PropertiesClass:= TcxLookupComboBoxProperties;
    FProps:= TcxLookupComboBoxProperties(pCol.Properties);
    FProps.KeyFieldNames:= pKeyFieldNames;
    FProps.ListFieldNames:= pListFieldNames;
    FProps.ListSource:= FDataSource;
  except
    FQry.Free;
    raise;
  end;
end;

function Get_Handle_Of_Application(partialTitle: string): HWND;
var
  hWndTemp: hWnd;
  iLenText: Integer;
  cTitletemp: array [0..254] of Char;
  sTitleTemp: string;
begin
  hWndTemp := FindWindow(nil, nil);
  while hWndTemp <> 0 do begin
    iLenText := GetWindowText(hWndTemp, cTitletemp, 255);
    sTitleTemp := cTitletemp;
    sTitleTemp := UpperCase(copy( sTitleTemp, 1, iLenText));
    partialTitle := UpperCase(partialTitle);
    if pos( partialTitle, sTitleTemp ) <> 0 then
      Break;
    hWndTemp := GetWindow(hWndTemp, GW_HWNDNEXT);
  end;
  result := hWndTemp;
end;


function VerficarSeAplicaticoEstarRodando(Nome:String):Boolean;
var rId:array[0..999] of DWord; i,NumProc,NumMod:DWord;
    HProc,HMod:THandle; sNome:String;
    Tamanho, Count:Integer;
    sNomeTratado:String;
begin
  result:=False;
  SetLength(sNome, 256);
// Aqui vc pega os IDs dos processos em execu��o
  EnumProcesses(@rId[0], 4000, NumProc);

// Aqui vc faz um for p/ pegar cada processo
  for i := 0 to NumProc div 4 do
  begin
// Aqui vc seleciona o processo
    HProc := OpenProcess(Process_Query_Information or Process_VM_Read, FALSE, rId[i]);
    if HProc = 0 then
      Continue;
// Aqui vc pega os m�dulos do processo
// Como vc s� quer o nome do programa, ent�o ser� sempre o primeiro
    EnumProcessModules(HProc, @HMod, 4, NumMod);
// Aqui vc pega o nome do m�dulo; como � o primeiro, � o nome do programa
    GetModuleBaseName(HProc, HMod, @sNome[1], 256);
    sNomeTratado := trim(sNome);
    Tamanho:=Length(SnomeTratado);
     Count:=1;
     While Count <= Tamanho do
       begin
         if SnomeTratado[Count]= '' Then
           Break;
        count:=Count+1;
       end;
     sNomeTratado:=Copy(SnomeTratado,1,Count-1);
    if AnsiUpperCase(sNomeTratado)=AnsiUpperCase(Nome) Then
      Result:=True;
// Aqui vc libera o handle do processo selecionado
    CloseHandle(HProc);
  end;
end;

function ProcessoExiste(ExeFileName: string): boolean;
const
  PROCESS_TERMINATE=$0001;
var
  ContinueLoop: BOOL;
  FSnapshotHandle: THandle;
  FProcessEntry32: TProcessEntry32;
begin
  result := false;

  FSnapshotHandle := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
  FProcessEntry32.dwSize := Sizeof(FProcessEntry32);
  ContinueLoop := Process32First(FSnapshotHandle,FProcessEntry32);

  while integer(ContinueLoop) <> 0 do
  begin
    if ((UpperCase(ExtractFileName(FProcessEntry32.szExeFile)) = UpperCase(ExeFileName))
      or (UpperCase(FProcessEntry32.szExeFile) = UpperCase(ExeFileName))) then
    begin
      Result := true;
      exit;
    end;
    ContinueLoop := Process32Next(FSnapshotHandle,FProcessEntry32);
    end;
    CloseHandle(FSnapshotHandle);
end;


procedure FazLookupCxLookup(AOwner: TComponent; pCxComboBox: TcxDBLookupComboBox; pConnection: TAdoConnection; pSql, pKeyFieldNames, pListFieldNames: String);
var
  FQry: TADOQuery;
  FDataSource: TDataSource;
  FProps: TcxLookupComboBoxProperties;
begin
  FQry:= TFuncoes.CriaQuery(pConnection, AOwner);
  try
    FQry.SQL.Text:= pSql;
    FDataSource:= TDataSource.Create(FQry);
    FDataSource.DataSet:= FQry;
    FQry.Open;

    FProps:= pCxComboBox.Properties;
    FProps.KeyFieldNames:= pKeyFieldNames;
    FProps.ListFieldNames:= pListFieldNames;
    FProps.ListSource:= FDataSource;
  except
    FQry.Free;
    raise;
  end;
end;

procedure FazdxLookup(AOwner: TComponent; pCxComboBox: TdxBarLookupCombo; pConnection: TAdoConnection; pSql, pKeyFieldNames, pListFieldNames: String);
var
  FQry: TADOQuery;
  FDataSource: TDataSource;
  //FProps: TdxBarLookupComboProperties;
begin
  FQry:= TFuncoes.CriaQuery(pConnection, AOwner);
  try
    FQry.SQL.Text:= pSql;
    FDataSource:= TDataSource.Create(FQry);
    FDataSource.DataSet:= FQry;
    FQry.Open;

    pCxComboBox.KeyField:= pKeyFieldNames;
    pCxComboBox.ListField:= pListFieldNames;
    pCxComboBox.ListSource:= FDataSource;
  except
    FQry.Free;
    raise;
  end;
end;

{ Dataset deve Destino deve estar em edi��o! � copiada apenas os valores dos campos de mesmo nome e da linha corrente; }
procedure CopiaValoresDataSet(Origem, Destino: TDataSet);
var
  I: Integer;
  FField: TField;
begin
  for I:= 0 to Origem.Fields.Count-1 do
  begin
    FField:= Destino.FindField(Origem.Fields[I].FieldName);
    if FField <> nil then
    begin
      FField.Value:= Origem.Fields[I].Value;
    end;
  end;
end;

// Usage  NewColor:= Blend(Color1, Color2, blending level 0 to 100);
function BlendColors(Color1, Color2: TColor; A: Byte): TColor;
var
  c1, c2: LongInt;
  r, g, b, v1, v2: byte;
begin
  A:= Round(2.55 * A);
  c1 := ColorToRGB(Color1);
  c2 := ColorToRGB(Color2);
  v1:= Byte(c1);
  v2:= Byte(c2);
  r:= A * (v1 - v2) shr 8 + v2;
  v1:= Byte(c1 shr 8);
  v2:= Byte(c2 shr 8);
  g:= A * (v1 - v2) shr 8 + v2;
  v1:= Byte(c1 shr 16);
  v2:= Byte(c2 shr 16);
  b:= A * (v1 - v2) shr 8 + v2;
  Result := (b shl 16) + (g shl 8) + r;
end;

function CurrToSqlStr(pValor: Currency): String;
begin
  Result:= StringReplace(CurrToStr(pValor), ',', '.', []);
end;

function FloatToSqlStr(pValor: Extended): String;
begin
  Result:= StringReplace(FloatToStr(pValor), ',', '.', []);
end;

function AreaIntToStrArea(pArea: Integer): String;
begin
  Result:= StringReplace(CurrToStr(AreaIntToArea(pArea)), ',', '.', []);
end;

function AreaIntToArea(pArea: Integer): Currency;
begin
  Result:= pArea / 1000;
end;

function CurIn(pValor: Currency; Valores: array of currency): Boolean;
var
  I: integer;
begin
  Result:= False;

  for I := Low(Valores) to High(Valores) do
    if pValor = Valores[I] then
    begin
      Result:= True;
      Exit;
    end;

end;

class function TFuncoes.CriaQuery(pConnection: TADOConnection; AOwner: TComponent = nil): TADOQuery;
begin
  if AOwner = nil then
    Result:= TADOQuery.Create(pConnection)
  else Result:= TADOQuery.Create(AOwner);
  try
    Result.Connection:= pConnection;
  except
    Result.Free;
    raise;
  end;
end;

class function TFuncoes.CriaStoredProc(pConnection: TADOConnection): TADOStoredProc;
begin
  Result:= TADOStoredProc.Create(pConnection);
  try
    Result.Connection:= pConnection;
  except
    Result.Free;
    raise;
  end;
end;

class function TFuncoes.Dias_Uteis(DtaI, DataF: TDateTime): Integer;
var
 contador:Integer;
begin
 if DtaI > DataF then
 begin
   result  := 0;
   exit
 end;
 Contador := 0;
 while (DtaI <= DataF) do
 begin
   if ((DayOfWeek(DtaI) <> 1) and (DayOfWeek(DtaI) <> 7)) then
     Inc(Contador);
   DtaI := DtaI + 1
 end;
 result := Contador;

end;

class function TFuncoes.RetornaValores(pConnection: TADOConnection;
  const pSql: String): Variant;
var
  vQuery: TADOQuery;
  I: Integer;
begin
  vQuery:= CriaQuery(pConnection);
  try
    vQuery.SQL.Text:= pSql;
    vQuery.Active:= True;

    if vQuery.Fields.Count = 0 then
    begin
      Result:= null;
    end
   else
    begin
      Result:= VarArrayCreate([0, vQuery.Fields.Count], varVariant);
      for I:= 0 to vQuery.Fields.Count - 1 do
        Result[I]:= vQuery.Fields[I].Value;
    end;

    vQuery.Close;
  finally
    vQuery.Free;
  end;
end;

class function TFuncoes.RetornaValor(pConnection: TADOConnection; const pSQL: String): Variant;
var
  vQuery: TADOQuery;
begin
  vQuery:= CriaQuery(pConnection);
  try
    vQuery.SQL.Text:= pSql;
    vQuery.Active:= True;
    Result:= vQuery.Fields[0].Value;
    vQuery.Close;
  finally
    vQuery.Free;
  end;
end;

class function TFuncoes.RetornaJSon(pConnection: TADOConnection;
  const pSql: String): String;
var
  vQuery: TADOQuery;
begin
  vQuery:= CriaQuery(pConnection);
  try
    vQuery.SQL.Text:= pSql;
    vQuery.Active:= True;
    Result:= DataSetToJson(vQuery);
    vQuery.Close;
  finally
    vQuery.Free;
  end;
end;

// Classe dos Itens do ObjectList
class function TFuncoes.RetornaListaObjetos(var ObjectList: TObjectList; ItemClass: TClass;
 pConnection: TAdoConnection; const pSql: String): Boolean;
var
  FJson: String;
begin
  FJson:= RetornaJSon(pConnection, pSql);

  Result:= ObjectLoadJSon(ObjectList, FJson, ItemClass);
end;

// Classe dos Itens do ObjectList
class function TFuncoes.RetornaListaObjetos<T>(var ObjectList: TObjectList<T>; ItemClass: TClass;
 pConnection: TAdoConnection; const pSql: String): Boolean;
var
  FResults, fVariant: Variant;
  I: Integer;
  Instance: TObject;
begin
  Result:= False;

  FResults:= TFuncoes.RetornaDocVariant(pConnection, pSql);

  if FResults._Count = 0 then Exit;

  for I := 0 to FResults._Count-1 do
  begin
    fVariant:= _ByRef(FResults._(I), JSON_OPTIONS[true]);
    Instance:= ItemClass.Create;
    Result:= ObjectLoadJSon(Instance, fVariant._JSON, nil);
    ObjectList.Add(Instance);
  end;
end;

class function TFuncoes.RetornaObjeto(var ObjectInstance;
  pConnection: TAdoConnection; const pSql: String): Boolean;
var
  FResults, fVariant: Variant;
begin
  Result:= False;

  FResults:= TFuncoes.RetornaDocVariant(pConnection, pSql);

  if FResults._Count = 0 then Exit;

  fVariant:= _ByRef(FResults._(0), JSON_OPTIONS[true]);

  Result:= ObjectLoadJSon(ObjectInstance, fVariant._JSON, nil);
end;

class function TFuncoes.RetornaRecord(var Rec; TypeInfo: pointer; pConnection: TAdoConnection;
  const pSql: String): Boolean;
var
  FResults, fVariant: Variant;
begin
  Result:= False;

  FResults:= TFuncoes.RetornaDocVariant(pConnection, pSql);

  if FResults._Count = 0 then Exit;

  fVariant:= _ByRef(FResults._(0), JSON_OPTIONS[true]);

  Result:= RecordLoadJSON(Rec, fVariant._JSON, TypeInfo);

  Result:= True;
end;

class function TFuncoes.RetornaArray(var pArray; TypeInfo: pointer;
      pConnection: TAdoConnection; const pSql: String): Boolean;
var
  FResults, fVariant: Variant;
  fJson: String;
  I: Integer;
  aDynArray: TDynArray;
begin
  Result:= False;

  aDynArray.Init(TypeInfo,pArray);

  fJSon:= TFuncoes.RetornaJSon(pConnection,pSql);

  aDynArray.LoadFromJSON(PUTF8Char(WinAnsiToUtf8(fJSon)));

  Result:= True;
end;

class function TFuncoes.RetornaDocVariant(pConnection: TADOConnection;
  const pSql: String): Variant;
begin
    Result:= _JSon(RetornaJSon(pConnection, pSql));
end;

class function TFuncoes.RetornaDouble(pConnection: TADOConnection; const pSQL: String; pValorDef: Double = 0): Double;
begin
  Result:= VarToFloatDef(RetornaValor(pConnection, pSQL), pValorDef);
end;

class function TFuncoes.RetornaInteiro(pConnection: TADOConnection;
  const pSQL: String; pValorDef: Integer): Integer;
begin
  Result:= VarToIntDef(RetornaValor(pConnection, pSQL), pValorDef);
end;

class function TFuncoes.ExecutaProcedure(pConnection: TADOConnection;
  const pSQL: String; pValorDef: Integer = 0): Integer;
var
  vQuery: TADOQuery;
begin
  Result:= pValorDef;
  vQuery:= CriaQuery(pConnection);
  try
    vQuery.SQL.Text:= pSql;
    Result:= vQuery.ExecSQL;
//    Result:= vQuery.Fields[0].Value;
  finally
    vQuery.Free;
  end;
end;


{class procedure TFuncoes.Proc_CopiarFields(pOrigem, pDestino: TClientDataSet);
var
  Field, NewField: TField;
  FieldDef: TFieldDef;
begin
  for Field in pOrigem.Fields do
  begin
    FieldDef := FieldDefs.AddFieldDef;
    FieldDef.DataType := Field.DataType;
    FieldDef.Size := Field.Size;
    FieldDef.Name := Field.FieldName;

    NewField := FieldDef.CreateField(Self);
    NewField.Visible := Field.Visible;
    NewField.DisplayLabel := Field.DisplayLabel;
    NewField.DisplayWidth := Field.DisplayWidth;
    NewField.EditMask := Field.EditMask;
  end;
end;                 }

class function TFuncoes.Func_CriaEClonaCDS(pCdsOrigem: TClientDataSet): TClientDataSet;
begin
  Result:= TClientDataSet.Create(pCdsOrigem);
  try
    Result.FieldDefs.Assign(pCdsOrigem.FieldDefs);
    Result.CreateDataSet;
    Result.Open;
    Result.CloneCursor(pCdsOrigem, True);
  except
    Result.Free;
    raise;
  end;
end;

class procedure TFuncoes.Tira_ParentBackground_Labels(pParent: TComponent);
var
  Cpt: Integer;
begin
  for Cpt := 0 to pParent.ComponentCount - 1 do
    if (pParent.Components[Cpt] is TLabel) then
      with (pParent.Components[Cpt] as TLabel) do
      begin
        ControlStyle := ControlStyle - [csParentBackground];
      end;
end;

function VarToIntDef(pVar: Variant; pValorDef: Integer = 0): Integer;
begin
  try
    if VarIsNull(pVar) then
      Result:= pValorDef
    else
      Result:= pVar;
  except
    Result:= pValorDef;
  end;
end;

function VarToFloatDef(pVar: Variant; pValorDef: Double = 0): Double;
begin
  try
    if VarIsNull(pVar) then
      Result:= pValorDef
    else
      Result:= pVar;
  except
    Result:= pValorDef;
  end;
end;

// pHoraIni = True = 00:00; = False: 23:59
function Func_Date_SqlServer(pData: TDateTime; pHoraIni: Boolean): String;
var
  vData: String;
begin
  vData:= FormatDateTime('yyyymmdd', pData);
  if pHoraIni then
    Result:= QuotedStr(vData+' 00:00')
  else
    Result:= QuotedStr(vData+' 23:59');

end;

function Func_DateTime_SqlServer(pData: TDateTime; withQuotes: Boolean = true): String;
begin
  if withQuotes then
    Result:= QuotedStr(FormatDateTime('yyyymmdd hh:mm', pData))
  else
    Result:= FormatDateTime('yyyymmdd hh:mm', pData);
end;

function Func_DateBetween_SqlServer(pDataIni, pDataFim: TDateTime): String;
begin
  Result:= ' BETWEEN '+Func_Date_SqlServer(pDataIni, True)+' AND '+Func_Date_SqlServer(pDataFim, False);
end;

function StrParaDate(fStr: String): TDate;
var
  vDia, vMes, vAno: Word;
begin
   if Length(fStr) < 10 then raise Exception.Create('Parametro incorreto: Fun��o StrParaDate');

   vAno:= StrToInt(copy(fStr, 0, 4));
   vMes:= StrToInt(copy(fStr, 6, 2));
   vDia:= StrToInt(copy(fStr, 9, 2));
   Result:= EncodeDate(vAno, vMes, vDia);
end;

procedure Proc_cxChart_Para_Bmp(pChart: TcxGridChartView; const pArquivo: String);
var
 ImageExport : TGraphic;
begin
   ImageExport := TGraphicClass.Create;
   try
     ImageExport := pChart.CreateImage(TBitmap);
     ImageExport.SaveToFile(pArquivo);
   finally
     ImageExport.free;
   end;
end;

procedure AbortMessage(Mensagem: String; const pTitulo: String; pFlags: Integer = 0);
begin
  PrjMessage(Mensagem, pTitulo, pFlags);
  Abort;
end;

procedure Func_DBGrid_Para_Excel(pGrid: TDBGrid);
var
    Linhas: TStringList;
    i, posicao, Count, Plan: integer;
    s, Cab, PlanName: string;
    planilha : variant;
begin
  Linhas := TStringList.Create;
  Clipboard.Open;
  try
    with pGrid do
    begin
      planilha:= CreateoleObject('Excel.Application');
      //planilha.caption :=QryNomeLayoutLayNome.Value; ->Alterado 26/07/2012
      planilha.WorkBooks.add(1);

      DataSource.DataSet.DisableControls;
      Posicao := DataSource.DataSet.RecNo;
      DataSource.DataSet.First;

      //Cabe�alho Par�metros
      Cab := 'Data Relat�rio';
      Cab := Cab + #9; // Tabula��o
      Cab := Cab + DateTimeToStr(Now());
      Linhas.Add(Cab);

      //Cabe�alho Colunas
      Cab := '';
      for i := 0 to Columns.Count - 1 do
      begin
        if Assigned(Columns.Items[i].Field) then
        begin
          if i > 0 then Cab := Cab + #9; // Tabula��o
          Cab := Cab + Columns.Items[i].Title.Caption;
        end;
      end;
      Linhas.Add(Cab);

      //Itens
      Count:=1;
      Plan:=1;
      while not DataSource.DataSet.Eof do
      begin
        s := '';
        for i := 0 to Columns.Count - 1 do
        begin
          if Assigned(Columns.Items[i].Field) then
          begin
            if i > 0 then s := s + #9; // Tabula��o
            s := s + Columns.Items[i].Field.Text;
          end;  
        end;
        Linhas.Add(s);
        Count:=Count+1;
        DataSource.DataSet.Next;
      end;
      DataSource.DataSet.RecNo := Posicao;
      DataSource.DataSet.EnableControls;
    end;

    Clipboard.SetTextBuf(Pointer(Linhas.Text));
  finally
    Linhas.Free;
    Clipboard.Close;
  end;
  planilha.WorkBooks[1].Sheets.Add;
  PlanName:=InttoStr(Plan)+'Audit_Project';
  planilha.WorkBooks[1].WorkSheets[1].Name :=PlanName ;
  planilha.Workbooks[1].WorkSheets[PlanName].Paste;
  Clipboard.Clear;

  showmessage('Arquivo Exportado com Sucesso.');
  planilha.visible := true;

end;

function ExcluiQuebra( pStr, pStrSubst: String): String;
begin
  Result:= StringReplace( pStr, #$D#$A, pStrSubst, [rfReplaceAll]);
  Result:= StringReplace( Result, #13#10, pStrSubst, [rfReplaceAll]);
end;

procedure GetTimeSettings(var pFormatSettings: TFormatSettings);
begin
  pFormatSettings.ShortTimeFormat:= 'hh:mm';
  pformatSettings.LongTimeFormat:= 'hh:mm';

  GetLocaleFormatSettings(LOCALE_SYSTEM_DEFAULT, pformatSettings);
end;

function HoraParaDateTimeDef(const pHora: String; pDef: TDateTime = 0): TDateTime;
var
  formatSettings : TFormatSettings;
begin
  GetTimeSettings(FormatSettings);

  Result:= StrToTimeDef(pHora, pDef, formatSettings);
end;

function DateTimeParaHoras(const pDate: TDateTime): String;
var
  formatSettings : TFormatSettings;
begin
  GetTimeSettings(FormatSettings);

  Result:= TimeToStr(pDate, formatSettings);
end;

procedure FazErro;
begin
  ShowMessage('Valor inv�lido!');
  Abort;
end;

function TryHoraParaInt(const pHora: String; out Resultado: Integer): Boolean;
var
  FPosDelimitador: Integer;
  FDias, FHoras, FMin: Integer;
begin
  Result:= False;
  FPosDelimitador:= Pos(':', pHora);

  if not TryStrToInt(Copy(pHora, 0, FPosDelimitador-1), FHoras) then
    Exit;

  if not TryStrToInt(Copy(pHora, FPosDelimitador+1, 2), FMin) then
    Exit;

  Resultado:= (FHoras*60) + FMin;
  Result:= True;
end;
              {
function TryHoraParaDateTime(const pHora: String; out Resultado: TDateTime): Boolean;
var
  formatSettings : TFormatSettings;
  FPosDelimitador: Integer;
  FDias, FHoras, FMin: Integer;
begin
  FPosDelimitador:= Pos(':', AText);

  FHoras:= Copy(pHora, 0, FPosDelimitador);

  if not TryStrToInt(Copy(pHora, 0, FPosDelimitador), FHoras) then
    FazErro;

  if not TryStrToInt(Copy(pHora, FPosDelimitador+1, 2, FMin) then
    FazErro;

  GetTimeSettings(FormatSettings);

  Result:= TryStrToTime(pHora, Resultado, formatSettings);
end;       }

function HoraParaInt(const AText: String): Integer;
var
  FAnos, FMes, FDias, FHoras, FMin, FSec, FMSec: Word;
  FMinutos: Integer;
  FTime: TDateTime;
  FPosDelimitador: Integer;
begin
  if Trim(AText) = '' then Exit;

  FPosDelimitador:= Pos(':', AText);
  if FPosDelimitador > 0 then begin
    if not TryHoraParaInt(AText, Result) then
      FazErro;
  end
 else
  begin
    FHoras:= 0;
    if not TryStrToInt(AText, FMinutos) then
      FazErro;

    FMin:= FMinutos;
    Result:= FMin+(FHoras*60);    
  end;
end;

function IntParaHora(FInt: Integer): String;
var
  FHoras, FMinutos: Integer;
  FTime: TDateTime;
  sMinutos: String;
begin
  FHoras:= FInt div 60;
  FMinutos:= FInt mod 60;

  sMinutos:= IntToStr(FMinutos);
  if Length(sMinutos) = 1 then
    sMinutos:= '0'+sMinutos;

  Result:= IntToStr(FHoras)+':'+sMinutos;
//  FTime:= EncodeDateTime( FHoras, FMinutos, 0, 0);
//  Result:= Copy(DateTimeParaHoras(FTime), 0, 5);

//  Result:= FormatFloat('#00', FHoras)+':'+FormatFloat('00', FMinutos);
end;


function TempoTotal(fMin: Integer): String;
var
  FMinutos, FHoras, FDias: Integer;
  sMinutos, sHoras: string;
begin
  FMinutos:= fMin mod 60;
  FHoras:= fMin div 60;
  FDias:= FHoras div cHorasDia;
  FHoras:= FHoras mod cHorasDia;

  sMinutos:= IntToStr(FMinutos);
  if Length(sMinutos) = 1 then
    sMinutos:= '0'+sMinutos;

  sHoras:= IntToStr(FHoras);
  if Length(sHoras) = 1 then
    sHoras:= '0'+sHoras;

  Result:= IntToStr(FDias) +' dias '+ IntToStr(FHoras)+' horas '+sMinutos +' min';
end;

function MontaDateTime(FDate : TDateTime; FHora : String) : TDateTime;
var
FResult,FIni : string;
begin
FIni := FormatDateTime('dd/mm/yyyy', FDate);
FResult := FIni +Chr(32)+ FHora;
Result := StrToDateTime(FResult);
end;

function GetDateTime(FDate : TDateTime; FHora : String) : Boolean;
{Essa funcao Retorna um booleano no comparativo de uma data hora informada no sitema
com a data hora no caso a Data esta em um DatePicker a hora em um MaskEdit}
var

FResult ,FFim, FIni : string;
begin
FFim := FormatDateTime('yyyy/mm/dd hh:mm:ss', Now);
FIni := FormatDateTime('yyyy/mm/dd', FDate);
FResult := FIni +Chr(32)+ FHora;

if FFim < FResult then
  begin
  Result := False;
  end
else
  if FFim >= FResult then
  begin
  Result := True;
  end
end;

// Fun��es de gr�fico
procedure GfCopiarSerie(pOrigem, pDestino: TChartSeries);
begin
  pDestino.Assign(pOrigem);
end;

function TryStrCurToFloat(const pStr: String; var Valor: Double): Boolean;
var
  s: string;
  c: Currency;
begin
  // Remove os separadores de milhares
  s := StringReplace(pStr, '.', '', [rfReplaceAll]);
  s := StringReplace(s, 'R$', '', [rfReplaceAll]);
  s := StringReplace(s, 'RS', '', [rfReplaceAll]);
  s := Trim(s);
  Result:= TryStrToFloat(s, Valor);
end;

function StringToStrList(const pDelimiter: Char; const pDelimitedText: String): TStringList;
begin
  Result:= TStringList.Create;
  try
    Result.Delimiter:= pDelimiter;
    Result.DelimitedText:= pDelimitedText;
  except
    Result.Free;
    raise;
  end;
end;

function StringToVarArray(const aSeparator, aString: String): Variant;
var
  FListaValores: TArrayOfString;
  I: Integer;
begin
  Result:= null;

  if aString = '' then Exit;

  FListaValores:= SplitString(',', aString);

  Result:= VarArrayCreate([0, High(FListaValores)], varVariant);

  for I:= 0 to High(FListaValores) do
    Result[I] := FListaValores[I];

end;

function SplitString(const aSeparator, aString: String; aMax: Integer = 0): TArrayOfString;
var
  i, strt, cnt: Integer;
  sepLen: Integer;

  procedure AddString(aEnd: Integer = -1);
  var
    endPos: Integer;
  begin
    if (aEnd = -1) then
      endPos := i
    else
      endPos := aEnd + 1;

    if (strt < endPos) then
      result[cnt] := Copy(aString, strt, endPos - strt)
    else
      result[cnt] := '';

    Inc(cnt);
  end;       
begin
  if (aString = '') or (aMax < 0) then
  begin
    SetLength(result, 0);
    EXIT;
  end;

  if (aSeparator = '') then
  begin
    SetLength(result, 1);
    result[0] := aString;
    EXIT;
  end;

  sepLen := Length(aSeparator);
  SetLength(result, (Length(aString) div sepLen) + 1);

  i     := 1;
  strt  := i;
  cnt   := 0;
  while (i <= (Length(aString)- sepLen + 1)) do
  begin
    if (aString[i] = aSeparator[1]) then
      if (Copy(aString, i, sepLen) = aSeparator) then
      begin
        AddString;

        if (cnt = aMax) then
        begin
          SetLength(result, cnt);
          EXIT;
        end;

        Inc(i, sepLen - 1);
        strt := i + 1;
      end;

    Inc(i);
  end;

  AddString(Length(aString));

  SetLength(result, cnt);
end;

function GetCellsFromExcelClipbrd: TCell;
var
  RowList: TStringList;
  ColList: TArrayOfString;
  I,F: Integer;
  FMaxColCnt: Integer;
begin
  FMaxColCnt:= 0;

  RowList:= TStringList.Create;
  try
    RowList.Text:= clipboard.AsText;
    for I:= 0 to RowList.Count - 1 do
    begin
      ColList:= SplitString(chr(9), RowList[I]);

      if FMaxColCnt < Length(ColList) then
      begin
        FMaxColCnt:= Length(ColList);
        SetLength(Result, RowList.Count, FMaxColCnt);
      end;

      for F:= 0 to High(ColList) do
      begin
        Result[I][F]:= ColList[F];
//          Memo1.Lines.Add('Row: '+IntToStr(I)+', Col: '+IntToStr(F)+' = '+ColList[F]);
      end;
      
    end;
  finally
    RowList.Free;
  end;
end;

function PrjMessage(const pText: String; const pTitulo: String; pFlags: Integer = 0): Integer;
begin
  {$IFDEF VER150}
    Result:= Application.MessageBox(PAnsiChar(pText),PAnsiChar(pTitulo),pFlags);
  {$ELSE}
    Result:= Application.MessageBox(PWideChar(pText),PWideChar(pTitulo),pFlags);
  {$ENDIF}
end;

procedure CopiaValorCampos(pSource, pDest: TDataSet; const pFieldsParaIgnorar: String = ''; pApenasExistentes: Boolean = false);
var
  vField: TField;
  I: Integer;
  FIgnoreList: TStringList;
begin
  FIgnoreList:= StringToStrList(';', UpperCase(pFieldsParaIgnorar));
  try
    for I:= 0 to pSource.FieldCount - 1 do
    begin
      if FIgnoreList.IndexOf(upperCase(pSource.Fields[I].FieldName)) >= 0 then
       Continue;

      vField:= pDest.FindField(pSource.Fields[I].FieldName);

      if not Assigned(vField) then
      begin
        if not pApenasExistentes then
          PrjMessage('Erro ao copiar Tabela!. Campo n�o existe: '+pSource.Fields[I].FieldName, 'e-Project', MB_ICONINFORMATION);
      end
      else if not vField.ReadOnly then
      begin
        pDest.FindField(pSource.Fields[I].FieldName).Value:= pSource.Fields[I].Value;
      end;
    end;
  finally
    FIgnoreList.Free;
  end;
end;

procedure SavalReportParaPDF(rReport: TppReport; sNomeArquivo: String);
var 
  wNomarq : string;
  oOldDevice: String;
  oAllowPrintToFile, oShowPrintDialog, oSaveAsTemplate, oSavePrinterSetup: Boolean;
  oTextFileName: String;

  procedure SaveOldConfig;
  begin
    with rReport do
    begin
      oOldDevice:= DeviceType;
      oAllowPrintToFile:= AllowPrintToFile;
      oShowPrintDialog:= ShowPrintDialog;
      oSaveAsTemplate:= SaveAsTemplate;
      oSavePrinterSetup:= SavePrinterSetup;
      oTextFileName:= TextFileName;
    end;
  end;

  procedure CarregaOldConfig;
  begin
    with rReport do
    begin
      DeviceType := oOldDevice;
      AllowPrintToFile := oAllowPrintToFile;
      ShowPrintDialog := oShowPrintDialog;
      SaveAsTemplate := oSaveAsTemplate;
      SavePrinterSetup := oSavePrinterSetup;
      TextFileName := oTextFileName;
    end;
  end;
begin
  try
    try
      SaveOldConfig;

      {:Gera o arquivo PDF para ser anexado para vers�o 10 acima}
      if not DirectoryExists(ExtractFilePath(sNomeArquivo)) then   // verifica se j� existe o diret�rio.
        ForceDirectories(ExtractFilePath(sNomeArquivo));

      with rReport do
      begin
//        wNomarq  := sNomeArquivo;
        DeviceType := 'PDF';
        AllowPrintToFile := True;
        ShowPrintDialog := False;
        SaveAsTemplate := True;
        SavePrinterSetup := True;
        TextFileName := sNomeArquivo; //Joga o pdf criado para a pasta aonde esta o programa l� no cliente.
        Print;
//        Pt_MessageDlg('O arquivo foi salvo com sucesso na pasta '+ExtractFilePath(Application.ExeName), mtInformation, [mbOK],0);
      end;
    except
      on E: Exception do
      begin
        if Pos('Cannot Open File', E.Message) > 0 then
          raise Exception.Create('Arquivo est� aberto. Feche o Adobe para gravar!')
        else
          raise E;
      end;
    end;
    
  finally
    CarregaOldConfig;
  end;
end;

function RetornaValoresCheckBox(Check:TCheckListBox) : string;
var
  Lista : String;
  i     : integer;
begin
  for I := 0 to Check.Items.Count -1 do // � fun��o for (Inicia zero at� a qtd de itens do CheckListBox
  // � Verificando qual est� selecionado
  begin
    if Check.Checked[i] = true then // � Se o item estiver selecionado ent�o �
    begin
    if lista = '' then // � Verificar se a lista esta em branco
    Lista := #13 + '' + copy(Check.Items[i],0,length(Check.Items[i]) - 950) + ''  // � se sim ele recebe o item selecionado da posicao 0 a 2
    else
    lista := lista + #13 + copy(Check.Items[i],0,length(Check.Items[i]) -950 ); // � se n�o ele recebe ele mesmo adicionando o
    // � pr�ximo item selecionado.
    end;

  end;

  Result := Lista;

end;

{ O c�digo abaixo usa Delphi puro, montando um TXT no bra�o, basicamente varremos o dataSet e vamos montando o arquivo texto separando campos por ponto e v�rgula }

procedure ExpTXT(DataSet: TDataSet; Arq: string);
var
  i: integer;
  sl: TStringList;
  st: string;
begin
  DataSet.First;
  sl := TStringList.Create;
  try
    st := '';
    for i := 0 to DataSet.Fields.Count - 1 do
      st := st + DataSet.Fields[i].DisplayLabel + ';';
    sl.Add(st);
    DataSet.First;
    while not DataSet.Eof do
    begin
      st := '';
      for i := 0 to DataSet.Fields.Count - 1 do
        st := st + DataSet.Fields[i].DisplayText + ';';
      sl.Add(st);
      DataSet.Next;
    end;
    sl.SaveToFile(Arq);
  finally
     sl.free;
  end;
end;

{ O c�digo abaixo usa a tecnologia OLE para criar uma planilha do Excel e enviar os dados do DataSet . OLE � uma tecnologia que pode ser usada desde o Delphi 2 e permite manipular aplica��es automaticamente, o que chamamos de Automation, usando interface COM }

procedure ExpXLS(DataSet: TDataSet; Arq: string);
var
  ExcApp: OleVariant;
  i,l: integer;
begin
  ExcApp := CreateOleObject('Excel.Application');
  ExcApp.Visible := True;
  ExcApp.WorkBooks.Add;
  DataSet.First;
  l := 1;
  DataSet.First;
  while not DataSet.EOF do
  begin
    for i := 0 to DataSet.Fields.Count - 1 do
      ExcApp.WorkBooks[1].Sheets[1].Cells[l,i + 1] :=
        DataSet.Fields[i].DisplayText;
    DataSet.Next;
    l := l + 1;
  end;
  ExcApp.WorkBooks[1].SaveAs(Arq);
end;

{ O c�digo abaixo usa a tecnologia OLE para criar uma tabela no WORD e enviar os dados do DataSet  }

procedure ExpDOC(DataSet: TDataSet; Arq: string);
var
  WordApp,WordDoc,WordTable,WordRange: Variant;
  Row,Column: integer;
begin
  WordApp := CreateOleobject('Word.basic');
  WordApp.Appshow;
  WordDoc := CreateOleobject('Word.Document');
  WordRange := WordDoc.Range;
  WordTable := WordDoc.tables.Add(
    WordDoc.Range,1,DataSet.FieldCount);
  for Column:=0 to DataSet.FieldCount-1 do
    WordTable.cell(1,Column+1).range.text:=
      DataSet.Fields.Fields[Column].FieldName;
  Row := 2;
  DataSet.First;
  while not DataSet.Eof do
  begin
     WordTable.Rows.Add;
     for Column:=0 to DataSet.FieldCount-1 do
       WordTable.cell(Row,Column+1).range.text :=
         DataSet.Fields.Fields[Column].DisplayText;
     DataSet.next;
     Row := Row+1;
  end;
  WordDoc.SaveAs(Arq);
  WordDoc := unAssigned;
end;


function InstanciaQuery(pConnection: TADOConnection): TADOQuery;
var
  pQuery : TADOQuery;
begin
  pQuery := TADOQuery.Create(nil);
  pQuery.CommandTimeout := 999999;
  pQuery.Connection := pConnection;
  pQuery.Prepared := True;
  pQuery.ParamCheck := False;
  pQuery.Active := False;

  Result := pQuery;

end;

function ConnectDBEmpresa(pHost, pDataBase, pUser, pUserBd , pPass :string): TADOConnection;
begin
  Result                   :=TADOConnection.Create(nil);
  Result.ConnectionString  :='Provider=SQLOLEDB.1;Password='+pPass+';User ID='+pUserBD+';Initial Catalog='+pDataBase+';Data Source='+pHost+'; Application Name=e-Audit - ' + pUser;
  Result.LoginPrompt       :=False;
  Result.Connected         :=true;
  Result.CommandTimeout    := 9999999;
end;

procedure GetFieldsComboFlat(pFlat : TComboFlat; pIndex: integer);
var
  x: integer;
begin
  pFlat.Items.Clear;

  x := YearOf(IncYear(Now, -15));

  while ( x <= YearOf(Now)) do
  begin
    pFlat.Items.Add(IntToStr(x));
    x := x + 1;
  end;

  pFlat.ItemIndex := pIndex;

end;


end.
