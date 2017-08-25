unit ServerUtils;

interface

Uses
    SysTypes, System.SysUtils, RegularExpressions, IdURI, IdGlobal;

type
    TServerUtils = class
      class function ParseRESTURL (const Cmd : string): TArguments;
      class function Result2JSON (wsResult : TResultErro) : String;
    end;

    TServerParams = class
    private
      fUsername : string;
      fPassword : String;
      fHasAuthenticacion : Boolean;
      function GetUserName : String;
      function GetPassword : String;
      function GetHasAuthentication : Boolean;
    Public
      property HasAuthentication : Boolean read fHasAuthenticacion write fHasAuthenticacion;
      property UserName : string read GetUserName write fUsername;
      property Password : string read GetPassword write fPassword;

      constructor Create; overload;
    end;

implementation

class function TServerUtils.ParseRESTURL (const Cmd : string): TArguments;
Var
     NewCmd       : String;
     iHttp        : Integer;
     ArraySize    : Integer;
     iBar1, IBar2 : Integer;
     Cont         : Integer;
begin
     NewCmd    := Cmd;
     ArraySize := TRegEx.Matches(NewCmd, '/').Count;
     SetLength(Result, ArraySize);
     NewCmd    := NewCmd + '/';

     iBar1 := Pos ('/', NewCmd);
     Delete (NewCmd, 1, iBar1);

     for Cont := 0 to ArraySize - 1 do
     begin
         iBar2 := Pos ('/', NewCmd);

         Result [Cont] := TIdURI.URLDecode (Copy (NewCmd, 1, iBar2 - 1), IndyTextEncoding (encUTF8));

         Delete (NewCmd, 1, iBar2);
     end;
end;

class function TServerUtils.Result2JSON (wsResult : TResultErro) : String;
Var
     SB: TStringBuilder;
begin
     SB := TStringBuilder.Create();
     SB.Append('{');
     SB.Append('"STATUS":"' + IntToStr(wsResult.STATUS) + '"');
     SB.Append(',"MENSAGEM":"' + wsResult.MENSAGEM + '"');
     SB.Append('}');

     Result := SB.ToString;
     SB.Free;
end;


constructor TServerParams.Create;
begin
    HasAuthentication := False;
end;


function TServerParams.GetUserName : String;
begin
     Result := fUsername;
end;

function TServerParams.GetPassword : String;
begin
     Result := fPassword;
end;

function TServerParams.GetHasAuthentication : Boolean;
Begin
     Result := fHasAuthenticacion;
End;

end.
