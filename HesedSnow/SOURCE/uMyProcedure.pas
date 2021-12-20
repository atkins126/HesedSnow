unit uMyProcedure;

interface

uses
  ComObj, ActiveX, Variants, Windows, Messages, SysUtils, Classes, Vcl.Dialogs,
  System.UITypes, JvStringGrid;
// �������� ������� �� �������
procedure CleanOutTable(tabl: String);
// ������� ������� � ��������� �������
procedure CleanOutTableAndIndex(tabl: String; pole: String);
// �������� ��� ���� �������� ���������
Function isStringAssign(s: String; combo: TStrings): boolean;
// ��������� '.' ��� ',' � ������������ ���������� ����������� ����� � ������� �����
procedure DecimalChange(s: string);
// ����� ����������� ����� � ������� ����� � �������
function isDesimal(): String;

function ParseStringUpakovka(s: string): String;

procedure AutoStringGridWidth(StringGrid: TJvStringGrid);
// �������� ����� �������

function RazdelenieIO(IO: String; ind: Integer): String;
function DCOUNT(str, Delimeter: string): Integer; // ������� ���� � ������

function DelSymbolIntoString(s, Delimeter:string): string; //������� ������ � ������

implementation

uses uDataModul, Registry;

function DelSymbolIntoString(s, Delimeter:string): string; //������� ������ � ������
var
i,j: integer;
begin
 for i := 0 to s.Length-1 do
 begin
  j:=pos(Delimeter,s);
  Delete(s,j,1);
 end;
Result := s;
end;

function DCOUNT(str, Delimeter: string): Integer;
// ������� ���� � �����
var
  StrL: TStringList;
  ParseStr: string;
begin
  try
    StrL := TStringList.Create;
    ParseStr := StringReplace(str, Delimeter, #13, [rfReplaceAll]);
    StrL.Text := ParseStr;
    Result := StrL.Count;
  finally
    StrL.Free;
  end;
end;

function RazdelenieIO(IO: String; ind: Integer): String;
// ���������� ����� �������� �� ��� �����
begin
  if (ind = 0) or (ind > 2) then
    Result := ''
  else if ind = 1 then
  begin // ���
    Result := Copy(IO, 1, Pos(' ', IO) - 1);
  end
  else
  begin // ��������
    Result := Copy(IO, Pos(' ', IO) + 1);
  end;
end;

procedure CleanOutTable(tabl: String);
begin // �������� ������� �� �������
  with DM.qQuery do
  begin
    Close;
    SQL.Clear;
    SQL.Add('Delete * From ' + tabl);
    ExecSQL;
    Close;
  end;
end;

procedure CleanOutTableAndIndex(tabl: String; pole: String);
begin
  with DM.qQuery do
  begin
    Close; // �������� ������ ������� pole - �������� ���� �� �������� �����
    SQL.Clear;
    SQL.Add('Alter Table ' + tabl + '  Alter Column ' + pole + ' Counter(1,1)');
    ExecSQL;
    Close;
    CleanOutTable(tabl); // ������� �������
  end;
end;

Function isStringAssign(s: string; combo: TStrings): boolean;
var
  I: Integer;
begin
  Result := False;
  for I := 0 to combo.Count - 1 do
  begin
    if combo.Strings[I] = s then
    begin
      Result := True;
      break;
    end;
  end;
end;

function ParseStringUpakovka(s: string): String;
var
  p: Integer;
begin
  s := Trim(Copy(s, 1, Pos('��', s) - 1));
  if not(s = '') then
  begin
    p := LastDelimiter(' ', s);
    Result := Copy(s, p, Length(s) - p + 1);
  end
  else
    Result := '0';
end;

procedure DecimalChange(s: string);
var
  Reg: TRegistry;
begin
  Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    Reg.OpenKey('Control Panel\International', False);
    Reg.WriteString('sDecimal', s { ����� ����������� } );
    // Reg.WriteString('sDecimal', '.' { ����� ����������� } );
  finally
    Reg.Free;
  end;
end;

function isDesimal(): String;
var
  Reg: TRegistry;
begin
  Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    Reg.OpenKey('Control Panel\International', False);
    // Reg.WriteString('sDecimal', '.' { ����� ����������� } );
    Result := Reg.ReadString('sDecimal');
  finally
    Reg.Free;
  end;
end;

{
  ----- �������� ������ ������� -----
  AutoStringGridWidth(��� JvStringGrid);
}
procedure AutoStringGridWidth(StringGrid: TJvStringGrid);
var
  X, Y, w: Integer;
  MaxWidth: Integer;
begin
  with StringGrid do
  //  ClientHeight := DefaultRowHeight * RowCount + 5;
  with StringGrid do
  begin
    for X := 0 to ColCount - 1 do
    begin
      MaxWidth := 0;
      for Y := 0 to RowCount - 1 do
      begin
        w := Canvas.TextWidth(Cells[X, Y]);
        if w > MaxWidth then
          MaxWidth := w;
      end;
      ColWidths[X] := MaxWidth + 5;
    end;
  end;
end;

end.
