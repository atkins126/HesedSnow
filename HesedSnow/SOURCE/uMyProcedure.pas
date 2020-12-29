unit uMyProcedure;

interface

uses
  ComObj, ActiveX, Variants, Windows, Messages, SysUtils, Classes, Vcl.Dialogs,
  System.UITypes;
// Очистить таблицу от записей
procedure CleanOutTable(tabl: String);
// Очистка таблицы и обнуление индекса
procedure CleanOutTableAndIndex(tabl: String);
// Проверка или есть название Ведомости
Function isStringAssign(s: String; combo: TStrings): boolean;

function ParseStringUpakovka(s: string): String;

implementation

uses uDataModul;

procedure CleanOutTable(tabl: String);
begin // Очистить таблицу от записей
  with DM.qQuery do
  begin
    Close;
    SQL.Clear;
    SQL.Add('Delete * From ' + tabl);
    ExecSQL;
    Close;
  end;
end;

procedure CleanOutTableAndIndex(tabl: String);
begin
  with DM.qQuery do
  begin
    Close; // обнуляем индекс таблицы
    SQL.Clear;
    SQL.Add('Alter Table ' + tabl + '  Alter Column Код Counter(1,1)');
    ExecSQL;
    Close;
    CleanOutTable(tabl); // очищаем таблицу
  end;
end;

Function isStringAssign(s: string; combo: TStrings): Boolean;
var
  I: Integer;
begin
  result := False;
  for I := 0 to combo.Count - 1 do
  begin
    if combo.Strings[i] = s then
    begin
      result := True;
      break;
    end;
   end;
end;

function ParseStringUpakovka(s: string): String;
var
  p: Integer;
begin
  s := Trim(Copy(s, 1, Pos('шт', s) - 1));
  if not(s = '') then
  begin
    p := LastDelimiter(' ', s);
    Result := Copy(s, p, Length(s) - p + 1);
  end
  else
    Result := '0';
end;



end.
