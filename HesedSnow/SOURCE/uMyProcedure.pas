unit uMyProcedure;

interface

uses
  ComObj, ActiveX, Variants, Windows, Messages, SysUtils, Classes, Vcl.Dialogs,
  System.UITypes;

procedure CleanOutTable(tabl: String); // Очистить таблицу от записей
procedure CleanOutTableAndIndex(tabl: String);
// Очистка таблицы и обнуление индекса
Function isStringAssign(s: String; combo: TStrings): boolean;
// Проверка или есть название Ведомости

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

end.
