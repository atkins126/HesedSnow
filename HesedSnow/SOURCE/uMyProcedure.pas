unit uMyProcedure;

interface

uses ComObj, ActiveX, Variants, Windows, Messages, SysUtils, Classes, Vcl.Dialogs, System.UITypes;

    procedure CleanOutTable(tabl:String); // �������� ������� �� �������
    procedure CleanOutTableAndIndex(tabl:String); // ������� ������� � ��������� �������

implementation

uses uDataModul;

procedure CleanOutTable(tabl: String);
begin  // �������� ������� �� �������
    with DM.qQuery do
    begin
      Close;
      SQL.Clear;
      SQL.Add('Delete * From '+ tabl);
     ExecSQL;
     Close;
    end;
end;

procedure CleanOutTableAndIndex(tabl: String);
begin
 with DM.qQuery do
  begin
   Close; //�������� ������ �������
   SQL.Clear;
   SQL.Add('Alter Table '+ tabl +'  Alter Column ��� Counter(1,1)');
   ExecSQL;
   Close;
CleanOutTable(tabl);  // ������� �������
  end;
end;

end.
