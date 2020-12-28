unit uBDlogic;

interface

uses
  Classes, System.Variants, System.SysUtils;

Function CuratorsList(table, pole: String): TStringList;
function ExportExcel(vKurators, vDir, vTitle: String): boolean;

implementation

uses uDataModul, uMyExcel;

{ ------- ������ ��������� ��� ��������� -------
 table -  ������� �� ������� ���������� ������
 pole - ���� �� �������� ������������ ������
}
Function CuratorsList(table, pole: String): TStringList;
var // ������ ��������� � ��������� ��� ���������
  i: Integer;
  f: TStringList;
begin
  with DM.qQuery do
  begin
    f := TStringList.Create;

    Active := false;
    SQL.Clear;
    SQL.Add('SELECT ' + table + '.' + pole + ' FROM ' + table + ' GROUP BY ' +
      table + '.' + pole + ';');
    Active := true;

    if RecordCount > 0 then
    begin
      for i := 0 to RecordCount - 1 do
      begin
        f.Add(FieldByName(pole).AsString);
        next;
      end;
      Result := f;
    end;
  end;
end;

{* ---------- ������� � ������ ��������� --------
  vKurators - �������
  vDir - ���������� ��� ���������� ���������
  vTitle -  �������� ���������
*}
function ExportExcel(vKurators, vDir, vTitle: String): boolean;
var   // ������������ ��������� � ������
  i: Integer;
  Sheets, ExcelApp: Variant;
  flName: String;
begin
  Result := false;
  if uMyExcel.RunExcel(false, false) = true then
  begin
    MyExcel.Workbooks.Add; // ��������� ����� �����
    Sheets := MyExcel.ActiveWorkBook.Sheets[1];
    Sheets.name := vKurators;
    ExcelApp := MyExcel.ActiveWorkBook.WorkSheets[1].columns;

    // ----------- ��������� ��������� ---------------
    MyExcel.ActiveSheet.PageSetup.Orientation := 2;
    MyExcel.ActiveSheet.PageSetup.LeftMargin :=
      MyExcel.Application.InchesToPoints(0.30);
    MyExcel.ActiveSheet.PageSetup.RightMargin :=
      MyExcel.Application.InchesToPoints(0.20);
    MyExcel.ActiveSheet.PageSetup.TopMargin :=
      MyExcel.Application.InchesToPoints(0.44);
    MyExcel.ActiveSheet.PageSetup.BottomMargin :=
      MyExcel.Application.InchesToPoints(0.44);
    Sheets.PageSetup.PrintTitleRows := '$2:$2';

    MyExcel.ActiveSheet.PageSetup.CenterFooter :=
      '&"Arial"&8���� &"Arial,����������"&P' +
      '&"Arial,�������" �� &"Arial,����������"&N';
    MyExcel.ActiveSheet.PageSetup.RightFooter := '&D';

    // ----------------- ������� -------------------
    // �
    ExcelApp.columns[1].columnwidth := 3.17;
    ExcelApp.columns[1].AutoFit;
    ExcelApp.columns[1].HorizontalAlignment := 3;
    ExcelApp.columns[1].VerticalAlignment := 2;
    // JDCID
    ExcelApp.columns[2].columnwidth := 11.71;
    ExcelApp.columns[2].HorizontalAlignment := 3;
    ExcelApp.columns[2].NumberFormat := '@';
    ExcelApp.columns[2].VerticalAlignment := 2;
    // FIO
    ExcelApp.columns[3].columnwidth := 20.00;
    ExcelApp.columns[3].WrapText := true;
    ExcelApp.columns[3].VerticalAlignment := 2;
    // �����
    ExcelApp.columns[4].columnwidth := 20.00;
    ExcelApp.columns[4].WrapText := true;
    ExcelApp.columns[4].VerticalAlignment := 2;
    // ���
    ExcelApp.columns[5].columnwidth := 12.00;
    ExcelApp.columns[5].NumberFormat := '@';
    ExcelApp.columns[5].VerticalAlignment := 2;
    // ���-��
    ExcelApp.columns[6].columnwidth := 3.29;
    ExcelApp.columns[6].VerticalAlignment := 2;
    ExcelApp.columns[6].HorizontalAlignment := 3;
    // �����
    ExcelApp.columns[7].columnwidth := 13.14;
    ExcelApp.columns[7].NumberFormat := '0.00" ���."';
    ExcelApp.columns[7].VerticalAlignment := 2;
    // ����
    ExcelApp.columns[8].columnwidth := 11.71;
    ExcelApp.columns[8].HorizontalAlignment := 3;
    ExcelApp.columns[8].NumberFormat := 'dd.mm.yyyy'; // '��.��.����';
    ExcelApp.columns[8].VerticalAlignment := 2;
    // �������
    ExcelApp.columns[9].columnwidth := 12.00;
    ExcelApp.columns[9].VerticalAlignment := 2;
    // ����������
    ExcelApp.columns[10].columnwidth := 27.00;
    ExcelApp.columns[10].VerticalAlignment := 2;

    // -------------- �������� ������� ----------------
    Sheets.select;
    Sheets.Range['A1:H1'].select;
    Sheets.Range['A1:H1'].Merge;
    MyExcel.Selection.HorizontalAlignment := xlCenter;
    MyExcel.Selection.Font.name := 'Calibri';
    MyExcel.Selection.Font.Size := 14;
    MyExcel.Selection.Font.Bold := true;
    Sheets.Cells[1, 1] := vTitle;

    Sheets.Range['I1:J1'].select;
    Sheets.Range['I1:J1'].Merge;
    MyExcel.Selection.Font.name := 'Calibri';
    MyExcel.Selection.Font.Size := 8;
    MyExcel.Selection.HorizontalAlignment := 4;
    MyExcel.Range['I1'].Value := '�������: ' + vKurators;

    // -------------- ����� ������� ---------------
    Sheets.Range['A2:J2'].select;
    MyExcel.Selection.Borders.LineStyle := xlContinuous; // �������
    MyExcel.Selection.Borders.Weight := xlThin; // ��������
    MyExcel.Selection.HorizontalAlignment := 3;
    MyExcel.Selection.VerticalAlignment := 2;
    MyExcel.Selection.Font.name := 'Calibri';
    MyExcel.Selection.Font.Size := 12;
    MyExcel.Selection.WrapText := true; // ��� ������� �� ������
    MyExcel.Selection.Borders[9].LineStyle := 9;

    Sheets.Range['A2'].Value := '�';
    Sheets.Range['B2'].Value := 'JDCID';
    Sheets.Range['C2'].Value := 'ϲ�';
    Sheets.Range['D2'].Value := '������';
    Sheets.Range['E2'].Value := '�������';
    Sheets.Range['F2'].Value := 'ʳ�';
    Sheets.Range['G2'].Value := '����';
    Sheets.Range['H2'].Value := '����';
    Sheets.Range['I2'].Value := 'ϳ����';
    Sheets.Range['J2'].Value := '�������';

    with DM.qQuery do
    begin
      Active := false;
      SQL.Clear;
      SQL.Add('SELECT * FROM Vidomist WHERE (((Vidomist.Curator)="' + vKurators
        + '"));');
      Active := true;
      First;
      i := 3; // c ������� ������ ���������
      while not Eof do
      begin
        Sheets.Range['' + 'A' + IntToStr(i) + ':J' + IntToStr(i) + ''].select;
        MyExcel.Selection.Borders.LineStyle := xlContinuous; // �������
        MyExcel.Selection.Borders.Weight := xlThin; // ��������
        MyExcel.Selection.Font.Size := 11;

        Sheets.Cells[i, 1] := '=ROW(A' + IntToStr(i - 2) + ')';
        Sheets.Cells[i, 2] := FieldByName('JDCID').AsString;
        Sheets.Cells[i, 3] := FieldByName('FIO').AsString;
        Sheets.Cells[i, 4] := FieldByName('Adress').AsString;
        Sheets.Cells[i, 5] := FieldByName('Mobila').AsString;
        Sheets.Cells[i, 6] := FieldByName('Count_usl').AsString;
        Sheets.Cells[i, 7] := CurrToStr(FieldByName('Cost_usl').AsCurrency);

        inc(i);
        next;
      end;

    end;

    // �����
    // ������� ������� � ����� ��� �������
    Sheets.Cells[i, 7] := ('=SUM(G3:G' + IntToStr(i - 1) + ')');
    Sheets.Cells[i, 5] := '�����: ';

    Sheets.Cells[i + 3, 5] := '�������� ��� "���������" ';
    Sheets.Range['E' + IntToStr(i + 3) + ':I' + IntToStr(i + 3) + ''].select;
    MyExcel.Selection.Borders[9].LineStyle := xlContinuous;
    MyExcel.Selection.Borders[9].Weight := xlThin;
    MyExcel.Selection.Font.Size := 12;
    Sheets.Cells[i + 3, 9] := '�. �������� ';

    Sheets.Cells[i + 5, 5] := '�������� ';
    Sheets.Range['E' + IntToStr(i + 5) + ':I' + IntToStr(i + 5) + ''].select;
    MyExcel.Selection.Borders[9].LineStyle := xlContinuous;
    MyExcel.Selection.Borders[9].Weight := xlThin;
    MyExcel.Selection.Font.Size := 14;

    flName := vDir + vKurators + '_' + vTitle + '.xlsx';
    SaveWorkBook(flName, 1); // ��������� ����
    Result := true;
  end;
  StopExcel;
  ExcelApp := Unassigned;
  Sheets := Unassigned;
end;

end.
