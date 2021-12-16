unit uBDlogic;

interface

uses
  Classes, System.Variants, System.SysUtils, System.Generics.Collections,
  Vcl.Dialogs;

Function CuratorsList(table, pole: String): TStringList;
Function GorodsList(kurator: String): TStringList;
function ExportExcel(vKurators, vDir, vTitle: String): boolean;
function SearchPoziciy(Tabl: String; pole: String; NaNk: string): boolean;
// ����� ������� � ������� ���������� ����
function InsertRecord(Tabl, poles, _fio, _inn: String; _isObrok: boolean;
  _ostatok: integer): boolean;
procedure ImportExcelToBD;
// ����� �������� ���� � �������
function SearchPoziciyString(Tabl, pole, Values, RezPole: string): String;
// ����������� � �� � ������
function ConnectBD(provider, user, psw, server, db, port: string): boolean;
// ������ �� ������ � ������� ������ � ������
function ImportExcelToUniTable(fName, table: string): boolean;

implementation

uses uDataModul, uMyExcel, uMyProcedure, uMainForm;

function ImportExcelToUniTable(fName, table: string): boolean;
var
  col,z,m,n: integer;
  CollectionNameColumIntoExcel : TDictionary<string, integer>;
  CollectionNameColumnTable: TDictionary<integer, string>;
  field_s, value_s, sql_text : WideString;
  stroka, kolonka: string;
begin
  Result := false;
  try
       if uMyExcel.RunExcel(false, false) = true then
      // �������� �� ������� � ������ Excel
    if uMyExcel.OpenWorkBook(fName, true) then
    // ��������� ����� Excel
    begin
      myForm.ProgressBar.Visible := true;
      MyExcel.ActiveWorkBook.Sheets[1];

      // ��������� ����������� �������
      col := MyExcel.ActiveCell.SpecialCells($000000B).Column;

      // ------ ���������� ��������� ������� �������� �������� �� ������ -------

      CollectionNameColumIntoExcel := TDictionary<string, integer>.Create();
      for z := 1 to col do
      begin
        if not CollectionNameColumIntoExcel.ContainsKey(MyExcel.Cells[1, z].value) then
          CollectionNameColumIntoExcel.Add(MyExcel.Cells[1, z].value, z)
        else
          CollectionNameColumIntoExcel.Add(MyExcel.Cells[1, z].value + z.ToString, z);
      end;



      // --- �������� ������ �������� ������ ������� �� �� -------------
      DM.UniTable1.Active := false;
      DM.UniTable1.TableName := table;
      DM.UniTable1.Active := true;

      field_s := '';
      value_s := '';
      CollectionNameColumnTable := TDictionary<integer, string>.Create();
      for z := 0 to DM.UniTable1.FieldCount -1 do
      begin
       CollectionNameColumnTable.Add(z+1, DM.UniTable1.Fields[z].FieldName);

       if (z = DM.UniTable1.FieldCount -1) then field_s := field_s + '`' + DM.UniTable1.Fields[z].FieldName + '`'
       else field_s := field_s + '`' + DM.UniTable1.Fields[z].FieldName + '`, ';
      end;











      m := 2; // �������� ���������� �� 2-� ������, �������� ��������� �������

      // ��������� ����������� ������
      n := MyExcel.ActiveCell.SpecialCells($000000B).Row;
      n := n + 1;
      with DM, myForm do
      begin
//        CleanOutTable('Vidomist'); // �������� �������
        ProgressBar.Min := 0;
        ProgressBar.Max := n;
        ProgressBar.Position := 1;
         while m <> n do // ���� ������� �� ������� EXCEL
        begin

        for z := 1 to col do
          begin
           kolonka := CollectionNameColumnTable.Items[z];
          stroka := MyExcel.Cells[m, StrToInt(CollectionNameColumIntoExcel.Items[kolonka].ToString)].value;//0002477265

           if z = n then
           value_s := value_s + '''' + stroka + ''''
           else
           value_s := value_s + '''' + stroka + ''', ';

          end;

        DM.UniQuery1.Active := False;
        DM.UniQuery1.SQL.Clear;

        sql_text := 'INSERT INTO `hesed_test`.`' + table + '` (' + field_s + ') VALUES (' + value_s + ');';

        DM.UniQuery1.SQL.Add(sql_text);
        DM.UniQuery1.ExecSQL;



           Inc(m);
          // Application.ProcessMessages;
          Sleep(25);
          ProgressBar.Position := m;
        end;




      end;
           CollectionNameColumIntoExcel.Free;
           CollectionNameColumnTable.Free;
           MyExcel.Application.DisplayAlerts := false;
           StopExcel;

    end;

  except
    on E: Exception do
    begin
     CollectionNameColumIntoExcel.Free;
     CollectionNameColumnTable.Free;
     StopExcel;
    end;
  end;

end;

{
  ---- ��������� ������ ���� � ���� ������ � ������� Vidomist ----
}
procedure ImportExcelToBD;
var
  m, n, col, z: integer;
  CollectionNameTable: TDictionary<string, integer>;
begin
  try
    if uMyExcel.RunExcel(false, false) = true then
      // �������� �� ������� � ������ Excel
      myForm.OpenDialog.Filter := '����� MS Excel|*.xls;*.xlsx|';
    if not myForm.OpenDialog.Execute then
      Exit;

    if uMyExcel.OpenWorkBook(myForm.OpenDialog.FileName, false) then
    // ��������� ����� Excel
    begin
      // MyExcel.Workbooks[1].Worksheets.Count; //���-�� ������ � ���������
      myForm.ProgressBar.Visible := true;
      MyExcel.ActiveWorkBook.Sheets[1];
      // ListExcel := MyExcel.ActiveWorkBook.Sheets[1];

      col := MyExcel.ActiveCell.SpecialCells($000000B).Column;
      // ��������� ����������� �������
      // ------------ ���������� ��������� ������� �������� �������� -------------

      CollectionNameTable := TDictionary<string, integer>.Create();
      for z := 1 to col do
      begin
        if not CollectionNameTable.ContainsKey(MyExcel.Cells[1, z].value) then
          CollectionNameTable.Add(MyExcel.Cells[1, z].value, z)
        else
          CollectionNameTable.Add(MyExcel.Cells[1, z].value + z.ToString, z);
      end;
      // ------------------------ ����� ������� ���  �������� --------------------

      m := 2; // �������� ���������� �� 2-� ������, �������� ��������� �������
      n := MyExcel.ActiveCell.SpecialCells($000000B).Row;
      // ��������� ����������� ������
      n := n + 1;
      with DM, myForm do
      begin
        CleanOutTable('Vidomist'); // �������� �������

        ProgressBar.Min := 0;
        ProgressBar.Max := n;
        ProgressBar.Position := 1;

        while m <> n do // ���� ������� �� ������� EXCEL
        begin

          tVedomost.Insert;

          tVedomost.FieldByName('Num_Uslugy').AsString :=
            MyExcel.Cells
            [m, StrToInt(CollectionNameTable.Items['�����'].ToString)].value;

          tVedomost.FieldByName('JDCID').AsString :=
            MyExcel.Cells
            [m, StrToInt(CollectionNameTable.Items['JDC ID'].ToString)].value;

          tVedomost.FieldByName('FIO').AsString :=
            MyExcel.Cells
            [m, StrToInt(CollectionNameTable.Items['������'].ToString)].value;

          tVedomost.FieldByName('Count_usl').AsCurrency :=
            MyExcel.Cells[m, StrToInt(CollectionNameTable.Items['����������']
            .ToString)].value;

          tVedomost.FieldByName('Cost_usl').AsCurrency :=
            MyExcel.Cells
            [m, StrToInt(CollectionNameTable.Items['��������� ������']
            .ToString)].value;

          tVedomost.FieldByName('Programma').AsString :=
            MyExcel.Cells[m, StrToInt(CollectionNameTable.Items['���������']
            .ToString)].value;

          tVedomost.FieldByName('Curator').AsString :=
            MyExcel.Cells
            [m, StrToInt(CollectionNameTable.Items['�������'].ToString)].value;

          // tVedomost.FieldByName('Zhertva').AsString :=
          // MyExcel.Cells
          // [m, StrToInt(CollectionNameTable.Items['��'].ToString)].value;

          if not CollectionNameTable.ContainsKey('��������� �������') then
            tVedomost.FieldByName('Mobila').AsString := ''
          else
            tVedomost.FieldByName('Mobila').AsString :=
              MyExcel.Cells
              [m, StrToInt(CollectionNameTable.Items['��������� �������']
              .ToString)].value;

          tVedomost.FieldByName('Adress').AsString :=
            MyExcel.Cells[m, StrToInt(CollectionNameTable.Items['������� �����']
            .ToString)].value;

          // tVedomost.FieldByName('S_Kem').AsString :=
          // MyExcel.Cells[m, StrToInt(CollectionNameTable.Items['� ��� ���������']
          // .ToString)].value;

          // tVedomost.FieldByName('SABA').AsCurrency :=
          // MyExcel.Cells
          // [m, StrToInt(CollectionNameTable.Items['��. ����� ��� �� (�����)']
          // .ToString)].value;

          // tVedomost.FieldByName('Type_Uchasnika').AsString :=
          // MyExcel.Cells[m, StrToInt(CollectionNameTable.Items['��� ���������']
          // .ToString)].value;

          tVedomost.FieldByName('Data_Plan').AsDateTime :=
            MyExcel.Cells
            [m, StrToInt(CollectionNameTable.Items['����������� ����']
            .ToString)].value;

          tVedomost.FieldByName('Gorod').AsString :=
            MyExcel.Cells
            [m, StrToInt(CollectionNameTable.Items['�����'].ToString)].value;

          tVedomost.Post;
          Inc(m);
          // Application.ProcessMessages;
          Sleep(25);
          ProgressBar.Position := m;
        end;

      end;

      ShowMessage('������������ ����� ��ϲ���!!!');
    end;

    MyExcel.Application.DisplayAlerts := false;
    StopExcel;
    CollectionNameTable.Clear;
    CollectionNameTable.Free;
    DM.tVedomost.Active := false;
    myForm.ProgressBar.Visible := false;
  except
    on E: EListError do
    begin
      CollectionNameTable.Free;
      CleanOutTable('Vidomist'); // �������� �������
      DM.tVedomost.Active := false;
      ShowMessage('�� ����� ������ �����, ���� ���������� ����');
      StopExcel;
    end;

  end;

end;

{ ------- ������ ��������� ��� ��������� -------
  table -  ������� �� ������� ���������� ������
  pole - ���� �� �������� ������������ ������
}
Function CuratorsList(table, pole: String): TStringList;
var // ������ ��������� � ��������� ��� ���������
  i: integer;
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

{ ------- ������ ������� ��� ��������� -------
  table -  ������� �� ������� ���������� ������
  pole - ���� �� �������� ������������ ������
}
Function GorodsList(kurator: String): TStringList;
var // ������ ��������� � ��������� ��� ���������
  i: integer;
  f: TStringList;
begin
  with DM.qQuery do
  begin
    f := TStringList.Create;

    Active := false;
    SQL.Clear;

    SQL.Add('SELECT Vidomist.Curator, Vidomist.Gorod FROM Vidomist GROUP BY Vidomist.Curator, Vidomist.Gorod '
      + 'HAVING (((Vidomist.Curator)="' + kurator + '"));');

    Active := true;

    if RecordCount > 0 then
    begin
      for i := 0 to RecordCount - 1 do
      begin
        f.Add(FieldByName('Gorod').AsString);
        next;
      end;
      Result := f;
    end;
  end;
end;

{ * ---------- ������� � ������ ��������� --------
  vKurators - �������
  vDir - ���������� ��� ���������� ���������
  vTitle -  �������� ���������
  * }
function ExportExcel(vKurators, vDir, vTitle: String): boolean;
var // ������������ ��������� � ������
  i, j, g: integer;
  Sheets, ExcelApp: Variant;
  flName: String;
  gorods: TStringList;
begin
  Result := false;
  if uMyExcel.RunExcel(false, false) = true then
  begin
    MyExcel.Workbooks.Add; // ��������� ����� �����
    i := 1; // ���-�� ������ � �����

    // 1 ������ ������ ��������� � ������ ���������
    gorods := TStringList.Create;
    gorods := GorodsList(vKurators);

    for g := 0 to gorods.Count - 1 do
    begin

      if MyExcel.Worksheets.Count >= i then
      begin
        Sheets := MyExcel.Worksheets[i];
      end
      else
      begin
        Sheets := MyExcel.Worksheets.Add(After := Sheets);
      end;

      // Sheets.name := Copy(vKurators, 1, 30) + ' ' + gorods[g];

      Sheets := MyExcel.ActiveWorkBook.Sheets[i];
      Sheets.name := gorods[g];

      ExcelApp := MyExcel.ActiveWorkBook.Worksheets[i].columns;

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
      MyExcel.Range['I1'].value := '�������: ' + vKurators;

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

      Sheets.Range['A2'].value := '�';
      Sheets.Range['B2'].value := 'JDCID';
      Sheets.Range['C2'].value := 'ϲ�';
      Sheets.Range['D2'].value := '������';
      Sheets.Range['E2'].value := '�������';
      Sheets.Range['F2'].value := 'ʳ�';
      Sheets.Range['G2'].value := '����';
      Sheets.Range['H2'].value := '����';
      Sheets.Range['I2'].value := 'ϳ����';
      Sheets.Range['J2'].value := '�������';

      with DM.qQuery do
      begin
        Active := false;
        SQL.Clear;
        SQL.Add('SELECT * FROM Vidomist WHERE (((Vidomist.Curator)="' +
          vKurators + '") AND ((Vidomist.Gorod)="' + gorods[g] + '"));');

        Active := true;
        First;
        j := 3; // c ������� ������ ���������
        while not Eof do
        begin
          Sheets.Range['' + 'A' + IntToStr(j) + ':J' + IntToStr(j) + ''].select;
          MyExcel.Selection.Borders.LineStyle := xlContinuous; // �������
          MyExcel.Selection.Borders.Weight := xlThin; // ��������
          MyExcel.Selection.Font.Size := 11;

          Sheets.Cells[j, 1] := '=ROW(A' + IntToStr(j - 2) + ')';
          Sheets.Cells[j, 2] := FieldByName('JDCID').AsString;
          Sheets.Cells[j, 3] := FieldByName('FIO').AsString;
          Sheets.Cells[j, 4] := FieldByName('Adress').AsString;
          Sheets.Cells[j, 5] := FieldByName('Mobila').AsString;
          Sheets.Cells[j, 6] := FieldByName('Count_usl').AsString;
          Sheets.Cells[j, 7] := CurrToStr(FieldByName('Cost_usl').AsCurrency);

          Inc(j);
          next;
        end;
        // �����
        // ������� ������� � ����� ��� �������
        Sheets.Cells[j, 7] := ('=SUM(G3:G' + IntToStr(j - 1) + ')');
        Sheets.Cells[j, 5] := '�����: ';

        Sheets.Cells[j + 3, 5] := '�������� ��� "���������" ';
        Sheets.Range['E' + IntToStr(j + 3) + ':I' + IntToStr(j + 3) +
          ''].select;
        MyExcel.Selection.Borders[9].LineStyle := xlContinuous;
        MyExcel.Selection.Borders[9].Weight := xlThin;
        MyExcel.Selection.Font.Size := 12;
        Sheets.Cells[j + 3, 9] := '�. �������� ';

        Sheets.Cells[j + 5, 5] := '�������� ';
        Sheets.Range['E' + IntToStr(j + 5) + ':I' + IntToStr(j + 5) +
          ''].select;
        MyExcel.Selection.Borders[9].LineStyle := xlContinuous;
        MyExcel.Selection.Borders[9].Weight := xlThin;
        MyExcel.Selection.Font.Size := 14;

      end;
      Inc(i);
    end;

    flName := vDir + vKurators + '_' + vTitle + '.xlsx';
    SaveWorkBook(flName, 1); // ��������� ����
    Result := true;

  end;
  StopExcel;
  ExcelApp := Unassigned;
  Sheets := Unassigned;
end;

{ * ------------------ ����� ������� � ������� ���������� ���� -----------
  Tabl - �������
  Pole - ���� � ������� �����
  NaNk -  ��� ������
}
function SearchPoziciy(Tabl, pole, NaNk: string): boolean;
begin
  with DM.qQuery do
  begin
    Active := false;
    Close;
    SQL.Clear;
    SQL.Add('SELECT * From ' + Tabl + ' Where (((' + Tabl + '.[' + pole +
      ']) = "' + NaNk + '"))');
    Active := true;
    if RecordCount > 0 then
      Result := true
    else
      Result := false;
  end;
end;

// ����� �������� ���� � �������
function SearchPoziciyString(Tabl, pole, Values, RezPole: string): String;
begin
  with DM.qQuery do
  begin
    Active := false;
    Close;
    SQL.Clear;
    SQL.Add('SELECT * From ' + Tabl + ' Where (((' + Tabl + '.[' + pole +
      ']) = "' + Values + '"))');
    Active := true;

    if RecordCount > 0 then
      Result := FieldByName(RezPole).AsString
    else
      Result := 'Error ����� �� ��������';
  end;
end;

{ * ------ ������� ����� ������ � ������� ---------------------------------
  Tabl - �������
  _poles - �������� ����� ����� ������� ������ �������
  _fio - ���
  _inn - ���
  _isObrok - ��������� ������������� ������������ �����
  _ostatok - ���������� �������� �� ����������� ������������
}
function InsertRecord(Tabl, poles, _fio, _inn: String; _isObrok: boolean;
  _ostatok: integer): boolean;
begin
  with DM.qQuery do
  begin
    Active := false;
    Close;
    SQL.Clear;
    SQL.Add('INSERT INTO ' + Tabl + '(' + poles +
      ' ) VALUES (:fio, :inn, :obrok, :ostatok)');
    Parameters.ParamByName('fio').value := _fio;
    Parameters.ParamByName('inn').value := _inn;
    Parameters.ParamByName('obrok').value := false;
    Parameters.ParamByName('ostatok').value := 0;
    ExecSQL;
    Close;
  end;
end;

{
  provider = MySQL,
  user = hesed_test,
  psw = +0aX3z%3oE,
  server = hesed.mysql.ukraine.com.ua,
  db = hesed_test,
  port = 3306
}
function ConnectBD(provider, user, psw, server, db, port: string): boolean;
begin
  Result := false;
  with DM do
  begin
    UniConnection.ConnectString := 'Provider Name=' + provider + ';User ID=' +
      user + ';Password=' + psw + ';Data Source=' + server + ';Database=' + db +
      ';Port=' + port + '';

    UniConnection.Connected := true;
    if UniConnection.Connected then
      Result := true;
  end;
end;

end.
