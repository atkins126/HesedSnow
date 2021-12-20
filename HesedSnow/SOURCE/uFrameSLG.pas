unit uFrameSLG;

interface

uses uFrameCustom, Data.DB, Vcl.DBGrids, acDBGrid, MyDBGrid, Vcl.Dialogs,
  System.Classes, System.Actions, Vcl.ActnList, Vcl.Grids, JvExGrids,
  JvStringGrid, Vcl.StdCtrls, sEdit, Vcl.ExtCtrls, Vcl.Controls, Vcl.Buttons,
  sPanel, sFrameAdapter, System.SysUtils, System.Generics.Collections,
  Vcl.Forms, System.Variants, System.UITypes, Winapi.ShellAPI, Winapi.Windows;


type
  TfrmSLG = class(TCustomInfoFrame)
    sPanel1: TsPanel;
    sPanel2: TsPanel;
    sPanel3: TsPanel;
    acList: TActionList;
    btnDelStroka: TAction;
    btnClearList: TAction;
    btnExportExcel: TAction;
    btnLoadTempFile: TAction;
    btnSaveTempFile: TAction;
    SpeedButton2: TSpeedButton;
    SpeedButton1: TSpeedButton;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    StringGrid: TJvStringGrid;
    Button3: TButton;
    Button4: TButton;
    Button5: TButton;
    Button6: TButton;
    OpenDialog: TOpenDialog;
    Button2: TButton;
    Splitter1: TSplitter;
    Splitter2: TSplitter;
    sEdit1: TsEdit;
    sEdit2: TsEdit;
    MyDBGrid1: TMyDBGrid;
    MyDBGrid2: TMyDBGrid;
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure sEdit1Change(Sender: TObject);
    procedure sEdit2Change(Sender: TObject);
    procedure btnDelStrokaUpdate(Sender: TObject);
    procedure btnClearListUpdate(Sender: TObject);
    procedure btnExportExcelUpdate(Sender: TObject);
    procedure btnLoadTempFileUpdate(Sender: TObject);
    procedure btnSaveTempFileUpdate(Sender: TObject);
    procedure MyDBGrid1KeyPress(Sender: TObject; var Key: Char);
    procedure MyDBGrid1MouseMove(Sender: TObject; Shift: TShiftState;
      X, Y: Integer);
    procedure StringGridMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure StringGridDragOver(Sender, Source: TObject; X, Y: Integer;
      State: TDragState; var Accept: Boolean);
    procedure StringGridDragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure MyDBGrid2KeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
    procedure ShapkaStringGrida;

  public
    { Public declarations }
    procedure AfterCreation; override;
    // procedure AutoStringGridWidth; // подгонка ширині колонок

  end;

implementation

{$R *.dfm}

uses uMyProcedure, uMainForm, uDataModul, uMyExcel;
{ TfrmSLG }

var
  SourceCol, SourceRow: Integer; // Источника колонка и строка координаты

procedure TfrmSLG.AfterCreation;
begin
  inherited;
  ShapkaStringGrida;
  StringGrid.RowCount := 1;
  DM.qUslugy.Active := true;
  DM.qQslg.Active := true;
end;

{ procedure TfrmSLG.AutoStringGridWidth;
  var
  X, Y, w: Integer;
  MaxWidth: Integer;
  begin
  with StringGrid do
  ClientHeight := DefaultRowHeight * RowCount + 5;
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
  end; }

// ------- actions -----------------------------------
procedure TfrmSLG.btnClearListUpdate(Sender: TObject);
begin // Очистить список
  inherited;
  if (StringGrid.RowCount > 1) then
    Button4.Enabled := true
  else
    Button4.Enabled := false;
end;

procedure TfrmSLG.btnDelStrokaUpdate(Sender: TObject);
begin // Удалить строку
  inherited;
  if (StringGrid.RowCount > 1) then
    Button3.Enabled := true
  else
    Button3.Enabled := false;
end;

procedure TfrmSLG.btnExportExcelUpdate(Sender: TObject);
begin // Єкспортировать в ексель
  inherited;
  if (StringGrid.RowCount > 1) then
    Button2.Enabled := true
  else
    Button2.Enabled := false;
end;

procedure TfrmSLG.btnLoadTempFileUpdate(Sender: TObject);
begin // Загрузить временный файл
  inherited;
  if (StringGrid.RowCount > 1) then
    Button6.Enabled := false
  else
    Button6.Enabled := true;
end;

procedure TfrmSLG.btnSaveTempFileUpdate(Sender: TObject);
begin // сохранить временный файл
  inherited;
  if (StringGrid.RowCount > 1) then
    Button5.Enabled := true
  else
    Button5.Enabled := false;
end;
// ========================================================

procedure TfrmSLG.Button2Click(Sender: TObject);
var // Експорт в Ексель
  Sheets, ExcelApp: variant;
  i, j: Integer;
  DirectoryNow, FileNameS: String;

begin
  inherited;
  if uMyExcel.RunExcel(false, false) = true then
    MyExcel.Workbooks.Add; // добавляем новую книгу

  Sheets := MyExcel.Worksheets.Add;
  Sheets.name := 'A1';
  ParametryStr; // Парметры страницы
  Sheets.PageSetup.PrintTitleRows := '$2:$2';

  // Параметры таблицы ( цифра номер столбца ) ширины столбцов

  ExcelApp := MyExcel.ActiveWorkBook.Worksheets[1].columns;

  ExcelApp.columns[1].columnwidth := 15;
  ExcelApp.columns[2].columnwidth := 15;
  ExcelApp.columns[2].NumberFormat := '@';
  ExcelApp.columns[3].columnwidth := 40;
  ExcelApp.columns[4].columnwidth := 40;
  ExcelApp.columns[5].columnwidth := 10;
  ExcelApp.columns[6].columnwidth := 10;
  ExcelApp.columns[7].columnwidth := 10;
  ExcelApp.columns[8].columnwidth := 10;
  ExcelApp.columns[9].columnwidth := 10;

  MyExcel.Range['A1'].value := 'Номер услуги';
  MyExcel.Range['B1'].value := 'JDC ID';
  MyExcel.Range['C1'].value := 'ФИО';
  MyExcel.Range['D1'].value := 'Средство личной гигиены';
  MyExcel.Range['E1'].value := 'Количество';
  MyExcel.Range['F1'].value := 'Стоимость';
  MyExcel.Range['G1'].value := 'Цена за единицу';
  MyExcel.Range['H1'].value := 'Пересчитать общую стоимость';
  MyExcel.Range['I1'].value := 'Примечание';

  for i := 1 to StringGrid.RowCount - 1 do // строки
    for j := 0 to StringGrid.ColCount - 1 do // столбцы
    begin
      MyExcel.Cells[i + 1, j + 1] := StringGrid.Cells[j, i];
    end;

  DirectoryNow := ExtractFilePath(ParamStr(0))+ 'Данные\СЛГ\';

  if not DirectoryExists('DirectoryNow') then ForceDirectories(DirectoryNow);
// ForceDirectories(ExtractFilePath(Application.ExeName) + '/folder1/folder2/newfolder');

  FileNameS := DirectoryNow + 'СЛГ_' + FormatDateTime('dd.mm.yyyy hh_mm_ss', Now) + '.xlsx';

  uMyExcel.SaveWorkBook(FileNameS, 1);

  Sheets := unassigned;
  ExcelApp := unassigned;
  uMyExcel.StopExcel;

  ShowMessage('Данные экспортированы и сохранены в файл');

 //   ShellExecute(Handle, 'open', PWideChar(DirectoryNow), nil, nil,SW_SHOWNORMAL);

   ShellExecute(Handle, 'open', PWideChar(DirectoryNow), nil, nil,SW_SHOWNORMAL);

end;

procedure TfrmSLG.Button3Click(Sender: TObject);
var // Удаляем выделенную позицию
  n: Integer;
begin
  inherited;
  // Если осталась одна строка, операцию удаления не выполнять
  if (StringGrid.RowCount = 1) then
    Exit;
  // Сдвиг строки вверх, начиная со строки, содержащей выделенную
  for n := StringGrid.Row to StringGrid.RowCount - 2 do
  begin
    StringGrid.Rows[n] := StringGrid.Rows[n + 1];
  end;
  // Удаление последней строки таблицы
  StringGrid.RowCount := StringGrid.RowCount - 1;

end;

procedure TfrmSLG.Button4Click(Sender: TObject);
var
  i, j: Integer;
begin // Очистить список
  inherited;
  with StringGrid do
  begin
    for i := FixedCols to ColCount - 1 do
    begin
      // for j:=FixedRows to RowCount-1 do
      for j := 1 to RowCount - 1 do
      begin
        Cells[i, j] := '';
      end;
    end;
  end;

  StringGrid.RowCount := 1;
end;

procedure TfrmSLG.Button5Click(Sender: TObject);
begin
  inherited;
  StringGrid.SaveToFile(ExtractFilePath(Application.ExeName) +
    'slg_chernovik.txt');
end;

procedure TfrmSLG.Button6Click(Sender: TObject);
begin
  inherited;
  if FileExists(ExtractFilePath(Application.ExeName) + 'slg_chernovik.txt') then
    StringGrid.LoadFromFile(ExtractFilePath(Application.ExeName) +
      'slg_chernovik.txt');
end;

procedure TfrmSLG.MyDBGrid1KeyPress(Sender: TObject; var Key: Char);
var
  last_row: Integer;
begin
  inherited;
  if (Key = #13) then // если нажат Enter
  begin
    last_row := StringGrid.RowCount;
    StringGrid.RowCount := StringGrid.RowCount + 1; // Добавление строки
    StringGrid.Cells[0, last_row] := MyDBGrid1.DataSource.DataSet.FieldByName
      ('RITM').AsString;
    StringGrid.Cells[1, last_row] := MyDBGrid1.DataSource.DataSet.FieldByName
      ('JDCID').AsString;
    StringGrid.Cells[2, last_row] := MyDBGrid1.DataSource.DataSet.FieldByName
      ('FIO').AsString;
    // StringGrid1.Cells[4, last_row] := DBGrid1.DataSource.DataSet.FieldByName('Number').AsString;
    AutoStringGridWidth(StringGrid);
          // переместить в конец стрингрида
      StringGrid.Row := StringGrid.RowCount - 1;
  end;
end;

procedure TfrmSLG.MyDBGrid2KeyPress(Sender: TObject; var Key: Char);
var // вібор средств гигиені
  select_row: Integer;
  colvo, upakovka: Integer;
  buttonSelected: Integer;
  izdelie: string;
  Price_inch: real;
label ups;

begin
  inherited;
  if (Key = #13) then // если нажат Enter
  begin
    select_row := StringGrid.Tag;
    if select_row = 0 then
      select_row := select_row + 1;
    if select_row <> 0 then
    begin

      sEdit1.text := '';
      izdelie := MyDBGrid2.DataSource.DataSet.FieldByName('Name_SLG').AsString;
      upakovka := MyDBGrid2.DataSource.DataSet.FieldByName('upakovka')
        .AsInteger;
      Price_inch := MyDBGrid2.DataSource.DataSet.FieldByName('Price_inch')
        .AsCurrency;

    ups:
      if (StringGrid.Cells[3, select_row] <> '') then
      begin
        buttonSelected := MessageDlg('Редактировать запись?', mtCustom,
          [mbYes, mbNo], 0);

        if buttonSelected = mrYes then
        begin
          StringGrid.Row := select_row;
          StringGrid.Cells[3, select_row] := izdelie;
          // Price_inch       upakovka
          colvo := StrToInt(InputBox(izdelie, 'Введите кол-во', ''));
          StringGrid.Cells[5, select_row] :=
            ((colvo / upakovka) * Price_inch).ToString;

          StringGrid.Cells[7, select_row] := '0';

          AutoStringGridWidth(StringGrid);
          if select_row <> StringGrid.RowCount - 1 then
            StringGrid.Row := StringGrid.Row + 1;
        end
        else
        begin
          select_row := select_row + 1;
          goto ups;
        end;

      end
      else
      begin
        StringGrid.Row := select_row;
        StringGrid.Cells[3, select_row] := izdelie;
        // Price_inch       upakovka
        colvo := StrToInt(InputBox(izdelie, 'Введите кол-во', ''));
        StringGrid.Cells[5, select_row] :=
          ((colvo / upakovka) * Price_inch).ToString;
        StringGrid.Cells[4, select_row] := colvo.ToString;
        StringGrid.Cells[7, select_row] := '0';
        AutoStringGridWidth(StringGrid);
        if select_row <> StringGrid.RowCount - 1 then
          StringGrid.Row := StringGrid.Row + 1;
      end;

    end;

  end;

end;

procedure TfrmSLG.sEdit1Change(Sender: TObject);
var
  a1, a2: String;
begin
  inherited;
  a1 := sEdit1.text + '%';
  a2 := QuotedStr(a1);
  // Функция добавляет символ одиночной кавычки ( ' ) в начало и конец строки
  with DM.qUslugy do
  begin
    Close;
    SQL.Clear;
    SQL.Add('SELECT Uslugy.[FIO], Uslugy.[JDCID],Uslugy.[RITM],Uslugy.[Number],Uslugy.[SABA] FROM Uslugy ');
    SQL.Add('WHERE Uslugy.[FIO] like ' + a2 + ' ORDER BY Uslugy.[FIO]');
    Open;
  end;
end;

procedure TfrmSLG.sEdit2Change(Sender: TObject);
var
  a1, a2: String;
begin
  inherited;
  a1 := sEdit2.text + '%';
  // Функция добавляет символ одиночной кавычки ( ' ) в начало и конец строки
  a2 := QuotedStr(a1);
  with DM.qQslg do
  begin
    Close;
    SQL.Clear;
    SQL.Add('SELECT * ');
    SQL.Add('FROM slg_items ');
    SQL.Add('WHERE slg_items.[Name_SLG] like ' + a2 +
      '  ORDER BY slg_items.[Name_SLG]');
    Open;
  end;
end;

procedure TfrmSLG.ShapkaStringGrida;
begin
  // Заголовки грида
  StringGrid.Cells[0, 0] := 'Номер услуги';
  StringGrid.Cells[1, 0] := 'JDC ID';
  StringGrid.Cells[2, 0] := 'ФИО';
  StringGrid.Cells[3, 0] := 'Средство личной гигиены';
  StringGrid.ColWidths[3] := 250;
  StringGrid.Cells[4, 0] := 'Количество';
  StringGrid.Cells[5, 0] := 'Стоимость';
  StringGrid.Cells[6, 0] := 'Цена за единицу';
  StringGrid.Cells[7, 0] := 'Пересчитать общую стоимость';
  StringGrid.Cells[8, 0] := 'Примечание';
end;

procedure TfrmSLG.SpeedButton1Click(Sender: TObject);
var
  z, col, m, n: Integer;
  CollectionNameTable: TDictionary<string, Integer>;
begin
  inherited;
  // ImportFromExcel(table, spisok poley, spisok typeDannih);
  // или переделать чтоб все данніе біли стринговіе

  with myForm, DM do
  begin
    OpenDialog.Filter := 'Файлы MS Excel|*.xls;*.xlsx|';
    if not OpenDialog.Execute then
      Exit;
    // проверка на инсталл и запуск Excel
    if uMyExcel.RunExcel(false, false) = true then
      // открываем книгу Excel
      if uMyExcel.OpenWorkBook(OpenDialog.FileName, false) then
      begin
        ProgressBar.Visible := true;
        MyExcel.ActiveWorkBook.Sheets[1];

        // последняя заполненная колонка
        col := MyExcel.ActiveCell.SpecialCells($000000B).Column;

        // ------------ пробежимся расставим индексы названий столбцов ---------
        CollectionNameTable := TDictionary<string, Integer>.Create();
        for z := 1 to col do
        begin
          if not CollectionNameTable.ContainsKey(MyExcel.Cells[1, z].value) then
            CollectionNameTable.Add(MyExcel.Cells[1, z].value, z)
          else
            CollectionNameTable.Add(MyExcel.Cells[1, z].value + IntToStr(z), z);
        end;
        // ================= конец пробега для  СТОЛБЦОВ ======================

        m := 2; // начинаем считывание со 2-й строки, оставляя заголовок колонки
        // последняя заполненная строка
        n := MyExcel.ActiveCell.SpecialCells($000000B).Row;
        n := n + 1;

        ProgressBar.Min := 0;
        ProgressBar.Max := n;
        ProgressBar.Position := 1;

        // очищаем таблицу средств личной гигиены
        CleanOutTable('slg_items');
        tSLG.Open;

        while m <> n do // цикл внешний по записям EXCEL
        begin
          tSLG.Insert;

          tSLG.FieldByName('Name_SLG').AsString :=
            MyExcel.Cells[m, StrToInt(CollectionNameTable.Items['Название СЛГ']
            .ToString)].value;

          if MyExcel.Cells
            [m, StrToInt(CollectionNameTable.Items['Действительно'].ToString)
            ].value = 'ВЕРНО' then
            tSLG.FieldByName('Active').AsBoolean := true
          else
            tSLG.FieldByName('Active').AsBoolean := false;

          tSLG.FieldByName('Price_inch').AsCurrency :=
            StrToCurr(MyExcel.Cells[m,
            StrToInt(CollectionNameTable.Items['Цена за единицу']
            .ToString)].value);

          tSLG.FieldByName('upakovka').AsInteger :=
            ParseStringUpakovka(MyExcel.Cells[m,
            StrToInt(CollectionNameTable.Items['Название СЛГ'].ToString)].value)
            .ToInteger;

          tSLG.Post;
          Inc(m);
          // Application.ProcessMessages;
          Sleep(25);
          ProgressBar.Position := m;
        end;

      end;

    MyExcel.Application.DisplayAlerts := false;
    StopExcel;
    CollectionNameTable.Clear;
    CollectionNameTable.Free;
    DM.tUslugy.Active := false;
    myForm.ProgressBar.Visible := false;
    DM.qQslg.Active := false;
    DM.qQslg.Active := true;
  end;

end;

procedure TfrmSLG.SpeedButton2Click(Sender: TObject);
var
  z, col, m, n: Integer;
  CollectionNameTable: TDictionary<string, Integer>;
begin
  inherited;

  with myForm, DM do
  begin

    OpenDialog.Filter := 'Файлы MS Excel|*.xls;*.xlsx|';
    if not OpenDialog.Execute then
      Exit;
    // проверка на инсталл и запуск Excel
    if uMyExcel.RunExcel(false, false) = true then
      // открываем книгу Excel
      if uMyExcel.OpenWorkBook(OpenDialog.FileName, false) then
      begin
        ProgressBar.Visible := true;
        MyExcel.ActiveWorkBook.Sheets[1];

        // последняя заполненная колонка
        col := MyExcel.ActiveCell.SpecialCells($000000B).Column;

        // ------------ пробежимся расставим индексы названий столбцов -------------

        CollectionNameTable := TDictionary<string, Integer>.Create();
        for z := 1 to col do
        begin
          if not CollectionNameTable.ContainsKey(MyExcel.Cells[1, z].value) then
            CollectionNameTable.Add(MyExcel.Cells[1, z].value, z)
          else
            CollectionNameTable.Add(MyExcel.Cells[1, z].value + z.ToString, z);
        end;
        // ================= конец пробега для  СТОЛБЦОВ ======================

        m := 2; // начинаем считывание со 2-й строки, оставляя заголовок колонки
        // последняя заполненная строка
        n := MyExcel.ActiveCell.SpecialCells($000000B).Row;
        n := n + 1;

        ProgressBar.Min := 0;
        ProgressBar.Max := n;
        ProgressBar.Position := 1;

        // очищаем таблицу средств личной гигиены
        CleanOutTable('Uslugy');
        tUslugy.Open;

        while m <> n do // цикл внешний по записям EXCEL
        begin
          tUslugy.Insert;

          tUslugy.FieldByName('RITM').AsString :=
            MyExcel.Cells
            [m, StrToInt(CollectionNameTable.Items['Номер'].ToString)].value;

          tUslugy.FieldByName('JDCID').AsString :=
            MyExcel.Cells
            [m, StrToInt(CollectionNameTable.Items['JDC ID'].ToString)].value;

          tUslugy.FieldByName('FIO').AsString :=
            MyExcel.Cells
            [m, StrToInt(CollectionNameTable.Items['Клиент'].ToString)].value;
          // [m, StrToInt(CollectionNameTable.Items['ФИО'].ToString)].value;

          tUslugy.FieldByName('SABA').AsCurrency :=
            MyExcel.Cells
            [m, StrToInt(CollectionNameTable.Items['Стоимость услуги']
            .ToString)].value;

          tUslugy.FieldByName('City').AsString :=
            MyExcel.Cells
            [m, StrToInt(CollectionNameTable.Items['Город'].ToString)].value;

          tUslugy.FieldByName('Number').AsString :=
            MyExcel.Cells[m, StrToInt(CollectionNameTable.Items['Количество']
            .ToString)].value;

          tUslugy.Post;
          Inc(m);
          // Application.ProcessMessages;
          Sleep(25);
          ProgressBar.Position := m;
        end;

      end;

    MyExcel.Application.DisplayAlerts := false;
    StopExcel;
    CollectionNameTable.Clear;
    CollectionNameTable.Free;
    DM.tUslugy.Active := false;
    myForm.ProgressBar.Visible := false;
    DM.qUslugy.Active := false;
    DM.qUslugy.Active := true;
  end;

end;

// ---------- реализуем Drag and Drop -------------------------

procedure TfrmSLG.MyDBGrid1MouseMove(Sender: TObject; Shift: TShiftState;
  X, Y: Integer);
begin
  inherited;
  if ssLeft in Shift then
    TMyDBGrid(Sender).BeginDrag(false);
end;

procedure TfrmSLG.StringGridDragOver(Sender, Source: TObject; X, Y: Integer;
  State: TDragState; var Accept: Boolean);
var
  Acol, ARow: Integer;
begin // разрешение на прием
  inherited;
  StringGrid.MouseToCell(X, Y, Acol, ARow);
  Accept := (Sender = Source) and (Acol > 0) and (ARow > 0) or
    (Source is TMyDBGrid);
end;

procedure TfrmSLG.StringGridMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  inherited;
  // Преобразовать координаты мыши X, Y в (для StringGrid) связанных номеров столбцов и строк
  StringGrid.MouseToCell(X, Y, SourceCol, SourceRow);
  // Разрешить перетаскивание, только если была нажата приемлемая ячейка (ячейка за фиксированным столбцом и строкой)
  if (SourceCol > 0) and (SourceRow > 0) then
    // Начните перетаскивание после перемещения мыши на 4 пикселя.
    StringGrid.BeginDrag(false, 4);
end;

procedure TfrmSLG.StringGridDragDrop(Sender, Source: TObject; X, Y: Integer);
var
  DestCol, DestRow, i, colvo, upakovka: Integer;
  izdelie: String;
  Price_inch: real;
begin // после перетаскивания
  inherited;
  with Sender as TStringGrid do
  begin
    // конвертировать координаты мыши.
    StringGrid.MouseToCell(X, Y, DestCol, DestRow);

    // Переместить содержимое из источника в место назначения
    if Source = StringGrid then // ---- СЕРЕДИНА стринггрид ---------
    begin
      for i := 3 to StringGrid.ColCount - 1 do
      begin
        StringGrid.Cells[i, DestRow] := StringGrid.Cells[i, SourceRow];
      end;

    end;
    // Drag into DBGrids
    if Source = MyDBGrid1 then // ------- СЛЕВА грид с ФИО ------
    begin
      DestCol := 1;
      DestRow := StringGrid.RowCount;
      StringGrid.RowCount := StringGrid.RowCount + 1; // Добавление строки

      StringGrid.Cells[DestCol - 1, DestRow] := MyDBGrid1.Fields[2].AsString;
      StringGrid.Cells[DestCol, DestRow] := MyDBGrid1.Fields[1].AsString;
      StringGrid.Cells[DestCol + 1, DestRow] := MyDBGrid1.Fields[0].AsString;
      AutoStringGridWidth(StringGrid);
      // переместить в конец стрингрида
      StringGrid.Row := StringGrid.RowCount - 1;
    end;
    if Source = MyDBGrid2 then // ---- СПРАВА грид -------------------------
    begin
      if StringGrid.RowCount > 1 then
      begin
        DestCol := 3;
        StringGrid.Cells[DestCol, DestRow] := MyDBGrid2.Fields[0].AsString;
        // StringGrid.Cells[DestCol + 1, DestRow] := MyDBGrid2.Fields[2].AsString;

        izdelie := MyDBGrid2.DataSource.DataSet.FieldByName('Name_SLG')
          .AsString;
        upakovka := MyDBGrid2.DataSource.DataSet.FieldByName('upakovka')
          .AsInteger;
        Price_inch := MyDBGrid2.DataSource.DataSet.FieldByName('Price_inch')
          .AsCurrency;

        colvo := StrToInt(InputBox(izdelie, 'Введите кол-во', ''));

        StringGrid.Cells[5, DestRow] :=
          ((colvo / upakovka) * Price_inch).ToString;

        StringGrid.Cells[4, DestRow] := colvo.ToString;

        StringGrid.Cells[7, DestRow] := '0';

        AutoStringGridWidth(StringGrid);

      end;
    end;

  end;

end;

end.
