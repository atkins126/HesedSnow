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
    // procedure AutoStringGridWidth; // �������� ����� �������

  end;

implementation

{$R *.dfm}

uses uMyProcedure, uMainForm, uDataModul, uMyExcel;
{ TfrmSLG }

var
  SourceCol, SourceRow: Integer; // ��������� ������� � ������ ����������

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
begin // �������� ������
  inherited;
  if (StringGrid.RowCount > 1) then
    Button4.Enabled := true
  else
    Button4.Enabled := false;
end;

procedure TfrmSLG.btnDelStrokaUpdate(Sender: TObject);
begin // ������� ������
  inherited;
  if (StringGrid.RowCount > 1) then
    Button3.Enabled := true
  else
    Button3.Enabled := false;
end;

procedure TfrmSLG.btnExportExcelUpdate(Sender: TObject);
begin // �������������� � ������
  inherited;
  if (StringGrid.RowCount > 1) then
    Button2.Enabled := true
  else
    Button2.Enabled := false;
end;

procedure TfrmSLG.btnLoadTempFileUpdate(Sender: TObject);
begin // ��������� ��������� ����
  inherited;
  if (StringGrid.RowCount > 1) then
    Button6.Enabled := false
  else
    Button6.Enabled := true;
end;

procedure TfrmSLG.btnSaveTempFileUpdate(Sender: TObject);
begin // ��������� ��������� ����
  inherited;
  if (StringGrid.RowCount > 1) then
    Button5.Enabled := true
  else
    Button5.Enabled := false;
end;
// ========================================================

procedure TfrmSLG.Button2Click(Sender: TObject);
var // ������� � ������
  Sheets, ExcelApp: variant;
  i, j: Integer;
  DirectoryNow, FileNameS: String;

begin
  inherited;
  if uMyExcel.RunExcel(false, false) = true then
    MyExcel.Workbooks.Add; // ��������� ����� �����

  Sheets := MyExcel.Worksheets.Add;
  Sheets.name := 'A1';
  ParametryStr; // �������� ��������
  Sheets.PageSetup.PrintTitleRows := '$2:$2';

  // ��������� ������� ( ����� ����� ������� ) ������ ��������

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

  MyExcel.Range['A1'].value := '����� ������';
  MyExcel.Range['B1'].value := 'JDC ID';
  MyExcel.Range['C1'].value := '���';
  MyExcel.Range['D1'].value := '�������� ������ �������';
  MyExcel.Range['E1'].value := '����������';
  MyExcel.Range['F1'].value := '���������';
  MyExcel.Range['G1'].value := '���� �� �������';
  MyExcel.Range['H1'].value := '����������� ����� ���������';
  MyExcel.Range['I1'].value := '����������';

  for i := 1 to StringGrid.RowCount - 1 do // ������
    for j := 0 to StringGrid.ColCount - 1 do // �������
    begin
      MyExcel.Cells[i + 1, j + 1] := StringGrid.Cells[j, i];
    end;

  DirectoryNow := ExtractFilePath(ParamStr(0))+ '������\���\';

  if not DirectoryExists('DirectoryNow') then ForceDirectories(DirectoryNow);
// ForceDirectories(ExtractFilePath(Application.ExeName) + '/folder1/folder2/newfolder');

  FileNameS := DirectoryNow + '���_' + FormatDateTime('dd.mm.yyyy hh_mm_ss', Now) + '.xlsx';

  uMyExcel.SaveWorkBook(FileNameS, 1);

  Sheets := unassigned;
  ExcelApp := unassigned;
  uMyExcel.StopExcel;

  ShowMessage('������ �������������� � ��������� � ����');

 //   ShellExecute(Handle, 'open', PWideChar(DirectoryNow), nil, nil,SW_SHOWNORMAL);

   ShellExecute(Handle, 'open', PWideChar(DirectoryNow), nil, nil,SW_SHOWNORMAL);

end;

procedure TfrmSLG.Button3Click(Sender: TObject);
var // ������� ���������� �������
  n: Integer;
begin
  inherited;
  // ���� �������� ���� ������, �������� �������� �� ���������
  if (StringGrid.RowCount = 1) then
    Exit;
  // ����� ������ �����, ������� �� ������, ���������� ����������
  for n := StringGrid.Row to StringGrid.RowCount - 2 do
  begin
    StringGrid.Rows[n] := StringGrid.Rows[n + 1];
  end;
  // �������� ��������� ������ �������
  StringGrid.RowCount := StringGrid.RowCount - 1;

end;

procedure TfrmSLG.Button4Click(Sender: TObject);
var
  i, j: Integer;
begin // �������� ������
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
  if (Key = #13) then // ���� ����� Enter
  begin
    last_row := StringGrid.RowCount;
    StringGrid.RowCount := StringGrid.RowCount + 1; // ���������� ������
    StringGrid.Cells[0, last_row] := MyDBGrid1.DataSource.DataSet.FieldByName
      ('RITM').AsString;
    StringGrid.Cells[1, last_row] := MyDBGrid1.DataSource.DataSet.FieldByName
      ('JDCID').AsString;
    StringGrid.Cells[2, last_row] := MyDBGrid1.DataSource.DataSet.FieldByName
      ('FIO').AsString;
    // StringGrid1.Cells[4, last_row] := DBGrid1.DataSource.DataSet.FieldByName('Number').AsString;
    AutoStringGridWidth(StringGrid);
          // ����������� � ����� ����������
      StringGrid.Row := StringGrid.RowCount - 1;
  end;
end;

procedure TfrmSLG.MyDBGrid2KeyPress(Sender: TObject; var Key: Char);
var // ���� ������� ������
  select_row: Integer;
  colvo, upakovka: Integer;
  buttonSelected: Integer;
  izdelie: string;
  Price_inch: real;
label ups;

begin
  inherited;
  if (Key = #13) then // ���� ����� Enter
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
        buttonSelected := MessageDlg('������������� ������?', mtCustom,
          [mbYes, mbNo], 0);

        if buttonSelected = mrYes then
        begin
          StringGrid.Row := select_row;
          StringGrid.Cells[3, select_row] := izdelie;
          // Price_inch       upakovka
          colvo := StrToInt(InputBox(izdelie, '������� ���-��', ''));
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
        colvo := StrToInt(InputBox(izdelie, '������� ���-��', ''));
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
  // ������� ��������� ������ ��������� ������� ( ' ) � ������ � ����� ������
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
  // ������� ��������� ������ ��������� ������� ( ' ) � ������ � ����� ������
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
  // ��������� �����
  StringGrid.Cells[0, 0] := '����� ������';
  StringGrid.Cells[1, 0] := 'JDC ID';
  StringGrid.Cells[2, 0] := '���';
  StringGrid.Cells[3, 0] := '�������� ������ �������';
  StringGrid.ColWidths[3] := 250;
  StringGrid.Cells[4, 0] := '����������';
  StringGrid.Cells[5, 0] := '���������';
  StringGrid.Cells[6, 0] := '���� �� �������';
  StringGrid.Cells[7, 0] := '����������� ����� ���������';
  StringGrid.Cells[8, 0] := '����������';
end;

procedure TfrmSLG.SpeedButton1Click(Sender: TObject);
var
  z, col, m, n: Integer;
  CollectionNameTable: TDictionary<string, Integer>;
begin
  inherited;
  // ImportFromExcel(table, spisok poley, spisok typeDannih);
  // ��� ���������� ���� ��� ����� ��� ���������

  with myForm, DM do
  begin
    OpenDialog.Filter := '����� MS Excel|*.xls;*.xlsx|';
    if not OpenDialog.Execute then
      Exit;
    // �������� �� ������� � ������ Excel
    if uMyExcel.RunExcel(false, false) = true then
      // ��������� ����� Excel
      if uMyExcel.OpenWorkBook(OpenDialog.FileName, false) then
      begin
        ProgressBar.Visible := true;
        MyExcel.ActiveWorkBook.Sheets[1];

        // ��������� ����������� �������
        col := MyExcel.ActiveCell.SpecialCells($000000B).Column;

        // ------------ ���������� ��������� ������� �������� �������� ---------
        CollectionNameTable := TDictionary<string, Integer>.Create();
        for z := 1 to col do
        begin
          if not CollectionNameTable.ContainsKey(MyExcel.Cells[1, z].value) then
            CollectionNameTable.Add(MyExcel.Cells[1, z].value, z)
          else
            CollectionNameTable.Add(MyExcel.Cells[1, z].value + IntToStr(z), z);
        end;
        // ================= ����� ������� ���  �������� ======================

        m := 2; // �������� ���������� �� 2-� ������, �������� ��������� �������
        // ��������� ����������� ������
        n := MyExcel.ActiveCell.SpecialCells($000000B).Row;
        n := n + 1;

        ProgressBar.Min := 0;
        ProgressBar.Max := n;
        ProgressBar.Position := 1;

        // ������� ������� ������� ������ �������
        CleanOutTable('slg_items');
        tSLG.Open;

        while m <> n do // ���� ������� �� ������� EXCEL
        begin
          tSLG.Insert;

          tSLG.FieldByName('Name_SLG').AsString :=
            MyExcel.Cells[m, StrToInt(CollectionNameTable.Items['�������� ���']
            .ToString)].value;

          if MyExcel.Cells
            [m, StrToInt(CollectionNameTable.Items['�������������'].ToString)
            ].value = '�����' then
            tSLG.FieldByName('Active').AsBoolean := true
          else
            tSLG.FieldByName('Active').AsBoolean := false;

          tSLG.FieldByName('Price_inch').AsCurrency :=
            StrToCurr(MyExcel.Cells[m,
            StrToInt(CollectionNameTable.Items['���� �� �������']
            .ToString)].value);

          tSLG.FieldByName('upakovka').AsInteger :=
            ParseStringUpakovka(MyExcel.Cells[m,
            StrToInt(CollectionNameTable.Items['�������� ���'].ToString)].value)
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

    OpenDialog.Filter := '����� MS Excel|*.xls;*.xlsx|';
    if not OpenDialog.Execute then
      Exit;
    // �������� �� ������� � ������ Excel
    if uMyExcel.RunExcel(false, false) = true then
      // ��������� ����� Excel
      if uMyExcel.OpenWorkBook(OpenDialog.FileName, false) then
      begin
        ProgressBar.Visible := true;
        MyExcel.ActiveWorkBook.Sheets[1];

        // ��������� ����������� �������
        col := MyExcel.ActiveCell.SpecialCells($000000B).Column;

        // ------------ ���������� ��������� ������� �������� �������� -------------

        CollectionNameTable := TDictionary<string, Integer>.Create();
        for z := 1 to col do
        begin
          if not CollectionNameTable.ContainsKey(MyExcel.Cells[1, z].value) then
            CollectionNameTable.Add(MyExcel.Cells[1, z].value, z)
          else
            CollectionNameTable.Add(MyExcel.Cells[1, z].value + z.ToString, z);
        end;
        // ================= ����� ������� ���  �������� ======================

        m := 2; // �������� ���������� �� 2-� ������, �������� ��������� �������
        // ��������� ����������� ������
        n := MyExcel.ActiveCell.SpecialCells($000000B).Row;
        n := n + 1;

        ProgressBar.Min := 0;
        ProgressBar.Max := n;
        ProgressBar.Position := 1;

        // ������� ������� ������� ������ �������
        CleanOutTable('Uslugy');
        tUslugy.Open;

        while m <> n do // ���� ������� �� ������� EXCEL
        begin
          tUslugy.Insert;

          tUslugy.FieldByName('RITM').AsString :=
            MyExcel.Cells
            [m, StrToInt(CollectionNameTable.Items['�����'].ToString)].value;

          tUslugy.FieldByName('JDCID').AsString :=
            MyExcel.Cells
            [m, StrToInt(CollectionNameTable.Items['JDC ID'].ToString)].value;

          tUslugy.FieldByName('FIO').AsString :=
            MyExcel.Cells
            [m, StrToInt(CollectionNameTable.Items['������'].ToString)].value;
          // [m, StrToInt(CollectionNameTable.Items['���'].ToString)].value;

          tUslugy.FieldByName('SABA').AsCurrency :=
            MyExcel.Cells
            [m, StrToInt(CollectionNameTable.Items['��������� ������']
            .ToString)].value;

          tUslugy.FieldByName('City').AsString :=
            MyExcel.Cells
            [m, StrToInt(CollectionNameTable.Items['�����'].ToString)].value;

          tUslugy.FieldByName('Number').AsString :=
            MyExcel.Cells[m, StrToInt(CollectionNameTable.Items['����������']
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

// ---------- ��������� Drag and Drop -------------------------

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
begin // ���������� �� �����
  inherited;
  StringGrid.MouseToCell(X, Y, Acol, ARow);
  Accept := (Sender = Source) and (Acol > 0) and (ARow > 0) or
    (Source is TMyDBGrid);
end;

procedure TfrmSLG.StringGridMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  inherited;
  // ������������� ���������� ���� X, Y � (��� StringGrid) ��������� ������� �������� � �����
  StringGrid.MouseToCell(X, Y, SourceCol, SourceRow);
  // ��������� ��������������, ������ ���� ���� ������ ���������� ������ (������ �� ������������� �������� � �������)
  if (SourceCol > 0) and (SourceRow > 0) then
    // ������� �������������� ����� ����������� ���� �� 4 �������.
    StringGrid.BeginDrag(false, 4);
end;

procedure TfrmSLG.StringGridDragDrop(Sender, Source: TObject; X, Y: Integer);
var
  DestCol, DestRow, i, colvo, upakovka: Integer;
  izdelie: String;
  Price_inch: real;
begin // ����� ��������������
  inherited;
  with Sender as TStringGrid do
  begin
    // �������������� ���������� ����.
    StringGrid.MouseToCell(X, Y, DestCol, DestRow);

    // ����������� ���������� �� ��������� � ����� ����������
    if Source = StringGrid then // ---- �������� ���������� ---------
    begin
      for i := 3 to StringGrid.ColCount - 1 do
      begin
        StringGrid.Cells[i, DestRow] := StringGrid.Cells[i, SourceRow];
      end;

    end;
    // Drag into DBGrids
    if Source = MyDBGrid1 then // ------- ����� ���� � ��� ------
    begin
      DestCol := 1;
      DestRow := StringGrid.RowCount;
      StringGrid.RowCount := StringGrid.RowCount + 1; // ���������� ������

      StringGrid.Cells[DestCol - 1, DestRow] := MyDBGrid1.Fields[2].AsString;
      StringGrid.Cells[DestCol, DestRow] := MyDBGrid1.Fields[1].AsString;
      StringGrid.Cells[DestCol + 1, DestRow] := MyDBGrid1.Fields[0].AsString;
      AutoStringGridWidth(StringGrid);
      // ����������� � ����� ����������
      StringGrid.Row := StringGrid.RowCount - 1;
    end;
    if Source = MyDBGrid2 then // ---- ������ ���� -------------------------
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

        colvo := StrToInt(InputBox(izdelie, '������� ���-��', ''));

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
