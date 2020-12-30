unit uFrameSLG;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, System.Generics.Collections,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  uFrameCustom, sFrameAdapter, Data.DB, Vcl.StdCtrls, Vcl.Grids, JvExGrids,
  JvStringGrid, Vcl.DBGrids, Vcl.ExtCtrls, Vcl.Buttons, System.Actions,
  Vcl.ActnList, sPanel, acDBGrid, sEdit;

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
    SpeedButton3: TSpeedButton;
    SpeedButton1: TSpeedButton;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    StringGrid1: TJvStringGrid;
    Button3: TButton;
    Button4: TButton;
    Button5: TButton;
    Button6: TButton;
    OpenDialog: TOpenDialog;
    Button2: TButton;
    Splitter1: TSplitter;
    Splitter2: TSplitter;
    sDBGrid1: TsDBGrid;
    sEdit1: TsEdit;
    sEdit2: TsEdit;
    sDBGrid2: TsDBGrid;
    procedure SpeedButton2Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure sEdit1Change(Sender: TObject);
    procedure sEdit2Change(Sender: TObject);
    procedure sDBGrid1DragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure sDBGrid1KeyPress(Sender: TObject; var Key: Char);
    procedure sDBGrid2KeyPress(Sender: TObject; var Key: Char);
    procedure btnDelStrokaUpdate(Sender: TObject);
    procedure btnClearListUpdate(Sender: TObject);
    procedure btnExportExcelUpdate(Sender: TObject);
    procedure btnLoadTempFileUpdate(Sender: TObject);
    procedure btnSaveTempFileUpdate(Sender: TObject);
  private
    { Private declarations }
    procedure ShapkaStringGrida;
    procedure AutoStringGridWidth; // �������� ����� �������
  public
    { Public declarations }
    procedure AfterCreation; override;
  end;

implementation

{$R *.dfm}

uses uMyProcedure, uMainForm, uDataModul, uMyExcel;
{ TfrmSLG }

procedure TfrmSLG.AfterCreation;
begin
  inherited;
  ShapkaStringGrida;
  StringGrid1.RowCount := 1;
  DM.qUslugy.Active := true;
  DM.qQslg.Active := true;
end;

procedure TfrmSLG.AutoStringGridWidth;
var
  X, Y, w: Integer;
  MaxWidth: Integer;
begin
  with StringGrid1 do
    ClientHeight := DefaultRowHeight * RowCount + 5;
  with StringGrid1 do
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

  // if StringGrid1.RowCount = 3 then StringGrid1.ColWidths[5] := 250;
end;

procedure TfrmSLG.btnClearListUpdate(Sender: TObject);
begin // �������� ������
  inherited;
  if (StringGrid1.RowCount > 1) then
    Button4.Enabled := true
  else
    Button4.Enabled := false;
end;

procedure TfrmSLG.btnDelStrokaUpdate(Sender: TObject);
begin // ������� ������
  inherited;
  if (StringGrid1.RowCount > 1) then
    Button3.Enabled := true
  else
    Button3.Enabled := false;
end;

procedure TfrmSLG.btnExportExcelUpdate(Sender: TObject);
begin // �������������� � ������
  inherited;
  if (StringGrid1.RowCount > 1) then
    Button2.Enabled := true
  else
    Button2.Enabled := false;
end;

procedure TfrmSLG.btnLoadTempFileUpdate(Sender: TObject);
begin
  inherited;
  if (StringGrid1.RowCount > 1) then
    Button6.Enabled := false
  else
    Button6.Enabled := true;
end;

procedure TfrmSLG.btnSaveTempFileUpdate(Sender: TObject);
begin
  inherited;
  if (StringGrid1.RowCount > 1) then
    Button5.Enabled := true
  else
    Button5.Enabled := false;
end;

procedure TfrmSLG.Button2Click(Sender: TObject);
var // ������� � ������
  Sheets, ExcelApp: variant;
  i, j: Integer;
  s, FileNameS: String;
  d: TDate;
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

  MyExcel.Range['A1'].value := '����� ������';
  MyExcel.Range['B1'].value := 'JDC ID';
  MyExcel.Range['C1'].value := '���';
  MyExcel.Range['D1'].value := '�������� ������ �������';
  MyExcel.Range['E1'].value := '����������';
  MyExcel.Range['F1'].value := '���������';
  MyExcel.Range['G1'].value := '����������';

  for i := 1 to StringGrid1.RowCount - 1 do // ������
    for j := 0 to StringGrid1.ColCount - 1 do // �������
    begin
      MyExcel.Cells[i + 1, j + 1] := StringGrid1.Cells[j, i];
    end;

  FileNameS := ExtractFilePath(ParamStr(0)) + '���_' +
    FormatDateTime('dd.mm.yyyy', Now) + '.xlsx';

  uMyExcel.SaveWorkBook(FileNameS, 1);

  Sheets := unassigned;
  ExcelApp := unassigned;
  uMyExcel.StopExcel;

  ShowMessage('������ �������������� � ��������� � ����');
end;

procedure TfrmSLG.Button3Click(Sender: TObject);
var // ������� ���������� �������
  n: Integer;
begin
  inherited;
  // ���� �������� ���� ������, ��������� �������� �� ���������
  if (StringGrid1.RowCount = 1) then
    Exit;
  // ����� ������ �����, ������� �� ������, ���������� ����������
  for n := StringGrid1.Row to StringGrid1.RowCount - 2 do
  begin
    StringGrid1.Rows[n] := StringGrid1.Rows[n + 1];
  end;
  // �������� ��������� ������ �������
  StringGrid1.RowCount := StringGrid1.RowCount - 1;

end;

procedure TfrmSLG.Button4Click(Sender: TObject);
var
  i, j: Integer;
begin
  inherited;
  with StringGrid1 do
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

  StringGrid1.RowCount := 1;
end;

procedure TfrmSLG.Button5Click(Sender: TObject);
begin
  inherited;
  StringGrid1.SaveToFile(ExtractFilePath(Application.ExeName) +
    'slg_chernovik.txt');
end;

procedure TfrmSLG.Button6Click(Sender: TObject);
begin
  inherited;
  if FileExists(ExtractFilePath(Application.ExeName) + 'slg_chernovik.txt') then
    StringGrid1.LoadFromFile(ExtractFilePath(Application.ExeName) +
      'slg_chernovik.txt');
end;

procedure TfrmSLG.sDBGrid1DragDrop(Sender, Source: TObject; X, Y: Integer);
begin
  inherited;
  // �������� �� ���������� ��������������

end;

procedure TfrmSLG.sDBGrid1KeyPress(Sender: TObject; var Key: Char);
var
  last_row: Integer;
begin
  inherited;
  if (Key = #13) then // ���� ����� Enter
  begin
    last_row := StringGrid1.RowCount;
    StringGrid1.RowCount := StringGrid1.RowCount + 1; // ���������� ������
    StringGrid1.Cells[0, last_row] := sDBGrid1.DataSource.DataSet.FieldByName
      ('RITM').AsString;
    StringGrid1.Cells[1, last_row] := sDBGrid1.DataSource.DataSet.FieldByName
      ('JDCID').AsString;
    StringGrid1.Cells[2, last_row] := sDBGrid1.DataSource.DataSet.FieldByName
      ('FIO').AsString;
    // StringGrid1.Cells[4, last_row] := DBGrid1.DataSource.DataSet.FieldByName('Number').AsString;
    AutoStringGridWidth;
  end;
end;

procedure TfrmSLG.sDBGrid2KeyPress(Sender: TObject; var Key: Char);
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
    select_row := StringGrid1.Tag;
    if select_row = 0 then
      select_row := select_row + 1;
    if select_row <> 0 then
    begin

      sEdit1.text := '';
      izdelie := sDBGrid2.DataSource.DataSet.FieldByName('Name_SLG').AsString;
      upakovka := sDBGrid2.DataSource.DataSet.FieldByName('upakovka').AsInteger;
      Price_inch := sDBGrid2.DataSource.DataSet.FieldByName('Price_inch')
        .AsCurrency;

    ups:
      if // (
        (StringGrid1.Cells[3, select_row] <> '')
      // or (select_row >= StringGrid1.RowCount - 1 ))
      then
      begin
        buttonSelected := MessageDlg('������������� ������?', mtCustom,
          [mbYes, mbNo], 0);

        if buttonSelected = mrYes then
        begin
          StringGrid1.Row := select_row;
          StringGrid1.Cells[3, select_row] := izdelie;
          // Price_inch       upakovka
          colvo := StrToInt(InputBox(izdelie, '������� ���-��', ''));
          StringGrid1.Cells[5, select_row] :=
            ((colvo / upakovka) * Price_inch).ToString;
          StringGrid1.Cells[4, select_row] := colvo.ToString;
          AutoStringGridWidth;
          if select_row <> StringGrid1.RowCount - 1 then
            StringGrid1.Row := StringGrid1.Row + 1;
        end
        else
        begin
          select_row := select_row + 1;
          goto ups;
        end;

      end
      else
      begin
        StringGrid1.Row := select_row;
        StringGrid1.Cells[3, select_row] := izdelie;
        // Price_inch       upakovka
        colvo := StrToInt(InputBox(izdelie, '������� ���-��', ''));
        StringGrid1.Cells[5, select_row] :=
          ((colvo / upakovka) * Price_inch).ToString;
        StringGrid1.Cells[4, select_row] := colvo.ToString;
        AutoStringGridWidth;
        if select_row <> StringGrid1.RowCount - 1 then
          StringGrid1.Row := StringGrid1.Row + 1;
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
  a2 := QuotedStr(a1);
  // ������� ��������� ������ ��������� ������� ( ' ) � ������ � ����� ������
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
  StringGrid1.Cells[0, 0] := '����� ������';
  StringGrid1.Cells[1, 0] := 'JDC ID';
  StringGrid1.Cells[2, 0] := '���';
  StringGrid1.Cells[3, 0] := '�������� ������ �������';
  StringGrid1.ColWidths[3] := 250;
  StringGrid1.Cells[4, 0] := '����������';
  StringGrid1.Cells[5, 0] := '���������';
  StringGrid1.Cells[6, 0] := '����������';
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

        // ------------ ���������� ��������� ������� �������� �������� -------------

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

          tSLG.FieldByName('Name_SLG').AsString := MyExcel.Cells[m, 1].value;

          if MyExcel.Cells[m, 3].value = '�����' then
            tSLG.FieldByName('Active').AsBoolean := true
          else
            tSLG.FieldByName('Active').AsBoolean := false;

          tSLG.FieldByName('Price_inch').AsCurrency :=
            StrToCurr(MyExcel.Cells[m, 4].value);

          tSLG.FieldByName('upakovka').AsInteger :=
            ParseStringUpakovka(MyExcel.Cells[m, 1].value).ToInteger;

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
            [m, StrToInt(CollectionNameTable.Items['���'].ToString)].value;

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

  end;

end;

end.
