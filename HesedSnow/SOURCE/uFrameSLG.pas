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
  private
    { Private declarations }
    procedure ShapkaStringGrida;
    procedure AutoStringGridWidth; // подгонка ширині колонок
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

procedure TfrmSLG.ShapkaStringGrida;
begin
  // Заголовки грида
  StringGrid1.Cells[0, 0] := 'Номер услуги';
  StringGrid1.Cells[1, 0] := 'JDC ID';
  StringGrid1.Cells[2, 0] := 'ФИО';
  StringGrid1.Cells[3, 0] := 'Средство личной гигиены';
  StringGrid1.ColWidths[3] := 250;
  StringGrid1.Cells[4, 0] := 'Количество';
  StringGrid1.Cells[5, 0] := 'Стоимость';
  StringGrid1.Cells[6, 0] := 'Примечание';
end;

procedure TfrmSLG.SpeedButton1Click(Sender: TObject);
begin
  inherited;
// ImportFromExcel(table, spisok poley, spisok typeDannih);
// или переделать чтоб все данніе біли стринговіе

end;

procedure TfrmSLG.SpeedButton2Click(Sender: TObject);
var
  z, col, m, n : Integer;
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
            [m, StrToInt(CollectionNameTable.Items['ФИО'].ToString)].value;

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

  end;

end;

end.
