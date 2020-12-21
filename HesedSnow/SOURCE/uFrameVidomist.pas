unit uFrameVidomist;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes,
  uFrameCustom, System.Generics.Collections,

  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Vcl.ComCtrls, sPanel, Vcl.ExtCtrls, sFrameAdapter, Vcl.StdCtrls,
  sButton, DBGridEhGrouping, ToolCtrlsEh, DBGridEhToolCtrls, DynVarsEh,
  EhLibVCL, GridsEh, DBAxisGridsEh, DBGridEh, sLabel, Vcl.Mask, DBCtrlsEh;

type
  TfrmVidomost = class(TCustomInfoFrame)
    sGradientPanel1: TsGradientPanel;
    sPanel1: TsPanel;
    sButton1: TsButton;
    OpenDialog: TOpenDialog;
    DBGridEhVedomist: TDBGridEh;
    sLabelFX1: TsLabelFX;
    btnCreateVidomist: TsButton;
    sPanel2: TsPanel;
    labInfoStatus: TsLabelFX;
    DBComboBoxEh1: TDBComboBoxEh;
    procedure sButton1Click(Sender: TObject);
    procedure btnCreateVidomistClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure AfterCreation; override;
    procedure BeforeDestruct; override;
    procedure ImportExcel;
  end;

implementation

{$R *.dfm}

uses uMyProcedure, uMyExcel, uDataModul, uMainForm;

procedure TfrmVidomost.AfterCreation;
begin
  inherited;
  DM.tVedomost.Open; //
  labInfoStatus.Caption := 'Завантажте дані для формування відомості';
end;

procedure TfrmVidomost.BeforeDestruct;
begin
  inherited;
  // надо сделать уничтожение фреймов, выгрузку из памяти
end;

procedure TfrmVidomost.btnCreateVidomistClick(Sender: TObject);
begin
  inherited;
//
end;

procedure TfrmVidomost.ImportExcel;
var
  i, m, n, col, z: Integer;
  ListExcel: Variant;
  CollectionNameTable: TDictionary<string, Integer>;
begin
  //
  if uMyExcel.RunExcel(False, False) = True then
    // проверка на инсталл и запуск Excel
    OpenDialog.Filter := 'Файлы MS Excel|*.xls;*.xlsx|';
  if not OpenDialog.Execute then
    Exit;

  if uMyExcel.OpenWorkBook(OpenDialog.FileName, False) then
  // открываем книгу Excel
  begin
    // MyExcel.Workbooks[1].Worksheets.Count; //кол-во листов в документе
    myForm.ProgressBar.Visible := true;
    MyExcel.ActiveWorkBook.Sheets[1];
    ListExcel := MyExcel.ActiveWorkBook.Sheets[1];

    col := MyExcel.ActiveCell.SpecialCells($000000B).Column;
    // последняя заполненная колонка
    // ------------ пробежимся расставим индексы названий столбцов -------------

    CollectionNameTable := TDictionary<string, Integer>.Create();
    for z := 1 to col do
    begin
      if not CollectionNameTable.ContainsKey(MyExcel.Cells[1, z].value) then
        CollectionNameTable.Add(MyExcel.Cells[1, z].value, z)
      else
        CollectionNameTable.Add(MyExcel.Cells[1, z].value + z.ToString, z);
    end;

    // ------------------------ конец пробега для  СТОЛБЦОВ --------------------

    m := 2; // начинаем считывание со 2-й строки, оставляя заголовок колонки
    n := MyExcel.ActiveCell.SpecialCells($000000B).Row;
    // последняя заполненная строка
    n := n + 1;
    with DM, myForm do
    begin
      CleanOutTable('Vidomist'); // обнуляем таблицу

      progressBar.Min := 0;
      progressBar.Max := n;
      progressBar.Position := 1;


      while m <> n do // цикл внешний по записям EXCEL
      begin

        tVedomost.Insert;

        tVedomost.FieldByName('Num_Uslugy').AsString :=
          MyExcel.Cells[m, StrToInt(CollectionNameTable.Items['Номер']
          .ToString)].value;

        tVedomost.FieldByName('JDCID').AsString :=
          MyExcel.Cells[m, StrToInt(CollectionNameTable.Items['JDC ID']
          .ToString)].value;

        tVedomost.FieldByName('FIO').AsString :=
          MyExcel.Cells
          [m, StrToInt(CollectionNameTable.Items['ФИО'].ToString)].value;

        tVedomost.FieldByName('Count_usl').AsCurrency :=
          MyExcel.Cells[m, StrToInt(CollectionNameTable.Items['Количество']
          .ToString)].value;

        tVedomost.FieldByName('Cost_usl').AsCurrency :=
          MyExcel.Cells
          [m, StrToInt(CollectionNameTable.Items['Стоимость услуги']
          .ToString)].value;

        tVedomost.FieldByName('Programma').AsString :=
          MyExcel.Cells[m, StrToInt(CollectionNameTable.Items['Программа']
          .ToString)].value;

        tVedomost.FieldByName('Curator').AsString :=
          MyExcel.Cells[m, StrToInt(CollectionNameTable.Items['Куратор']
          .ToString)].value;

        tVedomost.FieldByName('Zhertva').AsString :=
          MyExcel.Cells
          [m, StrToInt(CollectionNameTable.Items['ЖН'].ToString)].value;

        tVedomost.FieldByName('Mobila').AsString :=
          MyExcel.Cells
          [m, StrToInt(CollectionNameTable.Items['Мобильный телефон']
          .ToString)].value;

        tVedomost.FieldByName('Adress').AsString :=
          MyExcel.Cells[m, StrToInt(CollectionNameTable.Items['Главный адрес']
          .ToString)].value;

        tVedomost.FieldByName('S_Kem').AsString :=
          MyExcel.Cells[m, StrToInt(CollectionNameTable.Items['С кем проживает']
          .ToString)].value;

        tVedomost.FieldByName('SABA').AsCurrency :=
          MyExcel.Cells
          [m, StrToInt(CollectionNameTable.Items['Ср. доход для МП (Хесед)']
          .ToString)].value;

        tVedomost.FieldByName('Type_Uchasnika').AsString :=
          MyExcel.Cells[m, StrToInt(CollectionNameTable.Items['Тип участника']
          .ToString)].value;

        tVedomost.FieldByName('Data_Plan').AsDateTime :=
          MyExcel.Cells
          [m, StrToInt(CollectionNameTable.Items['Планируемая дата']
          .ToString)].value;

        tVedomost.Post;
        Inc(m);
        // Application.ProcessMessages;
        Sleep(25);
        progressBar.Position := m;
      end;
      // progressBar.Position := 50;
      labInfoStatus.Shadow.Color := clGreen;
      labInfoStatus.Caption := 'Завантаження даних УСПІШНО!!!';
    end;
  end;

  MyExcel.Application.DisplayAlerts := False;
  ListExcel := Unassigned;
  StopExcel;
  CollectionNameTable.Clear;
  CollectionNameTable.Free;
  DM.tVedomost.Active := False;
  myForm.ProgressBar.Visible := false;
end;

procedure TfrmVidomost.sButton1Click(Sender: TObject);
begin
  inherited;
  ImportExcel;
  DM.tVedomost.Active := True;
  DBGridEhVedomist.DataSource := DM.dsVedomist;
  btnCreateVidomist.Enabled := true;
end;

end.
