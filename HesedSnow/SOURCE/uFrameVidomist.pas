unit uFrameVidomist;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes,
  uFrameCustom, System.Generics.Collections,

  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Vcl.ComCtrls, sPanel, Vcl.ExtCtrls, sFrameAdapter, Vcl.StdCtrls,
  sButton, DBGridEhGrouping, ToolCtrlsEh, DBGridEhToolCtrls, DynVarsEh,
  EhLibVCL, GridsEh, DBAxisGridsEh, DBGridEh, sLabel, Vcl.Mask, DBCtrlsEh,
  System.Actions, Vcl.ActnList;

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
    cbVidomist: TDBComboBoxEh;
    actionList: TActionList;
    acbtnCreateVidomist: TAction;
    procedure sButton1Click(Sender: TObject);
    procedure btnCreateVidomistClick(Sender: TObject);
    procedure acbtnCreateVidomistUpdate(Sender: TObject);
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

uses uMyProcedure, uMyExcel, uDataModul, uMainForm, uBDlogic;

procedure TfrmVidomost.acbtnCreateVidomistUpdate(Sender: TObject);
begin
  inherited;
 if cbVidomist.Text.IsEmpty then btnCreateVidomist.Enabled := false else btnCreateVidomist.Enabled := true;
end;

procedure TfrmVidomost.AfterCreation;
begin
  inherited;
  DM.tVedomost.Open; //
  labInfoStatus.Caption := 'Завантажте дані для формування відомості';

  if FileExists(ExtractFilePath(ParamStr(0)) + 'items.dat') Then
  begin // да
    cbVidomist.Items.LoadFromFile(ExtractFilePath(ParamStr(0)) + 'items.dat');
  end;
end;

procedure TfrmVidomost.BeforeDestruct;
begin
  inherited;
  // надо сделать уничтожение фреймов, выгрузку из памяти
  cbVidomist.Items.SaveToFile(ExtractFilePath(ParamStr(0)) + 'items.dat');
end;

procedure TfrmVidomost.btnCreateVidomistClick(Sender: TObject);
var
i : integer;
_step : integer;
Kurators: TStringList;
DirectoryNow : String;
begin
  inherited;
  //если название Ведомости новое то добавить в историю Названий
  if not uMyProcedure.isStringAssign(cbVidomist.text, cbVidomist.Items) then
  begin
   cbVidomist.Items.Add(cbVidomist.text);
   cbVidomist.Items.SaveToFile(ExtractFilePath(ParamStr(0)) + 'items.dat');
  end;
    labInfoStatus.Shadow.Color := clBlue;
    labInfoStatus.Caption := 'Створюємо списки кураторів';
      myForm.ProgressBar.Visible := True;
      myForm.ProgressBar.Min := 0;
      myForm.ProgressBar.Max := 100;
      myForm.ProgressBar.Position := 1;

//1 делаем список кураторов у данной ведомости
Kurators := TStringList.Create;
Kurators := CuratorsList('Vidomist','Curator');
      myForm.ProgressBar.Position := 10;

//2 создаем папку для данной ведомости
DirectoryNow := cbVidomist.Text +' '+ FormatDateTime('mmmm yyyy', Now);
  if DirectoryExists(DirectoryNow) then
   begin
     DirectoryNow := DirectoryNow + '\';
   end
  else
   begin
     CreateDir(ExtractFilePath(ParamStr(0)) + DirectoryNow + '\');
     DirectoryNow := DirectoryNow + '\';
   end;
   DirectoryNow := ExtractFilePath(ParamStr(0)) + DirectoryNow;

   myForm.ProgressBar.Position := 20;
   _step := Trunc((myForm.ProgressBar.Max - myForm.ProgressBar.Position) / Kurators.Count);

//3 для каждого куратора создаем файл с данными
 for I := 0 to Kurators.Count - 1 do
 begin
  ExportExcel(Kurators.Strings[i], DirectoryNow, cbVidomist.Text);
  myForm.ProgressBar.Position := myForm.ProgressBar.Position + _step;
  labInfoStatus.Caption := 'Обробка даних по куратору: ' + Kurators.Strings[i];
 end;






//4 отправляем файл на почту для каждого куратора
//5 или печатаем файл на бумаге


      labInfoStatus.Shadow.Color := clGreen;
      labInfoStatus.Caption := 'Створення відомісті УСПІШНО!!!';
myForm.ProgressBar.Visible := false;
Kurators.Free;
end;

procedure TfrmVidomost.ImportExcel;
var
  i, m, n, col, z: integer;
//  ListExcel: Variant;
  CollectionNameTable: TDictionary<string, integer>;
begin

  if uMyExcel.RunExcel(False, False) = True then
    // проверка на инсталл и запуск Excel
    OpenDialog.Filter := 'Файлы MS Excel|*.xls;*.xlsx|';
  if not OpenDialog.Execute then
    Exit;

  if uMyExcel.OpenWorkBook(OpenDialog.FileName, False) then
  // открываем книгу Excel
  begin
    // MyExcel.Workbooks[1].Worksheets.Count; //кол-во листов в документе
    myForm.ProgressBar.Visible := True;
    MyExcel.ActiveWorkBook.Sheets[1];
//    ListExcel := MyExcel.ActiveWorkBook.Sheets[1];

    col := MyExcel.ActiveCell.SpecialCells($000000B).Column;
    // последняя заполненная колонка
    // ------------ пробежимся расставим индексы названий столбцов -------------

    CollectionNameTable := TDictionary<string, integer>.Create();
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

      ProgressBar.Min := 0;
      ProgressBar.Max := n;
      ProgressBar.Position := 1;

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
        ProgressBar.Position := m;
      end;
      labInfoStatus.Shadow.Color := clGreen;
      labInfoStatus.Caption := 'Завантаження даних УСПІШНО!!!';
    end;
  end;

  MyExcel.Application.DisplayAlerts := False;
//  ListExcel := Unassigned;
  StopExcel;
  CollectionNameTable.Clear;
  CollectionNameTable.Free;
  DM.tVedomost.Active := False;
  myForm.ProgressBar.Visible := False;
end;

procedure TfrmVidomost.sButton1Click(Sender: TObject);
begin
  inherited;
  labInfoStatus.Caption := 'Завантажте дані для формування відомості';
  ImportExcel;
  DM.tVedomost.Active := True;
  DBGridEhVedomist.DataSource := DM.dsVedomist;
end;

end.
