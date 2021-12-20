unit uFrameNalogy;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, uFrameCustom,
  sFrameAdapter,
  Vcl.StdCtrls, sButton, Vcl.ExtCtrls, sPanel, System.Generics.Collections,
  sLabel;

type
  TfrmNalogy = class(TCustomInfoFrame)
    sPanel2: TsPanel;
    sPanel3: TsPanel;
    btnImportNalog: TsButton;
    OpenDialog: TOpenDialog;
    labInfoStatus: TsLabelFX;
    sRollOutPanel1: TsRollOutPanel;
    procedure btnImportNalogClick(Sender: TObject);
  private
    { Private declarations }

  public
    { Public declarations }
  end;

implementation

{$R *.dfm}

uses uDataModul, uMainForm, uMyExcel, uMyProcedure, uBDlogic;

procedure TfrmNalogy.btnImportNalogClick(Sender: TObject);
var
  z, col, m, n: Integer;
  CollectionNameTable: TDictionary<string, Integer>;
  INN_value: String;
begin // ------------ Импорт отчета с налогами за период ----------------
  inherited;
  with myForm, DM do
  begin
    labInfoStatus.Shadow.Color := clBlue;
    labInfoStatus.Caption := 'Відкриваємо файл імпорту';
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
        labInfoStatus.Caption := ' Формуємо індекси стовпчиків';
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

        labInfoStatus.Caption := ' Обнуляємо таблиці';

        CleanOutTableAndIndex('nal_Otchets', 'ID_OT');
        CleanOutTableAndIndex('nal_Podopechnie', 'ID_PO');
        CleanOutTable('nal_Vedomost');

        while m <> n do // цикл внешний по записям EXCEL
        begin
          //1 внести данные о отчете в список отчетов


          //2 работа с самим файлом









          // проверить или ВЫПЛАТА не равна нулю
          if MyExcel.Cells
            [m, StrToInt(CollectionNameTable.Items['Стоимость'].ToString)
            ].value <> 0 then
          begin
            // 1. -------- проверка на новых подопечных -----------------------
            INN_value := MyExcel.Cells
              [m, StrToInt(CollectionNameTable.Items['ИНН'].ToString)].value;
            if INN_value <> '' then
            begin
              // если нет то внести нового подопечного
              if SearchPoziciy('nal_Podopechnie', 'po_INN', INN_value) = false
              then
                InsertRecord('nal_Podopechnie',
                  'po_FIO, po_INN, po_CalcNalogy, po_Ostatok',
                  MyExcel.Cells[m, StrToInt(CollectionNameTable.Items['ФИО']
                  .ToString)].value,
                  MyExcel.Cells[m, StrToInt(CollectionNameTable.Items['ИНН']
                  .ToString)].value, false, 0);

            end
            else
            begin
              // нет ИНН надо чего-то делать
            end;
            // ================= END проверки на подопечных =================



          end;


          ProgressBar.Position := m;
          Inc(m);
        end;
      end;

    myForm.ProgressBar.Visible := false;
    MyExcel.Application.DisplayAlerts := false;
    StopExcel;
    labInfoStatus.Shadow.Color := clGreen;
    labInfoStatus.Caption := 'Імпорт іспішно завершений!!!';
    // CollectionNameTable.Clear;
    // CollectionNameTable.Free;
    // DM.tUslugy.Active := false;
    // DM.qUslugy.Active := false;
    // DM.qUslugy.Active := true;
            { labInfoStatus.Shadow.Color := clred
              else
              labInfoStatus.Shadow.Color := clBlack; }
  end;
end;



end.
