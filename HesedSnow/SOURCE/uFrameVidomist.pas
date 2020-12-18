unit uFrameVidomist;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes,
  uFrameCustom, System.Generics.Collections,

  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Vcl.ComCtrls, sPanel, Vcl.ExtCtrls, sFrameAdapter, Vcl.StdCtrls,
  sButton, DBGridEhGrouping, ToolCtrlsEh, DBGridEhToolCtrls, DynVarsEh,
  EhLibVCL, GridsEh, DBAxisGridsEh, DBGridEh;

type
  TfrmVidomost = class(TCustomInfoFrame)
    sGradientPanel1: TsGradientPanel;
    sPanel1: TsPanel;
    sButton1: TsButton;
    OpenDialog: TOpenDialog;
    DBGridEhVedomist: TDBGridEh;
    procedure sButton1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure ImportExcel;
  end;

implementation

{$R *.dfm}

uses uMyProcedure, uMyExcel, uDataModul;

procedure TfrmVidomost.ImportExcel;
var
  i, m, n, col, z: Integer;
  ListExcel: Variant;
  CollectionNameTable: TDictionary<string, Integer>;
begin
  //
  if uMyExcel.RunExcel(False, False) = True then
    // �������� �� ������� � ������ Excel
    OpenDialog.Filter := '����� MS Excel|*.xls;*.xlsx|';
  if not OpenDialog.Execute then
    Exit;

  if uMyExcel.OpenWorkBook(OpenDialog.FileName, True) then
  // ��������� ����� Excel
  begin
    // MyExcel.Workbooks[1].Worksheets.Count; //���-�� ������ � ���������

    MyExcel.ActiveWorkBook.Sheets[1];
    ListExcel := MyExcel.ActiveWorkBook.Sheets[1];

    col := MyExcel.ActiveCell.SpecialCells($000000B).Column;
    // ��������� ����������� �������
    // ------------ ���������� ��������� ������� �������� �������� -------------

    CollectionNameTable := TDictionary<string, Integer>.Create();
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
           with DM do
           begin
             while m<>n do    //���� ������� �� ������� EXCEL
             begin
              tVedomost.Insert;

              tVedomost.FieldByName('Num_Uslugy').AsString :=
              MyExcel.Cells[m, StrToInt(CollectionNameTable.Items['�����'].ToString)].value;


               {  JDCID  FIO  Count_usl Cost_usl  Programma  Curator
                  Osobie_Proecty Zhertva Mobila Adress  S_Kem  SABA  Type_Uchasnika
                  Data_Plan
               }
               {
               �����	���	����������	��������� ������	���� ������	���������	�������� ��������� �������	������	������ �� ������������	���������	�������	������ �������	����������� ����	���������	JDC ID	����������� ���������	JDC ID	��	�������� �������	��������� �������	����� ��������	�����	������� �����	����� ����������	� ��� ���������	��. ����� ��� �� (�����)	��� ���������	�������������� ���������	�������� ��������	�������� ������������ ���������	������ ������� �� ��������� �������


               }


               tVedomost.Post;
             end;
           end;






  end;

  MyExcel.Application.DisplayAlerts := False;
  ListExcel := Unassigned;
  StopExcel;
end;

procedure TfrmVidomost.sButton1Click(Sender: TObject);
begin
  inherited;
  ImportExcel;
  DBGridEhVedomist.DataSource := DM.dsVedomist;
  DM.tVedomost.Active := true;
end;

end.
