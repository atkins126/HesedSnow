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
  labInfoStatus.Caption := '���������� ��� ��� ���������� �������';

  if FileExists(ExtractFilePath(ParamStr(0)) + 'items.dat') Then
  begin // ��
    cbVidomist.Items.LoadFromFile(ExtractFilePath(ParamStr(0)) + 'items.dat');
  end;
end;

procedure TfrmVidomost.BeforeDestruct;
begin
  inherited;
  // ���� ������� ����������� �������, �������� �� ������
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
  //���� �������� ��������� ����� �� �������� � ������� ��������
  if not uMyProcedure.isStringAssign(cbVidomist.text, cbVidomist.Items) then
  begin
   cbVidomist.Items.Add(cbVidomist.text);
   cbVidomist.Items.SaveToFile(ExtractFilePath(ParamStr(0)) + 'items.dat');
  end;
    labInfoStatus.Shadow.Color := clBlue;
    labInfoStatus.Caption := '��������� ������ ��������';
      myForm.ProgressBar.Visible := True;
      myForm.ProgressBar.Min := 0;
      myForm.ProgressBar.Max := 100;
      myForm.ProgressBar.Position := 1;

//1 ������ ������ ��������� � ������ ���������
Kurators := TStringList.Create;
Kurators := CuratorsList('Vidomist','Curator');
      myForm.ProgressBar.Position := 10;

//2 ������� ����� ��� ������ ���������
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

//3 ��� ������� �������� ������� ���� � �������
 for I := 0 to Kurators.Count - 1 do
 begin
  ExportExcel(Kurators.Strings[i], DirectoryNow, cbVidomist.Text);
  myForm.ProgressBar.Position := myForm.ProgressBar.Position + _step;
  labInfoStatus.Caption := '������� ����� �� ��������: ' + Kurators.Strings[i];
 end;






//4 ���������� ���� �� ����� ��� ������� ��������
//5 ��� �������� ���� �� ������


      labInfoStatus.Shadow.Color := clGreen;
      labInfoStatus.Caption := '��������� ������ ��ϲ���!!!';
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
    // �������� �� ������� � ������ Excel
    OpenDialog.Filter := '����� MS Excel|*.xls;*.xlsx|';
  if not OpenDialog.Execute then
    Exit;

  if uMyExcel.OpenWorkBook(OpenDialog.FileName, False) then
  // ��������� ����� Excel
  begin
    // MyExcel.Workbooks[1].Worksheets.Count; //���-�� ������ � ���������
    myForm.ProgressBar.Visible := True;
    MyExcel.ActiveWorkBook.Sheets[1];
//    ListExcel := MyExcel.ActiveWorkBook.Sheets[1];

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
          MyExcel.Cells[m, StrToInt(CollectionNameTable.Items['�����']
          .ToString)].value;

        tVedomost.FieldByName('JDCID').AsString :=
          MyExcel.Cells[m, StrToInt(CollectionNameTable.Items['JDC ID']
          .ToString)].value;

        tVedomost.FieldByName('FIO').AsString :=
          MyExcel.Cells
          [m, StrToInt(CollectionNameTable.Items['���'].ToString)].value;

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
          MyExcel.Cells[m, StrToInt(CollectionNameTable.Items['�������']
          .ToString)].value;

        tVedomost.FieldByName('Zhertva').AsString :=
          MyExcel.Cells
          [m, StrToInt(CollectionNameTable.Items['��'].ToString)].value;

        tVedomost.FieldByName('Mobila').AsString :=
          MyExcel.Cells
          [m, StrToInt(CollectionNameTable.Items['��������� �������']
          .ToString)].value;

        tVedomost.FieldByName('Adress').AsString :=
          MyExcel.Cells[m, StrToInt(CollectionNameTable.Items['������� �����']
          .ToString)].value;

        tVedomost.FieldByName('S_Kem').AsString :=
          MyExcel.Cells[m, StrToInt(CollectionNameTable.Items['� ��� ���������']
          .ToString)].value;

        tVedomost.FieldByName('SABA').AsCurrency :=
          MyExcel.Cells
          [m, StrToInt(CollectionNameTable.Items['��. ����� ��� �� (�����)']
          .ToString)].value;

        tVedomost.FieldByName('Type_Uchasnika').AsString :=
          MyExcel.Cells[m, StrToInt(CollectionNameTable.Items['��� ���������']
          .ToString)].value;

        tVedomost.FieldByName('Data_Plan').AsDateTime :=
          MyExcel.Cells
          [m, StrToInt(CollectionNameTable.Items['����������� ����']
          .ToString)].value;

        tVedomost.Post;
        Inc(m);
        // Application.ProcessMessages;
        Sleep(25);
        ProgressBar.Position := m;
      end;
      labInfoStatus.Shadow.Color := clGreen;
      labInfoStatus.Caption := '������������ ����� ��ϲ���!!!';
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
  labInfoStatus.Caption := '���������� ��� ��� ���������� �������';
  ImportExcel;
  DM.tVedomost.Active := True;
  DBGridEhVedomist.DataSource := DM.dsVedomist;
end;

end.
