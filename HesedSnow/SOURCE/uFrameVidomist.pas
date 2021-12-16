unit uFrameVidomist;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, ShellAPI,
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

  end;

implementation

{$R *.dfm}

uses uMyProcedure, uMyExcel, uDataModul, uMainForm, uBDlogic;

procedure TfrmVidomost.acbtnCreateVidomistUpdate(Sender: TObject);
begin
  inherited;
  if cbVidomist.Text.IsEmpty then
    btnCreateVidomist.Enabled := false
  else
    btnCreateVidomist.Enabled := true;
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
  i: integer;
  _step: integer;
  Kurators: TStringList;
  DirectoryNow: String;
begin
  inherited;
  // ���� �������� ��������� ����� �� �������� � ������� ��������
  if not uMyProcedure.isStringAssign(cbVidomist.Text, cbVidomist.Items) then
  begin
    cbVidomist.Items.Add(cbVidomist.Text);
    cbVidomist.Items.SaveToFile(ExtractFilePath(ParamStr(0)) + 'items.dat');
  end;
  labInfoStatus.Shadow.Color := clBlue;
  labInfoStatus.Caption := '��������� ������ ��������';
  myForm.ProgressBar.Visible := true;
  myForm.ProgressBar.Min := 0;
  myForm.ProgressBar.Max := 100;
  myForm.ProgressBar.Position := 1;

  // 1 ������ ������ ��������� � ������ ���������
  Kurators := TStringList.Create;
  Kurators := CuratorsList('Vidomist', 'Curator');
  myForm.ProgressBar.Position := 10;

  DirectoryNow := ExtractFilePath(ParamStr(0))+ '������\���������\';
  if not DirectoryExists('DirectoryNow') then ForceDirectories(DirectoryNow);

  // 2 ������� ����� ��� ������ ���������
  DirectoryNow := DirectoryNow + cbVidomist.Text + ' ' + FormatDateTime('mmmm yyyy', Now);

  if DirectoryExists(DirectoryNow) then
  begin
    DirectoryNow := DirectoryNow + '\';
  end
  else
  begin
    ForceDirectories(DirectoryNow);
    DirectoryNow := DirectoryNow + '\';
  end;

  myForm.ProgressBar.Position := 20;
  _step := Trunc((myForm.ProgressBar.Max - myForm.ProgressBar.Position) /
    Kurators.Count);

  // 3 ��� ������� �������� ������� ���� � �������
  for i := 0 to Kurators.Count - 1 do
  begin
    ExportExcel(Kurators.Strings[i], DirectoryNow, cbVidomist.Text);
    myForm.ProgressBar.Position := myForm.ProgressBar.Position + _step;
    labInfoStatus.Caption := '������� ����� �� ��������: ' +
      Kurators.Strings[i];
  end;

  // 4 ���������� ���� �� ����� ��� ������� ��������
  // 5 ��� �������� ���� �� ������

  labInfoStatus.Shadow.Color := clGreen;
  labInfoStatus.Caption := '��������� ������ ��ϲ���!!!';
  myForm.ProgressBar.Visible := false;
  Kurators.Free;
  // ��������� ����� � �����������
  ShellExecute(Handle, 'open', PWideChar(DirectoryNow), nil, nil,
    SW_SHOWNORMAL);
end;

procedure TfrmVidomost.sButton1Click(Sender: TObject);
begin
  inherited;
  labInfoStatus.Caption := '���������� ��� ��� ���������� �������';
  // ImportExcel;
  ImportExcelToBD;
  DM.tVedomost.Active := true;
  DBGridEhVedomist.DataSource := DM.dsVedomist;
  labInfoStatus.Shadow.Color := clGreen;
  labInfoStatus.Caption := '������������ ����� ��ϲ���!!!';
  ShowMessage('������������ ����� ��ϲ���!!!');
end;

end.
