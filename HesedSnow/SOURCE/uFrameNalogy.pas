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
begin // ------------ ������ ������ � �������� �� ������ ----------------
  inherited;
  with myForm, DM do
  begin
    labInfoStatus.Shadow.Color := clBlue;
    labInfoStatus.Caption := '³�������� ���� �������';
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
        labInfoStatus.Caption := ' ������� ������� ���������';
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

        labInfoStatus.Caption := ' ��������� �������';

        CleanOutTableAndIndex('nal_Otchets', 'ID_OT');
        CleanOutTableAndIndex('nal_Podopechnie', 'ID_PO');
        CleanOutTable('nal_Vedomost');

        while m <> n do // ���� ������� �� ������� EXCEL
        begin
          //1 ������ ������ � ������ � ������ �������


          //2 ������ � ����� ������









          // ��������� ��� ������� �� ����� ����
          if MyExcel.Cells
            [m, StrToInt(CollectionNameTable.Items['���������'].ToString)
            ].value <> 0 then
          begin
            // 1. -------- �������� �� ����� ���������� -----------------------
            INN_value := MyExcel.Cells
              [m, StrToInt(CollectionNameTable.Items['���'].ToString)].value;
            if INN_value <> '' then
            begin
              // ���� ��� �� ������ ������ �����������
              if SearchPoziciy('nal_Podopechnie', 'po_INN', INN_value) = false
              then
                InsertRecord('nal_Podopechnie',
                  'po_FIO, po_INN, po_CalcNalogy, po_Ostatok',
                  MyExcel.Cells[m, StrToInt(CollectionNameTable.Items['���']
                  .ToString)].value,
                  MyExcel.Cells[m, StrToInt(CollectionNameTable.Items['���']
                  .ToString)].value, false, 0);

            end
            else
            begin
              // ��� ��� ���� ����-�� ������
            end;
            // ================= END �������� �� ���������� =================



          end;


          ProgressBar.Position := m;
          Inc(m);
        end;
      end;

    myForm.ProgressBar.Visible := false;
    MyExcel.Application.DisplayAlerts := false;
    StopExcel;
    labInfoStatus.Shadow.Color := clGreen;
    labInfoStatus.Caption := '������ ������ ����������!!!';
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
