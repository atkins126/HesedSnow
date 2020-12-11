unit uFrameVidomist;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,

  uFrameCustom, Vcl.ComCtrls, sPanel, Vcl.ExtCtrls, sFrameAdapter, Vcl.StdCtrls,
  sButton;

type
  TfrmVidomost = class(TCustomInfoFrame)
    sGradientPanel1: TsGradientPanel;
    sPanel1: TsPanel;
    sButton1: TsButton;
    OpenDialog: TOpenDialog;
    procedure sButton1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure ImportExcel;
  end;

implementation

{$R *.dfm}

uses uMyProcedure, uMyExcel;

procedure TfrmVidomost.ImportExcel;
var
i : Integer;
 ListExcel: Variant;
begin
//
if uMyExcel.RunExcel(False,False) = True then    //проверка на инсталл и запуск Excel
    if  not OpenDialog.Execute then Exit;

    if uMyExcel.OpenWorkBook(OpenDialog.FileName, False)then //открываем книгу Excel
    begin
     MyExcel.ActiveWorkBook.Sheets[1];

          // тут на первый лист стать и шурувать по листам до последнего листа, проверяя записи на
     for I := 1 to MyExcel.Workbooks[1].Worksheets.Count do
      begin





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
end;

end.
