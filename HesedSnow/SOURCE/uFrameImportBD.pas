unit uFrameImportBD;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, uFrameCustom,
  sFrameAdapter, acArcControls, Vcl.ExtCtrls, Vcl.StdCtrls, Vcl.Buttons,
  sBitBtn, sPanel,
  System.ImageList, Vcl.ImgList, Data.DB, Vcl.Grids, Vcl.DBGrids, acDBGrid;

type
  TfrmImportBD = class(TCustomInfoFrame)
    btnConnect: TsBitBtn;
    ImageList1: TImageList;
    sDBGrid1: TsDBGrid;
    sPanel1: TsPanel;
    sPanel2: TsPanel;
    btnDeleteTable: TsBitBtn;
    sBitBtn1: TsBitBtn;
    OpenDialog: TOpenDialog;
    sBitBtn2: TsBitBtn;
    procedure btnConnectClick(Sender: TObject);
    procedure btnDeleteTableClick(Sender: TObject);
    procedure sBitBtn1Click(Sender: TObject);
    procedure sBitBtn2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

implementation

{$R *.dfm}

uses uBDlogic, uDataModul;

procedure TfrmImportBD.btnConnectClick(Sender: TObject);
begin
  if ConnectBD('MySQL', 'hesed_test', '+0aX3z%3oE',
    'hesed.mysql.ukraine.com.ua', 'hesed_test', '3306') = true then
  begin
    btnConnect.Glyph := nil;
    ImageList1.GetBitmap(0, btnConnect.Glyph);
    DM.UniQuery1.Active := true;
  end
  else
  begin
    btnConnect.Glyph := nil;
    ImageList1.GetBitmap(1, btnConnect.Glyph);
  end;
  // Winapi.Windows.SetFocus(0);
end;

procedure TfrmImportBD.btnDeleteTableClick(Sender: TObject);
begin

  with DM do
  begin
    UniQuery1.Sql.Text := 'TRUNCATE Uchastniky';
    UniQuery1.ExecSQL;
  end;
end;

procedure TfrmImportBD.sBitBtn1Click(Sender: TObject);
begin
    OpenDialog.Filter := 'Файлы MS Excel|*.xls;*.xlsx|';
    if not OpenDialog.Execute then Exit;
    if ImportExcelToUniTable(OpenDialog.FileName,'Uchastniky')= true then
    begin
     sBitBtn1.Caption := 'ok';

    end
    else
    begin
     sBitBtn1.Caption := 'no'

    end;






end;

procedure TfrmImportBD.sBitBtn2Click(Sender: TObject);
var
s: Widestring;
begin
 {
   with DM do
  begin
  try
  UniQuery1.Sql.Text := '';

    
    UniQuery1.Sql.Text := s;
//    UniQuery1.Insert;
    UniQuery1.ExecSQL;
  except
    showmessage('Не удалось загрузить данные!');
  end;
  end; }

end;



end.
