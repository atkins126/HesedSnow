unit uMenu;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons,
  sBitBtn, sFrameAdapter;

type
  TfrmMenu = class(TFrame)
    btnSLG: TsBitBtn;
    btnNalogy: TsBitBtn;
    sFrameAdapter1: TsFrameAdapter;
    sBitBtn1: TsBitBtn;
    procedure btnVidomistClick(Sender: TObject);
    procedure btnSLGClick(Sender: TObject);
    procedure btnNalogyClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

implementation

{$R *.dfm}

uses uFrameVidomist, uMainForm, uFrameSLG, uFrameNalogy;

procedure TfrmMenu.btnNalogyClick(Sender: TObject);
begin
  myForm.CreateNewFrame(TfrmNalogy, Sender);
end;

procedure TfrmMenu.btnSLGClick(Sender: TObject);
begin
  myForm.CreateNewFrame(TfrmSLG, Sender);
end;

procedure TfrmMenu.btnVidomistClick(Sender: TObject);
begin
  myForm.CreateNewFrame(TfrmVidomost, Sender);
end;

end.
