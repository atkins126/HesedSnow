unit uFrameVidomist;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,

  uFrameCustom, Vcl.ComCtrls, sPanel, Vcl.ExtCtrls, sFrameAdapter;

type
  TfrmVidomost = class(TCustomInfoFrame)
    sGradientPanel1: TsGradientPanel;
    sPanel1: TsPanel;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

implementation

{$R *.dfm}

end.
