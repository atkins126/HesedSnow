// ���� ������ ������������ Delphi IDE, ���� �������� *dfm object �� inherited
unit uMainForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, uFrameCustom,

  JvAppStorage, JvAppIniStorage,
  JvComponentBase, JvFormPlacement, acTitleBar, sSkinManager, sSkinProvider,
  Vcl.Controls, Vcl.StdCtrls, Vcl.ComCtrls, sComboBoxes, sLabel, JvExExtCtrls,
  JvExtComponent, JvClock, dxGDIPlusClasses, ES.BaseControls, ES.Images,
  Vcl.ExtCtrls, sPanel, sStatusBar, Vcl.Forms, sScrollBox, sFrameBar, sSplitter,
  sMonthCalendar;

type
  TFrameClass = class of TCustomInfoFrame;

  TmyForm = class(TForm)
    sFrameBar1: TsFrameBar;
    sStatusBar1: TsStatusBar;
    sSplitter1: TsSplitter;
    panConteiner: TsPanel;
    sPanel2: TsPanel;
    sGradientPanel2: TsGradientPanel;
    EsImage1: TEsImage;
    JvClock1: TJvClock;
    sPanel1: TsPanel;
    sLabelFX1: TsLabelFX;
    sSkinSelector1: TsSkinSelector;
    sSkinManager1: TsSkinManager;
    sTitleBar1: TsTitleBar;
    JvFormStorage1: TJvFormStorage;
    JvAppIniFileStorage1: TJvAppIniFileStorage;
    sMonthCalendar1: TsMonthCalendar;
    procedure sFrameBar1Items0CreateFrame(Sender: TObject;
      var Frame: TCustomFrame);
    procedure FormShow(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
    myFrame: TFrame;
    // --------------------
    procedure CreateNewFrame(FrameType: TFrameClass; Sender: TObject = nil);
    procedure UpdateFrame(Sender: TObject = nil);
    procedure UpdateFrameControls;
    // ----------------------
  end;

var
  myForm: TmyForm;
  OldFrame, CurrentFrame: TCustomInfoFrame;
  AppLoading: boolean = False; // ��������� �������� ����� �� ����� ��������
  // ����������
  FormShowed: boolean = False; // ��� ���������� ������������ ��� �������������
  // ������ ����� � ������� OnShow. ������������ ��� �������������� ���������
  // ������������� ����� ������� ����������� �����. ������� Form.OnShow
  // �������������� ����� ������� ������������ � ����� �� ������� ��� ��� ������.

implementation

{$R *.dfm}

uses uMenu, uDataModul;

{ TForm2 }

procedure TmyForm.CreateNewFrame(FrameType: TFrameClass; Sender: TObject);
begin
  if Assigned(CurrentFrame) then
    OldFrame := CurrentFrame;

  if OldFrame <> nil then
  begin // ��������� ���� ����������
    if OldFrame is FrameType then
      FreeAndNil(OldFrame);
  end;
  CurrentFrame := FrameType.Create(myForm);
  myForm.UpdateFrame(Sender);
end;

procedure TmyForm.FormShow(Sender: TObject);
begin
  if not FormShowed then
  begin
    AppLoading := True;
    FormShowed := True; // �������������� ��������� �������������
     // ��������� ������ ������� ������ ������ (TfrmMenu)
    sFrameBar1.OpenItem(0, False { ��� �������� } );
    // ������ ������� � ������ (������� �� spdBtn_CurrSkin)
//    TfrmMenu(sFrameBar1.Items[0].Frame).btnVidomist.OnClick(TfrmMenu(sFrameBar1.Items[0].Frame).btnVidomist);
//    GenerateSkinsList;  // ����� ��������� ������
    AppLoading := False;
  end;
end;

procedure TmyForm.sFrameBar1Items0CreateFrame(Sender: TObject;
  var Frame: TCustomFrame);
begin
  Frame := TfrmMenu.Create(nil);
  sSkinManager1.UpdateScale(Frame);
end;

procedure TmyForm.UpdateFrame(Sender: TObject);
begin
  if Assigned(CurrentFrame) then
  begin
    CurrentFrame.Visible := False;
    // ������������� ��������� ������ ������
    CurrentFrame.Align := alClient;
    CurrentFrame.Parent := panConteiner;
    UpdateFrameControls;
    // ���� Animated � sSkinManager1.Active, � �� AppLoading, ����� ��������
    CurrentFrame.SendToBack;
    CurrentFrame.Visible := True;
    if Assigned(OldFrame) then
      OldFrame.Visible := False;
  end
  else
  begin
    CurrentFrame.Visible := True;
{$IFNDEF DELPHI_XE}
    CurrentFrame.Repaint;
    // ����������� ����������� �������� ����������, �������� ���� ��� ������ �������� ���������� Delphi
{$ENDIF}
  end;
  if Assigned(OldFrame) then
    FreeAndNil(OldFrame);
end;

procedure TmyForm.UpdateFrameControls;
begin
  if CurrentFrame <> nil then
    CurrentFrame.AfterCreation;
end;

end.
