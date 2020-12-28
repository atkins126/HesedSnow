// чтоб фреймы показывались Delphi IDE, надо поменять *dfm object на inherited
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
  sMonthCalendar, acProgressBar;

type
  TFrameClass = class of TCustomInfoFrame;

  TmyForm = class(TForm)
    sFrameBar1: TsFrameBar;
    statusBar: TsStatusBar;
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
    ProgressBar: TsProgressBar;
    procedure sFrameBar1Items0CreateFrame(Sender: TObject;
      var Frame: TCustomFrame);
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure statusBarDrawPanel(statusBar: TStatusBar; Panel: TStatusPanel;
      const Rect: TRect);

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
  AppLoading: boolean = False; // Запретить анимацию кадра во время загрузки
  // приложения
  FormShowed: boolean = False; // Эта переменная используется при инициализации
  // первой формы в событии OnShow. Используется для предотвращения повторной
  // инициализации после каждого воссоздания формы. Событие Form.OnShow
  // обрабатывается после каждого переключения в режим со скинами или без скинов.

implementation

{$R *.dfm}

uses uMenu, uDataModul, uFrameVidomist;

{ TForm2 }

procedure TmyForm.CreateNewFrame(FrameType: TFrameClass; Sender: TObject);
begin
  if Assigned(CurrentFrame) then
    OldFrame := CurrentFrame;

  if OldFrame <> nil then
  begin // Выгрузить если существует
    if OldFrame is FrameType then
      FreeAndNil(OldFrame);
  end;
  CurrentFrame := FrameType.Create(myForm);
  myForm.UpdateFrame(Sender);
end;

procedure TmyForm.FormCreate(Sender: TObject);
begin
  with ProgressBar do
  begin
    parent := statusBar;
    Position := 1;
    statusBar.Panels[2].Style := psOwnerDraw;
  end;
end;

procedure TmyForm.statusBarDrawPanel(statusBar: TStatusBar; Panel: TStatusPanel;
  const Rect: TRect);
begin
  if Panel = statusBar.Panels[2] then
    with ProgressBar do
    begin
      top := Rect.top;
      left := Rect.left;
      width := Rect.Right - Rect.left - 15;
      height := Rect.Bottom - Rect.top;
    end;
end;

procedure TmyForm.FormShow(Sender: TObject);
begin
  if not FormShowed then
  begin
    AppLoading := true;
    FormShowed := true; // предотвращение повторной инициализации
    // Открываем первый элемент панели фрейма (TfrmMenu)
    sFrameBar1.OpenItem(0, False { Без анимации } );
    // Пример доступа к фрейму (нажмите на spdBtn_CurrSkin)
    // TfrmMenu(sFrameBar1.Items[0].Frame).btnVidomist.OnClick(TfrmMenu(sFrameBar1.Items[0].Frame).btnVidomist);
    // GenerateSkinsList;  // Поиск доступных скинов
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
    CurrentFrame.visible := False;
    // Устанавливаем положение нового фрейма
    CurrentFrame.Align := alClient;
    CurrentFrame.parent := panConteiner;
    UpdateFrameControls;
    // если Animated и sSkinManager1.Active, а не AppLoading, тогда начинаем
    CurrentFrame.SendToBack;
    CurrentFrame.visible := true;
    if Assigned(OldFrame) then
      OldFrame.visible := False;
  end
  else
  begin
    CurrentFrame.visible := true;
{$IFNDEF DELPHI_XE}
    CurrentFrame.Repaint;
    // Перекрасить графические элементы управления, обходной путь для старой проблемы обновления Delphi
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
