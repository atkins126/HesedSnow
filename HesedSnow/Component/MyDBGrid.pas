// потомок для реализации перетаскивания из DBGrid
unit MyDBGrid;

interface

uses Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, acDBGrid;

type
  TMyDBGrid = class(TsDBGrid)
  private
    { Private declarations }
    FOnMouseDown: TMouseEvent;

  protected
    { Protected declarations }
    procedure MouseDown(Button: TMouseButton; Shift: TShiftState;
      X, Y: Integer); override;

  published
    { Published declarations }
    property Row;
    property OnMouseDown read FOnMouseDown write FOnMouseDown;

  end;

procedure Register;

implementation

{ TMyDBGrid }

procedure TMyDBGrid.MouseDown(Button: TMouseButton; Shift: TShiftState;
  X, Y: Integer);
begin
  if Assigned(FOnMouseDown) then
    FOnMouseDown(Self, Button, Shift, X, Y);
  inherited MouseDown(Button, Shift, X, Y);
end;

procedure Register;
begin
  RegisterComponents('Samples', [TMyDBGrid]);
end;

end.
