program HesedSnow;

uses
  Vcl.Forms,
  uMainForm in '..\SOURCE\uMainForm.pas' {myForm},
  uFrameCustom in '..\SOURCE\uFrameCustom.pas' {CustomInfoFrame: TFrame},
  uMenu in '..\SOURCE\uMenu.pas' {frmMenu: TFrame},
  uFrameVidomist in '..\SOURCE\uFrameVidomist.pas' {frmVidomost: TFrame},
  uFrameSLG in '..\SOURCE\uFrameSLG.pas' {frmSLG: TFrame},
  uFrameNalogy in '..\SOURCE\uFrameNalogy.pas' {frmNalogy: TFrame},
  uMyProcedure in '..\SOURCE\uMyProcedure.pas',
  uMyExcel in '..\SOURCE\uMyExcel.pas',
  uDataModul in '..\SOURCE\uDataModul.pas' {DM: TDataModule},
  uBDlogic in '..\SOURCE\uBDlogic.pas',
  MyDBGrid in '..\Component\MyDBGrid.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TmyForm, myForm);
  Application.CreateForm(TDM, DM);
  Application.Run;
end.
