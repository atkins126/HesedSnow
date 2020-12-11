program HesedSnow;

uses
  Vcl.Forms,
  uMainForm in '..\SOURCE\uMainForm.pas' {myForm},
  uFrameCustom in '..\SOURCE\uFrameCustom.pas' {CustomInfoFrame: TFrame},
  uMenu in '..\SOURCE\uMenu.pas' {frmMenu: TFrame},
  uFrameVidomist in '..\SOURCE\uFrameVidomist.pas' {frmVidomost: TFrame},
  uFrameSLG in '..\SOURCE\uFrameSLG.pas' {frmSLG: TFrame},
  uFrameNalogy in '..\SOURCE\uFrameNalogy.pas' {frmNalogy: TFrame};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TmyForm, myForm);
  Application.Run;
end.
