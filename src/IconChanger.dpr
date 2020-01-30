program IconChanger;

uses
  Vcl.Forms,
  uMainForm in 'uMainForm.pas' {MainForm},
  uShortcut in 'uShortcut.pas',
  WinShell in 'WinShell.pas',
  Vcl.Themes,
  Vcl.Styles;

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  TStyleManager.TrySetStyle('Glow');
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
