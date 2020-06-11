program EmbeddedVissim;

uses
  Vcl.Forms,
  MainUnit in 'MainUnit.pas' {MainForm},
  Vissim.ComProvider.Utilities in 'Vissim.ComProvider.Utilities.pas',
  VISSIMLIB_TLB in 'VISSIMLIB_TLB.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
