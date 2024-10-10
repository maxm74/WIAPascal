program wia_tests;

uses
  Vcl.Forms,
  wia_tests_main in 'wia_tests_main.pas' {TFormWIATests};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TFormWIATests, FormWIATests);
  Application.Run;
end.
