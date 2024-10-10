program wia_demo;

uses
  Vcl.Forms,
  WIA_Demo_Form in 'WIA_Demo_Form.pas' {FormWIADemo};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TFormWIADemo, FormWIADemo);
  Application.Run;
end.
