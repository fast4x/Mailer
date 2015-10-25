program Mailer_x64;

{$mode objfpc}{$H+}

uses
  //{$IFDEF UNIX}{$IFDEF UseCThreads}
  cthreads,
  //{$ENDIF}{$ENDIF}
  Interfaces, // this includes the LCL widgetset
  Forms, main, pl_indycomp, pl_lnetcomp, sdflaz, setting, about
  { you can add units after this };

{$R *.res}

begin
  Application.Title:='fast Mailer x64';
  RequireDerivedFormResource := True;
  Application.Initialize;
  Application.CreateForm(TForm_main, Form_main);
  Application.CreateForm(TForm_setting, Form_setting);
  Application.CreateForm(TForm2, Form2);
  Application.Run;
end.
