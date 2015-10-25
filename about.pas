unit about;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, process, FileUtil, Forms, Controls, Graphics, Dialogs,
  ExtCtrls, StdCtrls;

type

  { TForm2 }

  TForm2 = class(TForm)
    Image1: TImage;
    Label1: TLabel;
    Label10: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    procedure Label2Click(Sender: TObject);
    procedure Label5Click(Sender: TObject);
    procedure Label7Click(Sender: TObject);
    procedure Label9Click(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  Form2: TForm2;

implementation
uses LCLIntf;
{$R *.lfm}

{ TForm2 }

procedure TForm2.Label2Click(Sender: TObject);
begin
  OpenURL('http://www.rinorusso.it');
end;

procedure TForm2.Label5Click(Sender: TObject);
begin
OpenURL('http://www.fasttools.it');
end;

procedure TForm2.Label7Click(Sender: TObject);
begin
 OpenURL('http://www.fasttools.it/fast-mailer');
end;

procedure TForm2.Label9Click(Sender: TObject);
begin
 OpenURL('http://task.fasttools.it/index.php?project=2');
end;

end.

