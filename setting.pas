unit setting;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, SdfData, db, FileUtil, Forms, Controls, Graphics, Dialogs,
  ComCtrls, ExtCtrls, DbCtrls, DBGrids, StdCtrls, DefaultTranslator, EditBtn, types;

type

  { TForm_setting }

  TForm_setting = class(TForm)
    CheckBox1: TCheckBox;
    DataSource1: TDataSource;
    DBGrid1: TDBGrid;
    DBNavigator1: TDBNavigator;
    emailmitt: TLabeledEdit;
    emailreply: TLabeledEdit;
    subject: TLabeledEdit;
    FileNameEdit1: TFileNameEdit;
    FileNameEdit2: TFileNameEdit;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    mailagent: TLabeledEdit;
    heloname: TLabeledEdit;
    nomemitt: TLabeledEdit;
    PageControl1: TPageControl;
    Panel1: TPanel;
    dbservers: TSdfDataSet;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    TabSheet4: TTabSheet;
    procedure dbserversAfterInsert(DataSet: TDataSet);
    procedure dbserversBeforePost(DataSet: TDataSet);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure TabSheet1ContextPopup(Sender: TObject; MousePos: TPoint;
      var Handled: Boolean);
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  Form_setting: TForm_setting;

implementation
uses main;
{$R *.lfm}

{ TForm_setting }

procedure TForm_setting.FormShow(Sender: TObject);
begin
  dbservers.Open;
end;

procedure TForm_setting.TabSheet1ContextPopup(Sender: TObject;
  MousePos: TPoint; var Handled: Boolean);
begin

end;

procedure TForm_setting.FormCreate(Sender: TObject);
begin
    dbservers.FileName:=extractfilepath(application.ExeName)+'smtpservers.conf';
    heloname.Text:= form_main.conf.ReadString('Server','HeloName','');
    mailagent.Text:= form_main.conf.ReadString('Server','MailAgent','');
    emailreply.Text:= form_main.conf.ReadString('Identity','ReplyEmail','');
    emailmitt.Text:= form_main.conf.ReadString('Identity','FromEmail','');
    nomemitt.Text:= form_main.conf.ReadString('Identity','FromName','');
    CheckBox1.Checked:=form_main.conf.ReadBool('Server','FileLog',false);

    subject.Text:=form_main.conf.ReadString('Model','Subject','');
    FileNameEdit1.Text:=form_main.conf.ReadString('Model','HTMLBody','');
    FileNameEdit2.Text:=form_main.conf.ReadString('Model','TEXTBody','');

end;

procedure TForm_setting.FormClose(Sender: TObject; var CloseAction: TCloseAction
  );
begin
  dbservers.Close;
  Form_main.LoadServersSmtp(Sender);
end;

procedure TForm_setting.dbserversAfterInsert(DataSet: TDataSet);
begin
  dbservers.FieldByName('Timeout').asinteger:=0;
end;

procedure TForm_setting.dbserversBeforePost(DataSet: TDataSet);
begin
  if dbservers.FieldByName('Timeout').AsString='' then dbservers.FieldByName('Timeout').AsInteger:=0;
end;

end.

