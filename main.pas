unit main;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, FileUtil, IdSMTP, IdMessage, IdAntiFreeze, Forms, Controls,
  Graphics, Dialogs, StdCtrls, CheckLst, Grids, ComCtrls, ExtCtrls, Buttons,
  types, LCLType, DefaultTranslator, Menus, IdAttachmentFile, IdComponent,
  IdText, IdLogDebug, IdIOHandlerStack, inifiles, IdIntercept, IdGlobal,
  IdLogFile, IdLogEvent, DOM, XMLWrite, XMLRead;

type

  { TForm_main }

  TForm_main = class(TForm)
    Button1: TButton;
    Button2: TButton;
    emailreply: TLabeledEdit;
    IdAntiFreeze1: TIdAntiFreeze;
    IdLogEvent1: TIdLogEvent;
    IdLogFile1: TIdLogFile;
    MenuItem1: TMenuItem;
    MenuItem2: TMenuItem;
    MenuItem3: TMenuItem;
    MenuItem4: TMenuItem;
    MSG: TIdMessage;
    OpenDialog3: TOpenDialog;
    OpenHtml: TOpenDialog;
    Opentxt: TOpenDialog;
    PopupMenuHtml: TPopupMenu;
    PopupMenuDest: TPopupMenu;
    PopupMenuTxt: TPopupMenu;
    SaveDialog1: TSaveDialog;
    smtp: TIdSMTP;
    Label1: TLabel;
    nomemitt: TLabeledEdit;
    emailmitt: TLabeledEdit;
    loggetto: TLabeledEdit;
    memolog: TMemo;
    memohtml: TMemo;
    memotesto: TMemo;
    OpenDialog2: TOpenDialog;
    PageControl1: TPageControl;
    PageControlMess: TPageControl;
    Panel1: TPanel;
    Panel2: TPanel;
    OpenDialog1: TOpenDialog;
    SpeedButton1: TSpeedButton;
    SpeedButton10: TSpeedButton;
    SpeedButton11: TSpeedButton;
    SpeedButton2: TSpeedButton;
    SpeedButton3: TSpeedButton;
    SpeedButton4: TSpeedButton;
    SpeedButton5: TSpeedButton;
    SpeedButton6: TSpeedButton;
    SpeedButton7: TSpeedButton;
    SpeedButton8: TSpeedButton;
    SpeedButton9: TSpeedButton;
    Splitter1: TSplitter;
    Splitter2: TSplitter;
    StatusBar1: TStatusBar;
    StringGrid1: TStringGrid;
    StringGrid2: TStringGrid;
    StringGrid3: TStringGrid;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    Tabhtml: TTabSheet;
    TabSheet4: TTabSheet;
    TabSheet5: TTabSheet;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure IdLogDebug1Receive(ASender: TIdConnectionIntercept;
      var ABuffer: TIdBytes);
    procedure IdLogEvent1Received(ASender: TComponent; const AText,
      AData: string);
    procedure IdLogEvent1Status(ASender: TComponent; const AText: string);
    procedure MenuItem1Click(Sender: TObject);
    procedure MenuItem2Click(Sender: TObject);
    procedure MenuItem3Click(Sender: TObject);
    procedure MenuItem4Click(Sender: TObject);
    procedure PageControlMessChange(Sender: TObject);
    procedure FillCell(Stringrid1 : TStringGrid; Col, Row : Integer;
BkCol, TextCol : TColor);
    procedure smtpConnected(Sender: TObject);
    procedure smtpDisconnected(Sender: TObject);
    procedure smtpStatus(ASender: TObject; const AStatus: TIdStatus;
      const AStatusText: string);
    procedure smtpWork(ASender: TObject; AWorkMode: TWorkMode; AWorkCount: Int64
      );
    procedure smtpWorkBegin(ASender: TObject; AWorkMode: TWorkMode;
      AWorkCountMax: Int64);
    procedure smtpWorkEnd(ASender: TObject; AWorkMode: TWorkMode);
    procedure SpeedButton10Click(Sender: TObject);
    procedure SpeedButton11Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);

    procedure SpeedButton3Click(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
    procedure SpeedButton5Click(Sender: TObject);
    procedure SpeedButton6Click(Sender: TObject);
    procedure SpeedButton8Click(Sender: TObject);
    procedure SpeedButton9Click(Sender: TObject);
    procedure StatusBar1DrawPanel(StatusBar: TStatusBar; Panel: TStatusPanel;
      const Rect: TRect);
    procedure StringGrid1SelectCell(Sender: TObject; aCol, aRow: Integer;
      var CanSelect: Boolean);
    procedure TabhtmlContextPopup(Sender: TObject; MousePos: TPoint;
      var Handled: Boolean);
   procedure Log(mess:string);
   procedure Log1(mess:string);
   procedure LoadServersSmtp(Sender: TObject);
  private
    { private declarations }

  public
    { public declarations }
    conf:Tinifile;
    progname, appname:String;
  end;

var
  Form_main: TForm_main;

resourcestring
  txt1 = 'E'' necessario caricare almeno un destinatario per inviare almeno una email.';
  txt2 = 'Attenzione';
  txt3 = 'Invio...';
  txt4 = 'Invio in corso per ';
  txt5 = 'Inviato';
  txt6 = 'Invio completato per ';
  txt7 = 'Destinatario ';
  txt8 = ' impossibile da raggiungere ';
  txt9 = 'Allegati ';
  txt10 = 'Allegati ';

  newprogetto = 'Nuovo progetto';
  tabnotempty = 'La tabella contiene già dei destinatari che verranno sostituiti, continuare?';
  attentionn = 'Attenzione';
  prepmessage = 'Preparazione del messaggio';
  serverconn = 'Connessione al server ';
  serverconnect ='Connesso al server...';
  serverdisconnect = 'Disconnesso al server...';
  deletedest = 'Eliminazione destinatario';
  msgdeletedest = 'Eliminare il destinatario selezionato?';
  deleteattach = 'Eliminazione allegato';
  msgdeleteattach = 'Eliminare l''allegato selezionato?';
  messtosend = 'Destinatari:';
  messsent = 'Inviati:';
  msgnewprog = 'Tutti i dati attualmente caricati verranno persi,continuare?';
  msgnewprogdesc = 'Nuovo progetto';
  resend = 'Inviare nuovamente a tutti i destinatari?';


implementation
uses setting,about;
{$R *.lfm}

{ TForm_main }

procedure TForm_main.Button1Click(Sender: TObject);
var
  Reply, BoxStyle, selfile : Integer;
begin
selfile:=1;
  if stringgrid1.RowCount>1 then begin;
     //BoxStyle := MB_ICONQUESTION + MB_YESNO;
       BoxStyle :=  MB_ICONWARNING + MB_YESNO;
       Reply := Application.MessageBox(pchar(tabnotempty), pchar(attentionn), BoxStyle);
       if Reply = IDYES then selfile:=1 else selfile:=0;
  end;
//  if Reply = IDYES then Application.MessageBox('Yes       ', 'Reply',MB_ICONINFORMATION)
  //  else Application.MessageBox('No         ', 'Reply', MB_ICONHAND);
  if selfile = 1 then
  if opendialog1.Execute then begin
     stringgrid1.LoadFromCSVFile(opendialog1.FileName,',',false);

  end;
  statusbar1.Panels[0].Text:=messtosend+' '+inttostr(stringgrid1.RowCount-1);
end;

procedure TForm_main.Button2Click(Sender: TObject);
var prog,b:longint;
    att:word;
    pendent: boolean;

begin

//     showmessage(inttostr(smtp.ReadTimeout));
//     exit;

     if stringgrid1.RowCount < 2 then begin
        Application.MessageBox(pchar(txt1), pchar(txt2),MB_ICONWARNING);
        EXIT;
     end;

     speedbutton2.Enabled:=false;

     pendent:=false;
     if stringgrid1.Columns.Count>4 then
     for b:=1 to stringgrid1.RowCount-1 do begin
          if stringgrid1.Cells[4,b]='' then begin
             pendent:=true;
             break;
          end;
     end;

     if not pendent then
     if stringgrid1.Columns.Count<4 then
     begin
        stringgrid1.Columns.Add;
        stringgrid1.Columns[3].Title.Caption:='Status';
        stringgrid1.Columns.Add;
        stringgrid1.Columns[4].Title.Caption:='StatusNum';
        stringgrid1.Columns[4].Visible:=false;

     end;

     log(messtosend+' '+inttostr(stringgrid1.RowCount-1));
     log(prepmessage);

     MSG.From.Name:=nomemitt.Text;   //'Aggiornamento dati Argo';
     MSG.From.Address:=emailmitt.Text;   //'rinorusso@cgilsicilia.it';

     MSG.MessageParts.Clear;

     with TIdText.Create(MSG.MessageParts, nil) do begin
      ContentType := 'multipart/alternative';
     end;

     if PageControlMess.TabIndex=0 then begin
     //MSG.Body.Text:=memohtml.Text;
     //MSG.ContentType:='text/html';
      with TIdText.Create(MSG.MessageParts, nil) do begin
      Body.Text := memohtml.Text;
      ContentType := 'text/html';
      ParentPart := 0;
      end;
     end
     else begin
     //MSG.Body.Text:=memotesto.Text;
     //MSG.ContentType:='text/plain';
      with TIdText.Create(MSG.MessageParts, nil) do begin
      Body.Text := memotesto.Text;
      ContentType := 'text/plain';
      ParentPart := 0;
      end;
     end;
     MSG.Subject:=loggetto.Text;


     for att:=1 to stringgrid2.RowCount-1 do
     with TIdAttachmentFile.Create(MSG.MessageParts, stringgrid2.Cells[0,att]) do begin
      //ContentType := 'application/x-zip-compressed';
      FileName := stringgrid2.Cells[0,att];
    end;
     (*
     if FileExists(stringgrid2.Cells[0,att]) then
          TIdAttachmentFile.Create(MSG.MessageParts, stringgrid2.Cells[0,att]) ;
          //MSG.MessageParts[0].ContentType:='application/pdf';
     *)

     MSG.ContentType := 'multipart/mixed';

    log('----------------------------------------------------------------------------------');

   for prog:=1 to stringgrid1.RowCount-1 do      // ripete il passaggio per ogni riga
   TRY

    if stringgrid1.Cells[4,prog]='1' then Continue;

    stringgrid1.Cells[3,prog]:=txt3;

    StringGrid1.Row:=prog;
    StringGrid1.Col:=2;
    StringGrid1.SetFocus;

    MSG.Recipients.Clear;
    WITH MSG.Recipients.Add DO BEGIN
      //Name:='<Name of recipient>';
      Address:=stringgrid1.Cells[2,prog];
    END;


    application.ProcessMessages;
    TRY
//      smtp.Host:='x.x.x.x'; // IP Address of smtp server
//      smtp.Port:=25; // Port address of smtp service (usually 25)

     //Timeout
     if strtoint(stringgrid3.Cells[6,stringgrid3.Row]) > 0 then
         smtp.ReadTimeout:=strtoint(stringgrid3.Cells[6,stringgrid3.Row]) else smtp.ReadTimeout:=-1;

      if smtp.Connected = false then begin
         log(serverconn+stringgrid3.Cells[0,stringgrid3.Row]);
         smtp.Host:=stringgrid3.Cells[1,stringgrid3.Row];
         smtp.Port:=strtoint(stringgrid3.Cells[2,stringgrid3.Row]);
         smtp.Username:=stringgrid3.Cells[4,stringgrid3.Row];
         smtp.Password:=stringgrid3.Cells[5,stringgrid3.Row];
         smtp.MailAgent:=form_setting.mailagent.Text;
         smtp.HeloName:=form_setting.heloname.Text;
         idlogevent1.Active:=true;
         smtp.Connect;

      end;

      TRY

       log(txt4+stringgrid1.Cells[2,prog]);
       statusbar1.Panels[2].Text:=txt4+stringgrid1.Cells[2,prog];
        application.ProcessMessages;
        smtp.Send(MSG);
        stringgrid1.Cells[3,prog]:=txt5;
        stringgrid1.Cells[4,prog]:='1';
        log(txt6+stringgrid1.Cells[2,prog]);
        log('----------------------------------------------------------------------------------');
        statusbar1.Panels[2].Text:=txt6+stringgrid1.Cells[2,prog];
        statusbar1.Panels[1].Text:=messsent+' '+inttostr(prog);
        application.ProcessMessages;
        //FillCell(Stringgrid1,3,prog,clYellow,clBlack);
      FINALLY
        //smtp.Disconnect
      END
    FINALLY
      application.ProcessMessages;



    END
  FINALLY
   // MSG.Free
  END;
  smtp.Disconnect;
  caption:=appname;
  if form_setting.CheckBox1.Checked then
     memolog.Lines.SaveToFile(extractfilepath(application.ExeName)+'fastMailer.log');

  speedbutton2.Enabled:=true;
  //statusbar1.Panels[1].Text:='';
  statusbar1.Panels[2].Text:='';

end;

procedure TForm_main.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  conf.WriteString('Server','Default',stringgrid3.Cells[0,stringgrid3.Row]);
  conf.WriteString('Server','HeloName',form_setting.heloname.Text);
  conf.WriteString('Server','MailAgent',form_setting.mailagent.Text);
  conf.WriteBool('Server','FileLog',form_setting.CheckBox1.Checked);



  conf.WriteString('Identity','FromName',form_setting.nomemitt.Text);
  conf.WriteString('Identity','FromEmail',form_setting.emailmitt.Text);
  conf.WriteString('Identity','ReplyEmail',form_setting.emailreply.Text);

  conf.WriteString('Model','Subject',form_setting.subject.Text);
  conf.WriteString('Model','HTMLBody',form_setting.FileNameEdit1.Text);
  conf.WriteString('Model','TEXTBody',form_setting.FileNameEdit2.Text);


end;

procedure TForm_main.FormCreate(Sender: TObject);
begin
  appname := 'fast Mailer';
  memohtml.Align:=alClient;
  memotesto.Align:=alClient;
  loadserverssmtp(sender);

//  SetDefaultLang('en');
//  GetLocaleFormatSettings($409, DefaultFormatSettings);
  conf:=Tinifile.Create(extractfilepath(application.ExeName)+'default.conf');
  //seleziono il profilo smtp dal file conf
  stringgrid3.Row := stringgrid3.Cols[0].IndexOf(conf.ReadString('Server','Default',''));

  emailreply.Text:=conf.ReadString('Identity','ReplyEmail','');
  emailmitt.Text:= conf.ReadString('Identity','FromEmail','');
  nomemitt.Text:= conf.ReadString('Identity','FromName','');

(*
  subject.Text:=form_main.conf.ReadString('Model','Subject','');
  FileNameEdit1.Text:=form_main.conf.WriteString('Model','HTMLBody','');
  FileNameEdit2.Text:=form_main.conf.WriteString('Model','TEXTBody','');
*)
end;

procedure TForm_main.IdLogDebug1Receive(ASender: TIdConnectionIntercept;
  var ABuffer: TIdBytes);
begin

end;

procedure TForm_main.IdLogEvent1Received(ASender: TComponent; const AText,
  AData: string);
begin
  log(AData);
end;

procedure TForm_main.IdLogEvent1Status(ASender: TComponent; const AText: string
  );
begin
  log(atext);
end;

procedure TForm_main.MenuItem1Click(Sender: TObject);
begin
  stringgrid1.Cells[3,stringgrid1.Row]:='';
  stringgrid1.Cells[4,stringgrid1.Row]:='';
end;

procedure TForm_main.MenuItem2Click(Sender: TObject);
var
  Reply, BoxStyle : Integer;
begin

  if stringgrid1.RowCount>1 then begin;
     //BoxStyle := MB_ICONQUESTION + MB_YESNO;
       BoxStyle :=  MB_ICONWARNING + MB_YESNO;
       Reply := Application.MessageBox(pchar(resend), pchar(attentionn), BoxStyle);
       if Reply = IDYES then begin
        stringgrid1.DeleteCol(4);
        stringgrid1.DeleteCol(3);
       end;
  end;
end;

procedure TForm_main.MenuItem3Click(Sender: TObject);
begin
  if openhtml.Execute then memohtml.Lines.LoadFromFile(openhtml.FileName);
end;

procedure TForm_main.MenuItem4Click(Sender: TObject);
begin
  if opentxt.Execute then memotesto.Lines.LoadFromFile(opentxt.FileName);
end;

procedure TForm_main.PageControlMessChange(Sender: TObject);
begin

end;


procedure TForm_main.FillCell(Stringrid1 : TStringGrid; Col, Row : Integer;
BkCol, TextCol : TColor);
var
Rect : TRect;
begin
Rect := StringGrid1.CellRect(Col, Row);
with StringGrid1.Canvas do begin
Brush.Color := BkCol;
FillRect(Rect);
Font.Color := TextCol;
TextOut(Rect.Left + 2, Rect.Top + 2, Stringrid1.Cells[Col, Row]);
end;
end;

procedure TForm_main.smtpConnected(Sender: TObject);
begin
  log(serverconnect);
end;

procedure TForm_main.smtpDisconnected(Sender: TObject);
begin
  log(serverdisconnect);
end;

procedure TForm_main.smtpStatus(ASender: TObject; const AStatus: TIdStatus;
  const AStatusText: string);
begin
  //log(AStatusText);
end;

procedure TForm_main.smtpWork(ASender: TObject; AWorkMode: TWorkMode;
  AWorkCount: Int64);
begin

end;

procedure TForm_main.smtpWorkBegin(ASender: TObject; AWorkMode: TWorkMode;
  AWorkCountMax: Int64);
begin

end;

procedure TForm_main.smtpWorkEnd(ASender: TObject; AWorkMode: TWorkMode);
begin

end;

procedure TForm_main.SpeedButton10Click(Sender: TObject);
var
  xdoc: TXMLDocument;                                  // variable to document
  RootNode, parentNode, nofilho: TDOMNode;
  a:longint; // variable to nodes
  status:string;
begin

if not savedialog1.Execute then exit;

//create a document
  xdoc := TXMLDocument.create;

  //create a root node
  RootNode := xdoc.CreateElement('fastMailerProject');
  Xdoc.Appendchild(RootNode);

  //create a parent node
    RootNode:= xdoc.DocumentElement;

    parentNode := xdoc.CreateElement('Recipients');
    //TDOMElement(parentNode).SetAttribute('id', '001');       // create atributes to parent node
    RootNode.Appendchild(parentNode);                          // save parent node

    for a:=1 to stringgrid1.RowCount-1 do begin
    //create a child node
      parentNode := xdoc.CreateElement('Recipient');                // create a child node
      TDOMElement(parentNode).SetAttribute('Group', stringgrid1.Cells[0,a]);     // create atributes
      TDOMElement(parentNode).SetAttribute('Name', stringgrid1.Cells[1,a]);     // create atributes
      if stringgrid1.ColCount>3 then status:=stringgrid1.Cells[4,a] else status:='';
      TDOMElement(parentNode).SetAttribute('Status', status);     // create atributes
      nofilho := xdoc.CreateCDATASection(stringgrid1.Cells[2,a]);         // insert a value to node
      parentNode.Appendchild(nofilho);                         // save node
      RootNode.ChildNodes.Item[0].AppendChild(parentNode);       // insert child node in respective parent node
    end;

      RootNode := xdoc.FindNode('fastMailerProject');
      //create a parent node
//        RootNode:= xdoc.DocumentElement;
        parentNode := xdoc.CreateElement('Attachments');
        //TDOMElement(parentNode).SetAttribute('id', '001');       // create atributes to parent node
        RootNode.Appendchild(parentNode);                          // save parent node

         for a:=1 to stringgrid2.RowCount-1 do begin
        //create a child node
          parentNode := xdoc.CreateElement('Attachment');                // create a child node
          TDOMElement(parentNode).SetAttribute('Size',  stringgrid2.Cells[1,a]);     // create atributes
          nofilho := xdoc.CreateCDATASection( stringgrid2.Cells[0,a]);         // insert a value to node
          parentNode.Appendchild(nofilho);                         // save node
          RootNode.ChildNodes.Item[1].AppendChild(parentNode);       // insert child node in respective parent node
         end;


            //create a parent node
      //        RootNode:= xdoc.DocumentElement;
              parentNode := xdoc.CreateElement('Params');
              //TDOMElement(parentNode).SetAttribute('id', '001');       // create atributes to parent node
              RootNode.Appendchild(parentNode);                          // save parent node

              //create a child node
                parentNode := xdoc.CreateElement('From');                // create a child node
                TDOMElement(parentNode).SetAttribute('Name', nomemitt.Text);     // create atributes
                TDOMElement(parentNode).SetAttribute('Reply', emailreply.Text);     // create atributes
                nofilho := xdoc.CreateCDATASection(emailmitt.Text);         // insert a value to node
                parentNode.Appendchild(nofilho);                         // save node
                RootNode.ChildNodes.Item[2].AppendChild(parentNode);       // insert child node in respective parent node

                //create a child node
                  parentNode := xdoc.CreateElement('HTMLMessage');                // create a child node
                  TDOMElement(parentNode).SetAttribute('Subject', loggetto.Text);     // create atributes
                  nofilho := xdoc.CreateCDATASection(memohtml.Text);         // insert a value to node
                  parentNode.Appendchild(nofilho);                         // save node
                  RootNode.ChildNodes.Item[2].AppendChild(parentNode);       // insert child node in respective parent node

                  //create a child node
                    parentNode := xdoc.CreateElement('TEXTMessage');                // create a child node
                    TDOMElement(parentNode).SetAttribute('Subject', loggetto.Text);     // create atributes
                    nofilho := xdoc.CreateCDATASection(memotesto.Text);         // insert a value to node
                    parentNode.Appendchild(nofilho);                         // save node
                    RootNode.ChildNodes.Item[2].AppendChild(parentNode);       // insert child node in respective parent node

                    //create a child node
                      parentNode := xdoc.CreateElement('SMTPServer');                // create a child node
                      //TDOMElement(parentNode).SetAttribute('Subject', loggetto.Text);     // create atributes
                      nofilho := xdoc.CreateCDATASection(stringgrid3.Cells[0,stringgrid3.Row]);         // insert a value to node
                      parentNode.Appendchild(nofilho);                         // save node
                      RootNode.ChildNodes.Item[2].AppendChild(parentNode);       // insert child node in respective parent node


      writeXMLFile(xDoc,savedialog1.FileName);                     // write to XML
      Xdoc.free;                                          // free memory

      progname:= extractfilename(savedialog1.FileName);
      caption:=appname+' - '+progname;


end;

procedure TForm_main.SpeedButton11Click(Sender: TObject);
var j: integer;
  Reply, BoxStyle, a: Integer;
begin

     BoxStyle := MB_ICONQUESTION + MB_YESNO;
     Reply := Application.MessageBox(pchar(msgnewprog), pchar(msgnewprogdesc), BoxStyle);
     if Reply = IDNO then exit;

  //caption:=appname+' - '+newprogetto;
  emailreply.Text:=conf.ReadString('Identity','ReplyEmail','');
  emailmitt.Text:= conf.ReadString('Identity','FromEmail','');
  nomemitt.Text:= conf.ReadString('Identity','FromName','');
  for j:=1 to stringgrid1.RowCount-1 do  stringgrid1.DeleteRow(1);
  for j:=1 to stringgrid2.RowCount-1 do  stringgrid2.DeleteRow(1);
  loggetto.Text:='';
  memohtml.Lines.Clear;
  memotesto.Lines.Clear;

  progname:= msgnewprogdesc;
  caption:=appname+' - '+progname;
  statusbar1.Panels[0].Text:=messtosend+' '+inttostr(stringgrid1.RowCount-1);

end;

procedure TForm_main.SpeedButton1Click(Sender: TObject);
var tmpsg:TStringgrid;
    b: integer;
begin
  if opendialog1.Execute then begin
     tmpsg:=TStringgrid.Create(self);
     tmpsg.ColCount:=3;
     tmpsg.RowCount:=1;
     tmpsg.FixedCols:=0;
     tmpsg.FixedRows:=1;
     tmpsg.LoadFromCSVFile(opendialog1.FileName,',',false);
    // showmessage(inttostr(tmpsg.RowCount));
     for b:=1 to tmpsg.RowCount-1 do begin
         stringgrid1.RowCount:=stringgrid1.RowCount+1;
        // showmessage(tmpsg.Rows[b].Text);
         stringgrid1.Rows[stringgrid1.RowCount-1].text:=tmpsg.Rows[b].Text;
     end;
  end;
  tmpsg.Free;
  statusbar1.Panels[0].Text:=messtosend+' '+inttostr(stringgrid1.RowCount-1);
end;


procedure TForm_main.SpeedButton3Click(Sender: TObject);
var a:word;
begin
   if opendialog2.Execute then begin
      for a:=0 to opendialog2.Files.Count-1 do begin
          stringgrid2.RowCount:=stringgrid2.RowCount+1;
          stringgrid2.Cells[0,stringgrid2.RowCount-1]:= opendialog2.Files[a];
          stringgrid2.Cells[1,stringgrid2.RowCount-1]:= inttostr(filesize(opendialog2.Files[a]))+' bytes';
      end;
      tabsheet2.Caption:=txt9+'('+inttostr(stringgrid2.RowCount-1)+')';
   end;
end;

procedure TForm_main.SpeedButton4Click(Sender: TObject);
var
  Reply, BoxStyle: Integer;
begin
   if stringgrid2.RowCount>1 then begin
       BoxStyle := MB_ICONQUESTION + MB_YESNO;
       Reply := Application.MessageBox(pchar(msgdeleteattach), pchar(deleteattach), BoxStyle);
       if Reply = IDYES then begin
        stringgrid2.DeleteRow(stringgrid2.Row);
        tabsheet2.Caption:=txt10+'('+inttostr(stringgrid2.RowCount-1)+')';
       end;
       end;
end;

procedure TForm_main.SpeedButton5Click(Sender: TObject);
begin
  form_setting.ShowModal;
end;

procedure TForm_main.SpeedButton6Click(Sender: TObject);
var
  Reply, BoxStyle, a: Integer;
begin
    if stringgrid1.RowCount>1 then begin
     BoxStyle := MB_ICONQUESTION + MB_YESNO;
     Reply := Application.MessageBox(pchar(msgdeletedest), pchar(deletedest), BoxStyle);
     if Reply = IDYES then
        if stringgrid1.RowCount > 1 then stringgrid1.DeleteRow(stringgrid1.Row);
     //for a:=1 to stringgrid1.RowCount-1 do  stringgrid1.DeleteRow(1); // elimina tutte le righe

    end;
end;

procedure TForm_main.SpeedButton8Click(Sender: TObject);
begin
  form2.showmodal;
end;

procedure TForm_main.SpeedButton9Click(Sender: TObject);
var
   Documento: TXMLDocument;
   Child: TDOMNode;
   j: Integer;
begin
  if not opendialog3.Execute then exit;

    if stringgrid1.Columns.Count<4 then
     begin
        stringgrid1.Columns.Add;
        stringgrid1.Columns[3].Title.Caption:='Status';
        stringgrid1.Columns.Add;
        stringgrid1.Columns[4].Title.Caption:='StatusNum';
        stringgrid1.Columns[4].Visible:=false;
     end;

  // empty recipients list
  for j:=1 to stringgrid1.RowCount-1 do  stringgrid1.DeleteRow(1);

  ReadXMLFile(Documento, opendialog3.FileName);
   Memolog.Lines.Clear;
   // using FirstChild and NextSibling properties
//   Child := Documento.DocumentElement.FirstChild;

// CARICA RECIPIENTS

   Child := Documento.DocumentElement.FindNode('Recipients').FirstChild;
   while Assigned(Child) do
   begin
     //Memolog.Lines.Add(Child.NodeName + ' ' + Child.Attributes.Item[0].NodeValue);
     stringgrid1.RowCount:=stringgrid1.RowCount+1;
     stringgrid1.cells[0,stringgrid1.RowCount-1]:=Child.Attributes.Item[0].NodeValue;
     stringgrid1.cells[1,stringgrid1.RowCount-1]:=Child.Attributes.Item[1].NodeValue;
     stringgrid1.cells[4,stringgrid1.RowCount-1]:=Child.Attributes.Item[2].NodeValue;
     if Child.Attributes.Item[2].NodeValue='1' then stringgrid1.cells[3,stringgrid1.RowCount-1]:=txt5;

//     Memolog.Lines.Add(Child.NodeName);
     // using ChildNodes method
     with Child.ChildNodes do
     try
       for j := 0 to (Count - 1) do
       //    Memolog.Lines.Add(Item[j].NodeName + ' ' + Item[j].FirstChild.NodeValue);
       //Memolog.Lines.Add(Item[j].NodeValue);
        stringgrid1.cells[2,stringgrid1.RowCount-1]:=Item[j].NodeValue;
     finally
       Free;
     end;
     Child := Child.NextSibling;
   end;



   // CARICA ATTACHMENTS
   // empty attachments list
  for j:=1 to stringgrid2.RowCount-1 do  stringgrid2.DeleteRow(1);

      Child := Documento.DocumentElement.FindNode('Attachments').FirstChild;
      while Assigned(Child) do
      begin
        stringgrid2.RowCount:=stringgrid2.RowCount+1;
        stringgrid2.cells[1,stringgrid2.RowCount-1]:=Child.Attributes.Item[0].NodeValue;
       // Memolog.Lines.Add(Child.NodeName + ' ' + Child.Attributes.Item[0].NodeValue);
   //     Memolog.Lines.Add(Child.NodeName);
        // using ChildNodes method
        with Child.ChildNodes do
        try
          for j := 0 to (Count - 1) do
        //    Memolog.Lines.Add(Item[j].NodeName + ' ' + Item[j].FirstChild.NodeValue);
         //  Memolog.Lines.Add(Item[j].NodeValue);
          stringgrid2.cells[0,stringgrid2.RowCount-1]:=Item[j].NodeValue;
        finally
          Free;
        end;
        Child := Child.NextSibling;
      end;

      // CARICA PARAMS

         Child := Documento.DocumentElement.FindNode('Params').FirstChild;
         while Assigned(Child) do
         begin
           //Memolog.Lines.Add(Child.NodeName + ' ' + Child.Attributes.Item[0].NodeValue);
           if lowercase(Child.NodeName)='from' then begin
                nomemitt.Text:=Child.Attributes.Item[0].NodeValue;
                emailreply.Text:=Child.Attributes.Item[1].NodeValue;
            end;
           if (lowercase(Child.NodeName)='htmlmessage') or (lowercase(Child.NodeName)='textmessage')
              then loggetto.Text:=Child.Attributes.Item[0].NodeValue;
//           if lowercase(Child.NodeName)='textmessage' then loggetto.Text:=Child.Attributes.Item[0].NodeValue;
      //     Memolog.Lines.Add(Child.NodeName);
           // using ChildNodes method
           with Child.ChildNodes do
           try
             for j := 0 to (Count - 1) do begin
           //    Memolog.Lines.Add(Item[j].NodeName + ' ' + Item[j].FirstChild.NodeValue);
            //  Memolog.Lines.Add(Item[j].NodeValue);
               if lowercase(Child.NodeName)='from' then emailmitt.Text:=Item[j].NodeValue;
               if lowercase(Child.NodeName)='htmlmessage' then memohtml.Text:=Item[j].NodeValue;
               if lowercase(Child.NodeName)='textmessage' then memotesto.Text:=Item[j].NodeValue;
               if lowercase(Child.NodeName)='smtpserver' then stringgrid3.Row := stringgrid3.Cols[0].IndexOf(Item[j].NodeValue);
             end;
           finally
             Free;
           end;
           Child := Child.NextSibling;
         end;


      Documento.Free;

      progname:= extractfilename(opendialog3.FileName);
      caption:=appname+' - '+progname;

      tabsheet2.Caption:=txt9+'('+inttostr(stringgrid2.RowCount-1)+')';
      statusbar1.Panels[0].Text:=messtosend+' '+inttostr(stringgrid1.RowCount-1);

end;

procedure TForm_main.StatusBar1DrawPanel(StatusBar: TStatusBar;
  Panel: TStatusPanel; const Rect: TRect);
begin


end;

procedure TForm_main.StringGrid1SelectCell(Sender: TObject; aCol, aRow: Integer;
  var CanSelect: Boolean);
begin

end;

procedure TForm_main.TabhtmlContextPopup(Sender: TObject; MousePos: TPoint;
  var Handled: Boolean);
begin

end;

procedure TForm_main.Log(mess:string);
begin
    memolog.Lines.Add(datetimetostr(now)+' - '+mess);
end;
procedure TForm_main.Log1(mess:string);
begin
//    memolog1.Lines.Add(datetimetostr(now)+' - '+mess);
end;
procedure TForm_main.LoadServersSmtp(Sender: TObject);
begin
    if fileexists(extractfilepath(application.ExeName)+'smtpservers.conf') then
          stringgrid3.LoadFromCSVFile(extractfilepath(application.ExeName)+'smtpservers.conf',',',false);
end;

end.

