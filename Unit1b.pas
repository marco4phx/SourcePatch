unit Unit1b;

(*
[4:25 PM] Dario Copertino

Esempio di path DBT (per un datasheet di produzione):
    \\Netserver\Database tecnico\OI\OI200021\OI200021_DSP.html
Nella nuova struttura:
    \\Netserver\05.DBtecnico\OI\OI200021\OI200021_DSP.html


*)

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.StdCtrls,
  Vcl.Buttons, Vcl.ComCtrls, Vcl.TabNotBk,
  FileCtrl;

type
  TForm1 = class(TForm)
    Label1: TLabel;
    OpenDialog1: TOpenDialog;
    btnSelect: TBitBtn;
    btnStart: TBitBtn;
    btnStop: TBitBtn;
    cbxSubdirs: TCheckBox;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    ListBox2: TListBox;
    ListBox1: TListBox;
    Memo1: TMemo;
    edtXlsSostituzioni: TEdit;
    Label2: TLabel;
    btnOpen2: TBitBtn;
    edtXlsCodici: TEdit;
    Label3: TLabel;
    btnOpen1: TBitBtn;
    OpenDialog2: TOpenDialog;
    SelectDialog: TOpenDialog;
    edtCurrDir: TEdit;
    procedure FormCreate(Sender: TObject);
    procedure btnSelectClick(Sender: TObject);
    procedure btnOpen1Click(Sender: TObject);
    procedure btnOpen2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1:      TForm1;
  currentDir: string;
  kAppPath:   string;

implementation

{$R *.dfm}

procedure TForm1.btnOpen1Click(Sender: TObject);
begin
    //OpenDialog1.InitialDir := 'C:\';
    OpenDialog1.InitialDir := kAppPath;
    OpenDialog1.FileName := '';
    ForceCurrentDirectory  := false;
    OpenDialog1.FilterIndex := 0; // visualizza solo *.XLS files.
    application.ProcessMessages;

    // qui solo scelta file
    if OpenDialog1.execute then begin
        edtXlsCodici.text := ExtractFileName(OpenDialog1.FileName);

//      edtXlsCodici.text := ExtractFileName(OpenDialog1.FileName);
        //edtXlsCodici.text := OpenDialog1.FileName;
//      btnStart.Enabled := True;
    end;
end;

procedure TForm1.btnOpen2Click(Sender: TObject);
begin
    //OpenDialog2.InitialDir := 'C:\';
    OpenDialog2.InitialDir := kAppPath;
    OpenDialog2.FileName := '';
    ForceCurrentDirectory  := false;
    OpenDialog2.FilterIndex := 0; // visualizza solo *.XLS files.
    application.ProcessMessages;

    // qui solo scelta file
    if OpenDialog2.execute then begin
        edtXlsSostituzioni.text := ExtractFileName(OpenDialog2.FileName);

//      edtXlsCodici.text := ExtractFileName(OpenDialog2.FileName);
        //edtXlsCodici.text := OpenDialog2.FileName;
//      btnStart.Enabled := True;
    end;
end;

procedure TForm1.btnSelectClick(Sender: TObject);
var
  OpenDialog: TFileOpenDialog;
begin
  currentDir := 'C:\';     // Set the starting directory
  //currentDir := kAppPath;  // Set the starting directory

  OpenDialog := TFileOpenDialog.Create(Form1);
  try
    OpenDialog.DefaultFolder := currentDir;
    OpenDialog.Options := OpenDialog.Options + [fdoPickFolders];
    if not OpenDialog.Execute then
        Abort;

    currentDir := OpenDialog.FileName;
    edtCurrDir.text := currentDir;
  finally
    OpenDialog.Free;
  end;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
    currentDir := GetCurrentDir;
    edtCurrDir.Text := currentDir;
    kAppPath := ExtractFilePath(Application.ExeName);
end;

end.
