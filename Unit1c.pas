unit Unit1c;

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
  VCL.FlexCel.Core, FlexCel.XlsAdapter,
  FileCtrl;

type
  TForm1 = class(TForm)
    Label1: TLabel;
    OpenDialog: TOpenDialog;
    btnSelect: TBitBtn;
    btnStart: TBitBtn;
    btnStop: TBitBtn;
    cbxSubdirs: TCheckBox;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    mmLog: TMemo;
    edtXlsSostituzioni: TEdit;
    Label2: TLabel;
    btnOpen2: TBitBtn;
    edtXlsCodici: TEdit;
    Label3: TLabel;
    btnOpen1: TBitBtn;
    SelectDialog: TOpenDialog;
    edtCurrDir: TEdit;
    mmCodici: TMemo;
    mmSostituzioni: TMemo;
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure btnSelectClick(Sender: TObject);
    procedure btnOpen1Click(Sender: TObject);
    procedure btnOpen2Click(Sender: TObject);
    procedure btnStartClick(Sender: TObject);
  private
    { Private declarations }
  public
    procedure ReadExcelFile(const aFile: string);
    { Public declarations }
  end;

var
  Form1:      TForm1;
  currentDir: string;
  kAppPath:   string;
  arrCodici:    array of string;      // array dynamic, len modificabile a runtime !
  arrLinkOrig:  array of string;
  arrLinkNew:   array of string;


implementation

{$R *.dfm}


procedure TForm1.FormCreate(Sender: TObject);
begin
    currentDir := GetCurrentDir;
    edtCurrDir.Text := currentDir;
    kAppPath := ExtractFilePath(Application.ExeName);
end;

procedure TForm1.ReadExcelFile(const aFile: string);
var
  xls: TXlsFile;
  row, colIndex: integer;
  rowCount: integer;
  XF: integer;
  cell: TCellValue;
  addr: TCellAddress;
  s: string;
begin
  xls := TXlsFile.Create(aFile);
  try
    xls.ActiveSheetByName := 'Sheet1';  //we'll read sheet1. We could loop over the existing sheets by using xls.SheetCount and xls.ActiveSheet
    rowCount := xls.RowCount;
    mmCodici.Lines.Add('Trovati '+ rowCount.ToString +' codici:');
    SetLength(arrCodici, rowCount);
    for row := 1 to rowCount do begin
        colIndex := 1;  // to xls.ColCountInRow(row) do //Don't use xls.ColCount as it is slow: See http://www.tmssoftware.biz/flexcel/doc/vcl/guides/performance-guide.html#avoid-calling-colcount
        XF := -1;

        cell := xls.GetCellValueIndexed(row, colIndex, XF);
        addr := TCellAddress.Create(row, xls.ColFromIndex(row, colIndex));
        s := (addr.CellRef + ' contiene ');
        if (cell.IsString) then begin
            s := s + 'la stringa: ' + cell.ToString;
            arrCodici[row] := cell.ToString;
        end
        else
            s := s + ('Error: Unknown cell type');

        mmCodici.Lines.Add(s);
    end;
    btnStart.Enabled := True;
  finally
    xls.Free;
  end;
end;

procedure TForm1.btnStartClick(Sender: TObject);
var
  i: integer;
  tStr: string;
begin
    for i := 1 to length(arrCodici) do begin
        tStr := arrCodici[i];
        mmLog.Lines.Add(tStr);
    end;
end;

procedure TForm1.btnOpen1Click(Sender: TObject);
begin
    //OpenDialog.InitialDir := 'C:\';
    OpenDialog.InitialDir := kAppPath;
    OpenDialog.FileName := '';
    ForceCurrentDirectory  := false;
    OpenDialog.FilterIndex := 0; // visualizza solo *.XLS files.
    application.ProcessMessages;

    // qui solo scelta file
    if OpenDialog.execute then begin
        edtXlsCodici.text := ExtractFileName(OpenDialog.FileName);

        // estrae codici da excel
        ReadExcelFile(OpenDialog.FileName);
//      btnStart.Enabled := True;
    end;
end;

procedure TForm1.btnOpen2Click(Sender: TObject);
begin
    //OpenDialog.InitialDir := 'C:\';
    OpenDialog.InitialDir := kAppPath;
    OpenDialog.FileName := '';
    ForceCurrentDirectory  := false;
    OpenDialog.FilterIndex := 0; // visualizza solo *.XLS files.
    application.ProcessMessages;

    // qui solo scelta file
    if OpenDialog.execute then begin
        edtXlsSostituzioni.text := ExtractFileName(OpenDialog.FileName);

        // estrae codici da excel
        ReadExcelFile(OpenDialog.FileName);
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


(*
                  // con codice Paziente PID eseguo lookup su Anagrafica
                  patIdx := AnsiIndexText (tRequestPID, demoPIDs );        // oppure IndexStr()
                  if patIdx < 0 then begin
                      // Errore PID non trovato implica problema grave !
                      LogSysEvent(svERROR, 202, 'TEST mode> PID not found='+ tRequestPID); // e su memo!
                      Next;
                      continue;
                  end;


*)

end.
