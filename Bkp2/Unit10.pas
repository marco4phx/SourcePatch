unit Unit10;

(*

todo:
    - lista cod.prodotti visualizzata con checkbox per scelta quali codici processare


    - checkbox overwrite file (or make copy)


[4:25 PM] Dario Copertino

Esempio di path DBT (per un datasheet di produzione):
    \\Netserver\Database tecnico\OI\OI200021\OI200021_DSP.html
    \\Netserver\Database tecnico\OI\OI010009\OI010009_DSP.html

Nella nuova struttura:
    \\Netserver\05.DBtecnico\OI\OI200021\OI200021_DSP.html

    \\Netserver\05.DbTecnico\OI\OI010009\OI010009_DSP.html


*)

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.StrUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.StdCtrls,
  Vcl.Buttons, Vcl.ComCtrls, Vcl.TabNotBk,
  VCL.FlexCel.Core, FlexCel.XlsAdapter,
  FileCtrl, Vcl.CheckLst;

type
  TForm1 = class(TForm)
    Label1: TLabel;
    OpenDialog: TOpenDialog;
    btnSelect: TBitBtn;
    btnStart: TBitBtn;
    btnStop: TBitBtn;
    cbxFileOverwrite: TCheckBox;
    PageControl1: TPageControl;
    TabSheet3: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet1: TTabSheet;
    mmLog: TMemo;
    edtXlsSostituzioni: TEdit;
    Label2: TLabel;
    btnOpen2: TBitBtn;
    edtXlsCodici: TEdit;
    Label3: TLabel;
    btnOpen1: TBitBtn;
    edtSourceDir: TEdit;
    cbxSkip1stLine: TCheckBox;
    btnTest: TBitBtn;
    BitBtn1: TBitBtn;
    lviewProdotti: TListView;
    cbxSelectAll: TCheckBox;
    edtDestDir: TEdit;
    Label4: TLabel;
    BitBtn2: TBitBtn;
    lviewSostituzioni: TListView;
    cbxForceDir: TCheckBox;
    BitBtn3: TBitBtn;
    procedure FormCreate(Sender: TObject);
    procedure btnSelectClick(Sender: TObject);
    procedure btnOpen1Click(Sender: TObject);
    procedure btnOpen2Click(Sender: TObject);
    procedure btnStartClick(Sender: TObject);
    procedure btnTestClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnStopClick(Sender: TObject);
    procedure cbxSelectAllClick(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
  private
    { Private declarations }
  public
    procedure ReadExcelCodici;
    procedure ReadExcelSostituzioni;
    function  FindFile(aPathAndFileName:string): boolean;
    function  PatchFile:integer;
    procedure SaveFile(filePathAndName:string);
    { Public declarations }
  end;

var
  Form1:      TForm1;
  currentDir: string;
  kAppPath:   string;
  arrCodici:    array of string;      // array dynamic, len modificabile a runtime !
  arrStringaOriginale:  array of string;
  arrStringaNuova:   array of string;
  FileHTML:     TStringList;
  strFileHTML:  TStrings;
  errCount:     integer;
  errMsg:       string;

const
    kOrigPath = '\\Netserver\Database tecnico\';
    kDestPath = '\\Netserver\05.DbTecnico\';

    kSuffix      = '_DSP.html';
    kSuffixCopia = '_DSP nuovo.html';




implementation

{$R *.dfm}


procedure TForm1.FormCreate(Sender: TObject);
begin
    FileHTML := TStringList.Create;
    strFileHTML := TStrings.Create;
    currentDir := GetCurrentDir;
    edtSourceDir.text := kOrigPath;
    edtDestDir.text   := kDestPath;

    kAppPath := ExtractFilePath(Application.ExeName);
    PageControl1.ActivePageIndex := 0;
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
    FileHTML.free;
    strFileHTML.free;
end;

procedure TForm1.btnTestClick(Sender: TObject);
var
  fMask: string;
begin
    (*
    fMask := '\OI030024_DSP.html';
    if FindFile(edtCurrDir.text + fMask) then
        if PatchFile > 0 then
            FileHTML.SaveToFile( edtCurrDir.text + fMask+'2' );
    *)
end;

procedure TForm1.cbxSelectAllClick(Sender: TObject);
var
  i: integer;
begin
    lviewProdotti.Items.BeginUpdate;
    try
      for i := 0 to lviewProdotti.Items.Count -1 do begin
        lviewProdotti.Items.item[i].checked := cbxSelectAll.Checked;
      end
    finally
      lviewProdotti.Items.EndUpdate;
    end;
end;

procedure TForm1.btnStartClick(Sender: TObject);
var
  i, nPatch: integer;
  tPath: string;
  PCode: string;
  fPrefix, fFolder: string;
  fToFind, fToSave: string;
  ListItem: TListItem;
begin
    btnStop.tag := 0;
    (*
    for i := 1 to length(arrCodici) do begin
        tStr := arrCodici[i];
        mmLog.Lines.Add(tStr);
    end;
    *)
    if length(arrCodici) = 0 then begin
        //Dialogs.MessageDlg('Confirm plugin Stop && Close ?', mtConfirmation, mbYesNo, 0, mbYes) = mrYes then
        //Dialogs.MessageDlg('SetSN Command  only !', mtWarning, [mbOK], 0);
        //esito := Dialogs.InputQuery('Insert a Serial Number', 'S/N:', NewSerial);
        //Dialogs.MessageDlg('Necessary define a Serial string !', mtError, [mbOK], 0);
        MessageDlg('Necessario specificare i Codici Prodotti !', mtWarning, [mbOk], 0, mbOk);
        exit;
    end;
    if length(arrStringaOriginale) = 0 then begin
        MessageDlg('Necessario specificare le stringhe per Sostituzioni !', mtWarning, [mbOk], 0, mbOk);
        exit;
    end;
    mmLog.Lines.Add(#13#10+ DatetimeToStr(now)+'   *** Inizio Elaborazione Codici ***');
    if cbxFileOverwrite.checked then
        mmLog.Lines.Add('OverWrite selected !')
    else
        mmLog.Lines.Add('Create copy of files.');

    btnStart.Enabled := False;
    Screen.Cursor := crHourglass;
    try

      for i := 0 to lviewProdotti.Items.Count -1 do begin   // Scorre codici prodotto

          application.processmessages;
          if btnStop.tag >0 then begin
              mmLog.Lines.Add('*** ABORT by USER Elaborazione Interrotta !'+#13#10 );
              exit;
          end;

          if lviewProdotti.Items[i].Checked  then begin
              errCount := 0;
              errMsg := '';
              PCode := lviewProdotti.Items[i].Caption;
              mmLog.Lines.Add( PCode );

              //  origine         \\Netserver\Database tecnico\  +  OI\OI010009\OI010009   + _DSP.html
              //  destinazione    \\Netserver\05.DbTecnico\      +  OI\OI010009\OI010009   + _DSP.html / _DSP copia.html

              fPrefix := copy(PCode,1,2);                       // 'OI'
              fFolder := fPrefix +'\'+ PCode  +'\'+ PCode;      //  'OI\OI010009\OI010009'
              fToFind := edtSourceDir.text + fFolder + kSuffix;

              if FindFile( fToFind ) then begin   // '_DSP.html'
                  nPatch := PatchFile();
                  if nPatch > 0 then begin

                      fToSave := edtDestDir.text + fFolder;
                      if cbxFileOverwrite.checked then
                          fToSave := fToSave + kSuffix        // '_DSP.html'
                      else
                          fToSave := fToSave + kSuffixCopia;  // '_DSP copia.html'

                      tPath := ExtractFilePath( fToSave );    // or GetDirectoryName
                      if DirectoryExists(tPath) then begin
                          SaveFile( fToSave );
                      end
                      else begin
                          //if not CreateDir('C:\temp') then
                          //  raise Exception.Create('Cannot create C:\temp')
                          //mmLog.Lines.Add('Attenzione il path non esiste: '+ tPath);

                          if cbxForceDir.checked then begin
                              // path non esiste, ma forzo la creazione
                              if not forcedirectories(tPath) then begin      // CreateDir() crea solo single dir.
                                  errMsg := '*** ERROR Cannot create directory, rc:'+ IntToStr(GetLastError);
                                  mmLog.Lines.Add(errMsg);
                                  inc(errCount);
                              end
                              else begin
                                  mmLog.Lines.Add('Path non esisteva, ma Creato!');
                                  SaveFile( fToSave );
                              end;
                          end
                          else begin
                              // path non esiste, rinuncio e segnalo.
                              errMsg := '*** ERROR il path non esiste: '+ tPath;
                              mmLog.Lines.Add(errMsg);
                              inc(errCount);
                          end;
                      end;

                  end;//PatchFile
              end;//FindFile

              lviewProdotti.Items.BeginUpdate;
              try
                ListItem:= lviewProdotti.Items.item[i];
                //ListItem.Caption resta invariato, es. 'OI010009'
                if errCount = 0 then begin
                    ListItem.SubItems[0]:= 'OK';
                    ListItem.SubItems[1]:= intToStr(nPatch);
                    ListItem.SubItems[2]:= fToSave;
                end
                else begin
                    ListItem.SubItems[0]:= 'ERROR !';
                    ListItem.SubItems[1]:= intToStr(errCount)+' errors';
                    ListItem.SubItems[2]:= errMsg;
                end;
              finally
                lviewProdotti.Items.EndUpdate;
              end;

          end;
      end;
    finally
      mmLog.Lines.Add(DatetimeToStr(now)+'   *** FINE Elaborazione Codici ***' +#13#10);
      Screen.Cursor := crDefault;
      btnStart.Enabled := True;
    end;


end;

procedure TForm1.SaveFile(filePathAndName:string);
begin
    try
        FileHTML.SaveToFile( filePathAndName );
    except
      errMsg := '*** ERROR cannot Save file: '+ filePathAndName;
      mmLog.Lines.Add(errMsg);
      inc(errCount);
    end;
end;

function TForm1.FindFile(aPathAndFileName:string): boolean;
var
  srec: TSearchRec;
  FileAttrs: Integer;
begin
    result := False;
    FileAttrs := 0;
    //FileAttrs := faAnyFile;
    mmLog.Lines.Add('Search ->  '+ aPathAndFileName);
    try
        if FindFirst(aPathAndFileName, FileAttrs, srec) = 0 then begin
            //if (srec.Attr and FileAttrs) = srec.Attr then // solo se necessita corrispondano anche attributi.
            FileHTML.LoadFromFile( aPathAndFileName );
            //strFileHTML.LoadFromFile( aPathAndFileName );
            mmLog.Lines.Add('Found  '+ srec.Name);  // IntToStr(srec.Size);
            result := True;
        end
        else begin
            inc(errCount);
            errMsg := '*** ERROR Source file NOT found: '+ aPathAndFileName;
            mmLog.Lines.Add(errMsg);
        end;
    finally
        FindClose(srec);  // FindFirst allocates resources (memory) that must be released by calling FindClose.
    end;
end;

function TForm1.PatchFile:integer;
var
  i, strPos, strIdx, pcount: integer;
  strToSearch: string;
  strNew: string;
  newFile: string;
begin
    pcount := 0;
    // scorre file alla ricerca di tutti i possibili codici da sostituire

    for i := 1 to length(arrStringaOriginale) do begin
        strToSearch := arrStringaOriginale[i];
        strNew      := arrStringaNuova[i];
      //mmLog.Lines.Add('search for '+strToSearch);
        // con codice strToSearch eseguo ricerca in FileHTML

        strPos := Pos(strToSearch, FileHTML.Text);  // ricerca flat su tutte Lines di TString ?
        //strPos := Pos(strToSearch, strFileHTML.Text);  //
        if strPos = 0 then begin
            // Errore strToSearch non trovato, proseguo con codice successivo
            continue;
        end;
        // Codice trovato, eseguo patch FileHTML
        inc(pcount);
        FileHTML.text := ReplaceText(FileHTML.text, strToSearch, strNew );
        mmLog.Lines.Add('trovato '+strToSearch +'  e sostituito con '+ strNew );

    end;//for

    result := pcount;
end;


procedure TForm1.BitBtn3Click(Sender: TObject);
begin
    mmLog.clear;
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
        //edtXlsSostituzioni.text := ExtractFileName(OpenDialog.FileName);
        edtXlsCodici.text := OpenDialog.FileName;
        // estrae codici da excel
        ReadExcelCodici();
//      btnStart.Enabled := True;
    end;
end;
procedure TForm1.ReadExcelCodici;
var
  xls: TXlsFile;
  row, colIndex: integer;
  rowCount: integer;
  XF: integer;
  cell: TCellValue;
  addr: TCellAddress;
  s: string;
  ListItem: TListItem;
begin
  xls := TXlsFile.Create( edtXlsCodici.text );
  try
    xls.ActiveSheetByName := 'Sheet1';  //we'll read sheet1. We could loop over the existing sheets by using xls.SheetCount and xls.ActiveSheet
    rowCount := xls.RowCount;
    lviewProdotti.Clear;
    lviewProdotti.Items.BeginUpdate;
    mmLog.Lines.Add('Aperto file Codici prodotto: '+ edtXlsCodici.text);
    mmLog.Lines.Add('Trovati '+ rowCount.ToString +' codici prodotto.');
    SetLength(arrCodici, rowCount);
    for row := 1 to rowCount do begin
        colIndex := 1;  // to xls.ColCountInRow(row) do //Don't use xls.ColCount as it is slow: See http://www.tmssoftware.biz/flexcel/doc/vcl/guides/performance-guide.html#avoid-calling-colcount
        XF := -1;

        cell := xls.GetCellValueIndexed(row, colIndex, XF);
        addr := TCellAddress.Create(row, xls.ColFromIndex(row, colIndex));
        (*
        s := (addr.CellRef + ' contiene ');
        if (cell.IsString) then begin
            s := s + 'la stringa: ' + cell.ToString;
            arrCodici[row] := cell.ToString;
        end
        else
            s := s + ('Error: Unknown cell type');
        *)

        // compila listview codici prodotto.
        ListItem:= lviewProdotti.Items.Add;
        if cell.IsString then begin
            arrCodici[row] := cell.ToString;
            ListItem.Caption := cell.ToString;  // init Codice
            with ListItem.SubItems do begin
                Add('');                        // init Esito
                Add('');                        // init Modifiche
                Add('');                        // init Path
                // if you need more columns, add here
            end;
        end
        else begin
            arrCodici[row] := 'ERROR';
            ListItem.Caption := 'ERRORE';
            with ListItem.SubItems do begin
                Add('Unknown cell type !');
            end;
        end;

    end;
    btnStart.Enabled := True;
  finally
    lviewProdotti.Items.EndUpdate;
    xls.Free;
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
        //edtXlsSostituzioni.text := ExtractFileName(OpenDialog.FileName);
        edtXlsSostituzioni.text := OpenDialog.FileName;
        // estrae codici da excel
        ReadExcelSostituzioni();
//      btnStart.Enabled := True;
    end;
end;
procedure TForm1.ReadExcelSostituzioni;
var
  xls: TXlsFile;
  row, colIndex: integer;
  rowCount, from: integer;
  XF: integer;
  cell: TCellValue;
  s: string;
  ListItem: TListItem;
begin
  xls := TXlsFile.Create( edtXlsSostituzioni.text );
  try
    xls.ActiveSheetByName := 'Sheet1';  //we'll read sheet1. We could loop over the existing sheets by using xls.SheetCount and xls.ActiveSheet
    rowCount := xls.RowCount;
    mmLog.Lines.Add('Aperto file sostituzioni: '+ edtXlsSostituzioni.text);
    mmLog.Lines.Add('Trovate '+ rowCount.ToString +' sostituzioni');

    lviewSostituzioni.Clear;
    lviewSostituzioni.Items.BeginUpdate;

    SetLength(arrStringaOriginale, rowCount);
    SetLength(arrStringaNuova, rowCount);

    if cbxSkip1stLine.checked then
        from := 2
    else
        from := 1;
    for row := from to rowCount do begin
        XF := -1;

        // compila listview codici prodotto.
        ListItem:= lviewSostituzioni.Items.Add;

        colIndex := 1;
        cell := xls.GetCellValueIndexed(row, colIndex, XF);
        if cell.IsString then begin
            s := cell.ToString;
            arrStringaOriginale[row] := s;
            ListItem.Caption         := s;
        end
        else begin
            arrStringaOriginale[row] := 'ERROR';
            ListItem.Caption := 'ERROR:Unknown cell type !';
        end;

        colIndex := 2;  // to xls.ColCountInRow(row) do //Don't use xls.ColCount as it is slow: See http://www.tmssoftware.biz/flexcel/doc/vcl/guides/performance-guide.html#avoid-calling-colcount
        cell := xls.GetCellValueIndexed(row, colIndex, XF);
        if cell.IsString then begin
            s := cell.ToString;
            arrStringaNuova[row] := s;
            ListItem.SubItems.Add(s);
        end
        else begin
            arrStringaNuova[row] := 'ERROR';
            ListItem.SubItems.Add( 'ERROR:Unknown cell type !' );
        end;

    end;//for
    btnStart.Enabled := True;
  finally
    lviewSostituzioni.Items.EndUpdate;
    xls.Free;
  end;
end;



procedure TForm1.btnStopClick(Sender: TObject);
begin
    //
    btnStop.tag := 1;
end;

procedure TForm1.BitBtn1Click(Sender: TObject);
begin
    mmLog.Lines.SaveToFile( currentDir + '\Log_'+ FormatDateTime('yyyymmdd-hhnnss', now) +'.txt');
end;


procedure TForm1.btnSelectClick(Sender: TObject);
var
  dlgSelectFolders: TFileOpenDialog;
begin
  //currentDir := 'C:\';     // Set the starting directory
  //currentDir := 'C:\Temp';   // Set the starting directory
  //currentDir := kAppPath;  // Set the starting directory

  dlgSelectFolders := TFileOpenDialog.Create(Form1);
  try
    if (Sender as TBitBtn).tag > 0 then
        dlgSelectFolders.DefaultFolder := edtDestDir.text
    else
        dlgSelectFolders.DefaultFolder := edtSourceDir.text;
    dlgSelectFolders.Options := dlgSelectFolders.Options + [fdoPickFolders];

    if not dlgSelectFolders.Execute then
        Abort;

    if (Sender as TBitBtn).tag > 0 then
        edtDestDir.text := dlgSelectFolders.FileName +'\'
    else
        edtSourceDir.text := dlgSelectFolders.FileName +'\';

  finally
    dlgSelectFolders.Free;
  end;
end;



end.
