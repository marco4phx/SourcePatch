               unit Sorcio;

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
  Winapi.Windows, Winapi.Messages,
  System.SysUtils, System.StrUtils, System.IOUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.StdCtrls,
  Vcl.Buttons, Vcl.ComCtrls, Vcl.TabNotBk,

  units.StrReplace,

  FileCtrl, Vcl.CheckLst, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf, FireDAC.Stan.Def,
  FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys, FireDAC.Phys.MSSQL,
  FireDAC.Phys.MSSQLDef, FireDAC.Stan.Param, FireDAC.DatS,
  FireDAC.DApt.Intf, FireDAC.DApt, FireDAC.VCLUI.Wait, FireDAC.Comp.UI,
  FireDAC.Phys.ODBCBase, Data.DB, FireDAC.Comp.DataSet, FireDAC.Comp.Client;

type
  TForm1 = class(TForm)
    Panel1: TPanel;
    pbarSostituzioni: TProgressBar;
    cbxSelectAllFiles: TCheckBox;
    lviewPasFiles: TListView;
    mmLog: TMemo;
    Label2: TLabel;
    edtSearchStr: TEdit;
    Label1: TLabel;
    edtPatchStr: TEdit;
    Label3: TLabel;
    CheckBox1: TCheckBox;
    edtFrom: TEdit;
    edtTo: TEdit;
    Label4: TLabel;
    Label5: TLabel;
    edtStep: TEdit;
    Label6: TLabel;
    btnStart: TBitBtn;
    btnStop: TBitBtn;
    cbxBackupFile: TCheckBox;
    btnRefresh: TBitBtn;
    CheckBox2: TCheckBox;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    Label7: TLabel;
    labCurrentDir: TLabel;    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure cbxSelectAllFilesClick(Sender: TObject);
    procedure btnRefreshClick(Sender: TObject);
    procedure btnStartClick(Sender: TObject);
    procedure btnStopClick(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
  private
    { Private declarations }
  public
    function CreateFileList(aPath:string): boolean;
    function LoadFile(aPathAndFileName:string): boolean;
    function PatchFile:integer;
    procedure SaveFile(filePathAndName:string);
    function FindFile(aPathAndFileName:string): boolean;
    { Public declarations }
  end;

var
  Form1:      TForm1;
  currentDir: string;
  kAppPath:   string;
  arrStringaOriginale:  array of string;
  arrStringaNuova:      array of string;
  FilePAS:     TStringList;
  errCount:     integer;
  errMsg:       string;
  vIOPdir:      string;
  vIOPfList:    string;

const
    kBkpPath  = 'Bkp2\';

    kSuffix      = '.pas';
    kSuffixCopia = '_BKP.pas';




implementation

{$R *.dfm}


procedure TForm1.FormCreate(Sender: TObject);
begin
    FilePAS := TStringList.Create;

//  currentDir := 'C:\DevDX\SourcePatch\'; // xdebug
    currentDir := GetCurrentDir;

    labCurrentDir.Caption := currentDir;
    kAppPath := ExtractFilePath(Application.ExeName);
    mmLog.Lines.Add( DatetimeToStr(now)+#9#9'  *** Log PAS Patcher ***');
    btnRefreshClick(Form1);
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
    //Dialogs.MessageDlg('Exiting the Delphi application.', mtInformation, [mbOk], 0, mbOk);
    FilePAS.free;
end;

function TForm1.CreateFileList(aPath:string): boolean;
var
  ListItem: TListItem;
  srec: TSearchRec;
  FileAttrs: Integer;
  count: integer;
begin
    result := False;
    FileAttrs := 0;
    //FileAttrs := faAnyFile;
    count := 0;

    if not DirectoryExists(aPath) then begin
        MessageDlg('Directory NOT Exists !', mtError, [mbOk], 0, mbOk);
        exit;
    end;
    mmLog.Lines.Add('Search .PAS in: '+ aPath);
    aPath := currentDir +'\*.pas';

    Screen.Cursor := crHourglass;
    btnRefresh.Enabled := False;
    try
        lviewPasFiles.Clear;
        lviewPasFiles.Items.BeginUpdate;

        if FindFirst(aPath, FileAttrs, srec) = 0 then begin
            repeat
                //if (srec.Attr and FileAttrs) <> srec.Attr then continue // solo se necessita corrispondano anche attributi.
                mmLog.Lines.Add('Found  '+ srec.Name);      // IntToStr(srec.Size);
                inc(count);

                // compila listview.
                ListItem := lviewPasFiles.Items.Add;
                ListItem.Caption := srec.name;      // File name
                with ListItem.SubItems do begin
                    Add('');                        // init Esito
                    Add('');                        // init Modifiche
                    Add('')
                    // if you need more columns, add here
                end;
            until FindNext(srec) <> 0;
        end
        else begin
            inc(errCount);
            errMsg := '*** Nessun .PAS trovato in: '+ aPath;
            mmLog.Lines.Add(errMsg);
        end;
    finally
      if count > 0 then begin
          result := True;
          btnStart.Enabled := True;
      end;
      lviewPasFiles.Items.EndUpdate;
      FindClose(srec);  // FindFirst allocates resources (memory) that must be released by calling FindClose.
      Screen.Cursor := crDefault;
      btnRefresh.Enabled := True;
    end;
end;

procedure TForm1.BitBtn1Click(Sender: TObject);
var
  Replacer: TFileSearchReplace;
  StartTime: TDateTime;
begin
  StartTime:=Now;
  Replacer:=TFileSearchReplace.Create('c:\Temp\123.txt');
  try
    Replacer.Replace('some text', 'some', [rfReplaceAll, rfIgnoreCase]);
  finally
    Replacer.Free;
  end;
  Caption:=FormatDateTime('nn:ss.zzz', Now - StartTime);
end;

procedure TForm1.BitBtn2Click(Sender: TObject);
begin
    ForceCurrentDirectory  := false;
    SelectDirectory('Select a directory', '', currentDir, [], self);
    currentDir := currentDir +'\';
    labCurrentDir.Caption := currentDir+'\';
    btnRefreshClick(Form1);
end;

procedure TForm1.btnRefreshClick(Sender: TObject);
begin
    CreateFileList(currentDir);
    pbarSostituzioni.position := 0;
end;

procedure TForm1.btnStartClick(Sender: TObject);
var
  i, nPatch: integer;
  FileName: string;
  fPrefix, fFolder: string;
  fToFind, fToSave: string;
  ListItem: TListItem;
  iRev: integer;
  vRev: string;
begin
    if lviewPasFiles.items.count = 0 then begin
        MessageDlg('Necessario Selezionare Files !', mtWarning, [mbOk], 0, mbOk);
        exit;
    end;
    if (length(edtSearchStr.text) = 0) or (length(edtPatchStr.text) = 0)then begin
        MessageDlg('Necessario specificare le stringhe per Sostituzioni !', mtWarning, [mbOk], 0, mbOk);
        exit;
    end;

    mmLog.Lines.Add(#13#10+ DatetimeToStr(now)+'   *** Inizio ricerca sostituzioni ***');
    if cbxBackupFile.checked then
        mmLog.Lines.Add('Create copy of source Files.')
    else
        mmLog.Lines.Add('OverWrite source Files !');

    btnStop.tag := 0;
    btnStart.Enabled := False;
    cbxBackupFile.Enabled := False;
    Screen.Cursor := crHourglass;
    try
        for i := 0 to lviewPasFiles.Items.Count -1 do begin   // Scorre files

            application.processmessages;
            if btnStop.tag >0 then begin
                mmLog.Lines.Add('*** ABORT by USER, Elaborazione Sostituzioni Interrotta !'+#13#10 );
                exit;
            end;

            if lviewPasFiles.Items[i].Checked then begin
                mmLog.Lines.Add(#13#10);
                errCount := 0;
                errMsg := '';
                FileName := lviewPasFiles.Items[i].Caption;
                mmLog.Lines.Add('patching file: '+ FileName );

                // scorre files
                fToFind := currentDir + FileName;

                if LoadFile( fToFind ) then begin   // '\\Netserver\...\source14.pas'
                    // il file trovato è ora in TStringList FilePAS

                    if cbxBackupFile.checked then begin
                        fToSave := currentDir + kBkpPath;
                        if DirectoryExists(fToSave) then begin
                            // qui cartella è già esistente.
                            // Verifica esistenza file...
                            fToSave := fToSave + TPath.GetFileNameWithoutExtension(FileName);
                            vRev := '';  // prima cerca filename senza revisione
                            iRev := 1;
                            while FindFile( fToSave + vRev + '.pas' ) and (iRev < 50) do begin    // max 50 revisioni!
                                // file esiste già, quindi aggiunge una revisione
                                vRev := '.rev'+ intToStr(iRev);    // se c'è già aggiunge '.revX';
                                inc( iRev );
                            end;
                            fToSave := fToSave + vRev + '.pas';
                            // Salva file
                            SaveFile( fToSave );
                            mmLog.Lines.Add('Backup saved in: '+ fToSave );
                        end
                        else begin
                            // backup path non esiste, rinuncio e segnalo.
                            errMsg := '*** ERROR il backup path non esiste: '+ fToSave;
                            mmLog.Lines.Add(errMsg);
                            inc(errCount);
                            break; // termina subito
                        end;
                    end;

                    nPatch := PatchFile();      // patch effettuata sulla TStringList FileHTML
                    if nPatch > 0 then
                        // il path esiste per certo visto che abbiamo già letto i source
                        SaveFile( fToFind )     // Patched File
                    else
                        mmLog.Lines.Add('Nessuna sostituzione effettuata in '+ FileName );

                end
                else begin
                    // path non esiste, rinuncio e segnalo.
                    errMsg := '*** ERROR file .pas Non trovato: '+ fToFind;
                    mmLog.Lines.Add(errMsg);
                    inc(errCount);
                    break; // termina subito
                end;//LoadFile

                lviewPasFiles.Items.BeginUpdate;
                try
                  ListItem:= lviewPasFiles.Items.item[i];
                  //ListItem.Caption resta invariato
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
                  lviewPasFiles.Items.EndUpdate;
                end;

            end;//if
        end;//for

    finally
      mmLog.Lines.Add(DatetimeToStr(now)+'   *** FINE Elaborazioni ***' +#13#10);
      Screen.Cursor := crDefault;
      btnStart.Enabled := True;
      cbxBackupFile.Enabled := True;
    end;
end;

procedure TForm1.btnStopClick(Sender: TObject);
begin
    btnStop.tag := 1;
end;

function TForm1.LoadFile(aPathAndFileName:string): boolean;
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
            FilePAS.LoadFromFile( aPathAndFileName );
            //strFileHTML.LoadFromFile( aPathAndFileName );
            mmLog.Lines.Add('Found  '+ srec.Name);      // IntToStr(srec.Size);
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
  len, strPos, posIdx, pcount: integer;
  strToSearch: string;
  strNew: string;
  strNum: string;
  strNewWithNum: string;
  From, step: integer;
begin
    pcount := 0;
    // scorre file alla ricerca di tutti i possibili codici da sostituire
    strToSearch := edtSearchStr.Text;
    strNew      := edtPatchStr.Text;
    len := length(strToSearch);
    from := strToInt(edtFrom.Text);
    step := strToInt(edtStep.Text);
    strNum := edtFrom.Text;
  //mmLog.Lines.Add('search for '+strToSearch);
    // con strToSearch eseguo ricerca in FilePAS
    pbarSostituzioni.position := 0;
    pbarSostituzioni.max := 100;   // default
    posIdx := 1;

    repeat
        strPos := PosEx(strToSearch, FilePAS.Text, posIdx );  // ricerca flat su tutte Lines di TString ?
        //strPos := Pos(strToSearch, strFileHTML.Text);  //
        if strPos = 0 then begin
            // strToSearch non trovato, esco
            break;
        end;
        // build patch string
        strNum := intToStr(from + (step * pcount));          //  '510'
        strNewWithNum := StringReplace(strNew, '%', strNum, []);    // "LogStep(%, "   sostituisce % con stringa numerica  "LogStep(510, "
        // eseguo 1 sostituzione
        //FilePAS.text := StringReplace(FilePAS.text, strToSearch, strNew, [rfIgnoreCase] );  // [rfReplaceAll, rfIgnoreCase]
        // No, StringReplace() non adatta perchè non ha ReplaceFrom ! e quindi va in loop cambiando sempre solo la prima ricorrenza !
        FilePAS.text := StuffString(FilePAS.text, strPos, len, strNewWithNum);
        // aggiorna posizione di ricerca
        posIdx := strPos + len;  // passa al prossima da sostituire.

        inc(pcount);
        pbarSostituzioni.position := pbarSostituzioni.position +1;
    until strPos = 0;
//    pbarSostituzioni.max := pcount;
    mmLog.Lines.Add('trovato '+ pcount.toString +' ricorrenze di "'+strToSearch +'"  e sostituita con '+ strNew );
    result := pcount;
end;

procedure TForm1.SaveFile(filePathAndName:string);
begin
    try
      //FileHTML.SaveToFile( filePathAndName );

        FilePAS.SaveToFile( filePathAndName, TEncoding.UTF8 ); // altrimenti se l'origine non è in unicode, perde le accentate.
        mmLog.Lines.Add('Saved file '+ filePathAndName );

      //FileHTML.SaveToFile( filePathAndName, TEncoding.Unicode );
      {
          Encoding            Class
          ANSI                System.SysUtils.TMBCSEncoding
          ASCII               System.SysUtils.TMBCSEncoding
          Big-endian UTF-16, Big-endian Unicode       System.SysUtils.TBigEndianUnicodeEncoding
          UTF-16, Unicode     System.SysUtils.TUnicodeEncoding
          UTF-7               System.SysUtils.TUTF7Encoding
          UTF-8               System.SysUtils.TUTF8Encoding
      }
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
    try
        if FindFirst(aPathAndFileName, FileAttrs, srec) = 0 then begin
            //if (srec.Attr and FileAttrs) = srec.Attr then // solo se necessita corrispondano anche attributi.
            result := True;
        end
    finally
        FindClose(srec);  // FindFirst allocates resources (memory) that must be released by calling FindClose.
    end;
end;

procedure TForm1.cbxSelectAllFilesClick(Sender: TObject);
var
  i: integer;
begin
    lviewPasFiles.Items.BeginUpdate;
    try
      for i := 0 to lviewPasFiles.Items.Count -1 do begin
        lviewPasFiles.Items.item[i].checked := cbxSelectAllFiles.Checked;
      end
    finally
      lviewPasFiles.Items.EndUpdate;
    end;
end;




end.
