               unit Unit1;

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
  FileCtrl, Vcl.CheckLst, FireDAC.Stan.Intf, FireDAC.Stan.Option,
  FireDAC.Stan.Error, FireDAC.UI.Intf, FireDAC.Phys.Intf, FireDAC.Stan.Def,
  FireDAC.Stan.Pool, FireDAC.Stan.Async, FireDAC.Phys, FireDAC.Phys.MSSQL,
  FireDAC.Phys.MSSQLDef, FireDAC.Stan.Param, FireDAC.DatS,
  FireDAC.DApt.Intf, FireDAC.DApt, FireDAC.VCLUI.Wait, FireDAC.Comp.UI,
  FireDAC.Phys.ODBCBase, Data.DB, FireDAC.Comp.DataSet, FireDAC.Comp.Client;

type
  TForm1 = class(TForm)
    OpenDialog: TOpenDialog;
    btnStart: TBitBtn;
    btnStop: TBitBtn;
    cbxFileOverwrite: TCheckBox;
    pageControl1: TPageControl;
    tabLog: TTabSheet;
    tabSostituzioni: TTabSheet;
    tabProdotti: TTabSheet;
    mmLog: TMemo;
    edtXlsSostituzioni: TEdit;
    Label2: TLabel;
    btnOpen2: TBitBtn;
    edtXlsCodici: TEdit;
    Label3: TLabel;
    btnOpen1: TBitBtn;
    cbxSkip1stLine: TCheckBox;
    btnTest: TBitBtn;
    BitBtn1: TBitBtn;
    lviewProdotti: TListView;
    cbxSelectAll_Prods: TCheckBox;
    lviewSostituzioni: TListView;
    cbxForceDir: TCheckBox;
    BitBtn3: TBitBtn;
    rgArea: TRadioGroup;
    Panel1: TPanel;
    pbarSostituzioni: TProgressBar;
    tabIOP: TTabSheet;
    lviewIOP: TListView;
    PageControl2: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    Label1: TLabel;
    Label4: TLabel;
    btnSelectSource: TBitBtn;
    edtSourceDir: TEdit;
    edtDestDir: TEdit;
    btnSelectTarget: TBitBtn;
    cbxReparto: TComboBox;
    Label5: TLabel;
    edtIopDir: TEdit;
    cbxSelectAll_IOP: TCheckBox;
    gammaConnection1: TFDConnection;
    FDPhysMSSQLDriverLink1: TFDPhysMSSQLDriverLink;
    FDGUIxWaitCursor1: TFDGUIxWaitCursor;
    btnLoad: TBitBtn;
    viewProdotti: TFDQuery;
    BitBtn2: TBitBtn;
    procedure FormCreate(Sender: TObject);
    procedure btnSelectSourceClick(Sender: TObject);
    procedure btnOpen1Click(Sender: TObject);
    procedure btnOpen2Click(Sender: TObject);
    procedure btnStartClick(Sender: TObject);
    procedure btnTestClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure btnStopClick(Sender: TObject);
    procedure cbxSelectAll_ProdsClick(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure BitBtn3Click(Sender: TObject);
    procedure rgAreaClick(Sender: TObject);
    procedure cbxRepartoChange(Sender: TObject);
    procedure cbxSelectAll_IOPClick(Sender: TObject);
    procedure btnLoadClick(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
  private
    { Private declarations }
  public
    procedure ReadExcelCodiciProd;
    procedure ReadExcelSostituzioni;
    procedure ReadExcelIOP(aPathAndFileName:string);
    function  FindFile(aPathAndFileName:string): boolean;
    function  PatchFile:integer;
    procedure SaveFile(filePathAndName:string);
    { Public declarations }
  end;

var
  Form1:      TForm1;
  currentDir: string;
  kAppPath:   string;
  arrStringaOriginale:  array of string;
  arrStringaNuova:      array of string;
  FileHTML:     TStringList;
  strFileHTML:  TStrings;
  errCount:     integer;
  errMsg:       string;
  vIOPdir:      string;
  vIOPfList:    string;

const
    kOrigPath = '\\Netserver\Database tecnico\';
    kDestPath = '\\Netserver\05.DbTecnico\';
    kIopPath  = '\\Netserver\06.Qualità\SGQ2_attività\01_Mansionari_IOP_MOD\';

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
    edtIopDir.text    := kIopPath;

    kAppPath := ExtractFilePath(Application.ExeName);
    tabIOP.TabVisible := False;
    pageControl1.ActivePageIndex := 0;
    pageControl2.ActivePageIndex := 0;
    mmLog.Lines.Add( DatetimeToStr(now)+'   >>> Log elaborazioni su DBTecnico');
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
    //Dialogs.MessageDlg('Exiting the Delphi application.', mtInformation, [mbOk], 0, mbOk);
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

    gammaConnection1.Connected := True;
    if gammaConnection1.Connected then
        MessageDlg('DB Gamma CONNESSO !', mtInformation, [mbOk], 0, mbOk)
    else
        MessageDlg('DB Gamma NON disponibile !', mtError, [mbOk], 0, mbOk);
end;

procedure TForm1.rgAreaClick(Sender: TObject);
begin
    if rgArea.itemindex = 0 then begin
        pageControl2.ActivePageIndex := 0;
        tabProdotti.TabVisible := True;
        tabIOP.TabVisible := False;
        pageControl1.ActivePageIndex := 0;
        mmLog.Lines.Add(#13#10+ DatetimeToStr(now)+'   >>> inizio Log elaborazioni su DBTecnico');
    end
    else begin
        pageControl2.ActivePageIndex := 1;
        tabProdotti.TabVisible := False;
        tabIOP.TabVisible := True;
        pageControl1.ActivePageIndex := 1;
        mmLog.Lines.Add(#13#10+ DatetimeToStr(now)+'   >>> inizio Log elaborazioni su IOP files');
    end;
end;

procedure TForm1.cbxSelectAll_ProdsClick(Sender: TObject);
var
  i: integer;
begin
    lviewProdotti.Items.BeginUpdate;
    try
      for i := 0 to lviewProdotti.Items.Count -1 do begin
        lviewProdotti.Items.item[i].checked := cbxSelectAll_Prods.Checked;
      end
    finally
      lviewProdotti.Items.EndUpdate;
    end;
end;

procedure TForm1.cbxSelectAll_IOPClick(Sender: TObject);
var
  i: integer;
begin
    lviewIOP.Items.BeginUpdate;
    try
      for i := 0 to lviewIOP.Items.Count -1 do begin
        lviewIOP.Items.item[i].checked := cbxSelectAll_IOP.Checked;
      end
    finally
      lviewIOP.Items.EndUpdate;
    end;
end;

procedure TForm1.cbxRepartoChange(Sender: TObject);
begin
    // scelta Reparto implica scelta file lista da cui caricare i nomi IOP
    case cbxReparto.itemindex of
    0: begin
         vIOPdir := kIopPath + '01.SviluppoCommerciale\02.IOP\';  // \\Netserver\06.Qualità\SGQ2_attività\01_Mansionari_IOP_MOD\01.SviluppoCommerciale\02.IOP
         vIOPfList := '01.Lista IOP.VEN.xlsx';
       end;
    1: begin
         vIOPdir := kIopPath + '02.R&D\02.IOP\';                  // \\Netserver\06.Qualità\SGQ2_attività\01_Mansionari_IOP_MOD\02.R&D\02.IOP
         vIOPfList := '01.Lista IOP.R&D.xlsx';
       end;
    2: begin
         vIOPdir := kIopPath + '03.Progettazione\02.IOP\';        // \\Netserver\06.Qualità\SGQ2_attività\01_Mansionari_IOP_MOD\03.Progettazione\02.IOP
         vIOPfList := '01.Lista IOP.PRG.xlsx';
       end;
    3: begin
         vIOPdir := kIopPath + '04.Produzione\02.IOP\';           // \\Netserver\06.Qualità\SGQ2_attività\01_Mansionari_IOP_MOD\04.Produzione\02.IOP
         vIOPfList := '01.Lista IOP.PRD.xlsx';
       end;
    4: begin
         vIOPdir := kIopPath + '05.Qualità\02.IOP\';              // \\Netserver\06.Qualità\SGQ2_attività\01_Mansionari_IOP_MOD\05.Qualità\02.IOP
         vIOPfList := '01.Lista IOP.SGQ.xlsx';
       end;
    5: begin
         vIOPdir := kIopPath + '06.Sicurezza\02.IOP\';            // \\Netserver\06.Qualità\SGQ2_attività\01_Mansionari_IOP_MOD\06.Sicurezza\02.IOP
         vIOPfList := '01.Lista IOP.SIC.xlsx';
       end;
    6: begin
         vIOPdir := kIopPath + '08.ICT\02.IOP\';                  // \\Netserver\06.Qualità\SGQ2_attività\01_Mansionari_IOP_MOD\08.ICT\02.IOP
         vIOPfList := '01.Lista IOP.ICT.xlsx';
       end;
    7: begin
         vIOPdir := kIopPath + '11.HR\02.IOP\';                   // \\Netserver\06.Qualità\SGQ2_attività\01_Mansionari_IOP_MOD\11.HR\02.IOP
         vIOPfList := '01.Lista IOP.HR.xlsx';
       end;
    8: begin
         vIOPdir := kIopPath + '12.FCA\02.IOP\';                  // \\Netserver\06.Qualità\SGQ2_attività\01_Mansionari_IOP_MOD\12.FCA\02.IOP
         vIOPfList := '01.Lista IOP.FCA.xlsx';
       end;
    end;
    edtIopDir.Text := vIOPdir;
    application.ProcessMessages;
    // estrae codici documenti IOP da excel
    ReadExcelIOP( vIOPdir + vIOPfList );
//      btnStart.Enabled := True;
end;
procedure TForm1.ReadExcelIOP(aPathAndFileName:string);
var
  xls:  TXlsFile;
  cell: TCellValue;
  addr: TCellAddress;
  row, colIndex: integer;
  rowCount, IopCount: integer;
  sIopType: string;
  ListItem: TListItem;
begin
  xls := TXlsFile.Create( aPathAndFileName );
  // \\Netserver\06.Qualità\SGQ2_attività\01_Mansionari_IOP_MOD\01.SviluppoCommerciale\02.IOP\01.Lista IOP.VEN.xlsx
  // \\Netserver\06.Qualità\SGQ2_attività\01_Mansionari_IOP_MOD\02.R&D\02.IOP\01.Lista IOP.R&D.xlsx
  // \\Netserver\06.Qualità\SGQ2_attività\01_Mansionari_IOP_MOD\03.Progettazione\02.IOP\01.Lista IOP.PRG.xlsx
  try
    xls.ActiveSheet := 1;        // we'll read 1st sheet1. We could loop over the existing sheets by using xls.SheetCount and xls.ActiveSheet

    rowCount := xls.RowCount;
    sIopType := LeftStr(RightStr(aPathAndFileName, 8), 3);          // = 'VEN', 'R&D', 'PRG' ...
    IopCount := 0;
    mmLog.Lines.Add('Aperto file Lista IOP: '+ aPathAndFileName);

    lviewIOP.Clear;
    lviewIOP.Items.BeginUpdate;

    for row := 1 to rowCount do begin        // codici iniziano sempre da riga 7 ?
        colIndex := 1;  // to xls.ColCountInRow(row) do //Don't use xls.ColCount as it is slow: See http://www.tmssoftware.biz/flexcel/doc/vcl/guides/performance-guide.html#avoid-calling-colcount
      //cell := xls.GetCellValueIndexed(row, colIndex, XF);  //GetCellValueIndexed only counts the used cells.
        cell := xls.GetCellValue(row, colIndex);
        addr := TCellAddress.Create(row, xls.ColFromIndex(row, colIndex));

        // compila listview IOP.
        if cell.IsString and StartsText(sIopType, cell.ToString) then begin
            ListItem := lviewIOP.Items.Add;
            inc(IopCount);
            ListItem.Caption := cell.ToString;  // init Codice
            with ListItem.SubItems do begin
                Add('');                        // init Esito
                Add('');                        // init Modifiche
                Add('');                        // init Path
                // if you need more columns, add here
            end;
        end;

    end;
    mmLog.Lines.Add('Trovati '+ IopCount.ToString +' documenti IOP.');
  finally
    lviewIOP.Items.EndUpdate;
    xls.Free;
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
    if rgArea.itemindex = 0 then begin
        if lviewProdotti.items.count = 0 then begin
            MessageDlg('Necessario aprire lista Codici Prodotto !', mtWarning, [mbOk], 0, mbOk);
            exit;
        end;
    end
    else
        if lviewIOP.items.count = 0 then begin
            MessageDlg('Necessario Selezionare reparto IOP !', mtWarning, [mbOk], 0, mbOk);
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

    btnStop.tag := 0;
    btnStart.Enabled := False;
    rgArea.Enabled := False;
    btnSelectSource.Enabled := False;
    btnSelectTarget.Enabled := False;
    btnOpen1.Enabled := False;
    btnOpen2.Enabled := False;
    cbxFileOverwrite.Enabled := False;
    cbxForceDir.Enabled := False;
    Screen.Cursor := crHourglass;
    try

      if rgArea.itemindex = 0 then

        // Area DbTecnico
        for i := 0 to lviewProdotti.Items.Count -1 do begin   // Scorre codici prodotto

            application.processmessages;
            if btnStop.tag >0 then begin
                mmLog.Lines.Add('*** ABORT by USER, Elaborazione DbTecnico Interrotta !'+#13#10 );
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
                    // il file trovato è ora in TStringList FileHTML
                    nPatch := PatchFile();  // patch effettuata sulla TStringList FileHTML
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
                            // Ora distinguo fra le Opzioni scelte:
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
        end

      else

        // Area IOP
        for i := 0 to lviewIOP.Items.Count -1 do begin   // Scorre IOP

            application.processmessages;
            if btnStop.tag >0 then begin
                mmLog.Lines.Add('*** ABORT by USER, Elaborazione IOP Interrotta !'+#13#10 );
                exit;
            end;

            if lviewIOP.Items[i].Checked  then begin
                errCount := 0;
                errMsg := '';
                PCode := lviewIOP.Items[i].Caption;
                mmLog.Lines.Add( PCode );

                // scorre files in Area IOP
                // \\Netserver\06.Qualità\SGQ2_attività\01_Mansionari_IOP_MOD\03.Progettazione\02.IOP\PRG.014.html
                fToFind := vIOPdir + PCode  +'.html';

                if FindFile( fToFind ) then begin   // '\\Netserver\...\PRG.014.html'
                    // il file trovato è ora in TStringList FileHTML
                    nPatch := PatchFile();      // patch effettuata sulla TStringList FileHTML
                    if nPatch > 0 then begin

                        if cbxFileOverwrite.checked then
                            fToSave := fToFind
                        else
                            fToSave := vIOPdir + PCode + ' copia.html';

                        tPath := ExtractFilePath( fToSave );    // or GetDirectoryName
                        if DirectoryExists(tPath) then begin
                            SaveFile( fToSave );
                        end
                        else begin
                            // Ora distinguo fra le Opzioni scelte:
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

                    end//PatchFile
                    else
                        mmLog.Lines.Add('Nessuna sostituzione effettuata in '+ PCode +'.html' );

                end;//FindFile

                lviewIOP.Items.BeginUpdate;
                try
                  ListItem:= lviewIOP.Items.item[i];
                  //ListItem.Caption resta invariato, es. 'PRG.014'
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
                  lviewIOP.Items.EndUpdate;
                end;

            end;
        end;

    finally
      mmLog.Lines.Add(DatetimeToStr(now)+'   *** FINE Elaborazioni ***' +#13#10);
      Screen.Cursor := crDefault;
      btnStart.Enabled := True;
      rgArea.Enabled := True;
      btnSelectSource.Enabled := True;
      btnSelectTarget.Enabled := True;
      btnOpen1.Enabled := True;
      btnOpen2.Enabled := True;
      cbxFileOverwrite.Enabled := True;
      cbxForceDir.Enabled := True;
    end;
end;

procedure TForm1.SaveFile(filePathAndName:string);
begin
    try
      //FileHTML.SaveToFile( filePathAndName );
        FileHTML.SaveToFile( filePathAndName, TEncoding.UTF8 ); // altrimenti se l'origine non è in unicode, perde le accentate.
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
    mmLog.Lines.Add('Search ->  '+ aPathAndFileName);
    try
        if FindFirst(aPathAndFileName, FileAttrs, srec) = 0 then begin
            //if (srec.Attr and FileAttrs) = srec.Attr then // solo se necessita corrispondano anche attributi.
            FileHTML.LoadFromFile( aPathAndFileName );
            //strFileHTML.LoadFromFile( aPathAndFileName );
            mmLog.Lines.Add('Found  '+ srec.Name);      // IntToStr(srec.Size);
            mmLog.Lines.Add(#9+'with encoding = '+ FileHTML.Encoding.ToString);  //
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

    pbarSostituzioni.position := 1;
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
        pbarSostituzioni.position := pbarSostituzioni.position +1;
    end;//for

    result := pcount;
end;



procedure TForm1.BitBtn2Click(Sender: TObject);
begin
    //...
end;

procedure TForm1.BitBtn3Click(Sender: TObject);
begin
    mmLog.clear;
end;

procedure TForm1.btnLoadClick(Sender: TObject);
var
  ListItem: TListItem;
begin
    Screen.Cursor := crHourglass;
    try
      gammaConnection1.Connected := True;
      if not gammaConnection1.Connected then begin
          MessageDlg('DB Gamma NON disponibile !', mtError, [mbOk], 0, mbOk);
          exit;
      end;
      mmLog.Lines.Add('DB Gamma CONNESSO !');
      //MessageDlg('DB Gamma CONNESSO !', mtInformation, [mbOk], 0, mbOk)
      viewProdotti.Open;
      if viewProdotti.Active then begin

          try
            lviewProdotti.Clear;
            lviewProdotti.Items.BeginUpdate;
            mmLog.Lines.Add('Aperto DB Gamma vista Codici prodotto.');
            mmLog.Lines.Add('Trovati '+ viewProdotti.RecordCount.ToString +' codici prodotto.');

            viewProdotti.first;
            while not viewProdotti.EOF do begin  // loop su tutti i prodotto.

                // compila listview codici prodotto.
                ListItem := lviewProdotti.Items.Add;
                ListItem.Caption := viewProdotti.FieldByName('Articolo').AsString;  // init Codice
                with ListItem.SubItems do begin
                    Add('');                        // init Esito
                    Add('');                        // init Modifiche
                    Add(viewProdotti.FieldByName('Descrizione').AsString);  // init Descr
                    // if you need more columns, add here
                end;
                viewProdotti.next;
            end;//while
            btnStart.Enabled := True;
          finally
            lviewProdotti.Items.EndUpdate;
            viewProdotti.Close;
          end;
      end
      else
          MessageDlg('Vista Prodotti NON disponibile !', mtError, [mbOk], 0, mbOk);
    finally
      Screen.Cursor := crDefault;
      gammaConnection1.Connected := False;
      btnLoad.Enabled := True;
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
        //edtXlsSostituzioni.text := ExtractFileName(OpenDialog.FileName);
        edtXlsCodici.text := OpenDialog.FileName;
        // estrae codici da excel
        ReadExcelCodiciProd();
//      btnStart.Enabled := True;
    end;
end;
procedure TForm1.ReadExcelCodiciProd;
var
  xls: TXlsFile;
  row, colIndex: integer;
  rowCount: integer;
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
    for row := 1 to rowCount do begin
        colIndex := 1;  // to xls.ColCountInRow(row) do //Don't use xls.ColCount as it is slow: See http://www.tmssoftware.biz/flexcel/doc/vcl/guides/performance-guide.html#avoid-calling-colcount

        cell := xls.GetCellValue(row, colIndex);
        addr := TCellAddress.Create(row, xls.ColFromIndex(row, colIndex));

        // compila listview codici prodotto.
        ListItem:= lviewProdotti.Items.Add;
        if cell.IsString then begin
            ListItem.Caption := cell.ToString;  // init Codice
            with ListItem.SubItems do begin
                Add('');                        // init Esito
                Add('');                        // init Modifiche
                Add('');                        // init Path
                // if you need more columns, add here
            end;
        end
        else begin
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
    pbarSostituzioni.Max := rowCount;

    if cbxSkip1stLine.checked then
        from := 2
    else
        from := 1;
    for row := from to rowCount do begin

        // compila listview codici prodotto.
        ListItem:= lviewSostituzioni.Items.Add;

        colIndex := 1;
        cell := xls.GetCellValue(row, colIndex);
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
        cell := xls.GetCellValue(row, colIndex);
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
    ShowMessage('OK, Log saved in TimeStamp file.');
end;


procedure TForm1.btnSelectSourceClick(Sender: TObject);
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
