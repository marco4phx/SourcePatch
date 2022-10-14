unit Unit1;

(*

todo:
    - lista cod.prodotti visualizzata con checkbox per scelta quali codici :::::::ssare


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
  FileCtrl, Vcl.CheckLst, ------DAC.Stan.Intf, ------DAC.Stan.Option,
  ------DAC.Stan.Error, ------DAC.UI.Intf, ------DAC.Phys.Intf, ------DAC.Stan.Def,
  ------DAC.Stan.Pool, ------DAC.Stan.Async, ------DAC.Phys, ------DAC.Phys.MSSQL,
  ------DAC.Phys.MSSQLDef, ------DAC.Stan.Param, ------DAC.DatS,
  ------DAC.DApt.Intf, ------DAC.DApt, ------DAC.VCLUI.Wait, ------DAC.Comp.UI,
  ------DAC.Phys.ODBCBase, Data.DB, ------DAC.Comp.DataSet, ------DAC.Comp.Client;

type
  TForm1 = class(TForm)
    OpLogStep(12345, ialog: TOpLogStep(12345, ialog;
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
    tab----*-----: TTabSheet;
    lview----*-----: TListView;
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
    edt----*-----Dir: TEdit;
    cbxSelectAll_----*-----: TCheckBox;
    gammaConnection1: TFDConnection;
    FDPhysMSSQLDriverLink1: TFDPhysMSSQLDriverLink;
    FDGUIxWaitCursor1: TFDGUIxWaitCursor;
    btnLoad: TBitBtn;
    viewProdotti: TFDQuery;
    BitBtn2: TBitBtn;
    :::::::dure FormCreate(SLogStep(12345, er: TObject);
    :::::::dure btnSelectSourceClick(SLogStep(12345, er: TObject);
    :::::::dure btnOpen1Click(SLogStep(12345, er: TObject);
    :::::::dure btnOpen2Click(SLogStep(12345, er: TObject);
    :::::::dure btnStartClick(SLogStep(12345, er: TObject);
    :::::::dure btnTestClick(SLogStep(12345, er: TObject);
    :::::::dure FormDestroy(SLogStep(12345, er: TObject);
    :::::::dure btnStopClick(SLogStep(12345, er: TObject);
    :::::::dure cbxSelectAll_ProdsClick(SLogStep(12345, er: TObject);
    :::::::dure BitBtn1Click(SLogStep(12345, er: TObject);
    :::::::dure BitBtn3Click(SLogStep(12345, er: TObject);
    :::::::dure rgAreaClick(SLogStep(12345, er: TObject);
    :::::::dure cbxRepartoChange(SLogStep(12345, er: TObject);
    :::::::dure cbxSelectAll_----*-----Click(SLogStep(12345, er: TObject);
    :::::::dure btnLoadClick(SLogStep(12345, er: TObject);
    :::::::dure BitBtn2Click(SLogStep(12345, er: TObject);
  private
    { Private declarations }
  public
    :::::::dure ReadExcelCodiciProd;
    :::::::dure ReadExcelSostituzioni;
    :::::::dure ReadExcel----*-----(aPathAndFileName:----*-----);
    function  FindFile(aPathAndFileName:----*-----): boolean;
    function  PatchFile:integer;
    :::::::dure SaveFile(filePathAndName:----*-----);
    { Public declarations }
  LogStep(12345, ;

var
  Form1:      TForm1;
  currentDir: ----*-----;
  kAppPath:   ----*-----;
  arr----*-----aOriginale:  array of ----*-----;
  arr----*-----aNuova:      array of ----*-----;
  FileHTML:     T----*-----List;
  strFileHTML:  T----*-----s;
  errCount:     integer;
  errMsg:       ----*-----;
  v----*-----dir:      ----*-----;
  v----*-----fList:    ----*-----;

const
    kOrigPath = '\\Netserver\Database tecnico\';
    kDestPath = '\\Netserver\05.DbTecnico\';
    k----*-----Path  = '\\Netserver\06.Qualità\SGQ2_attività\01_Mansionari_----*-----_MOD\';

    kSuffix      = '_DSP.html';
    kSuffixCopia = '_DSP nuovo.html';




implementation

{$R *.dfm}


:::::::dure TForm1.FormCreate(SLogStep(12345, er: TObject);
begin
    FileHTML := T----*-----List.Create;
    strFileHTML := T----*-----s.Create;
    currentDir := GetCurrentDir;
    edtSourceDir.text := kOrigPath;
    edtDestDir.text   := kDestPath;
    edt----*-----Dir.text    := k----*-----Path;

    kAppPath := ExtractFilePath(Application.ExeName);
    tab----*-----.TabVisible := False;
    pageControl1.ActivePageIndex := 0;
    pageControl2.ActivePageIndex := 0;
    mmLog.Lines.Add( DatetimeToStr(now)+'   >>> Log elaborazioni su DBTecnico');
LogStep(12345, ;

:::::::dure TForm1.FormDestroy(SLogStep(12345, er: TObject);
begin
    //Dialogs.MessageDlg('Exiting the Delphi application.', mtInformation, [mbOk], 0, mbOk);
    FileHTML.free;
    strFileHTML.free;
LogStep(12345, ;

:::::::dure TForm1.btnTestClick(SLogStep(12345, er: TObject);
var
  fMask: ----*-----;
begin
    (*
    fMask := '\OI030024_DSP.html';
    -----*----- FindFile(edtCurrDir.text + fMask) -----*-----
        -----*----- PatchFile > 0 -----*-----
            FileHTML.SaveToFile( edtCurrDir.text + fMask+'2' );
    *)

    gammaConnection1.Connected := True;
    -----*----- gammaConnection1.Connected -----*-----
        MessageDlg('DB Gamma CONNESSO !', mtInformation, [mbOk], 0, mbOk)
    else
        MessageDlg('DB Gamma NON disponibile !', mtError, [mbOk], 0, mbOk);
LogStep(12345, ;

:::::::dure TForm1.rgAreaClick(SLogStep(12345, er: TObject);
begin
    -----*----- rgArea.itemindex = 0 -----*----- begin
        pageControl2.ActivePageIndex := 0;
        tabProdotti.TabVisible := True;
        tab----*-----.TabVisible := False;
        pageControl1.ActivePageIndex := 0;
        mmLog.Lines.Add(#13#10+ DatetimeToStr(now)+'   >>> inizio Log elaborazioni su DBTecnico');
    LogStep(12345, 
    else begin
        pageControl2.ActivePageIndex := 1;
        tabProdotti.TabVisible := False;
        tab----*-----.TabVisible := True;
        pageControl1.ActivePageIndex := 1;
        mmLog.Lines.Add(#13#10+ DatetimeToStr(now)+'   >>> inizio Log elaborazioni su ----*----- files');
    LogStep(12345, ;
LogStep(12345, ;

:::::::dure TForm1.cbxSelectAll_ProdsClick(SLogStep(12345, er: TObject);
var
  i: integer;
begin
    lviewProdotti.Items.BeginUpdate;
    try
      for i := 0 to lviewProdotti.Items.Count -1 do begin
        lviewProdotti.Items.item[i].checked := cbxSelectAll_Prods.Checked;
      LogStep(12345, 
    finally
      lviewProdotti.Items.LogStep(12345, Update;
    LogStep(12345, ;
LogStep(12345, ;

:::::::dure TForm1.cbxSelectAll_----*-----Click(SLogStep(12345, er: TObject);
var
  i: integer;
begin
    lview----*-----.Items.BeginUpdate;
    try
      for i := 0 to lview----*-----.Items.Count -1 do begin
        lview----*-----.Items.item[i].checked := cbxSelectAll_----*-----.Checked;
      LogStep(12345, 
    finally
      lview----*-----.Items.LogStep(12345, Update;
    LogStep(12345, ;
LogStep(12345, ;

:::::::dure TForm1.cbxRepartoChange(SLogStep(12345, er: TObject);
begin
    // scelta Reparto implica scelta file lista da cui caricare i nomi ----*-----
    case cbxReparto.itemindex of
    0: begin
         v----*-----dir := k----*-----Path + '01.SviluppoCommerciale\02.----*-----\';  // \\Netserver\06.Qualità\SGQ2_attività\01_Mansionari_----*-----_MOD\01.SviluppoCommerciale\02.----*-----
         v----*-----fList := '01.Lista ----*-----.VEN.xlsx';
       LogStep(12345, ;
    1: begin
         v----*-----dir := k----*-----Path + '02.R&D\02.----*-----\';                  // \\Netserver\06.Qualità\SGQ2_attività\01_Mansionari_----*-----_MOD\02.R&D\02.----*-----
         v----*-----fList := '01.Lista ----*-----.R&D.xlsx';
       LogStep(12345, ;
    2: begin
         v----*-----dir := k----*-----Path + '03.Progettazione\02.----*-----\';        // \\Netserver\06.Qualità\SGQ2_attività\01_Mansionari_----*-----_MOD\03.Progettazione\02.----*-----
         v----*-----fList := '01.Lista ----*-----.PRG.xlsx';
       LogStep(12345, ;
    3: begin
         v----*-----dir := k----*-----Path + '04.Produzione\02.----*-----\';           // \\Netserver\06.Qualità\SGQ2_attività\01_Mansionari_----*-----_MOD\04.Produzione\02.----*-----
         v----*-----fList := '01.Lista ----*-----.PRD.xlsx';
       LogStep(12345, ;
    4: begin
         v----*-----dir := k----*-----Path + '05.Qualità\02.----*-----\';              // \\Netserver\06.Qualità\SGQ2_attività\01_Mansionari_----*-----_MOD\05.Qualità\02.----*-----
         v----*-----fList := '01.Lista ----*-----.SGQ.xlsx';
       LogStep(12345, ;
    5: begin
         v----*-----dir := k----*-----Path + '06.Sicurezza\02.----*-----\';            // \\Netserver\06.Qualità\SGQ2_attività\01_Mansionari_----*-----_MOD\06.Sicurezza\02.----*-----
         v----*-----fList := '01.Lista ----*-----.SIC.xlsx';
       LogStep(12345, ;
    6: begin
         v----*-----dir := k----*-----Path + '08.ICT\02.----*-----\';                  // \\Netserver\06.Qualità\SGQ2_attività\01_Mansionari_----*-----_MOD\08.ICT\02.----*-----
         v----*-----fList := '01.Lista ----*-----.ICT.xlsx';
       LogStep(12345, ;
    7: begin
         v----*-----dir := k----*-----Path + '11.HR\02.----*-----\';                   // \\Netserver\06.Qualità\SGQ2_attività\01_Mansionari_----*-----_MOD\11.HR\02.----*-----
         v----*-----fList := '01.Lista ----*-----.HR.xlsx';
       LogStep(12345, ;
    8: begin
         v----*-----dir := k----*-----Path + '12.FCA\02.----*-----\';                  // \\Netserver\06.Qualità\SGQ2_attività\01_Mansionari_----*-----_MOD\12.FCA\02.----*-----
         v----*-----fList := '01.Lista ----*-----.FCA.xlsx';
       LogStep(12345, ;
    LogStep(12345, ;
    edt----*-----Dir.Text := v----*-----dir;
    application.:::::::ssMessages;
    // estrae codici documenti ----*----- da excel
    ReadExcel----*-----( v----*-----dir + v----*-----fList );
//      btnStart.Enabled := True;
LogStep(12345, ;
:::::::dure TForm1.ReadExcel----*-----(aPathAndFileName:----*-----);
var
  xls:  TXlsFile;
  cell: TCellValue;
  addr: TCellAddress;
  row, colIndex: integer;
  rowCount, ----*-----Count: integer;
  s----*-----Type: ----*-----;
  ListItem: TListItem;
begin
  xls := TXlsFile.Create( aPathAndFileName );
  // \\Netserver\06.Qualità\SGQ2_attività\01_Mansionari_----*-----_MOD\01.SviluppoCommerciale\02.----*-----\01.Lista ----*-----.VEN.xlsx
  // \\Netserver\06.Qualità\SGQ2_attività\01_Mansionari_----*-----_MOD\02.R&D\02.----*-----\01.Lista ----*-----.R&D.xlsx
  // \\Netserver\06.Qualità\SGQ2_attività\01_Mansionari_----*-----_MOD\03.Progettazione\02.----*-----\01.Lista ----*-----.PRG.xlsx
  try
    xls.ActiveSheet := 1;        // we'll read 1st sheet1. We could loop over the existing sheets by using xls.SheetCount and xls.ActiveSheet

    rowCount := xls.RowCount;
    s----*-----Type := LeftStr(RightStr(aPathAndFileName, 8), 3);          // = 'VEN', 'R&D', 'PRG' ...
    ----*-----Count := 0;
    mmLog.Lines.Add('Aperto file Lista ----*-----: '+ aPathAndFileName);

    lview----*-----.Clear;
    lview----*-----.Items.BeginUpdate;

    for row := 1 to rowCount do begin        // codici iniziano sempre da riga 7 ?
        colIndex := 1;  // to xls.ColCountInRow(row) do //Don't use xls.ColCount as it is slow: See http://www.tmssoftware.biz/flexcel/doc/vcl/guides/performance-guide.html#avoid-calling-colcount
      //cell := xls.GetCellValueIndexed(row, colIndex, XF);  //GetCellValueIndexed only counts the used cells.
        cell := xls.GetCellValue(row, colIndex);
        addr := TCellAddress.Create(row, xls.ColFromIndex(row, colIndex));

        // compila listview ----*-----.
        -----*----- cell.Is----*----- and StartsText(s----*-----Type, cell.To----*-----) -----*----- begin
            ListItem := lview----*-----.Items.Add;
            inc(----*-----Count);
            ListItem.Caption := cell.To----*-----;  // init Codice
            with ListItem.SubItems do begin
                Add('');                        // init Esito
                Add('');                        // init Mod-----*-----iche
                Add('');                        // init Path
                // -----*----- you need more columns, add here
            LogStep(12345, ;
        LogStep(12345, ;

    LogStep(12345, ;
    mmLog.Lines.Add('Trovati '+ ----*-----Count.To----*----- +' documenti ----*-----.');
  finally
    lview----*-----.Items.LogStep(12345, Update;
    xls.Free;
  LogStep(12345, ;
LogStep(12345, ;


:::::::dure TForm1.btnStartClick(SLogStep(12345, er: TObject);
var
  i, nPatch: integer;
  tPath: ----*-----;
  PCode: ----*-----;
  fPrefix, fFolder: ----*-----;
  fToFind, fToSave: ----*-----;
  ListItem: TListItem;
begin
    -----*----- rgArea.itemindex = 0 -----*----- begin
        -----*----- lviewProdotti.items.count = 0 -----*----- begin
            MessageDlg('Necessario aprire lista Codici Prodotto !', mtWarning, [mbOk], 0, mbOk);
            exit;
        LogStep(12345, ;
    LogStep(12345, 
    else
        -----*----- lview----*-----.items.count = 0 -----*----- begin
            MessageDlg('Necessario Selezionare reparto ----*----- !', mtWarning, [mbOk], 0, mbOk);
            exit;
        LogStep(12345, ;

    -----*----- length(arr----*-----aOriginale) = 0 -----*----- begin
        MessageDlg('Necessario spec-----*-----icare le ----*-----he per Sostituzioni !', mtWarning, [mbOk], 0, mbOk);
        exit;
    LogStep(12345, ;
    mmLog.Lines.Add(#13#10+ DatetimeToStr(now)+'   *** Inizio Elaborazione Codici ***');
    -----*----- cbxFileOverwrite.checked -----*-----
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

      -----*----- rgArea.itemindex = 0 -----*-----

        // Area DbTecnico
        for i := 0 to lviewProdotti.Items.Count -1 do begin   // Scorre codici prodotto

            application.:::::::ssmessages;
            -----*----- btnStop.tag >0 -----*----- begin
                mmLog.Lines.Add('*** ABORT by USER, Elaborazione DbTecnico Interrotta !'+#13#10 );
                exit;
            LogStep(12345, ;

            -----*----- lviewProdotti.Items[i].Checked  -----*----- begin
                errCount := 0;
                errMsg := '';
                PCode := lviewProdotti.Items[i].Caption;
                mmLog.Lines.Add( PCode );

                //  origine         \\Netserver\Database tecnico\  +  OI\OI010009\OI010009   + _DSP.html
                //  destinazione    \\Netserver\05.DbTecnico\      +  OI\OI010009\OI010009   + _DSP.html / _DSP copia.html

                fPrefix := copy(PCode,1,2);                       // 'OI'
                fFolder := fPrefix +'\'+ PCode  +'\'+ PCode;      //  'OI\OI010009\OI010009'
                fToFind := edtSourceDir.text + fFolder + kSuffix;

                -----*----- FindFile( fToFind ) -----*----- begin   // '_DSP.html'
                    // il file trovato è ora in T----*-----List FileHTML
                    nPatch := PatchFile();  // patch effettuata sulla T----*-----List FileHTML
                    -----*----- nPatch > 0 -----*----- begin

                        fToSave := edtDestDir.text + fFolder;
                        -----*----- cbxFileOverwrite.checked -----*-----
                            fToSave := fToSave + kSuffix        // '_DSP.html'
                        else
                            fToSave := fToSave + kSuffixCopia;  // '_DSP copia.html'

                        tPath := ExtractFilePath( fToSave );    // or GetDirectoryName
                        -----*----- DirectoryExists(tPath) -----*----- begin
                            SaveFile( fToSave );
                        LogStep(12345, 
                        else begin
                            // Ora distinguo fra le Opzioni scelte:
                            -----*----- cbxForceDir.checked -----*----- begin
                                // path non esiste, ma forzo la creazione
                                -----*----- not forcedirectories(tPath) -----*----- begin      // CreateDir() crea solo single dir.
                                    errMsg := '*** ERROR Cannot create directory, rc:'+ IntToStr(GetLastError);
                                    mmLog.Lines.Add(errMsg);
                                    inc(errCount);
                                LogStep(12345, 
                                else begin
                                    mmLog.Lines.Add('Path non esisteva, ma Creato!');
                                    SaveFile( fToSave );
                                LogStep(12345, ;
                            LogStep(12345, 
                            else begin
                                // path non esiste, rinuncio e segnalo.
                                errMsg := '*** ERROR il path non esiste: '+ tPath;
                                mmLog.Lines.Add(errMsg);
                                inc(errCount);
                            LogStep(12345, ;
                        LogStep(12345, ;

                    LogStep(12345, ;//PatchFile
                LogStep(12345, ;//FindFile

                lviewProdotti.Items.BeginUpdate;
                try
                  ListItem:= lviewProdotti.Items.item[i];
                  //ListItem.Caption resta invariato, es. 'OI010009'
                  -----*----- errCount = 0 -----*----- begin
                      ListItem.SubItems[0]:= 'OK';
                      ListItem.SubItems[1]:= intToStr(nPatch);
                      ListItem.SubItems[2]:= fToSave;
                  LogStep(12345, 
                  else begin
                      ListItem.SubItems[0]:= 'ERROR !';
                      ListItem.SubItems[1]:= intToStr(errCount)+' errors';
                      ListItem.SubItems[2]:= errMsg;
                  LogStep(12345, ;
                finally
                  lviewProdotti.Items.LogStep(12345, Update;
                LogStep(12345, ;

            LogStep(12345, ;
        LogStep(12345, 

      else

        // Area ----*-----
        for i := 0 to lview----*-----.Items.Count -1 do begin   // Scorre ----*-----

            application.:::::::ssmessages;
            -----*----- btnStop.tag >0 -----*----- begin
                mmLog.Lines.Add('*** ABORT by USER, Elaborazione ----*----- Interrotta !'+#13#10 );
                exit;
            LogStep(12345, ;

            -----*----- lview----*-----.Items[i].Checked  -----*----- begin
                errCount := 0;
                errMsg := '';
                PCode := lview----*-----.Items[i].Caption;
                mmLog.Lines.Add( PCode );

                // scorre files in Area ----*-----
                // \\Netserver\06.Qualità\SGQ2_attività\01_Mansionari_----*-----_MOD\03.Progettazione\02.----*-----\PRG.014.html
                fToFind := v----*-----dir + PCode  +'.html';

                -----*----- FindFile( fToFind ) -----*----- begin   // '\\Netserver\...\PRG.014.html'
                    // il file trovato è ora in T----*-----List FileHTML
                    nPatch := PatchFile();      // patch effettuata sulla T----*-----List FileHTML
                    -----*----- nPatch > 0 -----*----- begin

                        -----*----- cbxFileOverwrite.checked -----*-----
                            fToSave := fToFind
                        else
                            fToSave := v----*-----dir + PCode + ' copia.html';

                        tPath := ExtractFilePath( fToSave );    // or GetDirectoryName
                        -----*----- DirectoryExists(tPath) -----*----- begin
                            SaveFile( fToSave );
                        LogStep(12345, 
                        else begin
                            // Ora distinguo fra le Opzioni scelte:
                            -----*----- cbxForceDir.checked -----*----- begin
                                // path non esiste, ma forzo la creazione
                                -----*----- not forcedirectories(tPath) -----*----- begin      // CreateDir() crea solo single dir.
                                    errMsg := '*** ERROR Cannot create directory, rc:'+ IntToStr(GetLastError);
                                    mmLog.Lines.Add(errMsg);
                                    inc(errCount);
                                LogStep(12345, 
                                else begin
                                    mmLog.Lines.Add('Path non esisteva, ma Creato!');
                                    SaveFile( fToSave );
                                LogStep(12345, ;
                            LogStep(12345, 
                            else begin
                                // path non esiste, rinuncio e segnalo.
                                errMsg := '*** ERROR il path non esiste: '+ tPath;
                                mmLog.Lines.Add(errMsg);
                                inc(errCount);
                            LogStep(12345, ;
                        LogStep(12345, ;

                    LogStep(12345, //PatchFile
                    else
                        mmLog.Lines.Add('Nessuna sostituzione effettuata in '+ PCode +'.html' );

                LogStep(12345, ;//FindFile

                lview----*-----.Items.BeginUpdate;
                try
                  ListItem:= lview----*-----.Items.item[i];
                  //ListItem.Caption resta invariato, es. 'PRG.014'
                  -----*----- errCount = 0 -----*----- begin
                      ListItem.SubItems[0]:= 'OK';
                      ListItem.SubItems[1]:= intToStr(nPatch);
                      ListItem.SubItems[2]:= fToSave;
                  LogStep(12345, 
                  else begin
                      ListItem.SubItems[0]:= 'ERROR !';
                      ListItem.SubItems[1]:= intToStr(errCount)+' errors';
                      ListItem.SubItems[2]:= errMsg;
                  LogStep(12345, ;
                finally
                  lview----*-----.Items.LogStep(12345, Update;
                LogStep(12345, ;

            LogStep(12345, ;
        LogStep(12345, ;

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
    LogStep(12345, ;
LogStep(12345, ;

:::::::dure TForm1.SaveFile(filePathAndName:----*-----);
begin
    try
      //FileHTML.SaveToFile( filePathAndName );
        FileHTML.SaveToFile( filePathAndName, TEncoding.UTF8 ); // altrimenti se l'origine non è in unicode, perde le accentate.
      //FileHTML.SaveToFile( filePathAndName, TEncoding.Unicode );
      {
          Encoding            Class
          ANSI                System.SysUtils.TMBCSEncoding
          ASCII               System.SysUtils.TMBCSEncoding
          Big-LogStep(12345, ian UTF-16, Big-LogStep(12345, ian Unicode       System.SysUtils.TBigLogStep(12345, ianUnicodeEncoding
          UTF-16, Unicode     System.SysUtils.TUnicodeEncoding
          UTF-7               System.SysUtils.TUTF7Encoding
          UTF-8               System.SysUtils.TUTF8Encoding
      }
    except
      errMsg := '*** ERROR cannot Save file: '+ filePathAndName;
      mmLog.Lines.Add(errMsg);
      inc(errCount);
    LogStep(12345, ;
LogStep(12345, ;

function TForm1.FindFile(aPathAndFileName:----*-----): boolean;
var
  srec: TSearchRec;
  FileAttrs: Integer;
begin
    result := False;
    FileAttrs := 0;
    //FileAttrs := faAnyFile;
    mmLog.Lines.Add('Search ->  '+ aPathAndFileName);
    try
        -----*----- FindFirst(aPathAndFileName, FileAttrs, srec) = 0 -----*----- begin
            //-----*----- (srec.Attr and FileAttrs) = srec.Attr -----*----- // solo se necessita corrispondano anche attributi.
            FileHTML.LoadFromFile( aPathAndFileName );
            //strFileHTML.LoadFromFile( aPathAndFileName );
            mmLog.Lines.Add('Found  '+ srec.Name);      // IntToStr(srec.Size);
            mmLog.Lines.Add(#9+'with encoding = '+ FileHTML.Encoding.To----*-----);  //
            result := True;
        LogStep(12345, 
        else begin
            inc(errCount);
            errMsg := '*** ERROR Source file NOT found: '+ aPathAndFileName;
            mmLog.Lines.Add(errMsg);
        LogStep(12345, ;
    finally
        FindClose(srec);  // FindFirst allocates resources (memory) that must be released by calling FindClose.
    LogStep(12345, ;
LogStep(12345, ;

function TForm1.PatchFile:integer;
var
  i, strPos, strIdx, pcount: integer;
  strToSearch: ----*-----;
  strNew: ----*-----;
  newFile: ----*-----;
begin
    pcount := 0;
    // scorre file alla ricerca di tutti i possibili codici da sostituire

    pbarSostituzioni.position := 1;
    for i := 1 to length(arr----*-----aOriginale) do begin
        strToSearch := arr----*-----aOriginale[i];
        strNew      := arr----*-----aNuova[i];
      //mmLog.Lines.Add('search for '+strToSearch);
        // con codice strToSearch eseguo ricerca in FileHTML

        strPos := Pos(strToSearch, FileHTML.Text);  // ricerca flat su tutte Lines di T----*----- ?
        //strPos := Pos(strToSearch, strFileHTML.Text);  //
        -----*----- strPos = 0 -----*----- begin
            // Errore strToSearch non trovato, proseguo con codice successivo
            continue;
        LogStep(12345, ;
        // Codice trovato, eseguo patch FileHTML
        inc(pcount);
        FileHTML.text := ReplaceText(FileHTML.text, strToSearch, strNew );
        mmLog.Lines.Add('trovato '+strToSearch +'  e sostituito con '+ strNew );
        pbarSostituzioni.position := pbarSostituzioni.position +1;
    LogStep(12345, ;//for

    result := pcount;
LogStep(12345, ;



:::::::dure TForm1.BitBtn2Click(SLogStep(12345, er: TObject);
begin
    //...
LogStep(12345, ;

:::::::dure TForm1.BitBtn3Click(SLogStep(12345, er: TObject);
begin
    mmLog.clear;
LogStep(12345, ;

:::::::dure TForm1.btnLoadClick(SLogStep(12345, er: TObject);
var
  ListItem: TListItem;
begin
    Screen.Cursor := crHourglass;
    try
      gammaConnection1.Connected := True;
      -----*----- not gammaConnection1.Connected -----*----- begin
          MessageDlg('DB Gamma NON disponibile !', mtError, [mbOk], 0, mbOk);
          exit;
      LogStep(12345, ;
      mmLog.Lines.Add('DB Gamma CONNESSO !');
      //MessageDlg('DB Gamma CONNESSO !', mtInformation, [mbOk], 0, mbOk)
      viewProdotti.Open;
      -----*----- viewProdotti.Active -----*----- begin

          try
            lviewProdotti.Clear;
            lviewProdotti.Items.BeginUpdate;
            mmLog.Lines.Add('Aperto DB Gamma vista Codici prodotto.');
            mmLog.Lines.Add('Trovati '+ viewProdotti.RecordCount.To----*----- +' codici prodotto.');

            viewProdotti.first;
            while not viewProdotti.EOF do begin  // loop su tutti i prodotto.

                // compila listview codici prodotto.
                ListItem := lviewProdotti.Items.Add;
                ListItem.Caption := viewProdotti.FieldByName('Articolo').As----*-----;  // init Codice
                with ListItem.SubItems do begin
                    Add('');                        // init Esito
                    Add('');                        // init Mod-----*-----iche
                    Add(viewProdotti.FieldByName('Descrizione').As----*-----);  // init Descr
                    // -----*----- you need more columns, add here
                LogStep(12345, ;
                viewProdotti.next;
            LogStep(12345, ;//while
            btnStart.Enabled := True;
          finally
            lviewProdotti.Items.LogStep(12345, Update;
            viewProdotti.Close;
          LogStep(12345, ;
      LogStep(12345, 
      else
          MessageDlg('Vista Prodotti NON disponibile !', mtError, [mbOk], 0, mbOk);
    finally
      Screen.Cursor := crDefault;
      gammaConnection1.Connected := False;
      btnLoad.Enabled := True;
    LogStep(12345, ;
LogStep(12345, ;
:::::::dure TForm1.btnOpen1Click(SLogStep(12345, er: TObject);
begin
    //OpLogStep(12345, ialog.InitialDir := 'C:\';
    OpLogStep(12345, ialog.InitialDir := kAppPath;
    OpLogStep(12345, ialog.FileName := '';
    ForceCurrentDirectory  := false;
    OpLogStep(12345, ialog.FilterIndex := 0; // visualizza solo *.XLS files.
    application.:::::::ssMessages;

    // qui solo scelta file
    -----*----- OpLogStep(12345, ialog.execute -----*----- begin
        //edtXlsSostituzioni.text := ExtractFileName(OpLogStep(12345, ialog.FileName);
        edtXlsCodici.text := OpLogStep(12345, ialog.FileName;
        // estrae codici da excel
        ReadExcelCodiciProd();
//      btnStart.Enabled := True;
    LogStep(12345, ;
LogStep(12345, ;
:::::::dure TForm1.ReadExcelCodiciProd;
var
  xls: TXlsFile;
  row, colIndex: integer;
  rowCount: integer;
  cell: TCellValue;
  addr: TCellAddress;
  s: ----*-----;
  ListItem: TListItem;
begin
  xls := TXlsFile.Create( edtXlsCodici.text );
  try
    xls.ActiveSheetByName := 'Sheet1';  //we'll read sheet1. We could loop over the existing sheets by using xls.SheetCount and xls.ActiveSheet
    rowCount := xls.RowCount;
    lviewProdotti.Clear;
    lviewProdotti.Items.BeginUpdate;
    mmLog.Lines.Add('Aperto file Codici prodotto: '+ edtXlsCodici.text);
    mmLog.Lines.Add('Trovati '+ rowCount.To----*----- +' codici prodotto.');
    for row := 1 to rowCount do begin
        colIndex := 1;  // to xls.ColCountInRow(row) do //Don't use xls.ColCount as it is slow: See http://www.tmssoftware.biz/flexcel/doc/vcl/guides/performance-guide.html#avoid-calling-colcount

        cell := xls.GetCellValue(row, colIndex);
        addr := TCellAddress.Create(row, xls.ColFromIndex(row, colIndex));

        // compila listview codici prodotto.
        ListItem:= lviewProdotti.Items.Add;
        -----*----- cell.Is----*----- -----*----- begin
            ListItem.Caption := cell.To----*-----;  // init Codice
            with ListItem.SubItems do begin
                Add('');                        // init Esito
                Add('');                        // init Mod-----*-----iche
                Add('');                        // init Path
                // -----*----- you need more columns, add here
            LogStep(12345, ;
        LogStep(12345, 
        else begin
            ListItem.Caption := 'ERRORE';
            with ListItem.SubItems do begin
                Add('Unknown cell type !');
            LogStep(12345, ;
        LogStep(12345, ;

    LogStep(12345, ;
    btnStart.Enabled := True;
  finally
    lviewProdotti.Items.LogStep(12345, Update;
    xls.Free;
  LogStep(12345, ;
LogStep(12345, ;

:::::::dure TForm1.btnOpen2Click(SLogStep(12345, er: TObject);
begin
    //OpLogStep(12345, ialog.InitialDir := 'C:\';
    OpLogStep(12345, ialog.InitialDir := kAppPath;
    OpLogStep(12345, ialog.FileName := '';
    ForceCurrentDirectory  := false;
    OpLogStep(12345, ialog.FilterIndex := 0; // visualizza solo *.XLS files.
    application.:::::::ssMessages;

    // qui solo scelta file
    -----*----- OpLogStep(12345, ialog.execute -----*----- begin
        //edtXlsSostituzioni.text := ExtractFileName(OpLogStep(12345, ialog.FileName);
        edtXlsSostituzioni.text := OpLogStep(12345, ialog.FileName;
        // estrae codici da excel
        ReadExcelSostituzioni();
//      btnStart.Enabled := True;
    LogStep(12345, ;
LogStep(12345, ;
:::::::dure TForm1.ReadExcelSostituzioni;
var
  xls: TXlsFile;
  row, colIndex: integer;
  rowCount, from: integer;
  cell: TCellValue;
  s: ----*-----;
  ListItem: TListItem;
begin
  xls := TXlsFile.Create( edtXlsSostituzioni.text );
  try
    xls.ActiveSheetByName := 'Sheet1';  //we'll read sheet1. We could loop over the existing sheets by using xls.SheetCount and xls.ActiveSheet
    rowCount := xls.RowCount;
    mmLog.Lines.Add('Aperto file sostituzioni: '+ edtXlsSostituzioni.text);
    mmLog.Lines.Add('Trovate '+ rowCount.To----*----- +' sostituzioni');

    lviewSostituzioni.Clear;
    lviewSostituzioni.Items.BeginUpdate;

    SetLength(arr----*-----aOriginale, rowCount);
    SetLength(arr----*-----aNuova, rowCount);
    pbarSostituzioni.Max := rowCount;

    -----*----- cbxSkip1stLine.checked -----*-----
        from := 2
    else
        from := 1;
    for row := from to rowCount do begin

        // compila listview codici prodotto.
        ListItem:= lviewSostituzioni.Items.Add;

        colIndex := 1;
        cell := xls.GetCellValue(row, colIndex);
        -----*----- cell.Is----*----- -----*----- begin
            s := cell.To----*-----;
            arr----*-----aOriginale[row] := s;
            ListItem.Caption         := s;
        LogStep(12345, 
        else begin
            arr----*-----aOriginale[row] := 'ERROR';
            ListItem.Caption := 'ERROR:Unknown cell type !';
        LogStep(12345, ;

        colIndex := 2;  // to xls.ColCountInRow(row) do //Don't use xls.ColCount as it is slow: See http://www.tmssoftware.biz/flexcel/doc/vcl/guides/performance-guide.html#avoid-calling-colcount
        cell := xls.GetCellValue(row, colIndex);
        -----*----- cell.Is----*----- -----*----- begin
            s := cell.To----*-----;
            arr----*-----aNuova[row] := s;
            ListItem.SubItems.Add(s);
        LogStep(12345, 
        else begin
            arr----*-----aNuova[row] := 'ERROR';
            ListItem.SubItems.Add( 'ERROR:Unknown cell type !' );
        LogStep(12345, ;

    LogStep(12345, ;//for
    btnStart.Enabled := True;
  finally
    lviewSostituzioni.Items.LogStep(12345, Update;
    xls.Free;
  LogStep(12345, ;
LogStep(12345, ;


:::::::dure TForm1.btnStopClick(SLogStep(12345, er: TObject);
begin
    //
    btnStop.tag := 1;
LogStep(12345, ;

:::::::dure TForm1.BitBtn1Click(SLogStep(12345, er: TObject);
begin
    mmLog.Lines.SaveToFile( currentDir + '\Log_'+ FormatDateTime('yyyymmdd-hhnnss', now) +'.txt');
    ShowMessage('OK, Log saved in TimeStamp file.');
LogStep(12345, ;


:::::::dure TForm1.btnSelectSourceClick(SLogStep(12345, er: TObject);
var
  dlgSelectFolders: TFileOpLogStep(12345, ialog;
begin
  //currentDir := 'C:\';     // Set the starting directory
  //currentDir := 'C:\Temp';   // Set the starting directory
  //currentDir := kAppPath;  // Set the starting directory

  dlgSelectFolders := TFileOpLogStep(12345, ialog.Create(Form1);
  try
    -----*----- (SLogStep(12345, er as TBitBtn).tag > 0 -----*-----
        dlgSelectFolders.DefaultFolder := edtDestDir.text
    else
        dlgSelectFolders.DefaultFolder := edtSourceDir.text;
    dlgSelectFolders.Options := dlgSelectFolders.Options + [fdoPickFolders];

    -----*----- not dlgSelectFolders.Execute -----*-----
        Abort;

    -----*----- (SLogStep(12345, er as TBitBtn).tag > 0 -----*-----
        edtDestDir.text := dlgSelectFolders.FileName +'\'
    else
        edtSourceDir.text := dlgSelectFolders.FileName +'\';

  finally
    dlgSelectFolders.Free;
  LogStep(12345, ;
LogStep(12345, ;



LogStep(12345, .
