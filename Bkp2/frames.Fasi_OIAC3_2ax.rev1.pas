unit frames.Fasi_OIAC3_2ax;
{$I Defines.inc}   // defines possibili anche in prj compiler options !

(*
    for btn in TArray<TButton>.Create(btn1, btn31,btn2) do
        btn.Enabled := false;

    if TRegEx.IsMatch(EditIP.Text, '^[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}$') then
        ShowMessage('IPv4 address is valid')

    ListBox1.Items.Add(Format('[%3.3d]', [Value]));

    // subscribe to a String message on the default message manager
    TMessageManager.DefaultManager.SubscribeToMessage(TMessage<String>,
        procedure(const Sender: TObject; const AMessage: TMessage)
        begin
            WriteLn('Called callback1 with value: ', TMessage<String>(AMessage).Value);
        end);
    // send a String message to the default message manager
    TMessageManager.DefaultManager.SendMessage(nil, TMessage<String>.Create('Hello Messaging'));
    // Remember that publishing a message is a synchronous process.
    // Your call to SendMessage will not return until each subscriber's MessageListener code is run in turn.


    procedure ExecWithLock(const ALockObj: TObject; const AProc: TProc);
    begin
        System.TMonitor.Enter(ALockObj);
        try
            AProc();
        finally
            System.TMonitor.Exit(ALockObj);
        end;
    end;
    //Using the preceding code, it is possible to write something like the following:
    ExecWithLock(Obj, procedure
    begin
        //Here you have thread safe access to Obj
    end);


    // Per gestione di redirect Function e procedure !
    https://stackoverflow.com/questions/14975415/what-is-the-difference-between-of-object-and-reference-to
    http://docwiki.embarcadero.com/RADStudio/Sydney/en/Procedural_Types_(Delphi)#About_Procedural_Types
    https://stackoverflow.com/questions/4186458/delphi-call-a-function-whose-name-is-stored-in-a-string

    per dinstinguere bene Steps e Fasi, che ripetono processi diversi, ma simili
    e che per questo usano la stessa GUI (TabSheet)
    devo poter usare lo stesso Tab per FASI diverse !
    Non è possibile collegare Tab vuoti a Tab esistenti (per evitare moltiplicazione resorse gui)
    quindi uso TTabControl + TCards  invece di classico TPageControl !

*)



interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,  System.StrUtils,
  System.Math, System.DateUtils, Generics.Collections, TypInfo, System.UITypes,
  SyncObjs,
  units.OptoiCAN,
  units.PCANBasic,
  units.Common,
  units.LocalKonst,
  units.Config,
  fraFasi_BASE,
  units.sortDictionary,
  Vcl.StdCtrls, Vcl.ComCtrls, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  VCL.FlexCel.Core, FlexCel.XlsAdapter,
  cxListView, dxSkinsCore, dxSkinsDefaultPainters,
  dxBarBuiltInMenu, cxGraphics, cxControls, cxLookAndFeels,
  cxLookAndFeelPainters, cxContainer, cxEdit, cxClasses,
  dxGaugeCustomScale, dxGaugeQuantitativeScale, dxGaugeCircularScale,
  dxGaugeControl, cxProgressBar, cxPC, Vcl.Menus,
  Vcl.PlatformDefaultStyleActnCtrls, Vcl.ActnPopup, System.Actions,
  Vcl.ActnList, System.ImageList, Vcl.ImgList, cxImageList,
  dxGDIPlusClasses, Vcl.ExtCtrls, cxTextEdit, cxMaskEdit, cxSpinEdit,
  cxMemo, cxRichEdit, dxCore, dxCoreClasses, dxGDIPlusAPI,
  dxRichEdit.NativeApi, dxRichEdit.Types, dxRichEdit.Options,
  dxRichEdit.Control, dxRichEdit.Control.SpellChecker,
  dxRichEdit.Dialogs.EventArgs, dxRichEdit.Platform.Win.Control,
  dxRichEdit.Control.Core, Vcl.Buttons, cxButtons, Vcl.WinXPanels,
  FlexCel.Render, FlexCel.Preview, cxGroupBox, cxRadioGroup, cxStyles,
  cxCustomData, cxFilter, cxData, cxDataStorage, cxNavigator, dxDateRanges,
  cxDataControllerConditionalFormattingRulesManagerDialog, Data.DB,
  cxDBData, cxGridLevel, cxGridCustomView, cxGridCustomTableView,
  cxGridTableView, cxGridDBTableView, cxGrid, DBAccess, Uni, cxCheckGroup,
  Vcl.MPlayer, cxDropDownEdit
  ;


type
  //TMyEvent = procedure(Sender:TObject) of object;
  //answer17 https://stackoverflow.com/questions/5786595/delphi-event-handling-how-to-create-own-event
  TMyEvent = TNotifyEvent;   // equivalente a precedente.

  TCalibAngle = ( X0_Y0, XN_Yxn, XP_Yxp, YP_Xyp, YN_Xyn );  // i valori precisi poi dipendono da tabella di calibrazione specifica per famiglia di devices.

  TTestAngle = ( deg60n, deg30n, deg15n, deg0, deg15p, deg30p, deg60p );  // su quale asse avviene il test dipende da altro parametro.

  TCorrBanco = class       // new class inherits directly from the predefined System.TObject class, in case you omit (ancestorClass).
    // per gestione coppie valori di Correzione Banco calibrazione.
    angolo: integer;
    asseTest: single;
    asseCross: single;
  end;
//TCorrezioniBanco = class(TDictionary<integer, TCorrBanco>);  // dictionary non utilizzabile perchè non è ordinato per key !
  TarrCorrezioniBanco = array [1..15] of TCorrBanco;           // quindi devo usare array, aggiungendo angolo alla class.

//STestRecords = class(TDictionary<TTestAngle, TSessionTests>);// no, uso sortDic perchè mi serve get per index.


  Toiac3Fasi2ax = class(TframeBase)
    popMenuAdmin: TPopupActionBar;
    menuInitCAN: TMenuItem;
    menuCloseCAN: TMenuItem;
    menuResetCAN: TMenuItem;
    menuCANgetStatus: TMenuItem;
    tmrCanRead: TTimer;
    PopMenuFaseC: TPopupMenu;
    SetupC1: TMenuItem;
    Step101: TMenuItem;
    Step111: TMenuItem;
    Step121: TMenuItem;
    Step131: TMenuItem;
    Step141: TMenuItem;
    Step151: TMenuItem;
    Step161: TMenuItem;
    N1: TMenuItem;
    PopMenuFaseB: TPopupMenu;
    MenuItem2: TMenuItem;
    MenuItem3: TMenuItem;
    MenuItem4: TMenuItem;
    MenuItem5: TMenuItem;
    MenuItem6: TMenuItem;
    MenuItem7: TMenuItem;
    MenuItem9: TMenuItem;
    Scarto1: TMenuItem;
    Scarto2: TMenuItem;
    Writecal1: TMenuItem;
    PopMenuFaseA: TPopupMenu;
    LoadDefaults1: TMenuItem;
    LetturaParametri1: TMenuItem;
    ScriviParametri1: TMenuItem;
    SalvaParametri1: TMenuItem;
    LoadCorrBanco1: TMenuItem;
    SetupA1: TMenuItem;
    popMenuFaseF: TPopupMenu;
    CreaReport30: TMenuItem;
    N4: TMenuItem;
    GenTestData: TMenuItem;
    lviewCanCMDs: TcxListView;
    btnWriteCAN: TButton;
    btnReadCAN: TButton;
    Button3: TButton;
    btnInitCAN: TButton;
    btnCloseCAN: TButton;
    btnWriteRaw: TButton;
    btnReadRaw: TButton;
    Label5: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label6: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label21: TLabel;
    edtCanID: TEdit;
    edtCanData: TEdit;
    edtSubIdx: TEdit;
    edtIndex: TEdit;
    edtRxCanID: TEdit;
    edtRxCanData: TEdit;
    edtRxSubIdx: TEdit;
    edtRxIndex: TEdit;
    edtCobID: TEdit;
    edtRawCanLen: TEdit;
    edtRawCanData: TEdit;
    edtCobIDrcvd: TEdit;
    edtRawCanLenRcvd: TEdit;
    edtRawCanDataRcvd: TEdit;
    Button4: TButton;
    Button5: TButton;
    Button6: TButton;
    Button7: TButton;
    Button8: TButton;
    Label20: TLabel;
    cbxDataType: TComboBox;
    Button1: TButton;
    tkBar2: TTrackBar;
    Label17: TLabel;
    lviewCFG: TcxListView;
    Label12: TLabel;
    edtToWrite: TEdit;
    cbxSelectAll: TCheckBox;
    N6: TMenuItem;
    TestsortDict1: TMenuItem;
    SaveDB1: TMenuItem;
    SavePDF1: TMenuItem;
    SaveHTML1: TMenuItem;
    SetupF: TMenuItem;
    SalvasuDB29: TMenuItem;
    panel_CD: TPanel;
    panel_B: TPanel;
    lviewCarattBanco: TcxListView;
    Label24: TLabel;
    Label25: TLabel;
    Label30: TLabel;
    Label32: TLabel;
    Label33: TLabel;
    edtScartoMax_test: TcxSpinEdit;
    edtErrMaxCross: TcxTextEdit;
    edtErrMaxTest: TcxTextEdit;
    Label26: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    edtScartoMax_calib: TcxSpinEdit;
    edtRefSamplesInRange: TcxSpinEdit;
    cxSpinEdit1: TcxSpinEdit;
    PopMenuFaseD: TPopupMenu;
    MenuItem1: TMenuItem;
    MenuItem8: TMenuItem;
    MenuItem12: TMenuItem;
    MenuItem13: TMenuItem;
    MenuItem14: TMenuItem;
    MenuItem15: TMenuItem;
    MenuItem16: TMenuItem;
    MenuItem17: TMenuItem;
    MenuItem18: TMenuItem;
    MenuItem22: TMenuItem;
    debugReadAngle1: TMenuItem;
    Label15: TLabel;
    edtResolREF2: TcxTextEdit;
    Label18: TLabel;
    edtResolDUT: TcxTextEdit;
    popMenuGauges: TPopupActionBar;
    StartREF2cyclicalRD: TMenuItem;
    StopREF2cyclicalRD: TMenuItem;
    MenuItem29: TMenuItem;
    MenuItem30: TMenuItem;
    MenuItem31: TMenuItem;
    N5: TMenuItem;
    ABORTStep1: TMenuItem;
    ResetAbortFlag1: TMenuItem;
    N7: TMenuItem;
    StopTIMERsamples1: TMenuItem;
    StartTIMERsamples1: TMenuItem;
    panel_Gauges: TPanel;
    Label16: TLabel;
    img1Seq1: TImage;
    Label27: TLabel;
    Label28: TLabel;
    lbDutMode: TLabel;
    Label31: TLabel;
    dxGaugeControl2: TdxGaugeControl;
    gaugeY: TdxGaugeCircularHalfScale;
    gaugeYLabel: TdxGaugeQuantitativeScaleCaption;
    dxGaugeControl1: TdxGaugeControl;
    gaugeX: TdxGaugeCircularHalfScale;
    gaugeXLabel: TdxGaugeQuantitativeScaleCaption;
    edtRefY: TEdit;
    edtRefX: TEdit;
    edtDutX: TEdit;
    edtDutY: TEdit;
    edtTargetX: TEdit;
    edtTargetY: TEdit;
    Button2: TButton;
    N2: TMenuItem;
    SaveREFParams1: TMenuItem;
    Button9: TButton;
    Button10: TButton;
    edtRxDataAsInt: TEdit;
    ReadallCalib1: TMenuItem;
    Label19: TLabel;
    cbxCanDataStrPrefix: TComboBox;
    lviewCalibVal: TcxListView;
    SetCalibintvalue1: TMenuItem;
    ClearLog1: TMenuItem;
    N3: TMenuItem;
    menuInvertiYref: TMenuItem;
    menuInvertiXref: TMenuItem;
    menuInvertiXdut: TMenuItem;
    menuInvertiYdut: TMenuItem;
    lbFondoScala: TLabel;
    pnlCardsF: TCardPanel;
    xlsPreview: TFlexCelPreviewer;
    cxButton3: TcxButton;
    Label4: TLabel;
    lbReportName: TLabel;
    cbxShowGrid: TCheckBox;
    cbxHeadings: TCheckBox;
    rgXlsZoom: TcxRadioGroup;
    cxButton2: TcxButton;
    btnSaveRtf: TcxButton;
    cxGrid1DBTableView2: TcxGridDBTableView;
    cxGrid1Level2: TcxGridLevel;
    cxGrid1: TcxGrid;
    N8: TMenuItem;
    viewPanel28: TMenuItem;
    viewPanel29: TMenuItem;
    Beeptest1: TMenuItem;
    popMenuFaseE: TPopupMenu;
    menuSetupE: TMenuItem;
    LetturaParametri25: TMenuItem;
    ScriviParametri26: TMenuItem;
    SalvaParametri27: TMenuItem;
    Label29: TLabel;
    cxGrid1Level3: TcxGridLevel;
    cxGrid1DBTableView3: TcxGridDBTableView;
    cxGrid1DBTableView2ASSEINTEST: TcxGridDBColumn;
    cxGrid1DBTableView2DEGPAIRID: TcxGridDBColumn;
    cxGrid1DBTableView2SCARTOAMMESSO: TcxGridDBColumn;
    cxGrid1DBTableView2TARGETX: TcxGridDBColumn;
    cxGrid1DBTableView2TARGETY: TcxGridDBColumn;
    cxGrid1DBTableView2REFSAMPLESINRANGE: TcxGridDBColumn;
    cxGrid1DBTableView2INTERVALLOREFREADS: TcxGridDBColumn;
    cxGrid1DBTableView2INTERVALLODUTREADS: TcxGridDBColumn;
    cxGrid1DBTableView2DUTSAMPLESSTABLE: TcxGridDBColumn;
    cxGrid1DBTableView2REFAX: TcxGridDBColumn;
    cxGrid1DBTableView2REFAY: TcxGridDBColumn;
    cxGrid1DBTableView2DUTAXdeg: TcxGridDBColumn;
    cxGrid1DBTableView2DUTAYdeg: TcxGridDBColumn;
    cxGrid1DBTableView2DUTAXmv: TcxGridDBColumn;
    cxGrid1DBTableView2DUTAYmv: TcxGridDBColumn;
    cxGrid1DBTableView2CORRECTREFTEST: TcxGridDBColumn;
    cxGrid1DBTableView2CORRECTREFCROSS: TcxGridDBColumn;
    cxGrid1DBTableView2ERRORASSETEST: TcxGridDBColumn;
    cxGrid1DBTableView2ERRORASSECROSS: TcxGridDBColumn;
    cxGrid1DBTableView3INDEX: TcxGridDBColumn;
    cxGrid1DBTableView3SUBIDX: TcxGridDBColumn;
    cxGrid1DBTableView3ALIAS: TcxGridDBColumn;
    cxGrid1DBTableView3TYPE: TcxGridDBColumn;
    cxGrid1DBTableView3VALUE: TcxGridDBColumn;
    ResetTreeviewicons1: TMenuItem;
    rgXlsDir: TcxRadioGroup;
    ClearTSPairs1: TMenuItem;
    LoadListaVerifica24: TMenuItem;
    Label34: TLabel;
    Label35: TLabel;
    Label36: TLabel;
    panel_A: TPanel;
    Label11: TLabel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    edtRef2Addr: TEdit;
    edtRef1Addr: TEdit;
    grpBoxListaParams: TGroupBox;
    cbxListaDefaultsCom: TCheckBox;
    cbxListaCalib1ax: TCheckBox;
    cbxListaCalib2ax: TCheckBox;
    cbxListaSN: TCheckBox;
    cbxListaDefaults1ax: TCheckBox;
    cbxListaDefaults2ax: TCheckBox;
    btnFindCanCH: TButton;
    Button12: TButton;
    Label23: TLabel;
    mmNoteFine: TMemo;
    btnOkNote: TcxButton;
    img1Seq4: TImage;
    imgFineProc: TImage;
    Label22: TLabel;
    Step28NoteCommenti1: TMenuItem;
    viewPanel30: TMenuItem;
    cbxListaDefaultsEND: TCheckBox;
    edtDUTaddr: TcxComboBox;
    Label39: TLabel;
    Label38: TLabel;
    gaugeXRange: TdxGaugeCircularScaleRange;
    gaugeYRange: TdxGaugeCircularScaleRange;
    cboxCanCH: TcxComboBox;
    Label40: TLabel;
    btnSearchCanNodes: TButton;
    Label41: TLabel;
    cboxCanREF: TcxComboBox;
    Label42: TLabel;
    cboxCanDUT: TcxComboBox;
    cbxCAN1baud: TcxComboBox;
    cbxCAN2baud: TcxComboBox;
    cbxCANbaud: TcxComboBox;
    StartCalibrazioneBANCO1: TMenuItem;
    procedure btnWriteCANClick(Sender: TObject);
    procedure btnReadCANClick(Sender: TObject);
    procedure tkBar2Change(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure cbxSelectAllClick(Sender: TObject);
    procedure btnInitCANClick(Sender: TObject);
    procedure btnCloseCANClick(Sender: TObject);
    procedure ResetCAN(Sender: TObject);
    procedure getCANstatus(Sender: TObject);
    procedure edtCanDataChange(Sender: TObject);
    procedure btnWriteRawClick(Sender: TObject);
    procedure btnReadRawClick(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure Button8Click(Sender: TObject);
    procedure tmrCanReadTimer(Sender: TObject);
    procedure tmrFasiBlinkTimer(Sender: TObject);
    procedure lviewCanCMDsDblClick(Sender: TObject);
    procedure SetupC1Click(Sender: TObject);
    procedure Step101Click(Sender: TObject);
    procedure Step111Click(Sender: TObject);
    procedure Step121Click(Sender: TObject);
    procedure Step131Click(Sender: TObject);
    procedure Step141Click(Sender: TObject);
    procedure Step151Click(Sender: TObject);
    procedure Step161Click(Sender: TObject);
    procedure MenuItem2Click(Sender: TObject);
    procedure MenuItem3Click(Sender: TObject);
    procedure MenuItem4Click(Sender: TObject);
    procedure MenuItem5Click(Sender: TObject);
    procedure MenuItem6Click(Sender: TObject);
    procedure MenuItem7Click(Sender: TObject);
    procedure MenuItem10Click(Sender: TObject);
    procedure MenuItem11Click(Sender: TObject);
    procedure ABORT1Click(Sender: TObject);
    procedure ResetAbort1Click(Sender: TObject);
    procedure Scarto1Click(Sender: TObject);
    procedure Scarto2Click(Sender: TObject);
    procedure Writecal1Click(Sender: TObject);
    procedure LoadDefaults1Click(Sender: TObject);
    procedure LetturaParametri1Click(Sender: TObject);
    procedure ScriviParametri1Click(Sender: TObject);
    procedure SalvaParametri1Click(Sender: TObject);
    procedure LoadCorrBanco1Click(Sender: TObject);
    procedure LoadCorrBancoY1Click(Sender: TObject);
    procedure SetupA1Click(Sender: TObject);
    procedure CreaReport30Click(Sender: TObject);
    procedure GenTestDataClick(Sender: TObject);
    procedure tabctrlFasiChange(Sender: TObject);
    procedure TestsortDict1Click(Sender: TObject);
    procedure rgXlsZoomxPropertiesChange(Sender: TObject);
    procedure cbxShowGridClick(Sender: TObject);
    procedure cbxHeadingsClick(Sender: TObject);
    procedure MenuItem1Click(Sender: TObject);
    procedure MenuItem8Click(Sender: TObject);
    procedure MenuItem12Click(Sender: TObject);
    procedure MenuItem13Click(Sender: TObject);
    procedure MenuItem14Click(Sender: TObject);
    procedure MenuItem15Click(Sender: TObject);
    procedure MenuItem16Click(Sender: TObject);
    procedure MenuItem17Click(Sender: TObject);
    procedure SetupFClick(Sender: TObject);
    procedure SalvasuDB29Click(Sender: TObject);
    procedure SaveDB1Click(Sender: TObject);
    procedure cxButton3Click(Sender: TObject);
    procedure debugReadAngle1Click(Sender: TObject);
    procedure StartDUTcyclicalRDClick(Sender: TObject);
    procedure StopDUTcyclicalRDClick(Sender: TObject);
    procedure StopREF2cyclicalRDClick(Sender: TObject);
    procedure StartREF2cyclicalRDClick(Sender: TObject);
    procedure edtRefSamplesInRangePropertiesChange(Sender: TObject);
    procedure edtScartoMax_calibPropertiesChange(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure SaveREFParams1Click(Sender: TObject);
    procedure Button9Click(Sender: TObject);
    procedure Button10Click(Sender: TObject);
    procedure ReadallCalib1Click(Sender: TObject);
    procedure SetCalibintvalue1Click(Sender: TObject);
    procedure ClearLog1Click(Sender: TObject);
    procedure viewPanel28Click(Sender: TObject);
    procedure viewPanel29Click(Sender: TObject);
    procedure Beeptest1Click(Sender: TObject);
    procedure LetturaParametri25Click(Sender: TObject);
    procedure ScriviParametri26Click(Sender: TObject);
    procedure SalvaParametri27Click(Sender: TObject);
    procedure menuSetupEClick(Sender: TObject);
    procedure ResetTreeviewicons1Click(Sender: TObject);
    procedure edtScartoMax_testPropertiesChange(Sender: TObject);
    procedure ClearTSPairs1Click(Sender: TObject);
    procedure LoadListaVerifica24Click(Sender: TObject);
    procedure btnFindCanCHClick(Sender: TObject);
    procedure Button12Click(Sender: TObject);
    procedure viewPanel30Click(Sender: TObject);
    procedure btnOkNoteClick(Sender: TObject);
    procedure edtDUTaddrChange(Sender: TObject);
    procedure cbxCAN1baudPropertiesChange(Sender: TObject);
    procedure btnSearchCanNodesClick(Sender: TObject);
    procedure cboxCanREFChange(Sender: TObject);
    procedure cboxCanDUTChange(Sender: TObject);
    procedure cboxCanCHChange(Sender: TObject);
    procedure cbxCAN2baudPropertiesChange(Sender: TObject);
    procedure cbxCANbaudPropertiesChange(Sender: TObject);
    procedure StartCalibrazioneBANCO1Click(Sender: TObject);
  private
    xlsReport: TXlsFile;
  //Xls: TExcelFile;
    ImgExport: TFlexCelImgExport;

    FMyEvent : TMyEvent;
    FLoading : Boolean;
    { Private declarations }
  public
    const kFrameVersion = 'F2.01';      // versione sorgente di questa Frame.

    procedure AfterConstruction; override;    // da verificare
    procedure BeforeDestruction; override;    //
    // https://stackoverflow.com/questions/3979298/how-to-simulate-an-ondestroy-event-on-a-tframe-in-delphi#_=_
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    (*  the DFM properties have not necessarily been read yet when the constructor is running.
        Instead, override the "Loaded" method of your frame class and put your code there.
        It's called after the properties are loaded from the DFM.
        Ma Non lo uso perchè non avviene in momento coerente !!!
    *)
    function  Scan4PCanInterfaces: boolean;
    procedure LoadCombo_4CanInterfaces(aComboBox: TcxComboBox);
    procedure SetAdminMode(isAdmin:boolean); override;  //reintroduce;
    function  LoadJobList:boolean;

    // FASI da definire per il collaudo
    function SetupFase_A(var JobParams: TJobParams): boolean;
    function SetupFase_B(var JobParams: TJobParams): boolean;
    function SetupFase_C(var JobParams: TJobParams): boolean;
    function SetupFase_D(var JobParams: TJobParams): boolean;
    function Setup4_CD(var JobParams: TJobParams; testAxis:char): boolean;
    function SetupFase_E(var JobParams: TJobParams): boolean;
    function SetupFase_F(var JobParams: TJobParams): boolean;
    function SetupFase_G(var JobParams: TJobParams): boolean;
    function Force_Resolution4EOL: boolean;

    // STEP da definire per il collaudo
    function RunStep_A1(var JobParams: TJobParams): boolean;
    function RunStep_A2(var JobParams: TJobParams): boolean;
    function RunStep_A3(var JobParams: TJobParams): boolean;
    function RunStep_A4(var JobParams: TJobParams): boolean;

    function RunStep_B5(var JobParams: TJobParams): boolean;
    function RunStep_B6(var JobParams: TJobParams): boolean;
    function RunStep_B7(var JobParams: TJobParams): boolean;
    function RunStep_B8(var JobParams: TJobParams): boolean;
    function RunStep_B9(var JobParams: TJobParams): boolean;
    function CalibrazioneDUT(var JobParams: TJobParams; CalDeg:TCalibAngle): boolean;
    function ValidazioneDUT(var JobParams: TJobParams; testDeg:TTestAngle; testAxis:char): boolean;

    function RunStep_C10(var JobParams: TJobParams): boolean;
    function RunStep_C11(var JobParams: TJobParams): boolean;
    function RunStep_C12(var JobParams: TJobParams): boolean;
    function RunStep_C13(var JobParams: TJobParams): boolean;
    function RunStep_C14(var JobParams: TJobParams): boolean;
    function RunStep_C15(var JobParams: TJobParams): boolean;
    function RunStep_C16(var JobParams: TJobParams): boolean;

    function RunStep_D17(var JobParams: TJobParams): boolean;
    function RunStep_D18(var JobParams: TJobParams): boolean;
    function RunStep_D19(var JobParams: TJobParams): boolean;
    function RunStep_D20(var JobParams: TJobParams): boolean;
    function RunStep_D21(var JobParams: TJobParams): boolean;
    function RunStep_D22(var JobParams: TJobParams): boolean;
    function RunStep_D23(var JobParams: TJobParams): boolean;

    function RunStep_E24(var JobParams: TJobParams): boolean;
    function RunStep_E25(var JobParams: TJobParams): boolean;
    function RunStep_E26(var JobParams: TJobParams): boolean;
    function RunStep_E27(var JobParams: TJobParams): boolean;

    function RunStep_F28(var JobParams: TJobParams): boolean;
    function RunStep_F29(var JobParams: TJobParams): boolean;
    function RunStep_F30(var JobParams: TJobParams): boolean;


    // Funzioni e procedure a supporto di quelle primarie di Fasi e Steps.
    function  Load_PCAN_CMDs_fromXLS(xlsFile: string): boolean;
    function  Load_CFG_CmdSet_fromXLS(xlsFile: string): boolean;
    function  LoadCarattBanco_fromXLS(xlsFile: string; testAxis:char): boolean;
    procedure SaveNodeParams(devAddr: string);
    procedure ReloadDefParams(devAddr: string);  // restore parametri di default da Flash (FW).
    function  calcWorkTimeFasi(begFase,endFase: string): TDateTime;
    function  SessionUpdateDB: boolean;
    procedure ResetJobs_Data;
    function  DutCalibrato(devAddr: string): boolean;

    // Funzioni di log sistema son in frame ancestor TframeBase.
    procedure panelPosition(aPanel:TPanel); // posiziona a coordinate fisse.
    procedure panelCenter(aPanel:TPanel);   // posiziona sempre al centro.

    //...  qui altre proc/func.

    // this goes inside the class definition just before the final closing end
    property MyEvent:TMyEvent read FMyEvent write FMyEvent;
    { Public declarations }
  end;

const
    kFrame4Devices : array [0..1] of string = ( 'OIAC3', 'OI2x' );  // device/modelli gestiti dalla frame
    kTimer4cyclicalTxd = '250';   // msec.  evita polling sul REF device.


var
  oiac3Fasi2ax: Toiac3Fasi2ax;
  JobRecord: TJobRecord;
  Ref2_resolution      : Smallint; // [millesimi, centesimi, decimi, interi = 1d,10d,100d,1000d = 1h,Ah,64h,3E8h]   2-byte signed integer.
//                     : integer;  NO!  Se non è dello stesso tipo della misura grezza, produce valori sballati !
  DUT_resolution       : Smallint; // Da notare che i valori da R/W nel subidx (1,10,100...) sono invertiti rispetto ai valori della risoluzione !
//                     : integer;  NO!
  vRef2Addr            : integer;
  currXRef, currYRef   : single;
  currXwRef, currYwRef : word;
  vDutAddr             : integer;
  currXDut, currYDut   : single;
  currXwDut, currYwDut : word;
  currSequence         : integer;
  currControl          : integer;
  scartoMax            : single;
  vRefSamplesInRange   : integer;
  vInvertX_ref      : boolean;
  vInvertY_ref      : boolean;
  vInvertX_dut      : boolean;
  vInvertY_dut      : boolean;
  vRangeError       : boolean;

[volatile] SampleReady: integer;

  arrCorrezioneBanco: TarrCorrezioniBanco;

  CalibRecords : TSPairs<TCalibAngle, TSessionCalibra>; // per gestione record prodotti in fase di Calibrazione.
  TestXRecords : TSPairs<TTestAngle, TSessionTests>;    // per gestione record prodotti in fase di test X.
  TestYRecords : TSPairs<TTestAngle, TSessionTests>;    // per gestione record prodotti in fase di test Y.
  // servono due dictionary perchè non posso duplicare la Key TTestAngle che è uguale per entrambi i set di misure X e Y !

  strTestAngle : array [0..integer(High(TTestAngle))] of string; // ( deg60n, deg30n, deg15n, deg0, deg15p, deg30p, deg60p );  // su quale asse avviene il test dipende da altro parametro.
  strCalibAngle: array [0..integer(High(TCalibAngle))] of string; // ( X0_Y0, X0_Yp60, X0_Yn60, Xp60_Y0, Xn60_Y0 );  // i valori precisi poi dipendono da tabella di calibrazione specifica per famiglia di devices.





implementation

{$R *.dfm}

uses forms.Main, data.Module;



constructor Toiac3Fasi2ax.Create(AOwner: TComponent);
var
  addr: string;
  i: integer;
begin
    inherited Create(AOwner);   // prima esegue Create della classe padre TframeBase !
    // "OnCreate" code here:
    LogStep(10,10, 10,  self.ClassName +' - onCreate.');
    LogStep(11,11, 11,  self.ClassName +' version: '+ kFrameVersion);

    strModuleTitle := 'Calibrazione DUT a 2 Assi';   // titolo della sequenza di test in questa Frame.
    try
        // Compila gestione logica delle fasi/step con relative operazioni.
        LoadJobList();          // definisce anche numSteps e numFasi !
        initJobList('OIAC3');   // inizializza strutture in memoria (StringList, Records) per gestione Fasi/Steps.
        ResetJobs_Data();       // reset anche counters della JobList.
    except
        on E: Exception do begin
            LogStep(12,12, 12,  'EXCEPTION in JobList Create >> '+ E.ClassName +#10#13#9+ ' description: '+E.Message);
            raise;
        end;
    end;

    // forza mode (interfaccia castrata o meno) a seconda del preset all'avvio in unit common,
    // od eventuale override fatto via menù, che richiede input password.
    SetAdminMode( vADMIN_mode );
end;

destructor Toiac3Fasi2ax.Destroy;  // da verificare che funzioni in Fmx !
begin
    LogStep(13,13, 13, self.ClassName +' - OnDestroy.');
    // prima nella sequenza Destroy.

    // "OnDestroy" code here:
    // Release memory, if obtained

    if chanCAN_isOpen then begin
        CAN_Release( CHref );
        if CHdut.Handle <> CHref.Handle then
            CAN_Release( CHdut );
    end;

    inherited Destroy;  // solo alla fine esegue infine OnDestroy della classe padre TframeBase !
end;

procedure Toiac3Fasi2ax.AfterConstruction;
var
  TestAngle  : TTestAngle;
  CalibAngle : TCalibAngle;
begin
    inherited; // prima esegue la sequenza AfterConstruction della classe padre TframeBase !
    LogStep(14,14, 14, self.ClassName +' - onAfterConstruction.');
    // "AfterConstruction" code here:

    // Ricava nomi degli Enum per uso display.
    for TestAngle := Low(TTestAngle) to High(TTestAngle) do
        strTestAngle[ord(TestAngle)] := TypInfo.GetEnumName(TypeInfo(TTestAngle), ord(TestAngle));
    for CalibAngle := Low(TCalibAngle) to High(TCalibAngle) do
        strCalibAngle[ord(CalibAngle)] := TypInfo.GetEnumName(TypeInfo(TCalibAngle), ord(CalibAngle));

    CalibRecords := TSPairs<TCalibAngle, TSessionCalibra>.Create();  // per gestione record prodotti in fasi di Calibrazione.
    TestXRecords := TSPairs<TTestAngle, TSessionTests>.Create();     // per gestione record prodotti in fasi di test X.
    TestYRecords := TSPairs<TTestAngle, TSessionTests>.Create();     // per gestione record prodotti in fasi di test Y.
    // servono due dictionary perchè non posso duplicare la Key TTestAngle che è uguale per entrambi i set di misure X e Y !

    panel_A.BevelOuter := bvNone;   // bevel utile solo a design time
    panel_B.BevelOuter := bvNone;   //
    panel_CD.BevelOuter := bvNone;  //

    panel_Gauges.BringToFront;      // per avere slide-up sempre pulito
end;

procedure Toiac3Fasi2ax.Beeptest1Click(Sender: TObject);
begin
    Winapi.Windows.Beep( 1500, 600) ;   // frequency, duration
end;

procedure Toiac3Fasi2ax.BeforeDestruction;
begin
    LogStep(15,15, 15, self.ClassName +' - onBeforeDestruction');
    // "BeforeDestruction" code here:

    // Stop any timer.
    tmrFasiBlink.Enabled := False;
    tmrCanRead.Enabled := False;

    with dmBase do begin
        //tabDut.Close;
        //tabReport.Close;
    end;

  //JobList.Clear;      // No,
  //JobList.Free;       // perchè viene fatto (inherited) in ancestor class TframeBase

    FreeAndNil(CalibRecords);  //
    FreeAndNil(TestXRecords);  // free solo non basta!
    FreeAndNil(TestYRecords);  //

    inherited;  // solo alla fine esegue BeforeDestroy della classe padre TframeBase !
end;

(*
procedure TframeFasi.Loaded; // Non usato perchè non avviene in momento coerente !
begin
    inherited;
    //Showmessage('Loaded');
    LogStep(16,16, 16, 'frameFasi  - onLoaded.');
end;
*)


procedure Toiac3Fasi2ax.ResetJobs_Data; // permette una ripetizione pulita della stessa Sessione !
begin
    inherited;   // esegue il Reset dei parametri Base.

    // esegue solo il Reset dei parametri creati in questa frame.
    if CalibRecords <> nil then         // per gestione record prodotti in fasi di Calibrazione.
        CalibRecords.Clear;             // Clear removes all keys and values from the dictionary. The Count property is set to 0.
    if TestXRecords <> nil then         // per gestione record prodotti in fasi di test X.
        TestXRecords.Clear;
    if TestYRecords <> nil then         // per gestione record prodotti in fasi di test Y.
        TestYRecords.Clear;
    // servono due dictionary perchè non posso duplicare la Key TTestAngle che è uguale per entrambi i set di misure X e Y !
    mainForm.mmLog.lines.add('Reset Session DATA done.');
end;

procedure Toiac3Fasi2ax.SetAdminMode(isAdmin:boolean);
begin
    inherited;   // prima esegue i setup della classe padre TframeBase !

    // gestione centralizzata di visualizzazioni Admin o User per tutta la frame!
    if isAdmin then begin
        //  ADMIN mode  = interfaccia full control
        PopMenuFaseA.AutoPopup := True;
        PopMenuFaseB.AutoPopup := True;
        PopMenuFaseC.AutoPopup := True;
        PopMenuFaseD.AutoPopup := True;
        PopMenuFaseE.AutoPopup := True;
        PopMenuFaseF.AutoPopup := True;
        grpBoxListaParams.Enabled := True;
        lviewCalibVal.Visible := True;
        lviewCarattBanco.Visible := True;
        //...
    end
    else begin
        //  User mode  = interfaccia castrata
        PopMenuFaseA.AutoPopup := False;
        PopMenuFaseB.AutoPopup := False;
        PopMenuFaseC.AutoPopup := False;
        PopMenuFaseD.AutoPopup := False;
        PopMenuFaseE.AutoPopup := False;
        PopMenuFaseF.AutoPopup := False;
        grpBoxListaParams.Enabled := False;
        lviewCalibVal.Visible := False;
        lviewCarattBanco.Visible := False;
        //...
    end;
end;



// Infatti Fase_A meglio visualizzarla solo dopo Start !

/////////////////////////////////  * CAN *  ///////////////////////////////////////////////////

procedure Toiac3Fasi2ax.edtCanDataChange(Sender: TObject);
begin
    if edtCanData.Text <> '' then begin
        cbxDataType.Enabled := True;
        cbxDataType.color := clWindow;
    end
    else begin
        cbxDataType.Enabled := False;
        cbxDataType.color := clbtnFace;
    end;
end;

procedure Toiac3Fasi2ax.edtDUTaddrChange(Sender: TObject);
begin
    vDutAddr := strCanValue2integer(edtDutAddr.text);
end;

procedure Toiac3Fasi2ax.cboxCanREFChange(Sender: TObject);
begin
    if cboxCanREF.ItemIndex > -1 then
        CHref.Handle := strCanValue2integer( cboxCanREF.text );
end;

procedure Toiac3Fasi2ax.cboxCanDUTChange(Sender: TObject);
begin
    if cboxCanDUT.ItemIndex > -1 then
        CHdut.Handle := strCanValue2integer( cboxCanDUT.text );
end;

procedure Toiac3Fasi2ax.cboxCanCHChange(Sender: TObject);
begin
    if cboxCanCH.ItemIndex > -1 then
        CHcan.Handle := strCanValue2integer( cboxCanCH.text );
end;


procedure Toiac3Fasi2ax.cbxCAN2baudPropertiesChange(Sender: TObject);
begin
    case cbxCAN2baud.itemIndex of
        0: CHdut.Baud := PCAN_BAUD_1M;
        1: CHdut.Baud := PCAN_BAUD_800K;
        2: CHdut.Baud := PCAN_BAUD_500K;
        3: CHdut.Baud := PCAN_BAUD_250K;
        4: CHdut.Baud := PCAN_BAUD_100K;
        5: CHdut.Baud := PCAN_BAUD_50K;
        6: CHdut.Baud := PCAN_BAUD_20K;
        7: CHdut.Baud := PCAN_BAUD_5K;
    else
        CHdut.Baud := PCAN_BAUD_500K;   // default
    end;
end;

procedure Toiac3Fasi2ax.cbxCAN1baudPropertiesChange(Sender: TObject);
begin
    case cbxCAN1baud.itemIndex of
        0: CHref.Baud := PCAN_BAUD_1M;
        1: CHref.Baud := PCAN_BAUD_800K;
        2: CHref.Baud := PCAN_BAUD_500K;
        3: CHref.Baud := PCAN_BAUD_250K;
        4: CHref.Baud := PCAN_BAUD_100K;
        5: CHref.Baud := PCAN_BAUD_50K;
        6: CHref.Baud := PCAN_BAUD_20K;
        7: CHref.Baud := PCAN_BAUD_5K;
    else
        CHref.Baud := PCAN_BAUD_500K;   // default
    end;
end;

procedure Toiac3Fasi2ax.cbxCANbaudPropertiesChange(Sender: TObject);
begin
    case cbxCANbaud.itemIndex of
        0: CHcan.Baud := PCAN_BAUD_1M;
        1: CHcan.Baud := PCAN_BAUD_800K;
        2: CHcan.Baud := PCAN_BAUD_500K;
        3: CHcan.Baud := PCAN_BAUD_250K;
        4: CHcan.Baud := PCAN_BAUD_100K;
        5: CHcan.Baud := PCAN_BAUD_50K;
        6: CHcan.Baud := PCAN_BAUD_20K;
        7: CHcan.Baud := PCAN_BAUD_5K;
    else
        CHcan.Baud := PCAN_BAUD_500K;   // default
    end;
end;

procedure Toiac3Fasi2ax.edtRefSamplesInRangePropertiesChange(Sender: TObject);
begin
    // = 10  num samples anche durante il test, per correzioni realtime.
    // lo spinedit value è un variant predefinito a vtInt.
    vRefSamplesInRange := edtRefSamplesInRange.value;
end;

procedure Toiac3Fasi2ax.edtScartoMax_calibPropertiesChange(Sender: TObject);
begin
  //scartoMax := strTofloat(edtScartoMax_calib.text);  // = ±0.02  aggiorna scarto anche durante il test, per check realtime.
    scartoMax := edtScartoMax_calib.Value;             // lo spinedit value è un variant predefinito a vtFloat.
end;

procedure Toiac3Fasi2ax.edtScartoMax_testPropertiesChange(Sender: TObject);
begin
  //scartoMax := strTofloat(edtScartoMax_test.text);  // = ±0.02  aggiorna scarto anche durante il test, per check realtime.
    scartoMax := edtScartoMax_test.Value;             // lo spinedit value è un variant predefinito a vtFloat.
end;

procedure Toiac3Fasi2ax.GenTestDataClick(Sender: TObject);
var
  TestAngle  : TTestAngle;
  CalibAngle : TCalibAngle;
  testRecord : TSessionTests;
  calRecord  : TSessionCalibra;
  CanCmd: TCanCmd;
  i: integer;
  aJobRecord: TJobRecord;
  start: TDatetime;
begin
    Randomize;
    // fill TestRecord sorted dictionary con dati random per test report su asse X e Y.
    Sessione.StartStamp := now;                 // TDateTime
    Sessione.EndStamp   := incHour(now, 3);
    Sessione.OperatoreID := 'TEST';

    // Accumula tempi delle Fasi desiderate
    // Fase_B - Calibrazione
    start := now;
    for i := 0 to JobList.Count-1 do begin
        aJobRecord := TJobRecord( JobList.Objects[ i ]);
        with aJobRecord do
            if GroupFase = 'Fase_B' then begin
                BeginStamp := start;
                EndStamp := incMinute(start, 2);
                start := incMinute(EndStamp, 1);     // 1 min. fisso di gap tra fasi/steps
            end;
    end;
    // Fase_C e D - Validazione
    for i := 0 to JobList.Count-1 do begin
        aJobRecord := TJobRecord( JobList.Objects[ i ]);
        with aJobRecord do
            if (GroupFase = 'Fase_C') or (GroupFase = 'Fase_D') then begin
                BeginStamp := start;
                EndStamp := incMinute(start, 2);
                start := incMinute(EndStamp, 1);     // 1 min. fisso di gap tra fasi/steps
            end;
    end;

    // set Calib dummy data per report
    CalibRecords.Clear;  // Clear removes all keys and values from the dictionary. The Count property is set to 0.
    for CalibAngle := Low(TCalibAngle) to High(TCalibAngle) do begin
        calRecord := TSessionCalibra.create;
        calRecord.AsseCorrente := 'CAL2A';
        calRecord.degPairID := strCalibAngle[ord(CalibAngle)];
        calRecord.RefAx := Random * 60;
        calRecord.RefAy := Random * 60;
        calRecord.DutAx_mV := trunc(Random * 45000);
        calRecord.DutAy_mV := trunc(Random * 45000);
        CalibRecords.Add( CalibAngle, calRecord );

        // set Calib dummy data, in dictionary per checklist view.
        case CalibAngle of
              X0_Y0: begin
                      if dicCanCmds.TryGetValue( _2ax_calX0_deg_, CanCmd ) then
                          CanCmd.WrValue := calRecord.DutAx_mV.ToString;  // Campo per calibrazione asse X del DUT.   5555-1
                      if dicCanCmds.TryGetValue( _2ax_calY0_deg_, CanCmd ) then
                          CanCmd.WrValue := calRecord.DutAy_mV.ToString;  // Campo per calibrazione asse Y del DUT.   5555-2
                     end;
             XN_Yxn: begin
                      if dicCanCmds.TryGetValue( _2ax_calXNeg_deg_, CanCmd ) then
                          CanCmd.WrValue := calRecord.DutAx_mV.ToString;  // Campo per calibrazione asse X del DUT.   5555-5
                      if dicCanCmds.TryGetValue( _2ax_calCross_Yxneg_, CanCmd ) then
                          CanCmd.WrValue := calRecord.DutAy_mV.ToString;  // Campo per calibrazione asse Y del DUT.   5555-31
                     end;
             XP_Yxp: begin
                      if dicCanCmds.TryGetValue( _2ax_calXPos_deg_, CanCmd ) then
                          CanCmd.WrValue := calRecord.DutAx_mV.ToString;  // Campo per calibrazione asse X del DUT.   5555-3
                      if dicCanCmds.TryGetValue( _2ax_calCross_Yxpos_, CanCmd ) then
                          CanCmd.WrValue := calRecord.DutAy_mV.ToString;  // Campo per calibrazione asse Y del DUT.   5555-30
                     end;
             YP_Xyp: begin
                      if dicCanCmds.TryGetValue( _2ax_calCross_Xypos_, CanCmd ) then
                          CanCmd.WrValue := calRecord.DutAx_mV.ToString;  // Campo per calibrazione asse Y del DUT.   5555-28
                      if dicCanCmds.TryGetValue( _2ax_calYPos_deg_, CanCmd ) then
                          CanCmd.WrValue := calRecord.DutAy_mV.ToString;  // Campo per calibrazione asse X del DUT.   5555-4
                     end;
             YN_Xyn: begin
                      if dicCanCmds.TryGetValue( _2ax_calCross_Xyneg_, CanCmd ) then
                          CanCmd.WrValue := calRecord.DutAx_mV.ToString;  // Campo per calibrazione asse Y del DUT.   5555-29
                      if dicCanCmds.TryGetValue( _2ax_calYNeg_deg_, CanCmd ) then
                          CanCmd.WrValue := calRecord.DutAy_mV.ToString;  // Campo per calibrazione asse X del DUT.   5555-6
                     end;
        end;
    end;

    // set Test dummy dataX
    TestXRecords.Clear;
    for TestAngle := Low(TTestAngle) to High(TTestAngle) do begin
        testRecord := TSessionTests.create;
        testRecord.AsseCorrente := 'TEST2AX';
        testRecord.degPairID := strTestAngle[ord(TestAngle)];
        testRecord.RefAx := Random * 60;
        testRecord.RefAy := Random * 60;
        testRecord.CorrRefTest := Random;
        testRecord.CorrRefCross := Random;
        testRecord.DutAx_deg := Random * 60;
        testRecord.DutAy_deg := Random * 60;
        testRecord.ErroreTest := Random;
        testRecord.ErroreCross := Random;
        TestXRecords.Add( TestAngle, testRecord);
    end;
    // set Test dummy dataY
    TestYRecords.Clear;
    for TestAngle := Low(TTestAngle) to High(TTestAngle) do begin
        testRecord := TSessionTests.create;
        testRecord.AsseCorrente := 'TEST2AY';
        testRecord.degPairID := strTestAngle[ord(TestAngle)];
        testRecord.RefAx := Random * 60;
        testRecord.RefAy := Random * 60;
        testRecord.CorrRefTest := Random;
        testRecord.CorrRefCross := Random;
        testRecord.DutAx_deg := Random * 60;
        testRecord.DutAy_deg := Random * 60;
        testRecord.ErroreTest := Random;
        testRecord.ErroreCross := Random;
        TestYRecords.Add( TestAngle, testRecord);
    end;
    LogStep(17,17, 17,  'Generato dati fittizi per demo Report.')
end;

procedure Toiac3Fasi2ax.getCANstatus(Sender: TObject);
begin
    if CAN_GetStatus( CHcan ) then
    	// The channel was successfully reset
		LogStep(18,18, 18, 'PCAN STATUS on channel '+ intToHex(CHcan.Handle, 2) +'h  -> OK') //PCAN_ERROR_OK.
	else
    	// An error occurred.  We show the error.
        LogStep(19,19, 19, 'PCAN STATUS on channel '+ intToHex(CHcan.Handle, 2) +'h  -> ERROR: '+ lastCAN_ErrMsg);
end;

procedure Toiac3Fasi2ax.ResetCAN(Sender: TObject);
begin
    if CAN_Reset( CHcan ) then
    	// The channel was successfully reset
		LogStep(20,20, 20, 'PCAN RESET on channel '+ intToHex(CHcan.Handle, 2) +'h  -> OK') //PCAN_ERROR_OK.
	else
    	// An error occurred.  We show the error.
        LogStep(21,21, 21, 'PCAN RESET on channel '+ intToHex(CHcan.Handle, 2) +'h  -> ERROR: '+ lastCAN_ErrMsg);
end;

procedure Toiac3Fasi2ax.btnCloseCANClick(Sender: TObject);
begin
    if CAN_Release( CHcan ) then
    	// The frame was successfully sent
		LogStep(22,22, 22, 'CLOSE channel '+ intToHex(CHcan.Handle, 2) +'h  -> OK') //PCAN_ERROR_OK.
	else
    	// An error occurred.  We show the error.
        LogStep(23,23, 23, 'CLOSE channel '+ intToHex(CHcan.Handle, 2) +'h  -> ERROR: '+ lastCAN_ErrMsg);
end;

procedure Toiac3Fasi2ax.btnInitCANClick(Sender: TObject);
begin
    if CAN_Init( CHcan ) then
    	// The frame was successfully sent
		LogStep(24,24, 24, 'INIT PCAN channel '+ intToHex(CHcan.Handle, 2) +'h  -> OK!') //PCAN_ERROR_OK.
	else
    	// An error occurred.  We show the error.
        LogStep(25,25, 25, 'INIT PCAN channel '+ intToHex(CHcan.Handle, 2) +'h  -> ERROR: '+ lastCAN_ErrMsg);
end;

procedure Toiac3Fasi2ax.btnOkNoteClick(Sender: TObject);
begin
    (sender as TComponent).Tag := 1;  // set flag pressed.
end;

procedure Toiac3Fasi2ax.lviewCanCMDsDblClick(Sender: TObject);
var
  ListItem: TListItem;
begin
    ListItem := lviewCanCMDs.Selected;
    edtIndex.Text := ListItem.caption;
    edtSubIdx.Text := ListItem.SubItems[0];
    edtCanData.Text := '';
  //cbxDataType.Text := ListItem.SubItems[3];   // no, perchè ItemIndex non cambia !
    cbxDataType.ItemIndex := cbxDataType.Items.IndexOf(ListItem.SubItems[3]); // ok, cambia ItemIndex e quindi anche il Text !
end;


procedure Toiac3Fasi2ax.btnWriteCANClick(Sender: TObject);
var
  Can: TCanFrame;
begin
    // Prepara la frame con indirizzi e Dati da inviare
    Can.RTR := False;
    Can.NodeID := edtCanID.text;
    Can.Index  := edtIndex.text;
    Can.Subidx := edtSubIdx.text;
    // nei campi CAN in input sono ammessi Hex, Dec, blank fra bytes, h finale
    // perchè verranno sempre validati e trimmati dalla CAN_SendFrame().
    // Can.Len viene calcolata dalla funzione CAN_SendFrame, con default Data Length $40 = AnySize (=8)
    if edtCanData.text <> '' then begin
        if cbxCanDataStrPrefix.text <> '' then begin
            Can.Data := cbxCanDataStrPrefix.text;
            Can.Data := Can.Data + Hex2String(edtCanData.text);
        end
        else
            Can.Data := edtCanData.text;
        Can.Datatype := cbxDataType.Items[cbxDataType.ItemIndex];    // int8, uns16, ...
      //Can.Datatype := cbxDataType.Text;                            //
    end
    else
        Can.Data := '';
    // invio can frame
    if CAN_SendFrame(Can) then
        // The frame was successfully sent
        LogStep(26,26, 26, 'CAN Frame SENT Successfully,  ID:'+Can.CobID +'   LenPL:'+ intToStr(Can.LenPayLoad) +'   Frame:'+ hexLarge(Can.Frame))
    else
        LogStep(27,27, 27,  lastCAN_ErrMsg );
end;

procedure Toiac3Fasi2ax.btnReadCANClick(Sender: TObject);
var
  Can: TCanFrame;
begin
    edtRxIndex.text := '';
    edtRxSubIdx.text := '';
    edtRxCanData.text := '';
    edtRxDataAsInt.text := '';
    // init e clean vari campi TCanFrame necessari pre receive, perchè la TCanFrame locale
    // non è stata inizializzata da precedente CAN_SendFrame() e con parametri vuoti darebbe mismatch vari...!
    // Trim all blanks, Hex symbol chars, the replacement is case insensitive, toglie any h , H !
    Can.NodeID := textCanValue2strHex(edtCanID.text, '2');
    Can.Index  := textCanValue2strHex(edtIndex.text, '4');
    Can.Subidx := textCanValue2strHex(edtSubIdx.text, '2');
    // legge response ok/err
    if CAN_ReceiveFrame( Can ) then begin
        // OK, the frame was successfully read.
        // Received Payload in Can.Data
        LogStep(28,28, 28, 'CAN Frame RECEIVED Successfully,  ID:'+Can.CobID +'   LenPL:'+intToStr(Can.LenPayLoad) +'   Frame:'+hexLarge(Can.Frame));
        edtRxCanID.text := Can.CobID;
        edtRxIndex.text := Can.Index;
        edtRxSubIdx.text := Can.SubIdx;
        edtRxCanData.text := Can.Data;
        edtRxDataAsInt.text := intToStr(strToIntDef('$'+ Can.Data, 0));
    end
    else
        LogStep(29,29, 29,  lastCAN_ErrMsg );
end;



//END///////////////////////////////  * CAN *


procedure Toiac3Fasi2ax.Button10Click(Sender: TObject);
begin
    if CAN_SendNMT( CHcan, cmdStopNode, edtCanID.text ) then begin   // sezione Admin
        // The Start was successfully sent
        LogStep(30,30, 30, 'CAN StopNode SENT Successfully to NodeID:'+ edtCanID.text );
        // quindi pulisco rx buffer.
        CAN_FlushRxBuffer(CHcan);
    end
    else begin
        // l' Errore è bloccante, quindi esco.
        LogStep(31,31, 31, 'CAN StopNode ERROR  on NodeID:'+ edtCanID.text );
        LogStep(32,32, 32,  lastCAN_ErrMsg );
    end;
end;

procedure Toiac3Fasi2ax.Button12Click(Sender: TObject);
begin
    ReloadDefParams( edtCanID.text );
end;

procedure Toiac3Fasi2ax.ReloadDefParams(devAddr: string);  // restore parametri di default da Flash (FW).
var
  Can: TCanFrame;
begin
    // Restore da Flash in Ram dei parametri di default, poi serve anche Save cfg.
    try
        // prima di tutto stop eventuali HeartBeat !
        Can.Command  := _Heartbeat_interval_;      // HeartBeat cyclical transmission [multiple of 1ms, 0 = disabled]
        Can.Data     := '0';                       // 0 msec = timer Stopped
        Can.NodeID   := devAddr;                   // str ID del device
        if not CAN_WriteParam( Can ) then begin
            LogStep(33,33, 33, 'CAN Stop HeartBeat ERROR on NodeID:'+ devAddr );
            LogStep(34,34, 34, lastCAN_ErrMsg);
            exit;
        end;
        // The parameter was Written successfully.
        LogStep(35,35, 35, 'CAN Stop HeartBeat command SENT Successfully to NodeID:'+ devAddr );

        if CAN_NodeRestoreDef( devAddr ) then begin
            // The frame was successfully sent
            LogStep(36,36, 36, 'CAN RestoreParams command SENT Successfully to NodeID:'+ devAddr );
        end
        else begin
            // l' Errore è bloccante, quindi esco.
            LogStep(37,37, 37, 'CAN RestoreParams ERROR on NodeID:'+ devAddr );
            LogStep(38,38, 38,  lastCAN_ErrMsg );
            exit;
        end;
        // Ok, niente msec. attesa perchè RD da flash è già completato.

        // Necessaria save in EEProm per rendere stabili i nuovi parametri di configurazione.
        SaveNodeParams(devAddr);
        // comprende Reset node e Flush can buffer.
    except
      on E: Exception do begin
          LogStep(39,39, 39,  'EXCEPTION on RELOAD & Reset node '+devAddr+' >> '+ E.ClassName +#10#13#9+ ' description: '+E.Message);
      end;
    end;
end;

procedure Toiac3Fasi2ax.Button1Click(Sender: TObject);
begin
    //mmLog.Lines.Add(#10+'Absolute NodeIndex = '+ inttoStr(trvJobList.Selected.AbsoluteIndex));// x debug.
	//LogText('Screen size: '+ Screen.Width.ToString +' x '+ Screen.Height.ToString);   // 1920 x 1080
    mainForm.trvJobList.SetFocus;     // necessario per avere campo full blue quando selected !
    mainForm.trvJobList.items[StrToInt(mainForm.edtSerialeTO.Text)].Selected := true;
//  trvJobList.items[StrToInt(edtSerialeTO.Text)].StateIndex := iconOK;
    mainForm.trvJobList.items[StrToInt(mainForm.edtSerialeTO.Text)].ImageIndex := iconOK;
    mainForm.trvJobList.items[StrToInt(mainForm.edtSerialeTO.Text)].SelectedIndex := iconOK;   // image quando selezionato.
(*
            currStepNode.ImageIndex := 1;      // main image
            currStepNode.SelectedIndex := 4;   // image quando selezionato.
            currStepNode.StateIndex := -1;     // 1=OffLine.  values of 0 or -1 mean no state icon.
  *)
end;


procedure Toiac3Fasi2ax.Button2Click(Sender: TObject);
begin
    SaveNodeParams( edtDUTaddr.text );
end;

procedure Toiac3Fasi2ax.SaveREFParams1Click(Sender: TObject);
begin
    SaveNodeParams( edtCanID.text );
end;

procedure Toiac3Fasi2ax.SaveNodeParams(devAddr: string);  // save in EEProm parametri di configurazione.
begin
    try
        // Save cfg
        if CAN_NodeSaveAll( devAddr ) then begin
            // The frame was successfully sent
            LogStep(40,40, 40, 'CAN SaveParams command SENT Successfully to NodeID:'+ devAddr );
        end
        else begin
            // l' Errore è bloccante, quindi esco.
            LogStep(41,41, 41, 'CAN SaveParams ERROR on NodeID:'+ devAddr );
            LogStep(42,42, 42,  lastCAN_ErrMsg );
            exit;
        end;
        // Ok, niente msec. attesa perchè WR in eeprom già completato.

        // Reset DUT
        if CAN_SendNMT( CHdut, cmdResetNode, devAddr ) then begin
            // The Reset was successfully sent
            LogStep(43,43, 43, 'CAN ResetNode SENT Successfully to NodeID:'+ devAddr );
        end
        else begin
            // l' Errore è bloccante, quindi esco.
            LogStep(44,44, 44, 'CAN ResetNode ERROR  on NodeID:'+ devAddr );
            LogStep(45,45, 45,  lastCAN_ErrMsg );
        end;
        // DUT reset & boot completed !
        // Read back cfg, non serve se WR senza errori...
        if not CAN_FlushRxBuffer(CHdut) then begin
            LogStep(46,46, 46,  lastCAN_ErrMsg );
            exit;
        end;
        LogStep(47,47, 47, 'OK FLUSH Can queue Successful.');    // preparato rx queue clean.
    except
      on E: Exception do begin
          LogStep(48,48, 48,  'EXCEPTION on SAVE & Reset node '+devAddr+' >> '+ E.ClassName +#10#13#9+ ' description: '+E.Message);
      end;
    end;
end;


procedure Toiac3Fasi2ax.Button3Click(Sender: TObject);
begin
    if not Load_PCAN_CMDs_fromXLS( kAppPath + dmBase.vCANcmdSource ) then
        LogText('ERROR Reading PCAN commands file: '+ dmBase.vCANcmdSource);
end;

procedure Toiac3Fasi2ax.Button4Click(Sender: TObject);
begin
    if CAN_FlushRxBuffer(CHcan) then begin
        LogStep(49,49, 49, 'FLUSH PCAN Successful.');
    end
end;

procedure Toiac3Fasi2ax.Button5Click(Sender: TObject);
begin
    if CAN_UnlockParams( edtCanID.text ) then begin
        // The frame was successfully sent
        LogStep(50,50, 50, 'CAN Unlock command CONFIRMED by NodeID:'+ edtCanID.text );
    end
    else begin
        // l' Errore è bloccante, quindi esco.
        LogStep(51,51, 51, 'CAN Unlock ERROR on NodeID:'+ edtCanID.text );
        LogStep(52,52, 52,  lastCAN_ErrMsg );
    end;
end;

procedure Toiac3Fasi2ax.Button6Click(Sender: TObject);
begin
    if CAN_LockParams( edtCanID.text ) then begin
        // The frame was successfully sent
        LogStep(53,53, 53, 'CAN Lock command CONFIRMED by NodeID:'+ edtCanID.text );
    end
    else begin
        // l' Errore è bloccante, quindi esco.
        LogStep(54,54, 54, 'CAN Lock ERROR on NodeID:'+ edtCanID.text );
        LogStep(55,55, 55,  lastCAN_ErrMsg );
    end;
end;

procedure Toiac3Fasi2ax.Button7Click(Sender: TObject);
begin
    if CAN_UnlockSN( edtCanID.text ) then begin
        // The frame was successfully sent
        LogStep(56,56, 56, 'CAN UnlockSN command CONFIRMED by NodeID:'+ edtCanID.text );
    end
    else begin
        // l' Errore è bloccante, quindi esco.
        LogStep(57,57, 57, 'CAN UnlockSN ERROR on NodeID:'+ edtCanID.text );
        LogStep(58,58, 58,  lastCAN_ErrMsg );
    end;
end;

procedure Toiac3Fasi2ax.Button8Click(Sender: TObject);
begin
    // Reset DUT
    if CAN_SendNMT( CHcan, cmdResetNode, edtCanID.text ) then begin
        // The Reset was successfully sent
        LogStep(59,59, 59, 'CAN ResetNode SENT Successfully to NodeID:'+ edtCanID.text );
        // dopo boot il DUT invia   COB:70A   Len:1   Data:00 00 00 00 00 00 00 00
        // quindi potrei verificare, ma almeno pulisco rx buffer.
        CAN_FlushRxBuffer(CHcan);
    end
    else begin
        // l' Errore è bloccante, quindi esco.
        LogStep(60,60, 60, 'CAN ResetNode ERROR  on NodeID:'+ edtCanID.text );
        LogStep(61,61, 61,  lastCAN_ErrMsg );
    end;
    // DUT reset & boot completed !
end;


procedure Toiac3Fasi2ax.Button9Click(Sender: TObject);
begin
    if CAN_SendNMT( CHcan, cmdStartNode, edtCanID.text ) then begin
        // The Start was successfully sent
        LogStep(62,62, 62, 'CAN StartNode SENT Successfully to NodeID:'+ edtCanID.text );
        // quindi pulisco rx buffer.
        CAN_FlushRxBuffer(CHcan);
    end
    else begin
        // l' Errore è bloccante, quindi esco.
        LogStep(63,63, 63, 'CAN StartNode ERROR  on NodeID:'+ edtCanID.text );
        LogStep(64,64, 64,  lastCAN_ErrMsg );
    end;
end;

procedure Toiac3Fasi2ax.btnWriteRawClick(Sender: TObject);
var
  txLen: integer;
begin
    edtCobIDrcvd.text := '';      // reset Rx fields
    edtRawCanLenRcvd.text := '';
    edtRawCanDataRcvd.text := '';

    // Send the message
    if CAN_WriteRawFrame(CHcan, edtCobID.text, edtRawCanData.text, txLen) then begin
    	// The frame was successfully sent
        edtRawCanLen.text := txLen.toString;
		LogStep(65,65, 65, '> CAN Frame SENT Successfully,  COB:'+edtCobID.text +'   Len:'+edtRawCanLen.text+'   Data:'+edtRawCanData.text); //PCAN_ERROR_OK.
    end
	else
    	// An error occurred.  We show the error.
        LogStep(66,66, 66,  '* '+ lastCAN_ErrMsg );
end;

procedure Toiac3Fasi2ax.btnReadRawClick(Sender: TObject);
var
  rcvdCob, RcvdLen, RcvdData: Ansistring;
begin
    rcvdCob := '';
    rcvdLen := '';
    rcvdData := '';
    if CAN_ReadRawFrame(CHcan, rcvdCob, RcvdLen, RcvdData) then begin
    	// The frame was successfully sent
        LogStep(67,67, 67, '< CAN Frame RECEIVED Successfully,  COB:'+rcvdCob +'   Len:'+RcvdLen +'   Data:'+RcvdData);
        edtCobIDrcvd.text := rcvdCob;
        edtRawCanLenRcvd.text := rcvdLen;
        edtRawCanDataRcvd.text := rcvdData;
    end
	else
        LogStep(68,68, 68,  '* '+ lastCAN_ErrMsg );
end;

function Toiac3Fasi2ax.Scan4PCanInterfaces: boolean;
var
  i: integer;
  Ch: string;
begin
    LogStep(69,69, 69, 'SCAN for PCAN-USB Adapters...');   // carica la relativa ComBox CAN-ID in Admin tab.
    result := False;
    numCanCH := 0;

    // Search CHs...
    if CAN_EnumChannels() then begin    // Cerca le interfacces USB PCan presenti

        if CanChannels.count > 0 then begin
            LogStep(70,70, 70, 'Found '+ CanChannels.count.ToString +' PCAN interfaces:');
            // dopo display lista dettagliata, la riduco per uso nelle combobox
            for i := 0 to CanChannels.count-1 do begin
                LogStep(71,71, 71, #9#9'Channel > '+ CanChannels.Strings[i] );
                Ch := Copy( CanChannels.Strings[i], Pos('(', CanChannels.Strings[i]) + 1, 3);
              //Ch := StringReplace(Ch, 'h',' ', [rfReplaceAll, rfIgnoreCase]);
                Ch := Trim(Ch);
                CanChannels.Strings[i] := Ch;
            end;
        end
        else begin
            lastCAN_ErrMsg := 'PCAN interfaces NOT Found !';
            LogStep(72,72, 72, 'WARNING: '+ lastCAN_ErrMsg );
            exit;
        end;

        LoadCombo_4CanInterfaces( cboxCanCH ); // carica la lista in cboxCanCH in Admin tab.
        numCanCH := CanChannels.count;
        result := True;
    end
    else begin
        // l' Errore è bloccante, quindi esco.
        LogStep(73,73, 73, 'ENUM PCAN interfaces ERROR: '+ lastCAN_ErrMsg );
    end;
end;

procedure Toiac3Fasi2ax.LoadCombo_4CanInterfaces(aComboBox: TcxComboBox);
var
  i: integer;
  tmpEvent: procedure(Sender: TObject) of object;
begin
    tmpEvent := aComboBox.Properties.OnChange;
    aComboBox.Properties.OnChange := nil;  // Sgancia evento
    try
        aComboBox.Properties.Items.Clear;
      //aComboBox.Properties.Items.Assign(CanChannels);    // no direct assign perchè devo estrarre solo l'id.
        for i := 0 to CanChannels.count -1 do
            aComboBox.Properties.Items.Add( CanChannels.Strings[i] );
    finally
        aComboBox.Properties.OnChange := tmpEvent;         // Ripristina evento
      //aComboBox.ItemIndex := aComboBox.Items.Count - 1;  // posiziona su ultima PCan
        aComboBox.itemIndex := 0;    // posiziona su prima PCan ed esegue onchange event, se esiste.
    end;
end;


procedure Toiac3Fasi2ax.btnFindCanCHClick(Sender: TObject);
begin
    Scan4PCanInterfaces();
end;

procedure Toiac3Fasi2ax.btnSearchCanNodesClick(Sender: TObject);
var
  i, nodes: integer;
begin
    LogStep(74,74, 74, 'SCAN for active Node IDs...');
    if not CAN_FlushRxBuffer(CHcan) then begin
        LogStep(75,75, 75,  lastCAN_ErrMsg );
        exit;
    end;

    // Search IDs...
    CanNodes.Clear;                             // reset lista nodi a cura chiamante per permettere di accumulare scan multipli su più Canali.
    if CAN_EnumNodes(CHcan, nodes) then begin  // Reset ALL nodes and wait for IDs
        // dopo boot il DUT invia solo la Boot frame   COB:70A   Len:1   Data:00 00 00 00 00 00 00 00
        if nodes > 0 then begin
            LogStep(76,76, 76, 'Found '+ CanNodes.count.ToString +' CAN Nodes on PCan interface:'+CHcan.Handle.ToHexString);
            for i := 0 to CanNodes.count-1 do
                LogStep(77,77, 77, #9#9'Node > '+ CanNodes.Strings[i] +'h');
        end
        else
            LogStep(78,78, 78, 'NO CAN Nodes Found on PCan interface:'+CHcan.Handle.ToHexString);
    end
    else begin
        // l' Errore è bloccante, quindi esco.
        LogStep(79,79, 79, 'ENUM CAN Nodes ERROR: '+ lastCAN_ErrMsg );
    end;
end;

procedure Toiac3Fasi2ax.cbxSelectAllClick(Sender: TObject);
var
  i: integer;
begin
    lviewCFG.Items.BeginUpdate;
    try
      for i := 0 to lviewCFG.Items.Count -1 do begin
          lviewCFG.Items.item[i].checked := cbxSelectAll.Checked;
      end
    finally
      lviewCFG.Items.EndUpdate;
    end;
end;

procedure Toiac3Fasi2ax.cbxShowGridClick(Sender: TObject);
begin
    if cbxShowGrid.checked then
        xlsReport.PrintGridLines := True
    else
        xlsReport.PrintGridLines := False;
    xlsPreview.InvalidatePreview;
end;

procedure Toiac3Fasi2ax.cbxHeadingsClick(Sender: TObject);
begin
    if cbxHeadings.checked then
        xlsReport.PrintHeadings := True
    else
        xlsReport.PrintHeadings := False;
    xlsPreview.InvalidatePreview;
end;

procedure Toiac3Fasi2ax.tkBar2Change(Sender: TObject);
begin
    tickPrescaler := tkBar2.Position;
end;

(*
procedure TframeFasi.tmrCanReadTimer(Sender: TObject);
var
  rcvdCob, RcvdLen, RcvdData: Ansistring;
  deg: single;
  val: word;
begin
    // Timer Lettura CAN Reference
    rcvdCob := '';
    rcvdLen := '';
    rcvdData := '';
    if CAN_ReadRawFrame(CHref, rcvdCob, RcvdLen, RcvdData) then
    	// The frame was successfully sent
        LogStep(80,80, 80, '<< RECEIVED CAN Frame,  COB:'+rcvdCob +'   Len:'+RcvdLen +'   Data:'+RcvdData);
	else
        LogStep(81,81, 81,  '* ERROR onTimer Receive CAN: '+ lastCAN_ErrMsg );
end;
*)

procedure Toiac3Fasi2ax.tmrCanReadTimer(Sender: TObject);
var
  rcvdCob : integer;
  RcvdLen : byte;
  RcvdData: ansiString;
  ivalX, ivalY: smallint;  // 2-byte signed integer.  NO integer !  produce valori sballati !
  tStr: ansiString;
begin
    // Timer Lettura CAN Reference.
    // letture con intervallo ampio 250 msec (x10 samples servono 2,5 sec.) per rendere stabile il posizionamento del DUT.
    // NB: l'intervallo time va dimensionato in modo da avere almeno il 10% di Read queue EMPTY
    //     così da garantire che la coda sia servita  in tempo reale, senza ritardi !
    //     Settando un Cyclic Interval del REF2 a 250 msec. il timer adeguato può essere 200 msec !
    //LogText('------------------------------------------------------------------------------');
    if CHdut.Handle = CHref.Handle then
        // gestione timer con 1 sola interfaccia PCAN.
        Repeat  // per svuotare eventuale coda e processare solo ultima Lettura !
          if CAN_ReadBinFrame(CHref, rcvdCob, RcvdLen, RcvdData) then begin
              // The CAN Read was successful
              {
              if vADMIN_mode then
                  LogText('<< RECEIVED CAN Frame,  COB:'+rcvdCob.ToHexString(3) +'   Len:'+RcvdLen.ToString +'   Data:'+dmBase.String2Hex(RcvdData))
              else
              }
              // filtro TPDO1 e gestisco REF2 frames...
              if rcvdCob = vRef2Addr + kBaseTPDOreceive then begin
                  WordRec(currXwRef).Lo := byte(RcvdData[1]);
                  WordRec(currXwRef).Hi := byte(RcvdData[2]);

                  WordRec(currYwRef).Lo := byte(RcvdData[3]);
                  WordRec(currYwRef).Hi := byte(RcvdData[4]);
                  //valT := byte(RcvdData[5]);                    // qui non serve

                  // NON applico 2's complement !
                  //ivalX := not(currXwRef) +1;
                  //ivalY := not(currYwRef) +1;
                  ivalX := currXwRef;   // perchè nel travaso da word a integer il segno viene considerato automaticamente !
                  ivalY := currYwRef;
                  //LogText(#9#9'Lon='+ ivalX.ToString +'   Lat='+ ivalY.ToString +'  T°='+valT.ToString); //x debug

                  currXRef := ivalX / Ref2_resolution;  // NB: devono essere dello stesso tipo
                  currYRef := ivalY / Ref2_resolution;  //     smallint o integer entrambi !!!  altrimenti danno risultati fuori scala !!(?)

                  // inversioni angolo REF, se richieste da opzione popmenù gauges
                  if vInvertX_ref then
                      currXRef := currXRef * -1;        // Necessario per REF invertire solo X !
                  if vInvertY_ref then
                      currYRef := currYRef * -1;        // Necessario per REF invertire solo Y !

                  edtRefX.text := format('%6.2f', [currXRef]);  //currXRef.ToString;
                  edtRefY.text := format('%6.2f', [currYRef]);  //currYRef.ToString;

                  mainForm.mmLog.seltext := ':';        // per evitare a user un log prolisso.
                  gaugeX.Value := currXRef;
                  gaugeY.Value := currYRef;
                //mainForm.mmLog.Lines.Add(#9+ format('%6.2f'+#9+'%6.2f', [currXRef, currYRef]));  //x Calibrazione Banco.
                //LogText(#9+ format('Ref2_X=%6.2f   Ref2_Y=%6.2f', [currXRef, currYRef]));  //x debug.

                  TInterlocked.BitTestAndSet(SampleReady, 0);  // Set SampleReady bit 0
                  // posso "set&forget" perchè nei relativi step campiono presenza sample a meno dei 250 msec del tmrCanRead !
              end//REF2 TPDO frames

              // filtro e gestisco DUT frames...
              else if rcvdCob = vDutAddr + kBaseTPDOreceive then begin
                  WordRec(currXwDut).Lo := byte(RcvdData[1]);
                  WordRec(currXwDut).Hi := byte(RcvdData[2]);

                  WordRec(currYwDut).Lo := byte(RcvdData[3]);
                  WordRec(currYwDut).Hi := byte(RcvdData[4]);
                  //valT := byte(RcvdData[5]);                    // qui non serve

                  // NON applico 2's complement !
                  //ivalX := not(currXwDut) +1;
                  //ivalY := not(currYwDut) +1;
                  ivalX := currXwDut;    // perchè nel travaso da word a integer il segno viene considerato automaticamente !
                  ivalY := currYwDut;
                  //LogText(#9#9'Lon='+ ivalX.ToString +'   Lat='+ ivalY.ToString +'  T°='+valT.ToString); //x debug

                  currXDut := ivalX / DUT_resolution;    // NB: devono essere dello stesso tipo
                  currYDut := ivalY / DUT_resolution;    //     smallint o integer entrambi !!!  altrimenti danno risultati fuori scala !!(?)

                  // inversioni angolo DUT, se richieste da opzione popmenù gauges
                  if vInvertX_dut then
                      currXDut := currXDut * -1;        // Necessario per DUT invertire solo X !
                  if vInvertY_dut then
                      currYDut := currYDut * -1;        // Necessario per DUT invertire solo Y !

                //edtDutX.text := format('%6.2f', [currXDut]);  //currXDut.ToString;  x DUT solo Gauges, perchè visualizzo mV.
                //edtDutY.text := format('%6.2f', [currYdut]);  //currYDut.ToString;

                  mainForm.mmLog.seltext := '.';        // per evitare a user un log prolisso.
                  gaugeX.Value := currXDut;
                  gaugeY.Value := currYDut;
                //mainForm.mmLog.Lines.Add(#9#9#9#9#9#9+ format('%6.2f'+#9+'%6.2f', [currXDut, currYDut]));  //x Calibrazione Banco.
                //LogText(#9#9#9#9#9#9+ format('Dut_X=%6.2f   Dut_Y=%6.2f', [currXDut, currYDut]));  //x debug.

                  TInterlocked.BitTestAndSet(SampleReady, 0);  // Set SampleReady bit 0
                  // posso "set&forget" perchè nei relativi step campiono presenza sample a meno dei 250 msec del tmrCanRead !
              end//DUT TPDO frames

              // filtro e gestisco DUT Emergency frames...
            //else if rcvdCob = vRef2Addr + kBaseEMCYreceive then begin   // nella realtà solo il DUT può dare Emergency frames,
              else if rcvdCob = vDutAddr +  kBaseEMCYreceive then begin   // perchè il REF può andare oltre il range del DUT !
                  // verifica se CobID ($80 + NID) è Emergency frame.
                  // indica problemi logici, non errori di protocollo
                  //var tStr: ansistring := leftStr(Data, 2); inline solo da delphi10.3 !
                  tStr := String2Hex( leftStr(RcvdData, 2) );
                  if dicEmergCode.TryGetValue( tStr, lastCAN_ErrMsg) then begin                             // se c'è il codice,
                      lastCAN_ErrMsg := 'CAN EmergencyCode: '+ String2Hex(RcvdData) +' = '+ lastCAN_ErrMsg; // assegna error msg,
                      // check e set della Flag per rilevazione letture range Out-of-Range ammesso
                      // per evitare validazioni con Dut fuori fondo scala e conseguenti calcoli errati!
                      if (tStr = '1050') or (tStr = '2050') then      // err X o Y rispettivamente.
                          vRangeError := True
                      else if (tStr = '0000') then
                          vRangeError := False;
                  end
                  else
                      lastCAN_ErrMsg := 'received Unexpected EmergencyCode >> '+ String2Hex(RcvdData);     // altrimenti default msg.

                  if vADMIN_mode then
                      LogText(#9'* WARNING onTimer Receive: '+ lastCAN_ErrMsg )
                  else
                  mainForm.mmLog.seltext := 'x';       // per evitare a user un log prolisso.

              end;//DUT EMCY frames...

          end
          else
              if lastCAN_Status <> kReceiveQueueEMPTY then                 // solo per evitare msg utile solo in debug,
                  LogText( '* ERROR onTimer Receive: '+ lastCAN_ErrMsg );  // visualizzando solo i significativi.
        until lastCAN_Status = kReceiveQueueEMPTY // svuoto eventuale coda e processo solo ultima Lettura !
    else
        // gestione timer con 2 interfacce PCAN.
        begin
            // gestisco frames da REF ch !
            Repeat  // per svuotare eventuale coda e processare solo ultima Lettura !
              if CAN_ReadBinFrame(CHref, rcvdCob, RcvdLen, RcvdData) then begin
                  // The CAN Read was successful
                  {
                  if vADMIN_mode then
                      LogText('<< RECEIVED CAN Frame,  COB:'+rcvdCob.ToHexString(3) +'   Len:'+RcvdLen.ToString +'   Data:'+dmBase.String2Hex(RcvdData))
                  else
                  }
                  // filtro TPDO1 e gestisco REF frames...
                  if rcvdCob = vRef2Addr + kBaseTPDOreceive then begin
                      WordRec(currXwRef).Lo := byte(RcvdData[1]);
                      WordRec(currXwRef).Hi := byte(RcvdData[2]);

                      WordRec(currYwRef).Lo := byte(RcvdData[3]);
                      WordRec(currYwRef).Hi := byte(RcvdData[4]);
                      //valT := byte(RcvdData[5]);                    // qui non serve

                      // NON applico 2's complement !
                      //ivalX := not(currXwRef) +1;
                      //ivalY := not(currYwRef) +1;
                      ivalX := currXwRef;   // perchè nel travaso da word a integer il segno viene considerato automaticamente !
                      ivalY := currYwRef;
                      //LogText(#9#9'Lon='+ ivalX.ToString +'   Lat='+ ivalY.ToString +'  T°='+valT.ToString); //x debug

                      currXRef := ivalX / Ref2_resolution;  // NB: devono essere dello stesso tipo
                      currYRef := ivalY / Ref2_resolution;  //     smallint o integer entrambi !!!  altrimenti danno risultati fuori scala !!(?)

                      // inversioni angolo REF, se richieste da opzione popmenù gauges
                      if vInvertX_ref then
                          currXRef := currXRef * -1;        // Necessario per REF invertire solo X !
                      if vInvertY_ref then
                          currYRef := currYRef * -1;        // Necessario per REF invertire solo Y !

                      mainForm.mmLog.seltext := ':';        // per evitare a user un log prolisso.
                      edtRefX.text := format('%6.2f', [currXRef]);  //currXRef.ToString;
                      edtRefY.text := format('%6.2f', [currYRef]);  //currYRef.ToString;

                      gaugeX.Value := currXRef;
                      gaugeY.Value := currYRef;
                    //LogText(#9#9+ format('Ref_X=%6.2f   Ref_Y=%6.2f', [currXRef, currYRef]));  //x debug.

                      TInterlocked.BitTestAndSet(SampleReady, 0);  // Set SampleReady bit 0
                      // posso "set&forget" perchè nei relativi step campiono presenza sample a meno dei 250 msec del tmrCanRead !
                  end;//REF2 TPDO frames

                  // filtro e gestisco REF Emergency frames...
                  // ma nella realtà solo il DUT può dare Emergency frames, perchè il REF può andare oltre il range del DUT !
                  //else if rcvdCob = vDutAddr +  kBaseEMCYreceive then begin
                  //end;//REF2 EMCY frames...

              end
              else
                  if lastCAN_Status <> kReceiveQueueEMPTY then                 // solo per evitare msg utile solo in debug,
                      LogText( '* ERROR onTimer REF Receive: '+ lastCAN_ErrMsg );  // visualizzando solo i significativi.
            until lastCAN_Status = kReceiveQueueEMPTY; // svuoto eventuale coda e processo solo ultima Lettura !

            // gestisco frames da DUT ch !
            Repeat  // su channel DUT per svuotare eventuale coda e processare solo ultima Lettura !
              if CAN_ReadBinFrame(CHdut, rcvdCob, RcvdLen, RcvdData) then begin
                  // The CAN Read was successful
                  {
                  if vADMIN_mode then
                      LogText('<< RECEIVED CAN Frame,  COB:'+rcvdCob.ToHexString(3) +'   Len:'+RcvdLen.ToString +'   Data:'+dmBase.String2Hex(RcvdData))
                  else
                  }
                  // filtro e gestisco DUT frames...
                  if rcvdCob = vDutAddr + kBaseTPDOreceive then begin
                      WordRec(currXwDut).Lo := byte(RcvdData[1]);
                      WordRec(currXwDut).Hi := byte(RcvdData[2]);

                      WordRec(currYwDut).Lo := byte(RcvdData[3]);
                      WordRec(currYwDut).Hi := byte(RcvdData[4]);
                      //valT := byte(RcvdData[5]);                    // qui non serve

                      // NON applico 2's complement !
                      //ivalX := not(currXwDut) +1;
                      //ivalY := not(currYwDut) +1;
                      ivalX := currXwDut;    // perchè nel travaso da word a integer il segno viene considerato automaticamente !
                      ivalY := currYwDut;
                      //LogText(#9#9'Lon='+ ivalX.ToString +'   Lat='+ ivalY.ToString +'  T°='+valT.ToString); //x debug

                      currXDut := ivalX / DUT_resolution;    // NB: devono essere dello stesso tipo
                      currYDut := ivalY / DUT_resolution;    //     smallint o integer entrambi !!!  altrimenti danno risultati fuori scala !!(?)

                      // inversioni angolo DUT, se richieste da opzione popmenù gauges
                      if vInvertX_dut then
                          currXDut := currXDut * -1;        // Necessario per DUT invertire solo X !
                      if vInvertY_dut then
                          currYDut := currYDut * -1;        // Necessario per DUT invertire solo Y !

                      mainForm.mmLog.seltext := '.';        // per evitare a user un log prolisso.
                    //edtDutX.text := format('%6.2f', [currXDut]);  //currXDut.ToString;  x DUT solo Gauges, perchè visualizzo mV.
                    //edtDutY.text := format('%6.2f', [currYdut]);  //currYDut.ToString;

                      gaugeX.Value := currXDut;
                      gaugeY.Value := currYDut;
                    //LogText(#9#9+ format('Dut_X=%6.2f   Dut_Y=%6.2f', [currXDut, currYDut]));  //x debug.

                      TInterlocked.BitTestAndSet(SampleReady, 0);  // Set SampleReady bit 0
                      // posso "set&forget" perchè nei relativi step campiono presenza sample a meno dei 250 msec del tmrCanRead !
                  end//DUT TPDO frames

                  // filtro e gestisco DUT Emergency frames...
                //else if rcvdCob = vRef1Addr + kBaseEMCYreceive then begin   // nella realtà solo il DUT può dare Emergency frames,
                  else if rcvdCob = vDutAddr +  kBaseEMCYreceive then begin   // perchè il REF può andare oltre il range del DUT !
                      // verifica se CobID ($80 + NID) è Emergency frame.
                      // indica problemi logici, non errori di protocollo
                      //var tStr: ansistring := leftStr(Data, 2); inline solo da delphi10.3 !
                      tStr := String2Hex( leftStr(RcvdData, 2) );
                      if dicEmergCode.TryGetValue( tStr, lastCAN_ErrMsg) then begin                             // se c'è il codice,
                          lastCAN_ErrMsg := 'CAN EmergencyCode: '+ String2Hex(RcvdData) +' = '+ lastCAN_ErrMsg; // assegna error msg,
                          // check e set della Flag per rilevazione letture range Out-of-Range ammesso
                          // per evitare validazioni con Dut fuori fondo scala e conseguenti calcoli errati!
                          if (tStr = '1050') or (tStr = '2050') then      // err X o Y rispettivamente.
                              vRangeError := True
                          else if (tStr = '0000') then
                              vRangeError := False;
                      end
                      else
                          lastCAN_ErrMsg := 'received Unexpected EmergencyCode >> '+ String2Hex(RcvdData);     // altrimenti default msg.

                      if vADMIN_mode then
                          LogText(#9'* WARNING onTimer DUT Receive: '+ lastCAN_ErrMsg )
                      else
                      mainForm.mmLog.seltext := 'x';       // per evitare a user un log prolisso.

                  end;//DUT EMCY frames...

              end
              else
                  if lastCAN_Status <> kReceiveQueueEMPTY then                 // solo per evitare msg utile solo in debug,
                      LogText( '* ERROR2 onTimer DUT Receive: '+ lastCAN_ErrMsg );  // visualizzando solo i significativi.
            until lastCAN_Status = kReceiveQueueEMPTY; // svuoto eventuale coda e processo solo ultima Lettura !
        end;
    {
    frameCrc32 := MAKELONG($1234, $5678);   // winapi
    WordRec(wd).Hi := $12;
    WordRec(wd).Lo := $34;
    LongRec(frameCrc32).Lo := $1234;
    LongRec(frameCrc32).Bytes[1] := $34;
    LongRec(frameCrc32).Bytes[3] := $78;

    //for i := 3 downto 0 do begin
    for i := 0 to 3 do begin
        LongRec(frameCrc32).Bytes[i] := Decoded_frame[Decoded_frameLen-4 +i];
    end;
    mmCom.Lines.Add('myCRC32 = '+ myCrc32.ToHexString(8));
    mmCom.Lines.Add('frameCrc32 = '+ frameCrc32.ToHexString(8));
    }
end;


/////////////////////////////////////////////////////////////////////////////////////////////////////

procedure Toiac3Fasi2ax.tmrFasiBlinkTimer(Sender: TObject);
begin
    // gestione blink dell'input pointer(icon) per le varie sequenze di input configurate.
    case currSequence of
      1: begin
           img1Seq1.Visible := not img1Seq1.Visible; // ovvero blink del control corrente.
           if vRangeError then
             lbFondoScala.Visible := not lbFondoScala.Visible;
         end;
//    2: imgSequence2.Visible := not imgSequence2.Visible;
//    3: imgSequence3.Visible := not imgSequence3.Visible;
      4: img1Seq4.Visible := not img1Seq4.Visible;
    end;
end;


//######################################################################################################################


// Carica in dictionary dicCanCmds la lista di Comandi PCAN prestabilita in file PCAN_CMDs.xlsx
function Toiac3Fasi2ax.Load_PCAN_CMDs_fromXLS(xlsFile: string): boolean;
var
  xls: TXlsFile;
  err, row, from: integer;
  rowCount, aLen: integer;
  cell: TCellValue;
  //addr: TCellAddress;
  tStr, alias, canIdx: string;
  CanCmd: TCanCmd;
  ListItem: TListItem;
begin
    result := False;
    xls := TXlsFile.Create( xlsFile );

    // Versione frame la visualizzo solo, perchè gestione troppo complessa
    LogStep(82,82, 82, 'Xls Keywords - FrameVersion: '+ xls.DocumentProperties.GetStandardProperty(TPropertyId.Keywords));    // Keywords = F1.01
    // Verifica compatibilità della frame con File xls dei comandi PCAN !
    tStr := xls.DocumentProperties.GetStandardProperty(TPropertyId.Category);   // Category = OIAC3
    LogStep(83,83, 83, 'Xls Category - Frame4Devices: '+ tStr);
    if IndexStr(tStr, kFrame4Devices) < 0 then begin
        LogSysEvent(svERROR, 408, 'EOL manager incompatibile con file Comandi PCan per device '+ tStr); // Category = OIAC3
        exit;
    end;

    try
      xls.ActiveSheetByName := 'PCAN_AllCMDs';  // We could loop over the existing sheets by using xls.SheetCount and xls.ActiveSheet
      rowCount := xls.RowCount;
      from := 2;                                // skip prima riga di testo.
      LogSysEvent(svINFO, 410, 'Aperto file Comandi CAN: '+ xlsFile);
      LogSysEvent(svINFO, 412, 'Trovati '+ intToStr(rowCount - from +1) +' comandi.');

      lviewCanCMDs.Clear;
      lviewCanCMDs.Items.BeginUpdate;
      lviewCanCMDs.Items.Clear;
      lviewCanCMDs.InnerListView.GridLines := True; // mostra grid, default è off.

      dicCanCmds.Clear;   // Clear removes all keys and values from the dictionary. The Count property is set to 0.
      canIdx := 'undef';                                // default value.
      err := 0;
      for row := from to rowCount do begin              // Skip 1st from-1 lines.
          // compila ListView con comandi CAN.
          try
              CanCmd := TCanCmd.Create;
              //addr := TCellAddress.Create(row, xls.ColFromIndex(row, colIndex));
              // colIndex := 1;  // to xls.ColCountInRow(row) do //Don't use xls.ColCount as it is slow: See http://www.tmssoftware.biz/flexcel/doc/vcl/guides/performance-guide.html#avoid-calling-colcount
              cell := xls.GetCellValue(row, 3);           // prima di tutt overifico presenza Alias.
              if cell.IsEmpty then begin                  // se manca, considero row di titolo da saltare,
                  cell := xls.GetCellValue(row, 1);       // ma di cui conservare l'index.
                  if cell.HasValue then                   // se Index presente
                      canIdx := cell.ToString;            // lo conservo e passo a prossimi subindex.
                  continue;
              end;

              ListItem := lviewCanCMDs.Items.Add;
            //ListItem.ImageIndex := -1;                    // per ora niente icon
              if cell.HasValue then
                  alias := cell.ToString;                   // Alias

              cell := xls.GetCellValue(row, 1);             // Can-Index
              if cell.HasValue then
                  CanCmd.Index := cell.ToString             // Index corrente.
              else
                  CanCmd.Index := canIdx;                   // Index da righe precedenti, perche siamo su dei subindex.
              ListItem.Caption := CanCmd.Index;

              cell := xls.GetCellValue(row, 2);             // Subindex
              if cell.HasValue then
                  CanCmd.SubIdx := cell.ToString
              else
                  CanCmd.SubIdx := '';
              ListItem.SubItems.Add( CanCmd.SubIdx );        // così non risulta mai nil !

              ListItem.SubItems.Add( alias );                // presenza Alias in col.3 già verificata.

              cell := xls.GetCellValue(row, 4);              // descrizione completa param
              if cell.HasValue then
                  CanCmd.Descrizione := cell.ToString
              else
                  CanCmd.Descrizione := '';
              ListItem.SubItems.Add( CanCmd.Descrizione );

              cell := xls.GetCellValue(row, 5);              // Data Type
              if cell.HasValue then
                  CanCmd.DataType := cell.ToString
              else
                  CanCmd.DataType := '';
              ListItem.SubItems.Add( CanCmd.DataType );
              if not dicDataType.TryGetValue(CanCmd.DataType, aLen) then begin
                  aLen := -1;                                // per indicare Len assente
                  LogStep(84,84, 84, 'Errore: CAN DataType NOT Found "'+ cell.ToString +'" in dictionary DataType.');
              end;
              CanCmd.paramLen := aLen;

              cell := xls.GetCellValue(row, 6);              // Accesso RO/RW
              if cell.HasValue then
                  CanCmd.Accesso := cell.ToString
              else
                  CanCmd.Accesso := '';
              ListItem.SubItems.Add( CanCmd.Accesso );

              cell := xls.GetCellValue(row, 7);              // Default value lo salvo anche in CanCmd, per eventuali usi.
              if cell.HasValue then
                  CanCmd.WrValue := cell.ToString
              else
                  CanCmd.WrValue := '';
              ListItem.SubItems.Add( CanCmd.WrValue );

              cell := xls.GetCellValue(row, 8);              // Range  (non serve in dictionary)
              if cell.HasValue then
                  ListItem.SubItems.Add( cell.ToString )
              else
                  ListItem.SubItems.Add('');

              cell := xls.GetCellValue(row, 9);             // Store
              if cell.HasValue then
                  CanCmd.StoreAs := cell.ToString
              else
                  CanCmd.StoreAs := '';
              ListItem.SubItems.Add( CanCmd.StoreAs );

              cell := xls.GetCellValue(row, 10);             // Extended value
              if cell.HasValue then
                  CanCmd.Extend := cell.ToString
              else
                  CanCmd.Extend := '';
              ListItem.SubItems.Add( CanCmd.Extend );

              // infine Insert nel Dictionary CanCmd con key = Alias
              dicCanCmds.add( alias, CanCmd);

              //ListItem.Caption := xls.GetCellValue(row, 12).ToString;
              //ListItem.SubItems.Add( xls.GetCellValue(row, 12).ToString );
              //if cell.HasValue then AngoloGarantito := strTointDef(cell.ToString, 0); // ok anche per celle General !
              //                      AngoloGarantito := trunc(cell.asNumber);          // solo per celle Number
              //if cell.HasValue then ErrMaxTest      := cell.asNumber;
          except
            on E: Exception do begin
              LogStep(85,85, 85,  'EXCEPTION in Load PCAN Cmd from XLS >> '+ E.ClassName +#10#13#9+ ' description: '+E.Message);
              LogSysEvent(svERROR, 413, 'XLS Unknown cell type in row '+ row.ToString  +' of tab '+xls.ActiveSheetByName);   //
              inc(err);
            end;
          end;
      end;//for row

      if err = 0 then
          result := True;  // se arriva fin qui, ovvero senza exceptions, il caricamento è ok.
    finally
      lviewCanCMDs.Items.EndUpdate;
      xls.Free;
      dicCanCmds.TrimExcess;
    end;

    // print lista per debug.
  //for alias in dicCanCmds.Keys do
  //    LogText(dicCanCmds.Items[alias].Index + ':'+#9+ alias );
end;

// Carica da file XLS di Configurazione PCAN_CFG_Optoi (o specifico per cliente) la lista dei Parametri da configurare nel DUT.
// NB: viene usata in 2 fasi, inizio e fine con parametri caricati diversi !
function Toiac3Fasi2ax.Load_CFG_CmdSet_fromXLS(xlsFile: string): boolean;
var
  xls: TXlsFile;
  err, i, r, row, from: integer;
  rowCount, aLen: integer;
  cell: TCellValue;
  //addr: TCellAddress;
  value, store, alias, canIdx: string;
  cfgListItem: TListItem;
  CanCmd: TCanCmd;
begin
    result := False;
    err := 0;
    xls := TXlsFile.Create( xlsFile );  // file xls di Default Optoi o Custom cliente.
    LogSysEvent(svINFO, 450, 'Aperto file Parametri CFG: '+ xlsFile);
    try
      xls.ActiveSheetByName := 'PCAN_CfgCMDs';  // We could loop over the existing sheets by using xls.SheetCount and xls.ActiveSheet
      LogStep(86,86, 86, 'Ordinamento file Parametri secondo WR Priority...');
      // ordina secondo campo WR Priority per poter scrivere i parametri nel DUT in ordine diverso dal deafult (Index+Subidx)
      // per i casi in cui i valori ammessi da certi parametri dipendono dal valore già inserito in altri.
      xls.Sort(
         TXlsCellRange.Create('A2:P150'), // no modo di ordinare the full sheet, perchè TXlsCellRange.Null dà runtime error !
         true,                            // Normally we want to sort the rows in the range. But we might want to sort columns and we can do it by setting this to false.
         Int32Array.Create(11),           // Sort by columns K (11)
         TArray<TSortOrder>.Create(TSortOrder.Ascending), //Sort column K ascending. If you want to sort everything in ascending order, just set this to nil.
         nil,                             // Use the standard comparer.
         TSortFormulaMode.ExcelLike       // Excel-like is faster, but won't change formulas outside the range being sorted.
      );                                  // Range A2:P150 sufficiente a contenere anche future nuove colonne e righe !

      rowCount := xls.RowCount;
      from := 2;                                // skip prima riga di testo.
      LogSysEvent(svINFO, 452, 'Contiene '+ intToStr(rowCount - from +1) +' parametri CAN.'); // non esatto per via di righe vuote nel file xls...

      lviewCFG.Clear;
      lviewCFG.Items.BeginUpdate;
      lviewCFG.Items.Clear;
      lviewCFG.InnerListView.GridLines := True; // mostra grid, default è off.

      canIdx := 'undef';                                // default value.
      for row := from to rowCount do begin              // Skip 1st from-1 lines.
          // compila ListView con comandi CAN.
          try
              cell := xls.GetCellValue(row, 3);           // prima di tutto verifico presenza Alias e
              if cell.IsEmpty then begin                  // se manca, considero row di titolo da saltare.
                  continue;
              end;
              alias := cell.ToString;                     // ok, Alias definito.
              cell := xls.GetCellValue(row, 1);           // poi verifico presenza Index e
              if cell.IsEmpty then begin                  // se manca, considero errore...
                  LogStep(87,87, 87, 'Errore: CAN command Index NOT Found for Alias "'+ alias +'"');
                  inc(err);
                  exit;
              end;
              canIdx := cell.ToString;                    // ok, canIdx definito.

              // Verifico sempre di caricare parametri già previsti in PCAN_CMDs.xlsx
              // che è il Master comandi (e Alias) cui fare sempre riferimento !
              if not dicCanCmds.TryGetValue( alias, CanCmd ) then begin
                  LogStep(88,88, 88, 'Errore: CAN command Alias NOT Found "'+ alias +'" in dictionary CanCmds.');
                  inc(err);
                  exit;
              end;

              // Verifica se è parametro da caricare come configurazione DEF, CFG1ax, CFG2ax.
              cell := xls.GetCellValue(row, 9);           // poi verifico Store column (DEF, CFG1ax, CFG2ax)
              if cell.IsEmpty then                        // se vuota, è row da saltare subito.
                  continue;
              store := cell.ToString;
              if (store <> 'DEF') and (store <> 'DEF1') and (store <> 'DEF2') and (store <> 'END') and
                 (store <> 'SNUM') and (store <> 'CFG1') and (store <> 'CFG2') then
                  continue;
                  // skip, perchè non sono parametri da configurare.

              if (store = 'DEF') and not cbxListaDefaultsCom.checked then
                  continue;
                  // skip, perchè non è parametro da configurare come default comune.
              if (store = 'END') and not cbxListaDefaultsEND.checked then
                  continue;
                  // skip, perchè non è parametro da configurare come default comune a Fine EOL.
              if (store = 'DEF1') and not cbxListaDefaults1ax.checked then
                  continue;
                  // skip, perchè non è parametro da configurare come default in device 1asse.
              if (store = 'DEF2') and not cbxListaDefaults2ax.checked then
                  continue;
                  // skip, perchè non è parametro da configurare come default in device 2asse.
              if (store = 'SNUM') and not cbxListaSN.checked then
                  continue;
                  // skip, perchè non è parametro da configurare come serial number.
              if (store = 'CFG1') and not cbxListaCalib1ax.checked then
                  continue;
                  // skip, perchè non è parametro da configurare come calib 1asse.
              if (store = 'CFG2') and not cbxListaCalib2ax.checked then
                  continue;
                  // skip, perchè non è parametro da configurare come calib 2assi.

              // E' parametro di configurazione (DEF, END, CFG1ax, CFG2ax, SNUM) quindi lo carico in Listview.
              cfgListItem := lviewCFG.Items.Add;            // prepara list item.
              cell := xls.GetCellValue(row, 1);             // Can-Index
              if cell.HasValue then
                  cfgListItem.Caption := cell.ToString      // Index.
              else
                  cfgListItem.Caption := canIdx;            // Index da righe precedenti, perche siamo su dei subindex.
            //cfgListItem.ImageIndex := -1;                 // qui niente icon
              cfgListItem.Checked := True;                  // marca record che va usato per store default values.
              cfgListItem.ImageIndex := -1;                 // per ora niente icon

              cell := xls.GetCellValue(row, 2);             // Subindex.
              if cell.HasValue then
                  cfgListItem.SubItems.Add( cell.ToString )
              else
                  cfgListItem.SubItems.Add('');

              cell := xls.GetCellValue(row, 4);             // Descrizione param.
              if cell.HasValue then
                  cfgListItem.SubItems.Add( cell.ToString )
              else
                  cfgListItem.SubItems.Add('');

              // Write value
              // Gestione specifica del Write-Value quando riguarda parametri di Calibrazione appena inseriti !
              if store = 'CFG2' then begin                  // param 2 Assi
                  // nel Dictionary dicCanCmds, per il cmdAlias corrente ho a disposizione i valori già assegnati in fase di Calibrazione,
                  // che ora posso assegnare in ListviewCFG come Write value, che verranno così verificati in prossimo step lettura params!
                  // Copia in listview il valore trovato in calibrazione, così
                  // da poterlo verificare in prossimi step di lettura sequenziale parametri dal DUT !
                  // ed eventualmente riprovare il Write.
                  value := CanCmd.WrValue;                  // NB: il campo CanCmd.WrValue viene valorizzato solo in fase Calibrazione!
              end
              else if store = 'CFG1' then begin             // param 1 Asse
                  // cerca l'Alias previsto da CFG1
                  // qui non usato.
                  value := '';                              //
              end
              else if store = 'SNUM' then begin             // Serial Number
                  // è un parametro 'SNUM' per cui prendo valori da record Sessione.
                  value := format( '%2.2x%6.6xh', [Sessione.Serial.ToInteger, Sessione.Lotto.ToInteger]);
                  CanCmd.WrValue := value;                  // ma tengo copia anche in dictionary dicCanCmds (per CFG2 c'è già).
              end
              else begin
                  // è un parametro 'DEF','DEF1','DEF2','END' Default quindi prendo value dal campo XLS.
                  cell := xls.GetCellValue(row, 7);         // Write value
                  if cell.HasValue then
                      value := cell.ToString
                  else
                      value := '';                          // così risulta empty, ma non nil !
                  CanCmd.WrValue := value;                  // ma tengo copia anche in dictionary dicCanCmds (per CFG2 c'è già).
              end;
              // infine carica il WrValue nella listview. Forse sarebbe utile avere 2 colonne per WrValue, una iniziale
              cfgListItem.SubItems.Add( value );            // per procedure EOL e una finale da applicare alla fine (?)

              // Data Type
              cell := xls.GetCellValue(row, 5);             // Data Type
              if cell.HasValue then
                  cfgListItem.SubItems.Add( cell.ToString )
              else
                  cfgListItem.SubItems.Add('');             // così risulta empty, ma non nil !
              if not dicDataType.TryGetValue(cell.ToString, aLen) then begin
                  LogStep(89,89, 89, 'Errore: CAN DataType NOT Found "'+ cell.ToString +'" in dictionary DataType.');
                  inc(err);
                  exit;
              end;

              // Read value, vuoto
              i := cfgListItem.SubItems.Add( '' );          // Read value, vuoto
              cfgListItem.SubItemImages[4] := 20;           // icon su colonna 4, ReadValue.
            //cfgListItem.ImageIndex := 20;                 // aggiorna icon su prima colonna.
              //cfgListItem.SubItem[]  purtroppo alignment è solo per Columns !

              // Range
              cell := xls.GetCellValue(row, 8);              // Range
              if cell.HasValue then
                  cfgListItem.SubItems.Add( cell.ToString )
              else
                  cfgListItem.SubItems.Add('');

              // Default value
              cell := xls.GetCellValue(row, 7);              // Default value = Write value
              if cell.HasValue then
                  value := cell.ToString
              else
                  value := '';                               // così risulta empty, ma non nil !
              cfgListItem.SubItems.Add( value );

              // Alias
              cfgListItem.SubItems.Add( alias );             // presenza Alias in col.3 già verificata.

              // Store
              cell := xls.GetCellValue(row, 9);             // Store
              if cell.HasValue then
                  cfgListItem.SubItems.Add( cell.ToString )
              else
                  cfgListItem.SubItems.Add('');

              // Extend
              cell := xls.GetCellValue(row, 10);             // Ext
              if cell.HasValue then
                  cfgListItem.SubItems.Add( cell.ToString )
              else
                  cfgListItem.SubItems.Add('');

              inc(r);
              //ListItem.Caption := xls.GetCellValue(row, 12).ToString;
              //ListItem.SubItems.Add( xls.GetCellValue(row, 12).ToString );
              //if cell.HasValue then AngoloGarantito := strTointDef(cell.ToString, 0); // ok anche per celle General !
              //                      AngoloGarantito := trunc(cell.asNumber);          // solo per celle Number
              //if cell.HasValue then ErrMaxTest      := cell.asNumber;
          except
            on E: Exception do begin
              LogStep(90,90, 90,  'EXCEPTION in Load CFG CmdSet from XLS >> '+ E.ClassName +#10#13#9+ ' description: '+E.Message);
              LogSysEvent(svERROR, 460, 'XLS Unknown cell type in row '+ row.ToString  +' of tab '+xls.ActiveSheetByName);   //
              inc(err);
            end;
          end;
      end;//for row

      if err = 0 then
          result := True;  // se arriva fin qui, ovvero senza exceptions, il caricamento è ok.
    finally
      lviewCFG.Items.EndUpdate;
      LogSysEvent(svINFO, 470, 'Caricati '+ r.ToString +' parametri di Configurazione.');   //
      xls.Free;
    end;
end;

procedure Toiac3Fasi2ax.TestsortDict1Click(Sender: TObject);
var
  sortDic: TSPairs<TTestAngle, TSessionTests>;
  Item: TPair<TTestAngle, TSessionTests>;
  Item2: TPair<TCalibAngle, TSessionCalibra>;
  TestAngle  : TTestAngle;
  CalibAngle : TCalibAngle;
  testRecord: TSessionTests;
  calRecord: TSessionCalibra;
begin

    CalibRecords.Clear;   // Clear removes all keys and values from the dictionary. The Count property is set to 0.
    for CalibAngle := Low(TCalibAngle) to High(TCalibAngle) do begin
        calRecord := TSessionCalibra.create;
        calRecord.degPairID := strCalibAngle[ord(CalibAngle)];
        calRecord.RefAx := Random * 60;
        calRecord.RefAy := Random * 60;
        calRecord.DutAx_mV := trunc(Random * 45000);
        calRecord.DutAy_mV := trunc(Random * 45000);
        CalibRecords.Add( CalibAngle, calRecord );
    end;

    for Item2 in CalibRecords do begin
        mainForm.mmLog.Lines.Add( 'Key > ' + strCalibAngle[ord( item2.Key )]);
        mainForm.mmLog.Lines.Add( #9'Value x: ' + floatToStr(item2.Value.RefAx));
    end;

    CalibRecords.Clear;
    for CalibAngle := Low(TCalibAngle) to High(TCalibAngle) do begin
        calRecord := TSessionCalibra.create;
        calRecord.degPairID := strCalibAngle[ord(CalibAngle)];
        calRecord.RefAx := Random * 60;
        calRecord.RefAy := Random * 60;
        calRecord.DutAx_mV := trunc(Random * 45000);
        calRecord.DutAy_mV := trunc(Random * 45000);
        CalibRecords.Add( CalibAngle, calRecord );
    end;
    CalibRecords.Clear;

    CalibRecords.Clear;

    for Item2 in CalibRecords do begin
        mainForm.mmLog.Lines.Add( 'Key > ' + strCalibAngle[ord( item2.Key )]);
        mainForm.mmLog.Lines.Add( #9'Value x: ' + floatToStr(item2.Value.RefAx));
    end;

exit;



    // mostra enums as string
    for TestAngle := Low(TTestAngle) to High(TTestAngle) do
        mainForm.mmLog.Lines.Add( 'TestAngle > ' + strTestAngle[ord(TestAngle)]);
    for CalibAngle := Low(TCalibAngle) to High(TCalibAngle) do
        mainForm.mmLog.Lines.Add( 'CalibAngle > ' + strCalibAngle[ord(CalibAngle)]);

    // crea Dictionary con enum+record
    // per ogni TestAngle(enum) collega un TestRecord
    sortDic := TSPairs<TTestAngle, TSessionTests>.Create();
    for TestAngle := Low(TTestAngle) to High(TTestAngle) do begin
        testRecord := TSessionTests.create;
        testRecord.degPairID := strTestAngle[ord(TestAngle)];
        sortDic.Add( TestAngle, testRecord);
    end;

(*
    Pairs.Add('Key2', 'Text2');
    Pairs.Add('Key3', 'Text3');
    Pairs.Add('Key4', 'Text4');
    Pairs.Add('Key5', 'Text5');
    for Item in Pairs do begin
        mainForm.mmLog.Lines.Add( item.Value.AsseCorrente);
    end;
*)
    // Loop su TSessionTests[] items per print records Asse TEST X e TEST Y (con relativi Cross).
    for Item in sortDic do begin
        mainForm.mmLog.Lines.Add( 'Test Angle > ' + strTestAngle[ord( item.Key )]);
        mainForm.mmLog.Lines.Add( #9'Test record: ' + item.Value.degPairID);
    end;

    // scorre il Dictionary
    for Item in sortDic do begin
        mainForm.mmLog.Lines.Add( 'Pair TestAngle: ' + item.Value.degPairID);
    end;

    mainForm.mmLog.Lines.Add( sortDic.Items[5].Value.degPairID );  // get by index

    if sortDic.ContainsKey( deg15p ) then begin
        testRecord := sortDic.GetValue( deg15p );                  // get by value
        mainForm.mmLog.Lines.Add( 'Found TestAngle: ' + testRecord.degPairID);
    end
    else
        mainForm.mmLog.Lines.Add( 'deg15p NOT Found!');
end;

function Toiac3Fasi2ax.LoadCarattBanco_fromXLS(xlsFile: string; testAxis:char): boolean;
var
  xls: TXlsFile;
  row, from: integer;
  deg, rowCount: integer;
  errCount: integer;
  cell: TCellValue;
//addr: TCellAddress;
  ListItem: TListItem;
  CorrBanco: TCorrBanco;
  valX, valY: single;
const
  kRowCount = 15;
begin
    // caricamento va fatto in funzione di quale asse è Test, scambiando le colonne restituite
    result := False;
    // verifica presenza file
    if dmBase.FindFile( xlsFile ) then
        LogSysEvent(svINFO, 1000, 'Found parameters file: '+ xlsFile)     // log in queue, fino al primo Dequeue!
    else begin
        LogSysEvent(svERROR, 1003, 'Params file NOT found: '+ xlsFile);
        exit;
    end;

    // carica dal foglio previsto
    xls := TXlsFile.Create( xlsFile );
    try
      xls.ActiveSheetByName := 'TAB.ERR';  //we'll read sheet3. We could loop over the existing sheets by using xls.SheetCount and xls.ActiveSheet
      rowCount := xls.RowCount;
      LogSysEvent(svINFO, 1005, 'XLS Aperto foglio: '+ xls.ActiveSheetByName);

      lviewCarattBanco.Clear;
      lviewCarattBanco.Items.BeginUpdate;
      lviewCarattBanco.Items.Clear;
      lviewCarattBanco.InnerListView.GridLines := True;           // mostra grid, default è off.
      // Assegna labels colonne a seconda dell'asse in Test.
      if testAxis = 'X' then begin
          lviewCarattBanco.Column[1].Caption := 'Test X';
          lviewCarattBanco.Column[2].Caption := 'Cross Y';
          from := 3;
      end
      else begin
          lviewCarattBanco.Column[1].Caption := 'Test Y';
          lviewCarattBanco.Column[2].Caption := 'Cross X';
          from := 20;
      end;

      errCount := 0;
      for row := from to (from + kRowCount) -1 do begin   // indirizza range di 15 righe a partire da From.
          // compila ListView con Valori caratterizzazione banco.
          try
              CorrBanco := TCorrBanco.Create;
              //addr := TCellAddress.Create(row, xls.ColFromIndex(row, colIndex));
              // colIndex := 1;  // to xls.ColCountInRow(row) do //Don't use xls.ColCount as it is slow: See http://www.tmssoftware.biz/flexcel/doc/vcl/guides/performance-guide.html#avoid-calling-colcount
              cell := xls.GetCellValue(row, 1);           // Angolo in gradi.
              if cell.IsEmpty then begin                  // se manca, errore quindi esco.
                  inc(errCount);
                  break;
              end;
              ListItem := lviewCarattBanco.Items.Add;
            //ListItem.ImageIndex := -1;                  // per ora niente icon
              ListItem.Caption := cell.ToString;          // NB:  ToString <> AsString  !
              deg := StrToInt(ListItem.Caption);          // deg (+70..0..-70)
              CorrBanco.angolo := deg;

              cell := xls.GetCellValue(row, 2);           // colIndex := 2;  // to xls.ColCountInRow(row) do //Don't use xls.ColCount as it is slow: See http://www.tmssoftware.biz/flexcel/doc/vcl/guides/performance-guide.html#avoid-calling-colcount
              if cell.IsEmpty then begin                  // se manca, errore quindi esco.
                  inc(errCount);
                  break;
              end;
              //ListItem.SubItems.Add( cell.ToString );
              valX := cell.asNumber;

              cell := xls.GetCellValue(row, 3);
              if cell.IsEmpty then begin                  // se manca, errore quindi esco.
                  inc(errCount);
                  break;
              end;
              //ListItem.SubItems.Add( cell.ToString );
              valY := cell.asNumber;

              // Assegna valori campi a seconda dell'asse in Test.
              if testAxis = 'X' then begin
                  CorrBanco.asseTest := valX;             // Test X
                  CorrBanco.asseCross := valY;            // Cross Y
              end
              else begin // 'Y'
                  CorrBanco.asseTest := valY;             // Test Y
                  CorrBanco.asseCross := valX;            // Cross X
              end;

              ListItem.SubItems.Add( CorrBanco.asseTest.ToString );
              ListItem.SubItems.Add( CorrBanco.asseCross.ToString );

              arrCorrezioneBanco[row - from +1] := CorrBanco;     // aggiungi record all'array.
          except
            on E: Exception do begin
              LogStep(91,91, 91,  'EXCEPTION in Load Caratterizzazione banco >> '+ E.ClassName +#10#13#9+ ' description: '+E.Message);
              LogSysEvent(svERROR, 1010, 'XLS Unknown cell type in row '+ row.ToString  +' of tab '+xls.ActiveSheetByName);
            end;
          end;
      end;//for

      if errCount > 0 then
          LogSysEvent(svINFO, 1011, 'XLS Errore in caricamento riga '+ row.ToString +' da sheet '+xls.ActiveSheetByName)
      else
          result := True; // se arriva fin qui, ovvero senza exceptions, il caricamento è ok.

    finally
      lviewCarattBanco.Items.EndUpdate;
      xls.Free;
      LogSysEvent(svINFO, 1012, 'XLS caricati '+ intToStr(row-from) +' Coefficienti di Correzione.');   //
    end;

    // print lista per debug.
    for deg := 1 to length(arrCorrezioneBanco) do begin
      LogText('Corr.Banco: '+ arrCorrezioneBanco[deg].angolo.ToString );
    end;

end;

function Toiac3Fasi2ax.calcWorkTimeFasi(begFase,endFase: string): TDateTime;
var
  i: integer;
  aJobRecord: TJobRecord;
  tBeg, tEnd: TDateTime;
begin
    LogStep(92,92, 92, 'Calcolo WorkTime '+ begFase +' e '+ endFase);
    i := JobList.IndexOf( begFase );
    if i >= 0 then begin
        aJobRecord := TJobRecord( JobList.Objects[ i ]);
        tBeg := aJobRecord.BeginStamp;
        tEnd := tBeg;   // provvisorio.
        // trovato inizio Fase di partenza
        // e da qui proseguo scorrendo tutti i restanti steps
        for i := aJobRecord.ListIndex to JobList.Count-1 do begin
            aJobRecord := TJobRecord( JobList.Objects[ i ]);
            if aJobRecord.GroupFase = endFase then
                tEnd := aJobRecord.EndStamp;
        end;
        // perciò alla fine avrò l'ultimo EndStamp della Fase finale
        // che mi permette di fare il cacolo Worktime preciso, ovvero
        // la differenza fra inizio Setup e termine dell'ultimo step
        result := tBeg-tEnd;
    end
    else begin
        LogStep(93,93, 93, 'ERR in calcolo WorkTime, '+ begFase +' non trovata!');
        result := NullDate;   // FormatDateTime('dd.mm.yyyy',-693594) = 00.00.0000
    end;
end;

function Toiac3Fasi2ax.SessionUpdateDB: boolean;
begin
    result := False;
    if not dmBase.pgConn.Connected then begin
        LogStep(94,94, 94, #9+'Errore in <SessionUpdateDB>: Database NON Connesso !');
        exit;
    end;
    try (*
        default isolation levels.
        in FireDAC the property value is xiUnspecified:
            DB2                     - xiReadCommitted
            InterBase and Firebird  - xiSnapshot
            MySQL and MariaDB       - xiRepeatableRead
            Oracle                  - xiReadCommitted
            Microsoft SQL Server    - xiReadCommitted
            SQLite                  - xiSerializible
            PostgreSQL              - xiReadCommitted
        http://docs.embarcadero.com/products/rad_studio/firedac/frames.html?frmname=topic&frmfile=Managing_Transactions.html
        *)
        // Prima di tutto verifica se esistono già Test Records per questo S/N e domanda se aggiungere...
        dmBase.pgConn.StartTransaction;  // gestione transazione usando try-except !
        try
            LogStep(95,95, 95, 'Aggiornamento dati Sessione EOL per S/N: '+ Sessione.Serial+'  su DB Produzione.' );
            // Deve esistere già un record Master perchè qui aggiorno solo timeStamp e counters
            with dmBase.tabSessionDut do begin
                // Cerco se esiste almeno una sessione di collaudo per questo [CodProd + Lotto + S/N + RecMode]
                LogStep(96,96, 96, 'Locate session Record...');
                CheckBrowseMode;
                if Locate('CODICEPRODOTTO;LOTTO;SERIAL;RECMODE', VarArrayOf([Sessione.CodiceProdotto, Sessione.Lotto, Sessione.Serial, Sessione.RecMode]), []) then begin
                    // andrebbe gestito se + di 1 record (versione per stesso s/n)
                    // ricerca versione va fatta usando in + RECORDSTAMP.
                    // Update Master Sessione
                    (*
                    "PK"                integer NOT NULL DEFAULT nextval('"Produzione"."DUTSESSIONS_PK_seq"'::regclass),
                    "WORKPLACE"         character varying(20) COLLATE pg_catalog."default",
                    "IDOPERATORE"       character varying(28) COLLATE pg_catalog."default"
                    "STARTSTAMP"        timestamp without time zone,
                    "ENDSTAMP"          timestamp without time zone,
                    "WORKTIME"          time without time zone,
                    "PAUSETIME"         double precision,
                    "MEMO"              text COLLATE pg_catalog."default",
                    *)
                    Edit;
                    LogStep(97,97, 97, 'Edit master tabSession...');
                    //
                    // Aggiornamento tempi preliminari con definitivi e salvataggio su DB.
                    // Solo ora che la Sessione è conclusa ho i dati definitivi da memorizzare su DB !
                    // quindi solo un update del record Master, perchè i detail sono già inseriti negli step finali...
                    Sessione.EndStamp := now;
                    Sessione.WorkTime := Sessione.EndStamp - Sessione.StartStamp;  // Valore aggiornato a reale fine Sessione.
                    //FieldByName('STARTSTAMP').AsDateTime := Sessione.StartStamp; // inizio collaudo già inserito.
                    FieldByName('ENDSTAMP').AsDateTime   := Sessione.EndStamp;     // valore definitivo di fine collaudo.
                    FieldByName('WORKTIME').AsFloat      := Sessione.WorkTime;     // tempo impiegato   AsSQLTimeStampOffset;
                    //FieldByName('PAUSETIME').AsSingle  := Sessione.Pausetime;    // tempo rimasta in pausa by user. hh:mm:ss
                    FieldByName('MEMO').AsString         := Sessione.Memo;         // commenti e attributi non codificati,
                    Post;                                                          // da userinput in memo nell'ultima fase ?
                end
                else begin
                    // errore Record Master non trovato, ma non posso fare molto se non rinunciare all'update e loggare...
                    dmBase.pgConn.Rollback;
                    dmBase.dmLogSysEvent(svERROR, 530, 'ERROR: DB Master record della Sessione non trovato; impossibile aggiornare !');
                    exit;
                end;
            end;
            dmBase.pgConn.Commit;
            LogStep(98,98, 98, 'OK dati di Sessione aggiornati.');
            result := True;
        except
          on E: Exception do begin
              dmBase.dmLogSysEvent(svEXCEPTION, 540, ' in Update Session record >> '+ E.ClassName +#10#13#9+ ' description: '+E.Message);
          end;

        end;
    finally
        if dmBase.pgConn.InTransaction then     // se trova pendenze impreviste
            dmBase.pgConn.Rollback;             // meglio Rollback !  finally è eseguito anche in caso di except !
    end;
end;


procedure Toiac3Fasi2ax.LoadCorrBanco1Click(Sender: TObject);
var
  sourceFile: string;
begin
    sourceFile := kAppPath + dmBase.vSessionSource;
    LogStep(99,99, 99, 'Loading Coefficienti di Correzione X from : '+ sourceFile);
    LoadCarattBanco_fromXLS( sourceFile, 'X' );
end;

procedure Toiac3Fasi2ax.LoadCorrBancoY1Click(Sender: TObject);
var
  sourceFile: string;
begin
    sourceFile := kAppPath + dmBase.vSessionSource;
    LogStep(100,100, 100, 'Loading Coefficienti di Correzione Y from : '+ sourceFile);
    LoadCarattBanco_fromXLS( sourceFile, 'Y' );
end;


//#########################  MENU Pulsanti per DEBUG  ###############################

procedure Toiac3Fasi2ax.SetCalibintvalue1Click(Sender: TObject);
var
  Can: TCanFrame;
  i: integer;
begin
    // Prepara la frame con indirizzi e Dati da inviare
    Can.RTR := False;
    Can.NodeID := edtCanID.text;
    Can.Index  := edtIndex.text;
    Can.Subidx := edtSubIdx.text;
    i := 33777;   // xtest
    Can.Datatype := 'VSTR';
    Can.Data := cbxCanDataStrPrefix.text;
    Can.Data := Can.Data + AnsiChar(Hi( i )) + AnsiChar(Lo( i )); // NB: usare Ansichar() e non char() o chr() che tronca a 7 bit !!!
    // invio can frame
    if CAN_SendFrame(Can) then
        // The frame was successfully sent
        LogStep(101,101, 101, 'CAN Frame SENT Successfully,  ID:'+Can.CobID +'   LenPL:'+ intToStr(Can.LenPayLoad) +'   Frame:'+ hexLarge(Can.Frame))
    else
        LogStep(102,102, 102,  lastCAN_ErrMsg );
end;

procedure Toiac3Fasi2ax.ReadallCalib1Click(Sender: TObject);
var
  CalDeg: TCalibAngle;  //  TCalibAngle = ( X0_Y0, X0_Yp60, X0_Yn60, Xp60_Y0, Xn60_Y0 );  // i valori precisi poi dipendono da tabella di calibrazione specifica per famiglia di devices.
  Can: TCanFrame;
  iValueReadX, iValueReadY: integer;
  kTargetRefX: single;
  kTargetRefY: single;
  CanCmd1, CanCmd2: string;
  val1, val2: string;
  ListItem: TListItem;
begin
    // prima di tutto assicura CAN channel "clean"
    if not CAN_CleanCH(CHdut) then begin
        // l' Errore è bloccante, quindi esco.
        LogStep(103,103, 103,  lastCAN_ErrMsg );
        exit;
    end;
    LogStep(104,104, 104, sLineBreak+'STATUS PCAN_OK');
    // The channel is Ready to work.

    // Reset DUT necessario per portare in RAM i parametri, nel caso siano appena stati calibrati (settati in eeprom)!
    if CAN_SendNMT( CHdut, cmdResetNode, edtDUTaddr.text ) then begin
        // The Reset was successfully sent
        LogStep(105,105, 105, 'CAN ResetNode SENT Successfully to NodeID:'+ edtDUTaddr.text );
    end
    else begin
        // l' Errore è bloccante, quindi esco.
        LogStep(106,106, 106, 'CAN ResetNode ERROR  on NodeID:'+ edtDUTaddr.text );
        LogStep(107,107, 107,  lastCAN_ErrMsg );
        exit;
    end;
    // DUT reset & boot completed !
    if not CAN_FlushRxBuffer(CHdut) then begin
        LogStep(108,108, 108,  lastCAN_ErrMsg );
        exit;
    end;
    LogStep(109,109, 109, 'OK FLUSH Can queue Successful.');    // preparato rx queue clean.

    // Sblocca area protetta 5555h del DUT per precauzione, se non fosse già da setup fase.
    if not CAN_UnlockParams( edtDUTaddr.text ) then begin
        // l' Errore è bloccante, quindi esco.
        LogStep(110,110, 110, 'ERROR on UNLOCK params for DUT:'+ edtDUTaddr.text );
        LogStep(111,111, 111,  lastCAN_ErrMsg );
        exit;     // unlock non riuscito, quindi non serve lock finally !
    end;
    LogStep(112,112, 112, 'OK, DUT Unlock command CONFIRMED by NodeID:'+ edtDUTaddr.text );

    lviewCalibVal.Items.BeginUpdate;
    try
        lviewCalibVal.Items.Clear;
        for CalDeg := X0_Y0 to YN_Xyn do begin
            // Valori come da tabella calibrazione per DUT corrente
            // al momento fissi, ma possibile da tabella xls.
            case CalDeg of
                  X0_Y0: begin
                           kTargetRefX := 0;
                           kTargetRefY := 0;
                           CanCmd1 := _2ax_calX0_deg_;       // Campo per calibrazione asse X del DUT.   5555-1
                           CanCmd2 := _2ax_calY0_deg_;       // Campo per calibrazione asse Y del DUT.   5555-2
                         end;
                 XN_Yxn: begin
                           kTargetRefX := 0.3;
                           kTargetRefY := 60;
                           CanCmd1 := _2ax_calXNeg_deg_;     // Campo per calibrazione asse X del DUT.   5555-5
                           CanCmd2 := _2ax_calCross_Yxneg_;  // Campo per calibrazione asse Y del DUT.   5555-31
                         end;
                 XP_Yxp: begin
                           kTargetRefX := 0.2;
                           kTargetRefY := -60;
                           CanCmd1 := _2ax_calXPos_deg_;     // Campo per calibrazione asse X del DUT.   5555-3
                           CanCmd2 := _2ax_calCross_Yxpos_;  // Campo per calibrazione asse Y del DUT.   5555-30
                         end;
                 YP_Xyp: begin
                           kTargetRefX := 59.9;
                           kTargetRefY := -0.1;
                           CanCmd1 := _2ax_calCross_Xypos_;  // Campo per calibrazione asse Y del DUT.   5555-28
                           CanCmd2 := _2ax_calYPos_deg_;     // Campo per calibrazione asse X del DUT.   5555-4
                         end;
                 YN_Xyn: begin
                           kTargetRefX := -60.1;
                           kTargetRefY := 0.1;
                           CanCmd1 := _2ax_calCross_Xyneg_;  // Campo per calibrazione asse Y del DUT.   5555-29
                           CanCmd2 := _2ax_calYNeg_deg_;     // Campo per calibrazione asse X del DUT.   5555-6
                         end;
            end;
            edtTargetX.Text := format('%6.2f°', [kTargetRefX]);
            edtTargetY.Text := format('%6.2f°', [kTargetRefY]);

            LogText(#9'Lettura calibrazioni angoli '+ strCalibAngle[ord(CalDeg)]+' X='+ edtTargetX.Text +' e Y='+ edtTargetY.Text );

            // Sblocco area protetta 5555h del DUT è già fatto nella relativa FASE preliminare.
            // legge valore X da DUT as UNS16. (5555-xx)
            Can.Command  := CanCmd1;
            Can.Data     := '';                            // '' perchè è Read quindi ospiterà il risultato.
            Can.NodeID   := edtDUTaddr.text;               // str ID del DUT.
            if not CAN_ReadParam( Can ) then begin         // Delay standard è 50 msec.
                LogStep(113,113, 113, 'ERROR on READ param ['+CanCmd1+'] for DUT:'+ edtDUTaddr.text );
                LogStep(114,114, 114,  lastCAN_ErrMsg );
                break;
            end;
            // The parameter was Read successfully.
            val1 := Can.Data;
            iValueReadX := StrToIntDef('$'+ val1, 0);      // $ perchè payload è stringa ma sempre Hex !
            LogStep(115,115, 115, #9#9+ CanCmd1 +' = $'+ val1 );

            // legge valore Y da DUT as UNS16. (5555-xx)
            Can.Command  := CanCmd2;
            Can.Data     := '';                            // '' perchè è Read quindi ospiterà il risultato.
            Can.NodeID   := edtDUTaddr.text;               // str ID del DUT.
            if not CAN_ReadParam( Can ) then begin         // Delay standard è 50 msec.
                LogStep(116,116, 116, 'ERROR on READ param ['+CanCmd2+'] for DUT:'+ edtDUTaddr.text );
                LogStep(117,117, 117,  lastCAN_ErrMsg );
                break;
            end;
            // The parameter was Read successfully.
            val2 := Can.Data;
            iValueReadY := StrToIntDef('$'+val2, 0);       // $ perchè payload è stringa ma sempre Hex !
            LogStep(118,118, 118, #9#9+ CanCmd2 +' = $'+ val2 );

            LogText( CanCmd1 +' = '+ val1 + ' ('+intToStr(strToIntDef('$'+ val1, 0))+')  ---  '+
                     CanCmd2 +' = '+ val2 + ' ('+intToStr(strToIntDef('$'+ val2, 0))+')');

            // compila lista per review.
            ListItem := lviewCalibVal.Items.Add;
          //ListItem.ImageIndex := -1;                  // per ora niente icon
            ListItem.Caption := strCalibAngle[ord(CalDeg)];
            ListItem.SubItems.Add( iValueReadX.ToString );
            ListItem.SubItems.Add( iValueReadY.ToString );
        end;//for
    finally
        lviewCalibVal.Items.EndUpdate;
        // Blocca subito l'area protetta 5555h del DUT, perchè il Reset node NON dà lo stesso risultato.
        if not CAN_LockParams( edtDUTaddr.text ) then begin
            // l' Errore è bloccante, quindi esco.
            LogStep(119,119, 119, 'ERROR on LOCK params for DUT:'+ edtDUTaddr.text );
            LogStep(120,120, 120,  lastCAN_ErrMsg );
        end
        else
            LogStep(121,121, 121, 'OK, DUT Lock command CONFIRMED by NodeID:'+ edtDUTaddr.text );
    end;
end;

function Toiac3Fasi2ax.DutCalibrato(devAddr: string): boolean;
var
  CalDeg: TCalibAngle;  //  TCalibAngle = ( X0_Y0, X0_Yp60, X0_Yn60, Xp60_Y0, Xn60_Y0 );  // i valori precisi poi dipendono da tabella di calibrazione specifica per famiglia di devices.
  Can: TCanFrame;
  iValueReadX, iValueReadY: integer;
  kDefaultX, kDefaultY: integer;
  kTargetRefX: single;
  kTargetRefY: single;
  CanCmd1, CanCmd2: string;
  val1, val2: string;
  calibrati: integer;
begin
    result := False;
    calibrati := 0;
    // prima di tutto assicura CAN channel "clean"
    if not CAN_CleanCH(CHdut) then begin
        // l' Errore è bloccante, quindi esco.
        LogStep(122,122, 122,  lastCAN_ErrMsg );
        exit;
    end;
    LogStep(123,123, 123, 'STATUS PCAN_OK, check DUT Calibrato...');
    // The channel is Ready to work.

    // Reset DUT necessario per portare in RAM i parametri, nel caso siano appena stati calibrati (settati in eeprom)!
    if CAN_SendNMT( CHdut, cmdResetNode, edtDUTaddr.text ) then begin
        // The Reset was successfully sent
        LogStep(124,124, 124, 'CAN ResetNode SENT Successfully to NodeID:'+ edtDUTaddr.text );
    end
    else begin
        // l' Errore è bloccante, quindi esco.
        LogStep(125,125, 125, 'CAN ResetNode ERROR  on NodeID:'+ edtDUTaddr.text );
        LogStep(126,126, 126,  lastCAN_ErrMsg );
        exit;
    end;
    // DUT reset & boot completed !
    if not CAN_FlushRxBuffer(CHdut) then begin
        LogStep(127,127, 127,  lastCAN_ErrMsg );
        exit;
    end;
    LogStep(128,128, 128, 'OK FLUSH Can queue Successful.');    // preparato rx queue clean.

    // Sblocca area protetta 5555h del DUT per precauzione, se non fosse già da setup fase.
    if not CAN_UnlockParams( edtDUTaddr.text ) then begin
        // l' Errore è bloccante, quindi esco.
        LogStep(129,129, 129, 'ERROR on UNLOCK params for DUT:'+ edtDUTaddr.text );
        LogStep(130,130, 130,  lastCAN_ErrMsg );
        exit;     // unlock non riuscito, quindi non serve lock finally !
    end;
    LogStep(131,131, 131, 'OK, DUT Unlock command CONFIRMED by NodeID:'+ edtDUTaddr.text );

    try
        for CalDeg := X0_Y0 to YN_Xyn do begin
            // verifica Valori come da tabella default calibrazione DUT.
            case CalDeg of
                  X0_Y0: begin
                           kTargetRefX := 0;
                           kTargetRefY := 0;
                           CanCmd1 := _2ax_calX0_deg_;       // Campo per calibrazione asse X del DUT.   5555-1
                           CanCmd2 := _2ax_calY0_deg_;       // Campo per calibrazione asse Y del DUT.   5555-2
                           kDefaultX := 25000;
                           kDefaultY := 25000;               // 61A8
                         end;
                 XN_Yxn: begin
                           kTargetRefX := 0.3;
                           kTargetRefY := 60;
                           CanCmd1 := _2ax_calXNeg_deg_;     // Campo per calibrazione asse X del DUT.   5555-5
                           CanCmd2 := _2ax_calCross_Yxneg_;  // Campo per calibrazione asse Y del DUT.   5555-31
                           kDefaultX := 16339;               // 3FD3
                           kDefaultY := 0;
                         end;
                 XP_Yxp: begin
                           kTargetRefX := 0.2;
                           kTargetRefY := -60;
                           CanCmd1 := _2ax_calXPos_deg_;     // Campo per calibrazione asse X del DUT.   5555-3
                           CanCmd2 := _2ax_calCross_Yxpos_;  // Campo per calibrazione asse Y del DUT.   5555-30
                           kDefaultX := 33660;               // 837C
                           kDefaultY := 0;
                         end;
                 YP_Xyp: begin
                           kTargetRefX := 59.9;
                           kTargetRefY := -0.1;
                           CanCmd1 := _2ax_calCross_Xypos_;  // Campo per calibrazione asse Y del DUT.   5555-28
                           CanCmd2 := _2ax_calYPos_deg_;     // Campo per calibrazione asse X del DUT.   5555-4
                           kDefaultX := 0;
                           kDefaultY := 33660;               // 837C
                         end;
                 YN_Xyn: begin
                           kTargetRefX := -60.1;
                           kTargetRefY := 0.1;
                           CanCmd1 := _2ax_calCross_Xyneg_;  // Campo per calibrazione asse Y del DUT.   5555-29
                           CanCmd2 := _2ax_calYNeg_deg_;     // Campo per calibrazione asse X del DUT.   5555-6
                           kDefaultX := 0;
                           kDefaultY := 16339;               // 3FD3
                         end;
            end;
            LogText(#9'Lettura calibrazioni angoli '+ strCalibAngle[ord(CalDeg)]);

            // Sblocco area protetta 5555h del DUT è già fatto nella relativa FASE preliminare.
            // legge valore X da DUT as UNS16. (5555-xx)
            Can.Command  := CanCmd1;
            Can.Data     := '';                            // '' perchè è Read quindi ospiterà il risultato.
            Can.NodeID   := edtDUTaddr.text;               // str ID del DUT.
            if not CAN_ReadParam( Can ) then begin         // Delay standard è 50 msec.
                LogStep(132,132, 132, 'ERROR on READ param ['+CanCmd1+'] for DUT:'+ edtDUTaddr.text );
                LogStep(133,133, 133,  lastCAN_ErrMsg );
                break;
            end;
            // The parameter was Read successfully.
            val1 := Can.Data;
            iValueReadX := StrToIntDef('$'+ val1, 0);      // $ perchè payload è stringa ma sempre Hex !

            // legge valore Y da DUT as UNS16. (5555-xx)
            Can.Command  := CanCmd2;
            Can.Data     := '';                            // '' perchè è Read quindi ospiterà il risultato.
            Can.NodeID   := edtDUTaddr.text;               // str ID del DUT.
            if not CAN_ReadParam( Can ) then begin         // Delay standard è 50 msec.
                LogStep(134,134, 134, 'ERROR on READ param ['+CanCmd2+'] for DUT:'+ edtDUTaddr.text );
                LogStep(135,135, 135,  lastCAN_ErrMsg );
                break;
            end;
            // The parameter was Read successfully.
            val2 := Can.Data;
            iValueReadY := StrToIntDef('$'+val2, 0);       // $ perchè payload è stringa ma sempre Hex !

            LogText( CanCmd1 +' = '+ val1 + ' ('+intToStr(strToIntDef('$'+ val1, 0))+')  ---  '+
                     CanCmd2 +' = '+ val2 + ' ('+intToStr(strToIntDef('$'+ val2, 0))+')');

            // Controllo se già calibrato
            if (iValueReadX <> kDefaultX) and (iValueReadY <> kDefaultY) then begin
                inc(calibrati); // alla prima coppia che trova fuori default
                break;          // lo considera già calibrato, quindi esce
            end;
        end;//for
    finally
        if calibrati > 0 then
            result := True;
        // Blocca subito l'area protetta 5555h del DUT, perchè il Reset node NON dà lo stesso risultato.
        if not CAN_LockParams( edtDUTaddr.text ) then begin
            // l' Errore è bloccante, quindi esco.
            LogStep(136,136, 136, 'ERROR on LOCK params for DUT:'+ edtDUTaddr.text );
            LogStep(137,137, 137,  lastCAN_ErrMsg );
        end
        else
            LogStep(138,138, 138, 'OK, DUT Lock command CONFIRMED by NodeID:'+ edtDUTaddr.text );
    end;
end;


procedure Toiac3Fasi2ax.StartCalibrazioneBANCO1Click(Sender: TObject);
var
  Can: TCanFrame;
begin
    // prepara DUT e REF address per usi vari.
    LogStep(139,139, 139, '*** CALIBRAZIONE BANCO ***');
    vRef2Addr := strCanValue2integer(edtRef2Addr.text);
    vDutAddr := strCanValue2integer(edtDutAddr.text);
    edtTargetX.text := '';
    edtTargetY.text := '';
    edtRefX.text := '';
    edtRefY.text := '';
    edtDutX.text := '';
    edtdutY.text := '';

    // Stop eventuale REF/DUT, ferma Timer ricezione ciclica valori dal node
    tmrCanRead.enabled := False;
    LogStep(140,140, 140, 'RESET REF e DUT...');
    // Poi Reset ALL nodes, casomai fossero già in Start mode.
    if not CAN_SendNMT( CHref, cmdResetNode, '00h' ) then begin  // Reset ALL Ref nodes!
        LogStep(141,141, 141, lastCAN_ErrMsg);
        exit;
    end;
    if CHdut.Handle <> CHref.Handle then
        if not CAN_SendNMT( CHdut, cmdResetNode, '00h' ) then begin  // Reset ALL Dut nodes!
            LogStep(142,142, 142, lastCAN_ErrMsg);
            exit;
        end;
    LogStep(143,143, 143, 'ALL nodes Reset Successfully');
    LogStep(144,144, 144, 'FLUSH queue All CanCH...');
    if not CAN_FlushRxBuffer(CHref) then begin
        LogStep(145,145, 145,  lastCAN_ErrMsg );
        exit;
    end;
    if CHdut.Handle <> CHref.Handle then
        if not CAN_FlushRxBuffer(CHdut) then begin
            LogStep(146,146, 146,  lastCAN_ErrMsg );
            exit;
        end;
    LogStep(147,147, 147, 'OK FLUSH All CanCH queue Successful.');    // preparato rx queue clean.

    // Setup GUI
    gaugeXLabel.Text := 'REF asse X';
    gaugeYLabel.Text := 'REF asse Y';

    // Start REF2 cyclic
    // serve leggere Ref2_resolution per calcolo DEGs quando acquisisco samples via can read.
    if not CAN_GetResolution( edtRef2Addr.text, Ref2_resolution) then begin // Delay standard è 25 msec
        LogStep(148,148, 148,  lastCAN_ErrMsg );
        exit;
    end;
    // The Device Resolution [1,10,100,1000] was Read successfully !
    LogStep(149,149, 149, 'OK Ref2_resolution READ Successfully from NodeID:'+ edtRef2Addr.text );
    LogStep(150,150, 150, #9#9'Ref2 Resolution = '+ Ref2_resolution.ToString );
    edtResolREF2.text := Ref2_resolution.ToString;

    // per START REF2 è necessario cyclical transmission già configurato su REF2 !
    Can.Command  := _Timer_for_cyclical_txd_;      // = Event timer for cyclical transmission [multiple of 1ms, 0 = disabled]
    Can.Data     := '500';                         // msec.
    Can.NodeID   := edtRef2Addr.text;              // str ID del device
    if not CAN_WriteParam( Can ) then begin
        LogStep(151,151, 151, lastCAN_ErrMsg);
        exit;
    end;
    // The parameter was Written successfully.
    LogStep(152,152, 152, 'OK REF Parameter '+Can.Index+'-'+Can.Subidx+'h STORED Successfully to NodeID:'+ edtRef2Addr.text );
    LogStep(153,153, 153, #9#9'REF2 '+ Can.Command +' = '+ kTimer4cyclicalTxd );

    // REF START node, enter operational state
    if not CAN_SendNMT( CHref, cmdStartNode, edtRef2Addr.text ) then begin
        // l' Errore è bloccante, quindi esco.
        LogStep(154,154, 154, 'ERROR on StartNode REF2:'+ edtRef2Addr.text );
        LogStep(155,155, 155,  lastCAN_ErrMsg );
        exit;
    end;
    // The Start was successfully sent
    LogStep(156,156, 156, 'OK StartNode SENT Successfully to REF2: '+ edtRef2Addr.text );
                                                                                                            // Start DUT cyclic
    // serve leggere DUT_resolution per calcolo DEGs quando acquisisco samples via can read.
    if not CAN_GetResolution( edtDUTaddr.text, DUT_resolution) then begin // Delay standard è 25 msec
        LogStep(157,157, 157,  lastCAN_ErrMsg );
        exit;
    end;
    // The Device Resolution [1,10,100,1000] was Read successfully !
    LogStep(158,158, 158, 'OK DUT Parameter '+Can.Index+'-'+Can.Subidx+'h READ Successfully from NodeID:'+ edtDutAddr.text );
    LogStep(159,159, 159, #9#9'DUT '+ Can.Command +' = '+ DUT_resolution.ToString );
    edtResolDUT.text := DUT_resolution.ToString;

    // per START DUT è necessario cyclical transmission già configurato su DUT !
    Can.Command  := _Timer_for_cyclical_txd_;      // = Event timer for cyclical transmission [multiple of 1ms, 0 = disabled]
    Can.Data     := '500';                         // msec.
    Can.NodeID   := edtDUTaddr.text;               // str ID del device
    if not CAN_WriteParam( Can ) then begin
        LogStep(160,160, 160, lastCAN_ErrMsg);
        exit;
    end;
    // The parameter was Written successfully.
    LogStep(161,161, 161, 'OK DUT Parameter '+Can.Index+'-'+Can.Subidx+'h STORED Successfully to NodeID:'+ edtDUTaddr.text );
    LogStep(162,162, 162, #9#9'DUT '+ Can.Command +' = '+ kTimer4cyclicalTxd );

    // DUT START node, enter operational state
    if not CAN_SendNMT( CHdut, cmdStartNode, edtDUTaddr.text ) then begin
        // l' Errore è bloccante, quindi esco.
        LogStep(163,163, 163, 'ERROR on StartNode DUT:'+ edtDUTaddr.text );
        LogStep(164,164, 164,  lastCAN_ErrMsg );
        exit;
    end;
    // The Start was successfully sent
    LogStep(165,165, 165, 'OK StartNode SENT Successfully to DUT: '+ edtDUTaddr.text );

    CAN_FlushRxBuffer(CHref);                      // prepara rx queue clean
    // inversioni angoli se richieste.
    vInvertX_ref := menuInvertiXref.checked;
    vInvertY_ref := menuInvertiYref.checked;
    CAN_FlushRxBuffer(CHdut);                      // prepara rx queue clean
    // inversioni angoli se richieste.
    vInvertX_dut := menuInvertiXdut.checked;
    vInvertY_dut := menuInvertiYdut.checked;

    LogText(#10#13+#10#13+#10#13);
    LogText(#9+ ' Ref2_X   Ref2_Y');
    LogText(#9#9#9#9#9#9+ ' Dut_X    Dut_Y');
    TInterlocked.BitTestAndClear(SampleReady, 0);  // Reset SampleReady bit 0
    tmrCanRead.enabled := True;                    // avvia timer ricezione ciclica valori da CAN
end;

procedure Toiac3Fasi2ax.StartDUTcyclicalRDClick(Sender: TObject);
var
  Can: TCanFrame;
begin
    // prepara DUT e REF address per usi vari.
    vRef2Addr := strCanValue2integer(edtRef2Addr.text);
    vDutAddr := strCanValue2integer(edtDutAddr.text);
    edtTargetX.text := '';
    edtTargetY.text := '';
    edtRefX.text := '';
    edtRefY.text := '';
    edtDutX.text := '';
    edtdutY.text := '';

    // Stop eventuale DUT/REF, ferma Timer ricezione ciclica valori dal node
    tmrCanRead.enabled := False;
    LogStep(166,166, 166, 'RESET REF e DUT...');
    // Poi Reset ALL nodes, casomai fossero già in Start mode.
    if not CAN_SendNMT( CHref, cmdResetNode, '00h' ) then begin  // Reset ALL Ref nodes!
        lastCAN_ErrMsg := 'ERROR on Reset REF Nodes: '+ lastCAN_ErrMsg;
        LogStep(167,167, 167, lastCAN_ErrMsg);
        exit;
    end;
    if CHdut.Handle <> CHref.Handle then
        if not CAN_SendNMT( CHdut, cmdResetNode, '00h' ) then begin  // Reset ALL Dut nodes!
            lastCAN_ErrMsg := 'ERROR on Reset DUT Nodes: '+ lastCAN_ErrMsg;
            LogStep(168,168, 168, lastCAN_ErrMsg);
            exit;
        end;
    LogStep(169,169, 169, 'ALL nodes Reset Successfully');
    LogStep(170,170, 170, 'FLUSH queue All CanCH...');
    if not CAN_FlushRxBuffer(CHref) then begin
        LogStep(171,171, 171,  'ERROR on Flush CHref: '+lastCAN_ErrMsg );
        exit;
    end;
    if CHdut.Handle <> CHref.Handle then
        if not CAN_FlushRxBuffer(CHdut) then begin
            LogStep(172,172, 172,  'ERROR on Flush CHdut: '+lastCAN_ErrMsg );
            exit;
        end;
    LogStep(173,173, 173, 'OK FLUSH All CanCH Successful.');    // preparato rx queue clean.

    // Setup GUI
    gaugeXLabel.Text := 'DUT asse X';
    gaugeYLabel.Text := 'DUT asse Y';

    // Start DUT cyclic
    // serve leggere DUT_resolution per calcolo DEGs quando acquisisco samples via can read.
    if not CAN_GetResolution( edtDUTaddr.text, DUT_resolution) then begin // Delay standard è 25 msec
        LogStep(174,174, 174,  lastCAN_ErrMsg );
        exit;
    end;
    // The Device Resolution [1,10,100,1000] was Read successfully !
    LogStep(175,175, 175, 'OK DUT Parameter '+Can.Index+'-'+Can.Subidx+'h READ Successfully from NodeID:'+ edtDutAddr.text );
    LogStep(176,176, 176, #9#9'DUT '+ Can.Command +' = '+ DUT_resolution.ToString );
    edtResolDUT.text := DUT_resolution.ToString;

    // per START DUT è necessario cyclical transmission già configurato su DUT !
    Can.Command  := _Timer_for_cyclical_txd_;      // = Event timer for cyclical transmission [multiple of 1ms, 0 = disabled]
    Can.Data     := kTimer4cyclicalTxd;            // 250 msec.
    Can.NodeID   := edtDUTaddr.text;               // str ID del device
    if not CAN_WriteParam( Can ) then begin
        LogStep(177,177, 177, lastCAN_ErrMsg);
        exit;
    end;
    // The parameter was Written successfully.
    LogStep(178,178, 178, 'OK DUT Parameter '+Can.Index+'-'+Can.Subidx+'h STORED Successfully to NodeID:'+ edtDUTaddr.text );
    LogStep(179,179, 179, #9#9'DUT '+ Can.Command +' = '+ kTimer4cyclicalTxd );

    // DUT START node, enter operational state
    if not CAN_SendNMT( CHdut, cmdStartNode, edtDUTaddr.text ) then begin
        // l' Errore è bloccante, quindi esco.
        LogStep(180,180, 180, 'ERROR on StartNode DUT:'+ edtDUTaddr.text );
        LogStep(181,181, 181,  lastCAN_ErrMsg );
        exit;
    end;
    // The Start was successfully sent
    LogStep(182,182, 182, 'OK StartNode SENT Successfully to DUT: '+ edtDUTaddr.text );
    CAN_FlushRxBuffer(CHdut);                     // prepara rx queue clean
    TInterlocked.BitTestAndClear(SampleReady, 0);  // Reset SampleReady bit 0

    // inversioni angoli se richieste.
    vInvertX_dut := menuInvertiXdut.checked;
    vInvertY_dut := menuInvertiYdut.checked;

    tmrCanRead.enabled := True;                    // avvia timer ricezione ciclica valori da CAN (Ref2)
end;

procedure Toiac3Fasi2ax.StartREF2cyclicalRDClick(Sender: TObject);
var
  Can: TCanFrame;
begin
    // prepara DUT e REF address per usi vari.
    vRef2Addr := strCanValue2integer(edtRef2Addr.text);
    vDutAddr := strCanValue2integer(edtDutAddr.text);
    edtTargetX.text := '';
    edtTargetY.text := '';
    edtRefX.text := '';
    edtRefY.text := '';
    edtDutX.text := '';
    edtdutY.text := '';

    // Stop eventuale REF/DUT, ferma Timer ricezione ciclica valori dal node
    tmrCanRead.enabled := False;
    LogStep(183,183, 183, 'RESET REF e DUT...');
    // Poi Reset ALL nodes, casomai fossero già in Start mode.
    if not CAN_SendNMT( CHref, cmdResetNode, '00h' ) then begin  // Reset ALL Ref nodes!
        LogStep(184,184, 184, lastCAN_ErrMsg);
        exit;
    end;
    if CHdut.Handle <> CHref.Handle then
        if not CAN_SendNMT( CHdut, cmdResetNode, '00h' ) then begin  // Reset ALL Dut nodes!
            LogStep(185,185, 185, lastCAN_ErrMsg);
            exit;
        end;
    LogStep(186,186, 186, 'ALL nodes Reset Successfully');
    LogStep(187,187, 187, 'FLUSH queue All CanCH...');
    if not CAN_FlushRxBuffer(CHref) then begin
        LogStep(188,188, 188,  lastCAN_ErrMsg );
        exit;
    end;
    if CHdut.Handle <> CHref.Handle then
        if not CAN_FlushRxBuffer(CHdut) then begin
            LogStep(189,189, 189,  lastCAN_ErrMsg );
            exit;
        end;
    LogStep(190,190, 190, 'OK FLUSH All CanCH queue Successful.');    // preparato rx queue clean.

    // Setup GUI
    gaugeXLabel.Text := 'REF asse X';
    gaugeYLabel.Text := 'REF asse Y';

    // Start REF2 cyclic
    // serve leggere Ref2_resolution per calcolo DEGs quando acquisisco samples via can read.
    if not CAN_GetResolution( edtRef2Addr.text, Ref2_resolution) then begin // Delay standard è 25 msec
        LogStep(191,191, 191,  lastCAN_ErrMsg );
        exit;
    end;
    // The Device Resolution [1,10,100,1000] was Read successfully !
    LogStep(192,192, 192, 'OK Ref2_resolution READ Successfully from NodeID:'+ edtRef2Addr.text );
    LogStep(193,193, 193, #9#9'Ref2 Resolution = '+ Ref2_resolution.ToString );
    edtResolREF2.text := Ref2_resolution.ToString;

    // per START REF2 è necessario cyclical transmission già configurato su REF2 !
    Can.Command  := _Timer_for_cyclical_txd_;      // = Event timer for cyclical transmission [multiple of 1ms, 0 = disabled]
    Can.Data     := kTimer4cyclicalTxd;            // 250 msec.
    Can.NodeID   := edtRef2Addr.text;              // str ID del device
    if not CAN_WriteParam( Can ) then begin
        LogStep(194,194, 194, lastCAN_ErrMsg);
        exit;
    end;
    // The parameter was Written successfully.
    LogStep(195,195, 195, 'OK REF Parameter '+Can.Index+'-'+Can.Subidx+'h STORED Successfully to NodeID:'+ edtRef2Addr.text );
    LogStep(196,196, 196, #9#9'REF2 '+ Can.Command +' = '+ kTimer4cyclicalTxd );

    // REF START node, enter operational state
    if not CAN_SendNMT( CHref, cmdStartNode, edtRef2Addr.text ) then begin
        // l' Errore è bloccante, quindi esco.
        LogStep(197,197, 197, 'ERROR on StartNode REF2:'+ edtRef2Addr.text );
        LogStep(198,198, 198,  lastCAN_ErrMsg );
        exit;
    end;
    // The Start was successfully sent
    LogStep(199,199, 199, 'OK StartNode SENT Successfully to REF2: '+ edtRef2Addr.text );
    CAN_FlushRxBuffer(CHref);                     // prepara rx queue clean
    TInterlocked.BitTestAndClear(SampleReady, 0);  // Reset SampleReady bit 0

    // inversioni angoli se richieste.
    vInvertX_ref := menuInvertiXref.checked;
    vInvertY_ref := menuInvertiYref.checked;

    tmrCanRead.enabled := True;                    // avvia timer ricezione ciclica valori da CAN (Ref2)
end;

procedure Toiac3Fasi2ax.StopDUTcyclicalRDClick(Sender: TObject);
begin
    // Stop DUT cyclic
    // ferma Timer ricezione ciclica valori da DUT/REF
    tmrCanRead.enabled := False;
    // Poi Reset ALL nodes, casomai fossero già in Start mode.
    if not CAN_SendNMT( CHref, cmdResetNode, '00h' ) then begin  // Reset ALL Ref nodes!
        LogStep(200,200, 200, lastCAN_ErrMsg);
        exit;
    end;
    if CHdut.Handle <> CHref.Handle then
        if not CAN_SendNMT( CHdut, cmdResetNode, '00h' ) then begin  // Reset ALL Dut nodes!
            LogStep(201,201, 201, lastCAN_ErrMsg);
            exit;
        end;
    LogStep(202,202, 202, 'ALL nodes Reset Successfully');
    if not CAN_FlushRxBuffer(CHref) then begin
        LogStep(203,203, 203,  lastCAN_ErrMsg );
        exit;
    end;
    LogStep(204,204, 204, 'OK FLUSH CanREF queue Successful.');    // preparato rx queue clean.
    if CHdut.Handle <> CHref.Handle then
        if not CAN_FlushRxBuffer(CHdut) then begin
            LogStep(205,205, 205,  lastCAN_ErrMsg );
            exit;
        end;
    LogStep(206,206, 206, 'OK FLUSH CanDUT queue Successful.');    // preparato rx queue clean.

    // Restore GUI su Ref
    gaugeXLabel.Text := 'REF asse X';
    gaugeYLabel.Text := 'REF asse Y';
    edtTargetX.text := '';
    edtTargetY.text := '';
    edtRefX.text := '';
    edtRefY.text := '';
    edtDutX.text := '';
    edtdutY.text := '';
    // Reset gauges
    gaugeX.Value := 0;
    gaugeY.Value := 0;

    LogStep(207,207, 207, 'OK Dut read STOPPED: '+ edtDUTaddr.text );
end;

procedure Toiac3Fasi2ax.StopREF2cyclicalRDClick(Sender: TObject);
begin
    // Stop REF2 cyclic
    // ferma Timer ricezione ciclica valori da DUT/REF
    tmrCanRead.enabled := False;
    // Poi Reset ALL nodes, casomai fossero già in Start mode.
    if not CAN_SendNMT( CHref, cmdResetNode, '00h' ) then begin  // Reset ALL Ref nodes!
        LogStep(208,208, 208, lastCAN_ErrMsg);
        exit;
    end;
    if CHdut.Handle <> CHref.Handle then
        if not CAN_SendNMT( CHdut, cmdResetNode, '00h' ) then begin  // Reset ALL Dut nodes!
            LogStep(209,209, 209, lastCAN_ErrMsg);
            exit;
        end;
    LogStep(210,210, 210, 'ALL nodes Reset Successfully');
    if not CAN_FlushRxBuffer(CHref) then begin
        LogStep(211,211, 211,  lastCAN_ErrMsg );
        exit;
    end;
    LogStep(212,212, 212, 'OK FLUSH CanREF queue Successful.');
    if CHdut.Handle <> CHref.Handle then
        if not CAN_FlushRxBuffer(CHdut) then begin
            LogStep(213,213, 213,  lastCAN_ErrMsg );
            exit;
        end;
    LogStep(214,214, 214, 'OK FLUSH CanDUT queue Successful.');    // preparato rx queue clean.

    // Restore GUI su Ref che è default.
    gaugeXLabel.Text := 'REF asse X';
    gaugeYLabel.Text := 'REF asse Y';
    edtTargetX.text := '';
    edtTargetY.text := '';
    edtRefX.text := '';
    edtRefY.text := '';
    edtDutX.text := '';
    edtdutY.text := '';
    // Reset gauges
    gaugeX.Value := 0;
    gaugeY.Value := 0;

    LogStep(215,215, 215, 'OK Ref2 cyclic read STOPPED: '+ edtRef2Addr.text );
end;

procedure Toiac3Fasi2ax.debugReadAngle1Click(Sender: TObject);
var
//iDeg: integer;      NO!  produce valori sballati !
  iDeg: smallint;  // 2-byte signed integer.
begin
    // serve leggere DUT_resolution per calcolo DEGs quando acquisisco samples via can read.
    if not CAN_GetResolution( edtDUTaddr.text, DUT_resolution) then begin // Delay standard è 25 msec
        LogStep(216,216, 216,  lastCAN_ErrMsg );
        exit;
    end;
    // The Device Resolution [1,10,100,1000] was Read successfully !
    LogStep(217,217, 217, #9#9'DUT resolution = '+ DUT_resolution.ToString );

    // legge Angolo X da DUT as INT16. (6010-0)
    if not CAN_Read_ANGLE( edtDUTaddr.text, 'X', iDeg) then begin // Delay standard è 25 msec
        LogStep(218,218, 218,  lastCAN_ErrMsg );
        exit;
    end;
    // The currXDut was Read successfully !
    currXDut := iDeg / DUT_resolution;            // applico risoluzione del Dut.
    // NB: devono essere dello stesso tipo, smallint o integer entrambi !!!
    // inversioni angolo DUT, se richieste da opzione popmenù gauges
    if vInvertX_dut then
        currXDut := currXDut * -1;        // Necessario per DUT invertire solo X !
    edtDutX.text := format('%6.2f', [currXDut]);  // currXDut.ToString;

    // legge Angolo Y da DUT as INT16. (6020-0)
    if not CAN_Read_ANGLE( edtDUTaddr.text, 'Y', iDeg) then begin // Delay standard è 25 msec
        LogStep(219,219, 219,  lastCAN_ErrMsg );
        exit;
    end;
    // The currYDut was Read successfully !
    currYDut := iDeg / DUT_resolution;             // applico risoluzione del Dut.
    // NB: devono essere dello stesso tipo, smallint o integer entrambi !!!
    // inversioni angolo DUT, se richieste da opzione popmenù gauges
    if vInvertY_dut then
        currYDut := currYDut * -1;        // Necessario per DUT invertire solo Y !
    edtDutY.text := format('%6.2f', [currYDut]);   // currYDut.ToString;
    LogText(#9#9+ format('DUT_X = %6.2f    DUT_Y = %6.2f', [currXDut, currYDut]));
end;

procedure Toiac3Fasi2ax.SetupA1Click(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 0;
    ResetJobs_Data();               // reset anche counters della JobList.
    mainForm.SetSession_Data();

    if SetupFase_A(Params) then
        LogStep(220,220, 220,  'Fase setup SUCCESS.')
    else
        LogStep(221,221, 221,  'Fase FAILED: '+ Params.retMsg );
end;

procedure Toiac3Fasi2ax.LoadDefaults1Click(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 1;
    RunStep_A1(Params);
end;

procedure Toiac3Fasi2ax.ClearLog1Click(Sender: TObject);
begin
    mainForm.mmLog.Clear;
end;

procedure Toiac3Fasi2ax.ClearTSPairs1Click(Sender: TObject);
begin
    CalibRecords.Clear; // Clear removes all keys and values from the dictionary. The Count property is set to 0.
end;

procedure Toiac3Fasi2ax.LetturaParametri1Click(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 2;
    RunStep_A2(Params);
end;


procedure Toiac3Fasi2ax.ScriviParametri1Click(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 3;
    RunStep_A3(Params);
end;

procedure Toiac3Fasi2ax.SalvaParametri1Click(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 4;
    RunStep_A4(Params);
end;


procedure Toiac3Fasi2ax.SaveDB1Click(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 32;
    RunStep_F29(Params);
end;

procedure Toiac3Fasi2ax.MenuItem2Click(Sender: TObject);
var
  Params: TJobParams;
begin
    SetupFase_B(Params);
end;

procedure Toiac3Fasi2ax.MenuItem3Click(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 6;
    RunStep_B5(Params);
end;

procedure Toiac3Fasi2ax.MenuItem4Click(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 7;
    RunStep_B6(Params);
end;

procedure Toiac3Fasi2ax.MenuItem5Click(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 8;
    RunStep_B7(Params);
end;

procedure Toiac3Fasi2ax.MenuItem6Click(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 9;
    RunStep_B8(Params);
end;

procedure Toiac3Fasi2ax.MenuItem7Click(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 10;
    RunStep_B9(Params);
end;

procedure Toiac3Fasi2ax.cxButton3Click(Sender: TObject);
begin
    SaveDB1Click(nil);
end;

procedure Toiac3Fasi2ax.viewPanel28Click(Sender: TObject);
begin
    pnlCardsF.ActiveCardIndex := 2;     //card_F28
end;

procedure Toiac3Fasi2ax.viewPanel29Click(Sender: TObject);
begin
    pnlCardsF.ActiveCardIndex := 0;     //card_F29
end;

procedure Toiac3Fasi2ax.viewPanel30Click(Sender: TObject);
begin
    pnlCardsF.ActiveCardIndex := 1;     //card_F30
end;

procedure Toiac3Fasi2ax.tabctrlFasiChange(Sender: TObject);
begin
    if tabctrlFasi.TabIndex < 0 then
        Exit;
    FLoading := True;
    try
      case tabctrlFasi.TabIndex of
        0: pnlCardsFasi.ActiveCardIndex := 0;      // o anche  pnlCardsFasi.ActiveCard := 'cardAdmin';
        1: begin
             pnlCardsFasi.ActiveCardIndex := 1;    // stessa Card per fasi A e E
             panel_A.visible := True;
             pnlCardsFasi.Cards[1].PopupMenu := PopMenuFaseA;
             cbxListaDefaultsCom.Checked := True;    //
             cbxListaDefaults1ax.checked := False;   //
             cbxListaDefaults2ax.checked := True;    //
             cbxListaDefaultsEND.Checked := False;   //
             cbxListaSN.Checked       := True;       //
             cbxListaCalib1ax.Checked := False;      // Definisce tipo di parametri da caricare in listview.
             cbxListaCalib2ax.Checked := False;      // qui solo parametri DEF.
           end;
        2: begin
             pnlCardsFasi.ActiveCardIndex := 2;    // stessa Card per fasi B,C e D
             panelPosition(panel_B);
             panel_B.visible  := True;
             panel_CD.visible := False;
             pnlCardsFasi.Cards[2].PopupMenu := PopMenuFaseB;
             lbDutMode.caption := 'DUT [mV]';
           end;
      3,4: begin
             pnlCardsFasi.ActiveCardIndex := 2;    // stessa Card per fasi B,C e D
             panel_B.visible  := False;
             panelPosition(panel_CD);
             panel_CD.visible := True;
             if tabctrlFasi.TabIndex = 3 then
                 pnlCardsFasi.Cards[2].PopupMenu := PopMenuFaseC
             else
                 pnlCardsFasi.Cards[2].PopupMenu := PopMenuFaseD;
             lbDutMode.caption := 'DUT [deg]';
           end;
        5: begin
             pnlCardsFasi.ActiveCardIndex := 1;      // stessa Card per fasi A e E
             panel_A.visible := False;
             pnlCardsFasi.Cards[1].PopupMenu := PopMenuFaseE;
             cbxListaDefaultsCom.Checked := True;    //
             cbxListaDefaults1ax.checked := False;   //
             cbxListaDefaults2ax.checked := True;    //
             if idxModuloSw = numModuliSw then           // se questo è l'ultimo modulo tra i ModuliSw previsti
                 cbxListaDefaultsEND.Checked := True     // abilito il caricamento dei defaults di Fine EOL !
             else
                 cbxListaDefaultsEND.Checked := False;
             cbxListaSN.Checked       := True;       //
             cbxListaCalib1ax.Checked := False;      // Definisce tipo di parametri da caricare in listview.
             cbxListaCalib2ax.Checked := True;       // qui parametri DEF e CFG2.
           end;
        6: pnlCardsFasi.ActiveCardIndex := 3;      // Fase F
        7: pnlCardsFasi.ActiveCardIndex := 4;      // Fase G
      end;
    finally
        FLoading := False;
    end;
end;

procedure Toiac3Fasi2ax.panelCenter(aPanel:TPanel);  // posiziona sempre al centro.
var
  i,k: integer;
begin
  i := frameFasi.Width div 2;
  k := aPanel.Width div 2;
  aPanel.Left := i-k;

  // Altdisponibile = (H pnlCardsFasi - H area occupata da gauges), rispetto a frame base che ha H fissa 600.
  i := (pnlCardsFasi.Height - panel_Gauges.Height) div 2;
  k := aPanel.Height div 2;
  aPanel.Top := i-k;

  aPanel.BevelOuter := bvNone;   // bevel utile solo a design time
end;

procedure Toiac3Fasi2ax.panelPosition(aPanel:TPanel); // posiziona a coordinate fisse.
begin
    if aPanel.Name = 'panel_CD' then begin
      //aPanel.Left := 210;
        aPanel.Left := tabctrlFasi.Width - aPanel.Width - 25;  // 25 distanza da bordo Dx.
        aPanel.Top := 20;
    end;
    if aPanel.Name = 'panel_B' then begin
        aPanel.Left := 330;
        aPanel.Top := 98;
    end;
    aPanel.BevelOuter := bvNone;   // bevel utile solo a design time
end;

procedure Toiac3Fasi2ax.Writecal1Click(Sender: TObject);
var
  Can: TCanFrame;
begin
    // Memorizza valore Y di calibrazione ottenuto nel DUT.
    Can.Command  := _2ax_calY0_deg_;              // Campo per calibrazione asse Y del DUT.   5555-2
    Can.Data     := 'Y0'+ chr($12) + chr($34);    // va trattato come VSTR perchè vuole prefisso 'Y0'+ value 16 bit senza swap !
    Can.NodeID   := edtDUTaddr.text;              // str ID del DUT
    if not CAN_WriteParam( Can ) then begin
        LogStep(222,222, 222, 'Error on DUT WrParam '+Can.Index+'-'+Can.Subidx+'h  msg:'+ lastCAN_ErrMsg);
        exit;
    end;
    // The parameter was Written successfully.
    LogStep(223,223, 223, 'OK, DUT Calibration  '+Can.Index+'-'+Can.Subidx+'h STORED Successfully to DutID:'+ edtDUTaddr.text );
end;

procedure Toiac3Fasi2ax.Scarto1Click(Sender: TObject);
begin
    edtScartoMax_calib.value := 2.5;   // allargo per debug.
end;

procedure Toiac3Fasi2ax.Scarto2Click(Sender: TObject);
begin
    edtScartoMax_calib.value := 0.02;   // stringo per debug.
end;


procedure Toiac3Fasi2ax.MenuItem10Click(Sender: TObject);
begin
    tmrCanRead.enabled := False;     // ferma ricezione ciclica valori da CAN (Ref2)
    dmBase.wait(300);                // ma non l'invio dal REF !
    CAN_FlushRxBuffer(CHref);       // prepara rx queue clean.
end;

procedure Toiac3Fasi2ax.MenuItem11Click(Sender: TObject);
begin
    CAN_FlushRxBuffer(CHref);       // prepara rx queue clean.
    tmrCanRead.enabled := True;      // ferma ricezione ciclica valori da CAN (Ref2)
end;

procedure Toiac3Fasi2ax.ABORT1Click(Sender: TObject);
begin
    doAbortJob := True;             // Set flag Abort. In alternativa usare schema TInterlocked.BitTestAndSet
    frameFasi.Busy := False;
end;

procedure Toiac3Fasi2ax.ResetAbort1Click(Sender: TObject);
begin
    doAbortJob := False;             // Reset flag Abort.
end;

procedure Toiac3Fasi2ax.ResetTreeviewicons1Click(Sender: TObject);
begin
    mainForm.resetTreeviewIcons();
end;

procedure Toiac3Fasi2ax.SetupC1Click(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 11;
    if SetupFase_C(Params) then
        LogStep(224,224, 224,  'Fase setup SUCCESS.')
    else
        LogStep(225,225, 225,  'Fase FAILED: '+ Params.retMsg );
end;

procedure Toiac3Fasi2ax.Step101Click(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 12;
    if RunStep_C10(Params) then
        LogStep(226,226, 226,  'Step SUCCESS.')
    else
        LogStep(227,227, 227,  'Step FAILED: '+ Params.retMsg );
end;

procedure Toiac3Fasi2ax.Step111Click(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 13;
    if RunStep_C11(Params) then
        LogStep(228,228, 228,  'Step SUCCESS.')
    else
        LogStep(229,229, 229,  'Step FAILED: '+ Params.retMsg );
end;

procedure Toiac3Fasi2ax.Step121Click(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 14;
    if RunStep_C12(Params) then
        LogStep(230,230, 230,  'Step SUCCESS.')
    else
        LogStep(231,231, 231,  'Step FAILED: '+ Params.retMsg );
end;

procedure Toiac3Fasi2ax.Step131Click(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 15;
    if RunStep_C13(Params) then
        LogStep(232,232, 232,  'Step SUCCESS.')
    else
        LogStep(233,233, 233,  'Step FAILED: '+ Params.retMsg );
end;

procedure Toiac3Fasi2ax.Step141Click(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 16;
    if RunStep_C14(Params) then
        LogStep(234,234, 234,  'Step SUCCESS.')
    else
        LogStep(235,235, 235,  'Step FAILED: '+ Params.retMsg );
end;

procedure Toiac3Fasi2ax.Step151Click(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 17;
    if RunStep_C15(Params) then
        LogStep(236,236, 236,  'Step SUCCESS.')
    else
        LogStep(237,237, 237,  'Step FAILED: '+ Params.retMsg );
end;

procedure Toiac3Fasi2ax.Step161Click(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 18;
    if RunStep_C16(Params) then
        LogStep(238,238, 238,  'Step SUCCESS.')
    else
        LogStep(239,239, 239,  'Step FAILED: '+ Params.retMsg );
end;


procedure Toiac3Fasi2ax.MenuItem1Click(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 19;
    if SetupFase_D(Params) then
        LogStep(240,240, 240,  'Fase setup SUCCESS.')
    else
        LogStep(241,241, 241,  'Fase FAILED: '+ Params.retMsg );
end;

procedure Toiac3Fasi2ax.MenuItem8Click(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 20;
    if RunStep_D17(Params) then
        LogStep(242,242, 242,  'Step SUCCESS.')
    else
        LogStep(243,243, 243,  'Step FAILED: '+ Params.retMsg );
end;

procedure Toiac3Fasi2ax.MenuItem12Click(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 21;
    if RunStep_D18(Params) then
        LogStep(244,244, 244,  'Step SUCCESS.')
    else
        LogStep(245,245, 245,  'Step FAILED: '+ Params.retMsg );
end;

procedure Toiac3Fasi2ax.MenuItem13Click(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 22;
    if RunStep_D19(Params) then
        LogStep(246,246, 246,  'Step SUCCESS.')
    else
        LogStep(247,247, 247,  'Step FAILED: '+ Params.retMsg );
end;

procedure Toiac3Fasi2ax.MenuItem14Click(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 23;
    if RunStep_D20(Params) then
        LogStep(248,248, 248,  'Step SUCCESS.')
    else
        LogStep(249,249, 249,  'Step FAILED: '+ Params.retMsg );
end;

procedure Toiac3Fasi2ax.MenuItem15Click(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 24;
    if RunStep_D21(Params) then
        LogStep(250,250, 250,  'Step SUCCESS.')
    else
        LogStep(251,251, 251,  'Step FAILED: '+ Params.retMsg );
end;

procedure Toiac3Fasi2ax.MenuItem16Click(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 25;
    if RunStep_D22(Params) then
        LogStep(252,252, 252,  'Step SUCCESS.')
    else
        LogStep(253,253, 253,  'Step FAILED: '+ Params.retMsg );
end;

procedure Toiac3Fasi2ax.MenuItem17Click(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 26;
    if RunStep_D23(Params) then
        LogStep(254,254, 254,  'Step SUCCESS.')
    else
        LogStep(255,255, 255,  'Step FAILED: '+ Params.retMsg );
end;

procedure Toiac3Fasi2ax.menuSetupEClick(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 27;
    if SetupFase_E(Params) then
        LogStep(256,256, 256,  'Fase setup SUCCESS.')
    else
        LogStep(257,257, 257,  'Fase FAILED: '+ Params.retMsg );
end;

procedure Toiac3Fasi2ax.LoadListaVerifica24Click(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 28;
    if RunStep_E24(Params) then
        LogStep(258,258, 258,  'Step SUCCESS.')
    else
        LogStep(259,259, 259,  'Step FAILED: '+ Params.retMsg );
end;

procedure Toiac3Fasi2ax.LetturaParametri25Click(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 29;
    if RunStep_E25(Params) then
        LogStep(260,260, 260,  'Step SUCCESS.')
    else
        LogStep(261,261, 261,  'Step FAILED: '+ Params.retMsg );
end;

procedure Toiac3Fasi2ax.ScriviParametri26Click(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 30;
    if RunStep_E26(Params) then
        LogStep(262,262, 262,  'Step SUCCESS.')
    else
        LogStep(263,263, 263,  'Step FAILED: '+ Params.retMsg );
end;

procedure Toiac3Fasi2ax.SalvaParametri27Click(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 31;
    if RunStep_E27(Params) then
        LogStep(264,264, 264,  'Step SUCCESS.')
    else
        LogStep(265,265, 265,  'Step FAILED: '+ Params.retMsg );
end;

procedure Toiac3Fasi2ax.SetupFClick(Sender: TObject);
var
  Params: TJobParams;
begin
    currJob := 32;
    if SetupFase_F(Params) then
        LogStep(266,266, 266,  'Fase setup SUCCESS.')
    else
        LogStep(267,267, 267,  'Fase FAILED: '+ Params.retMsg );
end;

procedure Toiac3Fasi2ax.SalvasuDB29Click(Sender: TObject);
var
  Params: TJobParams;
begin
    pnlCardsF.ActiveCardIndex := 0;    // qui necessario pechè non passa da mainForm.trvJobListChange()
    currJob := 34;
    if RunStep_F29(Params) then
        LogStep(268,268, 268,  'Step SUCCESS.')
    else
        LogStep(269,269, 269,  'Step FAILED: '+ Params.retMsg );
end;

procedure Toiac3Fasi2ax.CreaReport30Click(Sender: TObject);
var
  Params: TJobParams;
begin
    pnlCardsF.ActiveCardIndex := 1;    // qui necessario pechè non passa da mainForm.trvJobListChange()
    currJob := 35;
    if RunStep_F30(Params) then
        LogStep(270,270, 270,  'Step SUCCESS.')
    else
        LogStep(271,271, 271,  'Step FAILED: '+ Params.retMsg );
end;



//#########################  FINE Pulsanti per DEBUG  ###############################


procedure Toiac3Fasi2ax.rgXlsZoomxPropertiesChange(Sender: TObject);
begin
    case rgXlsZoom.ItemIndex of
      0: xlsPreview.Zoom := 0.5;
      1: xlsPreview.Zoom := 1;
      2: xlsPreview.AutofitPreview := TAutofitPreview.Width;
    end;
    xlsPreview.InvalidatePreview;
end;





//#####################################################################################################
//
//   Sezione con le funzioni (FASI/STEP) da personalizzare secondo le attività di test da eseguire
//           è la parte che cambia per ogni famiglia DUT da gestire.
//
function Toiac3Fasi2ax.LoadJobList:boolean;
var
  itemIdx: integer;
  urlDatasheet: string;
begin
    JobList.Clear;             // TStringList parte sempre da index 0 !
    numSteps := 0;
    numFasi  := 0;
    urlDatasheet := '\\netserver\05.DbTecnico\OI\'+ DutList[currDutIdx].CodArticolo +'\'+ DutList[currDutIdx].CodArticolo +'_DSP.html';

    JobRecord := TJobRecord.Create;
    JobRecord.RecType         := 'FASE';
    JobRecord.Description     := 'Configurazione interfaccia PCAN e setup DUT';
    JobRecord.urlDBtecnico    := urlDatasheet;
    JobRecord.urlMemoAudio    := kAppPath + kDataSource +'ConfiguraPCan.mp3';
    JobRecord.runSetup        := SetupFase_A;
    JobRecord.runStep         := nil;            // la Fase ha solo il Setup !
    JobRecord.StepCardPanel   := nil;            // Le Fasi non hanno subCards, mentre gli steps potrebbero.
    JobRecord.cardIdx         := -1;             // nemmeno un subCard index.
    JobRecord.GroupFase       := 'Fase_A';
    itemIdx := JobList.AddObject('Fase_A', TJobRecord(JobRecord));
    JobRecord.ListIndex       := itemIdx;  // per il passaggio diretto da Nodo TreeView ad item della stringList.
    inc(numFasi);

          JobRecord := TJobRecord.Create;
          JobRecord.RecType         := 'STEP';
          JobRecord.Description     := 'Carica parametri di Default da file XLS';
          JobRecord.urlDBtecnico    := urlDatasheet;
          JobRecord.runSetup        := nil;          // lo Step non ha un Setup !
          JobRecord.StepCardPanel   := nil;          // questo Step non ha una subCard specifica, ovvero tutti gli steps di questa Fase condividono la stessa Card base della Fase (pnlCardsFasi).
          JobRecord.cardIdx         := -1;           // quindi non ha nemmeno un index di subCard specifica.
          JobRecord.runStep         := RunStep_A1;
          JobRecord.GroupFase      := 'Fase_A';
          itemIdx := JobList.AddObject('Step_1', TJobRecord(JobRecord));
          JobRecord.ListIndex       := itemIdx;  // per il passaggio diretto da Nodo TreeView ad item della stringList.
          inc(numSteps);

          JobRecord := TJobRecord.Create;
          JobRecord.RecType         := 'STEP';
          JobRecord.Description     := 'Lettura parametri dal DUT e confronto con Default';
          JobRecord.urlDBtecnico    := urlDatasheet;
          JobRecord.runSetup        := nil;
          JobRecord.StepCardPanel   := nil;
          JobRecord.cardIdx         := -1;
          JobRecord.runStep         := RunStep_A2;
          JobRecord.GroupFase      := 'Fase_A';
          itemIdx := JobList.AddObject('Step_2', TJobRecord(JobRecord));
          JobRecord.ListIndex       := itemIdx;
          inc(numSteps);

          JobRecord := TJobRecord.Create;
          JobRecord.RecType         := 'STEP';
          JobRecord.Description     := 'Scrivi parametri di default nel DUT';
          JobRecord.urlDBtecnico    := urlDatasheet;
          JobRecord.runSetup        := nil;
          JobRecord.StepCardPanel   := nil;
          JobRecord.cardIdx         := -1;
          JobRecord.runStep         := RunStep_A3;
          JobRecord.GroupFase      := 'Fase_A';
          itemIdx := JobList.AddObject('Step_3', TJobRecord(JobRecord));
          JobRecord.ListIndex       := itemIdx;
          inc(numSteps);

          JobRecord := TJobRecord.Create;
          JobRecord.RecType         := 'STEP';
          JobRecord.Description     := 'Salva parametri in EEProm del DUT';
          JobRecord.urlDBtecnico    := urlDatasheet;
          JobRecord.runSetup        := nil;
          JobRecord.StepCardPanel   := nil;
          JobRecord.cardIdx         := -1;
          JobRecord.runStep         := RunStep_A4;
          JobRecord.GroupFase      := 'Fase_A';
          itemIdx := JobList.AddObject('Step_4', TJobRecord(JobRecord));
          JobRecord.ListIndex       := itemIdx;
          inc(numSteps);


    JobRecord := TJobRecord.Create;
    JobRecord.RecType         := 'FASE';
    JobRecord.Description     := 'Calibrazione Inclinometro';
    JobRecord.urlDBtecnico    := '\\netserver\06.Qualità\SGQ2_attività\01_Mansionari_IOP_MOD\04.Produzione\02.IOP\PRD.150.html';
    JobRecord.runSetup        := SetupFase_B;
    JobRecord.runStep         := nil;
    JobRecord.StepCardPanel   := nil;            // Le Fasi non hanno subCards, mentre gli steps potrebbero.
    JobRecord.cardIdx         := -1;             // nemmeno un subCard index.
    JobRecord.GroupFase       := 'Fase_B';
    itemIdx := JobList.AddObject('Fase_B', TJobRecord(JobRecord));
    JobRecord.ListIndex       := itemIdx;
    inc(numFasi);

          JobRecord := TJobRecord.Create;
          JobRecord.RecType         := 'STEP';
          JobRecord.Description     := 'Calibrazione DUT a inclinazione X=0°  Y=0°';
          JobRecord.urlDBtecnico    := '\\netserver\06.Qualità\SGQ2_attività\01_Mansionari_IOP_MOD\04.Produzione\02.IOP\PRD.150.html';
          JobRecord.runSetup        := nil;
          JobRecord.StepCardPanel   := nil;
          JobRecord.cardIdx         := -1;
          JobRecord.runStep         := RunStep_B5;
          JobRecord.GroupFase      := 'Fase_B';
          itemIdx := JobList.AddObject('Step_5', TJobRecord(JobRecord));
          JobRecord.ListIndex       := itemIdx;
          inc(numSteps);

          JobRecord := TJobRecord.Create;
          JobRecord.RecType         := 'STEP';
          JobRecord.Description     := 'Calibrazione DUT a inclinazione X=0,3°  Y=60°';
          JobRecord.urlDBtecnico    := '\\netserver\06.Qualità\SGQ2_attività\01_Mansionari_IOP_MOD\04.Produzione\02.IOP\PRD.150.html';
          JobRecord.runSetup        := nil;
          JobRecord.StepCardPanel   := nil;
          JobRecord.cardIdx         := -1;
          JobRecord.runStep         := RunStep_B6;
          JobRecord.GroupFase      := 'Fase_B';
          itemIdx := JobList.AddObject('Step_6', TJobRecord(JobRecord));
          JobRecord.ListIndex       := itemIdx;
          inc(numSteps);

          JobRecord := TJobRecord.Create;
          JobRecord.RecType         := 'STEP';
          JobRecord.Description     := 'Calibrazione DUT a inclinazione X=0,2°  Y=-60°';
          JobRecord.urlDBtecnico    := '\\netserver\06.Qualità\SGQ2_attività\01_Mansionari_IOP_MOD\04.Produzione\02.IOP\PRD.150.html';
          JobRecord.runSetup        := nil;
          JobRecord.StepCardPanel   := nil;
          JobRecord.cardIdx         := -1;
          JobRecord.runStep         := RunStep_B7;
          JobRecord.GroupFase      := 'Fase_B';
          itemIdx := JobList.AddObject('Step_7', TJobRecord(JobRecord));
          JobRecord.ListIndex       := itemIdx;
          inc(numSteps);

          JobRecord := TJobRecord.Create;
          JobRecord.RecType         := 'STEP';
          JobRecord.Description     := 'Calibrazione DUT a inclinazione X=59,9°  Y=-0,1°';
          JobRecord.urlDBtecnico    := '\\netserver\06.Qualità\SGQ2_attività\01_Mansionari_IOP_MOD\04.Produzione\02.IOP\PRD.150.html';
          JobRecord.runSetup        := nil;
          JobRecord.StepCardPanel   := nil;
          JobRecord.cardIdx         := -1;
          JobRecord.runStep         := RunStep_B8;
          JobRecord.GroupFase      := 'Fase_B';
          itemIdx := JobList.AddObject('Step_8', TJobRecord(JobRecord));
          JobRecord.ListIndex       := itemIdx;
          inc(numSteps);

          JobRecord := TJobRecord.Create;
          JobRecord.RecType         := 'STEP';
          JobRecord.Description     := 'Calibrazione DUT a inclinazione X=-60,1°  Y=0,1°';
          JobRecord.urlDBtecnico    := '\\netserver\06.Qualità\SGQ2_attività\01_Mansionari_IOP_MOD\04.Produzione\02.IOP\PRD.150.html';
          JobRecord.runSetup        := nil;
          JobRecord.StepCardPanel   := nil;
          JobRecord.cardIdx         := -1;
          JobRecord.runStep         := RunStep_B9;
          JobRecord.GroupFase      := 'Fase_B';
          itemIdx := JobList.AddObject('Step_9', TJobRecord(JobRecord));
          JobRecord.ListIndex       := itemIdx;
          inc(numSteps);


    JobRecord := TJobRecord.Create;
    JobRecord.RecType         := 'FASE';
    JobRecord.Description     := 'Test validazione DUT su asse X';
    JobRecord.urlDBtecnico    := urlDatasheet;
    JobRecord.runSetup        := SetupFase_C;
    JobRecord.runStep         := nil;
    JobRecord.StepCardPanel   := nil;            // Le Fasi non hanno subCards, mentre gli steps potrebbero.
    JobRecord.cardIdx         := -1;             // nemmeno un subCard index.
    JobRecord.GroupFase       := 'Fase_C';
    itemIdx := JobList.AddObject('Fase_C', TJobRecord(JobRecord));
    JobRecord.ListIndex       := itemIdx;
    inc(numFasi);

          JobRecord := TJobRecord.Create;
          JobRecord.RecType         := 'STEP';
          JobRecord.Description     := 'Validazione DUT a inclinazione X=-60°';
          JobRecord.urlDBtecnico    := urlDatasheet;
          JobRecord.runSetup        := nil;
          JobRecord.StepCardPanel   := nil;
          JobRecord.cardIdx         := -1;
          JobRecord.runStep         := RunStep_C10;
          JobRecord.GroupFase      := 'Fase_C';
          itemIdx := JobList.AddObject('Step_10', TJobRecord(JobRecord));
          JobRecord.ListIndex       := itemIdx;
          inc(numSteps);

          JobRecord := TJobRecord.Create;
          JobRecord.RecType         := 'STEP';
          JobRecord.Description     := 'Validazione DUT a inclinazione X=-30°';
          JobRecord.urlDBtecnico    := urlDatasheet;
          JobRecord.runSetup        := nil;
          JobRecord.StepCardPanel   := nil;
          JobRecord.cardIdx         := -1;
          JobRecord.runStep         := RunStep_C11;
          JobRecord.GroupFase      := 'Fase_C';
          itemIdx := JobList.AddObject('Step_11', TJobRecord(JobRecord));
          JobRecord.ListIndex       := itemIdx;
          inc(numSteps);

          JobRecord := TJobRecord.Create;
          JobRecord.RecType         := 'STEP';
          JobRecord.Description     := 'Validazione DUT a inclinazione X=-15°';
          JobRecord.urlDBtecnico    := urlDatasheet;
          JobRecord.runSetup        := nil;
          JobRecord.StepCardPanel   := nil;
          JobRecord.cardIdx         := -1;
          JobRecord.runStep         := RunStep_C12;
          JobRecord.GroupFase      := 'Fase_C';
          itemIdx := JobList.AddObject('Step_12', TJobRecord(JobRecord));
          JobRecord.ListIndex       := itemIdx;
          inc(numSteps);

          JobRecord := TJobRecord.Create;
          JobRecord.RecType         := 'STEP';
          JobRecord.Description     := 'Validazione DUT a inclinazione X=0°';
          JobRecord.urlDBtecnico    := urlDatasheet;
          JobRecord.runSetup        := nil;
          JobRecord.StepCardPanel   := nil;
          JobRecord.cardIdx         := -1;
          JobRecord.runStep         := RunStep_C13;
          JobRecord.GroupFase      := 'Fase_C';
          itemIdx := JobList.AddObject('Step_13', TJobRecord(JobRecord));
          JobRecord.ListIndex       := itemIdx;
          inc(numSteps);

          JobRecord := TJobRecord.Create;
          JobRecord.RecType         := 'STEP';
          JobRecord.Description     := 'Validazione DUT a inclinazione X=15°';
          JobRecord.urlDBtecnico    := urlDatasheet;
          JobRecord.runSetup        := nil;
          JobRecord.StepCardPanel   := nil;
          JobRecord.cardIdx         := -1;
          JobRecord.runStep         := RunStep_C14;
          JobRecord.GroupFase      := 'Fase_C';
          itemIdx := JobList.AddObject('Step_14', TJobRecord(JobRecord));
          JobRecord.ListIndex       := itemIdx;
          inc(numSteps);

          JobRecord := TJobRecord.Create;
          JobRecord.RecType         := 'STEP';
          JobRecord.Description     := 'Validazione DUT a inclinazione X=30°';
          JobRecord.urlDBtecnico    := urlDatasheet;
          JobRecord.runSetup        := nil;
          JobRecord.StepCardPanel   := nil;
          JobRecord.cardIdx         := -1;
          JobRecord.runStep         := RunStep_C15;
          JobRecord.GroupFase      := 'Fase_C';
          itemIdx := JobList.AddObject('Step_15', TJobRecord(JobRecord));
          JobRecord.ListIndex       := itemIdx;
          inc(numSteps);

          JobRecord := TJobRecord.Create;
          JobRecord.RecType         := 'STEP';
          JobRecord.Description     := 'Validazione DUT a inclinazione X=60°';
          JobRecord.urlDBtecnico    := urlDatasheet;
          JobRecord.runSetup        := nil;
          JobRecord.StepCardPanel   := nil;
          JobRecord.cardIdx         := -1;
          JobRecord.runStep         := RunStep_C16;
          JobRecord.GroupFase      := 'Fase_C';
          itemIdx := JobList.AddObject('Step_16', TJobRecord(JobRecord));
          JobRecord.ListIndex       := itemIdx;
          inc(numSteps);


    JobRecord := TJobRecord.Create;
    JobRecord.RecType         := 'FASE';
    JobRecord.Description     := 'Test validazione DUT su asse Y';
    JobRecord.urlDBtecnico    := '\\netserver\06.Qualità\SGQ2_attività\01_Mansionari_IOP_MOD\04.Produzione\02.IOP\PRD.161.html';
    JobRecord.runSetup        := SetupFase_D;
    JobRecord.runStep         := nil;
    JobRecord.StepCardPanel   := nil;            // Le Fasi non hanno subCards, mentre gli steps potrebbero.
    JobRecord.cardIdx         := -1;             // nemmeno un subCard index.
    JobRecord.GroupFase       := 'Fase_D';
    itemIdx := JobList.AddObject('Fase_D', TJobRecord(JobRecord));
    JobRecord.ListIndex       := itemIdx;
    inc(numFasi);

          JobRecord := TJobRecord.Create;
          JobRecord.RecType         := 'STEP';
          JobRecord.Description     := 'Validazione DUT a inclinazione Y=-60°';
          JobRecord.urlDBtecnico    := urlDatasheet;
          JobRecord.runSetup        := nil;
          JobRecord.StepCardPanel   := nil;
          JobRecord.cardIdx         := -1;
          JobRecord.runStep         := RunStep_D17;
          JobRecord.GroupFase      := 'Fase_D';
          itemIdx := JobList.AddObject('Step_17', TJobRecord(JobRecord));
          JobRecord.ListIndex       := itemIdx;
          inc(numSteps);

          JobRecord := TJobRecord.Create;
          JobRecord.RecType         := 'STEP';
          JobRecord.Description     := 'Validazione DUT a inclinazione Y=-30°';
          JobRecord.urlDBtecnico    := urlDatasheet;
          JobRecord.runSetup        := nil;
          JobRecord.StepCardPanel   := nil;
          JobRecord.cardIdx         := -1;
          JobRecord.runStep         := RunStep_D18;
          JobRecord.GroupFase      := 'Fase_D';
          itemIdx := JobList.AddObject('Step_18', TJobRecord(JobRecord));
          JobRecord.ListIndex       := itemIdx;
          inc(numSteps);

          JobRecord := TJobRecord.Create;
          JobRecord.RecType         := 'STEP';
          JobRecord.Description     := 'Validazione DUT a inclinazione Y=-15°';
          JobRecord.urlDBtecnico    := urlDatasheet;
          JobRecord.runSetup        := nil;
          JobRecord.StepCardPanel   := nil;
          JobRecord.cardIdx         := -1;
          JobRecord.runStep         := RunStep_D19;
          JobRecord.GroupFase      := 'Fase_D';
          itemIdx := JobList.AddObject('Step_19', TJobRecord(JobRecord));
          JobRecord.ListIndex       := itemIdx;
          inc(numSteps);

          JobRecord := TJobRecord.Create;
          JobRecord.RecType         := 'STEP';
          JobRecord.Description     := 'Validazione DUT a inclinazione Y=0°';
          JobRecord.urlDBtecnico    := urlDatasheet;
          JobRecord.runSetup        := nil;
          JobRecord.StepCardPanel   := nil;
          JobRecord.cardIdx         := -1;
          JobRecord.runStep         := RunStep_D20;
          JobRecord.GroupFase      := 'Fase_D';
          itemIdx := JobList.AddObject('Step_20', TJobRecord(JobRecord));
          JobRecord.ListIndex       := itemIdx;
          inc(numSteps);

          JobRecord := TJobRecord.Create;
          JobRecord.RecType         := 'STEP';
          JobRecord.Description     := 'Validazione DUT a inclinazione Y=15°';
          JobRecord.urlDBtecnico    := urlDatasheet;
          JobRecord.runSetup        := nil;
          JobRecord.StepCardPanel   := nil;
          JobRecord.cardIdx         := -1;
          JobRecord.runStep         := RunStep_D21;
          JobRecord.GroupFase      := 'Fase_D';
          itemIdx := JobList.AddObject('Step_21', TJobRecord(JobRecord));
          JobRecord.ListIndex       := itemIdx;
          inc(numSteps);

          JobRecord := TJobRecord.Create;
          JobRecord.RecType         := 'STEP';
          JobRecord.Description     := 'Validazione DUT a inclinazione Y=30°';
          JobRecord.urlDBtecnico    := urlDatasheet;
          JobRecord.runSetup        := nil;
          JobRecord.StepCardPanel   := nil;
          JobRecord.cardIdx         := -1;
          JobRecord.runStep         := RunStep_D22;
          JobRecord.GroupFase      := 'Fase_D';
          itemIdx := JobList.AddObject('Step_22', TJobRecord(JobRecord));
          JobRecord.ListIndex       := itemIdx;
          inc(numSteps);

          JobRecord := TJobRecord.Create;
          JobRecord.RecType         := 'STEP';
          JobRecord.Description     := 'Validazione DUT a inclinazione Y=60°';
          JobRecord.urlDBtecnico    := urlDatasheet;
          JobRecord.runSetup        := nil;
          JobRecord.StepCardPanel   := nil;
          JobRecord.cardIdx         := -1;
          JobRecord.runStep         := RunStep_D23;
          JobRecord.GroupFase      := 'Fase_D';
          itemIdx := JobList.AddObject('Step_23', TJobRecord(JobRecord));
          JobRecord.ListIndex       := itemIdx;
          inc(numSteps);


    JobRecord := TJobRecord.Create;
    JobRecord.RecType         := 'FASE';
    JobRecord.Description     := 'Configurazione per Verifica finale Paramentri DUT';
    JobRecord.urlDBtecnico    := urlDatasheet;
    JobRecord.runSetup        := SetupFase_E;
    JobRecord.runStep         := nil;            // la Fase ha solo il Setup !
    JobRecord.StepCardPanel   := nil;            // Le Fasi non hanno subCards, mentre gli steps potrebbero.
    JobRecord.cardIdx         := -1;             // nemmeno un subCard index.
    JobRecord.GroupFase       := 'Fase_E';
    itemIdx := JobList.AddObject('Fase_E', TJobRecord(JobRecord));
    JobRecord.ListIndex       := itemIdx;  // per il passaggio diretto da Nodo TreeView ad item della stringList.
    inc(numFasi);

          JobRecord := TJobRecord.Create;
          JobRecord.RecType         := 'STEP';
          JobRecord.Description     := 'Lista parametri di Default e Calibrazione 2 Assi';
          JobRecord.urlDBtecnico    := '';
          JobRecord.runSetup        := nil;
          JobRecord.StepCardPanel   := nil;
          JobRecord.cardIdx         := -1;
          JobRecord.runStep         := RunStep_E24;
          JobRecord.GroupFase      := 'Fase_E';
          itemIdx := JobList.AddObject('Step_24', TJobRecord(JobRecord));
          JobRecord.ListIndex       := itemIdx;
          inc(numSteps);

          JobRecord := TJobRecord.Create;
          JobRecord.RecType         := 'STEP';
          JobRecord.Description     := 'Verifica Lista parametri e contenuto DUT';
          JobRecord.urlDBtecnico    := '';
          JobRecord.runSetup        := nil;
          JobRecord.StepCardPanel   := nil;
          JobRecord.cardIdx         := -1;
          JobRecord.runStep         := RunStep_E25;
          JobRecord.GroupFase      := 'Fase_E';
          itemIdx := JobList.AddObject('Step_25', TJobRecord(JobRecord));
          JobRecord.ListIndex       := itemIdx;
          inc(numSteps);

          JobRecord := TJobRecord.Create;
          JobRecord.RecType         := 'STEP';
          JobRecord.Description     := 'Scrivi parametri di default nel DUT';
          JobRecord.urlDBtecnico    := '';
          JobRecord.runSetup        := nil;
          JobRecord.StepCardPanel   := nil;
          JobRecord.cardIdx         := -1;
          JobRecord.runStep         := RunStep_E26;
          JobRecord.GroupFase      := 'Fase_E';
          itemIdx := JobList.AddObject('Step_26', TJobRecord(JobRecord));
          JobRecord.ListIndex       := itemIdx;
          inc(numSteps);

          JobRecord := TJobRecord.Create;
          JobRecord.RecType         := 'STEP';
          JobRecord.Description     := 'Salva parametri in EEProm del DUT';
          JobRecord.urlDBtecnico    := '';
          JobRecord.runSetup        := nil;
          JobRecord.StepCardPanel   := nil;
          JobRecord.cardIdx         := -1;
          JobRecord.runStep         := RunStep_E27;
          JobRecord.GroupFase      := 'Fase_E';
          itemIdx := JobList.AddObject('Step_27', TJobRecord(JobRecord));
          JobRecord.ListIndex       := itemIdx;
          inc(numSteps);


    JobRecord := TJobRecord.Create;
    JobRecord.RecType         := 'FASE';
    JobRecord.Description     := 'Aggiornamenti DB e Report di fine Sessione';
    JobRecord.urlDBtecnico    := '';
    JobRecord.runSetup        := SetupFase_F;
    JobRecord.runStep         := nil;
    JobRecord.StepCardPanel   := nil;            // Le Fasi non hanno subCards, mentre gli steps potrebbero.
    JobRecord.cardIdx         := -1;             // nemmeno un subCard index.
    JobRecord.GroupFase       := 'Fase_F';
    itemIdx := JobList.AddObject('Fase_F', TJobRecord(JobRecord));
    JobRecord.ListIndex       := itemIdx;
    inc(numFasi);

          JobRecord := TJobRecord.Create;
          JobRecord.RecType         := 'STEP';
          JobRecord.Description     := 'Note di FINE Procedura';
          JobRecord.urlDBtecnico    := '';
          JobRecord.runSetup        := nil;
          JobRecord.StepCardPanel   := pnlCardsF;   // questo Step HA una subCard specifica, ovvero gli Steps di questa Fase hanno GUI su Cards distinte.
          JobRecord.cardIdx         := 2;           // il ActiveCardIndex per questo step è 0.
          JobRecord.runStep         := RunStep_F28;
          JobRecord.GroupFase      := 'Fase_F';
          itemIdx := JobList.AddObject('Step_28', TJobRecord(JobRecord));
          JobRecord.ListIndex       := itemIdx;
          inc(numSteps);

          JobRecord := TJobRecord.Create;
          JobRecord.RecType         := 'STEP';
          JobRecord.Description     := 'Salva dati su database server';
          JobRecord.urlDBtecnico    := '';
          JobRecord.runSetup        := nil;
          JobRecord.StepCardPanel   := pnlCardsF;   // questo Step HA una subCard specifica, ovvero gli Steps di questa Fase hanno GUI su Cards distinte.
          JobRecord.cardIdx         := 0;           // il ActiveCardIndex per questo step è 0.
          JobRecord.runStep         := RunStep_F29;
          JobRecord.GroupFase      := 'Fase_F';
          itemIdx := JobList.AddObject('Step_29', TJobRecord(JobRecord));
          JobRecord.ListIndex       := itemIdx;
          inc(numSteps);

          JobRecord := TJobRecord.Create;
          JobRecord.RecType         := 'STEP';
          JobRecord.Description     := 'Produzione Test report di validazione EOL';
          JobRecord.urlDBtecnico    := '';
          JobRecord.runSetup        := nil;
          JobRecord.StepCardPanel   := pnlCardsF;    // questo Step HA una subCard specifica, ovvero gli Steps di questa Fase hanno GUI su Cards distinte.
          JobRecord.cardIdx         := 1;            // il ActiveCardIndex per questo step è 1.
          JobRecord.runStep         := RunStep_F30;
          JobRecord.GroupFase      := 'Fase_F';
          itemIdx := JobList.AddObject('Step_30', TJobRecord(JobRecord));
          JobRecord.ListIndex       := itemIdx;
          inc(numSteps);


    JobRecord := TJobRecord.Create;
    JobRecord.RecType         := 'FASE';
    JobRecord.Description     := 'Chiusura connessioni e interfacce HW';
    JobRecord.urlDBtecnico    := '';
    JobRecord.runSetup        := SetupFase_G;
    JobRecord.runStep         := nil;
    JobRecord.StepCardPanel   := nil;            // Le Fasi non hanno subCards, mentre gli steps potrebbero.
    JobRecord.cardIdx         := -1;             // nemmeno un subCard index.
    JobRecord.GroupFase       := 'Fase_G';
    itemIdx := JobList.AddObject('Fase_G', TJobRecord(JobRecord));
    JobRecord.ListIndex       := itemIdx;
    inc(numFasi);

    //lenJobList := JobList.Count;   // ridondante.
    result := True;
end;


//#####################################################################################################
//
//   Sezione con l'implementazione vera e propria delle funzioni (FASI/STEP) per ogni test da eseguire.
//
////############################### FASI

function Toiac3Fasi2ax.SetupFase_A(var JobParams: TJobParams): boolean;
var
  vReportPath, kXls: string;
  addr, cmdFile, eolReport: string;
  i, nodes: integer;
begin
    // Carica in lviewCanCMDs tutti i Comandi CAN con relativi Alias e Defaults definiti in PCAN_CMDs.xlsx,
    // ma copia anche nel Dictionary  'dicCanCmds' con key = Alias  per usi successivi.
    // La lviewCanCMDs è in panel Admin, quindi non è visibile a User standard.
    // Il file name caricato è come da ini, solitamente 'PCAN_CMDs.xlsx'  e dal Sheet 'PCAN_AllCMDs'

  //mainForm.mmLog.lines.add('run SetupFase_A...');
    result := False;
    JobParams.retCode := resUNDEFINED;
    JobParams.retMsg := _Incompleto_;
    // al momento JobParams è veicolo solo per parametri di ritorno.
    // Ogni Fase/Step ha bisogno di notificare dati e progress  su Frame o Form, quindi è inevitabile
    // dover accedere a GUI controls da dentro la classe TEolSession.
    // La messagistica predefinita (JobRecord) è gestita dal chiamante, ma qui vanno visualizzate per forza
    // tutte le informazioni variabili !

    // init Com Channels
//  LogStep(272,272, 272, 'Inizializzazione Comm Channels...');
//  Scan4ComInterfaces();

    // init CAN Channels
    LogStep(273,273, 273, 'Inizializzazione CAN Channels...');
    CAN_Release(CHref);                     //
    if CHdut.Handle <> CHref.Handle then    // close if already open (eg. Retry step)
        CAN_Release(CHdut);                 //

    // Reset all Unknown prima di scan.
    cboxCanREF.ItemIndex := -1;
    cboxCanREF.Style.color := clBtnface;
    cboxCanDUT.ItemIndex := -1;
    cboxCanDUT.Style.color := clBtnface;
    cboxCanCH.ItemIndex := -1;
    cboxCanCH.Style.color := clBtnface;

    if not Scan4PCanInterfaces() then begin    // Cerca Canali PCAN disponibili
        SetAdminMode( True );                  // forza subito ADMIN_mode per permettere debug.
        JobParams.retMsg := 'ENUM PCAN Channels ERROR: '+ lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        LogStep(274,274, 274, JobParams.retMsg);
        exit;
    end;

    // init items di selezione CAN Busfields che sarano usati in tutti gli steps!
    LoadCombo_4CanInterfaces( cboxCanREF ); // carica la lista in cboxCanREF del tab Fase_A. setta anche CHref.Handle al primo item in lista
    LoadCombo_4CanInterfaces( cboxCanDUT ); // carica la lista in cboxCanDUT del tab Fase_A. setta anche CHdut.Handle al primo item in lista
    LoadCombo_4CanInterfaces( cboxCanCH );  // carica la lista in cboxCanDUT del tab Admin. setta anche CHcan.Handle al primo item in lista

    for i := 0 to CanChannels.count -1 do begin
        addr := CanChannels.Strings[i];
        if addr = '51h' then begin
            LogStep(275,275, 275, #9#9+ Addr +'  >> PCAN per i Reference');
            cboxCanREF.ItemIndex := i;
            cboxCanREF.Style.color := clLime;
        end
        else if addr = '52h' then begin
            LogStep(276,276, 276, #9#9+ Addr +'  >> PCAN per il Device in Test');
            cboxCanDUT.ItemIndex := i;
            cboxCanDUT.Style.color := clLime;
        end
        else begin
            // se non è dei precedenti, assumo sia un PCAN con ID imprevisto,
            // tipo altre interfacce, come step motor, ecc.
            LogStep(277,277, 277, #9#9+ Addr +'  >> PCAN Address imprevisto !');
            cboxCanREF.Style.color := clYellow;
            cboxCanDUT.Style.color := clYellow;
            JobParams.retMsg := 'ERROR PCan Address imprevisto >>  '+ Addr;
            JobParams.retCode := resERROR;
            LogStep(278,278, 278, JobParams.retMsg);
            exit;   // meglio fermare e far correggere.
        end;
    end;
    // Qui arriva con almeno 1 PCAN interface rilevata.
    cboxCanCH.ItemIndex := 0;                          // CHcan.Handle già settato in LoadCombo_4CanInterfaces()
    cboxCanCH.Style.color := clLime;
    // Per EOL è necessario almeno una PCAN interface, non necessariamente $51
    if CanChannels.count = 1 then begin
        // REFs e DUT vanno su stesso CanBus.
        cboxCanREF.ItemIndex := 0;
        cboxCanREF.Style.color := clLime;
        cboxCanDUT.ItemIndex := 0;
        cboxCanDUT.Style.color := clLime;
        LogStep(279,279, 279, 'WARNING - PCAN bus '+ CanChannels.Strings[0] +' è Condiviso fra REF e DUT !');
        // Override Handle
        CHref.Handle := strCanValue2integer( CanChannels.Strings[0] );    // il primo in lista, solitamente $51
        CHdut.Handle := CHref.Handle;                                     // ma lo è già dopo LoadCombo_4CanInterfaces()
    end
    else begin // REFs e DUT sono su CanBus separati.
        LogStep(280,280, 280, 'INFO - PCAN bus Channels dedicati per REF e DUT.');
        // Override Handle
        CHref.Handle := strCanValue2integer( CanChannels.Strings[0] );    // il primo CH in lista, solitamente $51
        CHdut.Handle := strCanValue2integer( CanChannels.Strings[1] );    // override secondo CH in lista, solitamente $52
    end;

    edtToWrite.text := '';              // Clean
    edtToWrite.Color := clBtnFace;      //
    // ferma Timer ricezione ciclica valori da DUT/REF, se non è già
    tmrCanRead.enabled := False;        // per sicurezza

    // Verifica se file Report è già stato prodotto ?
    //rptFile := vReportPath +'\'+ Sessione.CodiceProdotto +'_'+ Sessione.Lotto +'\'+ Sessione.CodiceProdotto +'_EOL_'+ Sessione.Lotto + Sessione.Serial +'.xls';
    LogStep(281,281, 281, 'Verifica se Report EOL già esistente...');
    //rptFile := vReportPath +'\'+ Sessione.CodiceProdotto +'_'+ Sessione.Lotto +'\'+ Sessione.CodiceProdotto +'_EOL_'+ Sessione.Lotto + Sessione.Serial +'.xls';
    kXls := '.xls';
    vReportPath := iniReportPath;                // vReportPath as from .INI  (solitamente Netserver)
    // Verifica accessibilità cartella principale su server:  vReportPath from .INI
    if DirectoryExists( vReportPath ) then begin     // Server disponibile e report di produzione
        // Server disponibile, salva report in rete.
        // Verifica esistenza cartella del Lotto, se no la crea.
        eolReport := vReportPath + Sessione.CodiceProdotto +'_L'+ Sessione.Lotto +'\'+ Sessione.CodiceProdotto +'_EOL_';
        eolReport := eolReport + Sessione.Lotto +'#'+ Sessione.Serial + '*.*';  // con *.* invece di kXls, cerco anche presenza di revisioni.
        LogStep(282,282, 282, #9+ eolReport);
        // Verifico se Nome-File Report già esistente.
        //   \\netserver\04.Produzione\01.PRD EL\Datalog\OI200047_L4\OI200047_EOL_4#12.xls
        if dmBase.FindFile( eolReport ) then begin
            // se il file esiste già, avviso utente che verrà sovrascritto.
            LogSysEvent(svWARNING, 1140, 'Report EOL già esistente per il Dut corrente !');
            MessageDlg('ATTENZIONE:' +sLineBreak+
                       'il Report EOL risulta già esistente per questo DUT !' +sLineBreak+ eolReport +sLineBreak+sLineBreak+
                       'Se è necessario conservarlo, fare una copia ADESSO !', mtWarning, [mbOk], 0, mbOk);
        end
        else
            LogStep(283,283, 283, 'OK report Non trovato.');
    end
    else begin
        // Server non disponibile o report di prova, salva sempre in locale.
        MessageDlg('ATTENZIONE:' +sLineBreak+ 'il Path  '+ vReportPath +sLineBreak+
                   'Non è raggiungibile !' +sLineBreak+ 'Verifica e riprova.', mtWarning, [mbOk], 0, mbOk);
        JobParams.retMsg := 'ERRORE Report path or Server Not available: '+ vReportPath;
        JobParams.retCode := resERROR;
        LogStep(284,284, 284, JobParams.retMsg);
        exit;                           // passa sempre per finally
    end;

    InitializeProtection();  // CAN Initialize the Critical Section.
    if not CAN_Init(CHref) then begin
        JobParams.retMsg := 'REF PCAN_ERROR: '+ lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        LogStep(285,285, 285, JobParams.retMsg);
        exit;
    end;
  	// CAN Ref interface is Ready
    LogStep(286,286, 286, 'OK REF PCAN Init.');
  //if numCanCH > 1 then          // CHref = CHdut
    if CHdut.Handle <> CHref.Handle then
        if not CAN_Init(CHdut) then begin
            JobParams.retMsg := 'DUT PCAN_ERROR: '+ lastCAN_ErrMsg;
            JobParams.retCode := resERROR;
            LogStep(287,287, 287, JobParams.retMsg);
            exit;
        end;
  	// CAN Dut interface is Ready
    LogStep(288,288, 288, 'PCAN Init OK'); //PCAN_ERROR_OK.

    // Definisce tipo di parametri da caricare in listview.
    cbxListaDefaultsCom.Checked := True;    // qui solo parametri DEF.
    cbxListaDefaults1ax.checked := False;   //
    cbxListaDefaults2ax.checked := True;    //
    cbxListaDefaultsEND.Checked := False;   //
    cbxListaSN.Checked       := True;       //
    cbxListaCalib1ax.Checked := False;
    cbxListaCalib2ax.Checked := False;

    edtToWrite.text := '';
    edtToWrite.Color := clBtnFace;

    cmdFile := kAppPath + dmBase.vCANcmdSource;
    LogStep(289,289, 289, 'Loading CMD file: '+ cmdFile);
    if not Load_PCAN_CMDs_fromXLS( cmdFile ) then begin
        JobParams.retMsg := 'ERROR on Load CAN Command list: '+ cmdFile;
        JobParams.retCode := resERROR;
        LogStep(290,290, 290, JobParams.retMsg);
        exit;
    end;
    LogStep(291,291, 291, 'OK PCAN Commands loaded.');

    // come prerequisito per tutti gli step della fase A resetto tutti i nodes !
    // questo stoppa  anche eventuali cyclical transmission dai REF.
    // Eseguo però un Enum nodes che resetta anche Tutti i device e flush dei Rx Buffer
    // solo per verificare presenza alimentazione

    LogStep(292,292, 292, 'RESET ALL nodes...');
    if not CAN_SendNMT( CHref, cmdResetNode, '00h' ) then begin  // Reset ALL Ref nodes!
        JobParams.retMsg := 'ERROR on Reset CHref Nodes: '+ lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        LogStep(293,293, 293, lastCAN_ErrMsg);
        exit;
    end;
    if CHdut.Handle <> CHref.Handle then
        if not CAN_SendNMT( CHdut, cmdResetNode, '00h' ) then begin  // Reset ALL Dut nodes!
            JobParams.retMsg := 'ERROR on Reset CHdut Nodes: '+ lastCAN_ErrMsg;
            JobParams.retCode := resERROR;
            LogStep(294,294, 294, lastCAN_ErrMsg);
            exit;
        end;
    LogStep(295,295, 295, 'ALL nodes Reset Successfully');
    dmBase.wait(500);                         // sembra necessario, forse il delay in CAN_SendNMT() non basta...!
    LogStep(296,296, 296, 'FLUSH All CAN queue...');
    if not CAN_FlushRxBuffer(CHref) then begin
        JobParams.retMsg := 'ERROR on Flush RefCAN buffer: '+ lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        if lastCAN_Status = PCAN_ERROR_BUSHEAVY then begin
            JobParams.retMsg := JobParams.retMsg + sLineBreak + 'Possibile mancanza di alimentazione o BaudRate errato !';
            // necessario Release+Init, perchè CAN_Reset() non risolve il lock, ma viene fatto al Retry !
            //CAN_Release();
            //CAN_Init();
        end;
        LogStep(297,297, 297,  JobParams.retMsg );
        exit;
    end;
    if CHdut.Handle <> CHref.Handle then
        if not CAN_FlushRxBuffer(CHdut) then begin
            JobParams.retMsg := 'ERROR on Flush DutCAN buffer: '+ lastCAN_ErrMsg;
            JobParams.retCode := resERROR;
            if lastCAN_Status = PCAN_ERROR_BUSHEAVY then begin
                JobParams.retMsg := JobParams.retMsg + sLineBreak + 'Possibile mancanza di alimentazione o BaudRate errato !';
                // necessario Release+Init, perchè CAN_Reset() non risolve il lock, ma viene fatto al Retry !
                //CAN_Release();
                //CAN_Init();
            end;
            LogStep(298,298, 298,  JobParams.retMsg );
            exit;
        end;
    LogStep(299,299, 299, 'ALL CAN queue Flushed Successfully.');    // preparato rx queue clean.

    // Eseguo però un Enum nodes perchè resetta Tutti i device e flush di Rx Buffer.
    // cerca almeno i Refs su Ch0
    // se cho=ch1 cerca anche Dut
    LogStep(300,300, 300, 'SCAN for active Node IDs...');

    // Search IDs...
    CanNodes.Clear;   // reset lista nodi a cura chiamante per permettere di accumulare scan multipli su più Canali.
    // Reset ALL nodes and wait for IDs
    // dopo boot il DUT invia solo la Boot frame   COB:70A   Len:1   Data:00 00 00 00 00 00 00 00
    if not CAN_EnumNodes(CHref, nodes) then begin        // Reset ALL nodes and wait for IDs
        JobParams.retMsg := 'ENUM CAN Nodes ERROR1: '+ lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        LogStep(301,301, 301, JobParams.retMsg);
        exit;
    end;
    LogStep(302,302, 302, #9'Found '+ nodes.ToString +' Nodes on PCAN '+ CHref.Handle.ToHexString);
    if CHdut.Handle <> CHref.Handle then begin
        if not CAN_EnumNodes(CHdut, nodes) then begin        // Reset ALL nodes and wait for IDs
            JobParams.retMsg := 'ENUM CAN Nodes ERROR2: '+ lastCAN_ErrMsg;
            JobParams.retCode := resERROR;
            LogStep(303,303, 303, JobParams.retMsg);
            exit;
        end;
        LogStep(304,304, 304, #9'Found '+ nodes.ToString +' Nodes on PCAN '+ CHdut.Handle.ToHexString +'h');
    end;
    if CanNodes.count = 0 then begin
        JobParams.retMsg := 'NO CAN Nodes Found !';
        JobParams.retCode := resERROR;
        LogStep(305,305, 305, JobParams.retMsg);
        exit;
    end;
    LogStep(306,306, 306, 'Total of '+ CanNodes.count.ToString +' CAN Nodes on '+numCanCH.toString+' PCan interface.');

    (*
    edtRef2Addr.Text := '';
    edtRef1Addr.Text := '';                      // init fields che sarano usati in tutti gli steps!
    edtDutAddr.Text := '';
    *)
    for i := 0 to CanNodes.count-1 do begin
        addr := CanNodes.Strings[i];
        if addr = '7E' then begin
            LogStep(307,307, 307, #9#9+ Addr +'h  >> node REF 1 asse');
            edtRef1Addr.Text  := addr +'h';
            edtRef1Addr.color := clLime;
        end
        else if addr = '7F' then begin
            LogStep(308,308, 308, #9#9+ Addr +'h  >> node REF 2 assi');
            edtRef2Addr.Text  := addr +'h';
            edtRef2Addr.color := clLime;
        end
        else if addr = '0A' then begin
            LogStep(309,309, 309, #9#9+ Addr +'h  >> node DUT');
            edtDutAddr.Text  := addr +'h';
            edtDutAddr.Style.color := clLime;
            edtCanID.Text  := addr +'h';
        end
        else begin
            // se non è dei precedenti, assumo sia un DUT con ID errato, che verrà corretto.
            // In futuro potrò anche avere altri nodes, come step motor, ecc.
            LogStep(310,310, 310, #9#9+ Addr +'h  >> node DUT ?');
            edtDutAddr.Text  := addr +'h';
            edtDutAddr.Style.color := clLime;
            edtCanID.Text  := addr +'h';
        end;
    end;
    // Per EOL 2 Assi sono necessari Ref=$7F e Dut=$0A o altro dut id.
    if (edtRef2Addr.Text = '') or (edtDutAddr.Text = '') then begin
        JobParams.retMsg := 'ERROR Necessary NODES not Found !';
        JobParams.retCode := resERROR;
        LogStep(311,311, 311, JobParams.retMsg);
        exit;
    end;

    // Cancello subito tracce di eventuale giro precedente.
    lviewCFG.Items.BeginUpdate;
    lviewCFG.Items.Clear;
    lviewCFG.InnerListView.GridLines := True; // mostra grid, default è off.
    lviewCFG.Items.EndUpdate;

    // Verifica se DUT è già calibrato 2 assi ?
    if DutCalibrato( edtDutAddr.Text ) then begin
        LogSysEvent(svWARNING, 1050, 'DUT Already configured!');
        if MessageDlg('ATTENZIONE:' +sLineBreak+
                      'il DUT sembra già calibrato !' +sLineBreak+
                      'Vuoi procedere comunque ?', mtWarning, [mbYes, mbNo], 0, mbNo) <> mrYes then
            exit;
    end;

    JobParams.retMsg := 'OK Setup CAN network.';
    JobParams.retCode := resOK;
    result := True;
end;

function Toiac3Fasi2ax.SetupFase_B(var JobParams: TJobParams): boolean;
var
  Can: TCanFrame;
  i: integer;
begin
{   come setup di fase B serve controllare status PCAN, presenza dei Reference device
    e innescare invio *continuo* delle misure di riferimento settando EventTimer e START !
}
    result := False;
    JobParams.retCode := resUNDEFINED;
    JobParams.retMsg := _Incompleto_;
    tmrCanRead.enabled := False;                          // per sicurezza

    gaugeXRange.ValueStart := 0;          // Hide marker dei target degs
    gaugeXRange.ValueEnd   := 0;
    gaugeYRange.ValueStart := 0;
    gaugeYRange.ValueEnd   := 0;

    if not CAN_GetStatus(CHref) then begin
        JobParams.retMsg := 'PCAN Ref Status ERROR: '+ lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        LogStep(312,312, 312, JobParams.retMsg);
        exit;
    end;
    if CHdut.Handle <> CHref.Handle then
        if not CAN_GetStatus(CHdut) then begin
            JobParams.retMsg := 'PCAN Dut Status ERROR: '+ lastCAN_ErrMsg;
            JobParams.retCode := resERROR;
            LogStep(313,313, 313, JobParams.retMsg);
            exit;
        end;
  	// The channels are Ready
    LogStep(314,314, 314, 'Fase_B: PCAN STATUS OK');

    // prepara DUT e REF2 address per usi vari.
    vRef2Addr := strCanValue2integer(edtRef2Addr.text);
    vDutAddr := strCanValue2integer(edtDutAddr.text);

    edtRefX.color := clBtnface;
    edtRefY.color := clBtnface;
    lviewCalibVal.Items.Clear;   // Pulisce evntuali dati di giri precedenti.

    // Prima meglio un Reset del Ref2 e Dut, casomai fossero già in Start mode.
    LogStep(315,315, 315, 'RESET REF2 e DUT...');
  //if not CAN_SendNMT( cmdResetNode, '00h' ) then begin  // Reset ALL nodes!
    if not CAN_SendNMT( CHref, cmdResetNode, edtRef2Addr.text ) then begin
        JobParams.retMsg := 'REF2 ResetNode '+edtRef2Addr.text+' ERROR: '+ lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        LogStep(316,316, 316, lastCAN_ErrMsg);
        exit;
    end;
    // The Reset was successfully sent
    if not CAN_SendNMT( CHdut, cmdResetNode, edtDutAddr.text ) then begin
        JobParams.retMsg := 'DUT ResetNode '+edtDutAddr.text+' ERROR: '+ lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        LogStep(317,317, 317, lastCAN_ErrMsg);
        exit;
    end;
    LogStep(318,318, 318, 'ALL Nodes Reset Successfully');
    LogStep(319,319, 319, 'FLUSH queue All Can CHANNELS...');
    // clean boot frames.
    if not CAN_FlushRxBuffer(CHref) then begin
        LogStep(320,320, 320,  'ERROR on CHref Flush: '+lastCAN_ErrMsg );
        exit;
    end;
    if CHdut.Handle <> CHref.Handle then
        if not CAN_FlushRxBuffer(CHdut) then begin
            LogStep(321,321, 321,  'ERROR on CHdut Flush: '+lastCAN_ErrMsg );
            exit;
        end;
    LogStep(322,322, 322, 'OK FLUSH All CanCH Done.');  // preparato rx queue clean.

    // Verifica presenza device REF2: 0=none 1=1asse  2=2assi
    i := CAN_CheckREF(edtRef2Addr.text);                  // per REF serve comando specifico, diverso da ID node.
    if i <> 2 then begin
        JobParams.retMsg := 'ERROR Unexpected REF2 device Type.';
        JobParams.retCode := resERROR;
        LogStep(323,323, 323, JobParams.retMsg);
        exit;
    end;
    // qui il tipo Ref trovato a questo indirizzo è corretto per i successivi usi previsti !!
    LogStep(324,324, 324, 'OK REF type '+ i.toString+'Axes present on Address: '+ edtRef2Addr.text); //PCAN_ERROR_OK.
    // Verifica presenza device DUT
    if not CAN_CheckNodeID(edtDutAddr.text) then begin
        JobParams.retMsg := 'DUT ID ERROR: '+ lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        LogStep(325,325, 325, JobParams.retMsg);
        exit;
    end;
    LogStep(326,326, 326, 'OK DUT present on Address: '+ edtDutAddr.text); //PCAN_ERROR_OK.

    // Assicura che REF e DUT lavorino a uguale risoluzione in Centesimi di grado !
    // in precedenza qui verificavo solo la risoluzione, ma ora la configuro e salvo esplicitamente,
    // come prerequisito prima di avviare la calibrazione e anche successivi test !
    if not Force_Resolution4EOL() then begin
        JobParams.retMsg := lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        exit;
    end;

    // per START REF2 è necessario cyclical transmission già configurato su REF !
    Can.Command  := _Timer_for_cyclical_txd_;      // = Event timer for cyclical transmission [multiple of 1ms, 0 = disabled]
    Can.Data     := kTimer4cyclicalTxd;            // 250 msec.
    Can.NodeID   := edtRef2Addr.text;              // str ID del device
    if not CAN_WriteParam( Can ) then begin
        JobParams.retMsg := 'REF2 Write '+ _Timer_for_cyclical_txd_ +' ERROR: '+ lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        LogStep(327,327, 327, JobParams.retMsg);
        exit;
    end;
    // The parameter was Written successfully.
    LogStep(328,328, 328, 'OK Ref2 '+ _Timer_for_cyclical_txd_ +' STORED Successfully to NodeID:'+ edtRef2Addr.text );
    LogStep(329,329, 329, #9#9'Ref2 '+ _Timer_for_cyclical_txd_ +' = '+ kTimer4cyclicalTxd );
    CalibRecords.Clear;          // Clear removes all keys and values from the dictionary. The Count property is set to 0.
    TestXRecords.Clear;          // reset sortDictionary con records dati dei test, anche se qui non usati.
    TestYRecords.Clear;          // comunque se rieseguo la calibrazione, i dati di test non sono più validi !
    LogStep(330,330, 330, 'OK, completato Setup per FaseB.' );

    dmBase.wait(500);
    JobParams.retMsg := _Completato_;
    JobParams.retCode := resOK;
    result := True;
end;

function Toiac3Fasi2ax.SetupFase_C(var JobParams: TJobParams): boolean;
begin
    result := Setup4_CD( JobParams, 'X' );
end;
function Toiac3Fasi2ax.SetupFase_D(var JobParams: TJobParams): boolean;
begin
    result := Setup4_CD( JobParams, 'Y' );
end;

function Toiac3Fasi2ax.Force_Resolution4EOL: boolean;
begin
    // La risoluzione durante calibrazioni e collaudo deve restare a default = Centesimi di grado
    // e permanere anche con die NodeReset, quindi va pure salvata !
    result := False;
    try
        // serve leggere REF2_resolution per calcolo DEGs quando acquisisco samples via can read.
        if not CAN_GetResolution( edtRef2Addr.text, Ref2_resolution) then begin  // Delay read è 25 msec
            lastCAN_ErrMsg := 'REF2 Read '+ _Angular_Resolution_ +' ERROR: '+ lastCAN_ErrMsg;
            LogStep(331,331, 331,  lastCAN_ErrMsg );
            exit;
        end;
        // The Device Resolution [1,10,100,1000] was Read successfully !
        // NB: il valore Letto in Ref2_resolution è uguale alla risoluzione, 100 = Centesimi di°
        LogStep(332,332, 332, 'OK Ref2 '+ _Angular_Resolution_ +' READ Successfully from NodeID:'+ edtRef2Addr.text );
        LogStep(333,333, 333, #9#9'REF2 '+ _Angular_Resolution_ +' = '+ Ref2_resolution.ToString );
        edtResolREF2.text := Ref2_resolution.ToString;

        // serve leggere DUT_resolution per calcolo DEGs quando acquisisco samples via can read.
        if not CAN_GetResolution( edtDUTaddr.text, DUT_resolution) then begin // Delay standard è 25 msec
            lastCAN_ErrMsg := 'DUT Read '+ _Angular_Resolution_ +' ERROR: '+ lastCAN_ErrMsg;
            LogStep(334,334, 334,  lastCAN_ErrMsg );
            exit;
        end;
        // The Device Resolution [1,10,100,1000] was Read successfully !
        // NB: il valore Letto in DUT_resolution è uguale alla risoluzione, 100 = Centesimi di°
        LogStep(335,335, 335, 'OK DUT '+ _Angular_Resolution_ +' READ Successfully from NodeID:'+ edtDutAddr.text );
        LogStep(336,336, 336, #9#9'DUT '+ _Angular_Resolution_ +' = '+ DUT_resolution.ToString );
        edtResolDUT.text := DUT_resolution.ToString;

        // NB:
        // per corrette comparazioni nei vari step è necessario che risoluzioni REF2 e DUT siano uguali e in Centesimi°.
        if REF2_resolution <> 100 then begin
            LogStep(337,337, 337, 'WARNING: Ref2 angular Resolution <> 1/100 [centesimi]  >>  will Override.' );
            if not CAN_SetResolution( edtRef2addr.text, degCentesimi) then begin // Delay standard è 25 msec
                // in Write uso TDegResolution perchè i valori non corrispondono come in Read !
                lastCAN_ErrMsg := 'REF2 write '+ _Angular_Resolution_ +' ERROR: '+ lastCAN_ErrMsg;
                LogStep(338,338, 338,  lastCAN_ErrMsg );
                exit;
            end;
            // The Device Resolution [1,10,{100},1000] was Written successfully ! (NB: valori invertiti rispetto a ciò che va in subidx!)
            REF2_resolution   := 100;
            edtResolREF2.text := REF2_resolution.ToString;
            LogStep(339,339, 339, 'OK new REF2 '+ _Angular_Resolution_ +' STORED Successfully to NodeID:'+ edtRef2addr.text );
            LogStep(340,340, 340, #9#9'REF2 '+ _Angular_Resolution_ +' = '+ REF2_resolution.ToString );
        end;
        if DUT_resolution <> 100 then begin
            LogStep(341,341, 341, 'WARNING: Dut angular Resolution <> 1/100 [centesimi]  >>  will Override.' );
            if not CAN_SetResolution( edtDUTaddr.text, degCentesimi) then begin // Delay standard è 25 msec
                // in Write uso TDegResolution perchè i valori non corrispondono come in Read !
                lastCAN_ErrMsg := 'DUT write '+ _Angular_Resolution_ +' ERROR: '+ lastCAN_ErrMsg;
                LogStep(342,342, 342,  lastCAN_ErrMsg );
                exit;
            end;
            // The Device Resolution [1,10,{100},1000] was Written successfully ! (NB: valori invertiti rispetto a ciò che va in subidx!)
            DUT_resolution   := 100;
            edtResolDUT.text := DUT_resolution.ToString;
            LogStep(343,343, 343, 'OK new DUT '+ _Angular_Resolution_ +' STORED Successfully to NodeID:'+ edtDutAddr.text );
            LogStep(344,344, 344, #9#9'DUT '+ _Angular_Resolution_ +' = '+ DUT_resolution.ToString );

            // NB: CAN_NodeSaveAll in eeprom permette di non perdere questo setup Risoluzione del DUT
            //     anche se poi avvengono dei Reset Node !!!
            //     che di solito annullano i setup temporanei (es. risoluzione in centesimi) ricaricando i default !
            if CAN_NodeSaveAll( edtDUTaddr.text ) then begin
                // Ok, niente msec. attesa perchè WR in eeprom già completato.
                LogStep(345,345, 345, 'CAN SaveParams command SENT Successfully to DUT:'+ edtDUTaddr.text );
            end
            else begin
                // l' Errore è bloccante, quindi esco.
                lastCAN_ErrMsg := 'CAN SaveParams ERROR on DUT:'+ edtDUTaddr.text +' - ERROR: '+ lastCAN_ErrMsg;
                LogStep(346,346, 346,  lastCAN_ErrMsg );
                exit;
            end;
            // Non serve anche NodeReset perchè parametro è già in Ram !?
            // e comunque non cambierebbe la risoluzione dato che è salvata !

            // Reset DUT per rendere operativi i parametri salvati in eeprom.
            if not CAN_SendNMT( CHdut, cmdResetNode, edtDUTaddr.text ) then begin // Reset solo DUT perchè i REF non vanno modificati, solo verificati.
          //if not CAN_SendNMT( cmdResetNode, '00h' ) then begin           // Reset ALL nodes!
                // l' Errore è bloccante, quindi esco.
                lastCAN_ErrMsg := 'DUT ResetNode ERROR: '+ lastCAN_ErrMsg;
                LogStep(347,347, 347,  lastCAN_ErrMsg );
                exit;
            end;
            // The Reset was successfully sent
            LogStep(348,348, 348, 'CAN ResetNode SENT Successfully to NodeID:'+ edtDUTaddr.text );
            // DUT reset & boot completed !
            if not CAN_FlushRxBuffer(CHdut) then begin
                lastCAN_ErrMsg := 'ERROR on Flush CAN buffer: '+ lastCAN_ErrMsg;
                LogStep(349,349, 349,  lastCAN_ErrMsg );
                exit;
            end;
            LogStep(350,350, 350, 'OK FLUSH DUT Can queue Successful.');
            // lascio Rx Can queue clean.
        end;
        result := True;
    except
      on E: Exception do begin
          LogStep(351,351, 351,  'EXCEPTION on Force Resolution to [degCentesimi] >> '+ E.ClassName +#10#13#9+ ' description: '+E.Message);
      end;
    end;
end;


function Toiac3Fasi2ax.Setup4_CD(var JobParams: TJobParams; testAxis:char): boolean;
var
  Can: TCanFrame;
  sourceFile: string;
  i: integer;
begin
{   come setup di fase C e D serve controllare status PCAN, presenza del Reference e DUT devices,
    caricare i valori di Caratterizzazione del banco inclinometri e massimo errore accettabile.
    Configurare DUT per invio *continuo* in gradi settando EventTimer cyclical transmission a 250ms
    mentre Start/Stop viene fatto invece nei singoli step.
}
    result := False;
    JobParams.retCode := resUNDEFINED;
    JobParams.retMsg := _Incompleto_;
    tmrCanRead.enabled := False;          // per sicurezza

    gaugeXRange.ValueStart := 0;          // Hide marker dei target degs
    gaugeXRange.ValueEnd   := 0;
    gaugeYRange.ValueStart := 0;
    gaugeYRange.ValueEnd   := 0;

    if not CAN_GetStatus(CHref) then begin
        JobParams.retMsg := 'PCAN Ref Status ERROR: '+ lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        if lastCAN_Status = PCAN_ERROR_BUSHEAVY then begin
            JobParams.retMsg := JobParams.retMsg + sLineBreak + 'Possibile mancanza di alimentazione o BaudRate errato !';
            // Necessario Close CAN !
            CAN_Release(CHref);           // necessario Release+Init, perchè CAN_Reset() non risolve il lock, ma forse
            CAN_Init(CHref);              // il reinit si...
            // Qui niente update lastCAN_ErrMsg perchè confonderebbe; va lasciato solo l'esito di CAN_GetStatus().
            // così il prossimo retry dovrebbe riuscire.
        end;
        LogStep(352,352, 352, JobParams.retMsg);
        exit;
    end;
    if CHdut.Handle <> CHref.Handle then
        if not CAN_GetStatus(CHdut) then begin
            JobParams.retMsg := 'PCAN Dut Status ERROR: '+ lastCAN_ErrMsg;
            JobParams.retCode := resERROR;
            if lastCAN_Status = PCAN_ERROR_BUSHEAVY then begin
                JobParams.retMsg := JobParams.retMsg + sLineBreak + 'Possibile mancanza di alimentazione o BaudRate errato !';
                // Necessario Close CAN !
                CAN_Release(CHdut);       // necessario Release+Init, perchè CAN_Reset() non risolve il lock, ma forse
                CAN_Init(CHdut);          // il reinit si...
                // Qui niente update lastCAN_ErrMsg perchè confonderebbe; va lasciato solo l'esito di CAN_GetStatus().
                // così il prossimo retry dovrebbe riuscire.
            end;
            LogStep(353,353, 353, JobParams.retMsg);
            exit;
        end;
  	// The channels are Ready
    LogStep(354,354, 354, 'Fase_C&D: STATUS PCAN_OK');

    // prepara DUT e REF2 address per usi vari.
    vRef2Addr := strCanValue2integer(edtRef2Addr.text);
    vDutAddr := strCanValue2integer(edtDutAddr.text);

    edtRefX.color := clBtnface;
    edtRefY.color := clBtnface;

    // Prima meglio un Reset del Ref2 e Dut, casomai fossero già in Start mode.
  //if not CAN_SendNMT( cmdResetNode, '00h' ) then begin  // Reset ALL nodes!
    if not CAN_SendNMT( CHref, cmdResetNode, edtRef2Addr.text ) then begin
        JobParams.retMsg := 'REF2 ResetNode '+edtRef2Addr.text+' ERROR: '+ lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        LogStep(355,355, 355, lastCAN_ErrMsg);
        exit;
    end;
    LogStep(356,356, 356, 'REF2 ResetNode SENT Successfully to NodeID:'+ edtRef2Addr.text );
    if not CAN_SendNMT( CHdut, cmdResetNode, edtDutAddr.text ) then begin
        JobParams.retMsg := 'DUT ResetNode '+edtDutAddr.text+' ERROR: '+ lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        LogStep(357,357, 357, lastCAN_ErrMsg);
        exit;
    end;
    LogStep(358,358, 358, 'ALL Nodes Reset Successful');
    LogStep(359,359, 359, 'FLUSH queue All Can CHANNELS...');
    if not CAN_FlushRxBuffer(CHref) then begin
        LogStep(360,360, 360,  'ERROR on CHref Flush: '+lastCAN_ErrMsg );
        exit;
    end;
    if CHdut.Handle <> CHref.Handle then
        if not CAN_FlushRxBuffer(CHdut) then begin
            LogStep(361,361, 361,  'ERROR on CHdut Flush: '+lastCAN_ErrMsg );
            exit;
        end;
    LogStep(362,362, 362, 'OK FLUSH All CanCH Successful.');    // preparato rx queue clean.

    // Verifica presenza device REF2: 0=none 1=1asse  2=2assi
    i := CAN_CheckREF(edtRef2Addr.text);     // per REF serve comando specifico, diverso da ID node.
    if i <> 2 then begin
        JobParams.retMsg := 'ERROR Unexpected REF2 device Type.';
        JobParams.retCode := resERROR;
        LogStep(363,363, 363, JobParams.retMsg);
        exit;
    end;
    // qui sarebbe da valutare se il tipo Ref trovato a questo indirizzo è corretto per i successivi Test previsti !!
    LogStep(364,364, 364, 'OK REF type '+ i.toString+'Axes present on Address: '+ edtRef2Addr.text); //PCAN_ERROR_OK.

    // Verifica presenza DUT.
    if not CAN_CheckNodeID(edtDutAddr.text) then begin
        JobParams.retMsg := 'DUT check NodeID ERROR: '+ lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        LogStep(365,365, 365, JobParams.retMsg);
        exit;
    end;
    LogStep(366,366, 366, 'OK DUT present on Address: '+ edtDutAddr.text); //PCAN_ERROR_OK.

    // Assicura che REF e DUT lavorino a uguale risoluzione in Centesimi di grado !
    // in precedenza qui forzavo la risoluzione a degCentesimi, ma ora la setto già in Fase_B !
    // come prerequisito prima di avviare la calibrazione e poi anche successivi test !
    // Ora dovrebbe risultare già a degCentesimi quindi solo una verifica.
    if not Force_Resolution4EOL() then begin
        JobParams.retMsg := lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        exit;
    end;

    // per START REF2 è necessario cyclical transmission già configurato su REF !
    Can.Command  := _Timer_for_cyclical_txd_;      // = Event timer for cyclical transmission [multiple of 1ms, 0 = disabled]
    Can.Data     := kTimer4cyclicalTxd;            // 250 msec.
    Can.NodeID   := edtRef2Addr.text;              // str ID del device
    if not CAN_WriteParam( Can ) then begin
        JobParams.retMsg := 'REF2 WrParam '+Can.Command+' ERROR: '+ lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        LogStep(367,367, 367, JobParams.retMsg);
        exit;
    end;
    // The parameter was Written successfully.
    LogStep(368,368, 368, 'OK Ref2 '+ Can.Command +' STORED Successfully to NodeID:'+ edtRef2Addr.text );
    LogStep(369,369, 369, #9#9'Ref2 '+ Can.Command +' = '+ kTimer4cyclicalTxd +' msec');

    // anche per START DUT è necessario cyclical transmission già configurato su DUT !
    Can.Command  := _Timer_for_cyclical_txd_;      // = Event timer for cyclical transmission [multiple of 1ms, 0 = disabled]
    Can.Data     := kTimer4cyclicalTxd;            // 250 msec.
    Can.NodeID   := edtDUTaddr.text;               // str ID del device
    if not CAN_WriteParam( Can ) then begin
        JobParams.retMsg := 'DUT WrParam '+Can.Command+' ERROR: '+ lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        LogStep(370,370, 370, lastCAN_ErrMsg);
        exit;
    end;
    // The parameter was Written successfully.
    LogStep(371,371, 371, 'OK DUT '+ Can.Command +' STORED Successfully to NodeID:'+ edtDUTaddr.text );
    LogStep(372,372, 372, #9#9'DUT '+ Can.Command +' = '+ kTimer4cyclicalTxd +' msec');

    // recupero valori di Errore Massimo accettabile, già caricati da XLS sheet 'OI20LIST'
    // per Asse in Test e asse Cross
    if (DutList[currDutIdx].ErrMaxTest = 0) or (DutList[currDutIdx].ErrMaxCross = 0) then begin
        // se non trovati, impossibile procedere con tests.
        JobParams.retMsg := 'ErrorMAX non trovati per device '+ DutList[currDutIdx].CodArticolo;
        JobParams.retCode := resERROR;
        LogStep(373,373, 373, JobParams.retMsg);
        exit;
    end;
    LogStep(374,374, 374, 'Valori MAX Error per device: '+ DutList[currDutIdx].CodArticolo );
    edtErrMaxTest.text  := format('%6.2f°', [DutList[currDutIdx].ErrMaxTest]);
    edtErrMaxCross.text := format('%6.2f°', [DutList[currDutIdx].ErrMaxCross]);

    LogStep(375,375, 375,  #9#9'ErrMaxTest = '+  edtErrMaxTest.text );
    LogStep(376,376, 376,  #9#9'ErrMaxCross = '+ edtErrMaxCross.text );

    // caricamento listview lviewCarattBanco
    sourceFile := kAppPath + dmBase.vSessionSource;
    LogStep(377,377, 377, 'Loading Coefficienti di Correzione from : '+ sourceFile);
    if not LoadCarattBanco_fromXLS( sourceFile, testAxis ) then begin
        JobParams.retMsg := 'ERROR on loading Coeff.Correction list: '+ sourceFile;
        JobParams.retCode := resERROR;
        LogStep(378,378, 378, JobParams.retMsg);
        exit;
    end;
    LogStep(379,379, 379, 'OK, Coefficienti di Correzione caricati.');

    // Essendo un setup comune per C e D non posso sapere di quale Fase posso fare il Clear,
    // ma posso evitare .Clear perchè le .Add al Dictionary sovrascrivono sempre la stessa Key (TTAngle).
    //TestXRecords.Clear;          //
    //TestYRecords.Clear;          // reset sortDictionary con records dati dei test, No!

    LogStep(380,380, 380, 'OK, completato Setup per FaseCD-'+testAxis );

    //dmBase.wait(500);
    JobParams.retMsg := _Completato_;
    JobParams.retCode := resOK;
    result := True;
end;

function Toiac3Fasi2ax.SetupFase_E(var JobParams: TJobParams): boolean;
begin
{   come setup di fase E serve solo un reset device per preparare R/W parameters
}
  //mainForm.mmLog.lines.add('run SetupFase_E...');
    result := False;
    JobParams.retCode := resUNDEFINED;
    JobParams.retMsg := _Incompleto_;
    tmrCanRead.enabled := False;            // per sicurezza

    // Definisce tipo di parametri da caricare in listview.
    cbxListaDefaultsCom.Checked := True;    //
    cbxListaDefaults1ax.checked := False;   //
    cbxListaDefaults2ax.checked := True;    //
    cbxListaSN.Checked       := True;       //
    cbxListaCalib1ax.Checked := False;      // qui parametri DEF e CFG2.
    cbxListaCalib2ax.Checked := True;
    if idxModuloSw = numModuliSw then           // Solo se questo è l'ultimo modulo tra i ModuliSw previsti
        cbxListaDefaultsEND.Checked := True     // abilito il caricamento dei defaults di Fine EOL !
    else
        cbxListaDefaultsEND.Checked := False;

    edtToWrite.text := '';                  // Clean
    edtToWrite.Color := clBtnFace;          //

    // come prerequisito per tutti gli step della fase E resetto tutti i nodes !
    // questo stoppa  anche eventuali cyclical transmission dai REF.
    if not CAN_SendNMT( CHref, cmdResetNode, edtRef2Addr.text ) then begin
        JobParams.retMsg := 'REF2 ResetNode '+edtRef2Addr.text+' ERROR: '+ lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        LogStep(381,381, 381, JobParams.retMsg);
        exit;
    end;
    // The Reset was successfully sent
    if not CAN_SendNMT( CHdut, cmdResetNode, edtDutAddr.text ) then begin
        JobParams.retMsg := 'DUT ResetNode '+edtDutAddr.text+' ERROR: '+ lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        LogStep(382,382, 382, JobParams.retMsg);
        exit;
    end;
    LogStep(383,383, 383, 'ALL Nodes Reset Successfully');
    // clean boot frames.
    dmBase.wait(300);       // anche solo per evitare click multipli involontari da parte user.
    LogStep(384,384, 384, 'FLUSH queue All Can CHANNELS...');
    if not CAN_FlushRxBuffer(CHref) then begin
        LogStep(385,385, 385,  lastCAN_ErrMsg );
        exit;
    end;
    if CHdut.Handle <> CHref.Handle then
        if not CAN_FlushRxBuffer(CHdut) then begin
            LogStep(386,386, 386,  lastCAN_ErrMsg );
            exit;
        end;
    LogStep(387,387, 387, 'OK FLUSH All CanCH Successful.');    // preparato rx queue clean.

    LogStep(388,388, 388, 'OK, completato Setup per Fase E' );
    JobParams.retMsg := _Completato_;
    JobParams.retCode := resOK;
    result := True;
end;

function Toiac3Fasi2ax.SetupFase_F(var JobParams: TJobParams): boolean;
var
  err: string;
  res: boolean;
begin
{
    come setup di fase F serve controllare la connessione al DB
}
    result := False;
    JobParams.retCode := resUNDEFINED;
    JobParams.retMsg := _Incompleto_;

    // Connessione al DB PostgreSQL 12.3
    LogStep(389,389, 389, 'Attempting connection to Optoi POSTGRES DB on '+ dmBase.pgConn.Server);
    Screen.Cursor := crHourglass;
    try
        try
            err := _Undefined_;
            dmBase.pgConn.Connected := False;
            (*
            pgConn.Server :=;
            pgConn.Port :=;
            pgConn.Database :=;
            pgConn.Username :=;
            pgConn.Password :=;
            *)
            dmBase.pgConn.connected := True;
        except
           on E: Exception do begin    // on EOverflow do ...  se trappo specifica exception.
                 // qui arriva dopo aver rilevato l'Exception nel pgConn Error.
                 err := E.ClassName +' - message: '+E.Message;
                 LogSysEvent(svEXCEPTION, 2300, 'DB Connection open '+#10#13+ err);
                 JobParams.retCode := resERROR;
                 JobParams.retMsg := 'DB Connection Error: '+ err;
                 exit;
           end;
              // niente else raise, perchè con la base ClassName E:Exception non esegue altro !!!  e senza dare errori...!
              // Else funziona solo se uso precisi ClassName, tipo E:EDatabaseError, EDivByZero, EHeapException, ...
              // vedi  http://www.delphibasics.co.uk/Article.asp?Name=Exceptions
        end;
    finally
        Screen.Cursor := crDefault;
    end;
    if dmBase.pgConn.Connected then begin
        LogSysEvent(svINFO, 2310, 'DB Connection open Successful.');
        LogStep(390,390, 390, #9'Connesso al Server '+ dmBase.pgConn.Server +':'+ dmBase.pgConn.Database);
        LogStep(391,391, 391, #9'Versione: '+ dmBase.pgConn.ServerVersionFull);  // o ServerVersion
    end
    else begin
        LogSysEvent(svERROR, 2320, 'DB Connection open FAILED !');
        JobParams.retMsg := 'ERROR on Connecting DB > '+ err;
        JobParams.retCode := resERROR;
        LogStep(392,392, 392, JobParams.retMsg);
        exit;
    end;

    // Apertura DatSets
    res := True;
    try
       dmBase.tabSessionDut.Active := True;
       LogStep(393,393, 393, 'Opened table: '+dmBase.tabSessionDut.TableName);
    except
       LogSysEvent(svEXCEPTION, 2330, 'Unable to Open Table: '+dmBase.tabSessionDut.TableName);
       res := False;
    end;
    try
       dmBase.tabTestData.Active := True;
       LogStep(394,394, 394, 'Opened table: '+dmBase.tabTestData.TableName);
    except
       LogSysEvent(svEXCEPTION, 2340, 'Unable to Open Table: '+dmBase.tabTestData.TableName);
       res := False;
    end;
    try
       dmBase.tabDutParams.Active := True;
       LogStep(395,395, 395, 'Opened table: '+dmBase.tabDutParams.TableName);
    except
       LogSysEvent(svEXCEPTION, 2342, 'Unable to Open Table: '+dmBase.tabDutParams.TableName);
       res := False;
    end;
    // Check open datasets Results
    if res then begin
        LogStep(396,396, 396, 'DB tables schema Produzione READY.');
    end
    else begin
        err := 'DB open Dataset ERROR !';
        LogSysEvent(svERROR, 2350, err);
        JobParams.retMsg := err;
        JobParams.retCode := resERROR;
        exit;
    end;

    LogStep(397,397, 397, 'OK, completato Setup per Fase F.' );
    dmBase.wait(1000);
    JobParams.retMsg := _Completato_;
    JobParams.retCode := resOK;
    result := True;
end;

function Toiac3Fasi2ax.SetupFase_G(var JobParams: TJobParams): boolean; // fine Procedure
var
  res: boolean;
  aDataset: string;
  err: string;
begin
    {
     Fase senza Steps che serve solo per attività di chiusura sessione.
     Se non va a buon fine, user può forzare pass.
    }
    result := False;
    JobParams.retCode := resUNDEFINED;
    JobParams.retMsg := _Incompleto_;

    if not CAN_Release( CHref ) then begin
        JobParams.retMsg := 'ERROR on CHref Release: '+ lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
      //exit; // comunque sia, eventuale errore non è bloccante per questa fase di chiusura !
    end;
    if CHdut.Handle <> CHref.Handle then
        if not CAN_Release( CHdut ) then begin
            JobParams.retMsg := 'ERROR on CHdut Release: '+ lastCAN_ErrMsg;
            JobParams.retCode := resERROR;
          //exit; // comunque sia, eventuale errore non è bloccante per questa fase di chiusura !
        end;
    FinalizeProtection();

    // ora aggiornamento tempi preliminari con definitivi e salvataggio su DB.
    SessionUpdateDB();   // per ora non bloccante.

    res := True;
    // Chiusura DataSets
    try
       dmBase.tabSessionDut.Active := False;
    except
       LogSysEvent(svEXCEPTION, 2600, 'on Close Table: '+dmBase.tabSessionDut.TableName);
       res := False;
    end;
    try
       dmBase.tabTestData.Active := False;
    except
       LogSysEvent(svEXCEPTION, 2610, 'on Close Table: '+dmBase.tabTestData.TableName);
       res := False;
    end;
    try
       dmBase.tabDutParams.Active := False;
    except
       LogSysEvent(svEXCEPTION, 2615, 'on Close Table: '+dmBase.tabDutParams.TableName);
       res := False;
    end;
    try
       {
       aDataset := qryProdotti.name;
       qrySelectLast.Close;
       aDataset := qryCommesse.name;
       qryCheck_Stornate.Close;
       }
       dmBase.pgConn.Connected := False;
    except
       LogSysEvent(svERROR, 2620, 'on Close Query: '+ aDataset);
       res := False;
    end;

    // Check close Results
    if res then begin
        LogSysEvent(svINFO, 2630, 'Production Datasets Closed.');
        end
    else begin
        err := 'Production Datasets Close ERROR !';
        LogSysEvent(svERROR, 2640, err );
        JobParams.retMsg := err;
        JobParams.retCode := resERROR;
      //exit; // comunque sia, eventuale errore non è bloccante per questa fase di chiusura !
    end;

    LogStep(398,398, 398, 'OK, completata ultima Fase G' );
    dmBase.wait(700);
    imgFineProc.visible := True;
    JobParams.retMsg := _Completato_;
    JobParams.retCode := resOK;
    result := True; // comunque sia, eventuale errore non è bloccante per questa fase di chiusura !
end;



////#######################  Steps Fase A  #######

function Toiac3Fasi2ax.RunStep_A1(var JobParams: TJobParams): boolean;
var
  cfgFile: string;
begin
    // Caricamento TList con Parametri CAN per Configurazione DUT.
    // Popola la lviewCFG con in Comandi CAN di default già acquisiti e presenti nella lviewCanCMDs !
    // In futuro sarà possibile acquisire tali Defaults e Cfg da file dedicati per modello/cliente
    // che saranno nel formato uguale al master PCAN_CMDs.xlsx  ma con valori dedicati a modelli e/o clienti specifici.

    result := False;
    JobParams.retCode := resUNDEFINED;
    JobParams.retMsg := _Incompleto_;
    dmBase.wait(1000);// anche solo per evitare click multipli involontari da parte user.
  //mainForm.mmLog.lines.add('RunStep_1...');

    edtToWrite.text := '';          // Clean
    edtToWrite.Color := clBtnFace;  //

    // Carica la lviewCFG da file Excel di configurazione selezionato
    // e solo i Comandi previsti di Default per configurare il DUT specifico modello/cliente.
    cfgFile := dmBase.vCAN_DefaultCFG;
    LogStep(399,399, 399, 'Loading CFG file '+ cfgFile);
    if not Load_CFG_CmdSet_fromXLS( cfgFile ) then begin
        JobParams.retMsg := 'Step1 ERROR on Load CFG Command list: '+ cfgFile;
        JobParams.retCode := resERROR;
        LogStep(400,400, 400, JobParams.retMsg);
        exit;
    end;
    LogStep(401,401, 401, 'OK CFG Command list loaded.');
    JobParams.retMsg := _Completato_;
    JobParams.retCode := resOK;
    result := True;
end;

function Toiac3Fasi2ax.RunStep_A2(var JobParams: TJobParams): boolean;
var
  i, tInt: integer;
  ListItem: TListItem;
  cmdAlias, tStr: string;
  WrValue: string;
  Can: TCanFrame;
  errCount, ToWrite: integer;
  CanCmd: TCanCmd;
begin
  //mainForm.mmLog.lines.add('RunStep_2...');
    LogText('   *** Lettura dati Configurazione DUT via CAN ***');
    result := False;
    errCount := 0;
    JobParams.retCode := resUNDEFINED;
    JobParams.retMsg := _Incompleto_;

    // prima di tutto assicura CAN channels "clean"
    if not CAN_CleanCH(CHref) then begin
    	// An error occurred.  We show the error.
        JobParams.retMsg := 'CAN_Clean CHref ERROR: '+ lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        inc(errCount);
        LogStep(402,402, 402,  lastCAN_ErrMsg );
        exit;
    end;
    if CHdut.Handle <> CHref.Handle then
        if not CAN_CleanCH(CHdut) then begin
            // An error occurred.  We show the error.
            JobParams.retMsg := 'CAN_Clean CHdut ERROR: '+ lastCAN_ErrMsg;
            JobParams.retCode := resERROR;
            inc(errCount);
            LogStep(403,403, 403,  lastCAN_ErrMsg );
            exit;
        end;
    LogStep(404,404, 404, JobList.Strings[currJob]+': STATUS PCAN_OK');
    // Channels are Ready to work, ma meglio stoppare eventuale cyclical transmission dal REF !

    // Meglio stoppare anche eventuali cyclical transmission dai REF,
    // ma un Reset ALL nodes è già fatto in SetupFase_A()
    // mentre lo Stop non è come in idle e il REF non risponde a CMDs !
    (*
    if not CAN_SendNMT( cmdStopNode, edtRef2Addr.text ) then begin
        JobParams.retMsg := 'ERROR on StopNode REF2: '+ lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        exit;
    end;
    LogStep(405,405, 405, 'OK REF2 node Stopped.'); //PCAN_ERROR_OK.
    *)

    // Sblocca area protetta 5555h del DUT per precauzione, se dovesse servire Write (again) params di Calibrazione,
    // ma serve anche solo per Leggere il S/N in locazione 5555 !
    if not CAN_UnlockParams( edtDUTaddr.text ) then begin
        // l' Errore è bloccante, quindi esco.
        JobParams.retMsg := 'DUT UnlockParams ERROR: '+ lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        inc(errCount);
        LogStep(406,406, 406,  JobParams.retMsg );
        exit(False);       // unlock non riuscito, quindi non serve lock finally !
    end;
    LogStep(407,407, 407, 'OK, DUT Unlock command CONFIRMED by NodeID:'+ edtDUTaddr.text );

    cbxSelectAll.Checked := True;
    cbxSelectAllClick(self);         // accerta di leggere sempre tutti i parametri.

    // Loop sul file CFG pe lettura parametri Default.
    ToWrite  := 0;                   // conta parametri che risultano diversi dal previsto default.
    edtToWrite.text := '';
    edtToWrite.Color := clBtnFace;
    try
        try
          for i := 0 to lviewCFG.Items.Count -1 do begin   // Scorre Parametri CAN
              // check per Cancel ad ogni giro.
              application.processmessages;
              if mainForm.btnAbort.tag > 0 then begin
                  LogText('*** ABORT by USER, lettura device Interrotta !'+#13#10 );
                  exit;
              end;
              dmBase.Wait(10); // msec. attesa fra ogni param.
              // clear Rx buffer ?  visto il reset a inizio loop, per ora non sembra necessario.

              // get param e setup frame da tx
              if lviewCFG.Items[i].Checked  then begin          // processo solo i checked, anche se ho travasato in lviewCFG solo i default,
                  ListItem:= lviewCFG.Items.item[i];            // perchè potrei aver modificato a mano.

                  // acquisisco Alias che è la Key di gestione frame comandi.
                  cmdAlias := ListItem.SubItems[7];
                  // con Alias mi procuro sempre il record completo dal dictionary TCanCmd,
                  // che uso sempre per composizione comandi, perchè da lviewCFG prendo solo dati variabili !
                  if not dicCanCmds.TryGetValue( cmdAlias, CanCmd) then begin
                      JobParams.retMsg := 'ERROR Alias NOT Found "'+ cmdAlias +'" in dictionary CanCmds.';
                      JobParams.retCode := resERROR;
                      inc(errCount);
                      LogStep(408,408, 408,  JobParams.retMsg );
                      exit;
                  end;

                  // Compone la frame da inviare per Lettura parametro.
                  Can.Index    := CanCmd.Index;                 // Index
                  Can.Subidx   := CanCmd.SubIdx;                // SubIdx
                  Can.Datatype := CanCmd.DataType;              // int8, uns16, ...
                  // e parti variabili
                  Can.NodeID   := edtDUTaddr.text;              // ID del device (str Hex o Dec) và rimesso perchè ogni CAN_SendFrame() lo modifica !
                  Can.RTR      := False;                        // fisso
                  Can.Data := '';                               // è una Read, quindi niente payload, ne Len payaload !
                  // Tutte le routine PCAN presuppongono parametri già in HEX (se c'è 'h' finale)
                  // quindi verifico che lo siano, altrimenti li converto in hex prima dell'uso!
                  // Sono ammessi blank fra bytes e h finale, perchè verranno rimossi dalla CAN_SendFrame.
                  // Can.Len viene calcolata dalla funzione CAN_Write, con default Data Length $40 = AnySize (=8)
                  // invio can frame
                  if CAN_SendFrame(Can) then begin
                      // The frame was successfully sent
                      // la funzione Write accetta Can.Data con spazi
                      LogStep(409,409, 409, 'CAN Frame SENT Successfully,  ID:'+Can.CobID +'   LenPL:'+ intToStr(Can.LenPayLoad) +'   Frame:'+ hexLarge(Can.Frame));
                  end
                  else begin
                      // singolo Errore non è bloccante, lascio portare a termine loop su tutti i params.
                      LogStep(410,410, 410,  lastCAN_ErrMsg );
                      inc(errCount);
                      continue;
                  end;

                  // attesa sincrona di answer da DUT.
                  dmBase.Wait(100); // msec. attesa prima di read.

                  // controlla response ok/err   (counter Errors per response finale)
                  if CAN_ReceiveFrame( Can ) then begin
                      // OK, the frame was successfully read.
                      // Received Payload in Can.Data
                      LogStep(411,411, 411, 'CAN Frame RECEIVED Successfully,  ID:'+Can.CobID +'   LenPL:'+intToStr(Can.LenPayLoad) +'   Frame:'+hexLarge(Can.Frame));

                      // qui verifica valore parametro rilevato con quello (Wr) previsto !
                      WrValue := ListItem.SubItems[2];                // prende Write value dalla listView
                      WrValue := ReplaceText(WrValue, ' ', '');       // prima sempre remove all blanks.
                      // Può essere in forma DEC o HEX con 'h' finale, ma la CAN_SendFrame() si arrangia a gestire i casi !
                      // Può subire override se sono previsti format specifici per parametro o gruppo !
                      // Avrei anche il valore salvato in CanCmd.WrValue ma meglio usare quello già visualizzato in lviewCFG !

                      // controlla se parametro previsto è espresso in HEX con 'h' finale
                      if AnsiContainsText(WrValue, 'h') then begin    // case insensitive 'h' or 'H'
                          // blanks sono già tolti, quindi elimino anche simboli hex per rendere
                          // direttamente confrontabile con parametro ricevuto via CAN.
                          WrValue := ReplaceText(WrValue, 'h', '');   // remove 'h' or 'H'
                      end
                      else begin
                          // è valore espresso in DECIMALE as string, quindi converto in integer
                          // per poi tornare a hex string e poterlo confrontare con Can.Data che è sempre hex string !
                          // blanks sono già stati tolti.
                          // Anche se valore decimale è NEGATIVO non serve usare complemento a 2.
                          WrValue := intToHex( StrToIntDef( WrValue, -1));  // -1 perchè 0 diventa uguale a locazione vuota e non dà errore!
                          WrValue := RightStr( WrValue, Length(Can.Data));  // cast a lunghezza prevista.
                      end;

                      // gestisco formato specifico che usa EXTEND.
                      if CanCmd.Extend <> '' then begin
                          // ha formato particolare by WT !  quindi trimmo prefisso.
                          WrValue := rightStr(WrValue, 4);  // e mantengo solo 16 bit.
                          Can.Data := rightStr(Can.Data, 4);  // e mantengo solo 16 bit.
                      end;

                      // ora posso fare confronto diretto dei valori stringa in forma Hex.
                      if WrValue.TrimLeft(['0']) <> Can.Data.TrimLeft(['0']) then begin // elimina eventuali 0 iniziali !?
                          // leave checked
                          // set icon Warning
                          ListItem.SubItemImages[4] := 2;            // Status Warning su colonna corrente.
                        //ListItem.ImageIndex := 2;                  // Status Warning su prima colonna.
                          inc(ToWrite);    // conta parametri che risultano diversi dal previsto default, ovvero to write in DUT.
                      end
                      else begin
                          // Uncheck
                          lviewCFG.Items[i].Checked := False;
                          // set icon Ok
                          ListItem.SubItemImages[4] := 6;            // Status OK su colonna corrente.
                        //ListItem.ImageIndex := 6;                  // Status OK su prima colonna.
                      end;

                      // in ogni caso visualizzo valore hex Ricevuto dal DUT.
                      ListItem.SubItems[4] := Can.Data + 'h';
                      // Payload ricevuto è sempre in Hex e quindi va evidenziato.
                      // Dove serve in decimale, sarà compito GUI convertire o trimmare hex leading zeroes !
                      // es:   intVal := StrToInt('$' + '0F');   // intVal = 15
                      // si può anche usare hexLarge() per separare e rendere più leggibili i singoli bytes.
                  end
                  else begin
                      // singolo Errore non è bloccante, lascio portare a termine loop su tutti i params.
                      LogStep(412,412, 412,  lastCAN_ErrMsg );
                      inc(errCount);
                      ListItem.SubItemImages[4] := 17;            // Status Errore su colonna corrente.
                      continue;
                  end;
              end;//if checked
          end;//for lviewCFG.Items

          // qui devo bloccare in caso di errore, perchè indica problemi di lettura parametri !
          if errCount > 0 then begin
              JobParams.retMsg := _Errore_;
              JobParams.retCode := resERROR;    // problemi di WR config è di tipo bloccante.
              result := False;
          end
          else begin
              if ToWrite = 0 then begin
                  // se la CFG richiesta è già memorizzata, posso predisporre SKIP di prossimi 2 Step (jobs) !
                  jobsToSkip := 2;
                  // salta prossimi 2 jobs, di Write cfg e Save cfg.
                  // NB: mai saltare ULTIMO step della Sessione !  che serve per closeup.
              end;
              JobParams.retMsg := _Completato_; // Ok, anche se trovati parametri da aggiornare,
              JobParams.retCode := resOK;       // perchè è attività a cura degli step successivi!
              result := True;
          end;

        except
          on E: Exception do begin
              JobParams.retCode := resERROR;
              tStr := 'in '+JobList.Strings[currJob]+' on CAN Parameters iteration >> '+ E.ClassName +#10#13#9+ ' description: '+E.Message;
              //LogStep(413,413, 413,  tStr );
              LogSysEvent(svEXCEPTION, 200, tStr);
          end;
        end;
    finally
        // Blocca subito l'area protetta 5555h del DUT, perchè il Reset node NON dà lo stesso risultato.
        if not CAN_LockParams( edtDUTaddr.text ) then begin
            // l' Errore è bloccante, quindi esco.
            JobParams.retMsg := 'DUT LockParams ERROR: '+ lastCAN_ErrMsg;
            JobParams.retCode := resERROR;
            inc(errCount);
            LogStep(414,414, 414, 'ERROR on LOCK params for DUT:'+ edtDUTaddr.text );
            LogStep(415,415, 415,  lastCAN_ErrMsg );
            result := False;
        end
        else
            LogStep(416,416, 416, 'OK, DUT Lock command CONFIRMED by NodeID:'+ edtDUTaddr.text );

        LogStep(417,417, 417, JobList.Strings[currJob]+' ERRORS = '+ intToStr(errCount));
        if ToWrite > 0 then
            edtToWrite.Color := clyellow   // = necessario alcuni overwrite con defaults.
        else
            edtToWrite.Color := clLime;    // = tutti OK.
        edtToWrite.text := intToStr(ToWrite);
        dmBase.wait(50);  // solo permettere application process messages e refresh GUI.
    end;
end;

function Toiac3Fasi2ax.RunStep_A3(var JobParams: TJobParams): boolean;
var
  i, iTmp: integer;
  ListItem: TListItem;
  tData, cmdAlias: string;
  Can: TCanFrame;
  errCount, countWritten: integer;
  CanCmd: TCanCmd;
begin
  //mainForm.mmLog.lines.add('RunStep_3...');
    LogText('   *** Inizio scrittura dati CFG via CAN ***');
    result := False;
    JobParams.retCode := resUNDEFINED;
    JobParams.retMsg := _Incompleto_;
    errCount := 0;

    // prima di tutto assicura CAN channels "clean"
    if not CAN_CleanCH(CHref) then begin
    	// An error occurred.  We show the error.
        JobParams.retMsg := 'CAN_Clean CHref ERROR: '+ lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        inc(errCount);
        LogStep(418,418, 418,  lastCAN_ErrMsg );
        exit;
    end;
    if CHdut.Handle <> CHref.Handle then
        if not CAN_CleanCH(CHdut) then begin
            // An error occurred.  We show the error.
            JobParams.retMsg := 'CAN_Clean CHdut ERROR: '+ lastCAN_ErrMsg;
            JobParams.retCode := resERROR;
            inc(errCount);
            LogStep(419,419, 419,  lastCAN_ErrMsg );
            exit;
        end;
    LogStep(420,420, 420, JobList.Strings[currJob]+': STATUS PCAN_OK');
    // Channels are Ready to work, ma meglio stoppare eventuale cyclical transmission dal REF !

    // Sblocca area protetta 5555h del DUT necessario se dovesse servire Write (again) params di Calibrazione,
    // ma serve anche solo per Leggere il S/N in locazione 5555 !
    if not CAN_UnlockParams( edtDUTaddr.text ) then begin
        // l' Errore è bloccante, quindi esco.
        JobParams.retMsg := 'DUT UnlockParams ERROR: '+ lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        LogStep(421,421, 421,  JobParams.retMsg );
        exit(False);       // unlock non riuscito, quindi non serve lock finally !
    end;
    LogStep(422,422, 422, 'OK, DUT Unlock command CONFIRMED by NodeID:'+ edtDUTaddr.text );

  //cbxSelectAll.Checked := True;    // per ora Write solo dei parametri checked
  //cbxSelectAllClick(self);         // ovvero trovati diversi dal previsto.

    // Loop su lviewCFG per Scrittura parametri, ma solo i selezionati !
    try
        errCount := 0;
        countWritten := 0;                // conta parametri scritti con default value.
        try
          for i := 0 to lviewCFG.Items.Count -1 do begin   // Scorre Parametri CAN
              // check per Cancel ad ogni giro.
              application.processmessages;
              if mainForm.btnAbort.tag > 0 then begin
                  LogText('*** ABORT by USER, configurazione device Interrotta !'+#13#10 );
                  exit;
              end;
              dmBase.Wait(10); // msec. attesa fra ogni param.
              // clear Rx buffer ?  visto il reset a inizio loop, per ora non sembra necessario.

              // acquisisco riga del parametro corrente selezionato.
              if lviewCFG.Items[i].Checked  then begin           // Write solo su quelli trovati, e marcati, come diversi da default.
                  ListItem:= lviewCFG.Items.item[i];             // ma anche se marcati esplicitamente by user.

                  // acquisisco Alias che è la Key di gestione frame comandi.
                  cmdAlias := ListItem.SubItems[7];
                  // con Alias mi procuro sempre il record completo dal dictionary TCanCmd,
                  // che uso sempre per composizione comandi, perchè da lviewCFG prendo solo dati variabili !
                  if not dicCanCmds.TryGetValue( cmdAlias, CanCmd) then begin
                      JobParams.retMsg := 'ERROR Alias NOT Found "'+ cmdAlias +'" in dictionary CanCmds.';
                      JobParams.retCode := resERROR;
                      LogStep(423,423, 423,  JobParams.retMsg );
                      exit;
                  end;

                  // Compone la frame da inviare per Scrittura parametro.
                  Can.Index    := CanCmd.Index;                    // Index
                  Can.Subidx   := CanCmd.SubIdx;                   // SubIdx
                  Can.Datatype := CanCmd.DataType;                 // int8, uns16, ...
                  // e part ivariabili
                  Can.NodeID   := edtDUTaddr.text;                 // ID del device (str Hex o Dec) và rimesso perchè CAN_SendFrame() lo modifica !
                  Can.RTR      := False;                           // fisso
                  Can.Data     := ListItem.SubItems[2];            // valore della colonna WrValue da scrivere nel DUT !
                  // Può essere in forma DEC o HEX con 'h' finale, ma la CAN_SendFrame() si arrangia a gestire i casi !
                  // Può subire override se sono previsti format specifici per parametro o gruppo !
                  // Avrei anche il valore salvato in CanCmd.WrValue ma meglio usare quello già visualizzato in lviewCFG !
                  tData := Can.Data;
                  tData := ReplaceText(tData, ' ', '');       // prima sempre remove all blanks.
                  // Can.Data ovvero il WrValue ListItem.SubItems[2]
                  // Può essere in forma DEC o HEX con 'h' finale, ma la CAN_SendFrame() si arrangia a gestire i casi !
                  // Può subire override se sono previsti format specifici per parametro o gruppo (EXTEND) !
                  // Avrei anche il valore salvato in CanCmd.WrValue ma meglio usare quello già visualizzato in lviewCFG !

                  // Se parametro previsto è espresso in HEX con 'h' finale, non serve modificare e lo visualizzo così com'è.
                  // Ma se è DECimale (+/-) lo converto in HEX per uniformare al formato della colonna ReadValue !
                  if not AnsiContainsText(tData, 'h') then begin    // case insensitive 'h' or 'H'
                      // è valore espresso in DECIMALE as string, quindi converto in integer
                      // per poi tornare a hex string e visualizzarlo nel ReadValue se la Write va a buon fine !
                      // Anche se valore decimale è NEGATIVO non serve usare complemento a 2.
                      tData := intToHex( StrToIntDef( tData, -1));      // -1 perchè 0 diventa uguale a locazione vuota e non dà errore!
                  end;

                  // *** qui controlli specifici su singolo param o gruppo ***

                  // gestione specifica per parametro Unlock inhibit !
                  if cmdAlias = _Inhibit_between_TPDOs_ then   // Sblocco se inhibit
                      CAN_setTPDO1_comms(edtDUTaddr.text, modeDisabled);

                  // gestione specifica per parametro S/N !
                  if cmdAlias = _Serial_number_ then begin     // S/N va scritto in index dedicato, diverso da quello di read.
                //if ListItem.SubItems[8] = 'SNUM' then begin
                      // inoltre va prima sbloccato lo specifico registro per Write S/N !
                      if not CAN_UnlockSN( edtDUTaddr.text ) then begin
                          // l' Errore è bloccante, quindi esco.
                          JobParams.retMsg := 'DUT UnlockSN ERROR: '+ lastCAN_ErrMsg;
                          JobParams.retCode := resERROR;
                          LogStep(424,424, 424, 'ERROR on UnlockSN for DUT:'+ edtDUTaddr.text );
                          LogStep(425,425, 425,  JobParams.retMsg );
                          inc(errCount);
                          exit;           // unlock non riuscito, quindi non serve lock finally !
                      end;
                      // Non serve e non esiste un comando di LockSN, ma solo alla fine per intera area 5555h.
                      LogStep(426,426, 426, 'OK, DUT UnlockSN CONFIRMED by NodeID:'+ edtDUTaddr.text );
                      // Ora posso procedere a Wr del S/N, usando però opportuni Index/Subidx,
                      // che sono *diversi* da quelli di Read e quindi li sostituisco!
                      if dicCanCmds.TryGetValue( _SN_Value_2wr_, CanCmd) then begin
                          Can.Index  := CanCmd.Index;        //
                          Can.Subidx := CanCmd.Subidx;       // Index e SubIdx diversi per Wr S/N !
                          Can.Datatype := CanCmd.Datatype;   //
                      end
                      else begin
                          JobParams.retMsg := 'ERROR Alias NOT Found "'+ cmdAlias +'" in dictionary CanCmds.';
                          JobParams.retCode := resERROR;
                          LogStep(427,427, 427,  JobParams.retMsg );
                          exit;
                      end;
                  end;

                  // gestione specifica per parametri CFG1 e CFG2 che necessitano formato specifico che usa EXTEND !
                  // e in genere sono espressi in Decimale.
                  if CanCmd.Extend <> '' then begin
                      // Parametri Calibrazione 1 e 2 Assi vanno scritti con format dedicato, diverso da standard !
                      // In questo caso uso la colonna WriteValue (Can.Data) che è già stata valorizzata in fase LoadDefaults
                      // con il campo CanCmd.WrValue salvato in fase di calibrazione.
                      tData := rightStr(tData, 4)+'h';  // Cast a 16 bit (4+h) prima di modificare in formato Ext by WT.
                      // NB: per verifica dei valori di calibrazione (CFG1 o CFG2) Non posso usare
                      //     quelli salvati nel dictionary CalibRecords perchè la Key non è il cmdAlias !
                      iTmp := strToInt(Can.Data);       // formatta payload in modo specifico previsto da WT per i parametri di tipo CFG1,2 !
                      Can.Data := CanCmd.Extend + Ansichar(LongRec(iTmp).Bytes[1]) + Ansichar(LongRec(iTmp).Bytes[0]); // NB: usare Ansichar() e non char() o chr() che tronca a 7 bit !!!
                      // Datatype VSTR vengono scritte così come sono, niente swap bytes !
                      // Area 5555h è già sbloccata quindi posso procedere con write.
                  end;

                  // *** Fine controlli specifici su singolo param o gruppo ***

                  // Tutte le routine PCAN presuppongono parametri già in HEX (anche se manca $ davanti)
                  // quindi verifico che lo siano, altrimenti li converto in hex prima dell'uso!
                  // Sono ammessi blank fra bytes e h finale, perchè verranno rimossi dalla CAN_SendFrame.
                  // Can.Len viene calcolata dalla funzione CAN_Write, con default Data Length $40 = AnySize (=8)
                  // invio can frame
                  if CAN_SendFrame(Can) then begin
                      // The frame was successfully sent
                      // la funzione Write accetta Can.Data con spazi
                      LogStep(428,428, 428, 'CAN Parameter SENT Successfully,  ID:'+Can.CobID +'   LenPL:'+ intToStr(Can.LenPayLoad) +'   Frame:'+ hexLarge(Can.Frame));
                  end
                  else begin
                      // singolo Errore non è bloccante, lascio portare a termine loop su tutti i params.
                      LogStep(429,429, 429,  lastCAN_ErrMsg );
                      inc(errCount);
                      continue;
                  end;

                  // attesa sincrona di answer da DUT.
                  dmBase.Wait(100); // msec. attesa prima di read.

                  // controlla response ok/err   (counter Errors per response finale)
                  if CAN_ReceiveFrame( Can ) then begin
                      // OK, the frame was successfully read.
                      inc(countWritten);
                      // Received Payload in Can.Data
                      LogStep(430,430, 430, 'CAN Parameter STORED Successfully,  ID:'+Can.CobID +'   LenPL:'+intToStr(Can.LenPayLoad) +'   Frame:'+hexLarge(Can.Frame));
                      // qui non ricevo parametri da verificare, ma solo se Write Ok/Error !
                      lviewCFG.Items[i].Checked := False;          // Uncheck
                      // set icon Ok
                      ListItem.SubItemImages[4] := 6;              // Status OK su colonna corrente.
                    //ListItem.ImageIndex := 6;                    // Status OK su prima colonna.
                      // e per evitare confusione a user, metto RdValue = WrValue ?
                      ListItem.SubItems[4] := tData;   // si, ma visualizzo la copia di Can.Data appena inviata,
                      // perchè ho ricevuto la frame di conferma invio, senza errori.
                      // e non  ListItem.SubItems[4] := ListItem.SubItems[2];  perchè spesso sono in formati diversi Hex/Dec,+/-, ecc..

                      // ora se il DUT-ID non era default, dovrei aggiornare TEdit con nuovo ID, BuadRate e prevedere reset per attivarli
                      // Baudrate richiede anche Reinit completo della CAN !
                      // Ma è meglio solo avviso e scrivere questi parametri solo a fine EOL grazie a StoreAs = 'END' in file di configurazione xls.
                      if cmdAlias = _Node_ID_ then begin
                          if edtDUTaddr.text <> tData then begin
                              LogStep(431,431, 431, #9#9'Warning CAN Node_ID change from:'+edtDUTaddr.text +'  To:'+ tData);
                              // Non aggiorno TEdit perchè non va usato subito, perchè attivare il nuovo ID serve un Reset,
                              //edtDUTaddr.text := tData;
                              // quindi negli step per parametri successivi resta ancora il Node_ID precedente !
                          end;
                      end;
                      if cmdAlias = _Baud_Rate_ then begin  // se il Baud_Rate non era default, devo aggiornare TEdit con nuovo rate,
                          if cbxCANbaud.text <> tData then begin
                              LogStep(432,432, 432, #9#9+'Warning CAN Baud_Rate change from:'+cbxCANbaud.text +'  To:'+ tData);
                              // Non aggiorno TEdit perchè non va usato subito, perchè attivare il nuovo Baud_Rate serve un Reset,
                              // che andrebbe anche in conflitto con i REF che restano a 500Kbit !
                              //edtCANbaud.text := tData;
                              // quindi negli step per parametri successivi resta ancora il Baud_Rate attuale !
                          end;
                      end;
                  end
                  else begin
                      // singolo Errore non è bloccante, lascio portare a termine loop su tutti i params.
                      // leave checked
                      // set icon Warning
                      ListItem.SubItemImages[4] := 22;           // Status Error su colonna corrente.
                    //ListItem.ImageIndex := 2;                  // Status Error su prima colonna.
                      LogStep(433,433, 433,  lastCAN_ErrMsg );
                      inc(errCount);
                      continue;
                  end;

                  // riBlocco se era inhibit
                  if cmdAlias = _Inhibit_between_TPDOs_ then             // se sbloccato per WR Inhibit
                      CAN_setTPDO1_comms(edtDUTaddr.text, modeEnabled);  // ripristino il blocco.

              end;//if checked
          end;//for lviewCFG.Items

          // qui devo bloccare in caso di errore, perchè indica scrittura configurazione incompleta !
          if errCount > 0 then begin
              JobParams.retMsg := _Errore_;
              JobParams.retCode := resERROR;    // problemi di WR config è di tipo bloccante.
              result := False;
          end
          else begin
              JobParams.retMsg := _Completato_; // Ok, anche se trovati parametri da aggiornare,
              JobParams.retCode := resOK;       // perchè è attività a cura degli step successivi!
              result := True;
          end;

        except
          on E: Exception do begin
              JobParams.retCode := resERROR;
              LogSysEvent(svEXCEPTION, 300, 'in '+JobList.Strings[currJob]+' on CAN Write Parameters >> '+ E.ClassName +#10#13#9+ ' description: '+E.Message);
          end;
        end;
    finally
        // Blocca subito l'area protetta 5555h del DUT, perchè il Reset node NON dà lo stesso risultato.
        if not CAN_LockParams( edtDUTaddr.text ) then begin
            // l' Errore è bloccante, quindi esco.
            JobParams.retMsg := 'DUT LockParams ERROR: '+ lastCAN_ErrMsg;
            JobParams.retCode := resERROR;
            inc(errCount);
            LogStep(434,434, 434, 'ERROR on LOCK params for DUT:'+ edtDUTaddr.text );
            LogStep(435,435, 435,  lastCAN_ErrMsg );
            result := False;             // necessario override result!
        end
        else
            LogStep(436,436, 436, 'OK, DUT Lock command CONFIRMED by NodeID:'+ edtDUTaddr.text );

        // NB: Reset DUT per rendere operativi i parametri salvati in eeprom
        //     non serve perchè viene fatto in Step seguente con Save e DUT reset !
        //     Comunque Read back cfg, non serve se WR senza errori...
        LogStep(437,437, 437,  #9'Parameters written: '+ intToStr(countWritten));
        LogStep(438,438, 438, JobList.Strings[currJob]+' ERRORS = '+ intToStr(errCount));
        edtToWrite.text := intToStr( StrToInt(edtToWrite.text) - countWritten);    // aggiorna label toWrite
        // sarebbe da contare i checked towrite che include anche i manuali...
        if StrToInt(edtToWrite.text) > 0 then
            edtToWrite.Color := clYellow   // = necessario ancora overwrite con defaults.
        else
            edtToWrite.Color := clLime;    // = tutti OK.
        dmBase.wait(50);  // solo permettere application process messages e refresh GUI.

        // NB: Reset CAN interface e DUT sarebbe necessario subito nel caso di cambio BaudRate o nodeID !
        //     ma NON è il caso, perchè andrebbe in crash la rete CAN con i REF che sono sempre a 500Kbit !!!
        //     e non è possible una gestione separate nè pensabile rinconfigurare pure i REF !!!
        //     Ha senso solo dare la possibilità di configurare manualmente il BaudRate ad inizio Fasi (Fase_A)
        //     così che user possa rileggere e verificare i parametri DUT...
        {
        if vResetCAN_todo then...
        }
    end;
end;

function Toiac3Fasi2ax.RunStep_A4(var JobParams: TJobParams): boolean;
var
  errCount: integer;
begin
  //mainForm.mmLog.lines.add('RunStep_4...');
    LogText('   *** Inizio Salvataggio Configurazione DUT ***');
    result := False;
    errCount := 0;
    JobParams.retCode := resUNDEFINED;
    JobParams.retMsg := _Incompleto_;

    // prima di tutto assicura CAN channel "clean"
    if not CAN_CleanCH(CHref) then begin
    	// An error occurred.  We show the error.
        JobParams.retMsg := 'CAN_CleanCH1 ERROR: '+ lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        inc(errCount);
        LogStep(439,439, 439,  lastCAN_ErrMsg );
        exit;
    end;
    if CHdut.Handle <> CHref.Handle then
        if not CAN_CleanCH(CHdut) then begin
            // An error occurred.  We show the error.
            JobParams.retMsg := 'CAN_CleanCH2 ERROR: '+ lastCAN_ErrMsg;
            JobParams.retCode := resERROR;
            inc(errCount);
            LogStep(440,440, 440,  lastCAN_ErrMsg );
            exit;
        end;
    LogStep(441,441, 441, JobList.Strings[currJob]+': STATUS PCAN_OK');
    // The channel is Ready to work.

    try
        try
            // Save cfg: parametri modificati via comandi vengono salvati in EEPROM in modo definitivo
            if CAN_NodeSaveAll( edtDUTaddr.text ) then begin
                // The frame was successfully sent
                LogStep(442,442, 442, 'CAN SaveParams command SENT Successfully to NodeID:'+ edtDUTaddr.text );
            end
            else begin
                // l' Errore è bloccante, quindi esco.
                LogStep(443,443, 443, 'CAN SaveParams ERROR on NodeID:'+ edtDUTaddr.text );
                LogStep(444,444, 444,  lastCAN_ErrMsg );
                inc(errCount);
                exit;
            end;
            // Ok, niente msec. attesa perchè WR in eeprom già completato.

            // Reset DUT per rendere operativi i parametri salvati in eeprom.
            if CAN_SendNMT( CHdut, cmdResetNode, edtDUTaddr.text ) then begin
                // The Reset was successfully sent
                LogStep(445,445, 445, 'CAN ResetNode SENT Successfully to NodeID:'+ edtDUTaddr.text );
            end
            else begin
                // l' Errore è bloccante, quindi esco.
                LogStep(446,446, 446, 'CAN ResetNode ERROR  on NodeID:'+ edtDUTaddr.text );
                LogStep(447,447, 447,  lastCAN_ErrMsg );
                inc(errCount);
                exit;
            end;
            // DUT reset & boot completed !
            // Read back cfg, non serve se WR senza errori...

            JobParams.retMsg := _Completato_;
            JobParams.retCode := resOK;
            result := True;
        except
          on E: Exception do begin
              JobParams.retCode := resERROR;
              LogSysEvent(svEXCEPTION, 400, 'in '+JobList.Strings[currJob]+' on CAN Save & Reset node >> '+ E.ClassName +#10#13#9+ ' description: '+E.Message);
          end;
        end;
    finally
        LogStep(448,448, 448, JobList.Strings[currJob]+' ERRORS = '+ intToStr(errCount));
        dmBase.wait(500);// anche solo per evitare click multipli involontari da parte user.
    end;
end;

////#######################  Steps Fase B  ####################################################

function Toiac3Fasi2ax.CalibrazioneDUT(var JobParams: TJobParams; CalDeg:TCalibAngle): boolean;
var
  Can: TCanFrame;
  i, errCount,
  iValueReadX, iValueReadY: integer;
  xREFinRange_counter, yREFinRange_counter: smallint;
  mediaX, mediaY : integer;
  mediaXf, mediaYf : single;
  TargetX, kTargetRefX: single;
  TargetY, kTargetRefY: single;
  CanCmd1, CanCmd2: string;
  cmdPrefix1, cmdPrefix2: string;
  ListItem: TListItem;
  calibRecord: TSessionCalibra;
  CanCmd: TCanCmd;
const
  kDutSamplesInterval = 50;  // che diventano 150 msec. perchè vanno sommati ai 50 già presenti nelle 2 CAN_ReadParam()
begin
{
    Serve aspettare posizionamento device a x0° e y0° per almeno 10 campioni con scarto massimo entro i ±0,03deg !
    quindi attendo almeno 10 letture consecutive <= 0,02 deg
    Quando ottengo tali condizioni, memorizzo il valore di Tensione del DUT e lo memorizzo nei suoi slot dedicati.
    Metto green icon sull'angolo in corso e passo al successivo.
}
    result := False;
    JobParams.retCode := resUNDEFINED;
    JobParams.retMsg := _Incompleto_;
    // Valori come da tabella calibrazione per DUT corrente
    // al momento fissi, ma possibile da tabella xls.
    case CalDeg of
          X0_Y0: begin
                   TargetX := 0;
                   TargetY := 0;
                   CanCmd1 := _2ax_calX0_deg_;       // Campo per calibrazione asse X del DUT.   5555-1
                   CanCmd2 := _2ax_calY0_deg_;       // Campo per calibrazione asse Y del DUT.   5555-2
                   cmdPrefix1 := 'X0';
                   cmdPrefix2 := 'Y0';
                 end;
         XN_Yxn: begin
                   TargetX := 0.3;
                   TargetY := 60;
                   CanCmd1 := _2ax_calXNeg_deg_;     // Campo per calibrazione asse X del DUT.   5555-5
                   CanCmd2 := _2ax_calCross_Yxneg_;  // Campo per calibrazione asse Y del DUT.   5555-31
                   cmdPrefix1 := 'XN';
                   cmdPrefix2 := 'CD';               // Yxn
                 end;
         XP_Yxp: begin
                   TargetX := 0.2;
                   TargetY := -60;
                   CanCmd1 := _2ax_calXPos_deg_;     // Campo per calibrazione asse X del DUT.   5555-3
                   CanCmd2 := _2ax_calCross_Yxpos_;  // Campo per calibrazione asse Y del DUT.   5555-30
                   cmdPrefix1 := 'XP';
                   cmdPrefix2 := 'CC';               // Yxp
                 end;
         YP_Xyp: begin
                   TargetX := 59.9;
                   TargetY := -0.1;
                   CanCmd1 := _2ax_calCross_Xypos_;  // Campo per calibrazione asse Y del DUT.   5555-28
                   CanCmd2 := _2ax_calYPos_deg_;     // Campo per calibrazione asse X del DUT.   5555-4
                   cmdPrefix1 := 'CA';               // Xyp
                   cmdPrefix2 := 'YP';
                 end;
         YN_Xyn: begin
                   TargetX := -60.1;
                   TargetY := 0.1;
                   CanCmd1 := _2ax_calCross_Xyneg_;  // Campo per calibrazione asse Y del DUT.   5555-29
                   CanCmd2 := _2ax_calYNeg_deg_;     // Campo per calibrazione asse X del DUT.   5555-6
                   cmdPrefix1 := 'CB';               // Xyn
                   cmdPrefix2 := 'YN';
                 end;
    end;
    edtTargetX.Text := format('%6.2f°', [TargetX]);
    edtTargetY.Text := format('%6.2f°', [TargetY]);

    kTargetRefX := (TargetX);     // Per ora no inversione.
    kTargetRefY := (TargetY);     //

    gaugeXRange.ValueStart := kTargetRefX - 0.5;   // Aggiunge marker visible ai target degs
    gaugeXRange.ValueEnd   := kTargetRefX + 0.5;
    gaugeYRange.ValueStart := kTargetRefY - 0.5;
    gaugeYRange.ValueEnd   := kTargetRefY + 0.5;

    LogText('   *** Inizio calibrazione ad angoli X='+ edtTargetX.Text +' e Y='+ edtTargetY.Text +'  ***');

    // prima di tutto assicura CAN channel "clean"
    if not CAN_CleanCH(CHref) then begin
    	// An error occurred.  We show the error.
        JobParams.retMsg := 'CAN_CleanCH1 ERROR: '+ lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        inc(errCount);
        LogStep(449,449, 449,  lastCAN_ErrMsg );
        exit;
    end;
    if CHdut.Handle <> CHref.Handle then
        if not CAN_CleanCH(CHdut) then begin
            // An error occurred.  We show the error.
            JobParams.retMsg := 'CAN_CleanCH2 ERROR: '+ lastCAN_ErrMsg;
            JobParams.retCode := resERROR;
            inc(errCount);
            LogStep(450,450, 450,  lastCAN_ErrMsg );
            exit;
        end;
    LogStep(451,451, 451,  JobList.Strings[ currJob ] +': STATUS PCAN_OK');
    // The channel is Ready to work.

    mainForm.showUserMessage('Posizionare il REF a inclinazioni richieste su assi X e Y'+sLineBreak+'poi attendere conferma avvenuta calibrazione.');
    // NB: ATTENZIONE a non fare altri Reset dei node !!!
    //     perchè annulla i setup temporanei necessari (es. risoluzione in centesimi) ricaricando i default !

    // Sblocca area protetta 5555h del DUT per precauzione, se non fosse già da setup fase.
    if not CAN_UnlockParams( edtDUTaddr.text ) then begin
        // l' Errore è bloccante, quindi esco.
        JobParams.retMsg := 'DUT UnlockParams ERROR: '+ lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        inc(errCount);
        LogStep(452,452, 452,  JobParams.retMsg );
        exit;           // unlock non riuscito, quindi non serve lock finally !
    end;
    LogStep(453,453, 453, 'OK, DUT Unlock command CONFIRMED by NodeID:'+ edtDUTaddr.text );

    try
        // Ref2 START node, enter operational state
        if not CAN_SendNMT( CHref, cmdStartNode, edtRef2Addr.text ) then begin
            // l' Errore è bloccante, quindi esco.
            JobParams.retMsg := 'REF2 StartNode ERROR: '+ lastCAN_ErrMsg;
            JobParams.retCode := resERROR;
            inc(errCount);
            LogStep(454,454, 454, 'ERROR on StartNode REF2:'+ edtRef2Addr.text );
            LogStep(455,455, 455,  lastCAN_ErrMsg );
            exit;
        end;
        // The Start was successfully sent
        LogStep(456,456, 456, 'OK StartNode SENT Successfully to Ref2: '+ edtRef2Addr.text );

        if not CAN_FlushRxBuffer(CHdut) then begin
            JobParams.retMsg := 'ERROR on Flush DutCAN buffer: '+ lastCAN_ErrMsg;
            JobParams.retCode := resERROR;
            LogStep(457,457, 457,  JobParams.retMsg );
            exit;
        end;
        TInterlocked.BitTestAndClear(SampleReady, 0);  // Reset SampleReady bit 0
        // inversioni angoli se richieste.
        vInvertX_ref := menuInvertiXref.checked;
        vInvertY_ref := menuInvertiYref.checked;
        vInvertX_dut := menuInvertiXdut.checked;
        vInvertY_dut := menuInvertiYdut.checked;
        tmrCanRead.enabled := True;                    // avvia timer ricezione ciclica valori da CAN (Ref2)

        edtRefX.color := clYellow;
        edtRefY.color := clYellow;
        edtDutX.text := '';
        edtDutY.text := '';
        edtDutX.color := clBtnface;
        edtDutY.color := clBtnface;
        currSequence := 1;                   // n° Sequenza per blink controls.
        tmrFasiBlink.Enabled := True;        // serve alla sequenza che parte subito.
        application.ProcessMessages();

        // Loop di lettura angolo del REFerence e write calibrazione in DUT.
        scartoMax := strTofloat(edtScartoMax_calib.text);  // = ±0.04
        vRefSamplesInRange := edtRefSamplesInRange.value;
        xREFinRange_counter := 0;
        yREFinRange_counter := 0;
        errCount := 0;

        inStepLoop := True;
        try
          while true do begin // resta in endless loop fino a regolazione ok, o user abort.
              // check per Cancel ad ogni giro.
              application.processmessages;
              if dmBase.WaitAbort(30) then begin
                  LogText('*** ABORT by USER, calibrazione Interrotta !'+#13#10 );
                  exit;
              end;

              // Attende valori Reference rilevati dal Timer, che siano in Range previsto (entro ±0,02deg)
              if SampleReady = 0 then
                  continue;
              // qui il loop Delay è solo 30 msec, mentre il timer è a 250 msec, quindi tutto il tempo di servire
              // i campioni validi trovati dal Timer.
              // Se resetto subito il flag sono sicuro di non perdere il prossimo set da parte del timer.

              //LogStep(458,458, 458, 'Flag SampleReady = '+ format('$%8.8x', [SampleReady]) );
              TInterlocked.BitTestAndClear(SampleReady, 0);     // Reset SampleReady bit 0
              //LogStep(459,459, 459, 'Flag SampleReady = '+ format('$%8.8x', [SampleReady]) +' =>> RELEASED');

              // verifica in due tempi per Ref X=xTargetRef° Y=yTargetRef°  con scarto max di  ±scartoMax° (±0,04deg)
              if (currXRef < (-scartoMax + kTargetRefX)) or (currXRef > (scartoMax + kTargetRefX)) then
                  xREFinRange_counter := 0     // reset count, perchè REFinRange devono essere consecutivi !
              else
                  inc(xREFinRange_counter);
              if (currYRef < (-scartoMax + kTargetRefY)) or (currYRef > (scartoMax + kTargetRefY)) then
                  yREFinRange_counter := 0     // reset count, perchè REFinRange devono essere consecutivi !
              else
                  inc(yREFinRange_counter);
              // ok, i REF samples sono in range richiesto.

              // Aggiorna colori Rosso, Giallo, Verde se n° samples sufficiente a calibrare.
              if xREFinRange_counter < vRefSamplesInRange then begin  // vRefSamplesInRange default 10
                  edtRefX.color := clYellow;      // ancora pochi X in range; ne attendo altri...
                  if kTargetRefX <> 0 then
                      if ((kTargetRefX > 0) and (currXRef < 0)) or ((kTargetRefX < 0) and (currXRef > 0)) then begin
                          edtRefX.color := clRed; // X è su angolo opposto, avvisa user.
                      end;
              end
              else
                  edtRefX.color := clLime;        // avvisa operatore X è ok.
              if yREFinRange_counter < vRefSamplesInRange then begin  // vRefSamplesInRange default 10
                  edtRefY.color := clYellow;      // ancora pochi Y in range; ne attendo altri...
                  if kTargetRefY <> 0 then
                      if ((kTargetRefY > 0) and (currYRef < 0)) or ((kTargetRefY < 0) and (currYRef > 0)) then begin
                          edtRefY.color := clRed; // Y è su angolo opposto, avvisa user.
                      end;
              end
              else
                  edtRefY.color := clLime;        // avvisa operatore Y è ok.

              // Samples devono essere n° sufficiente, ma contemporaneamente su entrambi gli assi.
              if (xREFinRange_counter < vRefSamplesInRange) or (yREFinRange_counter < vRefSamplesInRange) then
                  // ancora pochi X o Y in range; ne attendo altri...
                  continue;

              LogStep(460,460, 460, 'REF_in_Range counters   X='+ xREFinRange_counter.toString +'  Y='+ yREFinRange_counter.toString );

              // OK, ho ricevuto almeno 10 valori X e Y REF consecutivi nel range richiesto (entro ±0,02deg),
              // quindi ora che REF è stabile in range da almeno 10 letture
              // leggo valori dal DUT, non troppo ravvicinati, ne faccio media e la inserisco in campo calibrazione del DUT.

              // Ma non essendo a controllo motorizzato, non posso evitare che User muova ancora il piano durante l'attesa,
              // cosa che genera calcoli inconsistenti, quindi lo Avviso con BEEP quando deve fermarsi !
              Winapi.Windows.Beep( 1500, 800) ;   // Avviso user che deve fermarsi ! (frequency, duration)

              xREFinRange_counter := 0;
              yREFinRange_counter := 0;
              edtRefX.color := clLime;    // avvisa operatore di fermarsi.
              edtRefY.color := clLime;    //

              // qui ricevo ancora anche Cyclic transmission dal REF, che si mischiano alle risposte dei prossimi comandi,
              // se non le fermo o le distinguo...!
              // Visto che le letture dal DUT non devono essere troppo ravvicinate, posso spendere tempo per stoppare il REF!
              // prima ferma Timer ricezione ciclica valori da REF altrimenti tmrCanRead si mangia (perdo) i response dei prossimi cmd !
              tmrCanRead.enabled := False;
              if not CAN_SendNMT( CHref, cmdStopNode, edtRef2Addr.text ) then begin
                  JobParams.retMsg := 'ERROR on StopNode REF2: '+ lastCAN_ErrMsg;
                  JobParams.retCode := resERROR;
                  inc(errCount);
                  LogStep(461,461, 461,  JobParams.retMsg );
                  exit;
              end;
              LogStep(462,462, 462, 'OK REF2 node Stopped.');

              if not CAN_FlushRxBuffer(CHref) then begin
                  JobParams.retMsg := 'ERROR on Flush RefCAN buffer: '+ lastCAN_ErrMsg;
                  JobParams.retCode := resERROR;
                  inc(errCount);
                  LogStep(463,463, 463,  JobParams.retMsg );
                  exit;
              end;
              LogStep(464,464, 464, 'OK FLUSH RefCAN queue Successful.');    // preparato rx queue clean.

              // leggo Almeno 10 valori DUT non troppo ravvicinati, a 50-100 msec...
              // Al momento il delay interno all CAN_ReadParam() è fisso a 50 msec !
              mediaX := 0;
              mediaY := 0;
              for i := 1 to vRefSamplesInRange do begin  // default vRefSamplesInRange è 10.

                  dmBase.Wait(kDutSamplesInterval); // sono 150 msec. perchè vanno sommati ai 50 già presenti nelle CAN_ReadParam()

                  // legge valore X da DUT as UNS32. (5555-7)
                  Can.Command  := _ADXL203_X_mV_;
                  Can.Data     := '';                            // '' perchè è Read quindi ospiterà il risultato.
                  Can.NodeID   := edtDUTaddr.text;               // str ID del DUT.
                  if not CAN_ReadParam( Can ) then begin         // Delay standard è 50 msec.
                      JobParams.retMsg := 'DUT Read '+Can.Index+'-'+Can.Subidx+'h ERROR: '+ lastCAN_ErrMsg;
                      JobParams.retCode := resERROR;
                      inc(errCount);
                      LogStep(465,465, 465,  JobParams.retMsg );
                      break;
                  end;
                  // The parameter was Read successfully.
                  iValueReadX := StrToIntDef('$'+Can.Data, 0);       // $ perchè payload è stringa ma sempre Hex !
                  mediaX := mediaX + iValueReadX;
                  //LogStep(466,466, 466, #9#9'Get DUT '+ Can.Command +' = '+ Can.Data );
                  //LogStep(467,467, 467, 'OK, DUT Parameter ['+Can.Command+'] READ Successfully from DutID:'+ edtDUTaddr.text );

                  // legge valore Y da DUT as UNS32. (5555-8)
                  Can.Command  := _ADXL203_Y_mV_;
                  Can.Data     := '';                            // '' perchè è Read quindi ospiterà il risultato.
                  Can.NodeID   := edtDUTaddr.text;               // str ID del DUT
                  if not CAN_ReadParam( Can ) then begin         // Delay standard è 50 msec.
                      JobParams.retMsg := 'DUT Read '+Can.Index+'-'+Can.Subidx+'h ERROR: '+ lastCAN_ErrMsg;
                      JobParams.retCode := resERROR;
                      inc(errCount);
                      LogStep(468,468, 468,  JobParams.retMsg );
                      break;
                  end;
                  // The parameter was Read successfully.
                  iValueReadY := StrToIntDef('$'+Can.Data, 0);       // $ perchè payload è stringa ma sempre Hex !
                  mediaY := mediaY + iValueReadY;
                  //LogStep(469,469, 469, #9#9'Get DUT '+ Can.Command +' = '+ Can.Data );
                  //LogStep(470,470, 470, 'OK, DUT Parameter ['+Can.Command+'] READ Successfully from DutID:'+ edtDUTaddr.text );
                  LogStep(471,471, 471, #9'Got DUT    X = '+ iValueReadX.toString +'     Y = '+ iValueReadY.toString );

              end;//for vRefSamplesInRange

              // Calcolo media dei vRefSamplesInRange campioni e la memorizza nello slot prestabilito...
            //mediaX := mediaX div 10;  evito DIV per media perchè con 1 solo valore su 10 minore,
            //mediaY := mediaY div 10;  tronca su quel minore, perdendo il peso degli altri 9 !
              mediaXf := mediaX / vRefSamplesInRange;
              mediaYf := mediaY / vRefSamplesInRange;
              SetRoundMode(rmNearest);
              mediaX := round(mediaXf);
              mediaY := round(mediaYf);
              edtDutX.text := mediaX.toString;
              edtDutY.text := mediaY.toString;
              LogStep(472,472, 472, 'OK, DUT mediaX = '+edtDutX.text+'   mediaY = '+ edtDutY.text );

              // Sblocco area protetta 5555h del DUT è già fatto nella relativa FASE preliminare.
              // Memorizza valore X di calibrazione ottenuto nel DUT.
              Can.Command  := CanCmd1;                      // Campo per calibrazione assi del DUT.
              Can.Data     := cmdPrefix1 + Ansichar(LongRec(mediaX).Bytes[1]) + Ansichar(LongRec(mediaX).Bytes[0]); // NB: usare Ansichar() e non char() o chr() che tronca a 7 bit !!!
            //Can.Data     := cmdPrefix1 + Ansichar(wordrec(mvX).Bytes[1]) + Ansichar(wordrec(mvX).Bytes[0]);
              //  Data va trattato come VSTR perchè vuole prefisso 'X0'+ value 16 bit senza Lsb-Msb swap, come da indicazioni WT, quindi Msb+Lsb: byte[1]+byte[0]
              Can.NodeID   := edtDUTaddr.text;              // str ID del DUT
              if not CAN_WriteParam( Can ) then begin
                  JobParams.retMsg := 'DUT WrParam '+Can.Index+'-'+Can.Subidx+'h ERROR: '+ lastCAN_ErrMsg;
                  JobParams.retCode := resERROR;
                  inc(errCount);
                  LogStep(473,473, 473,  JobParams.retMsg );
                  break;
              end;
              LogStep(474,474, 474, #9#9'Set DUT param '+ Can.Command +' = '+ edtDutX.text +' ('+mediaX.ToHexString(4)+'h)');
              LogStep(475,475, 475, 'OK, DUT Calibration  '+Can.Index+'-'+Can.Subidx+'h STORED Successfully to DutID:'+ edtDUTaddr.text );

              // Memorizza valore Y di calibrazione ottenuto nel DUT.
              Can.Command  := CanCmd2;                      // Campo per calibrazione assi del DUT.
              Can.Data     := cmdPrefix2 + Ansichar(LongRec(mediaY).Bytes[1]) + Ansichar(LongRec(mediaY).Bytes[0]);
              //  Data va trattato come VSTR perchè vuole prefisso 'Y0'+ value 16 bit senza Lsb-Msb swap, come da indicazioni WT, quindi Msb+Lsb: byte[1]+byte[0]
              Can.NodeID   := edtDUTaddr.text;              // str ID del DUT
              if not CAN_WriteParam( Can ) then begin
                  JobParams.retMsg := 'DUT WrParam '+Can.Index+'-'+Can.Subidx+'h ERROR: '+ lastCAN_ErrMsg;
                  JobParams.retCode := resERROR;
                  inc(errCount);
                  LogStep(476,476, 476,  JobParams.retMsg );
                  break;
              end;
              // The parameter was Written successfully.
              LogStep(477,477, 477, #9#9'Set DUT param '+ Can.Command +' = '+ edtDutY.text +' ('+mediaY.ToHexString(4)+'h)');
              LogStep(478,478, 478, 'OK, DUT Calibration  '+Can.Index+'-'+Can.Subidx+'h STORED Successfully to DutID:'+ edtDUTaddr.text );

              // compila lista per review.
              lviewCarattBanco.Items.BeginUpdate;
              ListItem := lviewCalibVal.Items.Add;
            //ListItem.ImageIndex := -1;                  // per ora niente icon
              ListItem.Caption := strCalibAngle[ord(CalDeg)];
              ListItem.SubItems.Add( edtDutX.text );
              ListItem.SubItems.Add( edtDutY.text );
              lviewCarattBanco.Items.EndUpdate;

              // Per eventuale read-back di verifica calibrazioni è necessario prima un DUT node Reset per portare i nuovi parametri da eeprom a ram !
              // Blocco l'area protetta 5555h del DUT sia in Wr che Read, meglio a fine procedure collaudo ? No subito.

              break;      // infine esce sempre !!!
          end;//// endless loop

          // In tutti i casi compilo (o sovrascrivo) il record con il valore realmente inserito nel DUT.
          calibRecord := TSessionCalibra.create;
          with calibRecord do begin
            PK              := -1;                          // solo preset, sarà poi assegnato da DB.
            SessionID       := -1;                          // 1-N con Sessione, sarà poi assegnato da DB.
            Lotto           := Sessione.Lotto;              // necessario perchè s/n non è univoco!
            Serial          := Sessione.Serial;             // Relazione 1-N con il record SessionData e DutData.
            degPairID       := strCalibAngle[ord(CalDeg)];  // Mnemonico della combinazione Angolo in Calibrazione TCalibAngle.
            TargetX         := kTargetRefX;                 // deg. posizionamento richiesto
            TargetY         := kTargetRefY;
            ScartoAmmesso   := scartoMax;                   // centesimi deg. fra target e Ref ottenuto
            RefSamplesInRange  := vRefSamplesInRange;            // num. per considerare Ref valido e calcolarne valore Medio.
            IntervalloRefReads := strToInt(kTimer4cyclicalTxd);  // msec. fra letture consecutive del Reference; è il suo intervallo configurato di autoTX.
            IntervalloDutReads := 100;                           // 50+50 msec. fra letture consecutive del DUT
            RefAx        := currXRef;           // posizionamento ottenuto del Ref in deg
            RefAy        := currYRef;
            DutAx_mV     := mediaX;             // qui valori medi delle letture Dut in tensione $0000-$FFFF.
            DutAy_mV     := mediaY;             //
            AliasX       := CanCmd1;            // Alias per ora non usati, dato che a supporto delle letture sequenziali da DUT,
            AliasY       := CanCmd2;            // al momento copio i valori di Calib nell'alias del dictionary dicCanCmds.
          end;
          // Non serve vedere se il record per questo angolo in calibrazione è già presente nel dictionary,
          // perchè Add() lo modifica o lo aggiunge se non c'è già.
          CalibRecords.Add( CalDeg, calibRecord);
          // CalibRecords serve solo per display e report, non per R/W parametri di calibrazione.

          // infine Conservo i valori di Calibrazione trovati, nel Dictionary dicCanCmds, associati al proprio Alias
          // così potrò usarli in prossime letture sequenziali di verifica configurazione DUT !
          if not dicCanCmds.TryGetValue( CanCmd1, CanCmd ) then begin
              JobParams.retMsg := 'ERROR Alias NOT Found "'+ CanCmd1 +'" in dictionary CanCmds.';
              JobParams.retCode := resERROR;
              inc(errCount);
              LogStep(479,479, 479,  JobParams.retMsg );
              exit;
          end;
          CanCmd.WrValue := mediaX.toString;
          if not dicCanCmds.TryGetValue( CanCmd2, CanCmd ) then begin
              JobParams.retMsg := 'ERROR Alias NOT Found "'+ CanCmd2 +'" in dictionary CanCmds.';
              JobParams.retCode := resERROR;
              inc(errCount);
              LogStep(480,480, 480,  JobParams.retMsg );
              exit;
          end;
          CanCmd.WrValue := mediaY.toString;

          // Reset gauges
          gaugeX.Value := 0;
          gaugeY.Value := 0;

          // qui devo bloccare in caso di errore, perchè indica calibrazione incompleta !
          if errCount > 0 then begin
              edtRefX.color := clRed;
              edtRefY.color := clRed;
              JobParams.retMsg := _Errore_;
              JobParams.retCode := resERROR;    // problemi di acquisizione REF o WR calibrazione nel DUT sono di tipo bloccante.
              inc(errCount);
              result := False;
          end
          else begin
              JobParams.retMsg := _Completato_; // Ok, angolo 0° calibrato.
              JobParams.retCode := resOK;
              result := True;
          end;

        except
          on E: Exception do begin
              JobParams.retCode := resERROR;
              edtRefX.color := clBtnface;
              edtRefY.color := clBtnface;
              LogSysEvent(svEXCEPTION, 500, 'in '+ JobList.Strings[currJob] +' on DUT calibration >> '+ E.ClassName +#10#13#9+ ' description: '+E.Message);
          end;
        end;
    finally
        inStepLoop := False;               // va sempre in finally!
        tmrCanRead.enabled := False;       // Se non lo è già.
        currSequence := -1;                // Reset Sequenza blink current Job node.
        tmrFasiBlink.Enabled := False;     // termina blink.
        img1Seq1.Visible := False;         //

        // Blocca sempre l'area protetta 5555h del DUT, perchè il Reset node NON dà lo stesso risultato.
        if not CAN_LockParams( edtDUTaddr.text ) then begin
            // l' Errore è bloccante, quindi esco.
            JobParams.retMsg := 'DUT LockParams ERROR: '+ lastCAN_ErrMsg;
            JobParams.retCode := resERROR;
            LogStep(481,481, 481, 'ERROR on LOCK params for DUT:'+ edtDUTaddr.text );
            LogStep(482,482, 482,  lastCAN_ErrMsg );
            inc(errCount);
            result := False;
        end
        else
            LogStep(483,483, 483, 'OK, DUT Lock command CONFIRMED by NodeID:'+ edtDUTaddr.text );

      //TestRecord.free;  No, si perde il record di calibrazione appena memorizzato!  che serve in report !
        LogStep(484,484, 484,  JobList.Strings[currJob] +' ERRORS = '+ intToStr(errCount));
        CAN_SendNMT( CHref, cmdStopNode, edtRef2Addr.text );      // blind stop invio ciclico, se non è già a causa Abort.
        // Flush più che altro per lasciare pulizia a steps successivi.
        CAN_FlushRxBuffer(CHref);         // qui meglio no log, perchè non serve e posso perdere
        if CHdut.Handle <> CHref.Handle then           // eventuali altri errMsg di exit via finally !
            CAN_FlushRxBuffer(CHdut);
    end;
end;

function Toiac3Fasi2ax.RunStep_B5(var JobParams: TJobParams): boolean;
begin
    result := CalibrazioneDUT( JobParams, X0_Y0 );
end;
function Toiac3Fasi2ax.RunStep_B6(var JobParams: TJobParams): boolean;
begin
    result := CalibrazioneDUT( JobParams, XP_Yxp );
end;
function Toiac3Fasi2ax.RunStep_B7(var JobParams: TJobParams): boolean;
begin
    result := CalibrazioneDUT( JobParams, XN_Yxn );
end;
function Toiac3Fasi2ax.RunStep_B8(var JobParams: TJobParams): boolean;
begin
    result := CalibrazioneDUT( JobParams, YP_Xyp );
end;
function Toiac3Fasi2ax.RunStep_B9(var JobParams: TJobParams): boolean;
begin
    result := CalibrazioneDUT( JobParams, YN_Xyn );
end;


////#######################  Steps Fase C e D  #######################################################

function Toiac3Fasi2ax.ValidazioneDUT(var JobParams: TJobParams; testDeg:TTestAngle; testAxis:char): boolean;
var
  Can: TCanFrame;
  iDeg: smallint;
  i, errCount,
  iValueReadX, iValueReadY: integer;
  xREFinRange_counter, yREFinRange_counter: smallint;
  TargetX, kTargetRefX: single;  // Angoli entro intervalli previsti in XLS sheet 'TAB.ERR'
  TargetY, kTargetRefY: single;  //
  prevRead,
  prevAngleCross, prevAngleTest,
  pendTest, pendCross,
  testP1, testP0,
  crossP1, crossP0 : single;
  fattCORR_test, fattCORR_cross,
  RefCorretto_test, RefCorretto_cross,
  DutERR_test, DutERR_cross: single;
  DutStable: boolean;
  TestRecord: TSessionTests;
  currRef_test, currRef_cross,
  currDut_test, currDut_cross :single;
const
  kDutSamplesStable = 7;
  kDutSamplesScarto = 0.1;
  kDutSamplesInterval = 300;     // da sommare ai 25+50 ms impiegati dalla lettura per ogni parametro.
begin
{
    Avvio il test con campi Yellow e
    Attendo posizionamento del REF (RefAx, RefAy) ai gradi X,Y previsti da tabella per asse in TEST,
    ma questa volta entro un range ampio perchè per la verifica serve solo una posizione di massima, entro i 2°.
    A tali Reference devo *sottrarre* il fattore di correzione X e Ycross rispettivamente.
    CorrRefAx e CorrRefAy però vanno calcolati a seconda del range° (previsto in tabella di caratterizzazione banco)
    in cui cadono le misure RefAx,RefAy ovvero per entrambe si considerano i 2 valori dell'intervallo,
    e si utilizza il valore medio, che come detto va sottratto ai rispettivi RefAx,RefAy.
    Solo a questo punto i valori di Reference così compensati, possono essere confrontati con le letture dal DUT DevAx,DevAy.
    Se la differenza (errore) risulta inferiore ai valori ErrMaxTest ed ErrMaxCross il test è considerato superato,
    e i campi passano a green.
    Il procedimento viene applicato in parallelo su Asse di TEST e su quello CROSS.
}
    result := False;
    errCount := 0;
    JobParams.retCode := resUNDEFINED;
    JobParams.retMsg := _Incompleto_;
    // Angoli del REF da posizionare per test come da indicazioni Walter, con tolleranza larga +/- 2° !
    // viene testato un Asse per volta;
    // asse X posizionato a tutti gli angoli previsti, con Y=0
    // poi asse Y posizionato a tutti gli angoli previsti, con X=0
    // Angoli di test sono rilevati dal REFERENCE per punti prestabiliti fissi, ma possibile anche da DB o tabella xls.
    case testDeg of      // testDeg può essere riferito a X o Y, a seconda dell'asse che è in Test !  default X.
         deg60n: begin
                   TargetX := -60;
                   TargetY := 0;
                 end;
         deg30n: begin
                   TargetX := -30;
                   TargetY := 0;
                 end;
         deg15n: begin
                   TargetX := -15;
                   TargetY := 0;
                 end;
           deg0: begin
                   TargetX := 0;
                   TargetY := 0;
                 end;
         deg15p: begin
                   TargetX := 15;
                   TargetY := 0;
                 end;
         deg30p: begin
                   TargetX := 30;
                   TargetY := 0;
                 end;
         deg60p: begin
                   TargetX := 60;    // assegnazioni X e Y provvisorie
                   TargetY := 0;     // perchè definite da testAxis = 'X/Y'
                 end;
    end;
    kTargetRefX := (TargetX);     // Per ora no inversione.
    kTargetRefY := (TargetY);     //

    if testAxis = 'Y' then begin       // se Y swap Ref targets.
        kTargetRefY := kTargetRefX;
        kTargetRefX := 0;
    end;

    edtTargetX.Text := format('%6.2f°', [kTargetRefX]);
    edtTargetY.Text := format('%6.2f°', [kTargetRefY]);

    gaugeXRange.ValueStart := kTargetRefX - 0.5;          // Aggiunge marker visible ai target degs
    gaugeXRange.ValueEnd   := kTargetRefX + 0.5;
    gaugeYRange.ValueStart := kTargetRefY - 0.5;
    gaugeYRange.ValueEnd   := kTargetRefY + 0.5;

    LogText('   ****** Inizio Validazione ad angoli X='+ edtTargetX.Text +' e Y='+ edtTargetY.Text +'  ******');

    // prima di tutto assicura CAN channel "clean"
    if not CAN_CleanCH(CHref) then begin
    	// An error occurred.  We show the error.
        JobParams.retMsg := 'CAN_Clean CHref ERROR: '+ lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        inc(errCount);
        LogStep(485,485, 485,  lastCAN_ErrMsg );
        exit;
    end;
    if CHdut.Handle <> CHref.Handle then
        if not CAN_CleanCH(CHdut) then begin
            // An error occurred.  We show the error.
            JobParams.retMsg := 'CAN_Clean CHdut ERROR: '+ lastCAN_ErrMsg;
            JobParams.retCode := resERROR;
            LogStep(486,486, 486,  lastCAN_ErrMsg );
            exit;
        end;
    LogStep(487,487, 487,  JobList.Strings[ currJob ] +': STATUS PCAN_OK');
    // The channels are Ready to work.
    // NB: ATTENZIONE a non fare altri Reset del node !!!
    //     perchè annulla i setup temporanei necessari (es. risoluzione in centesimi) ricaricando i default !

    // Ref2 START node, enter operational state
    LogStep(488,488, 488, 'FLUSH queue All Can CHANNELS...');
    if not CAN_FlushRxBuffer(CHref) then begin    // prepara rx queue clean
        LogStep(489,489, 489,  'ERROR on CHref Flush: '+lastCAN_ErrMsg );
        exit;
    end;
    if CHdut.Handle <> CHref.Handle then
        if not CAN_FlushRxBuffer(CHdut) then begin
            LogStep(490,490, 490,  'ERROR on CHdut Flush: '+lastCAN_ErrMsg );
            exit;
        end;
    LogStep(491,491, 491, 'OK FLUSH All CanCH Successful.');    // preparato rx queue clean.

    if not CAN_SendNMT( CHref, cmdStartNode, edtRef2Addr.text ) then begin
        // l' Errore è bloccante, quindi esco.
        JobParams.retMsg := 'ERROR on REF2 StartNode: '+ lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        LogStep(492,492, 492, JobParams.retMsg);
        exit;
    end;
    // The Start was successfully sent
    LogStep(493,493, 493, 'OK StartNode SENT Successfully to REF: '+ edtRef2Addr.text );

    // DUT START node, enter operational state per tracciare EMCY frames quando viene portato fuori range massimo sballando i calcoli di validazione !
    if not CAN_SendNMT( CHdut, cmdStartNode, edtDUTaddr.text ) then begin
        // l' Errore è bloccante, quindi esco.
        JobParams.retMsg := 'ERROR on DUT StartNode: '+ lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        LogStep(494,494, 494,  JobParams.retMsg );
        exit;
    end;
    // The Start was successfully sent
    LogStep(495,495, 495, 'OK StartNode SENT Successfully to DUT: '+ edtDUTaddr.text );

    // NB: ATTENZIONE a non fare altri Reset dei node !!!
    //     perchè annulla i setup temporanei necessari (es. risoluzione in centesimi) ricaricando i default !
(*
    // rset node per assicurare prossime letture valori.
    if not CAN_SendNMT( CHdut, cmdResetNode, edtDutAddr.text ) then begin
        JobParams.retMsg := 'DUT ResetNode '+edtDutAddr.text+' ERROR: '+ lastCAN_ErrMsg;
        JobParams.retCode := resERROR;
        LogStep(496,496, 496, JobParams.retMsg);
        exit;
    end;
    LogStep(497,497, 497, 'DUT ResetNode SENT Successfully to NodeID:'+ edtDutAddr.text );
*)
    TInterlocked.BitTestAndClear(SampleReady, 0);  // Reset SampleReady bit 0
    // inversioni angoli se richieste.
    vInvertX_ref := menuInvertiXref.checked;
    vInvertY_ref := menuInvertiYref.checked;
    vInvertX_dut := menuInvertiXdut.checked;
    vInvertY_dut := menuInvertiYdut.checked;
    tmrCanRead.enabled := True;                    // avvia timer ricezione ciclica valori da CAN (Ref2)

    mainForm.showUserMessage('Posizionare il DUT a inclinazioni richieste su assi X e Y'+sLineBreak+'poi attendere conferma avvenuta validazione.');

    edtRefX.color := clYellow;
    edtRefY.color := clYellow;
    edtDutX.text := '';
    edtDutY.text := '';
    edtDutX.color := clBtnface;
    edtDutY.color := clBtnface;
    currSequence := 1;                   // n° Sequenza per blink controls.
    vRangeError := False;                // Anche se non ho certezza che sia già a fondo scala (!?) ma verificare è complicato...
    tmrFasiBlink.Enabled := True;        // serve alla sequenza che parte subito.
    application.ProcessMessages();

    // Loop di lettura angolo del DUT e calcoli per validazione.
    scartoMax := strTofloat(edtScartoMax_test.text);  // = ±2.0
    vRefSamplesInRange := edtRefSamplesInRange.value;
    xREFinRange_counter := 0;
    yREFinRange_counter := 0;
    try
        inStepLoop := True;
        try
          while true do begin // resta in endless loop fino a validazione ok, o user abort.
              // check per Cancel ad ogni giro.
              application.processmessages;
              if dmBase.WaitAbort(30) then begin
                  LogText('*** ABORT by USER, validazione DUT Interrotta !'+#13#10 );
                  exit;
              end;
              // qui il loop Delay è solo 30 msec, mentre il timer è a 250 msec, quindi tutto il tempo di servire
              // i campioni validi trovati dal Timer.

              // Attende valori Reference rilevati dal Timer
              if SampleReady = 0 then
                  continue;
              // che siano in Range previsto (entro ±2,0deg) lo verifico ora

              //LogStep(498,498, 498, 'Flag SampleReady = '+ format('$%8.8x', [SampleReady]) );
              // Se resetto subito il flag sono sicuro di non perdere il prossimo set da parte del timer.
              TInterlocked.BitTestAndClear(SampleReady, 0);     // Reset SampleReady bit 0
              //LogStep(499,499, 499, 'Flag SampleReady = '+ format('$%8.8x', [SampleReady]) +' =>> RELEASED');

              // verifica in due tempi per Ref X=xTargetRef° Y=yTargetRef°  con scarto max di  ±scartoMax° (±2,0deg)
              if (currXRef < (-scartoMax + kTargetRefX)) or (currXRef > (scartoMax + kTargetRefX)) then
                  xREFinRange_counter := 0    // reset count, perchè DUTinRange devono essere consecutivi !
              else
                  inc(xREFinRange_counter);
              if (currYRef < (-scartoMax + kTargetRefY)) or (currYRef > (scartoMax + kTargetRefY)) then
                  yREFinRange_counter := 0    // reset count, perchè DUTinRange devono essere consecutivi !
              else
                  inc(yREFinRange_counter);

              // Devo sempre evitare valori DUT fuori scala, innescati da valori REF portato oltre la capacità (range) del Dut !
              // DUT fuori fondo scala, sballa i calcoli di validazione,
              if vRangeError then begin          // quindi segnalo ed esco lasciando clYellow in attesa correzione da user...
                  xREFinRange_counter := 0;
                  yREFinRange_counter := 0;      // reset conteggi, implica aspettare di nuovo tutti i vRefSamplesInRange.
                  continue;
              end
              else
                  lbFondoScala.Visible := False; // toglie Warning.
              // ok, i REF samples sono in range richiesto.

              // Aggiorna colori Rosso, Giallo, Verde se n° samples sufficiente a calibrare.
              if xREFinRange_counter < vRefSamplesInRange then begin  // vRefSamplesInRange default 10
                  edtRefX.color := clYellow;      // ancora pochi X in range; ne attendo altri...
                  if kTargetRefX <> 0 then
                      if ((kTargetRefX > 0) and (currXRef < 0)) or ((kTargetRefX < 0) and (currXRef > 0)) then begin
                          edtRefX.color := clRed; // X è su angolo opposto, avvisa user.
                      end;
              end
              else
                  edtRefX.color := clLime;        // avvisa operatore X è ok.
              if yREFinRange_counter < vRefSamplesInRange then begin  // vRefSamplesInRange default 10
                  edtRefY.color := clYellow;      // ancora pochi Y in range; ne attendo altri...
                  if kTargetRefY <> 0 then
                      if ((kTargetRefY > 0) and (currYRef < 0)) or ((kTargetRefY < 0) and (currYRef > 0)) then begin
                          edtRefY.color := clRed; // Y è su angolo opposto, avvisa user.
                      end;
              end
              else
                  edtRefY.color := clLime;        // avvisa operatore Y è ok.

              // Samples devono essere n° sufficiente, ma contemporaneamente su entrambi gli assi.
              if (xREFinRange_counter < vRefSamplesInRange) or (yREFinRange_counter < vRefSamplesInRange) then
                  // ancora pochi X o Y in range; ne attendo altri...
                  continue;

              LogStep(500,500, 500, 'REF_in_Range counters   X='+ xREFinRange_counter.toString +'  Y='+ yREFinRange_counter.toString );

              // OK, ho ricevuto almeno 10 valori X e Y REF consecutivi nel range richiesto (entro ±2,0deg),
              // quindi ora che REF è stabile in range da almeno 10 letture
              // posso procedere con Correzioni valori reference e confronto con DUT.

              // Ma non essendo a controllo motorizzato, non posso evitare che User muova ancora il piano durante l'attesa,
              // cosa che genera calcoli inconsistenti, quindi lo Avviso con BEEP quando deve fermarsi !
              Winapi.Windows.Beep( 1500, 800) ;   // Avviso user che deve fermarsi ! (frequency, duration)

              xREFinRange_counter := 0;
              yREFinRange_counter := 0;
              edtRefX.color := clAqua;    // avvisa operatore di fermarsi, perchè REF è OK
              edtRefY.color := clAqua;    // ma vediamo se anche campi DUT diventeranno verdi...

              // qui ricevo ancora anche Cyclic transmission dal REF e DUT, che si mischiano alle risposte dei prossimi comandi,
              // se non le fermo o le distinguo...!
              // Visto che le letture dal DUT non devono essere troppo ravvicinate, posso spendere tempo per stoppare il REF!
              // prima ferma Timer ricezione ciclica valori dai nodes REF e DUT altrimenti tmrCanRead() si mangia (perdo) i response dei prossimi cmd !
              tmrCanRead.enabled := False;
//            if not CAN_SendNMT( CHref, cmdStopNode, edtRef2Addr.text ) then begin
              if not CAN_SendNMT( CHref, cmdResetNode, edtRef2Addr.text ) then begin
//                JobParams.retMsg := 'ERROR on StopNode REF: '+ lastCAN_ErrMsg;
                  JobParams.retMsg := 'ERROR on ResetNode REF: '+ lastCAN_ErrMsg;
                  JobParams.retCode := resERROR;
                  inc(errCount);
                  LogStep(501,501, 501,  JobParams.retMsg );
                  exit;
              end;
//            LogStep(502,502, 502, 'OK REF node Stopped.');
              LogStep(503,503, 503, 'DUT ResetNode SENT Successfully to Node:'+ edtDutAddr.text );

              // Stop DUT cyclic, ma con Reset perchè stop non permette successive letture dirette in mV !
              // Verificare che non si corrompa il setup risoluzione in centesimi, necessario per calcoli seguenti !
              if not CAN_SendNMT( CHdut, cmdResetNode, edtDutAddr.text ) then begin
                  JobParams.retMsg := 'DUT ResetNode '+edtDutAddr.text+' ERROR: '+ lastCAN_ErrMsg;
                  JobParams.retCode := resERROR;
                  inc(errCount);
                  LogStep(504,504, 504, JobParams.retMsg);
                  exit;
              end;
              LogStep(505,505, 505, 'DUT ResetNode SENT Successfully to Node:'+ edtDutAddr.text );
              dmBase.Wait(25);   // lascia digerire i comandi ai nodes.
              LogStep(506,506, 506, 'ALL Cyclic transmission stopped Successfully');

              // e svuota da eventuali residui
              LogStep(507,507, 507, 'FLUSH queue All Can CHANNELS...');
              if not CAN_FlushRxBuffer(CHref) then begin    // prepara rx queue clean
                  LogStep(508,508, 508,  'ERROR on CHref Flush: '+lastCAN_ErrMsg );
                  inc(errCount);
                  exit;
              end;
              if CHdut.Handle <> CHref.Handle then
                  if not CAN_FlushRxBuffer(CHdut) then begin
                      LogStep(509,509, 509,  'ERROR on CHdut Flush: '+lastCAN_ErrMsg );
                      inc(errCount);
                      exit;
                  end;
              LogStep(510,510, 510, 'OK FLUSH All CanCH Successful.');    // preparato rx queue clean.

              // OK, Ref è su Angolo di TEST richiesto, e angolo CROSS a 0°
              // entrambi con tolleranza larga (±2,0deg).
              LogText(#9#9+ format('    REF_X = %6.2f    REF_Y = %6.2f   > preliminare', [currXRef, currYRef]));

              // avvio subito acquisizione angoli DUT !
              // leggo Almeno 5 valori DUT non ravvicinati, a 300 msec.
              // il filtro massimo su DUT può impiegare anche *2 sec* prima di stabilizzare i valori.
              // Quindi attendo almeno 1,5 sec. per permettere al DUT di stabilizzare output.
              dmBase.Wait(1500);

              // Potrei anche restare in Read finchè i valori non sono sempre uguali, ma serve gestione del timeout...
              // quindi più facile e veloce far ripetere lo Step, visto che il DUT resta in posizione.
              LogStep(511,511, 511, 'Acquisizione campioni da DUT per validazione...');
              iDeg := 0;                // meglio <> da  range [-360..+360] ?
              prevRead := 0;
              DutStable := False;
              for i := 1 to kDutSamplesStable do begin  // DutStable in 1,5 sec fa 5 reads.

                  // legge Angolo X da DUT as INT16. (6010-0)
                  dmBase.Wait(50); // msec. attesa breve prima di read angle.
                  if not CAN_Read_ANGLE( edtDUTaddr.text, 'X', iDeg) then begin // Delay standard è 25 msec
                      JobParams.retMsg := 'DUT Read X-angle ERROR: '+ lastCAN_ErrMsg;
                      JobParams.retCode := resERROR;
                      inc(errCount);
                      LogStep(512,512, 512,  JobParams.retMsg );
                      break;
                  end;
                  // The currXDut was Read successfully !
                  currXDut := iDeg / DUT_resolution;            // applico risoluzione del Dut.
                  // NB: devono essere dello stesso tipo, smallint o integer entrambi !!!

                  // inversioni angolo DUT, se richieste da opzione popmenù gauges
                  if vInvertX_dut then
                      currXDut := currXDut * -1;                // Necessario per DUT invertire solo X !
                  edtDutX.text := format('%6.2f', [currXDut]);  // currXDut.ToString;

                  // legge Angolo Y da DUT as INT16. (6020-0)
                  dmBase.Wait(50); // msec. attesa breve prima di read angle.
                  if not CAN_Read_ANGLE( edtDUTaddr.text, 'Y', iDeg) then begin // Delay standard è 25 msec
                      JobParams.retMsg := 'DUT Read Y-angle ERROR: '+ lastCAN_ErrMsg;
                      JobParams.retCode := resERROR;
                      inc(errCount);
                      LogStep(513,513, 513,  JobParams.retMsg );
                      break;
                  end;
                  // The currYDut was Read successfully !
                  currYDut := iDeg / DUT_resolution;             // applico risoluzione del Dut.
                  // NB: devono essere dello stesso tipo, smallint o integer entrambi !!!

                  // inversioni angolo DUT, se richieste da opzione popmenù gauges
                  if vInvertY_dut then
                      currYDut := currYDut * -1;                 // Necessario per DUT invertire solo Y !
                  edtDutY.text := format('%6.2f', [currYDut]);   // currYDut.ToString;

                  LogText(#9#9+ format('[%d] DUT_X = %6.2f    DUT_Y = %6.2f', [i, currXDut, currYDut]));

                  // a seconda dell'asse in TEST, verifica che sia stabile entro 0,1° scarto fra letture successive
                  if testAxis = 'X' then begin
                      // ultime due letture su asse-X sono stabili entro 0.1 ?
                      if Abs(currXDut - prevRead) <= kDutSamplesScarto then begin
                          DutStable := True;
                          break;                    // si, procedo perchè entro tolleranza accettabile 0.1
                      end;
                      prevRead := currXDut;         // no, salva nuovo previous value e resta in read loop.
                  end
                  else begin
                      // ultime due letture su asse-Y sono stabili entro 0.1 ?
                      if Abs(currYDut - prevRead) <= kDutSamplesScarto then begin
                          DutStable := True;
                          break;                    // si, procedo perchè entro tolleranza accettabile 0.1
                      end;
                      prevRead := currYDut;         // no, salva nuovo previous value e resta in read loop.
                  end;

                  // 300 msec. delay e ignoro i 25+50 ms già presenti nelle CAN_ReadDUT().
                  dmBase.Wait( kDutSamplesInterval );

              end;//for 6 reads stable
              if not DutStable then begin
                  JobParams.retMsg := 'WARNING: Letture dal DUT danno un Angolo INSTABILE !';
                  JobParams.retCode := resWARNING;
                  inc(errCount);
                  LogStep(514,514, 514,  JobParams.retMsg );
                  break;
              end;
              LogText(format('OK letture Stabili DUT_X = %6.2f    DUT_Y = %6.2f', [currXDut, currYDut]));
//*
              //// Dopo vari Delay è necessario Aggiornare le letture dal Reference !
              // legge Angolo X da REF as INT16. (6010-0)
              if not CAN_Read_ANGLE( edtREF2addr.text, 'X', iDeg) then begin // Delay standard è 25 msec
                  JobParams.retMsg := 'REF2 Read X-angle ERROR: '+ lastCAN_ErrMsg;
                  JobParams.retCode := resERROR;
                  inc(errCount);
                  LogStep(515,515, 515,  JobParams.retMsg );
                  exit;
              end;
              // The currXRef was Read successfully !
              currXRef := iDeg / REF2_resolution;           // applico risoluzione del Ref.
              // NB: devono essere dello stesso tipo, smallint o integer entrambi !!!
              // inversioni angolo REF, se richieste da opzione popmenù gauges
              if vInvertX_ref then
                  currXRef := currXRef * -1;                // Necessario per REF invertire solo X !
              edtRefX.text := format('%6.2f', [currXRef]);  // currXRef.ToString;

              // legge Angolo Y da REF as INT16. (6020-0)
              if not CAN_Read_ANGLE( edtREF2addr.text, 'Y', iDeg) then begin // Delay standard è 25 msec
                  JobParams.retMsg := 'REF2 Read Y-angle ERROR: '+ lastCAN_ErrMsg;
                  JobParams.retCode := resERROR;
                  inc(errCount);
                  LogStep(516,516, 516,  JobParams.retMsg );
                  exit;
              end;
              // The currYRef was Read successfully !
              currYRef := iDeg / REF2_resolution;           // applico risoluzione del Ref.
              // NB: devono essere dello stesso tipo, smallint o integer entrambi !!!
              // inversioni angolo REF, se richieste da opzione popmenù gauges
              if vInvertY_ref then
                  currYRef := currYRef * -1;                 // Necessario per REF invertire solo Y !
              edtRefY.text := format('%6.2f', [currYRef]);   // currYRef.ToString;

              // Aggiorno letture REF Angolo di TEST richiesto, e angolo CROSS a 0°
              // entrambi con tolleranza larga (±2,0deg).
              LogText(#9#9+ format('    REF_X = %6.2f    REF_Y = %6.2f   > aggiornato', [currXRef, currYRef]));
//
              // Acquisiti sia Ref che Dut, assegno valori a Test e Cross a seconda dell'asse in TEST.
              if testAxis = 'X' then begin
                  // currXRef perchè qui asse di TEST = X
                  currRef_test  := currXRef;
                  currRef_cross := currYRef;
                  currDut_test  := currXDut;
                  currDut_cross := currYDut;
              end
              else begin
                  // currYRef perchè qui asse di TEST = Y
                  currRef_test  := currYRef;
                  currRef_cross := currXRef;
                  currDut_test  := currYDut;
                  currDut_cross := currXDut;
              end;
              // Ora individuo range P1,P0 e relativi valori di caratterizzazione banco, sia per asse TEST che CROSS,
              // per calcolare il coefficiente  di correzione dell'angolo REF.

              // loop su array CorrezioneBanco per trovare range P0, P1 e prevAngle per asse TEST.
              prevAngleTest := 1000;    // per evidenziare errore nel caso non venga definito.
              LogText('Correzione Banco per Asse in TEST: '+ testAxis );
              for i := 1 to length(arrCorrezioneBanco) do begin
                  // scorre i valori in array (partendo da +70 in giù) fino a quando trova il primo
                  // che risulta inferiore all'angolo REF misurato.
                  iDeg := arrCorrezioneBanco[i].angolo;
                  // il fatto che l'asse in TEST sia entro +70-70 deg è già verificato da sezione precedente
                  // che lo fa passare solo se entro le tolleranze previste.

                  // a seconda dell'asse in TEST
                  if currRef_test < iDeg then            // currRef_test può essere X o Y
                      continue;

                  // qui prosegue quando angolo currRef_test >= valore in arrCorrezioneBanco[]
                  prevAngleTest := iDeg;                 // salva per uso successivo.
                  TestP0 := arrCorrezioneBanco[i].asseTest;
                  TestP1 := arrCorrezioneBanco[i -1].asseTest;
                  // in LoadCarattBanco_fromXLS() i campi Test e Cross sono caricati in arrCorrezioneBanco[]
                  // a seconda di quale è l'asse in Test corrente !
                  break;
              end;
              LogText(#9#9+ format('identificato intervallo angoli P1:%d    P0:%d', [arrCorrezioneBanco[i-1].angolo, arrCorrezioneBanco[i].angolo])); // anche se var 'i' non sarebbe da usare fuori dal 'for'...
              LogText(#9#9+ format('TEST  P0 =%12.9f     P1 =%12.9f', [TestP0, TestP1]));

              // Secondo loop su array CorrezioneBanco per trovare range P0, P1 e prevAngle per asse CROSS.
              prevAngleCross := 1000;    // per evidenziare errore nel caso non venga definito.
              LogText('Correzione Banco per Asse CROSS');
              for i := 1 to length(arrCorrezioneBanco) do begin
                  // scorre i valori in array (partendo da +70 in giù) fino a quando trova il primo
                  // che risulta inferiore all'angolo REF misurato.
                  iDeg := arrCorrezioneBanco[i].angolo;
                  // il fatto che l'asse CROSS sia entro +70-70 deg è già verificato da sezione precedente
                  // che lo fa passare solo se entro le tolleranze previste.

                  // a seconda dell'asse CROSS
                  if currRef_cross < iDeg then           // currRef_cross può essere X o Y
                      continue;

                  // qui prosegue quando angolo currRef_test >= valore in arrCorrezioneBanco[]
                  prevAngleCross := iDeg;                // salva per uso successivo.
                  // in LoadCarattBanco_fromXLS() i campi Test e Cross sono caricati in arrCorrezioneBanco[]
                  // a seconda di quale è l'asse in Test corrente !
                  CrossP0 := arrCorrezioneBanco[i].asseCross;
                  CrossP1 := arrCorrezioneBanco[i -1].asseCross;
                  break;
              end;
              LogText(#9#9+ format('identificato intervallo angoli P1:%d    P0:%d', [arrCorrezioneBanco[i-1].angolo, arrCorrezioneBanco[i].angolo])); // anche se var 'i' non sarebbe da usare fuori dal 'for'...
              LogText(#9#9+ format('CROSS P0 =%12.9f     P1 =%12.9f', [CrossP0, CrossP1]));

              // calcolo fattore correzione angolo dell'asse in Test
              // Ricavo pendenza del Range identificato
              pendTest := (TestP1 - TestP0) / 10;
              pendCross := (CrossP1 - CrossP0) / 10;
              LogText(#9#9+ format('Pendenza range   pen_TEST =%12.9f     pen_CROSS =%12.9f', [pendTest, pendCross]));

              fattCORR_test := TestP0 + pendTest * ( currRef_test - prevAngleTest );
              fattCORR_cross := CrossP0 + pendCross * ( currRef_cross - prevAngleCross );
              LogText(#9#9+ format('Fattore Correzione corr_TEST =%12.9f     corr_CROSS =%12.9f', [fattCORR_test, fattCORR_cross]));

              RefCorretto_test := currRef_test - fattCORR_test;
              RefCorretto_cross := currRef_cross - fattCORR_cross;
              LogText(#9#9+ format('REF Corretto   refc_TEST =%12.9f    refc_CROSS =%12.9f', [RefCorretto_test, RefCorretto_cross]));

              // Ho già acquisito currXDut, currYDut con letture stabili ad intervalli ampi...  5 x 300 msec.
              // quindi posso fare la differenza
              DutERR_test  := currDut_test - RefCorretto_test;
              DutERR_cross := currDut_cross - RefCorretto_cross;
              LogText(#9#9+ format('DUT Errore   err_TEST =%12.9f    err_CROSS =%12.9f', [DutERR_test, DutERR_cross]));

              if Abs(DutERR_cross) > DutList[currDutIdx].ErrMaxCross then begin    // Abs() perchè errore ammesso è +/-
                  JobParams.retMsg := 'DUT Read CROSS angle OUT of TOLERANCE: '+ DutERR_cross.ToString;
                  JobParams.retCode := resERROR;
                  inc(errCount);
                  LogStep(517,517, 517,  JobParams.retMsg );
              end;
              if Abs(DutERR_test) > DutList[currDutIdx].ErrMaxTest then begin      // Abs() perchè errore ammesso è +/-
                  JobParams.retMsg := 'DUT Read TEST angle-'+testAxis +' OUT of TOLERANCE: '+ DutERR_test.ToString;
                  JobParams.retCode := resERROR;
                  inc(errCount);
                  LogStep(518,518, 518,  JobParams.retMsg );
              end;
              break;      // infine esce sempre !!!

          end;//// endless loop

          // qui devo bloccare in caso di errore, perchè indica validazione incompleta !
          if errCount > 0 then begin
              edtDutX.color := clRed;
              edtDutY.color := clRed;
              //JobParams.retMsg := _Errore_;     // già valorizzato !
              //JobParams.retCode := resERROR;    // problemi di acquisizione DUT, REF o validazione sono di tipo bloccante.
              Winapi.Windows.Beep( 500, 700) ;    // Avviso user dell'errore (frequency, duration)
              result := False;
          end
          else begin
              edtDutX.color := clLime;
              edtDutY.color := clLime;
              JobParams.retMsg := _Completato_;   // Ok Asse/angolo validato.
              JobParams.retCode := resOK;
              Winapi.Windows.Beep( 2000, 400) ;   // Avviso breve user che può proseguire (frequency, duration)
              result := True;
          end;

          // In tutti i casi compilo (o sovrascrivo) il record con i valori rilevati.
          TestRecord := TSessionTests.create;
          with TestRecord do begin
            PK              := -1;                          // solo preset, sarà poi assegnato da DB.
            SessionID       := -1;                          // 1-N con Sessione, sarà poi assegnato da DB.
            Lotto           := Sessione.Lotto;              // necessario perchè s/n non è univoco!
            Serial          := Sessione.Serial;             // Relazione 1-N con il record SessionData e DutData.
            AsseCorrente    := 'TEST2A'+ testAxis;          // per distinguere tipo di record misura su DB.
            degPairID       := strTestAngle[ord(testDeg)];  // Mnemonico della combinazione Angolo in test TTestAngle.
            TargetX         := kTargetRefX;                 // deg. posizionamento richiesto
            TargetY         := kTargetRefY;
            ScartoAmmesso   := scartoMax;                   // deg. fra target e Ref ottenuto
            RefSamplesInRange  := vRefSamplesInRange;            // num. per considerare Ref valido
            IntervalloRefReads := strToInt(kTimer4cyclicalTxd);  // msec. fra letture consecutive del Reference; è il suo intervallo configurato di autoTX.
            DutSamplesStable   := kDutSamplesStable;             // per considerare Dut valido
            IntervalloDutReads := kDutSamplesInterval;           // msec. fra letture consecutive del DUT
            RefAx        := currXRef;       // posizionamento ottenuto del Ref in deg
            RefAy        := currYRef;       // qui non è un valore medio, perchè lo scarto tollerato è ampio, a differenza della calibrazione.
            DutAx_deg    := currXDut;       // letture dut in gradi
            DutAy_deg    := currYDut;
            DutAx_mV     := -1;             // qui in Validazione niente letture dut in tensione.
            DutAy_mV     := -1;
            CorrRefTest  := RefCorretto_test;   // corrected Reference, per X o Y, a seconda quale è AsseTest.
            CorrRefCross := RefCorretto_cross;
            ErroreTest   := DutERR_test;    // per X o Y a seconda quale è AsseTest.
            ErroreCross  := DutERR_cross;   // Y o X opposto dell' AsseTest.
          end;
          // Non serve vedere se il record per questo angolo in test è già presente nel dictionary,
          // perchè Add() lo modifica o lo aggiunge se non c'è già.
          if testAxis = 'X' then
              // TestXRecords perchè qui asse di TEST = X
              TestXRecords.Add( testDeg, TestRecord)
          else
              // TestYRecords perchè qui asse di TEST = Y
              TestYRecords.Add( testDeg, TestRecord);

          // Reset gauges
          gaugeX.Value := 0;
          gaugeY.Value := 0;

        except
          on E: Exception do begin
              JobParams.retCode := resERROR;
              edtRefX.color := clBtnface;
              edtRefY.color := clBtnface;
              LogSysEvent(svEXCEPTION, 1000, 'in '+JobList.Strings[currJob]+' on DUT validation >> '+ E.ClassName +#10#13#9+ ' description: '+E.Message);
          end;
        end;
    finally
        inStepLoop := False;               // va sempre in finally!
        tmrCanRead.enabled := False;       // Se non lo è già.
        currSequence := -1;                // Reset Sequenza blink current Job node.
        tmrFasiBlink.Enabled := False;     // termina blink.
        img1Seq1.Visible := False;         //
        LogStep(519,519, 519, JobList.Strings[currJob]+' ERRORS = '+ intToStr(errCount));
      //TestRecord.free; No, si perde il record di test appena memorizzato!  che serve in report !
        CAN_SendNMT( CHref, cmdStopNode, edtRef2Addr.text );      // blind stop invio ciclico, se non è già a causa Abort.
        // Flush più che altro per lasciare pulizia a steps successivi.
        CAN_FlushRxBuffer(CHref);         // qui meglio no log, perchè non serve e posso perdere
        if CHdut.Handle <> CHref.Handle then           // eventuali altri errMsg di exit via finally !
            CAN_FlushRxBuffer(CHdut);
    end;
end;

function Toiac3Fasi2ax.RunStep_C10(var JobParams: TJobParams): boolean;
begin
    result := ValidazioneDUT( JobParams, deg60n, 'X' );
end;

function Toiac3Fasi2ax.RunStep_C11(var JobParams: TJobParams): boolean;
begin
    result := ValidazioneDUT( JobParams, deg30n, 'X' );
end;

function Toiac3Fasi2ax.RunStep_C12(var JobParams: TJobParams): boolean;
begin
    result := ValidazioneDUT( JobParams, deg15n, 'X' );
end;

function Toiac3Fasi2ax.RunStep_C13(var JobParams: TJobParams): boolean;
begin
    result := ValidazioneDUT( JobParams, deg0, 'X' );
end;

function Toiac3Fasi2ax.RunStep_C14(var JobParams: TJobParams): boolean;
begin
    result := ValidazioneDUT( JobParams, deg15p, 'X' );
end;

function Toiac3Fasi2ax.RunStep_C15(var JobParams: TJobParams): boolean;
begin
    result := ValidazioneDUT( JobParams, deg30p, 'X' );
end;

function Toiac3Fasi2ax.RunStep_C16(var JobParams: TJobParams): boolean;
begin
    result := ValidazioneDUT( JobParams, deg60p, 'X' );
end;


////#######################  Steps Fase D  #######################################################

function Toiac3Fasi2ax.RunStep_D17(var JobParams: TJobParams): boolean;
begin
    result := ValidazioneDUT( JobParams, deg60n, 'Y' );
end;

function Toiac3Fasi2ax.RunStep_D18(var JobParams: TJobParams): boolean;
begin
    result := ValidazioneDUT( JobParams, deg30n, 'Y' );
end;

function Toiac3Fasi2ax.RunStep_D19(var JobParams: TJobParams): boolean;
begin
    result := ValidazioneDUT( JobParams, deg15n, 'Y' );
end;

function Toiac3Fasi2ax.RunStep_D20(var JobParams: TJobParams): boolean;
begin
    result := ValidazioneDUT( JobParams, deg0, 'Y' );
end;

function Toiac3Fasi2ax.RunStep_D21(var JobParams: TJobParams): boolean;
begin
    result := ValidazioneDUT( JobParams, deg15p, 'Y' );
end;

function Toiac3Fasi2ax.RunStep_D22(var JobParams: TJobParams): boolean;
begin
    result := ValidazioneDUT( JobParams, deg30p, 'Y' );
end;

function Toiac3Fasi2ax.RunStep_D23(var JobParams: TJobParams): boolean;
begin
    result := ValidazioneDUT( JobParams, deg60p, 'Y' );
end;


////#######################  Steps Fase E  #######################################################

function Toiac3Fasi2ax.RunStep_E24(var JobParams: TJobParams): boolean;
begin
    //pnlCardsE.ActiveCardIndex := 0;      // card_E24   non dovrebbe servire perchè gestito da mainForm.trvJobListChange()
    result := RunStep_A1( JobParams );     // Porta in Listview tutti i parametri (Default e Calib2Assi) da verificare.
end;

function Toiac3Fasi2ax.RunStep_E25(var JobParams: TJobParams): boolean;
begin
    //pnlCardsE.ActiveCardIndex := 0;      // card_E25   non dovrebbe servire perchè gestito da mainForm.trvJobListChange()
    // qui presuppone i Defaults e Calib 2Assi già caricati in step precedente !
    result := RunStep_A2( JobParams );     // Legge e Verifica tutti i parametri presenti in Listview.
    // se non si trovano parametri errati, non servono i successivi 2 step, che vengono saltati con jobsToSkip !
end;

function Toiac3Fasi2ax.RunStep_E26(var JobParams: TJobParams): boolean;
begin
    //pnlCardsE.ActiveCardIndex := 0;      // card_E26   non dovrebbe servire perchè gestito da mainForm.trvJobListChange()
    result := RunStep_A3( JobParams );
end;

function Toiac3Fasi2ax.RunStep_E27(var JobParams: TJobParams): boolean;
begin
    //pnlCardsE.ActiveCardIndex := 0;      // card_E27   non dovrebbe servire perchè gestito da mainForm.trvJobListChange()
    result := RunStep_A4( JobParams );
end;

////#######################  Steps Fase F  #######################################################

function Toiac3Fasi2ax.RunStep_F28(var JobParams: TJobParams): boolean;
var
  errCount: integer;
begin
    LogStep(520,520, 520, JobList.Strings[currJob]+' - '+ TJobRecord( JobList.Objects[ currJob ]).Description);
    result := False;
    JobParams.retCode := resUNDEFINED;
    JobParams.retMsg := _Incompleto_;

    // Attesa input Note & Commenti.
    mainForm.showUserMessage('Inserire eventuali Note e commenti'+ sLineBreak +'poi procedere cliccando pulsante [OK]');

    btnOkNote.Tag := 0;
    btnOkNote.Enabled := True;           // abilita ora perchè default deve essere disabled !
    mmNoteFine.Enabled := True;          //
    mmNoteFine.Text := 'Nessuna';
    currSequence := 4;                   // n° Sequenza per blink controls.
    tmrFasiBlink.Enabled := True;        // serve alla sequenza che parte subito.
    application.ProcessMessages();
  //btnOkNote.setFocus;
    mmNoteFine.setFocus;

    // Attesa click pulsante btnConfirmEncoderReset
    errCount := 0;
    try
        inStepLoop := True;
        try
          while true do begin // resta in endless loop fino a button premuto, o user abort.
              // check per Cancel ad ogni giro.
              application.processmessages;
              if dmBase.WaitAbort(30) then begin
                  LogText('*** ABORT by USER, interrotto attesa inserimento Note finali !'+#13#10 );
                  exit;
              end;

              // Attende OK button click
              if btnOkNote.Tag > 0 then
                  break;
          end;

          // check esito finale
          if errCount > 0 then begin
              JobParams.retMsg := _Errore_;     // al momento generico
              JobParams.retCode := resERROR;
              result := False;
          end
          else begin
              JobParams.retMsg := _Completato_; // Ok, btn premuto senza altri errori.
              JobParams.retCode := resOK;
              LogStep(521,521, 521, 'OK acquisito campo Note.');
              result := True;
          end;

        except
          on E: Exception do begin
              inc(errCount);
              JobParams.retCode := resERROR;
              LogSysEvent(svEXCEPTION, 2615, 'in '+JobList.Strings[currJob]+' on wait for button btnOkNote >> '+ E.ClassName +#10#13#9+ ' description: '+E.Message);
          end;
        end;
    finally
        Sessione.Memo := mmNoteFine.Text;
        inStepLoop := False;               // va sempre in finally!
        btnOkNote.Tag := 0;                // reset flag.
        btnOkNote.Enabled := False;        // default deve essere disabled !
        mmNoteFine.Enabled := False;       //
        currSequence := -1;                // Reset Sequenza blink current Job node.
        tmrFasiBlink.Enabled := False;     // termina blink.
        img1Seq4.Visible := False;         //
        LogStep(522,522, 522, JobList.Strings[currJob]+' ERRORS = '+ intToStr(errCount));
    end;
end;


function Toiac3Fasi2ax.RunStep_F29(var JobParams: TJobParams): boolean;
var
  tStr: string;
  row, errCount, masterPK: integer;
  calibItem: TPair<TCalibAngle, TSessionCalibra>;
  testItem: TPair<TTestAngle, TSessionTests>;
  aDate: TDateTime;
  cfgListItem: TListItem;

  procedure addTestRecord(recType, masterKey:integer);
  begin
      with dmBase.tabTestData do begin
        Insert;
        (*
            "PK" integer NOT NULL DEFAULT nextval('"Produzione"."DUTDATA_OI20_PK_seq"'::regclass),
            "SESSIONID" integer,
            "SESSIONSTAMP" timestamp without time zone,
            "STARTSTAMP" timestamp without time zone,
            "ENDSTAMP" timestamp without time zone,
            "ASSEINTEST" character varying(8) COLLATE pg_catalog."default",
            "PAIRID" character varying(16) COLLATE pg_catalog."default",
            "SCARTOAMMESSO" real,
            "TARGETX" real,
            "TARGETY" real,
            "NUMREADS" smallint,
            "INTERVALLOREFREADS" smallint,
            "INTERVALLODUTREADS" smallint,
            "REFAX" real,
            "REFAY" real,
            "DUTAXdeg" real,
            "DUTAYdeg" real,
            "DUTAXmv" real,
            "DUTAYmv" real,
            "CORRECTREFTEST" double precision,
            "CORRECTREFCROSS" double precision,
            "ERRORASSETEST" double precision,
            "ERRORASSECROSS" double precision
        *)
        // Parte comune a tutti i record type.
        FieldByName('SESSIONID').AsInteger     := masterKey;                  // va preso da tabSessionDut.PK come master !
      //FieldByName('SERIAL').AsString         := testItem.Value.Serial;      // meglio diretto Sessione.Serial
      //FieldByName('LOTTO').AsString          := testItem.Value.Lotto;       // e Sessione.Lotto !
        FieldByName('SERIAL').AsString         := Sessione.Serial;            //
        FieldByName('LOTTO').AsString          := Sessione.Lotto ;            //
        FieldByName('RECMODE').AsInteger       := Sessione.RecMode;           // 0=Produzione 1=Test    (cbxDemoMode)

        // Parte specifica per il record type corrente.
        case recType of
        1:begin // calibItem: TPair<TCalibAngle, TSessionCalibra>    per gestione record prodotti in fase di Calibrazione.
            FieldByName('ASSEINTEST').AsString     := calibItem.Value.AsseCorrente;// X o Y !
            FieldByName('DEGPAIRID').AsString      := calibItem.Value.degPairID;   // Mnemonico della combinazione Angolo in Calibrazione TCalibAngle.
            FieldByName('TARGETX').AsSingle        := calibItem.Value.TargetX;
            FieldByName('TARGETY').AsSingle        := calibItem.Value.TargetY;
            FieldByName('SCARTOAMMESSO').AsSingle  := calibItem.Value.ScartoAmmesso;
            FieldByName('REFSAMPLESINRANGE').AsInteger  := calibItem.Value.RefSamplesInRange;    // toobig
            FieldByName('INTERVALLOREFREADS').AsInteger := calibItem.Value.IntervalloRefReads;   //
            FieldByName('INTERVALLODUTREADS').AsInteger := calibItem.Value.IntervalloDutReads;   //
            FieldByName('REFAX').AsSingle           := calibItem.Value.RefAx;
            FieldByName('REFAY').AsSingle           := calibItem.Value.RefAy;
            FieldByName('DUTAXmv').AsSingle         := calibItem.Value.DutAx_mV;                  // mV già settati a -1
            FieldByName('DUTAYmv').AsSingle         := calibItem.Value.DutAy_mV;                  // non vuoto.
            FieldByName('AliasX').AsString          := calibItem.Value.AliasX;
            FieldByName('AliasY').AsString          := calibItem.Value.AliasY;
          end;
        2:begin // testItem: TPair<TTestAngle, TSessionTests>  per gestione record prodotti in fase di test X e Y.
            FieldByName('ASSEINTEST').AsString     := testItem.Value.AsseCorrente;// X o Y !
            FieldByName('DEGPAIRID').AsString      := testItem.Value.degPairID;   // Mnemonico della combinazione Angolo in test TTestAngle.
            FieldByName('TARGETX').AsSingle        := testItem.Value.TargetX;
            FieldByName('TARGETY').AsSingle        := testItem.Value.TargetY;
            FieldByName('SCARTOAMMESSO').AsSingle  := testItem.Value.ScartoAmmesso;
            FieldByName('REFSAMPLESINRANGE').AsInteger  := testItem.Value.RefSamplesInRange;    // toobig
            FieldByName('INTERVALLOREFREADS').AsInteger := testItem.Value.IntervalloRefReads;   //
            FieldByName('DUTSAMPLESSTABLE').AsInteger   := testItem.Value.DutSamplesStable;     //
            FieldByName('INTERVALLODUTREADS').AsInteger := testItem.Value.IntervalloDutReads;   //
            FieldByName('REFAX').AsSingle           := testItem.Value.RefAx;
            FieldByName('REFAY').AsSingle           := testItem.Value.RefAy;
            FieldByName('DUTAXdeg').AsSingle        := testItem.Value.DutAx_deg;
            FieldByName('DUTAYdeg').AsSingle        := testItem.Value.DutAy_deg;
            FieldByName('DUTAXmv').AsSingle         := testItem.Value.DutAx_mV;                  // mV già settati a -1
            FieldByName('DUTAYmv').AsSingle         := testItem.Value.DutAy_mV;                  // non vuoto.
            FieldByName('CORRECTREFTEST').AsSingle  := testItem.Value.CorrRefTest;
            FieldByName('CORRECTREFCROSS').AsSingle := testItem.Value.CorrRefCross;
            FieldByName('ERRORASSETEST').AsSingle   := testItem.Value.ErroreTest;
            FieldByName('ERRORASSECROSS').AsSingle  := testItem.Value.ErroreCross;
          end;
        end;//case
        Post;
      end;
  end;

begin
    //mainForm.mmLog.lines.add('RunStep_F29...');
    //pnlCardsF.ActiveCardIndex := 0;      //card_F28  non dovrebbe servire perchè gestito da mainForm.trvJobListChange()
    result := False;
    errCount := 0;
    // Anche save su DB può essere Opzionale usando cbx "Force Pass" !
    // Verifica preliminare connessione a DB.
    if not dmBase.pgConn.Connected then begin
        LogStep(523,523, 523, 'ERRORE Database '+ dmBase.pgConn.Server +':'+ dmBase.pgConn.Database+' NON Connesso !');
        JobParams.retMsg := 'ERROR Database NOT Connected.';
        JobParams.retCode := resERROR;   // per ora bloccante.
        exit;
    end;
    try (*
        default isolation levels.
        in FireDAC the property value is xiUnspecified:
            DB2                     - xiReadCommitted
            InterBase and Firebird  - xiSnapshot
            MySQL and MariaDB       - xiRepeatableRead
            Oracle                  - xiReadCommitted
            Microsoft SQL Server    - xiReadCommitted
            SQLite                  - xiSerializible
            PostgreSQL              - xiReadCommitted
        http://docs.embarcadero.com/products/rad_studio/firedac/frames.html?frmname=topic&frmfile=Managing_Transactions.html
        *)
        // Prima di tutto verifica e avvisa se esistono già Test Records per questo S/N...
        dmBase.pgConn.StartTransaction;  // gestione transazione usando try-except !
        try
            LogStep(524,524, 524, 'Salvataggio dati test di collaudo per S/N: '+ Sessione.Serial+'  su DB Produzione.' );
            // Possono essere eseguite più Sessioni di test per uno stesso S/N, anche in momenti diversi
            // I dati su DB non si sovrappongono perchè i records Detail, come indice secondario hanno Session ID diversi.
            // Ma c'è overwrite dei Report che passano a revisione successiva !
            with dmBase.qryFindSession do begin
                Close;
                // Cerco se esiste almeno una sessione di collaudo per questo [CodProd + S/N + Lotto + RecMode]
                ParamByName('pCODPROD').Value := Sessione.CodiceProdotto;    // key = prod+lotto+serial  poi versioni con in + RECORDSTAMP
                ParamByName('pLOTTO').Value   := Sessione.Lotto;
                ParamByName('pSERIAL').Value  := Sessione.Serial;
                ParamByName('pRECMODE').Value := vRecMode;          // 0/1
                LogStep(525,525, 525, 'Execute Query...');
                Open;
                if RecordCount >= 1 then begin
                    // necessaria  qryFindSession.Options.QueryRecCount = True !
                    aDate := FieldByName('RecordStamp').AsDateTime;
                    // User può decidere di aggiungere sessioni di test per stesso device
                    MessageDlg('ATTENZIONE:' +sLineBreak+
                               'Esistono già dati di Test per questo S/N: '+ Sessione.Serial +sLineBreak+
                               'con data '+DateTimeToStr(aDate), mtWarning, [mbOk], 0, mbOk);
                    LogSysEvent(svWARNING, 2805, 'Test data already exists for S/N: '+ Sessione.Serial );
                    // Non bloccante perchè non c'è perdita di dati, ma semmai solo storico  nel report.
                end;
            end;
            LogStep(526,526, 526, 'Query check Done.');

            // insert Master Sessione
            // NB: crea 1 record Sessione per ogni Device collaudato, anche se ha stesso S/N.
            with dmBase.tabSessionDut do begin
              (*
              "PK"                integer NOT NULL DEFAULT nextval('"Produzione"."DUTSESSIONS_PK_seq"'::regclass),
              "WORKPLACE"         character varying(20) COLLATE pg_catalog."default",
              "IDOPERATORE"       character varying(28) COLLATE pg_catalog."default"
              "STARTSTAMP"        timestamp without time zone,
              "ENDSTAMP"          timestamp without time zone,
              "WORKTIME"          time without time zone,
              "PAUSETIME"         double precision,
              "MEMO"              text COLLATE pg_catalog."default",
              *)
              CheckBrowseMode;
              Insert;
              LogStep(527,527, 527, 'Insert master tabSessionDut...');
              FieldByName('DEVINFOPK').AsInteger     := Sessione.DevInfoPK;    // Al momento non disponbile. Relazione 1-1 per collegare ai dati Device
              FieldByName('RECORDSTAMP').AsDateTime  := Sessione.RecordStamp;
              FieldByName('WORKPLACE').AsString      := Sessione.PostazioneID;
              FieldByName('IDOPERATORE').AsString    := Sessione.OperatoreID;
              FieldByName('COMMESSA').AsString       := Sessione.Commessa;
              FieldByName('CODICEPRODOTTO').AsString := Sessione.CodiceProdotto;
              FieldByName('LOTTO').AsString          := Sessione.Lotto;
              FieldByName('SERIAL').AsString         := Sessione.Serial;
              //
              FieldByName('VERSIONFW').AsString    := Sessione.VersionFw;    // valori rilevati dal Device in lavorazione,
              FieldByName('VERSIONHW').AsString    := Sessione.VersionHw;    // utili anche per valutare il risultato dei test.
              FieldByName('VERSIONPCB').AsString   := Sessione.VersionPcb;
              //
              FieldByName('STARTSTAMP').AsDateTime := Sessione.StartStamp;   // inizio collaudo
              Sessione.EndStamp := incMinute(now, 1);                        // preliminare previsto, per poter stampare report. Va poi aggiornato a reale fine Sessione!
              FieldByName('ENDSTAMP').AsDateTime   := Sessione.EndStamp;     // fine collaudo
              Sessione.WorkTime := Sessione.EndStamp - Sessione.StartStamp;  // preliminare previsto, per poter stampare report. Va poi aggiornato a reale fine Sessione!
              FieldByName('WORKTIME').AsFloat      := Sessione.WorkTime;     // tempo impiegato   AsSQLTimeStampOffset;
              FieldByName('PAUSETIME').AsSingle    := Sessione.Pausetime;    // tempo rimasta in pausa by user. hh:mm:ss
              FieldByName('MEMO').AsString         := Sessione.Memo;         // commenti e atttributi non codificati.
              FieldByName('RECMODE').AsInteger     := Sessione.RecMode;      // 0=Produzione 1=Test    (cbxDemoMode)
              Post;
            //masterPK := dmBase.gammaConnection.GetLastAutoGenValue;        // solo in Firedac :-(
            //masterPK := dmBase.tabSessionDut.LastInsertId;                 // solo in UniDac e non va :-(
              masterPK := FieldByName('PK').AsInteger;   // con Options.DefaultValues := true  mi dà la PK assegnata dal Server !!! :-)
            end;                                         // che posso usare in tabDetail per collegare la Master !

            // insert Details Test data
            dmBase.tabTestData.CheckBrowseMode;
            LogStep(528,528, 528, 'Insert detail Test Records...');
            for calibItem in CalibRecords do      // record diverso per calibrazioni.
                addTestRecord( 1, masterPK );
            LogStep(529,529, 529, 'CalibrationRecords OK.');
            for testItem in TestXRecords do     // record per i test X.
                addTestRecord( 2, masterPK );   // masterPK per collegare i record detail al master tabSessionDut.
            LogStep(530,530, 530, 'TestXRecords OK.');
            for testItem in TestYRecords do     // record per i test Y.
                addTestRecord( 2, masterPK );
            LogStep(531,531, 531, 'TestYRecords OK.');

            // insert Details Params data, composti da listview già verificata
            // con aggiunti parametri di calibrazione che sono sempre dei CAN index.
            dmBase.tabDutParams.CheckBrowseMode;
            LogStep(532,532, 532, 'Insert detail tabDutParams Records...');
            // Carica Parametri da lista già verificata
            for row := 0 to lviewCFG.Items.Count-1 do begin
                // copia parametri previsti di default
                cfgListItem := lviewCFG.Items[ row ];    // intera riga della listview
                with dmBase.tabDutParams do begin
                  Insert;
                  (*
                      "PK" integer NOT NULL,
                      "SERIAL" character varying(20)[] COLLATE pg_catalog."default",
                      "SESSIONID" integer,
                      "INDEX" character varying(16)[] COLLATE pg_catalog."default",
                      "SUBIDX" character varying(8)[] COLLATE pg_catalog."default",
                      "ALIAS" character varying(40) COLLATE pg_catalog."default",
                      "TYPE" character varying(8)[] COLLATE pg_catalog."default",
                      "VALUE" character varying(16)[] COLLATE pg_catalog."default",
                  *)
                  FieldByName('SERIAL').AsString     := Sessione.Serial;
                  FieldByName('LOTTO').AsString      := Sessione.Lotto;
                  FieldByName('SESSIONID').AsInteger := masterPK;                // va preso da tabSessionDut.PK
                  FieldByName('INDEX').AsString      := cfgListItem.Caption;     // Index
                  FieldByName('SUBIDX').AsString     := cfgListItem.SubItems[0]; // SubIdx
                  FieldByName('ALIAS').AsString      := cfgListItem.SubItems[1]; // Alias
                  FieldByName('TYPE').AsString       := cfgListItem.SubItems[3]; // DataType
                  FieldByName('VALUE').AsString      := cfgListItem.SubItems[4]; // Read value
                  Post;
                end;
            end;//for
            LogStep(533,533, 533, 'tabDutParams OK.');

            dmBase.pgConn.Commit;

            // qui errore non sarebbe bloccante perchè è solo problema di store dati su DB, mentre il collaudo DUT si presume ok.
            if errCount = 0 then begin
                LogStep(534,534, 534, 'OK dati di collaudo salvati.');
                JobParams.retMsg := _Completato_;   // Ok save su DB completato.
                JobParams.retCode := resOK;
                result := True;
            end;

        except
          on E: Exception do begin
              dmBase.pgConn.Rollback;
              tStr := ' in RunStep_28 onSave test records >> '+ E.ClassName +#10#13#9+ ' description: '+E.Message;
              JobParams.retCode := resERROR;
              JobParams.retMsg  := _Exception_ + tStr;
              //LogStep(535,535, 535,  tStr );
              LogSysEvent(svEXCEPTION, 2830, tStr);
              // visto che ho anche il report XLS, problemi di salvataggio records non sarebbero di tipo bloccante...
          end;

        end;
    finally
        LogStep(536,536, 536, JobList.Strings[currJob]+' ERRORS = '+ intToStr(errCount));
        if dmBase.pgConn.InTransaction then     // se trova pendenze impreviste
            dmBase.pgConn.Rollback;             // meglio cancel !
        dmBase.wait(500);// anche solo per evitare click multipli involontari da parte user.
    end;
end;


function Toiac3Fasi2ax.RunStep_F30(var JobParams: TJobParams): boolean;
var
  tStr, xlsFile, rptFile: string;
  errCount: integer;
  testItem: TPair<TTestAngle, TSessionTests>;
  calItem: TPair<TCalibAngle, TSessionCalibra>;
  i, rowOffset: integer;
  ListItem: TListItem;
  aJobRecord: TJobRecord;
  tBeg, tEnd: TDateTime;
  iRev: integer;
  vRev, vReportPath, kXls: string;
  cell: TCellValue;

  function buildFileName: boolean;
  begin
      LogStep(537,537, 537, 'Build nome Report...');
      result := False;
      //rptFile := vReportPath +'\'+ Sessione.CodiceProdotto +'_'+ Sessione.Lotto +'\'+ Sessione.CodiceProdotto +'_EOL_'+ Sessione.Lotto + Sessione.Serial +'.xls';
      // Forzo destinazione per report di prova.
      if vRecMode = 1 then begin
          kXls := '_demo.xls';
          vReportPath := kAppPath + 'Reports\';        // override vReportPath to C: Local
      end
      else begin
          kXls := '.xls';
          vReportPath := iniReportPath;                // vReportPath as from .INI
      end;
      // permetto override destinazione report, ma
      // se Demo mode, forzo report locali !
      if vRecMode = 1 then
          rgXlsDir.ItemIndex := 0;
      if rgXlsDir.ItemIndex = 0 then
          vReportPath := kAppPath + 'Reports\'         // override vReportPath to C: Local
      else
          vReportPath := iniReportPath;                // vReportPath as from .INI  (solitamente Netserver)
      // Verifica accessibilità cartella principale su server:  vReportPath from .INI
      if DirectoryExists( vReportPath ) then begin     // Server disponibile e report di produzione
          // Server disponibile, salva report in rete.
          // Verifica esistenza cartella del Lotto, se no la crea.
          rptFile := vReportPath + Sessione.CodiceProdotto +'_L'+ Sessione.Lotto;
          if not DirectoryExists( rptFile ) then begin
              // crea nuova cartella per Lotto.
              if not CreateDir( rptFile ) then begin   // bloccante e l'utente potrà scegliere da GUI se salvare in locale.
                  JobParams.retMsg := 'ERROR impossibile creare il path: '+ rptFile;
                  JobParams.retCode := resERROR;
                  exit;
              end;
          end;
          // qui cartella del Lotto creata o già esistente.
          rptFile := vReportPath + Sessione.CodiceProdotto +'_L'+ Sessione.Lotto +'\'+ Sessione.CodiceProdotto +'_EOL_';
          rptFile := rptFile + Sessione.Lotto +'#'+ Sessione.Serial + kXls;
          result := True;
      end
      else begin
          // Server non disponibile o report di prova, salva sempre in locale.
          MessageDlg('ATTENZIONE:' +sLineBreak+ 'il Path  '+ vReportPath +sLineBreak+
                     'Non è raggiungibile !' +sLineBreak+ 'Verifica e riprova.', mtWarning, [mbOk], 0, mbOk);
          JobParams.retMsg := 'ERROR path Not Found: '+ vReportPath;
          JobParams.retCode := resERROR;
      end;
  end;//buildFileName()

begin
    //mainForm.mmLog.lines.add('RunStep_F29...');
    //pnlCardsF.ActiveCardIndex := 1;          //card_F29  non dovrebbe servire perchè gestito da mainForm.trvJobListChange()
    result := False;
    errCount := 0;
    Screen.Cursor := crHourglass;
    dmBase.wait(50);              // per attivare Screen.Cursor ?
    xlsReport := TXlsFile.Create(true);
    try
        try
            LogStep(538,538, 538, 'Generazione Report EOL per S/N: '+ Sessione.Serial );
            // Se siamo in versione FULL (2+1 assi), questo report 1Asse và aggiunto al file pre-esistente del precedente EOL a 2Assi !
{
            Sistema che prevede overwrite delle rev. precedenti !
            Build Nome-File previsto per questo SN.
            Verifico se Nome-File Report già esistente.  (solo 2ax precedente, oppure 2ax e 1ax precedente)
                si, Apro Nome-File esistente
                    posiziono su sheet "1 ASSE"
                    Read revision n° (x)
                    compila report
                    inc(rev n°)      (x+1)
                    Salva su Nome-File  (stesso file)

                no, Apro file Modello Full               (primo report 1ax)
                    Verifico se modello FULL
                        no, elimina sheet 2assi.
                    posiziono su sheet "1 ASSE"
                    Read revision n° (0)
                    compila report
                    inc(rev n°)
                    Salva su Nome-File  (file name NEW)
 }
            if not buildFileName() then begin       // genera file Name rptFile
                LogStep(539,539, 539, JobParams.retMsg);
                inc(errCount);
                exit;                               // passa sempre per finally
            end;
            //rptFile := vReportPath +'\'+ Sessione.CodiceProdotto +'_'+ Sessione.Lotto +'\'+ Sessione.CodiceProdotto +'_EOL_'+ Sessione.Lotto + Sessione.Serial +'.xls';

            // Verifico se Nome-File Report già esistente.  (solo 2ax precedente, oppure 2ax e 1ax precedente)
            if dmBase.FindFile( rptFile ) then begin
                // si, file esiste già, Apro file esistente e posiziono su sheet "2 ASSI".
                xlsReport.Open( rptFile );          // Open an existing xls file for modification.
            end
            else begin
                // no, Apro file Modello Full               (primo report 1ax)
                // previa verifica presenza file Base-Report
                xlsFile := kAppPath + 'DataSource\BaseReportALL.xlsx';
                if dmBase.FindFile( xlsFile ) then
                    LogSysEvent(svINFO, 2900, 'Found Report xls model: '+ xlsFile)     // log in queue, fino al primo Dequeue!
                else begin
                    LogSysEvent(svERROR, 2903, 'Report model NOT found: '+ xlsFile);
                    exit;
                end;
                xlsReport.Open( xlsFile );          // Open New Model file.

                // Verifico se modello FULL
                //     no, elimina sheet 1Asse.
                if not AnsiContainsText(DutList[currDutIdx].ModoMisura, 'full') then begin
                    xlsReport.DeleteSheet('1 ASSE');          // no, Elimina sheet non necessario.
                    LogStep(540,540, 540, 'INFO: Deleted 1 Axes unused sheet.');
                end;
            end;
            xlsReport.ActiveSheetByName := '2 ASSI';  // posiziona su report (sheet) per 2 Assi.
            cell := xlsReport.GetCellValue(10, 9);    // Read revision n°
            if cell.HasValue then begin
                vRev := cell.ToString;
                iRev := strToIntDef(vRev, 0) +1;      // rev.0 passa a rev.1   ecc...
            end
            else
                iRev := 0;                            // rev. vuoto è primo report quindi resta a rev.0 !
            xlsReport.SetCellValue(10, 9, iRev);      // update report revision.

            LogStep(541,541, 541, 'Print session Master fields...');
            xlsReport.SetCellValue(1, 2, DUT.CodiceProdotto);
            xlsReport.SetCellValue(2, 2, DUT.Alias);
            xlsReport.SetCellValue(3, 2, Sessione.Serial);  // str perchè potrebbe non esser solo numerico
            xlsReport.SetCellValue(4, 2, Sessione.Lotto);
            xlsReport.SetCellValue(5, 2, DateToStr(Sessione.RecordStamp));
            xlsReport.SetCellValue(6, 2, TimeToStr(Sessione.RecordStamp));
            xlsReport.SetCellValue(7, 2, Sessione.OperatoreID);
            xlsReport.SetCellValue(8, 2, TimeToStr(Sessione.StartStamp));
            xlsReport.SetCellValue(9, 2, TimeToStr(Sessione.EndStamp));
            xlsReport.SetCellValue(10, 2, FormatDateTime('hh:nn:ss', Sessione.EndStamp - Sessione.StartStamp));
            xlsReport.SetCellValue(10, 7, kMainVersion );
            xlsReport.SetCellValue(11, 2, Sessione.Memo );  // posizione Note da definire.

            // Accumula tempi delle Fasi desiderate usando BeginStamp della Fase e EndStamp dell'ultimo step (anche di altra fase).
            // workTime Fase_B
            LogStep(542,542, 542, 'Print Work times fields...');
            xlsReport.SetCellValue(9, 4, 'Calibrazione');
            i := JobList.IndexOf('Fase_B');                           // trovo inizio Fase B
            if i >= 0 then begin
                aJobRecord := TJobRecord( JobList.Objects[ i ]);
                tBeg := aJobRecord.BeginStamp;
            end
            else begin
                JobParams.retMsg := 'ERROR Fase_B Not Found on workTime calc.';
                JobParams.retCode := resERROR;
                LogStep(543,543, 543, JobParams.retMsg);
                exit;
            end;
            for i := aJobRecord.ListIndex to JobList.Count-1 do begin  // e da li proseguo scorrendo tutti i restanti steps,
                aJobRecord := TJobRecord( JobList.Objects[ i ]);       // perciò alla fine avrò l'ultimo EndStamp della Fase B.
                if aJobRecord.GroupFase = 'Fase_B' then
                    tEnd := aJobRecord.EndStamp;                       // assegna tutti, ma resta solo ultimo step di Fase-B !
            end;
            xlsReport.SetCellValue(10, 4, FormatDateTime('hh:nn:ss', tBeg-tEnd));

            // workTime Fase_C e D
            xlsReport.SetCellValue(9, 5, 'Validazione');
            i := JobList.IndexOf('Fase_C');                           // trovo inizio Fase C
            if i >= 0 then begin
                aJobRecord := TJobRecord( JobList.Objects[ i ]);
                tBeg := aJobRecord.BeginStamp;
                tEnd := tBeg;   // provvisorio.
            end
            else begin
                JobParams.retMsg := 'ERROR Fase_C Not Found on workTime calc.';
                JobParams.retCode := resERROR;
                LogStep(544,544, 544, JobParams.retMsg);
                exit;
            end;
            for i := aJobRecord.ListIndex to JobList.Count-1 do begin  // e da li proseguo scorrendo tutti i restanti steps,
                aJobRecord := TJobRecord( JobList.Objects[ i ]);       // perciò alla fine avrò l'ultimo EndStamp della Fase D.
                if aJobRecord.GroupFase = 'Fase_D' then
                    tEnd := aJobRecord.EndStamp;                       // assegna tutti, ma resta solo ultimo step di Fase-D !
            end;
            xlsReport.SetCellValue(10, 5, FormatDateTime('hh:nn:ss', tBeg-tEnd));

            // stampa i records di Calibrazione.
            LogStep(545,545, 545, 'Print calibration table...');
            rowOffset := 3;
            // stampa Titolo
            xlsReport.SetCellValue(rowOffset, 5, 'CALIBRAZIONI:');
            xlsReport.SetCellValue(rowOffset-1, 6, 'Combinazione Assi:');
            xlsReport.SetCellValue(rowOffset-1, 8, 'Vx [deg]');
            xlsReport.SetCellValue(rowOffset-1, 9, 'Vy [deg]');
            // e relativi records a seguire
            i := 0;
            for calItem in CalibRecords do begin
                xlsReport.SetCellValue(rowOffset+i, 6, calItem.Value.degPairID );
                xlsReport.SetCellValue(rowOffset+i, 8, calItem.Value.DutAx_mV );
                xlsReport.SetCellValue(rowOffset+i, 9, calItem.Value.DutAy_mV );
                inc(i);
            end;

            // stampa i records di Test in due gruppi di valori per assi X e Y.
            LogStep(546,546, 546, 'Print TestsX table...');
            SetRoundMode(rmNearest);                  // setting per RoundTo()
            rowOffset := 13;
            // stampa tipo asse
            xlsReport.SetCellValue(rowOffset, 1, 'TEST ASSE X');
            // e relativi records a seguire
            i := 0;
            for testItem in TestXRecords do begin
                inc(i);
                xlsReport.SetCellValue(rowOffset+i, 1, testItem.Value.degPairID);
                xlsReport.SetCellValue(rowOffset+i, 2, RoundTo(testItem.Value.RefAx, -2));
                xlsReport.SetCellValue(rowOffset+i, 3, RoundTo(testItem.Value.RefAy, -2));
                xlsReport.SetCellValue(rowOffset+i, 4, RoundTo(testItem.Value.CorrRefTest, -2));
                xlsReport.SetCellValue(rowOffset+i, 5, RoundTo(testItem.Value.CorrRefCross, -2));
                xlsReport.SetCellValue(rowOffset+i, 6, RoundTo(testItem.Value.DutAx_deg, -2));
                xlsReport.SetCellValue(rowOffset+i, 7, RoundTo(testItem.Value.DutAy_deg, -2));
                xlsReport.SetCellValue(rowOffset+i, 8, RoundTo(testItem.Value.ErroreTest, -2));
                xlsReport.SetCellValue(rowOffset+i, 9, RoundTo(testItem.Value.ErroreCross, -2));
            end;

            LogStep(547,547, 547, 'Print TestsY table...');
            rowOffset := 23;
            // stampa tipo asse
            xlsReport.SetCellValue(rowOffset, 1, 'TEST ASSE Y');
            // e relativi records a seguire
            i := 0;
            for testItem in TestYRecords do begin
                inc(i);
                xlsReport.SetCellValue(rowOffset+i, 1, testItem.Value.degPairID );
                xlsReport.SetCellValue(rowOffset+i, 2, RoundTo(testItem.Value.RefAx, -2));
                xlsReport.SetCellValue(rowOffset+i, 3, RoundTo(testItem.Value.RefAy, -2));
                xlsReport.SetCellValue(rowOffset+i, 4, RoundTo(testItem.Value.CorrRefTest, -2));
                xlsReport.SetCellValue(rowOffset+i, 5, RoundTo(testItem.Value.CorrRefCross, -2));
                xlsReport.SetCellValue(rowOffset+i, 6, RoundTo(testItem.Value.DutAx_deg, -2));
                xlsReport.SetCellValue(rowOffset+i, 7, RoundTo(testItem.Value.DutAy_deg, -2));
                xlsReport.SetCellValue(rowOffset+i, 8, RoundTo(testItem.Value.ErroreTest, -2));
                xlsReport.SetCellValue(rowOffset+i, 9, RoundTo(testItem.Value.ErroreCross, -2));
            end;

            // print parametri di configurazione del DUT
            LogStep(548,548, 548, 'Print CAN Parameters table...');
            (*
            xlsReport.SetCellValue(32, 1, 'Parametri DUT');    // Alias
            xlsReport.SetCellValue(33, 1, 'Index');
            xlsReport.SetCellValue(33, 2, 'Subindex');
            xlsReport.SetCellValue(33, 3, 'Descrizione');
            xlsReport.SetCellValue(33, 6, 'Tipo');
            xlsReport.SetCellValue(33, 7, 'Valore 2WR');
            xlsReport.SetCellValue(33, 8, 'Valore Hex');
            xlsReport.SetCellValue(33, 9, 'Range');
            *)
            for i := 0 to lviewCFG.Items.Count -1 do begin   // Scorre listview che contiene Parametri già acquisiti
                // get param items
                ListItem := lviewCFG.Items.item[i];
                xlsReport.SetCellValue(34+i, 1, lviewCFG.Items[i].Caption);  // index
                xlsReport.SetCellValue(34+i, 2, ListItem.SubItems[0]);       // 0-subidx
                xlsReport.SetCellValue(34+i, 3, ListItem.SubItems[1]);       // 1-descriz
                // Col 4 e 5 skipped perchè unite alla 3 Descrizione.
                xlsReport.SetCellValue(34+i, 6, ListItem.SubItems[3]);       // 3-Type.
                xlsReport.SetCellValue(34+i, 7, ListItem.SubItems[2]);       // 2-WrValue copiato as Value2Wr.
                xlsReport.SetCellValue(34+i, 8, ListItem.SubItems[4]);       // 4-RdValue  dovrebbe essere in HEX
                xlsReport.SetCellValue(34+i, 9, ListItem.SubItems[5]);       // 5-range
            end;//for lviewCFG.Items

            //Standard Document Properties - Most are only for xlsx files. In xls files FlexCel will only change the Creation Date and Modified Date.
            xlsReport.DocumentProperties.SetStandardProperty(TPropertyId.Author, 'EOL Manager');
            //You will normally not set CreateDateTime, since this is a new file and FlexCel will automatically use the current datetime.
            //But if you are editing a file and want to preserve the original creation date, you need to either set PreserveCreationDate to true:
            xlsReport.DocumentProperties.PreserveCreationDate := True;
            //You will normally not set LastSavedBy, since this is a new file.
            //If you don't set it, FlexCel will use the creator instead.
            xlsReport.DocumentProperties.SetStandardProperty(TPropertyId.LastSavedBy, Sessione.OperatoreID );
            xlsReport.SelectCell(1, 1, false);  //Cell selection and scroll position.
            // Report completato.

            // show preview del report
            xlsReport.HeadingColWidth := -1;
            xlsReport.HeadingRowHeight := -1;
            xlsReport.PrintGridLines := cbxShowGrid.checked;
            xlsReport.PrintHeadings := cbxHeadings.Checked;
            ImgExport := TFlexCelImgExport.Create(xlsReport, false);
            ImgExport.AllVisibleSheets := false;
            xlsPreview.Document := ImgExport;
            xlsPreview.InvalidatePreview;

            Screen.Cursor := crDefault;
            LogStep(549,549, 549, 'Creazione Report FILE rev.'+ iRev.ToString);
            //rptFile := vReportPath +'\'+ Sessione.CodiceProdotto +'_'+ Sessione.Lotto +'\'+ Sessione.CodiceProdotto +'_EOL_'+ Sessione.Lotto + Sessione.Serial +'.xls';
            try
                // Forzo destinazione per report di prova.
                if vRecMode = 1 then begin
                    kXls := '_demo.xls';
                    vReportPath := kAppPath + 'Reports\';        // override vReportPath to C: Local
                end
                else begin
                    kXls := '.xls';
                    vReportPath := iniReportPath;                // vReportPath as from .INI
                end;
                // permetto override destinazione report, ma
                // se Demo mode, forzo report locali !
                if vRecMode = 1 then
                    rgXlsDir.ItemIndex := 0;
                if rgXlsDir.ItemIndex = 0 then
                    vReportPath := kAppPath + 'Reports\'         // override vReportPath to C: Local
                else
                    vReportPath := iniReportPath;                // vReportPath as from .INI  (solitamente Netserver)
                // Verifica accessibilità cartella principale su server:  vReportPath from .INI
                if DirectoryExists( vReportPath ) then begin     // Server disponibile e report di produzione
                    // Server disponibile, salva report in rete.
                    // Verifica esistenza cartella del Lotto, se no la crea.
                    rptFile := vReportPath + Sessione.CodiceProdotto +'_L'+ Sessione.Lotto;
                    if not DirectoryExists( rptFile ) then begin
                        // crea nuova cartella per Lotto.
                        if not CreateDir( rptFile ) then begin   // bloccante e l'utente potrà scegliere da GUI se salvare in locale.
                            JobParams.retMsg := 'ERROR impossibile creare il path: '+ rptFile;
                            JobParams.retCode := resERROR;
                            LogStep(550,550, 550, JobParams.retMsg);
                            inc(errCount);
                            exit;
                        end;
                    end;
                    // qui cartella del Lotto creata o già esistente.
                    // Verifica esistenza file...
                    rptFile := vReportPath + Sessione.CodiceProdotto +'_L'+ Sessione.Lotto +'\'+ Sessione.CodiceProdotto +'_EOL_'+ Sessione.Lotto +'#'+ Sessione.Serial;
                    vRev := '';  // prima cerca filename senza revisione
                    iRev := 1;
                    while dmBase.FindFile( rptFile + vRev + kXls ) and (iRev < 50) do begin    // max 50 revisioni!
                        // file esiste già, quindi aggiunge una revisione
                        vRev := '.rev'+ intToStr(iRev);    // se c'è già aggiunge '.revX';
                        inc( iRev );
                    end;
                    rptFile := rptFile + vRev + kXls;
                    // Salva file as XLS
                    xlsReport.Save( rptFile );
                end
                else begin
                    // Server non disponibile o report di prova, salva sempre in locale.
                    MessageDlg('ATTENZIONE:' +sLineBreak+ 'il Path  '+ vReportPath +sLineBreak+
                               'Non è raggiungibile !' +sLineBreak+ 'Verifica e riprova.', mtWarning, [mbOk], 0, mbOk);
                    JobParams.retMsg := 'ERROR path Not Found: '+ vReportPath;
                    JobParams.retCode := resERROR;
                    LogStep(551,551, 551, JobParams.retMsg);
                    inc(errCount);
                    exit;
                end;
                lbReportName.caption := rptFile;  // mostra a user il path del report salvato.
            except
              on E: Exception do begin
                  inc(errCount);
                  tStr := 'in RunStep_29 on Save EOL Report >> '+ E.ClassName +#10#13#9+ ' description: '+E.Message;
                  JobParams.retCode := resERROR;
                  JobParams.retMsg  := _Exception_ + tStr;
                  //LogStep(552,552, 552,  tStr );
                  LogSysEvent(svEXCEPTION, 2925, tStr);
                  // problemi di salvatagio report non sarebbero di tipo bloccante, ma user può riprovare via GUI.
              end;
            end;

            // qui errore non sarebbe bloccante perchè è solo problema di report mentre il collaudo DUT si presume ok.
            if errCount = 0 then begin
                LogStep(553,553, 553, 'OK Report EOL compilato e salvato in:');
                LogStep(554,554, 554,  rptFile );
                JobParams.retMsg := _Completato_;   // Ok report completato.
                JobParams.retCode := resOK;
                result := True;
            end;

        except
          on E: Exception do begin
              inc(errCount);
              tStr := 'in RunStep_29 printing End-of-Test Report >> '+ E.ClassName +#10#13#9+ ' description: '+E.Message;
              JobParams.retCode := resERROR;
              JobParams.retMsg  := _Exception_ + tStr;
              //LogStep(555,555, 555,  tStr );
              LogSysEvent(svEXCEPTION, 2930, tStr);
              // problemi di composizione report non sarebbero di tipo bloccante, ma user può riprovare via GUI.
          end;
        end;
    finally
        Screen.Cursor := crDefault;
        LogStep(556,556, 556, JobList.Strings[currJob]+' ERRORS = '+ intToStr(errCount));
        dmBase.wait(500);// anche solo per evitare click multipli involontari da parte user.
    end;
end;



//######################################################################################################


initialization
    RegisterClass(Toiac3Fasi2ax);


finalization
    UnregisterClass(Toiac3Fasi2ax);



end.
