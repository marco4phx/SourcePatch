object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'Parser DB Tecnico - v1.2'
  ClientHeight = 861
  ClientWidth = 938
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  DesignSize = (
    938
    861)
  PixelsPerInch = 96
  TextHeight = 13
  object btnStart: TBitBtn
    Left = 461
    Top = 100
    Width = 75
    Height = 25
    Caption = 'Start'
    TabOrder = 1
    OnClick = btnStartClick
  end
  object btnStop: TBitBtn
    Left = 581
    Top = 100
    Width = 75
    Height = 25
    Caption = 'Stop'
    TabOrder = 2
    OnClick = btnStopClick
  end
  object cbxFileOverwrite: TCheckBox
    Left = 254
    Top = 93
    Width = 189
    Height = 17
    Caption = '  File overwrite (or make copy)'
    TabOrder = 0
  end
  object pageControl1: TPageControl
    Left = 0
    Top = 158
    Width = 938
    Height = 658
    ActivePage = tabIOP
    Align = alBottom
    Anchors = [akLeft, akTop, akRight, akBottom]
    Font.Charset = ANSI_CHARSET
    Font.Color = clNavy
    Font.Height = -13
    Font.Name = 'Segoe UI'
    Font.Style = []
    ParentFont = False
    TabOrder = 3
    object tabProdotti: TTabSheet
      Caption = 'Codici Prodotto da Aggiornare       '
      ImageIndex = 2
      object Label3: TLabel
        Left = 51
        Top = 14
        Width = 167
        Height = 13
        Caption = 'File EXCEL con lista Codici Prodotto'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
      end
      object edtXlsCodici: TEdit
        Left = 46
        Top = 30
        Width = 507
        Height = 21
        AutoSelect = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        ReadOnly = True
        TabOrder = 0
      end
      object btnOpen1: TBitBtn
        Left = 570
        Top = 28
        Width = 75
        Height = 25
        Caption = 'Open'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        OnClick = btnOpen1Click
      end
      object lviewProdotti: TListView
        Left = 0
        Top = 107
        Width = 930
        Height = 519
        Align = alBottom
        Anchors = [akLeft, akTop, akRight, akBottom]
        Checkboxes = True
        Columns = <
          item
            Caption = '     Codice Prodotto'
            Width = 150
          end
          item
            Alignment = taCenter
            Caption = 'Esito'
            Width = 70
          end
          item
            Caption = 'Modifiche'
            Width = 60
          end
          item
            Caption = 'Path'
            Width = 400
          end>
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        GridLines = True
        ParentFont = False
        TabOrder = 2
        ViewStyle = vsReport
      end
      object cbxSelectAll_Prods: TCheckBox
        Left = 6
        Top = 84
        Width = 136
        Height = 17
        Caption = '  Check ALL'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 3
        OnClick = cbxSelectAll_ProdsClick
      end
    end
    object tabIOP: TTabSheet
      Caption = 'IOP da Aggiornare         '
      ImageIndex = 3
      object lviewIOP: TListView
        Left = 0
        Top = 51
        Width = 930
        Height = 575
        Align = alBottom
        Anchors = [akLeft, akTop, akRight, akBottom]
        Checkboxes = True
        Columns = <
          item
            Caption = '     File IOP'
            Width = 150
          end
          item
            Alignment = taCenter
            Caption = 'Esito'
            Width = 70
          end
          item
            Caption = 'Modifiche'
            Width = 60
          end
          item
            Caption = 'Descrizione'
            Width = 400
          end>
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        GridLines = True
        ParentFont = False
        TabOrder = 0
        ViewStyle = vsReport
      end
      object cbxSelectAll_IOP: TCheckBox
        Left = 6
        Top = 28
        Width = 136
        Height = 17
        Caption = '  Check ALL'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        OnClick = cbxSelectAll_IOPClick
      end
    end
    object tabSostituzioni: TTabSheet
      Caption = 'Lista Sostituzioni           '
      ImageIndex = 1
      object Label2: TLabel
        Left = 51
        Top = 14
        Width = 148
        Height = 13
        Caption = 'File EXCEL con lista Sostituzioni'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
      end
      object edtXlsSostituzioni: TEdit
        Left = 46
        Top = 30
        Width = 507
        Height = 21
        AutoSelect = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        ReadOnly = True
        TabOrder = 0
      end
      object btnOpen2: TBitBtn
        Left = 570
        Top = 28
        Width = 75
        Height = 25
        Caption = 'Open'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        OnClick = btnOpen2Click
      end
      object cbxSkip1stLine: TCheckBox
        Left = 46
        Top = 57
        Width = 136
        Height = 17
        Caption = '  ignora prima riga'
        Checked = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        State = cbChecked
        TabOrder = 2
      end
      object lviewSostituzioni: TListView
        Left = 0
        Top = 107
        Width = 930
        Height = 519
        Align = alBottom
        Anchors = [akLeft, akTop, akRight, akBottom]
        Columns = <
          item
            Caption = 'Orginale'
            Width = 270
          end
          item
            Caption = 'Modificata'
            Width = 600
          end>
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        GridLines = True
        ParentFont = False
        TabOrder = 3
        ViewStyle = vsReport
      end
    end
    object tabLog: TTabSheet
      Caption = 'Log Operazioni         '
      object mmLog: TMemo
        Left = 0
        Top = 52
        Width = 930
        Height = 574
        Align = alBottom
        Anchors = [akLeft, akTop, akRight, akBottom]
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        Lines.Strings = (
          'Log parser DBTecnico')
        ParentFont = False
        ScrollBars = ssVertical
        TabOrder = 0
      end
      object BitBtn1: TBitBtn
        Left = 537
        Top = 13
        Width = 108
        Height = 25
        Caption = 'Save Log to File'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        OnClick = BitBtn1Click
      end
      object BitBtn3: TBitBtn
        Left = 72
        Top = 13
        Width = 108
        Height = 25
        Caption = 'Clear Log '
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        OnClick = BitBtn3Click
      end
    end
  end
  object btnTest: TBitBtn
    Left = 704
    Top = 100
    Width = 75
    Height = 25
    Caption = 'Test'
    Enabled = False
    TabOrder = 4
    OnClick = btnTestClick
  end
  object cbxForceDir: TCheckBox
    Left = 254
    Top = 115
    Width = 172
    Height = 17
    Caption = '  Force creation of dest folder'
    TabOrder = 5
  end
  object rgArea: TRadioGroup
    Left = 33
    Top = 25
    Width = 125
    Height = 88
    Caption = ' Area di Manutenzione '
    ItemIndex = 0
    Items.Strings = (
      'DB Tecnico'
      'IOP')
    TabOrder = 6
    OnClick = rgAreaClick
  end
  object Panel1: TPanel
    Left = 0
    Top = 816
    Width = 938
    Height = 45
    Align = alBottom
    TabOrder = 7
    object pbarSostituzioni: TProgressBar
      Left = 316
      Top = 13
      Width = 298
      Height = 17
      BarColor = clGreen
      TabOrder = 0
    end
  end
  object PageControl2: TPageControl
    Left = 177
    Top = 8
    Width = 744
    Height = 82
    ActivePage = TabSheet2
    Anchors = [akLeft, akTop, akRight]
    Style = tsFlatButtons
    TabOrder = 8
    object TabSheet1: TTabSheet
      Caption = 'TabSheet1'
      TabVisible = False
      object Label1: TLabel
        Left = 34
        Top = 14
        Width = 34
        Height = 13
        Caption = 'Origine'
      end
      object Label4: TLabel
        Left = 7
        Top = 45
        Width = 61
        Height = 13
        Caption = 'Destinazione'
      end
      object btnSelectSource: TBitBtn
        Left = 606
        Top = 9
        Width = 83
        Height = 25
        Caption = 'Select Source'
        TabOrder = 0
        OnClick = btnSelectSourceClick
      end
      object edtSourceDir: TEdit
        Left = 74
        Top = 11
        Width = 527
        Height = 21
        AutoSelect = False
        TabOrder = 1
        Text = '\\Netserver\Database tecnico\'
      end
      object edtDestDir: TEdit
        Left = 74
        Top = 42
        Width = 527
        Height = 21
        AutoSelect = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        Text = '\\Netserver\05.DbTecnico\ '
      end
      object btnSelectTarget: TBitBtn
        Tag = 1
        Left = 606
        Top = 40
        Width = 83
        Height = 25
        Caption = 'Select Target'
        TabOrder = 3
        OnClick = btnSelectSourceClick
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'TabSheet2'
      ImageIndex = 1
      TabVisible = False
      object Label5: TLabel
        Left = 193
        Top = 30
        Width = 46
        Height = 13
        Caption = 'Percorso:'
        Color = clBtnFace
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentColor = False
        ParentFont = False
      end
      object cbxReparto: TComboBox
        Left = 8
        Top = 27
        Width = 145
        Height = 21
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        Text = 'Scegliere un Reparto'
        OnChange = cbxRepartoChange
        Items.Strings = (
          '01.SviluppoCommerciale'
          '02.R&D'
          '03.Progettazione'
          '04.Produzione'
          '05.Qualit'#224
          '06.Sicurezza'
          '08.ICT'
          '11.HR'
          '12.FCA')
      end
      object edtIopDir: TEdit
        Left = 246
        Top = 27
        Width = 453
        Height = 21
        AutoSelect = False
        Enabled = False
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 1
        Text = '\\Netserver\06.Qualit'#224'\SGQ2_attivit'#224'\01_Mansionari_IOP_MOD'
      end
    end
  end
  object OpenDialog: TOpenDialog
    Filter = 'file Excel|*.XLSX|excel|*.XLS|All|*.*'
    Options = [ofHideReadOnly, ofPathMustExist, ofFileMustExist, ofNoNetworkButton, ofEnableSizing]
    Left = 417
    Top = 263
  end
  object FDConnection1: TFDConnection
    ConnectionName = 'ConnDbGamma'
    Params.Strings = (
      'Database=XDOPTO'
      'User_Name=sarossi'
      'Password=84a79o47!'
      'Server=185.6.193.4'
      'LoginTimeout=10'
      'DriverID=MSSQL')
    LoginPrompt = False
    Left = 819
    Top = 95
  end
  object FDTable1: TFDTable
    Left = 729
    Top = 139
  end
  object FDPhysMSSQLDriverLink1: TFDPhysMSSQLDriverLink
    ODBCDriver = 'SQL Server'
    Left = 830
    Top = 157
  end
  object FDGUIxWaitCursor1: TFDGUIxWaitCursor
    Provider = 'Forms'
    Left = 840
    Top = 23
  end
end
