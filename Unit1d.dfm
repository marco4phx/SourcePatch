object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'Parser DB Tecnico'
  ClientHeight = 699
  ClientWidth = 719
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
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 39
    Top = 31
    Width = 34
    Height = 13
    Caption = 'Origine'
  end
  object Label4: TLabel
    Left = 12
    Top = 62
    Width = 61
    Height = 13
    Caption = 'Destinazione'
  end
  object btnSelect: TBitBtn
    Left = 611
    Top = 26
    Width = 83
    Height = 25
    Caption = 'Select Source'
    TabOrder = 1
    OnClick = btnSelectClick
  end
  object edtSourceDir: TEdit
    Left = 79
    Top = 28
    Width = 527
    Height = 21
    AutoSelect = False
    ReadOnly = True
    TabOrder = 0
    Text = '\\Netserver\Database tecnico\'
  end
  object btnStart: TBitBtn
    Left = 274
    Top = 94
    Width = 75
    Height = 25
    Caption = 'Start'
    TabOrder = 3
    OnClick = btnStartClick
  end
  object btnStop: TBitBtn
    Left = 394
    Top = 94
    Width = 75
    Height = 25
    Caption = 'Stop'
    TabOrder = 4
    OnClick = btnStopClick
  end
  object cbxFileOverwrite: TCheckBox
    Left = 79
    Top = 94
    Width = 189
    Height = 17
    Caption = '  File overwrite (or make copy)'
    TabOrder = 2
  end
  object PageControl1: TPageControl
    Left = 0
    Top = 142
    Width = 719
    Height = 557
    ActivePage = TabSheet2
    Align = alBottom
    Anchors = [akLeft, akTop, akRight, akBottom]
    Font.Charset = ANSI_CHARSET
    Font.Color = clNavy
    Font.Height = -13
    Font.Name = 'Segoe UI'
    Font.Style = []
    ParentFont = False
    TabOrder = 5
    object TabSheet1: TTabSheet
      Caption = 'Codici Prodotto da Aggiornare       '
      ImageIndex = 2
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
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
      object CheckBox1: TCheckBox
        Left = 570
        Top = 65
        Width = 111
        Height = 17
        Caption = '  ignora prima riga'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
      end
      object lviewProdotti: TListView
        Left = 0
        Top = 107
        Width = 711
        Height = 418
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
        TabOrder = 3
        ViewStyle = vsReport
      end
      object cbxSelectAll: TCheckBox
        Left = 8
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
        TabOrder = 4
        OnClick = cbxSelectAllClick
      end
    end
    object TabSheet2: TTabSheet
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
        Width = 711
        Height = 418
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
    object TabSheet3: TTabSheet
      Caption = 'Log Operazioni         '
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
      object mmLog: TMemo
        Left = 0
        Top = 52
        Width = 711
        Height = 473
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
    end
  end
  object btnTest: TBitBtn
    Left = 517
    Top = 94
    Width = 75
    Height = 25
    Caption = 'Test'
    TabOrder = 6
    OnClick = btnTestClick
  end
  object edtDestDir: TEdit
    Left = 79
    Top = 59
    Width = 527
    Height = 21
    AutoSelect = False
    ReadOnly = True
    TabOrder = 7
    Text = '\\Netserver\05.DbTecnico\ '
  end
  object BitBtn2: TBitBtn
    Tag = 1
    Left = 611
    Top = 57
    Width = 83
    Height = 25
    Caption = 'Select Target'
    TabOrder = 8
    OnClick = btnSelectClick
  end
  object OpenDialog: TOpenDialog
    Filter = 'file Excel|*.XLSX|excel|*.XLS|All|*.*'
    Options = [ofHideReadOnly, ofPathMustExist, ofFileMustExist, ofNoNetworkButton, ofEnableSizing]
    Left = 375
    Top = 212
  end
end
