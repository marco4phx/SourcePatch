object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'Source Patcher - v1.0'
  ClientHeight = 514
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
  PixelsPerInch = 96
  TextHeight = 13
  object Label2: TLabel
    Left = 51
    Top = 22
    Width = 96
    Height = 13
    Caption = 'Stringa da sostituire'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object Label1: TLabel
    Left = 287
    Top = 22
    Width = 68
    Height = 13
    Caption = 'Nuova Stringa'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object Label3: TLabel
    Left = 287
    Top = 61
    Width = 121
    Height = 13
    Caption = '%  indica integer counter'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object Label4: TLabel
    Left = 635
    Top = 22
    Width = 48
    Height = 13
    Caption = 'From Num'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object Label5: TLabel
    Left = 726
    Top = 22
    Width = 36
    Height = 13
    Caption = 'To Num'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object Label6: TLabel
    Left = 816
    Top = 22
    Width = 35
    Height = 13
    Caption = 'Step of'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object Label7: TLabel
    Left = 51
    Top = 72
    Width = 81
    Height = 13
    Caption = 'Cartella corrente'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object labCurrentDir: TLabel
    Left = 51
    Top = 92
    Width = 15
    Height = 13
    Caption = 'C:\'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object Panel1: TPanel
    Left = 0
    Top = 300
    Width = 938
    Height = 214
    Align = alBottom
    TabOrder = 0
    object pbarSostituzioni: TProgressBar
      Left = 316
      Top = 13
      Width = 298
      Height = 17
      BarColor = clGreen
      TabOrder = 0
    end
    object mmLog: TMemo
      Left = 1
      Top = 43
      Width = 936
      Height = 170
      Align = alBottom
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      ScrollBars = ssVertical
      TabOrder = 1
    end
    object btnStart: TBitBtn
      Left = 182
      Top = 9
      Width = 75
      Height = 25
      Caption = 'Start'
      Enabled = False
      TabOrder = 2
      OnClick = btnStartClick
    end
    object btnStop: TBitBtn
      Left = 681
      Top = 9
      Width = 75
      Height = 25
      Caption = 'Stop'
      TabOrder = 3
      OnClick = btnStopClick
    end
    object btnRefresh: TBitBtn
      Left = 816
      Top = 9
      Width = 75
      Height = 25
      Caption = 'Refresh'
      TabOrder = 4
      OnClick = btnRefreshClick
    end
    object BitBtn2: TBitBtn
      Left = 46
      Top = 9
      Width = 75
      Height = 25
      Caption = 'Change DIR'
      TabOrder = 5
      OnClick = BitBtn2Click
    end
  end
  object cbxSelectAllFiles: TCheckBox
    Left = 6
    Top = 123
    Width = 88
    Height = 17
    Caption = '  Check ALL'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 1
    OnClick = cbxSelectAllFilesClick
  end
  object lviewPasFiles: TListView
    Left = 0
    Top = 146
    Width = 938
    Height = 154
    Align = alBottom
    Anchors = [akLeft, akTop, akRight, akBottom]
    Checkboxes = True
    Columns = <
      item
        Caption = '     File Name'
        Width = 200
      end
      item
        Alignment = taCenter
        Caption = 'Esito'
        Width = 70
      end
      item
        Caption = 'Num Modifiche'
        Width = 100
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
  object edtSearchStr: TEdit
    Left = 46
    Top = 38
    Width = 216
    Height = 21
    AutoSelect = False
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 3
    Text = 'LogStep('
  end
  object edtPatchStr: TEdit
    Left = 282
    Top = 38
    Width = 216
    Height = 21
    AutoSelect = False
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 4
    Text = 'LogStep(%,'
  end
  object CheckBox1: TCheckBox
    Left = 514
    Top = 40
    Width = 90
    Height = 17
    Caption = 'With Counter'
    Checked = True
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    State = cbChecked
    TabOrder = 5
    OnClick = cbxSelectAllFilesClick
  end
  object edtFrom: TEdit
    Left = 634
    Top = 38
    Width = 69
    Height = 21
    AutoSelect = False
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    ReadOnly = True
    TabOrder = 6
    Text = '10'
  end
  object edtTo: TEdit
    Left = 725
    Top = 38
    Width = 69
    Height = 21
    AutoSelect = False
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    ReadOnly = True
    TabOrder = 7
    Text = '5000'
  end
  object edtStep: TEdit
    Left = 815
    Top = 38
    Width = 40
    Height = 21
    AutoSelect = False
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    ReadOnly = True
    TabOrder = 8
    Text = '1'
  end
  object cbxBackupFile: TCheckBox
    Left = 514
    Top = 68
    Width = 90
    Height = 17
    Caption = 'Backup original'
    Checked = True
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    State = cbChecked
    TabOrder = 9
    OnClick = cbxSelectAllFilesClick
  end
  object CheckBox2: TCheckBox
    Left = 514
    Top = 10
    Width = 90
    Height = 17
    Caption = 'Case Sensitive'
    Enabled = False
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 10
    OnClick = cbxSelectAllFilesClick
  end
  object BitBtn1: TBitBtn
    Left = 681
    Top = 73
    Width = 75
    Height = 25
    Caption = 'Test'
    TabOrder = 11
    Visible = False
    OnClick = BitBtn1Click
  end
end
