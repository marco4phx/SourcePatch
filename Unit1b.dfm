object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'Parser DB Tecnico'
  ClientHeight = 504
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
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 55
    Top = 12
    Width = 199
    Height = 13
    Caption = 'Cartella in cui effettuare le sostituzioni'
  end
  object btnSelect: TBitBtn
    Left = 574
    Top = 26
    Width = 75
    Height = 25
    Caption = 'Select'
    TabOrder = 1
    OnClick = btnSelectClick
  end
  object edtCurrDir: TEdit
    Left = 50
    Top = 28
    Width = 507
    Height = 21
    AutoSelect = False
    ReadOnly = True
    TabOrder = 0
  end
  object btnStart: TBitBtn
    Left = 50
    Top = 94
    Width = 75
    Height = 25
    Caption = 'Start'
    Enabled = False
    TabOrder = 3
  end
  object btnStop: TBitBtn
    Left = 170
    Top = 94
    Width = 75
    Height = 25
    Caption = 'Stop'
    TabOrder = 4
  end
  object cbxSubdirs: TCheckBox
    Left = 50
    Top = 62
    Width = 136
    Height = 17
    Caption = '  include Sottocartelle'
    TabOrder = 2
  end
  object PageControl1: TPageControl
    Left = 0
    Top = 131
    Width = 719
    Height = 373
    ActivePage = TabSheet2
    Align = alBottom
    TabOrder = 5
    object TabSheet3: TTabSheet
      Caption = 'Codici Prodotto da Aggiornare'
      ImageIndex = 2
      ExplicitLeft = -1
      ExplicitTop = 22
      ExplicitWidth = 311
      ExplicitHeight = 363
      object Label3: TLabel
        Left = 51
        Top = 16
        Width = 167
        Height = 13
        Caption = 'File EXCEL con lista Codici Prodotto'
      end
      object ListBox2: TListBox
        Left = 0
        Top = 113
        Width = 711
        Height = 232
        Align = alBottom
        ItemHeight = 13
        Items.Strings = (
          'Codice')
        TabOrder = 0
        ExplicitTop = 97
      end
      object edtXlsCodici: TEdit
        Left = 46
        Top = 32
        Width = 507
        Height = 21
        AutoSelect = False
        ReadOnly = True
        TabOrder = 1
      end
      object btnOpen1: TBitBtn
        Left = 570
        Top = 30
        Width = 75
        Height = 25
        Caption = 'Open'
        TabOrder = 2
        OnClick = btnOpen1Click
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'Lista Sostituzioni'
      ImageIndex = 1
      ExplicitWidth = 281
      ExplicitHeight = 165
      object Label2: TLabel
        Left = 51
        Top = 14
        Width = 148
        Height = 13
        Caption = 'File EXCEL con lista Sostituzioni'
      end
      object ListBox1: TListBox
        Left = 0
        Top = 78
        Width = 711
        Height = 267
        Align = alBottom
        Columns = 2
        ItemHeight = 13
        Items.Strings = (
          'Originale'
          'Modificata')
        TabOrder = 0
      end
      object edtXlsSostituzioni: TEdit
        Left = 46
        Top = 30
        Width = 507
        Height = 21
        AutoSelect = False
        ReadOnly = True
        TabOrder = 1
      end
      object btnOpen2: TBitBtn
        Left = 570
        Top = 28
        Width = 75
        Height = 25
        Caption = 'Open'
        TabOrder = 2
        OnClick = btnOpen2Click
      end
    end
    object TabSheet1: TTabSheet
      Caption = 'Log Operazioni'
      ExplicitLeft = -167
      ExplicitTop = 58
      ExplicitWidth = 311
      ExplicitHeight = 255
      object Memo1: TMemo
        Left = 0
        Top = 0
        Width = 711
        Height = 345
        Align = alClient
        Lines.Strings = (
          'LOG:')
        TabOrder = 0
        ExplicitTop = 20
        ExplicitHeight = 325
      end
    end
  end
  object OpenDialog1: TOpenDialog
    Filter = 'file Excel|*.XLSX|excel|*.XLS|All|*.*'
    Options = [ofHideReadOnly, ofPathMustExist, ofFileMustExist, ofNoNetworkButton, ofEnableSizing]
    Left = 254
    Top = 181
  end
  object OpenDialog2: TOpenDialog
    Filter = 'file Excel|*.XLSX|excel|*.XLS|All|*.*'
    Options = [ofHideReadOnly, ofPathMustExist, ofFileMustExist, ofNoNetworkButton, ofEnableSizing]
    Left = 356
    Top = 183
  end
  object SelectDialog: TOpenDialog
    FileName = '*.'
    Filter = '*.'
    Options = [ofHideReadOnly, ofNoValidate, ofPathMustExist, ofNoNetworkButton, ofEnableSizing]
    Title = 'Cartella in cui effettuare le sostituzioni'
    Left = 348
    Top = 21
  end
end
