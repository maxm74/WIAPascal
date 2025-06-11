object FormWIATests: TFormWIATests
  Left = 391
  Top = 158
  BorderStyle = bsSingle
  Caption = 'WIA Tests'
  ClientHeight = 325
  ClientWidth = 680
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  TextHeight = 15
  object Label1: TLabel
    Left = 524
    Top = 94
    Width = 74
    Height = 15
    Caption = 'Device to Use:'
  end
  object Label2: TLabel
    Left = 524
    Top = 160
    Width = 66
    Height = 15
    Caption = 'Item to Use :'
  end
  object Label3: TLabel
    Left = 520
    Top = 290
    Width = 17
    Height = 15
    Caption = 'dpi'
  end
  object Memo2: TMemo
    Left = 0
    Top = 10
    Width = 514
    Height = 309
    TabOrder = 0
  end
  object edDevTest: TEdit
    Left = 624
    Top = 90
    Width = 50
    Height = 23
    TabOrder = 1
    Text = '0'
  end
  object btIntList: TButton
    Left = 520
    Top = 10
    Width = 154
    Height = 31
    Caption = 'List Wia'
    TabOrder = 2
    OnClick = btIntListClick
  end
  object btIntCap: TButton
    Left = 520
    Top = 220
    Width = 154
    Height = 31
    Caption = 'List Capabilities'
    TabOrder = 3
    OnClick = btIntCapClick
  end
  object btDownload: TButton
    Left = 520
    Top = 260
    Width = 154
    Height = 31
    Caption = 'Download/Acquire'
    TabOrder = 4
    OnClick = btDownloadClick
  end
  object btSelect: TButton
    Left = 520
    Top = 50
    Width = 154
    Height = 31
    Caption = 'Select Device'
    TabOrder = 5
    OnClick = btSelectClick
  end
  object edSelItemName: TEdit
    Left = 524
    Top = 180
    Width = 154
    Height = 23
    TabOrder = 6
  end
  object btListChilds: TButton
    Left = 520
    Top = 120
    Width = 154
    Height = 31
    Caption = 'List Childs'
    TabOrder = 7
    OnClick = btListChildsClick
  end
  object edDPI: TSpinEdit
    Left = 550
    Top = 290
    Width = 59
    Height = 24
    Increment = 50
    MaxValue = 600
    MinValue = 100
    TabOrder = 8
    Value = 100
  end
end
