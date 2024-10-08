object Form1: TForm1
  Left = 391
  Top = 158
  Caption = 'WIA Demos'
  ClientHeight = 260
  ClientWidth = 544
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
    Left = 419
    Top = 75
    Width = 74
    Height = 15
    Caption = 'Device to Use:'
  end
  object Label2: TLabel
    Left = 419
    Top = 128
    Width = 66
    Height = 15
    Caption = 'Item to Use :'
  end
  object Label3: TLabel
    Left = 416
    Top = 232
    Width = 17
    Height = 15
    Caption = 'dpi'
  end
  object Memo2: TMemo
    Left = 0
    Top = 8
    Width = 416
    Height = 247
    TabOrder = 0
  end
  object edDevTest: TEdit
    Left = 499
    Top = 72
    Width = 40
    Height = 23
    TabOrder = 1
    Text = '0'
  end
  object btIntList: TButton
    Left = 416
    Top = 8
    Width = 123
    Height = 25
    Caption = 'List Wia'
    TabOrder = 2
    OnClick = btIntListClick
  end
  object btIntCap: TButton
    Left = 416
    Top = 176
    Width = 123
    Height = 25
    Caption = 'List Capabilities'
    TabOrder = 3
    OnClick = btIntCapClick
  end
  object btDownload: TButton
    Left = 416
    Top = 208
    Width = 123
    Height = 25
    Caption = 'Download'
    TabOrder = 4
    OnClick = btDownloadClick
  end
  object btSelect: TButton
    Left = 416
    Top = 40
    Width = 123
    Height = 25
    Caption = 'Select Device'
    TabOrder = 5
    OnClick = btSelectClick
  end
  object edSelItemName: TEdit
    Left = 419
    Top = 144
    Width = 123
    Height = 23
    TabOrder = 6
  end
  object btListChilds: TButton
    Left = 416
    Top = 96
    Width = 123
    Height = 25
    Caption = 'List Childs'
    TabOrder = 7
    OnClick = btListChildsClick
  end
  object edDPI: TEdit
    Left = 440
    Top = 232
    Width = 47
    Height = 23
    TabOrder = 8
  end
end
