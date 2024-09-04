object Form1: TForm1
  Left = 391
  Top = 158
  Caption = 'WIA Demos'
  ClientHeight = 233
  ClientWidth = 540
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
    Left = 432
    Top = 192
    Width = 59
    Height = 15
    Caption = 'from Item :'
  end
  object Memo2: TMemo
    Left = 0
    Top = 8
    Width = 416
    Height = 224
    TabOrder = 0
  end
  object edDevTest: TEdit
    Left = 499
    Top = 72
    Width = 40
    Height = 23
    TabOrder = 1
    Text = '1'
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
    Top = 104
    Width = 123
    Height = 25
    Caption = 'List Capabilities'
    TabOrder = 3
    OnClick = btIntCapClick
  end
  object btListChilds: TButton
    Left = 416
    Top = 136
    Width = 123
    Height = 25
    Caption = 'List Childs'
    TabOrder = 4
    OnClick = btListChildsClick
  end
  object btDownload: TButton
    Left = 416
    Top = 168
    Width = 123
    Height = 25
    Caption = 'Download'
    TabOrder = 5
    OnClick = btDownloadClick
  end
  object btSelect: TButton
    Left = 416
    Top = 40
    Width = 123
    Height = 25
    Caption = 'Select Device'
    TabOrder = 6
    OnClick = btSelectClick
  end
  object edSelItemName: TEdit
    Left = 432
    Top = 208
    Width = 107
    Height = 23
    TabOrder = 7
  end
end
