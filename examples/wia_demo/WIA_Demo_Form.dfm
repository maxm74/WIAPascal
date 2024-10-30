object FormWIADemo: TFormWIADemo
  Left = 260
  Top = 102
  Caption = 'WIA Demo'
  ClientHeight = 576
  ClientWidth = 769
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  TextHeight = 15
  object ImageHolder: TImage
    Left = 0
    Top = 48
    Width = 769
    Height = 528
    Align = alClient
    Proportional = True
    Stretch = True
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 769
    Height = 48
    Align = alTop
    TabOrder = 0
    object lbProgress: TLabel
      Left = 304
      Top = 0
      Width = 3
      Height = 15
    end
    object lbSelected: TLabel
      Left = 12
      Top = 32
      Width = 3
      Height = 15
    end
    object btAcquire: TButton
      Left = 208
      Top = 8
      Width = 48
      Height = 32
      Caption = 'Acquire'
      Enabled = False
      TabOrder = 0
      OnClick = btAcquireClick
    end
    object btSelect: TButton
      Left = 8
      Top = 0
      Width = 88
      Height = 32
      Caption = 'Select Source'
      TabOrder = 1
      OnClick = btSelectClick
    end
    object progressBar: TProgressBar
      Left = 304
      Top = 16
      Width = 300
      Height = 24
      TabOrder = 2
    end
    object cbTest: TCheckBox
      Left = 657
      Top = 8
      Width = 44
      Height = 19
      Caption = 'Tests'
      TabOrder = 3
    end
    object edTests: TEdit
      Left = 707
      Top = 8
      Width = 56
      Height = 23
      TabOrder = 4
    end
    object cbEnumLocal: TCheckBox
      Left = 96
      Top = 8
      Width = 110
      Height = 19
      Caption = 'Only Connected'
      Checked = True
      State = cbChecked
      TabOrder = 5
    end
    object btNative: TButton
      Left = 260
      Top = 8
      Width = 40
      Height = 32
      Caption = 'Native'
      Enabled = False
      TabOrder = 6
      OnClick = btNativeClick
    end
  end
end
