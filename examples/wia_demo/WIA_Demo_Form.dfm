object FormWIADemo: TFormWIADemo
  Left = 260
  Height = 576
  Top = 102
  Width = 769
  Caption = 'WIA Demo'
  ClientHeight = 576
  ClientWidth = 769
  LCLVersion = '3.99.0.0'
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  object ImageHolder: TImage
    Left = 0
    Height = 528
    Top = 48
    Width = 769
    Align = alClient
    Proportional = True
    Stretch = True
  end
  object Panel1: TPanel
    Left = 0
    Height = 48
    Top = 0
    Width = 769
    Align = alTop
    ClientHeight = 48
    ClientWidth = 769
    TabOrder = 0
    object btAcquire: TButton
      Left = 208
      Height = 32
      Top = 8
      Width = 48
      Caption = 'Acquire'
      Enabled = False
      TabOrder = 0
      OnClick = btAcquireClick
    end
    object btSelect: TButton
      Left = 8
      Height = 32
      Top = 0
      Width = 88
      Caption = 'Select Source'
      TabOrder = 1
      OnClick = btSelectClick
    end
    object progressBar: TProgressBar
      Left = 304
      Height = 24
      Top = 16
      Width = 300
      TabOrder = 2
    end
    object cbTest: TCheckBox
      Left = 657
      Height = 19
      Top = 8
      Width = 44
      Caption = 'Tests'
      TabOrder = 3
    end
    object edTests: TEdit
      Left = 707
      Height = 23
      Top = 8
      Width = 56
      TabOrder = 4
    end
    object cbEnumLocal: TCheckBox
      Left = 96
      Height = 19
      Top = 8
      Width = 104
      Caption = 'Only Connected'
      Checked = True
      State = cbChecked
      TabOrder = 5
    end
    object lbProgress: TLabel
      Left = 304
      Height = 1
      Top = 0
      Width = 1
    end
    object lbSelected: TLabel
      Left = 12
      Height = 1
      Top = 32
      Width = 1
    end
    object btNative: TButton
      Left = 256
      Height = 32
      Top = 8
      Width = 40
      Caption = 'Native'
      Enabled = False
      TabOrder = 6
      OnClick = btNativeClick
    end
  end
end
