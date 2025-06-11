object FormWIADemo: TFormWIADemo
  Left = 264
  Top = 97
  Caption = 'WIA Demo'
  ClientHeight = 750
  ClientWidth = 1000
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
    Top = 130
    Width = 1000
    Height = 620
    Align = alClient
    Proportional = True
    Stretch = True
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 1000
    Height = 130
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 0
    ExplicitWidth = 998
    DesignSize = (
      1000
      130)
    object lbProgress: TLabel
      Left = 290
      Top = 0
      Width = 3
      Height = 15
    end
    object lbSelected: TLabel
      Left = 15
      Top = 48
      Width = 3
      Height = 15
    end
    object Label1: TLabel
      Left = 144
      Top = 67
      Width = 37
      Height = 15
      Caption = 'Pages :'
    end
    object Label2: TLabel
      Left = 290
      Top = 96
      Width = 27
      Height = 15
      Caption = 'Path:'
    end
    object btAcquire: TButton
      Left = 10
      Top = 80
      Width = 60
      Height = 40
      Caption = 'Acquire'
      Enabled = False
      TabOrder = 0
      OnClick = btAcquireClick
    end
    object btSelect: TButton
      Left = 10
      Top = 0
      Width = 126
      Height = 40
      Caption = 'Select Source'
      TabOrder = 1
      OnClick = btSelectClick
    end
    object progressBar: TProgressBar
      Left = 290
      Top = 24
      Width = 570
      Height = 30
      Anchors = [akLeft, akTop, akRight]
      TabOrder = 2
      ExplicitWidth = 568
    end
    object cbTest: TCheckBox
      Left = 880
      Top = 0
      Width = 54
      Height = 24
      Anchors = [akTop, akRight]
      Caption = 'Tests'
      TabOrder = 3
      ExplicitLeft = 878
    end
    object edTests: TEdit
      Left = 942
      Top = 0
      Width = 46
      Height = 23
      Anchors = [akTop, akRight]
      TabOrder = 4
      ExplicitLeft = 940
    end
    object cbEnumLocal: TCheckBox
      Left = 144
      Top = 16
      Width = 126
      Height = 24
      Caption = 'Only Connected'
      Checked = True
      State = cbChecked
      TabOrder = 5
    end
    object btNative: TButton
      Left = 75
      Top = 80
      Width = 61
      Height = 40
      Caption = 'Native'
      Enabled = False
      TabOrder = 6
      OnClick = btNativeClick
    end
    object edPages: TSpinEdit
      Left = 145
      Top = 88
      Width = 62
      Height = 24
      MaxValue = 255
      MinValue = 0
      TabOrder = 7
      Value = 0
    end
    object edPath: TEdit
      Left = 331
      Top = 92
      Width = 432
      Height = 23
      Anchors = [akLeft, akTop, akRight]
      TabOrder = 8
      ExplicitWidth = 430
    end
    object btBrowse: TButton
      Left = 772
      Top = 89
      Width = 86
      Height = 31
      Anchors = [akTop, akRight]
      Caption = 'Browse ...'
      TabOrder = 9
      OnClick = btBrowseClick
      ExplicitLeft = 770
    end
  end
end
