object FormWIADemo: TFormWIADemo
  Left = 234
  Top = 98
  Caption = 'FormWIADemo'
  ClientHeight = 575
  ClientWidth = 765
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
    Top = 40
    Width = 765
    Height = 535
    Align = alClient
    Proportional = True
    Stretch = True
    ExplicitWidth = 769
    ExplicitHeight = 536
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 765
    Height = 40
    Align = alTop
    TabOrder = 0
    object btAcquire: TButton
      Left = 104
      Top = 0
      Width = 75
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
      Left = 192
      Top = 5
      Width = 350
      Height = 20
      TabOrder = 2
    end
  end
end
