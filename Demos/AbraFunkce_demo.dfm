object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'Form1'
  ClientHeight = 413
  ClientWidth = 574
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object lbl1: TLabel
    Left = 16
    Top = 10
    Width = 16
    Height = 13
    Caption = 'lbl1'
  end
  object Memo1: TMemo
    Left = 48
    Top = 93
    Width = 505
    Height = 300
    TabOrder = 0
  end
  object btnDoIt1: TButton
    Left = 48
    Top = 40
    Width = 75
    Height = 25
    Caption = 'DoIt 1'
    TabOrder = 1
    OnClick = btnDoIt1Click
  end
  object edit1: TEdit
    Left = 48
    Top = 13
    Width = 401
    Height = 21
    TabOrder = 2
  end
  object Button1: TButton
    Left = 144
    Top = 40
    Width = 75
    Height = 25
    Caption = 'Button1'
    TabOrder = 3
    OnClick = Button1Click
  end
end
