object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'Form1'
  ClientHeight = 954
  ClientWidth = 1417
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object lblUrl: TLabel
    Left = 16
    Top = 10
    Width = 23
    Height = 13
    Caption = 'URL:'
  end
  object lblReqData: TLabel
    Left = 16
    Top = 40
    Width = 26
    Height = 13
    Caption = 'data:'
  end
  object btnGet: TButton
    Left = 472
    Top = 8
    Width = 121
    Height = 25
    Caption = 'GET'
    TabOrder = 0
    OnClick = btnGetClick
  end
  object Memo1: TMemo
    Left = 48
    Top = 37
    Width = 344
    Height = 121
    TabOrder = 1
  end
  object Memo2: TMemo
    Left = 8
    Top = 230
    Width = 1233
    Height = 707
    Lines.Strings = (
      '')
    ScrollBars = ssBoth
    TabOrder = 2
  end
  object btnPost: TButton
    Left = 472
    Top = 48
    Width = 121
    Height = 25
    Caption = 'Create by SO (POST)'
    TabOrder = 3
    OnClick = btnPostClick
  end
  object btnPut: TButton
    Left = 472
    Top = 79
    Width = 121
    Height = 25
    Caption = 'Update by SO (PUT)'
    TabOrder = 4
    OnClick = btnPutClick
  end
  object editUrl: TEdit
    Left = 48
    Top = 10
    Width = 401
    Height = 21
    TabOrder = 5
    Text = 'http://localhost/DES/periods?where=code+gt+2015'
  end
  object Memo3: TMemo
    Left = 800
    Top = 8
    Width = 497
    Height = 209
    TabOrder = 6
  end
  object btnCreateByAA: TButton
    Left = 616
    Top = 48
    Width = 105
    Height = 25
    Caption = 'Create by AArray'
    TabOrder = 7
    OnClick = btnCreateByAAClick
  end
  object btnUpdateByAa: TButton
    Left = 616
    Top = 79
    Width = 105
    Height = 25
    Caption = 'Update by AArray'
    TabOrder = 8
    OnClick = btnUpdateByAaClick
  end
  object btnSendEmail: TButton
    Left = 472
    Top = 133
    Width = 75
    Height = 25
    Caption = 'send email'
    TabOrder = 9
    OnClick = btnSendEmailClick
  end
  object btnSendSms: TButton
    Left = 576
    Top = 133
    Width = 75
    Height = 25
    Caption = 'send SMS'
    TabOrder = 10
    OnClick = btnSendSmsClick
  end
  object Button1: TButton
    Left = 657
    Top = 133
    Width = 75
    Height = 25
    Caption = 'Button1'
    TabOrder = 11
    OnClick = Button1Click
  end
  object IdSMTP1: TIdSMTP
    SASLMechanisms = <>
    Left = 768
    Top = 128
  end
  object IdMessage1: TIdMessage
    AttachmentEncoding = 'UUE'
    BccList = <>
    CCList = <>
    Encoding = meDefault
    FromList = <
      item
      end>
    Recipients = <>
    ReplyTo = <>
    ConvertPreamble = True
    Left = 832
    Top = 128
  end
end
