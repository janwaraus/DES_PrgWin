object DesU: TDesU
  Left = 0
  Top = 0
  Caption = 'DesU'
  ClientHeight = 319
  ClientWidth = 390
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object dbAbra: TZConnection
    ControlsCodePage = cCP_UTF16
    ClientCodepage = 'UTF8'
    Catalog = ''
    Properties.Strings = (
      'controls_cp=GET_ACP'
      'codepage=UTF8')
    ReadOnly = True
    HostName = ''
    Port = 0
    Database = ''
    User = ''
    Password = ''
    Protocol = 'firebirdd-2.1'
    Left = 8
    Top = 8
  end
  object qrAbra: TZQuery
    Connection = dbAbra
    Params = <>
    Left = 56
    Top = 8
  end
  object dbZakos: TZConnection
    ControlsCodePage = cCP_UTF16
    ClientCodepage = 'cp1250'
    Catalog = ''
    Properties.Strings = (
      'controls_cp=GET_ACP'
      'CLIENT_MULTI_STATEMENTS=1'
      'codepage=cp1250')
    HostName = ''
    Port = 0
    Database = ''
    User = ''
    Password = ''
    Protocol = 'mysql-5'
    Left = 8
    Top = 80
  end
  object qrZakos: TZQuery
    Connection = dbZakos
    Params = <>
    Left = 56
    Top = 80
  end
  object qrAbra2: TZQuery
    Connection = dbAbra
    Params = <>
    Left = 96
    Top = 8
  end
  object qrAbra3: TZQuery
    Connection = dbAbra
    Params = <>
    Left = 144
    Top = 8
  end
  object qrAbraOC: TZQuery
    Connection = dbAbra
    Params = <>
    Left = 216
    Top = 8
  end
  object idMessage: TIdMessage
    AttachmentEncoding = 'MIME'
    BccList = <>
    CCList = <>
    Encoding = meMIME
    FromList = <
      item
      end>
    Recipients = <>
    ReplyTo = <>
    ConvertPreamble = True
    Left = 12
    Top = 215
  end
  object idSMTP: TIdSMTP
    IOHandler = IdSSLIOHandler
    Port = 465
    SASLMechanisms = <>
    UseTLS = utUseImplicitTLS
    Left = 66
    Top = 215
  end
  object IdSSLIOHandler: TIdSSLIOHandlerSocketOpenSSL
    Destination = ':465'
    MaxLineAction = maException
    Port = 465
    DefaultPort = 0
    SSLOptions.Mode = sslmUnassigned
    SSLOptions.VerifyMode = []
    SSLOptions.VerifyDepth = 0
    Left = 136
    Top = 216
  end
  object qrZakosOC: TZQuery
    Connection = dbZakos
    Params = <>
    Left = 112
    Top = 80
  end
  object qrAbraOC2: TZQuery
    Connection = dbAbra
    Params = <>
    Left = 272
    Top = 8
  end
end
