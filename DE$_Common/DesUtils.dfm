object DesU: TDesU
  Left = 0
  Top = 0
  Caption = 'DesU'
  ClientHeight = 140
  ClientWidth = 224
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
    ControlsCodePage = cGET_ACP
    Catalog = ''
    Properties.Strings = (
      'controls_cp=GET_ACP')
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
    ControlsCodePage = cGET_ACP
    Catalog = ''
    Properties.Strings = (
      'controls_cp=GET_ACP')
    ReadOnly = True
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
end
