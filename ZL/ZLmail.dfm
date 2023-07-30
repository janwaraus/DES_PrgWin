object dmMail: TdmMail
  OldCreateOrder = False
  Height = 94
  Width = 252
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
    Left = 36
    Top = 8
  end
  object idSMTP: TIdSMTP
    IOHandler = idSSLHandler
    Port = 465
    SASLMechanisms = <>
    UseTLS = utUseImplicitTLS
    Left = 98
    Top = 8
  end
  object idSSLHandler: TIdSSLIOHandlerSocketOpenSSL
    Destination = ':465'
    MaxLineAction = maException
    Port = 465
    DefaultPort = 0
    SSLOptions.Method = sslvSSLv23
    SSLOptions.SSLVersions = [sslvSSLv2, sslvSSLv3, sslvTLSv1, sslvTLSv1_1, sslvTLSv1_2]
    SSLOptions.Mode = sslmUnassigned
    SSLOptions.VerifyMode = []
    SSLOptions.VerifyDepth = 0
    Left = 158
    Top = 8
  end
end
