object DesFastReport: TDesFastReport
  Left = 0
  Top = 0
  Caption = 'DesFastReport'
  ClientHeight = 411
  ClientWidth = 852
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object qrAbraDPH: TZQuery
    Connection = DesU.dbAbra
    ReadOnly = True
    SQL.Strings = (
      '')
    Params = <>
    Left = 20
    Top = 134
  end
  object qrAbraRadky: TZQuery
    Connection = DesU.dbAbra
    ReadOnly = True
    SQL.Strings = (
      '')
    Params = <>
    Left = 20
    Top = 78
  end
  object frxReport: TfrxReport
    Version = '2023.1.3'
    DotMatrixReport = False
    IniFile = '\Software\Fast Reports'
    PreviewOptions.Buttons = [pbPrint, pbLoad, pbSave, pbExport, pbZoom, pbFind, pbOutline, pbPageSetup, pbTools, pbEdit, pbNavigator, pbExportQuick]
    PreviewOptions.Zoom = 1.000000000000000000
    PrintOptions.Printer = 'Default'
    PrintOptions.PrintOnSheet = 0
    PrintOptions.ShowDialog = False
    ReportOptions.CreateDate = 41751.679037013900000000
    ReportOptions.LastChange = 45060.077223564820000000
    ScriptLanguage = 'PascalScript'
    StoreInDFM = False
    OnBeginDoc = frxReportBeginDoc
    OnGetValue = frxReportGetValue
    Left = 17
    Top = 14
  end
  object fdsRadky: TfrxDBDataset
    UserName = 'fdsRadky'
    CloseDataSource = False
    OpenDataSource = False
    DataSet = qrAbraRadky
    BCDToCurrency = False
    DataSetOptions = []
    Left = 88
    Top = 78
  end
  object fdsDPH: TfrxDBDataset
    UserName = 'fdsDPH'
    CloseDataSource = False
    OpenDataSource = False
    DataSet = qrAbraDPH
    BCDToCurrency = False
    DataSetOptions = []
    Left = 88
    Top = 134
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
    Left = 20
    Top = 216
  end
  object idSMTP: TIdSMTP
    SASLMechanisms = <>
    Left = 82
    Top = 216
  end
  object frxPDFExport1: TfrxPDFExport
    UseFileCache = True
    ShowProgress = True
    OverwritePrompt = False
    DataOnly = False
    EmbedFontsIfProtected = False
    InteractiveFormsFontSubset = 'A-Z,a-z,0-9,#43-#47 '
    OpenAfterExport = False
    PrintOptimized = False
    Outline = False
    Background = False
    HTMLTags = True
    Quality = 95
    Author = 'FastReport'
    Subject = 'FastReport PDF export'
    Creator = 'FastReport'
    ProtectionFlags = [ePrint, eModify, eCopy, eAnnot]
    HideToolbar = False
    HideMenubar = False
    HideWindowUI = False
    FitWindow = False
    CenterWindow = False
    PrintScaling = False
    PdfA = False
    PDFStandard = psNone
    PDFVersion = pv17
    Left = 96
    Top = 16
  end
  object frxBarCodeObject1: TfrxBarCodeObject
    Left = 184
    Top = 16
  end
end
