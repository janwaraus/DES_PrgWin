object fmZL: TfmZL
  Left = 755
  Top = 98
  Caption = 'Nezaplacen'#233' ZL'
  ClientHeight = 674
  ClientWidth = 786
  Color = clBtnFace
  Constraints.MinHeight = 615
  Constraints.MinWidth = 590
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object pnBottom: TPanel
    Left = 0
    Top = 554
    Width = 786
    Height = 120
    Align = alBottom
    TabOrder = 0
    object mmMail: TMemo
      Left = 1
      Top = 1
      Width = 784
      Height = 118
      Align = alClient
      Lines.Strings = (
        'V'#225#382'en'#253' pane, v'#225#382'en'#225' pan'#237','
        
          'dovolujeme si V'#225's upozornit, '#382'e je XXX dn'#237' po splatnosti z'#225'lohy ' +
          'na p'#345'ipojen'#237' k internetu a st'#225'le od V'#225's postr'#225'd'#225'me jej'#237' '
        #250'plnou '
        #250'hradu. Dlu'#382'n'#225' '#269#225'stka '#269'in'#237' YYY K'#269'.'#39
        
          'Pot'#283#353'ilo by n'#225's, kdybyste '#269#225'stku z'#225'lohy co nejd'#345#237've uhradili a m' +
          'y V'#225'm nemuseli omezovat poskytovan'#233' slu'#382'by.')
      TabOrder = 0
    end
  end
  object pnMain: TPanel
    Left = 0
    Top = 0
    Width = 786
    Height = 554
    Align = alClient
    TabOrder = 1
    object lbDo: TLabel
      Left = 10
      Top = 42
      Width = 51
      Height = 13
      Caption = 'Splatno do'
    end
    object lbOd: TLabel
      Left = 10
      Top = 4
      Width = 51
      Height = 13
      Caption = 'Splatno od'
    end
    object deDatumOd: TDateEdit
      Left = 10
      Top = 20
      Width = 90
      Height = 18
      BorderStyle = bsNone
      NumGlyphs = 2
      YearDigits = dyFour
      TabOrder = 0
    end
    object deDatumDo: TDateEdit
      Left = 10
      Top = 58
      Width = 90
      Height = 18
      BorderStyle = bsNone
      NumGlyphs = 2
      YearDigits = dyFour
      TabOrder = 1
    end
    object btKonec: TButton
      Left = 22
      Top = 462
      Width = 65
      Height = 21
      Caption = '&Konec'
      TabOrder = 3
      OnClick = btKonecClick
    end
    object asgPohledavky: TAdvStringGrid
      Left = 102
      Top = 1
      Width = 683
      Height = 552
      Align = alRight
      Anchors = [akLeft, akTop, akRight, akBottom]
      BorderStyle = bsNone
      ColCount = 14
      Ctl3D = True
      DefaultRowHeight = 18
      DrawingStyle = gdsClassic
      FixedColor = clWhite
      FixedCols = 0
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goDrawFocusSelected, goColSizing, goEditing]
      ParentCtl3D = False
      TabOrder = 2
      GridLineColor = 13948116
      GridFixedLineColor = 11250603
      OnGetAlignment = asgPohledavkyGetAlignment
      OnGetFormat = asgPohledavkyGetFormat
      OnClickSort = asgPohledavkyClickSort
      OnCanSort = asgPohledavkyCanSort
      OnClickCell = asgPohledavkyClickCell
      OnDblClickCell = asgPohledavkyDblClickCell
      OnCanEditCell = asgPohledavkyCanEditCell
      ActiveCellFont.Charset = DEFAULT_CHARSET
      ActiveCellFont.Color = 4474440
      ActiveCellFont.Height = -11
      ActiveCellFont.Name = 'MS Sans Serif'
      ActiveCellFont.Style = [fsBold]
      ActiveCellColor = 11565130
      ActiveCellColorTo = 11565130
      BorderColor = 11250603
      ColumnHeaders.Strings = (
        ' z'#225'kazn'#237'k'
        'po'#269'et ZL'
        ' dluh ZL'
        ' celkem'
        'FId        '
        'stav'
        'smlouva'
        'F.Code       '
        '      '
        ' mail'
        'CuId'
        'CId'
        'dny'
        'mobil SMS')
      ColumnSize.StretchColumn = 0
      ControlLook.FixedGradientFrom = clWhite
      ControlLook.FixedGradientTo = clWhite
      ControlLook.FixedGradientMirrorFrom = clWhite
      ControlLook.FixedGradientMirrorTo = clWhite
      ControlLook.FixedGradientHoverFrom = clGray
      ControlLook.FixedGradientHoverTo = clWhite
      ControlLook.FixedGradientHoverMirrorFrom = clWhite
      ControlLook.FixedGradientHoverMirrorTo = clWhite
      ControlLook.FixedGradientHoverBorder = 11645361
      ControlLook.FixedGradientDownFrom = clWhite
      ControlLook.FixedGradientDownTo = clWhite
      ControlLook.FixedGradientDownMirrorFrom = clWhite
      ControlLook.FixedGradientDownMirrorTo = clWhite
      ControlLook.FixedGradientDownBorder = 11250603
      ControlLook.ControlStyle = csWinXP
      ControlLook.DropDownHeader.Font.Charset = DEFAULT_CHARSET
      ControlLook.DropDownHeader.Font.Color = clWindowText
      ControlLook.DropDownHeader.Font.Height = -11
      ControlLook.DropDownHeader.Font.Name = 'Tahoma'
      ControlLook.DropDownHeader.Font.Style = []
      ControlLook.DropDownHeader.Visible = True
      ControlLook.DropDownHeader.Buttons = <>
      ControlLook.DropDownFooter.Font.Charset = DEFAULT_CHARSET
      ControlLook.DropDownFooter.Font.Color = clWindowText
      ControlLook.DropDownFooter.Font.Height = -11
      ControlLook.DropDownFooter.Font.Name = 'MS Sans Serif'
      ControlLook.DropDownFooter.Font.Style = []
      ControlLook.DropDownFooter.Visible = True
      ControlLook.DropDownFooter.Buttons = <>
      EnableHTML = False
      Filter = <>
      FilterDropDown.Font.Charset = DEFAULT_CHARSET
      FilterDropDown.Font.Color = clWindowText
      FilterDropDown.Font.Height = -11
      FilterDropDown.Font.Name = 'MS Sans Serif'
      FilterDropDown.Font.Style = []
      FilterDropDown.TextChecked = 'Checked'
      FilterDropDown.TextUnChecked = 'Unchecked'
      FilterDropDownClear = '(All)'
      FilterEdit.TypeNames.Strings = (
        'Starts with'
        'Ends with'
        'Contains'
        'Not contains'
        'Equal'
        'Not equal'
        'Clear')
      FixedColWidth = 66
      FixedRowHeight = 18
      FixedRowAlways = True
      FixedFont.Charset = EASTEUROPE_CHARSET
      FixedFont.Color = 3881787
      FixedFont.Height = -11
      FixedFont.Name = 'MS Sans Serif'
      FixedFont.Style = []
      FloatFormat = '%.2f'
      HoverButtons.Buttons = <>
      HTMLSettings.ImageFolder = 'images'
      HTMLSettings.ImageBaseName = 'img'
      Look = glCustom
      PrintSettings.DateFormat = 'dd/mm/yyyy'
      PrintSettings.Font.Charset = DEFAULT_CHARSET
      PrintSettings.Font.Color = clWindowText
      PrintSettings.Font.Height = -11
      PrintSettings.Font.Name = 'MS Sans Serif'
      PrintSettings.Font.Style = []
      PrintSettings.FixedFont.Charset = DEFAULT_CHARSET
      PrintSettings.FixedFont.Color = clWindowText
      PrintSettings.FixedFont.Height = -11
      PrintSettings.FixedFont.Name = 'MS Sans Serif'
      PrintSettings.FixedFont.Style = []
      PrintSettings.HeaderFont.Charset = DEFAULT_CHARSET
      PrintSettings.HeaderFont.Color = clWindowText
      PrintSettings.HeaderFont.Height = -11
      PrintSettings.HeaderFont.Name = 'MS Sans Serif'
      PrintSettings.HeaderFont.Style = []
      PrintSettings.FooterFont.Charset = DEFAULT_CHARSET
      PrintSettings.FooterFont.Color = clWindowText
      PrintSettings.FooterFont.Height = -11
      PrintSettings.FooterFont.Name = 'MS Sans Serif'
      PrintSettings.FooterFont.Style = []
      PrintSettings.PageNumSep = '/'
      RowHeaders.Strings = (
        ' ')
      ScrollBarAlways = saVert
      ScrollWidth = 16
      SearchFooter.ColorTo = 15790320
      SearchFooter.FindNextCaption = 'Find &next'
      SearchFooter.FindPrevCaption = 'Find &previous'
      SearchFooter.Font.Charset = DEFAULT_CHARSET
      SearchFooter.Font.Color = clWindowText
      SearchFooter.Font.Height = -11
      SearchFooter.Font.Name = 'MS Sans Serif'
      SearchFooter.Font.Style = []
      SearchFooter.HighLightCaption = 'Highlight'
      SearchFooter.HintClose = 'Close'
      SearchFooter.HintFindNext = 'Find next occurrence'
      SearchFooter.HintFindPrev = 'Find previous occurrence'
      SearchFooter.HintHighlight = 'Highlight occurrences'
      SearchFooter.MatchCaseCaption = 'Match case'
      SearchFooter.ResultFormat = '(%d of %d)'
      SelectionColor = 13744549
      SortSettings.Column = 0
      SortSettings.Show = True
      SortSettings.HeaderColor = clWhite
      SortSettings.HeaderColorTo = clWhite
      SortSettings.HeaderMirrorColor = clWhite
      SortSettings.HeaderMirrorColorTo = clWhite
      VAlignment = vtaCenter
      Version = '8.6.2.0'
      ColWidths = (
        66
        64
        56
        53
        37
        49
        43
        51
        40
        46
        33
        32
        64
        64)
    end
    object acbRada: TAdvComboBox
      Left = 10
      Top = 96
      Width = 63
      Height = 21
      Color = clWindow
      Version = '2.0.0.0'
      Visible = True
      ButtonWidth = 17
      Style = csDropDownList
      Flat = True
      FlatLineColor = clSilver
      EmptyTextStyle = []
      Ctl3D = False
      DropWidth = 0
      Enabled = True
      ItemIndex = -1
      Items.Strings = (
        'Mn'#237#353'ek'
        'O'#345'ech'
        'Stod'#367'lky'
        #381'i'#382'kov')
      LabelCaption = #344'ada ZL'
      LabelPosition = lpTopLeft
      LabelMargin = 1
      LabelAlwaysEnabled = True
      LabelFont.Charset = DEFAULT_CHARSET
      LabelFont.Color = clWindowText
      LabelFont.Height = -11
      LabelFont.Name = 'MS Sans Serif'
      LabelFont.Style = []
      ParentCtl3D = False
      TabOrder = 4
    end
    object btVyber: TButton
      Left = 22
      Top = 194
      Width = 65
      Height = 21
      Caption = '&Vyber'
      TabOrder = 5
      OnClick = btVyberClick
    end
    object btExport: TButton
      Left = 22
      Top = 226
      Width = 65
      Height = 21
      Caption = '&Export'
      TabOrder = 6
      OnClick = btExportClick
    end
    object btMail: TButton
      Left = 22
      Top = 258
      Width = 65
      Height = 21
      Caption = '&Mail'
      TabOrder = 7
      OnClick = btMailClick
    end
    object acbDruhSmlouvy: TAdvComboBox
      Left = 10
      Top = 134
      Width = 89
      Height = 21
      Color = clWindow
      Version = '2.0.0.0'
      Visible = True
      ButtonWidth = 17
      Style = csDropDownList
      Flat = True
      FlatLineColor = clSilver
      EmptyTextStyle = []
      Ctl3D = False
      DropWidth = 0
      Enabled = True
      ItemIndex = -1
      Items.Strings = (
        '%'
        '0'
        '1'
        '2'
        '3'
        '4'
        '5'
        '7')
      LabelCaption = 'Stav smlouvy'
      LabelPosition = lpTopLeft
      LabelMargin = 1
      LabelAlwaysEnabled = True
      LabelFont.Charset = DEFAULT_CHARSET
      LabelFont.Color = clWindowText
      LabelFont.Height = -11
      LabelFont.Name = 'MS Sans Serif'
      LabelFont.Style = []
      ParentCtl3D = False
      TabOrder = 8
    end
    object cbCast: TCheckBox
      Left = 14
      Top = 166
      Width = 87
      Height = 17
      Caption = 'I '#269#225'ste'#269'n'#283
      TabOrder = 9
    end
    object btOdpojit: TButton
      Left = 22
      Top = 432
      Width = 65
      Height = 21
      Caption = 'O&dpojit'
      TabOrder = 10
      OnClick = btOdpojitClick
    end
    object rgText: TRadioGroup
      Left = 22
      Top = 321
      Width = 65
      Height = 73
      ItemIndex = 0
      Items.Strings = (
        'Text 1'
        'Text 2'
        'Text 3')
      TabOrder = 11
      OnClick = rgTextClick
    end
    object btSMS: TButton
      Left = 22
      Top = 289
      Width = 65
      Height = 22
      Caption = '&SMS'
      TabOrder = 12
      OnClick = btSMSClick
    end
    object btOmezit: TButton
      Left = 22
      Top = 402
      Width = 65
      Height = 21
      Caption = '&Omezit'
      TabOrder = 13
      OnClick = btOmezitClick
    end
  end
  object dlgExport: TSaveDialog
    DefaultExt = '.xls'
    Filter = 'xls|*.xls'
    Options = [ofHideReadOnly, ofNoReadOnlyReturn, ofEnableSizing]
    Left = 134
    Top = 90
  end
  object idHTTP: TIdHTTP
    HandleRedirects = True
    ProxyParams.BasicAuthentication = False
    ProxyParams.ProxyPort = 0
    Request.ContentLength = 0
    Request.ContentRangeEnd = 0
    Request.ContentRangeStart = 0
    Request.ContentRangeInstanceLength = -1
    Request.ContentType = 'text/html'
    Request.Accept = 'text/html, */*'
    Request.BasicAuthentication = False
    Request.UserAgent = 'Mozilla/3.0 (compatible; Indy Library)'
    Request.Ranges.Units = 'bytes'
    Request.Ranges = <>
    HTTPOptions = [hoForceEncodeParams]
    Left = 130
    Top = 39
  end
  object IdAntiFreeze: TIdAntiFreeze
    OnlyWhenIdle = False
    Left = 190
    Top = 39
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
    Left = 250
    Top = 39
  end
  object idSMTP: TIdSMTP
    IOHandler = idSSLHandler
    Port = 465
    SASLMechanisms = <>
    UseTLS = utUseImplicitTLS
    Left = 310
    Top = 39
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
    Left = 368
    Top = 39
  end
end
