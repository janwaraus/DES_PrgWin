object fmMain: TfmMain
  Left = 141
  Top = 0
  Caption = 'Na'#269'ten'#237', oprava a ulo'#382'en'#237' bankovn'#237'ho v'#253'pisu'
  ClientHeight = 675
  ClientWidth = 1379
  Color = clBtnFace
  Constraints.MinHeight = 714
  Constraints.MinWidth = 1395
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnShow = FormShow
  DesignSize = (
    1379
    675)
  PixelsPerInch = 96
  TextHeight = 13
  object lblHlavicka: TLabel
    Left = 115
    Top = 5
    Width = 83
    Height = 16
    Caption = 'lbl Hlavicka'
    Color = clSkyBlue
    Font.Charset = EASTEUROPE_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Verdana'
    Font.Style = [fsBold]
    ParentColor = False
    ParentFont = False
    Layout = tlBottom
  end
  object lblVypisFioGpc: TLabel
    Left = 240
    Top = 36
    Width = 69
    Height = 13
    Caption = 'lblVypisFioGpc'
  end
  object lblVypisFioInfo: TLabel
    Left = 240
    Top = 71
    Width = 67
    Height = 13
    Caption = 'lblVypisFioInfo'
  end
  object lblVypisFioSporiciGpc: TLabel
    Left = 240
    Top = 115
    Width = 101
    Height = 13
    Caption = 'lblVypisFioSporiciGpc'
  end
  object lblVypisFioSporiciInfo: TLabel
    Left = 240
    Top = 149
    Width = 99
    Height = 13
    Caption = 'lblVypisFioSporiciInfo'
  end
  object lblVypisCsobInfo: TLabel
    Left = 240
    Top = 306
    Width = 77
    Height = 13
    Caption = 'lblVypisCsobInfo'
  end
  object lblVypisCsobGpc: TLabel
    Left = 240
    Top = 275
    Width = 79
    Height = 13
    Caption = 'lblVypisCsobGpc'
  end
  object lblVypisPayuGpc: TLabel
    Left = 238
    Top = 357
    Width = 79
    Height = 13
    Caption = 'lblVypisPayuGpc'
  end
  object lblVypisPayuInfo: TLabel
    Left = 240
    Top = 391
    Width = 77
    Height = 13
    Caption = 'lblVypisPayuInfo'
  end
  object lblVypisFiokontoGpc: TLabel
    Left = 240
    Top = 189
    Width = 96
    Height = 13
    Caption = 'lblVypisFiokontoGpc'
  end
  object lblVypisFiokontoInfo: TLabel
    Left = 240
    Top = 220
    Width = 94
    Height = 13
    Caption = 'lblVypisFiokontoInfo'
  end
  object lblHlavickaVpravo: TLabel
    Left = 819
    Top = 8
    Width = 107
    Height = 15
    Alignment = taRightJustify
    Caption = 'lbl Hlavicka Vpravo'
    Color = clSkyBlue
    Font.Charset = EASTEUROPE_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentColor = False
    ParentFont = False
    Layout = tlBottom
  end
  object btnStahniVypisy: TButton
    Left = 128
    Top = 472
    Width = 75
    Height = 22
    Caption = 'St'#225'hni v'#253'pisy'
    TabOrder = 9
    OnClick = btnStahniVypisyClick
  end
  object pnRight: TPanel
    Left = 932
    Top = 0
    Width = 447
    Height = 675
    Align = alRight
    TabOrder = 1
    object lblPrechoziPlatbyZUctu: TLabel
      Left = 10
      Top = 16
      Width = 137
      Height = 13
      Caption = 'P'#345'edchoz'#237' platby z '#250#269'tu'
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object lbPocetPlateb: TLabel
      Left = 354
      Top = 9
      Width = 60
      Height = 13
      Caption = 'Po'#269'et plateb'
    end
    object lblPrechoziPlatbySVs: TLabel
      Left = 8
      Top = 176
      Width = 128
      Height = 13
      Caption = 'P'#345'edchoz'#237' platby s VS'
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object lblVypisOverview: TLabel
      Left = 16
      Top = 464
      Width = 86
      Height = 13
      Caption = 'lbl Vypis Overview'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object Memo1: TMemo
      Left = 16
      Top = 556
      Width = 417
      Height = 105
      ScrollBars = ssVertical
      TabOrder = 0
    end
    object asgPredchoziPlatby: TAdvStringGrid
      Left = 8
      Top = 35
      Width = 433
      Height = 127
      DefaultRowHeight = 20
      DrawingStyle = gdsClassic
      Enabled = False
      FixedCols = 0
      RowCount = 8
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing]
      ParentFont = False
      ScrollBars = ssVertical
      TabOrder = 1
      OnGetAlignment = asgPredchoziPlatbyGetAlignment
      OnButtonClick = asgPredchoziPlatbyButtonClick
      ActiveCellFont.Charset = DEFAULT_CHARSET
      ActiveCellFont.Color = clWindowText
      ActiveCellFont.Height = -11
      ActiveCellFont.Name = 'Tahoma'
      ActiveCellFont.Style = [fsBold]
      ColumnHeaders.Strings = (
        ''
        'VS'
        #268'astka'
        'Datum'
        'Firma')
      ControlLook.FixedGradientHoverFrom = clGray
      ControlLook.FixedGradientHoverTo = clWhite
      ControlLook.FixedGradientDownFrom = clGray
      ControlLook.FixedGradientDownTo = clSilver
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
      FixedColWidth = 28
      FixedRowHeight = 20
      FixedFont.Charset = DEFAULT_CHARSET
      FixedFont.Color = clWindowText
      FixedFont.Height = -11
      FixedFont.Name = 'Tahoma'
      FixedFont.Style = [fsBold]
      FloatFormat = '%.2f'
      HoverButtons.Buttons = <>
      HTMLSettings.ImageFolder = 'images'
      HTMLSettings.ImageBaseName = 'img'
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
      ScrollWidth = 16
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
      ShowDesignHelper = False
      Version = '8.7.0.0'
      ColWidths = (
        28
        78
        82
        70
        167)
    end
    object asgPredchoziPlatbyVs: TAdvStringGrid
      Left = 4
      Top = 195
      Width = 433
      Height = 131
      ColCount = 4
      DefaultRowHeight = 20
      DrawingStyle = gdsClassic
      Enabled = False
      FixedCols = 0
      RowCount = 8
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing]
      ParentFont = False
      ScrollBars = ssVertical
      TabOrder = 2
      OnGetAlignment = asgPredchoziPlatbyVsGetAlignment
      ActiveCellFont.Charset = DEFAULT_CHARSET
      ActiveCellFont.Color = clWindowText
      ActiveCellFont.Height = -11
      ActiveCellFont.Name = 'Tahoma'
      ActiveCellFont.Style = [fsBold]
      ColumnHeaders.Strings = (
        'Bank. '#250#269'et'
        #268'astka'
        'Datum'
        'Firma')
      ControlLook.FixedGradientHoverFrom = clGray
      ControlLook.FixedGradientHoverTo = clWhite
      ControlLook.FixedGradientDownFrom = clGray
      ControlLook.FixedGradientDownTo = clSilver
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
      FixedColWidth = 134
      FixedRowHeight = 20
      FixedFont.Charset = DEFAULT_CHARSET
      FixedFont.Color = clWindowText
      FixedFont.Height = -11
      FixedFont.Name = 'Tahoma'
      FixedFont.Style = [fsBold]
      FloatFormat = '%.2f'
      HoverButtons.Buttons = <>
      HTMLSettings.ImageFolder = 'images'
      HTMLSettings.ImageBaseName = 'img'
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
      ScrollWidth = 16
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
      ShowDesignHelper = False
      Version = '8.7.0.0'
      ColWidths = (
        134
        82
        64
        143)
    end
    object editPocetPredchPlateb: TEdit
      Left = 423
      Top = 6
      Width = 13
      Height = 21
      TabOrder = 3
      Text = '4'
    end
    object Memo2: TMemo
      Left = 6
      Top = 332
      Width = 435
      Height = 106
      Align = alCustom
      ScrollBars = ssVertical
      TabOrder = 4
      Visible = False
    end
  end
  object btnVypisFiokonto: TButton
    Left = 128
    Top = 188
    Width = 97
    Height = 25
    Caption = 'Fiokonto v'#253'pis'
    TabOrder = 8
    OnClick = btnVypisFiokontoClick
  end
  object pnLeft: TPanel
    Left = 0
    Top = 0
    Width = 109
    Height = 531
    TabOrder = 2
    object lblPomocPrg: TLabel
      Left = 10
      Top = 365
      Width = 91
      Height = 13
      Caption = 'Pomocn'#233' programy'
    end
    object lblZobrazit: TLabel
      Left = 8
      Top = 75
      Width = 38
      Height = 13
      Caption = 'Zobrazit'
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
    end
    object btnNacti: TButton
      Left = 10
      Top = 31
      Width = 91
      Height = 23
      Caption = '&Na'#269'ti GPC'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 0
      OnClick = btnNactiClick
    end
    object btnZapisDoAbry: TButton
      Left = 10
      Top = 233
      Width = 91
      Height = 23
      Caption = '&Do Abry'
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 1
      OnClick = btnZapisDoAbryClick
    end
    object btnSparujPlatby: TButton
      Left = 14
      Top = 500
      Width = 75
      Height = 23
      Caption = 'Sp'#225'ruj platby'
      TabOrder = 2
      Visible = False
      OnClick = btnSparujPlatbyClick
    end
    object btnReconnect: TButton
      Left = 14
      Top = 471
      Width = 75
      Height = 23
      Caption = 'Reconnect'
      TabOrder = 3
      Visible = False
      OnClick = btnReconnectClick
    end
    object chbZobrazitBezproblemove: TCheckBox
      Left = 8
      Top = 94
      Width = 93
      Height = 17
      Caption = 'bezprobl'#233'mov'#233
      Checked = True
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      State = cbChecked
      TabOrder = 4
      OnClick = chbZobrazitBezproblemoveClick
    end
    object chbZobrazitDebety: TCheckBox
      Left = 8
      Top = 128
      Width = 65
      Height = 17
      Caption = 'debety'
      Checked = True
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      State = cbChecked
      TabOrder = 5
      OnClick = chbZobrazitDebetyClick
    end
    object chbZobrazitStandardni: TCheckBox
      Left = 8
      Top = 110
      Width = 93
      Height = 17
      Caption = 'standardn'#237
      Checked = True
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      State = cbChecked
      TabOrder = 6
      OnClick = chbZobrazitStandardniClick
    end
    object btnShowPrirazeniPnpForm: TButton
      Left = 8
      Top = 384
      Width = 91
      Height = 25
      Caption = 'P'#345'i'#345'azen'#237' PNP'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlue
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 7
      OnClick = btnShowPrirazeniPnpFormClick
    end
    object btnZavritVypis: TButton
      Left = 8
      Top = 271
      Width = 91
      Height = 23
      Caption = '&Zav'#345#237't v'#253'pis'
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 8
      OnClick = btnZavritVypisClick
    end
    object btnCustomers: TButton
      Left = 8
      Top = 415
      Width = 75
      Height = 23
      Caption = 'Z'#225'kazn'#237'ci'
      TabOrder = 9
      OnClick = btnCustomersClick
    end
    object btnHledej: TButton
      Left = 8
      Top = 190
      Width = 63
      Height = 21
      Caption = 'hledej...'
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 10
      OnClick = btnHledejClick
    end
    object editHledej: TEdit
      Left = 10
      Top = 163
      Width = 94
      Height = 21
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 11
    end
  end
  object btnVypisFio: TButton
    Left = 128
    Top = 30
    Width = 97
    Height = 25
    Caption = 'Fio v'#253'pis'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 3
    OnClick = btnVypisFioClick
  end
  object btnVypisFioSporici: TButton
    Left = 128
    Top = 110
    Width = 97
    Height = 25
    Caption = 'Fio spo'#345#237'c'#237' v'#253'pis'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 4
    OnClick = btnVypisFioSporiciClick
  end
  object btnVypisCsob: TButton
    Left = 128
    Top = 270
    Width = 99
    Height = 25
    Caption = #268'SOB v'#253'pis'
    TabOrder = 5
    OnClick = btnVypisCsobClick
  end
  object btnVypisPayU: TButton
    Left = 128
    Top = 352
    Width = 97
    Height = 25
    Caption = 'PayU v'#253'pis'
    Enabled = False
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 6
    OnClick = btnVypisPayUClick
  end
  object asgMain: TAdvStringGrid
    Left = 107
    Top = 26
    Width = 822
    Height = 506
    Anchors = [akLeft, akTop, akBottom]
    ColCount = 9
    DefaultRowHeight = 20
    DrawingStyle = gdsClassic
    Enabled = False
    FixedCols = 0
    RowCount = 5
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing, goEditing]
    ParentFont = False
    TabOrder = 0
    Visible = False
    OnClick = asgMainClick
    OnKeyUp = asgMainKeyUp
    OnGetCellColor = asgMainGetCellColor
    OnGetAlignment = asgMainGetAlignment
    OnCanEditCell = asgMainCanEditCell
    OnCellValidate = asgMainCellValidate
    OnCellsChanged = asgMainCellsChanged
    OnButtonClick = asgMainButtonClick
    OnCheckBoxClick = asgMainCheckBoxClick
    ActiveCellFont.Charset = DEFAULT_CHARSET
    ActiveCellFont.Color = clWindowText
    ActiveCellFont.Height = -11
    ActiveCellFont.Name = 'Tahoma'
    ActiveCellFont.Style = [fsBold]
    ColumnHeaders.Strings = (
      ''
      #268#225'stka'
      'VS'
      'SS'
      #268'. '#250#269'tu'
      'Protistrana / popis'
      'Datum'
      ''
      'pozn.')
    ControlLook.FixedGradientHoverFrom = clGray
    ControlLook.FixedGradientHoverTo = clWhite
    ControlLook.FixedGradientDownFrom = clGray
    ControlLook.FixedGradientDownTo = clSilver
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
    FixedColWidth = 80
    FixedRowHeight = 20
    FixedFont.Charset = DEFAULT_CHARSET
    FixedFont.Color = clWindowText
    FixedFont.Height = -11
    FixedFont.Name = 'Tahoma'
    FixedFont.Style = [fsBold]
    FloatFormat = '%.2f'
    HoverButtons.Buttons = <>
    HTMLSettings.ImageFolder = 'images'
    HTMLSettings.ImageBaseName = 'img'
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
    ScrollWidth = 16
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
    Version = '8.7.0.0'
    ColWidths = (
      80
      82
      79
      22
      132
      135
      64
      22
      184)
  end
  object pnBottom: TPanel
    Left = 0
    Top = 532
    Width = 932
    Height = 143
    TabOrder = 7
    object lblNalezeneDoklady: TLabel
      Left = 108
      Top = 8
      Width = 102
      Height = 13
      Caption = 'Doklady podle VS'
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object asgNalezeneDoklady: TAdvStringGrid
      Left = 108
      Top = 27
      Width = 812
      Height = 105
      ColCount = 8
      DefaultRowHeight = 20
      DrawingStyle = gdsClassic
      Enabled = False
      FixedCols = 0
      RowCount = 7
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing]
      ParentFont = False
      TabOrder = 0
      OnGetAlignment = asgNalezeneDokladyGetAlignment
      ActiveCellFont.Charset = DEFAULT_CHARSET
      ActiveCellFont.Color = clWindowText
      ActiveCellFont.Height = -11
      ActiveCellFont.Name = 'Tahoma'
      ActiveCellFont.Style = [fsBold]
      ColumnHeaders.Strings = (
        #268#237'slo dokladu'
        'Datum'
        'Firma'
        'P'#345'edpis'
        'Zaplaceno'
        'Dobropis.'
        'Nezaplaceno')
      ControlLook.FixedGradientHoverFrom = clGray
      ControlLook.FixedGradientHoverTo = clWhite
      ControlLook.FixedGradientDownFrom = clGray
      ControlLook.FixedGradientDownTo = clSilver
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
      FixedColWidth = 89
      FixedRowHeight = 20
      FixedFont.Charset = DEFAULT_CHARSET
      FixedFont.Color = clWindowText
      FixedFont.Height = -11
      FixedFont.Name = 'Tahoma'
      FixedFont.Style = [fsBold]
      FloatFormat = '%.2f'
      HoverButtons.Buttons = <>
      HTMLSettings.ImageFolder = 'images'
      HTMLSettings.ImageBaseName = 'img'
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
      ScrollWidth = 16
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
      Version = '8.7.0.0'
      ColWidths = (
        89
        69
        158
        66
        65
        65
        78
        186)
    end
    object chbVsechnyDoklady: TCheckBox
      Left = 824
      Top = 4
      Width = 103
      Height = 17
      Caption = 'v'#353'echny doklady'
      Enabled = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      OnClick = chbVsechnyDokladyClick
    end
  end
  object NactiGpcDialog: TOpenDialog
    Left = 1292
    Top = 486
  end
end
