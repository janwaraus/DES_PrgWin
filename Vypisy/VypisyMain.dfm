object fmMain: TfmMain
  Left = 141
  Top = 0
  Caption = 'Na'#269'ten'#237', oprava a ulo'#382'en'#237' bankovn'#237'ho v'#253'pisu'
  ClientHeight = 801
  ClientWidth = 1387
  Color = clBtnFace
  Constraints.MinHeight = 721
  Constraints.MinWidth = 1312
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnShow = FormShow
  DesignSize = (
    1387
    801)
  PixelsPerInch = 96
  TextHeight = 13
  object lblHlavicka: TLabel
    Left = 112
    Top = 5
    Width = 15
    Height = 16
    Caption = '...'
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
    Top = 256
    Width = 69
    Height = 13
    Caption = 'lblVypisFioGpc'
  end
  object lblVypisFioInfo: TLabel
    Left = 240
    Top = 288
    Width = 67
    Height = 13
    Caption = 'lblVypisFioInfo'
  end
  object lblVypisFioSporiciGpc: TLabel
    Left = 240
    Top = 352
    Width = 69
    Height = 13
    Caption = 'lblVypisFioGpc'
  end
  object lblVypisFioSporiciInfo: TLabel
    Left = 240
    Top = 392
    Width = 67
    Height = 13
    Caption = 'lblVypisFioInfo'
  end
  object lblVypisCsobInfo: TLabel
    Left = 240
    Top = 496
    Width = 77
    Height = 13
    Caption = 'lblVypisCsobInfo'
  end
  object lblVypisCsobGpc: TLabel
    Left = 240
    Top = 456
    Width = 79
    Height = 13
    Caption = 'lblVypisCsobGpc'
  end
  object lblVypisPayuGpc: TLabel
    Left = 240
    Top = 563
    Width = 79
    Height = 13
    Caption = 'lblVypisPayuGpc'
  end
  object lblVypisPayuInfo: TLabel
    Left = 240
    Top = 600
    Width = 77
    Height = 13
    Caption = 'lblVypisPayuInfo'
  end
  object pnBottom: TPanel
    Left = 0
    Top = 655
    Width = 1387
    Height = 146
    Align = alBottom
    TabOrder = 1
    object lblNalezeneDoklady: TLabel
      Left = 10
      Top = 6
      Width = 102
      Height = 13
      Caption = 'Doklady podle VS'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -12
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object asgNalezeneDoklady: TAdvStringGrid
      Left = -1
      Top = 24
      Width = 812
      Height = 105
      Cursor = crDefault
      ColCount = 9
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
      ScrollBars = ssBoth
      TabOrder = 0
      HoverRowCells = [hcNormal, hcSelected]
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
        'Nezaplaceno'
        'ID Dokladu')
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
      HoverButtons.Position = hbLeftFromColumnLeft
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
      SortSettings.DefaultFormat = ssAutomatic
      Version = '7.4.2.0'
      ColWidths = (
        89
        69
        158
        66
        65
        65
        78
        74
        110)
    end
    object chbVsechnyDoklady: TCheckBox
      Left = 270
      Top = 4
      Width = 103
      Height = 17
      Caption = 'v'#353'echny doklady'
      TabOrder = 1
      OnClick = chbVsechnyDokladyClick
    end
  end
  object pnRight: TPanel
    Left = 940
    Top = 0
    Width = 447
    Height = 655
    Align = alRight
    TabOrder = 2
    object lblPrechoziPlatbyZUctu: TLabel
      Left = 6
      Top = 10
      Width = 129
      Height = 13
      Caption = 'P'#345'echoz'#237' platby z '#250#269'tu'
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
      Left = 2
      Top = 217
      Width = 120
      Height = 13
      Caption = 'P'#345'echoz'#237' platby s VS'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Memo1: TMemo
      Left = 4
      Top = 408
      Width = 435
      Height = 131
      TabOrder = 0
    end
    object asgPredchoziPlatby: TAdvStringGrid
      Left = 0
      Top = 28
      Width = 417
      Height = 183
      Cursor = crDefault
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
      ScrollBars = ssBoth
      TabOrder = 1
      HoverRowCells = [hcNormal, hcSelected]
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
      HoverButtons.Position = hbLeftFromColumnLeft
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
      SortSettings.DefaultFormat = ssAutomatic
      Version = '7.4.2.0'
      ColWidths = (
        28
        72
        81
        64
        161)
    end
    object asgPredchoziPlatbyVs: TAdvStringGrid
      Left = 2
      Top = 233
      Width = 444
      Height = 169
      Cursor = crDefault
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
      ScrollBars = ssBoth
      TabOrder = 2
      HoverRowCells = [hcNormal, hcSelected]
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
      FixedColWidth = 137
      FixedRowHeight = 20
      FixedFont.Charset = DEFAULT_CHARSET
      FixedFont.Color = clWindowText
      FixedFont.Height = -11
      FixedFont.Name = 'Tahoma'
      FixedFont.Style = [fsBold]
      FloatFormat = '%.2f'
      HoverButtons.Buttons = <>
      HoverButtons.Position = hbLeftFromColumnLeft
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
      SortSettings.DefaultFormat = ssAutomatic
      Version = '7.4.2.0'
      ColWidths = (
        137
        82
        65
        150)
    end
    object editPocetPredchPlateb: TEdit
      Left = 423
      Top = 6
      Width = 13
      Height = 21
      TabOrder = 3
      Text = '7'
    end
    object Memo2: TMemo
      Left = 2
      Top = 545
      Width = 436
      Height = 216
      Lines.Strings = (
        'Memo2')
      TabOrder = 4
    end
  end
  object pnLeft: TPanel
    Left = 0
    Top = 0
    Width = 109
    Height = 655
    Align = alLeft
    TabOrder = 3
    object lbZobrazit: TLabel
      Left = 8
      Top = 72
      Width = 38
      Height = 13
      Caption = 'Zobrazit'
    end
    object btnNacti: TButton
      Left = 8
      Top = 26
      Width = 81
      Height = 21
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
      Left = 8
      Top = 233
      Width = 79
      Height = 21
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
      Left = 12
      Top = 562
      Width = 75
      Height = 17
      Caption = 'Sp'#225'ruj platby'
      TabOrder = 2
      Visible = False
      OnClick = btnSparujPlatbyClick
    end
    object btnReconnect: TButton
      Left = 24
      Top = 530
      Width = 65
      Height = 25
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
      State = cbChecked
      TabOrder = 6
      OnClick = chbZobrazitStandardniClick
    end
    object btnShowPrirazeniPnpForm: TButton
      Left = 15
      Top = 377
      Width = 89
      Height = 25
      Caption = 'P'#345'i'#345'azen'#237' PNP'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlue
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 7
      OnClick = btnShowPrirazeniPnpFormClick
    end
    object btnZavritVypis: TButton
      Left = 8
      Top = 283
      Width = 93
      Height = 25
      Caption = 'Zav'#345#237't v'#253'pis'
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
      Left = 14
      Top = 424
      Width = 75
      Height = 25
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
      TabOrder = 10
      OnClick = btnHledejClick
    end
    object editHledej: TEdit
      Left = 8
      Top = 163
      Width = 94
      Height = 21
      TabOrder = 11
    end
  end
  object btnVypisFio: TButton
    Left = 128
    Top = 256
    Width = 97
    Height = 25
    Caption = 'FIO v'#253'pis'
    TabOrder = 4
    OnClick = btnVypisFioClick
  end
  object btnVypisFioSporici: TButton
    Left = 128
    Top = 352
    Width = 97
    Height = 25
    Caption = 'FIO spo'#345#237'c'#237' v'#253'pis'
    TabOrder = 5
    OnClick = btnVypisFioSporiciClick
  end
  object btnVypisCsob: TButton
    Left = 128
    Top = 456
    Width = 99
    Height = 25
    Caption = #268'SOB v'#253'pis'
    TabOrder = 6
    OnClick = btnVypisCsobClick
  end
  object btnVypisPayU: TButton
    Left = 128
    Top = 560
    Width = 97
    Height = 25
    Caption = 'PayU v'#253'pis'
    Enabled = False
    TabOrder = 7
    OnClick = btnVypisPayUClick
  end
  object asgMain: TAdvStringGrid
    Left = 112
    Top = 29
    Width = 824
    Height = 624
    Cursor = crDefault
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
    ScrollBars = ssBoth
    TabOrder = 0
    Visible = False
    OnClick = asgMainClick
    OnKeyUp = asgMainKeyUp
    HoverRowCells = [hcNormal, hcSelected]
    OnGetCellColor = asgMainGetCellColor
    OnGetAlignment = asgMainGetAlignment
    OnCanEditCell = asgMainCanEditCell
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
      'Text'
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
    HoverButtons.Position = hbLeftFromColumnLeft
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
    SortSettings.DefaultFormat = ssAutomatic
    Version = '7.4.2.0'
    ColWidths = (
      80
      82
      78
      22
      132
      135
      64
      22
      195)
  end
  object NactiGpcDialog: TOpenDialog
    Left = 156
    Top = 166
  end
end
