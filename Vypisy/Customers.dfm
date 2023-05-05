object fmCustomers: TfmCustomers
  Left = 374
  Top = 132
  Caption = 'Z'#225'kazn'#237'ci v tabulce Customers'
  ClientHeight = 102
  ClientWidth = 891
  Color = clBtnFace
  Constraints.MinHeight = 134
  Constraints.MinWidth = 650
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnCreate = FormCreate
  OnKeyPress = FormKeyPress
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object lbJmeno: TLabel
    Left = 10
    Top = 14
    Width = 37
    Height = 13
    Caption = 'Jm'#233'no'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = 8269108
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object lbPrijmeni: TLabel
    Left = 159
    Top = 14
    Width = 92
    Height = 13
    Caption = 'P'#345#237'jmen'#237' / Firma'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = 8269108
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object lbVS: TLabel
    Left = 370
    Top = 14
    Width = 77
    Height = 13
    Caption = 'VS / smlouva'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = 8269108
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object edJmeno: TEdit
    Left = 54
    Top = 10
    Width = 83
    Height = 21
    TabOrder = 0
    Text = '*'
    OnKeyUp = edJmenoKeyUp
  end
  object edPrijmeni: TEdit
    Left = 260
    Top = 10
    Width = 95
    Height = 21
    TabOrder = 1
    Text = '*'
    OnKeyUp = edPrijmeniKeyUp
  end
  object btnNajdi: TButton
    Left = 555
    Top = 10
    Width = 65
    Height = 21
    Caption = '&Najdi'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 2
    OnClick = btnNajdiClick
  end
  object asgCustomers: TAdvStringGrid
    Left = 0
    Top = 42
    Width = 899
    Height = 72
    Cursor = crDefault
    Align = alCustom
    Anchors = [akLeft, akTop, akRight, akBottom]
    ColCount = 9
    DefaultRowHeight = 18
    DrawingStyle = gdsClassic
    Enabled = False
    FixedCols = 0
    RowCount = 2
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing, goEditing]
    ParentFont = False
    ScrollBars = ssBoth
    TabOrder = 3
    HoverRowCells = [hcNormal, hcSelected]
    OnGetAlignment = asgCustomersGetAlignment
    ActiveCellFont.Charset = DEFAULT_CHARSET
    ActiveCellFont.Color = clWindowText
    ActiveCellFont.Height = -11
    ActiveCellFont.Name = 'Tahoma'
    ActiveCellFont.Style = [fsBold]
    ColumnHeaders.Strings = (
      'Z'#225'kazn'#237'k'
      'Abrak'#243'd'
      'VS'
      'Smlouva'
      'Stav'
      'Fakturovat'
      #268#237'slo dokladu'
      'Datum'
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
    FixedColWidth = 185
    FixedRowHeight = 18
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
    ShowDesignHelper = False
    SortSettings.DefaultFormat = ssAutomatic
    Version = '7.4.2.0'
    ColWidths = (
      185
      90
      81
      106
      90
      90
      87
      64
      88)
  end
  object edVS: TEdit
    Left = 453
    Top = 10
    Width = 85
    Height = 21
    TabOrder = 4
    Text = '*'
    OnKeyUp = edVSKeyUp
  end
  object chbJenSNezaplacenym: TCheckBox
    Left = 631
    Top = 13
    Width = 162
    Height = 15
    Caption = 'jen s nezaplacen'#253'm dokladem'
    TabOrder = 5
  end
  object btnReset: TBitBtn
    Left = 821
    Top = 8
    Width = 57
    Height = 23
    Caption = '&Reset'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = 5790335
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 6
    OnClick = btnResetClick
  end
  object infoPanel: TPanel
    Left = 307
    Top = 41
    Width = 265
    Height = 41
    TabOrder = 7
    Visible = False
    object infoPanelLabel: TLabel
      Left = 19
      Top = 14
      Width = 70
      Height = 13
      Caption = 'infoPanelLabel'
    end
  end
end
