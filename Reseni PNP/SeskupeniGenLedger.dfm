object fmSeskupeniVDeniku: TfmSeskupeniVDeniku
  Left = 0
  Top = 0
  Caption = 'SeskupeniGenLedger'
  ClientHeight = 762
  ClientWidth = 1338
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  DesignSize = (
    1338
    762)
  PixelsPerInch = 96
  TextHeight = 13
  object lblLimit: TLabel
    Left = 717
    Top = 13
    Width = 55
    Height = 13
    Caption = 'Limit '#345#225'dk'#367':'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object lblFirma: TLabel
    Left = 217
    Top = 14
    Width = 30
    Height = 13
    Caption = 'Firma:'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object lblPomlcka: TLabel
    Left = 352
    Top = 51
    Width = 11
    Height = 13
    Caption = 'a'#382
  end
  object Label1: TLabel
    Left = 181
    Top = 51
    Width = 66
    Height = 13
    Caption = 'Datum od-do:'
  end
  object btnNactiData: TButton
    Left = 8
    Top = 8
    Width = 113
    Height = 25
    Caption = 'Na'#269'ti data pro '#250#269'et:'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 0
    OnClick = btnNactiDataClick
  end
  object btnProvedSeskupeni: TButton
    Left = 864
    Top = 8
    Width = 107
    Height = 25
    Caption = 'Prove'#271' seskupen'#237
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 1
    OnClick = btnProvedSeskupeniClick
  end
  object editKodUctu: TEdit
    Left = 125
    Top = 11
    Width = 60
    Height = 22
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 2
    Text = '32500'
  end
  object chb2: TCheckBox
    Left = 521
    Top = 12
    Width = 153
    Height = 17
    Caption = 'zobrazit i samostatn'#233' '#345#225'dky'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 3
  end
  object asgSeskupeniVDeniku: TAdvStringGrid
    Left = 8
    Top = 88
    Width = 1321
    Height = 666
    Cursor = crDefault
    Anchors = [akLeft, akTop, akBottom]
    BorderStyle = bsNone
    ColCount = 12
    DefaultRowHeight = 19
    DrawingStyle = gdsClassic
    FixedCols = 0
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing]
    ParentFont = False
    ScrollBars = ssBoth
    TabOrder = 4
    HoverRowCells = [hcNormal, hcSelected]
    OnGetAlignment = asgSeskupeniVDenikuGetAlignment
    OnCanSort = asgSeskupeniVDenikuCanSort
    OnClickCell = asgSeskupeniVDenikuClickCell
    OnCanEditCell = asgSeskupeniVDenikuCanEditCell
    ActiveCellFont.Charset = DEFAULT_CHARSET
    ActiveCellFont.Color = clWindowText
    ActiveCellFont.Height = -11
    ActiveCellFont.Name = 'Tahoma'
    ActiveCellFont.Style = [fsBold]
    ColumnHeaders.Strings = (
      ''
      'Jm'#233'no'
      'Text'
      #268#225'stka'
      #218#269'tov'#225'no'
      'Debit - MD'
      'Credit - D'
      'Doklad'
      'LastFirm ID'
      'AccGroupId'
      'GenLedger ID'
      'Sou'#269'et skupiny')
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
    ControlLook.DropDownFooter.Font.Name = 'Tahoma'
    ControlLook.DropDownFooter.Font.Style = []
    ControlLook.DropDownFooter.Visible = True
    ControlLook.DropDownFooter.Buttons = <>
    Filter = <>
    FilterDropDown.Font.Charset = DEFAULT_CHARSET
    FilterDropDown.Font.Color = clWindowText
    FilterDropDown.Font.Height = -11
    FilterDropDown.Font.Name = 'Tahoma'
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
    FixedColWidth = 20
    FixedRowHeight = 19
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
    PrintSettings.Font.Name = 'Tahoma'
    PrintSettings.Font.Style = []
    PrintSettings.FixedFont.Charset = DEFAULT_CHARSET
    PrintSettings.FixedFont.Color = clWindowText
    PrintSettings.FixedFont.Height = -11
    PrintSettings.FixedFont.Name = 'Tahoma'
    PrintSettings.FixedFont.Style = []
    PrintSettings.HeaderFont.Charset = DEFAULT_CHARSET
    PrintSettings.HeaderFont.Color = clWindowText
    PrintSettings.HeaderFont.Height = -11
    PrintSettings.HeaderFont.Name = 'Tahoma'
    PrintSettings.HeaderFont.Style = []
    PrintSettings.FooterFont.Charset = DEFAULT_CHARSET
    PrintSettings.FooterFont.Color = clWindowText
    PrintSettings.FooterFont.Height = -11
    PrintSettings.FooterFont.Name = 'Tahoma'
    PrintSettings.FooterFont.Style = []
    PrintSettings.PageNumSep = '/'
    SearchFooter.FindNextCaption = 'Find &next'
    SearchFooter.FindPrevCaption = 'Find &previous'
    SearchFooter.Font.Charset = DEFAULT_CHARSET
    SearchFooter.Font.Color = clWindowText
    SearchFooter.Font.Height = -11
    SearchFooter.Font.Name = 'Tahoma'
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
      20
      278
      231
      73
      72
      67
      59
      84
      71
      72
      84
      198)
  end
  object editLimit: TEdit
    Left = 776
    Top = 9
    Width = 65
    Height = 21
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 5
    Text = '100'
  end
  object editFirmName: TEdit
    Left = 253
    Top = 9
    Width = 121
    Height = 21
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 6
    Text = '%'
  end
  object rb1: TRadioButton
    Left = 1093
    Top = 15
    Width = 58
    Height = 17
    Caption = 'Metoda1'
    Checked = True
    Enabled = False
    TabOrder = 7
    TabStop = True
  end
  object rb2: TRadioButton
    Left = 1165
    Top = 15
    Width = 64
    Height = 17
    Caption = 'Metoda2'
    Enabled = False
    TabOrder = 8
  end
  object dtpDatumOd: TDateTimePicker
    Left = 253
    Top = 48
    Width = 93
    Height = 21
    Date = 42736.974757187500000000
    Time = 42736.974757187500000000
    TabOrder = 9
  end
  object dtpDatumDo: TDateTimePicker
    Left = 376
    Top = 48
    Width = 97
    Height = 21
    Date = 43100.974876666670000000
    Time = 43100.974876666670000000
    TabOrder = 10
  end
  object chb3: TCheckBox
    Left = 521
    Top = 50
    Width = 153
    Height = 17
    Caption = 'zobrazit i nulov'#233' sou'#269'ty'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 11
  end
end
