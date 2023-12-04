object fmSparovaniVDeniku: TfmSparovaniVDeniku
  Left = 0
  Top = 0
  Caption = 'Sp'#225'rov'#225'n'#237' v den'#237'ku'
  ClientHeight = 508
  ClientWidth = 1154
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesigned
  OnShow = FormShow
  DesignSize = (
    1154
    508)
  PixelsPerInch = 96
  TextHeight = 13
  object lblLimit: TLabel
    Left = 521
    Top = 17
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
  object Label1: TLabel
    Left = 8
    Top = 485
    Width = 967
    Height = 13
    Caption = 
      'Jsou nalezeny p'#225'ry z GL, kde polo'#382'ka A je sama ve skupin'#283' a na M' +
      'D m'#225' zadan'#253' '#250#269'et a k tomu se p'#225'ruje polo'#382'ka B, kter'#225' je tak'#233' sam' +
      'a ve skupin'#283' a zadan'#253' '#250#269'et m'#225' na D. Polo'#382'ky mus'#237' pat'#345'it stejn'#233' f' +
      'irm'#283'.'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
  end
  object asgSparovaniVDeniku: TAdvStringGrid
    Left = 0
    Top = 48
    Width = 1133
    Height = 431
    Anchors = [akLeft, akTop, akBottom]
    BorderStyle = bsNone
    ColCount = 13
    DefaultRowHeight = 19
    DrawingStyle = gdsClassic
    FixedCols = 0
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 0
    OnGetAlignment = asgSparovaniVDenikuGetAlignment
    OnCanSort = asgSparovaniVDenikuCanSort
    OnClickCell = asgSparovaniVDenikuClickCell
    OnCanEditCell = asgSparovaniVDenikuCanEditCell
    ActiveCellFont.Charset = DEFAULT_CHARSET
    ActiveCellFont.Color = clWindowText
    ActiveCellFont.Height = -11
    ActiveCellFont.Name = 'Tahoma'
    ActiveCellFont.Style = [fsBold]
    ColumnHeaders.Strings = (
      ''
      'Jm'#233'no'
      'A datum'
      'A text'
      'A '#269#225'stka'
      'A proti'#250#269'et'
      'B datum'
      'B text'
      'B '#269#225'stka'
      'B proti'#250#269'et'
      'Rozd'#237'l'
      'AccGroupId'
      'ID z'#225'znamu')
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
    FixedColWidth = 27
    FixedRowHeight = 19
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
    SearchFooter.ResultFormat = '(%d of %d)'
    SortSettings.Column = 2
    SortSettings.Show = True
    Version = '8.7.0.0'
    ColWidths = (
      27
      160
      64
      150
      62
      70
      64
      150
      62
      68
      64
      71
      89)
  end
  object btnProvedSparovani: TButton
    Left = 768
    Top = 8
    Width = 121
    Height = 25
    Caption = 'Prove'#271' sp'#225'rov'#225'n'#237
    TabOrder = 1
    OnClick = btnProvedSparovaniClick
  end
  object btnNactiData: TButton
    Left = 8
    Top = 8
    Width = 109
    Height = 25
    Caption = 'Na'#269'ti data pro '#250#269'et:'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 2
    OnClick = btnNactiDataClick
  end
  object editKodUctu: TEdit
    Left = 121
    Top = 11
    Width = 44
    Height = 22
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 3
    Text = '32500'
  end
  object chbPouzeNulovyRozdil: TCheckBox
    Left = 194
    Top = 14
    Width = 119
    Height = 17
    Caption = 'pouze nulov'#253' rozd'#237'l'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 4
  end
  object chbShodneProtiucty: TCheckBox
    Left = 335
    Top = 14
    Width = 149
    Height = 17
    Caption = 'proti'#250#269'ty mus'#237' b'#253't shodn'#233
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 5
  end
  object editLimit: TEdit
    Left = 580
    Top = 14
    Width = 53
    Height = 21
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 6
    Text = '100'
  end
end
