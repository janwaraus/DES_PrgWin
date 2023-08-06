object fmMain: TfmMain
  Left = 477
  Top = 140
  Caption = 'Zaplacen'#233' a nez'#250#269'tovan'#233' z'#225'lohov'#233' listy za p'#345'ipojen'#237' k internetu'
  ClientHeight = 281
  ClientWidth = 825
  Color = clBtnFace
  Constraints.MinHeight = 320
  Constraints.MinWidth = 680
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnActivate = FormActivate
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object apnTop: TAdvPanel
    Left = 0
    Top = 0
    Width = 825
    Height = 132
    Align = alTop
    Color = 16772326
    Constraints.MinHeight = 132
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlue
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 2
    UseDockManager = True
    Version = '2.6.3.2'
    BorderColor = clBlack
    Caption.Color = clHighlight
    Caption.ColorTo = clNone
    Caption.Font.Charset = DEFAULT_CHARSET
    Caption.Font.Color = clTeal
    Caption.Font.Height = -11
    Caption.Font.Name = 'MS Sans Serif'
    Caption.Font.Style = [fsBold]
    Caption.Indent = 0
    DoubleBuffered = True
    StatusBar.Font.Charset = DEFAULT_CHARSET
    StatusBar.Font.Color = clWindowText
    StatusBar.Font.Height = -11
    StatusBar.Font.Name = 'Tahoma'
    StatusBar.Font.Style = []
    Text = ''
    DesignSize = (
      825
      132)
    FullHeight = 200
    object apbProgress: TAdvProgressBar
      Left = 186
      Top = 103
      Width = 630
      Height = 22
      Anchors = [akLeft, akTop, akRight]
      BackgroundColor = 16772326
      Font.Charset = EASTEUROPE_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      Level0Color = clTeal
      Level0ColorTo = 14811105
      Level1ColorTo = 13303807
      Level2Color = 5483007
      Level2ColorTo = 11064319
      Level3ColorTo = 13290239
      Level1Perc = 101
      Level2Perc = 101
      Position = 50
      ShowBorder = True
      Version = '1.3.2.3'
      Visible = False
      ExplicitWidth = 469
    end
    object lbPozor1: TLabel
      Left = 200
      Top = 110
      Width = 385
      Height = 13
      Caption = 
        'Program vystav'#237' podle ZL fakturu a p'#345'ipoj'#237' platbu z'#225'lohov'#253'm list' +
        'em.'
    end
    object apnMail: TAdvPanel
      Left = 168
      Top = 95
      Width = 664
      Height = 37
      Anchors = [akLeft, akTop, akRight]
      ParentColor = True
      TabOrder = 3
      UseDockManager = True
      Visible = False
      Version = '2.6.3.2'
      BorderColor = clBlack
      Caption.Color = clHighlight
      Caption.ColorTo = clNone
      Caption.Font.Charset = DEFAULT_CHARSET
      Caption.Font.Color = clBlue
      Caption.Font.Height = -11
      Caption.Font.Name = 'MS Sans Serif'
      Caption.Font.Style = [fsBold]
      Caption.Indent = 0
      DoubleBuffered = True
      StatusBar.Font.Charset = DEFAULT_CHARSET
      StatusBar.Font.Color = clWindowText
      StatusBar.Font.Height = -11
      StatusBar.Font.Name = 'Tahoma'
      StatusBar.Font.Style = []
      Text = ''
      DesignSize = (
        664
        37)
      FullHeight = 95
      object fePriloha: TAdvFileNameEdit
        Left = 211
        Top = 8
        Width = 438
        Height = 21
        EmptyTextStyle = []
        LabelCaption = 'p'#345#237'loha'
        LabelPosition = lpLeftCenter
        LabelMargin = 24
        LabelFont.Charset = DEFAULT_CHARSET
        LabelFont.Color = clTeal
        LabelFont.Height = -11
        LabelFont.Name = 'MS Sans Serif'
        LabelFont.Style = [fsBold]
        Lookup.Font.Charset = DEFAULT_CHARSET
        Lookup.Font.Color = clWindowText
        Lookup.Font.Height = -11
        Lookup.Font.Name = 'Arial'
        Lookup.Font.Style = []
        Lookup.Separator = ';'
        Anchors = [akLeft, akTop, akRight]
        Color = clWindow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clTeal
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        ShortCut = 0
        TabOrder = 0
        Text = ''
        Visible = True
        Version = '1.7.1.3'
        ButtonStyle = bsButton
        ButtonWidth = 18
        Flat = False
        Etched = False
        Glyph.Data = {
          CE000000424DCE0000000000000076000000280000000C0000000B0000000100
          0400000000005800000000000000000000001000000000000000000000000000
          8000008000000080800080000000800080008080000080808000C0C0C0000000
          FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00D00000000DDD
          00000077777770DD00000F077777770D00000FF07777777000000FFF00000000
          00000FFFFFFF0DDD00000FFF00000DDD0000D000DDDDD0000000DDDDDDDDDD00
          0000DDDDD0DDD0D00000DDDDDD000DDD0000}
        ReadOnly = False
        FilterIndex = 0
        DialogOptions = []
        DialogKind = fdOpen
      end
    end
    object apnPrevod: TAdvPanel
      Left = 168
      Top = 96
      Width = 664
      Height = 37
      Anchors = [akLeft, akTop, akRight]
      ParentColor = True
      TabOrder = 1
      UseDockManager = True
      Visible = False
      Version = '2.6.3.2'
      BorderColor = clBlack
      Caption.Color = clHighlight
      Caption.ColorTo = clNone
      Caption.Font.Charset = DEFAULT_CHARSET
      Caption.Font.Color = clBlue
      Caption.Font.Height = -11
      Caption.Font.Name = 'MS Sans Serif'
      Caption.Font.Style = [fsBold]
      Caption.Indent = 0
      DoubleBuffered = True
      StatusBar.Font.Charset = DEFAULT_CHARSET
      StatusBar.Font.Color = clWindowText
      StatusBar.Font.Height = -11
      StatusBar.Font.Name = 'Tahoma'
      StatusBar.Font.Style = []
      Text = ''
      DesignSize = (
        664
        37)
      FullHeight = 95
      object cbNeprepisovat: TCheckBox
        Left = 67
        Top = 10
        Width = 237
        Height = 18
        Hint = 'Nep'#345'episovat u'#382' d'#345#237've vytvo'#345'en'#233' soubory DPF'
        Caption = ' nep'#345'episovat existuj'#237'c'#237' soubory PDF'
        Checked = True
        State = cbChecked
        TabOrder = 0
      end
      object btOdeslat: TButton
        Left = 578
        Top = 8
        Width = 71
        Height = 21
        Hint = 'Odesl'#225'n'#237' p'#345'eveden'#253'ch faktur na server'
        Anchors = [akTop, akRight]
        Caption = '&Na server'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 1
        OnClick = btOdeslatClick
      end
      object btSablona: TButton
        Left = 469
        Top = 8
        Width = 71
        Height = 21
        Hint = #218'prava '#353'ablony faktury'
        Anchors = [akTop, akRight]
        Caption = #352'&ablona'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 2
        OnClick = btSablonaClick
      end
    end
    object apnTisk: TAdvPanel
      Left = 168
      Top = 95
      Width = 664
      Height = 37
      Anchors = [akLeft, akTop, akRight]
      ParentColor = True
      TabOrder = 2
      UseDockManager = True
      Visible = False
      Version = '2.6.3.2'
      BorderColor = clBlack
      Caption.Color = clHighlight
      Caption.ColorTo = clNone
      Caption.Font.Charset = DEFAULT_CHARSET
      Caption.Font.Color = clBlue
      Caption.Font.Height = -11
      Caption.Font.Name = 'MS Sans Serif'
      Caption.Font.Style = [fsBold]
      Caption.Indent = 0
      DoubleBuffered = True
      StatusBar.Font.Charset = DEFAULT_CHARSET
      StatusBar.Font.Color = clWindowText
      StatusBar.Font.Height = -11
      StatusBar.Font.Name = 'Tahoma'
      StatusBar.Font.Style = []
      Text = ''
      FullHeight = 95
    end
    object apnMain: TAdvPanel
      Left = 168
      Top = 0
      Width = 664
      Height = 96
      Anchors = [akLeft, akTop, akRight]
      ParentColor = True
      TabOrder = 0
      UseDockManager = True
      Version = '2.6.3.2'
      BorderColor = clBlack
      BorderShadow = True
      Caption.Color = clHighlight
      Caption.ColorTo = clNone
      Caption.Font.Charset = DEFAULT_CHARSET
      Caption.Font.Color = clBlue
      Caption.Font.Height = -11
      Caption.Font.Name = 'MS Sans Serif'
      Caption.Font.Style = [fsBold]
      Caption.Indent = 0
      DoubleBuffered = True
      StatusBar.Font.Charset = DEFAULT_CHARSET
      StatusBar.Font.Color = clWindowText
      StatusBar.Font.Height = -11
      StatusBar.Font.Name = 'Tahoma'
      StatusBar.Font.Style = []
      Text = ''
      DesignSize = (
        664
        96)
      FullHeight = 200
      object aseRok: TAdvSpinEdit
        Left = 84
        Top = 54
        Width = 55
        Height = 21
        Color = clWindow
        Value = 2007
        FloatValue = 2007.000000000000000000
        HexDigits = 0
        HexValue = 7
        EditAlign = eaRight
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        IncrementFloat = 0.100000000000000000
        IncrementFloatPage = 1.000000000000000000
        LabelCaption = 'rok ZL'
        LabelPosition = lpLeftCenter
        LabelMargin = 8
        LabelFont.Charset = DEFAULT_CHARSET
        LabelFont.Color = clBlue
        LabelFont.Height = -12
        LabelFont.Name = 'MS Sans Serif'
        LabelFont.Style = [fsBold]
        MaxValue = 2020
        MinValue = 2007
        ParentFont = False
        TabOrder = 0
        Visible = True
        Version = '2.0.1.2'
        OnChange = aseRokChange
      end
      object btVytvorit: TButton
        Left = 578
        Top = 20
        Width = 71
        Height = 21
        Anchors = [akTop, akRight]
        Caption = '&Na'#269#237'st'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 1
        OnClick = btVytvoritClick
      end
      object btKonec: TButton
        Left = 578
        Top = 54
        Width = 71
        Height = 21
        Anchors = [akTop, akRight]
        Caption = '&Konec'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 2
        OnClick = btKonecClick
      end
      object deDatumDokladu: TAdvDateTimePicker
        Left = 274
        Top = 58
        Width = 92
        Height = 21
        Date = 41908.472326388890000000
        Format = 'dd.MM.yyyy'
        Time = 41908.472326388890000000
        DoubleBuffered = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Kind = dkDate
        ParentDoubleBuffered = False
        ParentFont = False
        TabOrder = 3
        OnChange = deDatumDokladuChange
        BorderStyle = bsSingle
        Ctl3D = True
        DateTime = 41908.472326388890000000
        Version = '1.3.6.5'
        LabelCaption = 'datum dokladu'
        LabelPosition = lpLeftCenter
        LabelMargin = 8
        LabelFont.Charset = DEFAULT_CHARSET
        LabelFont.Color = clBlue
        LabelFont.Height = -11
        LabelFont.Name = 'MS Sans Serif'
        LabelFont.Style = [fsBold]
      end
      object cbCast: TCheckBox
        Left = 186
        Top = 22
        Width = 180
        Height = 19
        Caption = 'i jen '#269#225'ste'#269'n'#283' zaplacen'#233'   '
        TabOrder = 4
      end
      object acbRada: TAdvComboBox
        Left = 84
        Top = 20
        Width = 55
        Height = 21
        Color = clWindow
        Version = '2.0.0.6'
        Visible = True
        BevelEdges = [beLeft, beTop]
        BorderColor = clGray
        ButtonWidth = 17
        FlatLineColor = clGray
        FlatParentColor = False
        EmptyTextStyle = []
        Ctl3D = True
        DropWidth = 0
        Enabled = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ItemIndex = -1
        Items.Strings = (
          'Mn'#237#353'ek'
          'O'#345'ech'
          'Stod'#367'lky'
          #381'i'#382'kov')
        LabelCaption = #345'ada ZL'
        LabelPosition = lpLeftCenter
        LabelMargin = 8
        LabelAlwaysEnabled = True
        LabelFont.Charset = DEFAULT_CHARSET
        LabelFont.Color = clBlue
        LabelFont.Height = -11
        LabelFont.Name = 'MS Sans Serif'
        LabelFont.Style = [fsBold]
        ParentCtl3D = False
        ParentFont = False
        ParentShowHint = False
        ShowHint = False
        TabOrder = 5
        Text = '%'
      end
    end
    object apnVyberCinnosti: TAdvPanel
      Left = 0
      Top = 0
      Width = 169
      Height = 132
      Align = alLeft
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlue
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentColor = True
      ParentFont = False
      TabOrder = 4
      UseDockManager = True
      Version = '2.6.3.2'
      BorderColor = clBlack
      Caption.Color = clHighlight
      Caption.ColorTo = clNone
      Caption.Font.Charset = DEFAULT_CHARSET
      Caption.Font.Color = clBlue
      Caption.Font.Height = -11
      Caption.Font.Name = 'MS Sans Serif'
      Caption.Font.Style = [fsBold]
      Caption.Indent = 0
      DoubleBuffered = True
      StatusBar.Font.Charset = DEFAULT_CHARSET
      StatusBar.Font.Color = clWindowText
      StatusBar.Font.Height = -11
      StatusBar.Font.Name = 'Tahoma'
      StatusBar.Font.Style = []
      Text = ''
      FullHeight = 200
      object glbVytvoreni: TGradientLabel
        Left = 36
        Top = 0
        Width = 133
        Height = 33
        AutoSize = False
        Caption = '  Vytvo'#345'en'#237' faktur'
        Color = clWhite
        ParentColor = False
        Layout = tlCenter
        OnClick = glbVytvoreniClick
        ColorTo = clMenu
        EllipsType = etNone
        GradientType = gtFullHorizontal
        GradientDirection = gdLeftToRight
        Indent = 0
        Orientation = goHorizontal
        TransparentText = False
        VAlignment = vaCenter
        Version = '1.2.1.0'
      end
      object glbPrevod: TGradientLabel
        Left = 36
        Top = 33
        Width = 133
        Height = 33
        AutoSize = False
        Caption = '  P'#345'evod do PDF'
        Color = clSilver
        ParentColor = False
        Layout = tlCenter
        OnClick = glbPrevodClick
        ColorTo = clGray
        EllipsType = etNone
        GradientType = gtFullHorizontal
        GradientDirection = gdLeftToRight
        Indent = 0
        Orientation = goHorizontal
        TransparentText = False
        VAlignment = vaCenter
        Version = '1.2.1.0'
      end
      object glbTisk: TGradientLabel
        Left = 36
        Top = 66
        Width = 133
        Height = 33
        AutoSize = False
        Caption = '  Tisk'
        Color = clSilver
        ParentColor = False
        Layout = tlCenter
        OnClick = glbTiskClick
        ColorTo = clGray
        EllipsType = etNone
        GradientType = gtFullHorizontal
        GradientDirection = gdLeftToRight
        Indent = 0
        Orientation = goHorizontal
        TransparentText = False
        VAlignment = vaCenter
        Version = '1.2.1.0'
      end
      object glbMail: TGradientLabel
        Left = 36
        Top = 99
        Width = 133
        Height = 33
        AutoSize = False
        Caption = '  Mail'
        Color = clSilver
        ParentColor = False
        Layout = tlCenter
        OnClick = glbMailClick
        ColorTo = clGray
        EllipsType = etNone
        GradientType = gtFullHorizontal
        GradientDirection = gdLeftToRight
        Indent = 0
        Orientation = goHorizontal
        TransparentText = False
        VAlignment = vaCenter
        Version = '1.2.1.0'
      end
      object arbVytvoreni: TAdvOfficeRadioButton
        Left = 11
        Top = 8
        Width = 17
        Height = 17
        TabOrder = 0
        TabStop = True
        OnClick = arbVytvoreniClick
        Alignment = taLeftJustify
        Caption = 'TAdvOfficeRadioButton'
        Checked = True
        ReturnIsTab = False
        Version = '1.8.1.2'
      end
      object arbPrevod: TAdvOfficeRadioButton
        Left = 11
        Top = 41
        Width = 17
        Height = 17
        TabOrder = 1
        OnClick = arbPrevodClick
        Alignment = taLeftJustify
        Caption = 'TAdvOfficeRadioButton'
        ReturnIsTab = False
        Version = '1.8.1.2'
      end
      object arbTisk: TAdvOfficeRadioButton
        Left = 11
        Top = 74
        Width = 17
        Height = 17
        TabOrder = 2
        OnClick = arbTiskClick
        Alignment = taLeftJustify
        Caption = ''
        ReturnIsTab = False
        Version = '1.8.1.2'
      end
      object arbMail: TAdvOfficeRadioButton
        Left = 11
        Top = 107
        Width = 17
        Height = 17
        TabOrder = 3
        OnClick = arbMailClick
        Alignment = taLeftJustify
        Caption = ''
        ReturnIsTab = False
        Version = '1.8.1.2'
      end
    end
  end
  object lbxLog: TListBox
    Left = 0
    Top = 132
    Width = 825
    Height = 149
    Align = alClient
    ItemHeight = 13
    ScrollWidth = 1600
    TabOrder = 0
    OnDblClick = lbxLogDblClick
  end
  object asgMain: TAdvStringGrid
    Left = 0
    Top = 132
    Width = 825
    Height = 149
    Align = alClient
    ColCount = 11
    DefaultRowHeight = 18
    DrawingStyle = gdsClassic
    FixedCols = 0
    RowCount = 2
    Font.Charset = EASTEUROPE_CHARSET
    Font.Color = clBlack
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goEditing]
    ParentFont = False
    TabOrder = 1
    Visible = False
    OnDblClick = asgMainDblClick
    OnGetAlignment = asgMainGetAlignment
    OnCanSort = asgMainCanSort
    OnClickCell = asgMainClickCell
    OnCanEditCell = asgMainCanEditCell
    ActiveCellFont.Charset = DEFAULT_CHARSET
    ActiveCellFont.Color = clWindowText
    ActiveCellFont.Height = -11
    ActiveCellFont.Name = 'MS Sans Serif'
    ActiveCellFont.Style = [fsBold]
    ActiveCellColor = 15387318
    ColumnHeaders.Strings = (
      'tisk'
      'doklad'
      'vystaveno'
      'zaplaceno'
      'dne'
      'z'#250#269'tov'#225'no'
      'jm'#233'no'
      'faktura'
      'IDI.Id'
      'F.Id'
      'II.Id')
    ColumnSize.Location = clIniFile
    ControlLook.FixedGradientFrom = clWhite
    ControlLook.FixedGradientTo = clBtnFace
    ControlLook.FixedGradientHoverFrom = 13619409
    ControlLook.FixedGradientHoverTo = 12502728
    ControlLook.FixedGradientHoverMirrorFrom = 12502728
    ControlLook.FixedGradientHoverMirrorTo = 11254975
    ControlLook.FixedGradientDownFrom = 8816520
    ControlLook.FixedGradientDownTo = 7568510
    ControlLook.FixedGradientDownMirrorFrom = 7568510
    ControlLook.FixedGradientDownMirrorTo = 6452086
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
    FixedRowHeight = 18
    FixedFont.Charset = EASTEUROPE_CHARSET
    FixedFont.Color = clWindowText
    FixedFont.Height = -11
    FixedFont.Name = 'MS Sans Serif'
    FixedFont.Style = []
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
    SearchFooter.ColorTo = 13160660
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
    SortSettings.Column = 2
    SortSettings.Show = True
    Version = '8.7.0.0'
    ColWidths = (
      28
      64
      73
      64
      64
      65
      180
      78
      64
      64
      64)
  end
end
