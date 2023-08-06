object fmMain: TfmMain
  Left = 553
  Top = 194
  Caption = 'Z'#225'lohov'#233' listy za p'#345'ipojen'#237' k internetu'
  ClientHeight = 361
  ClientWidth = 834
  Color = clBtnFace
  Constraints.MinHeight = 400
  Constraints.MinWidth = 850
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnActivate = FormActivate
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object apnTop: TAdvPanel
    Left = 0
    Top = 0
    Width = 834
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
      834
      132)
    FullHeight = 200
    object apbProgress: TAdvProgressBar
      Left = 270
      Top = 103
      Width = 480
      Height = 22
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
    end
    object apnMail: TAdvPanel
      Left = 168
      Top = 95
      Width = 673
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
        673
        37)
      FullHeight = 95
      object fePriloha: TAdvFileNameEdit
        Left = 211
        Top = 8
        Width = 447
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
      Width = 673
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
        673
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
        Left = 587
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
        Left = 478
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
      Width = 673
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
      Width = 673
      Height = 96
      Anchors = [akLeft, akTop, akRight]
      ParentColor = True
      TabOrder = 0
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
      DesignSize = (
        673
        96)
      FullHeight = 200
      object aseRok: TAdvSpinEdit
        Left = 50
        Top = 62
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
        LabelCaption = 'rok'
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
        Left = 587
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
        TabOrder = 3
        OnClick = btVytvoritClick
      end
      object btKonec: TButton
        Left = 587
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
        TabOrder = 4
        OnClick = btKonecClick
      end
      object aedOd: TAdvEdit
        Left = 466
        Top = 20
        Width = 93
        Height = 21
        EditType = etNumeric
        EmptyTextStyle = []
        FocusColor = clWindow
        LabelCaption = 'od VS'
        LabelPosition = lpLeftCenter
        LabelMargin = 10
        LabelFont.Charset = DEFAULT_CHARSET
        LabelFont.Color = clBlue
        LabelFont.Height = -12
        LabelFont.Name = 'MS Sans Serif'
        LabelFont.Style = [fsBold]
        Lookup.Font.Charset = DEFAULT_CHARSET
        Lookup.Font.Color = clWindowText
        Lookup.Font.Height = -11
        Lookup.Font.Name = 'Arial'
        Lookup.Font.Style = []
        Lookup.Separator = ';'
        Color = clWindow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 1
        Text = '0'
        Visible = True
        OnChange = aedOdChange
        OnExit = aedOdExit
        OnKeyUp = aedOdKeyUp
        Version = '4.0.3.6'
      end
      object aedDo: TAdvEdit
        Left = 466
        Top = 56
        Width = 93
        Height = 21
        EditType = etNumeric
        EmptyTextStyle = []
        FocusColor = clWindow
        LabelCaption = 'do VS'
        LabelPosition = lpLeftCenter
        LabelMargin = 10
        LabelFont.Charset = DEFAULT_CHARSET
        LabelFont.Color = clBlue
        LabelFont.Height = -12
        LabelFont.Name = 'MS Sans Serif'
        LabelFont.Style = [fsBold]
        Lookup.Font.Charset = DEFAULT_CHARSET
        Lookup.Font.Color = clWindowText
        Lookup.Font.Height = -11
        Lookup.Font.Name = 'Arial'
        Lookup.Font.Style = []
        Lookup.Separator = ';'
        Color = clWindow
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 2
        Text = '0'
        Visible = True
        OnChange = aedDoChange
        OnExit = aedDoExit
        OnKeyUp = aedDoKeyUp
        Version = '4.0.3.6'
      end
      object deDatumDokladu: TAdvDateTimePicker
        Left = 292
        Top = 20
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
        TabOrder = 5
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
      object argMesic: TAdvOfficeRadioGroup
        Left = 17
        Top = 10
        Width = 160
        Height = 47
        BorderStyle = bsNone
        CaptionFont.Charset = DEFAULT_CHARSET
        CaptionFont.Color = clWindowText
        CaptionFont.Height = -11
        CaptionFont.Name = 'Tahoma'
        CaptionFont.Style = []
        Transparent = False
        Version = '1.8.1.2'
        Caption = 'm'#283's'#237'c'
        ParentBackground = False
        ParentCtl3D = True
        TabOrder = 6
        OnClick = argMesicClick
        Columns = 4
        ItemIndex = 0
        Items.Strings = (
          '3'
          '6'
          '9'
          '12')
        ButtonVertAlign = tlBottom
        Ellipsis = False
      end
      object deDatumSplatnosti: TAdvDateTimePicker
        Left = 292
        Top = 56
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
        TabOrder = 7
        OnChange = deDatumSplatnostiChange
        BorderStyle = bsSingle
        Ctl3D = True
        DateTime = 41908.472326388890000000
        Version = '1.3.6.5'
        LabelCaption = 'datum splatnosti'
        LabelPosition = lpLeftCenter
        LabelMargin = 8
        LabelFont.Charset = DEFAULT_CHARSET
        LabelFont.Color = clBlue
        LabelFont.Height = -11
        LabelFont.Name = 'MS Sans Serif'
        LabelFont.Style = [fsBold]
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
        Caption = '  Vytvo'#345'en'#237
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
    Width = 834
    Height = 229
    Align = alClient
    ItemHeight = 13
    ScrollWidth = 1600
    TabOrder = 0
    OnDblClick = lbxLogDblClick
  end
  object asgMain: TAdvStringGrid
    Left = 0
    Top = 132
    Width = 834
    Height = 229
    Align = alClient
    ColCount = 8
    DefaultRowHeight = 18
    DrawingStyle = gdsClassic
    FixedCols = 0
    RowCount = 2
    Font.Charset = EASTEUROPE_CHARSET
    Font.Color = clWindowText
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
      'VS'
      'ZL'
      'K'#269
      'jm'#233'no'
      'mail'
      'ne reklama'
      'datum')
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
    ExplicitLeft = 80
    ExplicitTop = 124
    ColWidths = (
      28
      64
      73
      64
      187
      298
      62
      64)
  end
  object dbMain: TZConnection
    ControlsCodePage = cCP_UTF16
    ClientCodepage = 'cp1250'
    Catalog = ''
    Properties.Strings = (
      'controls_cp=GET_ACP'
      'codepage=cp1250')
    ReadOnly = True
    AfterConnect = dbMainAfterConnect
    HostName = ''
    Port = 0
    Database = ''
    User = ''
    Password = ''
    Protocol = 'mysql-5'
    Left = 33
    Top = 186
  end
  object qrMain: TZQuery
    Connection = dbMain
    CachedUpdates = True
    ReadOnly = True
    SQL.Strings = (
      '')
    Params = <>
    Left = 84
    Top = 185
  end
  object qrSmlouva: TZQuery
    Connection = dbMain
    CachedUpdates = True
    ReadOnly = True
    SQL.Strings = (
      '')
    Params = <>
    Left = 168
    Top = 185
  end
  object dbAbra: TZConnection
    ControlsCodePage = cCP_UTF16
    Catalog = ''
    Properties.Strings = (
      'controls_cp=GET_ACP')
    AfterConnect = dbAbraAfterConnect
    HostName = ''
    Port = 0
    Database = ''
    User = ''
    Password = ''
    Protocol = 'firebird-2.1'
    Left = 8
    Top = 262
  end
  object qrAbra: TZQuery
    Connection = dbAbra
    CachedUpdates = True
    ReadOnly = True
    SQL.Strings = (
      '')
    Params = <>
    Left = 60
    Top = 246
  end
  object qrAdresa: TZQuery
    Connection = dbAbra
    CachedUpdates = True
    ReadOnly = True
    SQL.Strings = (
      '')
    Params = <>
    Left = 104
    Top = 246
  end
  object qrRadky: TZQuery
    Connection = dbAbra
    ReadOnly = True
    SQL.Strings = (
      '')
    Params = <>
    Left = 164
    Top = 238
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
    ReportOptions.LastChange = 42011.851949502300000000
    ScriptLanguage = 'PascalScript'
    StoreInDFM = False
    OnBeginDoc = frxReportBeginDoc
    OnEndDoc = frxReportEndDoc
    OnGetValue = frxReportGetValue
    Left = 288
    Top = 250
  end
  object fdsRadky: TfrxDBDataset
    UserName = 'fdsRadky'
    CloseDataSource = False
    OpenDataSource = False
    DataSet = qrRadky
    BCDToCurrency = False
    DataSetOptions = []
    Left = 416
    Top = 258
  end
  object frxDesigner: TfrxDesigner
    DefaultScriptLanguage = 'PascalScript'
    DefaultFont.Charset = DEFAULT_CHARSET
    DefaultFont.Color = clWindowText
    DefaultFont.Height = -13
    DefaultFont.Name = 'Arial'
    DefaultFont.Style = []
    DefaultLeftMargin = 10.000000000000000000
    DefaultRightMargin = 10.000000000000000000
    DefaultTopMargin = 10.000000000000000000
    DefaultBottomMargin = 10.000000000000000000
    DefaultPaperSize = 9
    DefaultOrientation = poPortrait
    GradientEnd = 11982554
    GradientStart = clWindow
    TemplatesExt = 'fr3'
    Restrictions = []
    RTLLanguage = False
    MemoParentFont = False
    Left = 340
    Top = 250
  end
  object qrTmp: TZQuery
    Connection = dbAbra
    SQL.Strings = (
      '')
    Params = <>
    Left = 152
    Top = 286
  end
end
