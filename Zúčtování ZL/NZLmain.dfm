object fmMain: TfmMain
  Left = 477
  Top = 140
  Caption = 'Zaplacen'#233' a nez'#250#269'tovan'#233' z'#225'lohov'#233' listy za p'#345'ipojen'#237' k internetu'
  ClientHeight = 724
  ClientWidth = 1172
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
    Width = 1172
    Height = 132
    Align = alTop
    Color = clGradientActiveCaption
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
    ExplicitWidth = 1364
    FullHeight = 200
    object apnPrevod: TAdvPanel
      Left = 163
      Top = 0
      Width = 1009
      Height = 132
      Align = alClient
      Color = 15323589
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
      ExplicitLeft = 161
      ExplicitTop = 119
      ExplicitWidth = 1203
      ExplicitHeight = 122
      FullHeight = 95
      object deDatumFOxOd: TAdvDateTimePicker
        Left = 107
        Top = 15
        Width = 92
        Height = 21
        Date = 45108.472326388890000000
        Format = 'dd.MM.yyyy'
        Time = 45108.472326388890000000
        DoubleBuffered = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Kind = dkDate
        ParentDoubleBuffered = False
        ParentFont = False
        TabOrder = 0
        BorderStyle = bsSingle
        Ctl3D = True
        DateTime = 45108.472326388890000000
        Version = '1.3.6.5'
        LabelCaption = 'od'
        LabelPosition = lpLeftCenter
        LabelMargin = 8
        LabelFont.Charset = DEFAULT_CHARSET
        LabelFont.Color = clBlue
        LabelFont.Height = -11
        LabelFont.Name = 'MS Sans Serif'
        LabelFont.Style = [fsBold]
      end
      object deDatumFOxDo: TAdvDateTimePicker
        Left = 107
        Top = 48
        Width = 92
        Height = 21
        Date = 45138.472326388890000000
        Format = 'dd.MM.yyyy'
        Time = 45138.472326388890000000
        DoubleBuffered = True
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        Kind = dkDate
        ParentDoubleBuffered = False
        ParentFont = False
        TabOrder = 1
        BorderStyle = bsSingle
        Ctl3D = True
        DateTime = 45138.472326388890000000
        Version = '1.3.6.5'
        LabelCaption = 'do'
        LabelPosition = lpLeftCenter
        LabelMargin = 8
        LabelFont.Charset = DEFAULT_CHARSET
        LabelFont.Color = clBlue
        LabelFont.Height = -11
        LabelFont.Name = 'MS Sans Serif'
        LabelFont.Style = [fsBold]
      end
      object chbFO3: TCheckBox
        Left = 9
        Top = 18
        Width = 57
        Height = 17
        Caption = 'FO3'
        Checked = True
        DoubleBuffered = True
        ParentDoubleBuffered = False
        State = cbChecked
        TabOrder = 2
      end
      object chbFO2: TCheckBox
        Left = 9
        Top = 57
        Width = 57
        Height = 17
        Caption = 'FO2'
        Checked = True
        DoubleBuffered = True
        ParentDoubleBuffered = False
        State = cbChecked
        TabOrder = 3
      end
      object chbFO4: TCheckBox
        Left = 9
        Top = 80
        Width = 57
        Height = 17
        Caption = 'FO4'
        Checked = True
        DoubleBuffered = True
        ParentDoubleBuffered = False
        State = cbChecked
        TabOrder = 4
      end
      object btNacistFOx: TButton
        Left = 231
        Top = 10
        Width = 89
        Height = 41
        Caption = '&Na'#269#237'st FOx'
        TabOrder = 5
        OnClick = btNacistFOxClick
      end
      object apnVytvoritPDF: TAdvPanel
        Left = 373
        Top = 12
        Width = 300
        Height = 93
        Color = clGradientActiveCaption
        TabOrder = 6
        UseDockManager = True
        Version = '2.6.3.2'
        BorderColor = clBlack
        Caption.Color = clWhite
        Caption.ColorTo = clNone
        Caption.Font.Charset = DEFAULT_CHARSET
        Caption.Font.Color = clNone
        Caption.Font.Height = -11
        Caption.Font.Name = 'Tahoma'
        Caption.Font.Style = []
        Caption.GradientDirection = gdVertical
        Caption.Indent = 0
        Caption.ShadeLight = 255
        CollapsColor = clNone
        CollapsDelay = 0
        DoubleBuffered = True
        ShadowColor = clBlack
        ShadowOffset = 0
        StatusBar.BorderColor = clNone
        StatusBar.BorderStyle = bsSingle
        StatusBar.Font.Charset = DEFAULT_CHARSET
        StatusBar.Font.Color = 4473924
        StatusBar.Font.Height = -11
        StatusBar.Font.Name = 'Tahoma'
        StatusBar.Font.Style = []
        StatusBar.Color = clWhite
        StatusBar.GradientDirection = gdVertical
        Text = ''
        DesignSize = (
          300
          93)
        FullHeight = 200
        object cbNeprepisovat: TCheckBox
          Left = 113
          Top = 12
          Width = 192
          Height = 21
          Hint = 'Nep'#345'episovat u'#382' d'#345#237've vytvo'#345'en'#233' soubory DPF'
          Caption = ' nep'#345'episovat existuj'#237'c'#237' PDF'
          Checked = True
          State = cbChecked
          TabOrder = 0
        end
        object btVytvoritPDF: TButton
          Left = 8
          Top = 8
          Width = 89
          Height = 41
          Caption = 'Vytvo'#345'it PDF'
          TabOrder = 1
          OnClick = btVytvoritPDFClick
        end
        object btSablona: TButton
          Left = 130
          Top = 63
          Width = 71
          Height = 21
          Hint = #218'prava '#353'ablony faktury'
          Anchors = [akLeft]
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
        object btOdeslatNaServer: TButton
          Left = 217
          Top = 63
          Width = 71
          Height = 21
          Hint = 'Odesl'#225'n'#237' p'#345'eveden'#253'ch faktur na server'
          Anchors = [akLeft]
          Caption = '&Na server'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 3
          OnClick = btOdeslatNaServerClick
        end
      end
      object apnMailTisk: TAdvPanel
        Left = 696
        Top = 12
        Width = 121
        Height = 93
        Color = clGradientActiveCaption
        TabOrder = 7
        UseDockManager = True
        Version = '2.6.3.2'
        BorderColor = clBlack
        Caption.Color = clWhite
        Caption.ColorTo = clNone
        Caption.Font.Charset = DEFAULT_CHARSET
        Caption.Font.Color = clNone
        Caption.Font.Height = -11
        Caption.Font.Name = 'Tahoma'
        Caption.Font.Style = []
        Caption.GradientDirection = gdVertical
        Caption.Indent = 0
        Caption.ShadeLight = 255
        CollapsColor = clNone
        CollapsDelay = 0
        DoubleBuffered = True
        ShadowColor = clBlack
        ShadowOffset = 0
        StatusBar.BorderColor = clNone
        StatusBar.BorderStyle = bsSingle
        StatusBar.Font.Charset = DEFAULT_CHARSET
        StatusBar.Font.Color = 4473924
        StatusBar.Font.Height = -11
        StatusBar.Font.Name = 'Tahoma'
        StatusBar.Font.Style = []
        StatusBar.Color = clWhite
        StatusBar.GradientDirection = gdVertical
        Text = ''
        FullHeight = 200
        object btOdeslatMailem: TButton
          Left = 8
          Top = 16
          Width = 105
          Height = 25
          Caption = '&Odeslat e-maily'
          TabOrder = 0
          OnClick = btOdeslatMailemClick
        end
        object btTisk: TButton
          Left = 24
          Top = 61
          Width = 75
          Height = 25
          Caption = '&Vytisknout'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlue
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 1
          OnClick = btTiskClick
        end
      end
      object cbJenNezpracovaneFOx: TCheckBox
        Left = 231
        Top = 65
        Width = 119
        Height = 17
        Caption = 'jen nezpracovan'#233
        Checked = True
        DoubleBuffered = True
        ParentDoubleBuffered = False
        State = cbChecked
        TabOrder = 8
      end
      object editOrdNo: TAdvEdit
        Left = 126
        Top = 91
        Width = 73
        Height = 21
        EmptyTextStyle = []
        FlatLineColor = 11250603
        FocusColor = clWindow
        FocusFontColor = 3881787
        LabelCaption = #269'. fa'
        LabelFont.Charset = DEFAULT_CHARSET
        LabelFont.Color = clBlack
        LabelFont.Height = -11
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
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 9
        Text = ''
        Visible = True
        Version = '4.0.3.6'
      end
    end
    object apnVytvoreni: TAdvPanel
      Left = 163
      Top = 0
      Width = 1009
      Height = 132
      Align = alClient
      Color = 15323589
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
      ExplicitLeft = 161
      ExplicitWidth = 1195
      ExplicitHeight = 113
      FullHeight = 200
      object lbPozor1: TLabel
        Left = 5
        Top = 3
        Width = 381
        Height = 13
        Caption = 
          'Program vystav'#237' podle ZL fakturu a p'#345'ipoj'#237' platbu z'#225'lohov'#253'm list' +
          'em'
        Color = clNavy
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clNavy
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentColor = False
        ParentFont = False
      end
      object aseRok: TAdvSpinEdit
        Left = 191
        Top = 40
        Width = 55
        Height = 21
        Color = clWindow
        Value = 2023
        FloatValue = 2023.000000000000000000
        HexDigits = 0
        HexValue = 35
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
      end
      object btNacistZL: TButton
        Left = 291
        Top = 36
        Width = 100
        Height = 42
        Caption = '&Na'#269#237'st ZL'
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsBold]
        ParentFont = False
        TabOrder = 1
        OnClick = btNacistZLClick
      end
      object deDatumDokladu: TAdvDateTimePicker
        Left = 677
        Top = 45
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
        TabOrder = 2
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
        Left = 74
        Top = 73
        Width = 180
        Height = 19
        Caption = 'i jen '#269#225'ste'#269'n'#283' zaplacen'#233'   '
        TabOrder = 3
      end
      object acbRada: TAdvComboBox
        Left = 74
        Top = 40
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
        TabOrder = 4
        Text = '%'
      end
      object apnVytvorFOx: TAdvPanel
        Left = 444
        Top = 16
        Width = 130
        Height = 82
        Color = clGradientActiveCaption
        TabOrder = 5
        UseDockManager = True
        Version = '2.6.3.2'
        BorderColor = clBlack
        Caption.Color = clWhite
        Caption.ColorTo = clNone
        Caption.Font.Charset = DEFAULT_CHARSET
        Caption.Font.Color = clNone
        Caption.Font.Height = -11
        Caption.Font.Name = 'Tahoma'
        Caption.Font.Style = []
        Caption.GradientDirection = gdVertical
        Caption.Indent = 0
        Caption.ShadeLight = 255
        CollapsColor = clNone
        CollapsDelay = 0
        DoubleBuffered = True
        ShadowColor = clBlack
        ShadowOffset = 0
        StatusBar.BorderColor = clNone
        StatusBar.BorderStyle = bsSingle
        StatusBar.Font.Charset = DEFAULT_CHARSET
        StatusBar.Font.Color = 4473924
        StatusBar.Font.Height = -11
        StatusBar.Font.Name = 'Tahoma'
        StatusBar.Font.Style = []
        StatusBar.Color = clWhite
        StatusBar.GradientDirection = gdVertical
        Text = ''
        FullHeight = 200
        object btVytvoritFO: TButton
          Left = 11
          Top = 22
          Width = 100
          Height = 40
          Caption = '&Vytvo'#345'it FO3'
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsBold]
          ParentFont = False
          TabOrder = 0
          OnClick = btVytvoritFOClick
        end
      end
    end
    object apnVyberCinnosti: TAdvPanel
      Left = 0
      Top = 0
      Width = 163
      Height = 132
      Align = alLeft
      Color = 14733255
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlue
      Font.Height = -13
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
      ExplicitHeight = 144
      FullHeight = 122
      object glbVytvoreni: TGradientLabel
        Left = 36
        Top = 0
        Width = 125
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
        Width = 125
        Height = 33
        AutoSize = False
        Caption = '  PDF, Tisk, Mail'
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
      object btKonec: TButton
        Left = 36
        Top = 87
        Width = 75
        Height = 33
        Caption = '&Konec'
        TabOrder = 2
        OnClick = btKonecClick
      end
    end
    object apnProgress: TAdvPanel
      Left = 192
      Top = 93
      Width = 770
      Height = 30
      Color = clWhite
      TabOrder = 3
      UseDockManager = True
      Visible = False
      Version = '2.6.3.2'
      BorderColor = clBlack
      Caption.Color = clWhite
      Caption.ColorTo = clNone
      Caption.Font.Charset = DEFAULT_CHARSET
      Caption.Font.Color = clNone
      Caption.Font.Height = -11
      Caption.Font.Name = 'Tahoma'
      Caption.Font.Style = []
      Caption.GradientDirection = gdVertical
      Caption.Indent = 0
      Caption.ShadeLight = 255
      CollapsColor = clNone
      CollapsDelay = 0
      DoubleBuffered = True
      ShadowColor = clBlack
      ShadowOffset = 0
      StatusBar.BorderColor = clNone
      StatusBar.BorderStyle = bsSingle
      StatusBar.Font.Charset = DEFAULT_CHARSET
      StatusBar.Font.Color = 4473924
      StatusBar.Font.Height = -11
      StatusBar.Font.Name = 'Tahoma'
      StatusBar.Font.Style = []
      StatusBar.Color = clWhite
      StatusBar.GradientDirection = gdVertical
      Text = ''
      FullHeight = 200
      object apbProgress: TAdvProgressBar
        Left = 0
        Top = 0
        Width = 769
        Height = 29
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
      end
    end
  end
  object lbxLog: TListBox
    Left = 0
    Top = 504
    Width = 1172
    Height = 220
    Align = alBottom
    ItemHeight = 13
    ScrollWidth = 1600
    TabOrder = 0
    OnDblClick = lbxLogDblClick
  end
  object asgMain: TAdvStringGrid
    Left = 0
    Top = 132
    Width = 1172
    Height = 372
    Align = alClient
    ColCount = 17
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
      'PDF'
      'soubor PDF'
      'mail'
      'doklad ZL'
      'vystaveno'
      'zaplaceno'
      'dne'
      'z'#250#269'tov'#225'no'
      'jm'#233'no'
      'faktura'
      'VS'
      #269#225'stka'
      'mail'
      'IDI.Id'
      'II.Id'
      'mail zasl'#225'n'
      'datum fa')
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
    ExplicitTop = 399
    ExplicitWidth = 1185
    ExplicitHeight = 339
    ColWidths = (
      28
      90
      22
      64
      74
      65
      64
      64
      170
      64
      64
      64
      64
      64
      64
      64
      64)
  end
end
