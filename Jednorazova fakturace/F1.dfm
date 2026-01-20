object fmMain: TfmMain
  Left = 719
  Top = 184
  Caption = 'Zpracov'#225'n'#237' '#382#225'dost'#237' o fakturaci'
  ClientHeight = 679
  ClientWidth = 980
  Color = clBtnFace
  Constraints.MinHeight = 350
  Constraints.MinWidth = 680
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnActivate = FormActivate
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object lbxLog: TListBox
    Left = 0
    Top = 280
    Width = 980
    Height = 399
    TabStop = False
    Align = alClient
    BevelEdges = []
    Constraints.MinHeight = 32
    ItemHeight = 13
    ScrollWidth = 1600
    TabOrder = 0
  end
  object pnTop: TPanel
    Left = 0
    Top = 0
    Width = 980
    Height = 280
    Align = alTop
    Constraints.MinHeight = 280
    TabOrder = 1
    DesignSize = (
      980
      280)
    object btStart: TButton
      Left = 815
      Top = 20
      Width = 56
      Height = 21
      Anchors = [akTop, akRight]
      Caption = '&Start'
      TabOrder = 8
      OnClick = btStartClick
    end
    object aedTechnik: TAdvEdit
      Left = 891
      Top = 100
      Width = 84
      Height = 21
      TabStop = False
      EmptyTextStyle = []
      FocusColor = clBtnFace
      LabelCaption = 'Technik'
      LabelPosition = lpTopLeft
      LabelMargin = 2
      LabelFont.Charset = DEFAULT_CHARSET
      LabelFont.Color = clWindowText
      LabelFont.Height = -11
      LabelFont.Name = 'MS Sans Serif'
      LabelFont.Style = []
      Lookup.Font.Charset = DEFAULT_CHARSET
      Lookup.Font.Color = clWindowText
      Lookup.Font.Height = -11
      Lookup.Font.Name = 'Arial'
      Lookup.Font.Style = []
      Lookup.Separator = ';'
      Anchors = [akTop, akRight]
      Color = clWindow
      TabOrder = 11
      Text = ''
      Visible = True
      Version = '4.0.1.0'
    end
    object aedHotovost: TAdvEdit
      Left = 891
      Top = 140
      Width = 84
      Height = 21
      TabStop = False
      EditAlign = eaRight
      EditType = etFloat
      EmptyTextStyle = []
      FocusColor = clWindow
      Precision = 2
      LabelCaption = 'Zb'#253'v'#225' zaplatit'
      LabelPosition = lpTopLeft
      LabelMargin = 2
      LabelFont.Charset = DEFAULT_CHARSET
      LabelFont.Color = clWindowText
      LabelFont.Height = -11
      LabelFont.Name = 'MS Sans Serif'
      LabelFont.Style = []
      Lookup.Font.Charset = DEFAULT_CHARSET
      Lookup.Font.Color = clWindowText
      Lookup.Font.Height = -11
      Lookup.Font.Name = 'Arial'
      Lookup.Font.Style = []
      Lookup.Separator = ';'
      Anchors = [akTop, akRight]
      Color = clWindow
      TabOrder = 12
      Text = '0,00'
      Visible = True
      Version = '4.0.1.0'
    end
    object asgMain: TAdvStringGrid
      Left = 1
      Top = 0
      Width = 698
      Height = 220
      TabStop = False
      Anchors = [akLeft, akTop, akRight]
      ColCount = 11
      DefaultRowHeight = 18
      DrawingStyle = gdsClassic
      FixedCols = 0
      RowCount = 12
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goEditing]
      TabOrder = 10
      OnCanClickCell = asgMainCanClickCell
      OnCanEditCell = asgMainCanEditCell
      OnCheckBoxClick = asgMainCheckBoxClick
      ActiveCellFont.Charset = DEFAULT_CHARSET
      ActiveCellFont.Color = clWindowText
      ActiveCellFont.Height = -11
      ActiveCellFont.Name = 'MS Sans Serif'
      ActiveCellFont.Style = [fsBold]
      ActiveCellColor = 15387318
      ColumnHeaders.Strings = (
        ''
        ' k'#243'd Abra'
        'var. symbol'
        ' jm'#233'no'
        ' doklad'
        ' DL'
        'technik'
        ' Cu.Id'
        ' C.Id'
        ' F.Id'
        ' FO.Id')
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
      FilterDropDownClear = '(All)'
      FilterEdit.TypeNames.Strings = (
        'Starts with'
        'Ends with'
        'Contains'
        'Not contains'
        'Equal'
        'Not equal'
        'Clear')
      FixedColWidth = 19
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
      Navigation.AdvanceOnEnterLoop = False
      Navigation.AdvanceAutoEdit = False
      Navigation.AdvanceSkipReadOnlyCells = False
      Navigation.AutoComboSelect = False
      Navigation.AllowCtrlEnter = False
      Navigation.AllowClipboardRowGrow = False
      Navigation.AllowClipboardColGrow = False
      Navigation.ComboGetUpDown = False
      Navigation.CursorWalkAlwaysEdit = False
      Navigation.CopyHTMLTagsToClipboard = False
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
      SearchFooter.Font.Charset = DEFAULT_CHARSET
      SearchFooter.Font.Color = clWindowText
      SearchFooter.Font.Height = -11
      SearchFooter.Font.Name = 'MS Sans Serif'
      SearchFooter.Font.Style = []
      SearchFooter.ResultFormat = '(%d of %d)'
      ShowSelection = False
      ShowDesignHelper = False
      SortSettings.Column = 0
      Version = '8.6.2.0'
      ColWidths = (
        19
        55
        63
        97
        48
        40
        48
        40
        40
        40
        40)
    end
    object cbImport: TCheckBox
      Left = 715
      Top = 96
      Width = 147
      Height = 17
      Anchors = [akTop, akRight]
      Caption = '&Import z'#225'kazn'#237'ka do Abry'
      Enabled = False
      TabOrder = 3
    end
    object cbDoklad: TCheckBox
      Left = 715
      Top = 154
      Width = 112
      Height = 17
      Anchors = [akTop, akRight]
      Caption = 'Vytvo'#345'en'#237' &PP/FO '
      Enabled = False
      TabOrder = 4
    end
    object cbXLS: TCheckBox
      Left = 715
      Top = 173
      Width = 147
      Height = 17
      Anchors = [akTop, akRight]
      Caption = '&Hotovost do Technici.xls'
      Enabled = False
      TabOrder = 6
    end
    object cbClear: TCheckBox
      Left = 715
      Top = 191
      Width = 161
      Height = 17
      Anchors = [akTop, akRight]
      Caption = '&Vymaz'#225'n'#237' '#382#225'dosti o fakturaci'
      TabOrder = 7
    end
    object cbDL: TCheckBox
      Left = 715
      Top = 134
      Width = 112
      Height = 17
      Anchors = [akTop, akRight]
      Caption = 'Vytvo'#345'en'#237' &DL/PR'
      Enabled = False
      TabOrder = 5
    end
    object cbOprava: TCheckBox
      Left = 715
      Top = 115
      Width = 163
      Height = 17
      Anchors = [akTop, akRight]
      Caption = '&Oprava z'#225'kazn'#237'ka v datab'#225'zi'
      Enabled = False
      TabOrder = 2
    end
    object asgItems: TAdvStringGrid
      Left = -5
      Top = 226
      Width = 985
      Height = 59
      TabStop = False
      Anchors = [akLeft, akTop, akRight, akBottom]
      ColCount = 16
      Constraints.MinHeight = 40
      DefaultRowHeight = 18
      DrawingStyle = gdsClassic
      FixedCols = 0
      RowCount = 2
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goColSizing, goEditing]
      TabOrder = 13
      OnEditingDone = asgItemsEditingDone
      ActiveCellFont.Charset = DEFAULT_CHARSET
      ActiveCellFont.Color = clWindowText
      ActiveCellFont.Height = -11
      ActiveCellFont.Name = 'MS Sans Serif'
      ActiveCellFont.Style = [fsBold]
      ActiveCellColor = 15387318
      ColumnHeaders.Strings = (
        ''
        ' polo'#382'ka'
        ' '#269#225'stka'
        'druh'
        'cash'
        ' SN'
        ' MAC'
        ' n'#225'zev'
        'sklad'
        'ks'
        'SId'
        'SCId'
        'SBId'
        'sbId'
        'Doc'
        'DL')
      ColumnSize.Location = clIniFile
      ControlLook.FixedGradientFrom = clWhite
      ControlLook.FixedGradientTo = clSilver
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
      FilterDropDownClear = '(All)'
      FilterEdit.TypeNames.Strings = (
        'Starts with'
        'Ends with'
        'Contains'
        'Not contains'
        'Equal'
        'Not equal'
        'Clear')
      FixedColWidth = 18
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
      Look = glTMS
      Navigation.AlwaysEdit = True
      Navigation.AdvanceOnEnterLoop = False
      Navigation.AdvanceAutoEdit = False
      Navigation.AutoComboSelect = False
      Navigation.AllowCtrlEnter = False
      Navigation.AllowClipboardRowGrow = False
      Navigation.AllowClipboardColGrow = False
      Navigation.ComboGetUpDown = False
      Navigation.CursorWalkAlwaysEdit = False
      Navigation.LeftRightRowSelect = False
      Navigation.CopyHTMLTagsToClipboard = False
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
      SearchFooter.ColorTo = 15790320
      SearchFooter.Font.Charset = DEFAULT_CHARSET
      SearchFooter.Font.Color = clWindowText
      SearchFooter.Font.Height = -11
      SearchFooter.Font.Name = 'MS Sans Serif'
      SearchFooter.Font.Style = []
      SearchFooter.ResultFormat = '(%d of %d)'
      ShowDesignHelper = False
      SortSettings.Column = 0
      Version = '8.6.2.0'
      ColWidths = (
        18
        91
        52
        40
        40
        40
        40
        90
        41
        40
        40
        40
        39
        40
        40
        40)
    end
    object adeDatumDokladu: TAdvDateTimePicker
      Left = 715
      Top = 20
      Width = 82
      Height = 21
      Anchors = [akTop, akRight]
      Date = 41908.000000000000000000
      Format = 'dd.MM.yyyy'
      Time = 0.472326388888177500
      DoubleBuffered = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Kind = dkDate
      ParentDoubleBuffered = False
      ParentFont = False
      TabOrder = 0
      BorderStyle = bsSingle
      Ctl3D = True
      DateTime = 41908.472326388890000000
      Version = '1.3.6.0'
      LabelCaption = 'Datum dokladu'
      LabelPosition = lpTopLeft
      LabelMargin = 2
      LabelFont.Charset = DEFAULT_CHARSET
      LabelFont.Color = clBlack
      LabelFont.Height = -11
      LabelFont.Name = 'MS Sans Serif'
      LabelFont.Style = []
    end
    object adeDatumPlneni: TAdvDateTimePicker
      Left = 715
      Top = 60
      Width = 82
      Height = 21
      Anchors = [akTop, akRight]
      Date = 41908.000000000000000000
      Format = 'dd.MM.yyyy'
      Time = 0.472326388888177500
      DoubleBuffered = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Kind = dkDate
      ParentDoubleBuffered = False
      ParentFont = False
      TabOrder = 1
      BorderStyle = bsSingle
      Ctl3D = True
      DateTime = 41908.472326388890000000
      Version = '1.3.6.0'
      LabelCaption = 'Datum pln'#283'n'#237
      LabelPosition = lpTopLeft
      LabelMargin = 2
      LabelFont.Charset = DEFAULT_CHARSET
      LabelFont.Color = clBlack
      LabelFont.Height = -11
      LabelFont.Name = 'MS Sans Serif'
      LabelFont.Style = []
    end
    object aedDL: TAdvEdit
      Left = 891
      Top = 60
      Width = 84
      Height = 21
      TabStop = False
      EmptyTextStyle = []
      FocusColor = clBtnFace
      LabelCaption = 'Dodac'#237' list'
      LabelPosition = lpTopLeft
      LabelMargin = 2
      LabelFont.Charset = DEFAULT_CHARSET
      LabelFont.Color = clWindowText
      LabelFont.Height = -11
      LabelFont.Name = 'MS Sans Serif'
      LabelFont.Style = []
      Lookup.Font.Charset = DEFAULT_CHARSET
      Lookup.Font.Color = clWindowText
      Lookup.Font.Height = -11
      Lookup.Font.Name = 'Arial'
      Lookup.Font.Style = []
      Lookup.Separator = ';'
      Anchors = [akTop, akRight]
      Color = clWindow
      TabOrder = 15
      Text = ''
      Visible = True
      Version = '4.0.1.0'
    end
    object aedDoklad: TAdvEdit
      Left = 891
      Top = 20
      Width = 84
      Height = 21
      TabStop = False
      EmptyTextStyle = []
      FocusColor = clBtnFace
      LabelCaption = 'Doklad'
      LabelPosition = lpTopLeft
      LabelMargin = 2
      LabelFont.Charset = DEFAULT_CHARSET
      LabelFont.Color = clWindowText
      LabelFont.Height = -11
      LabelFont.Name = 'MS Sans Serif'
      LabelFont.Style = []
      Lookup.Font.Charset = DEFAULT_CHARSET
      Lookup.Font.Color = clWindowText
      Lookup.Font.Height = -11
      Lookup.Font.Name = 'Arial'
      Lookup.Font.Style = []
      Lookup.Separator = ';'
      Anchors = [akTop, akRight]
      Color = clWindow
      TabOrder = 14
      Text = ''
      Visible = True
      Version = '4.0.1.0'
    end
    object btReload: TButton
      Left = 815
      Top = 60
      Width = 56
      Height = 21
      Anchors = [akTop, akRight]
      Caption = '&Reload'
      TabOrder = 9
      OnClick = btReloadClick
    end
  end
  object qrItems: TZQuery
    Connection = DesU.dbZakos
    CachedUpdates = True
    SQL.Strings = (
      '')
    Params = <>
    ShowRecordTypes = [usUnmodified, usModified, usInserted, usDeleted]
    Left = 106
    Top = 128
  end
  object qrAbra: TZQuery
    Connection = DesU.dbAbra
    CachedUpdates = True
    ReadOnly = True
    SQL.Strings = (
      '')
    Params = <>
    Left = 106
    Top = 37
  end
  object qrMain: TZQuery
    Connection = DesU.dbZakos
    SortType = stIgnored
    CachedUpdates = True
    ReadOnly = True
    SQL.Strings = (
      '')
    Params = <>
    ShowRecordTypes = [usUnmodified, usModified, usInserted, usDeleted]
    WhereMode = wmWhereAll
    Left = 106
    Top = 83
  end
end
