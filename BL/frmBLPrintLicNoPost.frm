VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLPrintLicNoPost 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License License Printing With No Charge"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmBLPrintLicNoPost.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbPrintOpt 
      Height          =   384
      Left            =   7248
      TabIndex        =   12
      Tag             =   $"frmBLPrintLicNoPost.frx":08CA
      Top             =   4476
      Width           =   3564
      _Version        =   196608
      _ExtentX        =   6286
      _ExtentY        =   677
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   0
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   2
      SearchMethod    =   0
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483627
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ScrollHScale    =   2
      ScrollHInc      =   0
      ColsFrozen      =   0
      ScrollBarV      =   1
      NoIntegralHeight=   0   'False
      HighestPrecedence=   0
      AllowColResize  =   0
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   -1  'True
      DataAutoSizeCols=   2
      SearchIgnoreCase=   -1  'True
      ScrollBarH      =   1
      DataFieldList   =   ""
      ColumnEdit      =   -1
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   -1
      EditHeight      =   -1
      GrayAreaColor   =   -2147483633
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   150
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   -2147483637
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   0   'False
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   -2147483637
      ExtendRow       =   0
      ListPosition    =   0
      ButtonThreeDAppearance=   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      AutoSearchFill  =   -1  'True
      AutoSearchFillDelay=   200
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   1
      EditAlignV      =   0
      ColDesigner     =   "frmBLPrintLicNoPost.frx":09F1
   End
   Begin LpLib.fpCombo fpcmbBalanceType 
      Height          =   384
      Left            =   3792
      TabIndex        =   5
      Tag             =   $"frmBLPrintLicNoPost.frx":0CE8
      Top             =   4476
      Width           =   2784
      _Version        =   196608
      _ExtentX        =   4911
      _ExtentY        =   677
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   0
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   2
      SearchMethod    =   0
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483627
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ScrollHScale    =   2
      ScrollHInc      =   0
      ColsFrozen      =   0
      ScrollBarV      =   1
      NoIntegralHeight=   0   'False
      HighestPrecedence=   0
      AllowColResize  =   0
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   -1  'True
      DataAutoSizeCols=   2
      SearchIgnoreCase=   -1  'True
      ScrollBarH      =   1
      DataFieldList   =   ""
      ColumnEdit      =   -1
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   -1
      EditHeight      =   -1
      GrayAreaColor   =   -2147483633
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   5
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   -2147483637
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   0   'False
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   -2147483637
      ExtendRow       =   0
      ListPosition    =   0
      ButtonThreeDAppearance=   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      AutoSearchFill  =   -1  'True
      AutoSearchFillDelay=   200
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   1
      EditAlignV      =   0
      ColDesigner     =   "frmBLPrintLicNoPost.frx":0E7C
   End
   Begin LpLib.fpCombo fpcmbPrintFeesYN 
      Height          =   384
      Left            =   3792
      TabIndex        =   4
      Tag             =   $"frmBLPrintLicNoPost.frx":1173
      Top             =   3936
      Width           =   1104
      _Version        =   196608
      _ExtentX        =   1947
      _ExtentY        =   677
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   0
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   2
      SearchMethod    =   0
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483627
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ScrollHScale    =   2
      ScrollHInc      =   0
      ColsFrozen      =   0
      ScrollBarV      =   1
      NoIntegralHeight=   0   'False
      HighestPrecedence=   0
      AllowColResize  =   0
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   -1  'True
      DataAutoSizeCols=   2
      SearchIgnoreCase=   -1  'True
      ScrollBarH      =   1
      DataFieldList   =   ""
      ColumnEdit      =   -1
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   -1
      EditHeight      =   -1
      GrayAreaColor   =   -2147483633
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   5
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   -2147483637
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   0   'False
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   -2147483637
      ExtendRow       =   0
      ListPosition    =   0
      ButtonThreeDAppearance=   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      AutoSearchFill  =   -1  'True
      AutoSearchFillDelay=   200
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   1
      EditAlignV      =   0
      ColDesigner     =   "frmBLPrintLicNoPost.frx":12E8
   End
   Begin LpLib.fpCombo fpcmbSignature 
      Height          =   384
      Left            =   9312
      TabIndex        =   11
      Tag             =   $"frmBLPrintLicNoPost.frx":15DF
      Top             =   3648
      Width           =   984
      _Version        =   196608
      _ExtentX        =   1736
      _ExtentY        =   677
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   0
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   2
      SearchMethod    =   0
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483627
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ScrollHScale    =   2
      ScrollHInc      =   0
      ColsFrozen      =   0
      ScrollBarV      =   1
      NoIntegralHeight=   0   'False
      HighestPrecedence=   0
      AllowColResize  =   0
      AllowColDragDrop=   0
      ReadOnly        =   0   'False
      VScrollSpecial  =   0   'False
      VScrollSpecialType=   0
      EnableKeyEvents =   -1  'True
      EnableTopChangeEvent=   -1  'True
      DataAutoHeadings=   -1  'True
      DataAutoSizeCols=   2
      SearchIgnoreCase=   -1  'True
      ScrollBarH      =   1
      DataFieldList   =   ""
      ColumnEdit      =   -1
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   -1
      EditHeight      =   -1
      GrayAreaColor   =   -2147483633
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   5
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   -2147483637
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   0   'False
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   -2147483637
      ExtendRow       =   0
      ListPosition    =   0
      ButtonThreeDAppearance=   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      Redraw          =   -1  'True
      AutoSearchFill  =   -1  'True
      AutoSearchFillDelay=   200
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   1
      EditAlignV      =   0
      ColDesigner     =   "frmBLPrintLicNoPost.frx":16CD
   End
   Begin FPSpread.vaSpread vaSpread 
      Height          =   2265
      Left            =   2925
      TabIndex        =   13
      Tag             =   $"frmBLPrintLicNoPost.frx":19C4
      Top             =   5190
      Width           =   7890
      _Version        =   196613
      _ExtentX        =   13917
      _ExtentY        =   3995
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   5
      MaxRows         =   2000
      ShadowColor     =   13684944
      SpreadDesigner  =   "frmBLPrintLicNoPost.frx":1B09
      VisibleCols     =   4
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   300
      Left            =   528
      TabIndex        =   14
      Top             =   8016
      Width           =   684
      _Version        =   131072
      _ExtentX        =   1206
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   9405029
      ForeColor       =   8454143
      Text            =   ""
      Shape           =   0
      ShapeRoundWidth =   180
      ShapeRoundHeight=   180
      BorderWidth     =   -1
      BorderColor     =   -2147483630
      ThreeDWidth     =   -1
      ThreeDShadowColor=   -2147483632
      ThreeDHighlightColor=   16777215
      ThreeDText      =   0
      ThreeDTextHighlightColor=   16777215
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignPictureH   =   0
      AlignPictureV   =   0
      PictureStyle    =   0
      WordWrap        =   -1  'True
      ScaleMode       =   1
      ThreeDStyle     =   2
      Position        =   0
      PosBaseX        =   0
      PosBaseY        =   0
      PosOffsetX      =   -100
      PosOffsetY      =   300
      MaxWidth        =   5000
      CloudInset      =   100
      CloudMinWidth   =   600
      TailShape       =   2
      TailType        =   2
      TailBaseOffsetOutside=   300
      TailBaseOffsetInside=   100
      TailBaseAxisOutside=   0
      TailBaseAxisInside=   0
      TailBubbleCount =   3
      AlignTextH      =   1
      AlignTextV      =   1
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      AutoScan        =   0
      ScanAllDescendants=   -1  'True
      Interval        =   500
      IntervalNext    =   200
      AutoSize        =   -1  'True
      UseTagProp      =   -1  'True
      HideOnInactiveApp=   0   'False
      HideOnMouseDown =   2
      HideOnKeyDown   =   2
      HideOnFocus     =   0   'False
      ScanDisabledControls=   -1  'True
      ThreeDAppearance=   0
      FollowFocus     =   0   'False
      TemplateName    =   ""
   End
   Begin EditLib.fpDateTime fptxtVThru 
      Height          =   348
      Left            =   2160
      TabIndex        =   2
      Tag             =   "The date entered here will appear on the business license forms as the expiration date for this license."
      Top             =   2880
      Width           =   1740
      _Version        =   196608
      _ExtentX        =   3069
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   2
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   12648447
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   "04/28/2003"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
      DateMax         =   "00000000"
      DateMin         =   "00000000"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "00000000"
      TimeDefault     =   "000000"
      TimeStyle       =   0
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      PopUpType       =   1
      DateCalcY2KSplit=   60
      CaretPosition   =   0
      IncYear         =   1
      IncMonth        =   1
      IncDay          =   1
      IncHour         =   1
      IncMinute       =   1
      IncSecond       =   1
      ButtonColor     =   13684944
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpDateTime fptxtFromDate 
      Height          =   345
      Left            =   2160
      TabIndex        =   1
      Tag             =   "The date entered here will appear on the business license as the first day of the valid date range for this license."
      Top             =   2400
      Width           =   1740
      _Version        =   196608
      _ExtentX        =   3069
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   2
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   12648447
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   "04/28/2003"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
      DateMax         =   "00000000"
      DateMin         =   "00000000"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "00000000"
      TimeDefault     =   "000000"
      TimeStyle       =   0
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      PopUpType       =   1
      DateCalcY2KSplit=   60
      CaretPosition   =   0
      IncYear         =   1
      IncMonth        =   1
      IncDay          =   1
      IncHour         =   1
      IncMinute       =   1
      IncSecond       =   1
      ButtonColor     =   13684944
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpDateTime fptxtIssDate 
      Height          =   348
      Left            =   3072
      TabIndex        =   0
      Tag             =   $"frmBLPrintLicNoPost.frx":74DA
      Top             =   1440
      Width           =   1740
      _Version        =   196608
      _ExtentX        =   3069
      _ExtentY        =   614
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   2
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   12648447
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   "04/28/2003"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
      DateMax         =   "00000000"
      DateMin         =   "00000000"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "00000000"
      TimeDefault     =   "000000"
      TimeStyle       =   0
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      PopUpType       =   1
      DateCalcY2KSplit=   60
      CaretPosition   =   0
      IncYear         =   1
      IncMonth        =   1
      IncDay          =   1
      IncHour         =   1
      IncMinute       =   1
      IncSecond       =   1
      ButtonColor     =   13684944
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpDateTime fpBLYear 
      Height          =   345
      Left            =   3795
      TabIndex        =   3
      Tag             =   "The date entered here will appear on the business license as the active year for this license"
      Top             =   3450
      Width           =   1095
      _Version        =   196608
      _ExtentX        =   1931
      _ExtentY        =   609
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   2
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   12648447
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   "2018"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "yyyy"
      DateMax         =   "00000000"
      DateMin         =   "00000000"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "00000000"
      TimeDefault     =   "000000"
      TimeStyle       =   0
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      PopUpType       =   1
      DateCalcY2KSplit=   60
      CaretPosition   =   0
      IncYear         =   1
      IncMonth        =   1
      IncDay          =   1
      IncHour         =   1
      IncMinute       =   1
      IncSecond       =   1
      ButtonColor     =   13684944
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtHeading 
      Height          =   396
      Index           =   0
      Left            =   6048
      TabIndex        =   6
      Top             =   1488
      Width           =   4956
      _Version        =   196608
      _ExtentX        =   8742
      _ExtentY        =   698
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   38
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtHeading 
      Height          =   396
      Index           =   1
      Left            =   6048
      TabIndex        =   7
      Top             =   1920
      Width           =   4956
      _Version        =   196608
      _ExtentX        =   8742
      _ExtentY        =   698
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   30
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtHeading 
      Height          =   396
      Index           =   2
      Left            =   6048
      TabIndex        =   8
      Top             =   2352
      Width           =   4956
      _Version        =   196608
      _ExtentX        =   8742
      _ExtentY        =   698
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   30
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtHeading 
      Height          =   396
      Index           =   3
      Left            =   6048
      TabIndex        =   9
      Top             =   2784
      Width           =   4956
      _Version        =   196608
      _ExtentX        =   8742
      _ExtentY        =   698
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   45
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   636
      Left            =   6228
      TabIndex        =   23
      TabStop         =   0   'False
      Tag             =   "Press 'ESC' to exit this screen and return to the 'License Processing' menu."
      Top             =   7584
      Width           =   1644
      _Version        =   131072
      _ExtentX        =   2900
      _ExtentY        =   1122
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmBLPrintLicNoPost.frx":7568
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
      Height          =   636
      Left            =   8148
      TabIndex        =   24
      TabStop         =   0   'False
      Tag             =   "Press 'Process' to begin printing the business license forms using the parameters entered above."
      Top             =   7584
      Width           =   1644
      _Version        =   131072
      _ExtentX        =   2900
      _ExtentY        =   1122
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmBLPrintLicNoPost.frx":7747
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAlign 
      Height          =   636
      Left            =   4308
      TabIndex        =   25
      TabStop         =   0   'False
      Tag             =   $"frmBLPrintLicNoPost.frx":7926
      Top             =   7584
      Width           =   1644
      _Version        =   131072
      _ExtentX        =   2900
      _ExtentY        =   1122
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmBLPrintLicNoPost.frx":7A81
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   636
      Left            =   1860
      TabIndex        =   26
      TabStop         =   0   'False
      Tag             =   $"frmBLPrintLicNoPost.frx":7C5D
      Top             =   7584
      Width           =   2172
      _Version        =   131072
      _ExtentX        =   3831
      _ExtentY        =   1122
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   0   'False
      Value           =   0   'False
      GroupID         =   0
      GroupSelect     =   0
      DrawFocusRect   =   2
      DrawFocusRectCell=   -1
      GrayAreaPictureStyle=   0
      Static          =   0   'False
      BackStyle       =   1
      AutoSize        =   0
      AutoSizeOffsetTop=   0
      AutoSizeOffsetBottom=   0
      AutoSizeOffsetLeft=   0
      AutoSizeOffsetRight=   0
      DropShadowOffsetX=   3
      DropShadowOffsetY=   3
      DropShadowType  =   0
      DropShadowColor =   0
      Redraw          =   -1  'True
      ButtonDesigner  =   "frmBLPrintLicNoPost.frx":7D2D
   End
   Begin EditLib.fpText fptxtAuthorizedBy 
      Height          =   396
      Left            =   7632
      TabIndex        =   10
      Tag             =   "The person's name entered here will appear on the business license as the town offical responsible for issuing business licenses."
      Top             =   3216
      Width           =   3372
      _Version        =   196608
      _ExtentX        =   5948
      _ExtentY        =   698
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   -2147483642
      BorderWidth     =   1
      ButtonDisable   =   0   'False
      ButtonHide      =   0   'False
      ButtonIncrement =   1
      ButtonMin       =   0
      ButtonMax       =   100
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      AutoCase        =   0
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   30
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   7596
      Left            =   240
      Top             =   1056
      Width           =   11196
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Signature Line (Y/N)?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   300
      Left            =   6336
      TabIndex        =   32
      Top             =   3744
      Width           =   2796
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "License Type:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   300
      Left            =   8208
      TabIndex        =   31
      Top             =   4128
      Width           =   1500
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Authorized By:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   300
      Left            =   5808
      TabIndex        =   30
      Top             =   3312
      Width           =   1692
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   240
      X2              =   11424
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Click on a cell in the 'Select' column next to each customer for whom to print a license."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   1692
      Left            =   960
      TabIndex        =   28
      Top             =   5520
      Width           =   1500
   End
   Begin VB.Label lblBalloon 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "HELP BALLOONS ON"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   1908
      TabIndex        =   27
      Top             =   8256
      Width           =   2100
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "License Heading:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   300
      Left            =   7536
      TabIndex        =   22
      Top             =   1152
      Width           =   1980
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Balances To Print On License"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   300
      Left            =   432
      TabIndex        =   21
      Top             =   4512
      Width           =   3228
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Print License Fees (Y/N)?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   300
      Left            =   864
      TabIndex        =   20
      Top             =   3984
      Width           =   2796
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Business License For Year:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   300
      Left            =   672
      TabIndex        =   19
      Top             =   3504
      Width           =   2988
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Business License Date Range:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   300
      Left            =   1104
      TabIndex        =   18
      Top             =   2064
      Width           =   3276
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   300
      Left            =   1248
      TabIndex        =   17
      Top             =   2496
      Width           =   732
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   300
      Left            =   1488
      TabIndex        =   16
      Top             =   2928
      Width           =   492
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date License Issued:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   300
      Left            =   576
      TabIndex        =   15
      Top             =   1488
      Width           =   2412
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   756
      Index           =   1
      Left            =   1500
      Top             =   204
      Width           =   8652
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Business Licenses: No Posting"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   2508
      TabIndex        =   29
      Top             =   384
      Width           =   7068
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   1500
      Top             =   144
      Width           =   8652
   End
End
Attribute VB_Name = "frmBLPrintLicNoPost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class
  Dim UsePermLicNum As Boolean
  Dim Head1Laser$
  Dim Head2Laser$
  Dim Head3Laser$
  Dim Head1Text$
  Dim Head2Text$
  Dim Head3Text$
  Dim Head4Text$
  Dim FeeAmt1#
  Dim FeeAmt2#
  Dim FeeAmt3#
  Dim FeeAmt4#
  Dim FeeAmt5#
  Dim ThisIssFee#
  Dim ThisAcctBal#
Private Sub cmdAlign_Click()
  Dim SHeading1$
  Dim SHeading2$
  Dim SHeading3$
  Dim SHeading4$
  Dim Heading1 As Integer
  Dim Heading2 As Integer
  Dim Heading3 As Integer
  Dim Heading4 As Integer
  Dim tab1 As Integer
  Dim tab2 As Integer
  Dim Tab3 As Integer
  Dim Tab4 As Integer
  Dim ReportFile$
  Dim LPRINT As Integer
  Dim LCnt As Integer
  
  On Error GoTo ERRORSTUFF
  
  cmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  
  SHeading1$ = QPTrim$(fptxtHeading(0).Text)
  SHeading2$ = QPTrim$(fptxtHeading(1).Text)
  SHeading3$ = QPTrim$(fptxtHeading(2).Text)
  SHeading4$ = QPTrim$(fptxtHeading(3).Text)

  Heading1 = Len(SHeading1$)
  Heading2 = Len(SHeading2$)
  Heading3 = Len(SHeading3$)
  Heading4 = Len(SHeading4$)

  If Len(Heading1) > 0 Then tab1 = Heading1 / 2 Else tab1 = 0
  If Len(Heading2) > 0 Then tab2 = Heading2 / 2 Else tab2 = 0
  If Len(Heading3) > 0 Then Tab3 = Heading3 / 2 Else Tab3 = 0
  If Len(Heading4) > 0 Then Tab4 = Heading4 / 2 Else Tab4 = 0
  
  ReportFile$ = "LICMASK.PRT"
  LPRINT = FreeFile
  Open ReportFile$ For Output As #LPRINT

  ' Print Form Test
  Print #LPRINT, "TOP"
  For LCnt = 1 To 4
    Print #LPRINT, ""
  Next LCnt
  Print #LPRINT, Tab(37 - tab1); SHeading1$
  Print #LPRINT, Tab(37 - tab2); SHeading2$
  Print #LPRINT, Tab(37 - Tab3); SHeading3$
  Print #LPRINT, Tab(37 - Tab4); SHeading4$
  Print #LPRINT, Tab(66); Mid(fptxtVThru.Text, 7, 4)
  Print #LPRINT,
  Print #LPRINT, Tab(11); "Name of Some Business"
  Print #LPRINT, Tab(11); "Address Line 1"; Tab(58); "########"
  Print #LPRINT, Tab(11); "Address Line 2"
  Print #LPRINT, Tab(11); "Address Line 3"
  Print #LPRINT, Tab(55); Mid(fptxtFromDate.Text, 1, 6) + Mid(fptxtFromDate.Text, 9, 2);
  Print #LPRINT, Tab(64); Mid(fptxtVThru.Text, 1, 6) + Mid(fptxtVThru.Text, 9, 2)
  Print #LPRINT,
  Print #LPRINT,
  Print #LPRINT,
  Print #LPRINT, Tab(11); String$(35, "X")
  Print #LPRINT,
  Print #LPRINT,
  Print #LPRINT, Tab(5); "XXXXXXXX"; Tab(15); String$(30, "X"); Tab(62); "XXXXX.XX"
  For LCnt = 24 To 35
    Print #LPRINT, ""
  Next LCnt
  Print #LPRINT, Tab(62); "XXXXX.XX"
  Print #LPRINT,
  Print #LPRINT,
  Print #LPRINT,
  Print #LPRINT, Tab(62); "XXXXX.XX"
  Print #LPRINT,
  Print #LPRINT, "BOTTOM"
  
  Close
  ViewPrint ReportFile, "Business License Mask", True
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLPrintLic", "cmdAlign_Click", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    ClearInUse PWcnt
    Terminate
  
  
End Sub


Private Sub cmdExit_Click()
  frmBLPrintLicMenu.Show
  DoEvents
  Unload frmBLPrintLicNoPost
End Sub

Private Sub cmdProcess_Click()
  frmBLMessageBoxJr.cmdExit.Text = "ESC OK"
  frmBLMessageBoxJr.Label1.Caption = "Reminder: Printing business licenses here has no affect on a customer's license balance. No license fees are posted."
  frmBLMessageBoxJr.Label1.Top = 800
  frmBLMessageBoxJr.Show vbModal
  MainLog ("User processed non-posting business licenses and was reminded that no customer balances are affected in this process.")
  If fpcmbPrintOpt.Text = "Laser Form" Then
    Call PrintGraphics
  Else
    Call PrintText
  End If
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call LoadMe
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF5:
      SendKeys "%A"
      Call cmdAlign_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdProcess_Click
      KeyCode = 0
    Case vbKeyF1:
      SendKeys "%T"
      Call cmdHelp_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      ClearInUse PWcnt
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLPrintLicNoPost.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  Dim TownHandle As Integer
  Dim TownRec As TownSetUpType
  Dim ThisZip$
  Dim ThisYear As Integer
  Dim CustRec As ARCustRecType
  Dim CustIdxRec As CustNameIdxType
  Dim CustIdxHandle As Integer
  Dim CustIdxRecNum As Integer
  Dim CHandle As Integer
  Dim TotalAccts As Integer
  Dim x As Integer
  Dim cnt As Integer
  Dim CustCnt As Integer
  Dim Nextx As Integer
  Dim NewYear$
  Dim DHandle As Integer
  Dim ThisDate$
  Dim ThisHeader$
  
  lblBalloon.Visible = False
'  fptxtIssDate.ToolTipText = "Enter the date which will appear on the business licenses the first valid day of this license period."
'  fptxtFromDate.ToolTipText = "Enter the date which will appear on the business licenses indicating the first valid day."
'  fptxtVThru.ToolTipText = "The date entered here will appear on the license as the expiration date."
'  fpBLYear.ToolTipText = "Enter the year that will appear on the license as the primary year for this version of business license."
'  fpcmbPrintFeesYN.ToolTipText = "Select 'Yes' in this drop down list if you wish for fee amounts to appear on the business licenses."
'  fpcmbBalanceType.ToolTipText = "If you select 'Yes' in the 'Print License Fees (Y/N)?' field Then you can elect to display only fees generated for this license or current fees plus outstanding fees."
'  fptxtHeading(0).ToolTipText = "Optional line of text that will appear as the first line of the license header."
'  fptxtHeading(1).ToolTipText = "Optional line of text that will appear as the second line of the license header."
'  fptxtHeading(2).ToolTipText = "Optional line of text that will appear as the third line of the license header."
'  fptxtHeading(3).ToolTipText = "Optional line of text that will appear as the fourth line of the license header."
'  fptxtAuthorizedBy.ToolTipText = "You can elect to include (recommended) the name of a town employee who would be the first person to be contacted with questions/comments."
'  fpcmbPrintOpt.ToolTipText = "Business licenses can be printed graphically (same as laser) or in text (tractor fed) forms."
'  cmdAlign.ToolTipText = "Use this button to help line up license forms."
'  cmdExit.ToolTipText = "Press to return to the 'License Processing' menu."
'  cmdProcess.ToolTipText = "Press the 'Process' button to calculate fees for all customers earmarked for renewal."
'  cmdHelp.ToolTipText = "Press 'Turn Help On' to activate informational balloons for each field. Press 'Turn Help Off' to deactivate the informational balloons."
'  vaSpread.ToolTipText = "Click the 'Select' cell next to the customer for whom you wish a license printed."
  If Exist("validthrudate.dat") Then
    DHandle = FreeFile
    Open "validthrudate.dat" For Input As #DHandle
    Line Input #DHandle, ThisDate
    fptxtVThru = ThisDate
    Close DHandle
  Else
    fptxtVThru = Date
    NewYear = fptxtVThru.AdjustDate(fptxtVThru.DateValue, 1, 0, 0)
    fptxtVThru.DateValue = NewYear
  End If
  
  If Exist("appheader.dat") Then
    DHandle = FreeFile
    Open "appheader.dat" For Input As #DHandle
    Line Input #DHandle, ThisHeader
    fptxtHeading(0).Text = ThisHeader
    Close DHandle
    Head1Laser$ = ThisHeader
    Head1Text$ = ThisHeader
  Else
    fptxtHeading(0).Text = "MUNICIPAL LICENSE"
    Head1Laser$ = "MUNICIPAL LICENSE"
    Head1Text$ = QPTrim$(TownRec.TownName)
  End If
  
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  
'  Head1Laser$ = "MUNICIPAL LICENSE"
  Head2Laser$ = QPTrim$(TownRec.TownName)
  Select Case QPTrim$(TownRec.State)
    Case "NC"
      Head3Laser$ = "STATE OF NORTH CAROLINA"
    Case "SC"
      Head3Laser$ = "STATE OF SOUTH CAROLINA"
    Case "VA"
      Head3Laser$ = "STATE OF VIRGINIA"
    Case "GA"
      Head3Laser$ = "STATE OF GEORGIA"
    Case "AR"
      Head3Laser$ = "STATE OF ARKANSAS"
    Case "AL"
      Head3Laser$ = "STATE OF ALABAMA"
    Case "OK"
      Head3Laser$ = "STATE OF OKLAHOMA"
    Case Else
      Head3Laser$ = "UNKNOWN STATE"
  End Select
  
'  Head1Text$ = QPTrim$(TownRec.TownName)
  Head2Text$ = QPTrim$(TownRec.TownAdd1)
  Head3Text$ = QPTrim$(TownRec.TownAdd2)
  Head4Text$ = QPTrim$(TownRec.City) + ", " + QPTrim$(TownRec.State) + " " + QPTrim$(TownRec.ZipCode)
  
  fptxtFromDate = Date
'  fptxtVThru = Date
'  NewYear = fptxtVThru.AdjustDate(fptxtVThru.DateValue, 1, 0, 0)
'  fptxtVThru.DateValue = NewYear
  fptxtIssDate = Date
  fpcmbPrintFeesYN.Text = "Yes"
  fpcmbPrintFeesYN.AddItem "No"
  fpcmbPrintFeesYN.AddItem "Yes"
  fpcmbBalanceType.Text = "Current Balance Only"
  fpcmbBalanceType.AddItem "Current Balance Only"
  fpcmbBalanceType.AddItem "Total Balance"
  fpcmbSignature.Text = "Yes"
  fpcmbSignature.AddItem "Yes"
  fpcmbSignature.AddItem "No"
  
  fptxtAuthorizedBy.Text = QPTrim$(TownRec.Contact)
  
  fpcmbPrintOpt.Text = "Laser Form"
  fpcmbPrintOpt.AddItem "Laser Form"
  fpcmbPrintOpt.AddItem "Tractor Fed Form"
  If Mid(TownRec.ZipCode, 7, 1) = " " Then
    ThisZip = Mid(TownRec.ZipCode, 1, 5)
  Else
    ThisZip = QPTrim$(TownRec.ZipCode)
  End If
  
  fptxtHeading(3).Text = QPTrim$(TownRec.City) + ", " + QPTrim$(TownRec.State) + "  " + ThisZip
   OpenCustNameIdxFile CustIdxHandle
   CustIdxRecNum = LOF(CustIdxHandle) \ Len(CustIdxRec)
   If CustIdxRecNum = 0 Then 'file is there but there is nothing in it
     frmBLMessageBoxJr.Label1.Caption = "No Customers in index."
     frmBLMessageBoxJr.Label1.Top = 900
     frmBLMessageBoxJr.Show vbModal
     Close
     Exit Sub
   End If
   
   ReDim CustIdx(1 To CustIdxRecNum) As Integer
   For x = 1 To CustIdxRecNum
     Get CustIdxHandle, x, CustIdxRec
     CustIdx(x) = CustIdxRec.CustRec 'load array with record pointers
   Next x
   Close CustIdxHandle
   
   If Not Exist("ARCUST.DAT") Then
     frmBLMessageBoxJr.Label1.Caption = "Path to ARCUST.DAT could not be found"
     frmBLMessageBoxJr.Label1.Top = 900
     frmBLMessageBoxJr.Show vbModal
     Exit Sub
   End If
   
   OpenCustFile CHandle
   CustCnt = LOF(CHandle) / Len(CustRec)
   
   If CustCnt = 0 Then
     frmBLMessageBoxJr.Label1.Caption = "No Customer data on file."
     frmBLMessageBoxJr.Label1.Top = 900
     frmBLMessageBoxJr.Show vbModal
     Close
     Exit Sub
   End If
   
   For x = 1 To CustIdxRecNum
     Get CHandle, CustIdx(x), CustRec
     If CustRec.Deleted <> "Y" And QPTrim$(CustRec.SortName) <> "DELETED" And QPTrim$(CustRec.Inactive) <> "Y" Then
       Nextx = Nextx + 1
       vaSpread.Col = 2
       vaSpread.Row = Nextx
       vaSpread.Text = QPTrim$(CustRec.CustNumb)
       vaSpread.Col = 3
       vaSpread.Row = Nextx
       vaSpread.Text = QPTrim$(CustRec.CustName)
       vaSpread.Col = 4
       vaSpread.Row = Nextx
       vaSpread.Text = MakeRegDate(CustRec.VALID)
       vaSpread.Col = 5
       vaSpread.Row = Nextx
       vaSpread.Text = CustIdx(x)
     End If
   Next x
   vaSpread.MaxRows = Nextx
   Close
   
   Call FixSpread
   
End Sub

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "F1 &Turn Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    lblBalloon.Visible = True
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
  End If
End Sub

Private Sub fpcmbBalanceType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbBalanceType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbBalanceType.ListIndex = -1
  End If
  If fpcmbBalanceType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtHeading(0).SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbPrintFeesYN_Change()
  If QPTrim$(fpcmbPrintFeesYN.Text) = "" Then
    fpcmbPrintFeesYN.Text = "Yes"
  End If
  
  If fpcmbPrintFeesYN.Text = "Yes" Then
    fpcmbBalanceType.Enabled = True
  Else
    fpcmbBalanceType.Enabled = False
  End If
End Sub

Private Sub fpcmbPrintFeesYN_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbPrintFeesYN.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPrintFeesYN.ListIndex = -1
  End If
  If fpcmbPrintFeesYN.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbBalanceType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbPrintOpt_Change()
  If QPTrim$(fpcmbPrintOpt.Text) = "" Then
    fpcmbPrintOpt.Text = "Laser Form"
  End If
  
  If fpcmbPrintOpt.Text = "Laser Form" Then
    fptxtAuthorizedBy.Visible = True
    Label3.Visible = True
    fptxtHeading(3).Visible = False
    fptxtHeading(0) = Head1Laser$
    fptxtHeading(1) = Head2Laser$
    fptxtHeading(2) = Head3Laser$
    fptxtHeading(0).ToolTipText = "Enter the text you want to have printed on the first line of the laser license header. 'MUNICIPAL LICENSE' is the default entry."
    fptxtHeading(1).ToolTipText = "Enter the text you want to have printed on the second line of the laser license header."
    fptxtHeading(2).ToolTipText = "Enter the text you want to have printed on the third line of the laser license header."
    cmdAlign.Enabled = False
    fptxtIssDate.Enabled = True
    fpcmbSignature.Visible = True
    Label5.Visible = True
  Else
    fptxtAuthorizedBy.Visible = False
    Label3.Visible = False
    fptxtHeading(3).Visible = True
    fptxtHeading(0) = Head1Text$
    fptxtHeading(1) = Head2Text$
    fptxtHeading(2) = Head3Text$
    fptxtHeading(3) = Head4Text$
    cmdAlign.Enabled = True
    fptxtIssDate.Enabled = False
    fpcmbSignature.Visible = False
    Label5.Visible = False
  End If
End Sub

Private Sub fpcmbPrintOpt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbPrintOpt.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPrintOpt.ListIndex = -1
  End If
  If fpcmbPrintOpt.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      vaSpread.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbSignature_Change()
  If QPTrim$(fpcmbSignature.Text) = "" Then
    fpcmbSignature.Text = "Yes"
  End If
End Sub

Private Sub fpcmbSignature_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbSignature.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbSignature.ListIndex = -1
  End If
  If fpcmbSignature.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbPrintOpt.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fptxtHeading_Change(Index As Integer)
  If fpcmbPrintOpt.Text = "Laser Form" Then
    If Index = 0 Then
      Head1Laser$ = QPTrim$(fptxtHeading(Index).Text)
    ElseIf Index = 1 Then
      Head2Laser$ = QPTrim$(fptxtHeading(Index).Text)
    Else
      Head3Laser$ = QPTrim$(fptxtHeading(Index).Text)
    End If
  Else
    If Index = 0 Then
      Head1Text$ = QPTrim$(fptxtHeading(Index).Text)
    ElseIf Index = 1 Then
      Head2Text$ = QPTrim$(fptxtHeading(Index).Text)
    ElseIf Index = 2 Then
      Head3Text$ = QPTrim$(fptxtHeading(Index).Text)
    Else
      Head4Text$ = QPTrim$(fptxtHeading(Index).Text)
    End If
  End If
End Sub

Private Sub vaSpread_Click(ByVal Col As Long, ByVal Row As Long)
  'click on a row and an X will appear in the far left column...
  'click it again and the X disappears
  vaSpread.Col = 1
  vaSpread.Row = Row
  If vaSpread.Text = "X" Then
    vaSpread.Text = " "
  Else
    vaSpread.Text = "X"
  End If

End Sub

Private Sub PrintText()
  Dim ReportFile$
  Dim FF$, x As Double
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim CustCnt As Integer
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeHandle As Integer
  Dim NumOfARCatRecs As Integer
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Double
  Dim NameFlag As Boolean
  Dim NumFlag As Boolean
  Dim RptHandle As Integer
  Dim Page As Integer
  Dim TCat$, ZCnt&, cnt&
  Dim StoreExpireDate$
  Dim ExpireDate$
  Dim Year$
  Dim ChargeAcct$
  Dim NumOfTransRecs As Double
  Dim NextTransRec As Double
  Dim CategoryRecord1 As Integer
  Dim CategoryRecord2 As Integer
  Dim CategoryRecord3 As Integer
  Dim CategoryRecord4 As Integer
  Dim CategoryRecord5 As Integer
  Dim TotalBillAmt#
  Dim CustomerNumber As Integer
  Dim Prev As Long
  Dim CategoryDesc$
  Dim CategoryDesc1$
  Dim CategoryDesc2$
  Dim CategoryDesc3$
  Dim CategoryDesc4$
  Dim CategoryDesc5$, DidCnt As Integer
  Dim LICENSE#, ll As Integer
  Dim Heading1 As Integer
  Dim Heading2 As Integer
  Dim Heading3 As Integer
  Dim Heading4 As Integer
  Dim tab1 As Integer
  Dim tab2 As Integer
  Dim Tab3 As Integer
  Dim Tab4 As Integer
  Dim SHeading1$
  Dim SHeading2$
  Dim SHeading3$
  Dim SHeading4$
  Dim FromDate$
  Dim SCnt As Integer, LCnt As Integer
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim PrintFees As Boolean
  Dim IssFee As Double
  Dim BalanceFlag As Integer
  Dim XCnt As Integer
  Dim CustFee#, FeeAmt#
  Dim Prorate#
  Dim CatCode$, Snt&, Mult#
  Dim Revenue#
  Dim CHandle As Integer
  
  On Error GoTo ERRORSTUFF
  
  cmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  
  If Exist("artownsu.dat") Then
    OpenTownFile TownHandle
    Get TownHandle, 1, TownRec
    Close TownHandle
    IssFee = TownRec.IssFee
  Else
    IssFee = 0
  End If
  
  SHeading1$ = QPTrim$(fptxtHeading(0).Text)
  SHeading2$ = QPTrim$(fptxtHeading(1).Text)
  SHeading3$ = QPTrim$(fptxtHeading(2).Text)
  SHeading4$ = QPTrim$(fptxtHeading(3).Text)

  Heading1 = Len(SHeading1$)
  Heading2 = Len(SHeading2$)
  Heading3 = Len(SHeading3$)
  Heading4 = Len(SHeading4$)

  If Len(Heading1) > 0 Then tab1 = Heading1 / 2 Else tab1 = 0
  If Len(Heading2) > 0 Then tab2 = Heading2 / 2 Else tab2 = 0
  If Len(Heading3) > 0 Then Tab3 = Heading3 / 2 Else Tab3 = 0
  If Len(Heading4) > 0 Then Tab4 = Heading4 / 2 Else Tab4 = 0
  
  StoreExpireDate$ = fptxtVThru.Text
  ExpireDate$ = Mid(fptxtVThru.Text, 1, 6) + Mid(fptxtVThru.Text, 9, 2)
  Year$ = Mid(fptxtVThru.Text, 7, 4)
  
  FromDate$ = Mid(fptxtFromDate.Text, 1, 6) + Mid(fptxtFromDate.Text, 9, 2)
  ReportFile$ = "ARLCFREE.PRN"
  FF$ = Chr$(12)
  MaxLines = 58
  LineCnt = 0
  CustCnt = 0
  
  PrintFees = False
  If QPTrim$(fpcmbPrintFeesYN.Text) = "Yes" Then
    PrintFees = True
  End If
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  OpenCustFile CustHandle
  CustCnt = LOF(CustHandle) / Len(CustRec)
  
  If CustCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "There are no customers saved. Business license printing aborted."
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  ReDim XPrint(1 To 1) As Integer
  
  For x = 1 To CustCnt
    vaSpread.Col = 1
    vaSpread.Row = x
    If vaSpread.Text = "X" Then
      vaSpread.Col = 5
      XCnt = XCnt + 1
      ReDim Preserve XPrint(1 To XCnt) As Integer
      XPrint(XCnt) = CInt(vaSpread.Text)
    End If
  Next x
  
  If XCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "There are no customers selected. License printing aborted."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    Close
    fptxtFromDate.SetFocus
    Exit Sub
  End If
  
  frmBLShowPctComp.Label1 = "Printing Customer No Charge Business Licenses "
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  cmdHelp.Enabled = False
  cmdAlign.Enabled = False
  frmBLShowPctComp.cmdCancel.Visible = False
  
  If InStr(fpcmbBalanceType.Text, "Only") Then
    BalanceFlag = 1
  Else
    BalanceFlag = 2
  End If
  
  If PrintFees = False Then BalanceFlag = 1
  
  OpenCatCodeFile CodeHandle
  NumOfARCatRecs = LOF(CodeHandle) \ Len(CodeRec)
  If NumOfARCatRecs = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "Error: No category codes saved."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  For x = 1 To XCnt
    Get CustHandle, XPrint(x), CustRec
      If PrintFees = True Then
        Call SetFee(XPrint(x), NumOfARCatRecs)
        IssFee = ThisIssFee
      End If
      LICENSE# = Val(CustRec.LICENSE)
      CustomerNumber = Val(CustRec.CustNumb)
      For ll = 1 To 5
        Print #RptHandle,
      Next ll
      DidCnt = DidCnt + 1
      Print #RptHandle, Tab(37 - tab1); SHeading1$
      Print #RptHandle, Tab(37 - tab2); SHeading2$
      Print #RptHandle, Tab(37 - Tab3); SHeading3$
      Print #RptHandle, Tab(37 - Tab4); SHeading4$
      Print #RptHandle, Tab(66); fpBLYear.Text
      If CustRec.Prorate < 100 Then
        Print #RptHandle, Tab(11); "Cust #"; Tab(19); QPTrim$(Using("####0", CustomerNumber)); Tab(26); "Fee prorated at " + CStr(CustRec.Prorate) + "%"
      Else
        Print #RptHandle, Tab(11); "Cust #"; Tab(19); QPTrim$(Using("####0", CustomerNumber))
      End If
      Print #RptHandle, Tab(11); QPTrim$(CustRec.BillName)
      Print #RptHandle, Tab(11); QPTrim$(CustRec.ADDRESS1); Tab(58); Using("#######0", LICENSE#)
      Print #RptHandle, Tab(11); CustRec.ADDRESS2
      Print #RptHandle, Tab(11); RTrim$(CustRec.City); "  "; QPTrim$(CustRec.State); " "; QPTrim$(CustRec.ZipCode)
      Print #RptHandle, Tab(55); FromDate$;
      Print #RptHandle, Tab(64); ExpireDate$
      Print #RptHandle,
      Print #RptHandle,
      Print #RptHandle,
      Print #RptHandle, Tab(11); QPTrim$(CustRec.CustName)
      Print #RptHandle,
      Print #RptHandle,
      SCnt = 23
      If GetCatRecNum(QPTrim$(CustRec.BILLCAT1)) = 0 Then GoTo To2
      Print #RptHandle, Tab(5); QPTrim$(CustRec.BILLCAT1);
      If PrintFees = True Then
        Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC1);
        Print #RptHandle, Tab(62); Using("####0.00", FeeAmt1#)
      Else
        Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC1)
      End If
      SCnt = SCnt + 1
To2:
      If GetCatRecNum(QPTrim$(CustRec.BILLCAT2)) = 0 Then GoTo To3 'ExitFormPrint1
      Print #RptHandle, Tab(5); QPTrim$(CustRec.BILLCAT2);
      If PrintFees = True Then
        Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC2);
        Print #RptHandle, Tab(62); Using("####0.00", FeeAmt2#)
      Else
        Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC2)
      End If
      SCnt = SCnt + 1
To3:
      If GetCatRecNum(QPTrim$(CustRec.BILLCAT3)) = 0 Then GoTo To4 'ExitFormPrint1
      Print #RptHandle, Tab(5); QPTrim$(CustRec.BILLCAT3);
      If PrintFees = True Then
        Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC3);
        Print #RptHandle, Tab(62); Using("####0.00", FeeAmt3#)
      Else
        Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC3)
      End If
      SCnt = SCnt + 1
To4:
     If GetCatRecNum(QPTrim$(CustRec.BILLCAT4)) = 0 Then GoTo To5 'ExitFormPrint1
      Print #RptHandle, Tab(5); QPTrim$(CustRec.BILLCAT4);
      If PrintFees = True Then
        Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC4);
        Print #RptHandle, Tab(62); Using("####0.00", FeeAmt4#)
      Else
        Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC4)
      End If
      SCnt = SCnt + 1
To5:
     If GetCatRecNum(QPTrim$(CustRec.BILLCAT5)) = 0 Then GoTo ExitFormPrint1
      Print #RptHandle, Tab(5); QPTrim$(CustRec.BILLCAT5);
      If PrintFees = True Then
        Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC5);
        Print #RptHandle, Tab(62); Using("####0.00", FeeAmt5#)
      Else
        Print #RptHandle, Tab(15); QPTrim$(CustRec.DESC5)
      End If
      SCnt = SCnt + 1

ExitFormPrint1:
      If IssFee > 0 And PrintFees = True Then
        Print #RptHandle, Tab(15); "ISSUE FEE"; Tab(62); Using("####0.00", OldRound(IssFee))
        SCnt = SCnt + 1
      End If
      For LCnt = SCnt To 31
        Print #RptHandle,
      Next
      Print #RptHandle, ""

      For LCnt = 33 To 35
        Print #RptHandle, ""
      Next LCnt
      'Calc Total License Amount Here
      TotalBillAmt# = OldRound(FeeAmt1# + FeeAmt2# + FeeAmt3# + FeeAmt4# + FeeAmt5#)
      TotalBillAmt# = OldRound(TotalBillAmt# + IssFee)
      If PrintFees = True Then
        Print #RptHandle, Tab(62); Using("####0.00", TotalBillAmt#) ' - OldRound(CustRec.AcctBal))
      Else
        Print #RptHandle, "No"
      End If
      Print #RptHandle,
      Print #RptHandle,
      Print #RptHandle,
      If BalanceFlag = 1 Then
        Print #RptHandle,
      Else
        Print #RptHandle, Tab(62); Using("####0.00", ThisAcctBal)
      End If
      Print #RptHandle,
      Print #RptHandle, "~"
      MainLog ("Non-posting business license for customer # " + QPTrim$(CustRec.CustNumb) + "/" + QPTrim$(CustRec.CustName) + " printed to screen.")
      frmBLShowPctComp.ShowPctComp x, XCnt
  Next x
  
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  cmdHelp.Enabled = True
  cmdAlign.Enabled = True
  Print #RptHandle, Chr$(12);
  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
  Close         'Close all open files now
  
  'DidPrint works with the printing procedure to log an entry
  'if the user actually printed out the license forms
  
  DidPrint = 2
  
  ViewPrint ReportFile$, "Business License Printing", True
  KillFile ReportFile$
  
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLPrintLicNoPost", "PrintText", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    ClearInUse PWcnt
    Terminate
    Unload Me

End Sub

Private Sub PrintGraphics()

  Dim ReportFile$
  Dim x As Double, y As Integer
  Dim CustCnt As Integer
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeHandle As Integer
  Dim NumOfARCatRecs As Integer
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Double
  Dim CustNameIdxRec As CustSearchNameIdxType
  Dim CustNumIdxRec As CustNumIdxType
  Dim IdxHandle As Integer
  Dim NumOfCustIdxRecs As Double
  Dim NameFlag As Boolean
  Dim NumFlag As Boolean
  Dim RptHandle As Integer
  Dim TCat$, ZCnt&, cnt&
  Dim StoreExpireDate$
  Dim ExpireDate$
  Dim Year$
  Dim NumOfTransRecs As Double
  Dim NextTransRec As Double
  Dim CategoryRecord1 As Integer
  Dim CategoryRecord2 As Integer
  Dim CategoryRecord3 As Integer
  Dim CategoryRecord4 As Integer
  Dim CategoryRecord5 As Integer
  Dim TotalBillAmt#
  Dim PostDate$
  Dim CustLicNum$
  Dim Prev As Long
  Dim CategoryDesc$
  Dim CategoryDesc1$
  Dim CategoryDesc2$
  Dim CategoryDesc3$
  Dim CategoryDesc4$
  Dim CategoryDesc5$, DidCnt As Integer
  Dim LICENSE#, ll As Integer
  Dim TransRec As ARTransRecType
  Dim THandle As Integer
  Dim SHeading1$
  Dim SHeading2$
  Dim SHeading3$
  Dim IssueDate$
  Dim SCnt As Integer, LCnt As Integer
  Dim TempHandle As Integer
  Dim TempRec As TempTransPostType
  Dim TempNum As Integer
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim NumOfTempRecs As Integer
  Dim PrintFees As Boolean
  Dim dlm$
  Dim One As Integer
  Dim DHandle As Integer
  Dim PCnt As Integer
  Dim ThisCat As String * 35
  Dim BalanceFlag As Integer
  Dim CustFee#, FeeAmt#
  Dim Prorate#
  Dim CatCode$, Snt&, Mult#
  Dim Revenue#
  Dim CHandle As Integer
  Dim XCnt As Integer
  Dim AddEmptyFields As Integer
  
  On Error GoTo ERRORSTUFF
  
  If fpcmbSignature.Text = "Yes" Then
    PrintSign = True
  Else
    PrintSign = False
  End If
  
  dlm$ = "~"
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  
  SHeading1$ = QPTrim$(fptxtHeading(0).Text)
  SHeading2$ = QPTrim$(fptxtHeading(1).Text)
  SHeading3$ = QPTrim$(fptxtHeading(2).Text)

  StoreExpireDate$ = fptxtVThru.Text
  ExpireDate$ = fptxtVThru.Text
  Year$ = Mid(fptxtVThru.Text, 7, 4)
  
  IssueDate$ = fptxtIssDate.Text
  PostDate$ = fptxtIssDate.Text
  ReportFile$ = "BLRPTS\ARLASER.RPT"  'Report File Name
  CustCnt = 0
  
  PrintFees = False
  If QPTrim$(fpcmbPrintFeesYN.Text) = "Yes" Then
    PrintFees = True
  End If
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  OpenCustFile CustHandle
  CustCnt = LOF(CustHandle) / Len(CustRec)
  
  ReDim XPrint(1 To 1) As Integer
  
  For x = 1 To CustCnt
    vaSpread.Col = 1
    vaSpread.Row = x
    If vaSpread.Text = "X" Then
      vaSpread.Col = 5
      XCnt = XCnt + 1
      ReDim Preserve XPrint(1 To XCnt) As Integer
      XPrint(XCnt) = CInt(vaSpread.Text)
    End If
  Next x
  
  If XCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "There are no customers selected. License printing aborted."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    Close
    fptxtIssDate.SetFocus
    Exit Sub
  End If
  
  OpenTransFile THandle
  NumOfTransRecs = LOF(THandle) / Len(TransRec)
  Close THandle
  NextTransRec = NumOfTransRecs + 1
  
  frmBLShowPctComp.Label1 = "Printing Customer Business Licenses"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  cmdHelp.Enabled = False
  frmBLShowPctComp.cmdCancel.Visible = False

  If InStr(fpcmbBalanceType.Text, "Only") Then
    BalanceFlag = 1
  Else
    BalanceFlag = 2
  End If
  
  If PrintFees = False Then BalanceFlag = 1
  OpenCatCodeFile CodeHandle
  NumOfARCatRecs = LOF(CodeHandle) \ Len(CodeRec)
  If NumOfARCatRecs = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "Error: No category codes saved."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  For x = 1 To XCnt
    Get CustHandle, XPrint(x), CustRec
      Call SetFee(XPrint(x), NumOfARCatRecs)
      LICENSE# = Val(CustRec.LICENSE)
      CustLicNum = Val(CustRec.LICENSE)
      CategoryDesc$ = ""
      CategoryRecord1 = 0
      
      '                     0                1                 2             3
      Print #RptHandle, SHeading1$; dlm; SHeading2$; dlm; SHeading3$; dlm; Year$; dlm;
      If CustRec.Prorate < 100 Then
        '                       4                        5
        Print #RptHandle, CustLicNum; dlm; CStr(CustRec.Prorate); dlm;
      Else
        '                       4               5
        Print #RptHandle, CustLicNum; dlm; ""; dlm;
      End If
      '                              6                      7
      Print #RptHandle, QPTrim$(CustRec.Contact); dlm; LICENSE#; dlm;
      '                       8                          9
      Print #RptHandle, CustRec.ADDRESS1; dlm; QPTrim$(CustRec.City) + " ," + QPTrim$(CustRec.State) + " " + QPTrim$(CustRec.ZipCode); dlm;
      '                     10               11                      12                                  13
      Print #RptHandle, IssueDate$; dlm; ExpireDate$; dlm; QPTrim$(CustRec.CustName); dlm; QPTrim$(fptxtAuthorizedBy.Text); dlm;
      
      AddEmptyFields = 0
      
      If QPTrim$(CustRec.BILLCAT1) <> "" Then
        If PrintFees = True Then
          '                             14                          15                            16
          Print #RptHandle, QPTrim$(CustRec.BILLCAT1); dlm; GetCatDesc(CustRec.BILLCAT1); dlm; FeeAmt1#; dlm;
        Else
          '
          Print #RptHandle, QPTrim$(CustRec.BILLCAT1); dlm; GetCatDesc(CustRec.BILLCAT1); dlm; ""; dlm;
        End If
      Else
        AddEmptyFields = AddEmptyFields + 3
      End If
      
      If QPTrim$(CustRec.BILLCAT2) <> "" Then
        If PrintFees = True Then
          '                             17                          18                            19
          Print #RptHandle, QPTrim$(CustRec.BILLCAT2); dlm; GetCatDesc(CustRec.BILLCAT2); dlm; FeeAmt2#; dlm;
        Else
          '
          Print #RptHandle, QPTrim$(CustRec.BILLCAT2); dlm; GetCatDesc(CustRec.BILLCAT2); dlm; ""; dlm;
        End If
      Else
        '
        AddEmptyFields = AddEmptyFields + 3
      End If
        
      If QPTrim$(CustRec.BILLCAT3) <> "" Then
        If PrintFees = True Then
          '                             20                          21                            22
          Print #RptHandle, QPTrim$(CustRec.BILLCAT3); dlm; GetCatDesc(CustRec.BILLCAT3); dlm; FeeAmt3#; dlm;
        Else
          '
          Print #RptHandle, QPTrim$(CustRec.BILLCAT3); dlm; GetCatDesc(CustRec.BILLCAT3); dlm; ""; dlm;
        End If
      Else
        '
        AddEmptyFields = AddEmptyFields + 3
      End If
      
      If QPTrim$(CustRec.BILLCAT4) <> "" Then
        If PrintFees = True Then
          '                             23                          24                            25
          Print #RptHandle, QPTrim$(CustRec.BILLCAT4); dlm; GetCatDesc(CustRec.BILLCAT4); dlm; FeeAmt4#; dlm;
        Else
          '
          Print #RptHandle, QPTrim$(CustRec.BILLCAT4); dlm; GetCatDesc(CustRec.BILLCAT4); dlm; ""; dlm;
        End If
      Else
        '
        AddEmptyFields = AddEmptyFields + 3
      End If
      
      If QPTrim$(CustRec.BILLCAT5) <> "" Then
        If PrintFees = True Then
          '                             26                          27                            28
          Print #RptHandle, QPTrim$(CustRec.BILLCAT5); dlm; GetCatDesc(CustRec.BILLCAT5); dlm; FeeAmt5#; dlm;
        Else
          '
          Print #RptHandle, QPTrim$(CustRec.BILLCAT5); dlm; GetCatDesc(CustRec.BILLCAT5); dlm; ""; dlm;
        End If
      Else
        '
        AddEmptyFields = AddEmptyFields + 3
      End If
      
      For y = 1 To AddEmptyFields
        '
        Print #RptHandle, ""; dlm;
      Next y
      
      
      If PrintFees = True Then
        
        If OldRound(ThisIssFee#) > 0 Then
          '                            29
          Print #RptHandle, OldRound(ThisIssFee#); dlm;
        Else
          '                 29
          Print #RptHandle, "0"; dlm;
        End If
        
      'Calc Total License Amount Here
      Else
        '                  29
        Print #RptHandle, "0"; dlm;
      End If
        
      TotalBillAmt# = OldRound(FeeAmt1# + FeeAmt2# + FeeAmt3# + FeeAmt4# + FeeAmt5# + ThisIssFee)
      
      If PrintFees = True Then
        '                      30
        Print #RptHandle, TotalBillAmt#; dlm;
      Else
        '                 30
        Print #RptHandle, "No"; dlm;
      End If
      '                         31
      Print #RptHandle, fptxtFromDate.Text; dlm;
      
      If BalanceFlag = 2 Then
        '                      32                   33                 34                  35                          36                        37
        Print #RptHandle, CustRec.AcctBal; dlm; ThisAcctBal#; dlm; BalanceFlag; dlm; fpBLYear.Text; dlm; QPTrim$(CustRec.BillName); dlm; (CustRec.ServAdd)
      Else
        '                 32     33          34                  35                           36                                37
        Print #RptHandle, 0; dlm; 0; dlm; BalanceFlag; dlm; fpBLYear.Text; dlm; QPTrim$(CustRec.BillName); dlm; QPTrim$(CustRec.ServAdd)
      End If
      
      MainLog ("Non-posting business license for customer # " + QPTrim$(CustRec.CustNumb) + "/" + QPTrim$(CustRec.CustName) + " printed to screen. Page # " + CStr(x) + ".")
      frmBLShowPctComp.ShowPctComp x, XCnt
  Next x
  
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  cmdHelp.Enabled = True
  Close         'Close all open files now

    
  'DidPrint is used in conjunction with the printing process
  'to determine if the user actually printed the forms and if so
  'then the MainLog is updated and updated by page number...if the user
  'prints only one of a selection of possible forms then that page number
  'is recorded and corresponds with the MainLog entry above so you
  'can determine which of the available forms are available
  
  DidPrint = 2
  
  arBLLaser.Show
  frmBLLoadReport.Show
  
  Exit Sub

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLPrintLicNoPost", "PrintGraphics", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    ClearInUse PWcnt
    Terminate
  

End Sub

Private Sub FixSpread()
  Dim COne As Integer
  Dim CTwo As Integer
  Dim CThree As Integer
  Dim CFour As Integer
  Dim CFive As Integer
  Dim CSix As Integer
  Dim cnt As Integer
  '-1 means all rows or all columns....0 means headers
'    GoTo SkipAdjust
  Select Case ScreenW
    Case 1280
      If Screen.TwipsPerPixelX <> 12 Then
        COne = 5
        coladj = 10
        vaSpread.FontSize = 18
        vaSpread.RowHeight(-1) = 22
        vaSpread.RowHeight(0) = 22
      Else
        COne = 13
        coladj = 8.1
        vaSpread.RowHeight(-1) = 18
        vaSpread.RowHeight(0) = 18
      End If
    Case 1152
      If Screen.TwipsPerPixelX <> 12 Then
        COne = 14
        coladj = 8
        vaSpread.FontSize = 14
        vaSpread.RowHeight(0) = 18.5
        vaSpread.RowHeight(-1) = 18.5
      Else
        COne = 5.65
        coladj = 3.75
        vaSpread.RowHeight(0) = 16
        vaSpread.RowHeight(-1) = 17
      End If
    Case 1024
      If Screen.TwipsPerPixelX <> 12 Then
        COne = 13.49
        coladj = 5.65
        vaSpread.RowHeight(0) = 14
        vaSpread.RowHeight(-1) = 14
      Else
        COne = 0.75
        coladj = -0.5
      End If
    Case 800
      COne = 0
      coladj = -1.2
      vaSpread.Font.Size = 12
      vaSpread.RowHeight(-1) = 14
    Case Else
  End Select
SkipAdjust:
    vaSpread.ColWidth(1) = vaSpread.ColWidth(1)
    vaSpread.ColWidth(2) = vaSpread.ColWidth(2) + coladj
    vaSpread.ColWidth(3) = vaSpread.ColWidth(3) + coladj
    vaSpread.ColWidth(4) = vaSpread.ColWidth(4)

End Sub

Private Sub SetFee(ThisCust As Integer, NumOfARCatRecs As Integer)

  Dim CodeHandle As Integer
  Dim CodeRec As ARNewCatCodeRecType
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim ProrateFlag As Boolean
  Dim ProAt#, CatCode$, Snt&
  Dim FeeAmt#, Mult#, Revenue#
  Dim x As Double, Prorate#
  Dim TempRec As TempTransPostType
  Dim TempRec2 As TempTransPostType
  Dim OverLic As Double
  Dim OverPen As Double
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim IssFee As Double
  Dim TempIssFee As Double
  
  On Error GoTo ERRORSTUFF
  
  OpenTownFile TownHandle
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  
  FeeAmt1# = 0
  FeeAmt2# = 0
  FeeAmt3# = 0
  FeeAmt4# = 0
  FeeAmt5# = 0
  ThisIssFee# = 0
  ThisAcctBal# = 0
  
  OpenCatCodeFile CodeHandle
  
  OpenCustFile CustHandle
  Get CustHandle, ThisCust, CustRec
  Close CustHandle
  
  IssFee = TownRec.IssFee
  ProrateFlag = False
  
  'Clear OverPen & OverLic
  OverPen = 0
  OverLic = 0
  'Clear TempRec from previous calculations
  TempRec = TempRec2
  TempRec.CreditUsed = False
  'assign the temp fields currrent amounts...temp will
  'become permanent during posting
  TempRec.PenBal = CustRec.PenBal
  TempRec.LicBal = CustRec.LicBal
  
  TempRec.CatFeeBal1 = CustRec.FeeLicBal1
  TempRec.CatFeeBal2 = CustRec.FeeLicBal2
  TempRec.CatFeeBal3 = CustRec.FeeLicBal3
  TempRec.CatFeeBal4 = CustRec.FeeLicBal4
  TempRec.CatFeeBal5 = CustRec.FeeLicBal5
  
  'if a negative (credit) amount exists figure assign credit amount to
  'OverLic as a positive amount
  If TempRec.LicBal < 0 Then OverLic = Abs(TempRec.LicBal)
  
  If IssFee > 0 And OverLic > 0 Then 'we need to reduce the
  'the credit license balance to reflect the cost of the Issue Fee first
  'so we pay for the issuance fee out of whatever credit amount exists
    GoSub CalcLicBal 'if we are going to reduce the issuance fee
    'by the amount of the credit then we must also go into the
    'license balances and bring them closer to 0 because at
    'this point we know we have at least one license balance credit
  End If
  
  TempRec.LicBal = TempRec.CatFeeBal1 + TempRec.CatFeeBal2 + TempRec.CatFeeBal3 + TempRec.CatFeeBal4 + TempRec.CatFeeBal5
  If CustRec.PenBal < 0 Then OverPen = Abs(CustRec.PenBal)
  
  If IssFee > 0 And OverPen > 0 Then
    TempRec.PenBal = TempRec.PenBal + OverPen
    If OverPen >= IssFee Then
      IssFee = 0
      OverPen = OverPen - IssFee
    ElseIf OverPen < IssFee Then
      IssFee = IssFee - OverPen
      OverPen = 0
    End If
  End If
  
  TempRec.IssFeeBal = CustRec.IssuanceBal + IssFee
  TempRec.IssFee = IssFee 'TownRec.IssFee
  
  Prorate# = CustRec.Prorate
  
  If Prorate# >= 100 Or Prorate# <= 0 Then
    Prorate# = 100
  Else
    ProrateFlag = True
    Prorate# = OldRound(Prorate# * 0.01)
  End If

  CatCode$ = QPTrim$(CustRec.BILLCAT1)
  If Len(CatCode$) > 0 Then
    For Snt& = 1 To NumOfARCatRecs
      Get CodeHandle, Snt&, CodeRec
      'find the code that matches this category
      If QPTrim$(CodeRec.CatCode) = CatCode$ Then
        TempRec.CatCodeRec1 = Snt&
        'if it's a flat fate then do this
        If CodeRec.CodeType = "F" Then
          'get fee and assign accordingly
          If ProrateFlag = True Then
            TempRec.CatFee1 = OldRound(Prorate# * CodeRec.Fee)
          Else
            TempRec.CatFee1 = CodeRec.Fee
          End If
          If TempRec.CatFee1 < 0 Then TempRec.CatFee1 = 0
          'if there remains a negative penalty balance or a negative
          'category balance then these negative amounts are applied
          'to the new fee calculations
          If OverPen > 0 Or TempRec.CatFeeBal1 < 0 Then
            Call ApplyCredits2ThisFee(TempRec.CatFeeBal1, TempRec.CatFee1, OverPen)
            TempRec.CreditUsed = True
            TempRec.PenBal = OverPen
          Else
            TempRec.CatFeeBal1 = TempRec.CatFeeBal1 + TempRec.CatFee1
          End If
          GoTo C2
        End If
        'if it's a multiplier then do this
        If CodeRec.CodeType = "M" Then
          'get number of multipliers for this customer
          Mult = CustRec.REV1
          If ProrateFlag = True Then
            TempRec.CatFee1 = OldRound(Mult * CodeRec.Fee)
            TempRec.CatFee1 = OldRound(TempRec.CatFee1 * Prorate#)
          Else
            TempRec.CatFee1 = OldRound(Mult * CodeRec.Fee)
          End If
          If TempRec.CatFee1 < 0 Then TempRec.CatFee1 = 0
          'now apply any credits
          If OverPen > 0 Or TempRec.CatFeeBal1 < 0 Then
            Call ApplyCredits2ThisFee(TempRec.CatFeeBal1, TempRec.CatFee1, OverPen)
            TempRec.CreditUsed = True
            TempRec.PenBal = OverPen
          Else
            TempRec.CatFeeBal1 = TempRec.CatFeeBal1 + TempRec.CatFee1
          End If
          GoTo C2
        End If
        If CodeRec.CodeType = "S" Then
          'if it's a step rate then find the level
          'that applies to this customer's revenue
          Revenue# = CustRec.REV1
          If ProrateFlag = True Then
            If Revenue# <= CodeRec.Recpt1 Then
              TempRec.CatFee1 = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
              If TempRec.CatFee1 < CodeRec.BaseAmt1 Then TempRec.CatFee1 = CodeRec.BaseAmt1
              TempRec.CatFee1 = OldRound(Prorate# * TempRec.CatFee1)
              If OverPen > 0 Or TempRec.CatFeeBal1 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal1, TempRec.CatFee1, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal1 = TempRec.CatFeeBal1 + TempRec.CatFee1
              End If
              GoTo C2
            End If
            If Revenue# <= CodeRec.Recpt2 Then
              TempRec.CatFee1 = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
              If TempRec.CatFee1 < CodeRec.BaseAmt2 Then TempRec.CatFee1 = CodeRec.BaseAmt2
              TempRec.CatFee1 = OldRound(Prorate# * TempRec.CatFee1)
              If OverPen > 0 Or TempRec.CatFeeBal1 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal1, TempRec.CatFee1, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal1 = TempRec.CatFeeBal1 + TempRec.CatFee1
              End If
              GoTo C2
            End If
            If Revenue# <= CodeRec.Recpt3 Then
              TempRec.CatFee1 = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
              If TempRec.CatFee1 < CodeRec.BaseAmt3 Then TempRec.CatFee1 = CodeRec.BaseAmt3
              TempRec.CatFee1 = OldRound(Prorate# * TempRec.CatFee1)
              If OverPen > 0 Or TempRec.CatFeeBal1 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal1, TempRec.CatFee1, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal1 = TempRec.CatFeeBal1 + TempRec.CatFee1
              End If
              GoTo C2
            End If
            If Revenue# <= CodeRec.Recpt4 Then
              TempRec.CatFee1 = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
              If TempRec.CatFee1 < CodeRec.BaseAmt4 Then TempRec.CatFee1 = CodeRec.BaseAmt4
              TempRec.CatFee1 = OldRound(Prorate# * TempRec.CatFee1)
              If OverPen > 0 Or TempRec.CatFeeBal1 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal1, TempRec.CatFee1, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal1 = TempRec.CatFeeBal1 + TempRec.CatFee1
              End If
              GoTo C2
            End If
            If Revenue# <= CodeRec.Recpt5 Then
              TempRec.CatFee1 = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
              If TempRec.CatFee1 < CodeRec.BaseAmt5 Then TempRec.CatFee1 = CodeRec.BaseAmt5
              TempRec.CatFee1 = OldRound(Prorate# * TempRec.CatFee1)
              If OverPen > 0 Or TempRec.CatFeeBal1 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal1, TempRec.CatFee1, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal1 = TempRec.CatFeeBal1 + TempRec.CatFee1
              End If
              GoTo C2
            End If
            If Revenue# <= CodeRec.Recpt6 Then
              TempRec.CatFee1 = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
              If TempRec.CatFee1 < CodeRec.BaseAmt6 Then TempRec.CatFee1 = CodeRec.BaseAmt6
              TempRec.CatFee1 = OldRound(Prorate# * TempRec.CatFee1)
              If OverPen > 0 Or TempRec.CatFeeBal1 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal1, TempRec.CatFee1, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal1 = TempRec.CatFeeBal1 + TempRec.CatFee1
              End If
              GoTo C2
            End If
          Else 'ProrateFlag = False
            If Revenue# <= CodeRec.Recpt1 Then
              TempRec.CatFee1 = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
              If TempRec.CatFee1 < CodeRec.BaseAmt1 Then TempRec.CatFee1 = CodeRec.BaseAmt1
              If OverPen > 0 Or TempRec.CatFeeBal1 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal1, TempRec.CatFee1, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal1 = TempRec.CatFeeBal1 + TempRec.CatFee1
              End If
              GoTo C2
            End If
            If Revenue# <= CodeRec.Recpt2 Then
              TempRec.CatFee1 = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
              If TempRec.CatFee1 < CodeRec.BaseAmt2 Then TempRec.CatFee1 = CodeRec.BaseAmt2
              If OverPen > 0 Or TempRec.CatFeeBal1 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal1, TempRec.CatFee1, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal1 = TempRec.CatFeeBal1 + TempRec.CatFee1
              End If
              GoTo C2
            End If
            If Revenue# <= CodeRec.Recpt3 Then
              TempRec.CatFee1 = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
              If TempRec.CatFee1 < CodeRec.BaseAmt3 Then TempRec.CatFee1 = CodeRec.BaseAmt3
              If OverPen > 0 Or TempRec.CatFeeBal1 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal1, TempRec.CatFee1, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal1 = TempRec.CatFeeBal1 + TempRec.CatFee1
              End If
              GoTo C2
            End If
            If Revenue# <= CodeRec.Recpt4 Then
              TempRec.CatFee1 = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
              If TempRec.CatFee1 < CodeRec.BaseAmt4 Then TempRec.CatFee1 = CodeRec.BaseAmt4
              If OverPen > 0 Or TempRec.CatFeeBal1 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal1, TempRec.CatFee1, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal1 = TempRec.CatFeeBal1 + TempRec.CatFee1
              End If
              GoTo C2
            End If
            If Revenue# <= CodeRec.Recpt5 Then
              TempRec.CatFee1 = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
              If TempRec.CatFee1 < CodeRec.BaseAmt5 Then TempRec.CatFee1 = CodeRec.BaseAmt5
              If OverPen > 0 Or TempRec.CatFeeBal1 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal1, TempRec.CatFee1, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal1 = TempRec.CatFeeBal1 + TempRec.CatFee1
              End If
              GoTo C2
            End If
            If Revenue# <= CodeRec.Recpt6 Then
              TempRec.CatFee1 = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
              If TempRec.CatFee1 < CodeRec.BaseAmt6 Then TempRec.CatFee1 = CodeRec.BaseAmt6
              If OverPen > 0 Or TempRec.CatFeeBal1 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal1, TempRec.CatFee1, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal1 = TempRec.CatFeeBal1 + TempRec.CatFee1
              End If
              GoTo C2
            End If
          End If
        End If
      End If  'End Test for Code
    Next Snt&
  Else
    TempRec.CatFee1 = 0
  End If      'End Test for Cat 1


C2:             'Category #2
  CatCode$ = QPTrim$(CustRec.BILLCAT2)
  If Len(CatCode$) > 0 Then
    For Snt& = 1 To NumOfARCatRecs
      Get CodeHandle, Snt&, CodeRec
      If QPTrim$(CodeRec.CatCode) = CatCode$ Then
        TempRec.CatCodeRec2 = Snt&
        If CodeRec.CodeType = "F" Then
          If ProrateFlag = True Then
            TempRec.CatFee2 = OldRound(Prorate# * CodeRec.Fee)
          Else
            TempRec.CatFee2 = CodeRec.Fee
          End If
          If TempRec.CatFee2 < 0 Then TempRec.CatFee2 = 0
          If OverPen > 0 Or TempRec.CatFeeBal2 < 0 Then
            Call ApplyCredits2ThisFee(TempRec.CatFeeBal2, TempRec.CatFee2, OverPen)
            TempRec.CreditUsed = True
            TempRec.PenBal = OverPen
          Else
            TempRec.CatFeeBal2 = TempRec.CatFeeBal2 + TempRec.CatFee2
          End If
          GoTo C3
        End If
        If CodeRec.CodeType = "M" Then
          Mult = CustRec.REV2
          If ProrateFlag = True Then
            TempRec.CatFee2 = OldRound(Mult * CodeRec.Fee)
            TempRec.CatFee2 = OldRound(TempRec.CatFee2 * Prorate#)
          Else
            TempRec.CatFee2 = OldRound(Mult * CodeRec.Fee)
          End If
          If TempRec.CatFee2 < 0 Then TempRec.CatFee2 = 0
          If OverPen > 0 Or TempRec.CatFeeBal2 < 0 Then
            Call ApplyCredits2ThisFee(TempRec.CatFeeBal2, TempRec.CatFee2, OverPen)
            TempRec.CreditUsed = True
            TempRec.PenBal = OverPen
          Else
            TempRec.CatFeeBal2 = TempRec.CatFeeBal2 + TempRec.CatFee2
          End If
          GoTo C3
        End If
        If CodeRec.CodeType = "S" Then
          Revenue# = CustRec.REV2
          If ProrateFlag = True Then
            If Revenue# <= CodeRec.Recpt1 Then
              TempRec.CatFee2 = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
              If TempRec.CatFee2 < CodeRec.BaseAmt1 Then TempRec.CatFee2 = CodeRec.BaseAmt1
              TempRec.CatFee2 = OldRound(Prorate# * TempRec.CatFee2)
              If OverPen > 0 Or TempRec.CatFeeBal2 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal2, TempRec.CatFee2, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal2 = TempRec.CatFeeBal2 + TempRec.CatFee2
              End If
              GoTo C3
            End If
            If Revenue# <= CodeRec.Recpt2 Then
              TempRec.CatFee2 = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
              If TempRec.CatFee2 < CodeRec.BaseAmt2 Then TempRec.CatFee2 = CodeRec.BaseAmt2
              TempRec.CatFee2 = OldRound(Prorate# * TempRec.CatFee2)
              If OverPen > 0 Or TempRec.CatFeeBal2 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal2, TempRec.CatFee2, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal2 = TempRec.CatFeeBal2 + TempRec.CatFee2
              End If
              GoTo C3
            End If
            If Revenue# <= CodeRec.Recpt3 Then
              TempRec.CatFee2 = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
              If TempRec.CatFee2 < CodeRec.BaseAmt3 Then TempRec.CatFee2 = CodeRec.BaseAmt3
              TempRec.CatFee2 = OldRound(Prorate# * TempRec.CatFee2)
              If OverPen > 0 Or TempRec.CatFeeBal2 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal2, TempRec.CatFee2, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal2 = TempRec.CatFeeBal2 + TempRec.CatFee2
              End If
              GoTo C3
            End If
            If Revenue# <= CodeRec.Recpt4 Then
              TempRec.CatFee2 = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
              If TempRec.CatFee2 < CodeRec.BaseAmt4 Then TempRec.CatFee2 = CodeRec.BaseAmt4
              TempRec.CatFee2 = OldRound(Prorate# * TempRec.CatFee2)
              If OverPen > 0 Or TempRec.CatFeeBal2 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal2, TempRec.CatFee2, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal2 = TempRec.CatFeeBal2 + TempRec.CatFee2
              End If
              GoTo C3
            End If
            If Revenue# <= CodeRec.Recpt5 Then
              TempRec.CatFee2 = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
              If TempRec.CatFee2 < CodeRec.BaseAmt5 Then TempRec.CatFee2 = CodeRec.BaseAmt5
              TempRec.CatFee2 = OldRound(Prorate# * TempRec.CatFee2)
              If OverPen > 0 Or TempRec.CatFeeBal2 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal2, TempRec.CatFee2, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal2 = TempRec.CatFeeBal2 + TempRec.CatFee2
              End If
              GoTo C3
            End If
            If Revenue# <= CodeRec.Recpt6 Then
              TempRec.CatFee2 = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
              If TempRec.CatFee2 < CodeRec.BaseAmt6 Then TempRec.CatFee2 = CodeRec.BaseAmt6
              TempRec.CatFee2 = OldRound(Prorate# * TempRec.CatFee2)
              If OverPen > 0 Or TempRec.CatFeeBal2 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal2, TempRec.CatFee2, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal2 = TempRec.CatFeeBal2 + TempRec.CatFee2
              End If
              GoTo C3
            End If
          Else
            If Revenue# <= CodeRec.Recpt1 Then
              TempRec.CatFee2 = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
              If TempRec.CatFee2 < CodeRec.BaseAmt1 Then TempRec.CatFee2 = CodeRec.BaseAmt1
              If OverPen > 0 Or TempRec.CatFeeBal2 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal2, TempRec.CatFee2, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal2 = TempRec.CatFeeBal2 + TempRec.CatFee2
              End If
              GoTo C3
            End If
            If Revenue# <= CodeRec.Recpt2 Then
              TempRec.CatFee2 = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
              If TempRec.CatFee2 < CodeRec.BaseAmt2 Then TempRec.CatFee2 = CodeRec.BaseAmt2
              If OverPen > 0 Or TempRec.CatFeeBal2 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal2, TempRec.CatFee2, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal2 = TempRec.CatFeeBal2 + TempRec.CatFee2
              End If
              GoTo C3
            End If
            If Revenue# <= CodeRec.Recpt3 Then
              TempRec.CatFee2 = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
              If TempRec.CatFee2 < CodeRec.BaseAmt3 Then TempRec.CatFee2 = CodeRec.BaseAmt3
              If OverPen > 0 Or TempRec.CatFeeBal2 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal2, TempRec.CatFee2, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal2 = TempRec.CatFeeBal2 + TempRec.CatFee2
              End If
              GoTo C3
            End If
            If Revenue# <= CodeRec.Recpt4 Then
              TempRec.CatFee2 = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
              If TempRec.CatFee2 < CodeRec.BaseAmt4 Then TempRec.CatFee2 = CodeRec.BaseAmt4
              If OverPen > 0 Or TempRec.CatFeeBal2 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal2, TempRec.CatFee2, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal2 = TempRec.CatFeeBal2 + TempRec.CatFee2
              End If
              GoTo C3
            End If
            If Revenue# <= CodeRec.Recpt5 Then
              TempRec.CatFee2 = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
              If TempRec.CatFee2 < CodeRec.BaseAmt5 Then TempRec.CatFee2 = CodeRec.BaseAmt5
              If OverPen > 0 Or TempRec.CatFeeBal2 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal2, TempRec.CatFee2, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal2 = TempRec.CatFeeBal2 + TempRec.CatFee2
              End If
              GoTo C3
            End If
            If Revenue# <= CodeRec.Recpt6 Then
              TempRec.CatFee2 = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
              If TempRec.CatFee2 < CodeRec.BaseAmt6 Then TempRec.CatFee2 = CodeRec.BaseAmt6
              If OverPen > 0 Or TempRec.CatFeeBal2 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal2, TempRec.CatFee2, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal2 = TempRec.CatFeeBal2 + TempRec.CatFee2
              End If
              GoTo C3
            End If
          End If
        End If
      End If  'End Test for Code
    Next Snt&
  Else
    TempRec.CatFee2 = 0
  End If      'End Test for Cat 1


C3:
  CatCode$ = QPTrim$(CustRec.BILLCAT3)
  If Len(CatCode$) > 0 Then
    For Snt& = 1 To NumOfARCatRecs
      Get CodeHandle, Snt&, CodeRec
      If QPTrim$(CodeRec.CatCode) = CatCode$ Then
        TempRec.CatCodeRec3 = Snt&
        If CodeRec.CodeType = "F" Then
          If ProrateFlag = True Then
            TempRec.CatFee3 = OldRound(Prorate# * CodeRec.Fee)
          Else
            TempRec.CatFee3 = CodeRec.Fee
          End If
          If TempRec.CatFee3 < 0 Then TempRec.CatFee3 = 0
          If OverPen > 0 Or TempRec.CatFeeBal3 < 0 Then
            Call ApplyCredits2ThisFee(TempRec.CatFeeBal3, TempRec.CatFee3, OverPen)
            TempRec.CreditUsed = True
            TempRec.PenBal = OverPen
          Else
            TempRec.CatFeeBal3 = TempRec.CatFeeBal3 + TempRec.CatFee3
          End If
          GoTo c4
        End If
        If CodeRec.CodeType = "M" Then
          Mult = CustRec.REV3
          If ProrateFlag = True Then
            TempRec.CatFee3 = OldRound(Mult * CodeRec.Fee)
            TempRec.CatFee3 = OldRound(TempRec.CatFee3 * Prorate#)
          Else
            TempRec.CatFee3 = OldRound(Mult * CodeRec.Fee)
          End If
          If TempRec.CatFee3 < 0 Then TempRec.CatFee3 = 0
          If OverPen > 0 Or TempRec.CatFeeBal3 < 0 Then
            Call ApplyCredits2ThisFee(TempRec.CatFeeBal3, TempRec.CatFee3, OverPen)
            TempRec.CreditUsed = True
            TempRec.PenBal = OverPen
          Else
            TempRec.CatFeeBal3 = TempRec.CatFeeBal3 + TempRec.CatFee3
          End If
          GoTo c4
        End If
        If CodeRec.CodeType = "S" Then
          Revenue# = CustRec.REV3
          If ProrateFlag = True Then
            If Revenue# <= CodeRec.Recpt1 Then
              TempRec.CatFee3 = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
              If TempRec.CatFee3 < CodeRec.BaseAmt1 Then TempRec.CatFee3 = CodeRec.BaseAmt1
              TempRec.CatFee3 = OldRound(Prorate# * TempRec.CatFee3)
              If OverPen > 0 Or TempRec.CatFeeBal3 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal3, TempRec.CatFee3, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal3 = TempRec.CatFeeBal3 + TempRec.CatFee3
              End If
              GoTo c4
            End If
            If Revenue# <= CodeRec.Recpt2 Then
              TempRec.CatFee3 = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
              If TempRec.CatFee3 < CodeRec.BaseAmt2 Then TempRec.CatFee3 = CodeRec.BaseAmt2
              TempRec.CatFee3 = OldRound(Prorate# * TempRec.CatFee3)
              If OverPen > 0 Or TempRec.CatFeeBal3 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal3, TempRec.CatFee3, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal3 = TempRec.CatFeeBal3 + TempRec.CatFee3
              End If
              GoTo c4
            End If
            If Revenue# <= CodeRec.Recpt3 Then
              TempRec.CatFee3 = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
              If TempRec.CatFee3 < CodeRec.BaseAmt3 Then TempRec.CatFee3 = CodeRec.BaseAmt3
              TempRec.CatFee3 = OldRound(Prorate# * TempRec.CatFee3)
              If OverPen > 0 Or TempRec.CatFeeBal3 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal3, TempRec.CatFee3, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal3 = TempRec.CatFeeBal3 + TempRec.CatFee3
              End If
              GoTo c4
            End If
            If Revenue# <= CodeRec.Recpt4 Then
              TempRec.CatFee3 = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
              If TempRec.CatFee3 < CodeRec.BaseAmt4 Then TempRec.CatFee3 = CodeRec.BaseAmt4
              TempRec.CatFee3 = OldRound(Prorate# * TempRec.CatFee3)
              If OverPen > 0 Or TempRec.CatFeeBal3 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal3, TempRec.CatFee3, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal3 = TempRec.CatFeeBal3 + TempRec.CatFee3
              End If
              GoTo c4
            End If
            If Revenue# <= CodeRec.Recpt5 Then
              TempRec.CatFee3 = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
              If TempRec.CatFee3 < CodeRec.BaseAmt5 Then TempRec.CatFee3 = CodeRec.BaseAmt5
              TempRec.CatFee3 = OldRound(Prorate# * TempRec.CatFee3)
              If OverPen > 0 Or TempRec.CatFeeBal3 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal3, TempRec.CatFee3, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal3 = TempRec.CatFeeBal3 + TempRec.CatFee3
              End If
              GoTo c4
            End If
            If Revenue# <= CodeRec.Recpt6 Then
              TempRec.CatFee3 = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
              If TempRec.CatFee3 < CodeRec.BaseAmt6 Then TempRec.CatFee3 = CodeRec.BaseAmt6
              TempRec.CatFee3 = OldRound(Prorate# * TempRec.CatFee3)
              If OverPen > 0 Or TempRec.CatFeeBal3 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal3, TempRec.CatFee3, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal3 = TempRec.CatFeeBal3 + TempRec.CatFee3
              End If
              GoTo c4
            End If
          Else 'prorateflag = false
            If Revenue# <= CodeRec.Recpt1 Then
              TempRec.CatFee3 = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
              If TempRec.CatFee3 < CodeRec.BaseAmt1 Then TempRec.CatFee3 = CodeRec.BaseAmt1
              If OverPen > 0 Or TempRec.CatFeeBal3 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal3, TempRec.CatFee3, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal3 = TempRec.CatFeeBal3 + TempRec.CatFee3
              End If
              GoTo c4
            End If
            If Revenue# <= CodeRec.Recpt2 Then
              TempRec.CatFee3 = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
              If TempRec.CatFee3 < CodeRec.BaseAmt2 Then TempRec.CatFee3 = CodeRec.BaseAmt2
              If OverPen > 0 Or TempRec.CatFeeBal3 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal3, TempRec.CatFee3, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal3 = TempRec.CatFeeBal3 + TempRec.CatFee3
              End If
              GoTo c4
            End If
            If Revenue# <= CodeRec.Recpt3 Then
              TempRec.CatFee3 = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
              If TempRec.CatFee3 < CodeRec.BaseAmt3 Then TempRec.CatFee3 = CodeRec.BaseAmt3
              If OverPen > 0 Or TempRec.CatFeeBal3 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal3, TempRec.CatFee3, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal3 = TempRec.CatFeeBal3 + TempRec.CatFee3
              End If
              GoTo c4
            End If
            If Revenue# <= CodeRec.Recpt4 Then
              TempRec.CatFee3 = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
              If TempRec.CatFee3 < CodeRec.BaseAmt4 Then TempRec.CatFee3 = CodeRec.BaseAmt4
              If OverPen > 0 Or TempRec.CatFeeBal3 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal3, TempRec.CatFee3, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal3 = TempRec.CatFeeBal3 + TempRec.CatFee3
              End If
              GoTo c4
            End If
            If Revenue# <= CodeRec.Recpt5 Then
              TempRec.CatFee3 = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
              If TempRec.CatFee3 < CodeRec.BaseAmt5 Then TempRec.CatFee3 = CodeRec.BaseAmt5
              If OverPen > 0 Or TempRec.CatFeeBal3 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal3, TempRec.CatFee3, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal3 = TempRec.CatFeeBal3 + TempRec.CatFee3
              End If
              GoTo c4
            End If
            If Revenue# <= CodeRec.Recpt6 Then
              TempRec.CatFee3 = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
              If TempRec.CatFee3 < CodeRec.BaseAmt6 Then TempRec.CatFee3 = CodeRec.BaseAmt6
              If OverPen > 0 Or TempRec.CatFeeBal3 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal3, TempRec.CatFee3, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal3 = TempRec.CatFeeBal3 + TempRec.CatFee3
              End If
              GoTo c4
            End If
          End If
        End If
      End If  'End Test for Code
    Next Snt&
  Else
    TempRec.CatFee3 = 0
  End If      'End Test for Cat 3

c4:
  CatCode$ = QPTrim$(CustRec.BILLCAT4)
  If Len(CatCode$) > 0 Then
    For Snt& = 1 To NumOfARCatRecs
      Get CodeHandle, Snt&, CodeRec
      If QPTrim$(CodeRec.CatCode) = CatCode$ Then
        TempRec.CatCodeRec4 = Snt&
        If CodeRec.CodeType = "F" Then
          If ProrateFlag = True Then
            TempRec.CatFee4 = OldRound(Prorate# * CodeRec.Fee)
          Else
            TempRec.CatFee4 = CodeRec.Fee
          End If
          If TempRec.CatFee4 < 0 Then TempRec.CatFee4 = 0
          If OverPen > 0 Or TempRec.CatFeeBal4 < 0 Then
            Call ApplyCredits2ThisFee(TempRec.CatFeeBal4, TempRec.CatFee4, OverPen)
            TempRec.CreditUsed = True
            TempRec.PenBal = OverPen
          Else
            TempRec.CatFeeBal4 = TempRec.CatFeeBal4 + TempRec.CatFee4
          End If
          GoTo c5
        End If
        If CodeRec.CodeType = "M" Then
          Mult = CustRec.REV4
          If ProrateFlag = True Then
            TempRec.CatFee4 = OldRound(Mult * CodeRec.Fee)
            TempRec.CatFee4 = OldRound(TempRec.CatFee4 * Prorate#)
          Else
            TempRec.CatFee4 = OldRound(Mult * CodeRec.Fee)
          End If
          If TempRec.CatFee4 < 0 Then TempRec.CatFee4 = 0
          If OverPen > 0 Or TempRec.CatFeeBal4 < 0 Then
            Call ApplyCredits2ThisFee(TempRec.CatFeeBal4, TempRec.CatFee4, OverPen)
            TempRec.CreditUsed = True
            TempRec.PenBal = OverPen
          Else
            TempRec.CatFeeBal4 = TempRec.CatFeeBal4 + TempRec.CatFee4
          End If
          GoTo c5
        End If
        If CodeRec.CodeType = "S" Then
          Revenue# = CustRec.REV4
          If ProrateFlag = True Then
            If Revenue# <= CodeRec.Recpt1 Then
              TempRec.CatFee4 = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
              If TempRec.CatFee4 < CodeRec.BaseAmt1 Then TempRec.CatFee4 = CodeRec.BaseAmt1
              TempRec.CatFee4 = OldRound(Prorate# * TempRec.CatFee4)
              If OverPen > 0 Or TempRec.CatFeeBal4 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal4, TempRec.CatFee4, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal4 = TempRec.CatFeeBal4 + TempRec.CatFee4
              End If
              GoTo c5
            End If
            If Revenue# <= CodeRec.Recpt2 Then
              TempRec.CatFee4 = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
              If TempRec.CatFee4 < CodeRec.BaseAmt2 Then TempRec.CatFee4 = CodeRec.BaseAmt2
              TempRec.CatFee4 = OldRound(Prorate# * TempRec.CatFee4)
              If OverPen > 0 Or TempRec.CatFeeBal4 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal4, TempRec.CatFee4, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal4 = TempRec.CatFeeBal4 + TempRec.CatFee4
              End If
              GoTo c5
            End If
            If Revenue# <= CodeRec.Recpt3 Then
              TempRec.CatFee4 = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
              If TempRec.CatFee4 < CodeRec.BaseAmt3 Then TempRec.CatFee4 = CodeRec.BaseAmt3
              TempRec.CatFee4 = OldRound(Prorate# * TempRec.CatFee4)
              If OverPen > 0 Or TempRec.CatFeeBal4 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal4, TempRec.CatFee4, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal4 = TempRec.CatFeeBal4 + TempRec.CatFee4
              End If
              GoTo c5
            End If
            If Revenue# <= CodeRec.Recpt4 Then
              TempRec.CatFee4 = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
              If TempRec.CatFee4 < CodeRec.BaseAmt4 Then TempRec.CatFee4 = CodeRec.BaseAmt4
              TempRec.CatFee4 = OldRound(Prorate# * TempRec.CatFee4)
              If OverPen > 0 Or TempRec.CatFeeBal4 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal4, TempRec.CatFee4, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal4 = TempRec.CatFeeBal4 + TempRec.CatFee4
              End If
              GoTo c5
            End If
            If Revenue# <= CodeRec.Recpt5 Then
              TempRec.CatFee4 = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
              If TempRec.CatFee4 < CodeRec.BaseAmt5 Then TempRec.CatFee4 = CodeRec.BaseAmt5
              TempRec.CatFee4 = OldRound(Prorate# * TempRec.CatFee4)
              If OverPen > 0 Or TempRec.CatFeeBal4 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal4, TempRec.CatFee4, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal4 = TempRec.CatFeeBal4 + TempRec.CatFee4
              End If
              GoTo c5
            End If
            If Revenue# <= CodeRec.Recpt6 Then
              TempRec.CatFee4 = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
              If TempRec.CatFee4 < CodeRec.BaseAmt6 Then TempRec.CatFee4 = CodeRec.BaseAmt6
              TempRec.CatFee4 = OldRound(Prorate# * TempRec.CatFee4)
              If OverPen > 0 Or TempRec.CatFeeBal4 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal4, TempRec.CatFee4, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal4 = TempRec.CatFeeBal4 + TempRec.CatFee4
              End If
              GoTo c5
            End If
          Else 'ProrateFlag = False
            If Revenue# <= CodeRec.Recpt1 Then
              TempRec.CatFee4 = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
              If TempRec.CatFee4 < CodeRec.BaseAmt1 Then TempRec.CatFee4 = CodeRec.BaseAmt1
              If OverPen > 0 Or TempRec.CatFeeBal4 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal4, TempRec.CatFee4, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal4 = TempRec.CatFeeBal4 + TempRec.CatFee4
              End If
              GoTo c5
            End If
            If Revenue# <= CodeRec.Recpt2 Then
              TempRec.CatFee4 = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
              If TempRec.CatFee4 < CodeRec.BaseAmt2 Then TempRec.CatFee4 = CodeRec.BaseAmt2
              If OverPen > 0 Or TempRec.CatFeeBal4 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal4, TempRec.CatFee4, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal4 = TempRec.CatFeeBal4 + TempRec.CatFee4
              End If
              GoTo c5
            End If
            If Revenue# <= CodeRec.Recpt3 Then
              TempRec.CatFee4 = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
              If TempRec.CatFee4 < CodeRec.BaseAmt3 Then TempRec.CatFee4 = CodeRec.BaseAmt3
              If OverPen > 0 Or TempRec.CatFeeBal4 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal4, TempRec.CatFee4, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal4 = TempRec.CatFeeBal4 + TempRec.CatFee4
              End If
              GoTo c5
            End If
            If Revenue# <= CodeRec.Recpt4 Then
              TempRec.CatFee4 = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
              If TempRec.CatFee4 < CodeRec.BaseAmt4 Then TempRec.CatFee4 = CodeRec.BaseAmt4
              If OverPen > 0 Or TempRec.CatFeeBal4 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal4, TempRec.CatFee4, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal4 = TempRec.CatFeeBal4 + TempRec.CatFee4
              End If
              GoTo c5
            End If
            If Revenue# <= CodeRec.Recpt5 Then
              TempRec.CatFee4 = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
              If TempRec.CatFee4 < CodeRec.BaseAmt5 Then TempRec.CatFee4 = CodeRec.BaseAmt5
              If OverPen > 0 Or TempRec.CatFeeBal4 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal4, TempRec.CatFee4, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal4 = TempRec.CatFeeBal4 + TempRec.CatFee4
              End If
              GoTo c5
            End If
            If Revenue# <= CodeRec.Recpt6 Then
              TempRec.CatFee4 = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
              If TempRec.CatFee4 < CodeRec.BaseAmt6 Then TempRec.CatFee4 = CodeRec.BaseAmt6
              If OverPen > 0 Or TempRec.CatFeeBal4 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal4, TempRec.CatFee4, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal4 = TempRec.CatFeeBal4 + TempRec.CatFee4
              End If
              GoTo c5
            End If
          End If
        End If
      End If  'End Test for Code
    Next Snt&
  Else
    TempRec.CatFee4 = 0
  End If      'End Test for Cat 1

c5:
  CatCode$ = QPTrim$(CustRec.BILLCAT5)
  If Len(CatCode$) > 0 Then
    For Snt& = 1 To NumOfARCatRecs
      Get CodeHandle, Snt&, CodeRec
      If QPTrim$(CodeRec.CatCode) = CatCode$ Then
        TempRec.CatCodeRec5 = Snt&
        If CodeRec.CodeType = "F" Then
          If ProrateFlag = True Then
            TempRec.CatFee5 = OldRound(Prorate# * CodeRec.Fee)
          Else
            TempRec.CatFee5 = CodeRec.Fee
          End If
          If TempRec.CatFee5 < 0 Then TempRec.CatFee5 = 0
          If OverPen > 0 Or TempRec.CatFeeBal5 < 0 Then
            Call ApplyCredits2ThisFee(TempRec.CatFeeBal5, TempRec.CatFee5, OverPen)
            TempRec.CreditUsed = True
            TempRec.PenBal = OverPen
          Else
            TempRec.CatFeeBal5 = TempRec.CatFeeBal5 + TempRec.CatFee5
          End If
          GoTo FinishSaving
        End If
        If CodeRec.CodeType = "M" Then
          Mult = CustRec.REV5
          If ProrateFlag = True Then
            TempRec.CatFee5 = OldRound(Mult * CodeRec.Fee)
            TempRec.CatFee5 = OldRound(TempRec.CatFee5 * Prorate#)
          Else
            TempRec.CatFee5 = OldRound(Mult * CodeRec.Fee)
          End If
          If TempRec.CatFee5 < 0 Then TempRec.CatFee5 = 0
          If OverPen > 0 Or TempRec.CatFeeBal5 < 0 Then
            Call ApplyCredits2ThisFee(TempRec.CatFeeBal5, TempRec.CatFee5, OverPen)
            TempRec.CreditUsed = True
            TempRec.PenBal = OverPen
          Else
            TempRec.CatFeeBal5 = TempRec.CatFeeBal5 + TempRec.CatFee5
          End If
          GoTo FinishSaving
        End If
        If CodeRec.CodeType = "S" Then
          Revenue# = CustRec.REV5
          If ProrateFlag = True Then
            If Revenue# <= CodeRec.Recpt1 Then
              TempRec.CatFee5 = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
              If TempRec.CatFee5 < CodeRec.BaseAmt1 Then TempRec.CatFee5 = CodeRec.BaseAmt1
              TempRec.CatFee5 = OldRound(Prorate# * TempRec.CatFee5)
              If OverPen > 0 Or TempRec.CatFeeBal5 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal5, TempRec.CatFee5, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal5 = TempRec.CatFeeBal5 + TempRec.CatFee5
              End If
              GoTo FinishSaving
            End If
            If Revenue# <= CodeRec.Recpt2 Then
              TempRec.CatFee5 = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
              If TempRec.CatFee5 < CodeRec.BaseAmt2 Then TempRec.CatFee5 = CodeRec.BaseAmt2
              TempRec.CatFee5 = OldRound(Prorate# * TempRec.CatFee5)
              If OverPen > 0 Or TempRec.CatFeeBal5 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal5, TempRec.CatFee5, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal5 = TempRec.CatFeeBal5 + TempRec.CatFee5
              End If
              GoTo FinishSaving
            End If
            If Revenue# <= CodeRec.Recpt3 Then
              TempRec.CatFee5 = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
              If TempRec.CatFee5 < CodeRec.BaseAmt3 Then TempRec.CatFee5 = CodeRec.BaseAmt3
              TempRec.CatFee5 = OldRound(Prorate# * TempRec.CatFee5)
              If OverPen > 0 Or TempRec.CatFeeBal5 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal5, TempRec.CatFee5, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal5 = TempRec.CatFeeBal5 + TempRec.CatFee5
              End If
              GoTo FinishSaving
            End If
            If Revenue# <= CodeRec.Recpt4 Then
              TempRec.CatFee5 = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
              If TempRec.CatFee5 < CodeRec.BaseAmt4 Then TempRec.CatFee5 = CodeRec.BaseAmt4
              TempRec.CatFee5 = OldRound(Prorate# * TempRec.CatFee5)
              If OverPen > 0 Or TempRec.CatFeeBal5 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal5, TempRec.CatFee5, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal5 = TempRec.CatFeeBal5 + TempRec.CatFee5
              End If
              GoTo FinishSaving
            End If
            If Revenue# <= CodeRec.Recpt5 Then
              TempRec.CatFee5 = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
              If TempRec.CatFee5 < CodeRec.BaseAmt5 Then TempRec.CatFee5 = CodeRec.BaseAmt5
              TempRec.CatFee5 = OldRound(Prorate# * TempRec.CatFee5)
              If OverPen > 0 Or TempRec.CatFeeBal5 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal5, TempRec.CatFee5, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal5 = TempRec.CatFeeBal5 + TempRec.CatFee5
              End If
              GoTo FinishSaving
            End If
            If Revenue# <= CodeRec.Recpt6 Then
              TempRec.CatFee5 = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
              If TempRec.CatFee5 < CodeRec.BaseAmt6 Then TempRec.CatFee5 = CodeRec.BaseAmt6
              TempRec.CatFee5 = OldRound(Prorate# * TempRec.CatFee5)
              If OverPen > 0 Or TempRec.CatFeeBal5 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal5, TempRec.CatFee5, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal5 = TempRec.CatFeeBal5 + TempRec.CatFee5
              End If
              GoTo FinishSaving
            End If
          Else 'ProrateFlag = False
            If Revenue# <= CodeRec.Recpt1 Then
              TempRec.CatFee5 = OldRound(CodeRec.BaseAmt1 + (CodeRec.Percent1 / 100) * (Revenue# - CodeRec.Maximum1))
              If TempRec.CatFee5 < CodeRec.BaseAmt1 Then TempRec.CatFee5 = CodeRec.BaseAmt1
              If OverPen > 0 Or TempRec.CatFeeBal5 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal5, TempRec.CatFee5, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal5 = TempRec.CatFeeBal5 + TempRec.CatFee5
              End If
              GoTo FinishSaving
            End If
            If Revenue# <= CodeRec.Recpt2 Then
              TempRec.CatFee5 = OldRound(CodeRec.BaseAmt2 + (CodeRec.Percent2 / 100) * (Revenue# - CodeRec.Maximum2))
              If TempRec.CatFee5 < CodeRec.BaseAmt2 Then TempRec.CatFee5 = CodeRec.BaseAmt2
              If OverPen > 0 Or TempRec.CatFeeBal5 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal5, TempRec.CatFee5, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal5 = TempRec.CatFeeBal5 + TempRec.CatFee5
              End If
              GoTo FinishSaving
            End If
            If Revenue# <= CodeRec.Recpt3 Then
              TempRec.CatFee5 = OldRound(CodeRec.BaseAmt3 + (CodeRec.Percent3 / 100) * (Revenue# - CodeRec.Maximum3))
              If TempRec.CatFee5 < CodeRec.BaseAmt3 Then TempRec.CatFee5 = CodeRec.BaseAmt3
              If OverPen > 0 Or TempRec.CatFeeBal5 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal5, TempRec.CatFee5, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal5 = TempRec.CatFeeBal5 + TempRec.CatFee5
              End If
              GoTo FinishSaving
            End If
            If Revenue# <= CodeRec.Recpt4 Then
              TempRec.CatFee5 = OldRound(CodeRec.BaseAmt4 + (CodeRec.Percent4 / 100) * (Revenue# - CodeRec.Maximum4))
              If TempRec.CatFee5 < CodeRec.BaseAmt4 Then TempRec.CatFee5 = CodeRec.BaseAmt4
              If OverPen > 0 Or TempRec.CatFeeBal5 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal5, TempRec.CatFee5, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal5 = TempRec.CatFeeBal5 + TempRec.CatFee5
              End If
              GoTo FinishSaving
            End If
            If Revenue# <= CodeRec.Recpt5 Then
              TempRec.CatFee5 = OldRound(CodeRec.BaseAmt5 + (CodeRec.Percent5 / 100) * (Revenue# - CodeRec.Maximum5))
              If TempRec.CatFee5 < CodeRec.BaseAmt5 Then TempRec.CatFee5 = CodeRec.BaseAmt5
              If OverPen > 0 Or TempRec.CatFeeBal5 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal5, TempRec.CatFee5, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal5 = TempRec.CatFeeBal5 + TempRec.CatFee5
              End If
              GoTo FinishSaving
            End If
            If Revenue# <= CodeRec.Recpt6 Then
              TempRec.CatFee5 = OldRound(CodeRec.BaseAmt6 + (CodeRec.Percent6 / 100) * (Revenue# - CodeRec.Maximum6))
              If TempRec.CatFee5 < CodeRec.BaseAmt6 Then TempRec.CatFee5 = CodeRec.BaseAmt6
              If OverPen > 0 Or TempRec.CatFeeBal5 < 0 Then
                Call ApplyCredits2ThisFee(TempRec.CatFeeBal5, TempRec.CatFee5, OverPen)
                TempRec.CreditUsed = True
                TempRec.PenBal = OverPen
              Else
                TempRec.CatFeeBal5 = TempRec.CatFeeBal5 + TempRec.CatFee5
              End If
              GoTo FinishSaving
            End If
          End If
        End If
      End If  'End Test for Code
    Next Snt&
  Else
    TempRec.CatFee5 = 0
  End If      'End Test for Cat 1
FinishSaving:
  FeeAmt1# = TempRec.CatFee1
  FeeAmt2# = TempRec.CatFee2
  FeeAmt3# = TempRec.CatFee3
  FeeAmt4# = TempRec.CatFee4
  FeeAmt5# = TempRec.CatFee5
  ThisIssFee# = TownRec.IssFee
  TempRec.LicBal = TempRec.CatFeeBal1 + TempRec.CatFeeBal2 + TempRec.CatFeeBal3 + TempRec.CatFeeBal4 + TempRec.CatFeeBal5
  ThisAcctBal# = TempRec.PenBal + TempRec.LicBal + TempRec.IssFeeBal
  
  Close CodeHandle
  
  Exit Sub
  
CalcLicBal:
  
  'we only want to bring negative balances closer to zero here
  'since the overall license balance is negative then at least one
  'of the individual license balances has to be negative
  
  If TempRec.CatFeeBal1 >= 0 Then GoTo NextOne 'this isn't negative so move on
  'If the issue fee reduces the bal1 amount to zero but leaves a positive
  'amount in iss fee then carry the iss fee balance to the next category.
  'Otherwise bring the bal1 amount closer to zero and make iss fee zero and
  'then you're done
  If Abs(TempRec.CatFeeBal1) >= IssFee Then
    TempRec.CatFeeBal1 = TempRec.CatFeeBal1 + IssFee 'adding IssFee
    OverLic = OverLic - IssFee
    'brings the negative Bal1 closer to zero
    TempRec.CreditUsed = True
    IssFee = 0
    GoTo DoneHere
  Else
    'there is a balance in IssFee so reduce the
    'IssFee balance by the amount added to Bal1
    'and go to the next category
    IssFee = IssFee - Abs(TempRec.CatFeeBal1)
    TempRec.CreditUsed = True
    TempRec.CatFeeBal1 = 0 'was a negative so bring it up to zero
    OverLic = 0
  End If
  
NextOne:
  If TempRec.CatFeeBal2 >= 0 Then GoTo NextTwo
  If Abs(TempRec.CatFeeBal2) >= IssFee Then
    TempRec.CatFeeBal2 = TempRec.CatFeeBal2 + IssFee
    OverLic = OverLic - IssFee
    TempRec.CreditUsed = True
    IssFee = 0
    GoTo DoneHere
  Else
    IssFee = IssFee - Abs(TempRec.CatFeeBal2)
    TempRec.CreditUsed = True
    TempRec.CatFeeBal2 = 0
    OverLic = 0
  End If
  
NextTwo:
  If TempRec.CatFeeBal3 >= 0 Then GoTo NextThree
  If Abs(TempRec.CatFeeBal3) >= IssFee Then
    TempRec.CatFeeBal3 = TempRec.CatFeeBal3 + IssFee
    OverLic = OverLic - IssFee
    TempRec.CreditUsed = True
    IssFee = 0
    GoTo DoneHere
  Else
    IssFee = IssFee - Abs(TempRec.CatFeeBal3)
    TempRec.CreditUsed = True
    TempRec.CatFeeBal3 = 0
    OverLic = 0
  End If
 
NextThree:
  If TempRec.CatFeeBal4 >= 0 Then GoTo NextFour
  If Abs(TempRec.CatFeeBal4) >= IssFee Then
    TempRec.CatFeeBal4 = TempRec.CatFeeBal4 + IssFee
    OverLic = OverLic - IssFee
    TempRec.CreditUsed = True
    IssFee = 0
    GoTo DoneHere
  Else
    IssFee = IssFee - Abs(TempRec.CatFeeBal4)
    TempRec.CatFeeBal4 = 0
    TempRec.CreditUsed = True
    OverLic = 0
  End If
  
NextFour:
  If TempRec.CatFeeBal5 >= 0 Then GoTo DoneHere
  If Abs(TempRec.CatFeeBal5) >= IssFee Then
    TempRec.CatFeeBal5 = TempRec.CatFeeBal5 + IssFee
    OverLic = OverLic - IssFee
    TempRec.CreditUsed = True
    IssFee = 0
    GoTo DoneHere
  Else
    IssFee = IssFee - Abs(TempRec.CatFeeBal5)
    TempRec.CreditUsed = True
    TempRec.CatFeeBal5 = 0
    OverLic = 0
  End If

DoneHere:
  Return
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLPrintLicNoPost", "SetFee", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    ClearInUse PWcnt
    Terminate
  
  
End Sub

Private Sub ApplyCredits2ThisFee(ByRef ThisBal, ByVal ThisTFee As Double, ByRef OverPen As Double)
  
  On Error GoTo ERRORSTUFF
  
  'either ThisBal or OverPen has a negative value, possibly both
  If OverPen > 0 Then 'OverPen = a negative penalty balance
    If ThisTFee >= OverPen Then 'reduce fee by the credit in penalty and bring penalty balance up to 0
      ThisTFee = ThisTFee - OverPen
      OverPen = 0
'      CreditFlag = True
    ElseIf ThisTFee < OverPen Then 'reduce fee to 0 then bring penalty credit closer to 0
      OverPen = OverPen - ThisTFee
      ThisTFee = 0
    End If
  End If
  
  'bring any negative outstanding balance closer to zero
  'while reducing this license fee
  If ThisBal < 0 Then 'ThisBal is a negative license balance
    If ThisTFee >= Abs(ThisBal) Then
      ThisTFee = ThisTFee + ThisBal
      ThisBal = ThisTFee 'ThisBal now becomes whatever this category's fee is
    Else
      ThisBal = ThisBal + ThisTFee
      ThisTFee = 0
    End If
  End If
   
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLLicRegister", "ApplyCredits2ThisFee", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never happen.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    ClearInUse PWcnt
    Terminate
  
  
End Sub

Private Sub vaSpread_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
  If Col = 1 Then
    If vaSpread.Text <> "X" And vaSpread.Text <> "" Then
      vaSpread.Text = ""
      vaSpread.Text = "X"
    End If
  End If
End Sub

Private Sub vaSpread_KeyDown(KeyCode As Integer, Shift As Integer)
  KeyCode = KeyCode
End Sub

Private Sub vaSpread_KeyPress(KeyAscii As Integer)
  If KeyAscii = 88 Then Exit Sub
  If KeyAscii = 120 Then
    KeyAscii = 88
  Else
    KeyAscii = 0
  End If
End Sub
