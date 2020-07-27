VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVATaxCalcAdCol 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculate Advertising Charges"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxCalcAdCol.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbPrintOpt 
      Height          =   384
      Left            =   4272
      TabIndex        =   3
      Top             =   5136
      Width           =   3336
      _Version        =   196608
      _ExtentX        =   5884
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
      BackColor       =   16777215
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
      ColDesigner     =   "frmVATaxCalcAdCol.frx":08CA
   End
   Begin LpLib.fpCombo fpcmbPrintOrder 
      Height          =   384
      Left            =   4260
      TabIndex        =   4
      Top             =   6096
      Width           =   3372
      _Version        =   196608
      _ExtentX        =   5948
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
      AutoSearch      =   1
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
      AutoSearchFill  =   0   'False
      AutoSearchFillDelay=   500
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   1
      EditAlignV      =   0
      ColDesigner     =   "frmVATaxCalcAdCol.frx":0BC1
   End
   Begin EditLib.fpCurrency fpCurrChrg 
      Height          =   375
      Left            =   6180
      TabIndex        =   0
      Top             =   2561
      Width           =   1575
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "0"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   -2147483637
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
   Begin EditLib.fpDateTime fptxtCurrYear 
      Height          =   375
      Left            =   5970
      TabIndex        =   1
      Top             =   3281
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   -2147483637
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
      ThreeDTextHighlightColor=   -2147483637
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   0
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   0
      HideSelection   =   0   'False
      InvalidColor    =   -2147483643
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483643
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   "2018"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "yyyy"
      DateMax         =   "20350101"
      DateMin         =   "19800101"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "20010101"
      TimeDefault     =   "000000"
      TimeStyle       =   0
      BorderGrayAreaColor=   -2147483637
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
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
   Begin EditLib.fpDateTime fptxtPostDate 
      Height          =   375
      Left            =   5850
      TabIndex        =   2
      Top             =   4001
      Width           =   1740
      _Version        =   196608
      _ExtentX        =   3069
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      Text            =   "02/24/2005"
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
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
      Height          =   504
      Left            =   6408
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   7440
      Width           =   2064
      _Version        =   131072
      _ExtentX        =   3641
      _ExtentY        =   889
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
      ButtonDesigner  =   "frmVATaxCalcAdCol.frx":0EB8
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   504
      Left            =   3288
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   7440
      Width           =   2052
      _Version        =   131072
      _ExtentX        =   3619
      _ExtentY        =   889
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
      ButtonDesigner  =   "frmVATaxCalcAdCol.frx":1097
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   4455
      Left            =   2933
      Top             =   2321
      Width           =   5775
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Posting Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4050
      TabIndex        =   10
      Top             =   4121
      Width           =   1620
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Year:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   4650
      TabIndex        =   9
      Top             =   3386
      Width           =   1140
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Amount to Charge:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3900
      TabIndex        =   8
      Top             =   2681
      Width           =   2100
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Report Type:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   360
      Left            =   5115
      TabIndex        =   7
      Top             =   4841
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Printing Order:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   360
      Left            =   5100
      TabIndex        =   6
      Top             =   5801
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   660
      Index           =   1
      Left            =   1493
      Top             =   979
      Width           =   8655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Advertising Charges"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3113
      TabIndex        =   5
      Top             =   1144
      Width           =   5295
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   1493
      Top             =   874
      Width           =   8655
   End
End
Attribute VB_Name = "frmVATaxCalcAdCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  'Private Temp_Class As Resize_Class
  Dim UseOpt As String * 1
  Dim ThisOpt$
  Dim CurrTaxYear As Integer

Private Sub cmdExit_Click()
  frmVATaxAdvColMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdProcess_Click()
  Dim TaxCustRec As TaxCustType
  Dim PropertyRec As PropertyRecType
  Dim TaxTrans As TaxTransactionType
  Dim AdvTrans As InterestRecType
  Dim ATHandle As Integer
  Dim NumOfATRecs As Long
  Dim WhatYear As Integer
  Dim Year As Integer
  Dim TheDate$, CustAcct&
  Dim CustIdx As CustNameIdxType
  Dim CustIdxHandle As Integer
  Dim NumOfIdxRecs As Long
  Dim IdxCnt As Long, UsingIdx As Boolean
'  Dim UsingSrchIdx As Boolean
  Dim x As Long, cnt As Long
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TransRecord&
  Dim SrchIdx As SrchNameIdxType
  Dim NumOfSrchRecs As Long
  Dim SrchHandle As Integer
  Dim AmountCharged#
  Dim PRHandle As Integer
  Dim NumOfPRRecs As Long
  Dim TaxYear As Integer
  Dim PostDate$, CustCnt As Long
  Dim CurTransRecord&
  Dim TransRList As Long
  Dim Principle#, ThisPropDesc$
  Dim Paid#, Balance#, CurrOwnerPin&
  Dim BillNumber$, NME$
  Dim PropRec&, z1 As Long
  Dim TotalCharged#
  Dim ThisCustName$
  Dim ThisCustAcct&
  Dim NewOwnerFlag As Boolean
  Dim Limits As Integer
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim SSIdx As SocSecIdxType
  Dim SSIdxHandle As Integer
  Dim NumOfSSIdxRecs As Long
  
  On Error GoTo ERRORSTUFF
  
  AmountCharged# = CDbl(fpCurrChrg.Value)
  TaxYear = CInt(fptxtCurrYear.Text)
  
  If RevsAndGLsOK(Me, TaxYear, "R") = False Then
    Exit Sub
  End If
  
  If AmountCharged# = 0 Then
    Call TaxMsg(900, "The amount charged is zero. Processing aborted.")
    fpCurrChrg.SetFocus
    Exit Sub
  End If
  
  UsingIdx = False
'  UsingSrchIdx = False
  TheDate$ = Date$
  Year = CInt(fptxtCurrYear.Text)
  WhatYear = CurrTaxYear
  If Abs(CurrTaxYear - Year) > 5 Then
    If TaxMsgWOpts(800, "If " + Using("###0", Year) + " is the correct year then press F10 to continue. Otherwise, press ESC to review.", "F10 Continue", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtCurrYear.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
      MainLog ("WARNING: User warned that the current year they entered " + Using$("###0", Year) + " may be incorrect (more than 5 years from the current year of " + Using$("###0", CurrTaxYear) + ") and they continued anyway.")
    End If
  End If

  If fpcmbPrintOrder.Text = "2) Customer Name Order" Then
    UsingIdx = True
    OpenNameIdxFile CustIdxHandle, NumOfIdxRecs
    ReDim IdxRecs(1 To NumOfIdxRecs) As Long
    For x = 1 To NumOfIdxRecs
      Get CustIdxHandle, x, CustIdx
      IdxRecs(x) = CustIdx.CustRec
    Next x
    Close CustIdxHandle
  ElseIf fpcmbPrintOrder.Text = "3) Search Name Order" Then
    UsingIdx = True
    OpenSrchNameIdxFile SrchHandle, NumOfIdxRecs 'NumOfSrchRecs
    ReDim IdxRecs(1 To NumOfIdxRecs) As Long
    For x = 1 To NumOfIdxRecs 'NumOfSrchRecs
      Get SrchHandle, x, SrchIdx
      IdxRecs(x) = SrchIdx.CustRec
    Next x
    Close SrchHandle
  ElseIf fpcmbPrintOrder.Text = "4) Social Security Order" Then
    UsingIdx = True
    If Not Exist("TXSSIDX.DAT") Then
      If TaxMsgWOpts(800, "The social security number index has not been created. Press F10 if you would like to create this index or press ESC to abort interest calculation.", "F10 Make Index", "ESC Abort") = "abort" Then
        Unload frmVATaxMsgWOpts
        Close
        fpcmbPrintOrder.SetFocus
        Exit Sub
      Else
        Unload frmVATaxMsgWOpts
        Call CreateSSIdx
        Call Savemsg(900, "Index created successfully.")
      End If
    End If
    OpenSocSecIdxFile SSIdxHandle, NumOfIdxRecs 'NumOfSSIdxRecs
    ReDim IdxRecs(1 To NumOfIdxRecs) As Long
    For x = 1 To NumOfIdxRecs 'NumOfSrchRecs
      Get SSIdxHandle, x, SSIdx
      IdxRecs(x) = SSIdx.CustRec
    Next x
    Close SSIdxHandle
  ElseIf fpcmbPrintOrder.Text = "5) " + ThisOpt + " Order" Then
    UsingIdx = True
    OpenCustOptSearchFile OHandle, NumOfIdxRecs
    If NumOfIdxRecs = 0 Then
      Call TaxMsg(900, "There are no " + ThisOpt + "descriptions indexed.")
      Close OHandle
      Exit Sub
    End If
    ReDim IdxRecs(1 To NumOfIdxRecs) As Long
    For x = 1 To NumOfIdxRecs
      Get OHandle, x, OptRec
      IdxRecs(x) = OptRec.CustRec
    Next x
    Close OHandle
'    IdxFlag = True
'    OptFlag = True
    
  End If
  
  If InStr(fpcmbPrintOpt.Text, "Text") Then
    Call TaxMsg(900, "Pitch 10 is recommended for this report.")
  End If
  
  If Exist(TaxAdvFile) Then
    Kill TaxAdvFile             'kill any old work file
  End If
  
  OpenTaxCustFile TCHandle, NumOfTCRecs 'TaxCustRec
  OpenTaxTransFile TTHandle, NumOfTTRecs 'TaxTrans
  OpenAdvColRecFile ATHandle, NumOfATRecs 'AdvTrans
  OpenRealPropFile PRHandle, NumOfPRRecs
  CustCnt = 0
  frmVATaxShowPctComp.Label1 = "Calculating Advertising Charges"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  
  If UsingIdx = True Then
    NumOfTCRecs = NumOfIdxRecs
  End If
  
  For cnt& = 1 To NumOfTCRecs&
    If UsingIdx Then
      CustAcct& = IdxRecs(cnt&)
    Else
      CustAcct& = cnt&
    End If
    Get TCHandle, CustAcct&, TaxCustRec        'get cust record
    If TaxCustRec.Deleted <> 0 Then GoTo SkipIt
    TransRList = 0
    
    TransRecord& = TaxCustRec.LastTrans
    Do While TransRecord& <> 0
      Get TTHandle, TransRecord&, TaxTrans
      CurTransRecord& = TransRecord&
      If TaxTrans.BillType = "R" And TaxTrans.TranType = 1 And TaxTrans.TaxYear = TaxYear And Len(QPTrim$(TaxTrans.RealPin)) > 0 Then
        Principle# = OldRound#(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
        Principle# = OldRound#(Principle# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection)
        Principle# = OldRound#(Principle# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
        Paid# = OldRound#(TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd)
        Paid# = OldRound#(Paid# + TaxTrans.DiscAmt + TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd)
        Paid# = OldRound#(Paid# + TaxTrans.Revenue.LateListPd + TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.PPTRADisc)
          
        Balance# = OldRound#(Principle# - Paid#)
        TransRList = TransRList + 1
        If Balance# > 0 Then
          BillNumber$ = TaxTrans.Description
          BillNumber$ = ParseBillNum$(BillNumber$)
          ReDim PropDesc$(0 To 1)
          NewOwnerFlag = False
          CurrOwnerPin = 0
          PropRec& = FindPropRec(QPTrim$(TaxTrans.RealPin), TaxCustRec.PIN, CurrOwnerPin)
          If PropRec <= 0 Then GoTo SkipIt
          ThisCustName = ""
          ThisCustAcct = 0
          If CurrOwnerPin <> TaxCustRec.PIN Then
            NewOwnerFlag = True
            If CurrOwnerPin <> 0 Then
              Call FindNewOwner(ThisCustName$, CurrOwnerPin, ThisCustAcct&)
            Else
              ThisCustName = "UNKNOWN"
            End If
          End If
          z1 = 0
          Get PRHandle, PropRec&, PropertyRec
          ThisPropDesc$ = QPTrim$(PropertyRec.Map) + "\ " + QPTrim$(PropertyRec.BLOCK) + "\ " + QPTrim$(PropertyRec.LOTNUMB) + "\ " + QPTrim$(PropertyRec.PROPNOT1)
          NME$ = QPTrim$(TaxCustRec.CustName)
            
          TotalCharged# = OldRound(TotalCharged# + AmountCharged#)

          If AmountCharged# <> 0 Then   'Now Add Amt to Bill and Put Back
            AdvTrans.CustRec = CustAcct&
            AdvTrans.CustName = NME$
            AdvTrans.TaxYear = TaxYear
            AdvTrans.Amount = AmountCharged#
            AdvTrans.BillNumber = BillNumber$
            AdvTrans.BillRec = TransRecord&
            AdvTrans.InfoTxt = ThisPropDesc$
            If ThisCustAcct& > 0 Then
              AdvTrans.CustRec = ThisCustAcct& 'added 3/26/08
              AdvTrans.CustName = ThisCustName$ 'added 3/26/08
              AdvTrans.NewOwnerName = ThisCustName$
              AdvTrans.NewOwnerAcct = ThisCustAcct&
            Else
              AdvTrans.NewOwnerName = TaxCustRec.CustName
              AdvTrans.NewOwnerAcct = CustAcct&
            End If
            AdvTrans.DelFlag = 0
            AdvTrans.RealPin = TaxTrans.RealPin 'added 6/6/07
            Put ATHandle, , AdvTrans
          End If
          Balance# = 0
        End If
      End If
      TransRecord& = TaxTrans.LastTrans
    Loop
    
SkipIt:
    frmVATaxShowPctComp.ShowPctComp cnt, NumOfTCRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      Exit Sub
    End If
  Next cnt
  
  Close
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  
  If InStr(fpcmbPrintOpt.Text, "Graphical") Then
    Call PrintGraphics
  ElseIf InStr(fpcmbPrintOpt.Text, "Text") Then
    Call PrintText
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxCalcAdCol", "cmdProcess_Click", Erl)
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

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%E"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdProcess_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  'Set Temp_Class = New Resize_Class
  'Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpCalculate
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxCalcAdCol.")
      Call Terminate
      End
    End If
  End If

End Sub
'Private Sub Form_Resize()
'  If Me.WindowState <> vbMinimized Then
'    Me.Visible = False
'    'Temp_Class.ResizeControls Me
'    Me.Visible = True
'    Me.SetFocus
'    DoEvents
'  End If
'End Sub

Private Sub LoadMe()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  fptxtCurrYear.Text = CStr(TaxMasterRec.RTaxYear)
  CurrTaxYear = TaxMasterRec.RTaxYear
  fpcmbPrintOrder.Text = "1) Account Number Order"
  fpcmbPrintOrder.AddItem "1) Account Number Order"
  fpcmbPrintOrder.AddItem "2) Customer Name Order"
  fpcmbPrintOrder.AddItem "3) Search Name Order"
  fpcmbPrintOrder.AddItem "4) Social Security Order"
  ThisOpt = QPTrim$(TaxMasterRec.OptSrchCust)
  If ThisOpt <> "" Then
    fpcmbPrintOrder.AddItem "5) " + ThisOpt + " Order"
  End If
  
  fpcmbPrintOpt.Text = "Graphical"
  fpcmbPrintOpt.AddItem "Graphical"
  fpcmbPrintOpt.AddItem "Text"
  
  fptxtPostDate = Date
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
      fpcmbPrintOrder.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fptxtPostDate.SetFocus
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbPrintOrder_Change()
  If ThisOpt <> "" Then
    If InStr(fpcmbPrintOrder.Text, ThisOpt) Then
      UseOpt = "Y"
    Else
      UseOpt = "N"
    End If
  End If
End Sub

Private Sub fpcmbPrintOrder_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbPrintOrder.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPrintOrder.ListIndex = -1
  End If
  If fpcmbPrintOrder.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpCurrChrg.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcmbPrintOpt.SetFocus
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Function FindPropRec(RealPin$, CustPin As Long, CurrOwnerPin As Long) As Long
  Dim PropertyRec As PropertyRecType
  Dim PRHandle As Integer
  Dim NumOfPRRecs As Long
  Dim x As Long
  
  On Error GoTo ERRORSTUFF
  
  FindPropRec = 0
  OpenRealPropFile PRHandle, NumOfPRRecs 'PropertyRec
  For x = 1 To NumOfPRRecs
    Get PRHandle, x, PropertyRec
    If QPTrim$(PropertyRec.RealPin) = RealPin$ Then
      FindPropRec = x
      If PropertyRec.CustPin = 0 Then GoTo Deleted 'added 3/26/08
      If PropertyRec.CustPin <> CustPin Then
        CurrOwnerPin = PropertyRec.CustPin
      Else
        CurrOwnerPin = CustPin
      End If
      Exit For
    End If
Deleted:
  Next x
  
  Close PRHandle
  Exit Function
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxCalcAdCol", "FindPropRec", Erl)
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
End Function

Private Sub FindNewOwner(CustName$, CustPin&, CustAcct&)
  Dim TaxCustRec As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  
  On Error GoTo ERRORSTUFF
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCustRec
    If TaxCustRec.PIN = CustPin Then
      CustName = QPTrim$(TaxCustRec.CustName)
      CustAcct = x
      Exit For
    End If
  Next x
  Close TCHandle
  
  If x > NumOfTCRecs Then
    CustName$ = "UNKNOWN"
    CustAcct = 0
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxCalcAdCol", "FindNewOwner", Erl)
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
    
End Sub

Private Sub PrintGraphics()
  Dim AdvTrans As InterestRecType
  Dim ATHandle As Integer
  Dim NumOfATRecs As Long
  Dim dlm$
  Dim RptFile$
  Dim RptHandle As Integer
  Dim x As Long, y As Integer
  Dim TotAmt As Double
  Dim TCnt As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim Town$
  Dim PCnt As Long
  
  On Error GoTo ERRORSTUFF
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  Town$ = QPTrim$(TaxMasterRec.Name)
  dlm$ = "~"
  RptFile$ = "TAXRPTS\TAXADVCOL.RPT"
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  
  OpenAdvColRecFile ATHandle, NumOfATRecs
  For x = 1 To NumOfATRecs
    Get ATHandle, x, AdvTrans
    If AdvTrans.DelFlag = True Then GoTo SkipIt
    TotAmt = OldRound(TotAmt + AdvTrans.Amount)
    '                   0                 1                          2
    Print #RptHandle, Town$; dlm; fptxtCurrYear.Text; dlm; AdvTrans.CustRec; dlm;
    '                            3                            4                         5
    Print #RptHandle, QPTrim$(AdvTrans.CustName); dlm; AdvTrans.InfoTxt; dlm; AdvTrans.Amount; dlm;
    '                    6             7
    Print #RptHandle, TotAmt; dlm; NumOfATRecs
    PCnt = PCnt + 1
SkipIt:
  Next x
  
  Close
  
  If PCnt = 0 Then
    Call TaxMsg(900, "No advertising charges were warranted.")
    Exit Sub
  End If
  arVATaxAdvColRpt.Show
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxCalcAdCol", "PrintGraphics", Erl)
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
End Sub

Private Sub PrintText()
  Dim Page As Integer
  Dim LineCnt As Integer
  Dim MaxLines As Integer
  Dim RptHandle As Integer
  Dim RptFile$, FF$
  Dim AdvTrans As InterestRecType
  Dim ATHandle As Integer
  Dim NumOfATRecs As Long
  Dim TotAmt As Double
  Dim TCnt As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim Town$
  Dim TaxYear$
  Dim x As Long
  Dim ThisName As String * 35
  Dim ThisInfo As String * 30
  
  On Error GoTo ERRORSTUFF
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  MaxLines = 56
  Town$ = QPTrim$(TaxMasterRec.Name)
  TaxYear = CStr(fptxtCurrYear.Text)
  RptFile$ = "TAXRPTS\TAXADVCOL.RPT"
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  
  OpenAdvColRecFile ATHandle, NumOfATRecs
  GoSub PrintHeader
  
  For x = 1 To NumOfATRecs
    Get ATHandle, x, AdvTrans
    If AdvTrans.DelFlag = True Then GoTo SkipIt
    ThisName = QPTrim$(AdvTrans.CustName)
    ThisInfo = QPTrim$(AdvTrans.InfoTxt)
    TotAmt = OldRound(TotAmt + AdvTrans.Amount)
    Print #RptHandle, Using$("####0", AdvTrans.CustRec); Tab(10); ThisName; Tab(45); ThisInfo; Tab(76); Using$("$##,##0.00", AdvTrans.Amount)
    TCnt = TCnt + 1
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintHeader
    End If
SkipIt:
  Next x
  If LineCnt >= MaxLines - 4 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  Print #RptHandle, String$(85, "-")
  Print #RptHandle, "Transaction Count: "; Tab(26); Using$("####0", TCnt)
  Print #RptHandle, "Total Charges:     "; Tab(20); Using$("$###,##0.00", TotAmt)
  Print #RptHandle, FF$
  Close
  
  If TCnt = 0 Then
    Call TaxMsg(900, "No advertising charges were warranted.")
    Exit Sub
  End If
    
  ViewPrint RptFile, "Tax Advertising Charges Report", True
  
  Exit Sub
  
PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(30); "Tax Advertising Charges Report"
  Print #RptHandle, "Town: " + Town$; Tab(75); "Page # " + CStr(Page)
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle, "Tax Year: " + TaxYear
  Print #RptHandle, "Cust Num"; Tab(10); "Current Owner Name"; Tab(50); "Map\Block\Lot\Notes"; Tab(79); "Ad-Cost"
  Print #RptHandle, String(85, "-")
  LineCnt = 6
  
  Return
  
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxCalcAdCol", "PrintText", Erl)
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
End Sub
