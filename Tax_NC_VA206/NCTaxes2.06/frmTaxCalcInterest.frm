VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmTaxCalcInterest 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Interest Calculations"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmTaxCalcInterest.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbPrintOrder 
      Height          =   384
      Left            =   4080
      TabIndex        =   6
      Top             =   7068
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
      ColDesigner     =   "frmTaxCalcInterest.frx":08CA
   End
   Begin LpLib.fpCombo fpcmbPrintOpt 
      Height          =   384
      Left            =   4080
      TabIndex        =   5
      Tag             =   $"frmTaxCalcInterest.frx":0CA5
      Top             =   6228
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
      ColDesigner     =   "frmTaxCalcInterest.frx":0D74
   End
   Begin VB.TextBox fptxtOpt3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   4968
      Width           =   855
   End
   Begin VB.TextBox fptxtOpt2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4488
      Width           =   855
   End
   Begin VB.TextBox fptxtOpt1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4008
      Width           =   855
   End
   Begin VB.TextBox fptxtLateList 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   3528
      Width           =   855
   End
   Begin VB.TextBox fptxtAdvCol 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3048
      Width           =   855
   End
   Begin VB.TextBox fptxtInt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2568
      Width           =   855
   End
   Begin VB.TextBox fptxtPrinc 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2088
      Width           =   855
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
      Height          =   480
      Left            =   6720
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   7890
      Width           =   2070
      _Version        =   131072
      _ExtentX        =   3651
      _ExtentY        =   847
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
      ButtonDesigner  =   "frmTaxCalcInterest.frx":114F
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   480
      Left            =   2850
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   7890
      Width           =   2055
      _Version        =   131072
      _ExtentX        =   3625
      _ExtentY        =   847
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
      ButtonDesigner  =   "frmTaxCalcInterest.frx":132E
   End
   Begin EditLib.fpDoubleSingle fptxtCurrYrIntRate 
      Height          =   372
      Left            =   9480
      TabIndex        =   1
      ToolTipText     =   "If you wish to use a 5% penalty then enter 5 (not .5) in this field."
      Top             =   2148
      Width           =   852
      _Version        =   196608
      _ExtentX        =   1508
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
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   "."
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "100"
      MinValue        =   "0"
      NegFormat       =   1
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
   Begin EditLib.fpDoubleSingle fptxtPastYearIntRate 
      Height          =   372
      Left            =   9480
      TabIndex        =   2
      ToolTipText     =   "If you wish to use a 5% penalty then enter 5 (not .5) in this field."
      Top             =   2748
      Width           =   852
      _Version        =   196608
      _ExtentX        =   1508
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
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   "."
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "100"
      MinValue        =   "0"
      NegFormat       =   1
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
      Height          =   372
      Left            =   9000
      TabIndex        =   0
      Top             =   1548
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
      _ExtentY        =   661
      Enabled         =   0   'False
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
   Begin EditLib.fpDateTime fptxtFrom 
      Height          =   372
      Left            =   8400
      TabIndex        =   3
      Top             =   4440
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
      _ExtentY        =   656
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      AlignTextH      =   1
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
      Appearance      =   0
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
   Begin EditLib.fpDateTime fptxtTo 
      Height          =   372
      Left            =   8400
      TabIndex        =   4
      Top             =   5040
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
      _ExtentY        =   656
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      AlignTextH      =   1
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
      Appearance      =   0
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
   Begin VB.Shape Shape6 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1812
      Left            =   3720
      Top             =   5880
      Width           =   4092
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select A Date Range For Interest Calculations"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   612
      Left            =   6960
      TabIndex        =   32
      Top             =   3720
      Width           =   3420
   End
   Begin VB.Label Label33 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   276
      Left            =   7800
      TabIndex        =   31
      Top             =   5160
      Width           =   420
   End
   Begin VB.Label Label32 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   276
      Left            =   7560
      TabIndex        =   30
      Top             =   4560
      Width           =   660
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2052
      Left            =   6600
      Top             =   3588
      Width           =   4092
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Revenue Items Tagged For Interest:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   840
      TabIndex        =   15
      Top             =   1428
      Width           =   4572
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Opt3:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   276
      Left            =   1320
      TabIndex        =   22
      Top             =   5100
      Width           =   2580
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Opt2:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   276
      Left            =   1320
      TabIndex        =   21
      Top             =   4620
      Width           =   2580
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Opt1:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   276
      Left            =   1320
      TabIndex        =   20
      Top             =   4140
      Width           =   2580
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Late Listing:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   276
      Left            =   1320
      TabIndex        =   19
      Top             =   3660
      Width           =   2580
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Advertising/Col:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   276
      Left            =   1320
      TabIndex        =   18
      Top             =   3180
      Width           =   2580
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Interest:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   276
      Left            =   1320
      TabIndex        =   17
      Top             =   2700
      Width           =   2580
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Principle:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   276
      Left            =   1320
      TabIndex        =   16
      Top             =   2220
      Width           =   2580
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   4212
      Left            =   840
      Top             =   1428
      Width           =   4932
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1932
      Left            =   6600
      Top             =   1428
      Width           =   4092
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
      Left            =   4920
      TabIndex        =   14
      Top             =   6768
      Width           =   1812
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
      Left            =   4920
      TabIndex        =   13
      Top             =   5928
      Width           =   1812
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Past Year Interest Rate:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   276
      Left            =   6960
      TabIndex        =   12
      Top             =   2868
      Width           =   2340
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Current Year Interest Rate:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   276
      Left            =   6720
      TabIndex        =   11
      Top             =   2256
      Width           =   2580
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Current Tax Year:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   276
      Left            =   6960
      TabIndex        =   10
      Top             =   1656
      Width           =   1860
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Interest Calculations"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   3120
      TabIndex        =   7
      Top             =   624
      Width           =   5292
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   660
      Index           =   1
      Left            =   1500
      Top             =   456
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   1500
      Top             =   348
      Width           =   8652
   End
End
Attribute VB_Name = "frmTaxCalcInterest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim CurrTaxYear As Integer
  Dim Over As clsTextBoxOverRider
  Dim PrincInt As Boolean
  Dim IntInt As Boolean
  Dim AdvColInt As Boolean
  Dim LateListInt As Boolean
  Dim Opt1Int As Boolean
  Dim Opt2Int As Boolean
  Dim Opt3Int As Boolean
  Dim Years() As Integer
  Dim AtLeastOne As Boolean
  Dim YrCnt As Integer
  Dim ThisOpt$
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  frmTaxInterestMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdProcess_Click()
  Dim TaxCust As TaxCustType
  Dim TaxTrans As TaxTransactionType
  Dim IntTrans As InterestRecType
  Dim Year As Integer
  Dim TheDate$, CustAcct&
  Dim CustIdx As CustNameIdxType
  Dim CustIdxHandle As Integer
  Dim NumOfIdxRecs As Long
  Dim IdxCnt As Long, UsingNameIdx As Boolean
  Dim UsingSrchIdx As Boolean
  Dim x As Long, Cnt As Long
  Dim IRHandle As Integer
  Dim NumOfIRRecs As Long
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TransRecord&, IntAmount#, TIntAmount#
  Dim CurYearFlag As Boolean
  Dim PastYearFlag As Boolean
  Dim ThisBalance#, BillNumber$
  Dim TotBalance#
  Dim WhatYear As Integer
  Dim CurRate#, PastRate#, NME$
  Dim IntRecord&
  Dim SrchIdx As SrchNameIdxType
  Dim NumOfSrchRecs As Long
  Dim SrchHandle As Integer
  Dim SSIdx As SocSecIdxType
  Dim SSIdxHandle As Integer
  Dim NumOfSSIdxRecs As Long
  Dim UsingIdx As Boolean
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim HoldBal#
  Dim ThisDisc#
  Dim FromDate As Integer
  Dim ToDate As Integer
  Dim BYear As Integer
  Dim EYear As Integer
  
  'on error goto ERRORSTUFF
  
  BYear = CInt(fptxtFrom.Text)
  EYear = CInt(fptxtTo.Text)
  
  If AtLeastOne = False Then
    Call TaxMsg(800, "Currently no revenues are tagged for interest charges. Please refer to the 'System Setup' screen if you wish to earmark any revenues for interest charges.")
    Exit Sub
  End If
  
  If RevsAndGLsOK(Me, CurrTaxYear) = False Then
    Exit Sub
  End If
  
  FromDate = CInt(fptxtFrom.Text)
  ToDate = CInt(fptxtTo.Text)
  If ToDate < FromDate Then
    Call TaxMsg(900, "Error: The beginning date comes after the ending date.")
    fptxtFrom.SetFocus
    Exit Sub
  End If
  
  TheDate$ = Date$
  Year = CInt(fptxtCurrYear.Text)
  WhatYear = CurrTaxYear
  If Abs(CurrTaxYear - Year) > 5 Then
    If TaxMsgWOpts(800, "If " + Using("###0", Year) + " is the correct year then press F10 to continue. Otherwise, press ESC to review.", "F10 Continue", "ESC Review") = "abort" Then
      Unload frmTaxMsgWOpts
      fptxtCurrYear.SetFocus
      Exit Sub
    Else
      Unload frmTaxMsgWOpts
      MainLog ("WARNING: User warned that the current year they entered " + Using$("###0", Year) + " may be incorrect (more than 5 years from the current year of " + Using$("###0", CurrTaxYear) + ") and they continued anyway.")
    End If
  End If
  
  If CDbl(fptxtCurrYrIntRate.Text) = 0 Then
    If TaxMsgWOpts(900, "The current year interest rate is zero. If this is correct press F10 to continue. Otherwise, press ESC to review.", "F10 Continue", "ESC Review") = "abort" Then
      Unload frmTaxMsgWOpts
      fptxtCurrYrIntRate.SetFocus
      Exit Sub
    Else
      Unload frmTaxMsgWOpts
    End If
  End If
  
  If CDbl(fptxtPastYearIntRate.Text) = 0 Then
    If TaxMsgWOpts(900, "The past year interest rate is zero. If this is correct press F10 to continue. Otherwise, press ESC to review.", "F10 Continue", "ESC Review") = "abort" Then
      Unload frmTaxMsgWOpts
      fptxtPastYearIntRate.SetFocus
      Exit Sub
    Else
      Unload frmTaxMsgWOpts
    End If
  End If
  
  CurRate# = CDbl(fptxtCurrYrIntRate.Text)
  PastRate# = CDbl(fptxtPastYearIntRate.Text)
  UsingIdx = False
  
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
    For x = 1 To NumOfIdxRecs
      Get SrchHandle, x, SrchIdx
      IdxRecs(x) = SrchIdx.CustRec
    Next x
    Close SrchHandle
  ElseIf fpcmbPrintOrder.Text = "4) Social Security Order" Then
    UsingIdx = True
    If Not Exist("TXSSIDX.DAT") Then
      If TaxMsgWOpts(800, "The social security number index has not been created. Press F10 if you would like to create this index or press ESC to abort interest calculation.", "F10 Make Index", "ESC Abort") = "abort" Then
        Unload frmTaxMsgWOpts
        Close
        fpcmbPrintOrder.SetFocus
        Exit Sub
      Else
        Unload frmTaxMsgWOpts
        Call CreateSSIdx
        Call Savemsg(900, "Index created successfully.")
      End If
    End If
    OpenSocSecIdxFile SSIdxHandle, NumOfIdxRecs
    ReDim IdxRecs(1 To NumOfIdxRecs) As Long
    For x = 1 To NumOfIdxRecs
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
  End If
      
  If Exist(TaxIntFile) Then
    KillFile TaxIntFile              'kill any old work file
  End If
  
  If InStr(fpcmbPrintOpt.Text, "Text") Then
    Call TaxMsg(900, "Pitch 10 is recommended for this report.")
  End If
  
  OpenInterestRecFile IRHandle, NumOfIRRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  
  frmTaxShowPctComp.Label1 = "Calculating Interest"
  frmTaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  
  If UsingIdx = True Then NumOfTCRecs& = NumOfIdxRecs
  For Cnt& = 1 To NumOfTCRecs&
    If UsingIdx = True Then
      CustAcct& = IdxRecs(Cnt&)
    Else
      CustAcct& = Cnt&
    End If
    Get TCHandle, CustAcct&, TaxCust         'get cust record

    If TaxCust.Deleted <> 0 Then GoTo CalcIntSkip

    If TaxCust.Interest = "N" Then
      GoTo CalcIntSkip
    End If

    TransRecord& = TaxCust.LastTrans
    Do While TransRecord& > 0
      Get TTHandle, TransRecord&, TaxTrans
      CurYearFlag = False
      PastYearFlag = False
      TIntAmount# = 0
      If TaxTrans.TranType = 1 Then
        If TaxTrans.TaxYear < BYear Or TaxTrans.TaxYear > EYear Then GoTo NotInDate
        ThisDisc# = TaxTrans.DiscAmt
        IntAmount# = 0
        If PrincInt = True Then
          ThisBalance# = TaxTrans.Revenue.Principle1
          ThisBalance# = ThisBalance# - (TaxTrans.Revenue.Principle1Pd)
          ThisBalance# = OldRound#(ThisBalance#)
          HoldBal# = ThisBalance#
          ThisBalance# = OldRound(ThisBalance# - ThisDisc#)
          ThisDisc# = OldRound(ThisDisc# - HoldBal#)
          If ThisDisc# < 0 Then ThisDisc# = 0
          If ThisBalance# > 0 Then
            If TaxTrans.TaxYear = WhatYear Then CurYearFlag = True
            If TaxTrans.TaxYear <> WhatYear Then PastYearFlag = True
            If CurYearFlag Then
              IntAmount# = OldRound#(ThisBalance# * (CurRate# / 100))
            End If
            If PastYearFlag Then
              IntAmount# = OldRound#(ThisBalance# * (PastRate# / 100))
            End If
            TIntAmount# = OldRound(TIntAmount# + IntAmount#)
            TotBalance = OldRound(TotBalance# + ThisBalance#)
          End If
        End If
        ThisBalance = 0
        If IntInt = True Then
          ThisBalance# = TaxTrans.Revenue.Interest
          ThisBalance# = ThisBalance# - (TaxTrans.Revenue.InterestPd)
          ThisBalance# = OldRound#(ThisBalance#)
          HoldBal# = ThisBalance#
          ThisBalance# = OldRound(ThisBalance# - ThisDisc#)
          ThisDisc# = OldRound(ThisDisc# - HoldBal#)
          If ThisDisc# < 0 Then ThisDisc# = 0
          If ThisBalance# > 0 Then
            If TaxTrans.TaxYear = WhatYear Then CurYearFlag = True
            If TaxTrans.TaxYear <> WhatYear Then PastYearFlag = True
            If CurYearFlag Then
              IntAmount# = OldRound#(ThisBalance# * (CurRate# / 100))
            End If
            If PastYearFlag Then
              IntAmount# = OldRound#(ThisBalance# * (PastRate# / 100))
            End If
'            IntAmount# = OldRound(IntAmount# + ThisBalance#)
            TIntAmount# = OldRound(TIntAmount# + IntAmount#)
            TotBalance = OldRound(TotBalance# + ThisBalance#)
         End If
        End If
        ThisBalance = 0
        If AdvColInt = True Then
          ThisBalance# = TaxTrans.Revenue.Collection
          ThisBalance# = ThisBalance# - (TaxTrans.Revenue.CollectionPd)
          ThisBalance# = OldRound#(ThisBalance#)
          HoldBal# = ThisBalance#
          ThisBalance# = OldRound(ThisBalance# - ThisDisc#)
          ThisDisc# = OldRound(ThisDisc# - HoldBal#)
          If ThisDisc# < 0 Then ThisDisc# = 0
          If ThisBalance# > 0 Then
            If TaxTrans.TaxYear = WhatYear Then CurYearFlag = True
            If TaxTrans.TaxYear <> WhatYear Then PastYearFlag = True
            If CurYearFlag Then
              IntAmount# = OldRound#(ThisBalance# * (CurRate# / 100))
            End If
            If PastYearFlag Then
              IntAmount# = OldRound#(ThisBalance# * (PastRate# / 100))
            End If
            TIntAmount# = OldRound(TIntAmount# + IntAmount#)
            TotBalance = OldRound(TotBalance# + ThisBalance#)
'            IntAmount# = OldRound(IntAmount# + ThisBalance#)
          End If
        End If
        ThisBalance = 0
        If LateListInt = True Then
          ThisBalance# = TaxTrans.Revenue.LateList
          ThisBalance# = ThisBalance# - (TaxTrans.Revenue.LateListPd)
          ThisBalance# = OldRound#(ThisBalance#)
          HoldBal# = ThisBalance#
          ThisBalance# = OldRound(ThisBalance# - ThisDisc#)
          ThisDisc# = OldRound(ThisDisc# - HoldBal#)
          If ThisDisc# < 0 Then ThisDisc# = 0
          If ThisBalance# > 0 Then
            If TaxTrans.TaxYear = WhatYear Then CurYearFlag = True
            If TaxTrans.TaxYear <> WhatYear Then PastYearFlag = True
            If CurYearFlag Then
              IntAmount# = OldRound#(ThisBalance# * (CurRate# / 100))
            End If
            If PastYearFlag Then
              IntAmount# = OldRound#(ThisBalance# * (PastRate# / 100))
            End If
            TIntAmount# = OldRound(TIntAmount# + IntAmount#)
            TotBalance = OldRound(TotBalance# + ThisBalance#)
'            IntAmount# = OldRound(IntAmount# + ThisBalance#)
          End If
        End If
        ThisBalance = 0
        If Opt1Int = True Then
          ThisBalance# = TaxTrans.Revenue.RevOpt1
          ThisBalance# = ThisBalance# - (TaxTrans.Revenue.RevOpt1Pd)
          ThisBalance# = OldRound#(ThisBalance#)
          HoldBal# = ThisBalance#
          ThisBalance# = OldRound(ThisBalance# - ThisDisc#)
          ThisDisc# = OldRound(ThisDisc# - HoldBal#)
          If ThisDisc# < 0 Then ThisDisc# = 0
          If ThisBalance# > 0 Then
            If TaxTrans.TaxYear = WhatYear Then CurYearFlag = True
            If TaxTrans.TaxYear <> WhatYear Then PastYearFlag = True
            If CurYearFlag Then
              IntAmount# = OldRound#(ThisBalance# * (CurRate# / 100))
            End If
            If PastYearFlag Then
              IntAmount# = OldRound#(ThisBalance# * (PastRate# / 100))
            End If
            TIntAmount# = OldRound(TIntAmount# + IntAmount#)
            TotBalance = OldRound(TotBalance# + ThisBalance#)
'            IntAmount# = OldRound(IntAmount# + ThisBalance#)
          End If
        End If
        ThisBalance = 0
        If Opt2Int = True Then
          ThisBalance# = TaxTrans.Revenue.RevOpt2
          ThisBalance# = ThisBalance# - (TaxTrans.Revenue.RevOpt2Pd)
          ThisBalance# = OldRound#(ThisBalance#)
          HoldBal# = ThisBalance#
          ThisBalance# = OldRound(ThisBalance# - ThisDisc#)
          ThisDisc# = OldRound(ThisDisc# - HoldBal#)
          If ThisDisc# < 0 Then ThisDisc# = 0
          If ThisBalance# > 0 Then
            If TaxTrans.TaxYear = WhatYear Then CurYearFlag = True
            If TaxTrans.TaxYear <> WhatYear Then PastYearFlag = True
            If CurYearFlag Then
              IntAmount# = OldRound#(ThisBalance# * (CurRate# / 100))
            End If
            If PastYearFlag Then
              IntAmount# = OldRound#(ThisBalance# * (PastRate# / 100))
            End If
            TIntAmount# = OldRound(TIntAmount# + IntAmount#)
            TotBalance = OldRound(TotBalance# + ThisBalance#)
'            IntAmount# = OldRound(IntAmount# + ThisBalance#)
          End If
        End If
        ThisBalance = 0
        If Opt3Int = True Then
          ThisBalance# = TaxTrans.Revenue.RevOpt3
          ThisBalance# = ThisBalance# - (TaxTrans.Revenue.RevOpt3Pd)
          ThisBalance# = OldRound#(ThisBalance#)
          HoldBal# = ThisBalance#
          ThisBalance# = OldRound(ThisBalance# - ThisDisc#)
          ThisDisc# = OldRound(ThisDisc# - HoldBal#)
          If ThisDisc# < 0 Then ThisDisc# = 0
          If ThisBalance# > 0 Then
            If TaxTrans.TaxYear = WhatYear Then CurYearFlag = True
            If TaxTrans.TaxYear <> WhatYear Then PastYearFlag = True
            If CurYearFlag Then
              IntAmount# = OldRound#(ThisBalance# * (CurRate# / 100))
            End If
            If PastYearFlag Then
              IntAmount# = OldRound#(ThisBalance# * (PastRate# / 100))
            End If
'            IntAmount# = OldRound(IntAmount# + ThisBalance#)
            TIntAmount# = OldRound(TIntAmount# + IntAmount#)
            TotBalance = OldRound(TotBalance# + ThisBalance#)
          End If
        End If
        ThisBalance = 0

'        If OldRound#(IntAmount#) > 0 Then
        If OldRound#(TIntAmount#) > 0 Then
          BillNumber$ = TaxTrans.Description
          BillNumber$ = ParseBillNum$(BillNumber$)
          NME$ = RTrim$(TaxCust.CustName)
          NME$ = LTrim$(NME$)
          IntTrans.CustRec = CustAcct&
          IntTrans.CustPin = TaxCust.PIN
          IntTrans.CustName = NME$
          IntTrans.TaxYear = TaxTrans.TaxYear
'          IntTrans.Amount = OldRound#(IntAmount#)
          IntTrans.Amount = OldRound#(TIntAmount#)
          IntTrans.BillNumber = BillNumber$
          IntTrans.BillRec = TransRecord&
          IntTrans.RealPin = TaxTrans.RealPin
          IntTrans.PersPin = TaxTrans.PersPin
          IntTrans.CurYear = WhatYear
          IntRecord& = IntRecord& + 1
          Put IRHandle, IntRecord&, IntTrans
        End If
      End If
NotInDate:
    TransRecord& = TaxTrans.LastTrans
    Loop
CalcIntSkip:
    frmTaxShowPctComp.ShowPctComp Cnt, NumOfTCRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next Cnt
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True

  'CalcInt Calc END   *******************************
  Close
  
  If InStr(fpcmbPrintOpt.Text, "Graphical") Then
    Call PrintGraphics
  ElseIf InStr(fpcmbPrintOpt.Text, "Text") Then
    Call PrintText
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxCalcInterest", "cmdProcess_Click", Erl)
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
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpCalculateI
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxCalcInterest.")
      Call Terminate
      End
    End If
  End If

End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub LoadMe()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim x As Integer
  Dim DLen As Integer
  
  AtLeastOne = False
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
'  DLen = Len(Date)
'  fptxtTo.Text = CInt(Mid(fptxtTo.Text, DLen - 3, 4))
  fptxtTo.Text = Year(Date)
  x = CInt(fptxtTo.Text)
  x = x - 1
  
  fptxtCurrYear.Text = CStr(TaxMasterRec.TaxYear)
  CurrTaxYear = TaxMasterRec.TaxYear
  fptxtCurrYrIntRate = TaxMasterRec.CurrYrInt
  fptxtPastYearIntRate = TaxMasterRec.PastYrInt
  
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
  PrincInt = False
  IntInt = False
  AdvColInt = False
  LateListInt = False
  Opt1Int = False
  Opt2Int = False
  Opt3Int = False
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  If TaxMasterRec.IntPrncTaxYN = "Y" Then
    PrincInt = True
    fptxtPrinc.Text = "YES"
    AtLeastOne = True
  Else
    fptxtPrinc.Text = "NO"
  End If
  
  If TaxMasterRec.IntIntYN = "Y" Then
    IntInt = True
    fptxtInt.Text = "YES"
    AtLeastOne = True
  Else
    fptxtInt.Text = "NO"
  End If
  
  If TaxMasterRec.IntAdvYN = "Y" Then
    AdvColInt = True
    fptxtAdvCol.Text = "YES"
    AtLeastOne = True
  Else
    fptxtAdvCol.Text = "NO"
  End If
  
  If TaxMasterRec.IntLateLstYN = "Y" Then
    LateListInt = True
    fptxtLateList.Text = "YES"
    AtLeastOne = True
  Else
    fptxtLateList.Text = "NO"
  End If
  
  If QPTrim$(TaxMasterRec.OptRev1) = "" Then
    Label8.Visible = False
    fptxtOpt1.Visible = False
  Else
    Label8.Caption = QPTrim$(TaxMasterRec.OptRev1)
    If TaxMasterRec.IntOpt1YN = "Y" Then
      fptxtOpt1.Text = "YES"
      Opt1Int = True
      AtLeastOne = True
    Else
      fptxtOpt1.Text = "NO"
    End If
  End If
  
  If QPTrim$(TaxMasterRec.OptRev2) = "" Then
    Label12.Visible = False
    fpTxtOpt2.Visible = False
  Else
    Label12.Caption = QPTrim$(TaxMasterRec.OptRev2)
    If TaxMasterRec.IntOpt2YN = "Y" Then
      fpTxtOpt2.Text = "YES"
      Opt2Int = True
      AtLeastOne = True
    Else
      fpTxtOpt2.Text = "NO"
    End If
  End If
  
  If QPTrim$(TaxMasterRec.OptRev3) = "" Then
    Label13.Visible = False
    fpTxtOpt3.Visible = False
  Else
    Label13.Caption = QPTrim$(TaxMasterRec.OptRev3)
    If TaxMasterRec.IntOpt3YN = "Y" Then
      fpTxtOpt3.Text = "YES"
      Opt3Int = True
      AtLeastOne = True
    Else
      fpTxtOpt3.Text = "NO"
    End If
  End If
  
End Sub

Private Sub PrintGraphics()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim IntRec As InterestRecType
  Dim NumOfIRRecs As Long
  Dim IRHandle As Integer
  Dim x As Long, y As Integer
  Dim Town$
  Dim dlm$
  Dim RptHandle As Integer
  Dim RptFile$
  Dim SubRptHandle As Integer
  Dim SubRptFile$
  Dim TCnt As Long
  Dim TotInt As Double
  Dim TotCurrInt As Double
  Dim TotPastInt As Double
  
  'on error goto ERRORSTUFF
  
  dlm$ = "~"
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  Town$ = QPTrim$(TaxMasterRec.Name)
  
  RptFile$ = "TAXRPTS\TAXINT.RPT"     'Report File Name
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  
  OpenInterestRecFile IRHandle, NumOfIRRecs
  If NumOfIRRecs = 0 Then
    Call TaxMsg(900, "No interest calculations could be made.")
    Close
    Exit Sub
  End If
  Call GetYears
  ReDim YearAmts(1 To YrCnt) As Double
  
  For x = 1 To NumOfIRRecs
    Get IRHandle, x, IntRec
    '                   0               1                    2
    Print #RptHandle, Town; dlm; IntRec.CurYear; dlm; IntRec.CustRec; dlm;
    '                            3                           4
    Print #RptHandle, QPTrim$(IntRec.CustName); dlm; IntRec.BillNumber; dlm;
    '                        5                  6
    Print #RptHandle, IntRec.TaxYear; dlm; IntRec.Amount; dlm;
    TotInt = OldRound(TotInt + IntRec.Amount)
    If IntRec.TaxYear = CurrTaxYear Then
      TotCurrInt = OldRound(TotCurrInt + IntRec.Amount)
    Else
      TotPastInt = OldRound(TotPastInt + IntRec.Amount)
    End If
    TCnt = TCnt + 1
    '                    7             8                9             10
    Print #RptHandle, TotInt; dlm; TotCurrInt; dlm; TotPastInt; dlm; TCnt
    For y = 1 To YrCnt
      If IntRec.TaxYear = Years(y) Then
        YearAmts(y) = OldRound(YearAmts(y) + IntRec.Amount)
        Exit For
      End If
    Next y
  Next x
  
  Close

  SubRptFile$ = "TAXRPTS\SUBTAXINT.RPT"     'Report File Name
  SubRptHandle = FreeFile
  Open SubRptFile$ For Output As #SubRptHandle
  
  For x = 1 To YrCnt
    Print #SubRptHandle, Years(x); dlm; YearAmts(x)
  Next x
  
  Close

  arTaxInterestRpt.Show
  MainLog ("Interest calculations completed successfully for date range " + fptxtFrom.Text + " to " + fptxtTo.Text = ".")
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxCalcInterest", "PrintGraphics", Erl)
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

Private Sub GetYears()
  Dim IntRec As InterestRecType
  Dim NumOfIRRecs As Long
  Dim IRHandle As Integer
  Dim x As Long, y As Integer
  Dim BigNum As Integer
  Dim HoldNum As Integer
  Dim Thisx As Integer
  Dim Nextx As Integer
  
  'on error goto ERRORSTUFF
  
  OpenInterestRecFile IRHandle, NumOfIRRecs
  ReDim Years(1 To 1) As Integer
  YrCnt = 0
  For x = 1 To NumOfIRRecs
    Get IRHandle, x, IntRec
    If x = 1 Then
      YrCnt = 1
      ReDim Preserve Years(1 To YrCnt) As Integer
      Years(YrCnt) = IntRec.TaxYear
    Else
      For y = 1 To YrCnt
        If IntRec.TaxYear = Years(y) Then
          Exit For
        End If
      Next y
      If y > YrCnt Then
        YrCnt = YrCnt + 1
        ReDim Preserve Years(1 To YrCnt) As Integer
        Years(YrCnt) = IntRec.TaxYear
      End If
    End If
  Next x
  
  Close IRHandle
  
  BigNum = -1
  Nextx = 1
  If YrCnt = 0 Then Exit Sub
  Do
    For x = Nextx To YrCnt
      If Years(x) > BigNum Then
        BigNum = Years(x)
        Thisx = x
      End If
    Next x
    HoldNum = Years(Nextx)
    Years(Nextx) = Years(Thisx)
    Years(Thisx) = HoldNum
    Nextx = Nextx + 1
    If Nextx > YrCnt Then Exit Do
    BigNum = -1
  Loop
    
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxCalcInterest", "GetYears", Erl)
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
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim IntRec As InterestRecType
  Dim NumOfIRRecs As Long
  Dim IRHandle As Integer
  Dim x As Long, y As Integer
  Dim Town$
  Dim Page As Integer
  Dim LineCnt As Integer
  Dim MaxLines As Integer
  Dim RptHandle As Integer
  Dim RptFile$, FF$
  Dim TotInt As Double
  Dim TotCurrInt As Double
  Dim TotPastInt As Double
  Dim TCnt As Long
  
  'on error goto ERRORSTUFF
  
  MaxLines = 56
  FF$ = Chr(12)
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  Town$ = QPTrim$(TaxMasterRec.Name)
  
  RptFile$ = "TAXRPTS\TAXINT.PRN"     'Report File Name
  RptHandle = FreeFile
  Open RptFile$ For Output As #RptHandle
  
  Call GetYears
  ReDim YearAmts(1 To YrCnt) As Double
  GoSub PrintHeader
  
  OpenInterestRecFile IRHandle, NumOfIRRecs
  For x = 1 To NumOfIRRecs
    Get IRHandle, x, IntRec
    Print #RptHandle, Using$("####0", IntRec.CustRec); Tab(8); QPTrim$(IntRec.CustName);
    Print #RptHandle, Tab(60); Using$("####", IntRec.TaxYear); Tab(65); Using$("####0", IntRec.BillNumber);
    Print #RptHandle, Tab(70); Using$("$###,##0.00", IntRec.Amount)
    TCnt = TCnt + 1
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintHeader
    End If
    TotInt = OldRound(TotInt + IntRec.Amount)
    If IntRec.TaxYear = CurrTaxYear Then
      TotCurrInt = OldRound(TotCurrInt + IntRec.Amount)
    Else
      TotPastInt = OldRound(TotPastInt + IntRec.Amount)
    End If
    For y = 1 To YrCnt
      If IntRec.TaxYear = Years(y) Then
        YearAmts(y) = OldRound(YearAmts(y) + IntRec.Amount)
        Exit For
      End If
    Next y
  Next x
  
  Print #RptHandle, FF$
  Print #RptHandle, Tab(15); "Property Tax Billing: Interest Calculation Register"
  Print #RptHandle, "Town: "; Tab(8); Town$; Tab(70); "Page #: " + CStr(Page)
  Print #RptHandle, "Date: " + CStr(Date)
  Print #RptHandle, "Current Tax Year: " + fptxtCurrYear.Text
  Print #RptHandle, String(80, "-")
  Print #RptHandle, Tab(2); "Total Transactions:     "; Tab(27); Using$("#####0", TCnt)
  Print #RptHandle, Tab(2); "Total Interest Charged: "; Tab(27); Using$("$###,###,##0.00", TotInt)
  Print #RptHandle, Tab(2); "Total Current Interest: "; Tab(27); Using$("$###,###,##0.00", TotCurrInt)
  Print #RptHandle, Tab(2); "Total Past Interest:    "; Tab(27); Using("$###,###,##0.00", TotPastInt)
  Print #RptHandle,
  Print #RptHandle, Tab(2); "Interest Breakdown by Year:"
  Print #RptHandle, Tab(4); "Year"; Tab(12); "Interest Calculation"
  For x = 1 To YrCnt
    Print #RptHandle, Tab(4); Using$("###0", Years(x)); Tab(17); Using$("$###,###,##0.00", YearAmts(x))
  Next x
  
  Close

  ViewPrint RptFile, "Interest Calculations", True
  
  MainLog ("Interest calculations completed successfully for date range " + fptxtFrom.Text + " to " + fptxtTo.Text = ".")
  Exit Sub
  
PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(15); "Property Tax Billing: Interest Calculation Register"
  Print #RptHandle, "Town: "; Tab(8); Town$; Tab(70); "Page #: " + CStr(Page)
  Print #RptHandle, "Date: " + CStr(Date)
  Print #RptHandle, "Current Tax Year: " + fptxtCurrYear.Text
  Print #RptHandle, "Acct #"; Tab(8); "Customer Name"; Tab(58); "Tax Yr"; Tab(65); "Bill #"; Tab(73); "Interest"
  Print #RptHandle, String(80, "-")
  LineCnt = 6
  Return
  
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxCalcInterest", "PrintText", Erl)
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
        fptxtPastYearIntRate.SetFocus
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
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
      fptxtCurrYear.SetFocus
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
