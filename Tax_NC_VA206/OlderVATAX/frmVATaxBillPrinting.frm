VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVATaxBillPrinting 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Bill Printing Information"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxBillPrinting.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbPrintPriorYN 
      Height          =   390
      Left            =   2520
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   8040
      Width           =   900
      _Version        =   196608
      _ExtentX        =   1587
      _ExtentY        =   688
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      ColDesigner     =   "frmVATaxBillPrinting.frx":08CA
   End
   Begin LpLib.fpCombo fpcmbType 
      Height          =   405
      Left            =   5490
      TabIndex        =   0
      Top             =   1785
      Width           =   2700
      _Version        =   196608
      _ExtentX        =   4762
      _ExtentY        =   714
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
      ColDesigner     =   "frmVATaxBillPrinting.frx":0C69
   End
   Begin LpLib.fpCombo fpcmbBarCode 
      Height          =   405
      Left            =   6690
      TabIndex        =   45
      ToolTipText     =   "For a bar code to appear on a bill the tax customer must have a valid 9 digit zip code and a delivery point value."
      Top             =   2265
      Width           =   900
      _Version        =   196608
      _ExtentX        =   1587
      _ExtentY        =   714
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
      ColDesigner     =   "frmVATaxBillPrinting.frx":1008
   End
   Begin LpLib.fpCombo fpcmbCommentPlace 
      Height          =   360
      Left            =   600
      TabIndex        =   48
      ToolTipText     =   $"frmVATaxBillPrinting.frx":13A7
      Top             =   1800
      Visible         =   0   'False
      Width           =   2100
      _Version        =   196608
      _ExtentX        =   3704
      _ExtentY        =   635
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      ColDesigner     =   "frmVATaxBillPrinting.frx":1469
   End
   Begin VB.OptionButton OptMort 
      BackColor       =   &H008F8265&
      Caption         =   "Print By Mortgage Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4613
      TabIndex        =   55
      Top             =   7440
      Width           =   2415
   End
   Begin VB.OptionButton optLeave 
      BackColor       =   &H008F8265&
      Caption         =   "Print Unsorted"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2040
      TabIndex        =   54
      Top             =   7440
      Width           =   1695
   End
   Begin EditLib.fpText fptxtComment2 
      Height          =   375
      Left            =   6503
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   4695
      _Version        =   196608
      _ExtentX        =   8281
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
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      MaxLength       =   40
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
   Begin EditLib.fpDoubleSingle fpDblSnglMHRate 
      Height          =   372
      Left            =   9552
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4764
      Width           =   1332
      _Version        =   196608
      _ExtentX        =   2350
      _ExtentY        =   656
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
      NoSpecialKeys   =   2
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
      ControlType     =   1
      Text            =   "0.0000"
      DecimalPlaces   =   4
      DecimalPoint    =   "."
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpDateTime fptxtPersDueDate 
      Height          =   375
      Left            =   8880
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6660
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
   Begin EditLib.fpText fptxtRealOrder 
      Height          =   375
      Left            =   1170
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6300
      Width           =   2775
      _Version        =   196608
      _ExtentX        =   4890
      _ExtentY        =   656
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
      AutoAdvance     =   0   'False
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
      ControlType     =   1
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin EditLib.fpDoubleSingle fpDblSnglStartRealBill 
      Height          =   375
      Left            =   3090
      TabIndex        =   1
      Top             =   3450
      Width           =   1215
      _Version        =   196608
      _ExtentX        =   2138
      _ExtentY        =   656
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
      Text            =   "0"
      DecimalPlaces   =   -1
      DecimalPoint    =   ""
      FixedPoint      =   0   'False
      LeadZero        =   0
      MaxValue        =   "999999999"
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
   Begin EditLib.fpLongInteger fpLongRealTaxYear 
      Height          =   375
      Left            =   2850
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3900
      Width           =   1095
      _Version        =   196608
      _ExtentX        =   1926
      _ExtentY        =   656
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
      ControlType     =   1
      Text            =   "0"
      MaxValue        =   "2147483647"
      MinValue        =   "-2147483648"
      NegFormat       =   1
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
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
   Begin EditLib.fpText fptxtCurrForm 
      Height          =   396
      Left            =   5472
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Late notices are selected on the System Setup screen."
      Top             =   1188
      Width           =   2856
      _Version        =   196608
      _ExtentX        =   5038
      _ExtentY        =   698
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
      AutoCase        =   1
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
      ControlType     =   1
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   50
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
   Begin EditLib.fpDateTime fptxtRealDueDate 
      Height          =   375
      Left            =   2490
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6825
      Width           =   1740
      _Version        =   196608
      _ExtentX        =   3069
      _ExtentY        =   656
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
      Height          =   495
      Left            =   9150
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   8025
      Width           =   2040
      _Version        =   131072
      _ExtentX        =   3598
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmVATaxBillPrinting.frx":1808
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   495
      Left            =   4290
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   8040
      Width           =   2040
      _Version        =   131072
      _ExtentX        =   3598
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmVATaxBillPrinting.frx":19E7
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAlign 
      Height          =   495
      Left            =   6720
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   8025
      Width           =   2055
      _Version        =   131072
      _ExtentX        =   3625
      _ExtentY        =   873
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
      ButtonDesigner  =   "frmVATaxBillPrinting.frx":1BC3
   End
   Begin EditLib.fpDoubleSingle fpDblSnglStartPersBill 
      Height          =   372
      Left            =   8496
      TabIndex        =   7
      Top             =   3456
      Width           =   1212
      _Version        =   196608
      _ExtentX        =   2138
      _ExtentY        =   656
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
      Text            =   "0"
      DecimalPlaces   =   -1
      DecimalPoint    =   ""
      FixedPoint      =   0   'False
      LeadZero        =   0
      MaxValue        =   "999999999"
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
   Begin EditLib.fpDoubleSingle fpDblSnglRealRate 
      Height          =   375
      Left            =   2850
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4770
      Width           =   1335
      _Version        =   196608
      _ExtentX        =   2350
      _ExtentY        =   656
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
      ControlType     =   1
      Text            =   "0.0000"
      DecimalPlaces   =   4
      DecimalPoint    =   "."
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpDoubleSingle fpDblSnglPersRate 
      Height          =   372
      Left            =   6696
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4764
      Width           =   1332
      _Version        =   196608
      _ExtentX        =   2350
      _ExtentY        =   656
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
      NoSpecialKeys   =   2
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
      ControlType     =   1
      Text            =   "0.0000"
      DecimalPlaces   =   4
      DecimalPoint    =   "."
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpDoubleSingle fpDblSnglRealLateList 
      Height          =   375
      Left            =   2850
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5235
      Width           =   1335
      _Version        =   196608
      _ExtentX        =   2350
      _ExtentY        =   656
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
      ControlType     =   1
      Text            =   "0.0000"
      DecimalPlaces   =   4
      DecimalPoint    =   "."
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpDoubleSingle fpDblSnglPersLateList 
      Height          =   372
      Left            =   9552
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5676
      Width           =   1332
      _Version        =   196608
      _ExtentX        =   2350
      _ExtentY        =   656
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
      ControlType     =   1
      Text            =   "0.0000"
      DecimalPlaces   =   4
      DecimalPoint    =   "."
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpLongInteger fpLongPersTaxYear 
      Height          =   372
      Left            =   8376
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3900
      Width           =   1092
      _Version        =   196608
      _ExtentX        =   1926
      _ExtentY        =   656
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
      ControlType     =   1
      Text            =   "0"
      MaxValue        =   "2147483647"
      MinValue        =   "-2147483648"
      NegFormat       =   1
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
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
   Begin EditLib.fpText fptxtPersOrder 
      Height          =   375
      Left            =   5250
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   6660
      Width           =   2775
      _Version        =   196608
      _ExtentX        =   4890
      _ExtentY        =   656
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
      AutoAdvance     =   0   'False
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
      ControlType     =   1
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin EditLib.fpDoubleSingle fpDblSnglMTRate 
      Height          =   372
      Left            =   6696
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5220
      Width           =   1332
      _Version        =   196608
      _ExtentX        =   2350
      _ExtentY        =   656
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
      ControlType     =   1
      Text            =   "0.0000"
      DecimalPlaces   =   4
      DecimalPoint    =   "."
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpDoubleSingle fpDblSnglFERate 
      Height          =   372
      Left            =   6696
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5676
      Width           =   1332
      _Version        =   196608
      _ExtentX        =   2350
      _ExtentY        =   656
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
      ControlType     =   1
      Text            =   "0.0000"
      DecimalPlaces   =   4
      DecimalPoint    =   "."
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpDoubleSingle fpDblSnglMCRate 
      Height          =   372
      Left            =   9552
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5220
      Width           =   1332
      _Version        =   196608
      _ExtentX        =   2350
      _ExtentY        =   656
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
      NoSpecialKeys   =   2
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
      ControlType     =   1
      Text            =   "0.0000"
      DecimalPlaces   =   4
      DecimalPoint    =   "."
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpText fptxtComment 
      Height          =   375
      Left            =   1748
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   4695
      _Version        =   196608
      _ExtentX        =   8281
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
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      MaxLength       =   40
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
   Begin EditLib.fpDoubleSingle fpDSMnthlyPen 
      Height          =   375
      Left            =   9120
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
      _Version        =   196608
      _ExtentX        =   2350
      _ExtentY        =   656
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
      NoSpecialKeys   =   2
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
   Begin VB.OptionButton optZip 
      BackColor       =   &H008F8265&
      Caption         =   "Print By Zip Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7920
      TabIndex        =   56
      Top             =   7440
      Width           =   1815
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Prior Year Balance (Y/N?):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   840
      TabIndex        =   58
      Top             =   7920
      Width           =   1665
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   360
      X2              =   11220
      Y1              =   7800
      Y2              =   7800
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   11215
      X2              =   11215
      Y1              =   7320
      Y2              =   7800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   390
      X2              =   390
      Y1              =   7320
      Y2              =   7800
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Monthly Penalty %"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   8640
      TabIndex        =   53
      Top             =   1440
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Comment Placement"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   480
      TabIndex        =   51
      Top             =   1560
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Comment:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   443
      TabIndex        =   47
      Top             =   2400
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Use BarCode (Y/N?):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   4296
      TabIndex        =   46
      Top             =   2388
      Width           =   2268
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Merch Cap:"
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
      Height          =   372
      Left            =   8112
      TabIndex        =   44
      Top             =   5304
      Width           =   1332
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Rate"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   252
      Left            =   9672
      TabIndex        =   43
      Top             =   4500
      Width           =   1092
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mbl Homes:"
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
      Height          =   372
      Left            =   8112
      TabIndex        =   42
      Top             =   4860
      Width           =   1332
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Farm Equip:"
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
      Height          =   372
      Left            =   5112
      TabIndex        =   41
      Top             =   5736
      Width           =   1452
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mach/Tools:"
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
      Height          =   372
      Left            =   5112
      TabIndex        =   40
      Top             =   5304
      Width           =   1452
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PERSONAL PROPERTY"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   372
      Left            =   6456
      TabIndex        =   39
      Top             =   2904
      Width           =   3012
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "REAL PROPERTY"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   372
      Left            =   936
      TabIndex        =   38
      Top             =   2904
      Width           =   3012
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   456
      Left            =   4776
      Top             =   2820
      Width           =   6468
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   456
      Left            =   396
      Top             =   2820
      Width           =   4332
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   4080
      Left            =   4770
      Top             =   3300
      Width           =   6465
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Rate"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   252
      Left            =   6816
      TabIndex        =   37
      Top             =   4500
      Width           =   1092
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1776
      Left            =   5016
      Top             =   4380
      Width           =   5988
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pers Due Date:"
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
      Height          =   255
      Left            =   8760
      TabIndex        =   36
      Top             =   6300
      Width           =   1815
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pers Printing Order:"
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
      Left            =   5475
      TabIndex        =   35
      Top             =   6300
      Width           =   2295
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pers Tax Year:"
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
      Height          =   252
      Left            =   6576
      TabIndex        =   34
      Top             =   3984
      Width           =   1692
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Late List:"
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
      Height          =   252
      Left            =   8352
      TabIndex        =   33
      Top             =   5736
      Width           =   1092
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Pers Bill No:"
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
      Height          =   372
      Left            =   6096
      TabIndex        =   32
      Top             =   3528
      Width           =   2292
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Bill Type:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   360
      Left            =   3456
      TabIndex        =   31
      Top             =   1908
      Width           =   1932
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Real Due Date:"
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
      Height          =   255
      Left            =   570
      TabIndex        =   27
      Top             =   6885
      Width           =   1815
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Format In Use:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   3312
      TabIndex        =   26
      Top             =   1260
      Width           =   2028
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Real Printing Order:"
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
      Left            =   1410
      TabIndex        =   24
      Top             =   5925
      Width           =   2295
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1410
      Left            =   690
      Top             =   4380
      Width           =   3735
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   4080
      Left            =   390
      Top             =   3300
      Width           =   4335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Bill Printing Information"
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
      TabIndex        =   23
      Top             =   492
      Width           =   5292
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   660
      Index           =   1
      Left            =   1500
      Top             =   324
      Width           =   8652
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Real Tax Year:"
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
      Height          =   255
      Left            =   1050
      TabIndex        =   22
      Top             =   4005
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Real Bill No:"
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
      Left            =   690
      TabIndex        =   21
      Top             =   3525
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Real Estate:"
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
      Height          =   255
      Left            =   930
      TabIndex        =   20
      Top             =   4890
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pers Prop:"
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
      Height          =   372
      Left            =   5352
      TabIndex        =   19
      Top             =   4860
      Width           =   1212
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Real Late List:"
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
      Height          =   255
      Left            =   810
      TabIndex        =   18
      Top             =   5310
      Width           =   1815
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Rate"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   2970
      TabIndex        =   17
      Top             =   4485
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   1500
      Top             =   228
      Width           =   8652
   End
End
Attribute VB_Name = "frmVATaxBillPrinting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  'Private Temp_Class As Resize_Class
  Dim BillFormat$
  Dim MTTaxRate#, FETaxRate#, MHTaxRate#
  Dim MCTaxRate#, PersTaxRate#, PLateRate#
  Dim RLateRate#, RealRate#, PersYear As Integer
  Dim RealYear As Integer

Private Sub cmdAlign_Click()
  Dim Handle As Integer
  Dim TempHandle As Integer
  Dim cnt As Integer
  Dim TextLine$
  
  On Error GoTo ERRORSTUFF
  
  If fpcmbType.Text <> "REAL" And fpcmbType.Text <> "PERSONAL" Then
    Call TaxMsg(800, "No REAL or PERSONAL DESIGNATED. Alignment aborted.")
    Exit Sub
  End If
  
  If fpcmbType.Text = "REAL" Then
    Select Case fptxtCurrForm.Text
      Case "STANDARD"
        Call TaxMsg(900, "Pitch 10 is recommended for this bill.")
        If Exist("VASTANDRMSK.TXT") Then
          alnRpt = "VASTANDRMSK.TXT"
        Else
          Call TaxMsg(900, "Please install the file 'VASTANDRMSK.TXT'.")
          Exit Sub
        End If
      Case "MDLTWN"
        Call TaxMsg(900, "Pitch 10 is recommended for this bill.")
        If Exist("MdltwnRMask.TXT") Then
          alnRpt = "MdltwnRMask.TXT"
        Else
          Call TaxMsg(900, "Please install the file 'MdltwnRMask.TXT'.")
          Exit Sub
        End If
      Case "CDRBLUFF"
        Call TaxMsg(900, "Pitch 12 is recommended for this bill.")
        If Exist("CdrBluffRMask.TXT") Then
          alnRpt = "CdrBluffRMask.TXT"
        Else
          Call TaxMsg(900, "Please install the file 'CdrBluffRMask.TXT'.")
          Exit Sub
        End If
      Case Else
        Call TaxMsg(900, "No mask is available.")
        Exit Sub
    End Select
  End If
    
  If fpcmbType.Text = "PERSONAL" Then
    Select Case fptxtCurrForm.Text
      Case "STANDARD"
        Call TaxMsg(900, "Pitch 10 is recommended for this bill.")
        If Exist("VASTANDPMSK.TXT") Then
          alnRpt = "VASTANDPMSK.TXT"
        Else
          Call TaxMsg(900, "Please install the file 'VASTANDPMSK.TXT'.")
          Exit Sub
        End If
      Case "MDLTWN"
        Call TaxMsg(900, "Pitch 10 is recommended for this bill.")
        If Exist("MdltwnPMask.TXT") Then
          alnRpt = "MdltwnPMask.TXT"
        Else
          Call TaxMsg(900, "Please install the file 'MdltwnPMask.TXT'.")
          Exit Sub
        End If
      Case "CDRBLUFF"
        Call TaxMsg(900, "Pitch 10 is recommended for this bill.")
        If Exist("CdrBluffPMask.TXT") Then
          alnRpt = "CdrBluffPMask.TXT"
        Else
          Call TaxMsg(900, "Please install the file 'CdrBluffPMask.TXT'.")
          Exit Sub
        End If
      Case Else
        Call TaxMsg(900, "No mask is available.")
        Exit Sub
    End Select
  End If
  
  Handle = FreeFile
'  If fpcmbType.Text = "REAL" Then
'
'    alnRpt = "TAXREMSK.DAT"
'  Else
'    alnRpt = "TAXPPMSK.DAT"
'  End If
  Open alnRpt For Input As #Handle
  TempHandle = FreeFile
  Open "TAXALIGN.MSK" For Output As #TempHandle
  Do While Not eof(Handle)
    Line Input #Handle, TextLine   ' Read line into variable.
    Print #TempHandle, TextLine
  Loop
  Close
  alnRpt = "TAXALIGN.MSK"
  doAlign = True
  frmVATaxPrint.Show vbModal
  alnRpt = ""
  doAlign = False

  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxBillPrinting", "cmdAlign_Click", Erl)
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

Private Sub cmdExit_Click()
  Close 'added in case bill printing is stopped before completion
  'leaving a print file open
  frmVATaxBillPrintMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdProcess_Click()
  Dim RealBillInfo As VARETaxBillInfoType
  Dim PersBillInfo As VAPPTaxBillInfoType
  Dim BIHandle As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  If fpcmbType.Text = "REAL" Then
    If optLeave.Value = True Then
      KillFile "RZIPIDX.DAT"
      KillFile "MORTIDX.DAT"
    End If
    If optZip.Value = True Then
      If Not Exist("RZIPIDX.DAT") Then
        Call SortByZipCode
      End If
    End If
    If OptMort.Value = True Then
      If Not Exist("TAXMORT.DAT") Then
        Call TaxMsg(900, "No mortgage codes are saved. Attempt to index by mortgage codes is aborted.")
        OptMort.Value = False
        optLeave.Value = True
      End If
      If Not Exist("MORTIDX.DAT") Then
        Call SortByMortCode
      End If
    End If
    If fpDblSnglStartRealBill <= 0 Then
      Call TaxMsg(900, "Please enter a valid Real bill number.")
      fpDblSnglStartRealBill.SetFocus
      Exit Sub
    End If
    Select Case fptxtCurrForm.Text
      Case "STANDARD"
        Call TaxMsg(900, "Pitch 10 is recommended for this bill.")
        If InStr(TaxMasterRec.Name, "HALIFAX") Then
          Call PrintHalifaxStandardReal
        Else
          Call PrintRealVAStandard
        End If
      Case "LASER", "LASER ITEMIZED"
        Call PrintLaserReal1
      Case "EXPORT REAL"
        Call PrintExpReal
      Case "MDLTWN"
        Call PrintRealMiddletownBill
      Case "CDRBLUFF"
        Call TaxMsg(900, "Pitch 12 is recommended for this bill.")
        Call PrintCedarBluffReal
      Case Else
        Call TaxMsg(700, "The current bill format, " + QPTrim$(fptxtCurrForm.Text) + ", is not set up for bill printing at this time. Please select a different format from the Tax System Setup screen.")
        Close
        Exit Sub
    End Select
    OpenRealBillInfoFile BIHandle
    Get BIHandle, 1, RealBillInfo
    RealBillInfo.BillNum = fpDblSnglStartRealBill
    Put BIHandle, 1, RealBillInfo
    Close BIHandle
    Exit Sub
  End If
  
  If fpcmbType.Text = "PERSONAL" Then
    If optLeave.Value = True Then
      KillFile "PZIPIDX.DAT"
    End If
    If optZip.Value = True Then
      If Not Exist("PZIPIDX.DAT") Then
        Call SortByZipCode
      End If
    End If
    If fpDblSnglStartPersBill <= 0 Then
      Call TaxMsg(900, "Please enter a valid Personal bill number.")
      fpDblSnglStartPersBill.SetFocus
      Exit Sub
    End If
    Select Case fptxtCurrForm.Text
      Case "STANDARD"
        Call TaxMsg(900, "Pitch 10 is recommended for this bill.")
        If InStr(TaxMasterRec.Name, "HALIFAX") Then
          Call PrintHalifaxStandardPersonal
        Else
          Call PrintPersVAStandard
        End If
      Case "LASER"
        Call PrintLaserPers1
      Case "EXPORT PERSONAL"
        Call PrintPersExport
      Case "LASER ITEMIZED"
        Call PrintPersLaserItemized
      Case "MDLTWN"
        Call PrintPersMiddletownBill
      Case "CDRBLUFF"
        Call TaxMsg(900, "Pitch 10 is recommended for this bill.")
        Call PrintCedarBluffPersonal
      Case Else
        Call TaxMsg(700, "The current bill format, " + QPTrim$(fptxtCurrForm.Text) + ", is not set up for bill printing at this time. Please select a different format from the Tax System Setup screen.")
        Close
        Exit Sub
    End Select
    OpenPersBillInfoFile BIHandle
    Get BIHandle, 1, PersBillInfo
    PersBillInfo.BillNum = fpDblSnglStartPersBill
    Put BIHandle, 1, PersBillInfo
    Close BIHandle
    Exit Sub
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
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdProcess_Click
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
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxBillPrinting.")
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

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  'Set Temp_Class = New Resize_Class
  'Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  MainLog ("User opened frmVATaxBillPrinting.")
  Me.HelpContextID = hlpPrintTaxBills
  Call LoadMe
End Sub

Private Sub LoadMe()
  Dim RealBillInfo As VARETaxBillInfoType
  Dim BIHandle As Integer
  Dim PersBillInfo As VAPPTaxBillInfoType
  Dim IdxType As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim MortCodeRec As MortCodeRecType
  Dim MCHandle As Integer
  Dim NumMortCodes As Integer
  Dim x As Integer
  
  On Error GoTo ERRORSTUFF
  
  fpcmbPrintPriorYN.Visible = False
  Label29.Visible = False
  
  doAlign = False
  If Exist(RealTaxBillInfoFile) Then
    OpenRealBillInfoFile BIHandle
    Get BIHandle, 1, RealBillInfo
    Close BIHandle
    fpDblSnglRealRate = RealBillInfo.RealRate
    RealRate = RealBillInfo.RealRate
    fpDblSnglRealLateList = RealBillInfo.LATEPCT
    RLateRate = RealBillInfo.LATEPCT
    fpLongRealTaxYear = RealBillInfo.TaxYear
    RealYear = RealBillInfo.TaxYear
    fptxtRealDueDate = MakeRegDate(RealBillInfo.DueDate)
    If fptxtRealDueDate.Text = "12/31/1979" Then fptxtRealDueDate.Text = "N/A"
    Select Case QPTrim$(RealBillInfo.PRNORDER)
      Case "1"
        fptxtRealOrder.Text = "Account Number Order"
      Case "2"
        fptxtRealOrder.Text = "Customer Name Order"
      Case "3"
        fptxtRealOrder.Text = "Search Name Order"
      Case "4"
        fptxtRealOrder.Text = "Social Security Order"
      Case Else
        fptxtRealOrder.Text = "Unknown"
     End Select
    If RealBillInfo.BillNum > 0 Then
      fpDblSnglStartRealBill = RealBillInfo.BillNum
    Else
      fpDblSnglStartRealBill = 0
    End If
  Else
    fpDblSnglRealRate.Enabled = False
    fpDblSnglRealRate = 0
    fpDblSnglRealLateList = 0
    fpDblSnglRealLateList.Enabled = False
    fpLongRealTaxYear = 0
    fpLongRealTaxYear.Enabled = False
    fptxtRealDueDate = "N/A"
    fptxtRealDueDate.Enabled = False
    fpDblSnglStartRealBill = 0
    fpDblSnglStartRealBill.Enabled = False
    fptxtRealOrder.Text = "N/A"
    fptxtRealOrder.Enabled = False
  End If
   
  If Exist(PersTaxBillInfoFile) Then
    OpenPersBillInfoFile BIHandle
    Get BIHandle, 1, PersBillInfo
    Close BIHandle
    fpDblSnglPersRate = PersBillInfo.PERSRATE
    PersTaxRate# = PersBillInfo.PERSRATE
    fpDblSnglMCRate = PersBillInfo.MCRate
    MCTaxRate# = PersBillInfo.MCRate
    fpDblSnglFERate = PersBillInfo.FERate
    FETaxRate# = PersBillInfo.FERate
    fpDblSnglMTRate = PersBillInfo.MTRate
    MTTaxRate# = PersBillInfo.MTRate
    fpDblSnglMHRate = PersBillInfo.MHRate
    MHTaxRate# = PersBillInfo.MHRate
    fpDblSnglPersLateList = PersBillInfo.LATEPCT
    PLateRate# = PersBillInfo.LATEPCT
    fpLongPersTaxYear = PersBillInfo.TaxYear
    PersYear = PersBillInfo.TaxYear
    fptxtPersDueDate = MakeRegDate(PersBillInfo.DueDate)
    Select Case QPTrim$(PersBillInfo.PRNORDER)
      Case "1"
        fptxtPersOrder.Text = "Account Number Order"
      Case "2"
        fptxtPersOrder.Text = "Customer Name Order"
      Case "3"
        fptxtPersOrder.Text = "Search Name Order"
      Case "4"
        fptxtPersOrder.Text = "Social Security Order"
      Case Else
        fptxtPersOrder.Text = "Unknown"
    End Select
    If PersBillInfo.BillNum > 0 Then
      fpDblSnglStartPersBill = PersBillInfo.BillNum
    Else
      fpDblSnglStartPersBill = 0
    End If
  Else
    fpDblSnglPersRate = 0
    fpDblSnglPersRate.Enabled = False
    fpDblSnglMCRate = 0
    fpDblSnglMCRate.Enabled = False
    fpDblSnglFERate = 0
    fpDblSnglFERate.Enabled = False
    fpDblSnglMTRate = 0
    fpDblSnglMTRate.Enabled = False
    fpDblSnglMHRate = 0
    fpDblSnglMHRate.Enabled = False
    fpDblSnglPersLateList = 0
    fpDblSnglPersLateList.Enabled = False
    fpLongPersTaxYear = 0
    fpLongPersTaxYear.Enabled = False
    fptxtPersDueDate = "N/A"
    fptxtPersDueDate.Enabled = False
    fpDblSnglStartPersBill = 0
    fpDblSnglStartPersBill.Enabled = False
    fptxtPersOrder.Text = "N/A"
    fptxtPersOrder.Enabled = False
  End If
  
  If RealBillInfo.RealRate > 0 Then
    fpcmbType.Text = "REAL"
  ElseIf PersBillInfo.PERSRATE > 0 Then
    fpcmbType.Text = "PERSONAL"
  End If
  
  If RealBillInfo.RealRate > 0 And PersBillInfo.PERSRATE > 0 Then
    fpcmbType.AddItem "REAL"
    fpcmbType.AddItem "PERSONAL"
  End If
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  If fpcmbType.Text = "REAL" Then
    Call DisablePersonal
    Call EnableReal
  ElseIf fpcmbType.Text = "PERSONAL" Then
    Call EnablePersonal
    Call DisableReal
  End If
  
  If fpcmbType.Text = "REAL" Then
    OptMort.Enabled = True
    If Exist("MORTIDX.DAT") Then
      OptMort.Value = True
    ElseIf Exist("RZIPIDX.DAT") Then
      optZip.Value = True
    Else
      optLeave.Value = True
    End If
  ElseIf fpcmbType.Text = "PERSONAL" Then
    OptMort.Enabled = False
    If Exist("PZIPIDX.DAT") Then
      optZip.Value = True
    Else
      optLeave.Value = True
    End If
  End If
  
  BillFormat$ = Left$(TaxMasterRec.TaxForm, 1)
  If TaxMasterRec.TaxForm = 16716 Or TaxMasterRec.TaxForm = 20002 Then
    cmdAlign.Enabled = False
  End If
  Select Case TaxMasterRec.TaxForm
    Case 16716
      fptxtCurrForm.Text = "LASER"
    Case 29999
      fptxtCurrForm.Text = "EXPORT COMBINED"
      OptMort.Enabled = False
      optLeave.Enabled = False
      optZip.Enabled = False
    Case 30000
      fptxtCurrForm.Text = "STANDARD"
      If InStr(TaxMasterRec.Name, "HALIFAX") = 0 Then
        Label26.Visible = True
        Label27.Visible = True
        fpcmbCommentPlace.Text = "NO COMMENTS"
        fpcmbCommentPlace.AddItem "BOTTOM RIGHT"
        fpcmbCommentPlace.AddItem "BOTTOM LEFT"
        fpcmbCommentPlace.AddItem "NO COMMENTS"
        fptxtComment.Visible = True
        fptxtComment2.Visible = True
        fpcmbCommentPlace.Visible = True
        CheckComment
      End If
    Case 20000
      fptxtCurrForm.Text = "EXPORT REAL"
      fptxtRealDueDate.Visible = True
      fptxtRealDueDate.Text = Date
      Label9.Visible = True
      OptMort.Enabled = False
      optLeave.Enabled = False
      optZip.Enabled = False
    Case 20001
      fptxtCurrForm.Text = "EXPORT PERSONAL"
      Label15.Visible = True
      OptMort.Enabled = False
      optLeave.Enabled = False
      optZip.Enabled = False
    Case 20002
      fptxtCurrForm.Text = "LASER ITEMIZED"
    Case 20003
      fptxtCurrForm.Text = "MDLTWN"
    Case 20004
      fptxtCurrForm.Text = "CDRBLUFF"
      If fpcmbType.Text = "REAL" Then
        fpDSMnthlyPen.Visible = True
        Label28.Visible = True
        Call CheckComment
      End If
    Case Else
      fptxtCurrForm.Text = "UNKNOWN"
  End Select
  fptxtPersDueDate.Text = Date
  If fptxtCurrForm.Text = "LASER" Or fptxtCurrForm.Text = "LASER ITEMIZED" Then
    If fptxtCurrForm.Text = "LASER" Then
      fpcmbPrintPriorYN.Text = "N"
      fpcmbPrintPriorYN.AddItem "Y"
     fpcmbPrintPriorYN.AddItem "N"
      fpcmbPrintPriorYN.Visible = True
      Label29.Visible = True
    End If
    fpcmbBarCode.Visible = True
    Label25.Visible = True
  Else
    If fptxtCurrForm.Text <> "LASER" Then
      fpcmbPrintPriorYN.Visible = False
      Label29.Visible = False
    End If
    fpcmbBarCode.Visible = False
    Label25.Visible = False
  End If
  
  fpcmbBarCode.Text = "No"
  fpcmbBarCode.AddItem "No"
  fpcmbBarCode.AddItem "Yes"
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxBillPrinting", "LoadMe", Erl)
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

Private Sub MakeFile()
  Dim One As Integer
  Dim AHandle As Integer
  
  If fpcmbType.Text = "REAL" Then
    KillFile "txrblsprn.dat"
    One = 1
    AHandle = FreeFile
    Open "txrblsprn.dat" For Output As AHandle
    Print #AHandle, One
    Close AHandle
  ElseIf fpcmbType.Text = "PERSONAL" Then
    KillFile "txpblsprn.dat"
    One = 1
    AHandle = FreeFile
    Open "txpblsprn.dat" For Output As AHandle
    Print #AHandle, One
    Close AHandle
  End If
End Sub


Private Sub PrintMulti()
'  Dim RptHandle As Integer
'  Dim TaxBill As TaxBillType
'  Dim TBHandle As Integer
'  Dim NumOfTBRecs As Long
'  Dim x As Long, BillNo&
'  Dim WhatRec&, PrnCnt&
'  Dim PINTemp As String * 18
'  Dim CustName$, PCnt As Long
'  Dim RptFile$
'
'  On Error GoTo ERRORSTUFF
'
'  BillNo& = fpDblSnglStartBill.Value
'  RptHandle = FreeFile
'  RptFile$ = "TAXBIL.PRN"
'
'  Open RptFile For Output As RptHandle
'
'  OpenRealTaxBillFile TBHandle, NumOfTBRecs
'  frmVATaxShowPctComp.Label1 = "Printing Tax Bills"
'  frmVATaxShowPctComp.cmdCancel.Visible = False
'  frmVATaxShowPctComp.Show, Me
'  cmdProcess.Enabled = False
'  cmdExit.Enabled = False
'  cmdAlign.Enabled = False
'  EnableCloseButton Me.hwnd, False
'
'  For x = 1 To NumOfTBRecs
'    WhatRec& = x
'    Get TBHandle, WhatRec&, TaxBill
'    If TaxBill.BillNumber > 0 Then
'      If TaxBill.TotalBillDue > 0 Then
'        RSet PINTemp = TaxBill.RealPin
'        CustName$ = QPTrim$(TaxBill.CustName)
'        TaxBill.BillNumber = BillNo&
'        TaxBill.BillPrinted = True
'        Put TBHandle, WhatRec&, TaxBill
'        Call PrintStandard(RptHandle, TBHandle, TaxBill, CustName, PINTemp, PrnCnt)
'        BillNo& = BillNo& + 1
'        PrnCnt = PrnCnt + 1
'      End If
'    End If
'    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
'    If frmVATaxShowPctComp.Out = True Then
'      Close
'      frmVATaxShowPctComp.Out = False
'      Unload frmVATaxShowPctComp
'      cmdProcess.Enabled = True
'      cmdExit.Enabled = True
'      cmdAlign.Enabled = True
'      EnableCloseButton Me.hwnd, True
'      Exit Sub
'    End If
'  Next x
'  Unload frmVATaxShowPctComp
'  cmdProcess.Enabled = True
'  cmdExit.Enabled = True
'  cmdAlign.Enabled = True
'  EnableCloseButton Me.hwnd, True
'  Close TBHandle
'  Close RptHandle
'  ViewPrint RptFile$, "Tax Bill Printing", True
'  If PrnCnt > 0 Then
'    Call MakeFile 'Makefile creates the .dat identifying that bill printing took place
'  End If
'
'  Exit Sub
'
'ERRORSTUFF:
'   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxBillPrinting", "PrintMulti", Erl)
'     Case emrExitProc:
'       Resume Proc_Exit
'     Case emrResume:
'       Resume
'     Case emrResumeNext:
'       Resume Next
'     Case Else
'      '--- Technically, this should never happen.
'       Resume Proc_Exit
'   End Select
'
'Proc_Exit:
'  '--- Cleanup code goes here...
'    Close
'
End Sub

Private Sub PrintExpReal()
  Dim x As Long
  Dim TaxXRec As TaxBillExportRealType
  Dim TXHandle As Integer
  Dim TBRec As VARETaxBillType
  Dim TempCustRec As TaxCustType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim ReportFile As String
  Dim RptHandle As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim TaxYear$
  Dim FF10$
  Dim Map$
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRRecs As Long
  Dim BillNo&
  Dim PrnCnt As Long
  Dim MLBS As String
  Dim NumOfTCRecs As Long, CHandle As Integer
  Dim PastDue As Double
  
  On Error GoTo ERRORSTUFF
  PrnCnt = 0
  FF10$ = "#######.#0"
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  TaxYear$ = CStr(TaxMasterRec.RTaxYear)
  ReportFile$ = "LCRE" + TaxYear + ".TXT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  BillNo& = fpDblSnglStartRealBill.Value
  
  OpenRealTaxBillFile TBHandle, NumOfTBRecs
  
'dale hacked
  OpenTaxCustFile CHandle, NumOfTCRecs
  OpenRealPropFile RHandle, NumOfRRecs
  
  frmVATaxShowPctComp.Label1 = "Printing Tax Billing Data"
  frmVATaxShowPctComp.Show , Me
  frmVATaxShowPctComp.cmdCancel.Visible = False
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  
  For x = 1 To NumOfTBRecs
    Get TBHandle, x, TBRec
    If TBRec.BillNumber > 0 Then
      
      Get CHandle, TBRec.CustRec, TempCustRec
      Get RHandle, TempCustRec.FirstPropRec, RealRec
      
'      If OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt) > 0 Then
      If OldRound(TBRec.TotalBillDue) > 0 Then '5/11/07
        TBRec.BillNumber = BillNo&
        TBRec.BillPrinted = True
        Put TBHandle, x, TBRec
        BillNo& = BillNo& + 1
      End If
      RSet TaxXRec.TaxBillNum = Using$("######0", TBRec.BillNumber)
      RSet TaxXRec.CustName = QPTrim$(TBRec.CustName)
      RSet TaxXRec.Add1 = QPTrim$(TBRec.CustAdd1)
      RSet TaxXRec.Add2 = QPTrim$(TBRec.CustAdd2)
      RSet TaxXRec.Add3 = QPTrim$(TBRec.CustAdd3) + " " + QPTrim$(TBRec.CustZip)
      RSet TaxXRec.TaxYear = Using$("###0", TBRec.TaxYear)
      RSet TaxXRec.CustAcct = Using$("####0", TBRec.CustPin)
'      If TBRec.RealPropRecord > 0 Then
'        Get RHandle, TBRec.RealPropRecord, RealRec
      LSet TaxXRec.MapNum = QPTrim$(TBRec.RealPin)   'QPTrim$(RealRec.RealPin) 'changed 10/25/06
'      Else
'        RSet TaxXRec.MapNum = "NA"
'      End If

'Dale Hacked this up lol........;-}
      
      MLBS$ = QPTrim$(RealRec.Map) + " " + QPTrim$(RealRec.BLOCK) + " " + QPTrim$(RealRec.LOTNUMB)
      MLBS$ = MLBS$ + " " + QPTrim$(Str$(RealRec.PropSize)) + " " + QPTrim$(RealRec.LOTACRE)
      RSet TaxXRec.MBLASize = MLBS$
      
      RSet TaxXRec.PropDesc1 = QPTrim$(TBRec.RDesc1)
      RSet TaxXRec.TAXRATE = Using$("#.#0", TBRec.RealTaxRate)
      RSet TaxXRec.LandVal = Using$(FF10, TBRec.RealValue)
      RSet TaxXRec.BldgVal = Using$(FF10$, TBRec.BldgValue)
      RSet TaxXRec.RealVal = Using$(FF10$, TBRec.RealValue + TBRec.BldgValue)
      RSet TaxXRec.CurrTaxAmt = Using$(FF10$, OldRound(TBRec.RealTaxDue - TBRec.OverPayAmt))
      RSet TaxXRec.PropDesc2 = QPTrim$(TBRec.RDesc2)
      RSet TaxXRec.PropDesc3 = QPTrim$(TBRec.RDesc3)
      RSet TaxXRec.TotTaxAmt = Using$(FF10$, OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt))
      PrnCnt = PrnCnt + 1
      Print #RptHandle, TaxXRec.TaxBillNum; TaxXRec.CustName;
      Print #RptHandle, TaxXRec.Add1; TaxXRec.Add2; TaxXRec.Add3;
      Print #RptHandle, TaxXRec.TaxYear; TaxXRec.CustAcct;
      Print #RptHandle, TaxXRec.MapNum; TaxXRec.PropDesc1;
      Print #RptHandle, TaxXRec.TAXRATE; TaxXRec.LandVal;
      Print #RptHandle, TaxXRec.BldgVal; TaxXRec.RealVal;
      Print #RptHandle, TaxXRec.CurrTaxAmt; TaxXRec.PropDesc2;
      Print #RptHandle, TaxXRec.PropDesc3;
'dale hacked
      Print #RptHandle, TaxXRec.MBLASize;
'      Print #RptHandle, TaxXRec.TotTaxAmt;
'      PastDue = GetCustRealBalance(CLng(TBRec.CustRec), -1)
'      Print #RptHandle, Using$(FF10, PastDue)
      PastDue = GetCustRealBalance(CLng(TBRec.CustRec), -1)
      If PastDue < 0 Then
'        Print #RptHandle, Using$(FF10, OldRound(TaxXRec.TotTaxAmt + Abs(PastDue))); 'Total added past due on 9/19/07
        Print #RptHandle, Using$(FF10, OldRound(TaxXRec.TotTaxAmt - Abs(PastDue))); 'Inserted this line 11/4/2009
      Else
        Print #RptHandle, Using$(FF10, TaxXRec.TotTaxAmt);  'Total added past due on 9/19/07
      End If
      Print #RptHandle, Using$(FF10, PastDue); '+ or - past due
      Print #RptHandle, TaxXRec.TotTaxAmt 'net tax due 'added 9/19/07
    End If
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  
  Close
  
  If PrnCnt > 0 Then
    Call MakeFile 'Makefile creates the .dat identifying that bill printing took place
  End If
  
  Call TaxMsg(800, "The tax billing records have been successfully saved in the file named 'LCRE" + TaxYear + ".TXT' located in the Citipak folder.")
        
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxBillPrinting", "PrintExpReal", Erl)
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

Private Sub PrintExpPers()
  Dim x As Long, y As Integer
  Dim TaxXRec As TaxBillExportPersType
  Dim TXHandle As Integer
  Dim TBRec As VAPPTaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim ReportFile As String
  Dim RptHandle As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim TaxYear$
  Dim FF11$
  Dim Map$
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPRecs As Long
  Dim BillNo&
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim ThisSSN As String * 11
  Dim PersExemp As Double
  Dim NextRec As Long
  Dim ExempAmt As Double
  Dim PersDue As Double
  Dim TotPersVal As Double
  Dim FF9$
  Dim PrnCnt As Long
  Dim PastDue As Double
  Dim PastDueS As String
  
  On Error GoTo ERRORSTUFF
  
  PrnCnt = 0
  FF11$ = "########.#0"
  FF9$ = "######.#0"
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  TaxYear$ = CStr(TaxMasterRec.PTaxYear)
  ReportFile$ = "LCPP" + TaxYear + ".TXT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  BillNo& = fpDblSnglStartPersBill.Value
  OpenPersTaxBillFile TBHandle, NumOfTBRecs
'  OpenRealTaxBillFile TBHandle, NumOfTBRecs
  OpenPersPropFile PHandle, NumOfPRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  For x = 1 To NumOfTBRecs
    Get TBHandle, x, TBRec
    If TBRec.BillNumber > 0 Then
'      If OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt) > 0 Then 'assign and save bill number
      If OldRound(TBRec.TotalBillDue) > 0 Then '5/11/07
        TBRec.BillNumber = BillNo&
        TBRec.BillPrinted = True
        Put TBHandle, x, TBRec
        BillNo& = BillNo& + 1
      End If
      RSet TaxXRec.CustName = QPTrim$(TBRec.CustName)
      RSet TaxXRec.Add1 = QPTrim$(TBRec.CustAdd1)
      RSet TaxXRec.Add2 = QPTrim$(TBRec.CustAdd2)
      Get TCHandle, TBRec.CustRec, TaxCust
      RSet TaxXRec.City = QPTrim$(TaxCust.City)
      RSet TaxXRec.State = QPTrim$(TaxCust.State)
      RSet TaxXRec.Zip = QPTrim$(TaxCust.Zip)
      RSet TaxXRec.CustAcct = Using$("#####0", TBRec.CustRec)
      ThisSSN = QPTrim$(TaxCust.CSSN)
      ThisSSN = ReplaceString(ThisSSN, "-", "")
      Call InsertSSNDashes(ThisSSN)
      RSet TaxXRec.SSN1 = QPTrim$(ThisSSN)
      ThisSSN = QPTrim$(TaxCust.OSSN)
      ThisSSN = ReplaceString(ThisSSN, "-", "")
      Call InsertSSNDashes(ThisSSN)
      RSet TaxXRec.SSN2 = QPTrim$(ThisSSN)
      RSet TaxXRec.DueDate = fptxtPersDueDate.Text
      NextRec = TaxCust.FirstPersRec
      PersExemp = 0
      PastDue = GetCustPersBalance(CLng(TBRec.CustRec), -1)
      RSet PastDueS = Using$(FF11$, PastDue)
      If NextRec = 0 Then
        RSet TaxXRec.LessRelief = Using$(FF11$, 0)
        RSet TaxXRec.NetDue = Using$(FF11$, 0)
        RSet TaxXRec.RepeatDesc = ""
        RSet TaxXRec.RepeatID = ""
        RSet TaxXRec.RepeatValue = Using$(FF11$, 0)
        RSet TaxXRec.RepeatTaxRate = Using$("#.#0", 0)
        RSet TaxXRec.RepeatTotTax = Using$(FF9$, 0)
        RSet TaxXRec.RepeatTaxRelief = Using$(FF9, 0)
        RSet TaxXRec.RepeatTotTax = Using$(FF9$, 0)
      Else
        Get PHandle, NextRec, PersRec
        PersExemp = 0 '6/14/06 no more pers exemptions OldRound(PersExemp + PersRec.EXMPOTHR + PersRec.EXMPSENI)
        ExempAmt = OldRound(TBRec.PersTaxRate * PersExemp)
        PersDue = OldRound(ExempAmt + TBRec.PersTaxDue)
        RSet TaxXRec.LessRelief = Using$(FF11$, ExempAmt)
        RSet TaxXRec.NetDue = Using$(FF11$, OldRound(TBRec.PersTaxDue - TBRec.OverPayAmt))
        RSet TaxXRec.RepeatDesc = QPTrim$(PersRec.DESC1)
        RSet TaxXRec.RepeatID = QPTrim$(PersRec.PropPin)
        TotPersVal = OldRound(PersRec.CVALUE + PersRec.PersVal + PersRec.MCValue + PersRec.MHValue + PersRec.MTValue)
        RSet TaxXRec.RepeatValue = Using$(FF11$, TotPersVal)
        RSet TaxXRec.RepeatTaxRate = Using$("#.#0", TBRec.PersTaxRate)
        RSet TaxXRec.RepeatTotTax = Using$(FF9$, PersDue)
        RSet TaxXRec.RepeatTaxRelief = Using$(FF9, ExempAmt)
        RSet TaxXRec.RepeatTotTax = Using$(FF9$, OldRound(TBRec.PersTaxDue + TBRec.MCTaxDue + TBRec.MHTaxDue + TBRec.FETaxDue + TBRec.MTTaxDue + TBRec.OptRevTax1 + TBRec.OptRevTax2 + TBRec.OptRevTax3 - TBRec.PPTRADiscnt - TBRec.OverPayAmt))
      End If
      PrnCnt = PrnCnt + 1
      Print #RptHandle, TaxXRec.CustName; TaxXRec.Add1; TaxXRec.Add2;
      Print #RptHandle, TaxXRec.City; TaxXRec.State; TaxXRec.Zip;
      Print #RptHandle, TaxXRec.CustAcct; TaxXRec.SSN1; TaxXRec.SSN2;
      Print #RptHandle, TaxXRec.DueDate; TaxXRec.TotDue; TaxXRec.LessRelief;
      Print #RptHandle, TaxXRec.NetDue;
      Print #RptHandle, TaxXRec.RepeatDesc; TaxXRec.RepeatID;
      Print #RptHandle, TaxXRec.RepeatValue; TaxXRec.RepeatTaxRate;
      Print #RptHandle, TaxXRec.RepeatTotTax; TaxXRec.RepeatTaxRelief;
      Print #RptHandle, QPTrim$(TaxXRec.RepeatNetTax);
      Print #RptHandle, PastDueS
    End If
  Next x
  Close
  
  If PrnCnt > 0 Then
    Call MakeFile 'Makefile creates the .dat identifying that bill printing took place
  End If
  
  Call TaxMsg(800, "The tax billing records have been successfully saved in the file named 'LCPP" + TaxYear + ".TXT' located in the Citipak folder.")
        
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxBillPrinting", "PrintExpPers", Erl)
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

Private Sub fpcmbType_Change()
  If fpcmbType.Text = "REAL" Then
    OptMort.Enabled = True
    If Exist("RZipIdx.Dat") Then
      optZip.Value = True
    Else
      optLeave.Value = True
    End If
    Call DisablePersonal
    Call EnableReal
    If fptxtCurrForm.Text = "CDRBLUFF" Then
      fpDSMnthlyPen.Visible = True
      Label28.Visible = True
    End If
  ElseIf fpcmbType.Text = "PERSONAL" Then
    OptMort.Enabled = False
    If Exist("PZipIdx.Dat") Then
      optZip.Value = True
    Else
      optLeave.Value = True
    End If
    Call EnablePersonal
    Call DisableReal
    fpDSMnthlyPen.Visible = False
    Label28.Visible = False
  End If
End Sub

Private Sub fpcmbType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbType.ListIndex = -1
  End If
  If fpcmbType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      If fpcmbType.Text = "REAL" Then
        fptxtComment.SetFocus
      ElseIf fpLongPersTaxYear.Text = "PERSONAL" Then
        fptxtComment.SetFocus
      End If
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub PrintRealVAStandard()
 'checked OK against mask (TAXREMSK.DAT) on 10/21/2005
 'STANDARD REAL ESTATE BILL FORMAT AS SOLD BY SOUTHERN SOFTWARE
 'TAXRESTD.BI
  Dim TaxBill As VARETaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim BillInfo As VARETaxBillInfoType
  Dim BIHandle As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long, RealTaxRate#
  Dim RYear As Integer, RYearStr$
  Dim File$, LC As Integer
  Dim CustName As String * 45, WhatYear As Integer
  Dim RptFile#, WhatReal&
  Dim CarCount As Integer
  Dim TownName$, Add1$, Add2$, Add3$
  Dim RHandle As Integer
  Dim NumOfRRecs As Long
  Dim RealRec As PropertyRecType
  Dim TaxAmt#, LCnt As Integer
  Dim ThisDesc As String * 28
  Dim TotOth As Double
  Dim PrintComments As String
  Dim FF$
  Dim ZipRec As BillPrintRZipIdxType, CustArr As Long
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
  Dim BillNum As Long
'  Dim AHandle As Integer
'
'  AHandle = FreeFile
'  Open "realbill.dat" For Output As AHandle
  FF$ = Chr(12)
  If QPTrim$(fptxtComment.Text) <> "" Or QPTrim$(fptxtComment2.Text) <> "" Then
    If InStr(fpcmbCommentPlace.Text, "LEFT") Then
      PrintComments = "L"
    ElseIf InStr(fpcmbCommentPlace.Text, "RIGHT") Then
      PrintComments = "R"
    Else
      PrintComments = "N"
    End If
  End If
  
  RealTaxRate# = fpDblSnglRealRate
  WhatYear = RealYear
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  TownName = QPTrim$(TaxMasterRec.Name)
  Add1$ = QPTrim$(TaxMasterRec.Add1)
  Add2$ = QPTrim$(TaxMasterRec.Add2)
  Add3$ = QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TownState) + " " + QPTrim$(TaxMasterRec.Zip)
  
  OpenRealPropFile RHandle, NumOfRRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenRealTaxBillFile TBHandle, NumOfTBRecs
  OpenRealBillInfoFile BIHandle
  Get BIHandle, 1, BillInfo
  Close BIHandle
  File$ = StartPath$ + "/TxBStandRE.PRN"
  RptFile# = FreeFile
  Open File$ For Output As #RptFile
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  'Must Calc Late Fee Here
  frmVATaxShowPctComp.Label1 = "Creating Real Bills"
  frmVATaxShowPctComp.Show , Me
  frmVATaxShowPctComp.cmdCancel.Visible = False
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False

  If OptMort.Value = True Then '12/6/06
    OpenMortIdxFile MRHandle, NumOfMRRecs
    NumOfTBRecs = NumOfMRRecs
  ElseIf optZip.Value = True Then '12/6/06
    OpenRZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  BillNum = fpDblSnglStartRealBill.Value
  
  For x = 1 To NumOfTBRecs
    If OptMort.Value = True Then '12/6/06
      Get MRHandle, x, MortRec
      CustArr = MortRec.TaxBillRec
    ElseIf optZip.Value = True Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TaxBill
    Get TCHandle, TaxBill.CustRec, TaxCust
    CustName$ = QPTrim$(TaxCust.CustName)
    If TaxBill.BillNumber < 0 Then GoTo NotThisOne
'      Print #AHandle, CStr(TaxBill.CustRec) + "~" + Using("###,###,##0.00", TaxBill.TotalBillDue)
      Print #RptFile, "~"
      Print #RptFile, Tab(64); "TAX YEAR: "; WhatYear
      Print #RptFile, Tab(75); Using$("#####", TaxBill.BillNumber)
      Print #RptFile, " "
      Print #RptFile, " "
      Print #RptFile, " " 'added
      Print #RptFile, Tab(5); TownName$
      Print #RptFile, Tab(5); Add1$
      Print #RptFile, Tab(5); Add2$
      Print #RptFile, Tab(5); Add3$
      Print #RptFile, " "
      Print #RptFile, " "
      Print #RptFile, " " 'added
      Print #RptFile, " "
      Print #RptFile, Tab(5); "PIN:  " + QPTrim$(TaxBill.RealPin)
      Print #RptFile, Tab(5); "ACCT: " + Using$("#####", TaxBill.CustRec)
      Print #RptFile, Tab(5); CustName$
      Print #RptFile, Tab(5); Left$(TaxBill.CustAdd1, 35)
      Print #RptFile, Tab(5); Left$(TaxBill.CustAdd2, 35)
      Print #RptFile, Tab(5); QPTrim$(TaxBill.CustAdd3) + " " + InsertZipDash(TaxBill.CustZip)

      For LC = 19 To 20 'made 18 = 19
        Print #RptFile, " "
      Next LC
      Print #RptFile, Tab(41); "LAND"; Tab(51); "BUILDING"; Tab(61); "NET TOTAL"; Tab(72); "TOTAL DUE"
      Print #RptFile, " "
      'Line 23 Starts Here
      ThisDesc = QPTrim$(TaxBill.RDesc1)
      Print #RptFile, ThisDesc; 'QPTrim$(TaxBill.RDesc1);
      Print #RptFile, Tab(30); Using("#0.00", RealTaxRate#);
      If TaxBill.RealValue > TaxBill.ExptValue Then
        Print #RptFile, Tab(37); Using("######0.00", (TaxBill.RealValue - TaxBill.ExptValue)); ' - RTaxBill.PersValue));
        Print #RptFile, Tab(50); Using("#####0.00", TaxBill.BldgValue);
      ElseIf TaxBill.BldgValue > TaxBill.ExptValue Then
        Print #RptFile, Tab(37); Using("######0.00", TaxBill.RealValue); ' - RTaxBill.PersValue));
        Print #RptFile, Tab(50); Using("#####0.00", (TaxBill.BldgValue - TaxBill.ExptValue));
      ElseIf TaxBill.RealValue + TaxBill.BldgValue > TaxBill.ExptValue Then
        Print #RptFile, Tab(37); Using("######0.00", TaxBill.RealValue - (TaxBill.ExptValue * (TaxBill.RealValue / (TaxBill.RealValue + TaxBill.BldgValue)))); ' - RTaxBill.PersValue));
        Print #RptFile, Tab(50); Using("#####0.00", TaxBill.BldgValue - (TaxBill.ExptValue * (TaxBill.BldgValue / (TaxBill.RealValue + TaxBill.BldgValue)))); ' - RTaxBill.PersValue));
      Else
        Print #RptFile, Tab(37); Using("######0.00", TaxBill.RealValue); ' - RTaxBill.PersValue));
        Print #RptFile, Tab(50); Using("#####0.00", TaxBill.BldgValue);
      End If
      Print #RptFile, Tab(61); Using("#####0.00", OldRound(TaxBill.RealValue + TaxBill.BldgValue - TaxBill.ExptValue));
      Print #RptFile, Tab(71); Using("######0.00", OldRound(TaxBill.TotalBillDue)) ' - TaxBill.OverPayAmt))
      Print #RptFile, QPTrim$(TaxBill.RDesc2)
      TotOth = OldRound(TaxBill.OptRevTax1 + TaxBill.OptRevTax2 + TaxBill.OptRevTax3 + TaxBill.LateTaxDue)
      If TaxBill.OverPayAmt > 0 And TotOth > 0 Then
        Print #RptFile, "** Applied Credit Of " + QPTrim$(Using$("$##,##0.00", TaxBill.OverPayAmt)) + " **"; Tab(40); "Other Taxes Included: " + Using$("$#,##0.00", TotOth)
        For LCnt = 26 To 36: Print #RptFile, "": Next LCnt
      ElseIf TaxBill.OverPayAmt > 0 And TotOth = 0 Then
        Print #RptFile, "** Applied Credit Of " + QPTrim$(Using$("$##,##0.00", TaxBill.OverPayAmt)) + " **"
        For LCnt = 26 To 36: Print #RptFile, "": Next LCnt
      ElseIf TaxBill.OverPayAmt = 0 And TotOth > 0 Then
        Print #RptFile, Tab(40); "Other Taxes Included: " + Using$("$#,##0.00", TotOth)
        For LCnt = 26 To 36: Print #RptFile, "": Next LCnt
      Else
        For LCnt = 25 To 36: Print #RptFile, "": Next LCnt
      End If
      'Lines 25 to 36 are blank
     'Line 37 for Totals
      If PrintComments = "L" Then
        Print #RptFile, fptxtComment.Text; Tab(48); "Total Tax Due ";
        Print #RptFile, Using$("$#######0.00", OldRound(TaxBill.TotalBillDue - TaxBill.OverPayAmt))
        Print #RptFile, fptxtComment2.Text; Tab(48); "Tax Due Date: " + MakeRegDate(BillInfo.DueDate)
        Print #RptFile,
        Print #RptFile,
      ElseIf PrintComments = "R" Then
        Print #RptFile, Tab(48); "Total Tax Due ";
        Print #RptFile, Using$("$#######0.00", OldRound(TaxBill.TotalBillDue - TaxBill.OverPayAmt))
        Print #RptFile, Tab(48); "Tax Due Date: " + MakeRegDate(BillInfo.DueDate)
        Print #RptFile, Tab(48); fptxtComment.Text
        Print #RptFile, Tab(48); fptxtComment2.Text
      Else
        Print #RptFile, Tab(48); "Total Tax Due ";
        Print #RptFile, Using$("$#######0.00", OldRound(TaxBill.TotalBillDue - TaxBill.OverPayAmt))
        Print #RptFile, Tab(48); "Tax Due Date: " + MakeRegDate(BillInfo.DueDate)
        Print #RptFile,
        Print #RptFile,
      End If
      Print #RptFile, "BN"; Using$("#####", x) 'PrnCnt)
      Print #RptFile, "~"
      TaxBill.BillPrinted = True
      TaxBill.Comment = fptxtComment.Text
      TaxBill.Comment2 = fptxtComment2.Text
      TaxBill.CommentPlace = fpcmbCommentPlace.Text
      TaxBill.BillNumber = BillNum
      BillNum = BillNum + 1
      Put TBHandle, CustArr, TaxBill

NotThisOne:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  Print #RptFile, FF$
  
  Close
  Call MakeFile
  ViewPrint File$, "Real Property Tax Bills", True
  
End Sub

Private Sub PrintPersVAStandard()
  'checked OK against mask (taxppmsk.dat) on 10/21/2005
 'TAXPPSTD.BI
  Dim TaxBill As VAPPTaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim BillInfo As VAPPTaxBillInfoType
  Dim BIHandle As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  Dim PYear As Integer, PYearStr$
  Dim File$, LC As Integer
  Dim CustName$, WhatYear As Integer
  Dim RptFile#, WhatPers&
  Dim CarCount As Integer
  Dim TownName$, Add1$, Add2$, Add3$
  Dim PHandle As Integer, PPTRAVal#
  Dim NumOfPRecs As Long, PPTRADiscount#
  Dim PersRec As PersonalRecType
  Dim VehDesc$, PERC!
  Dim TaxAmt#, LCnt As Integer
  Dim MultiYear As Integer
  Dim TotOth As Double
  Dim PrintComments As String
  Dim FF$
  Dim Zip$, BillNum&
  Dim ZipRec As BillPrintRZipIdxType, CustArr As Long
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  
  FF$ = Chr(12)
  If QPTrim$(fptxtComment.Text) <> "" Or QPTrim$(fptxtComment2.Text) <> "" Then
    If InStr(fpcmbCommentPlace.Text, "LEFT") Then
      PrintComments = "L"
    ElseIf InStr(fpcmbCommentPlace.Text, "RIGHT") Then
      PrintComments = "R"
    Else
      PrintComments = "N"
    End If
  End If
  
  WhatYear = PersYear
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  MultiYear = TaxMasterRec.MultiYear
  TaxMasterRec.MaxVehTaxVal = OldRound(TaxMasterRec.MaxVehTaxVal)
  TownName = QPTrim$(TaxMasterRec.Name)
  Add1$ = QPTrim$(TaxMasterRec.Add1)
  Add2$ = QPTrim$(TaxMasterRec.Add2)
  Add3$ = QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TownState) + " " + QPTrim$(TaxMasterRec.Zip)
  
  If WhatYear = 1999 Then PERC! = 27.5
  If WhatYear = 2000 Then PERC! = 47.5
  If WhatYear >= 2001 Then PERC! = TaxMasterRec.PPTRADisc
  
  OpenPersPropFile PHandle, NumOfPRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenPersTaxBillFile TBHandle, NumOfTBRecs
  OpenPersBillInfoFile BIHandle
  Get BIHandle, 1, BillInfo
  Close BIHandle
  File$ = StartPath$ + "/TxBStandPP.PRN"
  RptFile# = FreeFile
  Open File$ For Output As #RptFile
  frmVATaxShowPctComp.Label1 = "Creating Personal Bills"
  frmVATaxShowPctComp.Show , Me
  frmVATaxShowPctComp.cmdCancel.Visible = False
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  
  If Exist("PZipIdx.Dat") Then '12/6/06
    OpenPZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  
  BillNum = fpDblSnglStartPersBill.Value
  For x = 1 To NumOfTBRecs
    If optZip.Value = True Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TaxBill
    If TaxBill.BillNumber < 0 Then GoTo Natta
    Get TCHandle, TaxBill.CustRec, TaxCust
    GoSub PrintIt
    TaxBill.BillNumber = BillNum
    TaxBill.BillPrinted = True
    TaxBill.Comment = fptxtComment.Text
    TaxBill.Comment2 = fptxtComment2.Text
    TaxBill.CommentPlace = fpcmbCommentPlace.Text
    Put TBHandle, CustArr, TaxBill
    BillNum = BillNum + 1
Natta:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  Print #RptFile, FF$
  Close
  
  If BillNum > 0 Then Call MakeFile 'added 8/21/07
  ViewPrint File$, "Personal Property Tax Bills", True
  
  Exit Sub
  
PrintIt:
  CustName$ = QPTrim$(TaxCust.CustName)
  Zip = InsertZipDash(TaxBill.CustZip)
  Print #RptFile, "~"
  Print #RptFile, Tab(63); "TAX YEAR: "; WhatYear
  Print #RptFile, Tab(75); Using$("#####", TaxBill.BillNumber)
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, Tab(5); TownName$
  Print #RptFile, Tab(5); Add1$
  Print #RptFile, Tab(5); Add2$
  Print #RptFile, Tab(5); Add3$
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, Tab(5); "Acct # "; Using$("#####0", TaxBill.CustRec)
  Print #RptFile, Tab(5); CustName$
  Print #RptFile, Tab(5); QPTrim$(TaxBill.CustAdd1)
  Print #RptFile, Tab(5); QPTrim$(TaxBill.CustAdd2)
  Print #RptFile, Tab(5); QPTrim$(TaxBill.CustAdd3) + " " + Zip ' QPTrim$(TaxBill.CustZip)
  For LC = 18 To 21
'  For LC = 19 To 21 'added
   Print #RptFile, " "
  Next LC
  Print #RptFile, Tab(37); "PROP VALUE"; Tab(50); "TAX AMOUNT"; Tab(62); "PPTRA DIS"; Tab(72); "TOTAL DUE"
 'Line 24 Starts Here
  Print #RptFile, " "
  Print #RptFile, "Personal Property"; Tab(32); Using$("#.00", PersTaxRate#);
   Print #RptFile, Tab(37); Using$("#####0.00", TaxBill.PersValue);
   Print #RptFile, Tab(51); Using$("#####0.00", TaxBill.PersTaxDue); ' - TaxBill.OverPayAmt);
   Print #RptFile, Tab(63); Using$("####0.00", TaxBill.PPTRADiscnt);
'   If InStr(TaxMasterRec.Name, "CHILHOWIE") Then '11/06/06
'     Print #RptFile, Tab(72); Using$("#####0.00", OldRound(TaxBill.PersTaxNet)) ' + TaxBill.ChillHowieFudge)) ' - TaxBill.OverPayAmt))
'   Else
     Print #RptFile, Tab(72); Using$("#####0.00", OldRound(TaxBill.PersTaxDue - TaxBill.PPTRADiscnt)) ' - TaxBill.OverPayAmt))
'   End If
   Print #RptFile, "Machinery/Tools"; Tab(32); Using$("#.00", MTTaxRate#);
   Print #RptFile, Tab(37); Using$("#####0.00", TaxBill.MTValue);
   Print #RptFile, Tab(51); Using$("#####0.00", TaxBill.MTTaxDue);
   Print #RptFile, Tab(72); Using$("#####0.00", TaxBill.MTTaxDue)
  Print #RptFile, "Farm Equipment";
   Print #RptFile, Tab(32); Using("#.00", FETaxRate#);
   Print #RptFile, Tab(37); Using$("#####0.00", TaxBill.FEValue);
   Print #RptFile, Tab(51); Using$("#####0.00", TaxBill.FETaxDue);
   Print #RptFile, Tab(72); Using$("#####0.00", TaxBill.FETaxDue)
  Print #RptFile, "Mobile Homes";
   Print #RptFile, Tab(32); Using$("#.00", MHTaxRate#);
   Print #RptFile, Tab(37); Using$("#####0.00", TaxBill.MHValue);
   Print #RptFile, Tab(51); Using$("#####0.00", TaxBill.MHTaxDue);
   Print #RptFile, Tab(72); Using$("#####0.00", TaxBill.MHTaxDue)
  Print #RptFile, "Merchant Capital";
   Print #RptFile, Tab(32); Using$("#.00", MCTaxRate#);
   Print #RptFile, Tab(37); Using$("#####0.00", TaxBill.MCValue);
   Print #RptFile, Tab(51); Using$("#####0.00", TaxBill.MCTaxDue);
   Print #RptFile, Tab(72); Using$("#####0.00", TaxBill.MCTaxDue)
   TotOth = OldRound(TaxBill.OptRevTax1 + TaxBill.OptRevTax2 + TaxBill.OptRevTax3)
   If TaxBill.OverPayAmt > 0 And TotOth = 0 Then '6/22/06
     Print #RptFile, " PPTRA Vehicle Information"; Tab(43); "** Applied Credit Of: " + QPTrim$(Using$("$##,##0.00", TaxBill.OverPayAmt)) + " **"
   ElseIf TaxBill.OverPayAmt > 0 And TotOth > 0 Then '6/22/06
     Print #RptFile, " PPTRA Vehicle Information"; Tab(30); "* Applied Credit Of: " + QPTrim$(Using$("$##,##0.00", TaxBill.OverPayAmt)) + " *"; Tab(60); "Other Taxes: " + Using$("#,##0.00", TotOth)
   ElseIf TaxBill.OverPayAmt = 0 And TotOth > 0 Then '6/22/06
     Print #RptFile, " PPTRA Vehicle Information"; Tab(60); "Other Taxes: " + Using$("#,##0.00", TotOth)
   Else
     Print #RptFile, " PPTRA Vehicle Information"
   End If
 'Line 30 to 35 here to print vehicles
  CarCount = 0
  WhatPers& = TaxCust.FirstPersRec
  Do
    Get PHandle, WhatPers&, PersRec
    PYearStr$ = CStr(PersRec.TaxBillYear)
    PYear = Val(PYearStr$)
    If PYear > 0 And PYear <> WhatYear And (PersRec.PersVal > 0 Or PersRec.CVALUE > 0 Or PersRec.MCValue > 0 Or PersRec.MHValue Or PersRec.MTValue > 0) Then
'    If PYear > 0 And PYear <> WhatYear Then
'      Return
        'Do Not Process This Record
    End If
    If PersRec.PPTRAYN = "Y" Then
      If OldRound#(PersRec.PersVal) > TaxMasterRec.MaxVehTaxVal Then
        If TaxMasterRec.MultiYear <> 0 Then
          PersRec.PersVal = OldRound(PersRec.PersVal)
        End If
        PPTRAVal# = TaxMasterRec.MaxVehTaxVal
      Else
        PPTRAVal# = OldRound#(PersRec.PersVal)
      End If

      If PPTRAVal# <= TaxMasterRec.MinVehTaxVal Then
        PPTRADiscount# = OldRound#((OldRound#(PPTRAVal# / 100) * PersTaxRate#))
      Else
        PPTRADiscount# = OldRound#((OldRound#((PPTRAVal# / 100) * (PERC! / 100)) * PersTaxRate#))
      End If

      VehDesc$ = " VIN# " + QPTrim$(PersRec.Vin)
      VehDesc$ = QPTrim$(VehDesc$)
      TaxAmt# = OldRound((PersTaxRate# / 100) * PersRec.PersVal)
      TaxBill.PersTaxDue = TaxBill.PersTaxDue
      TaxBill.PPTRADiscnt = TaxBill.PPTRADiscnt
      Print #RptFile, "*" + VehDesc$;
      Print #RptFile, Tab(37); Using$("#####0.00", PersRec.PersVal) ';
      CarCount = CarCount + 1
    End If
    
    If CarCount >= 6 Then
      Print #RptFile, ""
      Print #RptFile, Tab(48); "Total Tax Due ";
      Print #RptFile, Using$("$#######0.00", OldRound(TaxBill.TotalBillDue - TaxBill.OverPayAmt))
      Print #RptFile, Tab(48); "Tax Due Date: " + MakeRegDate(BillInfo.DueDate)  'DueDate$
      Print #RptFile, ""
      Print #RptFile,
      Print #RptFile,
      Print #RptFile,
      Print #RptFile, "BN"; Using("####0", x) ' PrnCnt
      Print #RptFile, "~"

      Print #RptFile, "~"
      Print #RptFile, Tab(62); "TAX YEAR: "; WhatYear
      Print #RptFile, Tab(75); Using$("####0", TaxBill.BillNumber)
      Print #RptFile, " "
      Print #RptFile, " "
      Print #RptFile, Tab(5); TownName$
      Print #RptFile, Tab(5); Add1$
      Print #RptFile, Tab(5); Add2$
      Print #RptFile, Tab(5); Add3$
      Print #RptFile, " "
      Print #RptFile, " "
      Print #RptFile, " "
      Print #RptFile, Tab(5); "Acct # " + Using$("####0", TaxBill.CustRec) + " Vehicle Listing Cont'd"
      Print #RptFile, Tab(5); CustName$
      Print #RptFile, Tab(5); QPTrim$(TaxBill.CustAdd1)
      Print #RptFile, Tab(5); QPTrim$(TaxBill.CustAdd2)
      Print #RptFile, Tab(5); QPTrim$(TaxBill.CustAdd3) + " " + Zip$ 'QPTrim$(TaxBill.CustZip)
      For LC = 18 To 21
       Print #RptFile, " "
      Next LC
      Print #RptFile, Tab(37); "PROP VALUE"; Tab(50); "TAX AMOUNT"; Tab(62); "PPTRA DIS."; Tab(72); "TOTAL DUE"
      Print #RptFile, " "
      Print #RptFile, "Vehicle Listing Continued ..."
      Print #RptFile, ""
      Print #RptFile, ""
      Print #RptFile, " PPTRA Vehicle Information"
      Print #RptFile, ""
      Print #RptFile, ""
      CarCount = 0
    End If
    WhatPers& = PersRec.NextRec
  Loop While WhatPers& > 0

   ' Finish the bill up here
  If CarCount < 6 Then
    For LCnt = CarCount To 6: Print #RptFile, "": Next LCnt
  End If
'  Print #RptFile,'10/24
'  Print #RptFile,
  If PrintComments = "L" Then 'comments left
    Print #RptFile, fptxtComment.Text; Tab(48); "Total Tax Due ";
    Print #RptFile, Using$("$#######0.00", OldRound(TaxBill.TotalBillDue - TaxBill.OverPayAmt))
    Print #RptFile, fptxtComment2.Text; Tab(48); "Tax Due Date: " + MakeRegDate(BillInfo.DueDate)
    Print #RptFile,
    Print #RptFile,
  ElseIf PrintComments = "R" Then 'comments right
    Print #RptFile, Tab(48); "Total Tax Due ";
    Print #RptFile, Using$("$#######0.00", OldRound(TaxBill.TotalBillDue - TaxBill.OverPayAmt))
    Print #RptFile, Tab(48); "Tax Due Date: " + MakeRegDate(BillInfo.DueDate)
    Print #RptFile, Tab(48); fptxtComment.Text
    Print #RptFile, Tab(48); fptxtComment2.Text
  Else
    Print #RptFile, Tab(48); "Total Tax Due ";
    Print #RptFile, Using$("$#######0.00", OldRound(TaxBill.TotalBillDue - TaxBill.OverPayAmt))
    Print #RptFile, Tab(48); "Tax Due Date: " + MakeRegDate(BillInfo.DueDate)
    Print #RptFile, ""
    Print #RptFile, ""
  End If
  Print #RptFile, ""
  Print #RptFile, "BN"; Using$("####0", x)
  Print #RptFile, "~"
  
  Call MakeFile
  Return
End Sub

Private Sub PrintLaserReal1()
  Dim ToPrint As String
  Dim TaxRptT As Integer
  Dim ReportFile As String
  Dim RptHandle As Integer
  Dim TBRec As VARETaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long
  Dim dlm$, BillNo&, PrnCnt As Long
  Dim TBDRec As TxBillLaser1DefaultsType
  Dim TBDHandle As Integer
  Dim ThisRate As Double
  Dim BZip$
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim CustArr As Long '12/6/06
  Dim ZipRec As BillPrintPZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
'  Dim AHandle As Integer
  
  On Error GoTo ERRORSTUFF
  
'  AHandle = FreeFile
'  Open "realbill.dat" For Output As AHandle
  dlm$ = "~"
  ReportFile$ = StartPath$ + "/TaxRBil1.PRN"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  BillNo& = fpDblSnglStartRealBill.Value
  
  frmVATaxShowPctComp.Label1 = "Printing Real Tax Bills"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  EnableCloseButton Me.hwnd, False
  
  OpenTxBillRealFile TBDHandle
  Get #TBDHandle, 1, TBDRec
  If fpcmbBarCode.Text = "No" Then
    TBDRec.UseBarCode = False
    Put #TBDHandle, 1, TBDRec
  ElseIf fpcmbBarCode.Text = "Yes" Then
    TBDRec.UseBarCode = True
    Put #TBDHandle, 1, TBDRec
  End If
  Close TBDHandle
  
  ARptVATempTaxBill.Head1 = QPTrim(TBDRec.TxtHead1)
  ARptVATempTaxBill.Head2 = QPTrim(TBDRec.TxtHead2)
  ARptVATempTaxBill.LblOpt1 = QPTrim(TBDRec.txtOpt1)
  ARptVATempTaxBill.LblOpt2 = QPTrim(TBDRec.TxtOpt2)
  ARptVATempTaxBill.LblOpt3 = QPTrim(TBDRec.TxtOpt3)
  ARptVATempTaxBill.LblOpt4 = QPTrim(TBDRec.TxtOpt4)
  ARptVATempTaxBill.LblPgph1 = QPTrim(TBDRec.txtPgph0)
  ARptVATempTaxBill.LblPgph2 = QPTrim(TBDRec.txtPgph1)
  ARptVATempTaxBill.LblPgph3 = QPTrim(TBDRec.txtPgph2)
  ARptVATempTaxBill.LblPgph4 = QPTrim(TBDRec.txtPgph3)
  ARptVATempTaxBill.LblPgph5 = QPTrim(TBDRec.txtPgph4)
  ARptVATempTaxBill.LblPgph6 = QPTrim(TBDRec.txtPgph5)
  ARptVATempTaxBill.LblPgph7 = QPTrim(TBDRec.txtPgph6)
  ARptVATempTaxBill.LblPgph8 = QPTrim(TBDRec.txtPgph7)
  ARptVATempTaxBill.LblOpt5 = QPTrim(TBDRec.TxtOpt5)
  ARptVATempTaxBill.LblHead4 = QPTrim(TBDRec.txtHead4)
  ARptVATempTaxBill.LblHead5 = QPTrim(TBDRec.txtHead5)
  ARptVATempTaxBill.LblHead6 = QPTrim(TBDRec.txtHead6)
  ARptVATempTaxBill.LblOpt6 = QPTrim(TBDRec.TxtOpt6)
  ARptVATempTaxBill.LblOpt7 = QPTrim(TBDRec.TxtOpt7)
  If TBDRec.dologo = 1 Then
    If Exist("towntaxlogo.bmp") Then
      ARptVATempTaxBill.Image1.Picture = LoadPicture("towntaxlogo.bmp")
      ARptVATempTaxBill.Image1.Visible = True
    End If
  End If
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenRealTaxBillFile TBHandle, NumOfTBRecs
  
  If OptMort.Value = True Then '12/6/06
    OpenMortIdxFile MRHandle, NumOfMRRecs
    NumOfTBRecs = NumOfMRRecs
  ElseIf optZip.Value = True Then '12/6/06
    OpenRZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  
  For x = 1 To NumOfTBRecs
    If OptMort.Value = True Then '12/6/06
      Get MRHandle, x, MortRec
      CustArr = MortRec.TaxBillRec
    ElseIf optZip.Value = True Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    
    Get TBHandle, CustArr, TBRec
    If fpcmbPrintPriorYN.Text = "N" Then '5/22/07
      TBRec.PrintPrior = False
    Else
      TBRec.PrintPrior = True
    End If
    If TBRec.BillNumber > 0 Then
'      If OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt) > 0 Then
      If OldRound(TBRec.TotalBillDue) > 0 Then '5/11/07
        TBRec.BillNumber = BillNo&
        TBRec.BillPrinted = True
        Put TBHandle, CustArr, TBRec
        GoSub GetBarCodeData
'        Print #AHandle, CStr(TBRec.CustRec) + "~" + Using("###,###,##0.00", TBRec.TotalBillDue)
        '                         0                         1
        Print #RptHandle, TBRec.BillNumber; dlm; QPTrim$(TBRec.CustName); dlm;
        '                           2                           3
        Print #RptHandle, QPTrim$(TBRec.CustAdd1); dlm; QPTrim$(TBRec.CustAdd2); dlm;
        '                             4                       5
        Print #RptHandle, QPTrim$(TBRec.CustAdd3); dlm; TBRec.CustPin; dlm;
        '                             6                            7
        Print #RptHandle, QPTrim$(TBRec.RealPin); dlm; QPTrim$(TBRec.RDesc1); dlm;
        '                        8                     9                    10
        Print #RptHandle, TBRec.RealValue; dlm; TBRec.BldgValue; dlm; TBRec.ExptValue; dlm;
        If OldRound(TBRec.RealTaxDue - TBRec.OverPayAmt) > 0 Then
          ThisRate = TBRec.RealTaxRate
        Else
          ThisRate = 0
        End If
        '                                         11                                             12
        Print #RptHandle, OldRound(TBRec.RealValue + TBRec.BldgValue - TBRec.ExptValue); dlm; ThisRate; dlm;
        '                                        13                              14                 15                       16
        Print #RptHandle, OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt); dlm; BZip; dlm; QPTrim$(TBRec.CustZip); dlm; TBDRec.dologo; dlm;
        '                        17                     18                     19
        Print #RptHandle, TBRec.OptRevTax1; dlm; TBRec.OptRevTax2; dlm; TBRec.OptRevTax3; dlm;
        '                       20                   21                  22                    23                     24                         25                    26
        Print #RptHandle, TBRec.Opt1Desc; dlm; TBRec.Opt2Desc; dlm; TBRec.Opt3Desc; dlm; TBRec.OverPayAmt; dlm; TBRec.LateTaxDue; dlm; TBRec.PriorYrBalance; dlm; TBRec.PrintPrior
        
        BillNo& = BillNo& + 1
        PrnCnt = PrnCnt + 1
      End If
    End If
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  EnableCloseButton Me.hwnd, True
  
  Close
  ARptVATempTaxBill.GetName ReportFile$
  ARptVATempTaxBill.Show
  If PrnCnt > 0 Then
    Call MakeFile 'Makefile creates the .dat identifying that bill printing took place
  End If
  
  Exit Sub
  
GetBarCodeData:
  If fpcmbBarCode.Text = "No" Then
    BZip = ""
    Return
  End If
  Get TCHandle, TBRec.CustPin, TaxCust
  If Len(QPTrim$(TaxCust.Zip)) < 10 Or Len(QPTrim$(TaxCust.DeliveryPt)) <> 2 Then
    BZip = ""
  Else
    BZip = QPTrim$(TaxCust.Zip) + QPTrim$(TaxCust.DeliveryPt)
  End If
  
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxBillPrinting", "PrintLaser1", Erl)
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

Private Sub PrintLaserPers1()
  Dim ToPrint As String
  Dim TaxRptT As Integer
  Dim ReportFile As String
  Dim RptHandle As Integer
  Dim TBRec As VAPPTaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long
  Dim dlm$, BillNo&, PrnCnt As Long
  Dim TBDRec As TxBillLaser1DefaultsType
  Dim TBDHandle As Integer
  Dim ThisRate As Double
  Dim TotValue As Double
  Dim ThisOpt1Desc As String * 15
  Dim ThisOpt2Desc As String * 15
  Dim ThisOpt3Desc As String * 15
  Dim BZip As String
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim CustArr As Long '12/6/06
  Dim ZipRec As BillPrintPZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
  Dim PriorBal# '5/22/07
  Dim PrintPriorYN As Boolean '5/22/07
  
  On Error GoTo ERRORSTUFF
  
  dlm$ = "~"
  ReportFile$ = StartPath$ + "/TaxPBil1.PRN"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  BillNo& = fpDblSnglStartPersBill.Value
  
  frmVATaxShowPctComp.Label1 = "Printing Personal Tax Bills"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  EnableCloseButton Me.hwnd, False
  
  OpenTxBillPersFile TBDHandle
  Get #TBDHandle, 1, TBDRec
  If fpcmbBarCode.Text = "No" Then
    TBDRec.UseBarCode = False
    Put #TBDHandle, 1, TBDRec
  ElseIf fpcmbBarCode.Text = "Yes" Then
    TBDRec.UseBarCode = True
    Put #TBDHandle, 1, TBDRec
  End If
  Close TBDHandle
  
  arVATaxBillPersLaser.Head1 = QPTrim(TBDRec.TxtHead1)
  arVATaxBillPersLaser.Head2 = QPTrim(TBDRec.TxtHead2)
  arVATaxBillPersLaser.LblOpt1 = QPTrim(TBDRec.txtOpt1)
  arVATaxBillPersLaser.LblOpt2 = QPTrim(TBDRec.TxtOpt2)
  arVATaxBillPersLaser.LblOpt3 = QPTrim(TBDRec.TxtOpt3)
  arVATaxBillPersLaser.LblOpt4 = QPTrim(TBDRec.TxtOpt4)
  arVATaxBillPersLaser.LblPgph1 = QPTrim(TBDRec.txtPgph0)
  arVATaxBillPersLaser.LblPgph2 = QPTrim(TBDRec.txtPgph1)
  arVATaxBillPersLaser.LblPgph3 = QPTrim(TBDRec.txtPgph2)
  arVATaxBillPersLaser.LblPgph4 = QPTrim(TBDRec.txtPgph3)
  arVATaxBillPersLaser.LblPgph5 = QPTrim(TBDRec.txtPgph4)
  arVATaxBillPersLaser.LblPgph6 = QPTrim(TBDRec.txtPgph5)
  arVATaxBillPersLaser.LblPgph7 = QPTrim(TBDRec.txtPgph6)
  arVATaxBillPersLaser.LblPgph8 = QPTrim(TBDRec.txtPgph7)
  arVATaxBillPersLaser.LblOpt5 = QPTrim(TBDRec.TxtOpt5)
  arVATaxBillPersLaser.LblHead4 = QPTrim(TBDRec.txtHead4)
  arVATaxBillPersLaser.LblHead5 = QPTrim(TBDRec.txtHead5)
  arVATaxBillPersLaser.LblHead6 = QPTrim(TBDRec.txtHead6)
  arVATaxBillPersLaser.LblOpt6 = QPTrim(TBDRec.TxtOpt6)
  arVATaxBillPersLaser.LblOpt7 = QPTrim(TBDRec.TxtOpt7)
  If TBDRec.dologo = 1 Then
    If Exist("towntaxlogo.bmp") Then
      arVATaxBillPersLaser.Image1.Picture = LoadPicture("towntaxlogo.bmp")
      arVATaxBillPersLaser.Image1.Visible = True
    End If
  End If
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenPersTaxBillFile TBHandle, NumOfTBRecs
  
  If Exist("PZipIdx.Dat") Then '12/6/06
    OpenPZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  
  For x = 1 To NumOfTBRecs
    If optZip.Value = True Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TBRec
    If fpcmbPrintPriorYN.Text = "N" Then '5/22/07
      TBRec.PrintPrior = False
    Else
      TBRec.PrintPrior = True
    End If
    If TBRec.BillNumber > 0 Then
'      If OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt) > 0 Then
      If OldRound(TBRec.TotalBillDue) > 0 Then '5/11/07
        TBRec.BillNumber = BillNo&
        TBRec.BillPrinted = True
        TBRec.Comment = QPTrim$(fptxtComment.Text)
        Put TBHandle, CustArr, TBRec
        If TBRec.PPTRAValue > 0 Then
          TotValue = OldRound(TBRec.PPTRAValue + TBRec.FEValue + TBRec.MCValue + TBRec.MHValue + TBRec.MTValue - (TBRec.ExptValue)) ' + TBRec.PPTRAValue))
        Else
          TotValue = OldRound(TBRec.PersValue + TBRec.FEValue + TBRec.MCValue + TBRec.MHValue + TBRec.MTValue - (TBRec.ExptValue)) ' + TBRec.PPTRAValue))
        End If
        GoSub GetBarCodeData
        '                         0                         1
        Print #RptHandle, TBRec.BillNumber; dlm; QPTrim$(TBRec.CustName); dlm;
        '                           2                           3
        Print #RptHandle, QPTrim$(TBRec.CustAdd1); dlm; QPTrim$(TBRec.CustAdd2); dlm;
        '                             4                      5
        Print #RptHandle, QPTrim$(TBRec.CustAdd3); dlm; TBRec.CustPin; dlm;
        '                     6               7
        Print #RptHandle, TotValue; dlm; TBRec.RDesc1; dlm;
        '                        8                   9                    10
        Print #RptHandle, TBRec.PersValue; dlm; TBRec.FEValue; dlm; TBRec.ExptValue; dlm;
        '                    11                      12                        13
        Print #RptHandle, TBRec.PPTRAValue; dlm; TBRec.PPTRADiscnt; dlm; OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt); dlm;
        '                       14                 15                   16
        Print #RptHandle, TBDRec.dologo; dlm; TBRec.MHValue; dlm; TBRec.MCValue; dlm;
        '                      17                     18                             19
        Print #RptHandle, TBRec.MTValue; dlm; OldRound(TBRec.PersTaxDue); dlm; TBRec.PersTaxNet; dlm;
        '                         20                    21                    22
        Print #RptHandle, TBRec.PersTaxRate; dlm; TBRec.FETaxDue; dlm; TBRec.FETaxRate; dlm;
        '                       23                    24                    25
        Print #RptHandle, TBRec.MCTaxDue; dlm; TBRec.MCTaxRate; dlm; TBRec.MHTaxDue; dlm;
        '                       26                    27                     28
        Print #RptHandle, TBRec.MHTaxRate; dlm; TBRec.MTTaxDue; dlm; TBRec.MTTaxRate; dlm;
        '                        29                     30                     31
        Print #RptHandle, TBRec.OptRevTax1; dlm; TBRec.OptRevTax2; dlm; TBRec.OptRevTax3; dlm;
        ThisOpt1Desc = QPTrim$(TBRec.OptRevDesc1)
        ThisOpt2Desc = QPTrim$(TBRec.OptRevDesc2)
        ThisOpt3Desc = QPTrim$(TBRec.OptRevDesc3)
        '                        32              33                 34              35             36                  37                   38
        Print #RptHandle, ThisOpt1Desc; dlm; ThisOpt2Desc; dlm; ThisOpt3Desc; dlm; BZip; dlm; TBRec.CustZip; dlm; TBDRec.dologo; dlm; TBRec.OverPayAmt; dlm;
        '                        39                        40
        Print #RptHandle, TBRec.PriorYrBalance; dlm; TBRec.PrintPrior
        
        BillNo& = BillNo& + 1
        PrnCnt = PrnCnt + 1
      End If
    End If
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  EnableCloseButton Me.hwnd, True
  
'  For x = 1 To 25
'    Get TBHandle, x, TBRec
'    If TBRec.BillPrinted = True Then
'      TBRec.BillNumber = TBRec.BillNumber
'    End If
'  Next x
  Close
  arVATaxBillPersLaser.GetName ReportFile$
  arVATaxBillPersLaser.Show
  If PrnCnt > 0 Then
    Call MakeFile 'Makefile creates the .dat identifying that bill printing took place
  End If
  
  Exit Sub
  
GetBarCodeData:
  If fpcmbBarCode.Text = "No" Then
    BZip = ""
    Return
  End If
  Get TCHandle, TBRec.CustPin, TaxCust
  If Len(QPTrim$(TaxCust.Zip)) < 10 Or Len(QPTrim$(TaxCust.DeliveryPt)) <> 2 Then
    BZip = ""
  Else
    BZip = QPTrim$(TaxCust.Zip) + QPTrim$(TaxCust.DeliveryPt)
  End If
  
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxBillPrinting", "PrintLaser1", Erl)
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

Private Sub DisablePersonal()
  fpDblSnglStartPersBill.Enabled = False
  fpLongPersTaxYear.Enabled = False
  fpDblSnglPersRate.Enabled = False
  fpDblSnglMHRate.Enabled = False
  fpDblSnglMTRate.Enabled = False
  fpDblSnglMCRate.Enabled = False
  fpDblSnglFERate.Enabled = False
  fpDblSnglPersLateList.Enabled = False
  fptxtPersOrder.Enabled = False
  fptxtPersDueDate.Enabled = False
End Sub
Private Sub DisableReal()
  fpDblSnglStartRealBill.Enabled = False
  fpLongRealTaxYear.Enabled = False
  fpDblSnglRealRate.Enabled = False
  fpDblSnglRealLateList.Enabled = False
  fptxtRealOrder.Enabled = False
  fptxtRealDueDate.Enabled = False
End Sub
Private Sub EnablePersonal()
  fpDblSnglStartPersBill.Enabled = True
  fpLongPersTaxYear.Enabled = True
  fpDblSnglPersRate.Enabled = True
  fpDblSnglMHRate.Enabled = True
  fpDblSnglMTRate.Enabled = True
  fpDblSnglMCRate.Enabled = True
  fpDblSnglFERate.Enabled = True
  fpDblSnglPersLateList.Enabled = True
  fptxtPersOrder.Enabled = True
  fptxtPersDueDate.Enabled = True
  Call CheckComment
End Sub
Private Sub EnableReal()
  fpDblSnglStartRealBill.Enabled = True
  fpLongRealTaxYear.Enabled = True
  fpDblSnglRealRate.Enabled = True
  fpDblSnglRealLateList.Enabled = True
  fptxtRealOrder.Enabled = True
  fptxtRealDueDate.Enabled = True
  Call CheckComment
End Sub

Private Sub PrintPersLaserItemized()
  Dim ToPrint As String
  Dim TaxRptT As Integer
  Dim ReportFile As String
  Dim RptHandle As Integer
  Dim TBRec As VAPPTaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim x As Long
  Dim dlm$, BillNo&, PrnCnt As Long
  Dim TBDRec As TxBillLaserItemized
  Dim TBDHandle As Integer
  Dim ThisRate As Double
  Dim TotValue As Double
  Dim ThisOpt1Desc As String * 15
  Dim ThisOpt2Desc As String * 15
  Dim ThisOpt3Desc As String * 15
  Dim BZip As String
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPRecs As Long
  Dim NumOfPers As Integer
  Dim NextRec As Long
  Dim thisVin As String
  Dim y As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim NumOfOpts As Integer
  Dim ZipRec As BillPrintPZipIdxType, CustArr As Long
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim AHandle As Integer
  Dim MinTaxedAmt As Double
  Dim PocahFlag As Boolean
  On Error GoTo ERRORSTUFF
  
'  AHandle = FreeFile
'  Open "persbill.dat" For Output As AHandle
'
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  NumOfOpts = 0
  If QPTrim$(TaxMasterRec.POptRev1) <> "" Then
    NumOfOpts = NumOfOpts + 1
  End If
  If QPTrim$(TaxMasterRec.POptRev2) <> "" Then
    NumOfOpts = NumOfOpts + 1
  End If
  If QPTrim$(TaxMasterRec.POptRev3) <> "" Then
    NumOfOpts = NumOfOpts + 1
  End If
  If InStr(TaxMasterRec.Name, "POCAHONTAS") > 0 Then
    PocahFlag = True
  End If
  
  MinTaxedAmt = TaxMasterRec.MinVehTaxVal
  
  dlm$ = "~"
  ReportFile$ = StartPath$ + "/TaxPLsrItem.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  BillNo& = fpDblSnglStartPersBill.Value
  
  frmVATaxShowPctComp.Label1 = "Printing Personal Tax Bills"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  EnableCloseButton Me.hwnd, False
  
  OpenLaserPersItemized TBDHandle
  Get #TBDHandle, 1, TBDRec
  If fpcmbBarCode.Text = "No" Then
    TBDRec.UseBarCode = False
    Put #TBDHandle, 1, TBDRec
  ElseIf fpcmbBarCode.Text = "Yes" Then
    TBDRec.UseBarCode = True
    Put #TBDHandle, 1, TBDRec
  End If
  Close TBDHandle
  
  If TBDRec.dologo = 1 Then
    If Exist("towntaxlogo.bmp") Then
      arVATaxLaserPersItemized.Image1.Picture = LoadPicture("towntaxlogo.bmp")
      arVATaxLaserPersItemized.Image1.Visible = True
    End If
  End If
  OpenTaxCustFile TCHandle, NumOfTCRecs
  ReDim VinArray(1 To 1) As String
  OpenPersPropFile PHandle, NumOfPRecs
  OpenPersTaxBillFile TBHandle, NumOfTBRecs
  
  If Exist("PZipIdx.Dat") Then '12/6/06
    OpenPZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  
  For x = 1 To NumOfTBRecs
    NumOfPers = 0
    If optZip.Value = True Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TBRec
'    If TBRec.CustRec = 85 Then Stop
    Get TCHandle, TBRec.CustRec, TaxCust
    If TBRec.BillNumber > 0 Then
'      If OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt) > 0 Then
      If OldRound(TBRec.TotalBillDue) > 0 Then '5/11/07
        TBRec.BillNumber = BillNo&
        TBRec.BillPrinted = True
        Put TBHandle, CustArr, TBRec
        If TBRec.PPTRAValue > 0 Then
          TotValue = OldRound(TBRec.PPTRAValue + TBRec.FEValue + TBRec.MCValue + TBRec.MHValue + TBRec.MTValue - (TBRec.ExptValue)) ' + TBRec.PPTRAValue))
        Else
          TotValue = OldRound(TBRec.PersValue + TBRec.FEValue + TBRec.MCValue + TBRec.MHValue + TBRec.MTValue - (TBRec.ExptValue)) ' + TBRec.PPTRAValue))
        End If
        GoSub GetBarCodeData
'        Print #AHandle, CStr(TBRec.CustRec) + "~" + Using("###,###,##0.00", TBRec.TotalBillDue)
        '                         0                         1
        Print #RptHandle, TBRec.BillNumber; dlm; QPTrim$(TBRec.CustName); dlm;
        '                           2                           3
        Print #RptHandle, QPTrim$(TBRec.CustAdd1); dlm; QPTrim$(TBRec.CustAdd2); dlm;
        '                             4                      5
        Print #RptHandle, QPTrim$(TBRec.CustAdd3); dlm; TBRec.CustPin; dlm;
        '                     6               7
        Print #RptHandle, TotValue; dlm; TBRec.RDesc1; dlm;
        '                        8                   9                    10
        Print #RptHandle, TBRec.PersValue; dlm; TBRec.FEValue; dlm; TBRec.ExptValue; dlm;
        '                    11                      12                               13
        Print #RptHandle, TBRec.PPTRAValue; dlm; TBRec.PPTRADiscnt; dlm; OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt); dlm;
        '                       14                 15                   16
        Print #RptHandle, TBDRec.dologo; dlm; TBRec.MHValue; dlm; TBRec.MCValue; dlm;
        '                      17                            18                                          19
        Print #RptHandle, TBRec.MTValue; dlm; OldRound(TBRec.PersTaxDue); dlm; TBRec.PersTaxNet; dlm;
        '                         20                    21                    22
        Print #RptHandle, TBRec.PersTaxRate; dlm; TBRec.FETaxDue; dlm; TBRec.FETaxRate; dlm;
        '                       23                    24                    25
        Print #RptHandle, TBRec.MCTaxDue; dlm; TBRec.MCTaxRate; dlm; TBRec.MHTaxDue; dlm;
        '                       26                    27                     28
        Print #RptHandle, TBRec.MHTaxRate; dlm; TBRec.MTTaxDue; dlm; TBRec.MTTaxRate; dlm;
        '                        29                     30                     31
        Print #RptHandle, TBRec.OptRevTax1; dlm; TBRec.OptRevTax2; dlm; TBRec.OptRevTax3; dlm;
        ThisOpt1Desc = QPTrim$(TBRec.OptRevDesc1)
        ThisOpt2Desc = QPTrim$(TBRec.OptRevDesc2)
        ThisOpt3Desc = QPTrim$(TBRec.OptRevDesc3)
        '                        32              33                 34              35             36                  37
        Print #RptHandle, ThisOpt1Desc; dlm; ThisOpt2Desc; dlm; ThisOpt3Desc; dlm; BZip; dlm; TBRec.CustZip; dlm; TBDRec.dologo; dlm;
        '                           38                           39                            40
        Print #RptHandle, QPTrim(TBDRec.TxtHead1); dlm; QPTrim(TBDRec.TxtHead2); dlm; QPTrim(TBDRec.txtOpt1); dlm;
        '                           41                           42                 43
        Print #RptHandle, QPTrim(TBDRec.TxtOpt2); dlm; QPTrim(TBDRec.TxtOpt3); dlm; ""; dlm;
        '                           44                           45                             46
        Print #RptHandle, QPTrim(TBDRec.txtPgph0); dlm; QPTrim(TBDRec.txtPgph1); dlm; QPTrim(TBDRec.txtPgph2); dlm;
        '                           47                           48                            49
        Print #RptHandle, QPTrim(TBDRec.txtPgph3); dlm; QPTrim(TBDRec.txtPgph4); dlm; "                     "; dlm;
        '                           50               51       52
        Print #RptHandle, "                  "; dlm; ""; dlm; ""; dlm;
        '                           53                           54                            55
        Print #RptHandle, QPTrim(TBDRec.txtHead3); dlm; QPTrim(TBDRec.txtHead4); dlm; QPTrim(TBDRec.txtHead5); dlm;
        '                 56       57
        Print #RptHandle, ""; dlm; ""; dlm;
        NextRec = TaxCust.FirstPersRec
        If NextRec > 0 Then
          Get PHandle, NextRec, PersRec
          NumOfPers = NumOfPers + 1
          '                           58                     59                    60
          Print #RptHandle, QPTrim$(PersRec.Vin); dlm; QPTrim$(PersRec.MakeMod); dlm; PersRec.PersVal; dlm;
          '                       61                     62                   63                  64                  65                66                      67
          Print #RptHandle, PersRec.MTValue; dlm; PersRec.MCValue; dlm; PersRec.CVALUE; dlm; PersRec.MHValue; dlm; NumOfPers; dlm; TBRec.OverPayAmt; dlm; TBRec.TotalBillDue
          NextRec = PersRec.NextRec
        Else
          '                 58       59       60       61       62       63       64          65                 66                      67
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; NumOfPers; dlm; TBRec.OverPayAmt; dlm; TBRec.TotalBillDue
        End If

        Do While NextRec > 0
          Get PHandle, NextRec, PersRec
          If PocahFlag Then  '092414
            If PersRec.PersVal < MinTaxedAmt Then
              GoTo SkipThisCar
            End If
          End If
          NumOfPers = NumOfPers + 1
          '                         0                         1                           2                                3
          Print #RptHandle, TBRec.BillNumber; dlm; QPTrim$(TBRec.CustName); dlm; QPTrim$(TBRec.CustAdd1); dlm; QPTrim$(TBRec.CustAdd2); dlm;
          '                             4                        5               6            7        8        9        10
          Print #RptHandle, QPTrim$(TBRec.CustAdd3); dlm; TBRec.CustPin; dlm; TotValue; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm;
          '                 11       12                          13
          Print #RptHandle, ""; dlm; ""; dlm; OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt); dlm;
          '                 14       15       16       17       18
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm;
          '                 19       20       21       22       23       24       25       26
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm;
          '                 27       28       29       30       31           32                  33                34
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ThisOpt1Desc; dlm; ThisOpt2Desc; dlm; ThisOpt3Desc; dlm;
          '                 35       36       37                38                            39                            40
          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; QPTrim(TBDRec.TxtHead1); dlm; QPTrim(TBDRec.TxtHead2); dlm; QPTrim(TBDRec.txtOpt1); dlm;
          '                            41                           42
          Print #RptHandle, QPTrim(TBDRec.TxtOpt2); dlm; QPTrim(TBDRec.TxtOpt3); dlm;
          '                 43                 44                           45                            46
          Print #RptHandle, ""; dlm; QPTrim(TBDRec.txtPgph0); dlm; QPTrim(TBDRec.txtPgph1); dlm; QPTrim(TBDRec.txtPgph2); dlm;
          '                            47                 48       49       50
          Print #RptHandle, QPTrim(TBDRec.txtPgph3); dlm; ""; dlm; ""; dlm; ""; dlm;
          '                 51       52                53                               54                        55
          Print #RptHandle, ""; dlm; ""; dlm; QPTrim(TBDRec.txtHead3); dlm; QPTrim(TBDRec.txtHead4); dlm; QPTrim(TBDRec.txtHead5); dlm;
          '                 56       57              58
          Print #RptHandle, ""; dlm; ""; dlm; QPTrim$(PersRec.Vin); dlm;
          '                         59                           60                   61                       62
          Print #RptHandle, QPTrim$(PersRec.MakeMod); dlm; PersRec.PersVal; dlm; PersRec.MTValue; dlm; PersRec.MCValue; dlm;
          '                        63                   64                65                  66                     67
          Print #RptHandle, PersRec.CVALUE; dlm; PersRec.MHValue; dlm; NumOfPers; dlm; TBRec.OverPayAmt; dlm; TBRec.TotalBillDue
          
          If NumOfOpts = 1 Then
            If NumOfPers = 13 Then NumOfPers = 0
          ElseIf NumOfOpts = 2 Then
            If NumOfPers = 12 Then NumOfPers = 0
          ElseIf NumOfOpts = 3 Then
            If NumOfPers = 11 Then NumOfPers = 0
          Else
            If NumOfPers = 14 Then NumOfPers = 0
          End If
SkipThisCar:
          NextRec = PersRec.NextRec
        Loop
        BillNo& = BillNo& + 1
        PrnCnt = PrnCnt + 1
      End If
    End If
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  EnableCloseButton Me.hwnd, True
  
  Close
  arVATaxLaserPersItemized.Show
  If PrnCnt > 0 Then
    Call MakeFile 'Makefile creates the .dat identifying that bill printing took place
  End If
  
  Exit Sub
  
GetBarCodeData:
  If fpcmbBarCode.Text = "No" Then
    BZip = ""
    Return
  End If
  Get TCHandle, TBRec.CustPin, TaxCust
  If Len(QPTrim$(TaxCust.Zip)) < 10 Or Len(QPTrim$(TaxCust.DeliveryPt)) <> 2 Then
    BZip = ""
  Else
    BZip = QPTrim$(TaxCust.Zip) + QPTrim$(TaxCust.DeliveryPt)
  End If
  
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxBillPrinting", "PrintPersLaserItemized", Erl)
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

Private Sub CheckComment()
  Dim TaxPBill As VAPPTaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim TaxRBill As VARETaxBillType
  Dim x As Long
  
  If fptxtCurrForm.Text = "STANDARD" Then
    If fpcmbType.Text = "REAL" Then
      If Exist(RealTaxBillFile) Then
        OpenRealTaxBillFile TBHandle, NumOfTBRecs
        Get TBHandle, 1, TaxRBill
        Close TBHandle
        If QPTrim$(TaxRBill.CommentPlace) = "" Then
          fpcmbCommentPlace.Text = "NO COMMENT"
        Else
          fpcmbCommentPlace.Text = TaxRBill.CommentPlace
        End If
        fptxtComment.Text = TaxRBill.Comment
        fptxtComment2.Text = TaxRBill.Comment2
      Else
        fpcmbCommentPlace.Text = "NO COMMENT"
      End If
    ElseIf fpDblSnglStartPersBill.Enabled = True Then
      If fpcmbType.Text = "PERSONAL" Then
        If Exist(PersTaxBillFile) Then
          OpenPersTaxBillFile TBHandle, NumOfTBRecs
          Get TBHandle, 1, TaxPBill
          Close TBHandle
          If QPTrim$(TaxPBill.CommentPlace) = "" Then
            fpcmbCommentPlace.Text = "NO COMMENT"
          Else
            fpcmbCommentPlace.Text = TaxPBill.CommentPlace
          End If
          fptxtComment.Text = TaxPBill.Comment
          fptxtComment2.Text = TaxPBill.Comment2
        Else
          fpcmbCommentPlace.Text = "NO COMMENT"
        End If
      End If
    End If
  ElseIf fptxtCurrForm.Text = "CDRBLUFF" Then
    If fpcmbType.Text = "REAL" Then
      If Exist(RealTaxBillFile) Then
        OpenRealTaxBillFile TBHandle, NumOfTBRecs
        For x = 1 To NumOfTBRecs
          Get TBHandle, x, TaxRBill
          If QPTrim$(TaxRBill.Comment) <> "" Then
            If IsNumeric(TaxRBill.Comment) Then
              fpDSMnthlyPen.Value = Val(TaxRBill.Comment)
            Else
              fpDSMnthlyPen.Value = 0
            End If
            Exit For
          End If
        Next x
        Close TBHandle
        If x > NumOfTBRecs Then
          fpDSMnthlyPen.Value = 0
        End If
      End If
    End If
  End If
  
End Sub

'Private Sub PrintRealLaserItemized()
'  Dim ToPrint As String
'  Dim TaxRptT As Integer
'  Dim ReportFile As String
'  Dim RptHandle As Integer
'  Dim TBRec As VARETaxBillType
'  Dim TBHandle As Integer
'  Dim NumOfTBRecs As Long
'  Dim x As Long
'  Dim dlm$, BillNo&, PrnCnt As Long
'  Dim TBDRec As TxBillLaserItemized
'  Dim TBDHandle As Integer
'  Dim ThisRate As Double
'  Dim TotValue As Double
'  Dim ThisOpt1Desc As String * 15
'  Dim ThisOpt2Desc As String * 15
'  Dim ThisOpt3Desc As String * 15
'  Dim BZip As String
'  Dim TaxCust As TaxCustType
'  Dim TCHandle As Integer
'  Dim NumOfTCRecs As Long
'  Dim RealRec As PropertyRecType
'  Dim RHandle As Integer
'  Dim NumOfRRecs As Long
'  Dim NumOfReal As Integer
'  Dim NextRec As Long
'  Dim thisVin As String
'  Dim y As Integer
'  Dim TaxMasterRec As TaxMasterType
'  Dim TMHandle As Integer
'  Dim NumOfOpts As Integer
'
'  On Error GoTo ERRORSTUFF
'
'  OpenTaxSetUpFile TMHandle
'  Get TMHandle, 1, TaxMasterRec
'  Close TMHandle
'  NumOfOpts = 0
'  If QPTrim$(TaxMasterRec.OptRev1) <> "" Then
'    NumOfOpts = NumOfOpts + 1
'  End If
'  If QPTrim$(TaxMasterRec.OptRev2) <> "" Then
'    NumOfOpts = NumOfOpts + 1
'  End If
'  If QPTrim$(TaxMasterRec.OptRev3) <> "" Then
'    NumOfOpts = NumOfOpts + 1
'  End If
'  dlm$ = "~"
'  ReportFile$ = StartPath$ + "/TaxRLsrItem.RPT"
'  RptHandle = FreeFile
'  Open ReportFile$ For Output As #RptHandle
'  BillNo& = fpDblSnglStartRealBill.Value
'
'  frmVATaxShowPctComp.Label1 = "Printing Real Tax Bills"
'  frmVATaxShowPctComp.cmdCancel.Visible = False
'  frmVATaxShowPctComp.Show , Me
'  cmdProcess.Enabled = False
'  cmdExit.Enabled = False
'  EnableCloseButton Me.hwnd, False
'
'  OpenLaserRealItemized TBDHandle
'  Get #TBDHandle, 1, TBDRec
'  If fpcmbBarCode.Text = "No" Then
'    TBDRec.UseBarCode = False
'    Put #TBDHandle, 1, TBDRec
'  ElseIf fpcmbBarCode.Text = "Yes" Then
'    TBDRec.UseBarCode = True
'    Put #TBDHandle, 1, TBDRec
'  End If
'  Close TBDHandle
'
'  If TBDRec.dologo = 1 Then
'    If Exist("towntaxlogo.bmp") Then
'      arVATaxLaserRealItemized.Image1.Picture = LoadPicture("towntaxlogo.bmp")
'      arVATaxLaserRealItemized.Image1.Visible = True
'    End If
'  End If
'  OpenTaxCustFile TCHandle, NumOfTCRecs
'  ReDim VinArray(1 To 1) As String
'  OpenRealPropFile RHandle, NumOfRRecs
'  OpenRealTaxBillFile TBHandle, NumOfTBRecs
'  For x = 1 To NumOfTBRecs
'    NumOfReal = 0
'    Get TBHandle, x, TBRec
'    Get TCHandle, TBRec.CustPin, TaxCust
'    If TBRec.BillNumber > 0 Then
'      If OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt) > 0 Then
'        TBRec.BillNumber = BillNo&
'        TBRec.BillPrinted = True
'        Put TBHandle, x, TBRec
'        TotValue = OldRound((TBRec.RealValue + TBRec.BldgValue) - TBRec.ExptValue)
'        GoSub GetBarCodeData
'        '                         0                         1
'        Print #RptHandle, TBRec.BillNumber; dlm; QPTrim$(TBRec.CustName); dlm;
'        '                           2                           3
'        Print #RptHandle, QPTrim$(TBRec.CustAdd1); dlm; QPTrim$(TBRec.CustAdd2); dlm;
'        '                             4                      5
'        Print #RptHandle, QPTrim$(TBRec.CustAdd3); dlm; TBRec.CustPin; dlm;
'        '                     6               7
'        Print #RptHandle, TotValue; dlm; TBRec.RDesc1; dlm;
'        '                        8                   9                    10
'        Print #RptHandle, TBRec.RealValue; dlm; TBRec.BldgValue; dlm; TBRec.ExptValue; dlm;
'        '                                      11
'        Print #RptHandle, OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt); dlm;
'        '                       12                                 13
'        Print #RptHandle, TBDRec.dologo; dlm; OldRound(TBRec.RealTaxDue - TBRec.OverPayAmt); dlm;
'        '                       14                       15
'        Print #RptHandle, TBRec.RealTaxDue; dlm; TBRec.RealTaxRate; dlm;
'        '                        16                     17                     18
'        Print #RptHandle, TBRec.OptRevTax1; dlm; TBRec.OptRevTax2; dlm; TBRec.OptRevTax3; dlm;
'        ThisOpt1Desc = QPTrim$(TBRec.Opt1Desc)
'        ThisOpt2Desc = QPTrim$(TBRec.Opt2Desc)
'        ThisOpt3Desc = QPTrim$(TBRec.Opt3Desc)
'        '                        19              20                 21              22             23                  24
'        Print #RptHandle, ThisOpt1Desc; dlm; ThisOpt2Desc; dlm; ThisOpt3Desc; dlm; BZip; dlm; TBRec.CustZip; dlm; TBDRec.dologo; dlm;
'        '                           25                           26                            27
'        Print #RptHandle, QPTrim(TBDRec.TxtHead1); dlm; QPTrim(TBDRec.TxtHead2); dlm; QPTrim(TBDRec.txtOpt1); dlm;
'        '                           28                           29                          30
'        Print #RptHandle, QPTrim(TBDRec.TxtOpt2); dlm; QPTrim(TBDRec.TxtOpt3); dlm; QPTrim(TBDRec.txtPgph0); dlm;
'        '                           31                           32                             33                          34
'        Print #RptHandle, QPTrim(TBDRec.txtPgph1); dlm; QPTrim(TBDRec.txtPgph2); dlm; QPTrim(TBDRec.txtPgph3); dlm; QPTrim$(TBDRec.txtPgph4); dlm;
'        '                           35                           36                            37
'        Print #RptHandle, QPTrim(TBDRec.txtHead3); dlm; QPTrim(TBDRec.txtHead4); dlm; QPTrim(TBDRec.txtHead5); dlm;
'        NextRec = TaxCust.FirstPropRec
'        If NextRec > 0 Then
'          NumOfReal = NumOfReal + 1
'          Get RHandle, NextRec, RealRec
'          '                        38                39
'          Print #RptHandle, RealRec.PROPVALU; dlm; NumOfReal; dlm;
'          NextRec = RealRec.NextRec
'        Else
'          '                 38         39
'          Print #RptHandle, ""; dlm; NumOfReal; dlm;
'        End If
'        '                    40                   41                    42                    43
'        Print #RptHandle, RealRec.Map; dlm; RealRec.BLOCK; dlm; RealRec.LOTNUMB; dlm; RealRec.PropAddr
'        Do While NextRec > 0
'          Get RHandle, NextRec, RealRec
'          NumOfReal = NumOfReal + 1
'          '                         0                         1                   2        3
'          Print #RptHandle, TBRec.BillNumber; dlm; QPTrim$(TBRec.CustName); dlm; ""; dlm; ""; dlm;
'          '                  4             5             6        7        8        9        10
'          Print #RptHandle, ""; dlm; TBRec.CustPin; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm;
'          '                                    11
'          Print #RptHandle, OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt); dlm;
'          '                 12       13       14       15       16
'          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm;
'          '                 17       18           19                  20                 21
'          Print #RptHandle, ""; dlm; ""; dlm; ThisOpt1Desc; dlm; ThisOpt2Desc; dlm; ThisOpt3Desc; dlm;
'          '                 22       23       24       25       26       27       28       29       30
'          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm;
'          '                 31       32       33       34       35       36       37             38                 39
'          Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; RealRec.PROPVALU; dlm; NumOfReal; dlm;
'          '                    40                   41                    42                    43
'          Print #RptHandle, RealRec.Map; dlm; RealRec.BLOCK; dlm; RealRec.LOTNUMB; dlm; RealRec.PropAddr
'
'          If NumOfOpts = 1 Then
'            If NumOfReal = 13 Then NumOfReal = 0
'          ElseIf NumOfOpts = 2 Then
'            If NumOfReal = 12 Then NumOfReal = 0
'          ElseIf NumOfOpts = 3 Then
'            If NumOfReal = 11 Then NumOfReal = 0
'          Else
'            If NumOfReal = 14 Then NumOfReal = 0
'          End If
'          NextRec = RealRec.NextRec
'        Loop
'        BillNo& = BillNo& + 1
'        PrnCnt = PrnCnt + 1
'      End If
'    End If
'    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
'    If frmVATaxShowPctComp.Out = True Then
'      Close
'      frmVATaxShowPctComp.Out = False
'      Unload frmVATaxShowPctComp
'      cmdProcess.Enabled = True
'      cmdExit.Enabled = True
'      EnableCloseButton Me.hwnd, True
'      Exit Sub
'    End If
'  Next x
'  Unload frmVATaxShowPctComp
'  cmdProcess.Enabled = True
'  cmdExit.Enabled = True
'  EnableCloseButton Me.hwnd, True
'
'  Close
'  arVATaxLaserRealItemized.Show
'  If PrnCnt > 0 Then
'    Call MakeFile 'Makefile creates the .dat identifying that bill printing took place
'  End If
'
'  Exit Sub
'
'GetBarCodeData:
'  If fpcmbBarCode.Text = "No" Then
'    BZip = ""
'    Return
'  End If
'  Get TCHandle, TBRec.CustPin, TaxCust
'  If Len(QPTrim$(TaxCust.Zip)) < 10 Or Len(QPTrim$(TaxCust.DeliveryPt)) <> 2 Then
'    BZip = ""
'  Else
'    BZip = QPTrim$(TaxCust.Zip) + QPTrim$(TaxCust.DeliveryPt)
'  End If
'
'  Return
'
'ERRORSTUFF:
'   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxBillPrinting", "PrintRealLaserItemized", Erl)
'     Case emrExitProc:
'       Resume Proc_Exit
'     Case emrResume:
'       Resume
'     Case emrResumeNext:
'       Resume Next
'     Case Else
'      '--- Technically, this should never happen.
'       Resume Proc_Exit
'   End Select
'
'Proc_Exit:
'  '--- Cleanup code goes here...
'    Close
'
'End Sub
'
Private Sub PrintPersExport()
  Dim TaxXRec As TaxBillExportPersType
  Dim TXHandle As Integer
  Dim TBRec As VAPPTaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim ReportFile As String
  Dim RptHandle As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim TaxYear$
  Dim FF11$
  Dim Map$
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPRecs As Long
  Dim BillNo&
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim ThisSSN As String * 11
  Dim PersExemp As Double
  Dim NextRec As Long
  Dim ExempAmt As Double
  Dim PersDue As Double
  Dim TotPersVal As Double
  Dim FF9$
  Dim PrnCnt As Long
  Dim PERC!
  Dim CustName$
  Dim TAXRATE#
  Dim Desc$
  Dim CarCount As Integer
  Dim WhatPer&
  Dim PYearStr$
  Dim PPTRAVal#
  Dim PPTRADiscount#
  Dim DueDate$
  Dim WhatYear$
  Dim x As Long
  Dim VehDesc$, PYear As Integer
  Dim TaxAmt#, WhatPers&
  Dim ThisDesc As String
  Dim PastDue As Double
  
  On Error GoTo ERRORSTUFF
  
  WhatYear = fpLongPersTaxYear.Text
  PrnCnt = 0
  FF11$ = "########.#0"
  FF9$ = "######.#0"
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  TaxYear$ = CStr(TaxMasterRec.PTaxYear)

  ReportFile$ = "LCPP" + TaxYear + ".TXT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  BillNo& = fpDblSnglStartPersBill.Value
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  PERC! = TaxMasterRec.PPTRADisc
  PersTaxRate# = fpDblSnglPersRate.Value ' / 100
  TaxYear$ = CStr(TaxMasterRec.PTaxYear)
  OpenPersPropFile PHandle, NumOfPRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenPersTaxBillFile TBHandle, NumOfTBRecs

  frmVATaxShowPctComp.Label1 = "Printing Tax Billing Data"
  frmVATaxShowPctComp.Show , Me
  frmVATaxShowPctComp.cmdCancel.Visible = False
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  
  For x = 1 To NumOfTBRecs
    Get TBHandle, x, TBRec
    If TBRec.BillNumber > 0 Then
'      If OldRound(TBRec.TotalBillDue - TBRec.OverPayAmt) > 0 Then 'assign and save bill number
      If OldRound(TBRec.TotalBillDue) > 0 Then '5/11/07
        TBRec.BillNumber = BillNo&
        TBRec.BillPrinted = True
        Put TBHandle, x, TBRec
        BillNo& = BillNo& + 1
      End If
      Get TCHandle, TBRec.CustRec, TaxCust
      CustName$ = QPTrim$(TaxCust.CustName)
      TAXRATE# = TBRec.PersTaxRate
      PrnCnt = PrnCnt + 1
      Print #RptHandle, QPTrim$(CStr(TBRec.TaxYear)); "~";
      Print #RptHandle, QPTrim$(CStr(TBRec.BillNumber)); "~";
      Print #RptHandle, CStr(TBRec.CustRec); "~";
      Print #RptHandle, QPTrim$(CustName$); "~";
      Print #RptHandle, QPTrim$(TBRec.CustAdd1); "~";
      Print #RptHandle, QPTrim$(TBRec.CustAdd2); "~";
      Print #RptHandle, QPTrim$(TBRec.CustAdd3); "~";
      Print #RptHandle, QPTrim$(TBRec.CustZip); "~";
      
      'Desc$ = QPTrim$(TBRec.RDesc1) + " " + QPTrim$(TBRec.RDesc2)
      'Print #RptHandle, QPTrim$(Desc$); "~";
      Print #RptHandle, QPTrim$(Using$("##.##", Str$(PersTaxRate#))); "~";
      Print #RptHandle, QPTrim$(Using$("##.##", Str$(MHTaxRate#))); "~";
      Print #RptHandle, QPTrim$(Using$("##.##", Str$(MCTaxRate#))); "~";
      Print #RptHandle, QPTrim$(Using$("##.##", Str$(FETaxRate#))); "~";
      Print #RptHandle, QPTrim$(Using$("##.##", Str$(MTTaxRate#))); "~";

      Print #RptHandle, QPTrim$(Str$(TBRec.PersValue)); "~";
      Print #RptHandle, QPTrim$(Str$(TBRec.MHValue)); "~";
      Print #RptHandle, QPTrim$(Str$(TBRec.MCValue)); "~";
      Print #RptHandle, QPTrim$(Str$(TBRec.FEValue)); "~";
      Print #RptHandle, QPTrim$(Str$(TBRec.MTValue)); "~";

      Print #RptHandle, QPTrim$(Using$("######.##", Str$(TBRec.PersTaxDue))); "~";
      Print #RptHandle, QPTrim$(Using$("######.##", Str$(TBRec.MHTaxDue))); "~";
      Print #RptHandle, QPTrim$(Using$("######.##", Str$(TBRec.MCTaxDue))); "~";
      Print #RptHandle, QPTrim$(Using$("######.##", Str$(TBRec.FETaxDue))); "~";
      Print #RptHandle, QPTrim$(Using$("######.##", Str$(TBRec.MTTaxDue))); "~";
      Print #RptHandle, QPTrim$(Using$("######.##", Str$(TBRec.OptRevTax1))); "~";
      Print #RptHandle, QPTrim$(Using$("######.##", Str$(TBRec.OptRevTax2))); "~";
      Print #RptHandle, QPTrim$(Using$("######.##", Str$(TBRec.OptRevTax3))); "~";
      Print #RptHandle, QPTrim$(Using$("######.##", Str$(TBRec.ExptValue))); "~";
      Print #RptHandle, QPTrim$(Using$("######.##", Str$(TBRec.PPTRADiscnt))); "~";
'      Print #RptHandle, QPTrim$(Using$("######.##", Str$(TBRec.TotalBillDue))); "~";
'      DueDate = MakeRegDate(TBRec.DueDate)
'      PastDue = GetCustPersBalance(CLng(TBRec.CustRec), -1)
''Dale hacked
'      Print #RptHandle, DueDate$; "~"; QPTrim$(Using$("######.##", Str$(PastDue))); "~PERS_START~";
      PastDue = GetCustPersBalance(CLng(TBRec.CustRec), -1)
      Print #RptHandle, QPTrim$(Using$("######.##", Str$(TBRec.TotalBillDue))); "~";
      DueDate = MakeRegDate(TBRec.DueDate)
'Dale hacked
      Print #RptHandle, DueDate$; "~"; QPTrim$(Using$("######.##", Str$(PastDue))); "~";
      If PastDue < 0 Then
'        Print #RptHandle, QPTrim$(Using$("######.##", Str$(OldRound(TBRec.TotalBillDue - PastDue)))); "~";
        Print #RptHandle, QPTrim$(Using$("######.##", Str$(OldRound(TBRec.TotalBillDue - Abs(PastDue))))); "~"; 'inserted this line 11/4/2009
      Else
        Print #RptHandle, QPTrim$(Using$("######.##", Str$(TBRec.TotalBillDue))); "~";
      End If
      Print #RptHandle, "PERS_START~";

      CarCount = 0
      WhatPers& = TaxCust.FirstPersRec
      Do
        Get PHandle, WhatPers&, PersRec
        PYear = PersRec.TaxBillYear
        If PYear > 0 And PYear <> WhatYear Then
          GoTo NextLoop
        Else
'Dale hacked
          PPTRADiscount# = 0
          If PersRec.PPTRAYN = "Y" Then 'changed on 10/30/06
            If OldRound#(PersRec.PersVal) > TaxMasterRec.MaxVehTaxVal Then
              PPTRAVal# = TaxMasterRec.MaxVehTaxVal
            Else
              PPTRAVal# = OldRound#(PersRec.PersVal)
            End If
            If PPTRAVal# <= TaxMasterRec.MinVehTaxVal Then
              PPTRADiscount# = OldRound#((PPTRAVal# / 100) * PersTaxRate#)
            Else
              PPTRADiscount# = OldRound#(((PPTRAVal# / 100) * (PERC! / 100)) * PersTaxRate#)
            End If
          End If
            VehDesc$ = "VIN:" + QPTrim$(PersRec.Vin)
            TaxAmt# = OldRound#((PersTaxRate# / 100) * PersRec.PersVal)
'Dale hacked
            ThisDesc$ = ""
            ThisDesc$ = QPTrim$(PersRec.DESC1) + " " + QPTrim$(PersRec.DESC2) + " " + QPTrim$(PersRec.DESC3) + " " + QPTrim$(PersRec.Desc4) + " " + QPTrim$(PersRec.Desc5)
            
            Print #RptHandle, QPTrim$(Using$("########", Str$(PersRec.PersVal))); "~";
            Print #RptHandle, QPTrim$(Using$("######.##", Str$(TaxAmt#))); "~";
            Print #RptHandle, QPTrim$(Using$("######.##", Str$(PPTRADiscount#))); "~";
'Dale hacked
            Print #RptHandle, ThisDesc$; "~";
            Print #RptHandle, VehDesc$; "~";
            Print #RptHandle, QPTrim$(PersRec.MakeMod); "~";
            Print #RptHandle, CStr(PersRec.ModYear); "~";
            CarCount = CarCount + 1
'Dale hacked
          'End If
        End If
NextLoop:
        WhatPers& = PersRec.NextRec
      Loop While WhatPers& > 0
      Print #RptHandle, "~PERS_END~"
    End If
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  Close
  If PrnCnt > 0 Then
    Call MakeFile 'Makefile creates the .dat identifying that bill printing took place
  End If

  Call TaxMsg(800, "The tax billing records have been successfully saved in the file named 'LCPP" + TaxYear + ".TXT' located in the Citipak folder.")

  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxBillPrinting", "PrintPersExport", Erl)
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
 ' --- Cleanup code goes here...
    Close


End Sub

Private Sub PrintRealMiddletownBill()
  Dim TaxBill As VARETaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim BillInfo As VARETaxBillInfoType
  Dim BIHandle As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long, RealTaxRate#
  Dim File$, WordLen As Integer
  Dim CustName As String * 45
  Dim RptFile#, ch$, y As Integer
  Dim TownName$, Add1$, Add2$, Add3$
  Dim RHandle As Integer, PrnCnt As Integer
  Dim NumOfRRecs As Long, TotOpt As Double
  Dim RealRec As PropertyRecType
  Dim TaxAmt#, Tab1 As Integer, Tab2 As Integer, Tab3 As Integer
  Dim DueDate$, WorkName$
  Dim ZipRec As BillPrintRZipIdxType, CustArr As Long
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long, BillNo As Long
  
  DueDate$ = fptxtRealDueDate
  RealTaxRate# = fpDblSnglRealRate
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  TownName = QPTrim$(TaxMasterRec.Name)
  Add1$ = QPTrim$(TaxMasterRec.Add1)
  Add2$ = QPTrim$(TaxMasterRec.Add2)
  Add3$ = QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TownState) + " " + QPTrim$(TaxMasterRec.Zip)
  
  OpenRealPropFile RHandle, NumOfRRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenRealTaxBillFile TBHandle, NumOfTBRecs
  OpenRealBillInfoFile BIHandle
  Get BIHandle, 1, BillInfo
  Close BIHandle
  File$ = StartPath$ + "/TxBMdltwnRE.PRN"
  RptFile# = FreeFile
  Open File$ For Output As #RptFile
  
  'Must Calc Late Fee Here
  frmVATaxShowPctComp.Label1 = "Creating Real Bills"
  frmVATaxShowPctComp.Show , Me
  frmVATaxShowPctComp.cmdCancel.Visible = False
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  
  GoSub LoadHeaders
  
  Tab1 = 44 - Len(TownName) / 2
  Tab2 = 44 - Len(Add1) / 2
  Tab3 = 44 - Len(Add3) / 2
  
  If OptMort.Value = True Then '12/6/06
    OpenMortIdxFile MRHandle, NumOfMRRecs
    NumOfTBRecs = NumOfMRRecs
  ElseIf optZip.Value = True Then '12/6/06
    OpenRZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  BillNo = fpDblSnglStartRealBill.Value
  
  For x = 1 To NumOfTBRecs
    If OptMort.Value = True Then '12/6/06
      Get MRHandle, x, MortRec
      CustArr = MortRec.TaxBillRec
    ElseIf optZip.Value = True Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TaxBill
    Get TCHandle, TaxBill.CustRec, TaxCust
    CustName$ = QPTrim$(TaxCust.CustName)
    If TaxBill.BillNumber < 0 Then GoTo NotThisOne
      'Must Calc Late Fee Here
      TotOpt = OldRound(TaxBill.OptRevTax1 + TaxBill.OptRevTax2 + TaxBill.OptRevTax3)
      Print #RptFile, "                                R E A L   E S T A T E"
      Print #RptFile, "                                 T A X   N O T I C E"
      Print #RptFile, Tab(Tab1); TownName
      Print #RptFile, Tab(Tab2); Add1
      Print #RptFile, Tab(Tab3); Add3
      Print #RptFile,
      Print #RptFile, "            VALUATION AMOUNT: "; Using$("$##,###,###.00", OldRound#(TaxBill.RealValue + TaxBill.BldgValue));
      Print #RptFile, Tab(50); "ACCT. #: "; CStr(TaxBill.CustRec)
      Print #RptFile, "                   EXEMPTION: "; Using$("$##,###,###.00", TaxBill.ExptValue);
      Print #RptFile, Tab(50); "PIN. #: "; QPTrim$(TaxBill.RealPin)
      Print #RptFile, "         LATE PENALTY AMOUNT: "; Using$("$##,###,###.00", TaxBill.LateTaxDue);
      Print #RptFile, Tab(50); "RECPT #: "; Using$("#####0", TaxBill.BillNumber)
      If TotOpt > 0 Then
        Print #RptFile, "              TAX AMOUNT DUE: "; Using$("$##,###,###.00", TaxBill.TotalBillDue) + "*";
      Else
        Print #RptFile, "              TAX AMOUNT DUE: "; Using$("$##,###,###.00", TaxBill.TotalBillDue);
      End If
      Print #RptFile, Tab(50); "TAX RATE %: "; Using$("#0.0000", RealTaxRate#)
      Print #RptFile, Tab(50); "TAX YEAR: "; CStr(TaxBill.TaxYear)
      Print #RptFile, Tab(50); "DUE DATE: "; DueDate$
      If TotOpt > 0 Then
        Print #RptFile, Tab(11); Left$(QPTrim$(CustName$), 45); Tab(50); "*PLUS " + QPTrim$(Using$("$##,##0.00", TotOpt)) + " IN ADDED TAXES."
      Else
        Print #RptFile, Tab(11); Left$(QPTrim$(CustName$), 45)
      End If
      Print #RptFile, Tab(11); Left$(QPTrim$(TaxBill.CustAdd1), 35)
      Print #RptFile, Tab(11); Left$(QPTrim$(TaxBill.CustAdd2), 35)
      Print #RptFile, Tab(11); QPTrim$(TaxBill.CustAdd3); " "; QPTrim$(TaxBill.CustZip)
      Print #RptFile,
      Print #RptFile,
      Print #RptFile, Tab(31); "T H A N K   Y O U ! ! !"
      Print #RptFile,
      Print #RptFile, "~"
      TaxBill.BillPrinted = True
      TaxBill.Comment = QPTrim$(fptxtComment.Text)
      TaxBill.BillNumber = BillNo
      BillNo = BillNo + 1
      Put TBHandle, CustArr, TaxBill

NotThisOne:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  
  Close
  Call MakeFile
  ViewPrint File$, "Real Property Tax Bills", True
  Exit Sub
  
LoadHeaders:
  WorkName = ""
  WordLen = Len(TownName)
  For y = 1 To WordLen
    ch = Mid(TownName, y, 1)
    WorkName = WorkName + ch + " "
  Next y
  TownName = WorkName
  
  WorkName = ""
  WordLen = Len(Add1)
  For y = 1 To WordLen
    ch = Mid(Add1, y, 1)
    WorkName = WorkName + ch + " "
  Next y
  Add1 = WorkName
  
  WorkName = ""
  WordLen = Len(Add2)
  For y = 1 To WordLen
    ch = Mid(Add2, y, 1)
    WorkName = WorkName + ch + " "
  Next y
  Add2 = WorkName
  
  WorkName = ""
  WordLen = Len(Add3)
  For y = 1 To WordLen
    ch = Mid(Add3, y, 1)
    WorkName = WorkName + ch + " "
  Next y
  Add3 = WorkName
  
  Return
    
End Sub

Private Sub PrintPersMiddletownBill()
  Dim TaxBill As VAPPTaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim BillInfo As VAPPTaxBillInfoType
  Dim BIHandle As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long, PersTaxRate#
  Dim PYear As Integer, PYearStr$
  Dim File$, LC As Integer
  Dim CustName$, WhatYear As Integer
  Dim RptFile#, WhatPers&
  Dim CarCount As Integer
  Dim TownName$, Add1$, Add2$, Add3$
  Dim PHandle As Integer, PPTRAVal#
  Dim NumOfPRecs As Long, PPTRADiscount#
  Dim PersRec As PersonalRecType
  Dim VehDesc$, PrnCnt As Integer
  Dim TaxAmt#, LCnt As Integer
  Dim TotOpt As Double
  Dim Tab1 As Integer, Tab2 As Integer, Tab3 As Integer, Tab4 As Integer
  Dim DueDate$, WorkName$
  Dim ZipRec As BillPrintPZipIdxType, CustArr As Long
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long, BillNo As Long
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  TaxMasterRec.MaxVehTaxVal = OldRound(TaxMasterRec.MaxVehTaxVal)
  TownName = QPTrim$(TaxMasterRec.Name)
  Add1$ = QPTrim$(TaxMasterRec.Add1)
  Add2$ = QPTrim$(TaxMasterRec.Add2)
  Add3$ = QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TownState) + " " + QPTrim$(TaxMasterRec.Zip)
  
  OpenPersPropFile PHandle, NumOfPRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenPersTaxBillFile TBHandle, NumOfTBRecs
  OpenPersBillInfoFile BIHandle
  Get BIHandle, 1, BillInfo
  Close BIHandle
  File$ = StartPath$ + "/TxMdltwnPP.PRN"
  RptFile# = FreeFile
  Open File$ For Output As #RptFile
  
  frmVATaxShowPctComp.Label1 = "Creating Personal Bills"
  frmVATaxShowPctComp.Show , Me
  frmVATaxShowPctComp.cmdCancel.Visible = False
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  PersTaxRate# = fpDblSnglPersRate
  Tab1 = 40 - Len(TownName) / 2
  Tab2 = 40 - Len(Add1) / 2
  Tab3 = 40 - Len(Add3) / 2
  
  If optZip.Value = True Then '12/6/06
    OpenPZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  BillNo = fpDblSnglStartPersBill.Value
  For x = 1 To NumOfTBRecs
    If optZip.Value = True Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TaxBill
    If TaxBill.BillNumber < 0 Then GoTo Natta
    Get TCHandle, TaxBill.CustRec, TaxCust
    WhatYear = TaxBill.TaxYear
    DueDate$ = MakeRegDate(TaxBill.DueDate)
    PrnCnt = PrnCnt + 1
    GoSub PrintIt
    TaxBill.BillPrinted = True
    TaxBill.Comment = fptxtComment.Text
    TaxBill.Comment2 = fptxtComment2.Text
    TaxBill.BillNumber = BillNo
    BillNo = BillNo + 1
    Put TBHandle, CustArr, TaxBill
Natta:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  
  Close
  Call MakeFile
  
  ViewPrint File$, "Personal Property Tax Bills", True
  Exit Sub

PrintIt:
  CustName$ = RTrim$(TaxCust.CustName)
  CustName$ = LTrim$(CustName$)
  TotOpt = OldRound(TaxBill.OptRevTax1 + TaxBill.OptRevTax2 + TaxBill.OptRevTax3)
  'Must Calc Late Fee Here
  
  Print #RptFile,
  Print #RptFile, Tab(Tab1); TownName
  Print #RptFile, Tab(Tab2); Add1
  Print #RptFile, Tab(Tab3); Add3
  Print #RptFile, Tab(27); "PERSONAL PROPERTY TAX BILL"
  Print #RptFile, Tab(30);
  For LC = 6 To 8
    Print #RptFile, " "
  Next
  Print #RptFile, Tab(10); "ACCT # "; Using$("######0", TaxBill.CustRec);
  Print #RptFile, Tab(63); "BILL # "; Using$("######0", TaxBill.BillNumber)
  Print #RptFile, Tab(10); Left$(CustName$, 25);
  Print #RptFile, Tab(63); "TAX YEAR: "; CStr(WhatYear)
  Print #RptFile, Tab(10); Left$(QPTrim$(TaxBill.CustAdd1), 25);
  Print #RptFile, Tab(63); "TAX RATE: "; Using("#0.##0", PersTaxRate#) + "%"
  Print #RptFile, Tab(10); Left$(QPTrim$(TaxBill.CustAdd2), 25)
  Print #RptFile, Tab(10); QPTrim$(TaxBill.CustAdd3); " "; QPTrim(TaxBill.CustZip)
  For LC = 14 To 17
    Print #RptFile, " "
  Next
  Print #RptFile, Tab(37); "PROPERTY"; Tab(51); "   TAX"; Tab(61); "   PPTRA"
  Print #RptFile, Tab(38); "  VALUE"; Tab(51); "AMOUNT"; Tab(61); "DISCOUNT"; Tab(71); "TOTAL DUE"
  'Line 23 Starts Here
  Print #RptFile, Tab(2); "Personal Property";
  Print #RptFile, Tab(38); Using$("###,##0", TaxBill.PersValue);
  Print #RptFile, Tab(47); Using$("###,##0.00", TaxBill.PersTaxDue);
  Print #RptFile, Tab(59); Using("###,##0.00", TaxBill.PPTRADiscnt);
  If TotOpt > 0 Then
    Print #RptFile, Tab(70); Using("###,##0.00", TaxBill.TotalBillDue) + "*"
  Else
    Print #RptFile, Tab(70); Using("###,##0.00", TaxBill.TotalBillDue)
  End If
  CarCount = 0
  Print #RptFile,
  Print #RptFile, " PPTRA Information"

  WhatPers& = TaxCust.FirstPersRec
  Do
    Get PHandle, WhatPers&, PersRec
    If PersRec.PPTRAYN = "Y" Then
      If OldRound#(PersRec.PersVal) > TaxMasterRec.MaxVehTaxVal Then
        PPTRAVal# = TaxMasterRec.MaxVehTaxVal
      Else
        PPTRAVal# = OldRound#(PersRec.PersVal)
      End If
      If PPTRAVal# <= TaxMasterRec.MinVehTaxVal Then
        PPTRADiscount# = OldRound#((OldRound#(PPTRAVal# / 100) * PersTaxRate#))
        PPTRADiscount# = OldRound(PPTRADiscount# / TaxMasterRec.MultiYear)
      Else
        PPTRADiscount# = OldRound#((OldRound#((PPTRAVal# / 100) * (TaxMasterRec.PPTRADisc / 100)) * PersTaxRate#))
        PPTRADiscount# = OldRound(PPTRADiscount# / TaxMasterRec.MultiYear)
      End If

      VehDesc$ = " VIN# " + QPTrim$(PersRec.Vin)
      VehDesc$ = QPTrim$(VehDesc$)
      TaxAmt# = (PersTaxRate# / 100) * PersRec.PersVal
      TaxAmt# = OldRound(TaxAmt# / TaxMasterRec.MultiYear)
      
      Print #RptFile, Tab(2); "*" + VehDesc$;
      Print #RptFile, Tab(38); Using("###,##0", PersRec.PersVal);
      Print #RptFile, Tab(47); Using("###,##0.00", TaxAmt#);
      Print #RptFile, Tab(59); Using("###,##0.00", PPTRADiscount#)
      CarCount = CarCount + 1
    End If

    WhatPers& = PersRec.NextRec
  Loop While WhatPers& > 0

  ' Finish the bill up here
  Print #RptFile, ""
  Print #RptFile, ""
  'Put Late Here and Add to Total
  If TotOpt > 0 Then
    Print #RptFile, "*Includes " + QPTrim$(Using("$##,##0.00", TotOpt)) + " in added taxes."; Tab(40); "Total Tax Due by "; DueDate$;
    Print #RptFile, Tab(69); Using("$###,##0.00", TaxBill.TotalBillDue) + "*"
  Else
    Print #RptFile, Tab(40); "Total Tax Due by "; DueDate$;
    Print #RptFile, Tab(69); Using("$###,##0.00", TaxBill.TotalBillDue)
  End If
  Print #RptFile,
  Print #RptFile,
  If CarCount > 0 Then
    Print #RptFile, " The Personal Property Relief Act provides that the tax on the first"
    Print #RptFile, " " + Using$("$##,##0.00", TaxMasterRec.MaxVehTaxVal) + " of value of your personal car, motorcycle and pickup  "
    Print #RptFile, " or panel truck under 7,501 pounds, which is a qualifying vehicle, has been"
    Print #RptFile, " reduced by " + Using$("#0.00", TaxMasterRec.PPTRADisc) + "% this year. If your qualifying vehicle's value is"
    Print #RptFile, " " + Using$("$#,##0.00", TaxMasterRec.MinVehTaxVal) + " or less, your tax has been eliminated. These reductions are"
    Print #RptFile, " based on the local tax rates in effect on July 1 or August 1, 1997,"
    Print #RptFile, " whichever was higher. Please contact the Town Office with any questions."
    Print #RptFile, ""
    Print #RptFile, ""
  End If

  Print #RptFile,
  Print #RptFile, " BN"; Using("#####", PrnCnt)
  Print #RptFile, Chr$(12);

  Return

End Sub

Private Sub PrintCedarBluffPersonal()
  Dim TaxBill As VAPPTaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim BillInfo As VAPPTaxBillInfoType
  Dim BIHandle As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long, PersTaxRate#
  Dim PYear As Integer, PYearStr$
  Dim File$, LC As Integer
  Dim CustName$, WhatYear As Integer
  Dim RptFile#, WhatPers&
  Dim CarCount As Integer
  Dim TownName$, Add1$, Add2$, Add3$
  Dim PHandle As Integer, PPTRAVal#
  Dim NumOfPRecs As Long, PPTRADiscount#
  Dim PersRec As PersonalRecType
  Dim VehDesc As String * 30, PrnCnt As Integer
  Dim TaxAmt#, LCnt As Integer
  Dim TotOth As Double
  Dim DueDate$, OptTot As Double
  Dim ZipRec As BillPrintPZipIdxType, CustArr As Long
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long, BillNo As Long
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  TaxMasterRec.MaxVehTaxVal = OldRound(TaxMasterRec.MaxVehTaxVal)
  TownName = QPTrim$(TaxMasterRec.Name)
  Add1$ = QPTrim$(TaxMasterRec.Add1)
  Add2$ = QPTrim$(TaxMasterRec.Add2)
  Add3$ = QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TownState) + " " + QPTrim$(TaxMasterRec.Zip)
  
  OpenPersPropFile PHandle, NumOfPRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenPersTaxBillFile TBHandle, NumOfTBRecs
  OpenPersBillInfoFile BIHandle
  Get BIHandle, 1, BillInfo
  Close BIHandle
  File$ = StartPath$ + "/TxCdrBluffPP.PRN"
  RptFile# = FreeFile
  Open File$ For Output As #RptFile
  
  frmVATaxShowPctComp.Label1 = "Creating Personal Bills"
  frmVATaxShowPctComp.Show , Me
  frmVATaxShowPctComp.cmdCancel.Visible = False
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  PersTaxRate# = fpDblSnglPersRate
  
  If Exist("PZipIdx.Dat") Then '12/6/06
    OpenPZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  BillNo = fpDblSnglStartPersBill.Value
  For x = 1 To NumOfTBRecs
    If optZip.Value = True Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TaxBill
    If TaxBill.BillNumber < 0 Then GoTo Natta
    Get TCHandle, TaxBill.CustRec, TaxCust
    WhatYear = TaxBill.TaxYear
    DueDate$ = MakeRegDate(TaxBill.DueDate)
    PrnCnt = PrnCnt + 1
    GoSub PrintIt
    TaxBill.BillPrinted = True
    TaxBill.Comment = fptxtComment.Text
    TaxBill.Comment2 = fptxtComment2.Text
    TaxBill.BillNumber = BillNo
    BillNo = BillNo + 1
    Put TBHandle, CustArr, TaxBill
Natta:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  
  Close
  Call MakeFile
  
  ViewPrint File$, "Personal Property Tax Bills", True
  Exit Sub

PrintIt:
   CustName$ = QPTrim$(TaxBill.CustName)
  'Must Calc Late Fee Here
   Print #RptFile, "~"; Tab(34); Str$(WhatYear); " PERSONAL PROPERTY"
   Print #RptFile, Tab(5); TownName$
   Print #RptFile, Tab(5); Add1$
   Print #RptFile, Tab(5); Add3$
   Print #RptFile, Tab(5); "   "
   Call InsertSSNDashes(TaxCust.CSSN)
   Print #RptFile, Tab(10); TaxCust.CSSN; Tab(65); "PP"; Using("##.###", PersTaxRate#)
   Print #RptFile, " "
   Print #RptFile, " "

  'Line 30 to 35 here to print vehicles
    CarCount = 0
    WhatPers& = TaxCust.FirstPersRec
    OptTot = OldRound(TaxBill.OptRevTax1 + TaxBill.OptRevTax2 + TaxBill.OptRevTax3)
    If OptTot > 0 Then
      Print #RptFile, "Other Taxes:"; Tab(64); Using("##,###.##", OptTot#)
    End If
    Do
      Get PHandle, WhatPers&, PersRec
'      PYear$ = Right$(PersRec.Desc5, 4)
      PYear = PersRec.TaxBillYear 'Val(PYear$)
'      If PYear > 0 And PYear <> WhatYear Then
'        'Do Not Process This Record
'      Else
        If Left$(PersRec.Desc5, 1) = "Y" Or PersRec.PPTRAYN = "Y" Then 'added PersRec.PPTRAYN on 11/26/07
          If OldRound#(PersRec.PersVal) > TaxMasterRec.MaxVehTaxVal Then
            PPTRAVal# = TaxMasterRec.MaxVehTaxVal
          Else
            PPTRAVal# = OldRound#(PersRec.PersVal)
          End If
          If PPTRAVal# <= TaxMasterRec.MinVehTaxVal Then
            PPTRADiscount# = OldRound#((OldRound#(PPTRAVal# / 100) * PersTaxRate#))
          Else
            PPTRADiscount# = OldRound#((OldRound#((PPTRAVal# / 100) * (TaxMasterRec.PPTRADisc / 100)) * PersTaxRate#))
          End If
        Else
          PPTRADiscount# = 0
        End If
'        VehDesc$ = QPTrim$(PersRec.Desc4) + " " + Left$(PersRec.DESC2, 22) + "   " + Left$(PersRec.Desc5, 1)
'        VehDesc$ = QPTrim$(VehDesc$)
        VehDesc$ = CStr(PersRec.ModYear) + " " + QPTrim$(PersRec.MakeMod) + " " + QPTrim$(PersRec.DESC1)
        VehDesc = Left(VehDesc, 28) + " " + PersRec.PPTRAYN 'changed from above at Cedar Bluff request on 11/26/07
        CarCount = CarCount + 1
        If PersRec.PersVal <= 0 Then
          TaxAmt# = (MHTaxRate# / 100) * PersRec.MHValue
          Print #RptFile, VehDesc$;
          Print #RptFile, Tab(33); Using("##,###,###", PersRec.MHValue);
          Print #RptFile, Tab(44); Using("###,###.##", TaxAmt#);
          Print #RptFile, Tab(54); Using("##,###.##", PPTRADiscount#);
          Print #RptFile, Tab(64); Using("##,###.##", TaxAmt#)
        Else
          TaxAmt# = (PersTaxRate# / 100) * PersRec.PersVal
          Print #RptFile, VehDesc$;
          Print #RptFile, Tab(33); Using("##,###,###", PersRec.PersVal);
          Print #RptFile, Tab(44); Using("###,###.##", TaxAmt#);
          Print #RptFile, Tab(54); Using("#####.##", PPTRADiscount#);
          Print #RptFile, Tab(64); Using("##,###.##", OldRound#(TaxAmt# - PPTRADiscount#))
        End If
'      End If
      If OptTot > 0 Then
        If CarCount = 4 Then CarCount = 5
      End If
      If (CarCount >= 5) And (PersRec.NextRec > 0) Then
        Print #RptFile, " "
        Print #RptFile, ""
        Print #RptFile, Tab(10); Using("#####", PrnCnt);
        Print #RptFile, Tab(36); DueDate$; Tab(66); "CONT'D"                 '; USING "$$#####,#.##"; TaxBill.TotalBillDue
        Print #RptFile,
        Print #RptFile, Tab(9); CustName$
        Print #RptFile, Tab(9); QPTrim$(TaxBill.CustAdd1)
        Print #RptFile, Tab(9); QPTrim$(TaxBill.CustAdd2)
        Print #RptFile, Tab(9); QPTrim$(TaxBill.CustAdd3); " "; TaxBill.CustZip
        Print #RptFile,
        Print #RptFile,
        Print #RptFile, "~"
        Print #RptFile, "~"
        Print #RptFile, Tab(5); TownName$
        Print #RptFile, Tab(5); Add1$
        Print #RptFile, Tab(5); Add2$
        Print #RptFile, Tab(5); Add3$;
        Print #RptFile, Tab(10); QPTrim$(TaxCust.CSSN)
        Print #RptFile, " "
        Print #RptFile, '"Vehicle Listing Continued ..."
        CarCount = 0
      ElseIf (CarCount >= 5) And (PersRec.NextRec <= 0) Then
        Print #RptFile, ""
      '  CarCount = 0
      End If
      WhatPers& = PersRec.NextRec
    Loop While WhatPers& > 0

  ' Finish the bill up here
    If CarCount < 5 Then
      For LCnt = CarCount To 5
        Print #RptFile, ""
      Next
    End If
'
    Print #RptFile, ""
    Print #RptFile, Tab(10); Using("#####", PrnCnt);
    Print #RptFile, Tab(36); DueDate$; Tab(62); Using("$###,###.##", TaxBill.TotalBillDue)
    Print #RptFile,
    Print #RptFile, Tab(9); CustName$
    Print #RptFile, Tab(9); TaxBill.CustAdd1
    Print #RptFile, Tab(9); TaxBill.CustAdd2
    Print #RptFile, Tab(9); QPTrim$(TaxBill.CustAdd3); " "; QPTrim$(TaxBill.CustZip)
    Print #RptFile,
    Print #RptFile,
    Print #RptFile, "~"
  Return

End Sub
Private Sub PrintCedarBluffReal()
  Dim TaxBill As VARETaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim BillInfo As VARETaxBillInfoType
  Dim BIHandle As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long, RealTaxRate#
  Dim File$, WordLen As Integer
  Dim CustName As String * 45
  Dim RptFile#, ch$, y As Integer
  Dim TownName$, Add1$, Add2$, Add3$
  Dim RHandle As Integer, PrnCnt As Integer
  Dim NumOfRRecs As Long, WhatYear As Integer
  Dim RealRec As PropertyRecType
  Dim TaxAmt#, Tab1 As Integer, Tab2 As Integer, Tab3 As Integer
  Dim DueDate$, WorkName$, OptTot As Double
  Dim ZipRec As BillPrintRZipIdxType, CustArr As Long
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long, BillNo As Long
  
  DueDate$ = fptxtRealDueDate
  RealTaxRate# = fpDblSnglRealRate
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  TownName = QPTrim$(TaxMasterRec.Name)
  Add1$ = QPTrim$(TaxMasterRec.Add1)
  Add2$ = QPTrim$(TaxMasterRec.Add2)
  Add3$ = QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TownState) + " " + QPTrim$(TaxMasterRec.Zip)

  OpenRealPropFile RHandle, NumOfRRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenRealTaxBillFile TBHandle, NumOfTBRecs
  OpenRealBillInfoFile BIHandle
  Get BIHandle, 1, BillInfo
  Close BIHandle
  File$ = StartPath$ + "/TxBCdrBluffRE.PRN"
  RptFile# = FreeFile
  Open File$ For Output As #RptFile

  OpenTaxCustFile TCHandle, NumOfTCRecs
  'Must Calc Late Fee Here
  frmVATaxShowPctComp.Label1 = "Creating Real Bills"
  frmVATaxShowPctComp.Show , Me
  frmVATaxShowPctComp.cmdCancel.Visible = False
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False

  If OptMort.Value = True Then '12/6/06
    OpenMortIdxFile MRHandle, NumOfMRRecs
    NumOfTBRecs = NumOfMRRecs
  ElseIf optZip.Value = True Then '12/6/06
    OpenRZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  
  BillNo = fpDblSnglStartRealBill.Value
  For x = 1 To NumOfTBRecs
    If OptMort.Value = True Then '12/6/06
      Get MRHandle, x, MortRec
      CustArr = MortRec.TaxBillRec
    ElseIf optZip.Value = True Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TaxBill
    Get TCHandle, TaxBill.CustRec, TaxCust
    If TaxBill.BillNumber < 0 Then GoTo NotThisOne
    CustName$ = QPTrim$(TaxBill.CustName)
    WhatYear = TaxBill.TaxYear
  'Must Calc Late Fee Here
    OptTot = OldRound(TaxBill.OptRevTax1 + TaxBill.OptRevTax2 + TaxBill.OptRevTax3)
    Print #RptFile, "~"
    Print #RptFile, Tab(50); CStr(WhatYear); Tab(78); Using("########", TaxBill.BillNumber)
    Print #RptFile,
    Print #RptFile, " "
    Print #RptFile, " "
    Print #RptFile, " "
    Print #RptFile, Tab(28); Using("#.##", RealTaxRate#);
    Print #RptFile, Tab(36); Using("###,###,###", TaxBill.RealValue);
    Print #RptFile, Tab(48); Using("##,###,###", TaxBill.BldgValue);
    Print #RptFile, Tab(61); Using("##,###,###", TaxBill.RealValue + TaxBill.BldgValue);
    Print #RptFile, Tab(75); Using("###,###.##", TaxBill.RealTaxDue);
    Print #RptFile, Tab(90); Using$("##0.00", fpDSMnthlyPen.Value) + "%"
    Print #RptFile, " "
    Print #RptFile, Tab(68); DueDate$
    Print #RptFile, QPTrim$(TaxBill.RDesc1)
    Print #RptFile, QPTrim$(TaxBill.RDesc2)
    Print #RptFile, ""
    Print #RptFile, ""
    Print #RptFile, ""
    Print #RptFile, Tab(7); "ACCT # "; CStr(TaxBill.CustRec)
    Print #RptFile, Tab(7); Left$(CustName$, 45)
    Print #RptFile, Tab(7); Left$(TaxBill.CustAdd1, 35)
    Print #RptFile, Tab(7); Left$(TaxBill.CustAdd2, 35)
    Print #RptFile, Tab(7); QPTrim$(TaxBill.CustAdd3); " "; TaxBill.CustZip
    Print #RptFile,
    If OptTot = 0 Then
      Print #RptFile, "BN"; Using("#####", PrnCnt)
    Else
      Print #RptFile, "BN"; Using("#####", PrnCnt); Tab(54); "Tax Due includes " + QPTrim$(Using$("$##,##0.00", OptTot)) + " in other taxes."
    End If
    Print #RptFile, "~"
    TaxBill.BillPrinted = True
    TaxBill.Comment = CStr(fpDSMnthlyPen.Text)
    BillNo = BillNo + 1
    TaxBill.BillNumber = BillNo
    Put TBHandle, CustArr, TaxBill

NotThisOne:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True

  Close
  Call MakeFile
  ViewPrint File$, "Real Property Tax Bills", True
  Exit Sub

End Sub

Private Sub PrintHalifaxStandardPersonal()
 'TAXPPSTD.BI
  Dim TaxBill As VAPPTaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim BillInfo As VAPPTaxBillInfoType
  Dim BIHandle As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  Dim PYear As Integer, PYearStr$
  Dim File$, LC As Integer
  Dim CustName$, WhatYear As Integer
  Dim RptFile#, WhatPers&
  Dim CarCount As Integer
  Dim TownName$, Add1$, Add2$, Add3$
  Dim PHandle As Integer, PPTRAVal#
  Dim NumOfPRecs As Long, PPTRADiscount#
  Dim PersRec As PersonalRecType
  Dim VehDesc$, PERC!
  Dim TaxAmt#, LCnt As Integer
  Dim MultiYear As Integer
  Dim TotOth As Double
  Dim PrintComments As String
  Dim ZipRec As BillPrintPZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim FF$, CustArr As Long, BillNo As Long
  
  FF$ = Chr(12)
  If QPTrim$(fptxtComment.Text) <> "" Or QPTrim$(fptxtComment2.Text) <> "" Then
    If InStr(fpcmbCommentPlace.Text, "LEFT") Then
      PrintComments = "L"
    ElseIf InStr(fpcmbCommentPlace.Text, "RIGHT") Then
      PrintComments = "R"
    Else
      PrintComments = "N"
    End If
  End If
  
  WhatYear = PersYear
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  MultiYear = TaxMasterRec.MultiYear
  TaxMasterRec.MaxVehTaxVal = OldRound(TaxMasterRec.MaxVehTaxVal)
  TownName = QPTrim$(TaxMasterRec.Name)
  Add1$ = QPTrim$(TaxMasterRec.Add1)
  Add2$ = QPTrim$(TaxMasterRec.Add2)
  Add3$ = QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TownState) + " " + QPTrim$(TaxMasterRec.Zip)
  
  If WhatYear = 1999 Then PERC! = 27.5
  If WhatYear = 2000 Then PERC! = 47.5
  If WhatYear >= 2001 Then PERC! = TaxMasterRec.PPTRADisc
  
  OpenPersPropFile PHandle, NumOfPRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenPersTaxBillFile TBHandle, NumOfTBRecs
  OpenPersBillInfoFile BIHandle
  Get BIHandle, 1, BillInfo
  Close BIHandle
  File$ = StartPath$ + "/TxBStandPP.PRN"
  RptFile# = FreeFile
  Open File$ For Output As #RptFile
  frmVATaxShowPctComp.Label1 = "Creating Personal Bills"
  frmVATaxShowPctComp.Show , Me
  frmVATaxShowPctComp.cmdCancel.Visible = False
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  
  If optZip.Value = True Then '12/6/06
    OpenPZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  BillNo = fpDblSnglStartPersBill.Value
  For x = 1 To NumOfTBRecs
    If optZip.Value = True Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TaxBill
    If TaxBill.BillNumber < 0 Then GoTo Natta
    Get TCHandle, TaxBill.CustRec, TaxCust
    GoSub PrintIt
    TaxBill.BillPrinted = True
    TaxBill.Comment = fptxtComment.Text
    TaxBill.Comment2 = fptxtComment2.Text
    TaxBill.CommentPlace = fpcmbCommentPlace.Text
    BillNo = BillNo + 1
    TaxBill.BillNumber = BillNo
    Put TBHandle, CustArr, TaxBill
Natta:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  Print #RptFile, FF$
  Close
  
  ViewPrint File$, "Personal Property Tax Bills", True
  
  Exit Sub
  
PrintIt:
  CustName$ = QPTrim$(TaxCust.CustName)
  Print #RptFile, "~"
  Print #RptFile, Tab(63); "TAX YEAR: "; WhatYear
  Print #RptFile, Tab(75); Using$("#####", TaxBill.BillNumber)
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, Tab(5); " "
  Print #RptFile, Tab(5); " "
  Print #RptFile, Tab(5); " "
  Print #RptFile, Tab(5); " "
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, " "
  Print #RptFile, Tab(5); "Acct # "; Using$("#####0", TaxBill.CustRec)
  Print #RptFile, Tab(5); CustName$
  Print #RptFile, Tab(5); QPTrim$(TaxBill.CustAdd1)
  Print #RptFile, Tab(5); QPTrim$(TaxBill.CustAdd2)
  Print #RptFile, Tab(5); QPTrim$(TaxBill.CustAdd3) + " " + QPTrim$(TaxBill.CustZip)
'  For LC = 18 To 21
  For LC = 19 To 21 'added
   Print #RptFile, " "
  Next LC
  Print #RptFile, Tab(37); "PROP VALUE"; Tab(50); "TAX AMOUNT"; Tab(62); "PPTRA DIS"; Tab(72); "TOTAL DUE"
  Print #RptFile, " "
 'Line 24 Starts Here
  Print #RptFile, "Personal Property"; Tab(32); Using$("#.00", PersTaxRate#);
   Print #RptFile, Tab(37); Using$("#####0.00", TaxBill.PersValue);
   Print #RptFile, Tab(51); Using$("#####0.00", TaxBill.PersTaxDue); ' - TaxBill.OverPayAmt);
   Print #RptFile, Tab(63); Using$("####0.00", TaxBill.PPTRADiscnt);
   Print #RptFile, Tab(72); Using$("#####0.00", OldRound(TaxBill.PersTaxDue - TaxBill.PPTRADiscnt)) ' - TaxBill.OverPayAmt))
   
  Print #RptFile, "Machinery/Tools"; Tab(32); Using$("#.00", MTTaxRate#);
   Print #RptFile, Tab(37); Using$("#####0.00", TaxBill.MTValue);
   Print #RptFile, Tab(51); Using$("#####0.00", TaxBill.MTTaxDue);
   Print #RptFile, Tab(72); Using$("#####0.00", TaxBill.MTTaxDue)
  Print #RptFile, "Farm Equipment";
   Print #RptFile, Tab(32); Using("#.00", FETaxRate#);
   Print #RptFile, Tab(37); Using$("#####0.00", TaxBill.FEValue);
   Print #RptFile, Tab(51); Using$("#####0.00", TaxBill.FETaxDue);
   Print #RptFile, Tab(72); Using$("#####0.00", TaxBill.FETaxDue)
  Print #RptFile, "Mobile Homes";
   Print #RptFile, Tab(32); Using$("#.00", MHTaxRate#);
   Print #RptFile, Tab(37); Using$("#####0.00", TaxBill.MHValue);
   Print #RptFile, Tab(51); Using$("#####0.00", TaxBill.MHTaxDue);
   Print #RptFile, Tab(72); Using$("#####0.00", TaxBill.MHTaxDue)
  Print #RptFile, "Merchant Capital";
   Print #RptFile, Tab(32); Using$("#.00", MCTaxRate#);
   Print #RptFile, Tab(37); Using$("#####0.00", TaxBill.MCValue);
   Print #RptFile, Tab(51); Using$("#####0.00", TaxBill.MCTaxDue);
   Print #RptFile, Tab(72); Using$("#####0.00", TaxBill.MCTaxDue)
   TotOth = OldRound(TaxBill.OptRevTax1 + TaxBill.OptRevTax2 + TaxBill.OptRevTax3)
   If TaxBill.OverPayAmt > 0 And TotOth = 0 Then '6/22/06
     Print #RptFile, " PPTRA Vehicle Information"; Tab(43); "** Applied Credit Of: " + QPTrim$(Using$("$##,##0.00", TaxBill.OverPayAmt)) + " **"
   ElseIf TaxBill.OverPayAmt > 0 And TotOth > 0 Then '6/22/06
     Print #RptFile, " PPTRA Vehicle Information"; Tab(30); "* Applied Credit Of: " + QPTrim$(Using$("$##,##0.00", TaxBill.OverPayAmt)) + " *"; Tab(60); "Other Taxes: " + Using$("#,##0.00", TotOth)
  ElseIf TaxBill.OverPayAmt = 0 And TotOth > 0 Then '6/22/06
     Print #RptFile, " PPTRA Vehicle Information"; Tab(60); "Other Taxes: " + Using$("#,##0.00", TotOth)
   Else
     Print #RptFile, " PPTRA Vehicle Information"
   End If
 'Line 30 to 35 here to print vehicles
  CarCount = 0
  WhatPers& = TaxCust.FirstPersRec
  Do
    Get PHandle, WhatPers&, PersRec
    PYearStr$ = CStr(PersRec.TaxBillYear)
    PYear = Val(PYearStr$)
    If PYear > 0 And PYear <> WhatYear Then
      Return
        'Do Not Process This Record
    End If
    If PersRec.PPTRAYN = "Y" Then
      If OldRound#(PersRec.PersVal) > TaxMasterRec.MaxVehTaxVal Then
        If TaxMasterRec.MultiYear <> 0 Then
          PersRec.PersVal = OldRound(PersRec.PersVal)
        End If
        PPTRAVal# = TaxMasterRec.MaxVehTaxVal
      Else
        PPTRAVal# = OldRound#(PersRec.PersVal)
      End If

      If PPTRAVal# <= TaxMasterRec.MinVehTaxVal Then
        PPTRADiscount# = OldRound#((OldRound#(PPTRAVal# / 100) * PersTaxRate#))
      Else
        PPTRADiscount# = OldRound#((OldRound#((PPTRAVal# / 100) * (PERC! / 100)) * PersTaxRate#))
      End If

      VehDesc$ = " VIN# " + QPTrim$(PersRec.Vin)
      VehDesc$ = QPTrim$(VehDesc$)
      TaxAmt# = OldRound((PersTaxRate# / 100) * PersRec.PersVal)
      TaxBill.PersTaxDue = TaxBill.PersTaxDue
      TaxBill.PPTRADiscnt = TaxBill.PPTRADiscnt
      Print #RptFile, "*" + VehDesc$;
      Print #RptFile, Tab(37); Using$("#####0.00", PersRec.PersVal) ';
      CarCount = CarCount + 1
    End If
    
    If CarCount >= 6 Then
      Print #RptFile, ""
      Print #RptFile, Tab(48); "Total Tax Due ";
      Print #RptFile, Using$("$#######0.00", OldRound(TaxBill.TotalBillDue - TaxBill.OverPayAmt))
      Print #RptFile, Tab(48); "Tax Due Date: " + MakeRegDate(BillInfo.DueDate)  'DueDate$
      Print #RptFile, ""
      Print #RptFile,
      Print #RptFile,
      Print #RptFile,
      Print #RptFile, "BN"; Using("####0", x) ' PrnCnt
      Print #RptFile, "~"

      Print #RptFile, "~"
      Print #RptFile, Tab(62); "TAX YEAR: "; WhatYear
      Print #RptFile, Tab(75); Using$("####0", TaxBill.BillNumber)
      Print #RptFile, " "
      Print #RptFile, " "
      Print #RptFile, Tab(5); TownName$
      Print #RptFile, Tab(5); Add1$
      Print #RptFile, Tab(5); Add2$
      Print #RptFile, Tab(5); Add3$
      Print #RptFile, " "
      Print #RptFile, " "
      Print #RptFile, " "
      Print #RptFile, Tab(5); "Acct # " + Using$("####0", TaxBill.CustRec) + " Vehicle Listing Cont'd"
      Print #RptFile, Tab(5); CustName$
      Print #RptFile, Tab(5); QPTrim$(TaxBill.CustAdd1)
      Print #RptFile, Tab(5); QPTrim$(TaxBill.CustAdd2)
      Print #RptFile, Tab(5); QPTrim$(TaxBill.CustAdd3) + " " + QPTrim$(TaxBill.CustZip)
      For LC = 18 To 21
       Print #RptFile, " "
      Next LC
      Print #RptFile, Tab(37); "PROP VALUE"; Tab(50); "TAX AMOUNT"; Tab(62); "PPTRA DIS."; Tab(73); "TOTAL DUE"
      Print #RptFile, " "
      Print #RptFile, "Vehicle Listing Continued ..."
      Print #RptFile, ""
      Print #RptFile, ""
      Print #RptFile, " PPTRA Vehicle Information"
      Print #RptFile, ""
      Print #RptFile, ""
      CarCount = 0
    End If
    WhatPers& = PersRec.NextRec
  Loop While WhatPers& > 0

   ' Finish the bill up here
  If CarCount < 6 Then
    For LCnt = CarCount To 6: Print #RptFile, "": Next LCnt
  End If
'  Print #RptFile,
  If PrintComments = "L" Then 'comments left
    Print #RptFile,
    Print #RptFile,
    Print #RptFile, fptxtComment.Text; Tab(48); "Total Tax Due ";
    Print #RptFile, Using$("$#######0.00", OldRound(TaxBill.TotalBillDue - TaxBill.OverPayAmt))
    Print #RptFile, fptxtComment2.Text; Tab(48); "Tax Due Date: " + MakeRegDate(BillInfo.DueDate)
  ElseIf PrintComments = "R" Then 'comments right
    Print #RptFile, Tab(48); "Total Tax Due ";
    Print #RptFile, Using$("$#######0.00", OldRound(TaxBill.TotalBillDue - TaxBill.OverPayAmt))
    Print #RptFile, Tab(48); "Tax Due Date: " + MakeRegDate(BillInfo.DueDate)
    Print #RptFile, Tab(48); fptxtComment.Text
    Print #RptFile, Tab(48); fptxtComment2.Text
  Else
    Print #RptFile, Tab(48); "Total Tax Due ";
    Print #RptFile, Using$("$#######0.00", OldRound(TaxBill.TotalBillDue - TaxBill.OverPayAmt))
    Print #RptFile, Tab(48); "Tax Due Date: " + MakeRegDate(BillInfo.DueDate)
  End If
  Print #RptFile, " " 'added
  Print #RptFile, " " 'added
  Print #RptFile, " " 'added
  Print #RptFile, " " 'added
  Print #RptFile, "BN"; Using$("####0", x)
  Print #RptFile, "~"
  
  Call MakeFile
  Return

End Sub

Private Sub PrintHalifaxStandardReal()
 'checked OK against mask (TAXREMSK.DAT) on 10/21/2005
 'STANDARD REAL ESTATE BILL FORMAT AS SOLD BY SOUTHERN SOFTWARE
 'TAXRESTD.BI
  Dim TaxBill As VARETaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim BillInfo As VARETaxBillInfoType
  Dim BIHandle As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long, RealTaxRate#
  Dim RYear As Integer, RYearStr$
  Dim File$, LC As Integer
  Dim CustName As String * 45, WhatYear As Integer
  Dim RptFile#, WhatReal&
  Dim CarCount As Integer
  Dim TownName$, Add1$, Add2$, Add3$
  Dim RHandle As Integer
  Dim NumOfRRecs As Long
  Dim RealRec As PropertyRecType
  Dim TaxAmt#, LCnt As Integer
  Dim ThisDesc As String * 28
  Dim TotOth As Double
  Dim PrintComments As String
  Dim FF$, CustArr As Long
  Dim ZipRec As BillPrintRZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim MortRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long, BillNo As Long
  
  FF$ = Chr(12)
  
  If QPTrim$(fptxtComment.Text) <> "" Or QPTrim$(fptxtComment2.Text) <> "" Then
    If InStr(fpcmbCommentPlace.Text, "LEFT") Then
      PrintComments = "L"
    ElseIf InStr(fpcmbCommentPlace.Text, "RIGHT") Then
      PrintComments = "R"
    Else
      PrintComments = "N"
    End If
  End If
  
  RealTaxRate# = fpDblSnglRealRate
  WhatYear = RealYear
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  OpenRealPropFile RHandle, NumOfRRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenRealTaxBillFile TBHandle, NumOfTBRecs
  OpenRealBillInfoFile BIHandle
  Get BIHandle, 1, BillInfo
  Close BIHandle
  File$ = StartPath$ + "/TxBStandRE.PRN"
  RptFile# = FreeFile
  Open File$ For Output As #RptFile
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  'Must Calc Late Fee Here
  frmVATaxShowPctComp.Label1 = "Creating Real Bills"
  frmVATaxShowPctComp.Show , Me
  frmVATaxShowPctComp.cmdCancel.Visible = False
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False

  If OptMort.Value = True Then '12/6/06
    OpenMortIdxFile MRHandle, NumOfMRRecs
    NumOfTBRecs = NumOfMRRecs
  ElseIf optZip.Value = True Then '12/6/06
    OpenRZipIdxFile ZHandle, NumOfZRecs
    NumOfTBRecs = NumOfZRecs
  End If
  BillNo = fpDblSnglStartRealBill.Value
  For x = 1 To NumOfTBRecs
    If OptMort.Value = True Then '12/6/06
      Get MRHandle, x, MortRec
      CustArr = MortRec.TaxBillRec
    ElseIf optZip.Value = True Then '12/6/06
      Get ZHandle, x, ZipRec
      CustArr = ZipRec.TaxBillRec
    Else
      CustArr = x '12/6/06
    End If
    Get TBHandle, CustArr, TaxBill
    Get TCHandle, TaxBill.CustRec, TaxCust
    CustName$ = QPTrim$(TaxCust.CustName)
    If TaxBill.BillNumber < 0 Then GoTo NotThisOne
      Print #RptFile, "~"
      Print #RptFile, Tab(64); "TAX YEAR: "; WhatYear
      Print #RptFile, Tab(75); Using$("#####", TaxBill.BillNumber)
      Print #RptFile, " "
      Print #RptFile, " "
      Print #RptFile, " "
      Print #RptFile, " "
      Print #RptFile, " "
      Print #RptFile, " "
      Print #RptFile, " "
      Print #RptFile, " "
      Print #RptFile, " "
      Print #RptFile, " "
      Print #RptFile, Tab(5); "PIN:  " + QPTrim$(TaxBill.RealPin)
      Print #RptFile, Tab(5); "ACCT: " + Using$("#####", TaxBill.CustRec)
      Print #RptFile, Tab(5); CustName$
      Print #RptFile, Tab(5); Left$(TaxBill.CustAdd1, 35)
      Print #RptFile, Tab(5); Left$(TaxBill.CustAdd2, 35)
      Print #RptFile, Tab(5); QPTrim$(TaxBill.CustAdd3) + " " + TaxBill.CustZip

      For LC = 19 To 20 'made 18 = 19
        Print #RptFile, " "
      Next LC
      Print #RptFile, Tab(41); "LAND"; Tab(51); "BUILDING"; Tab(61); "NET TOTAL"; Tab(72); "TOTAL DUE"
      Print #RptFile, " "
      'Line 23 Starts Here
      ThisDesc = QPTrim$(TaxBill.RDesc1)
      Print #RptFile, ThisDesc; 'QPTrim$(TaxBill.RDesc1);
      Print #RptFile, Tab(30); Using("#0.00", RealTaxRate#);
      If TaxBill.RealValue > TaxBill.ExptValue Then
        Print #RptFile, Tab(37); Using("######0.00", (TaxBill.RealValue - TaxBill.ExptValue)); ' - RTaxBill.PersValue));
        Print #RptFile, Tab(50); Using("#####0.00", TaxBill.BldgValue);
      ElseIf TaxBill.BldgValue > TaxBill.ExptValue Then
        Print #RptFile, Tab(37); Using("######0.00", TaxBill.RealValue); ' - RTaxBill.PersValue));
        Print #RptFile, Tab(50); Using("#####0.00", (TaxBill.BldgValue - TaxBill.ExptValue));
      ElseIf TaxBill.RealValue + TaxBill.BldgValue > TaxBill.ExptValue Then
        Print #RptFile, Tab(37); Using("######0.00", TaxBill.RealValue - (TaxBill.ExptValue * (TaxBill.RealValue / (TaxBill.RealValue + TaxBill.BldgValue)))); ' - RTaxBill.PersValue));
        Print #RptFile, Tab(50); Using("#####0.00", TaxBill.BldgValue - (TaxBill.ExptValue * (TaxBill.BldgValue / (TaxBill.RealValue + TaxBill.BldgValue)))); ' - RTaxBill.PersValue));
      Else
        Print #RptFile, Tab(37); Using("######0.00", TaxBill.RealValue); ' - RTaxBill.PersValue));
        Print #RptFile, Tab(50); Using("#####0.00", TaxBill.BldgValue);
      End If
      Print #RptFile, Tab(61); Using("#####0.00", OldRound(TaxBill.RealValue + TaxBill.BldgValue - TaxBill.ExptValue));
      Print #RptFile, Tab(71); Using("######0.00", OldRound(TaxBill.TotalBillDue)) ' - TaxBill.OverPayAmt))
      Print #RptFile, QPTrim$(TaxBill.RDesc2)
      TotOth = OldRound(TaxBill.OptRevTax1 + TaxBill.OptRevTax2 + TaxBill.OptRevTax3 + TaxBill.LateTaxDue)
      If TaxBill.OverPayAmt > 0 And TotOth > 0 Then
        Print #RptFile, "** Applied Credit Of " + QPTrim$(Using$("$##,##0.00", TaxBill.OverPayAmt)) + " **"; Tab(40); "Other Taxes Included: " + Using$("$#,##0.00", TotOth)
        For LCnt = 26 To 36: Print #RptFile, "": Next LCnt
      ElseIf TaxBill.OverPayAmt > 0 And TotOth = 0 Then
        Print #RptFile, "** Applied Credit Of " + QPTrim$(Using$("$##,##0.00", TaxBill.OverPayAmt)) + " **"
        For LCnt = 26 To 36: Print #RptFile, "": Next LCnt
      ElseIf TaxBill.OverPayAmt = 0 And TotOth > 0 Then
        Print #RptFile, Tab(40); "Other Taxes Included: " + Using$("$#,##0.00", TotOth)
        For LCnt = 26 To 36: Print #RptFile, "": Next LCnt
      Else
        For LCnt = 25 To 36: Print #RptFile, "": Next LCnt
      End If
      'Lines 25 to 36 are blank
     'Line 37 for Totals
      Print #RptFile, Tab(48); "Total Tax Due ";
      Print #RptFile, Using$("$#######0.00", OldRound(TaxBill.TotalBillDue - TaxBill.OverPayAmt))
      Print #RptFile, Tab(48); "Tax Due Date: " + MakeRegDate(BillInfo.DueDate)
      Print #RptFile,
      Print #RptFile,
      Print #RptFile,
      Print #RptFile, "BN"; Using$("#####", x) 'PrnCnt)
      Print #RptFile, "~"
      TaxBill.BillPrinted = True
      TaxBill.Comment = fptxtComment.Text
      TaxBill.Comment2 = fptxtComment2.Text
      TaxBill.CommentPlace = fpcmbCommentPlace.Text
      TaxBill.BillNumber = BillNo
      BillNo = BillNo + 1
      Put TBHandle, CustArr, TaxBill

NotThisOne:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  Print #RptFile, FF$
  Close
  Call MakeFile
  ViewPrint File$, "Real Property Tax Bills", True

End Sub

Private Sub SortByMortCode()
  Dim x As Long, z As Integer
  Dim NextRec As Long
  Dim MortRec As MortCodeRecType
  Dim MHandle As Integer
  Dim NumOfMCodes As Integer
  Dim TBRec As VARETaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim NoBill As Long
  Dim MortCnt As Long
  Dim MRRec As BillPrintMortIdxType
  Dim MRHandle As Integer
  Dim NumOfMRRecs As Long
  Dim TBMortCnt As Long
  
  KillFile "RZIPIDX.DAT"
  KillFile "MORTIDX.DAT"
  
  OpenMortCodeFile MHandle, NumOfMCodes
  OpenRealTaxBillFile TBHandle, NumOfTBRecs
  
  MortCnt = 0
  ReDim MortArr(1 To 1) As String
  frmVATaxShowPctComp.Label1 = "Sorting By Mortgage Code"
  frmVATaxShowPctComp.Show , Me
  For x = 1 To NumOfTBRecs
    Get TBHandle, x, TBRec
    If TBRec.BillNumber > 0 And QPTrim$(TBRec.MORTCODE) <> "" Then
      For z = 1 To TBMortCnt
        If QPTrim$(MortArr(z)) = QPTrim$(TBRec.MORTCODE) Then
          Exit For
        End If
      Next z
      If z > TBMortCnt Then
        TBMortCnt = TBMortCnt + 1
        ReDim Preserve MortArr(1 To TBMortCnt) As String
        MortArr(TBMortCnt) = QPTrim$(TBRec.MORTCODE)
      End If
    End If
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      Exit Sub
    End If
  Next x
  
  Unload frmVATaxShowPctComp
  
  frmVATaxShowPctComp.Label1 = "Sorting By Mortgage Code"
  frmVATaxShowPctComp.Show , Me
  
  OpenMortIdxFile MRHandle, NumOfMRRecs
  
  For z = 1 To TBMortCnt
    For x = 1 To NumOfTBRecs
      Get TBHandle, x, TBRec
      If TBRec.BillNumber > 0 Then
        If QPTrim$(TBRec.MORTCODE) = QPTrim$(MortArr(z)) Then
          MortCnt = MortCnt + 1
          MRRec.TaxBillRec = x
          Put MRHandle, MortCnt, MRRec
        End If
      End If
    Next x
    frmVATaxShowPctComp.ShowPctComp z, TBMortCnt
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      Exit Sub
    End If
  Next z
  Unload frmVATaxShowPctComp
  
  frmVATaxShowPctComp.Label1 = "Sorting By Mortgage Code"
  frmVATaxShowPctComp.Show , Me
  
  For x = 1 To NumOfTBRecs
    Get TBHandle, x, TBRec
    If TBRec.BillNumber > 0 And QPTrim$(TBRec.MORTCODE) = "" Then
      MortCnt = MortCnt + 1
      MRRec.TaxBillRec = x
      Put MRHandle, MortCnt, MRRec
    End If
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      Exit Sub
    End If
  Next x
  
  Unload frmVATaxShowPctComp
            
  Close
  
End Sub

Private Sub SortByZipCode1()
  Dim x As Long, z As Integer
  Dim NextRec As Long
  Dim TBRRec As VARETaxBillType
  Dim TBPRec As VAPPTaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim NoBill As Long
  Dim Nextx As Long
  Dim ThisZipcnt As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim ThisCustRec As Long
  Dim PZipRec As BillPrintPZipIdxType
  Dim RZipRec As BillPrintRZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim ZipCnt As Long
  Dim ThisZip$
  Dim Big$, SaveBig$
  Dim Hold$
  Dim Thisx As Integer
  
  If fpcmbType.Text = "PERSONAL" Then
    KillFile "PZIPIDX.DAT"
    OpenPersTaxBillFile TBHandle, NumOfTBRecs
    OpenPZipIdxFile ZHandle, NumOfZRecs
    GoTo SortP
  ElseIf fpcmbType.Text = "REAL" Then
    KillFile "RZIPIDX.DAT"
    KillFile "MORTIDX.DAT"
    OpenRealTaxBillFile TBHandle, NumOfTBRecs
    OpenRZipIdxFile ZHandle, NumOfZRecs
  End If

  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  ReDim ThisZipArr(1 To 1) As String
  
  frmVATaxShowPctComp.Label1 = "Sorting By Zip Code"
  frmVATaxShowPctComp.Show , Me
  
  For x = 1 To NumOfTBRecs
    Get TBHandle, x, TBRRec
    If TBRRec.BillNumber <= 0 Then GoTo SkipIt
    ThisCustRec = TBRRec.CustRec
    If ThisCustRec > 0 Then
      Get TCHandle, ThisCustRec, TaxCust
      ThisZip = QPTrim$(TaxCust.Zip)
      If Len(ThisZip) = 6 And Mid(ThisZip, 6, 1) = "-" Then
        ThisZip = Mid(ThisZip, 1, 5)
      End If
      For z = 1 To ThisZipcnt
        If ThisZip = ThisZipArr(z) Then Exit For
      Next z
      If z > ThisZipcnt Then
        ThisZipcnt = ThisZipcnt + 1
        ReDim Preserve ThisZipArr(1 To ThisZipcnt) As String
        ThisZipArr(ThisZipcnt) = QPTrim$(TaxCust.Zip)
      End If
    End If
SkipIt:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      Exit Sub
    End If
  Next x
  
  Unload frmVATaxShowPctComp
  
  GoSub SortLowToHigh
  
  frmVATaxShowPctComp.Label1 = "Sorting By Zip Code"
  frmVATaxShowPctComp.Show , Me
  
  For z = 1 To ThisZipcnt
    For x = 1 To NumOfTBRecs
    Get TBHandle, x, TBRRec
      If TBRRec.BillNumber > 0 Then
        ThisCustRec = TBRRec.CustRec
        If ThisCustRec > 0 Then
          Get TCHandle, ThisCustRec, TaxCust
          ThisZip = QPTrim$(TaxCust.Zip)
          If Len(ThisZip) = 6 And Mid(ThisZip, 6, 1) = "-" Then
            ThisZip = Mid(ThisZip, 1, 5)
          End If
          If ThisZip = ThisZipArr(z) Then
            ZipCnt = ZipCnt + 1
            RZipRec.TaxBillRec = x
            Put ZHandle, ZipCnt, RZipRec
          End If
        End If
      End If
    Next x
    frmVATaxShowPctComp.ShowPctComp z, ThisZipcnt
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      Exit Sub
    End If
  Next z
  Unload frmVATaxShowPctComp
  
  Close
  Exit Sub
  
SortLowToHigh:
  Big = "0"
  
  For x = 1 To ThisZipcnt
    If ThisZipArr(x) > Big Then
      Big = ThisZipArr(x)
    End If
  Next x
  
  Big = Big + "9"
  SaveBig = Big
  Nextx = 1
  Do While Nextx <= ThisZipcnt
    For x = Nextx To ThisZipcnt
      If ThisZipArr(x) < Big Then
        Big = ThisZipArr(x)
        Thisx = x
      End If
    Next x
    Hold = ThisZipArr(Nextx)
    ThisZipArr(Nextx) = ThisZipArr(Thisx)
    ThisZipArr(Thisx) = Hold
    Big = SaveBig
    Nextx = Nextx + 1
  Loop
  
  Return
    
SortP:
  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  ReDim ThisZipArr(1 To 1) As String
  
  frmVATaxShowPctComp.Label1 = "Sorting By Zip Code"
  frmVATaxShowPctComp.Show , Me
  For x = 1 To NumOfTBRecs
    Get TBHandle, x, TBPRec
    If TBPRec.BillNumber <= 0 Then GoTo SkipItP
    ThisCustRec = TBPRec.CustRec
    If ThisCustRec > 0 Then
      Get TCHandle, ThisCustRec, TaxCust
      ThisZip = QPTrim$(TaxCust.Zip)
      If Len(ThisZip) = 6 And Mid(ThisZip, 6, 1) = "-" Then
        ThisZip = Mid(ThisZip, 1, 5)
      End If
      For z = 1 To ThisZipcnt
        If ThisZip = ThisZipArr(z) Then Exit For
      Next z
      If z > ThisZipcnt Then
        ThisZipcnt = ThisZipcnt + 1
        ReDim Preserve ThisZipArr(1 To ThisZipcnt) As String
        ThisZipArr(ThisZipcnt) = QPTrim$(TaxCust.Zip)
      End If
    End If
SkipItP:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      Exit Sub
    End If
  Next x
  
  Unload frmVATaxShowPctComp
  GoSub SortLowToHigh
  
  frmVATaxShowPctComp.Label1 = "Sorting By Zip Code"
  frmVATaxShowPctComp.Show , Me
  
  For z = 1 To ThisZipcnt
    For x = 1 To NumOfTBRecs
    Get TBHandle, x, TBPRec
      If TBPRec.BillNumber > 0 Then
        ThisCustRec = TBPRec.CustRec
        If ThisCustRec > 0 Then
          Get TCHandle, ThisCustRec, TaxCust
          ThisZip = QPTrim$(TaxCust.Zip)
          If Len(ThisZip) = 6 And Mid(ThisZip, 6, 1) = "-" Then
            ThisZip = Mid(ThisZip, 1, 5)
          End If
          If ThisZip = ThisZipArr(z) Then
            ZipCnt = ZipCnt + 1
            PZipRec.TaxBillRec = x
            Put ZHandle, ZipCnt, PZipRec
          End If
        End If
      End If
    Next x
    frmVATaxShowPctComp.ShowPctComp z, ThisZipcnt
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      Exit Sub
    End If
  Next z
  Unload frmVATaxShowPctComp
  
  Close

End Sub

Private Sub SortByZipCode()
  Dim x As Long, z As Integer
  Dim NextRec As Long
  Dim TBRRec As VARETaxBillType
  Dim TBPRec As VAPPTaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim NoBill As Long
  Dim Nextx As Long
  Dim ThisZipcnt As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim ThisCustRec As Long
  Dim PZipRec As BillPrintPZipIdxType
  Dim RZipRec As BillPrintRZipIdxType
  Dim ZHandle As Integer
  Dim NumOfZRecs As Long
  Dim ZipCnt As Long
  Dim ThisZip$
  Dim Big$, SaveBig$
  Dim Hold$, Nextz As Long
  Dim Thisx As Integer
  
  If fpcmbType.Text = "PERSONAL" Then
    KillFile "PZIPIDX.DAT"
    OpenPersTaxBillFile TBHandle, NumOfTBRecs
    OpenPZipIdxFile ZHandle, NumOfZRecs
    GoTo SortP
  ElseIf fpcmbType.Text = "REAL" Then
    KillFile "RZIPIDX.DAT"
    KillFile "MORTIDX.DAT"
    OpenRealTaxBillFile TBHandle, NumOfTBRecs
    OpenRZipIdxFile ZHandle, NumOfZRecs
  End If

  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  ReDim ThisZipArr(1 To 1) As String
  
  frmVATaxShowPctComp.Label1 = "Sorting By Zip Code"
  frmVATaxShowPctComp.Show , Me
  
  For x = 1 To NumOfTBRecs
    Get TBHandle, x, TBRRec
    If TBRRec.BillNumber <= 0 Then GoTo SkipIt
    ThisCustRec = TBRRec.CustRec
    If ThisCustRec > 0 Then
      Get TCHandle, ThisCustRec, TaxCust
      ThisZip = QPTrim$(TaxCust.Zip)
      If Len(ThisZip) = 6 And Mid(ThisZip, 6, 1) = "-" Then
        ThisZip = Mid(ThisZip, 1, 5)
      End If
      For z = 1 To ThisZipcnt
        If ThisZip = ThisZipArr(z) Then Exit For
      Next z
      If z > ThisZipcnt Then
        ThisZipcnt = ThisZipcnt + 1
        ReDim Preserve ThisZipArr(1 To ThisZipcnt) As String
        ThisZipArr(ThisZipcnt) = ThisZip
      End If
    End If
SkipIt:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      Exit Sub
    End If
  Next x
  
  Unload frmVATaxShowPctComp
  
  GoSub SortLowToHigh
  
  ReDim ZipArray(1 To ThisZipcnt, 1 To NumOfTBRecs) As Long
  frmVATaxShowPctComp.Label1 = "Sorting By Zip Code"
  frmVATaxShowPctComp.Show , Me
  For x = 1 To NumOfTBRecs
    Get TBHandle, x, TBRRec
    If TBRRec.BillNumber > 0 Then
      ThisCustRec = TBRRec.CustRec
      If ThisCustRec > 0 Then
        ThisZip = QPTrim$(TBRRec.CustZip)
        If Len(ThisZip) = 6 And Mid(ThisZip, 6, 1) = "-" Then
          ThisZip = Mid(ThisZip, 1, 5)
        End If
        For z = 1 To ThisZipcnt
          If ThisZip = ThisZipArr(z) Then
            ZipArray(z, x) = x
            ZipCnt = ZipCnt + 1
            Exit For
          End If
        Next z
      End If
    End If
  Next x
  
  Nextz = 0
  For z = 1 To ThisZipcnt
    For x = 1 To NumOfTBRecs 'ZipCnt
      If ZipArray(z, x) > 0 Then
        RZipRec.TaxBillRec = x
        Nextz = Nextz + 1
        Put ZHandle, Nextz, RZipRec
      End If
    Next x
    frmVATaxShowPctComp.ShowPctComp z, ThisZipcnt
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      Exit Sub
    End If
 Next z
  
  Unload frmVATaxShowPctComp
  
  Close
  Exit Sub
  
SortLowToHigh:
  Big = "0"
  
  For x = 1 To ThisZipcnt
    If ThisZipArr(x) > Big Then
      Big = ThisZipArr(x)
    End If
  Next x
  
  Big = Big + "9"
  SaveBig = Big
  Nextx = 1
  Do While Nextx <= ThisZipcnt
    For x = Nextx To ThisZipcnt
      If ThisZipArr(x) < Big Then
        Big = ThisZipArr(x)
        Thisx = x
      End If
    Next x
    Hold = ThisZipArr(Nextx)
    ThisZipArr(Nextx) = ThisZipArr(Thisx)
    ThisZipArr(Thisx) = Hold
    Big = SaveBig
    Nextx = Nextx + 1
  Loop
  
  Return
    
SortP:
  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  ReDim ThisZipArr(1 To 1) As String
  
  frmVATaxShowPctComp.Label1 = "Sorting By Zip Code"
  frmVATaxShowPctComp.Show , Me
  For x = 1 To NumOfTBRecs
    Get TBHandle, x, TBPRec
    If TBPRec.BillNumber <= 0 Then GoTo SkipItP
    ThisCustRec = TBPRec.CustRec
    If ThisCustRec > 0 Then
      Get TCHandle, ThisCustRec, TaxCust
      ThisZip = QPTrim$(TaxCust.Zip)
      If Len(ThisZip) = 6 And Mid(ThisZip, 6, 1) = "-" Then
        ThisZip = Mid(ThisZip, 1, 5)
      End If
      For z = 1 To ThisZipcnt
        If ThisZip = ThisZipArr(z) Then Exit For
      Next z
      If z > ThisZipcnt Then
        ThisZipcnt = ThisZipcnt + 1
        ReDim Preserve ThisZipArr(1 To ThisZipcnt) As String
        ThisZipArr(ThisZipcnt) = ThisZip
      End If
    End If
SkipItP:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTBRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      Exit Sub
    End If
  Next x
  
  Unload frmVATaxShowPctComp
  GoSub SortLowToHigh
  
  frmVATaxShowPctComp.Label1 = "Sorting By Zip Code"
  frmVATaxShowPctComp.Show , Me
  
  ReDim ZipArray(1 To ThisZipcnt, 1 To NumOfTBRecs) As Long
  frmVATaxShowPctComp.Label1 = "Sorting By Zip Code"
  frmVATaxShowPctComp.Show , Me
  
  For x = 1 To NumOfTBRecs
    Get TBHandle, x, TBPRec
    If TBPRec.BillNumber > 0 Then
      ThisCustRec = TBPRec.CustRec
      If ThisCustRec > 0 Then
        ThisZip = QPTrim$(TBPRec.CustZip)
        If Len(ThisZip) = 6 And Mid(ThisZip, 6, 1) = "-" Then
          ThisZip = Mid(ThisZip, 1, 5)
        End If
        For z = 1 To ThisZipcnt
          If ThisZip = ThisZipArr(z) Then
            ZipArray(z, x) = x
            ZipCnt = ZipCnt + 1
            Exit For
          End If
        Next z
      End If
    End If
  Next x
  
  Nextz = 0
  For z = 1 To ThisZipcnt
    For x = 1 To NumOfTBRecs  'ZipCnt
      If ZipArray(z, x) > 0 Then
        PZipRec.TaxBillRec = x
        Nextz = Nextz + 1
        Put ZHandle, Nextz, PZipRec
      End If
    Next x
    frmVATaxShowPctComp.ShowPctComp z, ThisZipcnt
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      Exit Sub
    End If
 Next z
  
  Unload frmVATaxShowPctComp
'  For z = 1 To ThisZipcnt
'    For x = 1 To NumOfTBRecs
'    Get TBHandle, x, TBPRec
'      If TBPRec.BillNumber > 0 Then
'        ThisCustRec = TBPRec.CustRec
'        If ThisCustRec > 0 Then
'          Get TCHandle, ThisCustRec, TaxCust
'          ThisZip = QPTrim$(TaxCust.Zip)
'          If Len(ThisZip) = 6 And Mid(ThisZip, 6, 1) = "-" Then
'            ThisZip = Mid(ThisZip, 1, 5)
'          End If
'          If ThisZip = ThisZipArr(z) Then
'            ZipCnt = ZipCnt + 1
'            PZipRec.TaxBillRec = x
'            Put ZHandle, ZipCnt, PZipRec
'          End If
'        End If
'      End If
'    Next x
'    frmVATaxShowPctComp.ShowPctComp z, ThisZipcnt
'    If frmVATaxShowPctComp.Out = True Then
'      Close
'      frmVATaxShowPctComp.Out = False
'      Unload frmVATaxShowPctComp
'      Exit Sub
'    End If
'  Next z
'  Unload frmVATaxShowPctComp
  
  Close

End Sub



