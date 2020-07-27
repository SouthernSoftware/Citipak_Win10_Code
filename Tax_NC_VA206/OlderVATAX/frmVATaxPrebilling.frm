VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVATaxPrebilling 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Prebilling Information"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxPrebilling.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbInclNoBills 
      Height          =   405
      Left            =   5160
      TabIndex        =   11
      Top             =   6480
      Width           =   735
      _Version        =   196608
      _ExtentX        =   1296
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
      ColDesigner     =   "frmVATaxPrebilling.frx":08CA
   End
   Begin LpLib.fpCombo fpcmbTownships 
      Height          =   375
      Left            =   8400
      TabIndex        =   13
      Top             =   3120
      Width           =   2775
      _Version        =   196608
      _ExtentX        =   4895
      _ExtentY        =   661
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      ColDesigner     =   "frmVATaxPrebilling.frx":0C31
   End
   Begin LpLib.fpCombo fpcmbPrintOpt 
      Height          =   405
      Left            =   7560
      TabIndex        =   17
      Top             =   6360
      Width           =   2490
      _Version        =   196608
      _ExtentX        =   4392
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
      ColDesigner     =   "frmVATaxPrebilling.frx":0F98
   End
   Begin LpLib.fpCombo fpcmbPrintOrder 
      Height          =   405
      Left            =   7080
      TabIndex        =   19
      Top             =   7200
      Width           =   3375
      _Version        =   196608
      _ExtentX        =   5953
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
      ColDesigner     =   "frmVATaxPrebilling.frx":12FF
   End
   Begin LpLib.fpCombo fpcmbSplit 
      Height          =   390
      Left            =   2640
      TabIndex        =   0
      Top             =   1800
      Width           =   2295
      _Version        =   196608
      _ExtentX        =   4048
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
      ColDesigner     =   "frmVATaxPrebilling.frx":1666
   End
   Begin LpLib.fpCombo fpcmbCycle 
      Height          =   375
      Left            =   8040
      TabIndex        =   40
      Top             =   3600
      Width           =   3135
      _Version        =   196608
      _ExtentX        =   5530
      _ExtentY        =   661
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      Columns         =   2
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
      ColDesigner     =   "frmVATaxPrebilling.frx":19CD
   End
   Begin LpLib.fpCombo fpcmbCounty 
      Height          =   375
      Left            =   8040
      TabIndex        =   41
      Top             =   4080
      Width           =   3135
      _Version        =   196608
      _ExtentX        =   5530
      _ExtentY        =   661
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
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
      Columns         =   2
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
      ColDesigner     =   "frmVATaxPrebilling.frx":1D8C
   End
   Begin LpLib.fpCombo fpcmbSuppOnly 
      Height          =   405
      Left            =   5160
      TabIndex        =   10
      Top             =   5880
      Width           =   735
      _Version        =   196608
      _ExtentX        =   1296
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
      ColDesigner     =   "frmVATaxPrebilling.frx":214B
   End
   Begin EditLib.fpDateTime fptxtBillDate 
      Height          =   372
      Left            =   9360
      TabIndex        =   12
      Top             =   2400
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
   Begin EditLib.fpDoubleSingle fpDblSnglMTRate 
      Height          =   372
      Left            =   4800
      TabIndex        =   7
      Top             =   3720
      Width           =   1332
      _Version        =   196608
      _ExtentX        =   2355
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
      Text            =   "0.0000"
      DecimalPlaces   =   4
      DecimalPoint    =   "."
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
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
   Begin VB.OptionButton OptDiscYes 
      BackColor       =   &H008F8265&
      Caption         =   "Apply Discounts"
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
      Height          =   300
      Left            =   6600
      TabIndex        =   15
      Top             =   5400
      Width           =   1935
   End
   Begin VB.OptionButton OptDiscNo 
      BackColor       =   &H008F8265&
      Caption         =   "No Discounts"
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
      Height          =   255
      Left            =   6600
      TabIndex        =   14
      Top             =   5040
      Width           =   1932
   End
   Begin EditLib.fpDoubleSingle fpDblSnglRealRate 
      Height          =   372
      Left            =   1800
      TabIndex        =   2
      Top             =   2760
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
      ControlType     =   0
      Text            =   "0.0000"
      DecimalPlaces   =   4
      DecimalPoint    =   "."
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "0"
      NegFormat       =   1
      NegToggle       =   0   'False
      Separator       =   "."
      UseSeparator    =   -1  'True
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
      Left            =   4800
      TabIndex        =   3
      Top             =   2760
      Width           =   1332
      _Version        =   196608
      _ExtentX        =   2355
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
      Text            =   "0.0000"
      DecimalPlaces   =   4
      DecimalPoint    =   "."
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
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
   Begin EditLib.fpDoubleSingle fpDblSnglLateList 
      Height          =   372
      Left            =   4560
      TabIndex        =   9
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
      ControlType     =   0
      Text            =   "0.0000"
      DecimalPlaces   =   4
      DecimalPoint    =   "."
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
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
   Begin EditLib.fpDateTime fptxtRCurrYear 
      Height          =   348
      Left            =   7080
      TabIndex        =   1
      Top             =   1800
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
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
   Begin EditLib.fpDateTime fptxtDiscXDate 
      Height          =   372
      Left            =   9000
      TabIndex        =   16
      Top             =   5400
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
      Height          =   492
      Left            =   9120
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   7920
      Width           =   1572
      _Version        =   131072
      _ExtentX        =   2773
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmVATaxPrebilling.frx":24B2
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   492
      Left            =   6960
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   7920
      Width           =   1572
      _Version        =   131072
      _ExtentX        =   2773
      _ExtentY        =   868
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
      ButtonDesigner  =   "frmVATaxPrebilling.frx":2691
   End
   Begin EditLib.fpDoubleSingle fpDblSnglFERate 
      Height          =   372
      Left            =   1800
      TabIndex        =   4
      Top             =   3240
      Width           =   1332
      _Version        =   196608
      _ExtentX        =   2355
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
      Text            =   "0.0000"
      DecimalPlaces   =   4
      DecimalPoint    =   "."
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
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
   Begin EditLib.fpDoubleSingle fpDblSnglMHRate 
      Height          =   372
      Left            =   4800
      TabIndex        =   5
      Top             =   3240
      Width           =   1332
      _Version        =   196608
      _ExtentX        =   2355
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
      Text            =   "0.0000"
      DecimalPlaces   =   4
      DecimalPoint    =   "."
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
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
   Begin EditLib.fpDoubleSingle fpDblSnglMCRate 
      Height          =   372
      Left            =   1800
      TabIndex        =   6
      Top             =   3720
      Width           =   1332
      _Version        =   196608
      _ExtentX        =   2355
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
      Text            =   "0.0000"
      DecimalPlaces   =   4
      DecimalPoint    =   "."
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
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
   Begin EditLib.fpDateTime fptxtDueDate 
      Height          =   372
      Left            =   3360
      TabIndex        =   8
      Top             =   4476
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
   Begin EditLib.fpDateTime fptxtPCurrYear 
      Height          =   348
      Left            =   10080
      TabIndex        =   57
      Top             =   1800
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
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
   Begin VB.Line Line8 
      BorderColor     =   &H0080FFFF&
      BorderStyle     =   3  'Dot
      X1              =   240
      X2              =   6240
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Label Label31 
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
      Left            =   8280
      TabIndex        =   58
      Top             =   1872
      Width           =   1692
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Due Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   276
      Left            =   240
      TabIndex        =   56
      Top             =   4320
      Width           =   1932
   End
   Begin VB.Line Line15 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   6240
      X2              =   240
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Print Options:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   275
      Left            =   6240
      TabIndex        =   55
      Top             =   5880
      Width           =   1692
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Billing Date:"
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
      Left            =   7920
      TabIndex        =   54
      Top             =   2496
      Width           =   1332
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Billing Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   275
      Left            =   6240
      TabIndex        =   53
      Top             =   2280
      Width           =   1572
   End
   Begin VB.Line Line14 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   240
      X2              =   11280
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Other Options:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   276
      Left            =   240
      TabIndex        =   52
      Top             =   5040
      Width           =   1932
   End
   Begin VB.Line Line13 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   6240
      X2              =   240
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Rate Pct"
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
      Left            =   4800
      TabIndex        =   51
      ToolTipText     =   "Enter the percent amount as follows: 5% = 5; .5% = .5."
      Top             =   2472
      Width           =   1332
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Supplemental Only:"
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
      Left            =   2880
      TabIndex        =   50
      Top             =   5964
      Width           =   2172
   End
   Begin VB.Label lblSettings3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Max Vehicle Tax Value is:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   2400
      TabIndex        =   49
      Top             =   1200
      Width           =   3372
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "System Settings:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   275
      Left            =   240
      TabIndex        =   48
      Top             =   1080
      Width           =   2052
   End
   Begin VB.Line Line9 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   240
      X2              =   240
      Y1              =   1080
      Y2              =   1560
   End
   Begin VB.Line Line12 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   11280
      X2              =   11280
      Y1              =   1080
      Y2              =   1560
   End
   Begin VB.Line Line11 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   240
      X2              =   11280
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line10 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   240
      X2              =   11280
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblSettings2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Multi Year Value is:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   8520
      TabIndex        =   47
      Top             =   1200
      Width           =   2532
   End
   Begin VB.Label lblSettings1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PPTRA Discount is:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   5880
      TabIndex        =   46
      Top             =   1200
      Width           =   2532
   End
   Begin VB.Line Line7 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   5160
      X2              =   5160
      Y1              =   1680
      Y2              =   2280
   End
   Begin VB.Line Line6 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   240
      X2              =   240
      Y1              =   1668
      Y2              =   2518
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   11280
      X2              =   11280
      Y1              =   1680
      Y2              =   2880
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mach Tools:"
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
      Left            =   3360
      TabIndex        =   45
      Top             =   3804
      Width           =   1332
   End
   Begin VB.Label Label21 
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
      Left            =   360
      TabIndex        =   44
      Top             =   3804
      Width           =   1332
   End
   Begin VB.Label Label8 
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
      Left            =   3240
      TabIndex        =   43
      Top             =   3324
      Width           =   1452
   End
   Begin VB.Label Label1 
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
      Left            =   360
      TabIndex        =   42
      Top             =   3336
      Width           =   1332
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Discount Expiration Date "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   492
      Left            =   8640
      TabIndex        =   37
      Top             =   4896
      Width           =   2412
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0080FFFF&
      BorderStyle     =   3  'Dot
      X1              =   240
      X2              =   6240
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Rate Pct"
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
      Left            =   1800
      TabIndex        =   36
      ToolTipText     =   "Enter the percent amount as follows: 5% = 5; .5% = .5."
      Top             =   2472
      Width           =   1332
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   240
      X2              =   11280
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Early Payment Discount Setting:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   275
      Left            =   6240
      TabIndex        =   35
      Top             =   4560
      Width           =   3732
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   6240
      X2              =   11280
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Include Bills Even If Current Owed Is Zero:"
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
      Left            =   360
      TabIndex        =   34
      Top             =   6564
      Width           =   4692
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Tax Rates:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   276
      Left            =   240
      TabIndex        =   33
      Top             =   2280
      Width           =   1452
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Cycle:"
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
      Left            =   6480
      TabIndex        =   32
      Top             =   3720
      Width           =   1452
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select County:"
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
      Left            =   6360
      TabIndex        =   31
      Top             =   4176
      Width           =   1572
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Split Real/Personal:"
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
      Left            =   360
      TabIndex        =   30
      Top             =   1920
      Width           =   2172
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   6240
      X2              =   11280
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Billing Options:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   275
      Left            =   6240
      TabIndex        =   29
      Top             =   2880
      Width           =   1812
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   4812
      Left            =   6240
      Top             =   2880
      Width           =   5052
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   4812
      Left            =   240
      Top             =   2280
      Width           =   6012
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bill By Township:"
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
      Left            =   6360
      TabIndex        =   28
      Top             =   3228
      Width           =   1932
   End
   Begin VB.Label lblSupp1 
      BackStyle       =   0  'Transparent
      Caption         =   "System Tax Settings:"
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
      Height          =   732
      Left            =   480
      TabIndex        =   27
      Top             =   7440
      Width           =   5532
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Minimum Tax Setting:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   276
      Left            =   240
      TabIndex        =   26
      Top             =   7080
      Width           =   2532
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1332
      Left            =   240
      Top             =   7080
      Width           =   6012
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
      Left            =   7920
      TabIndex        =   25
      Top             =   6000
      Width           =   1812
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
      Left            =   7920
      TabIndex        =   24
      Top             =   6840
      Width           =   1812
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Late List Pct:"
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
      Left            =   3000
      TabIndex        =   23
      Top             =   5304
      Width           =   1452
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
      Left            =   3480
      TabIndex        =   22
      Top             =   2844
      Width           =   1212
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
      Height          =   252
      Left            =   240
      TabIndex        =   21
      Top             =   2844
      Width           =   1452
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
      Height          =   252
      Left            =   5280
      TabIndex        =   20
      Top             =   1872
      Width           =   1692
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   660
      Index           =   1
      Left            =   1488
      Top             =   228
      Width           =   8652
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Prebilling Information"
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
      Left            =   3144
      TabIndex        =   18
      Top             =   396
      Width           =   5292
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   1488
      Top             =   120
      Width           =   8652
   End
End
Attribute VB_Name = "frmVATaxPrebilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  'Private Temp_Class As Resize_Class
  Dim UseOpt As String * 1
  Dim UseSS As String * 1
  Dim ThisOpt$
  Dim WhatYearInt As Integer
  Dim WhatRYear$
  Dim WhatPYear$
  Dim BillNum&
  Dim RealRate#
  Dim PERSRATE#
  Dim LateList#
  Dim Order$
  Dim SupOnly$
  Dim RptType$
  Dim CustArr() As Long
  Dim OptCnt As Long
  Dim SSCnt As Long
  Dim MortCodes() As MortRecType
  Dim NumMortCodes As Integer
  Dim UsingIdx As Boolean
  Dim UsingMinTax As Integer
  Dim MinBill As Double
  Dim MaxVehVal As Double
  Dim MinVehVal As Double
  Dim LawChngDate As Integer
  Dim SpecTax As Double
  Dim SpecDesc As String
  Dim ThisTown$
  Dim Townships() As String
  Dim TSCnt As Integer
  Dim PPTRAYN As Boolean
  Dim RealOK As Boolean
  Dim PersOK As Boolean
  
Private Sub cmdExit_Click()
  KillFile "C:\CPWork\revrglbill.dat"
  frmVATaxBillingMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdProcess_Click()
  Dim RealBillInfo As VARETaxBillInfoType
  Dim PersBillInfo As VAPPTaxBillInfoType
  Dim BIHandle As Integer
  Dim IdxType As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim MortCodeRec As MortCodeRecType
  Dim MCHandle As Integer
  Dim x As Integer
  Dim FileName$
  Dim ThisRYear$
  Dim ThisPYear$
  Dim DirCnt As Integer
  Dim MyPath$
  Dim MyName$
  
  On Error GoTo ERRORSTUFF
  
  If fpcmbSplit.Text = "REAL" And RealOK = False Then
    Call TaxMsg(800, "Processing real bills cannot take place until the outstanding real payment file is posted.")
    Exit Sub
  ElseIf fpcmbSplit.Text = "PERSONAL" And PersOK = False Then
    Call TaxMsg(800, "Processing personal bills cannot take place until the outstanding personal payment file is posted.")
    Exit Sub
  End If
    
  ReDim DirContents(1 To 1) As String
  
  MyPath = StartPath + "\TAXBILLBU\"
  MyName$ = Dir(MyPath, vbDirectory)
  Do While MyName <> ""
    MyName = Dir
    If Len(MyName) > 4 Then
      DirCnt = DirCnt + 1
      ReDim Preserve DirContents(1 To DirCnt) As String
      DirContents(DirCnt) = MyPath + MyName
    End If
  Loop
  fpcmbCounty.Col = 1
  fpcmbCounty.Row = fpcmbCounty.ListIndex
  fpcmbCycle.Col = 1
  fpcmbCycle.Row = fpcmbCycle.ListIndex
  fpcmbTownships.Col = 1
  fpcmbTownships.Row = fpcmbTownships.ListIndex
  ThisRYear = CStr(fptxtRCurrYear.Text)
  ThisPYear = CStr(fptxtPCurrYear.Text)
  If fpcmbSuppOnly.Text = "Y" Then
    MainLog ("NOTE: User elected to process prebilling with supplemental only flag on. Warning against duplicate prebilling in current year was skipped.")
    GoTo SkipWarning
  End If
  If fpcmbSplit.Text = "PERSONAL" Then
    FileName = "TAXBILLBU\POSTP" + Mid(fpcmbCounty.ColText, 1, 3) + Mid(fpcmbCycle.ColText, 1, 3) + Mid(fpcmbTownships.Text, 1, 3) + ThisPYear ' + ".DAT"
    For x = 1 To DirCnt
      If InStr(DirContents(x), FileName) > 0 Then
        If TaxMsgWOpts(700, "Personal tax bills with these parameters have already been posted for tax year " + fptxtPCurrYear.Text + ". Press F10 to continue prebilling. Otherwise, press ESC to abort prebilling. NOTE: Press F10 if multiple billing within the year is adminstration policy.", "F10 Continue", "ESC Abort") = "abort" Then
          Unload frmVATaxMsgWOpts
          MainLog ("WARNING: User elected to abort personal prebilling processing for tax year " + fptxtPCurrYear.Text + " when they were warned the parameters for this year had already been billed and posted.")
          Close
          Exit Sub
        Else
          Unload frmVATaxMsgWOpts
          MainLog ("WARNING: User elected to continue personal prebilling processing for tax year " + fptxtPCurrYear.Text + " after they were warned the parameters for this year had already been billed and posted.")
          Exit For
        End If
      End If
    Next x
  ElseIf fpcmbSplit.Text = "REAL" Then
    FileName = "TAXBILLBU\POSTR" + Mid(fpcmbCounty.ColText, 1, 3) + Mid(fpcmbCycle.ColText, 1, 3) + Mid(fpcmbTownships.Text, 1, 3) + ThisRYear '+ ".DAT"
    For x = 1 To DirCnt
      If InStr(DirContents(x), FileName) > 0 Then
        If TaxMsgWOpts(600, "Real tax bills with these parameters have already been posted for tax year " + fptxtRCurrYear.Text + ". Press F10 to continue prebilling. Otherwise, press ESC to abort prebilling. NOTE: Press F10 if multiple billing within the year is adminstration policy.", "F10 Continue", "ESC Abort") = "abort" Then
          Unload frmVATaxMsgWOpts
          MainLog ("WARNING: User elected to abort real prebilling processing for tax year " + fptxtRCurrYear.Text + " when they were warned the parameters for this year had already been billed and posted.")
          Close
          Exit Sub
        Else
          Unload frmVATaxMsgWOpts
          MainLog ("WARNING: User elected to continue real prebilling processing for tax year " + fptxtRCurrYear.Text + " after they were warned the parameters for this year had already been billed and posted.")
          Exit For
        End If
      End If
    Next x
  End If
SkipWarning:
  If fpcmbSplit.Text = "REAL" Then
    If Exist("txrblsprn.dat") Then
      If TaxMsgWOpts(700, "NOTE: Real tax bills will need to be printed again upon completion of this process before posting can take place. Press F10 if you wish to continue. Otherwise, press ESC to abort prebilling.", "F10 Continue", "ESC Abort") = "abort" Then
        Unload frmVATaxMsgWOpts
        Close
        Exit Sub
      Else
        Unload frmVATaxMsgWOpts
        MainLog ("WARNING: User elected to continue the prebilling process after being warned that real tax bills will have to be reprinted.")
      End If
    End If
  ElseIf fpcmbSplit.Text = "PERSONAL" Then
    If Exist("txpblsprn.dat") Then
      If TaxMsgWOpts(700, "NOTE: Personal tax bills will need to be printed again upon completion of this process before posting can take place. Press F10 if you wish to continue. Otherwise, press ESC to abort prebilling.", "F10 Continue", "ESC Abort") = "abort" Then
        Unload frmVATaxMsgWOpts
        Close
        Exit Sub
      Else
        Unload frmVATaxMsgWOpts
        MainLog ("WARNING: User elected to continue the prebilling process after being warned that personal tax bills will have to be reprinted.")
      End If
    End If
  End If
  
  If Date2Num(fptxtDueDate.Text) < OldRound(Date2Num(Date) - 365) Or Date2Num(fptxtDueDate.Text) > OldRound(Date2Num(Date) + 365) Then
    If TaxMsgWOpts(700, "The due date is " + fptxtDueDate.Text + " and appears to be incorrect. If the date entered is OK then press F10 to continue. Otherwise, press ESC to edit.", "F10 Continue", "ESC Edit") = "abort" Then
      Unload frmVATaxMsgWOpts
      Close
      fptxtDueDate.SetFocus
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
      MainLog ("User warned that the due date being processed " + fptxtDueDate.Text + " is more than 365 days from today and the user elected to contniue anyway.")
    End If
  End If
  
  If fpDblSnglPersRate.Enabled = True Then
    If CheckPreReq = False Then
      Exit Sub
    Else
      If RevsAndGLsOK(Me, CInt(fptxtPCurrYear.Text), "P") = False Then
        Exit Sub
      End If
      KillFile PersTaxBillFile
      KillFile PersTaxBillInfoFile
      KillFile "PZIPIDX.DAT" 'added 12/706
      OpenPersBillInfoFile BIHandle
      PersBillInfo.BillNum = 0 'BillNum&
      PersBillInfo.PERSRATE = PERSRATE#
      PersBillInfo.LATEPCT = LateList#
      PersBillInfo.FERate = CDbl(fpDblSnglFERate.Value)
      PersBillInfo.MHRate = CDbl(fpDblSnglMHRate.Value)
      PersBillInfo.MCRate = CDbl(fpDblSnglMCRate.Value)
      PersBillInfo.MTRate = CDbl(fpDblSnglMTRate.Value)
      PersBillInfo.PRNORDER = Order$
      PersBillInfo.TaxYear = WhatPYear
      If fpcmbCounty.Enabled = True Then
        fpcmbCounty.Col = 1
        PersBillInfo.CountyPara = QPTrim$(fpcmbCounty.ColText)
      Else
        PersBillInfo.CountyPara = "ALL COUNTIES"
      End If
      If fpcmbTownships.Enabled = True Then
        PersBillInfo.TwnShpPara = QPTrim$(fpcmbTownships.Text)
      Else
        PersBillInfo.TwnShpPara = "ALL TOWNSHIPS"
      End If
      If fpcmbCycle.Enabled = True Then
        fpcmbCycle.Col = 1
        PersBillInfo.CyclePara = QPTrim$(fpcmbCycle.ColText)
      Else
        PersBillInfo.CyclePara = "No CYCLES"
      End If
      If fpcmbSplit.Enabled = True Then
        PersBillInfo.SplitPara = QPTrim$(fpcmbSplit.Text)
      Else
        PersBillInfo.SplitPara = "NO SPLIT"
      End If
      If OptDiscYes.Value = True Then 'added 9/20/05
        PersBillInfo.XDate = Date2Num(fptxtDiscXDate.Text)
      Else
        PersBillInfo.XDate = 0
      End If
      PersBillInfo.DueDate = Date2Num(fptxtDueDate.Text)
      Put BIHandle, 1, PersBillInfo
      Close BIHandle
    End If
  ElseIf fpDblSnglRealRate.Enabled = True Then
    If CheckPreReq = False Then
      Exit Sub
    Else
      If RevsAndGLsOK(Me, CInt(fptxtRCurrYear.Text), "R") = False Then
        Exit Sub
      End If
      KillFile RealTaxBillFile
      KillFile RealTaxBillInfoFile
      KillFile "RZIPIDX.DAT" 'added 12/706
      KillFile "MORTIDX.DAT" 'added 12/7/06
      OpenRealBillInfoFile BIHandle
      RealBillInfo.BillNum = 0 'BillNum&
      RealBillInfo.RealRate = RealRate#
      RealBillInfo.LATEPCT = LateList#
      RealBillInfo.PRNORDER = Order$
      RealBillInfo.TaxYear = WhatRYear
      RealBillInfo.XDate = Date2Num(fptxtDiscXDate)
      If fpcmbCounty.Enabled = True Then
        fpcmbCounty.Col = 1
        RealBillInfo.CountyPara = QPTrim$(fpcmbCounty.ColText)
      Else
        RealBillInfo.CountyPara = "ALL COUNTIES"
      End If
      If fpcmbTownships.Enabled = True Then
        RealBillInfo.TwnShpPara = QPTrim$(fpcmbTownships.Text)
      Else
        RealBillInfo.TwnShpPara = "ALL TOWNSHIPS"
      End If
      If fpcmbCycle.Enabled = True Then
        fpcmbCycle.Col = 1
        RealBillInfo.CyclePara = QPTrim$(fpcmbCycle.ColText)
      Else
        RealBillInfo.CyclePara = "No CYCLES"
      End If
      If fpcmbSplit.Enabled = True Then
        RealBillInfo.SplitPara = QPTrim$(fpcmbSplit.Text)
      Else
        RealBillInfo.SplitPara = "NO SPLIT"
      End If
      If OptDiscYes.Value = True Then 'added 9/20/05
        RealBillInfo.XDate = Date2Num(fptxtDiscXDate.Text)
      Else
        RealBillInfo.XDate = 0
      End If
      RealBillInfo.DueDate = Date2Num(fptxtDueDate.Text)
      Put BIHandle, 1, RealBillInfo
      Close BIHandle
    End If
  End If
  
  IdxType = CInt(Order$)
  Call MakeCustIdx(IdxType) 'populates global CustArr()
  
'  OpenTaxSetUpFile TMHandle 'remmed on 9/20/05
'  Get TMHandle, 1, TaxMasterRec
'  TaxMasterRec.TaxYear = CInt(fptxtRCurrYear.Text)
'  If OptDiscYes.Value = True Then 'remarked on 9/20/05
'    TaxMasterRec.DiscXDate = Date2Num(fptxtDiscXDate.Text)
'  Else
'    TaxMasterRec.DiscXDate = 0
'  End If
'  Put TMHandle, 1, TaxMasterRec
'  Close TMHandle
  
  GoSub LoadMortCodes

  If fpcmbPrintOpt.Text = "Graphical" Then
    If fpcmbSplit.Text = "PERSONAL" Then
      Call PrintPersGraphics
    Else
      Call PrintGraphics
    End If
  Else
    If fpcmbSplit.Text = "PERSONAL" Then
      Call TaxMsg(900, "Pitch 17 is recommended for this report.")
      Call PrintPersText
    Else
      Call TaxMsg(900, "Pitch 17 is recommended for this report.")
      Call PrintText
    End If
  End If
  Exit Sub
  
LoadMortCodes:
  ReDim MortCodes(1 To 1) As MortRecType
  OpenMortCodeFile MCHandle, NumMortCodes

  If NumMortCodes > 0 Then
    ReDim Preserve MortCodes(1 To NumMortCodes) As MortRecType
    For x = 1 To NumMortCodes
      Get MCHandle, x, MortCodeRec
      MortCodes(x).MORTCODE = MortCodeRec.MORTCODE
      MortCodes(x).MortRec = x
    Next
  End If
  Close MCHandle

  Return
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPrebilling", "cmdProcess_Click", Erl)
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      ClearInUse PWcnt
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxPrebilling.")
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
  MainLog ("User opened frmVATaxPrebilling.")
  UsingIdx = False
  SpecTax = 0
  UsingMinTax = 0
  SpecDesc = ""
  MinBill = 0
  Me.HelpContextID = hlpTaxPrebilling
  Call LoadMe
End Sub

Private Sub LoadMe()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim x As Long, y As Integer
  Dim TSRec As TownshipType
  Dim TSHandle As Integer
  Dim TSCnt As Integer
  
  On Error GoTo ERRORSTUFF
  
  RealOK = True
  PersOK = True
  
  If Check4PayBatch("P") = True Then
    frmVATaxUnpostedPaylist.BillType = "P"
    frmVATaxUnpostedPaylist.Label1.Caption = "Unposted payments files exist for personal property payments. Please post personal payments before proceeding with prebilling personal tax bills. The list shows the operators involved."
    frmVATaxUnpostedPaylist.Show vbModal
    PersOK = False
  End If

  If Check4PayBatch("R") = True Then
    frmVATaxUnpostedPaylist.BillType = "R"
    frmVATaxUnpostedPaylist.Label1.Caption = "Unposted payments files exist for real property payments. Please post real payments before proceeding with prebilling real tax bills. The list shows the operators involved."
    frmVATaxUnpostedPaylist.Show vbModal
    RealOK = False
  End If
  
  fptxtBillDate.Text = Date
  UseOpt = "N"
  UseSS = "N"
  ReDim MortCodes(1 To 1) As MortRecType
  fptxtDiscXDate.Text = CStr(Date)
  fptxtDueDate.Text = CStr(Date)
  lblSupp1.Caption = ""
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  PPTRAYN = True
  If TaxMasterRec.PPTRAYN <> "Y" Then
    PPTRAYN = False
  End If
  MaxVehVal = TaxMasterRec.MaxVehTaxVal
  MinVehVal = TaxMasterRec.MinVehTaxVal
  LawChngDate = TaxMasterRec.LawChngDate
  
  If PPTRAYN = True Then
    lblSettings1.Caption = "PPTRA Discount is " + QPTrim$(Using$("##0.00", TaxMasterRec.PPTRADisc)) + "%"
  Else
    lblSettings1.Caption = "PPTRA disabled"
  End If
  lblSettings2.Caption = "Multi Year Value is " + QPTrim$(Using$("#0", TaxMasterRec.MultiYear))
  lblSettings3.Caption = "Max Vehicle Tax Value is " + Using$("$###,##0.00", TaxMasterRec.MaxVehTaxVal)
  If TaxMasterRec.RealPersSplit = "Y" Then
    fpcmbSplit.AddItem "REAL"
    fpcmbSplit.AddItem "PERSONAL"
    fpcmbSplit.Text = "REAL"
    fpDblSnglPersRate.Enabled = False
    fpDblSnglFERate.Enabled = False
    fpDblSnglMHRate.Enabled = False
    fpDblSnglMCRate.Enabled = False
    fpDblSnglMTRate.Enabled = False
  Else
    fpcmbSplit.Text = "OPTION DISABLED"
    fpcmbSplit.Enabled = False
  End If
     
  If TaxMasterRec.UseCyclesYN = "Y" Then
    fpcmbCycle.AddItem "0" + Chr(9) + "NO CYCLE"
    For y = 1 To 5
      If TaxMasterRec.CycleNum(y) > 0 Then
        fpcmbCycle.AddItem CStr(TaxMasterRec.CycleNum(y)) + Chr(9) + QPTrim$(TaxMasterRec.CycleName(y))
      End If
    Next y
    fpcmbCycle.SearchText = CStr(0) + Chr(9) + "NO CYCLE"
    fpcmbCycle.Action = 0
    If fpcmbCycle.SearchIndex <> -1 Then
      fpcmbCycle.ListIndex = fpcmbCycle.SearchIndex
    Else
      fpcmbCycle.ListIndex = 0
    End If
  Else
    fpcmbCycle.Text = CStr(0) + Chr(9) + "OPTION DISABLED"
    fpcmbCycle.Enabled = False
  End If
     
  If TaxMasterRec.UseCountyYN = "Y" Then
    fpcmbCounty.AddItem "0" + Chr(9) + "ALL COUNTIES"
    For y = 1 To 5
      If TaxMasterRec.CountyNum(y) > 0 Then
        fpcmbCounty.AddItem CStr(TaxMasterRec.CountyNum(y)) + Chr(9) + QPTrim$(TaxMasterRec.CountyName(y))
      End If
    Next y
    fpcmbCounty.SearchText = CStr(0) + Chr(9) + "ALL COUNTIES"
    fpcmbCounty.Action = 0
    If fpcmbCounty.SearchIndex <> -1 Then
      fpcmbCounty.ListIndex = fpcmbCounty.SearchIndex
    Else
      fpcmbCounty.ListIndex = 0
    End If
  Else
    fpcmbCounty.Text = CStr(0) + Chr(9) + "OPTION DISABLED"
    fpcmbCounty.Enabled = False
  End If
     
  If TaxMasterRec.RTaxYear > 0 Then
    fptxtRCurrYear.Text = CStr(TaxMasterRec.RTaxYear)
  End If
  
  If TaxMasterRec.PTaxYear > 0 Then
    fptxtPCurrYear.Text = CStr(TaxMasterRec.PTaxYear)
  End If
  
  If TaxMasterRec.MinTxOpt >= 0 Then
    UsingMinTax = CInt(TaxMasterRec.MinTxOpt)
    MinBill = TaxMasterRec.MinBill
    Select Case UsingMinTax
      Case 0:
        lblSupp1.Caption = "* No minimum tax handling."
      Case 1:
        lblSupp1.Caption = "* Reduce tax bills to zero if bill is less than or equal to " + QPTrim$(Using("$#,##0.00", MinBill)) + "."
      Case 2:
        lblSupp1.Caption = "* Make tax bills " + QPTrim$(Using("$#,##0.00", MinBill)) + " if bill is less than  " + QPTrim$(Using("$#,##0.00", MinBill)) + "."
    End Select
  End If
  
  ThisTown = QPTrim$(TaxMasterRec.Name)
  fpcmbPrintOrder.Text = "1) Account Number Order"
  fpcmbPrintOrder.AddItem "1) Account Number Order"
  fpcmbPrintOrder.AddItem "2) Customer Name Order"
  fpcmbPrintOrder.AddItem "3) Search Name Order"
  fpcmbPrintOrder.AddItem "4) Social Security Order"
  ThisOpt = QPTrim$(TaxMasterRec.OptSrchCust)
  If ThisOpt <> "" Then
    fpcmbPrintOrder.AddItem "5) " + ThisOpt + " Order"
  End If
  
  fpcmbInclNoBills.Text = "N"
  fpcmbInclNoBills.AddItem "N"
  fpcmbInclNoBills.AddItem "Y"
  fpcmbPrintOpt.Text = "Graphical"
  fpcmbPrintOpt.AddItem "Graphical"
  fpcmbPrintOpt.AddItem "Text"
  
  fpcmbSuppOnly.Text = "N"
  fpcmbSuppOnly.AddItem "N"
  fpcmbSuppOnly.AddItem "Y"
  OptDiscNo.Value = True
  fptxtDiscXDate.Enabled = False
  fpcmbTownships.Clear
  fpcmbTownships.Text = "ALL TOWNSHIPS"
  fpcmbTownships.AddItem "ALL TOWNSHIPS"
  
  If Exist(TaxTownships) Then
    OpenTownshipFile TSHandle, TSCnt
    For x = 1 To TSCnt
      Get TSHandle, x, TSRec
      fpcmbTownships.AddItem QPTrim$(TSRec.TownShip) + " ONLY"
    Next x
  End If
  Close TSHandle
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPrebilling", "LoadMe", Erl)
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

Private Sub fpcmbCounty_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbCounty.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbCounty.ListIndex = -1
  End If
  If fpcmbCounty.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      If fpcmbSplit.Enabled = True Then
        fpcmbSplit.SetFocus
      Else
        OptDiscNo.SetFocus
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

Private Sub fpcmbCycle_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbCycle.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbCycle.ListIndex = -1
  End If
  If fpcmbCycle.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      If fpcmbCounty.Enabled = True Then
        fpcmbCounty.SetFocus
      ElseIf fpcmbSplit.Enabled = True Then
        fpcmbSplit.SetFocus
      Else
        OptDiscNo.SetFocus
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

Private Sub fpcmbInclNoBills_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbInclNoBills.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbInclNoBills.ListIndex = -1
  End If
  If fpcmbInclNoBills.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtBillDate.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
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
      fpcmbPrintOrder.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbPrintOrder_Change()
  If ThisOpt <> "" Then
    If InStr(fpcmbPrintOrder.Text, "5") Then
      UseOpt = "Y"
    Else
      UseOpt = "N"
    End If
  End If
  If InStr(fpcmbPrintOrder.Text, "4") Then
    UseSS = "Y"
  Else
    UseSS = "N"
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
      fpcmbSplit.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbSplit_Change()
  If fpcmbSplit.Text = "REAL" Then
    fpDblSnglRealRate.Enabled = True
    fpDblSnglPersRate.Enabled = False
    fpDblSnglFERate.Enabled = False
    fpDblSnglMHRate.Enabled = False
    fpDblSnglMCRate.Enabled = False
    fpDblSnglMTRate.Enabled = False
    Label13.Caption = "Real Discount Expiration Date"
    fpDblSnglLateList.Enabled = True
  Else
    fpDblSnglRealRate.Enabled = False
    fpDblSnglPersRate.Enabled = True
    fpDblSnglFERate.Enabled = True
    fpDblSnglMHRate.Enabled = True
    fpDblSnglMCRate.Enabled = True
    fpDblSnglMTRate.Enabled = True
    Label13.Caption = "Personal Discount Expiration Date"
    fpDblSnglLateList.Enabled = False
    fpDblSnglLateList.Value = 0
  End If
End Sub

Private Sub fpcmbSplit_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbSplit.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbSplit.ListIndex = -1
  End If
  If fpcmbSplit.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtRCurrYear.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbTownships_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbTownships.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbTownships.ListIndex = -1
  End If
  If fpcmbTownships.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      If fpcmbCycle.Enabled = True Then
        fpcmbCycle.SetFocus
      ElseIf fpcmbCounty.Enabled = True Then
        fpcmbCounty.SetFocus
      ElseIf fpcmbSplit.Enabled = True Then
        fpcmbSplit.SetFocus
      Else
        OptDiscNo.SetFocus
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

Private Function CheckPreReq() As Boolean
  Dim ThisYear As Integer
  Dim LenThisYear As Integer
  
  On Error GoTo ERRORSTUFF
  
  CheckPreReq = True 'nothing bad found yet
  WhatRYear = CInt(fptxtRCurrYear.Text)
  WhatPYear = CInt(fptxtPCurrYear.Text)
  BillNum& = 1 'fpDblSnglStartBill.Value
  RealRate# = fpDblSnglRealRate.Value
  PERSRATE# = fpDblSnglPersRate.Value
  LateList# = fpDblSnglLateList.Value
  Order$ = Mid(fpcmbPrintOrder.Text, 1, 1)
  SupOnly = QPTrim$(fpcmbSuppOnly.Text)
  RptType$ = Mid(fpcmbPrintOpt.Text, 1, 1)

  LenThisYear = Len(Date)
  LenThisYear = LenThisYear - 3
  ThisYear = Mid(Date, LenThisYear, Len(Date))
  If Abs(WhatRYear - ThisYear) > 10 Then
    If TaxMsgWOpts(800, "The real current year entered is more than 10 years from the current year. Press F10 if you wish to continue anyway. Otherwise, press ESC to review.", "F10 Continue", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtRCurrYear.SetFocus
      Close
      CheckPreReq = False
      Exit Function
    Else
      Unload frmVATaxMsgWOpts
      MainLog ("WARNING: frmVATaxPrebilling - User warned that the real date entered " + CStr(WhatRYear) + " is ten years more than today. The user continued saving anyway.")
    End If
  End If
  
  If Abs(WhatPYear - ThisYear) > 10 Then
    If TaxMsgWOpts(800, "The personal current year entered is more than 10 years from the current year. Press F10 if you wish to continue anyway. Otherwise, press ESC to review.", "F10 Continue", "ESC Review") = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtPCurrYear.SetFocus
      Close
      CheckPreReq = False
      Exit Function
    Else
      Unload frmVATaxMsgWOpts
      MainLog ("WARNING: frmVATaxPrebilling - User warned that the personal date entered " + CStr(WhatPYear) + " is ten years more than today. The user continued saving anyway.")
    End If
  End If
  
  If fpcmbSplit.Text = "REAL" Then
    If fpDblSnglRealRate.Value = 0 Then
      If TaxMsgWOpts(800, "The tax rate entered is zero. If you wish to continue anyway then press F10. Otherwise, press ESC to review.", "F10 Continue", "ESC Review") = "abort" Then
        Unload frmVATaxMsgWOpts
        fpDblSnglRealRate.SetFocus
        CheckPreReq = False
        Close
        Exit Function
      Else
        Unload frmVATaxMsgWOpts
        MainLog ("WARNING: frmVATaxPrebilling - User warned that the tax rate entered is zero but the user continued running the prebilling register anyway.")
      End If
    End If
  End If
  
  If fpcmbSplit.Text = "PERSONAL" Then
    If fpDblSnglPersRate.Value = 0 And fpDblSnglFERate.Value = 0 And fpDblSnglMHRate.Value = 0 And fpDblSnglMCRate.Value = 0 And fpDblSnglMTRate.Value = 0 Then
      If TaxMsgWOpts(800, "All tax rates entered are zero. If you wish to continue anyway then press F10. Otherwise, press ESC to review.", "F10 Continue", "ESC Review") = "abort" Then
        Unload frmVATaxMsgWOpts
        fpDblSnglPersRate.SetFocus
        CheckPreReq = False
        Close
        Exit Function
      Else
        Unload frmVATaxMsgWOpts
        MainLog ("WARNING: frmVATaxPrebilling - User warned that all tax rates entered are zero but the user continued running the personal prebilling register anyway.")
      End If
    End If
  End If
  
  Exit Function
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPrebilling", "CheckPreReq", Erl)
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
  
End Function

Private Sub MakeCustIdx(IdxType As Integer)
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  Dim HoldName$
  Dim HoldLong As Long
  Dim HoldNum As Long
  Dim Nextx As Long
  Dim BigName$
  Dim SmallName$
  Dim ThisSS$
  Dim Thisx As Long
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim OptFlag As Boolean
  Dim SSRec As SocSecIdxType
  Dim SSHandle As Integer
  Dim NumOfSSRecs As Long
  
  On Error GoTo ERRORSTUFF
  If UseOpt = "Y" Then
    OpenCustOptSearchFile OHandle, NumOfORecs
    If NumOfORecs = 0 Then
      Call TaxMsg(900, "There are no customers with an optional search description.")
      Close OHandle
      Exit Sub
    End If
    ReDim CustArr(1 To NumOfORecs) As Long
    OptCnt = NumOfORecs
    For x = 1 To NumOfORecs
      Get OHandle, x, OptRec
      CustArr(x) = OptRec.CustRec
    Next x
    Close OHandle
    UsingIdx = True
    Exit Sub
  End If
    
  If UseSS = "Y" Then
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
    OpenSocSecIdxFile SSHandle, NumOfSSRecs
    If NumOfSSRecs = 0 Then
      Call TaxMsg(900, "There are no customers with social security numbers saved.")
      Close SSHandle
      Exit Sub
    End If
    ReDim CustArr(1 To NumOfSSRecs) As Long
    SSCnt = NumOfSSRecs
    For x = 1 To NumOfSSRecs
      Get SSHandle, x, SSRec
      CustArr(x) = SSRec.CustRec
    Next x
    Close SSHandle
    UsingIdx = True
    Exit Sub
  End If
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  ReDim CustArr(1 To NumOfTCRecs) As Long
  If IdxType = 1 Then GoSub AcctNum
  
  ReDim CustName(1 To NumOfTCRecs) As String
  BigName = ""
  
  frmVATaxShowPctComp.Label1 = "Creating Customer Index"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
      CustArr(x) = x
      Select Case IdxType
        Case 2
          If QPTrim$(TaxCust.CustName) > BigName Then
            BigName = QPTrim$(TaxCust.CustName)
          End If
          CustName(x) = QPTrim$(TaxCust.CustName)
          UsingIdx = True
        Case 3
          If QPTrim$(TaxCust.SName) > BigName Then
            BigName = QPTrim$(TaxCust.SName)
          End If
          CustName(x) = QPTrim$(TaxCust.SName)
          UsingIdx = True
        Case 4
          ThisSS = ReplaceString(TaxCust.CSSN, "-", "")
          If ThisSS > BigName Then
            BigName = ThisSS
          End If
          CustName(x) = ThisSS
          UsingIdx = True
      End Select
    frmVATaxShowPctComp.ShowPctComp x, NumOfTCRecs
  Next x
  Close TCHandle
  
  Unload frmVATaxShowPctComp
  
  If QPTrim$(BigName) = "" Then
    Select Case IdxType
      Case 2
        Call TaxMsg(800, "The selection of customer name order is not possible because no customer names could be found. Print order will default to customer number order.")
        Close
        Exit Sub
      Case 3
        Call TaxMsg(800, "The selection of customer search name order is not possible because no customer search names could be found. Print order will default to customer number order.")
        Close
        Exit Sub
      Case 4
        Call TaxMsg(800, "The selection of customer social security number order is not possible because no customer social security numbers could be found. Print order will default to customer number order.")
        Close
        Exit Sub
    End Select
  End If
  
  SmallName = BigName + "z"
  Nextx = 1
  
  frmVATaxShowPctComp.Label1 = "Creating Customer Index"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  
  Do
    For x = Nextx To NumOfTCRecs
      If CustName(x) < SmallName Then
        SmallName = CustName(x)
        Thisx = x
      End If
    Next x
    HoldName = CustName(Thisx)
    HoldLong = CustArr(Thisx)
    CustName(Thisx) = CustName(Nextx)
    CustArr(Thisx) = CustArr(Nextx)
    CustName(Nextx) = HoldName
    CustArr(Nextx) = HoldLong
    Nextx = Nextx + 1
    SmallName = BigName
    If Nextx > NumOfTCRecs Then Exit Do
    frmVATaxShowPctComp.ShowPctComp Nextx, NumOfTCRecs
  Loop
  Unload frmVATaxShowPctComp
        
  Exit Sub
  
AcctNum:
  Dim BigNum As Long
  Dim SmallNum As Long
  BigNum = 0
  ReDim CustNum(1 To NumOfTCRecs) As Long
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    CustArr(x) = x
    CustNum(x) = TaxCust.PIN
    If TaxCust.PIN > BigNum Then
      BigNum = TaxCust.PIN
    End If
  Next x
  
  SmallNum = BigNum + 1
  Nextx = 1
  frmVATaxShowPctComp.Label1 = "Creating Customer Index"
  frmVATaxShowPctComp.cmdCancel.Visible = False
  frmVATaxShowPctComp.Show , Me
  
  Do
    For x = Nextx To NumOfTCRecs
      If CustNum(x) < SmallNum Then
        SmallNum = CustNum(x)
        Thisx = x
      End If
    Next x
    HoldNum = CustNum(Nextx)
    HoldLong = CustArr(Nextx)
    CustNum(Nextx) = CustNum(Thisx)
    CustArr(Nextx) = CustArr(Thisx)
    CustNum(Thisx) = HoldNum
    CustArr(Thisx) = HoldLong
    Nextx = Nextx + 1
    If Nextx > NumOfTCRecs Then Exit Do
    SmallNum = BigNum
    frmVATaxShowPctComp.ShowPctComp Nextx, NumOfTCRecs
  Loop
  Unload frmVATaxShowPctComp
  Close TCHandle
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPrebilling", "MakeCustIdx", Erl)
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

Private Sub PrintGraphics()
  Dim TBillRec As VARETaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim OPBillRec As TaxTransactionType
  Dim OPHandle As Integer
  Dim NumOfOPRecs As Long
  Dim RealRec As PropertyRecType
  Dim RRHandle As Integer
  Dim NumOfRRREcs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  Dim NewTBillRec As VARETaxBillType
  Dim LateAmt#
  Dim Inactive As Integer
  Dim PastFlagSet As Boolean
  Dim NoProp As Integer
  Dim CustName$, CitySt$, CustAcct&
  Dim ABalance#
  Dim Balance#
  Dim TransRecord&
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TempAddOn As TempTaxBillAddOn
  Dim AOHandle As Integer
  Dim AddOnCnt As Integer
  Dim MortCodeRec As MortCodeRecType
  Dim MCHandle As Integer
  Dim NumOfMCRecs As Integer
  Dim RealValue#
  Dim RealExmp#
  Dim RealCalcVal#
  Dim RealTaxDue#
  Dim MAPBLKLOT$
  Dim ThisProp&
  Dim CustMort$
  Dim MortCnt As Integer
  Dim ThisMort$
  Dim NextRealRec&
  Dim Discovery$
  Dim TotalReal#
  Dim TotalOverPay#
  Dim TotalEx#
  Dim NumBills&
  Dim RptHandle As Integer
  Dim TaxPreRptFile As String
  Dim TValue#
  Dim dlm$
  Dim TotalBills#
  Dim TotalLate#
  Dim TotalPast#
  Dim TotalBldgVal#
  Dim ThisTownship$
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim ThisTaxRec As Long
  Dim CountyNum As Long
  Dim CountyName$
  Dim CycleNum As Long
  Dim CycleName$
  Dim BillCnt As Long
  Dim SplitYN$
  Dim OptRevTax1 As Double
  Dim OptRevTax2 As Double
  Dim OptRevTax3 As Double
  Dim OptRev1Desc$
  Dim OptRev2Desc$
  Dim OptRev3Desc$
  Dim TotOpt1 As Double
  Dim TotOpt2 As Double
  Dim TotOpt3 As Double
  Dim TotReal As Double
  Dim OverPay As Boolean
  Dim OverPayAmt As Double
  Dim OPApplied As Double
  Dim ThisTBCnt As Long
  Dim ThisTest$
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim OptFlag As Boolean
  Dim PERC!
  Dim MultiYear As Integer
  Dim NextRec As Long
  Dim BldgValue#
  Dim GTotValue#
  Dim ThisPastDue#
  Dim Pct1#, Pct2#, Pct6#, Pct7#, Pct8#
  Dim PctTot#, PctTest#, ThisReal#
'  Dim AHandle As Integer
  
  On Error GoTo ERRORSTUFF
'  AHandle = FreeFile
'  Open "prebillreal.dat" For Output As AHandle
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  MultiYear = TaxMasterRec.MultiYear
  TotOpt1 = 0
  TotOpt2 = 0
  TotOpt3 = 0
  OptRev1Desc = QPTrim$(TaxMasterRec.OptRev1)
  OptRev2Desc = QPTrim$(TaxMasterRec.OptRev2)
  OptRev3Desc = QPTrim$(TaxMasterRec.OptRev3)
  
  fpcmbCycle.Col = 1
  CycleName = QPTrim$(fpcmbCycle.ColText)
  If CycleName = "" Then CycleName = "N/A"
  If fpcmbCycle.Enabled = True And CycleName <> "NO CYCLE" Then
    fpcmbCycle.Col = 0
    CycleNum = CLng(fpcmbCycle.ColText)
  Else
    CycleNum = -1
  End If
  
  fpcmbCounty.Col = 1
  CountyName = QPTrim$(fpcmbCounty.ColText)
  If CountyName = "" Then CountyName = "N/A"
  If fpcmbCounty.Enabled = True And CountyName <> "ALL COUNTIES" Then
    fpcmbCounty.Col = 0
    CountyNum = CLng(fpcmbCounty.ColText)
  Else
    CountyNum = -1
  End If
  
  If QPTrim$(fpcmbTownships.Text) = "ALL TOWNSHIPS" Then
    ThisTownship = "ALL"
  Else
    ThisTownship = QPTrim$(fpcmbTownships.Text)
    ThisTownship = Mid(ThisTownship, 1, Len(ThisTownship) - 4)
    ThisTownship = QPTrim$(ThisTownship)
  End If
  dlm$ = "~"
  If Exist(RealTaxBillFile) Then
    KillFile RealTaxBillFile
  End If
  If Exist(RealTaxBillOPFile) Then
    KillFile RealTaxBillOPFile
  End If
  If Exist("TMPBLADD.DAT") Then 'tax bill addon
    KillFile "TMPBLADD.DAT"
  End If
  
  AddOnCnt = 0
  OpenTaxBillAddOn AOHandle
  OpenRealTaxBillFile TBHandle, NumOfTBRecs
  OpenTaxPropFile RRHandle, NumOfRRREcs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenMortCodeFile MCHandle, NumOfMCRecs
  
  Inactive = 0
  
  frmVATaxShowPctComp.Label1 = "Creating Real Tax Pre-Billing Register"
  frmVATaxShowPctComp.Show , Me
  frmVATaxShowPctComp.cmdCancel.Visible = False
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  
  If UseOpt = "Y" Then
    NumOfTCRecs = OptCnt
  End If
  If UseSS = "Y" Then
    NumOfTCRecs = SSCnt
  End If
  
'  GoSub PPTRA
  
  For x = 1 To NumOfTCRecs
    TBillRec = NewTBillRec
    If UsingIdx = True Then
      Get TCHandle, CustArr(x), TaxCust
      ThisTaxRec = CustArr(x)
    Else
      Get TCHandle, x, TaxCust
      ThisTaxRec = x
    End If
    If TaxCust.FirstPropRec = 0 Then GoTo PreBillSkip
    If TaxCust.TaxExempt = "Y" Then
      GoTo PreBillSkip
    End If
    If ThisTownship <> "ALL" Then
      If QPTrim$(TaxCust.TownShip) <> ThisTownship Then
        GoTo PreBillSkip
      End If
    End If
    
    If CycleNum >= 0 Then
      If TaxCust.Cycle <> CycleNum Then
        GoTo PreBillSkip
      End If
    End If
    
    If CountyNum >= 0 Then
      If TaxCust.County4BillNum <> CountyNum Then
        GoTo PreBillSkip
      End If
    End If
    
    LateAmt# = 0
    OPApplied = 0
    If TaxCust.Deleted <> 0 Then
      GoTo PreBillSkip:
    End If
    If QPTrim$(TaxCust.Active) <> "Y" Then
      Inactive = Inactive + 1
      GoTo PreBillSkip:
    End If
    PastFlagSet = 0             'Initialize Past Balance Flag
    OverPayAmt = 0
    If UsingIdx = True Then
      OverPayAmt = GetCustRealBalance(CustArr(x), -1)
      If OverPayAmt < 0 Then
        OverPay = True
      Else
        OverPayAmt = 0
        OverPay = False
      End If
    Else
      OverPayAmt = GetCustRealBalance(x, -1)
      If OverPayAmt < 0 Then
        OverPay = True
      Else
        OverPayAmt = 0
        OverPay = False
      End If
    End If
    ThisTest = CStr(OverPayAmt)
    If InStr(ThisTest, "E") Then OverPayAmt = 0
    
    If TaxCust.FirstPropRec <= 0 Then
      NoProp = 1
      GoSub SetCustInfo
      GoSub WriteIt2Disk
      GoTo PreBillSkip
    End If
    
    NoProp = 0
    GoSub SetCustInfo
    GoSub GetRealInfo
    
PreBillSkip:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTCRecs
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
  
  Close RRHandle
  Close TCHandle
  Close MCHandle
  Close AOHandle
  
  If ThisTBCnt = 0 Then
    Call TaxMsg(900, "Using the parameters selected there are no customers who qualify for a tax charge.")
    Close
    fptxtRCurrYear.SetFocus
    Exit Sub
  End If
  
  TotalReal# = 0
  TotalEx# = 0
  NumBills& = 0
  
  TaxPreRptFile$ = "TAXRPTS\TaxRealPreBill.RPT"
  
  RptHandle = FreeFile
  Open TaxPreRptFile For Output As #RptHandle
  
  Close TBHandle
  OpenRealTaxBillFile TBHandle, NumOfTBRecs
  
  For x = 1 To NumOfTBRecs
    Get TBHandle, x, TBillRec
'    If Discovery$ = "Y" And TBillRec.BillNumber = -1 Then
    If TBillRec.BillNumber = -1 Then 'took out Discovery$ = "Y" on 12/5/06
       GoTo NotThisOne
    Else
      If TBillRec.TotalBillDue = 0 And TBillRec.PriorYrBalance = 0 Then
        If fpcmbInclNoBills.Text = "N" Then
          GoTo NotThisOne
        End If
      End If
'      Print #AHandle, CStr(TBillRec.CustRec) + "~" + Using$("$###,###,##0.00", TBillRec.RealTaxDue)
      '                        0
      Print #RptHandle, TBillRec.CustRec; dlm;
      '                            1
      Print #RptHandle, QPTrim$(TBillRec.CustName); dlm;
        '                            2
        Print #RptHandle, TBillRec.PriorYrBalance; dlm;
      If TBillRec.BillNumber = -1 Then
        '                   3
        Print #RptHandle, "N/A"; dlm;
      Else
        '                          3
        Print #RptHandle, TBillRec.BillNumber; dlm;
      End If
      '                           4
'      Print #RptHandle, OldRound(TBillRec.TotalBillDue); dlm;
'      Print #RptHandle, OldRound(TBillRec.RealTaxDue + TBillRec.LateTaxDue); dlm;  'comment 8/15/06
      Print #RptHandle, OldRound(TBillRec.RealTaxDue - TBillRec.LateTaxDue); dlm; 'made "- late tax due" on 8/15/06
      '                           5
      Print #RptHandle, TBillRec.RealValue; dlm;
      '                           6
      Print #RptHandle, TBillRec.ExptValue; dlm;
      TValue# = OldRound#(TBillRec.RealValue + TBillRec.BldgValue - TBillRec.ExptValue)
      If TValue# < 0 Then TValue# = 0
      '                    7
      Print #RptHandle, TValue#; dlm;
      '                           8
      Print #RptHandle, TBillRec.LateTaxDue; dlm;
      TotalReal# = OldRound#(TotalReal# + TBillRec.RealValue)
      TotalEx# = OldRound#(TotalEx# + TBillRec.ExptValue)
      TotalBills# = OldRound#(TotalBills# + TBillRec.RealTaxDue)
      TotalLate# = OldRound#(TotalLate# + TBillRec.LateTaxDue)
'      TotalPast# = OldRound#(TotalPast# + GetCustRealBalance(TBillRec.CustRec, -1))
      TotalBldgVal = OldRound(TotalBldgVal + TBillRec.BldgValue)
      TotalOverPay# = OldRound#(TotalOverPay# + TBillRec.OverPayAmt)

      If TBillRec.TotalBillDue > 0 Then
        NumBills& = NumBills& + 1
      End If
      '                     9
      Print #RptHandle, NumBills&; dlm;
      '                     10
      Print #RptHandle, TotalReal#; dlm;
      '                    11
      Print #RptHandle, TotalEx#; dlm;
      '                     12
      Print #RptHandle, TotalBills#; dlm;
      '                     13
      Print #RptHandle, TotalPast#; dlm;
      '                                   14
      Print #RptHandle, OldRound#(TotalPast# + TotalBills#); dlm;
      '                     15
      Print #RptHandle, TotalLate#; dlm;
      '                    16
      Print #RptHandle, Inactive; dlm;
      '                         17
      Print #RptHandle, TBillRec.RealPin; dlm;
      '                     18
      Print #RptHandle, ThisTown; dlm;
      '                    19
      Print #RptHandle, WhatRYear; dlm;
      '                     20                21              22
      Print #RptHandle, ThisTownship; dlm; CycleName; dlm; CycleNum; dlm;
      '                     23                24
      Print #RptHandle, CountyName$; dlm; CountyNum; dlm;
      '                     25                26                27
      Print #RptHandle, TBillRec.OptRevTax1; dlm; TBillRec.OptRevTax2; dlm; TBillRec.OptRevTax3; dlm;
       '                      28                 29                30
      Print #RptHandle, OptRev1Desc$; dlm; OptRev2Desc$; dlm; OptRev3Desc$; dlm;
      '                   31             32           33                     34                                  35
      Print #RptHandle, TotOpt1; dlm; TotOpt2; dlm; TotOpt3; dlm; -TBillRec.OverPayAmt; dlm; OldRound(TotalBills# - TotalOverPay); dlm;
      '                      36                          37                                        38
      Print #RptHandle, TotalOverPay#; dlm; CDbl(fpDblSnglRealRate.Value) / 100; dlm; CDbl(fpDblSnglLateList.Value) / 100; dlm;
      '                     39                40                  41
      Print #RptHandle, TotalBldgVal; dlm; MultiYear; dlm; TBillRec.BldgValue
    End If      'Test for Discovery Bills
NotThisOne:
  Next x
  
  Close
  
  arVATaxPreBillLS.Show
  frmVATaxLoadReport.Show
  KillFile "txrblsprn.dat"
  
  MainLog ("Prebilling graphics report for real property generated.")
  Exit Sub
  
SetCustInfo:
  ThisPastDue = 0
  TBillRec.CustRec = ThisTaxRec
  CustName$ = QPTrim$(TaxCust.CustName)
  TBillRec.CustName = CustName$
  TBillRec.CustAdd1 = QPTrim$(TaxCust.Addr1)
  TBillRec.CustAdd2 = QPTrim$(TaxCust.Addr2)
  CitySt$ = QPTrim$(TaxCust.City) + " " + TaxCust.State
  TBillRec.CustAdd3 = CitySt$
  TBillRec.CustZip = TaxCust.Zip
  TBillRec.CustPin = TaxCust.PIN
  TBillRec.TaxYear = WhatRYear
'  TBillRec.RDesc3 = TaxCust.CSSN 'commented out 10/25/06
  
  'Set Prior Balance if any
  ABalance = 0
  ABalance = GetCustRealBalance(TaxCust.Acct, -1) 'added this line on 8/11/2006    0
'  GoSub GetPastBalance
'  If ABalance# > 0 Then 'comment 8/14/06
    If PastFlagSet = 0 Then
      TBillRec.PriorYrBalance = ABalance#
    End If
    PastFlagSet = 1
'  End If 'comment 8/14/06
  Return
  
GetPastBalance:
  
  Balance# = 0
  ABalance# = 0
  
  If TaxCust.LastTrans > 0 Then
    OpenTaxTransFile TTHandle, NumOfTTRecs
    TransRecord& = TaxCust.LastTrans
    Do While TransRecord& <> 0
      Get TTHandle, TransRecord&, TaxTrans
      If TaxTrans.TranType = 1 And TaxTrans.BillType = "R" Then
        Balance# = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
        Balance# = OldRound(Balance# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection)
        Balance# = OldRound(Balance# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3) 'added for vs 2.05
        Balance# = OldRound(Balance# + TaxTrans.Revenue.PrePaidAmt) 'added for vs 2.05
        Balance# = OldRound(Balance# - (TaxTrans.Revenue.LateListPd + TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
        Balance# = OldRound(Balance# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd))
        Balance# = OldRound(Balance# - TaxTrans.DiscAmt) 'added for vs 2.05
        Balance# = OldRound#(Balance#)
      End If
      ABalance# = ABalance# + Balance#
      Balance# = 0
      TransRecord& = TaxTrans.LastTrans
    Loop
    Close TTHandle
  End If

Return

'PPTRA:
'  If WhatYear$ = "1998" Then
'    PERC! = 12.5
'  ElseIf WhatYear$ = "1999" Then
'    PERC! = 27.5
'  ElseIf WhatYear$ = "2000" Then
'    PERC! = 47.5
'  ElseIf WhatYear$ = "2001" Or WhatYear$ = "2002" Then
'    PERC! = 70
'  Else
'    PERC! = TaxMasterRec.PPTRADisc
'  End If
'
'Return
  
WriteIt2Disk:   'write the info out to disk here.
  If ThisPastDue = 0 Then
    ThisPastDue = ABalance#
    TotalPast# = OldRound(TotalPast# + ThisPastDue#)
  End If
  TBillRec.BillPrinted = False
  TBillRec.SetDscvry2No = "N" 'added 12/4/06
  If TBillRec.TotalBillDue > 0 Then
    TBillRec.BillNumber = BillNum&
  Else
    TBillRec.BillNumber = -1
  End If
  If UsingMinTax = 1 Then
    If TBillRec.TotalBillDue <= MinBill And TBillRec.BillNumber <> -1 Then 'added billnumber on 8/17/06
      TempAddOn.OldAmt = TBillRec.TotalBillDue
      TempAddOn.CustName = QPTrim$(TaxCust.CustName)
      TempAddOn.CustRec = x
      TempAddOn.Type = "Tax bills less than or equal to " + QPTrim$(Using$("$#,##0.00", MinBill)) + " become zero."
      AddOnCnt = AddOnCnt + 1
      TBillRec.TotalBillDue = 0
      TBillRec.SetDscvry2No = "Y" 'added 12/4/06
      TBillRec.RealTaxDue = 0 'added 12/4/06
      TBillRec.LateTaxDue = 0 'added 12/4/06
      TotOpt1# = OldRound(TotOpt1# - TBillRec.OptRevTax1) 'added 12/4/06
      TBillRec.OptRevTax1 = 0 'added 12/4/06
      TotOpt2# = OldRound(TotOpt2# - TBillRec.OptRevTax2) 'added 12/4/06
      TBillRec.OptRevTax2 = 0 'added 12/4/06
      TotOpt3# = OldRound(TotOpt3# - TBillRec.OptRevTax3) 'added 12/4/06
      TBillRec.OptRevTax3 = 0 'added 12/4/06
      TempAddOn.NewAmt = 0
      Put AOHandle, AddOnCnt, TempAddOn
    End If
  ElseIf UsingMinTax = 2 Then
    If TBillRec.TotalBillDue < MinBill And TBillRec.BillNumber <> -1 Then 'added billnumber on 8/17/06
      TempAddOn.OldAmt = TBillRec.TotalBillDue
      TempAddOn.CustName = QPTrim$(TaxCust.CustName)
      TempAddOn.CustRec = x
      TempAddOn.Type = "Tax bills less than " + QPTrim$(Using$("$#,##0.00", MinBill)) + " become " + QPTrim$(Using$("$#,##0.00", MinBill)) + "."
      AddOnCnt = AddOnCnt + 1
      TempAddOn.NewAmt = MinBill
      Put AOHandle, AddOnCnt, TempAddOn
      Pct1# = 0
      Pct2# = 0
      Pct6# = 0
      Pct7# = 0
      Pct8# = 0
      ThisReal = OldRound(TBillRec.RealTaxDue - (TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3))
      PctTot# = OldRound(TBillRec.TotalBillDue)
      If ThisReal > 0 Then
        Pct1# = OldRound#(ThisReal / PctTot#)
      End If
      If TBillRec.OptRevTax1 > 0 Then
        Pct6# = OldRound#(TBillRec.OptRevTax1 / PctTot#)
      End If
      If TBillRec.OptRevTax2 > 0 Then '
        Pct7# = OldRound#(TBillRec.OptRevTax2 / PctTot#)
      End If
      If TBillRec.OptRevTax3 > 0 Then
        Pct8# = OldRound#(TBillRec.OptRevTax3 / PctTot#)
      End If
      If TBillRec.LateTaxDue > 0 Then
        Pct2# = OldRound#(TBillRec.LateTaxDue / PctTot#)
      End If
      
      ThisReal = MinBill * Pct1
      
      TotOpt1# = OldRound(TotOpt1# - TBillRec.OptRevTax1)
      OptRevTax1# = OldRound(OptRevTax1# - TBillRec.OptRevTax1)
      TBillRec.OptRevTax1 = OldRound(MinBill * Pct6)
      TotOpt1# = OldRound(TotOpt1# + TBillRec.OptRevTax1)
      OptRevTax1# = OldRound(OptRevTax1# + TBillRec.OptRevTax1)
      
      TotOpt2# = OldRound(TotOpt2# - TBillRec.OptRevTax2)
      OptRevTax2# = OldRound(OptRevTax2# - TBillRec.OptRevTax2)
      TBillRec.OptRevTax2 = OldRound(MinBill * Pct7)
      TotOpt2# = OldRound(TotOpt2# + TBillRec.OptRevTax2)
      OptRevTax2# = OldRound(OptRevTax2# + TBillRec.OptRevTax2)
      
      TotOpt3# = OldRound(TotOpt3# - TBillRec.OptRevTax3)
      OptRevTax3# = OldRound(OptRevTax3# - TBillRec.OptRevTax3)
      TBillRec.OptRevTax3 = OldRound(MinBill * Pct8)
      TotOpt3# = OldRound(TotOpt3# + TBillRec.OptRevTax3)
      OptRevTax3# = OldRound(OptRevTax3# + TBillRec.OptRevTax3)
      
      TotalLate# = OldRound(TotalLate# - TBillRec.LateTaxDue)
      LateAmt# = OldRound(LateAmt# - TBillRec.LateTaxDue)
      TBillRec.LateTaxDue = OldRound(MinBill * Pct2)
      TotalLate# = OldRound(TotalLate# + TBillRec.LateTaxDue)
      LateAmt# = OldRound(LateAmt# + TBillRec.LateTaxDue)
      
      PctTest# = OldRound(ThisReal + TBillRec.LateTaxDue)
      PctTest# = OldRound(PctTest# + TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3)
      
      If MinBill < PctTest# Then
        If ThisReal > PctTest# - MinBill Then
          ThisReal = OldRound(ThisReal - (PctTest# - MinBill))
          TotalReal# = OldRound(TotalReal# - (PctTest# - MinBill))
        ElseIf TBillRec.OptRevTax1 > PctTest# - MinBill Then
          TBillRec.OptRevTax1 = OldRound(TBillRec.OptRevTax1 - (PctTest# - MinBill))
          TotOpt1# = OldRound(TotOpt1# - (PctTest# - MinBill))
          OptRevTax1# = OldRound(OptRevTax1# - (PctTest# - MinBill))
        ElseIf TBillRec.OptRevTax2 > PctTest# - MinBill Then
          TBillRec.OptRevTax2 = OldRound(TBillRec.OptRevTax2 - (PctTest# - MinBill))
          TotOpt2# = OldRound(TotOpt2# - (PctTest# - MinBill))
          OptRevTax2# = OldRound(OptRevTax2# - (PctTest# - MinBill))
        ElseIf TBillRec.OptRevTax3 > PctTest# - MinBill Then
          TBillRec.OptRevTax3 = OldRound(TBillRec.OptRevTax3 - (PctTest# - MinBill))
          TotOpt3# = OldRound(TotOpt3# - (PctTest# - MinBill))
          OptRevTax3# = OldRound(OptRevTax3# - (PctTest# - MinBill))
        ElseIf TBillRec.LateTaxDue > PctTest# - MinBill Then
          TBillRec.LateTaxDue = OldRound(TBillRec.LateTaxDue - (PctTest# - MinBill))
          TotalLate# = OldRound(TotalLate# - (PctTest# - MinBill))
          LateAmt# = OldRound(LateAmt# - (PctTest# - MinBill))
        End If
      ElseIf MinBill > PctTest# Then
        If ThisReal > MinBill - PctTest# Then
          ThisReal = OldRound(ThisReal + (MinBill - PctTest#))
          TotalReal# = OldRound(TotalReal# + (MinBill - PctTest#))
        ElseIf TBillRec.OptRevTax1 > MinBill - PctTest# Then
          TBillRec.OptRevTax1 = OldRound(TBillRec.OptRevTax1 + (MinBill - PctTest#))
          TotOpt1# = OldRound(TotOpt1# + (PctTest# - MinBill))
          OptRevTax1# = OldRound(OptRevTax1# + (PctTest# - MinBill))
        ElseIf TBillRec.OptRevTax2 > MinBill - PctTest# Then
          TBillRec.OptRevTax2 = OldRound(TBillRec.OptRevTax2 + (MinBill - PctTest#))
          TotOpt2# = OldRound(TotOpt2# + (PctTest# - MinBill))
          OptRevTax2# = OldRound(OptRevTax2# + (PctTest# - MinBill))
        ElseIf TBillRec.OptRevTax3 > MinBill - PctTest# Then
          TBillRec.OptRevTax3 = OldRound(TBillRec.OptRevTax3 + (MinBill - PctTest#))
          TotOpt3# = OldRound(TotOpt3# + (PctTest# - MinBill))
          OptRevTax3# = OldRound(OptRevTax3# + (PctTest# - MinBill))
        ElseIf TBillRec.LateTaxDue > MinBill - PctTest# Then
          TBillRec.LateTaxDue = OldRound(TBillRec.LateTaxDue + (MinBill - PctTest#))
          TotalLate# = OldRound(TotalLate# + (MinBill - PctTest#))
          LateAmt# = OldRound(LateAmt# + (MinBill - PctTest#))
        End If
      End If
      TBillRec.RealTaxDue = OldRound(ThisReal + TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3 + TBillRec.LateTaxDue)
      TBillRec.TotalBillDue = MinBill
    End If
  End If
  
  ThisTBCnt = ThisTBCnt + 1
  Put TBHandle, ThisTBCnt, TBillRec
  ThisTBCnt = LOF(TBHandle) / Len(TBillRec)
'  If TBillRec.TotalBillDue > 0 Then '8/15/06
  OPApplied = 0 'added 10/17/08
  If TBillRec.RealTaxDue > 0 Then
    If Abs(OverPayAmt) > 0 Then
      OPBillRec.Revenue.LateListPd = 0
      OPBillRec.Revenue.Principle1Pd = 0
      OPBillRec.Revenue.RevOpt1Pd = 0
      OPBillRec.Revenue.RevOpt2Pd = 0
      OPBillRec.Revenue.RevOpt3Pd = 0
      OpenRealTaxBillOverPayFile OPHandle, NumOfOPRecs
      Get TBHandle, ThisTBCnt, TBillRec
      OPBillRec.Revenue.PrePaidAmt = Abs(OverPayAmt)
      OPBillRec.BelongTo = ThisTBCnt 'BillNum&'9/9/05 Billnum is assigned
      'at bill printing so using this as a way of reference needed when posting
      If TBillRec.LateTaxDue > 0 And Abs(OverPayAmt) > 0 Then
        If Abs(OverPayAmt) > TBillRec.LateTaxDue Then
          OverPayAmt = OldRound(OverPayAmt + TBillRec.LateTaxDue)
          OPBillRec.Revenue.LateListPd = TBillRec.LateTaxDue
          OPApplied = OldRound(OPApplied + OPBillRec.Revenue.LateListPd)
        ElseIf Abs(OverPayAmt) <= TBillRec.LateTaxDue Then
          OPBillRec.Revenue.LateListPd = -OverPayAmt
'          OPApplied = OldRound(OPApplied - OverPayAmt)'changed to line below on 10/17/08
          OPApplied = OPBillRec.Revenue.PrePaidAmt '-OverPayAmt 'added OPBillRec.Revenue.PrePaidAmt 8/23/09
          OverPayAmt = 0
        End If
      End If
      
      If OldRound(TBillRec.RealTaxDue - (TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3)) > 0 And Abs(OverPayAmt) > 0 Then
        If Abs(OverPayAmt) > OldRound(TBillRec.RealTaxDue - (TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3)) Then
          OverPayAmt = OldRound(OverPayAmt + TBillRec.RealTaxDue - (TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3))
          OPBillRec.Revenue.Principle1Pd = OldRound(OPBillRec.Revenue.Principle1Pd + TBillRec.RealTaxDue - (TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3))
          OPApplied = OldRound(OPApplied + TBillRec.RealTaxDue - (TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3))
        ElseIf Abs(OverPayAmt) <= OldRound(TBillRec.RealTaxDue - (TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3)) Then
          OPBillRec.Revenue.Principle1Pd = OPBillRec.Revenue.Principle1Pd - OverPayAmt
'          OPApplied = OldRound(OPApplied - OverPayAmt)'changed to line below on 10/17/08
          OPApplied = OPBillRec.Revenue.PrePaidAmt '-OverPayAmt 'added OPBillRec.Revenue.PrePaidAmt 8/23/09
          OverPayAmt = 0
        End If
      End If
      
      If TBillRec.OptRevTax1 > 0 And Abs(OverPayAmt) > 0 Then
        If Abs(OverPayAmt) > TBillRec.OptRevTax1 Then
          OverPayAmt = OldRound(OverPayAmt + TBillRec.OptRevTax1)
          OPBillRec.Revenue.RevOpt1Pd = TBillRec.OptRevTax1
          OPApplied = OldRound(OPApplied + TBillRec.OptRevTax1)
        ElseIf Abs(OverPayAmt) <= TBillRec.OptRevTax1 Then
          OPBillRec.Revenue.RevOpt1Pd = -OverPayAmt
'          OPApplied = OldRound(OPApplied - OverPayAmt)'changed to line below on 10/17/08
          OPApplied = OPBillRec.Revenue.PrePaidAmt '-OverPayAmt 'added OPBillRec.Revenue.PrePaidAmt 8/23/09
          OverPayAmt = 0
        End If
      End If
      
      If TBillRec.OptRevTax2 > 0 And Abs(OverPayAmt) > 0 Then
        If Abs(OverPayAmt) > TBillRec.OptRevTax2 Then
          OverPayAmt = OldRound(OverPayAmt + TBillRec.OptRevTax2)
          OPBillRec.Revenue.RevOpt2Pd = TBillRec.OptRevTax2
          OPApplied = OldRound(OPApplied + TBillRec.OptRevTax2)
        ElseIf Abs(OverPayAmt) <= TBillRec.OptRevTax2 Then
          OPBillRec.Revenue.RevOpt2Pd = -OverPayAmt
'          OPApplied = OldRound(OPApplied - OverPayAmt)'changed to line below on 10/17/08
          OPApplied = OPBillRec.Revenue.PrePaidAmt '-OverPayAmt 'added OPBillRec.Revenue.PrePaidAmt 8/23/09
          OverPayAmt = 0
        End If
      End If
      
      If TBillRec.OptRevTax3 > 0 And Abs(OverPayAmt) > 0 Then
        If Abs(OverPayAmt) > TBillRec.OptRevTax3 Then
          OverPayAmt = OldRound(OverPayAmt + TBillRec.OptRevTax3)
          OPBillRec.Revenue.RevOpt3Pd = TBillRec.OptRevTax3
          OPApplied = OldRound(OPApplied + TBillRec.OptRevTax3)
        ElseIf Abs(OverPayAmt) <= TBillRec.OptRevTax3 Then
          OPBillRec.Revenue.RevOpt3Pd = -OverPayAmt
'          OPApplied = OldRound(OPApplied - OverPayAmt)'changed to line below on 10/17/08
          OPApplied = OPBillRec.Revenue.PrePaidAmt '-OverPayAmt 'added OPBillRec.Revenue.PrePaidAmt 8/23/09
          OverPayAmt = 0
        End If
      End If
      OPBillRec.Amount = OPApplied
      Put OPHandle, NumOfOPRecs + 1, OPBillRec
      Close OPHandle
    End If '#1
    Get TBHandle, ThisTBCnt, TBillRec
    TBillRec.OverPayAmt = OPApplied
    TBillRec.MORTCODE = TBillRec.MORTCODE
    Put TBHandle, ThisTBCnt, TBillRec
    'moved #1 end if to above from here on 10/17/08
    BillNum& = BillNum& + 1
  End If
  
  Return
  
GetRealInfo:
  ThisProp& = TaxCust.FirstPropRec
  TBillRec.Opt1Desc = QPTrim$(TaxMasterRec.OptRev1)
  TBillRec.Opt2Desc = QPTrim$(TaxMasterRec.OptRev2)
  TBillRec.Opt3Desc = QPTrim$(TaxMasterRec.OptRev3)
  If ThisProp& > 0 Then
    Do While ThisProp > 0
      Get RRHandle, ThisProp&, RealRec
      If RealRec.Deleted = True Then GoTo NextOne
      If SupOnly = "Y" And RealRec.PROPDISC <> "Y" Then GoTo NextOne
      If RealRec.Mock = "Y" Then GoTo NextOne
      CustMort$ = QPTrim$(RealRec.MORTCODE)
      If Len(CustMort$) > 0 And QPTrim$(CustMort) <> "NONE" Then
        For MortCnt = 1 To NumOfMCRecs
          ThisMort$ = QPTrim$(MortCodes(MortCnt).MORTCODE)
          If ThisMort$ = CustMort$ Then
            TBillRec.MortRec = MortCodes(MortCnt).MortRec
            TBillRec.MORTCODE = ThisMort 'added 8/22/05
            Exit For
          End If
        Next MortCnt
      Else
        TBillRec.MortRec = 0
        TBillRec.MORTCODE = ""
      End If
    
'      If RealRec.LastYrPrinted = WhatRYear Then Discovery$ = "Y" 'took out 12/5/06
      If MultiYear > 1 Then GoTo GoAhead
      If MultiYear = 1 And (RealRec.LastYrPrinted <> WhatRYear) Or (RealRec.PROPDISC = "Y") Or (RealRec.LastYrPrinted = WhatRYear) Then
      'above line needs work
GoAhead:
        RealValue# = RealRec.PROPVALU
        BldgValue# = RealRec.BldgVal
        OptRevTax1# = OldRound(FigureOptRevTax1(ThisProp&, RRHandle, "R"))
        TBillRec.OptRevTax1 = OptRevTax1# / MultiYear
        TotOpt1 = OldRound(TotOpt1 + TBillRec.OptRevTax1) 'OptRevTax1#)
        OptRevTax2# = OldRound(FigureOptRevTax2(ThisProp&, RRHandle, "R"))
        TBillRec.OptRevTax2 = OptRevTax2# / MultiYear
        TotOpt2 = OldRound(TotOpt2 + TBillRec.OptRevTax2) 'OptRevTax2#)
        OptRevTax3# = OldRound(FigureOptRevTax3(ThisProp&, RRHandle, "R"))
        TBillRec.OptRevTax3 = OptRevTax3# / MultiYear
        TotOpt3 = OldRound(TotOpt3 + TBillRec.OptRevTax3) 'OptRevTax3#)
        RealExmp# = RealRec.EXMPOTHR
        If RealRec.EXMPOTHR > 0 Then
          TempAddOn.OldAmt = 0
          TempAddOn.CustName = QPTrim$(TaxCust.CustName)
          TempAddOn.CustRec = x
          TempAddOn.Type = "Other discount of " + QPTrim$(Using$("$#,##0.00", RealRec.EXMPOTHR)) + " applied to real estate tax."
          AddOnCnt = AddOnCnt + 1
          TempAddOn.NewAmt = RealRec.EXMPOTHR
          Put AOHandle, AddOnCnt, TempAddOn
        End If
        RealCalcVal# = OldRound#((RealValue# + BldgValue - RealExmp#) / 100)
        If RealCalcVal# < 0 Then RealCalcVal# = 0
        RealTaxDue# = OldRound#((RealCalcVal# * RealRate#) + OptRevTax1# + OptRevTax2# + OptRevTax3#)
        RealTaxDue# = OldRound#(RealTaxDue / MultiYear) 'added 6/22/06
        If RealTaxDue# = 0 Then
          If OldRound(TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3) = 0 Then 'added 10/18/06
            If fpcmbInclNoBills.Text = "N" Then GoTo NextOne
          End If
        End If
      
        TBillRec.ExptValue = OldRound#(TBillRec.ExptValue + RealExmp#)
        
        TBillRec.RealValue = RealValue#
        TBillRec.BldgValue = BldgValue#
        TotReal = TotReal + RealTaxDue#
        TBillRec.RealTaxDue = RealTaxDue#
        TBillRec.RealPropRecord = ThisProp&
        TBillRec.RealTaxRate = RealRate#
        TBillRec.RDesc1 = QPTrim$(RealRec.PROPNOT1) + " " + QPTrim$(RealRec.PROPNOT2)
        MAPBLKLOT$ = RealRec.Map + " " + RealRec.BLOCK + " " + RealRec.LOTNUMB

        TBillRec.RDesc2 = RealRec.PROPNOT2 ' MAPBLKLOT$ changed 10/25/06
        TBillRec.RDesc3 = RealRec.PROPNOT3 'changed 10/25/06
        TBillRec.CustPin = TaxCust.PIN
        TBillRec.InternalPin = RealRec.InternalPin
        TBillRec.RealPin = QPTrim$(RealRec.RealPin)
        If RealRec.LateList = "Y" Then
          LateAmt# = OldRound#((RealTaxDue#) * (LateList# / 100))
        Else
          LateAmt# = 0
        End If
        TBillRec.LateTaxDue = LateAmt# 'changed from line above on 10/13/06
        TBillRec.TotalBillDue = OldRound#(RealTaxDue# + LateAmt#) 'changed to this from above on 10/13/06
        TBillRec.DueDate = Date2Num(fptxtDueDate)
      
        GoSub WriteIt2Disk
      End If      'End of Test For Current Year Tax Bill
NextOne:
      ThisProp = RealRec.NextRec
    Loop
  End If
  
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPrebilling", "PrintGraphics", Erl)
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

Private Sub ParseTownships()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Integer, y As Integer
  Dim ThisTS As String
  
  On Error GoTo ERRORSTUFF
  
  ReDim Townships(1 To 1) As String
  OpenTaxCustFile TCHandle, NumOfTCRecs
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If TSCnt = 0 And Len(TaxCust.TownShip) > 0 Then
      TSCnt = TSCnt + 1
      Townships(TSCnt) = QPTrim$(TaxCust.TownShip)
    ElseIf Len(TaxCust.TownShip) > 0 Then
      ThisTS = QPTrim$(TaxCust.TownShip)
      For y = 1 To TSCnt
        If QPTrim$(Townships(y)) = ThisTS Then
          Exit For
        End If
      Next y
      If y > TSCnt Then
        TSCnt = TSCnt + 1
        ReDim Preserve Townships(1 To TSCnt) As String
        Townships(TSCnt) = ThisTS
      End If
    End If
  Next x
  Close TCHandle
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPrebilling", "ParseTownships", Erl)
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

Private Sub fpDblSnglFERate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyTab Then
    fpDblSnglMHRate.SetFocus
  End If
End Sub

Private Sub fpDblSnglMCRate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyTab Then
    fpDblSnglMTRate.SetFocus
  End If

End Sub

Private Sub fpDblSnglMHRate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyTab Then
    fpDblSnglMCRate.SetFocus
  End If
End Sub

Private Sub fpDblSnglPersRate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyTab Then
    fpDblSnglFERate.SetFocus
  End If
End Sub

Private Sub OptDiscNo_Click()
  If OptDiscNo.Value = True Then
    fptxtDiscXDate.Enabled = False
  ElseIf OptDiscYes.Value = True Then
    fptxtDiscXDate.Enabled = True
  End If
End Sub

Private Sub OptDiscYes_Click()
  If OptDiscNo.Value = True Then
    fptxtDiscXDate.Enabled = False
  ElseIf OptDiscYes.Value = True Then
    fptxtDiscXDate.Enabled = True
  End If
End Sub

Private Function FigureOptRevTax1(RecNum As Long, RHandle As Integer, ThisType$) As Double
  Dim RRealRec As PropertyRecType
  Dim PRealRec As PersonalRecType
  Dim RateRec As OptRevRateTablesType
  Dim RateHandle As Integer
  Dim NumOfRecs As Integer
  Dim x As Integer, y As Integer
  Dim ThisTax As Double
  Dim NetRPropVal#
  
  On Error GoTo ERRORSTUFF
  
  ThisTax = 0
  FigureOptRevTax1 = 0
  If ThisType = "R" Then
    Get RHandle, RecNum, RRealRec
    If RRealRec.OptRev1Chrg = 0 Then Exit Function
    OpenTaxRateTables RateHandle, NumOfRecs
    For x = 1 To NumOfRecs
      Get RateHandle, x, RateRec
      If RateRec.Deleted = True Then GoTo Deleted
      NetRPropVal = OldRound(RRealRec.PROPVALU - (RRealRec.EXMPOTHR + RRealRec.EXMPSENI))
'      If RRealRec.OptRev1Chrg = x Then
      If RRealRec.OptRev1Chrg = RateRec.OptRevNum Then 'went from x to RateRec.OptRevNum 10/22/07
        If RateRec.Type = "F" Then
          ThisTax = RateRec.FlatAmt
        ElseIf RateRec.Type = "S" Then
          For y = 1 To 10
            If NetRPropVal >= RateRec.FromAmt(y) And NetRPropVal <= RateRec.ToAmt(y) Then
              ThisTax = RateRec.TaxFAmt(y)
              Exit For
            End If
          Next y
          If y < 11 Then
            Exit For
          End If
        ElseIf RateRec.Type = "P" Then
          For y = 1 To 10
            If NetRPropVal >= RateRec.FromAmt(y) And NetRPropVal <= RateRec.ToAmt(y) Then
              ThisTax = OldRound(NetRPropVal * RateRec.TaxPAmt(y) / 100)
              Exit For
            End If
          Next y
          If y < 11 Then
            Exit For
          End If
        End If
      End If
Deleted:
    Next x
  ElseIf ThisType = "P" Then
    ThisTax = 0
    Get RHandle, RecNum, PRealRec
    If PRealRec.OptRev1Chrg = 0 Then Exit Function
    OpenTaxRateTables RateHandle, NumOfRecs
    For x = 1 To NumOfRecs
      Get RateHandle, x, RateRec
      If RateRec.Deleted = True Then GoTo Deleted2
'      If PRealRec.OptRev1Chrg = x + 3 Then '10/19/07 go back and remove 3 and replace with variable
      If PRealRec.OptRev1Chrg = RateRec.OptRevNum Then 'added 10/22/07
        ThisTax = RateRec.FlatAmt
        Exit For
      End If
Deleted2:
    Next x
  End If
  
  Close RateHandle
  FigureOptRevTax1 = ThisTax
  
  Exit Function

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPrebilling", "FigureOptRevTax1", Erl)
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

End Function

Private Function FigureOptRevTax2(RecNum As Long, RHandle As Integer, ThisType$) As Double
  Dim RRealRec As PropertyRecType
  Dim PRealRec As PersonalRecType
  Dim RateRec As OptRevRateTablesType
  Dim RateHandle As Integer
  Dim NumOfRecs  As Integer
  Dim x As Integer, y As Integer
  Dim ThisTax As Double
  Dim NetRPropVal#
  
  On Error GoTo ERRORSTUFF
  
  ThisTax = 0
  FigureOptRevTax2 = 0
  If ThisType = "R" Then
    Get RHandle, RecNum, RRealRec
    OpenTaxRateTables RateHandle, NumOfRecs
    For x = 1 To NumOfRecs
      Get RateHandle, x, RateRec
      If RateRec.Deleted = True Then GoTo Deleted
      NetRPropVal = OldRound(RRealRec.PROPVALU - (RRealRec.EXMPOTHR + RRealRec.EXMPSENI))
'      If RRealRec.OptRev2Chrg = x Then
      If RRealRec.OptRev2Chrg = RateRec.OptRevNum Then 'went from x to RateRec.OptRevNum 10/22/07
        If RateRec.Type = "F" Then
          ThisTax = RateRec.FlatAmt
        ElseIf RateRec.Type = "S" Then
          For y = 1 To 10
            If NetRPropVal >= RateRec.FromAmt(y) And NetRPropVal <= RateRec.ToAmt(y) Then
              ThisTax = RateRec.TaxFAmt(y)
              Exit For
            End If
          Next y
          If y < 11 Then
            Exit For
          End If
        ElseIf RateRec.Type = "P" Then
          For y = 1 To 10
            If NetRPropVal >= RateRec.FromAmt(y) And NetRPropVal <= RateRec.ToAmt(y) Then
              ThisTax = OldRound(NetRPropVal * RateRec.TaxPAmt(y) / 100)
              Exit For
            End If
          Next y
          If y < 11 Then
            Exit For
          End If
        End If
      End If
Deleted:
    Next x
  ElseIf ThisType = "P" Then
    ThisTax = 0
    Get RHandle, RecNum, PRealRec
    If PRealRec.OptRev2Chrg = 0 Then Exit Function
    OpenTaxRateTables RateHandle, NumOfRecs
    For x = 1 To NumOfRecs
      Get RateHandle, x, RateRec
      If RateRec.Deleted = True Then GoTo Deleted2
'      If PRealRec.OptRev2Chrg = x + 3 Then '10/19/07 go back and remove 3 and replace with variable
      If PRealRec.OptRev2Chrg = RateRec.OptRevNum Then ' + 3 Then'added 10/22/07
        ThisTax = RateRec.FlatAmt
        Exit For
      End If
Deleted2:
    Next x
  End If
  Close RateHandle
  FigureOptRevTax2 = ThisTax
  
  Exit Function

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPrebilling", "FigureOptRevTax2", Erl)
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


End Function

Private Function FigureOptRevTax3(RecNum As Long, RHandle As Integer, ThisType$) As Double
  Dim RRealRec As PropertyRecType
  Dim PRealRec As PersonalRecType
  Dim RateRec As OptRevRateTablesType
  Dim RateHandle As Integer
  Dim NumOfRecs As Integer
  Dim x As Integer, y As Integer
  Dim ThisTax As Double
  Dim NetRPropVal#
  
  On Error GoTo ERRORSTUFF
  
  ThisTax = 0
  FigureOptRevTax3 = 0
  If ThisType = "R" Then
  Get RHandle, RecNum, RRealRec
    OpenTaxRateTables RateHandle, NumOfRecs
    For x = 1 To NumOfRecs
      Get RateHandle, x, RateRec
      If RateRec.Deleted = True Then GoTo Deleted
      NetRPropVal = OldRound(RRealRec.PROPVALU - (RRealRec.EXMPOTHR + RRealRec.EXMPSENI))
'      If RRealRec.OptRev3Chrg = x Then
      If RRealRec.OptRev3Chrg = RateRec.OptRevNum Then 'went from x to RateRec.OptRevNum 10/22/07
        If RateRec.Type = "F" Then
          ThisTax = RateRec.FlatAmt
        ElseIf RateRec.Type = "S" Then
          For y = 1 To 10
            If NetRPropVal >= RateRec.FromAmt(y) And NetRPropVal <= RateRec.ToAmt(y) Then
              ThisTax = RateRec.TaxFAmt(y)
              Exit For
            End If
          Next y
          If y < 11 Then
            Exit For
          End If
        ElseIf RateRec.Type = "P" Then
          For y = 1 To 10
            If NetRPropVal >= RateRec.FromAmt(y) And NetRPropVal <= RateRec.ToAmt(y) Then
              ThisTax = OldRound(NetRPropVal * RateRec.TaxPAmt(y) / 100)
              Exit For
            End If
          Next y
          If y < 11 Then
            Exit For
          End If
        End If
      End If
Deleted:
    Next x
  ElseIf ThisType = "P" Then
    Get RHandle, RecNum, PRealRec
    If PRealRec.OptRev3Chrg = 0 Then Exit Function
    OpenTaxRateTables RateHandle, NumOfRecs
    For x = 1 To NumOfRecs
      Get RateHandle, x, RateRec
      If RateRec.Deleted = True Then GoTo Deleted2
'      If PRealRec.OptRev3Chrg = x + 3 Then '10/19/07 go back and remove 3 and replace with variable
      If PRealRec.OptRev3Chrg = RateRec.OptRevNum Then 'added 10/22/07'x + 3 Then
        ThisTax = RateRec.FlatAmt
        Exit For
      End If
Deleted2:
    Next x
  End If
    
  Close RateHandle
  FigureOptRevTax3 = ThisTax
  
  Exit Function
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPrebilling", "FigureOptRevTax3", Erl)
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

End Function

Private Sub PrintText()
  Dim TBillRec As VARETaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim OPBillRec As TaxTransactionType
  Dim OPHandle As Integer
  Dim NumOfOPRecs As Long
  Dim RealRec As PropertyRecType
  Dim RRHandle As Integer
  Dim NumOfRRREcs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  Dim NewTBillRec As VARETaxBillType
'  Dim Done2Disk As Integer
  Dim LateAmt#
  Dim Inactive As Integer
  Dim PastFlagSet As Boolean
  Dim NoProp As Integer
  Dim CustName$, CitySt$, CustAcct&
  Dim ABalance#
  Dim Balance#
  Dim TransRecord&
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TempAddOn As TempTaxBillAddOn
  Dim AOHandle As Integer
  Dim AddOnCnt As Integer
  Dim MortCodeRec As MortCodeRecType
  Dim MCHandle As Integer
  Dim NumOfMCRecs As Integer
  Dim RealValue#
  Dim RealExmp#
  Dim RealCalcVal#
  Dim RealTaxDue#
  Dim MAPBLKLOT$
  Dim ThisProp&
  Dim CustMort$
  Dim MortCnt As Integer
  Dim ThisMort$
  Dim NextRealRec&
  Dim Discovery$
  Dim TotalRealVal#
  Dim TotalOverPay#
  Dim TotalEx#
  Dim TotalBldgVal#
  Dim NumBills&
  Dim RptHandle As Integer
  Dim TaxPreRptFile As String
  Dim TValue#
  Dim dlm$
  Dim TotalBills#
  Dim TotalLate#
  Dim TotalReal#
  Dim TotalPast#
  Dim ThisTownship$
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim ThisTaxRec As Long
  Dim CountyNum As Long
  Dim CountyName$
  Dim CycleNum As Long
  Dim CycleName$
  Dim BillCnt As Long
  Dim SplitYN$
  Dim OptRevTax1 As Double
  Dim OptRevTax2 As Double
  Dim OptRevTax3 As Double
  Dim OptRev1Desc$
  Dim OptRev2Desc$
  Dim OptRev3Desc$
  Dim TotOpt1 As Double
  Dim TotOpt2 As Double
  Dim TotOpt3 As Double
  Dim FF$, MaxLines As Integer, LineCnt As Integer
  Dim Line$, Page As Integer, Dot$
  Dim OverPay As Boolean
  Dim OverPayAmt As Double
  Dim OPApplied As Double
  Dim ThisTBCnt As Long
  Dim ThisTest$, Line1$, Town$
  Dim ThisRate As Double
  Dim MultiYear As Integer
  Dim BldgValue#
  Dim GTotValue#
  Dim ThisPastDue#
  Dim Pct1#, Pct2#, Pct6#, Pct7#, Pct8#
  Dim PctTot#, PctTest#, ThisReal#
  
  On Error GoTo ERRORSTUFF
  
  Line$ = String(80, "-")
  Dot$ = String(75, ".")
  Line1$ = String(75, "_")
  FF$ = Chr(12)
  MaxLines = 58
  LineCnt = 0
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  MultiYear = TaxMasterRec.MultiYear
  
  Town$ = QPTrim$(TaxMasterRec.Name)
  TotOpt1 = 0
  TotOpt2 = 0
  TotOpt3 = 0
  OptRev1Desc = QPTrim$(TaxMasterRec.OptRev1)
  OptRev2Desc = QPTrim$(TaxMasterRec.OptRev2)
  OptRev3Desc = QPTrim$(TaxMasterRec.OptRev3)
  
  fpcmbCycle.Col = 1
  CycleName = QPTrim$(fpcmbCycle.ColText)
  If CycleName = "" Then CycleName = "N/A"
  If fpcmbCycle.Enabled = True And CycleName <> "NO CYCLE" Then
    fpcmbCycle.Col = 0
    CycleNum = CLng(fpcmbCycle.ColText)
  Else
    CycleNum = -1
  End If
  
  fpcmbCounty.Col = 1
  CountyName = QPTrim$(fpcmbCounty.ColText)
  If CountyName = "" Then CountyName = "N/A"
  If fpcmbCounty.Enabled = True And CountyName <> "ALL COUNTIES" Then
    fpcmbCounty.Col = 0
    CountyNum = CLng(fpcmbCounty.ColText)
  Else
    CountyNum = -1
  End If
  
  If InStr(fpcmbTownships.Text, "ALL") Then
    ThisTownship = "ALL"
  Else
    ThisTownship = QPTrim$(fpcmbTownships.Text)
    ThisTownship = Mid(ThisTownship, 1, Len(ThisTownship) - 4)
    ThisTownship = QPTrim$(ThisTownship)
  End If
  dlm$ = "~"
  If Exist(RealTaxBillFile) Then
    KillFile RealTaxBillFile
  End If
  If Exist(RealTaxBillOPFile) Then
    KillFile RealTaxBillOPFile
  End If
  If Exist("TMPBLADD.DAT") Then 'tax bill addon
    KillFile "TMPBLADD.DAT"
  End If
  
  AddOnCnt = 0
  
  OpenTaxBillAddOn AOHandle
  OpenRealTaxBillFile TBHandle, NumOfTBRecs
  OpenTaxPropFile RRHandle, NumOfRRREcs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenMortCodeFile MCHandle, NumOfMCRecs
  
  Inactive = 0
  
  frmVATaxShowPctComp.Label1 = "Creating Tax Pre-Billing Register"
  frmVATaxShowPctComp.Show , Me
  frmVATaxShowPctComp.cmdCancel.Visible = False
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  If UseOpt = "Y" Then
    NumOfTCRecs = OptCnt
  End If
  If UseSS = "Y" Then
    NumOfTCRecs = SSCnt
  End If

  For x = 1 To NumOfTCRecs
    TBillRec = NewTBillRec
    OPApplied = 0
    If UsingIdx = True Then
      Get TCHandle, CustArr(x), TaxCust
      ThisTaxRec = CustArr(x)
    Else
      Get TCHandle, x, TaxCust
      ThisTaxRec = x
    End If
    If TaxCust.FirstPropRec = 0 Then GoTo PreBillSkip
    If TaxCust.TaxExempt = "Y" Then
      GoTo PreBillSkip
    End If
    
    If ThisTownship <> "ALL" Then
      If QPTrim$(TaxCust.TownShip) <> ThisTownship Then
        GoTo PreBillSkip
      End If
    End If
    
    If CycleNum >= 0 Then
      If TaxCust.Cycle <> CycleNum Then
        GoTo PreBillSkip
      End If
    End If
    
    If CountyNum >= 0 Then
      If TaxCust.County4BillNum <> CountyNum Then
        GoTo PreBillSkip
      End If
    End If
    
'    Done2Disk = 0
    LateAmt# = 0

    If TaxCust.Deleted <> 0 Then
      GoTo PreBillSkip:
    End If
    If TaxCust.Active <> "Y" Then
      Inactive = Inactive + 1
      GoTo PreBillSkip:
    End If
    
    PastFlagSet = 0             'Initialize Past Balance Flag
    OverPayAmt = 0
    If UsingIdx = True Then
      OverPayAmt = GetCustRealBalance(CustArr(x), -1)
      If OverPayAmt < 0 Then
        OverPay = True
      Else
        OverPayAmt = 0
        OverPay = False
      End If
    Else
      OverPayAmt = GetCustRealBalance(x, -1)
      If OverPayAmt < 0 Then
        OverPay = True
      Else
        OverPayAmt = 0
        OverPay = False
      End If
    End If
    ThisTest = CStr(OverPayAmt)
    If InStr(ThisTest, "E") Then OverPayAmt = 0

    If TaxCust.FirstPropRec <= 0 Then
      NoProp = 1
      GoSub SetCustInfo
      GoSub WriteIt2Disk
      GoTo PreBillSkip
    End If
    
    NoProp = 0
    GoSub SetCustInfo
    GoSub GetRealInfo
    If TBillRec.TotalBillDue = 0 Then
      If PastFlagSet = True Then
      End If
    End If
    BillCnt = BillCnt + 1
PreBillSkip:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTCRecs
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
  
  Close RRHandle
  Close TCHandle
  Close MCHandle
  Close AOHandle
  
  If ThisTBCnt = 0 Then
    Call TaxMsg(900, "Using the parameters selected there are no customers who qualify for a tax charge.")
    Close
    fptxtRCurrYear.SetFocus
    Exit Sub
  End If
  
  TotalRealVal# = 0
  TotalEx# = 0
  TotalBldgVal# = 0
  NumBills& = 0
  
  TaxPreRptFile$ = "TAXRPTS\TaxRealPreBill.PRN"
  
  RptHandle = FreeFile
  Open TaxPreRptFile For Output As #RptHandle
  GoSub PreBillHeading
  Close TBHandle
  OpenRealTaxBillFile TBHandle, NumOfTBRecs
  
  For x = 1 To NumOfTBRecs
    Get TBHandle, x, TBillRec
'    If Discovery$ = "Y" And TBillRec.BillNumber = -1 Then
    If TBillRec.BillNumber = -1 Then 'took out 12/5/06
       GoTo NotThisOne
    Else
      If TBillRec.TotalBillDue = 0 And TBillRec.PriorYrBalance = 0 Then
        If fpcmbInclNoBills.Text = "N" Then
          GoTo NotThisOne
        End If
      End If
      If TBillRec.OptRevTax1 > 0 Or TBillRec.OptRevTax2 > 0 Or TBillRec.OptRevTax3 > 0 And TBillRec.OverPayAmt > 0 Then
        If LineCnt >= MaxLines - 10 Then
          Print #RptHandle, Chr$(12);
          GoSub PreBillHeading
        End If
      ElseIf TBillRec.OptRevTax1 > 0 Or TBillRec.OptRevTax2 > 0 Or TBillRec.OptRevTax3 > 0 Then
        If LineCnt >= MaxLines - 10 Then
          Print #RptHandle, Chr$(12);
          GoSub PreBillHeading
        End If
      Else
        If LineCnt >= MaxLines - 3 Then
          Print #RptHandle, Chr$(12);
          GoSub PreBillHeading
        End If
      End If
      Print #RptHandle, Using("#####", TBillRec.CustRec);
      Print #RptHandle, Tab(8); Left$(TBillRec.CustName, 30);
      Print #RptHandle, Tab(49); Using("$###,##0.00", TBillRec.PriorYrBalance);
      If TBillRec.BillNumber = -1 Then
        Print #RptHandle, Tab(62); "N/A";
      Else
        Print #RptHandle, Tab(60); Using("########", TBillRec.BillNumber);
      End If
'      Print #RptHandle, Tab(69); Using("$###,##0.00", TBillRec.TotalBillDue) 'end
'      Print #RptHandle, Tab(16); Using("$###,###,##0", TBillRec.RealValue);
'      Print #RptHandle, Tab(68); Using("$#,###,##0.00", TBillRec.RealTaxDue - TBillRec.LateTaxDue) 'added late tax on 8/14/06
      Print #RptHandle, Tab(68); Using("$#,###,##0.00", TBillRec.TotalBillDue - TBillRec.LateTaxDue) 'added late tax on 8/14/06 and changed from line above
      Print #RptHandle, Using("$###,###,##0", TBillRec.RealValue);
      Print #RptHandle, Using$("$###,###,##0", TBillRec.BldgValue);
      Print #RptHandle, Tab(30); Using("$###,###,##0.00", TBillRec.ExptValue);
      TValue# = OldRound#(TBillRec.RealValue + TBillRec.BldgValue - TBillRec.ExptValue)
      If TValue# < 0 Then TValue# = 0
      Print #RptHandle, Tab(45); Using("$###,###,##0.00", TValue#);
      If TBillRec.LateTaxDue > 0 Then
        Print #RptHandle, Tab(58); "Late = "; Using("###0.00", TBillRec.LateTaxDue)
      Else
        Print #RptHandle, ""
      End If
      LineCnt = LineCnt + 2
      If TBillRec.OverPayAmt > 0 Then
        Print #RptHandle, Tab(5); "Credit Balance Applied to Tax Due: "; Tab(70); Using("$###,##0.00", -TBillRec.OverPayAmt)
        Print #RptHandle, Tab(5); "Total Tax Due: "; Tab(70); Using$("$###,##0.00", (TBillRec.TotalBillDue - TBillRec.OverPayAmt))
        LineCnt = LineCnt + 2
      End If
      If TBillRec.OptRevTax1 > 0 Or TBillRec.OptRevTax2 > 0 Or TBillRec.OptRevTax3 > 0 Then
        Print #RptHandle, Tab(5); Dot
        Print #RptHandle, Tab(5); "Optional Revenue Tax For: " + QPTrim$(TBillRec.CustName)
        LineCnt = LineCnt + 2
        If TBillRec.OptRevTax1 > 0 Then
          Print #RptHandle, Tab(30); OptRev1Desc$; Tab(50); Using$("$##,##0.00", TBillRec.OptRevTax1)
          LineCnt = LineCnt + 1
        End If
        If TBillRec.OptRevTax2 > 0 Then
          Print #RptHandle, Tab(30); OptRev2Desc$; Tab(50); Using$("$##,##0.00", TBillRec.OptRevTax2)
          LineCnt = LineCnt + 1
        End If
        If TBillRec.OptRevTax3 > 0 Then
          Print #RptHandle, Tab(30); OptRev3Desc$; Tab(50); Using$("$##,##0.00", TBillRec.OptRevTax3)
          LineCnt = LineCnt + 1
        End If
        Print #RptHandle, Tab(5); Dot$
        LineCnt = LineCnt + 1
      End If
      Print #RptHandle, Line
      LineCnt = LineCnt + 1
      TotalRealVal# = OldRound#(TotalRealVal# + TBillRec.RealValue)
      TotalBldgVal# = OldRound#(TotalBldgVal# + TBillRec.BldgValue)
      TotalEx# = OldRound#(TotalEx# + TBillRec.ExptValue)
'      TotalBills# = OldRound#(TotalBills# + TBillRec.RealTaxDue)
      TotalBills# = OldRound#(TotalBills# + TBillRec.TotalBillDue) '8/14/06 changed from above line
      TotalLate# = OldRound#(TotalLate# + TBillRec.LateTaxDue)
      TotalOverPay# = OldRound(TotalOverPay# + TBillRec.OverPayAmt)
      If TBillRec.TotalBillDue > 0 Then
        NumBills& = NumBills& + 1
      End If
      
      If LineCnt >= MaxLines Then
        Print #RptHandle, Chr$(12);
        GoSub PreBillHeading
      End If
    End If      'Test for Discovery Bills
NotThisOne:
  Next x
  
  Page = Page + 1
  Print #RptHandle, Chr$(12);
  Print #RptHandle, Tab(15); "Real Property Tax Billing : Pre-Billing Register"
  Print #RptHandle,
  Print #RptHandle, "Date: "; CStr(Date); Tab(65); "Page #"; Page
  Print #RptHandle, Line
  Print #RptHandle, "Number of Bills to Process: "; Using("###,##0", NumBills&)
  Print #RptHandle, "      Total Real Valuation: "; Using("$###,###,##0.00", TotalRealVal#)
  Print #RptHandle, "  Total Building Valuation: "; Using("$###,###,##0.00", TotalBldgVal#)
  If TotOpt1 > 0 Or TotOpt2 > 0 Or TotOpt3 > 0 Then
    Print #RptHandle, Tab(7); Dot
    If TotOpt1 > 0 Then
      Print #RptHandle, Tab(10); "Total For " + QPTrim$(OptRev1Desc$) + ":"; Tab(45); Using$("$###,###,##0.00", TotOpt1)
    End If
    If TotOpt2 > 0 Then
      Print #RptHandle, Tab(10); "Total For " + QPTrim$(OptRev2Desc$) + ":"; Tab(45); Using$("$###,###,##0.00", TotOpt2)
    End If
    If TotOpt3 > 0 Then
      Print #RptHandle, Tab(10); "Total For " + QPTrim$(OptRev3Desc$) + ":"; Tab(45); Using$("$###,###,##0.00", TotOpt3)
    End If
    Print #RptHandle, Tab(7); Dot
  End If
  Print #RptHandle, "          Total Exemptions: "; Using("$###,###,##0.00", TotalEx#)
  Print #RptHandle, "     Net Valuation to Bill: "; Using("$###,###,##0.00", OldRound(TotalRealVal# + TotalBldgVal# - TotalEx#))
  Print #RptHandle, "         Total Late Amount: "; Using("$###,###,##0.00", TotalLate#)
  If TotalOverPay > 0 Then
    Print #RptHandle, "      Amount Before Credit: "; Using("$###,###,##0.00", (TotalBills# + TotalLate#))
    Print #RptHandle, "            Credit Applied: "; Using("$###,###,##0.00", TotalOverPay#)
    Print #RptHandle, "    Total Tax Bill To Bill: "; Using("$###,###,##0.00", (TotalBills# - TotalOverPay# + TotalLate#))
    Print #RptHandle, "    Total Past Amt to Bill: "; Using("$###,###,##0.00", TotalPast#)
    Print #RptHandle, "Grand Total Amount to Bill: "; Using("$###,###,##0.00", OldRound#(TotalPast# + TotalBills# + TotalLate# - TotalOverPay#))
  Else
    Print #RptHandle, "      Total Amount to Bill: "; Using("$###,###,##0.00", TotalBills#)
    Print #RptHandle, "    Total Past Amt to Bill: "; Using("$###,###,##0.00", TotalPast#)
    Print #RptHandle, "Grand Total Amount to Bill: "; Using("$###,###,##0.00", OldRound#(TotalPast# + TotalBills# + TotalLate#))
  End If
  Print #RptHandle, ""
  Print #RptHandle, "          Inactive Skipped: "; Using("###,##0", Inactive)
  Print #RptHandle, Chr$(12)
  Close
  
  ViewPrint TaxPreRptFile, "Tax Real Pre-Billing Report", True
  
  KillFile "txrblsprn.dat"
  MainLog ("Prebilling text report for real property generated.")
  
  Exit Sub
  
SetCustInfo:
  ThisPastDue = 0
  TBillRec.CustRec = ThisTaxRec 'x              'cust acct rec
  CustName$ = QPTrim$(TaxCust.CustName)
  TBillRec.CustName = CustName$
  TBillRec.CustAdd1 = QPTrim$(TaxCust.Addr1)
  TBillRec.CustAdd2 = QPTrim$(TaxCust.Addr2)
  CitySt$ = QPTrim$(TaxCust.City) + " " + TaxCust.State
  TBillRec.CustAdd3 = CitySt$
  TBillRec.CustZip = TaxCust.Zip
  TBillRec.CustPin = TaxCust.PIN
  TBillRec.TaxYear = WhatRYear
'  TBillRec.RDesc3 = TaxCust.CSSN 'changed 10/25/06

  'Set Prior Balance if any
  ABalance = 0
  ABalance = GetCustRealBalance(TaxCust.Acct, -1) 'added this line on 8/11/2006    0
'  GoSub GetPastBalance 'comment out 8/14/06
'  If ABalance# > 0 Then
    If PastFlagSet = 0 Then 'comment out 8/14/06
      TBillRec.PriorYrBalance = ABalance# 'comment out 8/14/06
    End If 'comment out 8/14/06
    PastFlagSet = 1 'comment out 8/14/06
'  End If
  Return
  
PreBillHeading:
  Page = Page + 1
  Print #RptHandle, Tab(15); "Real Property Tax Billing : Pre-Billing Register"
  Print #RptHandle, Town
  Print #RptHandle, "Date: "; CStr(Date); Tab(73); "Page #"; CStr(Page)
  ThisRate = CDbl(fpDblSnglRealRate.Value)
  Print #RptHandle, "TownShip: " + ThisTownship; Tab(56); "Real Tax Rate: " + Using$("##0.0000", ThisRate) + "%"
  fpcmbCounty.Col = 1
  ThisRate = CDbl(fpDblSnglPersRate.Value)
  Print #RptHandle, "County: " + CountyName; Tab(55); "Late List Rate: " + Using$("##0.0000", ThisRate) + "%"
  fpcmbCycle.Col = 1
  ThisRate = CDbl(fpDblSnglLateList.Value)
  Print #RptHandle, "Cycle: " + CycleName; Tab(55); "Multiyear Value: " + Using$("#0", MultiYear)
  Print #RptHandle, "* = Tax Due Does Not Include Prior Year Balance"
  Print #RptHandle,
  Print #RptHandle, "Acct #"; Tab(8); "Customer Name"; Tab(48); "Prior Yr Bal"; Tab(62); "Bill Seq#"; Tab(72); " *Tax Due"
  Print #RptHandle, Tab(3); "Real Value"; Tab(18); "Bldg Value"; Tab(33); "Discnt Value"; Tab(47); "Net Valuation"
  Print #RptHandle, Line
  LineCnt = 11
Return

GetPastBalance:
  Balance# = 0
  ABalance# = 0
  If TaxCust.LastTrans > 0 Then
    OpenTaxTransFile TTHandle, NumOfTTRecs
    TransRecord& = TaxCust.LastTrans
    Do While TransRecord& <> 0
      Get TTHandle, TransRecord&, TaxTrans
      If TaxTrans.TranType = 1 And TaxTrans.BillType = "R" Then
        Balance# = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
        Balance# = OldRound(Balance# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection)
        Balance# = OldRound(Balance# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3) 'added for vs 2.05
        Balance# = OldRound(Balance# + TaxTrans.Revenue.PrePaidAmt) 'added for vs 2.05
        Balance# = OldRound(Balance# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
        Balance# = OldRound(Balance# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd))
        Balance# = OldRound(Balance# - (TaxTrans.Revenue.LateListPd + TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.PPTRADisc))
        Balance# = OldRound(Balance# - TaxTrans.DiscAmt) 'added for vs 2.05
        Balance# = OldRound#(Balance#)
      End If
      ABalance# = ABalance# + Balance#
      Balance# = 0
      TransRecord& = TaxTrans.LastTrans
    Loop
    Close TTHandle
  End If

Return
  
WriteIt2Disk:   'write the info out to disk here.
  If ThisPastDue = 0 Then
    ThisPastDue = ABalance#
    TotalPast# = OldRound(TotalPast# + ThisPastDue#)
  End If
  TBillRec.SetDscvry2No = "N" 'added 12/4/06
  TBillRec.BillPrinted = False
  If TBillRec.TotalBillDue > 0 Then
    TBillRec.BillNumber = BillNum&
  Else
    TBillRec.BillNumber = -1
  End If
  If UsingMinTax = 1 Then
    If TBillRec.TotalBillDue <= MinBill And TBillRec.BillNumber <> -1 Then '8/17/06 added billnumber = -1
      TempAddOn.OldAmt = TBillRec.TotalBillDue
      TempAddOn.CustName = QPTrim$(TaxCust.CustName)
      TempAddOn.CustRec = x
      TempAddOn.Type = "Tax bills less than or equal to " + QPTrim$(Using$("$#,##0.00", MinBill)) + " become zero."
      AddOnCnt = AddOnCnt + 1
      TBillRec.TotalBillDue = 0
      TBillRec.SetDscvry2No = "Y" 'added 12/4/06
      TBillRec.RealTaxDue = 0 'added 12/4/06
      TBillRec.LateTaxDue = 0 'added 12/4/06
      TotOpt1# = OldRound(TotOpt1# - TBillRec.OptRevTax1) 'added 12/4/06
      TBillRec.OptRevTax1 = 0 'added 12/4/06
      TotOpt2# = OldRound(TotOpt2# - TBillRec.OptRevTax2) 'added 12/4/06
      TBillRec.OptRevTax2 = 0 'added 12/4/06
      TotOpt3# = OldRound(TotOpt3# - TBillRec.OptRevTax3) 'added 12/4/06
      TBillRec.OptRevTax3 = 0 'added 12/4/06
      TempAddOn.NewAmt = 0
      Put AOHandle, AddOnCnt, TempAddOn
    End If
  ElseIf UsingMinTax = 2 Then
    If TBillRec.TotalBillDue < MinBill And TBillRec.BillNumber <> -1 Then '8/17/06 added billnumber = -1
      TempAddOn.OldAmt = TBillRec.TotalBillDue
      TempAddOn.CustName = QPTrim$(TaxCust.CustName)
      TempAddOn.CustRec = x
      TempAddOn.Type = "Tax bills less than " + QPTrim$(Using$("$#,##0.00", MinBill)) + " become " + QPTrim$(Using$("$#,##0.00", MinBill)) + "."
      AddOnCnt = AddOnCnt + 1
      TempAddOn.NewAmt = MinBill
      Put AOHandle, AddOnCnt, TempAddOn
      Pct1# = 0
      Pct2# = 0
      Pct6# = 0
      Pct7# = 0
      Pct8# = 0
      ThisReal = OldRound(TBillRec.RealTaxDue - (TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3))
      PctTot# = OldRound(TBillRec.TotalBillDue)
      If ThisReal > 0 Then
        Pct1# = OldRound#(ThisReal / PctTot#)
      End If
      If TBillRec.OptRevTax1 > 0 Then
        Pct6# = OldRound#(TBillRec.OptRevTax1 / PctTot#)
      End If
      If TBillRec.OptRevTax2 > 0 Then '
        Pct7# = OldRound#(TBillRec.OptRevTax2 / PctTot#)
      End If
      If TBillRec.OptRevTax3 > 0 Then
        Pct8# = OldRound#(TBillRec.OptRevTax3 / PctTot#)
      End If
      If TBillRec.LateTaxDue > 0 Then
        Pct2# = OldRound#(TBillRec.LateTaxDue / PctTot#)
      End If
      
      ThisReal = MinBill * Pct1
      
      TotOpt1# = OldRound(TotOpt1# - TBillRec.OptRevTax1)
      OptRevTax1# = OldRound(OptRevTax1# - TBillRec.OptRevTax1)
      TBillRec.OptRevTax1 = OldRound(MinBill * Pct6)
      TotOpt1# = OldRound(TotOpt1# + TBillRec.OptRevTax1)
      OptRevTax1# = OldRound(OptRevTax1# + TBillRec.OptRevTax1)
      
      TotOpt2# = OldRound(TotOpt2# - TBillRec.OptRevTax2)
      OptRevTax2# = OldRound(OptRevTax2# - TBillRec.OptRevTax2)
      TBillRec.OptRevTax2 = OldRound(MinBill * Pct7)
      TotOpt2# = OldRound(TotOpt2# + TBillRec.OptRevTax2)
      OptRevTax2# = OldRound(OptRevTax2# + TBillRec.OptRevTax2)
      
      TotOpt3# = OldRound(TotOpt3# - TBillRec.OptRevTax3)
      OptRevTax3# = OldRound(OptRevTax3# - TBillRec.OptRevTax3)
      TBillRec.OptRevTax3 = OldRound(MinBill * Pct8)
      TotOpt3# = OldRound(TotOpt3# + TBillRec.OptRevTax3)
      OptRevTax3# = OldRound(OptRevTax3# + TBillRec.OptRevTax3)
      
      TotalLate# = OldRound(TotalLate# - TBillRec.LateTaxDue)
      LateAmt# = OldRound(LateAmt# - TBillRec.LateTaxDue)
      TBillRec.LateTaxDue = OldRound(MinBill * Pct2)
      TotalLate# = OldRound(TotalLate# + TBillRec.LateTaxDue)
      LateAmt# = OldRound(LateAmt# + TBillRec.LateTaxDue)
      
      PctTest# = OldRound(ThisReal + TBillRec.LateTaxDue)
      PctTest# = OldRound(PctTest# + TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3)
      
      If MinBill < PctTest# Then
        If ThisReal > PctTest# - MinBill Then
          ThisReal = OldRound(ThisReal - (PctTest# - MinBill))
          TotalReal# = OldRound(TotalReal# - (PctTest# - MinBill))
        ElseIf TBillRec.OptRevTax1 > PctTest# - MinBill Then
          TBillRec.OptRevTax1 = OldRound(TBillRec.OptRevTax1 - (PctTest# - MinBill))
          TotOpt1# = OldRound(TotOpt1# - (PctTest# - MinBill))
          OptRevTax1# = OldRound(OptRevTax1# - (PctTest# - MinBill))
        ElseIf TBillRec.OptRevTax2 > PctTest# - MinBill Then
          TBillRec.OptRevTax2 = OldRound(TBillRec.OptRevTax2 - (PctTest# - MinBill))
          TotOpt2# = OldRound(TotOpt2# - (PctTest# - MinBill))
          OptRevTax2# = OldRound(OptRevTax2# - (PctTest# - MinBill))
        ElseIf TBillRec.OptRevTax3 > PctTest# - MinBill Then
          TBillRec.OptRevTax3 = OldRound(TBillRec.OptRevTax3 - (PctTest# - MinBill))
          TotOpt3# = OldRound(TotOpt3# - (PctTest# - MinBill))
          OptRevTax3# = OldRound(OptRevTax3# - (PctTest# - MinBill))
        ElseIf TBillRec.LateTaxDue > PctTest# - MinBill Then
          TBillRec.LateTaxDue = OldRound(TBillRec.LateTaxDue - (PctTest# - MinBill))
          TotalLate# = OldRound(TotalLate# - (PctTest# - MinBill))
          LateAmt# = OldRound(LateAmt# - (PctTest# - MinBill))
        End If
      ElseIf MinBill > PctTest# Then
        If ThisReal > MinBill - PctTest# Then
          ThisReal = OldRound(ThisReal + (MinBill - PctTest#))
          TotalReal# = OldRound(TotalReal# + (MinBill - PctTest#))
        ElseIf TBillRec.OptRevTax1 > MinBill - PctTest# Then
          TBillRec.OptRevTax1 = OldRound(TBillRec.OptRevTax1 + (MinBill - PctTest#))
          TotOpt1# = OldRound(TotOpt1# + (PctTest# - MinBill))
          OptRevTax1# = OldRound(OptRevTax1# + (PctTest# - MinBill))
        ElseIf TBillRec.OptRevTax2 > MinBill - PctTest# Then
          TBillRec.OptRevTax2 = OldRound(TBillRec.OptRevTax2 + (MinBill - PctTest#))
          TotOpt2# = OldRound(TotOpt2# + (PctTest# - MinBill))
          OptRevTax2# = OldRound(OptRevTax2# + (PctTest# - MinBill))
        ElseIf TBillRec.OptRevTax3 > MinBill - PctTest# Then
          TBillRec.OptRevTax3 = OldRound(TBillRec.OptRevTax3 + (MinBill - PctTest#))
          TotOpt3# = OldRound(TotOpt3# + (PctTest# - MinBill))
          OptRevTax3# = OldRound(OptRevTax3# + (PctTest# - MinBill))
        ElseIf TBillRec.LateTaxDue > MinBill - PctTest# Then
          TBillRec.LateTaxDue = OldRound(TBillRec.LateTaxDue + (MinBill - PctTest#))
          TotalLate# = OldRound(TotalLate# + (MinBill - PctTest#))
          LateAmt# = OldRound(LateAmt# + (MinBill - PctTest#))
        End If
      End If
      TBillRec.RealTaxDue = OldRound(ThisReal + TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3 + TBillRec.LateTaxDue)
      TBillRec.TotalBillDue = MinBill
    End If
  End If
  
  ThisTBCnt = ThisTBCnt + 1

  Put #TBHandle, ThisTBCnt, TBillRec
  ThisTBCnt = LOF(TBHandle) / Len(TBillRec)
  OPApplied = 0 'added 10/17/08
  If TBillRec.TotalBillDue > 0 Then
    If Abs(OverPayAmt) > 0 Then
      OPBillRec.Revenue.LateListPd = 0
      OPBillRec.Revenue.Principle1Pd = 0
      OPBillRec.Revenue.RevOpt1Pd = 0
      OPBillRec.Revenue.RevOpt2Pd = 0
      OPBillRec.Revenue.RevOpt3Pd = 0
      OpenRealTaxBillOverPayFile OPHandle, NumOfOPRecs
      Get TBHandle, ThisTBCnt, TBillRec
      OPBillRec.Revenue.PrePaidAmt = Abs(OverPayAmt)
      OPBillRec.BelongTo = ThisTBCnt 'BillNum& changed on 8/21/06
      If TBillRec.LateTaxDue > 0 And Abs(OverPayAmt) > 0 Then
        If Abs(OverPayAmt) > TBillRec.LateTaxDue Then
          OverPayAmt = OldRound(OverPayAmt + TBillRec.LateTaxDue)
          OPBillRec.Revenue.LateListPd = TBillRec.LateTaxDue
          OPApplied = OldRound(OPApplied + OPBillRec.Revenue.LateListPd)
        ElseIf Abs(OverPayAmt) <= TBillRec.LateTaxDue Then
          OPBillRec.Revenue.LateListPd = -OverPayAmt
'          OPApplied = OldRound(OPApplied - OverPayAmt)'changed to line below on 10/17/08
          OPApplied = OPBillRec.Revenue.PrePaidAmt '-OverPayAmt 'added OPBillRec.Revenue.PrePaidAmt 8/23/09
          OverPayAmt = 0
        End If
      End If
      
      If OldRound(TBillRec.RealTaxDue - (TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3)) > 0 And Abs(OverPayAmt) > 0 Then
        If Abs(OverPayAmt) > OldRound(TBillRec.RealTaxDue - (TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3)) Then
          OverPayAmt = OldRound(OverPayAmt + TBillRec.RealTaxDue - (TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3))
          OPBillRec.Revenue.Principle1Pd = OldRound(OPBillRec.Revenue.Principle1Pd + TBillRec.RealTaxDue - (TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3))
          OPApplied = OldRound(OPApplied + TBillRec.RealTaxDue - (TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3))
        ElseIf Abs(OverPayAmt) <= OldRound(TBillRec.RealTaxDue - (TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3)) Then
          OPBillRec.Revenue.Principle1Pd = OPBillRec.Revenue.Principle1Pd - OverPayAmt
'          OPApplied = OldRound(OPApplied - OverPayAmt)'changed to line below on 10/17/08
          OPApplied = OPBillRec.Revenue.PrePaidAmt '-OverPayAmt 'added OPBillRec.Revenue.PrePaidAmt 8/23/09
          OverPayAmt = 0
        End If
      End If
      
      If TBillRec.OptRevTax1 > 0 And Abs(OverPayAmt) > 0 Then
        If Abs(OverPayAmt) > TBillRec.OptRevTax1 Then
          OverPayAmt = OldRound(OverPayAmt + TBillRec.OptRevTax1)
          OPBillRec.Revenue.RevOpt1Pd = TBillRec.OptRevTax1
          OPApplied = OldRound(OPApplied + TBillRec.OptRevTax1)
        ElseIf Abs(OverPayAmt) <= TBillRec.OptRevTax1 Then
          OPBillRec.Revenue.RevOpt1Pd = -OverPayAmt
'          OPApplied = OldRound(OPApplied - OverPayAmt)'changed to line below on 10/17/08
          OPApplied = OPBillRec.Revenue.PrePaidAmt '-OverPayAmt 'added OPBillRec.Revenue.PrePaidAmt 8/23/09
          OverPayAmt = 0
        End If
      End If
      
      If TBillRec.OptRevTax2 > 0 And Abs(OverPayAmt) > 0 Then
        If Abs(OverPayAmt) > TBillRec.OptRevTax2 Then
          OverPayAmt = OldRound(OverPayAmt + TBillRec.OptRevTax2)
          OPBillRec.Revenue.RevOpt2Pd = TBillRec.OptRevTax2
          OPApplied = OldRound(OPApplied + TBillRec.OptRevTax2)
        ElseIf Abs(OverPayAmt) <= TBillRec.OptRevTax2 Then
          OPBillRec.Revenue.RevOpt2Pd = -OverPayAmt
'          OPApplied = OldRound(OPApplied - OverPayAmt)'changed to line below on 10/17/08
          OPApplied = OPBillRec.Revenue.PrePaidAmt '-OverPayAmt 'added OPBillRec.Revenue.PrePaidAmt 8/23/09
          OverPayAmt = 0
        End If
      End If
      
      If TBillRec.OptRevTax3 > 0 And Abs(OverPayAmt) > 0 Then
        If Abs(OverPayAmt) > TBillRec.OptRevTax3 Then
          OverPayAmt = OldRound(OverPayAmt + TBillRec.OptRevTax3)
          OPBillRec.Revenue.RevOpt3Pd = TBillRec.OptRevTax3
          OPApplied = OldRound(OPApplied + TBillRec.OptRevTax3)
        ElseIf Abs(OverPayAmt) <= TBillRec.OptRevTax3 Then
          OPBillRec.Revenue.RevOpt3Pd = -OverPayAmt
'          OPApplied = OldRound(OPApplied - OverPayAmt)'changed to line below on 10/17/08
          OPApplied = OPBillRec.Revenue.PrePaidAmt '-OverPayAmt 'added OPBillRec.Revenue.PrePaidAmt 8/23/09
          OverPayAmt = 0
        End If
      End If
      OPBillRec.Amount = OPApplied
      Put OPHandle, NumOfOPRecs + 1, OPBillRec
      Close OPHandle
    End If '#1
    Get TBHandle, ThisTBCnt, TBillRec
    TBillRec.OverPayAmt = OPApplied
    Put TBHandle, ThisTBCnt, TBillRec
    'moved #1 end if to above from here on 10/17/08
    BillNum& = BillNum& + 1
  End If
'  Done2Disk = -1

  Return
  
GetRealInfo:
  ThisProp& = TaxCust.FirstPropRec
  TBillRec.Opt1Desc = QPTrim$(TaxMasterRec.OptRev1)
  TBillRec.Opt2Desc = QPTrim$(TaxMasterRec.OptRev2)
  TBillRec.Opt3Desc = QPTrim$(TaxMasterRec.OptRev3)
  
  If ThisProp& > 0 Then
'    BldgValue = 0
'    RealValue = 0
    Do While ThisProp > 0
      BldgValue = 0
      RealValue = 0
      Get RRHandle, ThisProp&, RealRec
      If RealRec.Deleted = True Then GoTo SupOnlyNGSkip 'added 5/6/05
      If SupOnly = "Y" And RealRec.PROPDISC <> "Y" Then GoTo SupOnlyNGSkip
      If RealRec.Mock = "Y" Then GoTo SupOnlyNGSkip
      CustMort$ = QPTrim$(RealRec.MORTCODE)
      If Len(CustMort$) > 0 And QPTrim$(CustMort) <> "NONE" Then
        For MortCnt = 1 To NumOfMCRecs
          ThisMort$ = QPTrim$(MortCodes(MortCnt).MORTCODE)
          If ThisMort$ = CustMort$ Then
            TBillRec.MortRec = MortCodes(MortCnt).MortRec
            TBillRec.MORTCODE = ThisMort 'added 8/22/05
            Exit For
          End If
        Next MortCnt
      Else
        TBillRec.MortRec = 0
        TBillRec.MORTCODE = ""
      End If
    
'      If RealRec.LastYrPrinted = WhatRYear Then Discovery$ = "Y" 'took out 12/5/06
      If MultiYear > 1 Then GoTo GoAhead
      If MultiYear = 1 And (RealRec.LastYrPrinted <> WhatRYear) Or (RealRec.PROPDISC = "Y") Or (RealRec.LastYrPrinted = WhatRYear) Then
      'above is dumb
GoAhead:
'        If MultiYear <> 0 Then 'remmed out MultiYear on 6/22/06
'          RealValue# = OldRound((RealRec.PROPVALU / MultiYear) + RealValue#)
'          BldgValue# = OldRound((RealRec.BldgVal / MultiYear) + BldgValue#)
'        Else
          RealValue# = RealRec.PROPVALU
          BldgValue# = RealRec.BldgVal
'        End If
        OptRevTax1# = FigureOptRevTax1(ThisProp&, RRHandle, "R")
'        OptRevTax1# = OldRound#(OptRevTax1#) 'added on 6/22/06
        TBillRec.OptRevTax1 = OptRevTax1# / MultiYear
        TBillRec.Opt1Desc = QPTrim$(TaxMasterRec.OptRev1)
        TotOpt1 = OldRound(TotOpt1 + TBillRec.OptRevTax1) 'OptRevTax1#)
        OptRevTax2# = FigureOptRevTax2(ThisProp&, RRHandle, "R")
        TBillRec.OptRevTax2 = OptRevTax2# / MultiYear
        TBillRec.Opt2Desc = QPTrim$(TaxMasterRec.OptRev2)
        TotOpt2 = OldRound(TotOpt2 + TBillRec.OptRevTax2) 'OptRevTax2#)
        OptRevTax3# = FigureOptRevTax3(ThisProp&, RRHandle, "R")
        TBillRec.OptRevTax3 = OptRevTax3# / MultiYear
        TBillRec.Opt3Desc = QPTrim$(TaxMasterRec.OptRev3)
        TotOpt3 = OldRound(TotOpt3 + TBillRec.OptRevTax3) 'OptRevTax3#)
'        If MultiYear <> 0 Then 'remmed out MultiYear on 6/22/06
'          RealExmp# = OldRound(RealRec.EXMPOTHR / MultiYear)
'        Else
          RealExmp# = RealRec.EXMPOTHR
'        End If
        If RealRec.EXMPOTHR > 0 Then
          TempAddOn.OldAmt = 0
          TempAddOn.CustName = QPTrim$(TaxCust.CustName)
          TempAddOn.CustRec = x
          TempAddOn.Type = "Other discount of " + QPTrim$(Using$("$#,##0.00", RealRec.EXMPOTHR)) + " applied to real estate tax."
          AddOnCnt = AddOnCnt + 1
          TempAddOn.NewAmt = RealRec.EXMPOTHR
          Put AOHandle, AddOnCnt, TempAddOn
        End If
      
        RealCalcVal# = OldRound#((BldgValue# + RealValue# - RealExmp#) / 100)
        If RealCalcVal# < 0 Then RealCalcVal# = 0
        RealTaxDue# = OldRound#((RealCalcVal# * RealRate#) + OptRevTax1# + OptRevTax2# + OptRevTax3#)
        RealTaxDue# = OldRound#(RealTaxDue# / MultiYear) 'added on 6/22/06
        If RealTaxDue# = 0 Then
          If OldRound(TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3) = 0 Then 'added 10/18/06
            If fpcmbInclNoBills.Text = "N" Then GoTo SupOnlyNGSkip
          End If
        End If
        TBillRec.ExptValue = OldRound#(TBillRec.ExptValue + RealExmp#)
        TBillRec.RealValue = RealValue#
        TBillRec.BldgValue = BldgValue#
        TBillRec.TotalValue = OldRound(RealValue# + BldgValue#)
        TBillRec.RealTaxDue = RealTaxDue#
        TBillRec.RealPropRecord = ThisProp&
        TBillRec.RealTaxRate = RealRate#
        TBillRec.RDesc1 = QPTrim$(RealRec.PROPNOT1) + " " + QPTrim$(RealRec.PROPNOT2)
        MAPBLKLOT$ = RealRec.Map + " " + RealRec.BLOCK + " " + RealRec.LOTNUMB

        TBillRec.RDesc2 = RealRec.PROPNOT2 ' MAPBLKLOT$ changed 10/25/06
        TBillRec.RDesc3 = RealRec.PROPNOT3 'changed 10/25/06
        TBillRec.CustPin = TaxCust.PIN
        TBillRec.InternalPin = RealRec.InternalPin
        TBillRec.RealPin = RealRec.RealPin
        If RealRec.LateList = "Y" Then
          LateAmt# = OldRound#((RealTaxDue#) * (LateList# / 100))
        Else
          LateAmt# = 0
        End If
'        TBillRec.LateTaxDue = OldRound#(TBillRec.LateTaxDue + LateAmt#)
        TBillRec.LateTaxDue = LateAmt# 'changed from line above on 10/13/06
'        TBillRec.TotalBillDue = OldRound#(TBillRec.TotalBillDue + RealTaxDue# + LateAmt#)
        TBillRec.TotalBillDue = OldRound#(RealTaxDue# + LateAmt#) 'changed from above line 10/13/06
        TBillRec.DueDate = Date2Num(fptxtDueDate)

        GoSub WriteIt2Disk
      End If      'End of Test For Current Year Tax Bill
SupOnlyNGSkip:
      ThisProp = RealRec.NextRec
    Loop
  End If

  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPrebilling", "PrintText", Erl)
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
  
Private Sub PrintPersGraphics()
  Dim TBillRec As VAPPTaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim OPBillRec As TaxTransactionType
  Dim OPHandle As Integer
  Dim NumOfOPRecs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  Dim NewTBillRec As VAPPTaxBillType
  Dim LateAmt#
  Dim Inactive As Integer
  Dim PastFlagSet As Boolean
  Dim CustName$, CitySt$, CustAcct&
  Dim ABalance#
  Dim Balance#
  Dim TransRecord&
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TempAddOn As TempTaxBillAddOn
  Dim AOHandle As Integer
  Dim AddOnCnt As Integer
  Dim PersValue#
  Dim Discovery$
  Dim PersTaxDue#
  Dim NextPersRec&
  Dim TotalPers#
  Dim TotalOverPay#
  Dim TotalEx#
  Dim NumBills&
  Dim RptHandle As Integer
  Dim TaxPreRptFile As String
  Dim TValue#
  Dim dlm$
  Dim TotalBills#
  Dim TotalLate#
  Dim TotalPast#
  Dim ThisTownship$
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim ThisTaxRec As Long
  Dim CountyNum As Long
  Dim CountyName$
  Dim CycleNum As Long
  Dim CycleName$
  Dim BillCnt As Long
  Dim OverPay As Boolean
  Dim OverPayAmt As Double
  Dim OPApplied As Double
  Dim ThisTBCnt As Long
  Dim ThisTest$
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim OptFlag As Boolean
  Dim PERSRATE#
  Dim MHRate#
  Dim MCRate#
  Dim FERate#
  Dim MTRate#
  Dim PersTax#
  Dim MHTax#
  Dim MCTax#
  Dim FETax#
  Dim MTTax#
  Dim GTPersTax#
  Dim GTPersTaxNet#
  Dim GTMHTax#
  Dim GTMCTax#
  Dim GTFETax#
  Dim GTMTTax#
  Dim MHValue#
  Dim MCValue#
  Dim FEValue#
  Dim MTValue#
  Dim GTPersValue#
  Dim GTMHValue#
  Dim GTMCValue#
  Dim GTFEValue#
  Dim GTMTValue#
  Dim PropertyRec!
  Dim PPTRAVal#
  Dim GTPPTRADisc#
  Dim GTPPTRAVal#
  Dim PPTRADisc#
  Dim PERC!
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPRecs As Long
  Dim Factor!
  Dim Rate As Double
  Dim PYear$
  Dim PYearInt As Integer
  Dim TxFile As Integer
  Dim Prorate$
  Dim GTPTTRADisc#
  Dim GTPTTRAVal#
  Dim MultiYear As Integer
  Dim ThisPPTRADisc As Double
  Dim ThisCName As String * 36
  Dim TPersTaxDue As Double
  Dim TPPTRAVal As Double
  Dim TPPTRADisc As Double
  Dim ThisMaxVehVal As Double, ThisMinVehVal#
  Dim OptRevTax1 As Double
  Dim OptRevTax2 As Double
  Dim OptRevTax3 As Double
  Dim OptRev1Desc$
  Dim OptRev2Desc$
  Dim OptRev3Desc$
  Dim TotOpt1 As Double
  Dim TotOpt2 As Double
  Dim TotOpt3 As Double
  Dim TotReal As Double
  Dim ThisCnt As Integer
  Dim AHandle As Integer
  Dim ThisProp&
  Dim LunenBurgYN As Boolean
  Dim PersBalTest#
  Dim CalcDiff#
  Dim PPTRAValue#
  Dim ChilhowieGFudge As Single
  Dim Pct1#, Pct2#, Pct3#, Pct4#, Pct5#, Pct6#
  Dim Pct7#, Pct8#, Pct9#, PctTot#, PctTest#
  
  On Error GoTo ERRORSTUFF
  
'  AHandle = FreeFile
'  Open "prebillpers.dat" For Output As AHandle
  
  PERSRATE# = fpDblSnglPersRate
  MHRate# = fpDblSnglMHRate
  MCRate# = fpDblSnglMCRate
  FERate# = fpDblSnglFERate
  MTRate# = fpDblSnglMTRate
  
  Prorate$ = "Y"
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  MultiYear = TaxMasterRec.MultiYear
  TotOpt1 = 0
  TotOpt2 = 0
  TotOpt3 = 0
  OptRev1Desc = QPTrim$(TaxMasterRec.POptRev1)
  OptRev2Desc = QPTrim$(TaxMasterRec.POptRev2)
  OptRev3Desc = QPTrim$(TaxMasterRec.POptRev3)
  
  ThisPPTRADisc = TaxMasterRec.PPTRADisc
  
  fpcmbCycle.Col = 1
  CycleName = QPTrim$(fpcmbCycle.ColText)
  If CycleName = "" Then CycleName = "N/A"
  If fpcmbCycle.Enabled = True And CycleName <> "NO CYCLE" Then
    fpcmbCycle.Col = 0
    CycleNum = CLng(fpcmbCycle.ColText)
  Else
    CycleNum = -1
  End If
  
  fpcmbCounty.Col = 1
  CountyName = QPTrim$(fpcmbCounty.ColText)
  If CountyName = "" Then CountyName = "N/A"
  If fpcmbCounty.Enabled = True And CountyName <> "ALL COUNTIES" Then
    fpcmbCounty.Col = 0
    CountyNum = CLng(fpcmbCounty.ColText)
  Else
    CountyNum = -1
  End If
  
  If QPTrim$(fpcmbTownships.Text) = "ALL TOWNSHIPS" Then
    ThisTownship = "ALL"
  Else
    ThisTownship = QPTrim$(fpcmbTownships.Text)
    ThisTownship = Mid(ThisTownship, 1, Len(ThisTownship) - 4)
    ThisTownship = QPTrim$(ThisTownship)
  End If
  dlm$ = "~"
  If Exist(PersTaxBillFile) Then
    KillFile PersTaxBillFile
  End If
  If Exist(PersTaxBillOPFile) Then
    KillFile PersTaxBillOPFile
  End If
  If Exist("TMPPERSBLADD.DAT") Then 'tax bill addon
    KillFile "TMPPERSBLADD.DAT"
  End If
  
  AddOnCnt = 0
  OpenTaxBillPersAddOn AOHandle
  OpenPersPropFile PHandle, NumOfPRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenPersTaxBillFile TBHandle, NumOfTBRecs
  
  Inactive = 0
  
  frmVATaxShowPctComp.Label1 = "Creating Tax Personal Pre-Billing Register"
  frmVATaxShowPctComp.Show , Me
  frmVATaxShowPctComp.cmdCancel.Visible = False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  EnableCloseButton Me.hwnd, False
  If UseOpt = "Y" Then
    NumOfTCRecs = OptCnt
  End If
  If UseSS = "Y" Then
    NumOfTCRecs = SSCnt
  End If
  
  GoSub PPTRA
  
  For x = 1 To NumOfTCRecs
    TBillRec = NewTBillRec
    If UsingIdx = True Then
      Get TCHandle, CustArr(x), TaxCust
      ThisTaxRec = CustArr(x)
    Else
      Get TCHandle, x, TaxCust
      ThisTaxRec = x
    End If
'    If ThisTaxRec = 85 Then Stop
    If TaxCust.FirstPersRec = 0 Then GoTo PreBillSkip
    If TaxCust.TaxExempt = "Y" Then
      GoTo PreBillSkip
    End If
    If ThisTownship <> "ALL" Then
      If QPTrim$(TaxCust.TownShip) <> ThisTownship Then
        GoTo PreBillSkip
      End If
    End If
    If CycleNum >= 0 Then
      If TaxCust.Cycle <> CycleNum Then
        GoTo PreBillSkip
      End If
    End If
    
    If CountyNum >= 0 Then
      If TaxCust.County4BillNum <> CountyNum Then
        GoTo PreBillSkip
      End If
    End If
    
    LateAmt# = 0
    OPApplied = 0
    If TaxCust.Deleted <> 0 Then
      GoTo PreBillSkip:
    End If
    
    If QPTrim$(TaxCust.Active) <> "Y" Then
      Inactive = Inactive + 1
      GoTo PreBillSkip:
    End If
    
    PastFlagSet = 0             'Initialize Past Balance Flag
    OverPayAmt = 0
    If UsingIdx = True Then
      OverPayAmt = GetCustPersBalance(CustArr(x), -1)
      If OverPayAmt < 0 Then
        OverPay = True
      Else
        OverPayAmt = 0
        OverPay = False
      End If
    Else
      OverPayAmt = GetCustPersBalance(x, -1)
      If OverPayAmt < 0 Then
        OverPay = True
      Else
        OverPayAmt = 0
        OverPay = False
      End If
    End If
    ThisTest = CStr(OverPayAmt)
    If InStr(ThisTest, "E") Then OverPayAmt = 0
    If TaxCust.FirstPersRec <= 0 Then
      GoSub SetCustInfo
      GoSub WriteIt2Disk
      GoTo PreBillSkip
    End If
    
    GoSub SetCustInfo
    GoSub GetPersInfo
    
PreBillSkip:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTCRecs
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
  
'  Close TCHandle
  Close AOHandle
  Close PHandle
  
  If ThisTBCnt = 0 Then
    Call TaxMsg(900, "Using the parameters selected there are no customers who qualify for a tax charge.")
    Close
    fptxtRCurrYear.SetFocus
    Exit Sub
  End If
  
  TotalPers# = 0
  TotalEx# = 0
  NumBills& = 0
  
  TaxPreRptFile$ = "TAXRPTS\TaxPersPreBill.RPT"
  
  RptHandle = FreeFile
  Open TaxPreRptFile For Output As #RptHandle
  
  Close TBHandle
  OpenPersTaxBillFile TBHandle, NumOfTBRecs
  
  For x = 1 To NumOfTBRecs
    Get TBHandle, x, TBillRec
    Get TCHandle, TBillRec.CustRec, TaxCust
'    If Discovery$ = "Y" And TBillRec.BillNumber = -1 Then
    If TBillRec.BillNumber = -1 Then '12/5/06 took out the Discovery$ = "Y" part because
    'it pertains to only allowing one billing per year where in Windows they can bill multiple times
       GoTo NotThisOne
    Else
      If TBillRec.TotalBillDue = 0 And TBillRec.PriorYrBalance = 0 And TBillRec.PPTRAValue = 0 Then
        If fpcmbInclNoBills.Text = "N" Then
          GoTo NotThisOne
        End If
      End If
'      Print #AHandle, CStr(TBillRec.CustRec) + "~" + Using$("$###,###0.00", TBillRec.TotalBillDue)
      '                        0
      Print #RptHandle, TBillRec.CustRec; dlm;
      If QPTrim$(TBillRec.Prorate) <> "Y" Then
        '                            1
        Print #RptHandle, QPTrim$(TBillRec.CustName); dlm;
      Else
        ThisCName = QPTrim$(TBillRec.CustName)
        '
        Print #RptHandle, ThisCName + "**Prorated**"; dlm;
      End If
        
'      If TBillRec.PriorYrBalance > 0 Then '8/14/06
        '                            2
        Print #RptHandle, TBillRec.PriorYrBalance; dlm;
'      Else '8/14/06
'        '                 2
'        Print #RptHandle, 0; dlm; '8/14/06
'      End If '8/14/06
      If TBillRec.BillNumber = -1 Then
        '                   3
        Print #RptHandle, "N/A"; dlm;
      Else
        '                          3
        Print #RptHandle, TBillRec.BillNumber; dlm;
      End If
      '                           4
      Print #RptHandle, OldRound(TBillRec.TotalBillDue - TBillRec.LateTaxDue); dlm; '8/14/06 updated to add late tax in this field
      TValue# = OldRound#(TBillRec.PersValue + TBillRec.FEValue + TBillRec.MCValue + TBillRec.MHValue + TBillRec.MTValue - TBillRec.ExptValue)
      
      If TValue# < 0 Then TValue# = 0
      '                    5
      Print #RptHandle, TValue#; dlm;
      '                           6
      Print #RptHandle, TBillRec.LateTaxDue; dlm;
      TotalLate# = OldRound#(TotalLate# + TBillRec.LateTaxDue)
'      TotalPast# = OldRound#(TotalPast# + TBillRec.PriorYrBalance)
      If TBillRec.PriorYrBalance > 0 Then 'added 8/14/06
        TotalPast# = OldRound#(TotalPast# + TBillRec.PriorYrBalance)
      Else 'added 8/14/06
'        TotalPast# = 0 'added 8/14/06
      End If 'added 8/14/06
      If InStr(TaxMasterRec.Name, "CHILHOWIE") Then
        If TBillRec.TotalBillDue = 0 Then 'And TBillRec.ChillHowieFudge > 0 Then
          TotalLate# = TotalLate# - TBillRec.LateTaxDue
          GTFETax = GTFETax - TBillRec.FETaxDue
          GTMHTax = GTMHTax - TBillRec.MHTaxDue
          GTMCTax = GTMCTax - TBillRec.MCTaxDue
          GTMTTax = GTMTTax - TBillRec.MTTaxDue
          TotOpt2 = TotOpt2 - TBillRec.OptRevTax2
          TotOpt3 = TotOpt3 - TBillRec.OptRevTax3
          GTPersTax = GTPersTax - TBillRec.PersTaxDue
          GTPPTRADisc# = GTPPTRADisc - TBillRec.PersTaxDue ' - TBillRec.PPTRADiscnt)
        ElseIf TBillRec.TotalBillDue > 0 Then
          If TBillRec.PersTaxDue > 0 Then  '11/9/06
            GTPersTax = GTPersTax + TBillRec.OptRevTax1
            GTPersTaxNet = GTPersTaxNet + TBillRec.OptRevTax1
          ElseIf TBillRec.MHTaxDue > 0 Then
            GTMHTax = GTMHTax + TBillRec.OptRevTax1
          ElseIf TBillRec.MCTaxDue > 0 Then
            GTMCTax = GTMCTax + TBillRec.OptRevTax1
          ElseIf TBillRec.MTTaxDue > 0 Then
            GTMTTax = GTMTTax + TBillRec.OptRevTax1
          ElseIf TBillRec.FETaxDue > 0 Then
            GTFETax = GTFETax + TBillRec.OptRevTax1
          ElseIf TBillRec.OptRevTax2 > 0 Then
            TotOpt2 = TotOpt2 + TBillRec.OptRevTax1
          ElseIf TBillRec.OptRevTax3 > 0 Then
            TotOpt3 = TotOpt3 + TBillRec.OptRevTax1
          ElseIf TBillRec.LateTaxDue > 0 Then
            TotalLate# = TotalLate# + TBillRec.OptRevTax1
          End If
        End If
        TotalBills = TotalBills + TBillRec.TotalBillDue
        GoTo ChilhowieSkip4
      End If
      
      TotalBills# = OldRound(TotalBills# + TBillRec.PersTaxNet + TBillRec.FETaxDue + TBillRec.MCTaxDue + TBillRec.MHTaxDue + TBillRec.MTTaxDue + TBillRec.LateTaxDue)
      TotalBills# = OldRound(TotalBills# + TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3)
ChilhowieSkip4:
      
      TotalOverPay# = OldRound#(TotalOverPay# + TBillRec.OverPayAmt)

      If TBillRec.TotalBillDue > 0 Then
        NumBills& = NumBills& + 1
      End If
      '                     7
      Print #RptHandle, NumBills&; dlm;
      '                     8
      Print #RptHandle, TotalPers#; dlm;
      '                     9
      Print #RptHandle, TotalBills#; dlm;
      '                     10
      Print #RptHandle, TotalPast#; dlm;
      '                                   11
      Print #RptHandle, OldRound#(TotalPast# + TotalBills# - TotalOverPay); dlm; 'added -TotalOverPay on 8/15/06
      '                     12
      Print #RptHandle, TotalLate#; dlm;
      '                    13
      Print #RptHandle, Inactive; dlm;
      '                     14
      Print #RptHandle, ThisTown; dlm;
      '                    15
      Print #RptHandle, WhatPYear; dlm;
      '                     16                17              18
      Print #RptHandle, ThisTownship; dlm; CycleName; dlm; CycleNum; dlm;
      '                     19                20
      Print #RptHandle, CountyName$; dlm; CountyNum; dlm;
      '                          21                                  22
      Print #RptHandle, -TBillRec.OverPayAmt; dlm; OldRound(TotalBills# - TotalOverPay); dlm;
      '                      23                              24                                   25
      Print #RptHandle, TotalOverPay#; dlm; CDbl(fpDblSnglPersRate.Value) / 100; dlm; CDbl(fpDblSnglLateList.Value) / 100; dlm;
      '                                    26                              27                                      28                                           29
      Print #RptHandle, CDbl(fpDblSnglFERate.Value) / 100; dlm; CDbl(fpDblSnglMHRate.Value) / 100; dlm; CDbl(fpDblSnglMCRate.Value) / 100; dlm; CDbl(fpDblSnglMTRate.Value) / 100; dlm;
      '                       30                         31                      32                      33                     34                       35
      Print #RptHandle, TBillRec.PersTaxDue; dlm; TBillRec.MHTaxDue; dlm; TBillRec.MCTaxDue; dlm; TBillRec.MTTaxDue; dlm; TBillRec.FETaxDue; dlm; TBillRec.PersValue; dlm;
      '                        36                     37                    38
      Print #RptHandle, TBillRec.MHValue; dlm; TBillRec.MTValue; dlm; TBillRec.MCValue; dlm; TBillRec.FEValue; dlm;
      '                    40               41             42             43             44               45                46               47               48              49
      Print #RptHandle, GTPersTax#; dlm; GTMHTax#; dlm; GTMTTax#; dlm; GTMCTax#; dlm; GTFETax#; dlm; GTPersValue#; dlm; GTMHValue#; dlm; GTMTValue#; dlm; GTMCValue#; dlm; GTFEValue#; dlm;
      '                         50                         51                     52                53                              54
      Print #RptHandle, TBillRec.PPTRADiscnt; dlm; TBillRec.PPTRAValue; dlm; GTPPTRADisc#; dlm; GTPPTRAVal#; dlm; OldRound(TBillRec.PersTaxDue); dlm;
      '                                    55                           56              57              58               59                              60
      Print #RptHandle, OldRound(GTPersValue# - GTPPTRAVal#); dlm; GTPersTaxNet#; dlm; PERC!; dlm; MultiYear; dlm; OldRound(ThisPPTRADisc / 100); dlm; PPTRAYN; dlm;
      '                     61                            62                        63
      Print #RptHandle, TBillRec.OptRevTax1; dlm; TBillRec.OptRevTax2; dlm; TBillRec.OptRevTax3; dlm;
       '                      64                65                 66
      Print #RptHandle, OptRev1Desc$; dlm; OptRev2Desc$; dlm; OptRev3Desc$; dlm;
      '                   67            68            69               70
      Print #RptHandle, TotOpt1; dlm; TotOpt2; dlm; TotOpt3; dlm; ChilhowieGFudge
      
    End If      'Test for Discovery Bills
NotThisOne:
  Next x
  
  Close
  arVATaxPreBillPers.Show
  frmVATaxLoadReport.Show
   
  KillFile "txpblsprn.dat"
  MainLog ("Prebilling graphics report for personal property generated.")
  Exit Sub
  
PPTRA:
  If WhatPYear$ = "1998" Then
    PERC! = 12.5
  ElseIf WhatPYear$ = "1999" Then
    PERC! = 27.5
  ElseIf WhatPYear$ = "2000" Then
    PERC! = 47.5
  ElseIf WhatPYear$ = "2001" Or WhatPYear$ = "2002" Then
    PERC! = 70
  Else
    PERC! = TaxMasterRec.PPTRADisc
  End If

Return

SetCustInfo:
  TBillRec.CustRec = ThisTaxRec
  CustName$ = QPTrim$(TaxCust.CustName)
  TBillRec.CustName = CustName$
  TBillRec.CustAdd1 = QPTrim$(TaxCust.Addr1)
  TBillRec.CustAdd2 = QPTrim$(TaxCust.Addr2)
  CitySt$ = QPTrim$(TaxCust.City) + " " + TaxCust.State
  TBillRec.CustAdd3 = CitySt$
  TBillRec.CustZip = TaxCust.Zip
  TBillRec.CustPin = TaxCust.PIN
  TBillRec.TaxYear = WhatPYear
  TBillRec.RDesc3 = TaxCust.CSSN

  'Set Prior Balance if any
  ABalance = 0
  ABalance = GetCustPersBalance(TaxCust.Acct, -1) 'added this line on 8/11/2006    0
'  GoSub GetPastBalance
  If ABalance# > 0 Then
    If PastFlagSet = 0 Then
      TBillRec.PriorYrBalance = ABalance#
    End If
    PastFlagSet = 1
  End If
  Return
  
GetPastBalance:
  
  Balance# = 0
  ABalance# = 0
  
  If TaxCust.LastTrans > 0 Then
    OpenTaxTransFile TTHandle, NumOfTTRecs
    TransRecord& = TaxCust.LastTrans
    Do While TransRecord& <> 0
      Get TTHandle, TransRecord&, TaxTrans
      If TaxTrans.TranType = 1 And TaxTrans.BillType = "P" Then
        Balance# = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
        Balance# = OldRound(Balance# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection)
        Balance# = OldRound(Balance# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3) 'added for vs 2.05
        Balance# = OldRound(Balance# + TaxTrans.Revenue.PrePaidAmt) 'added for vs 2.05
        Balance# = OldRound(Balance# - (TaxTrans.Revenue.LateListPd + TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
        Balance# = OldRound(Balance# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd))
        Balance# = OldRound(Balance# - TaxTrans.DiscAmt) 'added for vs 2.05
        Balance# = OldRound(Balance# - (TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
        Balance# = OldRound#(Balance#)
      End If
      ABalance# = ABalance# + Balance#
      Balance# = 0
      TransRecord& = TaxTrans.LastTrans
    Loop
    Close TTHandle
  End If

Return
  
WriteIt2Disk:   'write the info out to disk here.
  TBillRec.BillPrinted = False
  TBillRec.SetDscvry2No = "N" 'added 12/4/06
  If TBillRec.TotalBillDue > 0 Then
    TBillRec.BillNumber = BillNum&
  Else
    TBillRec.BillNumber = -1
  End If
  If UsingMinTax = 1 Then
    If TBillRec.TotalBillDue <= MinBill And TBillRec.BillNumber <> -1 Then '8/17/06 added billnumber = -1
      TempAddOn.OldAmt = TBillRec.TotalBillDue
      TempAddOn.CustName = QPTrim$(TaxCust.CustName)
      TempAddOn.CustRec = x
      TempAddOn.Type = "Tax bills less than or equal to " + QPTrim$(Using$("$#,##0.00", MinBill)) + " become zero."
      AddOnCnt = AddOnCnt + 1
      TBillRec.TotalBillDue = 0
      TBillRec.SetDscvry2No = "Y" 'added 12/4/06
      GTPersTax# = OldRound(GTPersTax# - TBillRec.PersTaxDue)
      TBillRec.PersTaxDue = 0 'added 12/4/06
      GTPersTaxNet# = OldRound(GTPersTaxNet# - TBillRec.PersTaxNet) 'added 12/4/06
      TBillRec.PersTaxNet = 0 'added 12/4/06
      GTMCTax# = OldRound(GTMCTax# - TBillRec.MCTaxDue#) 'added 12/4/06
      TBillRec.MCTaxDue = 0 'added 12/4/06
      GTMHTax# = OldRound(GTMHTax# - TBillRec.MHTaxDue#) 'added 12/4/06
      TBillRec.MHTaxDue = 0 'added 12/4/06
      GTMTTax# = OldRound(GTMTTax# - TBillRec.MTTaxDue#) 'added 12/4/06
      TBillRec.MTTaxDue = 0 'added 12/4/06
      GTFETax# = OldRound(GTFETax# - TBillRec.FETaxDue#) 'added 12/4/06
      TBillRec.FETaxDue = 0 'added 12/4/06
      TotOpt1# = OldRound(TotOpt1# - TBillRec.OptRevTax1) 'added 12/4/06
      TBillRec.OptRevTax1 = 0 'added 12/4/06
      TotOpt2# = OldRound(TotOpt2# - TBillRec.OptRevTax2) 'added 12/4/06
      TBillRec.OptRevTax2 = 0 'added 12/4/06
      TotOpt3# = OldRound(TotOpt3# - TBillRec.OptRevTax3) 'added 12/4/06
      TBillRec.OptRevTax3 = 0 'added 12/4/06
      TotalLate# = OldRound(TotalLate# - TBillRec.LateTaxDue) 'added 12/4/06
      TBillRec.LateTaxDue = 0 'added 12/4/06
      GTPPTRADisc# = OldRound(GTPPTRADisc# - TBillRec.PPTRADiscnt) 'added 12/4/06
      TBillRec.PPTRADiscnt = 0 'added 12/4/06
      TempAddOn.NewAmt = 0
      Put AOHandle, AddOnCnt, TempAddOn
    End If
  ElseIf UsingMinTax = 2 And InStr(TaxMasterRec.Name, "CHILHOWIE") = 0 Then 'added Chilhowie on 11/3/06
    If TBillRec.TotalBillDue < MinBill And TBillRec.BillNumber <> -1 Then '8/17/06 added billnumber = -1
      TempAddOn.OldAmt = TBillRec.TotalBillDue
      TempAddOn.CustName = QPTrim$(TaxCust.CustName)
      TempAddOn.CustRec = x
      TempAddOn.Type = "Tax bills less than " + QPTrim$(Using$("$#,##0.00", MinBill)) + " become " + QPTrim$(Using$("$#,##0.00", MinBill)) + "."
      AddOnCnt = AddOnCnt + 1
      TempAddOn.NewAmt = MinBill
      Put AOHandle, AddOnCnt, TempAddOn
      Pct1# = 0
      Pct2# = 0
      Pct3# = 0
      Pct4# = 0
      Pct5# = 0
      Pct6# = 0
      Pct7# = 0
      Pct8# = 0
      Pct9# = 0
      PctTot# = OldRound(TBillRec.PersTaxNet + TBillRec.FETaxDue + TBillRec.MCTaxDue + TBillRec.MHTaxDue + TBillRec.MTTaxDue)
      PctTot# = OldRound(PctTot# + TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3 + TBillRec.LateTaxDue)
      If TBillRec.PersTaxNet > 0 Then
        Pct1# = OldRound#(TBillRec.PersTaxNet / PctTot#)
      End If
      If TBillRec.FETaxDue > 0 Then
        Pct2# = OldRound#(TBillRec.FETaxDue / PctTot#)
      End If
      If TBillRec.MCTaxDue > 0 Then
        Pct3# = OldRound#(TBillRec.MCTaxDue / PctTot#)
      End If
      If TBillRec.MHTaxDue > 0 Then
        Pct4# = OldRound#(TBillRec.MHTaxDue / PctTot#)
      End If
      If TBillRec.MTTaxDue > 0 Then
        Pct5# = OldRound#(TBillRec.MTTaxDue / PctTot#)
      End If
      If TBillRec.OptRevTax1 > 0 Then
        Pct6# = OldRound#(TBillRec.OptRevTax1 / PctTot#)
      End If
      If TBillRec.OptRevTax2 > 0 Then '
        Pct7# = OldRound#(TBillRec.OptRevTax2 / PctTot#)
      End If
      If TBillRec.OptRevTax3 > 0 Then
        Pct8# = OldRound#(TBillRec.OptRevTax3 / PctTot#)
      End If
      If TBillRec.LateTaxDue > 0 Then
        Pct9# = OldRound#(TBillRec.LateTaxDue / PctTot#)
      End If
      
      GTPersTaxNet = OldRound(GTPersTaxNet - TBillRec.PersTaxNet)
      TBillRec.PersTaxNet = MinBill * Pct1
      GTPersTaxNet = OldRound(GTPersTaxNet + TBillRec.PersTaxNet)
      
      'new code could go here
      '----------------------------------------------------
      If PPTRADisc# = 0 Then
        GTPersTax# = OldRound(GTPersTax# - TBillRec.PersTaxDue)
        TBillRec.PersTaxDue = TBillRec.PersTaxNet
        GTPersTax# = OldRound(GTPersTax# + TBillRec.PersTaxDue)
      Else
        GTPPTRADisc# = OldRound(GTPPTRADisc# - TBillRec.PPTRADiscnt)
        GTPersTax# = OldRound(GTPersTax# - TBillRec.PersTaxDue)
        TBillRec.PersTaxDue = TBillRec.PersTaxNet + PPTRADisc#
        TBillRec.PPTRADiscnt = PPTRADisc# 'TBillRec.PersTaxDue - TBillRec.PersTaxNet
        GTPPTRADisc# = OldRound(GTPPTRADisc# + TBillRec.PPTRADiscnt)
        GTPersTax# = OldRound(GTPersTax# + TBillRec.PersTaxDue)
      End If
      '----------------------------------------------------
      
      GTFETax# = OldRound(GTFETax# - TBillRec.FETaxDue)
      FETax# = OldRound(FETax# - TBillRec.FETaxDue)
      TBillRec.FETaxDue = OldRound(MinBill * Pct2)
      GTFETax# = OldRound(GTFETax# + TBillRec.FETaxDue)
      FETax# = OldRound(FETax# + TBillRec.FETaxDue)
      
      GTMCTax# = OldRound(GTMCTax# - TBillRec.MCTaxDue)
      MCTax# = OldRound(MCTax# - TBillRec.MCTaxDue)
      TBillRec.MCTaxDue = OldRound(MinBill * Pct3)
      GTMCTax# = OldRound(GTMCTax# + TBillRec.MCTaxDue)
      MCTax# = OldRound(MCTax# + TBillRec.MCTaxDue)
      
      GTMHTax# = OldRound(GTMHTax# - TBillRec.MHTaxDue)
      MHTax# = OldRound(MHTax# - TBillRec.MHTaxDue)
      TBillRec.MHTaxDue = OldRound(MinBill * Pct4)
      GTMHTax# = OldRound(GTMHTax# + TBillRec.MHTaxDue)
      MHTax# = OldRound(MHTax# + TBillRec.MHTaxDue)
      
      GTMTTax# = OldRound(GTMTTax# - TBillRec.MTTaxDue)
      MTTax# = OldRound(MTTax# - TBillRec.MTTaxDue)
      TBillRec.MTTaxDue = OldRound(MinBill * Pct5)
      GTMTTax# = OldRound(GTMTTax# + TBillRec.MTTaxDue)
      MTTax# = OldRound(MTTax# + TBillRec.MTTaxDue)
      
      TotOpt1# = OldRound(TotOpt1# - TBillRec.OptRevTax1)
      OptRevTax1# = OldRound(OptRevTax1# - TBillRec.OptRevTax1)
      TBillRec.OptRevTax1 = OldRound(MinBill * Pct6)
      TotOpt1# = OldRound(TotOpt1# + TBillRec.OptRevTax1)
      OptRevTax1# = OldRound(OptRevTax1# + TBillRec.OptRevTax1)
      
      TotOpt2# = OldRound(TotOpt2# - TBillRec.OptRevTax2)
      OptRevTax2# = OldRound(OptRevTax2# - TBillRec.OptRevTax2)
      TBillRec.OptRevTax2 = OldRound(MinBill * Pct7)
      TotOpt2# = OldRound(TotOpt2# + TBillRec.OptRevTax2)
      OptRevTax2# = OldRound(OptRevTax2# + TBillRec.OptRevTax2)
      
      TotOpt3# = OldRound(TotOpt3# - TBillRec.OptRevTax3)
      OptRevTax3# = OldRound(OptRevTax3# - TBillRec.OptRevTax3)
      TBillRec.OptRevTax3 = OldRound(MinBill * Pct8)
      TotOpt3# = OldRound(TotOpt3# + TBillRec.OptRevTax3)
      OptRevTax3# = OldRound(OptRevTax3# + TBillRec.OptRevTax3)
      
      TotalLate# = OldRound(TotalLate# - TBillRec.LateTaxDue)
      LateAmt# = OldRound(LateAmt# - TBillRec.LateTaxDue)
      TBillRec.LateTaxDue = OldRound(MinBill * Pct9)
      TotalLate# = OldRound(TotalLate# + TBillRec.LateTaxDue)
      LateAmt# = OldRound(LateAmt# + TBillRec.LateTaxDue)
      
      PctTest# = OldRound(TBillRec.PersTaxNet + TBillRec.FETaxDue + TBillRec.MCTaxDue + TBillRec.MHTaxDue + TBillRec.MTTaxDue)
      PctTest# = OldRound(PctTest# + TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3 + TBillRec.LateTaxDue)
      
      If MinBill < PctTest# Then
        If TBillRec.PersTaxNet > PctTest# - MinBill Then
          TBillRec.PersTaxNet = OldRound(TBillRec.PersTaxNet - (PctTest# - MinBill))
          TBillRec.PersTaxDue = OldRound(TBillRec.PersTaxDue - (PctTest# - MinBill))
          GTPersTax# = OldRound(GTPersTax# - (PctTest# - MinBill))
          GTPersTaxNet# = OldRound(GTPersTaxNet# - (PctTest# - MinBill))
        ElseIf TBillRec.FETaxDue > PctTest# - MinBill Then
          TBillRec.FETaxDue = OldRound(TBillRec.FETaxDue - (PctTest# - MinBill))
          GTFETax# = OldRound(GTFETax# - (PctTest# - MinBill))
          FETax# = OldRound(FETax# - (PctTest# - MinBill))
        ElseIf TBillRec.MCTaxDue > PctTest# - MinBill Then
          TBillRec.MCTaxDue = OldRound(TBillRec.MCTaxDue - (PctTest# - MinBill))
          GTMCTax# = OldRound(GTMCTax# - (PctTest# - MinBill))
          MCTax# = OldRound(MCTax# - (PctTest# - MinBill))
        ElseIf TBillRec.MHTaxDue > PctTest# - MinBill Then
          TBillRec.MHTaxDue = OldRound(TBillRec.MHTaxDue - (PctTest# - MinBill))
          GTMHTax# = OldRound(GTMHTax# - (PctTest# - MinBill))
          MHTax# = OldRound(MHTax# - (PctTest# - MinBill))
        ElseIf TBillRec.MTTaxDue > PctTest# - MinBill Then
          TBillRec.MTTaxDue = OldRound(TBillRec.MTTaxDue - (PctTest# - MinBill))
          GTMTTax# = OldRound(GTMTTax# - (PctTest# - MinBill))
          MTTax# = OldRound(MTTax# - (PctTest# - MinBill))
        ElseIf TBillRec.OptRevTax1 > PctTest# - MinBill Then
          TBillRec.OptRevTax1 = OldRound(TBillRec.OptRevTax1 - (PctTest# - MinBill))
          TotOpt1# = OldRound(TotOpt1# - (PctTest# - MinBill))
          OptRevTax1# = OldRound(OptRevTax1# - (PctTest# - MinBill))
        ElseIf TBillRec.OptRevTax2 > PctTest# - MinBill Then
          TBillRec.OptRevTax2 = OldRound(TBillRec.OptRevTax2 - (PctTest# - MinBill))
          TotOpt2# = OldRound(TotOpt2# - (PctTest# - MinBill))
          OptRevTax2# = OldRound(OptRevTax2# - (PctTest# - MinBill))
        ElseIf TBillRec.OptRevTax3 > PctTest# - MinBill Then
          TBillRec.OptRevTax3 = OldRound(TBillRec.OptRevTax3 - (PctTest# - MinBill))
          TotOpt3# = OldRound(TotOpt3# - (PctTest# - MinBill))
          OptRevTax3# = OldRound(OptRevTax3# - (PctTest# - MinBill))
        ElseIf TBillRec.LateTaxDue > PctTest# - MinBill Then
          TBillRec.LateTaxDue = OldRound(TBillRec.LateTaxDue - (PctTest# - MinBill))
          TotalLate# = OldRound(TotalLate# - (PctTest# - MinBill))
          LateAmt# = OldRound(LateAmt# - (PctTest# - MinBill))
        End If
      ElseIf MinBill > PctTest# Then
        If TBillRec.PersTaxNet > MinBill - PctTest# Then
          TBillRec.PersTaxNet = OldRound(TBillRec.PersTaxNet + (MinBill - PctTest#))
          TBillRec.PersTaxDue = OldRound(TBillRec.PersTaxDue + (MinBill - PctTest#))
          GTPersTax# = GTPersTax# + MinBill - PctTest#
          GTPersTaxNet# = GTPersTaxNet# + MinBill - PctTest#
        ElseIf TBillRec.FETaxDue > MinBill - PctTest# Then
          TBillRec.FETaxDue = OldRound(TBillRec.FETaxDue + (MinBill - PctTest#))
          GTFETax# = OldRound(GTFETax# + (PctTest# - MinBill))
          FETax# = OldRound(FETax# + (PctTest# - MinBill))
        ElseIf TBillRec.MCTaxDue > MinBill - PctTest# Then
          TBillRec.MCTaxDue = OldRound(TBillRec.MCTaxDue + (MinBill - PctTest#))
          GTMCTax# = OldRound(GTMCTax# + (PctTest# - MinBill))
          MCTax# = OldRound(MCTax# + (PctTest# - MinBill))
        ElseIf TBillRec.MHTaxDue > MinBill - PctTest# Then
          TBillRec.MHTaxDue = OldRound(TBillRec.MHTaxDue + (MinBill - PctTest#))
          GTMHTax# = OldRound(GTMHTax# + (PctTest# - MinBill))
          MHTax# = OldRound(MHTax# + (PctTest# - MinBill))
        ElseIf TBillRec.MTTaxDue > MinBill - PctTest# Then
          TBillRec.MTTaxDue = OldRound(TBillRec.MTTaxDue + (MinBill - PctTest#))
          GTMTTax# = OldRound(GTMTTax# + (PctTest# - MinBill))
          MTTax# = OldRound(MTTax# + (PctTest# - MinBill))
        ElseIf TBillRec.OptRevTax1 > MinBill - PctTest# Then
          TBillRec.OptRevTax1 = OldRound(TBillRec.OptRevTax1 + (MinBill - PctTest#))
          TotOpt1# = OldRound(TotOpt1# + (PctTest# - MinBill))
          OptRevTax1# = OldRound(OptRevTax1# + (PctTest# - MinBill))
        ElseIf TBillRec.OptRevTax2 > MinBill - PctTest# Then
          TBillRec.OptRevTax2 = OldRound(TBillRec.OptRevTax2 + (MinBill - PctTest#))
          TotOpt2# = OldRound(TotOpt2# + (PctTest# - MinBill))
          OptRevTax2# = OldRound(OptRevTax2# + (PctTest# - MinBill))
        ElseIf TBillRec.OptRevTax3 > MinBill - PctTest# Then
          TBillRec.OptRevTax3 = OldRound(TBillRec.OptRevTax3 + (MinBill - PctTest#))
          TotOpt3# = OldRound(TotOpt3# + (PctTest# - MinBill))
          OptRevTax3# = OldRound(OptRevTax3# + (PctTest# - MinBill))
        ElseIf TBillRec.LateTaxDue > MinBill - PctTest# Then
          TBillRec.LateTaxDue = OldRound(TBillRec.LateTaxDue + (MinBill - PctTest#))
          TotalLate# = OldRound(TotalLate# + (MinBill - PctTest#))
          LateAmt# = OldRound(LateAmt# + (MinBill - PctTest#))
        End If
      End If
      TBillRec.TotalBillDue = MinBill
    End If
  End If
  
  ThisTBCnt = ThisTBCnt + 1
  TBillRec.OptRevDesc1 = OptRev1Desc
  TBillRec.OptRevDesc2 = OptRev2Desc
  TBillRec.OptRevDesc3 = OptRev3Desc
  TBillRec.BillNumber = TBillRec.BillNumber
  Put TBHandle, ThisTBCnt, TBillRec
  ThisTBCnt = LOF(TBHandle) / Len(TBillRec)
  If TBillRec.TotalBillDue > 0 Then
    If Abs(OverPayAmt) > 0 Then
      OPBillRec.Revenue.LateListPd = 0
      OPBillRec.Revenue.Principle1Pd = 0
      OPBillRec.Revenue.Principle2Pd = 0
      OPBillRec.Revenue.Principle3Pd = 0
      OPBillRec.Revenue.Principle4Pd = 0
      OPBillRec.Revenue.Principle5Pd = 0
      OPBillRec.Revenue.RevOpt1Pd = 0
      OPBillRec.Revenue.RevOpt2Pd = 0
      OPBillRec.Revenue.RevOpt3Pd = 0
      OpenPersTaxBillOverPayFile OPHandle, NumOfOPRecs
      Get TBHandle, ThisTBCnt, TBillRec
      OPBillRec.Revenue.PrePaidAmt = Abs(OverPayAmt)
      OPBillRec.BelongTo = ThisTBCnt 'BillNum&'9/9/05 Billnum is assigned
      'at bill printing so using this as a way of reference needed when posting
      If TBillRec.LateTaxDue > 0 And Abs(OverPayAmt) > 0 Then
        If Abs(OverPayAmt) > TBillRec.LateTaxDue Then
          OverPayAmt = OldRound(OverPayAmt + TBillRec.LateTaxDue)
          OPBillRec.Revenue.LateListPd = TBillRec.LateTaxDue
          OPApplied = OldRound(OPApplied + OPBillRec.Revenue.LateListPd)
        ElseIf Abs(OverPayAmt) <= TBillRec.LateTaxDue Then
          OPBillRec.Revenue.LateListPd = -OverPayAmt
          OPApplied = OldRound(OPApplied - OverPayAmt)
          OverPayAmt = 0
        End If
      End If
      
      If OldRound(TBillRec.PersTaxDue - TBillRec.PPTRADiscnt) > 0 And Abs(OverPayAmt) > 0 Then
        If Abs(OverPayAmt) > OldRound(TBillRec.PersTaxDue - TBillRec.PPTRADiscnt) Then
          OverPayAmt = OldRound(OverPayAmt + (TBillRec.PersTaxDue - TBillRec.PPTRADiscnt))
          OPBillRec.Revenue.Principle1Pd = OldRound(TBillRec.PersTaxDue - TBillRec.PPTRADiscnt)
          OPApplied = OldRound(OPApplied + OPBillRec.Revenue.Principle1Pd)
        ElseIf Abs(OverPayAmt) <= OldRound(TBillRec.PersTaxDue - TBillRec.PPTRADiscnt) Then
          OPBillRec.Revenue.Principle1Pd = -OverPayAmt
          OPApplied = OldRound(OPApplied - OverPayAmt)
          OverPayAmt = 0
        End If
      End If

      If TBillRec.MTTaxDue > 0 And Abs(OverPayAmt) > 0 Then
        If Abs(OverPayAmt) > TBillRec.MTTaxDue Then
          OverPayAmt = OldRound(OverPayAmt + TBillRec.MTTaxDue)
          OPBillRec.Revenue.Principle2Pd = TBillRec.MTTaxDue
          OPApplied = OldRound(OPApplied + TBillRec.MTTaxDue)
        ElseIf Abs(OverPayAmt) <= TBillRec.MTTaxDue Then
          OPBillRec.Revenue.Principle2Pd = -OverPayAmt
          OPApplied = OldRound(OPApplied - OverPayAmt)
          OverPayAmt = 0
        End If
      End If
      
      If TBillRec.MCTaxDue > 0 And Abs(OverPayAmt) > 0 Then
        If Abs(OverPayAmt) > TBillRec.MCTaxDue Then
          OverPayAmt = OldRound(OverPayAmt + TBillRec.MCTaxDue)
          OPBillRec.Revenue.Principle3Pd = TBillRec.MCTaxDue
          OPApplied = OldRound(OPApplied + TBillRec.MCTaxDue)
        ElseIf Abs(OverPayAmt) <= TBillRec.MCTaxDue Then
          OPBillRec.Revenue.Principle3Pd = -OverPayAmt
          OPApplied = OldRound(OPApplied - OverPayAmt)
          OverPayAmt = 0
        End If
      End If
      
      If TBillRec.FETaxDue > 0 And Abs(OverPayAmt) > 0 Then
        If Abs(OverPayAmt) > TBillRec.FETaxDue Then
          OverPayAmt = OldRound(OverPayAmt + TBillRec.FETaxDue)
          OPBillRec.Revenue.Principle4Pd = TBillRec.FETaxDue
          OPApplied = OldRound(OPApplied + TBillRec.FETaxDue)
        ElseIf Abs(OverPayAmt) <= TBillRec.FETaxDue Then
          OPBillRec.Revenue.Principle4Pd = -OverPayAmt
          OPApplied = OldRound(OPApplied - OverPayAmt)
          OverPayAmt = 0
        End If
      End If
      
      If TBillRec.MHTaxDue > 0 And Abs(OverPayAmt) > 0 Then
        If Abs(OverPayAmt) > TBillRec.MHTaxDue Then
          OverPayAmt = OldRound(OverPayAmt + TBillRec.MHTaxDue)
          OPBillRec.Revenue.Principle5Pd = TBillRec.MHTaxDue
          OPApplied = OldRound(OPApplied + TBillRec.MHTaxDue)
        ElseIf Abs(OverPayAmt) <= TBillRec.MHTaxDue Then
          OPBillRec.Revenue.Principle5Pd = -OverPayAmt
          OPApplied = OldRound(OPApplied - OverPayAmt)
          OverPayAmt = 0
        End If
      End If
      
      If TBillRec.OptRevTax1 > 0 And Abs(OverPayAmt) > 0 Then
        If Abs(OverPayAmt) > TBillRec.OptRevTax1 Then
          OverPayAmt = OldRound(OverPayAmt + TBillRec.OptRevTax1)
          OPBillRec.Revenue.RevOpt1Pd = TBillRec.OptRevTax1
          OPApplied = OldRound(OPApplied + TBillRec.OptRevTax1)
        ElseIf Abs(OverPayAmt) <= TBillRec.OptRevTax1 Then
          OPBillRec.Revenue.RevOpt1Pd = -OverPayAmt
          OPApplied = OldRound(OPApplied - OverPayAmt)
          OverPayAmt = 0
        End If
      End If
      
      If TBillRec.OptRevTax2 > 0 And Abs(OverPayAmt) > 0 Then
        If Abs(OverPayAmt) > TBillRec.OptRevTax2 Then
          OverPayAmt = OldRound(OverPayAmt + TBillRec.OptRevTax2)
          OPBillRec.Revenue.RevOpt2Pd = TBillRec.OptRevTax2
          OPApplied = OldRound(OPApplied + TBillRec.OptRevTax2)
        ElseIf Abs(OverPayAmt) <= TBillRec.OptRevTax2 Then
          OPBillRec.Revenue.RevOpt2Pd = -OverPayAmt
          OPApplied = OldRound(OPApplied - OverPayAmt)
          OverPayAmt = 0
        End If
      End If
      
      If TBillRec.OptRevTax3 > 0 And Abs(OverPayAmt) > 0 Then
        If Abs(OverPayAmt) > TBillRec.OptRevTax3 Then
          OverPayAmt = OldRound(OverPayAmt + TBillRec.OptRevTax3)
          OPBillRec.Revenue.RevOpt3Pd = TBillRec.OptRevTax3
          OPApplied = OldRound(OPApplied + TBillRec.OptRevTax3)
        ElseIf Abs(OverPayAmt) <= TBillRec.OptRevTax3 Then
          OPBillRec.Revenue.RevOpt3Pd = -OverPayAmt
          OPApplied = OldRound(OPApplied - OverPayAmt)
          OverPayAmt = 0
        End If
      End If
      
      OPBillRec.Amount = OPApplied
      Put OPHandle, NumOfOPRecs + 1, OPBillRec
      Close OPHandle
      Get TBHandle, ThisTBCnt, TBillRec
      TBillRec.OverPayAmt = OPApplied
      Put TBHandle, ThisTBCnt, TBillRec
    End If
      
    BillNum& = BillNum& + 1
  End If
  
  Return
  
GetPersInfo:
  LunenBurgYN = False
  PersValue# = 0
  PersTaxDue# = 0
  MHValue# = 0
  MCValue# = 0
  FEValue# = 0
  MTValue# = 0
  PropertyRec! = TaxCust.FirstPersRec
  PPTRAVal# = 0
  PPTRAVal# = 0
  TPersTaxDue# = 0
  TPPTRAVal# = 0
  TPPTRADisc# = 0
  OptRevTax1# = 0
  OptRevTax2# = 0
  OptRevTax3# = 0
  Do While PropertyRec! > 0
    PPTRADisc# = 0
    Get PHandle, PropertyRec!, PersRec
    ThisProp = PropertyRec!
    ThisMaxVehVal = MaxVehVal
    ThisMinVehVal = MinVehVal
    If PersRec.Deleted = True Then GoTo KeepGoing
    If SupOnly = "Y" And PersRec.DISCOV <> "Y" Then GoTo KeepGoing
'    If PersRec.PersVal < ThisMinVehVal Then
'      GoTo KeepGoing
'    End If
'    If PersRec.DISCOV = "N" And SupOnly = "Y" Then GoTo KeepGoing
'    If PersRec.LastYrPrinted = WhatPYear Then Discovery$ = "Y" 'took out 12/5/06
    If MultiYear > 1 Or PersRec.DISCOV = "Y" Then GoTo GoAhead 'new for 2.05
    If (PersRec.LastYrPrinted <> WhatPYear) Or (PersRec.DISCOV = "Y") Or (PersRec.LastYrPrinted = WhatPYear) Then
GoAhead:
      PYear$ = CInt(PersRec.TaxBillYear)
      If Val(PYear$) > 0 Then
        PYearInt = Val(PYear$)
      Else
        GoTo KeepGoing
      End If
    
      If InStr(TaxMasterRec.Name, "CHILHOWIE") Then GoTo SkipOpt1
      OptRevTax1# = OldRound(FigureOptRevTax1(ThisProp&, PHandle, "P"))
      TBillRec.OptRevTax1 = OptRevTax1# + TBillRec.OptRevTax1
'      TotOpt1 = OldRound(TotOpt1 + OptRevTax1#)
SkipOpt1:
      OptRevTax2# = OldRound(FigureOptRevTax2(ThisProp&, PHandle, "P"))
      TBillRec.OptRevTax2 = OptRevTax2# + TBillRec.OptRevTax2
'      TotOpt2 = OldRound(TotOpt2 + OptRevTax2#)
      If InStr(TaxMasterRec.Name, "LUNENBURG") Then 'added 10/12/06
        OptRevTax3# = 0
      Else
        OptRevTax3# = OldRound(FigureOptRevTax3(ThisProp&, PHandle, "P"))
      End If
      TBillRec.OptRevTax3 = OptRevTax3# + TBillRec.OptRevTax3
'      TotOpt3 = OldRound(TotOpt3 + OptRevTax3#)
    
      Factor! = 1
      Rate = PersRec.ProrateVal
      If Rate > 0 Then
        Factor! = Rate / 12
        TBillRec.Prorate = "Y"
      Else
        TBillRec.Prorate = "N"
        Factor = 1
      End If
      
      PersValue# = OldRound((PersRec.PersVal * Factor!) + PersValue)
      FEValue# = OldRound((PersRec.CVALUE) + FEValue#)
      MHValue# = OldRound((PersRec.MHValue) + MHValue#)
      MCValue# = OldRound((PersRec.MCValue) + MCValue#)
      MTValue# = OldRound((PersRec.MTValue) + MTValue#)
      If ABalance# > 0 Then
        If LawChngDate > 0 Then
          If Date2Num(fptxtBillDate.Text) >= LawChngDate Then
            GoTo NoDisc 'new for 2006
          End If
        End If
      End If
      If PPTRAYN = False Then GoTo NoDisc
      
      If PersRec.PPTRAYN = "Y" Then
        If OldRound(PersRec.PersVal * Factor!) > ThisMaxVehVal Then
          PPTRAVal# = ThisMaxVehVal
        Else
          PPTRAVal# = OldRound(PersRec.PersVal * Factor!)
        End If
        If PPTRAVal# <= (ThisMinVehVal * Factor!) Then
          PPTRADisc# = OldRound(((PPTRAVal# / 100) * Factor) * PERSRATE#) '2/21/06
        Else
          PPTRADisc# = OldRound(((PPTRAVal# / 100) * (PERC! / 100)) * PERSRATE#)
        End If
        TPPTRAVal# = OldRound(TPPTRAVal# + PPTRAVal#)
      End If
NoDisc:
      PersTaxDue = OldRound#(((PersRec.PersVal / 100) * Factor!) * PERSRATE#)
      If PersTaxDue < 0 Then PersTaxDue = 0
    End If
    If PersRec.OptRev3Chrg > 0 Then LunenBurgYN = True
    
    TPersTaxDue = OldRound(TPersTaxDue + PersTaxDue) ' - (OptRevTax1# + OptRevTax2# + OptRevTax3#))
    TPPTRADisc# = OldRound(PPTRADisc + TPPTRADisc#)
'    If PersRec.PersVal > 0 Then
''      Print #AHandle, QPTrim$(TaxCust.CountyAcctString) + "~" + Using$("$###,###0.00", PersRec.PersVal) + "~" + Using$("$##,##0.00", PersTaxDue)
'      Print #AHandle, CStr(TaxCust.Acct) + "~" + Using$("$###,###0.00", PersRec.PersVal)
'    End If
KeepGoing:
    PropertyRec! = PersRec.NextRec
  Loop
  TBillRec.OptRevTax1 = TBillRec.OptRevTax1 / MultiYear
  TotOpt1 = TotOpt1 + TBillRec.OptRevTax1
  TBillRec.OptRevTax2 = TBillRec.OptRevTax2 / MultiYear
  TotOpt2 = TotOpt2 + TBillRec.OptRevTax2
  TBillRec.OptRevTax3 = TBillRec.OptRevTax3 / MultiYear
  TotOpt3 = TotOpt3 + TBillRec.OptRevTax3
  
  TPersTaxDue = OldRound(TPersTaxDue / MultiYear) 'adding multiyear is new on 6/22/06
  
  TPPTRADisc# = OldRound(TPPTRADisc# / MultiYear) 'adding multiyear is new on 6/22/06
  
  MHTax# = OldRound((MHValue# / 100) * MHRate#)
  MHTax# = OldRound(MHTax# / MultiYear) 'adding multiyear is new on 6/22/06
  MCTax# = OldRound((MCValue# / 100) * MCRate#)
  MCTax# = OldRound(MCTax# / MultiYear) 'adding multiyear is new on 6/22/06
  
  FETax# = OldRound((FEValue# / 100) * FERate#)
  FETax# = OldRound(FETax# / MultiYear) 'adding multiyear is new on 6/22/06
  MTTax# = OldRound((MTValue# / 100) * MTRate#)
  MTTax# = OldRound(MTTax# / MultiYear) 'adding multiyear is new on 6/22/06
  
  If OldRound(TPersTaxDue + MHTax# + MCTax# + FETax# + MTTax# - TPPTRADisc#) = 0 Then
    If OldRound(TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3) = 0 Then 'added 10/18/06
      If fpcmbInclNoBills.Text = "N" Then
        If InStr(TaxMasterRec.Name, "LUNENBURG") Then
          If OldRound((TPersTaxDue) - (TBillRec.OptRevTax1 + TBillRec.OptRevTax2)) = 0 Then
            Return
          End If
        Else
          Return
        End If
      End If
    End If
  End If
  TBillRec.FETaxDue = FETax#
  TBillRec.FETaxRate = CDbl(fpDblSnglFERate.Value)
  TBillRec.FEValue = FEValue# ' * MultiYear
  TBillRec.MCTaxDue = MCTax#
  TBillRec.MCTaxRate = CDbl(fpDblSnglMCRate.Value)
  TBillRec.MCValue = MCValue# ' * MultiYear
  TBillRec.MHTaxDue = MHTax#
  TBillRec.MHTaxRate = CDbl(fpDblSnglMHRate.Value)
  TBillRec.MHValue = MHValue# ' * MultiYear
  TBillRec.MTTaxDue = MTTax#
  TBillRec.MTTaxRate = CDbl(fpDblSnglMTRate.Value)
  TBillRec.MTValue = MTValue# '* MultiYear
  GTMHTax# = OldRound(GTMHTax# + MHTax#)
  GTMCTax# = OldRound(GTMCTax# + MCTax#)
  GTFETax# = OldRound(GTFETax# + FETax#)
  GTMTTax# = OldRound(GTMTTax# + MTTax#)
  GTPersTax# = OldRound(GTPersTax# + TPersTaxDue#)
  TBillRec.PersTaxNet = OldRound(TPersTaxDue# - TPPTRADisc#)
  TBillRec.MultiYrVal = MultiYear
  TBillRec.PersTaxDue = TPersTaxDue#
  TBillRec.PPTRADiscnt = TPPTRADisc#
  GTPersTaxNet# = OldRound(GTPersTaxNet# + TBillRec.PersTaxNet)
  GTPTTRADisc# = OldRound(GTPTTRADisc# + TPPTRADisc#)
  TBillRec.PersValue = PersValue# ' * MultiYear
  GTPersValue# = OldRound(GTPersValue# + PersValue#)
  GTFEValue# = OldRound(GTFEValue# + TBillRec.FEValue#)
  GTMCValue# = OldRound(GTMCValue# + TBillRec.MCValue#)
  GTMHValue# = OldRound(GTMHValue# + TBillRec.MHValue#)
  GTMTValue# = OldRound(GTMTValue# + TBillRec.MTValue#)
  TBillRec.CustPin = TaxCust.PIN
  TBillRec.InternalPin = PersRec.InternalPin
  TBillRec.PersPin = QPTrim$(PersRec.PropPin)
  TBillRec.ExptValue = 0
  GTPPTRADisc = OldRound(GTPPTRADisc + TPPTRADisc#)
  TBillRec.PPTRAValue = TPPTRAVal#
  GTPPTRAVal = OldRound(GTPPTRAVal# + TPPTRAVal#)
  TBillRec.PersPropRecord = TaxCust.FirstPersRec
  TBillRec.PersTaxRate = PERSRATE#
  
  If InStr(TaxMasterRec.Name, "CHILHOWIE") Then 'added 11/3/06
    PersBalTest# = OldRound(TPersTaxDue + MHTax# + MCTax# + FETax# + MTTax# + TBillRec.OptRevTax2 + TBillRec.OptRevTax3)
    CalcDiff# = OldRound(TaxMasterRec.MinBill - PersBalTest#)
    If PersBalTest# > 0 And PersBalTest# < TaxMasterRec.MinBill Then
      TBillRec.ChillHowieFudge = CalcDiff#
      ChilhowieGFudge = OldRound(ChilhowieGFudge + TBillRec.ChillHowieFudge)
      If TBillRec.PPTRADiscnt > 0 And TBillRec.PPTRAValue <= TaxMasterRec.MinVehTaxVal Then
         TBillRec.PPTRADiscnt = PersBalTest# + CalcDiff#
      End If
    End If
    TBillRec.OptRevTax1 = TBillRec.ChillHowieFudge
    TotOpt1 = OldRound(TotOpt1 + TBillRec.OptRevTax1)
    TBillRec.TotalBillDue = OldRound(TBillRec.PersTaxDue + TBillRec.FETaxDue + TBillRec.MCTaxDue + TBillRec.MHTaxDue + TBillRec.MTTaxDue - TBillRec.PPTRADiscnt)
    TBillRec.TotalBillDue = OldRound(TBillRec.TotalBillDue + TBillRec.OptRevTax2 + TBillRec.OptRevTax3 + TBillRec.ChillHowieFudge)
    GoTo ChilhowieSkip3
  End If
  
  TBillRec.TotalBillDue = OldRound(TBillRec.PersTaxDue + TBillRec.FETaxDue + TBillRec.MCTaxDue + TBillRec.MHTaxDue + TBillRec.MTTaxDue - TBillRec.PPTRADiscnt)
  TBillRec.TotalBillDue = OldRound(TBillRec.TotalBillDue + TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3)
ChilhowieSkip3:
  
  If PersRec.LateList = "Y" Then
    LateAmt# = OldRound(LateAmt# + OldRound(TBillRec.PersTaxDue - TBillRec.PPTRADiscnt) * (LateList# / 100))
    TBillRec.LateTaxDue = LateAmt#
    TBillRec.TotalBillDue = OldRound(TBillRec.TotalBillDue + TBillRec.LateTaxDue)
  End If
  TBillRec.DueDate = Date2Num(fptxtDueDate.Text)
  TBillRec.RDesc1 = QPTrim$(PersRec.DESC1)
Lunenburg:
  Dim ThisTotBill As Double
  ThisTotBill = OldRound((TBillRec.TotalBillDue + TBillRec.PPTRADiscnt) - (TBillRec.OptRevTax1 + TBillRec.OptRevTax2))
  If InStr(TaxMasterRec.Name, "LUNENBURG") Then 'added 10/12/06
    If LunenBurgYN = True Then
      If ThisTotBill < 10 Then
        TBillRec.OptRevTax3 = ThisTotBill
      ElseIf ThisTotBill >= 10 And ThisTotBill <= 100 Then
        TBillRec.OptRevTax3 = 10
      Else
        TBillRec.OptRevTax3 = OldRound(ThisTotBill * 0.1)
      End If
      TotOpt3 = OldRound(TotOpt3 + TBillRec.OptRevTax3)
    End If
    TBillRec.TotalBillDue = OldRound(TBillRec.PersTaxDue + TBillRec.FETaxDue + TBillRec.MCTaxDue + TBillRec.MHTaxDue + TBillRec.MTTaxDue - TBillRec.PPTRADiscnt)
    TBillRec.TotalBillDue = OldRound(TBillRec.TotalBillDue + TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3)
'    TotOpt3 = TBillRec.OptRevTax3
  End If
  
  GoSub WriteIt2Disk

  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPrebilling", "PrintGraphics", Erl)
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

Private Sub PrintPersText()
  Dim TBillRec As VAPPTaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim OPBillRec As TaxTransactionType
  Dim OPHandle As Integer
  Dim NumOfOPRecs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  Dim NewTBillRec As VAPPTaxBillType
  Dim LateAmt#
  Dim Inactive As Integer
  Dim PastFlagSet As Boolean
  Dim CustName$, CitySt$, CustAcct&
  Dim ABalance#
  Dim Balance#
  Dim TransRecord&
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim TempAddOn As TempTaxBillAddOn
  Dim AOHandle As Integer
  Dim AddOnCnt As Integer
  Dim PersValue#
  Dim Discovery$
  Dim PersTaxDue#
  Dim NextPersRec&
  Dim TotalPers#
  Dim TotalOverPay#
  Dim TotalEx#
  Dim NumBills&
  Dim RptHandle As Integer
  Dim TaxPreRptFile As String
  Dim TValue#
  Dim dlm$
  Dim TotalBills#
  Dim TotalLate#
  Dim TotalPast#
  Dim ThisTownship$
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim ThisTaxRec As Long
  Dim CountyNum As Long
  Dim CountyName$
  Dim CycleNum As Long
  Dim CycleName$
  Dim BillCnt As Long
  Dim OverPay As Boolean
  Dim OverPayAmt As Double
  Dim OPApplied As Double
  Dim ThisTBCnt As Long
  Dim ThisTest$
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim OptFlag As Boolean
  Dim PERSRATE#
  Dim MHRate#
  Dim MCRate#
  Dim FERate#
  Dim MTRate#
  Dim PersTax#
  Dim MHTax#
  Dim MCTax#
  Dim FETax#
  Dim MTTax#
  Dim GTPersTax#
  Dim GTPersTaxNet#
  Dim GTMHTax#
  Dim GTMCTax#
  Dim GTFETax#
  Dim GTMTTax#
  Dim OptRev1Tot#
  Dim OptRev2Tot#
  Dim OptRev3Tot#
  Dim MHValue#
  Dim MCValue#
  Dim FEValue#
  Dim MTValue#
  Dim GTPersValue#
  Dim GTMHValue#
  Dim GTMCValue#
  Dim GTFEValue#
  Dim GTMTValue#
  Dim PropertyRec!
  Dim PPTRAVal#
  Dim GTPPTRADisc#
  Dim GTPPTRAVal#
  Dim PPTRADisc#
  Dim PERC!
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPRecs As Long
  Dim Factor!
  Dim Rate As Double
  Dim PYear$
  Dim PYearInt As Integer
  Dim TxFile As Integer
  Dim Prorate$
  Dim GTPTTRADisc#
  Dim GTPTTRAVal#
  Dim MultiYear As Integer
  Dim MaxLines As Integer
  Dim FF$
  Dim LineCnt As Integer
  Dim Page As Integer
  Dim ThisCName As String * 45
  Dim ThisRate As Double
  Dim TPersTaxDue As Double
  Dim TPPTRAVal As Double
  Dim TPPTRADisc As Double
  Dim ThisMaxVehVal As Double, ThisMinVehVal As Double
  Dim OptRevTax1 As Double
  Dim OptRevTax2 As Double
  Dim OptRevTax3 As Double
  Dim OptRev1Desc As String * 20
  Dim OptRev2Desc As String * 20
  Dim OptRev3Desc As String * 20
  Dim TotOpt1 As Double
  Dim TotOpt2 As Double
  Dim TotOpt3 As Double
  Dim TotReal As Double
  Dim LunenBurgYN As Boolean
  Dim AHandle As Integer
  Dim PersBalTest#
  Dim CalcDiff#
  Dim PPTRAValue#
  Dim ChilhowieGFudge As Single
  Dim Pct1#, Pct2#, Pct3#, Pct4#, Pct5#, Pct6#
  Dim Pct7#, Pct8#, Pct9#, PctTot#, PctTest#
  
  On Error GoTo ERRORSTUFF
  AHandle = FreeFile
  Open "latelisttext.dat" For Output As AHandle
  FF$ = Chr(12)
  MaxLines = 56
  
  PERSRATE# = fpDblSnglPersRate
  MHRate# = fpDblSnglMHRate
  MCRate# = fpDblSnglMCRate
  FERate# = fpDblSnglFERate
  MTRate# = fpDblSnglMTRate
  
  Prorate$ = "Y"
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  MultiYear = TaxMasterRec.MultiYear
  If QPTrim$(TaxMasterRec.POptRev1) <> "" Then
    RSet OptRev1Desc$ = QPTrim$(TaxMasterRec.POptRev1)
  Else
    OptRev1Desc$ = ""
  End If
  
  If QPTrim$(TaxMasterRec.POptRev2) <> "" Then
    RSet OptRev2Desc$ = QPTrim$(TaxMasterRec.POptRev2)
  Else
    OptRev2Desc$ = ""
  End If
  
  If QPTrim$(TaxMasterRec.POptRev3) <> "" Then
    RSet OptRev3Desc$ = QPTrim$(TaxMasterRec.POptRev3)
  Else
    OptRev3Desc$ = ""
  End If
  
  fpcmbCycle.Col = 1
  CycleName = QPTrim$(fpcmbCycle.ColText)
  If CycleName = "" Then CycleName = "N/A"
  If fpcmbCycle.Enabled = True And CycleName <> "NO CYCLE" Then
    fpcmbCycle.Col = 0
    CycleNum = CLng(fpcmbCycle.ColText)
  Else
    CycleNum = -1
  End If
  
  fpcmbCounty.Col = 1
  CountyName = QPTrim$(fpcmbCounty.ColText)
  If CountyName = "" Then CountyName = "N/A"
  If fpcmbCounty.Enabled = True And CountyName <> "ALL COUNTIES" Then
    fpcmbCounty.Col = 0
    CountyNum = CLng(fpcmbCounty.ColText)
  Else
    CountyNum = -1
  End If
  
  
  If QPTrim$(fpcmbTownships.Text) = "ALL TOWNSHIPS" Then
    ThisTownship = "ALL"
  Else
    ThisTownship = QPTrim$(fpcmbTownships.Text)
    ThisTownship = Mid(ThisTownship, 1, Len(ThisTownship) - 4)
    ThisTownship = QPTrim$(ThisTownship)
  End If
  dlm$ = "~"
  If Exist(PersTaxBillFile) Then
    KillFile PersTaxBillFile
  End If
  If Exist(PersTaxBillOPFile) Then
    KillFile PersTaxBillOPFile
  End If
  If Exist("TMPPERSBLADD.DAT") Then 'tax bill addon
    KillFile "TMPPERSBLADD.DAT"
  End If
  
  AddOnCnt = 0
  OpenTaxBillPersAddOn AOHandle
  OpenPersPropFile PHandle, NumOfPRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenPersTaxBillFile TBHandle, NumOfTBRecs
  
  Inactive = 0
  
  frmVATaxShowPctComp.Label1 = "Creating Tax Personal Pre-Billing Register"
  frmVATaxShowPctComp.Show , Me
  frmVATaxShowPctComp.cmdCancel.Visible = False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  EnableCloseButton Me.hwnd, False
  If UseOpt = "Y" Then
    NumOfTCRecs = OptCnt
  End If
  If UseSS = "Y" Then
    NumOfTCRecs = SSCnt
  End If
  
  GoSub PPTRA
  
  For x = 1 To NumOfTCRecs
    TBillRec = NewTBillRec
    If UsingIdx = True Then
      Get TCHandle, CustArr(x), TaxCust
      ThisTaxRec = CustArr(x)
    Else
      Get TCHandle, x, TaxCust
      ThisTaxRec = x
    End If
    If TaxCust.FirstPersRec = 0 Then GoTo PreBillSkip
    If TaxCust.TaxExempt = "Y" Then
      GoTo PreBillSkip
    End If
    If ThisTownship <> "ALL" Then
      If QPTrim$(TaxCust.TownShip) <> ThisTownship Then
        GoTo PreBillSkip
      End If
    End If
    
    If CycleNum >= 0 Then
      If TaxCust.Cycle <> CycleNum Then
        GoTo PreBillSkip
      End If
    End If
    
    If CountyNum >= 0 Then
      If TaxCust.County4BillNum <> CountyNum Then
        GoTo PreBillSkip
      End If
    End If
    
    LateAmt# = 0
    OPApplied = 0
    If TaxCust.Deleted <> 0 Then
      GoTo PreBillSkip:
    End If
    If QPTrim$(TaxCust.Active) <> "Y" Then
      Inactive = Inactive + 1
      GoTo PreBillSkip:
    End If
    PastFlagSet = 0             'Initialize Past Balance Flag
    OverPayAmt = 0
    If UsingIdx = True Then
      OverPayAmt = GetCustPersBalance(CustArr(x), -1)
      If OverPayAmt < 0 Then
        OverPay = True
      Else
        OverPayAmt = 0
        OverPay = False
      End If
    Else
      OverPayAmt = GetCustPersBalance(x, -1)
      If OverPayAmt < 0 Then
        OverPay = True
      Else
        OverPayAmt = 0
        OverPay = False
      End If
    End If
    ThisTest = CStr(OverPayAmt)
    If InStr(ThisTest, "E") Then OverPayAmt = 0
    
    If TaxCust.FirstPersRec <= 0 Then
      GoSub SetCustInfo
      GoSub WriteIt2Disk
      GoTo PreBillSkip
    End If
    
    GoSub SetCustInfo
    GoSub GetPersInfo
    
PreBillSkip:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTCRecs
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
  
'  Close TCHandle
  Close AOHandle
  Close PHandle
  
  If ThisTBCnt = 0 Then
    Call TaxMsg(900, "Using the parameters selected there are no customers who qualify for a tax charge.")
    Close
    fptxtPCurrYear.SetFocus
    Exit Sub
  End If
  
  TotalPers# = 0
  NumBills& = 0
  
  TaxPreRptFile$ = "TAXRPTS\TaxPersPreBill.PRN"
  
  RptHandle = FreeFile
  Open TaxPreRptFile For Output As #RptHandle
  
  GoSub PreBillHeading
  
  Close TBHandle
  OpenPersTaxBillFile TBHandle, NumOfTBRecs
  
  For x = 1 To NumOfTBRecs
    Get TBHandle, x, TBillRec
    Get TCHandle, TBillRec.CustRec, TaxCust
'    If TBillRec.TotalBillDue = 0 And (TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax1) > 0 Then Stop
'    If Discovery$ = "Y" And TBillRec.BillNumber = -1 Then
    If TBillRec.BillNumber = -1 Then 'commented out "Discovery$ = Y" on 12/5/06
       GoTo NotThisOne
    Else
      If TBillRec.TotalBillDue = 0 And TBillRec.PriorYrBalance = 0 Then
        If fpcmbInclNoBills.Text = "N" Then
          GoTo NotThisOne
        End If
      End If
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PreBillHeading
      End If
      If TBillRec.TotalBillDue > 0 Then
        NumBills& = NumBills& + 1
      End If
      If QPTrim$(TBillRec.Prorate) <> "Y" Then
        ThisCName = QPTrim$(TBillRec.CustName)
      Else
        ThisCName = QPTrim$(TBillRec.CustName) + "**Prorated**"
      End If
      Print #RptHandle, Using$("####0", TBillRec.CustRec); Tab(8); ThisCName;
      Print #RptHandle, Tab(56); Using$("$##,###,##0.00", TBillRec.PriorYrBalance); Tab(75); Using$("###0", NumBills&);
      Print #RptHandle, Tab(80); Using$("$##,###,##0.00", TBillRec.TotalBillDue - TBillRec.LateTaxDue) 'added late tax on 8/14/06
      LineCnt = LineCnt + 1
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PreBillHeading
        Print #RptHandle, Using$("####0", TBillRec.CustRec); Tab(8); ThisCName;
        Print #RptHandle, Tab(56); Using$("$##,###,##0.00", TBillRec.PriorYrBalance); Tab(75); Using$("###0", NumBills&);
        Print #RptHandle, Tab(80); Using$("$##,###,##0.00", TBillRec.TotalBillDue)
        LineCnt = LineCnt + 1
      End If
      Print #RptHandle, Using$("$##,###,##0.00", TBillRec.PersValue); Tab(15); Using$("$##,###,##0.00", TBillRec.FEValue);
      Print #RptHandle, Tab(30); Using$("$##,###,##0.00", TBillRec.MHValue); Tab(44); Using$("$##,###,##0.00", TBillRec.MCValue);
      Print #RptHandle, Tab(60); Using$("$##,###,##0.00", TBillRec.MTValue); Tab(74); Using$("$##,###,##0.00", TBillRec.LateTaxDue)
      LineCnt = LineCnt + 1
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PreBillHeading
        Print #RptHandle, Using$("####0", TBillRec.CustRec); Tab(8); ThisCName;
        Print #RptHandle, Tab(56); Using$("$##,###,##0.00", TBillRec.PriorYrBalance); Tab(75); Using$("###0", NumBills&);
'        Print #RptHandle, Tab(80); Using$("$##,###,##0.00", TBillRec.PersTaxNet)
        Print #RptHandle, Tab(80); Using$("$##,###,##0.00", TBillRec.TotalBillDue)
        LineCnt = LineCnt + 1
      End If
      Print #RptHandle, Using$("$##,###,##0.00", TBillRec.PersTaxDue); Tab(15); Using$("$##,###,##0.00", TBillRec.FETaxDue);
      Print #RptHandle, Tab(30); Using$("$##,###,##0.00", TBillRec.MHTaxDue); Tab(44); Using$("$##,###,##0.00", TBillRec.MCTaxDue);
      Print #RptHandle, Tab(60); Using$("$##,###,##0.00", TBillRec.MTTaxDue)
      LineCnt = LineCnt + 1
      If TBillRec.PPTRADiscnt > 0 Then
        If LineCnt > MaxLines - 3 Then
          Print #RptHandle, FF$
          GoSub PreBillHeading
          Print #RptHandle, Using$("####0", TBillRec.CustRec); Tab(8); ThisCName;
          Print #RptHandle, Tab(56); Using$("$##,###,##0.00", TBillRec.PriorYrBalance); Tab(74); Using$("###0", NumBills&);
          Print #RptHandle, Tab(80); Using$("$##,###,##0.00", TBillRec.TotalBillDue)
          LineCnt = LineCnt + 1
        End If
        If TBillRec.OverPayAmt > 0 Then
          If PPTRAYN = True Then
            Print #RptHandle, Using$("$##,###,##0.00", TBillRec.PPTRAValue); Tab(17); "<----PPTRA Value";
          End If
          Print #RptHandle, Tab(40); "Credit Bal Applied to Tax Due: "; Tab(72); Using$("$##,###,##0.00", -TBillRec.OverPayAmt)
          LineCnt = LineCnt + 1
          If PPTRAYN = True Then
            Print #RptHandle, Using$("$##,###,##0.00", TBillRec.PPTRADiscnt); Tab(17); "<----PPTRA Discount";
            LineCnt = LineCnt + 1
          End If
          Print #RptHandle, Tab(57); "Total Tax Due: "; Tab(72); Using$("$##,###,##0.00", (TBillRec.TotalBillDue - TBillRec.OverPayAmt))
          LineCnt = LineCnt + 1
        Else
          If PPTRAYN = True Then
            Print #RptHandle, Using$("$##,###,##0.00", TBillRec.PPTRAValue); Tab(17); "<----PPTRA Value"
            Print #RptHandle, Using$("$##,###,##0.00", TBillRec.PPTRADiscnt); Tab(17); "<----PPTRA Discount"
            LineCnt = LineCnt + 2
          End If
        End If
      ElseIf TBillRec.PPTRADiscnt <= 0 Then
        If TBillRec.OverPayAmt > 0 Then
          Print #RptHandle, Tab(40); "Credit Bal Applied to Tax Due: "; Tab(72); Using$("$##,###,##0.00", -TBillRec.OverPayAmt)
          LineCnt = LineCnt + 1
          If PPTRAYN = True Then
            Print #RptHandle, Using$("$##,###,##0.00", TBillRec.PPTRADiscnt); Tab(17); "<----PPTRA Discount"
            LineCnt = LineCnt + 1
          End If
        End If
      End If
      If LineCnt > MaxLines - 6 Then
        Print #RptHandle, FF$
        GoSub PreBillHeading
        Print #RptHandle, Using$("####0", TBillRec.CustRec); Tab(8); ThisCName;
        Print #RptHandle, Tab(56); Using$("$##,###,##0.00", TBillRec.PriorYrBalance); Tab(74); Using$("###0", NumBills&);
        Print #RptHandle, Tab(80); Using$("$##,###,##0.00", TBillRec.TotalBillDue)
        LineCnt = LineCnt + 1
      End If
      If TBillRec.OptRevTax1 > 0 Or TBillRec.OptRevTax2 > 0 Or TBillRec.OptRevTax3 > 0 Then
        Print #RptHandle, Tab(5); "OPTIONAL REVENUE"
        Print #RptHandle, Tab(5); String(88, ".")
        LineCnt = LineCnt + 2
        If TBillRec.OptRevTax1 > 0 And TBillRec.OptRevTax2 = 0 And TBillRec.OptRevTax3 = 0 Then
          Print #RptHandle, Tab(10); OptRev1Desc$
          Print #RptHandle, Tab(20); Using$("$##,##0.00", TBillRec.OptRevTax1)
        ElseIf TBillRec.OptRevTax1 = 0 And TBillRec.OptRevTax2 > 0 And TBillRec.OptRevTax3 = 0 Then
          Print #RptHandle, Tab(40); OptRev2Desc$
          Print #RptHandle, Tab(50); Using$("$##,##0.00", TBillRec.OptRevTax2)
        ElseIf TBillRec.OptRevTax1 = 0 And TBillRec.OptRevTax2 = 0 And TBillRec.OptRevTax3 > 0 Then
          Print #RptHandle, Tab(70); OptRev3Desc$
          Print #RptHandle, Tab(80); Using$("$##,##0.00", TBillRec.OptRevTax3)
        ElseIf TBillRec.OptRevTax1 > 0 And TBillRec.OptRevTax2 > 0 And TBillRec.OptRevTax3 = 0 Then
          Print #RptHandle, Tab(10); OptRev1Desc$;
          Print #RptHandle, Tab(40); OptRev2Desc$
          Print #RptHandle, Tab(20); Using$("$##,##0.00", TBillRec.OptRevTax1);
          Print #RptHandle, Tab(50); Using$("$##,##0.00", TBillRec.OptRevTax2)
        ElseIf TBillRec.OptRevTax1 > 0 And TBillRec.OptRevTax2 = 0 And TBillRec.OptRevTax3 > 0 Then
          Print #RptHandle, Tab(10); OptRev1Desc$;
          Print #RptHandle, Tab(70); OptRev3Desc$
          Print #RptHandle, Tab(20); Using$("$##,##0.00", TBillRec.OptRevTax1);
          Print #RptHandle, Tab(80); Using$("$##,##0.00", TBillRec.OptRevTax3)
        ElseIf TBillRec.OptRevTax1 = 0 And TBillRec.OptRevTax2 > 0 And TBillRec.OptRevTax3 > 0 Then
          Print #RptHandle, Tab(40); OptRev2Desc$;
          Print #RptHandle, Tab(70); OptRev3Desc$
          Print #RptHandle, Tab(50); Using$("$##,##0.00", TBillRec.OptRevTax2);
          Print #RptHandle, Tab(80); Using$("$##,##0.00", TBillRec.OptRevTax3)
        ElseIf TBillRec.OptRevTax1 > 0 And TBillRec.OptRevTax2 > 0 And TBillRec.OptRevTax3 > 0 Then
          Print #RptHandle, Tab(10); OptRev1Desc$;
          Print #RptHandle, Tab(40); OptRev2Desc$;
          Print #RptHandle, Tab(70); OptRev3Desc$
          Print #RptHandle, Tab(20); Using$("$##,##0.00", TBillRec.OptRevTax1);
          Print #RptHandle, Tab(50); Using$("$##,##0.00", TBillRec.OptRevTax2);
          Print #RptHandle, Tab(80); Using$("$##,##0.00", TBillRec.OptRevTax3)
        End If
      End If
      Print #RptHandle, String(93, "-")
      LineCnt = LineCnt + 1
      TValue# = OldRound#(TBillRec.PersValue + TBillRec.FEValue + TBillRec.MCValue + TBillRec.MHValue + TBillRec.MTValue - TBillRec.ExptValue)
      TotalLate# = OldRound#(TotalLate# + TBillRec.LateTaxDue)
'      TotalPast# = OldRound#(TotalPast# + TBillRec.PriorYrBalance) 'too out 8/14/06
      If TBillRec.PriorYrBalance > 0 Then 'added 8/14/06
        TotalPast# = OldRound#(TotalPast# + TBillRec.PriorYrBalance)
      Else 'added 8/14/06
'        TotalPast# = 0 'added 8/14/06
      End If 'added 8/14/06
      If InStr(TaxMasterRec.Name, "CHILHOWIE") Then
        If TBillRec.TotalBillDue = 0 Then 'And TBillRec.ChillHowieFudge > 0 Then
          TotalLate# = TotalLate# - TBillRec.LateTaxDue
          GTFETax = GTFETax - TBillRec.FETaxDue
          GTMHTax = GTMHTax - TBillRec.MHTaxDue
          GTMCTax = GTMCTax - TBillRec.MCTaxDue
          GTMTTax = GTMTTax - TBillRec.MTTaxDue
          TotOpt2 = TotOpt2 - TBillRec.OptRevTax2
          TotOpt3 = TotOpt3 - TBillRec.OptRevTax3
          GTPersTax = GTPersTax - TBillRec.PersTaxDue
          GTPPTRADisc# = GTPPTRADisc - TBillRec.PersTaxDue ' - TBillRec.PPTRADiscnt)
        ElseIf TBillRec.TotalBillDue > 0 Then
          If TBillRec.PersTaxDue > 0 Then  '11/9/06
            GTPersTax = GTPersTax + TBillRec.OptRevTax1
            GTPersTaxNet = GTPersTaxNet + TBillRec.OptRevTax1
          ElseIf TBillRec.MHTaxDue > 0 Then
            GTMHTax = GTMHTax + TBillRec.OptRevTax1
          ElseIf TBillRec.MCTaxDue > 0 Then
            GTMCTax = GTMCTax + TBillRec.OptRevTax1
          ElseIf TBillRec.MTTaxDue > 0 Then
            GTMTTax = GTMTTax + TBillRec.OptRevTax1
          ElseIf TBillRec.FETaxDue > 0 Then
            GTFETax = GTFETax + TBillRec.OptRevTax1
          ElseIf TBillRec.OptRevTax2 > 0 Then
            TotOpt2 = TotOpt2 + TBillRec.OptRevTax1
          ElseIf TBillRec.OptRevTax3 > 0 Then
            TotOpt3 = TotOpt3 + TBillRec.OptRevTax1
          ElseIf TBillRec.LateTaxDue > 0 Then
            TotalLate# = TotalLate# + TBillRec.OptRevTax1
          End If
        End If
        TotalBills = TotalBills + TBillRec.TotalBillDue
        GoTo ChilhowieSkip4
      End If
      
      TotalBills# = OldRound(TotalBills# + TBillRec.PersTaxNet + TBillRec.FETaxDue + TBillRec.MCTaxDue + TBillRec.MHTaxDue + TBillRec.MTTaxDue + TBillRec.LateTaxDue) ' - TBillRec.PPTRADiscnt) 'changed 11/6/06
      TotalBills# = OldRound(TotalBills# + TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3)
ChilhowieSkip4:
      TotalOverPay# = OldRound#(TotalOverPay# + TBillRec.OverPayAmt)
      
    End If      'Test for Discovery Bills
NotThisOne:
  Next x
  
  Print #RptHandle, FF$
  GoSub PrintSummaryHeader
  Print #RptHandle, "Number Of Bills To Process: "; Tab(29); Using$("##,##0", NumBills); Tab(48); "Amount To Bill Before Overpay:"; Tab(80); Using$("$##,###,##0.00", TotalBills#)
  Print #RptHandle, "Total Personal Valuation: "; Tab(28); Using$("$##,###,##0.00", GTPersValue#); Tab(48); "Total Overpay Applied:"; Tab(80); Using$("$##,###,##0.00", TotalOverPay#)
  If PPTRAYN = True Then
    Print #RptHandle, "Total PPTRA Valuation:"; Tab(28); Using$("$##,###,##0.00", GTPPTRAVal#); Tab(48); "Total Current Amount To Bill:"; Tab(80); Using$("$##,###,##0.00", OldRound(TotalBills# - TotalOverPay))
    Print #RptHandle, "Net Personal Valuation:"; Tab(28); Using$("$##,###,##0.00", OldRound(GTPersValue# - GTPPTRAVal#)); Tab(48); "Total Past Amount To Bill:"; Tab(80); Using$("$##,###,##0.00", TotalPast#)
  End If
  Print #RptHandle,
  Print #RptHandle, "Total Farm Eq Valuation:"; Tab(28); Using$("$##,###,##0.00", GTFEValue#); Tab(48); "Grand Total Amount To Bill:"; Tab(80); Using$("$##,###,##0.00", OldRound#(TotalPast# + TotalBills# - TotalOverPay)) 'added -TotalOverPay on 8/15/06
  Print #RptHandle, "Total Mobile Hm Valuation:"; Tab(28); Using$("$##,###,##0.00", GTMHValue#); Tab(48); "Inactive Skipped:"; Tab(89); Using$("#,##0", Inactive)
  Print #RptHandle, "Total Merch Cap Valuation:"; Tab(28); Using$("$##,###,##0.00", GTMCValue#)
  Print #RptHandle, "Total Mach/Tools Valuation:"; Tab(28); Using$("$##,###,##0.00", GTMTValue#)
  Print #RptHandle,
  Print #RptHandle, "Total Personal Taxes:"; Tab(28); Using$("$##,###,##0.00", GTPersTax#)
  If InStr(TaxMasterRec.Name, "CHILHOWIE") Then 'added 11/3/06
    Print #RptHandle, "Chilhowie " + QPTrim$(Using$("$##.00", TaxMasterRec.MinBill)) + " Min Charge:"; Tab(28); Using$("$##,###,##0.00", ChilhowieGFudge)
  End If
  If PPTRAYN = True Then
    Print #RptHandle, "Total PPTRA Disc Amt:"; Tab(28); Using$("$##,###,##0.00", GTPPTRADisc#)
    Print #RptHandle, "Net Personal Taxes:"; Tab(28); Using$("$##,###,##0.00", GTPersTaxNet#)
  End If
  Print #RptHandle,
  Print #RptHandle, "Total Farm Equip Taxes:"; Tab(28); Using$("$##,###,##0.00", GTFETax#)
  Print #RptHandle, "Total Mobile Hm Taxes:"; Tab(28); Using$("$##,###,##0.00", GTMHTax#)
  Print #RptHandle, "Total Merch Cap Taxes:"; Tab(28); Using$("$##,###,##0.00", GTMCTax#)
  Print #RptHandle, "Total Mach/Tools Taxes:"; Tab(28); Using$("$##,###,##0.00", GTMTTax#)
  Print #RptHandle, "Total Late List Taxes:"; Tab(28); Using$("$##,###,##0.00", TotalLate#)
  If OptRev1Tot# > 0 Then
    Print #RptHandle, "Total " + QPTrim$(OptRev1Desc) + ":"; Tab(28); Using$("$##,###,##0.00", OptRev1Tot#)
  End If
  If OptRev2Tot# > 0 Then
    Print #RptHandle, "Total " + QPTrim$(OptRev2Desc) + ":"; Tab(28); Using$("$##,###,##0.00", OptRev2Tot#)
  End If
  If OptRev3Tot# > 0 Then
    Print #RptHandle, "Total " + QPTrim$(OptRev3Desc) + ":"; Tab(28); Using$("$##,###,##0.00", OptRev3Tot#)
  End If
  Close
  ViewPrint TaxPreRptFile$, "Personal Property Tax Billing", True
  KillFile "txpblsprn.dat"
  MainLog ("Prebilling text report for personal property generated.")
  
  Exit Sub
  
SetCustInfo:
  TBillRec.CustRec = ThisTaxRec
  CustName$ = QPTrim$(TaxCust.CustName)
  TBillRec.CustName = CustName$
  TBillRec.CustAdd1 = QPTrim$(TaxCust.Addr1)
  TBillRec.CustAdd2 = QPTrim$(TaxCust.Addr2)
  CitySt$ = QPTrim$(TaxCust.City) + " " + TaxCust.State
  TBillRec.CustAdd3 = CitySt$
  TBillRec.CustZip = TaxCust.Zip
  TBillRec.CustPin = TaxCust.PIN
  TBillRec.TaxYear = WhatPYear
  TBillRec.RDesc3 = TaxCust.CSSN

  'Set Prior Balance if any
  ABalance = 0
  ABalance = GetCustPersBalance(TaxCust.Acct, -1) 'added this line on 8/11/2006    0
'  GoSub GetPastBalance
'  If ABalance# > 0 Then '8/14/06
    If PastFlagSet = 0 Then
      TBillRec.PriorYrBalance = ABalance#
    End If
    PastFlagSet = 1
'  End If '8/14/06
  Return
  
GetPastBalance:
  
  Balance# = 0
  ABalance# = 0
  
  If TaxCust.LastTrans > 0 Then
    OpenTaxTransFile TTHandle, NumOfTTRecs
    TransRecord& = TaxCust.LastTrans
    Do While TransRecord& <> 0
      Get TTHandle, TransRecord&, TaxTrans
      If TaxTrans.TranType = 1 And TaxTrans.BillType = "P" Then
        Balance# = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
        Balance# = OldRound(Balance# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection)
        Balance# = OldRound(Balance# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3) 'added for vs 2.05
        Balance# = OldRound(Balance# + TaxTrans.Revenue.PrePaidAmt) 'added for vs 2.05
        Balance# = OldRound(Balance# - (TaxTrans.Revenue.LateListPd + TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
        Balance# = OldRound(Balance# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd))
        Balance# = OldRound(Balance# - TaxTrans.DiscAmt) 'added for vs 2.05
        Balance# = OldRound(Balance# - (TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
        Balance# = OldRound#(Balance#)
      End If
      ABalance# = ABalance# + Balance#
      Balance# = 0
      TransRecord& = TaxTrans.LastTrans
    Loop
    Close TTHandle
  End If

Return
  
PPTRA:
  If WhatPYear$ = "1998" Then
    PERC! = 12.5
  ElseIf WhatPYear$ = "1999" Then
    PERC! = 27.5
  ElseIf WhatPYear$ = "2000" Then
    PERC! = 47.5
  ElseIf WhatPYear$ = "2001" Or WhatPYear$ = "2002" Then
    PERC! = 70
  Else
    PERC! = TaxMasterRec.PPTRADisc
  End If

Return

WriteIt2Disk:   'write the info out to disk here.
  TBillRec.BillPrinted = False
  TBillRec.SetDscvry2No = "N" 'added 12/4/06
  If TBillRec.TotalBillDue > 0 Then
    TBillRec.BillNumber = BillNum&
  Else
    TBillRec.BillNumber = -1
  End If
  If UsingMinTax = 1 Then
    If TBillRec.TotalBillDue <= MinBill And TBillRec.BillNumber <> -1 Then '8/17/06 added billnumber = -1
      TempAddOn.OldAmt = TBillRec.TotalBillDue
      TempAddOn.CustName = QPTrim$(TaxCust.CustName)
      TempAddOn.CustRec = x
      TempAddOn.Type = "Tax bills less than or equal to " + QPTrim$(Using$("$#,##0.00", MinBill)) + " become zero."
      AddOnCnt = AddOnCnt + 1
      TBillRec.TotalBillDue = 0
      TBillRec.SetDscvry2No = "Y" 'added 12/4/06
      GTPersTax# = OldRound(GTPersTax# - TBillRec.PersTaxDue)
      TBillRec.PersTaxDue = 0 'added 12/4/06
      GTPersTaxNet# = OldRound(GTPersTaxNet# - TBillRec.PersTaxNet) 'added 12/4/06
      TBillRec.PersTaxNet = 0 'added 12/4/06
      GTMCTax# = OldRound(GTMCTax# - TBillRec.MCTaxDue#) 'added 12/4/06
      TBillRec.MCTaxDue = 0 'added 12/4/06
      GTMHTax# = OldRound(GTMHTax# - TBillRec.MHTaxDue#) 'added 12/4/06
      TBillRec.MHTaxDue = 0 'added 12/4/06
      GTMTTax# = OldRound(GTMTTax# - TBillRec.MTTaxDue#) 'added 12/4/06
      TBillRec.MTTaxDue = 0 'added 12/4/06
      GTFETax# = OldRound(GTFETax# - TBillRec.FETaxDue#) 'added 12/4/06
      TBillRec.FETaxDue = 0 'added 12/4/06
      TotOpt1# = OldRound(TotOpt1# - TBillRec.OptRevTax1) 'added 12/4/06
      TBillRec.OptRevTax1 = 0 'added 12/4/06
      TotOpt2# = OldRound(TotOpt2# - TBillRec.OptRevTax2) 'added 12/4/06
      TBillRec.OptRevTax2 = 0 'added 12/4/06
      TotOpt3# = OldRound(TotOpt3# - TBillRec.OptRevTax3) 'added 12/4/06
      TBillRec.OptRevTax3 = 0 'added 12/4/06
      TotalLate# = OldRound(TotalLate# - TBillRec.LateTaxDue) 'added 12/4/06
      TBillRec.LateTaxDue = 0 'added 12/4/06
      GTPPTRADisc# = OldRound(GTPPTRADisc# - TBillRec.PPTRADiscnt) 'added 12/4/06
      TBillRec.PPTRADiscnt = 0 'added 12/4/06
      TempAddOn.NewAmt = 0
      Put AOHandle, AddOnCnt, TempAddOn
    End If
  ElseIf UsingMinTax = 2 And InStr(TaxMasterRec.Name, "CHILHOWIE") = 0 Then 'added Chilhowie on 11/3/06
    If TBillRec.TotalBillDue < MinBill And TBillRec.BillNumber <> -1 Then '8/17/06 added billnumber = -1
      TempAddOn.OldAmt = TBillRec.TotalBillDue
      TempAddOn.CustName = QPTrim$(TaxCust.CustName)
      TempAddOn.CustRec = x
      TempAddOn.Type = "Tax bills less than " + QPTrim$(Using$("$#,##0.00", MinBill)) + " become " + QPTrim$(Using$("$#,##0.00", MinBill)) + "."
      AddOnCnt = AddOnCnt + 1
      TempAddOn.NewAmt = MinBill
      Put AOHandle, AddOnCnt, TempAddOn
      Pct1# = 0
      Pct2# = 0
      Pct3# = 0
      Pct4# = 0
      Pct5# = 0
      Pct6# = 0
      Pct7# = 0
      Pct8# = 0
      Pct9# = 0
      PctTot# = OldRound(TBillRec.PersTaxNet + TBillRec.FETaxDue + TBillRec.MCTaxDue + TBillRec.MHTaxDue + TBillRec.MTTaxDue)
      PctTot# = OldRound(PctTot# + TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3 + TBillRec.LateTaxDue)
      If TBillRec.PersTaxNet > 0 Then
        Pct1# = OldRound#(TBillRec.PersTaxNet / PctTot#)
      End If
      If TBillRec.FETaxDue > 0 Then
        Pct2# = OldRound#(TBillRec.FETaxDue / PctTot#)
      End If
      If TBillRec.MCTaxDue > 0 Then
        Pct3# = OldRound#(TBillRec.MCTaxDue / PctTot#)
      End If
      If TBillRec.MHTaxDue > 0 Then
        Pct4# = OldRound#(TBillRec.MHTaxDue / PctTot#)
      End If
      If TBillRec.MTTaxDue > 0 Then
        Pct5# = OldRound#(TBillRec.MTTaxDue / PctTot#)
      End If
      If TBillRec.OptRevTax1 > 0 Then
        Pct6# = OldRound#(TBillRec.OptRevTax1 / PctTot#)
      End If
      If TBillRec.OptRevTax2 > 0 Then '
        Pct7# = OldRound#(TBillRec.OptRevTax2 / PctTot#)
      End If
      If TBillRec.OptRevTax3 > 0 Then
        Pct8# = OldRound#(TBillRec.OptRevTax3 / PctTot#)
      End If
      If TBillRec.LateTaxDue > 0 Then
        Pct9# = OldRound#(TBillRec.LateTaxDue / PctTot#)
      End If
      
      GTPersTaxNet = OldRound(GTPersTaxNet - TBillRec.PersTaxNet)
      TBillRec.PersTaxNet = MinBill * Pct1
      GTPersTaxNet = OldRound(GTPersTaxNet + TBillRec.PersTaxNet)
      
      'new code could go here
      '----------------------------------------------------
      If PPTRADisc# = 0 Then
        GTPersTax# = OldRound(GTPersTax# - TBillRec.PersTaxDue)
        TBillRec.PersTaxDue = TBillRec.PersTaxNet
        GTPersTax# = OldRound(GTPersTax# + TBillRec.PersTaxDue)
      Else
        GTPPTRADisc# = OldRound(GTPPTRADisc# - TBillRec.PPTRADiscnt)
        GTPersTax# = OldRound(GTPersTax# - TBillRec.PersTaxDue)
        TBillRec.PersTaxDue = TBillRec.PersTaxNet + PPTRADisc#
        TBillRec.PPTRADiscnt = PPTRADisc# 'TBillRec.PersTaxDue - TBillRec.PersTaxNet
        GTPPTRADisc# = OldRound(GTPPTRADisc# + TBillRec.PPTRADiscnt)
        GTPersTax# = OldRound(GTPersTax# + TBillRec.PersTaxDue)
      End If
      '----------------------------------------------------
      
      GTFETax# = OldRound(GTFETax# - TBillRec.FETaxDue)
      FETax# = OldRound(FETax# - TBillRec.FETaxDue)
      TBillRec.FETaxDue = OldRound(MinBill * Pct2)
      GTFETax# = OldRound(GTFETax# + TBillRec.FETaxDue)
      FETax# = OldRound(FETax# + TBillRec.FETaxDue)
      
      GTMCTax# = OldRound(GTMCTax# - TBillRec.MCTaxDue)
      MCTax# = OldRound(MCTax# - TBillRec.MCTaxDue)
      TBillRec.MCTaxDue = OldRound(MinBill * Pct3)
      GTMCTax# = OldRound(GTMCTax# + TBillRec.MCTaxDue)
      MCTax# = OldRound(MCTax# + TBillRec.MCTaxDue)
      
      GTMHTax# = OldRound(GTMHTax# - TBillRec.MHTaxDue)
      MHTax# = OldRound(MHTax# - TBillRec.MHTaxDue)
      TBillRec.MHTaxDue = OldRound(MinBill * Pct4)
      GTMHTax# = OldRound(GTMHTax# + TBillRec.MHTaxDue)
      MHTax# = OldRound(MHTax# + TBillRec.MHTaxDue)
      
      GTMTTax# = OldRound(GTMTTax# - TBillRec.MTTaxDue)
      MTTax# = OldRound(MTTax# - TBillRec.MTTaxDue)
      TBillRec.MTTaxDue = OldRound(MinBill * Pct5)
      GTMTTax# = OldRound(GTMTTax# + TBillRec.MTTaxDue)
      MTTax# = OldRound(MTTax# + TBillRec.MTTaxDue)
      
      TotOpt1# = OldRound(TotOpt1# - TBillRec.OptRevTax1)
      OptRevTax1# = OldRound(OptRevTax1# - TBillRec.OptRevTax1)
      TBillRec.OptRevTax1 = OldRound(MinBill * Pct6)
      TotOpt1# = OldRound(TotOpt1# + TBillRec.OptRevTax1)
      OptRevTax1# = OldRound(OptRevTax1# + TBillRec.OptRevTax1)
      
      TotOpt2# = OldRound(TotOpt2# - TBillRec.OptRevTax2)
      OptRevTax2# = OldRound(OptRevTax2# - TBillRec.OptRevTax2)
      TBillRec.OptRevTax2 = OldRound(MinBill * Pct7)
      TotOpt2# = OldRound(TotOpt2# + TBillRec.OptRevTax2)
      OptRevTax2# = OldRound(OptRevTax2# + TBillRec.OptRevTax2)
      
      TotOpt3# = OldRound(TotOpt3# - TBillRec.OptRevTax3)
      OptRevTax3# = OldRound(OptRevTax3# - TBillRec.OptRevTax3)
      TBillRec.OptRevTax3 = OldRound(MinBill * Pct8)
      TotOpt3# = OldRound(TotOpt3# + TBillRec.OptRevTax3)
      OptRevTax3# = OldRound(OptRevTax3# + TBillRec.OptRevTax3)
      
      TotalLate# = OldRound(TotalLate# - TBillRec.LateTaxDue)
      LateAmt# = OldRound(LateAmt# - TBillRec.LateTaxDue)
      TBillRec.LateTaxDue = OldRound(MinBill * Pct9)
      TotalLate# = OldRound(TotalLate# + TBillRec.LateTaxDue)
      LateAmt# = OldRound(LateAmt# + TBillRec.LateTaxDue)
      
      PctTest# = OldRound(TBillRec.PersTaxNet + TBillRec.FETaxDue + TBillRec.MCTaxDue + TBillRec.MHTaxDue + TBillRec.MTTaxDue)
      PctTest# = OldRound(PctTest# + TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3 + TBillRec.LateTaxDue)
      
      If MinBill < PctTest# Then
        If TBillRec.PersTaxNet > PctTest# - MinBill Then
          TBillRec.PersTaxNet = OldRound(TBillRec.PersTaxNet - (PctTest# - MinBill))
          TBillRec.PersTaxDue = OldRound(TBillRec.PersTaxDue - (PctTest# - MinBill))
          GTPersTax# = OldRound(GTPersTax# - (PctTest# - MinBill))
          GTPersTaxNet# = OldRound(GTPersTaxNet# - (PctTest# - MinBill))
        ElseIf TBillRec.FETaxDue > PctTest# - MinBill Then
          TBillRec.FETaxDue = OldRound(TBillRec.FETaxDue - (PctTest# - MinBill))
          GTFETax# = OldRound(GTFETax# - (PctTest# - MinBill))
          FETax# = OldRound(FETax# - (PctTest# - MinBill))
        ElseIf TBillRec.MCTaxDue > PctTest# - MinBill Then
          TBillRec.MCTaxDue = OldRound(TBillRec.MCTaxDue - (PctTest# - MinBill))
          GTMCTax# = OldRound(GTMCTax# - (PctTest# - MinBill))
          MCTax# = OldRound(MCTax# - (PctTest# - MinBill))
        ElseIf TBillRec.MHTaxDue > PctTest# - MinBill Then
          TBillRec.MHTaxDue = OldRound(TBillRec.MHTaxDue - (PctTest# - MinBill))
          GTMHTax# = OldRound(GTMHTax# - (PctTest# - MinBill))
          MHTax# = OldRound(MHTax# - (PctTest# - MinBill))
        ElseIf TBillRec.MTTaxDue > PctTest# - MinBill Then
          TBillRec.MTTaxDue = OldRound(TBillRec.MTTaxDue - (PctTest# - MinBill))
          GTMTTax# = OldRound(GTMTTax# - (PctTest# - MinBill))
          MTTax# = OldRound(MTTax# - (PctTest# - MinBill))
        ElseIf TBillRec.OptRevTax1 > PctTest# - MinBill Then
          TBillRec.OptRevTax1 = OldRound(TBillRec.OptRevTax1 - (PctTest# - MinBill))
          TotOpt1# = OldRound(TotOpt1# - (PctTest# - MinBill))
          OptRevTax1# = OldRound(OptRevTax1# - (PctTest# - MinBill))
        ElseIf TBillRec.OptRevTax2 > PctTest# - MinBill Then
          TBillRec.OptRevTax2 = OldRound(TBillRec.OptRevTax2 - (PctTest# - MinBill))
          TotOpt2# = OldRound(TotOpt2# - (PctTest# - MinBill))
          OptRevTax2# = OldRound(OptRevTax2# - (PctTest# - MinBill))
        ElseIf TBillRec.OptRevTax3 > PctTest# - MinBill Then
          TBillRec.OptRevTax3 = OldRound(TBillRec.OptRevTax3 - (PctTest# - MinBill))
          TotOpt3# = OldRound(TotOpt3# - (PctTest# - MinBill))
          OptRevTax3# = OldRound(OptRevTax3# - (PctTest# - MinBill))
        ElseIf TBillRec.LateTaxDue > PctTest# - MinBill Then
          TBillRec.LateTaxDue = OldRound(TBillRec.LateTaxDue - (PctTest# - MinBill))
          TotalLate# = OldRound(TotalLate# - (PctTest# - MinBill))
          LateAmt# = OldRound(LateAmt# - (PctTest# - MinBill))
        End If
      ElseIf MinBill > PctTest# Then
        If TBillRec.PersTaxNet > MinBill - PctTest# Then
          TBillRec.PersTaxNet = OldRound(TBillRec.PersTaxNet + (MinBill - PctTest#))
          TBillRec.PersTaxDue = OldRound(TBillRec.PersTaxDue + (MinBill - PctTest#))
          GTPersTax# = GTPersTax# + MinBill - PctTest#
          GTPersTaxNet# = GTPersTaxNet# + MinBill - PctTest#
        ElseIf TBillRec.FETaxDue > MinBill - PctTest# Then
          TBillRec.FETaxDue = OldRound(TBillRec.FETaxDue + (MinBill - PctTest#))
          GTFETax# = OldRound(GTFETax# + (PctTest# - MinBill))
          FETax# = OldRound(FETax# + (PctTest# - MinBill))
        ElseIf TBillRec.MCTaxDue > MinBill - PctTest# Then
          TBillRec.MCTaxDue = OldRound(TBillRec.MCTaxDue + (MinBill - PctTest#))
          GTMCTax# = OldRound(GTMCTax# + (PctTest# - MinBill))
          MCTax# = OldRound(MCTax# + (PctTest# - MinBill))
        ElseIf TBillRec.MHTaxDue > MinBill - PctTest# Then
          TBillRec.MHTaxDue = OldRound(TBillRec.MHTaxDue + (MinBill - PctTest#))
          GTMHTax# = OldRound(GTMHTax# + (PctTest# - MinBill))
          MHTax# = OldRound(MHTax# + (PctTest# - MinBill))
        ElseIf TBillRec.MTTaxDue > MinBill - PctTest# Then
          TBillRec.MTTaxDue = OldRound(TBillRec.MTTaxDue + (MinBill - PctTest#))
          GTMTTax# = OldRound(GTMTTax# + (PctTest# - MinBill))
          MTTax# = OldRound(MTTax# + (PctTest# - MinBill))
        ElseIf TBillRec.OptRevTax1 > MinBill - PctTest# Then
          TBillRec.OptRevTax1 = OldRound(TBillRec.OptRevTax1 + (MinBill - PctTest#))
          TotOpt1# = OldRound(TotOpt1# + (PctTest# - MinBill))
          OptRevTax1# = OldRound(OptRevTax1# + (PctTest# - MinBill))
        ElseIf TBillRec.OptRevTax2 > MinBill - PctTest# Then
          TBillRec.OptRevTax2 = OldRound(TBillRec.OptRevTax2 + (MinBill - PctTest#))
          TotOpt2# = OldRound(TotOpt2# + (PctTest# - MinBill))
          OptRevTax2# = OldRound(OptRevTax2# + (PctTest# - MinBill))
        ElseIf TBillRec.OptRevTax3 > MinBill - PctTest# Then
          TBillRec.OptRevTax3 = OldRound(TBillRec.OptRevTax3 + (MinBill - PctTest#))
          TotOpt3# = OldRound(TotOpt3# + (PctTest# - MinBill))
          OptRevTax3# = OldRound(OptRevTax3# + (PctTest# - MinBill))
        ElseIf TBillRec.LateTaxDue > MinBill - PctTest# Then
          TBillRec.LateTaxDue = OldRound(TBillRec.LateTaxDue + (MinBill - PctTest#))
          TotalLate# = OldRound(TotalLate# + (MinBill - PctTest#))
          LateAmt# = OldRound(LateAmt# + (MinBill - PctTest#))
        End If
      End If
      TBillRec.TotalBillDue = MinBill
    End If
  End If
  
  ThisTBCnt = ThisTBCnt + 1
  TBillRec.OptRevDesc1 = OptRev1Desc
  TBillRec.OptRevDesc2 = OptRev2Desc
  TBillRec.OptRevDesc3 = OptRev3Desc
  Put TBHandle, ThisTBCnt, TBillRec
  ThisTBCnt = LOF(TBHandle) / Len(TBillRec)
  If TBillRec.TotalBillDue > 0 Then
    If Abs(OverPayAmt) > 0 Then
      OPBillRec.Revenue.LateListPd = 0
      OPBillRec.Revenue.Principle1Pd = 0
      OPBillRec.Revenue.Principle2Pd = 0
      OPBillRec.Revenue.Principle3Pd = 0
      OPBillRec.Revenue.Principle4Pd = 0
      OPBillRec.Revenue.Principle5Pd = 0
      OPBillRec.Revenue.RevOpt1Pd = 0
      OPBillRec.Revenue.RevOpt2Pd = 0
      OPBillRec.Revenue.RevOpt3Pd = 0
      OpenPersTaxBillOverPayFile OPHandle, NumOfOPRecs
      Get TBHandle, ThisTBCnt, TBillRec
      OPBillRec.Revenue.PrePaidAmt = Abs(OverPayAmt)
      OPBillRec.BelongTo = ThisTBCnt 'BillNum&'9/9/05 Billnum is assigned
      'at bill printing so using this as a way of reference needed when posting
      If TBillRec.LateTaxDue > 0 And Abs(OverPayAmt) > 0 Then
        If Abs(OverPayAmt) > TBillRec.LateTaxDue Then
          OverPayAmt = OldRound(OverPayAmt + TBillRec.LateTaxDue)
          OPBillRec.Revenue.LateListPd = TBillRec.LateTaxDue
          OPApplied = OldRound(OPApplied + OPBillRec.Revenue.LateListPd)
        ElseIf Abs(OverPayAmt) <= TBillRec.LateTaxDue Then
          OPBillRec.Revenue.LateListPd = -OverPayAmt
          OPApplied = OldRound(OPApplied - OverPayAmt)
          OverPayAmt = 0
        End If
      End If
      
      If OldRound(TBillRec.PersTaxDue - TBillRec.PPTRADiscnt) > 0 And Abs(OverPayAmt) > 0 Then
        If Abs(OverPayAmt) > OldRound(TBillRec.PersTaxDue - TBillRec.PPTRADiscnt) Then
          OverPayAmt = OldRound(OverPayAmt + (TBillRec.PersTaxDue - TBillRec.PPTRADiscnt))
          OPBillRec.Revenue.Principle1Pd = OldRound(TBillRec.PersTaxDue - TBillRec.PPTRADiscnt)
          OPApplied = OldRound(OPApplied + OPBillRec.Revenue.Principle1Pd)
        ElseIf Abs(OverPayAmt) <= OldRound(TBillRec.PersTaxDue - TBillRec.PPTRADiscnt) Then
          OPBillRec.Revenue.Principle1Pd = -OverPayAmt
          OPApplied = OldRound(OPApplied - OverPayAmt)
          OverPayAmt = 0
        End If
      End If
'      If TBillRec.PersTaxDue > 0 And Abs(OverPayAmt) > 0 Then
'        If Abs(OverPayAmt) > TBillRec.PersTaxDue Then
'          OverPayAmt = OldRound(OverPayAmt + TBillRec.PersTaxDue)
'          OPBillRec.Revenue.Principle1Pd = TBillRec.PersTaxDue
'          OPApplied = OldRound(OPApplied + OPBillRec.Revenue.Principle1Pd)
'        ElseIf Abs(OverPayAmt) <= TBillRec.PersTaxDue Then
'          OPBillRec.Revenue.Principle1Pd = -OverPayAmt
'          OPApplied = OldRound(OPApplied - OverPayAmt)
'          OverPayAmt = 0
'        End If
'      End If

      If TBillRec.MTTaxDue > 0 And Abs(OverPayAmt) > 0 Then
        If Abs(OverPayAmt) > TBillRec.MTTaxDue Then
          OverPayAmt = OldRound(OverPayAmt + TBillRec.MTTaxDue)
          OPBillRec.Revenue.Principle2Pd = TBillRec.MTTaxDue
          OPApplied = OldRound(OPApplied + TBillRec.MTTaxDue)
        ElseIf Abs(OverPayAmt) <= TBillRec.MTTaxDue Then
          OPBillRec.Revenue.Principle2Pd = -OverPayAmt
          OPApplied = OldRound(OPApplied - OverPayAmt)
          OverPayAmt = 0
        End If
      End If
      
      If TBillRec.MCTaxDue > 0 And Abs(OverPayAmt) > 0 Then
        If Abs(OverPayAmt) > TBillRec.MCTaxDue Then
          OverPayAmt = OldRound(OverPayAmt + TBillRec.MCTaxDue)
          OPBillRec.Revenue.Principle3Pd = TBillRec.MCTaxDue
          OPApplied = OldRound(OPApplied + TBillRec.MCTaxDue)
        ElseIf Abs(OverPayAmt) <= TBillRec.MCTaxDue Then
          OPBillRec.Revenue.Principle3Pd = -OverPayAmt
          OPApplied = OldRound(OPApplied - OverPayAmt)
          OverPayAmt = 0
        End If
      End If
      
      If TBillRec.FETaxDue > 0 And Abs(OverPayAmt) > 0 Then
        If Abs(OverPayAmt) > TBillRec.FETaxDue Then
          OverPayAmt = OldRound(OverPayAmt + TBillRec.FETaxDue)
          OPBillRec.Revenue.Principle4Pd = TBillRec.FETaxDue
          OPApplied = OldRound(OPApplied + TBillRec.FETaxDue)
        ElseIf Abs(OverPayAmt) <= TBillRec.FETaxDue Then
          OPBillRec.Revenue.Principle4Pd = -OverPayAmt
          OPApplied = OldRound(OPApplied - OverPayAmt)
          OverPayAmt = 0
        End If
      End If
      
      If TBillRec.MHTaxDue > 0 And Abs(OverPayAmt) > 0 Then
        If Abs(OverPayAmt) > TBillRec.MHTaxDue Then
          OverPayAmt = OldRound(OverPayAmt + TBillRec.MHTaxDue)
          OPBillRec.Revenue.Principle5Pd = TBillRec.MHTaxDue
          OPApplied = OldRound(OPApplied + TBillRec.MHTaxDue)
        ElseIf Abs(OverPayAmt) <= TBillRec.MHTaxDue Then
          OPBillRec.Revenue.Principle5Pd = -OverPayAmt
          OPApplied = OldRound(OPApplied - OverPayAmt)
          OverPayAmt = 0
        End If
      End If
      
      If TBillRec.OptRevTax1 > 0 And Abs(OverPayAmt) > 0 Then
        If Abs(OverPayAmt) > TBillRec.OptRevTax1 Then
          OverPayAmt = OldRound(OverPayAmt + TBillRec.OptRevTax1)
          OPBillRec.Revenue.RevOpt1Pd = TBillRec.OptRevTax1
          OPApplied = OldRound(OPApplied + TBillRec.OptRevTax1)
        ElseIf Abs(OverPayAmt) <= TBillRec.OptRevTax1 Then
          OPBillRec.Revenue.RevOpt1Pd = -OverPayAmt
          OPApplied = OldRound(OPApplied - OverPayAmt)
          OverPayAmt = 0
        End If
      End If
      
      If TBillRec.OptRevTax2 > 0 And Abs(OverPayAmt) > 0 Then
        If Abs(OverPayAmt) > TBillRec.OptRevTax2 Then
          OverPayAmt = OldRound(OverPayAmt + TBillRec.OptRevTax2)
          OPBillRec.Revenue.RevOpt2Pd = TBillRec.OptRevTax2
          OPApplied = OldRound(OPApplied + TBillRec.OptRevTax2)
        ElseIf Abs(OverPayAmt) <= TBillRec.OptRevTax2 Then
          OPBillRec.Revenue.RevOpt2Pd = -OverPayAmt
          OPApplied = OldRound(OPApplied - OverPayAmt)
          OverPayAmt = 0
        End If
      End If
      
      If TBillRec.OptRevTax3 > 0 And Abs(OverPayAmt) > 0 Then
        If Abs(OverPayAmt) > TBillRec.OptRevTax3 Then
          OverPayAmt = OldRound(OverPayAmt + TBillRec.OptRevTax3)
          OPBillRec.Revenue.RevOpt3Pd = TBillRec.OptRevTax3
          OPApplied = OldRound(OPApplied + TBillRec.OptRevTax3)
        ElseIf Abs(OverPayAmt) <= TBillRec.OptRevTax3 Then
          OPBillRec.Revenue.RevOpt3Pd = -OverPayAmt
          OPApplied = OldRound(OPApplied - OverPayAmt)
          OverPayAmt = 0
        End If
      End If
      
      OPBillRec.Amount = OPApplied
      Put OPHandle, NumOfOPRecs + 1, OPBillRec
      Close OPHandle
      Get TBHandle, ThisTBCnt, TBillRec
      TBillRec.OverPayAmt = OPApplied
      Put TBHandle, ThisTBCnt, TBillRec
    End If
      
    BillNum& = BillNum& + 1
  End If
  
  Return
  
GetPersInfo:
  LunenBurgYN = False
  PersValue# = 0
  PersTaxDue# = 0
  MHValue# = 0
  MCValue# = 0
  FEValue# = 0
  MTValue# = 0
  PropertyRec! = TaxCust.FirstPersRec
  PPTRAVal# = 0
  PPTRADisc# = 0
  TPersTaxDue# = 0
  TPPTRAVal# = 0
  TPPTRADisc# = 0
  OptRevTax1# = 0
  OptRevTax2# = 0
  OptRevTax3# = 0
  Dim ThisProp&
  Do While PropertyRec! > 0
    PPTRADisc# = 0
    Get PHandle, PropertyRec!, PersRec
    ThisProp = PropertyRec!
    ThisMaxVehVal = MaxVehVal
    ThisMinVehVal = MinVehVal
    If PersRec.Deleted = True Then GoTo KeepGoing
'    If PersRec.LastYrPrinted = WhatPYear Then Discovery$ = "Y" 'took out 12/5/06
    If SupOnly = "Y" And PersRec.DISCOV = "N" Then GoTo KeepGoing
    If MultiYear > 1 Or PersRec.DISCOV = "Y" Then GoTo GoAhead
    If (PersRec.LastYrPrinted <> WhatPYear) Then
GoAhead:
    PYear$ = CInt(PersRec.TaxBillYear)
    If Val(PYear$) > 0 Then
      PYearInt = Val(PYear$)
    Else
      GoTo KeepGoing
    End If
'    If MultiYear > 1 Then
'      If PYear <> WhatYear$ Then GoTo KeepGoing
'    Else
'      If PYear$ <> WhatYear Or PersRec.LastYrPrinted = WhatYear Then GoTo KeepGoing
'    End If
    If InStr(TaxMasterRec.Name, "CHILHOWIE") Then GoTo SkipOpt1

    OptRevTax1# = OldRound(FigureOptRevTax1(ThisProp&, PHandle, "P"))
    TBillRec.OptRevTax1 = OptRevTax1# + TBillRec.OptRevTax1
'    TotOpt1 = OldRound(TotOpt1 + OptRevTax1#)
SkipOpt1:
    OptRevTax2# = OldRound(FigureOptRevTax2(ThisProp&, PHandle, "P"))
    TBillRec.OptRevTax2 = OptRevTax2# + TBillRec.OptRevTax2
'    TotOpt2 = OldRound(TotOpt2 + OptRevTax2#)
    
    If InStr(TaxMasterRec.Name, "LUNENBURG") Then 'added 10/12/06
      OptRevTax3# = 0
      TBillRec.OptRevTax3 = OptRevTax3# + TBillRec.OptRevTax3
    Else
      OptRevTax3# = OldRound(FigureOptRevTax3(ThisProp&, PHandle, "P"))
      TBillRec.OptRevTax3 = OptRevTax3# + TBillRec.OptRevTax3
    End If
'    TotOpt3 = OldRound(TotOpt3 + OptRevTax3#)
    
    Factor! = 1
    TBillRec.Prorate = "N"
    Rate = PersRec.ProrateVal
    If Rate > 0 Then
      Factor! = Rate / 12
      TBillRec.Prorate = "Y"
    Else
      Factor = 1
    End If
'    If MultiYear <> 0 Then 'remmed out 6/22/06
'      PersRec.PersVal = OldRound(PersRec.PersVal / MultiYear)
'      PersRec.CVALUE = OldRound(PersRec.CVALUE / MultiYear)
'      PersRec.MHValue = OldRound(PersRec.MHValue / MultiYear)
'      PersRec.MCValue = OldRound(PersRec.MCValue / MultiYear)
'      PersRec.MTValue = OldRound(PersRec.MTValue / MultiYear)
'      ThisMaxVehVal = OldRound(ThisMaxVehVal / MultiYear)
'    End If
      
    PersValue# = OldRound#(PersRec.PersVal * Factor!) + PersValue
    FEValue# = OldRound#(PersRec.CVALUE) + FEValue#
    MHValue# = OldRound#(PersRec.MHValue) + MHValue#
    MCValue# = OldRound#(PersRec.MCValue) + MCValue#
    MTValue# = OldRound#(PersRec.MTValue) + MTValue#
      
    If ABalance# > 0 Then
      If LawChngDate > 0 Then
        If Date2Num(fptxtBillDate.Text) >= LawChngDate Then
          GoTo NoDisc 'new for 2006
        End If
      End If
    End If
    If PPTRAYN = False Then GoTo NoDisc
    If PersRec.PPTRAYN = "Y" Then
      If OldRound#(PersRec.PersVal * Factor!) > ThisMaxVehVal Then
        PPTRAVal# = ThisMaxVehVal
      Else
        PPTRAVal# = OldRound#(PersRec.PersVal * Factor!)
      End If
'      If MultiYear <> 0 Then'remmed out 6/22/06
'        ThisMinVehVal = OldRound(ThisMinVehVal / MultiYear)
'      End If
      If PPTRAVal# <= (ThisMinVehVal * Factor!) Then
        PPTRADisc# = OldRound(((PPTRAVal# / 100) * Factor) * PERSRATE#) '2/21/06
      Else
        PPTRADisc# = OldRound#(((PPTRAVal# / 100) * (PERC! / 100)) * PERSRATE#)
      End If
      TPPTRAVal# = OldRound(TPPTRAVal# + PPTRAVal#)
    End If
NoDisc:
    PersTaxDue = OldRound#(((PersRec.PersVal / 100) * Factor!) * PERSRATE#)
    If PersTaxDue < 0 Then PersTaxDue = 0
    End If
    TPersTaxDue = OldRound(TPersTaxDue + PersTaxDue)
    TPPTRADisc# = OldRound(PPTRADisc + TPPTRADisc#)
    If PersRec.OptRev3Chrg > 0 Then LunenBurgYN = True
  
KeepGoing:
    PropertyRec! = PersRec.NextRec
  Loop
  
  TBillRec.OptRevTax1 = TBillRec.OptRevTax1 / MultiYear
  TBillRec.OptRevTax2 = TBillRec.OptRevTax2 / MultiYear
  TBillRec.OptRevTax3 = TBillRec.OptRevTax3 / MultiYear

  TPersTaxDue = OldRound(TPersTaxDue / MultiYear) 'added on 6/22/06
  TPPTRADisc# = OldRound(TPPTRADisc / MultiYear) 'added on 6/22/06
  
  MHTax# = OldRound((MHValue# / 100) * MHRate#)
  MHTax# = OldRound(MHTax# / MultiYear) 'added on 6/22/06
  MCTax# = OldRound((MCValue# / 100) * MCRate#)
  MCTax# = OldRound(MCTax# / MultiYear) 'added on 6/22/06
  FETax# = OldRound((FEValue# / 100) * FERate#)
  FETax# = OldRound(FETax# / MultiYear) 'added on 6/22/06
  MTTax# = OldRound((MTValue# / 100) * MTRate#)
  MTTax# = OldRound(MTTax# / MultiYear) 'added on 6/22/06
'  If OldRound(TPersTaxDue + MHTax# + MCTax# + FETax# + MTTax# - TPPTRADisc#) = 0 Then
'    If OldRound(TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3) = 0 Then 'added 10/18/06
'      If fpcmbInclNoBills.Text = "N" Then Return
'    End If
'  End If
  If OldRound(TPersTaxDue + MHTax# + MCTax# + FETax# + MTTax# - TPPTRADisc#) = 0 Then
    If OldRound(TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3) = 0 Then 'added 10/18/06
      If fpcmbInclNoBills.Text = "N" Then
        If InStr(TaxMasterRec.Name, "LUNENBURG") Then
          If OldRound((TPersTaxDue) - (TBillRec.OptRevTax1 + TBillRec.OptRevTax2)) = 0 Then
            Return
          End If
        Else
          Return
        End If
      End If
    End If
  End If
  TBillRec.FETaxDue = FETax#
  TBillRec.FEValue = CDbl(fpDblSnglFERate.Value)
  TBillRec.FEValue = FEValue#
  TBillRec.MCTaxDue = MCTax#
  TBillRec.MCValue = CDbl(fpDblSnglMCRate.Value)
  TBillRec.MCValue = MCValue#
  TBillRec.MHTaxDue = MHTax#
  TBillRec.MHValue = CDbl(fpDblSnglMHRate.Value)
  TBillRec.MHValue = MHValue#
  TBillRec.MTTaxDue = MTTax#
  TBillRec.MTValue = CDbl(fpDblSnglMTRate.Value)
  TBillRec.MTValue = MTValue#
  GTMHTax# = OldRound(GTMHTax# + MHTax#)
  GTMCTax# = OldRound(GTMCTax# + MCTax#)
  GTFETax# = OldRound(GTFETax# + FETax#)
  GTMTTax# = OldRound(GTMTTax# + MTTax#)
  OptRev1Tot# = OldRound(OptRev1Tot# + TBillRec.OptRevTax1)
  OptRev2Tot# = OldRound(OptRev2Tot# + TBillRec.OptRevTax2)
  OptRev3Tot# = OldRound(OptRev3Tot# + TBillRec.OptRevTax3)
  GTPersTax# = OldRound(GTPersTax# + TPersTaxDue#)
  TBillRec.PersTaxNet = OldRound(TPersTaxDue# - TPPTRADisc#)
  TBillRec.MultiYrVal = MultiYear
  TBillRec.PersTaxDue = TPersTaxDue#
  TBillRec.PPTRADiscnt = TPPTRADisc
  GTPersTaxNet# = OldRound(GTPersTaxNet# + TBillRec.PersTaxNet)
  GTPTTRADisc# = OldRound(GTPTTRADisc# + TPPTRADisc#)
  TBillRec.PersValue = PersValue#
  GTPersValue# = OldRound(GTPersValue# + PersValue#)
  GTFEValue# = OldRound(GTFEValue# + TBillRec.FEValue#)
  GTMCValue# = OldRound(GTMCValue# + TBillRec.MCValue#)
  GTMHValue# = OldRound(GTMHValue# + TBillRec.MHValue#)
  GTMTValue# = OldRound(GTMTValue# + TBillRec.MTValue#)
  TBillRec.CustPin = TaxCust.PIN
  TBillRec.InternalPin = PersRec.InternalPin
  TBillRec.PersPin = QPTrim$(PersRec.PropPin)
  TBillRec.ExptValue = 0
  GTPPTRADisc = OldRound(GTPPTRADisc + TPPTRADisc#)
  TBillRec.PPTRAValue = TPPTRAVal#
  GTPPTRAVal = OldRound(GTPPTRAVal# + TPPTRAVal#)
  TBillRec.PersPropRecord = TaxCust.FirstPersRec
  TBillRec.PersTaxRate = PERSRATE#
  If InStr(TaxMasterRec.Name, "CHILHOWIE") Then 'added 11/3/06
    PersBalTest# = OldRound(TPersTaxDue + MHTax# + MCTax# + FETax# + MTTax# + TBillRec.OptRevTax2 + TBillRec.OptRevTax3)
    CalcDiff# = OldRound(TaxMasterRec.MinBill - PersBalTest#)
    If PersBalTest# > 0 And PersBalTest# < TaxMasterRec.MinBill Then
      TBillRec.ChillHowieFudge = CalcDiff#
      ChilhowieGFudge = OldRound(ChilhowieGFudge + TBillRec.ChillHowieFudge)
      If TBillRec.PPTRADiscnt > 0 And TBillRec.PPTRAValue <= TaxMasterRec.MinVehTaxVal Then
         TBillRec.PPTRADiscnt = PersBalTest# + CalcDiff#
      End If
    End If
    TBillRec.OptRevTax1 = TBillRec.ChillHowieFudge
    TotOpt1 = OldRound(TotOpt1 + TBillRec.OptRevTax1)
    TBillRec.TotalBillDue = OldRound(TBillRec.PersTaxDue + TBillRec.FETaxDue + TBillRec.MCTaxDue + TBillRec.MHTaxDue + TBillRec.MTTaxDue - TBillRec.PPTRADiscnt)
    TBillRec.TotalBillDue = OldRound(TBillRec.TotalBillDue + TBillRec.OptRevTax2 + TBillRec.OptRevTax3 + TBillRec.ChillHowieFudge)
    GoTo ChilhowieSkip3
  End If
  
  TBillRec.TotalBillDue = OldRound#(TBillRec.PersTaxDue + TBillRec.FETaxDue + TBillRec.MCTaxDue + TBillRec.MHTaxDue + TBillRec.MTTaxDue - TBillRec.PPTRADiscnt)
  TBillRec.TotalBillDue = OldRound(TBillRec.TotalBillDue + TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3)
ChilhowieSkip3:
  If PersRec.LateList = "Y" Then
    LateAmt# = OldRound#((LateAmt# + OldRound(TBillRec.PersTaxDue - TBillRec.PPTRADiscnt)) * (LateList# / 100))
    TBillRec.LateTaxDue = LateAmt#
    TBillRec.TotalBillDue = OldRound#(TBillRec.TotalBillDue + TBillRec.LateTaxDue)
  End If
  TBillRec.DueDate = Date2Num(fptxtDueDate.Text)
  TBillRec.RDesc1 = QPTrim$(PersRec.DESC1)
  Dim ThisTotBill As Double
  ThisTotBill = OldRound((TBillRec.TotalBillDue + TBillRec.PPTRADiscnt) - (TBillRec.OptRevTax1 + TBillRec.OptRevTax2))
  If InStr(TaxMasterRec.Name, "LUNENBURG") Then 'added 10/12/06
    If LunenBurgYN = True Then
      If ThisTotBill < 10 Then
        TBillRec.OptRevTax3 = ThisTotBill
      ElseIf ThisTotBill >= 10 And ThisTotBill <= 100 Then
        TBillRec.OptRevTax3 = 10
      Else
        TBillRec.OptRevTax3 = OldRound(ThisTotBill * 0.1)
      End If
      OptRev3Tot# = OldRound(OptRev3Tot# + TBillRec.OptRevTax3)
    End If
    TBillRec.TotalBillDue = OldRound(TBillRec.PersTaxDue + TBillRec.FETaxDue + TBillRec.MCTaxDue + TBillRec.MHTaxDue + TBillRec.MTTaxDue - TBillRec.PPTRADiscnt)
    TBillRec.TotalBillDue = OldRound(TBillRec.TotalBillDue + TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3)
  End If
 
  GoSub WriteIt2Disk

  Return

PreBillHeading:
  Page = Page + 1
  Print #RptHandle, Tab(20); "Personal Property Tax Billing : Pre-Billing Register"
  Print #RptHandle, ThisTown; Tab(73); "Page #"; CStr(Page)
  ThisRate = CDbl(fpDblSnglPersRate.Value)
  Print #RptHandle, "Date: "; CStr(Date); Tab(62); "Pers Tax Rate:         " + Using$("##0.0000", ThisRate) + "%"
  ThisRate = CDbl(fpDblSnglFERate.Value)
  Print #RptHandle, "TownShip: " + ThisTownship; Tab(62); "Farm Eq Tax Rate:      " + Using$("##0.0000", ThisRate) + "%"
  fpcmbCounty.Col = 1
  ThisRate = CDbl(fpDblSnglMHRate.Value)
  If fpcmbCounty.ColText = "" Then
    Print #RptHandle, "County: " + "N/A"; Tab(62); "Mobile Hm Tax Rate:    " + Using$("##0.0000", ThisRate) + "%"
  Else
    Print #RptHandle, "County: " + QPTrim$(fpcmbCounty.ColText); Tab(62); "Mobile Hm Tax Rate:    " + Using$("##0.0000", ThisRate) + "%"
  End If
  fpcmbCycle.Col = 1
  ThisRate = CDbl(fpDblSnglMCRate.Value)
  If fpcmbCycle.ColText = "" Then
    Print #RptHandle, "Cycle: " + "N/A"; Tab(62); "Merch Cap Tax Rate:    " + Using$("##0.0000", ThisRate) + "%"
  Else
    Print #RptHandle, "Cycle: " + QPTrim$(fpcmbCycle.ColText); Tab(62); "Merch Cap Tax Rate:    " + Using$("##0.0000", ThisRate) + "%"
  End If
  ThisRate = CDbl(fpDblSnglMTRate.Value)
  Print #RptHandle, "* = Tax Due Does Not Include Prior Year Balance"; Tab(62); "Machine/Tools Tax Rate:" + Using$("##0.0000", ThisRate) + "%"
  ThisRate = CDbl(fpDblSnglLateList.Value)
  If PPTRAYN = True Then
    Print #RptHandle, "Current PPTRA Disc: " + Using$("#0.0000", OldRound(PERC! / 100)) + "%"; Tab(62); "Late List Rate:        " + Using$("##0.0000", ThisRate) + "%"
  Else
    Print #RptHandle, Tab(62); "Late List Rate:        " + Using$("##0.0000", ThisRate) + "%"
  End If
  Print #RptHandle, "Multi Year Value: " + Using$("#0", MultiYear)
  Print #RptHandle,
  
  Print #RptHandle, "Acct #"; Tab(8); "Customer Name"; Tab(58); "Prior Yr Bal"; Tab(73); "Bill Seq#"; Tab(86); "*Tax Due"
  Print #RptHandle, "    Pers Value"; Tab(16); "Farm Eq Value"; Tab(32); "Mob Hm Value"; Tab(45); "Merch Cap Val"; Tab(60); "Mach/Tools Val"; Tab(75); "Late List Tax"
  Print #RptHandle, "      Pers Tax"; Tab(16); "  Farm Eq Tax"; Tab(32); "  Mob Hm Tax"; Tab(45); "Merch Cap Tax"; Tab(60); "Mach/Tools Tax"
  Print #RptHandle, String(93, "-")
  LineCnt = 14
  
  Return

PrintSummaryHeader:
  Page = Page + 1
  Print #RptHandle, Tab(20); "Personal Property Tax Billing : Pre-Billing Register"
  Print #RptHandle, Tab(40); "Summary"
  Print #RptHandle, ThisTown; Tab(73); "Page #"; CStr(Page)
  ThisRate = CDbl(fpDblSnglPersRate.Value)
  Print #RptHandle, "Date: "; CStr(Date); Tab(62); "Pers Tax Rate:         " + Using$("##0.0000", ThisRate) + "%"
  ThisRate = CDbl(fpDblSnglFERate.Value)
  Print #RptHandle, "TownShip: " + ThisTownship; Tab(62); "Farm Eq Tax Rate:      " + Using$("##0.0000", ThisRate) + "%"
  fpcmbCounty.Col = 1
  ThisRate = CDbl(fpDblSnglMHRate.Value)
  If fpcmbCounty.ColText = "" Then
    Print #RptHandle, "County: " + "N/A"; Tab(62); "Mobile Hm Tax Rate:    " + Using$("##0.0000", ThisRate) + "%"
  Else
    Print #RptHandle, "County: " + QPTrim$(fpcmbCounty.ColText); Tab(62); "Mobile Hm Tax Rate:    " + Using$("##0.0000", ThisRate) + "%"
  End If
  fpcmbCycle.Col = 1
  ThisRate = CDbl(fpDblSnglMCRate.Value)
  If fpcmbCycle.ColText = "" Then
    Print #RptHandle, "Cycle: " + "N/A"; Tab(62); "Merch Cap Tax Rate:    " + Using$("##0.0000", ThisRate) + "%"
  Else
    Print #RptHandle, "Cycle: " + QPTrim$(fpcmbCycle.ColText); Tab(62); "Merch Cap Tax Rate:    " + Using$("##0.0000", ThisRate) + "%"
  End If
  ThisRate = CDbl(fpDblSnglMTRate.Value)
  Print #RptHandle, Tab(62); "Machine/Tools Tax Rate:" + Using$("##0.0000", ThisRate) + "%"
  ThisRate = CDbl(fpDblSnglLateList.Value)
  If PPTRAYN = True Then
    Print #RptHandle, "Current PPTRA Disc: " + Using$("#0.0000", OldRound(PERC! / 100)) + "%"; Tab(62); "Late List Rate:        " + Using$("##0.0000", ThisRate) + "%"
  Else
    Print #RptHandle, Tab(62); "Late List Rate:        " + Using$("##0.0000", ThisRate) + "%"
  End If
  Print #RptHandle, "Multi Year Value: " + Using$("#0", MultiYear)
  
  Print #RptHandle, String(93, "-")
  
  LineCnt = 11
  
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPrebilling", "PrintGraphics", Erl)
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

Private Sub fpcmbSuppOnly_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbSuppOnly.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbSuppOnly.ListIndex = -1
  End If
  If fpcmbSuppOnly.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbInclNoBills.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
End Sub

'Code used to experiment with PrintPersGraphics and retaining total values of
'properties when multiyear calculations are used to allow property totals to
'print on bills

'  Dim TBillRec As VAPPTaxBillType
'  Dim TBHandle As Integer
'  Dim NumOfTBRecs As Long
'  Dim OPBillRec As TaxTransactionType
'  Dim OPHandle As Integer
'  Dim NumOfOPRecs As Long
'  Dim TaxCust As TaxCustType
'  Dim TCHandle As Integer
'  Dim NumOfTCRecs As Long
'  Dim x As Long
'  Dim NewTBillRec As VAPPTaxBillType
'  Dim LateAmt#
'  Dim Inactive As Integer
'  Dim PastFlagSet As Boolean
'  Dim CustName$, CitySt$, CustAcct&
'  Dim ABalance#
'  Dim Balance#
'  Dim TransRecord&
'  Dim TaxTrans As TaxTransactionType
'  Dim TTHandle As Integer
'  Dim NumOfTTRecs As Long
'  Dim TempAddOn As TempTaxBillAddOn
'  Dim AOHandle As Integer
'  Dim AddOnCnt As Integer
'  Dim PersValue#
'  Dim Discovery$
'  Dim PersTaxDue#
'  Dim NextPersRec&
'  Dim TotalPers#
'  Dim TotalOverPay#
'  Dim TotalEx#
'  Dim NumBills&
'  Dim RptHandle As Integer
'  Dim TaxPreRptFile As String
'  Dim TValue#
'  Dim dlm$
'  Dim TotalBills#
'  Dim TotalLate#
'  Dim TotalPast#
'  Dim ThisTownship$
'  Dim TaxMasterRec As TaxMasterType
'  Dim TMHandle As Integer
'  Dim ThisTaxRec As Long
'  Dim CountyNum As Long
'  Dim CountyName$
'  Dim CycleNum As Long
'  Dim CycleName$
'  Dim BillCnt As Long
'  Dim OverPay As Boolean
'  Dim OverPayAmt As Double
'  Dim OPApplied As Double
'  Dim ThisTBCnt As Long
'  Dim ThisTest$
'  Dim OptRec As OptCustIdxType
'  Dim OHandle As Integer
'  Dim NumOfORecs As Long
'  Dim OptFlag As Boolean
'  Dim PERSRATE#
'  Dim MHRate#
'  Dim MCRate#
'  Dim FERate#
'  Dim MTRate#
'  Dim PersTax#
'  Dim MHTax#
'  Dim MCTax#
'  Dim FETax#
'  Dim MTTax#
'  Dim GTPersTax#
'  Dim GTPersTaxNet#
'  Dim GTMHTax#
'  Dim GTMCTax#
'  Dim GTFETax#
'  Dim GTMTTax#
'  Dim MHValue#
'  Dim MCValue#
'  Dim FEValue#
'  Dim MTValue#
'  Dim GTPersValue#
'  Dim GTMHValue#
'  Dim GTMCValue#
'  Dim GTFEValue#
'  Dim GTMTValue#
'  Dim PropertyRec!
'  Dim PPTRAVal#
'  Dim GTPPTRADisc#
'  Dim GTPPTRAVal#
'  Dim PPTRADisc#
'  Dim PERC!
'  Dim PersRec As PersonalRecType
'  Dim PHandle As Integer
'  Dim NumOfPRecs As Long
'  Dim Factor!
'  Dim Rate As Double
'  Dim PYear$
'  Dim PYearInt As Integer
'  Dim TxFile As Integer
'  Dim Prorate$
'  Dim GTPTTRADisc#
'  Dim GTPTTRAVal#
'  Dim MultiYear As Integer
'  Dim ThisPPTRADisc As Double
'  Dim ThisCName As String * 36
'  Dim TPersTaxDue As Double
'  Dim TPPTRAVal As Double
'  Dim TPPTRADisc As Double
'  Dim ThisMaxVehVal As Double, ThisMinVehVal#
'  Dim OptRevTax1 As Double
'  Dim OptRevTax2 As Double
'  Dim OptRevTax3 As Double
'  Dim OptRev1Desc$
'  Dim OptRev2Desc$
'  Dim OptRev3Desc$
'  Dim TotOpt1 As Double
'  Dim TotOpt2 As Double
'  Dim TotOpt3 As Double
'  Dim TotReal As Double
'  Dim fullPersVal As Double
'  Dim fullFEVal As Double
'  Dim fullMHVal As Double
'  Dim fullMCVal As Double
'  Dim fullMTVal As Double
'  Dim fullPPTRAVal As Double
'  Dim TfullPPTRAVal As Double
'
'  On Error GoTo ERRORSTUFF
'
'  PERSRATE# = fpDblSnglPersRate
'  MHRate# = fpDblSnglMHRate
'  MCRate# = fpDblSnglMCRate
'  FERate# = fpDblSnglFERate
'  MTRate# = fpDblSnglMTRate
'
'  Prorate$ = "Y"
'  OpenTaxSetUpFile TMHandle
'  Get TMHandle, 1, TaxMasterRec
'  Close TMHandle
'
'  MultiYear = TaxMasterRec.MultiYear
'  TotOpt1 = 0
'  TotOpt2 = 0
'  TotOpt3 = 0
'  OptRev1Desc = QPTrim$(TaxMasterRec.POptRev1)
'  OptRev2Desc = QPTrim$(TaxMasterRec.POptRev2)
'  OptRev3Desc = QPTrim$(TaxMasterRec.POptRev3)
'
'  ThisPPTRADisc = TaxMasterRec.PPTRADisc
'
'  fpcmbCycle.Col = 1
'  CycleName = QPTrim$(fpcmbCycle.ColText)
'  If CycleName = "" Then CycleName = "N/A"
'  If fpcmbCycle.Enabled = True And CycleName <> "NO CYCLE" Then
'    fpcmbCycle.Col = 0
'    CycleNum = CLng(fpcmbCycle.ColText)
'  Else
'    CycleNum = -1
'  End If
'
'  fpcmbCounty.Col = 1
'  CountyName = QPTrim$(fpcmbCounty.ColText)
'  If CountyName = "" Then CountyName = "N/A"
'  If fpcmbCounty.Enabled = True And CountyName <> "ALL COUNTIES" Then
'    fpcmbCounty.Col = 0
'    CountyNum = CLng(fpcmbCounty.ColText)
'  Else
'    CountyNum = -1
'  End If
'
'  If QPTrim$(fpcmbTownships.Text) = "ALL TOWNSHIPS" Then
'    ThisTownship = "ALL"
'  Else
'    ThisTownship = QPTrim$(fpcmbTownships.Text)
'    ThisTownship = Mid(ThisTownship, 1, Len(ThisTownship) - 4)
'    ThisTownship = QPTrim$(ThisTownship)
'  End If
'  dlm$ = "~"
'  If Exist(PersTaxBillFile) Then
'    KillFile PersTaxBillFile
'  End If
'  If Exist(PersTaxBillOPFile) Then
'    KillFile PersTaxBillOPFile
'  End If
'  If Exist("TMPPERSBLADD.DAT") Then 'tax bill addon
'    KillFile "TMPPERSBLADD.DAT"
'  End If
'
'  AddOnCnt = 0
'  OpenTaxBillPersAddOn AOHandle
'  OpenPersPropFile PHandle, NumOfPRecs
'  OpenTaxCustFile TCHandle, NumOfTCRecs
'  OpenPersTaxBillFile TBHandle, NumOfTBRecs
'
'  Inactive = 0
'
'  frmVATaxShowPctComp.Label1 = "Creating Tax Personal Pre-Billing Register"
'  frmVATaxShowPctComp.Show , Me
'  frmVATaxShowPctComp.cmdCancel.Visible = False
'  cmdExit.Enabled = False
'  cmdProcess.Enabled = False
'  EnableCloseButton Me.hwnd, False
'  If UseOpt = "Y" Then
'    NumOfTCRecs = OptCnt
'  End If
'  If UseSS = "Y" Then
'    NumOfTCRecs = SSCnt
'  End If
'
'  GoSub PPTRA
'
'  For x = 1 To NumOfTCRecs
'    TBillRec = NewTBillRec
'    If UsingIdx = True Then
'      Get TCHandle, CustArr(x), TaxCust
'      ThisTaxRec = CustArr(x)
'    Else
'      Get TCHandle, x, TaxCust
'      ThisTaxRec = x
'    End If
'    If TaxCust.TaxExempt = "Y" Then
'      GoTo PreBillSkip
'    End If
'    If ThisTownship <> "ALL" Then
'      If QPTrim$(TaxCust.TownShip) <> ThisTownship Then
'        GoTo PreBillSkip
'      End If
'    End If
'    If CycleNum >= 0 Then
'      If TaxCust.Cycle <> CycleNum Then
'        GoTo PreBillSkip
'      End If
'    End If
'
'    If CountyNum >= 0 Then
'      If TaxCust.County4BillNum <> CountyNum Then
'        GoTo PreBillSkip
'      End If
'    End If
'
'    LateAmt# = 0
'    OPApplied = 0
'    If TaxCust.Deleted <> 0 Then
'      GoTo PreBillSkip:
'    End If
'
'    If QPTrim$(TaxCust.Active) <> "Y" Then
'      Inactive = Inactive + 1
'      GoTo PreBillSkip:
'    End If
'
'    PastFlagSet = 0             'Initialize Past Balance Flag
'    OverPayAmt = 0
'    If UsingIdx = True Then
'      OverPayAmt = GetCustPersBalance(CustArr(x), -1)
'      If OverPayAmt < 0 Then
'        OverPay = True
'      Else
'        OverPayAmt = 0
'        OverPay = False
'      End If
'    Else
'      OverPayAmt = GetCustPersBalance(x, -1)
'      If OverPayAmt < 0 Then
'        OverPay = True
'      Else
'        OverPayAmt = 0
'        OverPay = False
'      End If
'    End If
'    ThisTest = CStr(OverPayAmt)
'    If InStr(ThisTest, "E") Then OverPayAmt = 0
''    If TaxCust.Acct = 1236 Then Stop
'    If TaxCust.FirstPersRec <= 0 Then
'      GoSub SetCustInfo
'      GoSub WriteIt2Disk
'      GoTo PreBillSkip
'    End If
'
'    GoSub SetCustInfo
'    GoSub GetPersInfo
'
'PreBillSkip:
'    frmVATaxShowPctComp.ShowPctComp x, NumOfTCRecs
'    If frmVATaxShowPctComp.Out = True Then
'      Close
'      frmVATaxShowPctComp.Out = False
'      Unload frmVATaxShowPctComp
'      EnableCloseButton Me.hwnd, True
'      cmdExit.Enabled = True
'      cmdProcess.Enabled = True
'      Exit Sub
'    End If
'  Next x
'
'  Unload frmVATaxShowPctComp
'  EnableCloseButton Me.hwnd, True
'  cmdExit.Enabled = True
'  cmdProcess.Enabled = True
'
'  Close TCHandle
'  Close AOHandle
'  Close PHandle
'
'  If ThisTBCnt = 0 Then
'    Call TaxMsg(900, "Using the parameters selected there are no customers who qualify for a tax charge.")
'    Close
'    fptxtRCurrYear.SetFocus
'    Exit Sub
'  End If
'
'  TotalPers# = 0
'  TotalEx# = 0
'  NumBills& = 0
'
'  TaxPreRptFile$ = "TAXRPTS\TaxPersPreBill.RPT"
'
'  RptHandle = FreeFile
'  Open TaxPreRptFile For Output As #RptHandle
'
'  Close TBHandle
'  OpenPersTaxBillFile TBHandle, NumOfTBRecs
'
'  For x = 1 To NumOfTBRecs
'    Get TBHandle, x, TBillRec
'    If Discovery$ = "Y" And TBillRec.BillNumber = -1 Then
'       GoTo NotThisOne
'    Else
'      If TBillRec.TotalBillDue = 0 And TBillRec.PriorYrBalance = 0 And TBillRec.PPTRAValue = 0 Then
'        If fpcmbInclNoBills.Text = "N" Then
'          GoTo NotThisOne
'        End If
'      End If
'
'      '                        0
'      Print #RptHandle, TBillRec.CustRec; dlm;
'      If QPTrim$(TBillRec.Prorate) <> "Y" Then
'        '                            1
'        Print #RptHandle, QPTrim$(TBillRec.CustName); dlm;
'      Else
'        ThisCName = QPTrim$(TBillRec.CustName)
'        '
'        Print #RptHandle, ThisCName + "**Prorated**"; dlm;
'      End If
'
'      If TBillRec.PriorYrBalance > 0 Then
'        '                            2
'        Print #RptHandle, TBillRec.PriorYrBalance; dlm;
'      Else
'        '                 2
'        Print #RptHandle, 0; dlm;
'      End If
'      If TBillRec.BillNumber = -1 Then
'        '                   3
'        Print #RptHandle, "N/A"; dlm;
'      Else
'        '                          3
'        Print #RptHandle, TBillRec.BillNumber; dlm;
'      End If
'      '                           4
'      Print #RptHandle, OldRound(TBillRec.TotalBillDue); dlm;
'      TValue# = OldRound#(TBillRec.PersValue + TBillRec.FEValue + TBillRec.MCValue + TBillRec.MHValue + TBillRec.MTValue - TBillRec.ExptValue)
'
'      If TValue# < 0 Then TValue# = 0
'      '                    5
'      Print #RptHandle, TValue#; dlm;
'      '                           6
'      Print #RptHandle, TBillRec.LateTaxDue; dlm;
'      TotalLate# = OldRound#(TotalLate# + TBillRec.LateTaxDue)
'      TotalPast# = OldRound#(TotalPast# + TBillRec.PriorYrBalance)
'      TotalBills# = OldRound(TotalBills# + TBillRec.PersTaxNet + TBillRec.FETaxDue + TBillRec.MCTaxDue + TBillRec.MHTaxDue + TBillRec.MTTaxDue + TBillRec.LateTaxDue)
'      TotalBills# = OldRound(TotalBills# + TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3)
'      TotalOverPay# = OldRound#(TotalOverPay# + TBillRec.OverPayAmt)
'
'      If TBillRec.TotalBillDue > 0 Then
'        NumBills& = NumBills& + 1
'      End If
'      '                     7
'      Print #RptHandle, NumBills&; dlm;
'      '                     8
'      Print #RptHandle, TotalPers#; dlm;
'      '                     9
'      Print #RptHandle, TotalBills#; dlm;
'      '                     10
'      Print #RptHandle, TotalPast#; dlm;
'      '                                   11
'      Print #RptHandle, OldRound#(TotalPast# + TotalBills#); dlm;
'      '                     12
'      Print #RptHandle, TotalLate#; dlm;
'      '                    13
'      Print #RptHandle, Inactive; dlm;
'      '                     14
'      Print #RptHandle, ThisTown; dlm;
'      '                    15
'      Print #RptHandle, WhatPYear; dlm;
'      '                     16                17              18
'      Print #RptHandle, ThisTownship; dlm; CycleName; dlm; CycleNum; dlm;
'      '                     19                20
'      Print #RptHandle, CountyName$; dlm; CountyNum; dlm;
'      '                          21                                  22
'      Print #RptHandle, -TBillRec.OverPayAmt; dlm; OldRound(TotalBills# - TotalOverPay); dlm;
'      '                      23                              24                                   25
'      Print #RptHandle, TotalOverPay#; dlm; CDbl(fpDblSnglPersRate.Value) / 100; dlm; CDbl(fpDblSnglLateList.Value) / 100; dlm;
'      '                                    26                              27                                      28                                           29
'      Print #RptHandle, CDbl(fpDblSnglFERate.Value) / 100; dlm; CDbl(fpDblSnglMHRate.Value) / 100; dlm; CDbl(fpDblSnglMCRate.Value) / 100; dlm; CDbl(fpDblSnglMTRate.Value) / 100; dlm;
'      '                       30                         31                      32                      33                     34                       35
'      Print #RptHandle, TBillRec.PersTaxDue; dlm; TBillRec.MHTaxDue; dlm; TBillRec.MTTaxDue; dlm; TBillRec.MCTaxDue; dlm; TBillRec.FETaxDue; dlm; TBillRec.PersValue; dlm;
'      '                        36                     37                    38
'      Print #RptHandle, TBillRec.MHValue; dlm; TBillRec.MTValue; dlm; TBillRec.MCValue; dlm; TBillRec.FEValue; dlm;
'      '                    40               41             42             43             44               45                46               47               48              49
'      Print #RptHandle, GTPersTax#; dlm; GTMHTax#; dlm; GTMTTax#; dlm; GTMCTax#; dlm; GTFETax#; dlm; GTPersValue#; dlm; GTMHValue#; dlm; GTMTValue#; dlm; GTMCValue#; dlm; GTFEValue#; dlm;
'      '                         50                         51                     52                53                              54
'      Print #RptHandle, TBillRec.PPTRADiscnt; dlm; TBillRec.PPTRAValue; dlm; GTPPTRADisc#; dlm; GTPPTRAVal#; dlm; OldRound(TBillRec.PersTaxDue); dlm;
'      '                                    55                           56              57              58               59                              60
'      Print #RptHandle, OldRound(GTPersValue# - GTPPTRAVal#); dlm; GTPersTaxNet#; dlm; PERC!; dlm; MultiYear; dlm; OldRound(ThisPPTRADisc / 100); dlm; PPTRAYN; dlm;
'      '                     61                            62                        63
'      Print #RptHandle, TBillRec.OptRevTax1; dlm; TBillRec.OptRevTax2; dlm; TBillRec.OptRevTax3; dlm;
'       '                      64                65                 66
'      Print #RptHandle, OptRev1Desc$; dlm; OptRev2Desc$; dlm; OptRev3Desc$; dlm;
'      '                   67            68            69
'      Print #RptHandle, TotOpt1; dlm; TotOpt2; dlm; TotOpt3
'
'    End If      'Test for Discovery Bills
'NotThisOne:
'  Next x
'
'  Close
'  arVATaxPreBillPers.Show
'  frmVATaxLoadReport.Show
'
'  KillFile "txpblsprn.dat"
'  MainLog ("Prebilling graphics report for personal property generated.")
'  Exit Sub
'
'PPTRA:
'  If WhatPYear$ = "1998" Then
'    PERC! = 12.5
'  ElseIf WhatPYear$ = "1999" Then
'    PERC! = 27.5
'  ElseIf WhatPYear$ = "2000" Then
'    PERC! = 47.5
'  ElseIf WhatPYear$ = "2001" Or WhatPYear$ = "2002" Then
'    PERC! = 70
'  Else
'    PERC! = TaxMasterRec.PPTRADisc
'  End If
'
'Return
'
'SetCustInfo:
'  TBillRec.CustRec = ThisTaxRec
'  CustName$ = QPTrim$(TaxCust.CustName)
'  TBillRec.CustName = CustName$
'  TBillRec.CustAdd1 = QPTrim$(TaxCust.Addr1)
'  TBillRec.CustAdd2 = QPTrim$(TaxCust.Addr2)
'  CitySt$ = QPTrim$(TaxCust.City) + " " + TaxCust.State
'  TBillRec.CustAdd3 = CitySt$
'  TBillRec.CustZip = TaxCust.Zip
'  TBillRec.CustPin = TaxCust.PIN
'  TBillRec.TaxYear = WhatPYear
'  TBillRec.RDesc3 = TaxCust.CSSN
'
'  'Set Prior Balance if any
'  ABalance = 0
'  GoSub GetPastBalance
'  If ABalance# > 0 Then
'    If PastFlagSet = 0 Then
'      TBillRec.PriorYrBalance = ABalance#
'    End If
'    PastFlagSet = 1
'  End If
'  Return
'
'GetPastBalance:
'
'  Balance# = 0
'  ABalance# = 0
'
'  If TaxCust.LastTrans > 0 Then
'    OpenTaxTransFile TTHandle, NumOfTTRecs
'    TransRecord& = TaxCust.LastTrans
'    Do While TransRecord& <> 0
'      Get TTHandle, TransRecord&, TaxTrans
'      If TaxTrans.TranType = 1 And TaxTrans.BillType = "P" Then
'        Balance# = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
'        Balance# = OldRound(Balance# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection)
'        Balance# = OldRound(Balance# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3) 'added for vs 2.05
'        Balance# = OldRound(Balance# + TaxTrans.Revenue.PrePaidAmt) 'added for vs 2.05
'        Balance# = OldRound(Balance# - (TaxTrans.Revenue.LateListPd + TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
'        Balance# = OldRound(Balance# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd))
'        Balance# = OldRound(Balance# - TaxTrans.DiscAmt) 'added for vs 2.05
'        Balance# = OldRound(Balance# - (TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
'        Balance# = OldRound#(Balance#)
'      End If
'      ABalance# = ABalance# + Balance#
'      Balance# = 0
'      TransRecord& = TaxTrans.LastTrans
'    Loop
'    Close TTHandle
'  End If
'
'Return
'
'WriteIt2Disk:   'write the info out to disk here.
'
'  TBillRec.BillPrinted = False
'  If TBillRec.TotalBillDue > 0 Then
'    TBillRec.BillNumber = BillNum&
'  Else
'    TBillRec.BillNumber = -1
'  End If
'  If UsingMinTax = 1 Then
'    If TBillRec.TotalBillDue <= MinBill Then
'      TempAddOn.OldAmt = TBillRec.TotalBillDue
'      TempAddOn.CustName = QPTrim$(TaxCust.CustName)
'      TempAddOn.CustRec = x
'      TempAddOn.Type = "Tax bills less than or equal to " + QPTrim$(Using$("$#,##0.00", MinBill)) + " become zero."
'      AddOnCnt = AddOnCnt + 1
'      TBillRec.TotalBillDue = 0
'      TempAddOn.NewAmt = 0
'      Put AOHandle, AddOnCnt, TempAddOn
'    End If
'  ElseIf UsingMinTax = 2 Then
'    If TBillRec.TotalBillDue < MinBill Then
'      TempAddOn.OldAmt = TBillRec.TotalBillDue
'      TempAddOn.CustName = QPTrim$(TaxCust.CustName)
'      TempAddOn.CustRec = x
'      TempAddOn.Type = "Tax bills less than " + QPTrim$(Using$("$#,##0.00", MinBill)) + " become " + QPTrim$(Using$("$#,##0.00", MinBill)) + "."
'      AddOnCnt = AddOnCnt + 1
'      TempAddOn.NewAmt = MinBill
'      Put AOHandle, AddOnCnt, TempAddOn
'      TBillRec.TotalBillDue = MinBill
'    End If
'  End If
'
'  ThisTBCnt = ThisTBCnt + 1
'  TBillRec.OptRevDesc1 = OptRev1Desc
'  TBillRec.OptRevDesc2 = OptRev2Desc
'  TBillRec.OptRevDesc3 = OptRev3Desc
'  Put TBHandle, ThisTBCnt, TBillRec
'  ThisTBCnt = LOF(TBHandle) / Len(TBillRec)
'  If TBillRec.TotalBillDue > 0 Then
'    If Abs(OverPayAmt) > 0 Then
'      OPBillRec.Revenue.LateListPd = 0
'      OPBillRec.Revenue.Principle1Pd = 0
'      OPBillRec.Revenue.Principle2Pd = 0
'      OPBillRec.Revenue.Principle3Pd = 0
'      OPBillRec.Revenue.Principle4Pd = 0
'      OPBillRec.Revenue.Principle5Pd = 0
'      OPBillRec.Revenue.RevOpt1Pd = 0
'      OPBillRec.Revenue.RevOpt2Pd = 0
'      OPBillRec.Revenue.RevOpt3Pd = 0
'      OpenPersTaxBillOverPayFile OPHandle, NumOfOPRecs
'      Get TBHandle, ThisTBCnt, TBillRec
'      OPBillRec.Revenue.PrePaidAmt = Abs(OverPayAmt)
'      OPBillRec.BelongTo = ThisTBCnt 'BillNum&'9/9/05 Billnum is assigned
'      'at bill printing so using this as a way of reference needed when posting
'      If TBillRec.LateTaxDue > 0 And Abs(OverPayAmt) > 0 Then
'        If Abs(OverPayAmt) > TBillRec.LateTaxDue Then
'          OverPayAmt = OldRound(OverPayAmt + TBillRec.LateTaxDue)
'          OPBillRec.Revenue.LateListPd = TBillRec.LateTaxDue
'          OPApplied = OldRound(OPApplied + OPBillRec.Revenue.LateListPd)
'        ElseIf Abs(OverPayAmt) <= TBillRec.LateTaxDue Then
'          OPBillRec.Revenue.LateListPd = -OverPayAmt
'          OPApplied = OldRound(OPApplied - OverPayAmt)
'          OverPayAmt = 0
'        End If
'      End If
'
'      If OldRound(TBillRec.PersTaxDue - TBillRec.PPTRADiscnt) > 0 And Abs(OverPayAmt) > 0 Then
'        If Abs(OverPayAmt) > OldRound(TBillRec.PersTaxDue - TBillRec.PPTRADiscnt) Then
'          OverPayAmt = OldRound(OverPayAmt + (TBillRec.PersTaxDue - TBillRec.PPTRADiscnt))
'          OPBillRec.Revenue.Principle1Pd = OldRound(TBillRec.PersTaxDue - TBillRec.PPTRADiscnt)
'          OPApplied = OldRound(OPApplied + OPBillRec.Revenue.Principle1Pd)
'        ElseIf Abs(OverPayAmt) <= OldRound(TBillRec.PersTaxDue - TBillRec.PPTRADiscnt) Then
'          OPBillRec.Revenue.Principle1Pd = -OverPayAmt
'          OPApplied = OldRound(OPApplied - OverPayAmt)
'          OverPayAmt = 0
'        End If
'      End If
'
'      If TBillRec.MTTaxDue > 0 And Abs(OverPayAmt) > 0 Then
'        If Abs(OverPayAmt) > TBillRec.MTTaxDue Then
'          OverPayAmt = OldRound(OverPayAmt + TBillRec.MTTaxDue)
'          OPBillRec.Revenue.Principle2Pd = TBillRec.MTTaxDue
'          OPApplied = OldRound(OPApplied + TBillRec.MTTaxDue)
'        ElseIf Abs(OverPayAmt) <= TBillRec.MTTaxDue Then
'          OPBillRec.Revenue.Principle2Pd = -OverPayAmt
'          OPApplied = OldRound(OPApplied - OverPayAmt)
'          OverPayAmt = 0
'        End If
'      End If
'
'      If TBillRec.MCTaxDue > 0 And Abs(OverPayAmt) > 0 Then
'        If Abs(OverPayAmt) > TBillRec.MCTaxDue Then
'          OverPayAmt = OldRound(OverPayAmt + TBillRec.MCTaxDue)
'          OPBillRec.Revenue.Principle3Pd = TBillRec.MCTaxDue
'          OPApplied = OldRound(OPApplied + TBillRec.MCTaxDue)
'        ElseIf Abs(OverPayAmt) <= TBillRec.MCTaxDue Then
'          OPBillRec.Revenue.Principle3Pd = -OverPayAmt
'          OPApplied = OldRound(OPApplied - OverPayAmt)
'          OverPayAmt = 0
'        End If
'      End If
'
'      If TBillRec.FETaxDue > 0 And Abs(OverPayAmt) > 0 Then
'        If Abs(OverPayAmt) > TBillRec.FETaxDue Then
'          OverPayAmt = OldRound(OverPayAmt + TBillRec.FETaxDue)
'          OPBillRec.Revenue.Principle4Pd = TBillRec.FETaxDue
'          OPApplied = OldRound(OPApplied + TBillRec.FETaxDue)
'        ElseIf Abs(OverPayAmt) <= TBillRec.FETaxDue Then
'          OPBillRec.Revenue.Principle4Pd = -OverPayAmt
'          OPApplied = OldRound(OPApplied - OverPayAmt)
'          OverPayAmt = 0
'        End If
'      End If
'
'      If TBillRec.MHTaxDue > 0 And Abs(OverPayAmt) > 0 Then
'        If Abs(OverPayAmt) > TBillRec.MHTaxDue Then
'          OverPayAmt = OldRound(OverPayAmt + TBillRec.MHTaxDue)
'          OPBillRec.Revenue.Principle5Pd = TBillRec.MHTaxDue
'          OPApplied = OldRound(OPApplied + TBillRec.MHTaxDue)
'        ElseIf Abs(OverPayAmt) <= TBillRec.MHTaxDue Then
'          OPBillRec.Revenue.Principle5Pd = -OverPayAmt
'          OPApplied = OldRound(OPApplied - OverPayAmt)
'          OverPayAmt = 0
'        End If
'      End If
'
'      If TBillRec.OptRevTax1 > 0 And Abs(OverPayAmt) > 0 Then
'        If Abs(OverPayAmt) > TBillRec.OptRevTax1 Then
'          OverPayAmt = OldRound(OverPayAmt + TBillRec.OptRevTax1)
'          OPBillRec.Revenue.RevOpt1Pd = TBillRec.OptRevTax1
'          OPApplied = OldRound(OPApplied + TBillRec.OptRevTax1)
'        ElseIf Abs(OverPayAmt) <= TBillRec.OptRevTax1 Then
'          OPBillRec.Revenue.RevOpt1Pd = -OverPayAmt
'          OPApplied = OldRound(OPApplied - OverPayAmt)
'          OverPayAmt = 0
'        End If
'      End If
'
'      If TBillRec.OptRevTax2 > 0 And Abs(OverPayAmt) > 0 Then
'        If Abs(OverPayAmt) > TBillRec.OptRevTax2 Then
'          OverPayAmt = OldRound(OverPayAmt + TBillRec.OptRevTax2)
'          OPBillRec.Revenue.RevOpt2Pd = TBillRec.OptRevTax2
'          OPApplied = OldRound(OPApplied + TBillRec.OptRevTax2)
'        ElseIf Abs(OverPayAmt) <= TBillRec.OptRevTax2 Then
'          OPBillRec.Revenue.RevOpt2Pd = -OverPayAmt
'          OPApplied = OldRound(OPApplied - OverPayAmt)
'          OverPayAmt = 0
'        End If
'      End If
'
'      If TBillRec.OptRevTax3 > 0 And Abs(OverPayAmt) > 0 Then
'        If Abs(OverPayAmt) > TBillRec.OptRevTax3 Then
'          OverPayAmt = OldRound(OverPayAmt + TBillRec.OptRevTax3)
'          OPBillRec.Revenue.RevOpt3Pd = TBillRec.OptRevTax3
'          OPApplied = OldRound(OPApplied + TBillRec.OptRevTax3)
'        ElseIf Abs(OverPayAmt) <= TBillRec.OptRevTax3 Then
'          OPBillRec.Revenue.RevOpt3Pd = -OverPayAmt
'          OPApplied = OldRound(OPApplied - OverPayAmt)
'          OverPayAmt = 0
'        End If
'      End If
'
'      OPBillRec.Amount = OPApplied
'      Put OPHandle, NumOfOPRecs + 1, OPBillRec
'      Close OPHandle
'      Get TBHandle, ThisTBCnt, TBillRec
'      TBillRec.OverPayAmt = OPApplied
'      Put TBHandle, ThisTBCnt, TBillRec
'    End If
'
'    BillNum& = BillNum& + 1
'  End If
'
'  Return
'
'GetPersInfo:
'  PersValue# = 0
'  PersTaxDue# = 0
'  MHValue# = 0
'  MCValue# = 0
'  FEValue# = 0
'  MTValue# = 0
'  PropertyRec! = TaxCust.FirstPersRec
'  PPTRAVal# = 0
'  PPTRAVal# = 0
''  PPTRADisc# = 0
'  TPersTaxDue# = 0
'  TPPTRAVal# = 0
'  TPPTRADisc# = 0
'  OptRevTax1# = 0
'  OptRevTax2# = 0
'  OptRevTax3# = 0
'  fullPersVal = 0
'  fullFEVal = 0
'  fullMHVal = 0
'  fullMCVal = 0
'  fullMTVal = 0
'  fullPPTRAVal = 0
'  TfullPPTRAVal = 0
''  If TaxCust.Acct = 115 Then Stop
'  Dim ThisProp&
'  Do While PropertyRec! > 0
'    PPTRADisc# = 0
'    Get PHandle, PropertyRec!, PersRec
'    ThisProp = PropertyRec!
'    ThisMaxVehVal = MaxVehVal
'    ThisMinVehVal = MinVehVal
'    If PersRec.Deleted = True Then GoTo KeepGoing
'    If SupOnly = "Y" And PersRec.DISCOV <> "Y" Then GoTo KeepGoing
'    If PersRec.DISCOV = "N" And SupOnly = "Y" Then GoTo KeepGoing
'    If PersRec.LastYrPrinted = WhatPYear Then Discovery$ = "Y"
'    If MultiYear > 1 Or PersRec.DISCOV = "Y" Then GoTo GoAhead 'new for 2.05
'    If (PersRec.LastYrPrinted <> WhatPYear) Or (PersRec.DISCOV = "Y") Or (PersRec.LastYrPrinted = WhatPYear) Then
'GoAhead:
'      PYear$ = CInt(PersRec.TaxBillYear)
'      If Val(PYear$) > 0 Then
'        PYearInt = Val(PYear$)
'      Else
'        GoTo KeepGoing
'      End If
'
'      OptRevTax1# = FigureOptRevTax1(ThisProp&, PHandle, "P")
'      TBillRec.OptRevTax1 = OptRevTax1# + TBillRec.OptRevTax1
'      TotOpt1 = OldRound(TotOpt1 + OptRevTax1#)
'      OptRevTax2# = FigureOptRevTax2(ThisProp&, PHandle, "P")
'      TBillRec.OptRevTax2 = OptRevTax2# + TBillRec.OptRevTax2
'      TotOpt2 = OldRound(TotOpt2 + OptRevTax2#)
'      OptRevTax3# = FigureOptRevTax3(ThisProp&, PHandle, "P")
'      TBillRec.OptRevTax3 = OptRevTax3# + TBillRec.OptRevTax3
'      TotOpt3 = OldRound(TotOpt3 + OptRevTax3#)
'
'      Factor! = 1
'      Rate = PersRec.ProrateVal
'      If Rate > 0 Then
'        Factor! = Rate / 12
'        TBillRec.Prorate = "Y"
'      Else
'        TBillRec.Prorate = "N"
'        Factor = 1
'      End If
'
'      fullPersVal = OldRound(fullPersVal + PersRec.PersVal)
'      fullFEVal = OldRound(fullFEVal + PersRec.CVALUE)
'      fullMHVal = OldRound(fullMHVal + PersRec.MHValue)
'      fullMCVal = OldRound(fullMCVal + PersRec.MCValue)
'      fullMTVal = OldRound(fullMTVal + PersRec.MTValue)
'
'      If MultiYear <> 0 Then 'new for 2.05
'        PersRec.PersVal = OldRound(PersRec.PersVal / MultiYear)
'        PersRec.CVALUE = OldRound(PersRec.CVALUE / MultiYear)
'        PersRec.MHValue = OldRound(PersRec.MHValue / MultiYear)
'        PersRec.MCValue = OldRound(PersRec.MCValue / MultiYear)
'        PersRec.MTValue = OldRound(PersRec.MTValue / MultiYear)
'        ThisMaxVehVal = OldRound(ThisMaxVehVal / MultiYear)
'        ThisMinVehVal = OldRound(ThisMinVehVal / MultiYear)
'      End If
'
'      PersValue# = OldRound#(PersRec.PersVal * Factor!) + PersValue
'      FEValue# = OldRound#(PersRec.CVALUE) + FEValue#
'      MHValue# = OldRound#(PersRec.MHValue) + MHValue#
'      MCValue# = OldRound#(PersRec.MCValue) + MCValue#
'      MTValue# = OldRound#(PersRec.MTValue) + MTValue#
''      PersValue# = OldRound#(fullPersVal * Factor!) + PersValue#
''      FEValue# = OldRound#(fullFEVal) + FEValue#
''      MHValue# = OldRound#(fullMHVal) + MHValue#
''      MCValue# = OldRound#(fullMCVal) + MCValue#
''      MTValue# = OldRound#(fullMTVal) + MTValue#
'
'      If ABalance# > 0 Then
'        If LawChngDate > 0 Then
'          If Date2Num(fptxtBillDate.Text) >= LawChngDate Then
'            GoTo NoDisc 'new for 2006
'          End If
'        End If
'      End If
'
'      If PPTRAYN = False Then GoTo NoDisc
'
'      If PersRec.PPTRAYN = "Y" Then
'        If OldRound#(PersRec.PersVal * Factor!) > ThisMaxVehVal Then
''        If OldRound#(fullPersVal * Factor!) > ThisMaxVehVal Then
'          PPTRAVal# = ThisMaxVehVal
'          fullPPTRAVal = PPTRAVal * MultiYear
'        Else
'          PPTRAVal# = OldRound#(PersRec.PersVal * Factor!)
'          fullPPTRAVal = PPTRAVal * MultiYear
''          PPTRAVal# = OldRound#(fullPersVal * Factor!)
'        End If
'        If MultiYear <> 0 Then
'          ThisMinVehVal = OldRound(ThisMinVehVal / MultiYear)
'        End If
'        If PPTRAVal# <= (ThisMinVehVal * Factor!) Then
''          PPTRADisc# = OldRound#((PPTRAVal# / 100) * PERC! / 100)'remmed out 2/21/06
'          PPTRADisc# = OldRound(((PPTRAVal# / 100) * Factor) * PERSRATE#) '2/21/06
'        Else
'          PPTRADisc# = OldRound#(((PPTRAVal# / 100) * (PERC! / 100)) * PERSRATE#)
'        End If
'        TPPTRAVal# = OldRound(TPPTRAVal# + PPTRAVal#)
'        TfullPPTRAVal = OldRound(TfullPPTRAVal + fullPPTRAVal)
'      End If
'NoDisc:
''      PersTaxDue = OldRound#(((fullPersVal / 100) * Factor!) * PERSRATE#)
'      PersTaxDue = OldRound#(((PersRec.PersVal / 100) * Factor!) * PERSRATE#)
'      If PersTaxDue < 0 Then PersTaxDue = 0
''      If PPTRADisc > PersTaxDue Then 'remmed out on 2/21/06
''        PersTaxDue = 0
''        PPTRADisc = 0
''      End If
'    End If
'
'    TPersTaxDue = OldRound(TPersTaxDue + PersTaxDue) ' - (OptRevTax1# + OptRevTax2# + OptRevTax3#))
'    TPPTRADisc# = OldRound(PPTRADisc + TPPTRADisc#)
'
'KeepGoing:
'    PropertyRec! = PersRec.NextRec
'  Loop
'  MHTax# = OldRound((MHValue# / 100) * MHRate#)
'  MCTax# = OldRound((MCValue# / 100) * MCRate#)
'  FETax# = OldRound((FEValue# / 100) * FERate#)
'  MTTax# = OldRound((MTValue# / 100) * MTRate#)
'  If OldRound(TPersTaxDue + MHTax# + MCTax# + FETax# + MTTax# - TPPTRADisc#) = 0 Then
'    If fpcmbInclNoBills.Text = "N" Then Return
'  End If
'  TBillRec.FETaxDue = FETax#
'  TBillRec.FETaxRate = CDbl(fpDblSnglFERate.Value)
'  TBillRec.FEValue = fullFEVal 'FEValue# ' * MultiYear
'  TBillRec.MCTaxDue = MCTax#
'  TBillRec.MCTaxRate = CDbl(fpDblSnglMCRate.Value)
'  TBillRec.MCValue = fullMCVal 'MCValue# ' * MultiYear
'  TBillRec.MHTaxDue = MHTax#
'  TBillRec.MHTaxRate = CDbl(fpDblSnglMHRate.Value)
'  TBillRec.MHValue = fullMHVal 'MHValue# ' * MultiYear
'  TBillRec.MTTaxDue = MTTax#
'  TBillRec.MTTaxRate = CDbl(fpDblSnglMTRate.Value)
'  TBillRec.MTValue = fullMTVal 'MTValue# '* MultiYear
'  GTMHTax# = OldRound(GTMHTax# + MHTax#)
'  GTMCTax# = OldRound(GTMCTax# + MCTax#)
'  GTFETax# = OldRound(GTFETax# + FETax#)
'  GTMTTax# = OldRound(GTMTTax# + MTTax#)
'  GTPersTax# = OldRound(GTPersTax# + TPersTaxDue#)
'  TBillRec.PersTaxNet = OldRound(TPersTaxDue# - TPPTRADisc#)
''  TBillRec.PersTaxNet = OldRound(TPersTaxDue# - TPPTRADisc# + OptRevTax1# + OptRevTax2# + OptRevTax3#)
'  TBillRec.MultiYrVal = MultiYear
'  TBillRec.PersTaxDue = TPersTaxDue#
'  TBillRec.PPTRADiscnt = TPPTRADisc#
'  GTPersTaxNet# = OldRound(GTPersTaxNet# + TBillRec.PersTaxNet)
'  GTPTTRADisc# = OldRound(GTPTTRADisc# + TPPTRADisc#)
'  TBillRec.PersValue = fullPersVal 'PersValue# ' * MultiYear
'  GTPersValue# = OldRound(GTPersValue# + fullPersVal) 'PersValue#)
'  GTFEValue# = OldRound(GTFEValue# + fullFEVal) 'TBillRec.FEValue#)
'  GTMCValue# = OldRound(GTMCValue# + fullMCVal) 'TBillRec.MCValue#)
'  GTMHValue# = OldRound(GTMHValue# + fullMHVal) 'TBillRec.MHValue#)
'  GTMTValue# = OldRound(GTMTValue# + fullMTVal) 'TBillRec.MTValue#)
'  TBillRec.CustPin = TaxCust.PIN
'  TBillRec.InternalPin = PersRec.InternalPin
'  TBillRec.PersPin = QPTrim$(PersRec.PropPin)
'  TBillRec.ExptValue = 0
'  GTPPTRADisc = OldRound(GTPPTRADisc + TPPTRADisc#)
'  TBillRec.PPTRAValue = TfullPPTRAVal 'TPPTRAVal#
'  GTPPTRAVal = OldRound(GTPPTRAVal# + TfullPPTRAVal) 'TPPTRAVal#)
'  TBillRec.PersPropRecord = TaxCust.FirstPersRec
'  TBillRec.PersTaxRate = PERSRATE#
'  TBillRec.TotalBillDue = OldRound#(TBillRec.PersTaxDue + TBillRec.FETaxDue + TBillRec.MCTaxDue + TBillRec.MHTaxDue + TBillRec.MTTaxDue - TBillRec.PPTRADiscnt)
'  TBillRec.TotalBillDue = OldRound(TBillRec.TotalBillDue + OptRevTax1# + OptRevTax2# + OptRevTax3#)
'  If PersRec.LateList = "Y" Then
'    LateAmt# = OldRound#((LateAmt# + OldRound(TBillRec.PersTaxDue - TBillRec.PPTRADiscnt)) * (LateList# / 100))
'    TBillRec.LateTaxDue = LateAmt#
'    TBillRec.TotalBillDue = OldRound#(TBillRec.TotalBillDue + TBillRec.LateTaxDue)
'  End If
'  TBillRec.DueDate = Date2Num(fptxtDueDate.Text)
'  TBillRec.RDesc1 = QPTrim$(PersRec.DESC1)
'  GoSub WriteIt2Disk
'
'  Return
'
'ERRORSTUFF:
'   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPrebilling", "PrintGraphics", Erl)
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
'    ClearInUse PWcnt
'    Terminate

