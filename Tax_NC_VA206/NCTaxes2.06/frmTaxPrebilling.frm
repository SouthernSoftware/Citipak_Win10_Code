VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmTaxPrebilling 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Prebilling Information"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmTaxPrebilling.frx":0000
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
      TabIndex        =   5
      Top             =   5520
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
      ColDesigner     =   "frmTaxPrebilling.frx":08CA
   End
   Begin LpLib.fpCombo fpcmbCycle 
      Height          =   390
      Left            =   8040
      TabIndex        =   7
      Top             =   2160
      Width           =   3135
      _Version        =   196608
      _ExtentX        =   5530
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
      ColDesigner     =   "frmTaxPrebilling.frx":0CA5
   End
   Begin LpLib.fpCombo fpcmbCounty 
      Height          =   390
      Left            =   8040
      TabIndex        =   8
      Top             =   2640
      Width           =   3135
      _Version        =   196608
      _ExtentX        =   5530
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
      ColDesigner     =   "frmTaxPrebilling.frx":10D8
   End
   Begin LpLib.fpCombo fpcmbTownships 
      Height          =   390
      Left            =   8400
      TabIndex        =   6
      Top             =   1680
      Width           =   2775
      _Version        =   196608
      _ExtentX        =   4895
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
      ColDesigner     =   "frmTaxPrebilling.frx":150B
   End
   Begin LpLib.fpCombo fpcmbPrintOpt 
      Height          =   405
      Left            =   7560
      TabIndex        =   13
      Top             =   6120
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
      ColDesigner     =   "frmTaxPrebilling.frx":18E6
   End
   Begin LpLib.fpCombo fpcmbPrintOrder 
      Height          =   405
      Left            =   7080
      TabIndex        =   16
      Top             =   6960
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
      ColDesigner     =   "frmTaxPrebilling.frx":1CC1
   End
   Begin LpLib.fpCombo fpcmbSuppOnly 
      Height          =   405
      Left            =   3960
      TabIndex        =   4
      Top             =   4755
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
      ColDesigner     =   "frmTaxPrebilling.frx":209C
   End
   Begin LpLib.fpCombo fpcmbSplit 
      Height          =   390
      Left            =   8880
      TabIndex        =   9
      Top             =   3120
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
      ColDesigner     =   "frmTaxPrebilling.frx":2477
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
      TabIndex        =   11
      Top             =   4920
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
      TabIndex        =   10
      Top             =   4440
      Width           =   1932
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
      Height          =   495
      Left            =   6120
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   7800
      Width           =   1560
      _Version        =   131072
      _ExtentX        =   2752
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
      ButtonDesigner  =   "frmTaxPrebilling.frx":2852
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   495
      Left            =   3960
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   7800
      Width           =   1560
      _Version        =   131072
      _ExtentX        =   2752
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
      ButtonDesigner  =   "frmTaxPrebilling.frx":2A31
   End
   Begin EditLib.fpDoubleSingle fpDblSnglRealRate 
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   2640
      Width           =   1335
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
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1335
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
      ControlType     =   1
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
      Left            =   3600
      TabIndex        =   3
      Top             =   4080
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
   Begin EditLib.fpDateTime fptxtCurrYear 
      Height          =   345
      Left            =   3240
      TabIndex        =   0
      Top             =   1680
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
      TabIndex        =   12
      Top             =   4920
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
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Expiration Date For This Billing "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   612
      Left            =   9000
      TabIndex        =   36
      Top             =   4200
      Width           =   1812
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0080FFFF&
      BorderStyle     =   3  'Dot
      X1              =   240
      X2              =   6240
      Y1              =   3720
      Y2              =   3720
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
      Height          =   255
      Left            =   3600
      TabIndex        =   35
      ToolTipText     =   "Enter the percent amount as follows: 5% = 5; .5% = .5."
      Top             =   2355
      Width           =   1335
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0080FFFF&
      BorderStyle     =   3  'Dot
      X1              =   240
      X2              =   6240
      Y1              =   2280
      Y2              =   2280
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
      Height          =   255
      Left            =   6240
      TabIndex        =   34
      Top             =   3720
      Width           =   3735
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   6240
      X2              =   11280
      Y1              =   5640
      Y2              =   5640
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
      TabIndex        =   33
      Top             =   5604
      Width           =   4692
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Bill Settings:"
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
      Height          =   255
      Left            =   240
      TabIndex        =   32
      Top             =   1320
      Width           =   1815
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
      Height          =   375
      Left            =   6480
      TabIndex        =   31
      Top             =   2280
      Width           =   1455
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
      Height          =   375
      Left            =   6360
      TabIndex        =   30
      Top             =   2730
      Width           =   1575
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
      Height          =   375
      Left            =   6600
      TabIndex        =   29
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   6240
      X2              =   11280
      Y1              =   3720
      Y2              =   3720
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
      Height          =   255
      Left            =   6240
      TabIndex        =   28
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   6255
      Left            =   6240
      Top             =   1320
      Width           =   5055
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   4935
      Left            =   240
      Top             =   1320
      Width           =   6015
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
      Height          =   375
      Left            =   6360
      TabIndex        =   27
      Top             =   1785
      Width           =   1935
   End
   Begin VB.Label lblSupp1 
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
      Height          =   735
      Left            =   480
      TabIndex        =   26
      Top             =   6600
      Width           =   5535
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
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   6240
      Width           =   2535
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1335
      Left            =   240
      Top             =   6240
      Width           =   6015
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
      TabIndex        =   24
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label Label8 
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
      Height          =   375
      Left            =   1680
      TabIndex        =   23
      Top             =   4835
      Width           =   2175
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
      TabIndex        =   22
      Top             =   6600
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
      Height          =   255
      Left            =   1920
      TabIndex        =   21
      Top             =   4155
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Property:"
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
      Left            =   1200
      TabIndex        =   20
      Top             =   3200
      Width           =   2175
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
      Left            =   1680
      TabIndex        =   19
      Top             =   2720
      Width           =   1695
   End
   Begin VB.Label Label4 
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
      Height          =   255
      Left            =   2040
      TabIndex        =   18
      Top             =   1755
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   660
      Index           =   1
      Left            =   1493
      Top             =   345
      Width           =   8655
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
      Height          =   390
      Left            =   3143
      TabIndex        =   14
      Top             =   510
      Width           =   5295
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   1493
      Top             =   240
      Width           =   8655
   End
End
Attribute VB_Name = "frmTaxPrebilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim UseOpt As String * 1
  Dim UseSS As String * 1
  Dim ThisOpt$
  Dim WhatYear As Integer
  Dim BillNum&
  Dim REALRATE#
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
  Dim SpecTax As Double
  Dim SpecDesc As String
  Dim ThisTown$
  Dim Townships() As String
  Dim TSCnt As Integer
  
Private Sub cmdExit_Click()
  Dim BillInfo As TaxBillInfoType
  Dim BIHandle As Integer
  
  KillFile "C:\CPWork\revglbill.dat"
  frmTaxBillingMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdProcess_Click()
  Dim BillInfo As TaxBillInfoType
  Dim BIHandle As Integer
  Dim IdxType As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim MortCodeRec As MortCodeRecType
  Dim MCHandle As Integer
  Dim x As Integer
  Dim ThisYear$
  Dim FileName$
  
  'on error goto ERRORSTUFF
  
  fpcmbCounty.Col = 1
  fpcmbCounty.Row = fpcmbCounty.ListIndex
  fpcmbCycle.Col = 1
  fpcmbCycle.Row = fpcmbCycle.ListIndex
  fpcmbTownships.Col = 1
  fpcmbTownships.Row = fpcmbTownships.ListIndex
  ThisYear = CStr(fptxtCurrYear.Text)
  FileName = "TAXBILLBU/POSTT" + Mid(fpcmbCounty.ColText, 1, 3) + Mid(fpcmbCycle.ColText, 1, 3) + Mid(fpcmbTownships.Text, 1, 3) + ThisYear + ".DAT"
  If Exist(FileName) Then
    If TaxMsgWOpts(600, "Personal tax bills with these parameters have already been posted for tax year " + fptxtCurrYear.Text + ". Press F10 to continue prebilling. Otherwise, press ESC to abort prebilling. NOTE: If you wish to print a posted bill then please exit this screen and go to 'Reprint Posted Tax Bills'. ", "F10 Continue", "ESC Abort") = "abort" Then
      Unload frmTaxMsgWOpts
      MainLog ("WARNING: User elected to abort prebilling processing for tax year " + fptxtCurrYear.Text + " when they were warned the parameters for this year had already been billed and posted.")
      Close
      Exit Sub
    Else
      Unload frmTaxMsgWOpts
      MainLog ("WARNING: User elected to continue prebilling processing for tax year " + fptxtCurrYear.Text + " after they were warned the parameters for this year had already been billed and posted.")
    End If
  End If
  
  If Exist("txblsprn.dat") Then
    If TaxMsgWOpts(800, "NOTE: Tax bills will need to be printed again upon completion of this process before posting can take place. Press F10 if you wish to continue. Otherwise, press ESC to abort prebilling.", "F10 Continue", "ESC Abort") = "abort" Then
      Unload frmTaxMsgWOpts
      Close
      Exit Sub
    Else
      Unload frmTaxMsgWOpts
      MainLog ("WARNING: User elected to continue the prebilling process after being warned that tax bills will have to be reprinted.")
    End If
  End If
  
  If CheckPreReq = False Then
    Exit Sub
  Else
    If RevsAndGLsOK(Me, CInt(fptxtCurrYear.Text)) = False Then
      Exit Sub
    End If
    KillFile TaxBillFile
    KillFile TaxBillInfoFile
    OpenBillInfoFile BIHandle
      BillInfo.BillNum = 0 'BillNum&
      BillInfo.PERSRATE = PERSRATE#
      BillInfo.LATEPCT = LateList#
      BillInfo.PRNORDER = Order$
      BillInfo.REALRATE = REALRATE#
      BillInfo.TaxYear = WhatYear
      If fpcmbCounty.Enabled = True Then
        fpcmbCounty.Col = 1
        BillInfo.CountyPara = QPTrim$(fpcmbCounty.ColText)
      Else
        BillInfo.CountyPara = "ALL COUNTIES"
      End If
      If fpcmbTownships.Enabled = True Then
        BillInfo.TwnShpPara = QPTrim$(fpcmbTownships.Text)
      Else
        BillInfo.TwnShpPara = "ALL TOWNSHIPS"
      End If
      If fpcmbCycle.Enabled = True Then
        fpcmbCycle.Col = 1
        BillInfo.CyclePara = QPTrim$(fpcmbCycle.ColText)
      Else
        BillInfo.CyclePara = "No CYCLES"
      End If
      If fpcmbSplit.Enabled = True Then
        BillInfo.SplitPara = QPTrim$(fpcmbSplit.Text)
      Else
        BillInfo.SplitPara = "NO SPLIT"
      End If
      If OptDiscYes.Value = True Then 'added 9/20/05
        BillInfo.XDate = Date2Num(fptxtDiscXDate.Text)
      Else
        BillInfo.XDate = 0
      End If
      
      Put BIHandle, 1, BillInfo
      Close BIHandle
  End If
  
  IdxType = CInt(Order$)
  Call MakeCustIdx(IdxType) 'populates global CustArr()
  
'  OpenTaxSetUpFile TMHandle 'remmed on 9/20/05
'  Get TMHandle, 1, TaxMasterRec
'  TaxMasterRec.TaxYear = CInt(fptxtCurrYear.Text)
'  If OptDiscYes.Value = True Then 'remarked on 9/20/05
'    TaxMasterRec.DiscXDate = Date2Num(fptxtDiscXDate.Text)
'  Else
'    TaxMasterRec.DiscXDate = 0
'  End If
'  Put TMHandle, 1, TaxMasterRec
'  Close TMHandle
  
  GoSub LoadMortCodes

  If fpcmbPrintOpt.Text = "Graphical" Then
    Call PrintGraphics
  Else
    Call PrintText
  End If
  
  KillFile "ZIPIDX.DAT" 'added 12/706
  KillFile "MORTIDX.DAT" 'added 12/7/06
  
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPrebilling", "cmdProcess_Click", Erl)
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
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxPrebilling.")
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

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  MainLog ("User opened frmTaxPrebilling.")
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
  
  'on error goto ERRORSTUFF
  
  UseOpt = "N"
  UseSS = "N"
  ReDim MortCodes(1 To 1) As MortRecType
  fptxtDiscXDate.Text = CStr(Date)
  lblSupp1.Caption = ""
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  If TaxMasterRec.RealPersSplit = "Y" Then
    fpcmbSplit.Text = "NO SPLIT"
    fpcmbSplit.AddItem "NO SPLIT"
    fpcmbSplit.AddItem "REAL ONLY"
    fpcmbSplit.AddItem "PERSONAL ONLY"
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
     
  If TaxMasterRec.TaxYear > 0 Then
    fptxtCurrYear.Text = CStr(TaxMasterRec.TaxYear)
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
  
  fpcmbSuppOnly.Text = "N"
  fpcmbSuppOnly.AddItem "N"
  fpcmbSuppOnly.AddItem "Y"
  fpcmbInclNoBills.Text = "N"
  fpcmbInclNoBills.AddItem "N"
  fpcmbInclNoBills.AddItem "Y"
  fpcmbPrintOpt.Text = "Graphical"
  fpcmbPrintOpt.AddItem "Graphical"
  fpcmbPrintOpt.AddItem "Text"
  
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPrebilling", "LoadMe", Erl)
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
      fpcmbTownships.SetFocus
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
  UseOpt = "N"
  UseSS = "N"
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
      fptxtCurrYear.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
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
      OptDiscNo.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

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

Private Sub fpDblSnglRealRate_Change()
  fpDblSnglPersRate.Value = fpDblSnglRealRate.Value
End Sub

Private Function CheckPreReq() As Boolean
  Dim ThisYear As Integer
  Dim LenThisYear As Integer
  
  'on error goto ERRORSTUFF
  
  CheckPreReq = True 'nothing bad found yet
  WhatYear = CInt(fptxtCurrYear.Text)
  BillNum& = 1 'fpDblSnglStartBill.Value
  REALRATE# = fpDblSnglRealRate.Value
  PERSRATE# = fpDblSnglPersRate.Value
  LateList# = fpDblSnglLateList.Value
  Order$ = Mid(fpcmbPrintOrder.Text, 1, 1)
  SupOnly = QPTrim$(fpcmbSuppOnly.Text)
  RptType$ = Mid(fpcmbPrintOpt.Text, 1, 1)

  LenThisYear = Len(Date)
  LenThisYear = LenThisYear - 3
  ThisYear = Mid(Date, LenThisYear, Len(Date))
  If Abs(WhatYear - ThisYear) > 10 Then
    If TaxMsgWOpts(800, "The year entered is more than 10 years from the current year. Press F10 if you wish to continue anyway. Otherwise, press ESC to review.", "F10 Continue", "ESC Review") = "abort" Then
      Unload frmTaxMsgWOpts
      fptxtCurrYear.SetFocus
      Close
      CheckPreReq = False
      Exit Function
    Else
      Unload frmTaxMsgWOpts
      MainLog ("WARNING: frmTaxPrebilling - User warned that the date entered " + CStr(WhatYear) + " is ten years more than today. The user continued saving anyway.")
    End If
  End If
  
  If fpDblSnglRealRate.Value = 0 Then
    If TaxMsgWOpts(800, "The tax rate entered is zero. If you wish to continue anyway then press F10. Otherwise, press ESC to review.", "F10 Continue", "ESC Review") = "abort" Then
      Unload frmTaxMsgWOpts
      fpDblSnglRealRate.SetFocus
      CheckPreReq = False
      Close
      Exit Function
    Else
      Unload frmTaxMsgWOpts
      MainLog ("WARNING: frmTaxPrebilling - User wanred that the tax rate entered is zero but the user continued running the prebilling register anyway.")
    End If
  End If
  Exit Function
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPrebilling", "CheckPreReq", Erl)
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
  
  'on error goto ERRORSTUFF
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
  Next x
  Close TCHandle
  
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
  Loop
        
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
  frmTaxShowPctComp.Label1 = "Creating Customer Index"
  frmTaxShowPctComp.cmdCancel.Visible = False
  frmTaxShowPctComp.Show , Me
  
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
    frmTaxShowPctComp.ShowPctComp Nextx, NumOfTCRecs
  Loop
  Unload frmTaxShowPctComp
  Close TCHandle
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPrebilling", "MakeCustIdx", Erl)
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
  Dim TBillRec As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim OPBillRec As TaxTransactionType
  Dim OPHandle As Integer
  Dim NumOfOPRecs As Long
  Dim RealRec As PropertyRecType
  Dim RRHandle As Integer
  Dim NumOfRRREcs As Long
  Dim PersRec As PersonalRecType
  Dim PRHandle As Integer
  Dim NumOfPRRecs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  Dim NewTBillRec As TaxBillType
  Dim Done2Disk As Integer
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
  Dim GotPersVal As Integer
  Dim PersExmp#
  Dim PersValue#
  Dim Discovery$
  Dim PersCalcVal#
  Dim PersTaxDue#
  Dim NextPersRec&
  Dim TotalPers#
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
  Dim BillInfo As TaxBillInfoType
  Dim BIHandle As Integer
  Dim TotReal As Double, TotPers As Double
  Dim OverPay As Boolean
  Dim OverPayAmt As Double
  Dim OPApplied As Double
  Dim ThisTBCnt As Long
  Dim ThisTest$
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim OptFlag As Boolean
  Dim ThisPersExmp As Double
  Dim ThisPersValue As Double
  Dim Pct1#, Pct2#, Pct6#, Pct7#, Pct8#, Pct9#
  Dim PctTot#, PctTest#, ThisReal#, ThisPers#
  
  'on error goto ERRORSTUFF
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  TotOpt1 = 0
  TotOpt2 = 0
  TotOpt3 = 0
  OptRev1Desc = QPTrim$(TaxMasterRec.OptRev1)
  OptRev2Desc = QPTrim$(TaxMasterRec.OptRev2)
  OptRev3Desc = QPTrim$(TaxMasterRec.OptRev3)
  
  fpcmbCycle.Col = 1
  CycleName = QPTrim$(fpcmbCycle.ColText)
  If fpcmbCycle.Enabled = True And CycleName <> "NO CYCLE" Then
    fpcmbCycle.Col = 0
    CycleNum = CLng(fpcmbCycle.ColText)
  Else
    CycleNum = -1
  End If
  
  fpcmbCounty.Col = 1
  CountyName = QPTrim$(fpcmbCounty.ColText)
  If fpcmbCounty.Enabled = True And CountyName <> "ALL COUNTIES" Then
    fpcmbCounty.Col = 0
    CountyNum = CLng(fpcmbCounty.ColText)
  Else
    CountyNum = -1
  End If
  
  If fpcmbSplit.Enabled = True Then
    If fpcmbSplit.Text = "NO SPLIT" Then
      SplitYN = "No"
    ElseIf fpcmbSplit.Text = "REAL ONLY" Then
      SplitYN = "Real"
    ElseIf fpcmbSplit.Text = "PERSONAL ONLY" Then
      SplitYN = "Pers"
    Else
      SplitYN = "No"
    End If
  Else
    SplitYN = "No"
  End If
  
  If QPTrim$(fpcmbTownships.Text) = "ALL TOWNSHIPS" Then
    ThisTownship = "ALL"
  Else
    ThisTownship = QPTrim$(fpcmbTownships.Text)
    ThisTownship = Mid(ThisTownship, 1, Len(ThisTownship) - 4)
    ThisTownship = QPTrim$(ThisTownship)
  End If
  dlm$ = "~"
  If Exist(TaxBillFile) Then
    KillFile TaxBillFile
  End If
  If Exist(TaxBillOPFile) Then
    KillFile TaxBillOPFile
  End If
  If Exist("TMPBLADD") Then 'tax bill addon
    KillFile "TMPBLADD"
  End If
  
  AddOnCnt = 0
  OpenTaxBillAddOn AOHandle
  OpenTaxBillFile TBHandle, NumOfTBRecs
  OpenTaxPropFile RRHandle, NumOfRRREcs
  OpenTaxPersFile PRHandle, NumOfPRRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenMortCodeFile MCHandle, NumOfMCRecs
  
  Inactive = 0
  
  frmTaxShowPctComp.Label1 = "Creating Tax Pre-Billing Register"
  frmTaxShowPctComp.Show , Me
  frmTaxShowPctComp.cmdCancel.Visible = False
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
    If UsingIdx = True Then
      Get TCHandle, CustArr(x), TaxCust
      ThisTaxRec = CustArr(x)
    Else
      Get TCHandle, x, TaxCust
      ThisTaxRec = x
    End If
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
    
    Done2Disk = 0
    LateAmt# = 0
    OPApplied = 0
    If TaxCust.Deleted <> 0 Then
      GoTo PreBillSkip:
    End If
    If QPTrim$(TaxCust.Active) <> "Y" Then
      Inactive = Inactive + 1
      TaxCust.CustName = TaxCust.CustName
      GoTo PreBillSkip:
    End If
    PastFlagSet = 0             'Initialize Past Balance Flag
    OverPayAmt = 0
    If UsingIdx = True Then
      OverPayAmt = GetCustBalance(CustArr(x), -1)
      If OverPayAmt < 0 Then
        OverPay = True
      Else
        OverPayAmt = 0
        OverPay = False
      End If
    Else
      OverPayAmt = GetCustBalance(x, -1)
      If OverPayAmt < 0 Then
        OverPay = True
      Else
        OverPayAmt = 0
        OverPay = False
      End If
    End If
    ThisTest = CStr(OverPayAmt)
    If InStr(ThisTest, "E") Then OverPayAmt = 0
    
    'If ThisTaxRec = 3428 Then Stop
    
    If TaxCust.FirstPersRec <= 0 And TaxCust.FirstPropRec <= 0 Then
      NoProp = 1
      GoSub SetCustInfo
      GoSub WriteIt2Disk
      GoTo PreBillSkip
    End If
    
    NoProp = 0
    GoSub SetCustInfo
    GoSub GetPersInfo
    GoSub GetRealInfo
    
PreBillSkip:
    frmTaxShowPctComp.ShowPctComp x, NumOfTCRecs
'    If x / NumOfTCRecs = 0.85 Then Stop
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      Exit Sub
    End If
  Next x
  
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  
  Close RRHandle
  Close PRHandle
  Close TCHandle
  Close MCHandle
  Close AOHandle
  
  If ThisTBCnt = 0 Then
    Call TaxMsg(900, "Using the parameters selected there are no customers who qualify for a tax charge.")
    Close
    fptxtCurrYear.SetFocus
    Exit Sub
  End If
  
  TotalPers# = 0
  TotalReal# = 0
  TotalEx# = 0
  NumBills& = 0
  
  TaxPreRptFile$ = "TAXRPTS\TaxPreBill.RPT"
  
  RptHandle = FreeFile
  Open TaxPreRptFile For Output As #RptHandle
  
  Close TBHandle
  OpenTaxBillFile TBHandle, NumOfTBRecs
'Graphics printing
  For x = 1 To NumOfTBRecs
    Get TBHandle, x, TBillRec
    'If TBillRec.BillNumber = -1 Then 'took out Discovery$ = "Y" on 12/5/06
    If TBillRec.BillNumber = -1 And TBillRec.ExptValue = 0 Then 'took out Discovery$ = "Y" on 12/5/06
       GoTo NotThisOne
    Else
      If TBillRec.TotalBillDue = 0 And TBillRec.PriorYrBalance = 0 And TBillRec.ExptValue = 0 Then
        If OldRound(TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3) = 0 Then 'added 10/18/06
          If fpcmbInclNoBills.Text = "N" Then 'GoTo NextOne
            GoTo NotThisOne
          End If
        End If
      End If
      '                        0
      Print #RptHandle, TBillRec.CustRec; dlm;
      '                            1
      Print #RptHandle, QPTrim$(TBillRec.CustName); dlm;
'      If TBillRec.PriorYrBalance > 0 Then'comment out 8/14/06
        '                            2
        Print #RptHandle, TBillRec.PriorYrBalance; dlm;
'      Else'comment out 8/14/06
'        '                 2
'        Print #RptHandle, 0; dlm;'comment out 8/14/06
'      End If'comment out 8/14/06
      If TBillRec.BillNumber = -1 Then
        '                   3
        Print #RptHandle, "N/A"; dlm;
      Else
        '                          3
        Print #RptHandle, TBillRec.BillNumber; dlm;
      End If
      '                           4
      Print #RptHandle, OldRound(TBillRec.TotalBillDue - TBillRec.LateTaxDue); dlm; '8/14/06 updated to add late tax in this field
      '                           5
      Print #RptHandle, TBillRec.RealValue; dlm;
      '                           6
      Print #RptHandle, TBillRec.PersValue; dlm;
      '                           7
      Print #RptHandle, TBillRec.ExptValue; dlm;
      TValue# = OldRound#(TBillRec.RealValue + TBillRec.PersValue - TBillRec.ExptValue)
      If TValue# < 0 Then TValue# = 0
      '                    8
      Print #RptHandle, TValue#; dlm;
      '                           9
      Print #RptHandle, TBillRec.LateTaxDue; dlm;
      TotalPers# = OldRound#(TotalPers# + TBillRec.PersValue)
      TotalReal# = OldRound#(TotalReal# + TBillRec.RealValue)
      TotalEx# = OldRound#(TotalEx# + TBillRec.ExptValue)
      TotalBills# = OldRound#(TotalBills# + TBillRec.PersTaxDue + TBillRec.RealTaxDue)
      TotalLate# = OldRound#(TotalLate# + TBillRec.LateTaxDue)
'      TotalPast# = OldRound#(TotalPast# + TBillRec.PriorYrBalance)
      If TBillRec.PriorYrBalance > 0 Then 'added 8/14/06
        TotalPast# = OldRound#(TotalPast# + TBillRec.PriorYrBalance)
      Else 'added 8/14/06
'        TotalPast# = 0 'added 8/14/06
      End If 'added 8/14/06
      TotalOverPay# = OldRound#(TotalOverPay# + TBillRec.OverPayAmt)

      If TBillRec.TotalBillDue > 0 Then
        NumBills& = NumBills& + 1
      End If
      '                     10
      Print #RptHandle, NumBills&; dlm;
      '                     11
      Print #RptHandle, TotalReal#; dlm;
      '                     12
      Print #RptHandle, TotalPers#; dlm;
      '                    13
      Print #RptHandle, TotalEx#; dlm;
      '                     14
      Print #RptHandle, TotalBills#; dlm;
      '                     15
      Print #RptHandle, TotalPast#; dlm;
      '                                   16
      Print #RptHandle, OldRound#(TotalPast# + TotalBills#); dlm;
      '                     17
      Print #RptHandle, TotalLate#; dlm;
      '                    18
      Print #RptHandle, Inactive; dlm;
      '                         19
      Print #RptHandle, TBillRec.RealPin; dlm;
      '                     20
      Print #RptHandle, ThisTown; dlm;
      '                    21
      Print #RptHandle, WhatYear; dlm;
      '                     22                23              24
      Print #RptHandle, ThisTownship; dlm; CycleName; dlm; CycleNum; dlm;
      '                     25                26
      Print #RptHandle, CountyName$; dlm; CountyNum; dlm;
      '                     27                28                29
      Print #RptHandle, TBillRec.OptRevTax1; dlm; TBillRec.OptRevTax2; dlm; TBillRec.OptRevTax3; dlm;
       '                      30                 31                32
      Print #RptHandle, OptRev1Desc$; dlm; OptRev2Desc$; dlm; OptRev3Desc$; dlm;
      '                   33             34           35                     36                                  37
      Print #RptHandle, TotOpt1; dlm; TotOpt2; dlm; TotOpt3; dlm; -TBillRec.OverPayAmt; dlm; OldRound(TotalBills# - TotalOverPay); dlm;
      '                      38                          39                                40                                    41
      Print #RptHandle, TotalOverPay#; dlm; CDbl(fpDblSnglRealRate.Value) / 100; dlm; CDbl(fpDblSnglPersRate.Value) / 100; dlm; CDbl(fpDblSnglLateList.Value) / 100
    End If      'Test for Discovery Bills
NotThisOne:
  Next x
  
  Close
  TotReal = TotReal
  TotPers = TotPers
  arTaxPreBillLS.Show
  frmTaxLoadReport.Show
   
  KillFile "txblsprn.dat"
  
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
  TBillRec.TaxYear = WhatYear
  TBillRec.RDesc3 = TaxCust.CSSN

  'Set Prior Balance if any
  ABalance = GetCustBalance(TaxCust.Acct, -1) 'added this line on 8/11/2006    0
'  GoSub GetPastBalance '8/14/06
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
      If TaxTrans.TranType = 1 Then
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
  
WriteIt2Disk:   'write the info out to disk here.

  TBillRec.BillPrinted = False
  If TBillRec.TotalBillDue > 0 Then
'    If ShoresFlag Then
'      TBillRec(1).CarShore = 60
'      TStorm# = Round(TStorm# + TBillRec(1).CarShore)
'      StormCnt& = StormCnt& + 1
'      'TBillRec(1).TotalBillDue = Round#(TBillRec(1).TotalBillDue + TBillRec(1).CarShore)
'    End If
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
      TBillRec.SetDscvry2No = "Y" 'added 12/5/06
      TBillRec.RealTaxDue = 0 'added 12/5/06
      TBillRec.LateTaxDue = 0 'added 12/5/06
      TotOpt1# = OldRound(TotOpt1# - TBillRec.OptRevTax1) 'added 12/5/06
      TBillRec.OptRevTax1 = 0 'added 12/5/06
      TotOpt2# = OldRound(TotOpt2# - TBillRec.OptRevTax2) 'added 12/5/06
      TBillRec.OptRevTax2 = 0 'added 12/5/06
      TotOpt3# = OldRound(TotOpt3# - TBillRec.OptRevTax3) 'added 12/5/06
      TBillRec.OptRevTax3 = 0 'added 12/5/06
      TBillRec.PersTaxDue = 0 'added 12/4/06
      TotalLate# = OldRound(TotalLate# - TBillRec.LateTaxDue) 'added 12/4/06
      TBillRec.LateTaxDue = 0 'added 12/4/06
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
      Pct9# = 0
      ThisReal = OldRound(TBillRec.RealTaxDue - (TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3))
      ThisPers = OldRound(TBillRec.PersTaxDue)
      PctTot# = OldRound(TBillRec.TotalBillDue)
      If ThisReal > 0 Then
        Pct1# = OldRound#(ThisReal / PctTot#)
      End If
      If TBillRec.LateTaxDue > 0 Then
        Pct2# = OldRound#(TBillRec.LateTaxDue / PctTot#)
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
      If ThisPers > 0 Then
        Pct9# = OldRound#(TBillRec.PersTaxDue / PctTot#)
      End If
      
      ThisReal = MinBill * Pct1
      ThisPers = MinBill * Pct9
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
      
      PctTest# = OldRound(ThisReal + ThisPers + TBillRec.LateTaxDue)
      PctTest# = OldRound(PctTest# + TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3)
      
      If MinBill < PctTest# Then
        If ThisReal > PctTest# - MinBill Then
          ThisReal = OldRound(ThisReal - (PctTest# - MinBill))
          TotalReal# = OldRound(TotalReal# - (PctTest# - MinBill))
        ElseIf ThisPers > PctTest# - MinBill Then
          ThisPers = OldRound(ThisPers - (PctTest# - MinBill))
          TotalPers# = OldRound(TotalPers# - (PctTest# - MinBill))
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
        ElseIf ThisPers > MinBill - PctTest# Then
          ThisPers = OldRound(ThisPers + (MinBill - PctTest#))
          TotalPers# = OldRound(TotalPers# + (MinBill - PctTest#))
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
      TBillRec.PersTaxDue = ThisPers
      TBillRec.TotalBillDue = MinBill
    End If
  End If
  
'  If SYLVAFlag Then
'    If (TBillRec(1).TotalBillDue > 0) And (TBillRec(1).TotalBillDue < 5.01) Then
'      'L5 = L5 + 1
'      'STOP
'      'TBillRec(1).TotalBillDue = 0
'      TBillRec(1).BillNumber = -100
'    End If
'  End If
  ThisTBCnt = ThisTBCnt + 1
'  TBillRec.MORTCODE = TBillRec.MORTCODE
  Put TBHandle, ThisTBCnt, TBillRec
  ThisTBCnt = LOF(TBHandle) / Len(TBillRec)
  OPApplied = 0 'added 10/17/08
  If TBillRec.TotalBillDue > 0 Then
    If Abs(OverPayAmt) > 0 Then
      OPBillRec.Revenue.LateListPd = 0
      OPBillRec.Revenue.Principle1Pd = 0
      OPBillRec.Revenue.RevOpt1Pd = 0
      OPBillRec.Revenue.RevOpt2Pd = 0
      OPBillRec.Revenue.RevOpt3Pd = 0
      OpenTaxBillOverPayFile OPHandle, NumOfOPRecs
      
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
      
      If TBillRec.PersTaxDue > 0 And Abs(OverPayAmt) > 0 Then
        If Abs(OverPayAmt) > TBillRec.PersTaxDue Then
          OverPayAmt = OldRound(OverPayAmt + TBillRec.PersTaxDue)
          OPBillRec.Revenue.Principle1Pd = TBillRec.PersTaxDue
          OPApplied = OldRound(OPApplied + OPBillRec.Revenue.Principle1Pd)
        ElseIf Abs(OverPayAmt) <= TBillRec.PersTaxDue Then
          OPBillRec.Revenue.Principle1Pd = -OverPayAmt
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
'          OPApplied = OldRound(OPApplied - OverPayAmt) 'changed to line below on 10/17/08
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
'          OPApplied = OldRound(OPApplied - OverPayAmt) 'changed to line below on 10/17/08
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
'          OPApplied = OldRound(OPApplied - OverPayAmt) 'changed to line below on 10/17/08
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
'          OPApplied = OldRound(OPApplied - OverPayAmt) 'changed to line below on 10/17/08
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
  Done2Disk = -1
  
  Return
  
GetPersInfo:
  If SplitYN = "Real" Then Return 'added 5/12/05
  
  GotPersVal = 0
  PersExmp# = 0
  PersValue# = 0
  PersCalcVal# = 0
  PersTaxDue# = 0 'added 7/3/2008
  ThisPersValue# = 0 'added 10/21/08
  
  If TaxCust.FirstPersRec > 0 Then
    Get PRHandle, TaxCust.FirstPersRec, PersRec
    If PersRec.Deleted = True Then GoTo SupOnlyNGSkip3 'added 5/6/05
'    If PersRec.LastYrPrinted = WhatYear Then Discovery$ = "Y" 'took out 12/5/06
    If SupOnly = "Y" And PersRec.DISCOV = "N" Then GoTo SupOnlyNGSkip3 'added 10/18/05
    If (PersRec.LastYrPrinted <> WhatYear) Or (PersRec.DISCOV = "Y") Or (PersRec.LastYrPrinted = WhatYear) Then 'added third argument
    'to accommodate the ability to
      If SupOnly = "Y" And PersRec.DISCOV <> "Y" Then
        GoTo SupOnlyNGSkip3
      End If
    'IF PersRec(1).LastYrPrinted <> WhatYear THEN
      PersValue# = OldRound#(PersRec.PersVal + PersRec.MHVALUE + PersRec.MCVALUE)
      PersValue# = OldRound#(PersValue# + PersRec.CVALUE + PersRec.MTVALUE)
      PersExmp# = OldRound#(PersRec.EXMPSENI + PersRec.EXMPOTHR)
      If PersRec.EXMPSENI > 0 Then
        TempAddOn.OldAmt = 0
        TempAddOn.CustName = QPTrim$(TaxCust.CustName)
        TempAddOn.CustRec = x
        TempAddOn.Type = "Senior discount of " + QPTrim$(Using$("$#,##0.00", PersRec.EXMPSENI)) + " applied to personal property tax."
        AddOnCnt = AddOnCnt + 1
        TempAddOn.NewAmt = PersRec.EXMPSENI
        Put AOHandle, AddOnCnt, TempAddOn
      End If
      If PersRec.EXMPOTHR > 0 Then
        TempAddOn.OldAmt = 0
        TempAddOn.CustName = QPTrim$(TaxCust.CustName)
        TempAddOn.CustRec = x
        TempAddOn.Type = "Other discount of " + QPTrim$(Using$("$#,##0.00", PersRec.EXMPOTHR)) + " applied to personal property tax."
        AddOnCnt = AddOnCnt + 1
        TempAddOn.NewAmt = PersRec.EXMPOTHR
        Put AOHandle, AddOnCnt, TempAddOn
      End If
      PersCalcVal# = OldRound#((PersValue# - PersExmp#) / 100)
    Else
      PersCalcVal# = 0
    End If
    If PersCalcVal# < 0 Then
      PersCalcVal# = 0
    End If
    
    PersTaxDue# = OldRound#(PersCalcVal# * PERSRATE#)
    If PersTaxDue# < 0 Then PersTaxDue# = 0
    TotPers = TotPers + PersTaxDue#
    TBillRec.CustPin = TaxCust.PIN
    TBillRec.InternalPin = PersRec.InternalPin
    TBillRec.PersPin = QPTrim$(PersRec.PropPin)
    TBillRec.ExptValue = PersExmp#
    TBillRec.PersValue = PersValue#
    ThisPersValue# = PersValue# 'added for 2.05 8/24/06
    
    TBillRec.PersTaxDue = PersTaxDue#
    TBillRec.PersPropRecord = TaxCust.FirstPersRec
    TBillRec.PersTaxRate = PERSRATE#
    TBillRec.RealTaxRate = REALRATE#
    TBillRec.MORTCODE = ""
    TBillRec.MortRec = 0
'    TBillRec.TotalBillDue = OldRound#(TBillRec.PersTaxDue)
    TBillRec.TotalBillDue = TBillRec.PersTaxDue
    If TBillRec.TotalBillDue > 0 Then
      GotPersVal = -1
    End If
    If PersRec.LateList = "Y" Then
      LateAmt# = OldRound#((LateAmt# + TBillRec.PersTaxDue) * (LateList# / 100))
      TBillRec.LateTaxDue = LateAmt#
      TBillRec.TotalBillDue = OldRound#(TBillRec.PersTaxDue + LateAmt#)
    End If
  Else
    'no need to set the variables all are zero already
  End If

SupOnlyNGSkip3:
  NextPersRec& = PersRec.NextRec
  If NextPersRec& > 0 Then
    Do
'      TBillRec = NewTBillRec        'remmed 8/24/06 'make a new empty record each time
      GoSub SetCustInfo
      Get PRHandle, NextPersRec&, PersRec
      If PersRec.Deleted = True Then GoTo SupOnlyNGSkip4 'added 5/6/05
      If SupOnly = "Y" And PersRec.DISCOV = "N" Then GoTo SupOnlyNGSkip4 'added 10/18/05
'      If PersRec.LastYrPrinted = WhatYear Then Discovery$ = "Y" 'took out 12/5/06
      If (PersRec.LastYrPrinted <> WhatYear) Or (PersRec.DISCOV = "Y") Or (PersRec.LastYrPrinted = WhatYear) Then 'don't need this line
        If SupOnly = "Y" And PersRec.DISCOV <> "Y" Then
          GoTo SupOnlyNGSkip4
        End If
'        ThisPersValue# = PersValue# 'added for 2.05 '8/24/06
'        PersValue# = OldRound#(PersRec.PersVal + PersRec.MHVALUE + PersRec.MCVALUE) 'remmed out 8/24/06
'        PersValue# = OldRound#(PersValue# + PersRec.CVALUE + PersRec.MTVALUE)'remmed out 8/24/06
        PersValue# = OldRound#(PersRec.PersVal + PersRec.MHVALUE + PersRec.MCVALUE + PersRec.CVALUE + PersRec.MTVALUE) 'added PersRec.PersVal 0n 8/24/06
        ThisPersValue# = ThisPersValue# + PersValue# 'added 8/24/06
        'ThisPersExmp# = PersExmp# 'added for 2.05
        PersExmp# = OldRound#(PersRec.EXMPSENI + PersRec.EXMPOTHR)
        ThisPersExmp# = PersExmp# 'added for 2.05
        If PersRec.EXMPSENI > 0 Then
          TempAddOn.OldAmt = 0
          TempAddOn.CustName = QPTrim$(TaxCust.CustName)
          TempAddOn.CustRec = x
          TempAddOn.Type = "Senior discount of " + QPTrim$(Using$("$#,##0.00", PersRec.EXMPSENI)) + " applied to personal property tax."
          AddOnCnt = AddOnCnt + 1
          TempAddOn.NewAmt = PersRec.EXMPSENI
          Put AOHandle, AddOnCnt, TempAddOn
        End If
        If PersRec.EXMPOTHR > 0 Then
          TempAddOn.OldAmt = 0
          TempAddOn.CustName = QPTrim$(TaxCust.CustName)
          TempAddOn.CustRec = x
          TempAddOn.Type = "Other discount of " + QPTrim$(Using$("$#,##0.00", PersRec.EXMPOTHR)) + " applied to personal property tax."
          AddOnCnt = AddOnCnt + 1
          TempAddOn.NewAmt = PersRec.EXMPOTHR
          Put AOHandle, AddOnCnt, TempAddOn
        End If
        PersCalcVal# = OldRound#((PersValue# - PersExmp#) / 100)
      Else
        PersCalcVal# = 0
      End If
      If PersCalcVal# < 0 Then
        PersCalcVal# = 0
      End If
      PersTaxDue# = OldRound#(PersTaxDue# + PersCalcVal# * PERSRATE#) 'added inside brackets -> PersTaxDue...for 2.05
      If PersTaxDue# < 0 Then PersTaxDue# = 0
      TBillRec.CustPin = TaxCust.PIN
      TBillRec.InternalPin = PersRec.InternalPin
      TBillRec.PersPin = QPTrim$(PersRec.PropPin)
      'TBillRec.ExptValue = PersExmp# + ThisPersExmp# 'added ThisPersExmp for 2.05
      TBillRec.ExptValue = TBillRec.ExptValue + ThisPersExmp#
'      TBillRec.PersValue = PersValue# + ThisPersValue# 'added ThisPersValue# for 2.05
      TBillRec.PersValue = ThisPersValue# 'added 0n 08/24/06
      TBillRec.PersTaxDue = PersTaxDue#
      TBillRec.PersPropRecord = TaxCust.FirstPersRec
      TBillRec.PersTaxRate = PERSRATE#
      TBillRec.RealTaxRate = REALRATE#
      TBillRec.MORTCODE = ""
      TBillRec.MortRec = 0
'      TBillRec.TotalBillDue = OldRound#(TBillRec.TotalBillDue + TBillRec.PersTaxDue) 'remmed 8/24/06
      TBillRec.TotalBillDue = OldRound#((TBillRec.PersTaxDue) + LateAmt)  '08/24/06 'added LateAmt on 8/9/07
      If TBillRec.TotalBillDue > 0 Then
        GotPersVal = -1
      End If
      If PersRec.LateList = "Y" Then
'        LateAmt# = OldRound#((LateAmt# + TBillRec.PersTaxDue) * (LateList# / 100))
        LateAmt# = OldRound#(LateAmt# + (PersCalcVal# * PERSRATE#) * (LateList# / 100)) ' TBillRec.PersTaxDue) * (LateList# / 100))
        TBillRec.LateTaxDue = LateAmt#
'        TBillRec.TotalBillDue = OldRound#(TBillRec.PersTaxDue + LateAmt#)
        TBillRec.TotalBillDue = OldRound#(PersTaxDue# + LateAmt#)
      End If
SupOnlyNGSkip4:
      NextPersRec& = PersRec.NextRec
    Loop While NextPersRec& > 0
  End If
  
  Return

GetRealInfo:
  If SplitYN = "Pers" Then Return 'added 5/12/05
  ThisProp& = TaxCust.FirstPropRec
  If ThisProp& > 0 Then
    Get RRHandle, ThisProp&, RealRec
    If RealRec.Deleted = True Then GoTo SupOnlyNGSkip 'added 5/6/05
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
    
'    If RealRec.LastYrPrinted = WhatYear Then Discovery$ = "Y" 'took out 12/5/06
    If SupOnly = "Y" And RealRec.PROPDISC = "N" Then GoTo SupOnlyNGSkip 'added 10/18/05
    If (RealRec.LastYrPrinted <> WhatYear) Or (RealRec.PROPDISC = "Y") Or (RealRec.LastYrPrinted = WhatYear) Then
      If SupOnly = "Y" And RealRec.PROPDISC <> "Y" Then
        GoTo SupOnlyNGSkip
      End If
      RealValue# = RealRec.PROPVALU
      OptRevTax1# = FigureOptRevTax1(ThisProp&, RRHandle)
      TBillRec.OptRevTax1 = OptRevTax1#
      TotOpt1 = OldRound(TotOpt1 + OptRevTax1#)
      OptRevTax2# = FigureOptRevTax2(ThisProp&, RRHandle)
      TBillRec.OptRevTax2 = OptRevTax2#
      TotOpt2 = OldRound(TotOpt2 + OptRevTax2#)
      OptRevTax3# = FigureOptRevTax3(ThisProp&, RRHandle)
      TBillRec.OptRevTax3 = OptRevTax3#
      TotOpt3 = OldRound(TotOpt3 + OptRevTax3#)
      RealExmp# = OldRound#(RealRec.EXMPSENI + RealRec.EXMPOTHR)
      If RealRec.EXMPSENI > 0 Then
        TempAddOn.OldAmt = 0
        TempAddOn.CustName = QPTrim$(TaxCust.CustName)
        TempAddOn.CustRec = x
        TempAddOn.Type = "Senior discount of " + QPTrim$(Using$("$#,##0.00", RealRec.EXMPSENI)) + " applied to real estate tax."
        AddOnCnt = AddOnCnt + 1
        TempAddOn.NewAmt = RealRec.EXMPSENI
        Put AOHandle, AddOnCnt, TempAddOn
      End If
      If RealRec.EXMPOTHR > 0 Then
        TempAddOn.OldAmt = 0
        TempAddOn.CustName = QPTrim$(TaxCust.CustName)
        TempAddOn.CustRec = x
        TempAddOn.Type = "Other discount of " + QPTrim$(Using$("$#,##0.00", RealRec.EXMPOTHR)) + " applied to real estate tax."
        AddOnCnt = AddOnCnt + 1
        TempAddOn.NewAmt = RealRec.EXMPOTHR
        Put AOHandle, AddOnCnt, TempAddOn
      End If
      RealCalcVal# = OldRound#((RealValue# - RealExmp#) / 100)
      If RealCalcVal# < 0 Then RealCalcVal# = 0
      RealTaxDue# = OldRound#((RealCalcVal# * REALRATE#) + OptRevTax1# + OptRevTax2# + OptRevTax3#)
      
      TBillRec.ExptValue = OldRound#(TBillRec.ExptValue + RealExmp#)
      TBillRec.RealValue = RealValue#
      TotReal = TotReal + RealTaxDue#
      TBillRec.RealTaxDue = RealTaxDue#
      TBillRec.RealPropRecord = ThisProp&
      TBillRec.RealTaxRate = REALRATE#
      TBillRec.RDesc1 = QPTrim$(RealRec.PROPNOT1) + " " + QPTrim$(RealRec.PROPNOT2)
      'TBillRec(1).RDesc2 = RealRec(1).LOTNUMB
      MAPBLKLOT$ = RealRec.Map + " " + RealRec.BLOCK + " " + RealRec.LOTNUMB

      TBillRec.RDesc2 = MAPBLKLOT$
      TBillRec.CustPin = TaxCust.PIN
      TBillRec.InternalPin = RealRec.InternalPin
      TBillRec.RealPin = QPTrim$(RealRec.RealPin)
      TBillRec.LASize = CStr(RealRec.PropSize)
      TBillRec.LotOrAcre = QPTrim$(RealRec.LOTACRE)
      If RealRec.LateList = "Y" Then
        LateAmt# = OldRound#((RealTaxDue#) * (LateList# / 100))
      Else
        LateAmt# = 0
      End If
      TBillRec.LateTaxDue = OldRound#(TBillRec.LateTaxDue + LateAmt#)
      TBillRec.TotalBillDue = OldRound#(TBillRec.TotalBillDue + RealTaxDue# + LateAmt#)
      '****** put the first rec to disk
      GoSub WriteIt2Disk
    End If      'End of Test For Current Year Tax Bill

SupOnlyNGSkip:
    NextRealRec& = RealRec.NextRec
    If NextRealRec& > 0 Then
      Do
        If TBillRec.PersTaxDue > 0 And TBillRec.RealTaxDue = 0 Then '10/5/07
          GoTo Skip
        ElseIf TBillRec.PersTaxDue > 0 And TBillRec.RealTaxDue > 0 Then
          PersTaxDue = 0
        End If
        TBillRec = NewTBillRec        'make a new empty record each time
        GoSub SetCustInfo
Skip:
        Get RRHandle, NextRealRec&, RealRec
        ThisProp& = NextRealRec&
        If RealRec.Deleted = True Then GoTo SupOnlyNGSkip2 'added 5/6/05
        If SupOnly = "Y" And RealRec.PROPDISC = "N" Then GoTo SupOnlyNGSkip2 'added 10/18/05
        If RealRec.Mock = "Y" Then GoTo SupOnlyNGSkip2
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
'        If RealRec.LastYrPrinted = WhatYear Then Discovery$ = "Y" 'took out 12/5/06
        If (RealRec.LastYrPrinted <> WhatYear) Or (RealRec.PROPDISC = "Y") Or (RealRec.LastYrPrinted = WhatYear) Then 'don't need this line
          If SupOnly = "Y" And RealRec.PROPDISC <> "Y" Then
            GoTo SupOnlyNGSkip2
          End If
          RealValue# = RealRec.PROPVALU
          RealExmp# = OldRound#(RealRec.EXMPSENI + RealRec.EXMPOTHR)
          OptRevTax1# = FigureOptRevTax1(ThisProp&, RRHandle)
          TBillRec.OptRevTax1 = OptRevTax1#
          TotOpt1 = OldRound(TotOpt1 + OptRevTax1#)
          OptRevTax2# = FigureOptRevTax2(ThisProp&, RRHandle)
          TBillRec.OptRevTax2 = OptRevTax2#
          TotOpt2 = OldRound(TotOpt2 + OptRevTax2#)
          OptRevTax3# = FigureOptRevTax3(ThisProp&, RRHandle)
          TBillRec.OptRevTax3 = OptRevTax3#
          TotOpt3 = OldRound(TotOpt3 + OptRevTax3#)
          If RealRec.EXMPSENI > 0 Then
            TempAddOn.OldAmt = 0
            TempAddOn.CustName = QPTrim$(TaxCust.CustName)
            TempAddOn.CustRec = x
            TempAddOn.Type = "Senior discount of " + QPTrim$(Using$("$#,##0.00", RealRec.EXMPSENI)) + " applied to real estate tax."
            AddOnCnt = AddOnCnt + 1
            TempAddOn.NewAmt = RealRec.EXMPSENI
            Put AOHandle, AddOnCnt, TempAddOn
          End If
          If RealRec.EXMPOTHR > 0 Then
            TempAddOn.OldAmt = 0
            TempAddOn.CustName = QPTrim$(TaxCust.CustName)
            TempAddOn.CustRec = x
            TempAddOn.Type = "Other discount of " + QPTrim$(Using$("$#,##0.00", RealRec.EXMPOTHR)) + " applied to real estate tax."
            AddOnCnt = AddOnCnt + 1
            TempAddOn.NewAmt = RealRec.EXMPOTHR
            Put AOHandle, AddOnCnt, TempAddOn
          End If
          RealCalcVal# = OldRound#((RealValue# - RealExmp#) / 100)
          If RealCalcVal# < 0 Then RealCalcVal# = 0
          RealTaxDue# = OldRound#((RealCalcVal# * REALRATE#) + OptRevTax1# + OptRevTax2# + OptRevTax3#)
          TBillRec.ExptValue = RealExmp#
          TBillRec.RealValue = RealValue#
          TBillRec.RealTaxDue = RealTaxDue#
          TBillRec.RealPropRecord = NextRealRec&
          TBillRec.RealTaxRate = REALRATE#
          TBillRec.RDesc1 = RealRec.PROPNOT1
          TBillRec.RDesc2 = RealRec.PROPNOT2
          TBillRec.CustPin = TaxCust.PIN
          TBillRec.InternalPin = RealRec.InternalPin
          TBillRec.RealPin = QPTrim$(RealRec.RealPin)
          TBillRec.LASize = CStr(RealRec.PropSize)
          TBillRec.LotOrAcre = QPTrim$(RealRec.LOTACRE)
          If RealRec.LateList = "Y" Then
            LateAmt# = OldRound#(RealTaxDue# * (LateList# / 100))
            TBillRec.LateTaxDue = LateAmt#
          Else
            LateAmt# = 0
          End If
          TBillRec.TotalBillDue = OldRound#(RealTaxDue# + LateAmt# + PersTaxDue#) '10/5/07 added PersTaxDue
          TBillRec.MORTCODE = QPTrim$(RealRec.MORTCODE)
          GoSub WriteIt2Disk
        End If

'020501
SupOnlyNGSkip2:
        NextRealRec& = RealRec.NextRec
      Loop While NextRealRec& > 0
    End If
  Else
    GoSub WriteIt2Disk
  End If

'testing
  If GotPersVal And (Done2Disk = False) Then
    GoSub WriteIt2Disk
  End If
  
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPrebilling", "PrintGraphics", Erl)
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
  
  'on error goto ERRORSTUFF
  
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPrebilling", "ParseTownships", Erl)
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

Private Function FigureOptRevTax1(RealRecNum As Long, RRHandle As Integer) As Double
  Dim RealRec As PropertyRecType
  Dim RateRec As OptRevRateTablesType
  Dim RateHandle As Integer
  Dim NumOfRRREcs As Integer
  Dim x As Integer, y As Integer
  Dim ThisTax As Double
  Dim ThisPropVal#
  
  'on error goto ERRORSTUFF
  
  ThisTax = 0
  FigureOptRevTax1 = 0
  Get RRHandle, RealRecNum, RealRec
  If RealRec.OptRev1Chrg = 0 Then Exit Function
  OpenTaxRateTables RateHandle, NumOfRRREcs
  For x = 1 To NumOfRRREcs
    Get RateHandle, x, RateRec
    If RateRec.Deleted = True Then GoTo Deleted
    ThisPropVal = OldRound(RealRec.PROPVALU - (RealRec.EXMPOTHR + RealRec.EXMPSENI))
    If RealRec.OptRev1Chrg = x Then
      If RateRec.Type = "F" Then
        ThisTax = RateRec.FlatAmt
      ElseIf RateRec.Type = "S" Then
        For y = 1 To 10
          If ThisPropVal >= RateRec.FromAmt(y) And ThisPropVal <= RateRec.ToAmt(y) Then
            ThisTax = RateRec.TaxFAmt(y)
            Exit For
          End If
        Next y
        If y < 11 Then
          Exit For
        End If
      ElseIf RateRec.Type = "P" Then
        For y = 1 To 10
          If ThisPropVal >= RateRec.FromAmt(y) And ThisPropVal <= RateRec.ToAmt(y) Then
            ThisTax = OldRound(ThisPropVal * RateRec.TaxPAmt(y) / 100)
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
  Close RateHandle
  FigureOptRevTax1 = ThisTax
  
  Exit Function

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPrebilling", "FigureOptRevTax1", Erl)
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

Private Function FigureOptRevTax2(RealRecNum As Long, RRHandle As Integer) As Double
  Dim RealRec As PropertyRecType
  Dim RateRec As OptRevRateTablesType
  Dim RateHandle As Integer
  Dim NumOfRRREcs As Integer
  Dim x As Integer, y As Integer
  Dim ThisTax As Double
  Dim ThisPropVal#
  
  'on error goto ERRORSTUFF
  
  ThisTax = 0
  FigureOptRevTax2 = 0
  Get RRHandle, RealRecNum, RealRec
  OpenTaxRateTables RateHandle, NumOfRRREcs
  For x = 1 To NumOfRRREcs
    Get RateHandle, x, RateRec
    If RateRec.Deleted = True Then GoTo Deleted
    ThisPropVal = OldRound(RealRec.PROPVALU - (RealRec.EXMPOTHR + RealRec.EXMPSENI))
    If RealRec.OptRev2Chrg = x Then
      If RateRec.Type = "F" Then
        ThisTax = RateRec.FlatAmt
      ElseIf RateRec.Type = "S" Then
        For y = 1 To 10
          If ThisPropVal >= RateRec.FromAmt(y) And ThisPropVal <= RateRec.ToAmt(y) Then
            ThisTax = RateRec.TaxFAmt(y)
            Exit For
          End If
        Next y
        If y < 11 Then
          Exit For
        End If
      ElseIf RateRec.Type = "P" Then
        For y = 1 To 10
          If ThisPropVal >= RateRec.FromAmt(y) And ThisPropVal <= RateRec.ToAmt(y) Then
            ThisTax = OldRound(ThisPropVal * RateRec.TaxPAmt(y) / 100)
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
  Close RateHandle
  FigureOptRevTax2 = ThisTax
  
  Exit Function

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPrebilling", "FigureOptRevTax2", Erl)
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

Private Function FigureOptRevTax3(RealRecNum As Long, RRHandle As Integer) As Double
  Dim RealRec As PropertyRecType
  Dim RateRec As OptRevRateTablesType
  Dim RateHandle As Integer
  Dim NumOfRRREcs As Integer
  Dim x As Integer, y As Integer
  Dim ThisTax As Double
  Dim ThisPropVal#
  
  'on error goto ERRORSTUFF
  
  ThisTax = 0
  FigureOptRevTax3 = 0
  Get RRHandle, RealRecNum, RealRec
  OpenTaxRateTables RateHandle, NumOfRRREcs
  For x = 1 To NumOfRRREcs
    Get RateHandle, x, RateRec
    If RateRec.Deleted = True Then GoTo Deleted
    ThisPropVal = OldRound(RealRec.PROPVALU - (RealRec.EXMPOTHR + RealRec.EXMPSENI))
    If RealRec.OptRev3Chrg = x Then
      If RateRec.Type = "F" Then
        ThisTax = RateRec.FlatAmt
      ElseIf RateRec.Type = "S" Then
        For y = 1 To 10
          If ThisPropVal >= RateRec.FromAmt(y) And ThisPropVal <= RateRec.ToAmt(y) Then
            ThisTax = RateRec.TaxFAmt(y)
            Exit For
          End If
        Next y
        If y < 11 Then
          Exit For
        End If
      ElseIf RateRec.Type = "P" Then
        For y = 1 To 10
          If ThisPropVal >= RateRec.FromAmt(y) And ThisPropVal <= RateRec.ToAmt(y) Then
            ThisTax = OldRound(ThisPropVal * RateRec.TaxPAmt(y) / 100)
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
  Close RateHandle
  FigureOptRevTax3 = ThisTax
  
  Exit Function
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPrebilling", "FigureOptRevTax3", Erl)
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
  Dim TBillRec As TaxBillType
  Dim TBHandle As Integer
  Dim NumOfTBRecs As Long
  Dim OPBillRec As TaxTransactionType
  Dim OPHandle As Integer
  Dim NumOfOPRecs As Long
  Dim RealRec As PropertyRecType
  Dim RRHandle As Integer
  Dim NumOfRRREcs As Long
  Dim PersRec As PersonalRecType
  Dim PRHandle As Integer
  Dim NumOfPRRecs As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  Dim NewTBillRec As TaxBillType
  Dim Done2Disk As Integer
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
  Dim GotPersVal As Integer
  Dim PersExmp#
  Dim PersValue#
  Dim Discovery$
  Dim PersCalcVal#
  Dim PersTaxDue#
  Dim NextPersRec&
  Dim TotalPers#
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
  Dim ThisPersExmp As Double
  Dim ThisPersValue As Double
  Dim Pct1#, Pct2#, Pct6#, Pct7#, Pct8#, Pct9#
  Dim PctTot#, PctTest#, ThisReal#, ThisPers#
  
  'on error goto ERRORSTUFF
  
  Line$ = String(80, "-")
  Dot$ = String(75, ".")
  Line1$ = String(75, "_")
  FF$ = Chr(12)
  MaxLines = 58
  LineCnt = 0
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  Town$ = QPTrim$(TaxMasterRec.Name)
  TotOpt1 = 0
  TotOpt2 = 0
  TotOpt3 = 0
  OptRev1Desc = QPTrim$(TaxMasterRec.OptRev1)
  OptRev2Desc = QPTrim$(TaxMasterRec.OptRev2)
  OptRev3Desc = QPTrim$(TaxMasterRec.OptRev3)
  
  fpcmbCycle.Col = 1
  CycleName = QPTrim$(fpcmbCycle.ColText)
  If fpcmbCycle.Enabled = True And CycleName <> "NO CYCLE" Then
    fpcmbCycle.Col = 0
    CycleNum = CLng(fpcmbCycle.ColText)
  Else
    CycleNum = -1
  End If
  
  fpcmbCounty.Col = 1
  CountyName = QPTrim$(fpcmbCounty.ColText)
  If fpcmbCounty.Enabled = True And CountyName <> "ALL COUNTIES" Then
    fpcmbCounty.Col = 0
    CountyNum = CLng(fpcmbCounty.ColText)
  Else
    CountyNum = -1
  End If
  
  If fpcmbSplit.Enabled = True Then
    If fpcmbSplit.Text = "NO SPLIT" Then
      SplitYN = "No"
    ElseIf fpcmbSplit.Text = "REAL ONLY" Then
      SplitYN = "Real"
    ElseIf fpcmbSplit.Text = "PERSONAL ONLY" Then
      SplitYN = "Pers"
    Else
      SplitYN = "No"
    End If
  Else
    SplitYN = "No"
  End If
  
  If InStr(fpcmbTownships.Text, "ALL") Then
    ThisTownship = "ALL"
  Else
    ThisTownship = QPTrim$(fpcmbTownships.Text)
    ThisTownship = Mid(ThisTownship, 1, Len(ThisTownship) - 4)
    ThisTownship = QPTrim$(ThisTownship)
  End If
  dlm$ = "~"
  If Exist(TaxBillFile) Then
    KillFile TaxBillFile
  End If
  If Exist(TaxBillOPFile) Then
    KillFile TaxBillOPFile
  End If
  If Exist("TMPBLADD") Then 'tax bill addon
    KillFile "TMPBLADD"
  End If
  
  AddOnCnt = 0
  OpenTaxBillAddOn AOHandle
  OpenTaxBillFile TBHandle, NumOfTBRecs
  OpenTaxPropFile RRHandle, NumOfRRREcs
  OpenTaxPersFile PRHandle, NumOfPRRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenMortCodeFile MCHandle, NumOfMCRecs
  
  Inactive = 0
  
  frmTaxShowPctComp.Label1 = "Creating Tax Pre-Billing Register"
  frmTaxShowPctComp.Show , Me
  frmTaxShowPctComp.cmdCancel.Visible = False
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
    If TaxCust.TaxExempt = "Y" Then
      GoTo PreBillSkip
    End If
 ' If x = 1602 Then Stop
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
    
    Done2Disk = 0
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
      OverPayAmt = GetCustBalance(CustArr(x), -1)
      If OverPayAmt < 0 Then
        OverPay = True
      Else
        OverPayAmt = 0
        OverPay = False
      End If
    Else
      OverPayAmt = GetCustBalance(x, -1)
      If OverPayAmt < 0 Then
        OverPay = True
      Else
        OverPayAmt = 0
        OverPay = False
      End If
    End If
    ThisTest = CStr(OverPayAmt)
    If InStr(ThisTest, "E") Then OverPayAmt = 0

    If TaxCust.FirstPersRec <= 0 And TaxCust.FirstPropRec <= 0 Then
      NoProp = 1
      GoSub SetCustInfo
      GoSub WriteIt2Disk
      GoTo PreBillSkip
    End If
    
    NoProp = 0
    GoSub SetCustInfo
    GoSub GetPersInfo
    GoSub GetRealInfo
    If TBillRec.TotalBillDue = 0 Then
      If PastFlagSet = True Then
      End If
    End If
    BillCnt = BillCnt + 1
PreBillSkip:
    frmTaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      Exit Sub
    End If
  Next x
  
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  
  Close RRHandle
  Close PRHandle
  Close TCHandle
  Close MCHandle
  Close AOHandle
  
  If ThisTBCnt = 0 Then
    Call TaxMsg(900, "Using the parameters selected there are no customers who qualify for a tax charge.")
    Close
    fptxtCurrYear.SetFocus
    Exit Sub
  End If
  
  TotalPers# = 0
  TotalReal# = 0
  TotalEx# = 0
  NumBills& = 0
  
  TaxPreRptFile$ = "TAXRPTS\TaxPreBill.RPT"
  
  RptHandle = FreeFile
  Open TaxPreRptFile For Output As #RptHandle
  GoSub PreBillHeading
  Close TBHandle
  OpenTaxBillFile TBHandle, NumOfTBRecs
  
  For x = 1 To NumOfTBRecs
    Get TBHandle, x, TBillRec
'    If Discovery$ = "Y" And TBillRec.BillNumber = -1 Then
    If TBillRec.BillNumber = -1 And TBillRec.ExptValue = 0 Then  'took out Discovery$ = "Y" on 12/5/06
       GoTo NotThisOne
    Else
      If TBillRec.TotalBillDue = 0 And TBillRec.PriorYrBalance = 0 And TBillRec.ExptValue = 0 Then
        If OldRound(TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3) = 0 Then 'added 10/18/06
          If fpcmbInclNoBills.Text = "N" Then 'GoTo NextOne
            GoTo NotThisOne
          End If
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
      Print #RptHandle, Tab(69); Using("$###,##0.00", TBillRec.TotalBillDue - TBillRec.LateTaxDue) 'added late tax on 8/14/06
      Print #RptHandle, Using("$###,###,##0", TBillRec.RealValue);
      Print #RptHandle, Tab(15); Using("$###,###,##0.00", TBillRec.PersValue);
      Print #RptHandle, Tab(30); Using("$###,###,##0.00", TBillRec.ExptValue);
      TValue# = OldRound#(TBillRec.RealValue + TBillRec.PersValue - TBillRec.ExptValue)
      If TValue# < 0 Then TValue# = 0
      Print #RptHandle, Tab(45); Using("$###,###,##0.00", TValue#);
      If TBillRec.LateTaxDue > 0 Then
        Print #RptHandle, Tab(58); "Late = "; Using("###0.00", TBillRec.LateTaxDue)
      Else
        Print #RptHandle, ""
      End If
      LineCnt = LineCnt + 2
      If TBillRec.OverPayAmt > 0 Then
        Print #RptHandle, Tab(5); "Credit Balance Applied to Tax Due: "; Tab(69); Using("$###,##0.00", -TBillRec.OverPayAmt)
        Print #RptHandle, Tab(5); "Total Tax Due: "; Tab(69); Using$("$###,##0.00", (TBillRec.TotalBillDue - TBillRec.OverPayAmt))
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
      TotalPers# = OldRound#(TotalPers# + TBillRec.PersValue)
      TotalReal# = OldRound#(TotalReal# + TBillRec.RealValue)
      TotalEx# = OldRound#(TotalEx# + TBillRec.ExptValue)
'      TotalBills# = OldRound#(TotalBills# + TBillRec.TotalBillDue)
      TotalBills# = OldRound#(TotalBills# + TBillRec.PersTaxDue + TBillRec.RealTaxDue)
      TotalLate# = OldRound#(TotalLate# + TBillRec.LateTaxDue)
'      TotalPast# = OldRound#(TotalPast# + TBillRec.PriorYrBalance)
      If TBillRec.PriorYrBalance > 0 Then 'added 8/14/06
        TotalPast# = OldRound#(TotalPast# + TBillRec.PriorYrBalance)
      Else 'added 8/14/06
'        TotalPast# = 0 'added 8/14/06
      End If 'added 8/14/06
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
  Print #RptHandle, Tab(20); "Property Tax Billing : Pre-Billing Register"
  Print #RptHandle,
  Print #RptHandle, "Date: "; CStr(Date); Tab(65); "Page #"; Page
  Print #RptHandle, Line
  Print #RptHandle, "Number of Bills to Process: "; Using("###,##0", NumBills&)
  Print #RptHandle, "      Total Real Valuation: "; Using("$###,###,##0.00", TotalReal#)
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
  Print #RptHandle, "      Total Pers Valuation: "; Using("$###,###,##0.00", TotalPers#)
  Print #RptHandle, "          Total Exemptions: "; Using("$###,###,##0.00", TotalEx#)
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
'  Print #RptHandle, "         Total Late Amount: "; Using("$###,###,##0.00", TotalLate#)
  Print #RptHandle, "          Inactive Skipped: "; Using("###,##0", Inactive)
  Print #RptHandle, Chr$(12)
  Close
  
  ViewPrint TaxPreRptFile, "Tax Pre-Billing Report", True
  
  KillFile "txblsprn.dat"
  
  Exit Sub
  
SetCustInfo:
  TBillRec.CustRec = ThisTaxRec 'x              'cust acct rec
  CustName$ = QPTrim$(TaxCust.CustName)
  TBillRec.CustName = CustName$
  TBillRec.CustAdd1 = QPTrim$(TaxCust.Addr1)
  TBillRec.CustAdd2 = QPTrim$(TaxCust.Addr2)
  CitySt$ = QPTrim$(TaxCust.City) + " " + TaxCust.State
  TBillRec.CustAdd3 = CitySt$
  TBillRec.CustZip = TaxCust.Zip
  TBillRec.CustPin = TaxCust.PIN
  TBillRec.TaxYear = WhatYear
  TBillRec.RDesc3 = TaxCust.CSSN

  'Set Prior Balance if any
  ABalance = 0
  ABalance = GetCustBalance(TaxCust.Acct, -1) 'added this line on 8/11/2006    0
'  GoSub GetPastBalance
'  If ABalance# > 0 Then '8/14/06
    If PastFlagSet = 0 Then
      TBillRec.PriorYrBalance = ABalance#
    End If
    PastFlagSet = 1
'  End If '8/14/06
  Return
  
PreBillHeading:
  Dim ThisRate As Double
  Page = Page + 1
  Print #RptHandle, Tab(20); "Property Tax Billing : Pre-Billing Register"
  Print #RptHandle, Town
  Print #RptHandle, "Date: "; CStr(Date); Tab(73); "Page #"; CStr(Page)
  ThisRate = CDbl(fpDblSnglRealRate.Value)
  Print #RptHandle, "TownShip: " + ThisTownship; Tab(56); "Real Tax Rate: " + Using$("##0.0000", ThisRate) + "%"
  fpcmbCounty.Col = 1
  ThisRate = CDbl(fpDblSnglPersRate.Value)
  Print #RptHandle, "County: " + QPTrim$(fpcmbCounty.ColText); Tab(56); "Pers Tax Rate: " + Using$("##0.0000", ThisRate) + "%"
  fpcmbCycle.Col = 1
  ThisRate = CDbl(fpDblSnglLateList.Value)
  Print #RptHandle, "Cycle: " + QPTrim$(fpcmbCycle.ColText); Tab(55); "Late List Rate: " + Using$("##0.0000", ThisRate) + "%"
  Print #RptHandle, "* = Tax Due Does Not Include Prior Year Balance"
  Print #RptHandle,
  Print #RptHandle, "Acct #"; Tab(8); "Customer Name"; Tab(48); "Prior Yr Bal"; Tab(62); "Bill Seq#"; Tab(72); "*Tax Due"
  Print #RptHandle, "Real Value"; Tab(20); "Pers Value"; Tab(33); "Discnt Value"; Tab(47); "Net Valuation"
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
      If TaxTrans.TranType = 1 Then
        Balance# = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
        Balance# = OldRound(Balance# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection)
        Balance# = OldRound(Balance# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3) 'added for vs 2.05
        Balance# = OldRound(Balance# + TaxTrans.Revenue.PrePaidAmt) 'added for vs 2.05
        Balance# = OldRound(Balance# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
        Balance# = OldRound(Balance# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd))
        Balance# = OldRound(Balance# - (TaxTrans.Revenue.LateListPd + TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd))
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

  TBillRec.BillPrinted = False
  If TBillRec.TotalBillDue > 0 Then
'    If ShoresFlag Then
'      TBillRec(1).CarShore = 60
'      TStorm# = Round(TStorm# + TBillRec(1).CarShore)
'      StormCnt& = StormCnt& + 1
'      'TBillRec(1).TotalBillDue = Round#(TBillRec(1).TotalBillDue + TBillRec(1).CarShore)
'    End If
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
      TBillRec.SetDscvry2No = "Y" 'added 12/5/06
      TBillRec.RealTaxDue = 0 'added 12/5/06
      TBillRec.LateTaxDue = 0 'added 12/5/06
      TotOpt1# = OldRound(TotOpt1# - TBillRec.OptRevTax1) 'added 12/5/06
      TBillRec.OptRevTax1 = 0 'added 12/5/06
      TotOpt2# = OldRound(TotOpt2# - TBillRec.OptRevTax2) 'added 12/5/06
      TBillRec.OptRevTax2 = 0 'added 12/5/06
      TotOpt3# = OldRound(TotOpt3# - TBillRec.OptRevTax3) 'added 12/5/06
      TBillRec.OptRevTax3 = 0 'added 12/5/06
      TBillRec.PersTaxDue = 0 'added 12/4/06
      TotalLate# = OldRound(TotalLate# - TBillRec.LateTaxDue) 'added 12/4/06
      TBillRec.LateTaxDue = 0 'added 12/4/06
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
      Pct9# = 0
      ThisReal = OldRound(TBillRec.RealTaxDue - (TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3))
      ThisPers = OldRound(TBillRec.PersTaxDue)
      PctTot# = OldRound(TBillRec.TotalBillDue)
      If ThisReal > 0 Then
        Pct1# = OldRound#(ThisReal / PctTot#)
      End If
      If TBillRec.LateTaxDue > 0 Then
        Pct2# = OldRound#(TBillRec.LateTaxDue / PctTot#)
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
      If ThisPers > 0 Then
        Pct9# = OldRound#(TBillRec.PersTaxDue / PctTot#)
      End If
      
      ThisReal = MinBill * Pct1
      ThisPers = MinBill * Pct9
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
      
      PctTest# = OldRound(ThisReal + ThisPers + TBillRec.LateTaxDue)
      PctTest# = OldRound(PctTest# + TBillRec.OptRevTax1 + TBillRec.OptRevTax2 + TBillRec.OptRevTax3)
      
      If MinBill < PctTest# Then
        If ThisReal > PctTest# - MinBill Then
          ThisReal = OldRound(ThisReal - (PctTest# - MinBill))
          TotalReal# = OldRound(TotalReal# - (PctTest# - MinBill))
        ElseIf ThisPers > PctTest# - MinBill Then
          ThisPers = OldRound(ThisPers - (PctTest# - MinBill))
          TotalPers# = OldRound(TotalPers# - (PctTest# - MinBill))
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
        ElseIf ThisPers > MinBill - PctTest# Then
          ThisPers = OldRound(ThisPers + (MinBill - PctTest#))
          TotalPers# = OldRound(TotalPers# + (MinBill - PctTest#))
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
      TBillRec.PersTaxDue = ThisPers
      TBillRec.TotalBillDue = MinBill
    End If
  End If
  
'  If SYLVAFlag Then
'    If (TBillRec(1).TotalBillDue > 0) And (TBillRec(1).TotalBillDue < 5.01) Then
'      'L5 = L5 + 1
'      'STOP
'      'TBillRec(1).TotalBillDue = 0
'      TBillRec(1).BillNumber = -100
'    End If
'  End If
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
      OpenTaxBillOverPayFile OPHandle, NumOfOPRecs
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
      
      If TBillRec.PersTaxDue > 0 And Abs(OverPayAmt) > 0 Then
        If Abs(OverPayAmt) > TBillRec.PersTaxDue Then
          OverPayAmt = OldRound(OverPayAmt + TBillRec.PersTaxDue)
          OPBillRec.Revenue.Principle1Pd = TBillRec.PersTaxDue
          OPApplied = OldRound(OPApplied + OPBillRec.Revenue.Principle1Pd)
        ElseIf Abs(OverPayAmt) <= TBillRec.PersTaxDue Then
          OPBillRec.Revenue.Principle1Pd = -OverPayAmt
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
  Done2Disk = -1

  Return
  
GetPersInfo:
  If SplitYN = "Real" Then Return 'added 5/12/05
  
  GotPersVal = 0
  PersExmp# = 0
  PersValue# = 0
  PersCalcVal# = 0
  PersTaxDue# = 0 'added 7/3/2008
  ThisPersValue# = 0 'added 10/21/08
  
'  If ThisTaxRec = 3428 Then Stop
  
  If TaxCust.FirstPersRec > 0 Then
    Get PRHandle, TaxCust.FirstPersRec, PersRec
    If PersRec.Deleted = True Then GoTo SupOnlyNGSkip3 'added 5/6/05
'    If PersRec.LastYrPrinted = WhatYear Then Discovery$ = "Y" 'took out 12/5/06
    If SupOnly = "Y" And PersRec.DISCOV <> "Y" Then GoTo SupOnlyNGSkip3 'added 10/18/05
    If (PersRec.LastYrPrinted <> WhatYear) Or (PersRec.DISCOV = "Y") Or (PersRec.LastYrPrinted = WhatYear) Then
'      If SupOnly = "Y" And PersRec.DISCOV <> "Y" Then
'        GoTo SupOnlyNGSkip3
'      End If
    'IF PersRec(1).LastYrPrinted <> WhatYear THEN
      PersValue# = OldRound#(PersRec.PersVal + PersRec.MHVALUE + PersRec.MCVALUE)
      PersValue# = OldRound#(PersValue# + PersRec.CVALUE + PersRec.MTVALUE)
      PersExmp# = OldRound#(PersRec.EXMPSENI + PersRec.EXMPOTHR)
      If PersRec.EXMPSENI > 0 Then
        TempAddOn.OldAmt = 0
        TempAddOn.CustName = QPTrim$(TaxCust.CustName)
        TempAddOn.CustRec = x
        TempAddOn.Type = "Senior discount of " + QPTrim$(Using$("$#,##0.00", PersRec.EXMPSENI)) + " applied to personal property tax."
        AddOnCnt = AddOnCnt + 1
        TempAddOn.NewAmt = PersRec.EXMPSENI
        Put AOHandle, AddOnCnt, TempAddOn
      End If
      If PersRec.EXMPOTHR > 0 Then
        TempAddOn.OldAmt = 0
        TempAddOn.CustName = QPTrim$(TaxCust.CustName)
        TempAddOn.CustRec = x
        TempAddOn.Type = "Other discount of " + QPTrim$(Using$("$#,##0.00", PersRec.EXMPOTHR)) + " applied to personal property tax."
        AddOnCnt = AddOnCnt + 1
        TempAddOn.NewAmt = PersRec.EXMPOTHR
        Put AOHandle, AddOnCnt, TempAddOn
      End If
      PersCalcVal# = OldRound#((PersValue# - PersExmp#) / 100)
    Else
      PersCalcVal# = 0
    End If
    If PersCalcVal# < 0 Then
      PersCalcVal# = 0
    End If
  
    PersTaxDue# = OldRound#(PersCalcVal# * PERSRATE#) 'added inside brackets -> PersTaxDue#...for 2.05
    TBillRec.CustPin = TaxCust.PIN
    TBillRec.InternalPin = PersRec.InternalPin
    TBillRec.PersPin = QPTrim$(PersRec.PropPin)
    TBillRec.ExptValue = PersExmp#
    TBillRec.PersValue = PersValue#
    ThisPersValue# = PersValue# 'added for 2.05 8/24/06
    TBillRec.PersTaxDue = PersTaxDue#
    TBillRec.PersPropRecord = TaxCust.FirstPersRec
    TBillRec.PersTaxRate = PERSRATE#
    TBillRec.RealTaxRate = REALRATE#
    TBillRec.MORTCODE = ""
    TBillRec.MortRec = 0
    TBillRec.TotalBillDue = OldRound#(TBillRec.PersTaxDue)
    If TBillRec.TotalBillDue > 0 Then
      GotPersVal = -1
    End If
    If PersRec.LateList = "Y" Then
      LateAmt# = OldRound#((LateAmt# + TBillRec.PersTaxDue) * (LateList# / 100))
      TBillRec.LateTaxDue = LateAmt#
      TBillRec.TotalBillDue = OldRound#(TBillRec.PersTaxDue + LateAmt#)
    End If
  Else
    'no need to set the variables all are zero already
  End If

SupOnlyNGSkip3:
  NextPersRec& = PersRec.NextRec
  If NextPersRec& > 0 Then
    Do
'      TBillRec = NewTBillRec        'make a new empty record each time 'commented out 8/24/06
      GoSub SetCustInfo
      Get PRHandle, NextPersRec&, PersRec
      If PersRec.Deleted = True Then GoTo SupOnlyNGSkip4 'added 5/6/05
'      If PersRec.LastYrPrinted = WhatYear Then Discovery$ = "Y" 'took out 12/5/06
      If SupOnly = "Y" And PersRec.DISCOV <> "Y" Then GoTo SupOnlyNGSkip4 'added 10/18/05
      If (PersRec.LastYrPrinted <> WhatYear) Or (PersRec.DISCOV = "Y") Or (PersRec.LastYrPrinted = WhatYear) Then 'don't need this line
'        If SupOnly = "Y" And PersRec.DISCOV <> "Y" Then
'          GoTo SupOnlyNGSkip4
'        End If
'        ThisPersValue# = PersValue# 'added for 2.05 commented out 8/24/06
'        PersValue# = OldRound#(PersRec.PersVal + PersRec.MHVALUE + PersRec.MCVALUE)'commented out 8/24/06
'        PersValue# = OldRound#(PersValue# + PersRec.CVALUE + PersRec.MTVALUE)'commented out 8/24/06
        PersValue# = OldRound#(PersRec.PersVal + PersRec.MHVALUE + PersRec.MCVALUE + PersRec.CVALUE + PersRec.MTVALUE) 'added PersRec.PersVal 0n 8/24/06
        ThisPersValue# = ThisPersValue# + PersValue# 'added 8/14/08
        PersExmp# = OldRound#(PersRec.EXMPSENI + PersRec.EXMPOTHR)
        ThisPersExmp# = PersExmp# 'added for 2.05
        If PersRec.EXMPSENI > 0 Then
          TempAddOn.OldAmt = 0
          TempAddOn.CustName = QPTrim$(TaxCust.CustName)
          TempAddOn.CustRec = x
          TempAddOn.Type = "Senior discount of " + QPTrim$(Using$("$#,##0.00", PersRec.EXMPSENI)) + " applied to personal property tax."
          AddOnCnt = AddOnCnt + 1
          TempAddOn.NewAmt = PersRec.EXMPSENI
          Put AOHandle, AddOnCnt, TempAddOn
        End If
        If PersRec.EXMPOTHR > 0 Then
          TempAddOn.OldAmt = 0
          TempAddOn.CustName = QPTrim$(TaxCust.CustName)
          TempAddOn.CustRec = x
          TempAddOn.Type = "Other discount of " + QPTrim$(Using$("$#,##0.00", PersRec.EXMPOTHR)) + " applied to personal property tax."
          AddOnCnt = AddOnCnt + 1
          TempAddOn.NewAmt = PersRec.EXMPOTHR
          Put AOHandle, AddOnCnt, TempAddOn
        End If
        PersCalcVal# = OldRound#((PersValue# - PersExmp#) / 100)
      Else
        PersCalcVal# = 0
      End If
      If PersCalcVal# < 0 Then
        PersCalcVal# = 0
      End If
  
      PersTaxDue# = OldRound#(PersTaxDue# + PersCalcVal# * PERSRATE#)
      TBillRec.CustPin = TaxCust.PIN
      TBillRec.InternalPin = PersRec.InternalPin
      TBillRec.PersPin = QPTrim$(PersRec.PropPin)
'      TBillRec.ExptValue = PersExmp#
'      TBillRec.PersValue = PersValue#
      'TBillRec.ExptValue = PersExmp# + ThisPersExmp# 'added ThisPersExmp for 2.05
      TBillRec.ExptValue = TBillRec.ExptValue + ThisPersExmp# 'added ThisPersExmp for 2.05
      TBillRec.PersValue = PersValue# + ThisPersValue# 'added ThisPersValue# for 2.05 'commented out 8/24/06
      TBillRec.PersValue = ThisPersValue# 'added 0n 8/24/06
      TBillRec.PersTaxDue = PersTaxDue#
      TBillRec.PersPropRecord = TaxCust.FirstPersRec
      TBillRec.PersTaxRate = PERSRATE#
      TBillRec.RealTaxRate = REALRATE#
      TBillRec.MORTCODE = ""
      TBillRec.MortRec = 0
'      TBillRec.TotalBillDue = OldRound#(TBillRec.TotalBillDue + TBillRec.PersTaxDue)'added line below and commented out this line on 7/9/07
      TBillRec.TotalBillDue = OldRound#((TBillRec.PersTaxDue) + LateAmt)  '08/24/06 'added LateAmt on 8/9/07
      If TBillRec.TotalBillDue > 0 Then
        GotPersVal = -1
      End If
      If PersRec.LateList = "Y" Then
'        LateAmt# = OldRound#((LateAmt# + TBillRec.PersTaxDue) * (LateList# / 100))
        LateAmt# = OldRound#(LateAmt# + (PersCalcVal# * PERSRATE#) * (LateList# / 100)) '9.6.06 TBillRec.PersTaxDue) * (LateList# / 100))
        TBillRec.LateTaxDue = LateAmt#
        TBillRec.TotalBillDue = OldRound#(TBillRec.PersTaxDue + LateAmt#)
      End If
SupOnlyNGSkip4:
      NextPersRec& = PersRec.NextRec
    Loop While NextPersRec& > 0
  End If
  
  Return

GetRealInfo:
  If SplitYN = "Pers" Then Return 'added 5/12/05
  ThisProp& = TaxCust.FirstPropRec
  If ThisProp& > 0 Then
    Get RRHandle, ThisProp&, RealRec
    If RealRec.Deleted = True Then GoTo SupOnlyNGSkip 'added 5/6/05
    If SupOnly = "Y" And RealRec.PROPDISC <> "Y" Then GoTo SupOnlyNGSkip 'added 10/18/05
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
    
'    If RealRec.LastYrPrinted = WhatYear Then Discovery$ = "Y" '12/5/06
    If (RealRec.LastYrPrinted <> WhatYear) Or (RealRec.PROPDISC = "Y") Or (RealRec.LastYrPrinted = WhatYear) Then 'don't need this line
'      If SupOnly = "Y" And RealRec.PROPDISC <> "Y" Then
'        GoTo SupOnlyNGSkip
'      End If
      RealValue# = RealRec.PROPVALU
      OptRevTax1# = FigureOptRevTax1(ThisProp&, RRHandle)
      TBillRec.OptRevTax1 = OptRevTax1#
      TotOpt1 = OldRound(TotOpt1 + OptRevTax1#)
      OptRevTax2# = FigureOptRevTax2(ThisProp&, RRHandle)
      TBillRec.OptRevTax2 = OptRevTax2#
      TotOpt2 = OldRound(TotOpt2 + OptRevTax2#)
      OptRevTax3# = FigureOptRevTax3(ThisProp&, RRHandle)
      TBillRec.OptRevTax3 = OptRevTax3#
      TotOpt3 = OldRound(TotOpt3 + OptRevTax3#)
      RealExmp# = OldRound#(RealRec.EXMPSENI + RealRec.EXMPOTHR)
      If RealRec.EXMPSENI > 0 Then
        TempAddOn.OldAmt = 0
        TempAddOn.CustName = QPTrim$(TaxCust.CustName)
        TempAddOn.CustRec = x
        TempAddOn.Type = "Senior discount of " + QPTrim$(Using$("$#,##0.00", RealRec.EXMPSENI)) + " applied to real estate tax."
        AddOnCnt = AddOnCnt + 1
        TempAddOn.NewAmt = RealRec.EXMPSENI
        Put AOHandle, AddOnCnt, TempAddOn
      End If
      If RealRec.EXMPOTHR > 0 Then
        TempAddOn.OldAmt = 0
        TempAddOn.CustName = QPTrim$(TaxCust.CustName)
        TempAddOn.CustRec = x
        TempAddOn.Type = "Other discount of " + QPTrim$(Using$("$#,##0.00", RealRec.EXMPOTHR)) + " applied to real estate tax."
        AddOnCnt = AddOnCnt + 1
        TempAddOn.NewAmt = RealRec.EXMPOTHR
        Put AOHandle, AddOnCnt, TempAddOn
      End If
      
      RealCalcVal# = OldRound#((RealValue# - RealExmp#) / 100)
      If RealCalcVal# < 0 Then RealCalcVal# = 0
      RealTaxDue# = OldRound#((RealCalcVal# * REALRATE#) + OptRevTax1# + OptRevTax2# + OptRevTax3#)
      TBillRec.ExptValue = OldRound#(TBillRec.ExptValue + RealExmp#)
      TBillRec.RealValue = RealValue#
      TBillRec.RealTaxDue = RealTaxDue#
      TBillRec.RealPropRecord = ThisProp&
      TBillRec.RealTaxRate = REALRATE#
      TBillRec.RDesc1 = QPTrim$(RealRec.PROPNOT1) + " " + QPTrim$(RealRec.PROPNOT2)
      'TBillRec(1).RDesc2 = RealRec(1).LOTNUMB
      MAPBLKLOT$ = RealRec.Map + " " + RealRec.BLOCK + " " + RealRec.LOTNUMB

      TBillRec.RDesc2 = MAPBLKLOT$
      TBillRec.CustPin = TaxCust.PIN
      TBillRec.InternalPin = RealRec.InternalPin
      TBillRec.RealPin = RealRec.RealPin
      TBillRec.LASize = CStr(RealRec.PropSize)
      TBillRec.LotOrAcre = QPTrim$(RealRec.LOTACRE)
      If RealRec.LateList = "Y" Then
        LateAmt# = OldRound#((RealTaxDue#) * (LateList# / 100))
      Else
        LateAmt# = 0
      End If
      TBillRec.LateTaxDue = OldRound#(TBillRec.LateTaxDue + LateAmt#)
      TBillRec.TotalBillDue = OldRound#(TBillRec.TotalBillDue + RealTaxDue# + LateAmt#)
      '****** put the first rec to disk
      GoSub WriteIt2Disk
    End If      'End of Test For Current Year Tax Bill

SupOnlyNGSkip:
    NextRealRec& = RealRec.NextRec
    If NextRealRec& > 0 Then
      Do
        If TBillRec.PersTaxDue > 0 And TBillRec.RealTaxDue = 0 Then '10/5/07
          GoTo Skip
        ElseIf TBillRec.PersTaxDue > 0 And TBillRec.RealTaxDue > 0 Then
          PersTaxDue = 0
        End If
        TBillRec = NewTBillRec        'make a new empty record each time
        GoSub SetCustInfo
Skip:
        Get RRHandle, NextRealRec&, RealRec
        ThisProp& = NextRealRec&
        If RealRec.Deleted = True Then GoTo SupOnlyNGSkip2 'added 5/6/05
        If SupOnly = "Y" And RealRec.PROPDISC <> "Y" Then GoTo SupOnlyNGSkip2 'added 10/18/05
        If RealRec.Mock = "Y" Then GoTo SupOnlyNGSkip2
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
'        If RealRec.LastYrPrinted = WhatYear Then Discovery$ = "Y" '12/5/06
        If (RealRec.LastYrPrinted <> WhatYear) Or (RealRec.PROPDISC = "Y") Or (RealRec.LastYrPrinted = WhatYear) Then
'          If SupOnly = "Y" And RealRec.PROPDISC <> "Y" Then
'            GoTo SupOnlyNGSkip2
'          End If
          RealValue# = RealRec.PROPVALU
          RealExmp# = OldRound#(RealRec.EXMPSENI + RealRec.EXMPOTHR)
          OptRevTax1# = FigureOptRevTax1(ThisProp&, RRHandle)
          TBillRec.OptRevTax1 = OptRevTax1#
          TotOpt1 = OldRound(TotOpt1 + OptRevTax1#)
          OptRevTax2# = FigureOptRevTax2(ThisProp&, RRHandle)
          TBillRec.OptRevTax2 = OptRevTax2#
          TotOpt2 = OldRound(TotOpt2 + OptRevTax2#)
          OptRevTax3# = FigureOptRevTax3(ThisProp&, RRHandle)
          TBillRec.OptRevTax3 = OptRevTax3#
          TotOpt3 = OldRound(TotOpt3 + OptRevTax3#)
          If RealRec.EXMPSENI > 0 Then
            TempAddOn.OldAmt = 0
            TempAddOn.CustName = QPTrim$(TaxCust.CustName)
            TempAddOn.CustRec = x
            TempAddOn.Type = "Senior discount of " + QPTrim$(Using$("$#,##0.00", PersRec.EXMPSENI)) + " applied to real estate tax."
            AddOnCnt = AddOnCnt + 1
            TempAddOn.NewAmt = RealRec.EXMPSENI
            Put AOHandle, AddOnCnt, TempAddOn
          End If
          If RealRec.EXMPOTHR > 0 Then
            TempAddOn.OldAmt = 0
            TempAddOn.CustName = QPTrim$(TaxCust.CustName)
            TempAddOn.CustRec = x
            TempAddOn.Type = "Other discount of " + QPTrim$(Using$("$#,##0.00", PersRec.EXMPOTHR)) + " applied to real estate tax."
            AddOnCnt = AddOnCnt + 1
            TempAddOn.NewAmt = RealRec.EXMPOTHR
            Put AOHandle, AddOnCnt, TempAddOn
          End If
          RealCalcVal# = OldRound#((RealValue# - RealExmp#) / 100)
          If RealCalcVal# < 0 Then RealCalcVal# = 0
          RealTaxDue# = OldRound#((RealCalcVal# * REALRATE#) + OptRevTax1# + OptRevTax2# + OptRevTax3#)
          TBillRec.ExptValue = RealExmp#
          TBillRec.RealValue = RealValue#
          TBillRec.RealTaxDue = RealTaxDue#
          TBillRec.RealPropRecord = NextRealRec&
          TBillRec.RealTaxRate = REALRATE#
          TBillRec.RDesc1 = RealRec.PROPNOT1
          TBillRec.RDesc2 = RealRec.PROPNOT2
          TBillRec.CustPin = TaxCust.PIN
          TBillRec.InternalPin = RealRec.InternalPin
          TBillRec.RealPin = RealRec.RealPin
          TBillRec.LASize = CStr(RealRec.PropSize)
          TBillRec.LotOrAcre = QPTrim$(RealRec.LOTACRE)
          If RealRec.LateList = "Y" Then
            LateAmt# = OldRound#(RealTaxDue# * (LateList# / 100))
            TBillRec.LateTaxDue = LateAmt#
          Else
            LateAmt# = 0
          End If
          TBillRec.TotalBillDue = OldRound#(RealTaxDue# + LateAmt# + PersTaxDue#) '10/5/07 added PersTaxDue
          GoSub WriteIt2Disk
        End If

'020501
SupOnlyNGSkip2:
        NextRealRec& = RealRec.NextRec
      Loop While NextRealRec& > 0
    End If
  Else
    GoSub WriteIt2Disk
  End If

'testing
  If GotPersVal And (Done2Disk = False) Then
    GoSub WriteIt2Disk
  End If
  
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxPrebilling", "PrintText", Erl)
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
  
