VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVATaxRealProp 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Real Property Information"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbClass 
      Height          =   390
      Left            =   2715
      TabIndex        =   13
      Top             =   5400
      Width           =   3255
      _Version        =   196608
      _ExtentX        =   5741
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
      ColDesigner     =   "frmVATaxRealProp.frx":0000
   End
   Begin LpLib.fpCombo fpcmbMortCode 
      Height          =   390
      Left            =   8115
      TabIndex        =   23
      Top             =   5280
      Width           =   2220
      _Version        =   196608
      _ExtentX        =   3916
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
      ColDesigner     =   "frmVATaxRealProp.frx":0367
   End
   Begin LpLib.fpCombo fpcmbLateListYN 
      Height          =   390
      Left            =   10200
      TabIndex        =   21
      Top             =   4320
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
      ColDesigner     =   "frmVATaxRealProp.frx":06CE
   End
   Begin LpLib.fpCombo fpcmbDiscoveryYN 
      Height          =   390
      Left            =   10200
      TabIndex        =   19
      Top             =   3840
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
      ColDesigner     =   "frmVATaxRealProp.frx":0A35
   End
   Begin LpLib.fpCombo fpcmbLotAcre 
      Height          =   390
      Left            =   5145
      TabIndex        =   6
      Top             =   3480
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
      ColDesigner     =   "frmVATaxRealProp.frx":0D9C
   End
   Begin LpLib.fpCombo fpcmbTownship 
      Height          =   390
      Left            =   1515
      TabIndex        =   12
      Top             =   4920
      Width           =   4455
      _Version        =   196608
      _ExtentX        =   7858
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
      ColDesigner     =   "frmVATaxRealProp.frx":1103
   End
   Begin LpLib.fpCombo fpcmbLienYN 
      Height          =   390
      Left            =   7680
      TabIndex        =   20
      Top             =   4320
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
      ColDesigner     =   "frmVATaxRealProp.frx":146A
   End
   Begin LpLib.fpCombo fpcmbOptRev1 
      Height          =   360
      Left            =   8460
      TabIndex        =   24
      Top             =   6000
      Width           =   2220
      _Version        =   196608
      _ExtentX        =   3916
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
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmVATaxRealProp.frx":17D1
   End
   Begin LpLib.fpCombo fpcmbOptRev2 
      Height          =   360
      Left            =   8460
      TabIndex        =   25
      Top             =   6480
      Width           =   2220
      _Version        =   196608
      _ExtentX        =   3916
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
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmVATaxRealProp.frx":1B90
   End
   Begin LpLib.fpCombo fpcmbOptRev3 
      Height          =   360
      Left            =   8460
      TabIndex        =   26
      Top             =   6960
      Width           =   2220
      _Version        =   196608
      _ExtentX        =   3916
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
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmVATaxRealProp.frx":1F4F
   End
   Begin LpLib.fpCombo fpcmbGoToPins 
      Height          =   390
      Left            =   7920
      TabIndex        =   17
      Top             =   2280
      Width           =   2460
      _Version        =   196608
      _ExtentX        =   4339
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
      ColDesigner     =   "frmVATaxRealProp.frx":230E
   End
   Begin EditLib.fpDoubleSingle fpdblSize 
      Height          =   375
      Left            =   4470
      TabIndex        =   9
      Top             =   3960
      Width           =   1575
      _Version        =   196608
      _ExtentX        =   2773
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
      DecimalPlaces   =   -1
      DecimalPoint    =   "."
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
      NegFormat       =   1
      NegToggle       =   0   'False
      Separator       =   ","
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
   Begin EditLib.fpCurrency fpCurrRealVal 
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   3480
      Width           =   2055
      _Version        =   196608
      _ExtentX        =   3625
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpText fptxtThisCust 
      Height          =   390
      Left            =   2850
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   645
      Width           =   6015
      _Version        =   196608
      _ExtentX        =   10610
      _ExtentY        =   688
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
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
   Begin EditLib.fpDateTime fptxtDate 
      Height          =   372
      Left            =   9060
      TabIndex        =   1
      Top             =   1320
      Width           =   1752
      _Version        =   196608
      _ExtentX        =   3090
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
   Begin EditLib.fpText fptxtRecord 
      Height          =   396
      Left            =   2280
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2052
      _Version        =   196608
      _ExtentX        =   3619
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
      MaxLength       =   25
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
   Begin EditLib.fpText fptxtRealPin 
      Height          =   390
      Left            =   5760
      TabIndex        =   0
      Top             =   1320
      Width           =   2535
      _Version        =   196608
      _ExtentX        =   4471
      _ExtentY        =   688
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
      AutoAdvance     =   0   'False
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   20
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
   Begin EditLib.fpCurrency fpCurrBldgVal 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   3000
      Width           =   2055
      _Version        =   196608
      _ExtentX        =   3619
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
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpCurrency fpCurrOther 
      Height          =   375
      Left            =   9240
      TabIndex        =   18
      Top             =   3240
      Width           =   1335
      _Version        =   196608
      _ExtentX        =   2355
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   ""
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
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
   Begin EditLib.fpText fptxtDesc 
      Height          =   396
      Index           =   0
      Left            =   432
      TabIndex        =   14
      Top             =   6240
      Width           =   5532
      _Version        =   196608
      _ExtentX        =   9763
      _ExtentY        =   688
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      AlignTextH      =   0
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   31
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
   Begin EditLib.fpText fptxtDesc 
      Height          =   396
      Index           =   1
      Left            =   432
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   6600
      Width           =   5532
      _Version        =   196608
      _ExtentX        =   9763
      _ExtentY        =   688
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      AlignTextH      =   0
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   31
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   -1  'True
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
   Begin EditLib.fpText fptxtDesc 
      Height          =   396
      Index           =   2
      Left            =   432
      TabIndex        =   16
      Top             =   6960
      Width           =   5532
      _Version        =   196608
      _ExtentX        =   9763
      _ExtentY        =   688
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      AlignTextH      =   0
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   31
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
   Begin EditLib.fpText fptxtMap 
      Height          =   390
      Left            =   1035
      TabIndex        =   7
      Top             =   3960
      Width           =   855
      _Version        =   196608
      _ExtentX        =   1503
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   6
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
   Begin EditLib.fpText fptxtBlock 
      Height          =   390
      Left            =   1035
      TabIndex        =   10
      Top             =   4440
      Width           =   855
      _Version        =   196608
      _ExtentX        =   1503
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   6
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
   Begin EditLib.fpText fptxtLot 
      Height          =   390
      Left            =   2715
      TabIndex        =   8
      Top             =   3960
      Width           =   855
      _Version        =   196608
      _ExtentX        =   1503
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   6
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
   Begin EditLib.fpText fptxtLandGIS 
      Height          =   390
      Left            =   5640
      TabIndex        =   4
      Top             =   3000
      Width           =   1695
      _Version        =   196608
      _ExtentX        =   2990
      _ExtentY        =   688
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   20
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
   Begin EditLib.fpText fptxtPropAdd 
      Height          =   390
      Left            =   2760
      TabIndex        =   2
      Top             =   2145
      Width           =   4695
      _Version        =   196608
      _ExtentX        =   8281
      _ExtentY        =   688
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
   Begin EditLib.fpText fptxtLienDesc 
      Height          =   390
      Left            =   7680
      TabIndex        =   22
      Top             =   4800
      Width           =   3135
      _Version        =   196608
      _ExtentX        =   5530
      _ExtentY        =   688
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
   Begin EditLib.fpText fptxtImage 
      Height          =   390
      Left            =   3600
      TabIndex        =   11
      Top             =   4440
      Width           =   1575
      _Version        =   196608
      _ExtentX        =   2773
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   10
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
   Begin EditLib.fpText fptxtOptSearch 
      Height          =   390
      Left            =   8640
      TabIndex        =   27
      Top             =   7560
      Width           =   2295
      _Version        =   196608
      _ExtentX        =   4048
      _ExtentY        =   688
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
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
      AlignTextV      =   1
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   20
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtOptSearchDesc 
      Height          =   396
      Left            =   2520
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   7560
      Width           =   3612
      _Version        =   196608
      _ExtentX        =   6376
      _ExtentY        =   688
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
      AlignTextV      =   1
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   0   'False
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
      MaxLength       =   25
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   420
      Left            =   9720
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   8160
      Width           =   1452
      _Version        =   131072
      _ExtentX        =   2561
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmVATaxRealProp.frx":26CD
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   420
      Left            =   480
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   8160
      Width           =   1452
      _Version        =   131072
      _ExtentX        =   2561
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmVATaxRealProp.frx":28A9
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDelete 
      Height          =   420
      Left            =   2040
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   8160
      Width           =   1452
      _Version        =   131072
      _ExtentX        =   2561
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmVATaxRealProp.frx":2A85
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAdd 
      Height          =   420
      Left            =   8280
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   8160
      Width           =   1332
      _Version        =   131072
      _ExtentX        =   2350
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmVATaxRealProp.frx":2C62
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPageDown 
      Height          =   420
      Left            =   3600
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   8160
      Width           =   1452
      _Version        =   131072
      _ExtentX        =   2561
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmVATaxRealProp.frx":2E3C
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPageUp 
      Height          =   420
      Left            =   5160
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   8160
      Width           =   1452
      _Version        =   131072
      _ExtentX        =   2561
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmVATaxRealProp.frx":3018
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHist 
      Height          =   420
      Left            =   6720
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   8160
      Width           =   1452
      _Version        =   131072
      _ExtentX        =   2561
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmVATaxRealProp.frx":31F4
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdImage 
      Height          =   390
      Left            =   5280
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   4440
      Width           =   735
      _Version        =   131072
      _ExtentX        =   1296
      _ExtentY        =   688
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
      ButtonDesigner  =   "frmVATaxRealProp.frx":33D2
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdMortDetail 
      Height          =   390
      Left            =   10440
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   5280
      Width           =   630
      _Version        =   131072
      _ExtentX        =   1111
      _ExtentY        =   688
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
      ButtonDesigner  =   "frmVATaxRealProp.frx":35AA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDetail1 
      Height          =   405
      Left            =   10725
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   5970
      Width           =   630
      _Version        =   131072
      _ExtentX        =   1111
      _ExtentY        =   714
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
      ButtonDesigner  =   "frmVATaxRealProp.frx":3783
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDetail2 
      Height          =   405
      Left            =   10725
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   6450
      Width           =   630
      _Version        =   131072
      _ExtentX        =   1111
      _ExtentY        =   714
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
      ButtonDesigner  =   "frmVATaxRealProp.frx":395C
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDetail3 
      Height          =   405
      Left            =   10725
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   6930
      Width           =   630
      _Version        =   131072
      _ExtentX        =   1111
      _ExtentY        =   714
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
      ButtonDesigner  =   "frmVATaxRealProp.frx":3B35
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdGo 
      Height          =   390
      Left            =   10440
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   2280
      Width           =   855
      _Version        =   131072
      _ExtentX        =   1508
      _ExtentY        =   688
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
      ButtonDesigner  =   "frmVATaxRealProp.frx":3D0E
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   2520
      Y2              =   5880
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1572
      Left            =   120
      Top             =   5880
      Width           =   6132
   End
   Begin VB.Line Line10 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   7560
      X2              =   7560
      Y1              =   3000
      Y2              =   2280
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Go To Pin#"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   7560
      TabIndex        =   77
      Top             =   1920
      Width           =   1620
   End
   Begin VB.Line Line9 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   7560
      X2              =   11400
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bldg Value:"
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
      Height          =   270
      Left            =   360
      TabIndex        =   75
      Top             =   3120
      Width           =   1140
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Opt'l Revenue"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   6240
      TabIndex        =   55
      Top             =   5760
      Width           =   1740
   End
   Begin VB.Line Line8 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   6240
      X2              =   6240
      Y1              =   3720
      Y2              =   5980
   End
   Begin VB.Label Label29 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Property Classification:"
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
      Left            =   312
      TabIndex        =   62
      Top             =   5496
      Width           =   2340
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Opt Search Desc:"
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
      Height          =   270
      Left            =   600
      TabIndex        =   60
      Top             =   7680
      Width           =   1740
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Opt Search Entry:"
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
      Height          =   270
      Left            =   6480
      TabIndex        =   59
      Top             =   7680
      Width           =   1860
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Image Name:"
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
      Height          =   270
      Left            =   2280
      TabIndex        =   58
      Top             =   4560
      Width           =   1260
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   12
      Left            =   480
      Top             =   5880
      Width           =   5772
   End
   Begin VB.Line Line6 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   11400
      X2              =   11400
      Y1              =   1920
      Y2              =   6000
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Lien Desc:"
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
      Height          =   270
      Left            =   6600
      TabIndex        =   57
      Top             =   4920
      Width           =   1020
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Lien Y/N?:"
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
      Height          =   270
      Left            =   6600
      TabIndex        =   56
      Top             =   4440
      Width           =   1020
   End
   Begin VB.Line Line7 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   11400
      X2              =   6240
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   11400
      X2              =   6192
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   11400
      X2              =   11400
      Y1              =   5880
      Y2              =   7440
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Other Data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   6240
      TabIndex        =   54
      Top             =   3720
      Width           =   1305
   End
   Begin VB.Label Label75 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Township:"
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
      Left            =   432
      TabIndex        =   53
      Top             =   5016
      Width           =   1020
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Opt'l Rev3 Y/N?:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   6315
      TabIndex        =   52
      Top             =   7080
      Width           =   2100
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Opt'l Rev2 Y/N?:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   6312
      TabIndex        =   51
      Top             =   6600
      Width           =   2100
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Opt'l Rev1 Y/N?:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   6312
      TabIndex        =   50
      Top             =   6120
      Width           =   2100
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   6220
      X2              =   8280
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   735
      Left            =   120
      Top             =   1920
      Width           =   7455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   555
      X2              =   6195
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Address of Property"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   276
      Left            =   108
      TabIndex        =   49
      Top             =   1920
      Width           =   2580
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Land Rec/GIS Key:"
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
      Height          =   270
      Left            =   3720
      TabIndex        =   48
      Top             =   3120
      Width           =   1860
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mortgage Code:"
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
      Height          =   270
      Left            =   6435
      TabIndex        =   47
      Top             =   5400
      Width           =   1620
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Size:"
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
      Height          =   270
      Left            =   3840
      TabIndex        =   46
      Top             =   4080
      Width           =   540
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Lot/Acre?:"
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
      Height          =   270
      Left            =   4080
      TabIndex        =   45
      Top             =   3600
      Width           =   1020
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Lot:"
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
      Height          =   270
      Left            =   2115
      TabIndex        =   44
      Top             =   4080
      Width           =   540
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Block:"
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
      Height          =   270
      Left            =   315
      TabIndex        =   43
      Top             =   4560
      Width           =   660
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Map:"
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
      Height          =   270
      Left            =   435
      TabIndex        =   42
      Top             =   4080
      Width           =   540
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Property Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   105
      TabIndex        =   41
      Top             =   2640
      Width           =   2190
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   1020
      Index           =   1
      Left            =   1470
      Top             =   180
      Width           =   8655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Real Property Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2040
      TabIndex        =   40
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label71 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
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
      Height          =   270
      Left            =   8145
      TabIndex        =   39
      Top             =   1440
      Width           =   780
   End
   Begin VB.Label Label72 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pin Number:"
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
      Left            =   4440
      TabIndex        =   38
      Top             =   1428
      Width           =   1260
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Discovery Y/N?:"
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
      Height          =   270
      Left            =   8640
      TabIndex        =   37
      Top             =   3960
      Width           =   1500
   End
   Begin VB.Label lblMode 
      BackStyle       =   0  'Transparent
      Caption         =   "Mode:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7320
      TabIndex        =   36
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Record Sequence:"
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
      Left            =   360
      TabIndex        =   35
      Top             =   1440
      Width           =   1860
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Real Value:"
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
      Height          =   270
      Left            =   360
      TabIndex        =   34
      Top             =   3615
      Width           =   1140
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Late List Y/N?:"
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
      Height          =   270
      Left            =   8640
      TabIndex        =   33
      Top             =   4440
      Width           =   1500
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Exemptions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   7560
      TabIndex        =   32
      Top             =   2880
      Width           =   1620
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   855
      Left            =   7560
      Top             =   2880
      Width           =   3855
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Other:"
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
      Height          =   270
      Left            =   8400
      TabIndex        =   31
      Top             =   3360
      Width           =   660
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Descriptions - Notes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   276
      Left            =   120
      TabIndex        =   30
      Top             =   5880
      Width           =   2580
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   12
      Left            =   6240
      Top             =   5880
      Width           =   12
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1080
      Left            =   1470
      Top             =   120
      Width           =   8655
   End
End
Attribute VB_Name = "frmVATaxRealProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim CustName$
  Public WhichRec As Integer
  Dim RealRecs() As Long
  Public NumOfCustRERecs As Integer
  Dim TempRealPIN$
  Dim TempPROPDATE As Integer
  Dim TempGISPOS   As String
  Dim TempMAP      As String
  Dim TempBLOCK    As String
  Dim TempLOTNUMB  As String
  Dim TempLOTACRE  As String
  Dim TempPropSize As Double
  Dim TempPROPDISC As String
  Dim TempLATELIST As String
  Dim TempMORTCODE As String
  Dim TempPROPVALU As Double
  Dim TempBldgVal As Double
  Dim TempEXMPOTHR As Double
  Dim TempPROPNOT1 As String
  Dim TempPROPNOT2 As String
  Dim TempPROPNOT3 As String
  Dim TempPropAddr As String
  Dim TempLienYN As String
  Dim TempLienDesc As String
  Dim TempOptRev1Chrg As Integer
  Dim TempOptRev2Chrg As Integer
  Dim TempOptRev3Chrg As Integer
  Dim TempSearchName$
  Dim TempClass$
  Dim DontExit As Boolean
  Dim RateDesc() As String * 20
  Public ThisImage As String
  Dim GOptSearchDesc As String
  Public ThisBillType$
  
Private Sub cmdAdd_Click()
  On Error GoTo ERRORSTUFF
  
  If Check4Changes(WhichRec) = True Then
    Exit Sub
  End If
  
  If NumOfCustRERecs = 0 Then
    WhichRec = 0
  Else
    WhichRec = NumOfCustRERecs + 1
  End If
  
  Call LoadAdd(WhichRec)
  
  cmdAdd.Enabled = False
  cmdPageDown.Enabled = False
  cmdPageUp.Enabled = False
  cmdDelete.Enabled = False
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxRealProp", "cmdAdd_Click", Erl)
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

Private Sub cmdDelete_Click()
  Dim CustName$
  Dim ThisPin$
  Dim PersVal$
  Dim ThisMap$
  Dim ThisBlock$
  Dim ThisLot$
  Dim ThisBal As Double
  
  On Error GoTo ERRORSTUFF
  If Check4UnpostedBilling(RealRecs(WhichRec), "R") = True Then Exit Sub
  
  ThisPin$ = QPTrim$(fptxtRealPin.Text)
  If Val(ThisPin) > 0 Then
    ThisBal = GetRealBalance(ThisPin)
    If ThisBal <> 0 Then
      Call TaxMsg(900, "This property has an outstanding balance of " + QPTrim$(Using$("$###,###,##0.00", ThisBal)) + ". Please resolve this balance before deleting.")
      Exit Sub
    End If
  End If
  
  frmVATaxMsgWOpts.Label1.Caption = "Are you sure you wish to delete this record? Press F10 to continue the deletion. Otherwise, press ESC to abort the deletion."
  frmVATaxMsgWOpts.Label1.Top = 800
  frmVATaxMsgWOpts.cmdExit.Text = "ESC Abort"
  frmVATaxMsgWOpts.cmdCont.Text = "F10 Delete OK"
  Me.ZOrder 0
  frmVATaxCustAddEdit.Visible = False
  If EditCust = True Then
    frmVATaxCustLookup.Visible = False
  End If
  If AddCust = True Then
    frmVATaxCustMaintMenu.Visible = False
  End If
  frmVATaxMsgWOpts.Show vbModal
  If EditCust = True Then
    frmVATaxCustLookup.Visible = True
  End If
  If AddCust = True Then
    frmVATaxCustMaintMenu.Visible = True
  End If
  frmVATaxCustAddEdit.Visible = True
  Me.Show
  If frmVATaxMsgWOpts.fptxtChoice.Text = "abort" Then
    Unload frmVATaxMsgWOpts
    fptxtRealPin.SetFocus
    Exit Sub
  Else
    Unload frmVATaxMsgWOpts
  End If
  CustName$ = QPTrim$(fptxtThisCust.Text)
  PersVal$ = QPTrim$(fpCurrRealVal.Text)
  ThisMap$ = QPTrim$(fptxtMap.Text)
  ThisBlock$ = QPTrim$(fptxtBlock.Text)
  ThisLot$ = QPTrim$(fptxtLot.Text)
  
  Call DelRealAbstract(RealRecs(), WhichRec, GCustNum)
'  If RealRecs(0) = 0 Then Exit Sub
  Call GetRealRecList(RealRecs(), GCustNum, CustName)
  MainLog ("REAL PROPERTY DELETION: User deleted the following real property for : " + CustName + " - Pin Number: " + ThisPin + " - Map: " + ThisMap + " - Block: " + ThisBlock + " - Lot: " + ThisLot + " - Value: " + PersVal + ".")
  NumOfCustRERecs = RealRecs(0)
  If RealRecs(0) = 0 Then
    WhichRec = 0
    Call LoadMe
  Else
    WhichRec = 1
    Call LoadAgain(WhichRec)
  End If
  frmVATaxMsg.Label1.Caption = "The real property was deleted successfully."
  frmVATaxMsg.Label1.Top = 900
  Me.ZOrder 0
  frmVATaxCustAddEdit.Visible = False
  If EditCust = True Then
    frmVATaxCustLookup.Visible = False
  End If
  If AddCust = True Then
    frmVATaxCustMaintMenu.Visible = False
  End If
  frmVATaxMsg.Show vbModal
  If EditCust = True Then
    frmVATaxCustLookup.Visible = True
  End If
  If AddCust = True Then
    frmVATaxCustMaintMenu.Visible = True
  End If
  frmVATaxCustAddEdit.Visible = True
  Me.Show
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxRealProp", "cmdDelete_Click", Erl)
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

Private Sub cmdDetail1_Click()
  Dim One As Integer
  Dim ThisFile As Integer
  Dim FileName$
  
  If fpcmbOptRev1.Enabled = False Then Exit Sub
  If InStr(fpcmbOptRev1.Text, "NOT IN USE") Then
    Call TaxMsg(800, "'NOT IN USE' prevents any rate detail lookup activity. If there is no access to this optional revenue please check to see if rates have been saved for it.")
    Exit Sub
  End If
  FileName = "C:\CPWork\rdetail1.dat"
  ThisFile = FreeFile
  Open FileName For Output As ThisFile
  One = 1
  Print #ThisFile, One
  Close ThisFile
  
  frmVATaxRateDetail.Show vbModal
  
End Sub

Private Sub cmdDetail2_Click()
  Dim One As Integer
  Dim ThisFile As Integer
  Dim FileName$
  
  If fpcmbOptRev2.Enabled = False Then Exit Sub
  If InStr(fpcmbOptRev2.Text, "NOT IN USE") Then
    Call TaxMsg(800, "'NOT IN USE' prevents any rate detail lookup activity. If there is no access to this optional revenue please check to see if rates have been saved for it.")
    Exit Sub
  End If
  
  FileName = "C:\CPWork\rdetail2.dat"
  ThisFile = FreeFile
  Open FileName For Output As ThisFile
  One = 1
  Print #ThisFile, One
  Close ThisFile
  
  frmVATaxRateDetail.Show vbModal

End Sub

Private Sub cmdDetail3_Click()
  Dim One As Integer
  Dim ThisFile As Integer
  Dim FileName$
  
  If fpcmbOptRev3.Enabled = False Then Exit Sub
  If InStr(fpcmbOptRev3.Text, "NOT IN USE") Then
    Call TaxMsg(800, "'NOT IN USE' prevents any rate detail lookup activity. If there is no access to this optional revenue please check to see if rates have been saved for it.")
    Exit Sub
  End If
  
  FileName = "C:\CPWork\rdetail3.dat"
  ThisFile = FreeFile
  Open FileName For Output As ThisFile
  One = 1
  Print #ThisFile, One
  Close ThisFile
  
  frmVATaxRateDetail.Show vbModal

End Sub

Private Sub cmdExit_Click()
  On Error GoTo ERRORSTUFF

  If cmdAdd.Enabled = False Then
    frmVATaxMsgWOpts.Label1.Caption = "Do you wish to exit without saving any changes? Press F10 to save. Press ESC to exit without saving."
    frmVATaxMsgWOpts.Label1.Top = 900
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Save Changes"
    frmVATaxMsgWOpts.cmdExit.Text = "ESC OK to Exit"
    Me.ZOrder 0
    frmVATaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmVATaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = False
    End If
    frmVATaxMsgWOpts.Show vbModal
    If EditCust = True Then
      frmVATaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = True
    End If
    frmVATaxCustAddEdit.Visible = True
    frmVATaxCustAddEdit.Refresh '12/19/07
    Me.Show
    If frmVATaxMsgWOpts.fptxtChoice.Text = "abort" Then
      Unload frmVATaxMsgWOpts
      Unload Me
      ReDim RealRecs(0 To 0) As Long 'added 10/20/06
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
      Call cmdSave_Click
      If DontExit = True Then
        DontExit = False
        Exit Sub
      Else
        Unload Me
        ReDim RealRecs(0 To 0) As Long 'added 10/20/06
        Exit Sub
      End If
    End If
  End If
  
  If Check4Changes(WhichRec) = True Then
    Exit Sub
  End If
  
  If DontExit = False Then
    Unload Me
    ReDim RealRecs(0 To 0) As Long 'added 10/20/06
  Else
    DontExit = False
  End If
  
  frmVATaxCustAddEdit.Refresh '12/19/07
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxRealProp", "cmdExit_Click", Erl)
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

Private Sub cmdHist_Click()
  If QPTrim$(fptxtRealPin) = "" Then
    Call TaxMsg(800, "Property history can only be generated if the property has a pin number.")
    Exit Sub
  End If
  frmVATaxLoadingRpt.Show
  frmVATaxRealTInfo.Show vbModal
End Sub

Private Sub cmdImage_Click()
  On Error GoTo ErrorExit
  ThisImage = QPTrim$(fptxtImage.Text) + ".bmp"
  ThisImage = CurrCitiPath + "TaxImages\" + ThisImage
  frmVATaxImage.Show vbModal
  DoEvents
  Exit Sub
ErrorExit:

End Sub

Private Sub cmdMortDetail_Click()
  If fpcmbMortCode.Enabled = False Then
    Exit Sub
  End If
  
  If InStr(fpcmbMortCode.Text, "NONE") Then
    Call TaxMsg(900, "'NONE' prevents any mortgage code detail lookup activity.")
    Exit Sub
  End If
  
  frmVATaxMortDetail.Show vbModal
End Sub

Private Sub cmdPageUp_Click()

  On Error GoTo ERRORSTUFF
  
  If Check4Changes(WhichRec) = True Then
    Exit Sub
  End If
  
  If WhichRec = NumOfCustRERecs Then
    frmVATaxMsg.Label1.Caption = "No more pages above this one."
    frmVATaxMsg.Label1.Top = 900
    Me.ZOrder 0
    frmVATaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmVATaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = False
    End If
    frmVATaxMsg.Show vbModal
    If EditCust = True Then
      frmVATaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = True
    End If
    frmVATaxCustAddEdit.Visible = True
    Me.ZOrder 0
    frmVATaxCustAddEdit.ZOrder 1
    Exit Sub
  End If
  
  WhichRec = WhichRec + 1
  Call LoadAgain(WhichRec)
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxRealProp", "cmdPageUp_Click", Erl)
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

Private Sub cmdPageDown_Click()

  On Error GoTo ERRORSTUFF
  
  If Check4Changes(WhichRec) = True Then
    Exit Sub
  End If
  
  If WhichRec = 0 Or WhichRec = 1 Then
    frmVATaxMsg.Label1.Caption = "No more pages below this one."
    frmVATaxMsg.Label1.Top = 900
    Me.ZOrder 0
    frmVATaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmVATaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = False
    End If
    frmVATaxMsg.Show vbModal
    If EditCust = True Then
      frmVATaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = True
    End If
    frmVATaxCustAddEdit.Visible = True
    Me.ZOrder 0
    frmVATaxCustAddEdit.ZOrder 1
    Exit Sub
  End If
  
  WhichRec = WhichRec - 1
  Call LoadAgain(WhichRec)
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxRealProp", "cmdPageDown_Click", Erl)
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

Private Sub cmdSave_Click()
  Dim RealPropRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim CustPin&
  Dim TaxRec As TaxCustType
  Dim THandle As Integer
  Dim NumOfCustRecs As Long
  Dim WhatPers&
  Dim LastPers&
  Dim CustPinRec As PINRecType
  Dim CPHandle As Integer
  Dim NumOfCPRecs As Long
  Dim IntPinRec As InternalPinType
  Dim IHandle As Integer
  Dim NumOfIntPins As Long
  Dim NextIntPin As Long
  Dim NextRec As Long
  Dim CustName$
  
  On Error GoTo ERRORSTUFF
  
  If TempRealPIN$ <> QPTrim$(fptxtRealPin.Text) Then
    If InPayBatchYN(GCustNum, ThisBillType) = True Then
       Exit Sub
    End If
  End If
       
  If QPTrim$(fptxtRealPin.Text) = "" Then
    frmVATaxMsg.Label1.Caption = "The 'Pin Number' field is a requirement. Please enter a 'Pin Number' value."
    frmVATaxMsg.Label1.Top = 900
    Me.ZOrder 0
    frmVATaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmVATaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = False
    End If
    frmVATaxMsg.Show vbModal
    If EditCust = True Then
      frmVATaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = True
    End If
    frmVATaxCustAddEdit.Visible = True
    Me.Show
    fptxtRealPin.SetFocus
    DontExit = True
    Exit Sub
  End If
  
  If WhichRec <= NumOfCustRERecs Then
    If Check4DupPins(QPTrim$(fptxtRealPin.Text), RealRecs(WhichRec)) = True Then
      Close
      Exit Sub
    End If
  Else
    If Check4DupPins(QPTrim$(fptxtRealPin.Text), 0) = True Then
      Close
      Exit Sub
    End If
  End If
  
  If fpCurrRealVal.Value = 0 And fpCurrBldgVal.Value = 0 Then
    frmVATaxMsgWOpts.Label1.Caption = "No real estate or building values have been entered. Press F10 to save anyway. Otherwise, press ESC to abort the save procedure."
    frmVATaxMsgWOpts.Label1.Top = 800
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Save Anyway"
    frmVATaxMsgWOpts.cmdExit.Text = "ESC Abort Save"
    Me.ZOrder 0
    frmVATaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmVATaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = False
    End If
    frmVATaxMsgWOpts.Show vbModal
    If EditCust = True Then
      frmVATaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = True
    End If
    frmVATaxCustAddEdit.Visible = True
    Me.Show
    If frmVATaxMsgWOpts.fptxtChoice.Text = "abort" Then
      Unload frmVATaxMsgWOpts
      fpCurrRealVal.SetFocus
      DontExit = True
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
    End If
  End If
  
  If OldRound(CDbl(fpCurrRealVal.Value) + CDbl(fpCurrBldgVal.Value)) < CDbl(fpCurrOther.Value) Then
    Call TaxMsg(900, "Exemptions cannot be more then the total property value.")
    fpCurrOther.SetFocus
    Exit Sub
  End If

  If GOptSearchDesc = "Not Saved" And QPTrim$(fptxtOptSearch.Text) <> "" Then
    Call TaxMsg(700, "The optional search description has not been saved (located on the town setup screen). Be advised that optional search names are not used until an optional search description is saved.")
    MainLog ("User warned that the optional property search description has not yet been saved and the name they entered will be inactive until a description is saved on the town setup screen.")
  End If
    
  OpenTaxCustFile THandle, NumOfCustRecs
  Get THandle, GCustNum, TaxRec
  CustPin& = TaxRec.PIN
  RealPropRec.RealPin = QPTrim$(fptxtRealPin.Text)
  RealPropRec.PROPDATE = Date2Num(fptxtDate.Text)
  RealPropRec.PROPVALU = fpCurrRealVal
  RealPropRec.BldgVal = fpCurrBldgVal
  RealPropRec.EXMPOTHR = fpCurrOther
  
  RealPropRec.PROPDISC = fpcmbDiscoveryYN.Text
  RealPropRec.LateList = fpcmbLateListYN.Text
  RealPropRec.LienYN = QPTrim$(fpcmbLienYN.Text)
  RealPropRec.LienDesc = QPTrim$(fptxtLienDesc.Text)
  If fpcmbOptRev1.Enabled = True Then
    fpcmbOptRev1.Col = 1
    If QPTrim$(fpcmbOptRev1.ColText) = "" Or QPTrim$(fpcmbOptRev1.ColText) = "0" Then 'added = "0" 10/22/07
      RealPropRec.OptRev1Chrg = 0
    Else
      RealPropRec.OptRev1Chrg = 1 'changed 10/22/07 'CInt(fpcmbOptRev1.ColText)
    End If
  Else
    RealPropRec.OptRev1Chrg = 0
  End If
  If fpcmbOptRev2.Enabled = True Then
    fpcmbOptRev2.Col = 1
    If QPTrim$(fpcmbOptRev2.ColText) = "" Or QPTrim$(fpcmbOptRev2.ColText) = "0" Then 'added = "0" 10/22/07
      RealPropRec.OptRev2Chrg = 0
    Else
      RealPropRec.OptRev2Chrg = 2 'changed 10/22/07'CInt(fpcmbOptRev2.ColText)
    End If
  Else
    RealPropRec.OptRev2Chrg = 0
  End If
  If fpcmbOptRev3.Enabled = True Then
    fpcmbOptRev3.Col = 1
    If QPTrim$(fpcmbOptRev3.ColText) = "" Or QPTrim$(fpcmbOptRev3.ColText) = "0" Then 'added = "0" 10/22/07
      RealPropRec.OptRev3Chrg = 0
    Else
      RealPropRec.OptRev3Chrg = 3 'changed 10/22/07'CInt(fpcmbOptRev3.ColText)
    End If
  Else
    RealPropRec.OptRev3Chrg = 0
  End If
  RealPropRec.TownShip = QPTrim$(fpcmbTownship.Text)
  RealPropRec.ICPDesc = QPTrim$(fpcmbClass.Text)
  RealPropRec.PROPNOT1 = fptxtDesc(0).Text
  RealPropRec.PROPNOT2 = fptxtDesc(1).Text
  RealPropRec.PROPNOT3 = fptxtDesc(2).Text
  RealPropRec.Deleted = 0
  RealPropRec.CustPin = CustPin&
  RealPropRec.MORTCODE = QPTrim$(fpcmbMortCode.Text)
  RealPropRec.PropSize = fpdblSize
  RealPropRec.GISPOS = QPTrim$(fptxtLandGIS.Text)
  RealPropRec.Map = QPTrim$(fptxtMap.Text)
  RealPropRec.BLOCK = QPTrim$(fptxtBlock.Text)
  RealPropRec.LOTNUMB = QPTrim$(fptxtLot.Text)
  RealPropRec.LOTACRE = fpcmbLotAcre.Text
  RealPropRec.PropAddr = QPTrim$(fptxtPropAdd.Text)
  RealPropRec.LastYrPrinted = 0
  RealPropRec.Image = QPTrim$(fptxtImage.Text)
  RealPropRec.OptSearch = QPTrim$(fptxtOptSearch.Text)
  RealPropRec.Blank = ""
  RealPropRec.Fill1 = ""
  OpenRealPropFile RHandle, NumOfRealRecs
    
  WhatPers = NumOfRealRecs + 1
 
  If WhichRec = 0 Then 'first pers prop record for this customer
    RealPropRec.LastYrPrinted = 0
    TaxRec.FirstPropRec = WhatPers&
    Put THandle, GCustNum, TaxRec
    Close THandle
    ReDim Preserve RealRecs(0 To 1) As Long
    RealRecs(0) = 1 '# of props for this customer
    RealRecs(1) = WhatPers 'record # for this prop
    NumOfCustRERecs = 1
    RealPropRec.NextRec = 0
    Put RHandle, WhatPers, RealPropRec
    OpenIntPinFile IHandle, NumOfIntPins
    NextIntPin = NumOfIntPins + 1
    RealPropRec.InternalPin = NextIntPin
    IntPinRec.PIN = RealPropRec.InternalPin
    Put IHandle, NextIntPin, IntPinRec
    Close IHandle
    fptxtRecord.Text = CStr(NumOfCustRERecs) + " of " + CStr(NumOfCustRERecs)
  ElseIf WhichRec > NumOfCustRERecs Then 'adding to existing real prop
    NumOfCustRERecs = NumOfCustRERecs + 1
    ReDim Preserve RealRecs(0 To WhichRec) As Long
    RealRecs(0) = RealRecs(0) + 1
    RealRecs(WhichRec) = WhatPers
    RealPropRec.NextRec = 0
    Put RHandle, WhatPers, RealPropRec
    Get RHandle, RealRecs(NumOfCustRERecs - 1), RealPropRec
    RealPropRec.NextRec = WhatPers
    Put RHandle, RealRecs(NumOfCustRERecs - 1), RealPropRec
    OpenIntPinFile IHandle, NumOfIntPins
    NextIntPin = NumOfIntPins + 1
    RealPropRec.InternalPin = NextIntPin
    IntPinRec.PIN = RealPropRec.InternalPin
    Put IHandle, NextIntPin, IntPinRec
    Close IHandle
    fptxtRecord.Text = CStr(NumOfCustRERecs) + " of " + CStr(NumOfCustRERecs)
  Else 'editing existing data only on screen fields
    Get RHandle, RealRecs(WhichRec), RealPropRec
    RealPropRec.RealPin = QPTrim$(fptxtRealPin.Text)
    RealPropRec.PROPDATE = Date2Num(fptxtDate.Text)
    RealPropRec.PROPVALU = fpCurrRealVal
    RealPropRec.BldgVal = fpCurrBldgVal
    RealPropRec.EXMPOTHR = fpCurrOther
    RealPropRec.PROPDISC = fpcmbDiscoveryYN.Text
    RealPropRec.LateList = fpcmbLateListYN.Text
    RealPropRec.LienYN = QPTrim$(fpcmbLienYN.Text)
    RealPropRec.LienDesc = QPTrim$(fptxtLienDesc.Text)
    If fpcmbOptRev1.Enabled = True Then
      fpcmbOptRev1.Col = 1
      If QPTrim$(fpcmbOptRev1.ColText) = "" Or QPTrim$(fpcmbOptRev1.ColText) = "0" Then 'added = "0" 10/22/07
        RealPropRec.OptRev1Chrg = 0
      Else
        RealPropRec.OptRev1Chrg = 1 ' changed 10/22/07 CInt(fpcmbOptRev1.ColText)
      End If
    Else
      RealPropRec.OptRev1Chrg = 0
    End If
    If fpcmbOptRev2.Enabled = True Then
      fpcmbOptRev2.Col = 1
      If QPTrim$(fpcmbOptRev2.ColText) = "" Or QPTrim$(fpcmbOptRev2.ColText) = "0" Then 'added = "0" 10/22/07
        RealPropRec.OptRev2Chrg = 0
      Else
        RealPropRec.OptRev2Chrg = 2 'changed 10/22/07 CInt(fpcmbOptRev2.ColText)
      End If
    Else
      RealPropRec.OptRev2Chrg = 0
    End If
    If fpcmbOptRev3.Enabled = True Then
      fpcmbOptRev3.Col = 1
      If QPTrim$(fpcmbOptRev3.ColText) = "" Or QPTrim$(fpcmbOptRev3.ColText) = "0" Then 'added = "0" 10/22/07
        RealPropRec.OptRev3Chrg = 0
       Else
        RealPropRec.OptRev3Chrg = 3 'changed 10/22/07 CInt(fpcmbOptRev3.ColText)
      End If
    Else
      RealPropRec.OptRev3Chrg = 0
    End If
    RealPropRec.TownShip = QPTrim$(fpcmbTownship.Text)
    RealPropRec.ICPDesc = QPTrim$(fpcmbClass.Text)
    RealPropRec.PROPNOT1 = fptxtDesc(0).Text
    RealPropRec.PROPNOT2 = fptxtDesc(1).Text
    RealPropRec.PROPNOT3 = fptxtDesc(2).Text
    RealPropRec.MORTCODE = QPTrim$(fpcmbMortCode.Text)
    RealPropRec.PropSize = fpdblSize
    RealPropRec.GISPOS = QPTrim$(fptxtLandGIS.Text)
    RealPropRec.Map = QPTrim$(fptxtMap.Text)
    RealPropRec.BLOCK = QPTrim$(fptxtBlock.Text)
    RealPropRec.LOTNUMB = QPTrim$(fptxtLot.Text)
    RealPropRec.LOTACRE = fpcmbLotAcre.Text
    RealPropRec.Image = QPTrim$(fptxtImage.Text)
    RealPropRec.PropAddr = QPTrim$(fptxtPropAdd.Text)
    RealPropRec.OptSearch = QPTrim$(fptxtOptSearch.Text)
    Put RHandle, RealRecs(WhichRec), RealPropRec
    Call LogSaves
  End If
  
  Close RHandle
  Close THandle
  
  Call MakeRealPINFile
  
  ReDim RealRecs(0 To 0) As Long
  Call GetRealRecList(RealRecs(), GCustNum, CustName)
  
  cmdAdd.Enabled = True
  cmdPageDown.Enabled = True
  cmdPageUp.Enabled = True
  cmdDelete.Enabled = True
  Call AssignTemps
  Call LoadGoToPinsCmb
  
  Me.ZOrder 0
  frmVATaxCustAddEdit.Visible = False
  If EditCust = True Then
    frmVATaxCustLookup.Visible = False
  End If
  If AddCust = True Then
    frmVATaxCustMaintMenu.Visible = False
  End If
  If QPTrim$(fptxtOptSearch.Text) <> "" Then
    Call CreateOptRealIdx
  End If
  Call Savemsg(900, "Your real property data has been saved successfully.")
  DontExit = False

  If EditCust = True Then
    frmVATaxCustLookup.Visible = True
  End If
  If AddCust = True Then
    frmVATaxCustMaintMenu.Visible = True
  End If
  frmVATaxCustAddEdit.Visible = True
  Me.Show
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxRealProp", "cmdSave_Click", Erl)
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
      SendKeys "%C"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
      Call cmdSave_Click
      KeyCode = 0
    Case vbKeyF2:
      Call cmdGo_Click
      KeyCode = 0
    Case vbKeyF3:
      SendKeys "%D"
      Call cmdDelete_Click
      KeyCode = 0
    Case vbKeyF4:
      SendKeys "%H"
      Call cmdHist_Click
      KeyCode = 0
    Case vbKeyF8:
      SendKeys "%A"
      Call cmdAdd_Click
      KeyCode = 0
    Case vbKeyPageUp:
      Call cmdPageUp_Click
      KeyCode = 0
    Case vbKeyPageDown:
      Call cmdPageDown_Click
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
  DontExit = False
  Me.HelpContextID = hlpRealEstate
  Call LoadMe
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxRealProp.")
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

Public Sub LoadMe()
  Dim RealPropRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim x As Integer
  Dim PropSize$
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim TSRec As TownshipType
  Dim TSCnt As Integer
  Dim TSHandle As Integer
  Dim TblRec As OptRevRateTablesType
  Dim TRHandle As Integer
  Dim NumOfTRRecs As Integer
  Dim ThisDesc As String * 20
  Dim MortRec As MortCodeRecType
  Dim MCHandle As Integer
  Dim NumOfMCRecs As Integer
  
  On Error GoTo ERRORSTUFF
  ThisBillType = "R"
  fpcmbMortCode.Clear
  fpcmbMortCode.Enabled = True
  OpenMortCodeFile MCHandle, NumOfMCRecs
  If NumOfMCRecs = 0 Then
    fpcmbMortCode.Text = "UNAVAILABLE"
    fpcmbMortCode.Enabled = False
  Else
    fpcmbMortCode.AddItem "NONE"
    For x = 1 To NumOfMCRecs
      Get MCHandle, x, MortRec
      If MortRec.Deleted = True Then GoTo SkipMort
      fpcmbMortCode.AddItem QPTrim$(MortRec.MORTCODE)
SkipMort:
    Next x
  End If
  Close MCHandle
  fpcmbTownship.Clear
  fpcmbClass.Clear
  
  
1:
  If Exist(TaxTownships) Then
    OpenTownshipFile TSHandle, TSCnt
    For x = 1 To TSCnt
      Get TSHandle, x, TSRec
      fpcmbTownship.AddItem QPTrim$(TSRec.TownShip)
    Next x
    Close TSHandle
  End If
  
  If Exist(TaxSetupName) Then
    OpenTaxSetUpFile TMHandle
    Get TMHandle, 1, TaxMasterRec
    Close TMHandle
    fpcmbClass.Text = "UNCLASSIFIED"
    fpcmbClass.AddItem "UNCLASSIFIED"
    fpcmbClass.AddItem "INDUSTRIAL"
    fpcmbClass.AddItem "COMMERCIAL"
    fpcmbClass.AddItem "PRIVATE"
    For x = 1 To 6
      If QPTrim$(TaxMasterRec.ClassName(x)) <> "" Then
        fpcmbClass.AddItem QPTrim$(TaxMasterRec.ClassName(x))
      End If
    Next x
  Else
    fpcmbClass.Text = "UNCLASSIFIED"
    fpcmbClass.AddItem "UNCLASSIFIED"
    fpcmbClass.AddItem "INDUSTRIAL"
    fpcmbClass.AddItem "COMMERCIAL"
    fpcmbClass.AddItem "PRIVATE"
  End If
  
2:
  If QPTrim$(TaxMasterRec.OptRev1) > "" Then
    Label18.Caption = QPTrim$(TaxMasterRec.OptRev1)
    fpcmbOptRev1.Clear
  Else
    Label18.Caption = "REVENUE #1 UNSAVED"
    fpcmbOptRev1.Enabled = False
  End If
  
  If QPTrim$(TaxMasterRec.OptRev2) > "" Then
    Label20.Caption = QPTrim$(TaxMasterRec.OptRev2)
    fpcmbOptRev2.Clear
  Else
    Label20.Caption = "REVENUE #2 UNSAVED"
    fpcmbOptRev2.Enabled = False
  End If
3:
  If QPTrim$(TaxMasterRec.OptRev3) > "" Then
    Label21.Caption = QPTrim$(TaxMasterRec.OptRev3)
    fpcmbOptRev3.Clear
  Else
    Label21.Caption = "REVENUE #3 UNSAVED"
    fpcmbOptRev3.Enabled = False
  End If
  
  If QPTrim$(TaxMasterRec.OptSrchProp) <> "" Then
    fptxtOptSearchDesc.Text = QPTrim$(TaxMasterRec.OptSrchProp)
'    fptxtOptSearchDesc.ControlType = ControlTypeNormal
    GOptSearchDesc = QPTrim$(TaxMasterRec.OptSrchProp)
  Else
'    fptxtOptSearchDesc.ControlType = ControlTypeReadOnly
    GOptSearchDesc = "Not Saved"
  End If
4:
  If Exist(TaxRateTableFile) Then
    OpenTaxRateTables TRHandle, NumOfTRRecs
    ReDim RateDesc(0 To NumOfTRRecs) As String * 20
    RateDesc(0) = "NOT IN USE"
    fpcmbOptRev1.InsertRow = QPTrim$(RateDesc(0)) + Chr(9) + CStr(0)
    fpcmbOptRev2.InsertRow = QPTrim$(RateDesc(0)) + Chr(9) + CStr(0)
    fpcmbOptRev3.InsertRow = QPTrim$(RateDesc(0)) + Chr(9) + CStr(0)
    For x = 1 To NumOfTRRecs
      Get TRHandle, x, TblRec
      If TblRec.Deleted = True And TblRec.RevType = "R" Then
        fpcmbOptRev1.ToolTipText = "Please note: If there is no access to the optional revenue in the drop down list then the rate settings for this option have not been set."
        fpcmbOptRev2.ToolTipText = "Please note: If there is no access to the optional revenue in the drop down list then the rate settings for this option have not been set."
        fpcmbOptRev3.ToolTipText = "Please note: If there is no access to the optional revenue in the drop down list then the rate settings for this option have not been set."
      End If
      If TblRec.Deleted = True Or TblRec.RevType = "P" Then GoTo Deleted
      RateDesc(x) = QPTrim$(TblRec.Desc)
      If TblRec.OptRevNum = 1 Then
        fpcmbOptRev1.InsertRow = RateDesc(x) + Chr(9) + CStr(x)
      ElseIf TblRec.OptRevNum = 2 Then
        fpcmbOptRev2.InsertRow = RateDesc(x) + Chr(9) + CStr(x)
      ElseIf TblRec.OptRevNum = 3 Then
        fpcmbOptRev3.InsertRow = RateDesc(x) + Chr(9) + CStr(x)
      End If
Deleted:
    Next x
  Else
    ReDim RateDesc(0 To 0) As String * 20
    RateDesc(0) = "NOT IN USE"
    fpcmbOptRev1.InsertRow = RateDesc(0) + Chr(9) + CStr(0)
    fpcmbOptRev2.InsertRow = RateDesc(0) + Chr(9) + CStr(0)
    fpcmbOptRev3.InsertRow = RateDesc(0) + Chr(9) + CStr(0)
  End If
  Close TRHandle
5:
  fpcmbDiscoveryYN.AddItem "Y"
  fpcmbDiscoveryYN.AddItem "N"
  fpcmbLateListYN.AddItem "Y"
  fpcmbLateListYN.AddItem "N"
  fpcmbLotAcre.AddItem "L"
  fpcmbLotAcre.AddItem "A"
  fpcmbLienYN.AddItem "Y"
  fpcmbLienYN.AddItem "N"
  ReDim RealRecs(0 To 0) As Long
  Call GetRealRecList(RealRecs(), GCustNum, CustName)
  fptxtThisCust.Text = CustName
  NumOfCustRERecs = RealRecs(0)
6:
  If NumOfCustRERecs = 0 Then
    WhichRec = 0
    fptxtRealPin.Text = ""
    fptxtRecord.Text = "None Saved"
    lblMode.Caption = "Mode: Adding"
    fptxtDate.Text = Date
    fpcmbDiscoveryYN.Text = "N"
    fpcmbLateListYN.Text = "N"
    fpcmbLotAcre.Text = ""
    fpCurrRealVal = 0
    fpCurrBldgVal = 0
    fpCurrOther = 0
    fptxtDesc(0).Text = ""
    fptxtDesc(1).Text = ""
    fptxtDesc(2).Text = ""
    fptxtMap.Text = ""
    fptxtBlock.Text = ""
    fptxtLot.Text = ""
    fptxtImage.Text = ""
    fptxtLandGIS.Text = ""
    fpdblSize.Value = 0
    fpcmbLienYN.Text = "N"
    fptxtLienDesc.Text = ""
    fptxtPropAdd.Text = ""
    If NumOfMCRecs = 0 Then
      fpcmbMortCode.Text = "UNAVAILABLE"
      fpcmbMortCode.Enabled = False
    Else
      fpcmbMortCode.Text = "NONE"
    End If
    fptxtOptSearch.Text = ""
    fpcmbTownship = ""
    fpcmbClass.Text = "UNCLASSIFIED"
    
7:
    fpcmbOptRev1.SearchText = "NOT IN USE" + Chr(9) + "0"
    fpcmbOptRev1.Action = 0
    If fpcmbOptRev1.SearchIndex <> -1 Then
      fpcmbOptRev1.ListIndex = fpcmbOptRev1.SearchIndex
    Else
      fpcmbOptRev1.ListIndex = 0
    End If
    fpcmbOptRev2.SearchText = "NOT IN USE" + Chr(9) + "0"
    fpcmbOptRev2.Action = 0
    If fpcmbOptRev2.SearchIndex <> -1 Then
      fpcmbOptRev2.ListIndex = fpcmbOptRev2.SearchIndex
    Else
      fpcmbOptRev2.ListIndex = 0
    End If
    fpcmbOptRev3.SearchText = "NOT IN USE" + Chr(9) + "0"
    fpcmbOptRev3.Action = 0
    If fpcmbOptRev3.SearchIndex <> -1 Then
      fpcmbOptRev3.ListIndex = fpcmbOptRev3.SearchIndex
    Else
      fpcmbOptRev3.ListIndex = 0
    End If
  Else
8:
    OpenRealPropFile RHandle, NumOfRealRecs
    Get RHandle, RealRecs(1), RealPropRec
    Close RHandle
    RealPropRec.CustPin = RealPropRec.CustPin
    WhichRec = 1
    fptxtRecord.Text = "1 of " + CStr(NumOfCustRERecs)
    lblMode.Caption = "Mode: Editing"
    fptxtDate.Text = MakeRegDate(RealPropRec.PROPDATE)
    fptxtRealPin.Text = QPTrim$(RealPropRec.RealPin)
9:
    If RealPropRec.PROPDISC <> "Y" Then
      fpcmbDiscoveryYN.Text = "N"
    Else
      fpcmbDiscoveryYN.Text = "Y"
    End If
    If RealPropRec.LateList <> "Y" Then
      fpcmbLateListYN.Text = "N"
    Else
      fpcmbLateListYN.Text = "Y"
    End If
10:
    fpCurrRealVal = RealPropRec.PROPVALU
    fpCurrBldgVal = RealPropRec.BldgVal
    fpCurrOther = RealPropRec.EXMPOTHR
    fptxtDesc(0).Text = RealPropRec.PROPNOT1
    fptxtDesc(1).Text = RealPropRec.PROPNOT2
    fptxtDesc(2).Text = RealPropRec.PROPNOT3
    fpcmbMortCode.Text = QPTrim$(RealPropRec.MORTCODE)
    If QPTrim$(RealPropRec.MORTCODE) = "UNAVAILA" And NumOfMCRecs > 0 Then
      fpcmbMortCode.Text = "NONE"
    ElseIf fpcmbMortCode.Text = "UNAVAILA" Then
      fpcmbMortCode.Text = "UNAVAILABLE"
    End If
    
    fptxtOptSearch.Text = QPTrim$(RealPropRec.OptSearch)

    PropSize$ = CStr(RealPropRec.PropSize)
11:
    If InStr(PropSize, "E") Then
      fpdblSize = 0
    Else
      fpdblSize = RealPropRec.PropSize
    End If
    fptxtLandGIS.Text = QPTrim$(RealPropRec.GISPOS)
    fptxtMap.Text = QPTrim$(RealPropRec.Map)
    fptxtBlock.Text = QPTrim$(RealPropRec.BLOCK)
    fptxtLot.Text = QPTrim$(RealPropRec.LOTNUMB)
    fpcmbLotAcre.Text = RealPropRec.LOTACRE
    If QPTrim$(RealPropRec.Image) <> "" Then
      fptxtImage.Text = QPTrim$(RealPropRec.Image)
    Else
      fptxtImage.Text = "NONE SAVED"
    End If
    fptxtPropAdd.Text = QPTrim$(RealPropRec.PropAddr)
    fpcmbLienYN.Text = RealPropRec.LienYN
    fptxtLienDesc.Text = QPTrim$(RealPropRec.LienDesc)
    fpcmbOptRev1.SearchText = RateDesc(RealPropRec.OptRev1Chrg) + Chr(9) + CStr(RealPropRec.OptRev1Chrg)
    fpcmbOptRev1.Action = 0
12:
    If fpcmbOptRev1.SearchIndex <> -1 Then
      fpcmbOptRev1.ListIndex = fpcmbOptRev1.SearchIndex
    Else
      fpcmbOptRev1.ListIndex = 0
    End If
    fpcmbOptRev2.SearchText = RateDesc(RealPropRec.OptRev2Chrg) + Chr(9) + CStr(RealPropRec.OptRev2Chrg)
    fpcmbOptRev2.Action = 0
    If fpcmbOptRev2.SearchIndex <> -1 Then
      fpcmbOptRev2.ListIndex = fpcmbOptRev2.SearchIndex
    Else
      fpcmbOptRev2.ListIndex = 0
    End If
    fpcmbOptRev3.SearchText = RateDesc(RealPropRec.OptRev3Chrg) + Chr(9) + CStr(RealPropRec.OptRev3Chrg)
    fpcmbOptRev3.Action = 0
    If fpcmbOptRev3.SearchIndex <> -1 Then
      fpcmbOptRev3.ListIndex = fpcmbOptRev3.SearchIndex
    Else
      fpcmbOptRev3.ListIndex = 0
    End If
13:
    fpcmbOptRev1.Col = 0
    ThisDesc = fpcmbOptRev1.ColText
    fpcmbOptRev2.Col = 0
    ThisDesc = fpcmbOptRev2.ColText
    fpcmbOptRev3.Col = 0
    ThisDesc = fpcmbOptRev3.ColText
    fpcmbTownship.Text = QPTrim$(RealPropRec.TownShip)
    fpcmbClass.Text = QPTrim$(RealPropRec.ICPDesc)
14:
    Call AssignTemps
  End If
  Call LoadGoToPinsCmb
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxRealProp", "LoadMe", Erl)
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

Public Sub LoadAgain(WhichRec)
  Dim RealPropRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim PropSize$
  Dim ThisDesc As String * 20
  
  On Error GoTo ERRORSTUFF
  
  OpenRealPropFile RHandle, NumOfRealRecs
  Get RHandle, RealRecs(WhichRec), RealPropRec
  Close RHandle
  fptxtRecord.Text = CStr(WhichRec) + " of " + CStr(NumOfCustRERecs)
  lblMode.Caption = "Mode: Editing"
  fptxtRealPin.Text = QPTrim$(RealPropRec.RealPin)
  If RealPropRec.PROPDISC <> "Y" Then
    fpcmbDiscoveryYN.Text = "N"
  Else
    fpcmbDiscoveryYN.Text = "Y"
  End If
  If RealPropRec.LateList <> "Y" Then
    fpcmbLateListYN.Text = "N"
  Else
    fpcmbLateListYN.Text = "Y"
  End If
  PropSize$ = CStr(RealPropRec.PropSize)
  If InStr(PropSize, "E") Then
    fpdblSize = 0
  Else
    fpdblSize = RealPropRec.PropSize
  End If
  fpCurrRealVal = RealPropRec.PROPVALU
  fptxtDate.Text = MakeRegDate(RealPropRec.PROPDATE)
  fpCurrBldgVal = RealPropRec.BldgVal
  fpCurrOther = RealPropRec.EXMPOTHR
  fptxtDesc(0).Text = RealPropRec.PROPNOT1
  fptxtDesc(1).Text = RealPropRec.PROPNOT2
  fptxtDesc(2).Text = RealPropRec.PROPNOT3
  fptxtRealPin = QPTrim$(RealPropRec.RealPin)
  fpcmbMortCode.Text = QPTrim$(RealPropRec.MORTCODE)
  fptxtLandGIS.Text = QPTrim$(RealPropRec.GISPOS)
  fptxtMap.Text = QPTrim$(RealPropRec.Map)
  fptxtBlock.Text = QPTrim$(RealPropRec.BLOCK)
  fptxtLot.Text = QPTrim$(RealPropRec.LOTNUMB)
  fpcmbLotAcre.Text = RealPropRec.LOTACRE
  fptxtPropAdd.Text = QPTrim$(RealPropRec.PropAddr)
  fpcmbLienYN.Text = RealPropRec.LienYN
  fptxtLienDesc.Text = QPTrim$(RealPropRec.LienDesc)
  fptxtImage.Text = QPTrim$(RealPropRec.Image)
  fptxtOptSearch.Text = QPTrim$(RealPropRec.OptSearch)
  fpcmbOptRev1.SearchText = RateDesc(RealPropRec.OptRev1Chrg) + Chr(9) + CStr(RealPropRec.OptRev1Chrg)
  fpcmbOptRev1.Action = 0
  If fpcmbOptRev1.SearchIndex <> -1 Then
    fpcmbOptRev1.ListIndex = fpcmbOptRev1.SearchIndex
  Else
    fpcmbOptRev1.ListIndex = 0
  End If
  fpcmbOptRev2.SearchText = RateDesc(RealPropRec.OptRev2Chrg) + Chr(9) + CStr(RealPropRec.OptRev2Chrg)
  fpcmbOptRev2.Action = 0
  If fpcmbOptRev2.SearchIndex <> -1 Then
    fpcmbOptRev2.ListIndex = fpcmbOptRev2.SearchIndex
  Else
    fpcmbOptRev2.ListIndex = 0
  End If
  fpcmbOptRev3.SearchText = RateDesc(RealPropRec.OptRev3Chrg) + Chr(9) + CStr(RealPropRec.OptRev3Chrg)
  fpcmbOptRev3.Action = 0
  If fpcmbOptRev3.SearchIndex <> -1 Then
    fpcmbOptRev3.ListIndex = fpcmbOptRev3.SearchIndex
  Else
    fpcmbOptRev3.ListIndex = 0
  End If
  fptxtOptSearch.Text = QPTrim$(RealPropRec.OptSearch)
  fpcmbTownship.Text = QPTrim$(RealPropRec.TownShip)
  fpcmbClass.Text = QPTrim$(RealPropRec.ICPDesc)

  Call AssignTemps
  Call LoadGoToPinsCmb
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxRealProp", "LoadAgain", Erl)
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
Private Sub LoadAdd(WhichRec)
  
  On Error GoTo ERRORSTUFF
  
  If NumOfCustRERecs > 0 Then
    fptxtRecord.Text = "Adding Record # " + CStr(WhichRec)
  Else
    fptxtRecord.Text = "Adding 1st Record"
  End If
  lblMode.Caption = "Mode: Adding"
  fptxtDate.Text = Date
  fptxtRealPin.Text = ""
  fpcmbDiscoveryYN.Text = "N"
  fpcmbLateListYN.Text = "N"
  fpCurrRealVal = 0
  fpCurrBldgVal = 0
  fpCurrOther = 0
  fptxtDesc(0).Text = ""
  fptxtDesc(1).Text = ""
  fptxtDesc(2).Text = ""
  fpcmbMortCode.Text = ""
  fpdblSize = 0
  fptxtLandGIS.Text = ""
  fptxtMap.Text = ""
  fptxtBlock.Text = ""
  fptxtLot.Text = ""
  fpcmbLotAcre.Text = ""
  fptxtPropAdd.Text = ""
  fpcmbTownship.Text = ""
  fptxtImage.Text = ""
  fpcmbLienYN.Text = "N"
  fptxtLienDesc.Text = ""
  fptxtOptSearch.Text = ""
  fpcmbOptRev1.SearchText = "NOT IN USE" + Chr(9) + "0"
  fpcmbOptRev1.Action = 0
  If fpcmbOptRev1.SearchIndex <> -1 Then
    fpcmbOptRev1.ListIndex = fpcmbOptRev1.SearchIndex
  Else
    fpcmbOptRev1.ListIndex = 0
  End If
  fpcmbOptRev2.SearchText = "NOT IN USE" + Chr(9) + "0"
  fpcmbOptRev2.Action = 0
  If fpcmbOptRev2.SearchIndex <> -1 Then
    fpcmbOptRev2.ListIndex = fpcmbOptRev2.SearchIndex
  Else
    fpcmbOptRev2.ListIndex = 0
  End If
  fpcmbOptRev3.SearchText = "NOT IN USE" + Chr(9) + "0"
  fpcmbOptRev3.Action = 0
  If fpcmbOptRev3.SearchIndex <> -1 Then
    fpcmbOptRev3.ListIndex = fpcmbOptRev3.SearchIndex
  Else
    fpcmbOptRev3.ListIndex = 0
  End If
  fptxtRealPin.SetFocus
  fpcmbTownship.Text = ""
  fpcmbClass.Text = "UNCLASSIFIED"
  Call LoadGoToPinsCmb
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxRealProp", "LoadAdd", Erl)
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

Private Sub fpcmbDiscoveryYN_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbDiscoveryYN.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbDiscoveryYN.ListIndex = -1
  End If
  If fpcmbDiscoveryYN.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbLateListYN.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbClass_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbTownship.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbTownship.ListIndex = -1
  End If
  If fpcmbTownship.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpCurrOther.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbLateListYN_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbLateListYN.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbLateListYN.ListIndex = -1
  End If
  If fpcmbLateListYN.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbLienYN.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbLienYN_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbLienYN.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbLienYN.ListIndex = -1
  End If
  If fpcmbLienYN.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtLienDesc.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbLotAcre_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbLotAcre.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbLotAcre.ListIndex = -1
  End If
  If fpcmbLotAcre.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtMap.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbMortCode_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbMortCode.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbMortCode.ListIndex = -1
  End If
  If fpcmbMortCode.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      If fpcmbOptRev1.Enabled = True Then
        fpcmbOptRev1.SetFocus
      ElseIf fpcmbOptRev2.Enabled = True Then
        fpcmbOptRev2.SetFocus
      ElseIf fpcmbOptRev3.Enabled = True Then
        fpcmbOptRev3.SetFocus
      Else
        fptxtRealPin.SetFocus
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

Private Sub fpcmbOptRev1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbOptRev1.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbOptRev1.ListIndex = -1
  End If
  If fpcmbOptRev1.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      If fpcmbOptRev2.Enabled = True Then
        fpcmbOptRev2.SetFocus
      ElseIf fpcmbOptRev3.Enabled = True Then
        fpcmbOptRev3.SetFocus
      Else
        fptxtDesc(0).SetFocus
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

Private Sub fpcmbOptRev2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbOptRev2.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbOptRev2.ListIndex = -1
  End If
  If fpcmbOptRev2.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      If fpcmbOptRev3.Enabled = True Then
        fpcmbOptRev3.SetFocus
      Else
        fptxtDesc(0).SetFocus
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

Private Sub fpcmbOptRev3_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbOptRev3.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbOptRev3.ListIndex = -1
  End If
  If fpcmbOptRev3.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtDesc(0).SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbTownship_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbTownship.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbTownship.ListIndex = -1
  End If
  If fpcmbTownship.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbClass.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fptxtOptSearch_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fptxtRealPin.SetFocus
  ElseIf KeyCode = vbKeyUp Then
    fptxtOptSearchDesc.SetFocus
  End If
End Sub

Private Sub fptxtOptSearchDesc_LostFocus()
  fptxtOptSearchDesc.Text = GOptSearchDesc
End Sub

Private Sub fptxtRealPin_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fptxtDate.SetFocus
  ElseIf KeyCode = vbKeyUp Then
    fptxtDesc(2).SetFocus
  End If
End Sub

Private Sub AssignTemps()
  TempRealPIN = QPTrim$(fptxtRealPin.Text)
  TempPROPDATE = Date2Num(fptxtDate.Text)
  TempGISPOS = QPTrim$(fptxtLandGIS.Text)
  TempMAP = QPTrim$(fptxtMap.Text)
  TempBLOCK = QPTrim$(fptxtBlock.Text)
  TempLOTNUMB = QPTrim$(fptxtLot.Text)
  TempLOTACRE = QPTrim$(fpcmbLotAcre.Text)
  TempPropSize = fpdblSize.Value
  TempPROPDISC = QPTrim$(fpcmbDiscoveryYN.Text)
  TempLATELIST = QPTrim$(fpcmbLateListYN.Text)
  TempMORTCODE = QPTrim$(fpcmbMortCode.Text)
  TempPROPVALU = fpCurrRealVal.Value
  TempBldgVal = fpCurrBldgVal.Value
  TempEXMPOTHR = fpCurrOther.Value
  TempPROPNOT1 = fptxtDesc(0).Text
  TempPROPNOT2 = fptxtDesc(1).Text
  TempPROPNOT3 = fptxtDesc(2).Text
  TempPropAddr = QPTrim$(fptxtPropAdd.Text)
  TempLienYN = QPTrim$(fpcmbLienYN.Text)
  TempLienDesc = QPTrim$(fptxtLienDesc.Text)
  fpcmbOptRev1.Col = 1
  TempOptRev1Chrg = CInt(fpcmbOptRev1.ColText)
  fpcmbOptRev2.Col = 1
  TempOptRev2Chrg = CInt(fpcmbOptRev2.ColText)
  fpcmbOptRev3.Col = 1
  TempOptRev3Chrg = CInt(fpcmbOptRev3.ColText)
  TempSearchName = QPTrim$(fptxtOptSearch.Text)
  TempClass = QPTrim$(fpcmbClass.Text)
End Sub

Private Function Check4Changes(WhichRec As Integer) As Boolean
  Dim ThisControl As Control
  Dim ThisDesc$
  Dim ThatDesc$
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim ThisDbl#
  Dim ThatDbl#
  Dim ThisInt As Integer
  Dim ThatInt As Integer
  Dim choice$
  Dim NoEntry As Boolean
  
  On Error GoTo ERRORSTUFF
  NoEntry = True
  Check4Changes = False
  
  If RealRecs(WhichRec) > 0 Then
    OpenRealPropFile RHandle, NumOfRealRecs
    Get RHandle, RealRecs(WhichRec), RealRec
  Else
    GoSub EntryCheck
    If NoEntry = True Then Exit Function
    frmVATaxMsgWOpts.Label1.Caption = "Do you wish to exit without saving any changes? Press F10 to save. Press ESC to exit without saving."
    frmVATaxMsgWOpts.Label1.Top = 900
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Save Changes"
    frmVATaxMsgWOpts.cmdExit.Text = "ESC OK to Exit"
    Me.ZOrder 0
    frmVATaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmVATaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = False
    End If
    frmVATaxMsgWOpts.Show vbModal
    If EditCust = True Then
      frmVATaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = True
    End If
    frmVATaxCustAddEdit.Visible = True
    Me.Show
    If frmVATaxMsgWOpts.fptxtChoice.Text = "abort" Then
      Unload frmVATaxMsgWOpts
      Exit Function
    Else
      Unload frmVATaxMsgWOpts
      Call cmdSave_Click
      Exit Function
    End If
  End If
    
  Set ThisControl = fptxtRealPin
  ThisDesc = QPTrim$(fptxtRealPin.Text)
  ThatDesc = TempRealPIN
  If ThatDesc <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Pin Number' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmVATaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmVATaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = False
    End If
    frmVATaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmVATaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = True
    End If
    frmVATaxCustAddEdit.Visible = True
    Me.Show
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      If WhichRec <= NumOfCustRERecs Then
        If Check4DupPins(QPTrim$(fptxtRealPin.Text), RealRecs(WhichRec)) = True Then
          choice = "review"
          GoSub HandleChoice
        End If
      Else
        If Check4DupPins(QPTrim$(fptxtRealPin.Text), 0) = True Then
          choice = "review"
          GoSub HandleChoice
        End If
      End If
      RealRec.RealPin = QPTrim$(ThisControl.Text)
      Put RHandle, RealRecs(WhichRec), RealRec
      Call Savemsg(900, "The Pin Number has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtDate
  ThisDesc = fptxtDate.Text
  ThatDesc = MakeRegDate(TempPROPDATE)
  If ThatDesc <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Date' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmVATaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmVATaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = False
    End If
    frmVATaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmVATaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = True
    End If
    frmVATaxCustAddEdit.Visible = True
    Me.Show
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RealRec.PROPDATE = Date2Num(ThisDesc)
      Put RHandle, RealRecs(WhichRec), RealRec
      Call Savemsg(900, "The Date has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fptxtPropAdd
  ThisDesc = QPTrim$(fptxtPropAdd.Text)
  ThatDesc = TempPropAddr
  If ThatDesc <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Property Address' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmVATaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmVATaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = False
    End If
    frmVATaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmVATaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = True
    End If
    frmVATaxCustAddEdit.Visible = True
    Me.Show
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RealRec.PropAddr = QPTrim$(ThisControl.Text)
      Put RHandle, RealRecs(WhichRec), RealRec
      Call Savemsg(900, "Property Address has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtLandGIS
  ThisDesc = QPTrim$(fptxtLandGIS.Text)
  ThatDesc = TempGISPOS
  If ThatDesc <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Land Rec/GIS Key' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmVATaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmVATaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = False
    End If
    frmVATaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmVATaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = True
    End If
    frmVATaxCustAddEdit.Visible = True
    Me.Show
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RealRec.GISPOS = QPTrim$(ThisControl.Text)
      Put RHandle, RealRecs(WhichRec), RealRec
      Call Savemsg(900, "Land Rec/GIS Key has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpCurrRealVal
  ThisDbl = fpCurrRealVal.Value
  ThatDbl = TempPROPVALU
  If ThatDbl <> ThisDbl Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Real Value' field has been changed from " + QPTrim$(Using$("$###,##0.00", ThatDbl)) + " to " + QPTrim$(Using$("$###,##0.00", ThisDbl)) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmVATaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmVATaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = False
    End If
    frmVATaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmVATaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = True
    End If
    frmVATaxCustAddEdit.Visible = True
    Me.Show
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RealRec.PROPVALU = ThisDbl
      Put RHandle, RealRecs(WhichRec), RealRec
      Call Savemsg(900, "Real Value has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fptxtMap
  ThisDesc = QPTrim$(fptxtMap.Text)
  ThatDesc = TempMAP
  If ThatDesc <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Map' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmVATaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmVATaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = False
    End If
    frmVATaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmVATaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = True
    End If
    frmVATaxCustAddEdit.Visible = True
    Me.Show
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RealRec.Map = QPTrim$(ThisControl.Text)
      Put RHandle, RealRecs(WhichRec), RealRec
      Call Savemsg(900, "Map has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtBlock
  ThisDesc = QPTrim$(fptxtBlock.Text)
  ThatDesc = TempBLOCK
  If ThatDesc <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Block' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmVATaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmVATaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = False
    End If
    frmVATaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmVATaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = True
    End If
    frmVATaxCustAddEdit.Visible = True
    Me.Show
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RealRec.BLOCK = QPTrim$(ThisControl.Text)
      Put RHandle, RealRecs(WhichRec), RealRec
      Call Savemsg(900, "Block has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtLot
  ThisDesc = QPTrim$(fptxtLot.Text)
  ThatDesc = TempLOTNUMB
  If ThatDesc <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Lot' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmVATaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmVATaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = False
    End If
    frmVATaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmVATaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = True
    End If
    frmVATaxCustAddEdit.Visible = True
    Me.Show
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RealRec.LOTNUMB = QPTrim$(ThisControl.Text)
      Put RHandle, RealRecs(WhichRec), RealRec
      Call Savemsg(900, "Lot has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpcmbLotAcre
  ThisDesc = QPTrim$(fpcmbLotAcre.Text)
  ThatDesc = TempLOTACRE
  If ThatDesc <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Lot/Acre?' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmVATaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmVATaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = False
    End If
    frmVATaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmVATaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = True
    End If
    frmVATaxCustAddEdit.Visible = True
    Me.Show
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RealRec.LOTACRE = QPTrim$(ThisControl.Text)
      Put RHandle, RealRecs(WhichRec), RealRec
      Call Savemsg(900, "Lot/Acre? has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpCurrBldgVal
  ThisDbl = fpCurrBldgVal.Value
  ThatDbl = TempBldgVal#
  If ThatDbl <> ThisDbl Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Building Value' field has been changed from " + QPTrim$(Using$("$###,##0.00", ThatDbl)) + " to " + QPTrim$(Using$("$###,##0.00", ThisDbl)) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmVATaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmVATaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = False
    End If
    frmVATaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmVATaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = True
    End If
    frmVATaxCustAddEdit.Visible = True
    Me.Show
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RealRec.BldgVal = ThisDbl
      Put RHandle, RealRecs(WhichRec), RealRec
      Call Savemsg(900, "Building Value has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fpCurrOther
  ThisDbl = fpCurrOther.Value
  ThatDbl = TempEXMPOTHR#
  If ThatDbl <> ThisDbl Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Other' field has been changed from " + QPTrim$(Using$("$###,##0.00", ThatDbl)) + " to " + QPTrim$(Using$("$###,##0.00", ThisDbl)) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmVATaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmVATaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = False
    End If
    frmVATaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmVATaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = True
    End If
    frmVATaxCustAddEdit.Visible = True
    Me.Show
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RealRec.EXMPOTHR = ThisDbl
      Put RHandle, RealRecs(WhichRec), RealRec
      Call Savemsg(900, "Other has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fpdblSize
  ThisDbl = fpdblSize.Value
  ThatDbl = TempPropSize
  If ThatDbl <> ThisDbl Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Size' field has been changed from " + QPTrim$(Using$("$###,##0.00", ThatDbl)) + " to " + QPTrim$(Using$("$###,##0.00", ThisDbl)) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmVATaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmVATaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = False
    End If
    frmVATaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmVATaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = True
    End If
    frmVATaxCustAddEdit.Visible = True
    Me.Show
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RealRec.PropSize = ThisDbl
      Put RHandle, RealRecs(WhichRec), RealRec
      Call Savemsg(900, "Size has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fpcmbDiscoveryYN
  ThisDesc = QPTrim$(fpcmbDiscoveryYN.Text)
  ThatDesc = TempPROPDISC
  If ThatDesc <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Discovery Y/N?' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmVATaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmVATaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = False
    End If
    frmVATaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmVATaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = True
    End If
    frmVATaxCustAddEdit.Visible = True
    Me.Show
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RealRec.PROPDISC = QPTrim$(ThisControl.Text)
      Put RHandle, RealRecs(WhichRec), RealRec
      Call Savemsg(900, "Discovery Y/N? has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpcmbLateListYN
  ThisDesc = QPTrim$(fpcmbLateListYN.Text)
  ThatDesc = TempLATELIST
  If ThatDesc <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Late List Y/N?' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmVATaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmVATaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = False
    End If
    frmVATaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmVATaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = True
    End If
    frmVATaxCustAddEdit.Visible = True
    Me.Show
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RealRec.LateList = QPTrim$(ThisControl.Text)
      Put RHandle, RealRecs(WhichRec), RealRec
      Call Savemsg(900, "LateList Y/N? has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpcmbMortCode
  ThisDesc = QPTrim$(fpcmbMortCode.Text)
  If QPTrim$(ThisDesc) = "" Then ThisDesc = "BLANK"
  ThatDesc = TempMORTCODE
  If QPTrim$(ThatDesc) = "" Then ThatDesc = "BLANK"
  If ThatDesc <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Mortgage Code' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmVATaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmVATaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = False
    End If
    frmVATaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmVATaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = True
    End If
    frmVATaxCustAddEdit.Visible = True
    Me.Show
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RealRec.MORTCODE = QPTrim$(ThisControl.Text)
      Put RHandle, RealRecs(WhichRec), RealRec
      Call Savemsg(900, "Mortgage Code has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtDesc(0)
  ThisDesc = fptxtDesc(0).Text
  ThatDesc = TempPROPNOT1
  If ThatDesc <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Description Line #1' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmVATaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmVATaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = False
    End If
    frmVATaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmVATaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = True
    End If
    frmVATaxCustAddEdit.Visible = True
    Me.Show
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RealRec.PROPNOT1 = QPTrim$(ThisControl.Text)
      Put RHandle, RealRecs(WhichRec), RealRec
      Call Savemsg(900, "Description Line #1 has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtDesc(1)
  ThisDesc = fptxtDesc(1).Text
  ThatDesc = TempPROPNOT2
  If ThatDesc <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Description Line #2' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmVATaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmVATaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = False
    End If
    frmVATaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmVATaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = True
    End If
    frmVATaxCustAddEdit.Visible = True
    Me.Show
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RealRec.PROPNOT2 = QPTrim$(ThisControl.Text)
      Put RHandle, RealRecs(WhichRec), RealRec
      Call Savemsg(900, "Description Line #2 has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtDesc(2)
  ThisDesc = fptxtDesc(2).Text
  ThatDesc = TempPROPNOT3
  If ThatDesc <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Description Line #3' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmVATaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmVATaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = False
    End If
    frmVATaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmVATaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = True
    End If
    frmVATaxCustAddEdit.Visible = True
    Me.Show
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RealRec.PROPNOT3 = QPTrim$(ThisControl.Text)
      Put RHandle, RealRecs(WhichRec), RealRec
      Call Savemsg(900, "Description Line #3 has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  If fpcmbOptRev1.Enabled = True Then
    Set ThisControl = fpcmbOptRev1
    fpcmbOptRev1.Col = 1
    ThisInt = fpcmbOptRev1.ColText
    ThatInt = TempOptRev1Chrg
    If ThatInt <> ThisInt Then
      frmVATaxMsgW4Opts.Label1.Caption = "The " + QPTrim$(Label18.Caption) + " field has been changed from " + QPTrim$(RateDesc(ThatInt)) + " to " + QPTrim$(RateDesc(ThisInt)) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmVATaxMsgW4Opts.Label1.Top = 575
      Me.ZOrder 0
      frmVATaxCustAddEdit.Visible = False
      If EditCust = True Then
        frmVATaxCustLookup.Visible = False
      End If
      If AddCust = True Then
        frmVATaxCustMaintMenu.Visible = False
      End If
      frmVATaxMsgW4Opts.Show vbModal
      If EditCust = True Then
        frmVATaxCustLookup.Visible = True
      End If
      If AddCust = True Then
        frmVATaxCustMaintMenu.Visible = True
      End If
      frmVATaxCustAddEdit.Visible = True
      Me.Show
      choice = frmVATaxMsgW4Opts.fptxtChoice.Text
      Unload frmVATaxMsgW4Opts
      If choice = "save" Then
        RealRec.OptRev1Chrg = ThisInt
        Put RHandle, RealRecs(WhichRec), RealRec
        Call Savemsg(900, QPTrim$(Label18.Caption) + " has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    End If
  End If

  If fpcmbOptRev2.Enabled = True Then
    Set ThisControl = fpcmbOptRev2
    fpcmbOptRev2.Col = 1
    ThisInt = fpcmbOptRev2.ColText
    ThatInt = TempOptRev2Chrg
    If ThatInt <> ThisInt Then
      frmVATaxMsgW4Opts.Label1.Caption = "The " + QPTrim$(Label20.Caption) + " field has been changed from " + QPTrim$(RateDesc(ThatInt)) + " to " + QPTrim$(RateDesc(ThisInt)) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmVATaxMsgW4Opts.Label1.Top = 575
      Me.ZOrder 0
      frmVATaxCustAddEdit.Visible = False
      If EditCust = True Then
        frmVATaxCustLookup.Visible = False
      End If
      If AddCust = True Then
        frmVATaxCustMaintMenu.Visible = False
      End If
      frmVATaxMsgW4Opts.Show vbModal
      If EditCust = True Then
        frmVATaxCustLookup.Visible = True
      End If
      If AddCust = True Then
        frmVATaxCustMaintMenu.Visible = True
      End If
      frmVATaxCustAddEdit.Visible = True
      Me.Show
      choice = frmVATaxMsgW4Opts.fptxtChoice.Text
      Unload frmVATaxMsgW4Opts
      If choice = "save" Then
        RealRec.OptRev2Chrg = ThisInt
        Put RHandle, RealRecs(WhichRec), RealRec
        Call Savemsg(900, QPTrim$(Label20.Caption) + " has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    End If
  End If

  If fpcmbOptRev3.Enabled = True Then
    Set ThisControl = fpcmbOptRev3
    fpcmbOptRev3.Col = 1
    ThisInt = fpcmbOptRev3.ColText
    ThatInt = TempOptRev3Chrg
    If ThatInt <> ThisInt Then
      frmVATaxMsgW4Opts.Label1.Caption = "The " + QPTrim$(Label21.Caption) + " field has been changed from " + QPTrim$(RateDesc(ThatInt)) + " to " + QPTrim$(RateDesc(ThisInt)) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
      frmVATaxMsgW4Opts.Label1.Top = 575
      Me.ZOrder 0
      frmVATaxCustAddEdit.Visible = False
      If EditCust = True Then
        frmVATaxCustLookup.Visible = False
      End If
      If AddCust = True Then
        frmVATaxCustMaintMenu.Visible = False
      End If
      frmVATaxMsgW4Opts.Show vbModal
      If EditCust = True Then
        frmVATaxCustLookup.Visible = True
      End If
      If AddCust = True Then
        frmVATaxCustMaintMenu.Visible = True
      End If
      frmVATaxCustAddEdit.Visible = True
      Me.Show
      choice = frmVATaxMsgW4Opts.fptxtChoice.Text
      Unload frmVATaxMsgW4Opts
      If choice = "save" Then
        RealRec.OptRev3Chrg = ThisInt
        Put RHandle, RealRecs(WhichRec), RealRec
        Call Savemsg(900, QPTrim$(Label21.Caption) + " has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    End If
  End If
  
  Set ThisControl = fpcmbLienYN
  ThisDesc = QPTrim$(fpcmbLienYN.Text)
  ThatDesc = QPTrim$(TempLienYN)
  If ThatDesc <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Lien Y/N?' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmVATaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmVATaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = False
    End If
    frmVATaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmVATaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = True
    End If
    frmVATaxCustAddEdit.Visible = True
    Me.Show
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RealRec.LienYN = QPTrim$(ThisControl.Text)
      Put RHandle, RealRecs(WhichRec), RealRec
      Call Savemsg(900, "Lien Y/N? has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtLienDesc
  ThisDesc = QPTrim$(fptxtLienDesc.Text)
  If QPTrim$(ThisDesc) = "" Then
    ThisDesc = "BLANK"
  End If
  ThatDesc = QPTrim$(TempLienDesc)
  If QPTrim$(ThatDesc) = "" Then
    ThatDesc = "BLANK"
  End If
  If ThatDesc <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Lien Description' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmVATaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmVATaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = False
    End If
    frmVATaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmVATaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = True
    End If
    frmVATaxCustAddEdit.Visible = True
    Me.Show
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RealRec.LienDesc = QPTrim$(ThisControl.Text)
      Put RHandle, RealRecs(WhichRec), RealRec
      Call Savemsg(900, "Lien Description has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtOptSearch
  ThisDesc = fptxtOptSearch.Text
  ThatDesc = TempSearchName
  If ThatDesc <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Optional Search Name' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmVATaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmVATaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = False
    End If
    frmVATaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmVATaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = True
    End If
    frmVATaxCustAddEdit.Visible = True
    Me.Show
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RealRec.OptSearch = QPTrim$(ThisControl.Text)
      Put RHandle, RealRecs(WhichRec), RealRec
      Call CreateOptRealIdx
      Call Savemsg(900, "Optional Search Name has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpcmbClass
  ThisDesc = QPTrim$(fpcmbClass.Text)
  If QPTrim$(ThisDesc) = "" Then ThisDesc = "BLANK"
  ThatDesc = TempClass
  If QPTrim$(ThatDesc) = "" Then ThatDesc = "BLANK"
  If ThatDesc <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Property Classification' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
    frmVATaxMsgW4Opts.Label1.Top = 575
    Me.ZOrder 0
    frmVATaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmVATaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = False
    End If
    frmVATaxMsgW4Opts.Show vbModal
    If EditCust = True Then
      frmVATaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = True
    End If
    frmVATaxCustAddEdit.Visible = True
    Me.Show
    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
    Unload frmVATaxMsgW4Opts
    If choice = "save" Then
      RealRec.ICPDesc = QPTrim$(ThisControl.Text)
      Put RHandle, RealRecs(WhichRec), RealRec
      Call Savemsg(900, "Property Classification has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
  
  Close RHandle
  
  Exit Function
  
EntryCheck:
  If QPTrim$(fptxtRealPin.Text) <> "" Then
    NoEntry = False
    Return
  ElseIf QPTrim$(fptxtLandGIS.Text) <> "" Then
    NoEntry = False
    Return
  ElseIf fpCurrRealVal.Value <> 0 Then
    NoEntry = False
    Return
  ElseIf QPTrim$(fptxtMap.Text) <> "" Then
    NoEntry = False
    Return
  ElseIf QPTrim$(fptxtBlock.Text) <> "" Then
    NoEntry = False
    Return
  ElseIf QPTrim$(fptxtLot.Text) <> "" Then
    NoEntry = False
    Return
  ElseIf fpdblSize.Value <> 0 Then
    NoEntry = False
    Return
  ElseIf fpCurrBldgVal.Value <> 0 Then
    NoEntry = False
    Return
  ElseIf fpCurrOther.Value <> 0 Then
    NoEntry = False
    Return
  ElseIf QPTrim$(fpcmbMortCode.Text) <> "NONE" Then
    NoEntry = False
    Return
  ElseIf QPTrim$(fpcmbClass.Text) <> "UNCLASSIFIED" Then
    NoEntry = False
    Return
  ElseIf QPTrim$(fptxtDesc(0).Text) <> "" Then
    NoEntry = False
    Return
  ElseIf QPTrim$(fptxtDesc(1).Text) <> "" Then
    NoEntry = False
    Return
  ElseIf QPTrim$(fptxtDesc(2).Text) <> "" Then
    NoEntry = False
    Return
  ElseIf QPTrim$(fptxtOptSearch.Text) <> "" Then
    NoEntry = False
    Return
  End If

  Return
  
HandleChoice:
    Select Case choice
      Case "abandon"
        Close RHandle
        Unload Me
        ReDim RealRecs(0 To 0) As Long 'added 10/20/06
        Exit Function
      Case "dontsave"
      Case "review"
        ThisControl.SetFocus
        Close RHandle
        Check4Changes = True
        Exit Function
      Case Else
    End Select
      
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxRealProp", "Check4Changes", Erl)
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

Private Sub LogSaves()
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  
  On Error GoTo ERRORSTUFF
  
  OpenRealPropFile RHandle, NumOfRealRecs
  Get RHandle, RealRecs(WhichRec), RealRec
  Close RHandle
  
  If QPTrim$(TempRealPIN$) <> QPTrim$(RealRec.RealPin) Then
    MainLog ("For " + QPTrim$(CustName$) + " Real Estate #" + CStr(WhichRec) + ": the Real Pin # was changed from " + QPTrim$(TempRealPIN$) + " to " + QPTrim$(RealRec.RealPin) + " and saved.")
  End If
  
  If TempPROPDATE% <> RealRec.PROPDATE Then
    MainLog ("For " + QPTrim$(CustName$) + " Real Estate #" + CStr(WhichRec) + ": the date was changed from " + MakeRegDate(TempPROPDATE) + " to " + MakeRegDate(RealRec.PROPDATE) + " and saved.")
  End If
  
  If QPTrim$(TempGISPOS$) <> QPTrim$(RealRec.GISPOS) Then
    MainLog ("For " + QPTrim$(CustName$) + " Real Estate #" + CStr(WhichRec) + ": the LandRec/GIS Key was changed from " + QPTrim$(TempGISPOS$) + " to " + QPTrim$(RealRec.GISPOS) + " and saved.")
  End If
  
  If QPTrim$(TempMAP$) <> QPTrim$(RealRec.Map) Then
    MainLog ("For " + QPTrim$(CustName$) + " Real Estate #" + CStr(WhichRec) + ": the Map was changed from " + QPTrim$(TempMAP$) + " to " + QPTrim$(RealRec.Map) + " and saved.")
  End If
  
  If QPTrim$(TempBLOCK$) <> QPTrim$(RealRec.BLOCK) Then
    MainLog ("For " + QPTrim$(CustName$) + " Real Estate #" + CStr(WhichRec) + ": the Block was changed from " + QPTrim$(TempBLOCK) + " to " + QPTrim$(RealRec.BLOCK) + " and saved.")
  End If
  
  If QPTrim$(TempLOTNUMB$) <> QPTrim$(RealRec.LOTNUMB) Then
    MainLog ("For " + QPTrim$(CustName$) + " Real Estate #" + CStr(WhichRec) + ": the Lot Number was changed from " + QPTrim$(TempLOTNUMB) + " to " + QPTrim$(RealRec.LOTNUMB) + " and saved.")
  End If
  
  If QPTrim$(TempLOTACRE$) <> QPTrim$(RealRec.LOTACRE) Then
    MainLog ("For " + QPTrim$(CustName$) + " Real Estate #" + CStr(WhichRec) + ": the Lot/Acre? was changed from " + QPTrim$(TempLOTACRE) + " to " + QPTrim$(RealRec.LOTACRE) + " and saved.")
  End If
  
  If TempPropSize# <> RealRec.PropSize Then
    MainLog ("For " + QPTrim$(CustName$) + " Real Estate #" + CStr(WhichRec) + ": the Property Size was changed from " + CStr(TempPropSize) + " to " + CStr(RealRec.PropSize) + " and saved.")
  End If
  
  If QPTrim$(TempPROPDISC$) <> QPTrim$(RealRec.PROPDISC) Then
    MainLog ("For " + QPTrim$(CustName$) + " Real Estate #" + CStr(WhichRec) + ": the Discovery was changed from " + QPTrim$(TempPROPDISC$) + " to " + QPTrim$(RealRec.PROPDISC) + " and saved.")
  End If
  
  If QPTrim$(TempLATELIST$) <> QPTrim$(RealRec.LateList) Then
    MainLog ("For " + QPTrim$(CustName$) + " Real Estate #" + CStr(WhichRec) + ": the Late List was changed from " + QPTrim$(TempLATELIST$) + " to " + QPTrim$(RealRec.LateList) + " and saved.")
  End If
  
  If QPTrim$(TempMORTCODE$) <> QPTrim$(RealRec.MORTCODE) Then
    MainLog ("For " + QPTrim$(CustName$) + " Real Estate #" + CStr(WhichRec) + ": the Mortgage Code was changed from " + QPTrim$(TempMORTCODE$) + " to " + QPTrim$(RealRec.MORTCODE) + " and saved.")
  End If
  
  If TempPROPVALU# <> RealRec.PROPVALU Then
    MainLog ("For " + QPTrim$(CustName$) + " Real Estate #" + CStr(WhichRec) + ": the Property Value was changed from " + QPTrim$(Using("$###,###,##0.00", TempPROPVALU)) + " to " + QPTrim$(Using("$###,###,##0.00", RealRec.PROPVALU)) + " and saved.")
  End If
  
  If TempBldgVal# <> RealRec.BldgVal Then
    MainLog ("For " + QPTrim$(CustName$) + " Real Estate #" + CStr(WhichRec) + ": the Building Value was changed from " + QPTrim$(Using("$###,###,##0.00", TempBldgVal)) + " to " + QPTrim$(Using("$###,###,##0.00", RealRec.BldgVal)) + " and saved.")
  End If
  
  If TempEXMPOTHR# <> RealRec.EXMPOTHR Then
    MainLog ("For " + QPTrim$(CustName$) + " Real Estate #" + CStr(WhichRec) + ": the Other Exemption was changed from " + QPTrim$(Using("$###,###,##0.00", TempEXMPOTHR)) + " to " + QPTrim$(Using("$###,###,##0.00", RealRec.EXMPOTHR)) + " and saved.")
  End If
  
  If QPTrim$(TempPROPNOT1$) <> QPTrim$(RealRec.PROPNOT1) Then
    MainLog ("For " + QPTrim$(CustName$) + " Real Estate #" + CStr(WhichRec) + ": the Notes Line #1 was changed from " + QPTrim$(TempPROPNOT1$) + " to " + QPTrim$(RealRec.PROPNOT1) + " and saved.")
  End If
  
  If QPTrim$(TempPROPNOT2$) <> QPTrim$(RealRec.PROPNOT2) Then
    MainLog ("For " + QPTrim$(CustName$) + " Real Estate #" + CStr(WhichRec) + ": the Notes Line #2 was changed from " + QPTrim$(TempPROPNOT2$) + " to " + QPTrim$(RealRec.PROPNOT2) + " and saved.")
  End If
  
  If QPTrim$(TempPROPNOT3$) <> QPTrim$(RealRec.PROPNOT3) Then
    MainLog ("For " + QPTrim$(CustName$) + " Real Estate #" + CStr(WhichRec) + ": the Notes Line #3 was changed from " + QPTrim$(TempPROPNOT3$) + " to " + QPTrim$(RealRec.PROPNOT3) + " and saved.")
  End If
  
  If QPTrim$(TempPropAddr$) <> QPTrim$(RealRec.PropAddr) Then
    MainLog ("For " + QPTrim$(CustName$) + " Real Estate #" + CStr(WhichRec) + ": the Property Address was changed from " + QPTrim$(TempPropAddr$) + " to " + QPTrim$(RealRec.PropAddr) + " and saved.")
  End If
  
  If QPTrim$(TempLienYN$) <> QPTrim$(RealRec.LienYN) Then
    MainLog ("For " + QPTrim$(CustName$) + " Real Estate #" + CStr(WhichRec) + ": the LienYN? was changed from " + QPTrim$(TempLienYN$) + " to " + QPTrim$(RealRec.LienYN) + " and saved.")
  End If
  
  If QPTrim$(TempLienDesc$) <> QPTrim$(RealRec.LienDesc) Then
    MainLog ("For " + QPTrim$(CustName$) + " Real Estate #" + CStr(WhichRec) + ": the Lien Description was changed from " + QPTrim$(TempLienDesc$) + " to " + QPTrim$(RealRec.LienDesc) + " and saved.")
  End If
  
  If QPTrim$(TempClass) <> QPTrim$(RealRec.ICPDesc) Then
    MainLog ("For " + QPTrim$(CustName$) + " Real Estate #" + CStr(WhichRec) + ": the Class Description was changed from " + QPTrim$(TempClass$) + " to " + QPTrim$(RealRec.ICPDesc) + " and saved.")
  End If
  
  If TempOptRev1Chrg% <> RealRec.OptRev1Chrg Then
    MainLog ("For " + QPTrim$(CustName$) + " Real Estate #" + CStr(WhichRec) + ": the Opt'l Rev 1 Y/N? was changed from " + QPTrim$(RateDesc(TempOptRev1Chrg)) + " to " + QPTrim$(RateDesc(RealRec.OptRev1Chrg)) + " and saved.")
  End If
  
  If TempOptRev2Chrg% <> RealRec.OptRev2Chrg Then
    MainLog ("For " + QPTrim$(CustName$) + " Real Estate #" + CStr(WhichRec) + ": the Opt'l Rev 2 Y/N? was changed from " + QPTrim$(RateDesc(TempOptRev2Chrg)) + " to " + QPTrim$(RateDesc(RealRec.OptRev2Chrg)) + " and saved.")
  End If
  
  If TempOptRev3Chrg% <> RealRec.OptRev3Chrg Then
    MainLog ("For " + QPTrim$(CustName$) + " Real Estate #" + CStr(WhichRec) + ": the Opt'l Rev 3 Y/N? was changed from " + QPTrim$(RateDesc(TempOptRev3Chrg)) + " to " + QPTrim$(RateDesc(RealRec.OptRev3Chrg)) + " and saved.")
  End If
  
  If TempSearchName <> QPTrim$(RealRec.OptSearch) Then
    MainLog ("For " + QPTrim$(CustName$) + " Optional Search Name was changed from " + QPTrim$(TempSearchName) + " to " + QPTrim$(RealRec.OptSearch) + " and saved.")
  End If
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxRealProp", "LogSaves", Erl)
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

Private Function Check4DupPins(PinNum$, RecNum As Long) As Boolean
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim x As Long
  
  On Error GoTo ERRORSTUFF
  
  Check4DupPins = False
  OpenRealPropFile RHandle, NumOfRealRecs
  For x = 1 To NumOfRealRecs
    Get RHandle, x, RealRec
    If x <> RecNum Then
      If RealRec.Deleted = -1 Then GoTo Deleted
      If RealRec.CustPin = 0 Then GoTo Deleted
      If QPTrim$(RealRec.RealPin) = PinNum Then
        Check4DupPins = True
        Call TaxMsg(900, "The pin number entered is already in use. Please enter a unique pin number.")
        fptxtRealPin.SetFocus
        Close RHandle
        Exit Function
      End If
    End If
Deleted:
  Next x
  
  Close RHandle
  
  Exit Function

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxRealProp", "Check4DupPins", Erl)
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
Private Sub cmdGo_Click()
  If QPTrim$(fpcmbGoToPins.Text) = "" Then Exit Sub
  If Check4Changes(WhichRec) = True Then
    Exit Sub
  End If
  fpcmbGoToPins.Col = 1
  If fpcmbGoToPins.ColText = "" Then Exit Sub
  WhichRec = fpcmbGoToPins.ColText
  Call LoadAgain(WhichRec)
End Sub

Private Sub fpcmbGoToPins_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbGoToPins.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbGoToPins.ListIndex = -1
  End If
  If fpcmbGoToPins.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpCurrOther.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Public Sub LoadGoToPinsCmb()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NextRec As Long
  Dim NumOfTCRecs As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRRecs As Long
  Dim ThisPin As String
  Dim ThatPin As String
  Dim RecCnt As Integer
  ThisPin = QPTrim$(fptxtRealPin.Text)
  OpenTaxCustFile TCHandle, NumOfTCRecs
  Get TCHandle, GCustNum, TaxCust
  Close TCHandle
  
  NextRec = TaxCust.FirstPropRec
  If NextRec = 0 Then
    fpcmbGoToPins.Action = ActionClear
    Exit Sub
  End If
  
  fpcmbGoToPins.Action = ActionClear
  
  OpenRealPropFile RHandle, NumOfRRecs
  Do While NextRec > 0
    Get RHandle, NextRec, RealRec
    If RealRec.Deleted <> 0 Then GoTo SkipIt
    RecCnt = RecCnt + 1
    ThatPin = QPTrim$(RealRec.RealPin)
    If ThatPin <> "" And ThatPin = ThisPin Then
      fpcmbGoToPins.Text = ThatPin & Chr(9) & CStr(RecCnt)
    End If
    fpcmbGoToPins.AddItem ThatPin & Chr(9) & CStr(RecCnt)
SkipIt:
    NextRec = RealRec.NextRec
  Loop
     
  Close RHandle
  
End Sub

