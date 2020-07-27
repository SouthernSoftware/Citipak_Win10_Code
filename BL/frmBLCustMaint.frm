VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLCustEdit 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Customer Maintenance"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmBLCustMaint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbIOType 
      Height          =   405
      Left            =   9645
      TabIndex        =   22
      Tag             =   "Select the proximity, inside or outside, to the city limits where this business is situated."
      Top             =   4995
      Width           =   1740
      _Version        =   196608
      _ExtentX        =   3069
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
      ColDesigner     =   "frmBLCustMaint.frx":08CA
   End
   Begin LpLib.fpCombo fpcmbPrintNext 
      Height          =   405
      Left            =   10560
      TabIndex        =   25
      Tag             =   $"frmBLCustMaint.frx":0BF9
      Top             =   6285
      Width           =   825
      _Version        =   196608
      _ExtentX        =   1455
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
      ColDesigner     =   "frmBLCustMaint.frx":0D18
   End
   Begin LpLib.fpCombo fpcmbInactiveAcct 
      Height          =   405
      Left            =   9645
      TabIndex        =   26
      Tag             =   $"frmBLCustMaint.frx":1047
      Top             =   6720
      Width           =   1740
      _Version        =   196608
      _ExtentX        =   3069
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
      ColDesigner     =   "frmBLCustMaint.frx":113F
   End
   Begin LpLib.fpCombo fptxtCode 
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   12
      Tag             =   $"frmBLCustMaint.frx":146E
      Top             =   5430
      Width           =   5295
      _Version        =   196608
      _ExtentX        =   9340
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
      Columns         =   3
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   0
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
      ColDesigner     =   "frmBLCustMaint.frx":162F
   End
   Begin LpLib.fpCombo fptxtCode 
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   14
      Tag             =   $"frmBLCustMaint.frx":19E2
      Top             =   5850
      Width           =   5295
      _Version        =   196608
      _ExtentX        =   9340
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
      Columns         =   3
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   0
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
      ColDesigner     =   "frmBLCustMaint.frx":1BA4
   End
   Begin LpLib.fpCombo fptxtCode 
      Height          =   375
      Index           =   2
      Left            =   480
      TabIndex        =   16
      Tag             =   $"frmBLCustMaint.frx":1F57
      Top             =   6285
      Width           =   5295
      _Version        =   196608
      _ExtentX        =   9340
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
      Columns         =   3
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   0
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
      ColDesigner     =   "frmBLCustMaint.frx":2119
   End
   Begin LpLib.fpCombo fptxtCode 
      Height          =   375
      Index           =   3
      Left            =   480
      TabIndex        =   18
      Tag             =   $"frmBLCustMaint.frx":24CC
      Top             =   6720
      Width           =   5295
      _Version        =   196608
      _ExtentX        =   9340
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
      Columns         =   3
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   0
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
      ColDesigner     =   "frmBLCustMaint.frx":268E
   End
   Begin LpLib.fpCombo fptxtCode 
      Height          =   375
      Index           =   4
      Left            =   480
      TabIndex        =   20
      Tag             =   $"frmBLCustMaint.frx":2A41
      Top             =   7155
      Width           =   5295
      _Version        =   196608
      _ExtentX        =   9340
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
      Columns         =   3
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   0
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
      ColDesigner     =   "frmBLCustMaint.frx":2C03
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDelete 
      Height          =   540
      Left            =   3690
      TabIndex        =   31
      TabStop         =   0   'False
      Tag             =   $"frmBLCustMaint.frx":2FB6
      Top             =   7770
      Width           =   1845
      _Version        =   131072
      _ExtentX        =   3254
      _ExtentY        =   952
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
      ButtonDesigner  =   "frmBLCustMaint.frx":3087
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCustList 
      Height          =   390
      Left            =   6450
      TabIndex        =   28
      TabStop         =   0   'False
      Tag             =   $"frmBLCustMaint.frx":3264
      Top             =   1395
      Width           =   1815
      _Version        =   131072
      _ExtentX        =   3201
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
      ButtonDesigner  =   "frmBLCustMaint.frx":335A
   End
   Begin EditLib.fpMask fpMaskPhone 
      Height          =   396
      Left            =   9648
      TabIndex        =   23
      Tag             =   "Enter the phone number for this business."
      Top             =   5424
      Width           =   1740
      _Version        =   196608
      _ExtentX        =   3069
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
      AllowOverflow   =   0   'False
      BestFit         =   0   'False
      ClipMode        =   0
      DataFormatEx    =   0
      Mask            =   "(###)###-####"
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      RequireFill     =   0   'False
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      AutoTab         =   0   'False
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtCustNum 
      Height          =   390
      Left            =   10320
      TabIndex        =   30
      TabStop         =   0   'False
      Tag             =   "This number is assigned automatically by the program. It is not possible to edit this field."
      Top             =   1440
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
      _ExtentY        =   688
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
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
      ControlType     =   1
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
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtSearchName 
      Height          =   340
      Left            =   7560
      TabIndex        =   9
      Tag             =   "Enter a short name for this business in this field. In some searches this name is used to speed up the search process."
      Top             =   2675
      Width           =   3105
      _Version        =   196608
      _ExtentX        =   5477
      _ExtentY        =   600
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
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
      MaxLength       =   10
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtBusName 
      Height          =   390
      Left            =   2040
      TabIndex        =   1
      Tag             =   "Enter the name of this business as popularly known to it's customers."
      Top             =   1875
      Width           =   4380
      _Version        =   196608
      _ExtentX        =   7726
      _ExtentY        =   698
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
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
      MaxLength       =   35
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtAddress1 
      Height          =   390
      Left            =   2040
      TabIndex        =   2
      Tag             =   "Enter the address this business considers it's primary address. Usually this is a street address."
      Top             =   2355
      Width           =   4380
      _Version        =   196608
      _ExtentX        =   7726
      _ExtentY        =   698
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
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
      MaxLength       =   35
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtAddress2 
      Height          =   390
      Left            =   2040
      TabIndex        =   3
      Tag             =   "Enter the secondary business address here. Usually a secondary address would be a post office box or a suite number, etc."
      Top             =   2835
      Width           =   4380
      _Version        =   196608
      _ExtentX        =   7726
      _ExtentY        =   698
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
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
      MaxLength       =   35
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtCity 
      Height          =   390
      Left            =   2040
      TabIndex        =   4
      Tag             =   "Enter the city where this business receives it's mail."
      Top             =   3315
      Width           =   4380
      _Version        =   196608
      _ExtentX        =   7726
      _ExtentY        =   698
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
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
      MaxLength       =   20
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtBillingName 
      Height          =   390
      Left            =   2040
      TabIndex        =   0
      Tag             =   "Enter the name this business will use in correspondance regarding any business license affairs."
      Top             =   1395
      Width           =   4380
      _Version        =   196608
      _ExtentX        =   7726
      _ExtentY        =   698
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
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
      MaxLength       =   35
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtContact 
      Height          =   345
      Left            =   6795
      TabIndex        =   10
      Tag             =   $"frmBLCustMaint.frx":353E
      Top             =   3400
      Width           =   4665
      _Version        =   196608
      _ExtentX        =   8229
      _ExtentY        =   600
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
      BackColor       =   16777215
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
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
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtState 
      Height          =   390
      Left            =   2040
      TabIndex        =   5
      Tag             =   $"frmBLCustMaint.frx":35CA
      Top             =   3795
      Width           =   540
      _Version        =   196608
      _ExtentX        =   952
      _ExtentY        =   698
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
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
      CharValidationText=   "A B C D E F G H I J K L M N O P Q R S T U V W X Y Z"
      MaxLength       =   2
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   -1  'True
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpMask fptxtZip 
      Height          =   390
      Left            =   4635
      TabIndex        =   6
      Tag             =   "Enter the five digit or nine digit postal code for this business."
      Top             =   3795
      Width           =   1785
      _Version        =   196608
      _ExtentX        =   3154
      _ExtentY        =   698
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
      AllowOverflow   =   0   'False
      BestFit         =   0   'False
      ClipMode        =   0
      DataFormatEx    =   0
      Mask            =   "#####-####"
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      RequireFill     =   0   'False
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      AutoTab         =   0   'False
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpDateTime fptxtValidThru 
      Height          =   348
      Left            =   9648
      TabIndex        =   24
      Tag             =   $"frmBLCustMaint.frx":3662
      Top             =   5880
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
      Text            =   "11/20/2002"
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
   Begin EditLib.fpText fptxtProrate 
      Height          =   348
      Left            =   9648
      TabIndex        =   27
      Tag             =   $"frmBLCustMaint.frx":3716
      Top             =   7152
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
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
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 % ."
      MaxLength       =   7
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   540
      Left            =   9552
      TabIndex        =   32
      TabStop         =   0   'False
      Tag             =   $"frmBLCustMaint.frx":38CD
      Top             =   7776
      Width           =   1836
      _Version        =   131072
      _ExtentX        =   3238
      _ExtentY        =   952
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
      ButtonDesigner  =   "frmBLCustMaint.frx":39D7
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   540
      Left            =   7620
      TabIndex        =   33
      TabStop         =   0   'False
      Tag             =   $"frmBLCustMaint.frx":3BB5
      Top             =   7776
      Width           =   1836
      _Version        =   131072
      _ExtentX        =   3238
      _ExtentY        =   952
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
      ButtonDesigner  =   "frmBLCustMaint.frx":3C59
   End
   Begin EditLib.fpText fptxtRev 
      Height          =   348
      Index           =   0
      Left            =   5808
      TabIndex        =   13
      Tag             =   $"frmBLCustMaint.frx":3E35
      Top             =   5424
      Width           =   1644
      _Version        =   196608
      _ExtentX        =   2900
      _ExtentY        =   614
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
      ThreeDInsideHighlightColor=   -2147483637
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   -2147483628
      ThreeDOutsideShadowColor=   -2147483632
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
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
      ThreeDTextHighlightColor=   -2147483637
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
      CaretOverWrite  =   0
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
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 . , $"
      MaxLength       =   35
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtRev 
      Height          =   348
      Index           =   1
      Left            =   5808
      TabIndex        =   15
      Tag             =   $"frmBLCustMaint.frx":4096
      Top             =   5856
      Width           =   1644
      _Version        =   196608
      _ExtentX        =   2900
      _ExtentY        =   614
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
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
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 . , $"
      MaxLength       =   35
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtRev 
      Height          =   348
      Index           =   2
      Left            =   5808
      TabIndex        =   17
      Tag             =   $"frmBLCustMaint.frx":42F7
      Top             =   6288
      Width           =   1644
      _Version        =   196608
      _ExtentX        =   2900
      _ExtentY        =   614
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
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
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 . , $"
      MaxLength       =   35
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtRev 
      Height          =   348
      Index           =   3
      Left            =   5808
      TabIndex        =   19
      Tag             =   $"frmBLCustMaint.frx":4558
      Top             =   6720
      Width           =   1644
      _Version        =   196608
      _ExtentX        =   2900
      _ExtentY        =   614
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
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
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 . , $"
      MaxLength       =   35
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtRev 
      Height          =   348
      Index           =   4
      Left            =   5808
      TabIndex        =   21
      Tag             =   $"frmBLCustMaint.frx":47B9
      Top             =   7152
      Width           =   1644
      _Version        =   196608
      _ExtentX        =   2900
      _ExtentY        =   614
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
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
      CharValidationText=   "1 2 3 4 5 6 7 8 9 0 . , $"
      MaxLength       =   35
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdList 
      Height          =   396
      Left            =   4536
      TabIndex        =   29
      TabStop         =   0   'False
      Tag             =   $"frmBLCustMaint.frx":4A1A
      Top             =   4272
      Width           =   1884
      _Version        =   131072
      _ExtentX        =   3323
      _ExtentY        =   698
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
      ButtonDesigner  =   "frmBLCustMaint.frx":4AED
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   540
      Left            =   480
      TabIndex        =   62
      TabStop         =   0   'False
      Tag             =   "Place the cursor over any field and a pop-up balloon will appear containing help information about that field."
      ToolTipText     =   "Press here to exit this screen."
      Top             =   7776
      Width           =   2220
      _Version        =   131072
      _ExtentX        =   3916
      _ExtentY        =   952
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
      ButtonDesigner  =   "frmBLCustMaint.frx":4CD0
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   300
      Left            =   240
      TabIndex        =   66
      Top             =   240
      Width           =   735
      _Version        =   131072
      _ExtentX        =   1291
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
      ShapeRoundWidth =   195
      ShapeRoundHeight=   195
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
      MaxWidth        =   6000
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
   Begin EditLib.fpText fptxtLicNum 
      Height          =   390
      Left            =   2040
      TabIndex        =   7
      Tag             =   $"frmBLCustMaint.frx":4EB3
      Top             =   4275
      Width           =   2460
      _Version        =   196608
      _ExtentX        =   4339
      _ExtentY        =   698
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
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
      MaxLength       =   35
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdTransHist 
      Height          =   540
      Left            =   5664
      TabIndex        =   68
      TabStop         =   0   'False
      Tag             =   $"frmBLCustMaint.frx":503C
      Top             =   7776
      Width           =   1836
      _Version        =   131072
      _ExtentX        =   3238
      _ExtentY        =   952
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
      ButtonDesigner  =   "frmBLCustMaint.frx":50DD
   End
   Begin EditLib.fpText fptxtServAdd 
      Height          =   340
      Left            =   6795
      TabIndex        =   11
      Tag             =   "Enter the physical address of this business if necessary."
      Top             =   4080
      Width           =   4665
      _Version        =   196608
      _ExtentX        =   8229
      _ExtentY        =   600
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
      BackColor       =   16777215
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
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
      MaxLength       =   35
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtSSNFID 
      Height          =   345
      Left            =   8040
      TabIndex        =   8
      Tag             =   "This number is assigned automatically by the program. It is not possible to edit this field."
      Top             =   1965
      Width           =   1980
      _Version        =   196608
      _ExtentX        =   3492
      _ExtentY        =   609
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483637
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
      MaxLength       =   15
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Label Label33 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SSN/FID:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   8520
      TabIndex        =   71
      Top             =   1608
      Width           =   1020
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cust Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   9720
      TabIndex        =   70
      Top             =   1080
      Width           =   1740
   End
   Begin VB.Label Label32 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Service Address:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   7800
      TabIndex        =   69
      Top             =   3792
      Width           =   2136
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   7584
      X2              =   7584
      Y1              =   4848
      Y2              =   7536
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
      Left            =   576
      TabIndex        =   67
      Top             =   8352
      Width           =   2076
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "/Flat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   300
      Left            =   6816
      TabIndex        =   65
      Top             =   5040
      Width           =   540
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   4800
      TabIndex        =   64
      Top             =   5040
      Width           =   828
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Required Fields = *"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   240
      TabIndex        =   63
      Top             =   1080
      Width           =   1785
   End
   Begin VB.Label lblSave 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Indexing and Saving..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   61
      Top             =   4560
      Width           =   3750
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   4416
      X2              =   11424
      Y1              =   4848
      Y2              =   4848
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   480
      X2              =   1536
      Y1              =   4848
      Y2              =   4848
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "/Multi"
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
      Height          =   300
      Left            =   6288
      TabIndex        =   60
      Top             =   5040
      Width           =   588
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   6645
      X2              =   6645
      Y1              =   4410
      Y2              =   4842
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "5)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   192
      TabIndex        =   59
      Top             =   7200
      Width           =   300
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "4)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   192
      TabIndex        =   58
      Top             =   6768
      Width           =   300
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   192
      TabIndex        =   57
      Top             =   6336
      Width           =   300
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   192
      TabIndex        =   56
      Top             =   5904
      Width           =   300
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   192
      TabIndex        =   55
      Top             =   5472
      Width           =   300
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*Prorate:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   8592
      TabIndex        =   54
      Top             =   7248
      Width           =   972
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*Inactive/Active:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   7680
      TabIndex        =   53
      Top             =   6816
      Width           =   1884
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*Set Renewal Flag (Y/N)?:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   7680
      TabIndex        =   52
      Top             =   6384
      Width           =   2844
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*Valid Thru:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   7872
      TabIndex        =   51
      Top             =   5952
      Width           =   1692
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*License #:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   690
      TabIndex        =   50
      Top             =   4365
      Width           =   1260
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Work Phone #:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   7872
      TabIndex        =   49
      Top             =   5520
      Width           =   1692
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "In/Out City Lmts:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   7632
      TabIndex        =   48
      Top             =   5088
      Width           =   1932
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "*Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   432
      TabIndex        =   47
      Top             =   5040
      Width           =   828
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   2256
      TabIndex        =   46
      Top             =   5040
      Width           =   1308
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Step"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   5760
      TabIndex        =   45
      Top             =   5040
      Width           =   684
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "BILLING CATEGORIES"
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
      Left            =   1824
      TabIndex        =   44
      Top             =   4704
      Width           =   2364
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*Zip Code:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   3285
      TabIndex        =   43
      Top             =   3885
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*State:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   1125
      TabIndex        =   42
      Top             =   3885
      Width           =   825
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contact/Owner Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   7920
      TabIndex        =   41
      Top             =   3085
      Width           =   2370
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*Business Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   75
      TabIndex        =   40
      Top             =   1965
      Width           =   1890
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*City:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   1170
      TabIndex        =   39
      Top             =   3405
      Width           =   780
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address Line 2:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   210
      TabIndex        =   38
      Top             =   2925
      Width           =   1740
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*Address Line 1:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   165
      TabIndex        =   37
      Top             =   2445
      Width           =   1785
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*Billing Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   360
      TabIndex        =   36
      Top             =   1485
      Width           =   1590
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*Search Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   300
      Left            =   8160
      TabIndex        =   35
      Top             =   2360
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   750
      Index           =   1
      Left            =   1500
      Top             =   240
      Width           =   8655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Maintenance"
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
      Left            =   3945
      TabIndex        =   34
      Top             =   390
      Width           =   3750
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   1500
      Top             =   195
      Width           =   8655
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   6645
      X2              =   6645
      Y1              =   1725
      Y2              =   4413
   End
End
Attribute VB_Name = "frmBLCustEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class
  Dim AddFlag As Boolean
  Dim FirstTime As Boolean
  Dim TempBillName$
  Dim TempSearchName$
  Dim TempCustName$
  Dim TempCustNumb$
  Dim TempAddress1$
  Dim TempAddress2$
  Dim TempCity$
  Dim TempState$
  Dim TempZip$
  Dim TempContact$
  Dim TempServAdd$
  Dim TempBillCat1$
  Dim TempDESC1$
  Dim TempType1$
  Dim TempRev1 As Double
  Dim TempBillCat2$
  Dim TempDESC2$
  Dim TempType2$
  Dim TempRev2 As Double
  Dim TempBillCat3$
  Dim TempDESC3$
  Dim TempType3$
  Dim TempRev3 As Double
  Dim TempBillCat4$
  Dim TempType4$
  Dim TempDESC4$
  Dim TempRev4 As Double
  Dim TempBillCat5$
  Dim TempType5$
  Dim TempDESC5$
  Dim TempRev5 As Double
  Dim TempLocation$
  Dim TempWPHONE$
  Dim TempVALID As Integer
  Dim TempLICENSE$
  Dim TempIssueLicense$
  Dim TempInactive$
  Dim TempProrate As Double
  Dim PermNum As Boolean
  Dim TempSSNFID$
  Dim CreditAmt(0 To 4) As Double
  Dim TempGCustNum As Integer
  Dim NFlag As Boolean
  
Private Sub cmdCatList_Click()
  frmBLCategoryList.Show vbModal
End Sub

Private Sub cmdCustList_Click()
  frmBLCustomerList.Show vbModal
End Sub

Private Sub cmdDelete_Click()
  Dim CHandle As Integer
  Dim CustRec As ARCustRecType
  
  On Error GoTo ERRORSTUFF
  If GCustNum = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No active customer selected. Cannot delete."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  If EmpInLicProcess(CStr(GCustNum)) = True Then
    frmBLMessageBoxJr.Label1.Caption = "This customer is involved in an unposted license fee file and cannot be deleted."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  If EmpInPenProcess(CStr(GCustNum)) = True Then
    frmBLMessageBoxJr.Label1.Caption = "This customer is involved in an unposted penalty fee file and cannot be deleted."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  If EmpInPayProcess(CStr(GCustNum)) = True Then
    frmBLMessageBoxJr.Label1.Caption = "This customer is involved in an unposted payment file and cannot be deleted."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  'user given choice of continuing or exiting without deleting
  frmBLMessageBoxJrWOpts.Label1.Caption = "Deleting a business permanently removes that business and all it's data from memory. Do you wish to continue?"
  frmBLMessageBoxJrWOpts.Label1.Top = 800
  frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
  frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Abort"
  frmBLMessageBoxJrWOpts.Show vbModal
  
  If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
    Close
    Unload frmBLMessageBoxJrWOpts
    Exit Sub
  Else
    Unload frmBLMessageBoxJrWOpts
  End If
  
  OpenCustFile CHandle
  
  Get CHandle, GCustNum, CustRec
    
    'if this customer has a balance pending then this trap alerts
    'the user
    If CustRec.AcctBal <> 0 Then
      frmBLMessageBoxJr.Label1.Top = 600
      frmBLMessageBoxJr.Label1.Caption = "This customer has a balance of " + QPTrim$(Using("$###,##0.00", CustRec.AcctBal)) + " remaining. Please resolve this outstanding balance before deleting. If a payment has been made but not posted then that amount is assumed as still outstanding."
      frmBLMessageBoxJr.Show vbModal
      Close
      Exit Sub
    End If
    'This is where the deletion actually happens...the rest of the program
    'looks for these two pieces of data and filters out customers who have
    'this saved...so technically the data still exists if we need to get at it
    CustRec.Deleted = "Y"
    CustRec.SortName = "DELETED"
  Put CHandle, GCustNum, CustRec
  Close CHandle
  
  'indexes are processed to eliminate this customer from
  'those lists
  Call CreateCustNameIdx
  Call CreateCustNumIdx
  Call CreateLicNumIdx
  Call CreateCustSearchNameIdx
  
  frmBLMessageBoxJr.Label1.Caption = QPTrim$(CustRec.CustName) + " has been deleted."
  frmBLMessageBoxJr.Label1.Top = 800
  frmBLMessageBoxJr.Show vbModal
  MainLog (QPTrim$(CustRec.CustName) + " was deleted using the delete button on the customer edit screen. Warning was issued stating that continuing would remove this customer's data permanently.")
  Unload frmBLCustomerLookup
  frmBLCustMaintMenu.Show
  DoEvents
  Unload Me
  
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustEdit", "cmdDelete_Click", Erl)
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

Public Sub cmdExit_Click()
  Dim ChangeFlag As Boolean
  Dim CustFile As Integer
  Dim DoWhatFlag As SaveChangeOptions1
  Dim CustRec As ARCustRecType
  Dim Control As Controls, x As Integer
  Dim ThisInactive$
  Dim ThisVal$
  
  On Error GoTo ERRORSTUFF
  
  If GCustNum = 0 Then
    frmBLMessageBoxJrWOpts.Label1.Caption = "Are you sure you want to exit without saving any changes?"
    frmBLMessageBoxJrWOpts.Label1.Top = 900
    frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 OK To Exit"
    frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Don't Exit"
    frmBLMessageBoxJrWOpts.Show vbModal
    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
      Unload frmBLMessageBoxJrWOpts
      Close
      If fptxtBillingName.Enabled = True Then
        fptxtBillingName.SetFocus
      End If
      Exit Sub
    Else
      Unload frmBLMessageBoxJrWOpts
      MainLog ("User has exited the Add Customer screen without saving any changes after being warned they were exiting without saving changes.")
      GoTo CustNumIsZero 'user is exiting
      'without saving new record entries...also skips the change
      'check feature and if custlist is open then the number
      'double clicked will be brought up to this screen
    End If
  End If
10:
  ChangeFlag = False
  OpenCustFile CustFile
  Get CustFile, GCustNum, CustRec
  Close CustFile

  'the next series of code is designed to keep the user from
  'exiting without saving any changes made...user may have erroneously thought
  'he had already saved the data
  
  If QPTrim$(CustRec.CustName) <> QPTrim$(fptxtBusName.Text) Then
    ChangeFlag = True
    If fptxtBusName.Enabled = True Then
      fptxtBusName.SetFocus
    End If
    GoTo ChangeFound
  End If

  If QPTrim$(CustRec.CustNumb) <> QPTrim$(fptxtCustNum.Text) Then
    ChangeFlag = True
    If fptxtCustNum.Enabled = True Then
      fptxtCustNum.SetFocus
    End If
    GoTo ChangeFound
  End If

  If QPTrim$(CustRec.ADDRESS1) <> QPTrim$(fptxtAddress1.Text) Then
    ChangeFlag = True
    fptxtAddress1.SetFocus
    GoTo ChangeFound
  End If

  If QPTrim$(CustRec.ADDRESS2) <> QPTrim$(fptxtAddress2.Text) Then
    ChangeFlag = True
    fptxtAddress2.SetFocus
    GoTo ChangeFound
  End If

  If QPTrim$(CustRec.City) <> QPTrim$(fptxtCity.Text) Then
    ChangeFlag = True
    fptxtCity.SetFocus
    GoTo ChangeFound
  End If
20:
  If QPTrim$(CustRec.State) <> QPTrim$(fptxtState.Text) Then
    ChangeFlag = True
    fptxtState.SetFocus
    GoTo ChangeFound
  End If

  If QPTrim$(CustRec.ZipCode) = "-" Then 'means no zip is saved
    If QPTrim$(fptxtZip.Text) = "" Then
      GoTo OverrideZipMask
    End If
  End If
      
  If QPTrim$(CustRec.ZipCode) <> QPTrim$(fptxtZip.Text) Then
    ChangeFlag = True
    fptxtZip.SetFocus
    GoTo ChangeFound
  End If

OverrideZipMask:

  If QPTrim$(CustRec.BillName) <> QPTrim$(fptxtBillingName.Text) Then
    ChangeFlag = True
    If fptxtBillingName.Enabled = True Then
      fptxtBillingName.SetFocus
    End If
    GoTo ChangeFound
  End If
30:
  If QPTrim$(CustRec.SortName) <> QPTrim$(fptxtSearchName.Text) Then
    ChangeFlag = True
    If fptxtSearchName.Enabled = True Then
      fptxtSearchName.SetFocus
    End If
    GoTo ChangeFound
  End If

  If QPTrim$(CustRec.Contact) <> QPTrim$(fptxtContact.Text) Then
    ChangeFlag = True
    fptxtContact.SetFocus
    GoTo ChangeFound
  End If

  If QPTrim$(CustRec.ServAdd) <> QPTrim$(fptxtServAdd.Text) Then
    ChangeFlag = True
    fptxtServAdd.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(CustRec.SSNFID) <> QPTrim$(fptxtSSNFID.Text) Then
    ChangeFlag = True
    fptxtSSNFID.SetFocus
    GoTo ChangeFound
  End If

  If QPTrim$(CustRec.LICENSE) <> QPTrim$(fptxtLicNum.Text) Then
    ChangeFlag = True
    If fptxtLicNum.Enabled = True Then
      fptxtLicNum.SetFocus
    End If
    GoTo ChangeFound
  End If

  fptxtCode(0).Col = 0
  If QPTrim$(CustRec.BILLCAT1) <> QPTrim$(fptxtCode(0).ColText) Then
    ChangeFlag = True
    If fptxtCode(0).Enabled = True Then
      fptxtCode(0).SetFocus
    End If
    GoTo ChangeFound
  End If

  'strip out $ and , before comparing
  ThisVal = ReplaceString(fptxtRev(0).Text, "$", "")
  ThisVal = ReplaceString(ThisVal, ",", "")
40:
  If CustRec.REV1 <> Val(ThisVal) Then '(fptxtRev(0).Text) Then
    ChangeFlag = True
    If fptxtRev(0).Enabled = True Then
      fptxtRev(0).SetFocus
    End If
    GoTo ChangeFound
  End If

  fptxtCode(1).Col = 0
  If QPTrim$(CustRec.BILLCAT2) <> QPTrim$(fptxtCode(1).ColText) Then
    ChangeFlag = True
    If fptxtCode(1).Enabled = True Then
      fptxtCode(1).SetFocus
    End If
    GoTo ChangeFound
  End If

  ThisVal = ReplaceString(fptxtRev(1).Text, "$", "")
  ThisVal = ReplaceString(ThisVal, ",", "")
  
  If CustRec.REV2 <> Val(ThisVal) Then 'Val(fptxtRev(1).Text) Then
    ChangeFlag = True
    If fptxtRev(1).Enabled = True Then
      fptxtRev(1).SetFocus
    End If
    GoTo ChangeFound
  End If

  fptxtCode(2).Col = 0
  If QPTrim$(CustRec.BILLCAT3) <> QPTrim$(fptxtCode(2).ColText) Then
    ChangeFlag = True
    If fptxtCode(2).Enabled = True Then
      fptxtCode(2).SetFocus
    End If
    GoTo ChangeFound
  End If
50:
  
  ThisVal = ReplaceString(fptxtRev(2).Text, "$", "")
  ThisVal = ReplaceString(ThisVal, ",", "")
  
  If CustRec.REV3 <> Val(ThisVal) Then 'Val(fptxtRev(2).Text) Then
    ChangeFlag = True
    If fptxtRev(2).Enabled = True Then
      fptxtRev(2).SetFocus
    End If
    GoTo ChangeFound
  End If

  fptxtCode(3).Col = 0
  If QPTrim$(CustRec.BILLCAT4) <> QPTrim$(fptxtCode(3).ColText) Then
    ChangeFlag = True
    If fptxtCode(3).Enabled = True Then
      fptxtCode(3).SetFocus
    End If
    GoTo ChangeFound
  End If

  ThisVal = ReplaceString(fptxtRev(3).Text, "$", "")
  ThisVal = ReplaceString(ThisVal, ",", "")
  
  If CustRec.REV4 <> Val(ThisVal) Then 'Val(fptxtRev(3).Text) Then
    ChangeFlag = True
    If fptxtRev(3).Enabled = True Then
      fptxtRev(3).SetFocus
    End If
    GoTo ChangeFound
  End If

  fptxtCode(4).Col = 0
  If QPTrim$(CustRec.BILLCAT5) <> QPTrim$(fptxtCode(4).ColText) Then
    ChangeFlag = True
    If fptxtCode(4).Enabled = True Then
      fptxtCode(4).SetFocus
    End If
    GoTo ChangeFound
  End If
60:
  
  ThisVal = ReplaceString(fptxtRev(4).Text, "$", "")
  ThisVal = ReplaceString(ThisVal, ",", "")
  
  If CustRec.REV5 <> Val(ThisVal) Then ' Val(fptxtRev(4).Text) Then
    ChangeFlag = True
    If fptxtRev(4).Enabled = True Then
      fptxtRev(4).SetFocus
    End If
    GoTo ChangeFound
  End If

  If QPTrim$(CustRec.CustLocation) <> Mid(fpcmbIOType.Text, 1, 1) Then
    ChangeFlag = True
    fpcmbIOType.SetFocus
    GoTo ChangeFound
  End If

  If TrimPhone(CustRec.WPHONE) <> TrimPhone(fpMaskPhone.Text) Then
    ChangeFlag = True
    fpMaskPhone.SetFocus
    GoTo ChangeFound
  End If

  If CustRec.VALID <> Date2Num(fptxtValidThru.Text) Then
    ChangeFlag = True
    If fptxtValidThru.Enabled = True Then
      fptxtValidThru.SetFocus
    End If
    GoTo ChangeFound
  End If
  
  If QPTrim$(CustRec.IssueLicense) <> "Y" And QPTrim$(CustRec.IssueLicense) <> "N" And _
    fpcmbPrintNext.Text = "" Then
    GoTo ILISOK
  End If
  
  If QPTrim$(CustRec.IssueLicense) <> Mid(fpcmbPrintNext.Text, 1, 1) Then
    ChangeFlag = True
    If fpcmbPrintNext.Enabled = True Then
      fpcmbPrintNext.SetFocus
    End If
    GoTo ChangeFound
  End If
  
ILISOK: 'I = Issue and L = License
70:

  If QPTrim$(CustRec.Inactive) = "N" Then
    ThisInactive = "Active"
  Else
    ThisInactive = "Inactive"
  End If
  
  If ThisInactive <> QPTrim$(fpcmbInactiveAcct.Text) Then
    ChangeFlag = True
    If fpcmbInactiveAcct.Enabled = True Then
      fpcmbInactiveAcct.SetFocus
    End If
    GoTo ChangeFound
  End If
  
InactiveIsBlank:

  If CustRec.Prorate = 0 Then GoTo ProrateIsBlank
  
  If CustRec.Prorate <> Val(ReplaceString(fptxtProrate.Text, "%", "")) Then
    ChangeFlag = True
    If fptxtProrate.Enabled = True Then
      fptxtProrate.SetFocus
    End If
    GoTo ChangeFound
  End If
  
ProrateIsBlank:
80:
ChangeFound:
  If ChangeFlag = True Then
    ChangeFlag = False
    ItemChangeFlag = True 'global sends message to the customer list
    '(if it is open) to wait before bringing up the new customer's
    'data
    DoWhatFlag = PromptSaveChanges(Me)
90:
    Select Case DoWhatFlag
    Case SaveChangeOptions1.scoSaveChanges
      Call cmdSave_Click
      Exit Sub 'don't exit screen
    Case SaveChangeOptions1.scoReviewChanges 'review is just bringing back the current form
      If Exist("custlistopen.dat") Then 'the new data is being called up
      'from double clicking a name on the customer list option...so close it
      'now and kill the identifier file because the user wants to wait and
      'look over the current data before bringing up the new data
        Unload frmBLCustomerList
        KillFile "custlistopen.dat"
      End If
      Exit Sub
    Case SaveChangeOptions1.scoAbandonChanges 'abandon
      If Exist("custlistopen.dat") Then
        ItemChangeFlag = False 'if the customer list is open then
        'close it and kill the identifier file because the user
        'wants to just exit this screen
        KillFile "custlistopen.dat"
        Exit Sub
      End If
      'if the user arrived at this screen by way of editing
      'an existing customer then the CustomerLookup screen will be open
      'so when he abandons this screen he returns to the
      'CustomerLookup screen...if this screen opened from the Add
      'Customer option then when the user abandons this screen then
      'the Customer Maintenance Menu comes up
      KillFile "customeredit.dat"
      FromCustEdit = True
      frmBLCustomerLookup.Show
      Call frmBLCustomerLookup.cmdSearch_Click
      DoEvents
      Unload frmBLCustEdit
      Exit Sub
    Case Else:
    End Select
  End If
CustNumIsZero:
100:
  If Exist("custlistopen.dat") Then
    KillFile ("custlistopen.dat")
    Exit Sub
  End If
  
  If Not Exist("custlookup.dat") Then
    frmBLCustMaintMenu.Show
  Else
    frmBLCustomerLookup.Show
  End If

110:
  If Exist("customeredit.dat") Then
    KillFile "customeredit.dat"
  End If
  DoEvents
  FromCustEdit = True
  
120:
  Unload frmBLCustEdit
  
  Exit Sub
  
BackGround:
130:
  For x = 1 To 35 'Me.Controls.Count...changes all screen
  'control backgrounds to white
    If Me.Controls(x).BackColor = &HFFFF& Then
      Me.Controls(x).BackColor = &H80000005
    End If
  Next x
  
  Return
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustEdit", "cmdExit_Click", Erl)
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

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "F1 &Turn Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    lblBalloon.Visible = True
    cmdHelp.ToolTipText = ""
    fptxtBillingName.ToolTipText = ""
    fptxtBusName.ToolTipText = ""
    fptxtAddress1.ToolTipText = ""
    fptxtAddress2.ToolTipText = ""
    fptxtCity.ToolTipText = ""
    fptxtState.ToolTipText = ""
    fptxtZip.ToolTipText = ""
    fptxtLicNum.ToolTipText = ""
    cmdList.ToolTipText = ""
    fptxtCustNum.ToolTipText = ""
    fptxtSearchName.ToolTipText = ""
    fptxtContact.ToolTipText = ""
    fptxtCode(0).ToolTipText = ""
    fptxtCode(1).ToolTipText = ""
    fptxtCode(2).ToolTipText = ""
    fptxtCode(3).ToolTipText = ""
    fptxtCode(4).ToolTipText = ""
    fptxtRev(0).ToolTipText = ""
    fptxtRev(1).ToolTipText = ""
    fptxtRev(2).ToolTipText = ""
    fptxtRev(3).ToolTipText = ""
    fptxtRev(4).ToolTipText = ""
    cmdCustList.ToolTipText = ""
    fpcmbIOType.ToolTipText = ""
    fpMaskPhone.ToolTipText = ""
    fptxtValidThru.ToolTipText = ""
    fpcmbPrintNext.ToolTipText = ""
    fpcmbInactiveAcct.ToolTipText = ""
    fptxtProrate.ToolTipText = ""
    cmdDelete.ToolTipText = ""
    cmdExit.ToolTipText = ""
    cmdSave.ToolTipText = ""
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
'    cmdHelp.ToolTipText = "Click on this button to activate informational balloons for each field."
'    fptxtBillingName.ToolTipText = "Enter the name this business uses for offical business license related correspondance."
'    fptxtBusName.ToolTipText = "Enter the name by which this business is known to it's customers."
'    fptxtAddress1.ToolTipText = "Enter the primary address here (generally the street address)."
'    fptxtAddress2.ToolTipText = "Enter the secondary address here (generally a post office box or suite name, etc.)."
'    fptxtCity.ToolTipText = "Enter the city where this business receives it's mail."
'    fptxtState.ToolTipText = "Enter the state where this business receives it's mail."
'    fptxtZip.ToolTipText = "Enter a five digit or nine digit postal code for this business."
'    fptxtLicNum.ToolTipText = "Enter the license number here."
'    cmdList.ToolTipText = "Press this button to bring up a list of currently saved license numbers. "
'    fptxtCustNum.ToolTipText = "This number is automatically assigned. This field cannot be edited."
'    fptxtSearchName.ToolTipText = "Enter an abbreviated name for this business which is used to make it easier for a computer search for this business."
'    fptxtContact.ToolTipText = "Enter a primary contact name for this business."
'    fptxtCode(0).ToolTipText = "Select a category code from the drop down box."
'    fptxtCode(1).ToolTipText = "Select a category code from the drop down box."
'    fptxtCode(2).ToolTipText = "Select a category code from the drop down box."
'    fptxtCode(3).ToolTipText = "Select a category code from the drop down box."
'    fptxtCode(4).ToolTipText = "Select a category code from the drop down box."
'    fptxtRev(0).ToolTipText = "If the category uses a multiplier then enter the number of profit centers this business has. If the category uses step rate then enter annual revenue. Otherwise enter zero."
'    fptxtRev(1).ToolTipText = "If the category uses a multiplier then enter the number of profit centers this business has. If the category uses step rate then enter annual revenue. Otherwise enter zero."
'    fptxtRev(2).ToolTipText = "If the category uses a multiplier then enter the number of profit centers this business has. If the category uses step rate then enter annual revenue. Otherwise enter zero."
'    fptxtRev(3).ToolTipText = "If the category uses a multiplier then enter the number of profit centers this business has. If the category uses step rate then enter annual revenue. Otherwise enter zero."
'    fptxtRev(4).ToolTipText = "If the category uses a multiplier then enter the number of profit centers this business has. If the category uses step rate then enter annual revenue. Otherwise enter zero."
'    cmdCustList.ToolTipText = "Press this button to bring up a complete list of customers currently saved."
'    fpcmbIOType.ToolTipText = "Indicate the proximity to the city limits of this business."
'    fpMaskPhone.ToolTipText = "Enter the phone number for this business."
'    fptxtValidThru.ToolTipText = "Enter the date on which the current business license for this business will expire."
'    fpcmbPrintNext.ToolTipText = "Select Yes and this customer will be included in the next business license processing regardless of it's expiration date."
'    fpcmbInactiveAcct.ToolTipText = "Select Active for customers that will be processed for a business license. Otherwise select Inactive."
'    fptxtProrate.ToolTipText = "If this business is coming on line in the middle of the year then enter what percentage of the total annual fee will apply during the partial year."
'    cmdDelete.ToolTipText = "Press here to remove this business from the current list of businesses."
'    cmdExit.ToolTipText = "Press here to exit this screen."
'    cmdSave.ToolTipText = "Press here to commit the data on this screen to memory."
    
  End If

End Sub

Private Sub cmdSave_Click()
  Dim DHandle As Integer
  Dim CustRec As ARCustRecType
  Dim SaveHere As Integer
  Dim NumOfCustRecs As Integer
  Dim IdxFlag As Boolean
  Dim ReIndexFlag As Boolean
  Dim x As Integer
  Dim ThisRev$
  Dim Answer As VbMsgBoxResult
  On Error GoTo ERRORSTUFF
  
  Call CleanUpCodeEntries
  
  If GCustNum > 0 And GCustNum <> QPTrim$(fptxtCustNum.Text) Then
    frmBLMessageBoxJrWOpts.Label1.Caption = "The customer number, " & QPTrim$(fptxtCustNum.Text) & ", on the screen and the saved customer number, " & CStr(GCustNum) & ", for this customer are not the same. Do you wish to continue anyway?"
    frmBLMessageBoxJrWOpts.Label1.Top = 600
    frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
    frmBLMessageBoxJrWOpts.cmdExit.Text = "Esc Edit"
    frmBLMessageBoxJrWOpts.Show vbModal
    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
      Unload frmBLMessageBoxJrWOpts
      MainLog ("User warned that the customer number " & QPTrim$(fptxtCustNum.Text) & " on the screen and the saved customer number, " & CStr(GCustNum) & ", for this customer were not the same and they elected to save anyway.")
    Else
      Unload frmBLMessageBoxJrWOpts
      Close
      Exit Sub
    End If
  End If
    
  If GCustNum > 0 Then
    If EmpInLicProcess(CStr(GCustNum)) = True And Exist("artmppst.dat") Then
      frmBLMessageBoxJrWOpts.Label1.Caption = "This customer is currently being processed for a business license renewal. If you wish to continue to save the edit for this customer all temporary business license files will be deleted. You will be required to re-process the business license fees operation. Do you wish to continue to save anyway?"
      frmBLMessageBoxJrWOpts.Label1.Top = 500
      frmBLMessageBoxJrWOpts.Label1.Height = 1300
      frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
      frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
      frmBLMessageBoxJrWOpts.Show vbModal
      If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
        Unload frmBLMessageBoxJrWOpts
        Close
        Exit Sub
      Else
        Unload frmBLMessageBoxJrWOpts
        KillFile "artmppst.dat"
        KillFile "artmplic.dat"
        KillFile "licprnOK.dat"
        MainLog ("User warned that continuing to edit customer # " + CStr(GCustNum) + " would delete the 'attmppst.dat' and 'artmplic.dat' files and the user elected to continue anyway.")
      End If
    End If
  End If
        
  If GCustNum > 0 Then
    If EmpInPenProcess(CStr(GCustNum)) = True And Exist("artmppst.dat") Then
      frmBLMessageBoxJrWOpts.Label1.Caption = "This customer is currently being processed for a penalty fee. If you wish to continue to save the edit for this customer the temporary penalty fee file will be deleted. You will be required to re-process the penalty fees operation. Do you wish to continue to save anyway?"
      frmBLMessageBoxJrWOpts.Label1.Top = 500
      frmBLMessageBoxJrWOpts.Label1.Height = 1300
      frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
      frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
      frmBLMessageBoxJrWOpts.Show vbModal
      If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
        Unload frmBLMessageBoxJrWOpts
        Close
        Exit Sub
      Else
        Unload frmBLMessageBoxJrWOpts
        KillFile "artmppen.dat"
        MainLog ("User warned that continuing to edit customer # " + CStr(GCustNum) + " would delete the 'attmppen.dat' file and the user elected to continue anyway.")
      End If
    End If
  End If
  
  'The next code examines the entries in the categories
  'fields looking for "F" which is flat rate...if F is found
  'and a value is in the Revenue field then a pop up comes up
  'telling the user that even though they entered a value for the
  'F category it isn't necessary because flat rate fees are
  'determined at the category edit screen...the program gives
  'the user the option of saving the unnecessary value or changing
  'it to zero for them and then saving
  For x = 0 To 4
    fptxtCode(x).Col = 2
    If QPTrim$(fptxtCode(x).ColText) = "F" Then
      If Val(fptxtRev(x).Text) <> 0 Then
        fptxtCode(x).Col = 0
        frmBLMessageBoxJrWOpts.Label1.Caption = "Code number " + fptxtCode(x).ColText + " uses a flat rate to calculate fees. Therefore no value is necessary for this category. Do you wish to reset your revenue entry for this category to zero?"
        frmBLMessageBoxJrWOpts.Label1.Top = 700
        frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Make Zero"
        frmBLMessageBoxJrWOpts.cmdExit.Text = "Esc Save As Is"
        frmBLMessageBoxJrWOpts.Show vbModal
        If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
'          fptxtRev(x).Text = Using("###0", 0)
          Unload frmBLMessageBoxJrWOpts
        Else
          Unload frmBLMessageBoxJrWOpts
          MainLog ("User notified that the category they are saving, " + fptxtCode(x).Text + ", does not need a value for revenue because it uses a flat rate. The user opted to save the revenue as, " + fptxtRev(x).Text + " anyway.")
        End If
      End If
    ElseIf QPTrim$(fptxtCode(x).ColText) = "M" Then
      If Val(fptxtRev(x).Text) = 0 Then
        fptxtCode(x).Col = 0
        frmBLMessageBoxJrWOpts.Label1.Caption = "Code number " + fptxtCode(x).ColText + " uses a multiplier to calculate fees. A unit value is necessary or no fee will be charged for this category. Press ESC if you wish to edit this category or press F10 to save it with no value."
        frmBLMessageBoxJrWOpts.Label1.Top = 700
        frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
        frmBLMessageBoxJrWOpts.cmdExit.Text = "Esc Edit"
        frmBLMessageBoxJrWOpts.Show vbModal
        If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
          Unload frmBLMessageBoxJrWOpts
          MainLog ("User notified that the category they are saving, " + fptxtCode(x).Text + ", needs a value for revenue because it uses a multiplier rate. The user opted to save the revenue as, " + fptxtRev(x).Text + " anyway.")
        Else
          Unload frmBLMessageBoxJrWOpts
          Close
          If fptxtRev(x).Enabled = True Then
            fptxtRev(x).SetFocus
          End If
          Exit Sub
        End If
      End If
    ElseIf QPTrim$(fptxtCode(x).ColText) = "S" Then
      If Val(ReplaceString(fptxtRev(x).Text, "$", "")) = 0 Then
        fptxtCode(x).Col = 0
        frmBLMessageBoxJrWOpts.Label1.Caption = "Code number " + fptxtCode(x).ColText + " uses a step rate to calculate fees. A revenue value is necessary or no fee will be charged for this category. Press ESC if you wish to edit this category or press F10 to save it with no value."
        frmBLMessageBoxJrWOpts.Label1.Top = 600
        frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
        frmBLMessageBoxJrWOpts.cmdExit.Text = "Esc Edit"
        frmBLMessageBoxJrWOpts.Show vbModal
        If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
          Unload frmBLMessageBoxJrWOpts
          MainLog ("User notified that the category they are saving, " + fptxtCode(x).Text + ", needs a value for revenue because it uses a step rate. The user opted to save the revenue as, " + fptxtRev(x).Text + " anyway.")
        Else
          Unload frmBLMessageBoxJrWOpts
          Close
          If fptxtRev(x).Enabled = True Then
            fptxtRev(x).SetFocus
          End If
          Exit Sub
        End If
      End If
    End If
  Next x
  
  'error check for eliminating redundant category codes
  If Check4DupCodes = True Then Exit Sub
  'error check that looks for values in the revenue field
  'but no values in the code field
  If CodeAmtsOK = False Then Exit Sub
    
  ReIndexFlag = False
  IdxFlag = False
  
  If Check4ValidLic(QPTrim$(fptxtLicNum.Text)) = False Then
    Exit Sub
  End If
  
  If Mid(fpcmbPrintNext.Text, 1, 1) = "Y" And Mid(fpcmbInactiveAcct.Text, 1, 1) = "I" Then
    fpcmbPrintNext.BackColor = &H80FFFF
    frmBLMessageBoxJr.Label1.Caption = "If a customer is set to be 'Inactive' then 'Set Renewal Flag (Y/N)?' cannot be set to 'Yes'. 'Set Renewal Flag (Y/N)?' reset to 'No'."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
    fpcmbPrintNext.BackColor = &HFFFFFF
    fpcmbPrintNext.Text = "No"
    fpcmbPrintNext.SetFocus
    Close
    Exit Sub
  End If
  
  If GCustNum = 0 Then
    IdxFlag = True
    AddFlag = True
  Else
    OpenCustFile DHandle
    SaveHere = GCustNum
    Get DHandle, SaveHere, CustRec
  End If
  
  If QPTrim$(fptxtBusName.Text) = "" Then
    fptxtBusName.BackColor = &H80FFFF
    frmBLMessageBoxJr.Label1.Caption = "Please enter a business name for this customer."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtBusName.BackColor = &HFFFFFF
    If fptxtBusName.Enabled = True Then
      fptxtBusName.SetFocus
    End If
    Close
    Exit Sub
  End If
  
  CustRec.CustName = QPTrim$(fptxtBusName.Text)
  
  If QPTrim$(fptxtAddress1.Text) = "" And QPTrim$(fptxtAddress2.Text) = "" Then
    fptxtAddress1.BackColor = &H80FFFF
    frmBLMessageBoxJr.Label1.Caption = "Please enter an address for this customer."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtAddress1.BackColor = &HFFFFFF
    fptxtAddress1.SetFocus
    Close
    Exit Sub
  End If
  
  CustRec.ADDRESS1 = QPTrim$(fptxtAddress1.Text)
  CustRec.ADDRESS2 = QPTrim$(fptxtAddress2.Text)
  
  If QPTrim$(fptxtCity.Text) = "" Then
    fptxtCity.BackColor = &H80FFFF
    frmBLMessageBoxJr.Label1.Caption = "Please enter the city for this business."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtCity.BackColor = &HFFFFFF
    fptxtCity.SetFocus
    Close
    Exit Sub
  End If
  
  CustRec.City = QPTrim$(fptxtCity.Text)
  
  If QPTrim$(fptxtState.Text) = "" Then
    fptxtState.BackColor = &H80FFFF
    frmBLMessageBoxJr.Label1.Caption = "Please enter the state for this business."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtState.BackColor = &HFFFFFF
    fptxtState.SetFocus
    Close
    Exit Sub
  End If
  
  CustRec.State = QPTrim$(fptxtState.Text)
  
  CustRec.ZipCode = QPTrim$(fptxtZip.Text)
  
  If QPTrim$(fptxtBillingName.Text) = "" Then
    fptxtBillingName.BackColor = &H80FFFF
    frmBLMessageBoxJr.Label1.Caption = "Please enter a billing name for this business."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtBillingName.BackColor = &HFFFFFF
    If fptxtBillingName.Enabled = True Then
      fptxtBillingName.SetFocus
    End If
    Close
    Exit Sub
  End If
  
  If QPTrim$(fptxtBillingName.Text) <> TempBillName Then
    lblSave.Visible = True
    DoEvents
    ReIndexFlag = True 'this value is used to
    'create the customer index sort so if it has
    'changed the index needs to be redone
  End If
  CustRec.BillName = QPTrim$(fptxtBillingName.Text)
  
  If GCustNum > 0 Then 'if this is an existing business and the
  'license numbers are set to permanent then if the number is changed
  'this warning comes up
    If PermNum = True Then
      If QPTrim$(fptxtLicNum.Text) <> QPTrim$(CustRec.LICENSE) Then
        frmBLMessageBoxJrWOpts.Label1.Caption = "Currently business license numbers are permanently set (option selected on Town Setup screen). However, a change has been made to this business license number. Are you sure you want to continue?"
        frmBLMessageBoxJrWOpts.Label1.Top = 600
        frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
        frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
        frmBLMessageBoxJrWOpts.Show vbModal
        If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
          Close
          Unload frmBLMessageBoxJrWOpts
          fptxtLicNum.Text = QPTrim(CustRec.LICENSE)
          If fptxtLicNum.Enabled = True Then
            fptxtLicNum.SetFocus
          End If
          lblSave.Visible = False
          Exit Sub
        Else
          MainLog ("User warned that they are changing this customer's business license number even though the setting is for permanent numbers. They elected to save anyway.")
          Unload frmBLMessageBoxJrWOpts
        End If
      End If
    End If
  End If
  
  If QPTrim$(fptxtLicNum.Text) = "" Then
    fptxtLicNum.BackColor = &H80FFFF
    frmBLMessageBoxJr.Label1.Caption = "Please enter a business license number for this business."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtLicNum.BackColor = &HFFFFFF
    If fptxtLicNum.Enabled = True Then
      fptxtLicNum.SetFocus
    End If
    lblSave.Visible = False
    Close
    Exit Sub
  End If
    
  If QPTrim$(fptxtSearchName.Text) = "" Then
    fptxtSearchName.BackColor = &H80FFFF
    frmBLMessageBoxJr.Label1.Caption = "Please enter a search name for this business."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtSearchName.BackColor = &HFFFFFF
    If fptxtSearchName.Enabled = True Then
      fptxtSearchName.SetFocus
    End If
    lblSave.Visible = False
    Close
    Exit Sub
  End If
  
  If QPTrim$(fptxtSearchName.Text) <> TempSearchName Then
    lblSave.Visible = True
    DoEvents
    ReIndexFlag = True
  End If
  
  CustRec.SortName = QPTrim$(fptxtSearchName.Text)
  CustRec.Contact = QPTrim$(fptxtContact.Text)
  CustRec.ServAdd = QPTrim$(fptxtServAdd.Text)
  CustRec.SSNFID = QPTrim$(fptxtSSNFID.Text)
  For x = 0 To 4
    If QPTrim$(fptxtCode(x).Text) <> "" Then
      Exit For
    End If
  Next x
  
  If x = 5 Then
    fptxtCode(0).BackColor = &H80FFFF
    frmBLMessageBoxJr.Label1.Caption = "Please enter at least one code for this business."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    fptxtCode(0).BackColor = &HFFFFFF
    If fptxtCode(0).Enabled = True Then
      fptxtCode(0).SetFocus
    End If
    lblSave.Visible = False
    Close
    Exit Sub
  End If
  
  fptxtCode(0).Col = 0
  CustRec.BILLCAT1 = QPTrim$(fptxtCode(0).ColText)
  fptxtCode(0).Col = 1
  CustRec.DESC1 = QPTrim$(fptxtCode(0).ColText)
  ThisRev = ReplaceString$(fptxtRev(0), "$", "")
  ThisRev = ReplaceString$(ThisRev, ",", "")
  CustRec.REV1 = Val(ThisRev)
  fptxtCode(1).Col = 0
  CustRec.BILLCAT2 = QPTrim$(fptxtCode(1).ColText)
  fptxtCode(1).Col = 1
  CustRec.DESC2 = QPTrim$(fptxtCode(1).ColText)
  ThisRev = ReplaceString$(fptxtRev(1), "$", "")
  ThisRev = ReplaceString$(ThisRev, ",", "")
  CustRec.REV2 = Val(ThisRev)
  fptxtCode(2).Col = 0
  CustRec.BILLCAT3 = QPTrim$(fptxtCode(2).ColText)
  fptxtCode(2).Col = 1
  CustRec.DESC3 = QPTrim$(fptxtCode(2).ColText)
  ThisRev = ReplaceString$(fptxtRev(2), "$", "")
  ThisRev = ReplaceString$(ThisRev, ",", "")
  CustRec.REV3 = Val(ThisRev)
  fptxtCode(3).Col = 0
  CustRec.BILLCAT4 = QPTrim$(fptxtCode(3).ColText)
  fptxtCode(3).Col = 1
  CustRec.DESC4 = QPTrim$(fptxtCode(3).ColText)
  ThisRev = ReplaceString$(fptxtRev(3), "$", "")
  ThisRev = ReplaceString$(ThisRev, ",", "")
  CustRec.REV4 = Val(ThisRev)
  fptxtCode(4).Col = 0
  CustRec.BILLCAT5 = QPTrim$(fptxtCode(4).ColText)
  fptxtCode(4).Col = 1
  CustRec.DESC5 = QPTrim$(fptxtCode(4).ColText)
  ThisRev = ReplaceString$(fptxtRev(4), "$", "")
  ThisRev = ReplaceString$(ThisRev, ",", "")
  CustRec.REV5 = Val(ThisRev)
  CustRec.CustLocation = QPTrim$(fpcmbIOType.Text)
  CustRec.WPHONE = fpMaskPhone
  
  CustRec.LICENSE = QPTrim$(fptxtLicNum.Text)
  
  If Date2Num(fptxtValidThru.Text) = 0 Then
    fptxtValidThru.BackColor = &H80FFFF
    frmBLMessageBoxJr.Label1.Caption = "Please enter a valid through date for this business."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtValidThru.BackColor = &HFFFFFF
    If fptxtValidThru.Enabled = True Then
      fptxtValidThru.SetFocus
    End If
    lblSave.Visible = False
    Close
    Exit Sub
  ElseIf CheckValDate(fptxtValidThru.Text) = False Then
    fptxtValidThru.BackColor = &H80FFFF
    frmBLMessageBoxJr.Label1.Caption = "The valid through date entered for this customer is not valid. Please enter a valid date."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtValidThru.BackColor = &HFFFFFF
    If fptxtValidThru.Enabled = True Then
      fptxtValidThru.SetFocus
    End If
    lblSave.Visible = False
    Close
    Exit Sub
  ElseIf Mid(fptxtValidThru.Text, 7, 4) > 2050 Or Mid(fptxtValidThru.Text, 7, 4) < 1979 Then
    fptxtValidThru.BackColor = &H80FFFF
    frmBLMessageBoxJr.Label1.Caption = "Please enter a valid through date between 1979 and 2050."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtValidThru.BackColor = &HFFFFFF
    If fptxtValidThru.Enabled = True Then
      fptxtValidThru.SetFocus
    End If
    lblSave.Visible = False
    Close
    Exit Sub
  Else
    CustRec.VALID = Date2Num(fptxtValidThru.Text)
  End If
  
  If QPTrim$(fpcmbPrintNext.Text) = "" Then
    fpcmbPrintNext.BackColor = &H80FFFF
    frmBLMessageBoxJr.Label1.Caption = "Please select Yes or No in the 'Set Renewal Flag (Y/N)?' field."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fpcmbPrintNext.BackColor = &HFFFFFF
    If fpcmbPrintNext.Enabled = True Then
      fpcmbPrintNext.SetFocus
    End If
    lblSave.Visible = False
    Close
    Exit Sub
  End If
  
  CustRec.IssueLicense = QPTrim$(fpcmbPrintNext.Text)
  
  If QPTrim$(fpcmbInactiveAcct.Text) = "" Then
    fpcmbInactiveAcct.BackColor = &H80FFFF
    frmBLMessageBoxJr.Label1.Caption = "Please indicate if this business is Active or Inactive."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fpcmbInactiveAcct.BackColor = &HFFFFFF
    If fpcmbInactiveAcct.Enabled = True Then
      fpcmbInactiveAcct.SetFocus
    End If
    lblSave.Visible = False
    Close
    Exit Sub
  End If
  
  If QPTrim$(fpcmbInactiveAcct.Text) = "Active" Then
    CustRec.Inactive = "N"
  Else
    CustRec.Inactive = "Y"
  End If
  
  CustRec.Prorate = Val(ReplaceString(fptxtProrate.Text, "%", ""))
  If CustRec.Prorate > 100 Then
    fptxtProrate.BackColor = &H80FFFF
    frmBLMessageBoxJr.Label1.Caption = "Amounts greater than 100% in the 'Prorate' field are invalid. Please enter a 'Prorate' amount of 100% or less."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
    fptxtProrate.BackColor = &HFFFFFF
    If fptxtProrate.Enabled = True Then
      fptxtProrate.SetFocus
    End If
    Close
    Exit Sub
  ElseIf CustRec.Prorate < 0 Then
    fptxtProrate.BackColor = &H80FFFF
    frmBLMessageBoxJr.Label1.Caption = "Amounts less than 0% in the 'Prorate' field are invalid. Please enter a 'Prorate' amount of 0% or more."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
    fptxtProrate.BackColor = &HFFFFFF
    If fptxtProrate.Enabled = True Then
      fptxtProrate.SetFocus
    End If
    Close
    Exit Sub
  End If
  
'************
  If PWUser = "Sosoft Support" Then
    CustRec.CustNumb = QPTrim$(fptxtCustNum.Text)
  End If
  
  If GCustNum = 0 Then
    OpenCustFile DHandle
    NumOfCustRecs = LOF(DHandle) \ Len(CustRec)
    SaveHere = NumOfCustRecs + 1
    CustRec.CustNumb = SaveHere
    CustRec.Fee1 = 0
    CustRec.FeeLicBal1 = 0
    CustRec.FeeLicPay1 = 0
    CustRec.Fee2 = 0
    CustRec.FeeLicBal2 = 0
    CustRec.FeeLicPay2 = 0
    CustRec.Fee3 = 0
    CustRec.FeeLicBal3 = 0
    CustRec.FeeLicPay3 = 0
    CustRec.Fee4 = 0
    CustRec.FeeLicBal4 = 0
    CustRec.FeeLicPay4 = 0
    CustRec.Fee5 = 0
    CustRec.FeeLicBal5 = 0
    CustRec.FeeLicPay5 = 0
    CustRec.IssuanceFee = 0
    CustRec.FeeAmt = 0
    CustRec.AcctBal = 0
    CustRec.Deleted = ""          '(yY)=deleted, anything else isn't
    CustRec.FirstTrans = 0
    CustRec.LastTrans = 0
    CustRec.LicBal = 0
    CustRec.FeeBal = 0
    CustRec.PenBal = 0
    CustRec.RoomtoGrow = ""
    CustRec.ChkByte = ""
    CustRec.IssuanceBal = 0
    CustRec.IssuancePay = 0
  End If
  
  Put DHandle, SaveHere, CustRec
  Close DHandle
  Call LogSaves(SaveHere)
  
  If IdxFlag = True Or ReIndexFlag = True Then
    lblSave.Visible = True
    DoEvents
    IdxFlag = False
    ReIndexFlag = False
    lblSave.BackColor = &HC0FFFF
    lblSave.Caption = "Indexing Names"
    DoEvents
    Call CreateCustNameIdx
    lblSave.BackColor = &H80C0FF
    lblSave.Caption = "Indexing Numbers"
    DoEvents
    Call CreateLicNumIdx
    lblSave.BackColor = &H80C0FF
    lblSave.Caption = "Indexing License Numbers"
    DoEvents
    Call CreateCustNumIdx
    lblSave.BackColor = &HC0FFFF
    lblSave.Caption = "Indexing Sort Names"
    DoEvents
    Call CreateCustSearchNameIdx
    lblSave.Visible = False
    DoEvents
  End If
  lblSave.Visible = False
  frmBLSucSave.Label1.Caption = "Data for " + QPTrim$(fptxtBusName.Text) + " has been saved."
  frmBLSucSave.Label1.Top = 700
  frmBLSucSave.Show vbModal
  
  If Exist("custlistopen.dat") Then '
    KillFile ("custlistopen.dat")
    ItemChangeFlag = False
    If fptxtBusName.Enabled = True Then
      If fptxtBusName.Enabled = True Then
        fptxtBusName.SetFocus
      End If
    Else
      fptxtAddress1.SetFocus
    End If
    Close
    Exit Sub
  End If
  
  If AddFlag = True Then 'entering a list of several items is tedious
  'if after each save the program returns to the menu so this feature allows
  'the user to speed up the entry process
    AddFlag = False
    frmBLMessageBoxJrWOpts.Label1.Caption = "Do you wish to add another new customer?"
    frmBLMessageBoxJrWOpts.Label1.Top = 900
    frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Add New"
    frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
    frmBLMessageBoxJrWOpts.Show vbModal
    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
      Unload frmBLMessageBoxJrWOpts
      fptxtBillingName.SetFocus
      Call LoadMe
      Close
      Exit Sub
    Else
      Unload frmBLMessageBoxJrWOpts
      frmBLCustMaintMenu.Show
      DoEvents
      KillFile ("customeredit.dat")
      Unload frmBLCustEdit
      Close
      Exit Sub
    End If
  Else 'just editing an existing customer...sends user back to menu upon
  'completion
    FromCustEdit = True
'    Call frmBLCustomerLookup.RefreshSearchList
    frmBLCustomerLookup.Show
    Call frmBLCustomerLookup.cmdSearch_Click
    DoEvents
    KillFile ("customeredit.dat")
    Unload frmBLCustEdit
  End If
  
  Close
  
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustEdit", "cmdSave", Erl)
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

Private Sub cmdTransHist_Click()
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim ThisCustNum As Integer
  
  If QPTrim$(fptxtCustNum.Text) = "" Then
    Exit Sub
  End If
  
  ThisCustXNum = CInt(fptxtCustNum.Text)
  
  If Check4ValidCustNum(QPTrim$(fptxtCustNum.Text)) = False Then
    frmBLMessageBoxJr.Label1.Caption = "The customer number entered is not valid. Please enter a valid customer number."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    Call ClearScreen
    If fptxtBusName.Enabled = True Then
      fptxtBusName.SetFocus
    End If
    Exit Sub
  End If
  
  If Exist("transhistjr.dat") Then Exit Sub
  
  If ThisCustXNum > 0 Then
    OpenCustFile CustHandle
    Get CustHandle, ThisCustXNum, CustRec
    Close CustHandle
    If CustRec.LastTrans = 0 Then
      frmBLMessageBoxJr.Label1.Caption = "This customer has no transaction activity."
      frmBLMessageBoxJr.Label1.Top = 800
      frmBLMessageBoxJr.Show vbModal
      If fptxtBusName.Enabled = True Then
        fptxtBusName.SetFocus
      End If
      Exit Sub
    Else
      DoEvents
      Load frmBLTransHistJr
      DoEvents
      frmBLTransHistJr.Show vbModal
      DoEvents
      Me.Hide
    End If
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
    Case vbKeyF10:
      SendKeys "%S"
      Call cmdSave_Click
      KeyCode = 0
    Case vbKeyF7:
      SendKeys "%L"
      Call cmdCustList_Click
      KeyCode = 0
    Case vbKeyF5:
      SendKeys "%c"
      Call cmdList_Click
      KeyCode = 0
    Case vbKeyF4:
      SendKeys "%H"
      Call cmdTransHist_Click
      KeyCode = 0
    Case vbKeyF1:
      SendKeys "%T"
      Call cmdHelp_Click
      KeyCode = 0
    Case vbKeyF2:
      SendKeys "%D"
      Call cmdDelete_Click
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
      KillFile "customeredit.dat"
      ClearInUse PWcnt
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLCustEdit.")
      Call Terminate
      End
    End If
  End If
End Sub

Public Sub LoadMe()
  Dim x As Integer, y As Integer
  Dim CustFile As Integer
  Dim NumOfCustRecs As Integer
  Dim RecNo As Integer
  Dim GLFundLen As Integer
  Dim GLAcctLen As Integer
  Dim GLDetLen As Integer
  Dim ValidCust As Boolean
  Dim cnt As Integer
  Dim CustRecLen As Integer
  Dim CustNum$
  Dim CustIdxHandle As Integer
  Dim CustIdx As CustNameIdxType
  Dim CustIdxNum As Integer
  Dim One As Integer
  Dim DHandle As Integer
  Dim CustRec As ARCustRecType
  Dim TownHandle As Integer
  Dim TownRec As TownSetUpType
  Dim NumOfTownRecs As Integer
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeIdxRec As CatCodeIdxType
  Dim CodeIdxHandle As Integer
  Dim CodeIdxRecNum As Integer
  Dim CHandle As Integer
  Dim TotalAccts As Integer
  Dim CatCodeCnt As Integer
  Dim Nextx As Integer
  Dim RevType As String * 1
  Dim ThisText$

  On Error GoTo ERRORSTUFF
  NFlag = False
  
  'bizarre: had to put this code in to make these
  'fields load with the correct width
'  For x = 0 To 4
'    fptxtCode(x).Width = 6000
'  Next x
  
  lblBalloon.Visible = False
'  fptxtLicNum.Tag = "The default license number for new customers is one more than the existing largest saved license number. License numbers can be assigned permanently or they can change as desired depending on how the license number setting is saved on the Town Setup screen. If permanent license numbers are being used then if you change a saved license number you will be warned if you attempt to save it. "
  
'  cmdHelp.ToolTipText = "Click on this button to activate informational balloons for each field."
'  fptxtBillingName.ToolTipText = "Enter the name this business uses for offical business license related correspondance."
'  fptxtBusName.ToolTipText = "Enter the name by which this business is known to it's customers."
'  fptxtAddress1.ToolTipText = "Enter the primary address here (generally the street address)."
'  fptxtAddress2.ToolTipText = "Enter the secondary address here (generally a post office box or suite name, etc.)."
'  fptxtCity.ToolTipText = "Enter the city where this business receives it's mail."
'  fptxtState.ToolTipText = "Enter the state where this business receives it's mail."
'  fptxtZip.ToolTipText = "Enter a five digit or nine digit postal code for this business."
'  fptxtLicNum.ToolTipText = "Enter the license number here."
'  cmdList.ToolTipText = "Press this button to bring up a list of currently saved license numbers. "
'  fptxtCustNum.ToolTipText = "This number is automatically assigned. This field cannot be edited."
'  fptxtSearchName.ToolTipText = "Enter an abbreviated name for this business which is used to make it easier for a computer search for this business."
'  fptxtContact.ToolTipText = "Enter a primary contact name for this business."
'  fptxtCode(0).ToolTipText = "Select a category code from the drop down box."
'  fptxtCode(1).ToolTipText = "Select a category code from the drop down box."
'  fptxtCode(2).ToolTipText = "Select a category code from the drop down box."
'  fptxtCode(3).ToolTipText = "Select a category code from the drop down box."
'  fptxtCode(4).ToolTipText = "Select a category code from the drop down box."
'  fptxtRev(0).ToolTipText = "If the category uses a multiplier then enter the number of profit centers this business has. If the category uses step rate then enter annual revenue. Otherwise enter zero."
'  fptxtRev(1).ToolTipText = "If the category uses a multiplier then enter the number of profit centers this business has. If the category uses step rate then enter annual revenue. Otherwise enter zero."
'  fptxtRev(2).ToolTipText = "If the category uses a multiplier then enter the number of profit centers this business has. If the category uses step rate then enter annual revenue. Otherwise enter zero."
'  fptxtRev(3).ToolTipText = "If the category uses a multiplier then enter the number of profit centers this business has. If the category uses step rate then enter annual revenue. Otherwise enter zero."
'  fptxtRev(4).ToolTipText = "If the category uses a multiplier then enter the number of profit centers this business has. If the category uses step rate then enter annual revenue. Otherwise enter zero."
'  cmdCustList.ToolTipText = "Press this button to bring up a complete list of customers currently saved."
'  fpcmbIOType.ToolTipText = "Indicate the proximity to the city limits of this business."
'  fpMaskPhone.ToolTipText = "Enter the phone number for this business."
'  fptxtValidThru.ToolTipText = "Enter the date on which the current business license for this business will expire."
'  fpcmbPrintNext.ToolTipText = "Select Yes and this customer will be included in the next business license processing regardless of it's expiration date."
'  fpcmbInactiveAcct.ToolTipText = "Select Active for customers that will be processed for a business license. Otherwise select Inactive."
'  fptxtProrate.ToolTipText = "If this business is coming on line in the middle of the year then enter what percentage of the total annual fee will apply during the partial year."
'  cmdDelete.ToolTipText = "Press here to remove this business from the current list of businesses."
'  cmdExit.ToolTipText = "Press here to exit this screen."
'  cmdSave.ToolTipText = "Press here to commit the data on this screen to memory."

  If PWUser = "Sosoft Support" Then
    fptxtCustNum.ControlType = ControlTypeNormal
  End If
  
  'the following code allows the user to view the screen if a customer
  'is involved in either a license processing or a penalty processing
  'and warns him about saving (deletes temporary files)...making changes while a
  'customer is involved in one of the mentioned processes would make
  'the data that's posted inaccurate
  
  If GCustNum = 0 Then
    cmdTransHist.Enabled = False
    cmdDelete.Enabled = False
  Else
    cmdTransHist.Enabled = True
    cmdDelete.Enabled = True
  End If
  
  If GCustNum > 0 And Exist("artmppst.dat") Then
    If EmpInLicProcess(CStr(GCustNum)) = True Then
      frmBLMessageBoxJr.Label1.Caption = "This customer is currently being processed for a business license renewal. Saving data for this customer automatically deletes ALL temporary business license files. You will be required to run the license register again."
      frmBLMessageBoxJr.Label1.Top = 600
      frmBLMessageBoxJr.Show vbModal
      MainLog ("User warned that this customer, " + CStr(GCustNum) + ", is in a temporary license file and that saving data will delete the 'artmppst.dat' file.")
    End If
  End If

  If GCustNum > 0 And Exist("artmppen.dat") Then
    If EmpInPenProcess(CStr(GCustNum)) = True Then
      frmBLMessageBoxJr.Label1.Caption = "This customer is currently being processed for a penalty fee. Saving data for this customer automatically deletes ALL temporary penalty fee files. You will be required to run penalty fees again."
      frmBLMessageBoxJr.Label1.Top = 600
      frmBLMessageBoxJr.Show vbModal
      MainLog ("User warned that this customer, " + CStr(GCustNum) + ", is in a temporary penalty fee file and that saving data will delete the 'artmppen.dat' file.")
    End If
  End If
  
  OpenTownFile TownHandle
  NumOfTownRecs = LOF(TownHandle) / Len(TownRec)
  If NumOfTownRecs > 0 Then
    Get TownHandle, 1, TownRec
  End If
  Close TownHandle

  PermNum = False

  If QPTrim$(TownRec.LicNumPermYN) = "Yes" Then
    PermNum = True
  End If

  lblSave.Visible = False

  If Not Exist("arcatcodeidx.dat") Then 'no file there
    frmBLMessageBoxJr.Label1.Caption = "No category codes have been saved. Please save data for at least one category code. Loading aborted."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If

  OpenCatCodeIdxFile CodeIdxHandle
  CodeIdxRecNum = LOF(CodeIdxHandle) \ Len(CodeIdxRec)
  If CodeIdxRecNum = 0 Then 'file is there but there is nothing in it
    frmBLMessageBoxJr.Label1.Caption = "Category codes have not been indexed. Please call Southern Software at 1-800-842-8190. Loading aborted."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If

  ReDim CodeIdx(1 To CodeIdxRecNum) As Integer
  For x = 1 To CodeIdxRecNum
    Get CodeIdxHandle, x, CodeIdxRec
    CodeIdx(x) = CodeIdxRec.CatCodeRec 'load array with record pointers
  Next x
  Close CodeIdxHandle

  If Not Exist("ARCODE.DAT") Then
    frmBLMessageBoxJr.Label1.Caption = "The file 'arcode.dat' could not be found. Please call Southern Software at 1-800-842-8190. Loading aborted."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If

  OpenCatCodeFile CHandle
  CatCodeCnt = LOF(CHandle) / Len(CodeRec)

  If CatCodeCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No category codes have been saved. Please save data for at least one category code. Loading aborted."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If

  For x = 0 To 4
    fptxtRev(x).BackColor = &H80000005
  Next x

  If fptxtCode(0).ListCount = 0 Then 'keeps the list from rebuilding
  'when reloading
    For x = 0 To 4
      For y = 1 To CodeIdxRecNum
        Get CHandle, CodeIdx(y), CodeRec
        If Len(QPTrim(CodeRec.CatCode)) = 0 Then GoTo BadCode
        fptxtCode(x).InsertRow = QPTrim$(CodeRec.CatCode) & Chr(9) & QPTrim$(CodeRec.CODEDESC) & Chr(9) & QPTrim$(CodeRec.CodeType)
BadCode:
      Next y
    Next x
  End If
  
  Close CHandle

  FirstTime = True
  One = 1
  DHandle = FreeFile
  Open "customeredit.dat" For Output As DHandle Len = 2
  Print #DHandle, One
  Close DHandle
  AddFlag = False

  GoSub LoadGLAcctInfo

  OpenCustFile CustFile
  NumOfCustRecs = LOF(CustFile) \ Len(CustRec)
  
  TempGCustNum = GCustNum
  If fpcmbInactiveAcct.ListCount = 0 Then
    fpcmbInactiveAcct.AddItem "Active"
    fpcmbInactiveAcct.AddItem "Inactive"
  End If
  If fpcmbIOType.ListCount = 0 Then
    fpcmbIOType.AddItem "Inside"
    fpcmbIOType.AddItem "Outside"
  End If
  If fpcmbPrintNext.ListCount = 0 Then
    fpcmbPrintNext.AddItem "Yes"
    fpcmbPrintNext.AddItem "No"
  End If

  If GCustNum > 0 Then
    Get CustFile, GCustNum, CustRec
    MainLog ("Customer number " + CStr(GCustNum) + "/" + QPTrim$(CustRec.BillName) + " opened for editing on the 'Customer Maintenance' screen.")
    fptxtBusName.Text = QPTrim$(CustRec.CustName)
    TempCustName$ = QPTrim$(CustRec.CustName)
    
    
    fptxtCustNum.Text = QPTrim$(CustRec.CustNumb)
    TempCustNumb$ = QPTrim$(CustRec.CustNumb)
    fptxtAddress1.Text = QPTrim$(CustRec.ADDRESS1)
    TempAddress1$ = QPTrim$(CustRec.ADDRESS1)
    fptxtAddress2.Text = QPTrim$(CustRec.ADDRESS2)
    TempAddress2$ = QPTrim$(CustRec.ADDRESS2)
    fptxtCity.Text = QPTrim$(CustRec.City)
    TempCity$ = QPTrim$(CustRec.City)
    fptxtState.Text = QPTrim$(CustRec.State)
    TempState$ = QPTrim$(CustRec.State)
    If QPTrim$(CustRec.ZipCode) = "-" Then
      fptxtZip.Text = ""
    Else
      fptxtZip.Text = QPTrim$(CustRec.ZipCode)
    End If
    TempZip$ = QPTrim$(CustRec.ZipCode)
    fptxtBillingName.Text = QPTrim$(CustRec.BillName)
    TempBillName = QPTrim$(CustRec.BillName)
    fptxtSearchName.Text = QPTrim$(CustRec.SortName)
    TempSearchName = QPTrim$(CustRec.SortName)

    fptxtContact.Text = QPTrim$(CustRec.Contact)
    TempContact$ = QPTrim$(CustRec.Contact)
    fptxtServAdd.Text = QPTrim$(CustRec.ServAdd)
    TempServAdd = QPTrim$(CustRec.ServAdd)
    fptxtSSNFID.Text = QPTrim$(CustRec.SSNFID)
    TempSSNFID = QPTrim$(CustRec.SSNFID)
    TempBillCat1$ = QPTrim$(CustRec.BILLCAT1)
    CreditAmt(0) = CustRec.FeeLicBal1
    TempDESC1$ = QPTrim$(CustRec.DESC1)
    'Check4Multi is needed because we don't know what type of
    'fee method is used by the categories saved for the customer
    'This information is used to determine the kind of colors
    'and fields will be created for the Rev\Multi\Flat fields
    RevType = Check4Multi(QPTrim$(CustRec.BILLCAT1))
    TempType1 = RevType
    'the following code is required to properly load the display field
    'of the drop down box so that its data can be retrieved and saved...not doing it
    'like this prevents any recognition when an attempt is made to read the any data
    'in the field...data can only be recognized if the display is populated
    'from data in the list    fptxtCode(0).SearchText = QPTrim$(CustRec.BILLCAT1)
    fptxtCode(0).Text = ""
    fptxtCode(0).SearchText = QPTrim$(CustRec.BILLCAT1)
    fptxtCode(0).Action = 0
    If fptxtCode(0).SearchIndex <> -1 Then
      DoEvents
      fptxtCode(0).ListIndex = fptxtCode(0).SearchIndex
    End If
    
    If CustRec.FeeLicBal1 <> 0 Then
      fptxtCode(0).Enabled = False
      fptxtCode(0).Tag = "This drop down box contains all the category codes currently saved. You can assign this business one of the codes from this list. The program will not permit a business to have the same category code more than once. Each line contains the category code, category description and the fee assessment method for each category. A category cannot be changed or deleted if there is an outstanding balance for that category. This category has an outstanding balance of " + QPTrim$(Using("$#,###,##0.00", CustRec.FeeLicBal1)) + "."
    Else
      fptxtCode(0).Enabled = True
      fptxtCode(0).Tag = "This drop down box contains all the category codes currently saved. You can assign this business one of the codes from this list. The program will not permit a business to have the same category code more than once. Each line contains the category code, category description and the fee assessment method for each category. A category cannot be changed or deleted if there is an outstanding balance for that category."
    End If
    
    If RevType = "M" Then
      fptxtRev(0).BackColor = &H80FFFF
      fptxtRev(0) = CustRec.REV1
'      fptxtRev(0) = Using$("###0", CustRec.REV1)
      TempRev1 = CustRec.REV1
    ElseIf RevType = "S" Then
      fptxtRev(0).BackColor = &HFFFFFF
      fptxtRev(0) = Using("$###,###,##0.00", CustRec.REV1)
      TempRev1 = CustRec.REV1
    ElseIf RevType = "F" Then
      fptxtRev(0).BackColor = &HC0C0FF
      fptxtRev(0) = CustRec.REV1
      TempRev1 = CustRec.REV1
'      fptxtRev(0) = Using$("###0", CustRec.REV1)
    Else
      fptxtRev(0).Text = ""
      TempRev1 = -1
    End If
'    TempRev1 = CustRec.REV1
      
NextCat2:
    TempBillCat2$ = QPTrim$(CustRec.BILLCAT2)
    CreditAmt(1) = CustRec.FeeLicBal2
    TempDESC2$ = QPTrim$(CustRec.DESC2)
    RevType = Check4Multi(QPTrim$(CustRec.BILLCAT2))
    TempType2 = RevType
    fptxtCode(1).Text = ""
    fptxtCode(1).SearchText = QPTrim$(CustRec.BILLCAT2)
    fptxtCode(1).Action = 0
    If fptxtCode(1).SearchIndex <> -1 Then
      DoEvents
      fptxtCode(1).ListIndex = fptxtCode(1).SearchIndex
    End If
    
    If CustRec.FeeLicBal2 <> 0 Then
      fptxtCode(1).Enabled = False
      fptxtCode(1).Tag = "This drop down box contains all the category codes currently saved. You can assign this business one of the codes from this list. The program will not permit a business to have the same category code more than once. Each line contains the category code, category description and the fee assessment method for each category. A category cannot be changed or deleted if there is an outstanding balance for that category. This category has an outstanding balance of " + QPTrim$(Using("$#,###,##0.00", CustRec.FeeLicBal2)) + "."
    Else
      fptxtCode(1).Enabled = True
      fptxtCode(1).Tag = "This drop down box contains all the category codes currently saved. You can assign this business one of the codes from this list. The program will not permit a business to have the same category code more than once. Each line contains the category code, category description and the fee assessment method for each category. A category cannot be changed or deleted if there is an outstanding balance for that category."
    End If
    
    If RevType = "M" Then
      fptxtRev(1).BackColor = &H80FFFF
      fptxtRev(1) = CustRec.REV2
'      fptxtRev(1) = Using$("###0", CustRec.REV2)
      TempRev2 = CustRec.REV2
    ElseIf RevType = "S" Then
      fptxtRev(1).BackColor = &HFFFFFF
      fptxtRev(1) = Using("$###,###,##0.00", CustRec.REV2)
      TempRev2 = CustRec.REV2
    ElseIf RevType = "F" Then
      fptxtRev(1).BackColor = &HC0C0FF
      fptxtRev(1) = CustRec.REV2
'      fptxtRev(1) = Using$("###0", CustRec.REV2)
      TempRev2 = CustRec.REV2
    Else
      fptxtRev(1).BackColor = &HFFFFFF
      fptxtRev(1).Text = ""
      TempRev2 = -1
    End If
'    TempRev2 = CustRec.REV2

NextCat3:
    TempBillCat3$ = QPTrim$(CustRec.BILLCAT3)
    CreditAmt(2) = CustRec.FeeLicBal3
    TempDESC3$ = QPTrim$(CustRec.DESC3)
    RevType = Check4Multi(QPTrim$(CustRec.BILLCAT3))
    TempType3 = RevType
    fptxtCode(2).Text = ""
    fptxtCode(2).SearchText = QPTrim$(CustRec.BILLCAT3)
    fptxtCode(2).Action = 0
    If fptxtCode(2).SearchIndex <> -1 Then
      DoEvents
      fptxtCode(2).ListIndex = fptxtCode(2).SearchIndex
    End If
    
    If CustRec.FeeLicBal3 <> 0 Then
      fptxtCode(2).Enabled = False
      fptxtCode(2).Tag = "This drop down box contains all the category codes currently saved. You can assign this business one of the codes from this list. The program will not permit a business to have the same category code more than once. Each line contains the category code, category description and the fee assessment method for each category. A category cannot be changed or deleted if there is an outstanding balance for that category. This category has an outstanding balance of " + QPTrim$(Using("$#,###,##0.00", CustRec.FeeLicBal3)) + "."
    Else
      fptxtCode(2).Enabled = True
      fptxtCode(2).Tag = "This drop down box contains all the category codes currently saved. You can assign this business one of the codes from this list. The program will not permit a business to have the same category code more than once. Each line contains the category code, category description and the fee assessment method for each category. A category cannot be changed or deleted if there is an outstanding balance for that category."
    End If
    
    If RevType = "M" Then
      fptxtRev(2).BackColor = &H80FFFF
      fptxtRev(2) = CustRec.REV3
'      fptxtRev(2) = Using$("###0", CustRec.REV3)
      TempRev3 = CustRec.REV3
    ElseIf RevType = "S" Then
      fptxtRev(2).BackColor = &HFFFFFF
      fptxtRev(2) = Using("$###,###,##0.00", CustRec.REV3)
      TempRev3 = CustRec.REV3
    ElseIf RevType = "F" Then
      fptxtRev(2).BackColor = &HC0C0FF
      fptxtRev(2) = CustRec.REV3
'      fptxtRev(2) = Using$("###0", CustRec.REV3)
      TempRev3 = CustRec.REV3
    Else
      fptxtRev(2).BackColor = &HFFFFFF
      fptxtRev(2).Text = ""
      TempRev3 = -1
    End If
'    TempRev3 = CustRec.REV3

NextCat4:
    TempBillCat4$ = QPTrim$(CustRec.BILLCAT4)
    CreditAmt(3) = CustRec.FeeLicBal4
    TempDESC4$ = QPTrim$(CustRec.DESC4)
    RevType = Check4Multi(QPTrim$(CustRec.BILLCAT4))
    TempType4 = RevType
    fptxtCode(3).Text = ""
    fptxtCode(3).SearchText = QPTrim$(CustRec.BILLCAT4)
    fptxtCode(3).Action = 0
    If fptxtCode(3).SearchIndex <> -1 Then
      DoEvents
      fptxtCode(3).ListIndex = fptxtCode(3).SearchIndex
    End If
    
    If CustRec.FeeLicBal4 <> 0 Then
      fptxtCode(3).Enabled = False
      fptxtCode(3).Tag = "This drop down box contains all the category codes currently saved. You can assign this business one of the codes from this list. The program will not permit a business to have the same category code more than once. Each line contains the category code, category description and the fee assessment method for each category. A category cannot be changed or deleted if there is an outstanding balance for that category. This category has an outstanding balance of " + QPTrim$(Using("$#,###,##0.00", CustRec.FeeLicBal4)) + "."
    Else
      fptxtCode(3).Enabled = True
      fptxtCode(3).Tag = "This drop down box contains all the category codes currently saved. You can assign this business one of the codes from this list. The program will not permit a business to have the same category code more than once. Each line contains the category code, category description and the fee assessment method for each category. A category cannot be changed or deleted if there is an outstanding balance for that category."
    End If
    
    If RevType = "M" Then
      fptxtRev(3).BackColor = &H80FFFF
      fptxtRev(3) = CustRec.REV4
'      fptxtRev(3) = Using$("###0", CustRec.REV4)
      TempRev4 = CustRec.REV4
    ElseIf RevType = "S" Then
      fptxtRev(3).BackColor = &HFFFFFF
      fptxtRev(3) = Using("$###,###,##0.00", CustRec.REV4)
      TempRev4 = CustRec.REV4
    ElseIf RevType = "F" Then
      fptxtRev(3).BackColor = &HC0C0FF
      fptxtRev(3) = CustRec.REV4
'      fptxtRev(3) = Using$("###0", CustRec.REV4)
      TempRev4 = CustRec.REV4
    Else
      fptxtRev(3).BackColor = &HFFFFFF
      fptxtRev(3).Text = ""
      TempRev4 = -1
    End If
'    TempRev4 = CustRec.REV4

NextCat5:
    TempBillCat5$ = QPTrim$(CustRec.BILLCAT5)
    CreditAmt(4) = CustRec.FeeLicBal5
    TempDESC5$ = QPTrim$(CustRec.DESC5)
    RevType = Check4Multi(QPTrim$(CustRec.BILLCAT5))
    TempType5 = RevType
    fptxtCode(4).Text = ""
    fptxtCode(4).SearchText = QPTrim$(CustRec.BILLCAT5)
    fptxtCode(4).Action = 0
    If fptxtCode(4).SearchIndex <> -1 Then
      DoEvents
      fptxtCode(4).ListIndex = fptxtCode(4).SearchIndex
    End If
    
    If CustRec.FeeLicBal5 <> 0 Then
      fptxtCode(4).Enabled = False
      fptxtCode(4).Tag = "This drop down box contains all the category codes currently saved. You can assign this business one of the codes from this list. The program will not permit a business to have the same category code more than once. Each line contains the category code, category description and the fee assessment method for each category. A category cannot be changed or deleted if there is an outstanding balance for that category. This category has an outstanding balance of " + QPTrim$(Using("$#,###,##0.00", CustRec.FeeLicBal5)) + "."
    Else
      fptxtCode(4).Enabled = True
      fptxtCode(4).Tag = "This drop down box contains all the category codes currently saved. You can assign this business one of the codes from this list. The program will not permit a business to have the same category code more than once. Each line contains the category code, category description and the fee assessment method for each category. A category cannot be changed or deleted if there is an outstanding balance for that category."
    End If
    
    If RevType = "M" Then
      fptxtRev(4).BackColor = &H80FFFF
      fptxtRev(4) = CustRec.REV5
'      fptxtRev(4) = Using$("###0", CustRec.REV5)
      TempRev5 = CustRec.REV5
    ElseIf RevType = "S" Then
      fptxtRev(4).BackColor = &HFFFFFF
      fptxtRev(4) = Using("$###,###,##0.00", CustRec.REV5)
      TempRev5 = CustRec.REV5
    ElseIf RevType = "F" Then
      fptxtRev(4).BackColor = &HC0C0FF
      fptxtRev(4) = CustRec.REV5
'      fptxtRev(4) = Using$("###0", CustRec.REV5)
      TempRev5 = CustRec.REV5
    Else
      fptxtRev(4).BackColor = &HFFFFFF
      fptxtRev(4).Text = ""
      TempRev5 = -1
    End If
'    TempRev5 = CustRec.REV5
    
    'check to make sure all categories loaded OK and if not
    'then tell the user there is a problem
    For x = 0 To 4
      If QPTrim$(fptxtCode(x).SearchText) <> "" Then
        If fptxtCode(x).SearchIndex = -1 Then
           frmBLMessageBoxJr.Label1.Caption = "The category code " + QPTrim$(fptxtCode(x).SearchText) + " saved for this customer on 'Billing Categories' row " + CStr(x + 1) + " could not be found. Please resolve this issue before continuing."
           frmBLMessageBoxJr.Label1.Top = 700
           frmBLMessageBoxJr.Show vbModal
           MainLog ("User warned that the category code saved for this customer " + QPTrim$(fptxtCode(x).SearchText) + " on 'Billing Categories' row " + CStr(x + 1) + " could not be found and they needed to resolve this issue before continuing.")
        End If
      End If
    Next x
 
EndCat:
    If QPTrim$(CustRec.CustLocation) = "I" Then
      fpcmbIOType.Text = "Inside"
    ElseIf QPTrim$(CustRec.CustLocation) = "O" Then
      fpcmbIOType.Text = "Outside"
    Else
      fpcmbIOType.Text = ""
    End If
    TempLocation$ = QPTrim$(fpcmbIOType.Text)

    fpMaskPhone = TrimPhone(CustRec.WPHONE)
    TempWPHONE$ = CustRec.WPHONE
    fptxtLicNum.Text = QPTrim$(CustRec.LICENSE)
    TempLICENSE$ = QPTrim$(CustRec.LICENSE)
    fptxtValidThru.Text = MakeRegDate(CustRec.VALID)
    TempVALID = CustRec.VALID
    If QPTrim$(CustRec.IssueLicense) <> "Y" And QPTrim$(CustRec.IssueLicense) <> "N" Then
      fpcmbPrintNext.Text = "No"
    Else
      If QPTrim$(CustRec.IssueLicense) = "N" Then
        fpcmbPrintNext.Text = "No"
      Else
        fpcmbPrintNext.Text = "Yes"
      End If
    End If

    TempIssueLicense$ = QPTrim$(CustRec.IssueLicense)

    If QPTrim$(CustRec.Inactive) = "Y" Then
      fpcmbInactiveAcct.Text = "Inactive"
    Else
      fpcmbInactiveAcct.Text = "Active"
    End If

    TempInactive$ = QPTrim$(CustRec.Inactive)
    If CustRec.Prorate = 0 Then
      fptxtProrate.Text = "100.00%"
    Else
      fptxtProrate.Text = CStr(CustRec.Prorate) + "%"
    End If
    TempProrate = CustRec.Prorate
  Else 'zero out
    'The Temp fields are used during the LogSave routine
    TempBillName$ = "New Addition"
    TempSearchName$ = "New Addition"
    TempCustName$ = "New Addition"
    TempCustNumb$ = "New Addition"
    TempAddress1$ = "New Addition"
    TempAddress2$ = "New Addition"
    TempCity$ = "New Addition"
    TempState$ = "New Addition"
    TempZip$ = "New Addition"
    TempContact$ = "New Addition"
    AddFlag = True 'tells program that this is a new addition
    'and the user might want to keep adding categories without
    'returning to the main menu everytime a save is made
    fptxtBusName.Text = ""

    'customer number is automatically assigned
    fptxtCustNum.Text = NumOfCustRecs + 1

    If NumOfTownRecs > 0 Then
      fptxtCity.Text = QPTrim$(TownRec.City)
      fptxtState.Text = QPTrim$(TownRec.State)
      fptxtZip.Text = Mid(TownRec.ZipCode, 1, 5)
    Else
      fptxtCity.Text = ""
      fptxtState.Text = ""
      fptxtZip.Text = ""
    End If
    fptxtAddress1.Text = ""
    fptxtAddress2.Text = ""
    fptxtBillingName.Text = ""
    fptxtSearchName.Text = ""
    fptxtContact.Text = ""
    fptxtServAdd.Text = ""
    fptxtCode(0).Text = ""
    fptxtRev(0) = ""
    fptxtCode(1).Text = ""
    fptxtRev(1).Text = ""
    fptxtCode(2).Text = ""
    fptxtRev(2) = ""
    fptxtCode(3).Text = ""
    fptxtRev(3) = ""
    fptxtCode(4).Text = ""
    fptxtRev(4) = ""
    fpcmbIOType.Text = "Inside"
    fpMaskPhone.Text = ""
    'License number is automatically assigned
    fptxtLicNum.Text = FirstLicenseNum + 1
    fptxtValidThru.Text = Date$
    fpcmbPrintNext.Text = "No"
    fpcmbInactiveAcct.Text = "Active"
    fptxtProrate.Text = "100%"
    fptxtSSNFID.Text = ""
    fptxtServAdd.Text = ""
  End If
  
  Close

  Exit Sub

LoadGLAcctInfo:
  GetAcctStruct GLFundLen, GLAcctLen, GLDetLen
Return

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustEdit", "LoadMe", Erl)
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

Private Sub fpcmbInactiveAcct_Change()
  If QPTrim$(fpcmbInactiveAcct.Text) = "" Then
    fpcmbInactiveAcct.Text = "Active"
  End If
End Sub
Private Sub fpcmbIOType_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  If KeyCode = vbKeySpace Then
    fpcmbIOType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbIOType.ListIndex = -1
  End If
  If fpcmbIOType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbInactiveAcct_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  If KeyCode = vbKeySpace Then
    fpcmbInactiveAcct.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbInactiveAcct.ListIndex = -1
  End If
  If fpcmbInactiveAcct.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      If fptxtProrate.Enabled = True Then
        fptxtProrate.SetFocus
      End If
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        If fpcmbInactiveAcct.Enabled = True Then
          fpcmbPrintNext.SetFocus
        End If
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbPrintNext_Change()
  If QPTrim$(fpcmbPrintNext.Text) = "" Then
    fpcmbPrintNext.Text = "No"
  End If
End Sub

Private Sub fpcmbPrintNext_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  If KeyCode = vbKeyY Then
    fpcmbPrintNext.Text = "Yes"
  ElseIf KeyCode = vbKeyN Then
    fpcmbPrintNext.Text = "No"
  End If
  
  If KeyCode = vbKeySpace Then
    fpcmbPrintNext.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPrintNext.ListIndex = -1
  End If
  If fpcmbPrintNext.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbInactiveAcct.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fptxtValidThru.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub fpcmbPrintNext_LostFocus()
  If QPTrim$(fpcmbPrintNext.Text) = "" Then
    fpcmbPrintNext.Text = "N"
  End If
End Sub

Private Sub fptxtBillingName_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    If fptxtBusName.Enabled = True Then
      fptxtBusName.SetFocus
    End If
  ElseIf KeyCode = vbKeyUp Then
    If fptxtProrate.Enabled = True Then
      fptxtProrate.SetFocus
    End If
  End If
End Sub

Private Sub fptxtBillingName_LostFocus()
  
  If FirstTime = True And GCustNum = 0 Then
    If Len(QPTrim$(fptxtBusName.Text)) = 0 Then
      If Len(QPTrim$(fptxtBillingName.Text)) > 0 Then
        fptxtBusName.Text = QPTrim$(fptxtBillingName.Text)
        FirstTime = False
      End If
    End If
  End If
End Sub

Private Sub fptxtBusName_LostFocus()
  If Len(QPTrim$(fptxtBusName.Text)) > 0 Then
    FirstTime = False
  End If
End Sub

Public Function Check4Multi(CatCode$) As String
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeHandle As Integer
  Dim NumOfCodeRecs As Integer
  Dim x As Integer
  
  On Error GoTo ERRORSTUFF
  
  'assign the function the "N' for NA value
  Check4Multi = "N"
  If QPTrim$(CatCode) = "" Then Exit Function
  OpenCatCodeFile CodeHandle
  
  NumOfCodeRecs = LOF(CodeHandle) / Len(CodeRec)
  If NumOfCodeRecs = 0 Then
    Close
    Exit Function
  End If
  
  For x = 1 To NumOfCodeRecs
    Get CodeHandle, x, CodeRec
    'compare the category in the catcode field until
    'is is found in the category list and then retrieve
    'the fee method
    If QPTrim$(CodeRec.CatCode) = QPTrim$(CatCode) Then
      If CodeRec.CodeType = "M" Then
        Check4Multi = "M"
        Close CodeHandle
        Exit Function
      ElseIf CodeRec.CodeType = "S" Then
        Check4Multi = "S"
        Close CodeHandle
        Exit Function
      ElseIf CodeRec.CodeType = "F" Then
        Check4Multi = "F"
        Close CodeHandle
        Exit Function
      End If
    End If
  Next x
  Close CodeHandle
  
  If NFlag = False Then
    If Check4Multi = "N" Then
      NFlag = True
      frmBLMessageBoxJr.Label1.Caption = "PROBLEM: This customer has a category code with no category type. Please examine the category codes for this customer. If an 'N' is indicated as the type then that category needs to be edited on the 'Category Edit' screen."
      frmBLMessageBoxJr.Label1.Top = 600
      frmBLMessageBoxJr.Show vbModal
    End If
  End If
    
  Exit Function
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustEdit", "Check4Multi", Erl)
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

Private Sub fptxtCode_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  
'  Call GetType(Index)
  If KeyCode = vbKeySpace Then
    fptxtCode(Index).ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fptxtCode(Index).ListIndex = -1
  End If
  
  If fptxtCode(Index).ListDown <> True Then
    If KeyCode = vbKeyDown Then
      If fptxtRev(Index).Enabled = True Then
        fptxtRev(Index).SetFocus
      End If
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        If Index = 0 Then
          If fptxtServAdd.Enabled = True Then
            fptxtServAdd.SetFocus
          End If
          KeyCode = 0
        Else
          If fptxtRev(Index - 1).Enabled = True Then
            fptxtRev(Index - 1).SetFocus
          End If
          KeyCode = 0
        End If
      End If
    End If
  End If

End Sub

Private Sub fptxtCode_LostFocus(Index As Integer)
  If QPTrim$(fptxtCode(Index).Text) <> "" And QPTrim(fptxtRev(Index).Text) = "" Then
'    fptxtCode(Index).SetFocus 'removed on 6/16/04
    Call fptxtCode_KeyDown(Index, vbKeyDown, 0)
  End If
  If QPTrim$(fptxtRev(Index).Text) <> "" Then
    If QPTrim$(fptxtCode(Index).Text) = "" Then
      fptxtRev(Index).Text = ""
      fptxtRev(Index).BackColor = &HFFFFFF
    End If
  End If
End Sub

Private Function Check4ValidCodeNum(CatCode$, IndexNum As Integer) As Boolean
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeHandle As Integer
  Dim NumOfCodeRecs As Integer
  Dim x As Integer

  On Error GoTo ERRORSTUFF

  'The function is false until proven otherwise
  Check4ValidCodeNum = False
  OpenCatCodeFile CodeHandle

  NumOfCodeRecs = LOF(CodeHandle) / Len(CodeRec)
  If NumOfCodeRecs = 0 Then
    Close
    Exit Function
  End If

  For x = 1 To NumOfCodeRecs
    Get CodeHandle, x, CodeRec
    If QPTrim$(CodeRec.CatCode) = QPTrim$(CatCode$) Then
      Check4ValidCodeNum = True
      fptxtCode(IndexNum).Col = 1
      fptxtCode(IndexNum).ColText = QPTrim$(CatCode$)
      If fptxtRev(IndexNum).Enabled = True Then
        fptxtRev(IndexNum).SetFocus
      End If
      Exit For
    End If
  Next x

  Close CodeHandle

  If Check4ValidCodeNum = False Then
    frmBLMessageBoxJrWOpts.Label1.Caption = "The Code you entered is not valid. Would you like to add this as a new code number?"
    frmBLMessageBoxJrWOpts.Label1.Top = 800
    frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Add"
    frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
    frmBLMessageBoxJrWOpts.Show vbModal
    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
      Unload frmBLMessageBoxJrWOpts
      frmBLCatEdit.Show
      DoEvents
      Unload frmBLCustEdit
      Exit Function
    Else
      Unload frmBLMessageBoxJrWOpts
      fptxtCode(IndexNum).Col = 0
      fptxtCode(IndexNum).ColText = ""
      fptxtCode(IndexNum).Col = 1
      fptxtCode(IndexNum).ColText = ""
      fptxtRev(IndexNum).Text = Using$("$###,###,##0.00", 0)
      If fptxtCode(IndexNum).Enabled = True Then
        fptxtCode(IndexNum).SetFocus
      End If
    End If
  End If

  Exit Function

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustEdit", "Check4ValidCodeNum", Erl)
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

Private Sub fptxtLicNum_KeyPress(KeyAscii As Integer)
  Dim ThisLic As Double
  
  If KeyAscii = 61 Then
    If IsNumeric(fptxtLicNum.Text) Then
      ThisLic = CDbl(fptxtLicNum.Text)
      ThisLic = ThisLic + 1
      fptxtLicNum.Text = CStr(ThisLic)
      KeyAscii = 0
    End If
  End If

End Sub

Private Sub fptxtProrate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    If fptxtBillingName.Enabled = True Then
      fptxtBillingName.SetFocus
    End If
  ElseIf KeyCode = vbKeyUp Then
    If fpcmbInactiveAcct.Enabled Then
      fpcmbInactiveAcct.SetFocus
    End If
  End If
End Sub

Private Sub fptxtProrate_LostFocus()
  Dim Prorate$
  Dim ThisProrate As Double
  
  On Error Resume Next
  If QPTrim$(fptxtProrate.Text) = "" Then
    fptxtProrate.Text = "100%"
  ElseIf InStr(fptxtProrate.Text, ".") Then
    ThisProrate = CDbl(ReplaceString(fptxtProrate.Text, "%", ""))
    fptxtProrate.Text = CStr(ThisProrate)
    frmBLMessageBoxJr.Label1.Caption = "Please enter whole percentage numbers only."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtProrate.Text = "100%"
    If fptxtProrate.Enabled = True Then
      fptxtProrate.SetFocus
    End If
  End If
  
  Prorate = ReplaceString(fptxtProrate.Text, "%", "")
  fptxtProrate.Text = Prorate + "%"
End Sub

Private Sub fptxtRev_GotFocus(Index As Integer)
  If QPTrim$(fptxtCode(Index).Text) = "" Then Exit Sub

  fptxtCode(Index).Col = 2

  If fptxtCode(Index).ColText = "F" Then
    fptxtRev(Index).BackColor = &HC0C0FF
    If InStr(fptxtRev(Index).Text, "$") Then
      fptxtRev(Index).Text = ReplaceString$(fptxtRev(Index).Text, "$", "")
      If InStr(fptxtRev(Index).Text, ",") Then
        fptxtRev(Index).Text = ReplaceString(fptxtRev(Index).Text, ",", "")
      End If
      fptxtRev(Index).Text = Val(fptxtRev(Index).Text)
    End If
  ElseIf fptxtCode(Index).ColText = "M" Then
    fptxtRev(Index).BackColor = &H80FFFF
    If InStr(fptxtRev(Index).Text, "$") Then
      fptxtRev(Index).Text = ReplaceString$(fptxtRev(Index).Text, "$", "")
      If InStr(fptxtRev(Index).Text, ",") Then
        fptxtRev(Index).Text = ReplaceString(fptxtRev(Index).Text, ",", "")
      End If
      fptxtRev(Index).Text = Val(fptxtRev(Index).Text)
    End If
  Else
    fptxtRev(Index).BackColor = &HFFFFFF
  End If
  
End Sub

Private Sub fptxtRev_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    If Index = 4 Then
      fpcmbIOType.SetFocus
    Else
      If fptxtCode(Index + 1).Enabled = True Then
        fptxtCode(Index + 1).SetFocus
      End If
    End If
  ElseIf KeyCode = vbKeyUp Then
    If fptxtCode(Index).Enabled = True Then
      fptxtCode(Index).SetFocus
    End If
  End If
    
End Sub

Private Function Check4ValidCustNum(CustNumb$) As Boolean
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim x As Integer
  Dim NumOfCustRecs As Integer
  
  On Error Resume Next
  'function is true until proven otherwise
  Check4ValidCustNum = True
  OpenCustFile CustHandle
  NumOfCustRecs = LOF(CustHandle) / Len(CustRec)
  For x = 1 To NumOfCustRecs
    Get CustHandle, x, CustRec
    'if this customer number is the same as the one the screen opened
    'with then we're OK...go ahead and exit
    If GCustNum > 0 And GCustNum = QPTrim$(CustNumb) Then Exit For
    'otherwise start trying to find a duplicate in the customer number list
    If QPTrim$(CustRec.CustNumb) = QPTrim$(CustNumb) Then
      'found a copy so make the function false and leave
      Check4ValidCustNum = False
      Exit For
    End If
  Next x

  If Check4ValidCustNum = False Then
    frmBLMessageBoxJr.Label1.Caption = "The customer number you entered is already in use. Please select another number."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
    If fptxtCustNum.Enabled = True Then
      fptxtCustNum.SetFocus
    End If
  End If

  Close CustHandle

End Function

Private Function Check4ValidLic(ThisLic$) As Boolean
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim x As Integer
  Dim NumOfCustRecs As Integer
  Dim BigNum As Double
  
  On Error GoTo ERRORSTUFF
  'the function is true unless proven otherwise
  Check4ValidLic = True
  OpenCustFile CustHandle
  NumOfCustRecs = LOF(CustHandle) / Len(CustRec)
  For x = 1 To NumOfCustRecs
    Get CustHandle, x, CustRec
    'don't compare the current license number with itself
    If GCustNum > 0 And GCustNum = x Then GoTo SkipIt
    If CustRec.Deleted = "Y" Or QPTrim$(CustRec.SortName) = "DELETED" Then GoTo SkipIt
    If QPTrim$(CustRec.LICENSE) = QPTrim$(ThisLic$) Then 'found a copy
      'exit after making function false
      Check4ValidLic = False
      Exit For
    End If
SkipIt:
  Next x

  If Check4ValidLic = False Then
    frmBLMessageBoxJr.Label1.Caption = "The customer license number you entered is already in use. Please select a different license number." 'Any number larger than " + CStr(BigNum) + " will be valid."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
    If fptxtLicNum.Enabled = True Then
      fptxtLicNum.SetFocus
    End If
  End If
  
  Close CustHandle
  
  Exit Function
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustEdit", "Check4ValidLic", Erl)
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

Private Sub ClearScreen()
  Dim x As Integer
  
  fptxtBillingName.Text = ""
  fpcmbInactiveAcct.Text = ""
  fpcmbIOType.Text = ""
  fpcmbPrintNext.Text = ""
  fpMaskPhone.Text = ""
  fptxtAddress1.Text = ""
  fptxtAddress2.Text = ""
  fptxtBusName.Text = ""
  fptxtCity.Text = ""
  For x = 0 To 4
    fptxtCode(x).Col = 0
    fptxtCode(x).ColText = 0
    fptxtCode(x).Col = 1
    fptxtCode(x).ColText = ""
    fptxtCode(x).Text = ""
    fptxtRev(x).Text = ""
  Next x
  fptxtContact.Text = ""
  fptxtCustNum.Text = ""
  fptxtLicNum.Text = ""
  fptxtProrate.Text = ""
  fptxtSearchName.Text = ""
  fptxtState.Text = ""
  fptxtValidThru.Text = ""
  fptxtZip.Text = ""

End Sub

Private Sub LogSaves(SaveHere As Integer)
  Dim DHandle As Integer
  Dim CustRec As ARCustRecType
  Dim NumOfCustRecs As Integer
   'this sub simply records anything saved in the arlog
  On Error GoTo ERRORSTUFF
  If SaveHere <= 0 Then Exit Sub
  OpenCustFile DHandle
  Get DHandle, SaveHere, CustRec
  Close DHandle
  
  If QPTrim$(TempCustName$) <> QPTrim$(CustRec.CustName) Then
    MainLog ("For " + QPTrim(CustRec.CustName) + ":  Customer name, " + QPTrim$(TempCustName$) + ", changed and saved as " + QPTrim$(CustRec.CustName) + ".")
  End If
  If QPTrim$(TempCustNumb$) <> QPTrim$(CustRec.CustNumb) Then
    MainLog ("For " + QPTrim(CustRec.CustName) + ":  Customer number, " + QPTrim$(TempCustNumb$) + ", changed and saved as " + QPTrim$(CustRec.CustNumb) + ".")
  End If
  If QPTrim$(TempAddress1$) <> QPTrim$(CustRec.ADDRESS1) Then
    MainLog ("For " + QPTrim(CustRec.CustName) + ":  Address 1, " + QPTrim$(TempAddress1$) + ", changed and saved as " + QPTrim$(CustRec.ADDRESS1) + ".")
  End If
  If QPTrim$(TempAddress2$) <> QPTrim$(CustRec.ADDRESS2) Then
    MainLog ("For " + QPTrim(CustRec.CustName) + ":  Address 2, " + QPTrim$(TempAddress2$) + ", changed and saved as " + QPTrim$(CustRec.ADDRESS2) + ".")
  End If
  If QPTrim$(TempCity$) <> QPTrim$(CustRec.City) Then
    MainLog ("For " + QPTrim(CustRec.CustName) + ":  City, " + QPTrim$(TempCity$) + ", changed and saved as " + QPTrim$(CustRec.City) + ".")
  End If
  If QPTrim$(TempState$) <> QPTrim$(CustRec.State) Then
    MainLog ("For " + QPTrim(CustRec.CustName) + ":  State, " + QPTrim$(TempState$) + ", changed and saved as " + QPTrim$(CustRec.State) + ".")
  End If
  If QPTrim$(TempZip$) <> QPTrim$(CustRec.ZipCode) Then
    MainLog ("For " + QPTrim(CustRec.CustName) + ":  Zip Code, " + QPTrim$(TempZip$) + ", changed and saved as " + QPTrim$(CustRec.ZipCode) + ".")
  End If
  If QPTrim$(TempSearchName$) <> QPTrim$(CustRec.SortName) Then
    MainLog ("For " + QPTrim(CustRec.CustName) + ":  Search Name, " + QPTrim$(TempSearchName$) + ", changed and saved as " + QPTrim$(CustRec.SortName) + ".")
  End If
  If QPTrim$(TempBillName$) <> QPTrim$(CustRec.BillName) Then
    MainLog ("For " + QPTrim(CustRec.CustName) + ":  Billing Name, " + QPTrim$(TempBillName$) + ", changed and saved as " + QPTrim$(CustRec.BillName) + ".")
  End If
  If QPTrim$(TempContact$) <> QPTrim$(CustRec.Contact) Then
    MainLog ("For " + QPTrim(CustRec.CustName) + ":  Contact Name, " + QPTrim$(TempContact$) + ", changed and saved as " + QPTrim$(CustRec.Contact) + ".")
  End If
  If QPTrim$(TempServAdd$) <> QPTrim$(CustRec.ServAdd) Then
    MainLog ("For " + QPTrim(CustRec.CustName) + ":  Service Address, " + QPTrim$(TempServAdd$) + ", changed and saved as " + QPTrim$(CustRec.ServAdd) + ".")
  End If
  If QPTrim$(TempSSNFID$) <> QPTrim$(CustRec.SSNFID) Then
    MainLog ("For " + QPTrim(CustRec.CustName) + ":  SSNFID, " + QPTrim$(TempSSNFID$) + ", changed and saved as " + QPTrim$(CustRec.SSNFID) + ".")
  End If
  If QPTrim$(TempBillCat1$) <> QPTrim$(CustRec.BILLCAT1) Then
    MainLog ("For " + QPTrim(CustRec.CustName) + ":  Billing Category 1, " + QPTrim$(TempBillCat1$) + ", changed and saved as " + QPTrim$(CustRec.BILLCAT1) + ".")
  End If
  If TempRev1 <> CustRec.REV1 Then
    MainLog ("For " + QPTrim(CustRec.CustName) + ":  Revenue 1, " + CStr(TempRev1) + ", changed and saved as " + CStr(CustRec.REV1) + ".")
  End If
  If QPTrim$(TempBillCat2$) <> QPTrim$(CustRec.BILLCAT2) Then
    MainLog ("For " + QPTrim(CustRec.CustName) + ":  Billing Category 2, " + QPTrim$(TempBillCat2$) + ", changed and saved as " + QPTrim$(CustRec.BILLCAT2) + ".")
  End If
  If TempRev2 <> CustRec.REV2 Then
    MainLog ("For " + QPTrim(CustRec.CustName) + ":  Revenue 2, " + CStr(TempRev2) + ", changed and saved as " + CStr(CustRec.REV2) + ".")
  End If
  If QPTrim$(TempBillCat3$) <> QPTrim$(CustRec.BILLCAT3) Then
    MainLog ("For " + QPTrim(CustRec.CustName) + ":  Billing Category 3, " + QPTrim$(TempBillCat3$) + ", changed and saved as " + QPTrim$(CustRec.BILLCAT3) + ".")
  End If
  If TempRev3 <> CustRec.REV3 Then
    MainLog ("For " + QPTrim(CustRec.CustName) + ":  Revenue 3, " + CStr(TempRev3) + ", changed and saved as " + CStr(CustRec.REV3) + ".")
  End If
  If QPTrim$(TempBillCat4$) <> QPTrim$(CustRec.BILLCAT4) Then
    MainLog ("For " + QPTrim(CustRec.CustName) + ":  Billing Category 4, " + QPTrim$(TempBillCat4$) + ", changed and saved as " + QPTrim$(CustRec.BILLCAT4) + ".")
  End If
  If TempRev4 <> CustRec.REV4 Then
    MainLog ("For " + QPTrim(CustRec.CustName) + ":  Revenue 4, " + CStr(TempRev4) + ", changed and saved as " + CStr(CustRec.REV4) + ".")
  End If
  If QPTrim$(TempBillCat5$) <> QPTrim$(CustRec.BILLCAT5) Then
    MainLog ("For " + QPTrim(CustRec.CustName) + ":  Billing Category 5, " + QPTrim$(TempBillCat5$) + ", changed and saved as " + QPTrim$(CustRec.BILLCAT5) + ".")
  End If
  If TempRev5 <> CustRec.REV5 Then
    MainLog ("For " + QPTrim(CustRec.CustName) + ":  Revenue 5, " + CStr(TempRev5) + ", changed and saved as " + CStr(CustRec.REV5) + ".")
  End If
  If QPTrim$(TempLocation) = "Inside" Then TempLocation = "I"
  If QPTrim$(TempLocation) = "Outside" Then TempLocation = "O"
  If QPTrim$(TempLocation$) <> QPTrim$(CustRec.CustLocation) Then
    MainLog ("For " + QPTrim(CustRec.CustName) + ":  Location, " + QPTrim$(TempLocation$) + ", changed and saved as " + QPTrim$(CustRec.CustLocation) + ".")
  End If
  If TrimPhone(TempWPHONE$) <> TrimPhone(CustRec.WPHONE) Then
    MainLog ("For " + QPTrim(CustRec.CustName) + ":  Phone number, " + QPTrim$(TempWPHONE$) + ", changed and saved as " + QPTrim$(CustRec.WPHONE) + ".")
  End If
  If TempVALID <> CustRec.VALID Then
    MainLog ("For " + QPTrim(CustRec.CustName) + ":  Valid To, " + MakeRegDate(TempVALID) + ", changed and saved as " + MakeRegDate(CustRec.VALID) + ".")
  End If
  If QPTrim$(TempLICENSE$) <> QPTrim$(CustRec.LICENSE) Then
    MainLog ("For " + QPTrim(CustRec.CustName) + ":  License, " + QPTrim$(TempLICENSE$) + ", changed and saved as " + QPTrim$(CustRec.LICENSE) + ".")
  End If
  If QPTrim$(TempIssueLicense$) <> QPTrim$(CustRec.IssueLicense) Then
    MainLog ("For " + QPTrim(CustRec.CustName) + ":  Issue License, " + QPTrim$(TempIssueLicense$) + ", changed and saved as " + QPTrim$(CustRec.IssueLicense) + ".")
  End If
  If QPTrim$(TempInactive$) <> QPTrim$(CustRec.Inactive) Then
    MainLog ("For " + QPTrim(CustRec.CustName) + ":  Inactive, " + QPTrim$(TempInactive$) + ", changed and saved as " + QPTrim$(CustRec.Inactive) + ".")
  End If
  If TempProrate <> CustRec.Prorate Then
    MainLog ("For " + QPTrim(CustRec.CustName) + ":  Prorate, " + CStr(TempProrate) + ", changed and saved as " + CStr(CustRec.Prorate) + ".")
  End If
  
  Exit Sub

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustEdit", "LogSaves", Erl)
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
Private Sub cmdList_Click()
  frmBLLicenseNumList.Show vbModal, Me
  If fptxtLicNum.Enabled = True Then
    fptxtLicNum.SetFocus
  End If
End Sub

Private Function AnyCreditBal() As Boolean 'NOT BEING USED
  Dim x As Integer
  Dim CreditBal(0 To 4) As Double
  Dim TotCreditBal As Double
  Dim RevType As String * 1
  
  On Error GoTo ERRORSTUFF
  'function is false unless proven otherwise
  AnyCreditBal = False
  For x = 0 To 4
    If CreditAmt(x) <> 0 Then 'CreditAmt is a local global assigned
    'in the LoadMe sub
      Select Case x 'this select looks to see if any categories have changed
      'and if so then checks to see if any balance exists for the old
      'category...if it does then the user has to resolve the balance
      'before any change can be made
      Case 0:
        fptxtCode(x).Col = 0
        If QPTrim$(fptxtCode(x).ColText) <> QPTrim$(TempBillCat1) Then
          fptxtCode(x).BackColor = &H80FFFF
          fptxtRev(x).BackColor = &H80FFFF
          AnyCreditBal = True
          If CreditAmt(x) < 0 Then
            frmBLMessageBoxJr.Label1.Caption = "The old category has a current balance of -$" + QPTrim$(Using("##,##0.00", Abs(CreditAmt(x)))) + ". Please resolve this outstanding balance before continuing."
            frmBLMessageBoxJr.Label1.Top = 700
            frmBLMessageBoxJr.Show vbModal
          Else
            frmBLMessageBoxJr.Label1.Caption = "The old category has a current balance of $" + QPTrim$(Using("##,##0.00", Abs(CreditAmt(x)))) + ". Please resolve this outstanding balance before continuing."
            frmBLMessageBoxJr.Label1.Top = 700
            frmBLMessageBoxJr.Show vbModal
          End If
          're-populate the affected fields with the old data
          fptxtCode(x).Col = 0
          fptxtCode(x).ColText = TempBillCat1
          fptxtCode(x).Col = 1
          fptxtCode(x).ColText = TempDESC1
          fptxtRev(x).Text = TempRev1
          fptxtCode(x).BackColor = &H80000005
          fptxtRev(x).BackColor = &H80000005
          fptxtCode(x).SetFocus
        End If
      Case 1:
        fptxtCode(x).Col = 0
        If QPTrim$(fptxtCode(x).ColText) <> QPTrim$(TempBillCat2) Then
          fptxtCode(x).BackColor = &H80FFFF
          fptxtRev(x).BackColor = &H80FFFF
          AnyCreditBal = True
          If CreditAmt(x) < 0 Then
            frmBLMessageBoxJr.Label1.Caption = "The old category has a current balance of -$" + QPTrim$(Using("##,##0.00", Abs(CreditAmt(x)))) + ". Please resolve this outstanding balance before continuing."
            frmBLMessageBoxJr.Label1.Top = 700
            frmBLMessageBoxJr.Show vbModal
          Else
            frmBLMessageBoxJr.Label1.Caption = "The old category has a current balance of $" + QPTrim$(Using("##,##0.00", Abs(CreditAmt(x)))) + ". Please resolve this outstanding balance before continuing."
            frmBLMessageBoxJr.Label1.Top = 700
            frmBLMessageBoxJr.Show vbModal
          End If
          fptxtCode(x).Col = 0
          fptxtCode(x).ColText = TempBillCat2
          fptxtRev(x).Text = TempRev2
          fptxtCode(x).BackColor = &H80000005
          fptxtRev(x).BackColor = &H80000005
          fptxtCode(x).SetFocus
        End If
      Case 2:
        fptxtCode(x).Col = 0
        If QPTrim$(fptxtCode(x).ColText) <> QPTrim$(TempBillCat3) Then
          fptxtCode(x).BackColor = &H80FFFF
          fptxtRev(x).BackColor = &H80FFFF
          AnyCreditBal = True
          If CreditAmt(x) < 0 Then
            frmBLMessageBoxJr.Label1.Caption = "The old category has a current balance of -$" + QPTrim$(Using("##,##0.00", Abs(CreditAmt(x)))) + ". Please resolve this outstanding balance before continuing."
            frmBLMessageBoxJr.Label1.Top = 700
            frmBLMessageBoxJr.Show vbModal
          Else
            frmBLMessageBoxJr.Label1.Caption = "The old category has a current balance of $" + QPTrim$(Using("##,##0.00", Abs(CreditAmt(x)))) + ". Please resolve this outstanding balance before continuing."
            frmBLMessageBoxJr.Label1.Top = 700
            frmBLMessageBoxJr.Show vbModal
          End If
          fptxtCode(x).Col = 0
          fptxtCode(x).ColText = TempBillCat3
          fptxtRev(x).Text = TempRev3
          fptxtCode(x).BackColor = &H80000005
          fptxtRev(x).BackColor = &H80000005
          fptxtCode(x).BackColor = &H80000005
          fptxtRev(x).BackColor = &H80000005
          fptxtCode(x).SetFocus
        End If
      Case 3:
        fptxtCode(x).Col = 0
        If QPTrim$(fptxtCode(x).ColText) <> QPTrim$(TempBillCat4) Then
          fptxtCode(x).BackColor = &H80FFFF
          fptxtRev(x).BackColor = &H80FFFF
          AnyCreditBal = True
          If CreditAmt(x) < 0 Then
            frmBLMessageBoxJr.Label1.Caption = "The old category has a current balance of -$" + QPTrim$(Using("##,##0.00", Abs(CreditAmt(x)))) + ". Please resolve this outstanding balance before continuing."
            frmBLMessageBoxJr.Label1.Top = 700
            frmBLMessageBoxJr.Show vbModal
          Else
            frmBLMessageBoxJr.Label1.Caption = "The old category has a current balance of $" + QPTrim$(Using("##,##0.00", Abs(CreditAmt(x)))) + ". Please resolve this outstanding balance before continuing."
            frmBLMessageBoxJr.Label1.Top = 700
            frmBLMessageBoxJr.Show vbModal
          End If
          fptxtCode(x).Col = 0
          fptxtCode(x).ColText = TempBillCat4
          fptxtRev(x).Text = TempRev4
          fptxtCode(x).BackColor = &H80000005
          fptxtRev(x).BackColor = &H80000005
          fptxtCode(x).SetFocus
        End If
      Case 4:
        fptxtCode(x).Col = 0
        If QPTrim$(fptxtCode(x).ColText) <> QPTrim$(TempBillCat5) Then
          fptxtCode(x).BackColor = &H80FFFF
          fptxtRev(x).BackColor = &H80FFFF
          AnyCreditBal = True
          If CreditAmt(x) < 0 Then
            frmBLMessageBoxJr.Label1.Caption = "The old category has a current balance of -$" + QPTrim$(Using("##,##0.00", Abs(CreditAmt(x)))) + ". Please resolve this outstanding balance before continuing."
            frmBLMessageBoxJr.Label1.Top = 700
            frmBLMessageBoxJr.Show vbModal
          Else
            frmBLMessageBoxJr.Label1.Caption = "The old category has a current balance of $" + QPTrim$(Using("##,##0.00", Abs(CreditAmt(x)))) + ". Please resolve this outstanding balance before continuing."
            frmBLMessageBoxJr.Label1.Top = 700
            frmBLMessageBoxJr.Show vbModal
          End If
          fptxtCode(x).Col = 0
          fptxtCode(x).ColText = TempBillCat5
          fptxtRev(x).Text = TempRev5
          fptxtCode(x).BackColor = &H80000005
          fptxtRev(x).BackColor = &H80000005
          fptxtCode(x).SetFocus
        End If
      End Select
    End If
  Next x
  
  If AnyCreditBal = False Then Exit Function
  
  fptxtCode(0).Col = 2
  RevType = fptxtCode(0).ColText
  If RevType = "M" Then
    fptxtRev(0).BackColor = &H80FFFF
    fptxtRev(0) = Using$("###0", TempRev1)
  ElseIf RevType = "S" Then
    fptxtRev(0) = Using("$###,###,##0.00", TempRev1)
  ElseIf RevType = "F" Then
    fptxtRev(0) = Using$("###0", TempRev1)
  Else
    fptxtRev(0).Text = ""
  End If
  
  fptxtCode(1).Col = 2
  RevType = fptxtCode(1).ColText
  If RevType = "M" Then
    fptxtRev(1).BackColor = &H80FFFF
    fptxtRev(1) = Using$("###0", TempRev2)
  ElseIf RevType = "S" Then
    fptxtRev(1) = Using("$###,###,##0.00", TempRev2)
  ElseIf RevType = "F" Then
    fptxtRev(1) = Using$("###0", TempRev2)
  Else
    fptxtRev(1).Text = ""
  End If

  fptxtCode(2).Col = 2
  RevType = fptxtCode(2).ColText
  If RevType = "M" Then
    fptxtRev(2).BackColor = &H80FFFF
    fptxtRev(2) = Using$("###0", TempRev3)
  ElseIf RevType = "S" Then
    fptxtRev(2) = Using("$###,###,##0.00", TempRev3)
  ElseIf RevType = "F" Then
    fptxtRev(2) = Using$("###0", TempRev3)
  Else
    fptxtRev(2).Text = ""
  End If

  fptxtCode(3).Col = 2
  RevType = fptxtCode(3).ColText
  If RevType = "M" Then
    fptxtRev(3).BackColor = &H80FFFF
    fptxtRev(3) = Using$("###0", TempRev4)
  ElseIf RevType = "S" Then
    fptxtRev(3) = Using("$###,###,##0.00", TempRev4)
  ElseIf RevType = "F" Then
    fptxtRev(3) = Using$("###0", TempRev4)
  Else
    fptxtRev(3).Text = ""
  End If

  fptxtCode(4).Col = 2
  RevType = fptxtCode(4).ColText
  If RevType = "M" Then
    fptxtRev(4).BackColor = &H80FFFF
    fptxtRev(4) = Using$("###0", TempRev5)
  ElseIf RevType = "S" Then
    fptxtRev(4) = Using("$###,###,##0.00", TempRev5)
  ElseIf RevType = "F" Then
    fptxtRev(4) = Using$("###0", TempRev5)
  Else
    fptxtRev(4).Text = ""
  End If
  
  Exit Function
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustEdit", "AnyCreditBal", Erl)
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

Private Sub TightenCodeEntries() 'NOT BEING USED
  Dim x As Integer
  Dim ThisIndex As Integer
  Dim Full As Integer
  Dim CodeRec As ARNewCatCodeRecType
  Dim CHandle As Integer
  Dim CatCodeCnt As Integer
  Dim y As Integer
  Dim ChgCnt As Integer
  Dim RevType As String * 1
  Dim Col1$
  Dim Col2$
  Dim Col0$
  
  'it is important that category codes are saved such that
  'no gaps are between category rows...when license fees are processed
  'when the program reached a category code that is empty then
  'it no longer calculates so if a gap is left then any categories
  'saved after a gap will not be processed...that is why this sub
  'exists
  
  On Error GoTo ERRORSTUFF
  OpenCatCodeFile CHandle
  CatCodeCnt = LOF(CHandle) / Len(CodeRec)
  Close CHandle
  
  If CatCodeCnt = 0 Then
    Exit Sub
  End If
  
  Do 'keep looping until all the gaps are ironed out
    For x = 0 To 3 'look for an empty row
      If QPTrim$(fptxtCode(x).Text) = "" Then 'got one so...
        For y = x + 1 To 4 'start above the empty row
          If fptxtCode(y).Text <> "" Then 'OK, a row above an empty row
            'is found to be populated
            fptxtCode(y).Col = 0 'assign column 0 to Col0
            Col0$ = QPTrim$(fptxtCode(y).ColText)
            fptxtCode(y).Col = 1 'assign column 1 to Col1
            Col1$ = QPTrim$(fptxtCode(y).ColText)
            fptxtCode(y).Col = 2 'assign column 2 to Col2
            Col2$ = QPTrim$(fptxtCode(y).ColText)
            'now populate the empty row with the data saved from the row above it
            fptxtCode(x).Text = Col0$ & Chr(9) & Col1$ + Chr(9) & Col2$
            'make the row above empty
            fptxtCode(y).Text = ""
            'now format the Rev/Multi/Flat field according to the
            'freshly populated row's fee method
            fptxtCode(x).Col = 2
            'it's necessary for any Rev/Multi/Flat field with a category
            'code to have a value of either 0 or $0.00 because
            'when assigning backcolors the Using requires a value in the field
            If QPTrim$(fptxtCode(x).ColText) = "F" Or QPTrim$(fptxtCode(x).ColText) = "M" Then
              fptxtRev(x).Text = Using$("$###,###,##0.00", Val(fptxtRev(y).Text))
            Else
              fptxtRev(x).Text = Using("$###,###,##0.00", Val(fptxtRev(y).Text))
            End If
            fptxtRev(y).Text = Using$("$###,###,##0.00", 0)
            ChgCnt = ChgCnt + 1
            Exit For
          End If
        Next y
      End If
    Next x
    If x = 4 Then Exit Do
  Loop
  
  'now go through and check the fee methods and make sure
  'the values and back colors are correct
  If ChgCnt > 0 Then
    ChgCnt = 0
    For x = 0 To 4
      fptxtCode(x).Col = 2
      RevType = QPTrim$(fptxtCode(x).ColText)
      If RevType = "M" Then
        fptxtRev(x).BackColor = &H80FFFF
        fptxtRev(x) = Using$("###0", fptxtRev(x).Text)
        ChgCnt = ChgCnt + 1
      ElseIf RevType = "S" Then
        fptxtRev(x).BackColor = &HFFFFFF
        fptxtRev(x) = Using("$###,###,##0.00", fptxtRev(x).Text)
        ChgCnt = ChgCnt + 1
      ElseIf RevType = "F" Then
        fptxtRev(x).BackColor = &HC0C0FF
        fptxtRev(x) = Using$("###0", fptxtRev(x).Text)
        ChgCnt = ChgCnt + 1
      Else
        fptxtRev(x).BackColor = &HFFFFFF
        fptxtRev(x).Text = " "
      End If
    Next x
    fptxtRev(ChgCnt - 1).SetFocus
  End If
  
  Exit Sub

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustEdit", "TightenCodeEntries", Erl)
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

Private Function CodeAmtsOK() As Boolean
  Dim x As Integer
  'this function makes sure that if there is a category code then
  'there is also a value in the rev/multi/flat field
  On Error Resume Next
  CodeAmtsOK = True
  For x = 0 To 4
    If Val(fptxtRev(x).Text) <> 0 Then
      fptxtCode(x).Col = 0
      If QPTrim$(fptxtCode(x).ColText) = "" Then
        frmBLMessageBoxJr.Label1.Caption = "A value has been entered for 'Revenues/Multi' but no code is indicated. Please correct this situation before saving."
        frmBLMessageBoxJr.Label1.Top = 700
        frmBLMessageBoxJr.Show vbModal
        If fptxtCode(x).Enabled = True Then
          fptxtCode(x).SetFocus
        End If
        CodeAmtsOK = False
        Exit Function
      End If
    End If
  Next x
  
End Function
Private Function Check4DupCodes() As Boolean
  Dim x As Integer
  Dim Nextx As Integer
  Dim ThisCode$
    
  Check4DupCodes = False
  Nextx = 0
  Do
    fptxtCode(Nextx).Col = 0
    ThisCode = QPTrim$(fptxtCode(Nextx).ColText)
    For x = 0 To 4 'first look for duplicate code numbers...meaning codes
    'that are already in use for this customer
      If x <> Nextx Then
        If ThisCode = "" Then GoTo NoText
        fptxtCode(x).Col = 0
        If QPTrim$(fptxtCode(x).ColText) = ThisCode Then
        'if a match is found (not the indexed code) then a duplicate exists
          frmBLMessageBoxJr.Label1.Caption = "This code is already in use for this business. Please select a different code."
          frmBLMessageBoxJr.Label1.Top = 800
          frmBLMessageBoxJr.Show vbModal
          Check4DupCodes = True
'          Select Case Nextx 'put back the data that was in the
          Select Case x 'put back the data that was in the
'          'indexed fields before the change was made
            Case 0:
'              fptxtCode(x).Col = 0
'              fptxtCode(x).ColText = TempBillCat1
'              fptxtCode(x).Col = 1
'              fptxtCode(x).ColText = TempDESC1
'              fptxtCode(x).Col = 2
'              fptxtCode(x).ColText = TempType1
'              If TempType1 = "F" Then
'                fptxtRev(x).BackColor = &H80FFFF
'              ElseIf TempType1 = "M" Then
'                fptxtRev(x).BackColor = &HFFFFFF
'              Else
'                fptxtRev(x).BackColor = &HC0C0FF
'              End If
              If fptxtCode(x).Enabled = True Then
                fptxtCode(x).SetFocus
              End If
'              If fptxtRev(x).Enabled = True Then
'                If TempRev1 >= 0 Then
'                  fptxtRev(x).Text = TempRev1
'                Else
'                  fptxtRev(x).Text = ""
'                End If
'              End If
            Case 1:
'              fptxtCode(x).Col = 0
'              fptxtCode(x).ColText = TempBillCat2
'              fptxtCode(x).Col = 1
'              fptxtCode(x).ColText = TempDESC2
'              fptxtCode(x).Col = 2
'              fptxtCode(x).ColText = TempType2
'              If TempType2 = "F" Then
'                fptxtRev(x).BackColor = &HC0C0FF
'              ElseIf TempType2 = "M" Then
'                fptxtRev(x).BackColor = &H80FFFF
'              Else
'                fptxtRev(x).BackColor = &HFFFFFF
'              End If
              If fptxtCode(x).Enabled = True Then
                fptxtCode(x).SetFocus
              End If
'              If fptxtRev(x).Enabled = True Then
'                If TempRev2 >= 0 Then
'                  fptxtRev(x).Text = TempRev2
'                Else
'                  fptxtRev(x).Text = ""
'                End If
'              End If
            Case 2:
'              fptxtCode(x).Col = 0
'              fptxtCode(x).ColText = TempBillCat3
'              fptxtCode(x).Col = 1
'              fptxtCode(x).ColText = TempDESC3
'              fptxtCode(x).Col = 2
'              fptxtCode(x).ColText = TempType3
'              If TempType3 = "F" Then
'                fptxtRev(x).BackColor = &H80FFFF
'              ElseIf TempType3 = "M" Then
'                fptxtRev(x).BackColor = &HFFFFFF
'              Else
'                fptxtRev(x).BackColor = &HC0C0FF
'              End If
              If fptxtCode(x).Enabled = True Then
                fptxtCode(x).SetFocus
              End If
'              If fptxtRev(x).Enabled = True Then
'                If TempRev3 >= 0 Then
'                  fptxtRev(x).Text = TempRev3
'                Else
'                  fptxtRev(x).Text = ""
'                End If
'              End If
            Case 3:
'              fptxtCode(x).Col = 0
'              fptxtCode(x).ColText = TempBillCat4
'              fptxtCode(x).Col = 1
'              fptxtCode(x).ColText = TempDESC4
'              fptxtCode(x).Col = 2
'              fptxtCode(x).ColText = TempType4
'              If TempType4 = "F" Then
'                fptxtRev(x).BackColor = &H80FFFF
'              ElseIf TempType4 = "M" Then
'                fptxtRev(x).BackColor = &HFFFFFF
'              Else
'                fptxtRev(x).BackColor = &HC0C0FF
'              End If
'              If TempBillCat4 = "" Then
'                fptxtCode(x).Text = ""
'              End If
              If fptxtCode(x).Enabled = True Then
                fptxtCode(x).SetFocus
              End If
'              If fptxtRev(x).Enabled = True Then
'                If TempRev4 >= 0 Then
'                  fptxtRev(x).Text = TempRev4
'                Else
'                  fptxtRev(x).Text = ""
'                End If
'              End If
            Case 4:
'              fptxtCode(x).Col = 0
'              fptxtCode(x).ColText = TempBillCat5
'              fptxtCode(x).Col = 1
'              fptxtCode(x).ColText = TempDESC5
'              fptxtCode(x).Col = 2
'              fptxtCode(x).ColText = TempType5
'              If TempType5 = "F" Then
'                fptxtRev(x).BackColor = &H80FFFF
'              ElseIf TempType5 = "M" Then
'                fptxtRev(x).BackColor = &HFFFFFF
'              Else
'                fptxtRev(x).BackColor = &HC0C0FF
'              End If
              If fptxtCode(x).Enabled = True Then
                fptxtCode(x).SetFocus
              End If
'              If fptxtRev(x).Enabled = True Then
'                If TempRev5 >= 0 Then
'                  fptxtRev(x).Text = TempRev5
'                Else
'                  fptxtRev(x).Text = ""
'                End If
'              End If
            End Select
            Exit Do
        End If
      End If
NoText:
    Next x
    If Nextx = 4 Then Exit Do
    Nextx = Nextx + 1
    fptxtCode(Nextx).Col = 0
    If QPTrim$(fptxtCode(Nextx).ColText) = "" Then Exit Do
  Loop
End Function

Private Sub fptxtRev_LostFocus(Index As Integer)
  fptxtCode(Index).Col = 1
  If Len(fptxtCode(Index).Text) = 0 Then
    fptxtRev(Index).Text = ""
    Exit Sub
  End If
  
'  If QPTrim$(fptxtRev(Index).Text) = "" Then Exit Sub
  fptxtCode(Index).Col = 2
'  If fptxtCode(Index).ColText = "" Then Exit Sub
  If fptxtCode(Index).ColText = "S" Then
    If Len(fptxtRev(Index).Text) = 0 Then fptxtRev(Index).Text = 0
    fptxtRev(Index).Text = Using("$###,###,##0.00", fptxtRev(Index).Text)
  Else
    If Len(fptxtRev(Index).Text) = 0 Then
      fptxtRev(Index).Text = 0
    End If
'    fptxtRev(Index).Text = Using("###0", fptxtRev(Index).Text)
  End If

End Sub

Private Sub fptxtZip_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    If fptxtLicNum.Enabled = True Then
      fptxtLicNum.SetFocus
    End If
  ElseIf KeyCode = vbKeyUp Then
    If fptxtState.Enabled = True Then
      fptxtState.SetFocus
    End If
  End If
End Sub

Private Sub GetType(Index As Integer)

  Dim x As Integer
  Dim RevType As String * 1
  Dim ThisCode$
  Dim ChgCnt As Integer
  Dim ThisRev$
  
  On Error GoTo ERRORSTUFF
  x = Index
'  For x = 0 To 4
    'get the column needing to be read
    fptxtCode(x).Col = 0
    'assign column data to ThisCode
    ThisCode = QPTrim$(fptxtCode(x).ColText)
    'get the column containing the fee method
    fptxtCode(x).Col = 2
    'assign RevType that value
    RevType = fptxtCode(x).ColText
    If RevType = "N" Then
      fptxtRev(x).Text = "0"
      Exit Sub
    End If
    If RevType = "M" Then 'set the revenue field to yellow and to units instead of dollars
      fptxtRev(x).BackColor = &H80FFFF
      'If the rev field is blank then give it a zero
      If QPTrim$(fptxtRev(x).Text) = "" Then fptxtRev(x).Text = 0
      'now format the data according to the M method
      If ThisCode <> "" Then
        fptxtRev(x).Text = Using$("###0", fptxtRev(x).Text)
      Else
        fptxtRev(x).Text = Using$("###0", 0)
      End If
    ElseIf RevType = "S" Then
      'back color for S is white
      fptxtRev(x).BackColor = &HFFFFFF
      'If the rev field is blank then give it a zero
      If QPTrim$(fptxtRev(x).Text) = "" Then fptxtRev(x).Text = 0
      'now format the data according to the S method
      If ThisCode <> "" Then
        fptxtRev(x).Text = QPTrim$(Using("$###,###,##0.00", fptxtRev(x).Text))
      Else
        fptxtRev(x).Text = Using("$###,###,##0.00", 0)
      End If
    ElseIf RevType = "F" Then
      'back color for F is light red
      fptxtRev(x).BackColor = &HC0C0FF
      'If the rev field is blank then give it a zero
      If QPTrim$(fptxtRev(x).Text) = "" Then fptxtRev(x).Text = 0
      'now format the data according to the F method
      If ThisCode <> "" Then
        fptxtRev(x).Text = Using$("###0", Val(fptxtRev(x).Text))
      Else
        fptxtRev(x).Text = Using$("###0", 0)
      End If
    ElseIf RevType = " " Then
      fptxtRev(x).BackColor = &HFFFFFF
      ChgCnt = x
    End If
'  Next x
  
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustEdit", "fptxtCode_LostFocus", Erl)
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

Private Sub CleanUpCodeEntries()
  Dim x As Integer
  
  For x = 0 To 4
    If QPTrim$(fptxtCode(x).Text) = "" Then
      fptxtRev(x).Text = ""
      fptxtRev(x).BackColor = &H80000005
    End If
  Next x

End Sub
