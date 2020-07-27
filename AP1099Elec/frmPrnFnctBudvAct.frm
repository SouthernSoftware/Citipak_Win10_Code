VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmPrnFnctBudvAct 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Function Budget V Actual"
   ClientHeight    =   8640
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   Icon            =   "frmPrnFnctBudvAct.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboIncAcct 
      Height          =   384
      Left            =   6024
      TabIndex        =   5
      Top             =   5064
      Width           =   1404
      _Version        =   196608
      _ExtentX        =   2476
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
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmPrnFnctBudvAct.frx":08CA
   End
   Begin LpLib.fpCombo fpcboReportOn 
      Height          =   384
      Left            =   6024
      TabIndex        =   6
      Top             =   5652
      Width           =   1404
      _Version        =   196608
      _ExtentX        =   2476
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
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmPrnFnctBudvAct.frx":0CF1
   End
   Begin LpLib.fpCombo fpcboFunction2 
      Height          =   384
      Left            =   6024
      TabIndex        =   2
      Top             =   3312
      Width           =   3804
      _Version        =   196608
      _ExtentX        =   6710
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
      Columns         =   3
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
      ColDesigner     =   "frmPrnFnctBudvAct.frx":1118
   End
   Begin LpLib.fpCombo fpcboFunction1 
      Height          =   384
      Left            =   6024
      TabIndex        =   1
      Top             =   2736
      Width           =   3804
      _Version        =   196608
      _ExtentX        =   6710
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
      Columns         =   3
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
      ColDesigner     =   "frmPrnFnctBudvAct.frx":15C3
   End
   Begin LpLib.fpCombo fpcboSubtot 
      Height          =   384
      Left            =   6024
      TabIndex        =   4
      Top             =   4488
      Width           =   1404
      _Version        =   196608
      _ExtentX        =   2476
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
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmPrnFnctBudvAct.frx":1A6E
   End
   Begin LpLib.fpCombo fpcboRepNum 
      Height          =   384
      Left            =   6024
      TabIndex        =   3
      Top             =   3900
      Width           =   3948
      _Version        =   196608
      _ExtentX        =   6964
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
      Columns         =   2
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   -1
      ColumnWidthScale=   2
      RowHeight       =   -1
      WrapList        =   0   'False
      WrapWidth       =   0
      AutoSearch      =   2
      SearchMethod    =   2
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
      ColDesigner     =   "frmPrnFnctBudvAct.frx":1E95
   End
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   384
      Left            =   6024
      TabIndex        =   7
      Top             =   6240
      Width           =   1908
      _Version        =   196608
      _ExtentX        =   3365
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
      Columns         =   1
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
      ScrollBarH      =   3
      DataFieldList   =   ""
      ColumnEdit      =   0
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   3504
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
      ColDesigner     =   "frmPrnFnctBudvAct.frx":2314
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Esc E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   10080
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7488
      Width           =   1332
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F10 &Print"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   8400
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7488
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   8280
      Width           =   12192
      _ExtentX        =   21505
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7133
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "4:23 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "12/6/2004"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin EditLib.fpDateTime txtDate 
      Height          =   372
      Left            =   6024
      TabIndex        =   0
      Top             =   2160
      Width           =   1692
      _Version        =   196608
      _ExtentX        =   2984
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
      Text            =   "11/06/2001"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
      DateMax         =   "20350101"
      DateMin         =   "19800101"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "19800101"
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
      ButtonColor     =   14737632
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Include Acct Numbers:"
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
      Height          =   324
      Index           =   5
      Left            =   2784
      TabIndex        =   19
      Top             =   5088
      Width           =   3060
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Report On:"
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
      Height          =   324
      Index           =   6
      Left            =   4440
      TabIndex        =   18
      Top             =   5664
      Width           =   1404
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select Report Type: "
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
      Left            =   3528
      TabIndex        =   17
      Top             =   6240
      Width           =   2388
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Print Revenues on Seperate Page:"
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
      Height          =   324
      Index           =   4
      Left            =   1152
      TabIndex        =   16
      Top             =   4512
      Width           =   4692
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Report Number:"
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
      Height          =   324
      Index           =   2
      Left            =   3816
      TabIndex        =   15
      Top             =   3948
      Width           =   2028
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Function:"
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
      Height          =   324
      Index           =   1
      Left            =   3408
      TabIndex        =   14
      Top             =   2796
      Width           =   2436
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Date:"
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
      Height          =   324
      Index           =   0
      Left            =   4272
      TabIndex        =   13
      Top             =   2220
      Width           =   1572
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3210
      Top             =   600
      Width           =   5772
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Function Budget vs Actual"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3318
      TabIndex        =   12
      Top             =   840
      Width           =   5604
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   5244
      Left            =   1932
      Top             =   1776
      Width           =   8340
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Function:"
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
      Height          =   324
      Index           =   1
      Left            =   3336
      TabIndex        =   11
      Top             =   3372
      Width           =   2508
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      Height          =   972
      Left            =   3210
      Top             =   480
      Width           =   5772
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuPrnScn 
         Caption         =   "Prin&t Screen"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmPrnFnctBudvAct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim GLSetup As GLSetupRecType
Dim GLAcct    As GLAcctRecType
Dim GLFundIdx As GLFundIndexType
Dim GLAcctidx As GLAcctIndexType
Dim GLTrans   As GLTransRecType
Dim GLFund As GLFundRecType
Dim GLFNCT As GLFNCTRecType
Dim GLFNCTIdx As GLFNCTIndexType
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim FY1BegDate As Integer, FY1EndDate As Integer, FY2BegDate As Integer, FY2EndDate As Integer
Dim StartFnct As String, EndFnct As String, FYStartDate As Integer
Dim ActiveYear As Integer
Private Sub cmdExit_Click()
  frmGLFunctionReports.Show
  Unload Me
End Sub
Private Sub fpcboRptType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRptType.ListDown = True
  End If
  If fpcboRptType.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      cmdPrint.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboReportOn.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        ClearInUse PWcnt
      End If
    End If
  End If
End Sub

Private Function ValidDate()
  Dim TempDate As Integer
  GetFYDates FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate
  If CheckValDate(txtDate) = True Then
    TempDate = DateDiff("d", "12/31/1979", txtDate)
    ValidDate = True
    If TempDate >= FY2BegDate Then
      ActiveYear = 2
      FYStartDate = FY2BegDate
    Else
      ActiveYear = 1
      FYStartDate = FY1BegDate
    End If
  Else
    MsgBox "Date Is Not Valid. Please Correct.", vbOKOnly, "Invalid Date"
    ValidDate = False
    Exit Function
  
  End If
End Function
Private Function ValidFunctions()
  If fpcboFunction1.Text <> "" And fpcboFunction2.Text <> "" Then
    fpcboFunction1.Col = 0
    fpcboFunction2.Col = 0
    If fpcboFunction1.ColText > fpcboFunction2.ColText Then
      MsgBox "Invalid Function Selection, The Beginning Function Should Be Less or Equal to Ending Function.", vbOKOnly, "Invalid Selection"
      ValidFunctions = False
    Else
      ValidFunctions = True
      StartFnct = QPTrim(fpcboFunction1.ColText)
      EndFnct = QPTrim(fpcboFunction2.ColText)
    End If
  Else
    MsgBox "Function Selections May Not Be Left Blank.", vbOKOnly, "Invalid Selection"
  End If
End Function

Private Sub cmdPrint_Click()
  If ValidDate = True Then
    If ValidFunctions = True Then
      If fpcboRptType.ListIndex = 0 Then
        rptopt = 1
      ElseIf fpcboRptType.ListIndex = 1 Then
        rptopt = 2
      End If
      If rptopt = 1 Then
        PrintBgtAct
      ElseIf rptopt = 2 Then
'1 is graphic, 2 is text
        PrintBgtAct2 (rptopt)
      End If
    End If
  End If
End Sub
Private Sub fpcboFunction2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboFunction2.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboFunction2.ListIndex = -1
    fpcboFunction2.Action = ActionClearSearchBuffer
  End If
  If fpcboFunction2.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboRepNum.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboFunction1.SetFocus
        KeyCode = 0
      End If
    End If
  End If
     

End Sub
Private Sub fpcboFunction1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboFunction1.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboFunction1.ListIndex = -1
    fpcboFunction1.Action = ActionClearSearchBuffer
  End If
  If fpcboFunction1.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboFunction2.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        txtDate.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub
'Private Sub fpcboDetSum_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeySpace Then
'    fpcboDetSum.ListDown = True
'  End If
'  If KeyCode = vbKeyDelete Then
'    fpcboDetSum.ListIndex = -1
'    fpcboDetSum.Action = ActionClearSearchBuffer
'  End If
'  If fpcboDetSum.ListDown <> True Then
'    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
'      fpcboRepNum.SetFocus
'      KeyCode = 0
'    Else
'      If KeyCode = vbKeyUp Then
'        fpcboFund2.SetFocus
'        KeyCode = 0
'      End If
'    End If
'  End If
'
'End Sub
Private Sub fpcboRepNum_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboRepNum.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboRepNum.ListIndex = -1
    fpcboRepNum.Action = ActionClearSearchBuffer
  End If
  If fpcboRepNum.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboSubtot.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboFunction2.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub
'Private Sub fpcboPagebrk_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeySpace Then
'    fpcboPagebrk.ListDown = True
'  End If
'  If KeyCode = vbKeyDelete Then
'    fpcboPagebrk.ListIndex = -1
'    fpcboPagebrk.Action = ActionClearSearchBuffer
'  End If
'  If fpcboPagebrk.ListDown <> True Then
'    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
'      fpcboSubtot.SetFocus
'      KeyCode = 0
'    Else
'      If KeyCode = vbKeyUp Then
'        fpcboRepNum.SetFocus
'        KeyCode = 0
'      End If
'    End If
'  End If
'
'End Sub
Private Sub fpcboSubtot_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboSubtot.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboSubtot.ListIndex = -1
    fpcboSubtot.Action = ActionClearSearchBuffer
  End If
  If fpcboSubtot.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboIncAcct.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboRepNum.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub
Private Sub fpcboIncAcct_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboIncAcct.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboIncAcct.ListIndex = -1
    fpcboIncAcct.Action = ActionClearSearchBuffer
  End If
  If fpcboIncAcct.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboReportOn.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboSubtot.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub
Private Sub fpcboReportOn_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboReportOn.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboReportOn.ListIndex = -1
    fpcboReportOn.Action = ActionClearSearchBuffer
  End If
  If fpcboReportOn.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboRptType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboIncAcct.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
'    Case vbKeyDown, vbKeyReturn:
'      SendKeys "{Tab}"
'      KeyCode = 0
'    Case vbKeyUp:
'      SendKeys "+{Tab}"
'      KeyCode = 0
    Case vbKeyEscape:
      cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      cmdPrint_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  Me.HelpContextID = hlpFunctionBudtVAct
  'FundList fpcboFund1
  'fpcboFund1.RemoveItem 0
  'FundstoList fpcboFund1
 ' FundstoList fpcboFund2
  'FundList fpcboFund2
  'fpcboFund2.RemoveItem 0
  txtDate.Text = Format(Now, "mm/dd/yyyy")
  'fpcboDetSum.AddItem "Detail"
  'fpcboDetSum.AddItem "Summary"
 'fpcboDetSum.ListIndex = 0
  fpcboRepNum.InsertRow = "1" & Chr$(9) & "Bgt,Mo,YTD,Var:Req."
  fpcboRepNum.InsertRow = "2" & Chr$(9) & "Bgt,Enc,YTD,Var:Req."
  fpcboRepNum.InsertRow = "3" & Chr$(9) & "Bgt,QTD,YTD,Var:Req."
  fpcboRepNum.ListIndex = 0
'  fpcboPagebrk.AddItem "Yes"
'  fpcboPagebrk.AddItem "No"
'  fpcboPagebrk.ListIndex = 1
  fpcboSubtot.AddItem "Yes"
  fpcboSubtot.AddItem "No"
  fpcboSubtot.ListIndex = 1
  fpcboIncAcct.AddItem "Yes"
  fpcboIncAcct.AddItem "No"
  fpcboIncAcct.ListIndex = 0
  fpcboReportOn.AddItem "Revenues Only"
  fpcboReportOn.AddItem "Expenditures Only"
  fpcboReportOn.AddItem "Both Revenues & Expenditures"
  fpcboReportOn.ListIndex = 2
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
  FillFNCTList fpcboFunction1
  FillFNCTList fpcboFunction2
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub

Private Sub PrintBgtAct()
  Dim CommaFmt As String, TotalFmt As String, RunBalFmt As String
  Dim SumLine As String, BgtFmt As String, BSumLine As String, PSumLine As String
  Dim DivLine As String, DivLine2 As String, FF As String, RptTitle As String
  Dim MaxLines As Integer, Col1 As Integer, Col2 As Integer, Col3 As Integer
  Dim Col4 As Integer, Col5 As Integer, EndDate As Integer, PRNFile As Integer
  Dim M As String, HollyFlag As Boolean, Pitch12 As String, ThisFnct As String
  Dim DoingDetail As Boolean, SubTotalRevenues As Boolean, DeptOnNewPage As Boolean
  Dim WhichReport As Integer, GetMonth As Boolean, GetQtr As Boolean
  Dim RptMonth As String, ReportFile As String, FnctCode As String
  Dim FundIdxFile As Integer, NumFunds As Integer, Acct As Integer
  Dim FnctIdxFile As Integer, NumFncts As Integer, FnctName As String
  Dim AcctIdxFileNum As Integer, NumGLAccts As Integer, FundName As String
  Dim AcctFileNum As Integer, NumGLAcctRecs As Integer, Rec As Integer
  Dim TransFileNum As Integer, NumTrans As Long, NextTr As Long, AcctNum As String
  Dim BGTBal As Double, YTDBal As Double, MTDBal As Double, UsingFnct As Boolean
  Dim ECnt As Integer, RCnt As Integer, d As String, TransMonth As String
  Dim InThisQtr As Boolean, FirstTime As Boolean, Fnct As Long, FNCTRec As Long
  Dim Dept As String, LastDept As String, LastDeptName As String, cnt As Integer
  Dim Account As String, BudgetAmt As Double, DeptRecNum As Integer, DeptName As String
  Dim Pct As String, Variance As Double, ToPrint As String, Linecnt As Integer
  Dim MTDSum As Double, BgtSum As Double, YTDSum As Double, DeptBgtSum As Double
  Dim DeptYTDSum As Double, DeptENCSum As Double, DeptMTDSum As Double
  Dim FnctRevMTD As Double, FnctRevBgt As Double, FnctRevYTD As Double
  Dim EncSum As Double, FnctExpMTD As Double, FnctExpBgt As Double
  Dim FnctExpYTD As Double, FnctEncYTD As Double, EncBal As Double
  Dim DeptSummary As String, PageNum As Integer, Newrp As String
  Dim NumRpt As Integer, lab14 As String
  Dim GetEnc As Boolean, IncAcct As Boolean
  Dim DoingRevenues As Boolean, DoingExp As Boolean

  Select Case Mid$(fpcboReportOn.Text, 1, 1)
    Case "B"
      DoingRevenues = True
      DoingExp = True
      'rpt = 3
    Case "R"
      DoingRevenues = True
      DoingExp = False
      'rpt = 2
    Case "E"
      DoingRevenues = False
      DoingExp = True
      'rpt = 1
  End Select

'On Local Error GoTo GotErr

'  If InStr(UCase$(GLUserName), "HOLLY SPR") > 0 Then
'    HollyFlag = True
'    Pitch12$ = Chr$(27) + Chr$(38) + Chr$(107) + Chr$(52) + Chr$(83)
'  End If
'Do Not Need this Already Have Funds
  'GetFnctCodes FirstFund$, LastFund$
  CommaFmt$ = "###,###,###.##"  'format takes 14 chars
  TotalFmt$ = "#,###,###,###.##" 'format takes 16 chars
  RunBalFmt$ = "##########.##"
  SumLine$ = String$(16, "-")   'column summary line
  BgtFmt$ = "###,###,###"         'format takes 11 chars
  BSumLine$ = String$(11, "-")  'summary line for budget columns
  PSumLine$ = "----"            'summary line for Pct columns
  DivLine$ = String$(115, "-")   'dashed line
  DivLine2$ = String$(115, "=")  'Double Line
  FF$ = Chr$(12)
  ReDim Desc$(1)
  RptTitle$ = "Function Budget vs. Actual "
  MaxLines = 55
  '--Column offsets for printing amounts
  Col1 = 40    'Budget
  Col2 = 54   'Month or Actual YTD
  Col3 = 71    'YTD or Var
  Col4 = 88    'Enc
  Col5 = 105    'Pct
  EndDate = DateDiff("d", "12/31/1979", txtDate)
  M$ = Right$(txtDate, 2) + Left$(txtDate, 2)
'Format(DateAdd("d", FY2BegDate, "12-31-1979"), "mm/dd/yy")
  FrmShowPctComp.Label1 = "Printing Function Budget vs Actual Report"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmPrnFnctBudvAct, True
'  If fpcboDetSum.Text = "Detail" Then DoingDetail = True
  If fpcboSubtot.Text = "Yes" Then SubTotalRevenues = True
  If fpcboIncAcct.ListIndex = 1 Then
    IncAcct = False
  Else
    IncAcct = True
  End If
  
  '  SubTotalRevenues = True
  'ELSE
  '  SubTotalRevenues = False
  'END IF
'  If DoingDetail Then
'   ' ARptBudVAct.detopt = 1
'    '--if break on depts
'    If fpcboPagebrk.Text = "Yes" Then
'      DeptOnNewPage = True
'      'ARptBudVAct.deptpage = True
'    'Else
'      'ARptBudVAct.deptpage = False
'    End If
'  End If
  fpcboRepNum.Col = 0
  WhichReport = Val(fpcboRepNum.Text)
  Select Case WhichReport
  Case 1        'Bgt, Month, YTD
    GetMonth = True
    GetQtr = False
    GetEnc = False
    RptMonth = M$
''''''''''''''''123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456''''For 96
    Desc$(1) = "Description                                Budget          Month             YTD           Variance    Pct"
    lab14$ = "MTD"
    NumRpt = 1
    MTDBal# = 0
  Case 2        'Bgt, Enc, YTD, Variance
    GetMonth = False
    GetQtr = False
    GetEnc = True
    Desc$(1) = "Description                                Budget          Encumb            YTD           Variance    Pct"
    lab14$ = "Encumb"
    NumRpt = 2
  Case 3        'Bgt, Enc, YTD, Variance
    GetMonth = False
    GetQtr = True
    GetEnc = False
    Desc$(1) = "Description                                Budget           QTD              YTD           Variance    Pct"
    lab14 = "QTD"
    NumRpt = 3
  End Select
  If GetEnc = True Then
        'if need enc totals do this here
    FixPOEncumbRpt EndDate, FYStartDate
  End If
  MTDBal# = 0
  Newrp = "BGTACT"
  GetRPTName Newrp
  ReportFile$ = Newrp
  'ReportFile$ = Unique$(Path$)
  'P17$ = CHR$(27) + "(s17H"
  PRNFile = FreeFile
  Open ReportFile$ For Output As #PRNFile
  'PRINT #PrnFile, P17$;
'  If HollyFlag Then
'    Print #PRNFile, Pitch12$;
'  End If
  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum
  NumGLAcctRecs = LOF(AcctFileNum) / Len(GLAcct)
  OpenTransFile TransFileNum, NumTrans&
  OpenFundIdx FundIdxFile, NumFunds
  OpenFnctIdx FnctIdxFile, NumFncts
  
  ReDim RevAccts%(1 To NumGLAccts)              'Holds all rev acct record num
  ReDim ExpAccts%(1 To NumGLAccts)              'Holds all exp acct record num
'  ReDim FundList(1) As String                           'List of all active Funds
'  GetFundList FundList$(), NumFunds
  ReDim FnctList(1 To NumFncts) As String                           'List of all active Funds
  GetFnctList FnctList$(), NumFncts

  '--Build a list of revenue and exp accounts
  For Acct = 1 To NumGLAccts
    '--Initialize
    BGTBal# = 0
    YTDBal# = 0
    MTDBal# = 0
    Get AcctIdxFileNum, Acct, GLAcctidx
    Get AcctFileNum, GLAcctidx.RecNum, GLAcct
    '--Find what fund this account is in

    If GLAcct.FNCTRec > 0 Then
    FnctCode$ = GetFnctCode(GLAcct.FNCTRec) 'Left$(GLAcct.Num, GLFundLen)

    '--See if the account is in a fund we want to see
    If FnctCode$ >= StartFnct$ And FnctCode$ <= EndFnct$ Then

      '--Account is in fund, check to see if its proper type
      '--We want only revenue or expenditure accounts
      If GLAcct.Typ = "R" Or GLAcct.Typ = "E" Then

        '--Assign the Account Record Number to proper list
        Select Case GLAcct.Typ
        Case "E"
          ECnt = ECnt + 1
          ExpAccts%(ECnt) = GLAcctidx.RecNum

        Case "R"
          RCnt = RCnt + 1
          RevAccts%(RCnt) = GLAcctidx.RecNum
        End Select

        '--Get account balances
        '--There should be no beginning balances in rev & exp accts
        'YTDBal# = Round#(Acct.BegBal)           'get the beginning balance

        NextTr& = GLAcct.FrstTran 'get the first trans for this acct

        Do Until NextTr& = 0    'keep going 'til we run out

          Get TransFileNum, NextTr&, GLTrans

          '--Get MTD Account Balance if necessary
          If GLTrans.TRDATE >= FYStartDate And GLTrans.TRDATE <= EndDate Then

          If GetMonth Then
            'Lookhere change num2month to reflect year & month
            d$ = Format(DateAdd("d", GLTrans.TRDATE, "12-31-1979"), "mm/dd/yyyy")
            TransMonth = Right$(d$, 2) + Left$(d$, 2)
            'TransMonth = Num2Month%(GLTrans.TRDATE)
            If TransMonth = RptMonth Then
              Select Case GLAcct.Typ
              Case "E"
                MTDBal# = Round#(MTDBal# + GLTrans.DrAmt - GLTrans.CrAmt)
              Case "R"
                MTDBal# = Round#(MTDBal# + GLTrans.CrAmt - GLTrans.DrAmt)
              End Select
            End If
          End If

          If GetQtr Then
            InThisQtr = InQtr(GLTrans.TRDATE, EndDate)
            If InThisQtr Then
              Select Case GLAcct.Typ
                Case "E"
                  MTDBal# = Round#(MTDBal# + GLTrans.DrAmt - GLTrans.CrAmt)
                Case "R"
                  MTDBal# = Round#(MTDBal# + GLTrans.CrAmt - GLTrans.DrAmt)
              End Select
            End If
          End If

          '--Get YTD Account Balance
          '--original
          '--Carthage 9/17/96
          'IF Trans.TrDate <= EndDate THEN
          '--08/07/96 keeping funds open --
          '--Does'nt work when 2 years are open...what to do?
          'IF Trans.TrDate <= EndDate THEN
            Select Case GLAcct.Typ
            Case "E"
              YTDBal# = Round#(YTDBal# + GLTrans.DrAmt - GLTrans.CrAmt)
            Case "R"
              YTDBal# = Round#(YTDBal# + GLTrans.CrAmt - GLTrans.DrAmt)
            End Select
          End If

          NextTr& = GLTrans.NextTran              'Get the next transaction

        Loop
        '--Put the new totals in the file
        GLAcct.MTD = Round#(MTDBal#)
        GLAcct.YTD = Round#(YTDBal#)
        Put AcctFileNum, GLAcctidx.RecNum, GLAcct

      End If    '--test for rev or exp accts
    End If      '--End of acct in fund range test
    End If      '--if function rec stored in glacct > 0
    FrmShowPctComp.ShowPctComp Acct, NumGLAccts
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnFnctBudvAct, True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If

  Next          'Process next account

  ActivateControls frmPrnFnctBudvAct, True
  '--Now write the report to file.
'''''  PrintHelp "Generating Report..."
  FirstTime = True 'flag not to print page feed first time thru
  '--Process each function
  For Fnct = 1 To NumFncts
    ThisFnct$ = FnctList$(Fnct)
    If ThisFnct$ >= StartFnct$ And ThisFnct$ <= EndFnct$ Then
      UsingFnct = True
      FNCTRec = FindFnct(ThisFnct$)             'Get the fund name
      FnctName$ = QPTrim$(GetFnctTitle(FNCTRec))
     'Dept$ = ""

     ' LastDept$ = ""
     ' LastDeptName$ = ""
    Else
      UsingFnct = False
    End If

    If UsingFnct Then
      '--print a form feed for each new function
      If FirstTime Then
        FirstTime = False
      Else
       ' Print #PRNFile, FF$
      End If

      'RE$ = "Revenues"
      'If Not SubTotalRevenues Or Not DoingDetail Then
      '  GoSub PrintBVAPageHeader

      'End If
      If DoingRevenues Then

      '--Search thru list of revenue accounts
      For cnt = 1 To RCnt
        Rec = RevAccts%(cnt)
        Get AcctFileNum, Rec, GLAcct
        FnctCode$ = GetFnctCode(GLAcct.FNCTRec) 'Left$(GLAcct.Num, GLFundLen)

        If FnctCode$ = ThisFnct$ Then
          
          If IncAcct = True Then
            Account$ = QPTrim$(GLAcct.Num) + "  " + QPTrim$(Mid(GLAcct.Title, 1, 22))
          Else
            Account$ = QPTrim$(GLAcct.Title)
          End If
          Select Case ActiveYear
          Case 1
            BudgetAmt# = GLAcct.Bgt
          Case 2
            BudgetAmt# = GLAcct.NYApp
          End Select
          '----------
          'If SubTotalRevenues Then
            '--Extract the Dept$ from the G/L Acct
'            Dept$ = Mid$(GLAcct.Num, GLFundLen + 2, GLAcctLen)
'
'            '--if a new dept, get its name from the Dept name file
'            If Dept$ <> LastDept$ Then
'              DeptRecNum = FindDept(Dept$)
'              If DeptRecNum > 0 Then
'                DeptName$ = QPTrim$(GetDeptTitle$(DeptRecNum))
'              Else
'                'DeptName$ = "Department " + Dept$
'                DeptName$ = " "
'              End If
'            End If
'
'            '--Print Department Header first time thru
'            If Len(LastDeptName$) = 0 Then
'              '--if we're not printing departments on separate pages then
'              '--print a new page header
'              If DeptOnNewPage = False Then
'                If DoingDetail = True Then
'                  'Print #PRNFile, FF$
'                  'GoSub PrintBVAPageHeader
'                Else
'                  'Print #PRNFile,
'                  '*&*&*&*&*&
'                  'Print #PRNFile, "Revenues"
'                End If
'              End If
'              LastDeptName$ = DeptName$
'              LastDept$ = Dept$
'              If DoingDetail Then
'
'                'GoSub PrintDeptHeader
'              End If
'            End If
'
'            '--see if we need to subtotal dept
'            If Len(LastDept$) > 0 Then
'              If Dept$ <> LastDept$ Then
'                'GoSub PrintDeptTotals
'                If DoingDetail Then
'                  'GoSub PrintDeptHeader
'                End If
'              End If
'
'            End If

          'End If
          '===========
          Pct$ = GetPct$(GLAcct.YTD, BudgetAmt#)
          'If DoingDetail Then
            Variance# = Round#(GLAcct.YTD - BudgetAmt#) 'Acct.Bgt
            'ToPrint$ = Space$(96)
            'LSet ToPrint$ = Left$(Account$, 28)
            'Mid$(ToPrint$, Col1) = Using$(BgtFmt$, Str$(BudgetAmt#))   'chang
            ToPrint$ = FnctCode$ & "~" & GLAcct.Typ & "~~" & FnctName$ & "~" & Account$ & "~" & BudgetAmt#
            Select Case WhichReport
            Case 1, 3
'              Mid$(ToPrint$, Col2) = Using$(CommaFmt$, Str$(GLAcct.MTD))
'              Mid$(ToPrint$, Col3) = Using$(CommaFmt$, Str$(GLAcct.YTD))
'              Mid$(ToPrint$, Col4) = Using$(CommaFmt$, Str$(Variance#))
'              Mid$(ToPrint$, Col5) = Pct$
              ToPrint$ = ToPrint$ & "~" & GLAcct.MTD
              ToPrint$ = ToPrint$ & "~" & GLAcct.YTD
              ToPrint$ = ToPrint$ & "~" & Variance# & "~" & Pct$
            Case 2
'              Mid$(ToPrint$, Col2) = Using$(CommaFmt$, Str$(GLAcct.Encumb))
'              Mid$(ToPrint$, Col3) = Using$(CommaFmt$, Str$(GLAcct.YTD))
'              Mid$(ToPrint$, Col4) = Using$(CommaFmt$, Str$(Variance#))
'              Mid$(ToPrint$, Col5) = Pct$
              ToPrint$ = ToPrint$ & "~" & GLAcct.Encumb
              ToPrint$ = ToPrint$ & "~" & GLAcct.YTD
              ToPrint$ = ToPrint$ & "~" & Variance# & "~" & Pct$
            End Select
            Print #PRNFile, QPTrim$(ToPrint$)
            Linecnt = Linecnt + 1
            If Linecnt > MaxLines Then
              'Print #PRNFile, FF$
              'GoSub PrintBVAPageHeader
            End If
          'End If
          'changed
          If GetMonth Or GetQtr Then
            MTDSum# = MTDSum# + GLAcct.MTD
          End If
          BgtSum# = BgtSum# + BudgetAmt# 'Acct.Bgt
          YTDSum# = YTDSum# + GLAcct.YTD
'          DeptBgtSum# = DeptBgtSum# + BudgetAmt#
'          DeptYTDSum# = DeptYTDSum# + GLAcct.YTD
'          DeptENCSum# = DeptENCSum# + GLAcct.Encumb
'          DeptMTDSum# = DeptMTDSum# + GLAcct.MTD
'          LastDept$ = Dept$
'          LastDeptName$ = DeptName$
        End If
      Next      'Revenue Acct
      '--Summarize Revenues
'      If DoingDetail Then
'        'GoSub PrintSummaryLines
'      End If
      ToPrint$ = Space$(115)
      'LSet ToPrint$ = "Total Revenues"
      Pct$ = GetPct$(YTDSum#, BgtSum#)
      Variance# = YTDSum# - BgtSum#
      Select Case WhichReport
      Case 1, 3
        Variance# = YTDSum# - BgtSum#
       ' ToPrint$ = BgtSum# & "," & MTDSum# & "," & YTDSum# & "," & Variance# & "," & Pct$
        '--Reset vars
        FnctRevMTD# = MTDSum#
        MTDSum# = 0
      Case 2
        Variance# = YTDSum# - BgtSum#
       ' ToPrint$ = BgtSum# & "," & " 0" & "," & YTDSum# & "," & Variance# & "," & Pct$
      End Select
      'Print #PRNFile, RTrim$(ToPrint$)
      Linecnt = Linecnt + 1
      If Linecnt > MaxLines Then
        'Print #PRNFile, FF$
        'GoSub PrintBVAPageHeader
      End If
      End If
      FnctRevBgt# = BgtSum#
      FnctRevYTD# = YTDSum#
      EncSum# = 0
      BgtSum# = 0
      YTDSum# = 0
      'initialize dept variables
'      DeptBgtSum# = 0
'      DeptMTDSum# = 0
'      DeptYTDSum# = 0
'      DeptENCSum# = 0
'      LastDept$ = ""
'      LastDeptName$ = ""
      '--Search exp accounts list for accounts in this fund
      'RE$ = "Expenditures"
      If DoingExp Then

      For cnt = 1 To ECnt
        Rec = ExpAccts%(cnt)
        Get AcctFileNum, Rec, GLAcct

        FnctCode$ = GetFnctCode(GLAcct.FNCTRec) 'Left$(GLAcct.Num, GLFundLen)
        If FnctCode$ = ThisFnct$ Then
          FnctName$ = QPTrim$(GetFnctTitle(FNCTRec))
          If IncAcct = True Then
            Account$ = QPTrim$(GLAcct.Num) + "  " + QPTrim$(Mid(GLAcct.Title, 1, 22))
          Else
            Account$ = QPTrim$(GLAcct.Title)
          End If

          Select Case ActiveYear
          Case 1
            BudgetAmt# = GLAcct.Bgt
          Case 2
            BudgetAmt# = GLAcct.NYApp
          End Select
          '--Extract the Dept$ from the G/L Acct
'          Dept$ = Mid$(GLAcct.Num, GLFundLen + 2, GLAcctLen)

          '--if a new dept, get its name from the Dept name file
'          If Dept$ <> LastDept$ Then
'            DeptRecNum = FindDept(Dept$)
'            If DeptRecNum > 0 Then
'              DeptName$ = QPTrim$(GetDeptTitle$(DeptRecNum))
'            Else
'              'DeptName$ = "Department " + Dept$
'              DeptName$ = " "
'            End If
'          End If

          '--Print Department Header first time thru
'          If Len(LastDeptName$) = 0 Then
'            '--if we're not printing departments on separate pages then
'            '--print a new page header
'            If DeptOnNewPage = False Then
'              If DoingDetail = True Then
'                'Print #PRNFile, FF$
'                'GoSub PrintBVAPageHeader
'              Else
'                'Print #PRNFile,
'                'Print #PRNFile, "Expenditures"
'              End If
'            End If
'            LastDeptName$ = DeptName$
'            LastDept$ = Dept$
'            If DoingDetail Then
'              'GoSub PrintDeptHeader
'            End If
'          End If

          '--see if we need to subtotal dept
'          If Len(LastDept$) > 0 Then
'            If Dept$ <> LastDept$ Then
'              'GoSub PrintDeptTotals
'              If DoingDetail Then
'                'GoSub PrintDeptHeader
'              End If
'            End If
'          End If

          'If DoingDetail Then   'Print Account Detail
            ToPrint$ = Space$(115)

            ToPrint$ = FnctCode$ & "~" & GLAcct.Typ & "~" & Dept$ & "~" & FnctName$ & "~" & Account$
            Select Case WhichReport
            Case 1, 3
              Pct$ = GetPct$(GLAcct.YTD, BudgetAmt#) 'Acct.Bgt
              Variance# = Round(QPTrim(BudgetAmt# - GLAcct.YTD))
              ToPrint$ = ToPrint$ & "~" & BudgetAmt# & "~" & GLAcct.MTD & "~" & GLAcct.YTD
              ToPrint$ = ToPrint$ & "~" & Variance# & "~" & Pct$
            Case 2
              Pct$ = GetPct$(GLAcct.Encumb + GLAcct.YTD, BudgetAmt#) 'Acct.Bgt
              Variance# = Round(QPTrim(BudgetAmt# - GLAcct.Encumb - GLAcct.YTD))
              ToPrint$ = ToPrint$ & "~" & BudgetAmt# & "~" & Str$(GLAcct.Encumb) & "~" & Str$(GLAcct.YTD)
              ToPrint$ = ToPrint$ & "~" & Variance# & "~" & Pct$
            End Select
            Print #PRNFile, QPTrim$(ToPrint$)
          'End If

          If GetMonth Or GetQtr Then
            MTDSum# = MTDSum# + GLAcct.MTD
           ' DeptMTDSum# = DeptMTDSum# + GLAcct.MTD
          End If

          BgtSum# = BgtSum# + BudgetAmt#
          YTDSum# = YTDSum# + GLAcct.YTD
          EncSum# = EncSum# + GLAcct.Encumb

'          DeptBgtSum# = DeptBgtSum# + BudgetAmt#
'          DeptYTDSum# = DeptYTDSum# + GLAcct.YTD
'
'          DeptENCSum# = DeptENCSum# + GLAcct.Encumb
'
'          LastDept$ = Dept$
'          LastDeptName$ = DeptName$

        End If
      Next      'Exp Acct
      End If
      '--Summarize last Dept after loop
      'GoSub PrintDeptTotals

      '--Now summarize all expenditures
      'GoSub PrintSummaryLines   'Print dashed line after last

      '--print total exp for fund
'      If DeptOnNewPage Then
'        'Print #PRNFile, FF$
'        'GoSub PrintBVAPageHeader
'      End If

      Pct$ = GetPct$(YTDSum#, BgtSum#)
      'ToPrint$ = Space$(96)
      'LSet ToPrint$ = "Total Expenditures for Fund:"
'      Select Case WhichReport
'      Case 1, 3
'        Variance# = BgtSum# - YTDSum#
'        Mid$(ToPrint$, Col1) = Using$(BgtFmt$, Str$(BgtSum#))
'        Mid$(ToPrint$, Col2) = Using$(CommaFmt$, Str$(MTDSum#))
'        Mid$(ToPrint$, Col3) = Using$(CommaFmt$, Str$(YTDSum#))
'        Mid$(ToPrint$, Col4) = Using$(CommaFmt$, Str$(Variance#))
'        Mid$(ToPrint$, Col5) = Pct$
'        FundExpMTD# = MTDSum#
'        MTDSum# = 0
'      Case 2
'        Variance# = BgtSum# - EncSum# - YTDSum#
'        Mid$(ToPrint$, Col1) = Using$(BgtFmt$, Str$(BgtSum#))
'        Mid$(ToPrint$, Col2) = Using$(CommaFmt$, Str$(EncSum#))
'        Mid$(ToPrint$, Col3) = Using$(CommaFmt$, Str$(YTDSum#))
'        Mid$(ToPrint$, Col4) = Using$(CommaFmt$, Str$(Variance#))
'        Mid$(ToPrint$, Col5) = Pct$
'      End Select
'      Print #PRNFile, RTrim$(ToPrint$)

      '--Summarize Exp
      FnctExpBgt# = BgtSum#
      FnctExpYTD# = YTDSum#
      FnctEncYTD# = EncSum#

      BgtSum# = 0
      YTDSum# = 0
      EncSum# = 0

      '--Summarize fund
      If GetMonth Or GetQtr Then 'changed
        MTDBal# = Round#(FnctRevMTD# - FnctExpMTD#)
      End If
      BGTBal# = Round#(FnctRevBgt# - FnctExpBgt#)
      YTDBal# = Round#(FnctRevYTD# - FnctExpYTD#)
      EncBal# = Round#(FnctEncYTD#)
'      Print #PRNFile,
      '--print the net
      'ToPrint$ = Space$(96)
      'LSet ToPrint$ = "Revenues Over/(Under) Expenditures"
      'Mid$(ToPrint$, Col1) = Using$(BgtFmt$, Str$(BGTBal#))
'      Select Case WhichReport
'      Case 1, 3
'        Mid$(ToPrint$, Col2) = Using$(CommaFmt$, Str$(MTDBal#))
'        Mid$(ToPrint$, Col3) = Using$(CommaFmt$, Str$(YTDBal#))
'        '--Reset MTD Variables
'        FundRevMTD# = 0
'        FundExpMTD# = 0
'        DeptMTDSum# = 0
'      Case 2
'        Mid$(ToPrint$, Col3) = Using$(CommaFmt$, Str$(YTDBal#))
'      End Select

      'FOR zz = 1 TO 96

     ' Print #PRNFile, RTrim$(ToPrint$)

      '--Blank lines between funds
    '  Print #PRNFile,
    '  Print #PRNFile,
      Linecnt = Linecnt + 2
      '--Reset variables for next fund
      FnctRevBgt# = 0
      FnctRevYTD# = 0
      FnctExpBgt# = 0
      FnctExpYTD# = 0
      FnctEncYTD# = 0
'      DeptBgtSum# = 0
'      DeptYTDSum# = 0
    End If      'Using fund test
  Next     'process next fund
'  Print #PRNFile, FF$
  Close
   Load frmLoadingRpt
'   If DoingDetail Then
    ARptFnctBudVAct.detopt = 1
   ' --if break on depts
'    If fpcboPagebrk.Text = "Yes" Then
'      ARptBudVAct.deptpage = True
'    Else
  If SubTotalRevenues Then
      ARptFnctBudVAct.deptpage = True
  Else
      ARptFnctBudVAct.deptpage = False
  End If
'    End If
'**** Did arptbudvact1 to figure out problem w/original after updating
'     the active reports , but figured out so could use original...
'   ARptBudVAct1.Label4.Caption = lab14
'   ARptBudVAct1.rptnum = NumRpt
'   ARptBudVAct1.labelEnd.Caption = ("Ending Date: " + txtDate)
'   ARptBudVAct1.txtDate = Now
'   ARptBudVAct1.txtTown = GLUserName$
'   ARptBudVAct1.GetName ReportFile$
'   'ARptBudVAct.Visible = False
'    ARptBudVAct1.startrpt
'End If
'  Else
   If DoingRevenues = False Or DoingExp = False Then
    ARptFnctBudVAct.overunder = False
   Else
    ARptFnctBudVAct.overunder = True
   End If

   ARptFnctBudVAct.Label4.Caption = lab14
   ARptFnctBudVAct.rptnum = NumRpt
   ARptFnctBudVAct.labelEnd.Caption = ("Ending Date: " + txtDate)
   ARptFnctBudVAct.txtDate = Now
   ARptFnctBudVAct.txtTown = GLUserName$
   ARptFnctBudVAct.GetName ReportFile$
   'ARptBudVAct.Visible = False
    ARptFnctBudVAct.startrpt
   'Unload frmLoadingRpt
   'ARptBudVAct.Show 1, Me
'  End If
  '====End Report Processing
'  ViewPrint ReportFile$, RptTitle$, True
'  KillFile ReportFile$
  'End Report Printing========================================================
Exit Sub
'PrintBVAPageHeader:
'  PageNum = PageNum + 1
'  Print #PRNFile, GLUserName; Tab(60); "Run Date: " + Date$; "      Page: "; PageNum
'  Print #PRNFile, FundName$ + " " + RptTitle$
'  Print #PRNFile, "Period Ending: " + txtDate
'  Print #PRNFile,
'  Print #PRNFile, DESC$(1)
'  Print #PRNFile, String$(96, "-")
'  Linecnt = 6
'Return
'PrintSummaryLines:
'    '--Print summary lines
'    ToPrint$ = Space$(96)
'    Mid$(ToPrint$, Col1) = BSumLine$
'    Mid$(ToPrint$, (Col2 - 2)) = SumLine$
'    Mid$(ToPrint$, (Col3 - 2)) = SumLine$
'    Mid$(ToPrint$, (Col4 - 2)) = SumLine$
'    Mid$(ToPrint$, Col5) = PSumLine$
'    Print #PRNFile, RTrim$(ToPrint$)
'    Linecnt = Linecnt + 1
'Return
'
'
'PrintDeptTotals:
'  If DoingDetail Then
'    GoSub PrintSummaryLines
'    DeptSummary$ = LastDeptName$ + " Totals"
'  Else
'    DeptSummary$ = LastDept$ + " " + LastDeptName$
'  End If
'
'  'IF INSTR(DeptSummary$, "4910") > 0 THEN STOP
'  'IF INSTR(DeptSummary$, "ZONING") > 0 THEN STOP
'
'  ToPrint$ = Space$(96)
'  LSet ToPrint$ = DeptSummary$
'  Select Case WhichReport
'     Case 1, 3
'       Pct$ = GetPct$(DeptYTDSum#, DeptBgtSum#)
'       Variance# = DeptBgtSum# - DeptYTDSum#
'       Mid$(ToPrint$, Col1) = Using$(BgtFmt$, Str$(DeptBgtSum#))
'       Mid$(ToPrint$, Col2) = Using$(CommaFmt$, Str$(DeptMTDSum#))
'       Mid$(ToPrint$, Col3) = Using$(CommaFmt$, Str$(DeptYTDSum#))
'       Mid$(ToPrint$, (Col4 - 2)) = Using$(TotalFmt$, Str$(Variance#))
'       Mid$(ToPrint$, Col5) = Pct$
'       DeptMTDSum# = 0
'     Case 2
'       Pct$ = GetPct$(DeptYTDSum# + DeptENCSum#, DeptBgtSum#)
'       Variance# = DeptBgtSum# - DeptENCSum# - DeptYTDSum#
'       Mid$(ToPrint$, Col1) = Using$(BgtFmt$, Str$(DeptBgtSum#))
'       Mid$(ToPrint$, Col2) = Using$(CommaFmt$, Str$(DeptENCSum#))
'       Mid$(ToPrint$, Col3) = Using$(CommaFmt$, Str$(DeptYTDSum#))
'       Mid$(ToPrint$, (Col4 - 2)) = Using$(TotalFmt$, Str$(Variance#))
'       Mid$(ToPrint$, Col5) = Pct$
'     Case Else
'  End Select
'  Print #PRNFile, RTrim$(ToPrint$)
'  Linecnt = Linecnt + 1
'  If DoingDetail Then
'    '--formfeed if were at max
'    If Linecnt >= MaxLines Then
'      Print #PRNFile, FF$
'      GoSub PrintBVAPageHeader
'    Else
'      '--print a blank line after totals
'      Print #PRNFile,
'      Linecnt = Linecnt + 1
'    End If
'  End If
'  DeptBgtSum# = 0
'  DeptYTDSum# = 0
'  DeptENCSum# = 0
'Return
'
'
'PrintDeptHeader:
'  If DeptOnNewPage Then
'    Print #PRNFile, FF$
'    GoSub PrintBVAPageHeader
'  End If
'  ToPrint$ = Space$(80)
'  LSet ToPrint$ = DeptName$
'  Print #PRNFile, RTrim$(ToPrint$)
'  Linecnt = Linecnt + 1
'Return
'
'
GotErr:
'  Select Case Err
'''    Case 70
'''      Cls
'''      QPrintRC "Access Denied. Try again later.", 12, 1, 12
'''      QPrintRC "Press any key to continue.", 14, 1, 11
'''    Case Else
'''      Cls
      Unload FrmShowPctComp
      MsgBox "An Error has halted the process, Error Code: " + Str(Err), vbOKOnly, "Error"

'''      QPrintRC "Press any key exit.", 13, 1, 14
'''   End Select
'''
'''   K$ = INPUT$(1)
'''   Exit Sub
'''
'''Return
  Exit Sub
CancelExit:
  Exit Sub
End Sub
Private Sub PrintBgtAct2(rptopt)
  Dim CommaFmt As String, TotalFmt As String, RunBalFmt As String
  Dim SumLine As String, BgtFmt As String, BSumLine As String, PSumLine As String
  Dim DivLine As String, DivLine2 As String, FF As String, RptTitle As String
  Dim MaxLines As Integer, Col1 As Integer, Col2 As Integer, Col3 As Integer
  Dim Col4 As Integer, Col5 As Integer, EndDate As Integer, PRNFile As Integer
  Dim M As String, HollyFlag As Boolean, Pitch12 As String, ThisFnct As String
  Dim DoingDetail As Boolean, SubTotalRevenues As Boolean, DeptOnNewPage As Boolean
  Dim WhichReport As Integer, GetMonth As Boolean, GetQtr As Boolean
  Dim RptMonth As String, ReportFile As String, FnctCode As String
  Dim FundIdxFile As Integer, NumFunds As Integer, Acct As Integer
  Dim FnctIdxFile As Integer, NumFncts As Integer, FnctName As String
  Dim AcctIdxFileNum As Integer, NumGLAccts As Integer, FundName As String
  Dim AcctFileNum As Integer, NumGLAcctRecs As Integer, Rec As Integer
  Dim TransFileNum As Integer, NumTrans As Long, NextTr As Long, AcctNum As String
  Dim BGTBal As Double, YTDBal As Double, MTDBal As Double, UsingFnct As Boolean
  Dim ECnt As Integer, RCnt As Integer, d As String, TransMonth As String
  Dim InThisQtr As Boolean, FirstTime As Boolean, Fnct As Long, FNCTRec As Long
  Dim Dept As String, LastDept As String, LastDeptName As String, cnt As Integer
  Dim Account As String, BudgetAmt As Double, DeptRecNum As Integer, DeptName As String
  Dim Pct As String, Variance As Double, ToPrint As String, Linecnt As Integer
  Dim MTDSum As Double, BgtSum As Double, YTDSum As Double, DeptBgtSum As Double
  Dim DeptYTDSum As Double, DeptENCSum As Double, DeptMTDSum As Double
  Dim FnctRevMTD As Double, FnctRevBgt As Double, FnctRevYTD As Double
  Dim EncSum As Double, FnctExpMTD As Double, FnctExpBgt As Double
  Dim FnctExpYTD As Double, FnctEncYTD As Double, EncBal As Double
  Dim DeptSummary As String, PageNum As Integer, Newrp As String
  Dim GetEnc As Boolean, IncAcct As Boolean
  Dim DoingRevenues As Boolean, DoingExp As Boolean

  Select Case Mid$(fpcboReportOn.Text, 1, 1)
    Case "B"
      DoingRevenues = True
      DoingExp = True
      'rpt = 3
    Case "R"
      DoingRevenues = True
      DoingExp = False
      'rpt = 2
    Case "E"
      DoingRevenues = False
      DoingExp = True
      'rpt = 1
  End Select
'''On Local Error GoTo GotErr

  If InStr(UCase$(GLUserName), "HOLLY SPR") > 0 Then
    HollyFlag = True
    If rptopt = 2 Then Pitch12$ = Chr$(27) + Chr$(38) + Chr$(107) + Chr$(52) + Chr$(83)
  End If


'Do Not Need this Already Have Funds
  'GetFnctCodes FirstFund$, LastFund$
  CommaFmt$ = "###,###,###.##"  'format takes 14 chars
  TotalFmt$ = "#,###,###,###.##" 'format takes 16 chars
  RunBalFmt$ = "##########.##"
  SumLine$ = String$(16, "-")   'column summary line

  BgtFmt$ = "###,###,###"         'format takes 11 chars
  BSumLine$ = String$(11, "-")  'summary line for budget columns
  PSumLine$ = "----"            'summary line for Pct columns
  DivLine$ = String$(115, "-")   'dashed line
  DivLine2$ = String$(115, "=")  'Double Line
  FF$ = Chr$(12)
  ReDim Desc$(1)
  RptTitle$ = "Function Budget vs. Actual "
  MaxLines = 55

  '--Column offsets for printing amounts
  Col1 = 40    'Budget
  Col2 = 54   'Month or Actual YTD
  Col3 = 71    'YTD or Var
  Col4 = 88    'Enc
  Col5 = 105    'Pct
  
  EndDate = DateDiff("d", "12/31/1979", txtDate)
  M$ = Right$(txtDate, 2) + Left$(txtDate, 2)
'Format(DateAdd("d", FY2BegDate, "12-31-1979"), "mm/dd/yy")
  FrmShowPctComp.Label1 = "Printing Function Budget vs Actual Report"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmPrnFnctBudvAct
  If fpcboIncAcct.ListIndex = 1 Then
    IncAcct = False
  Else
    IncAcct = True
  End If

  'If fpcboDetSum.Text = "Detail" Then DoingDetail = True

  If fpcboSubtot.Text = "Yes" Then SubTotalRevenues = True
  '  SubTotalRevenues = True
  'ELSE
  '  SubTotalRevenues = False
  'END IF

' If DoingDetail Then
    '--if break on depts
'    If fpcboPagebrk.Text = "Yes" Then
'      DeptOnNewPage = True
'    End If
 ' End If
  fpcboRepNum.Col = 0
  WhichReport = Val(fpcboRepNum.Text)
  Select Case WhichReport
  Case 1        'Bgt, Month, YTD
    GetMonth = True
    GetQtr = False
    GetEnc = False
    RptMonth = M$
''''''''''''''''123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456''''For 96
    Desc$(1) = "Description                                Budget          Month             YTD           Variance    Pct"
    MTDBal# = 0
  Case 2        'Bgt, Enc, YTD, Variance
    GetMonth = False
    GetQtr = False
    GetEnc = True
    Desc$(1) = "Description                                Budget          Encumb            YTD           Variance    Pct"
  Case 3        'Bgt, Enc, YTD, Variance
    GetMonth = False
    GetQtr = True
    GetEnc = False
    Desc$(1) = "Description                                Budget           QTD              YTD           Variance    Pct"
  End Select
  If GetEnc = True Then
        'if need enc totals do this here
    FixPOEncumbRpt EndDate, FYStartDate
  End If

  MTDBal# = 0
  Newrp = "BGTACT"
  GetRPTName Newrp
  ReportFile$ = Newrp

  'ReportFile$ = Unique$(Path$)
  'P17$ = CHR$(27) + "(s17H"
  PRNFile = FreeFile
  Open ReportFile$ For Output As #PRNFile

  'PRINT #PrnFile, P17$;

  If HollyFlag Then
    If rptopt = 2 Then Print #PRNFile, Pitch12$;
  End If

  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum
  NumGLAcctRecs = LOF(AcctFileNum) / Len(GLAcct)
  OpenTransFile TransFileNum, NumTrans&
  OpenFundIdx FundIdxFile, NumFunds
  OpenFnctIdx FnctIdxFile, NumFncts
  
  ReDim RevAccts%(1 To NumGLAccts)              'Holds all rev acct record num
  ReDim ExpAccts%(1 To NumGLAccts)              'Holds all exp acct record num
'  ReDim FundList(1) As String                           'List of all active Funds
'  GetFundList FundList$(), NumFunds
  ReDim FnctList(1 To NumFncts) As String                           'List of all active Funds
  GetFnctList FnctList$(), NumFncts

  '--Build a list of revenue and exp accounts
  For Acct = 1 To NumGLAccts
    '--Initialize
    BGTBal# = 0
    YTDBal# = 0
    MTDBal# = 0

    Get AcctIdxFileNum, Acct, GLAcctidx
    Get AcctFileNum, GLAcctidx.RecNum, GLAcct

    '--Find what fund this account is in
    If GLAcct.FNCTRec > 0 Then
      FnctCode$ = GetFnctCode(GLAcct.FNCTRec) 'Left$(GLAcct.Num, GLFundLen)

    '--See if the account is in a fund we want to see
    If FnctCode$ >= StartFnct$ And FnctCode$ <= EndFnct$ Then

      '--Account is in fund, check to see if its proper type
      '--We want only revenue or expenditure accounts
      If GLAcct.Typ = "R" Or GLAcct.Typ = "E" Then

        '--Assign the Account Record Number to proper list
        Select Case GLAcct.Typ
        Case "E"
          ECnt = ECnt + 1
          ExpAccts%(ECnt) = GLAcctidx.RecNum

        Case "R"
          RCnt = RCnt + 1
          RevAccts%(RCnt) = GLAcctidx.RecNum
        End Select

        '--Get account balances
        '--There should be no beginning balances in rev & exp accts
        'YTDBal# = Round#(Acct.BegBal)           'get the beginning balance

        NextTr& = GLAcct.FrstTran 'get the first trans for this acct

        Do Until NextTr& = 0    'keep going 'til we run out

          Get TransFileNum, NextTr&, GLTrans

          '--Get MTD Account Balance if necessary
          If GLTrans.TRDATE >= FYStartDate And GLTrans.TRDATE <= EndDate Then

          If GetMonth Then
            'Lookhere change num2month to reflect year & month
            d$ = Format(DateAdd("d", GLTrans.TRDATE, "12-31-1979"), "mm/dd/yyyy")
            TransMonth = Right$(d$, 2) + Left$(d$, 2)
            'TransMonth = Num2Month%(GLTrans.TRDATE)
            If TransMonth = RptMonth Then
              Select Case GLAcct.Typ
              Case "E"
                MTDBal# = MTDBal# + Round#(GLTrans.DrAmt) - Round#(GLTrans.CrAmt)
              Case "R"
                MTDBal# = MTDBal# + Round#(GLTrans.CrAmt) - Round#(GLTrans.DrAmt)
              End Select
            End If
          End If

          If GetQtr Then
            InThisQtr = InQtr(GLTrans.TRDATE, EndDate)
            If InThisQtr Then
              Select Case GLAcct.Typ
                Case "E"
                  MTDBal# = MTDBal# + Round#(GLTrans.DrAmt) - Round#(GLTrans.CrAmt)
                Case "R"
                  MTDBal# = MTDBal# + Round#(GLTrans.CrAmt) - Round#(GLTrans.DrAmt)
              End Select
            End If
          End If

          '--Get YTD Account Balance
          '--original
          '--Carthage 9/17/96
          'IF Trans.TrDate <= EndDate THEN
          '--08/07/96 keeping funds open --
          '--Does'nt work when 2 years are open...what to do?
          'IF Trans.TrDate <= EndDate THEN
            Select Case GLAcct.Typ
            Case "E"
              YTDBal# = YTDBal# + Round#(GLTrans.DrAmt) - Round#(GLTrans.CrAmt)
            Case "R"
              YTDBal# = YTDBal# + Round#(GLTrans.CrAmt) - Round#(GLTrans.DrAmt)
            End Select
          End If

          NextTr& = GLTrans.NextTran              'Get the next transaction

        Loop
        '--Put the new totals in the file
        GLAcct.MTD = Round#(MTDBal#)
        GLAcct.YTD = Round#(YTDBal#)
        Put AcctFileNum, GLAcctidx.RecNum, GLAcct

      End If    '--test for rev or exp accts
    End If      '--End of acct in fund range test
    End If      '--glacct function rec no > 0
    FrmShowPctComp.ShowPctComp Acct, NumGLAccts
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnFnctBudvAct
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
  Next          'Process next account

  ActivateControls frmPrnFnctBudvAct
 

  '--Now write the report to file.
'''''  PrintHelp "Generating Report..."

  FirstTime = True 'flag not to print page feed first time thru

  '--Process each function
  For Fnct = 1 To NumFncts
    ThisFnct$ = FnctList$(Fnct)
    If ThisFnct$ >= StartFnct$ And ThisFnct$ <= EndFnct$ Then
      UsingFnct = True
      FNCTRec = FindFnct(ThisFnct$)             'Get the fund name
      FnctName$ = QPTrim$(GetFnctTitle(FNCTRec))
      
     'Dept$ = ""

     ' LastDept$ = ""
     ' LastDeptName$ = ""
    Else
      UsingFnct = False
    End If

    If UsingFnct Then
      '--print a form feed for each new function
      If FirstTime Then
        FirstTime = False
      Else
        Print #PRNFile, FF$
      End If

      'RE$ = "Revenues" OrNot DoingDetail
      'If Not SubTotalRevenues Then
        GoSub PrintBVAPageHeader
       
      'End If
      If DoingRevenues = True Then
      '--Search thru list of revenue accounts
      For cnt = 1 To RCnt
        Rec = RevAccts%(cnt)
        Get AcctFileNum, Rec, GLAcct
        FnctCode$ = GetFnctCode(GLAcct.FNCTRec) 'Left$(GLAcct.Num, GLFundLen)

        If FnctCode$ = ThisFnct$ Then
          If IncAcct = True Then
            Account$ = QPTrim$(GLAcct.Num) + "  " + QPTrim$(GLAcct.Title)
          Else
            Account$ = QPTrim$(GLAcct.Title)
          End If
          Select Case ActiveYear
          Case 1
            BudgetAmt# = GLAcct.Bgt
          Case 2
            BudgetAmt# = GLAcct.NYApp
          End Select

          '----------
'          If SubTotalRevenues Then
            '--Extract the Dept$ from the G/L Acct
'            Dept$ = Mid$(GLAcct.Num, GLFundLen + 2, GLAcctLen)
'
'            '--if a new dept, get its name from the Dept name file
'            If Dept$ <> LastDept$ Then
'              DeptRecNum = FindDept(Dept$)
'              If DeptRecNum > 0 Then
'                DeptName$ = QPTrim$(GetDeptTitle$(DeptRecNum))
'              Else
'                'DeptName$ = "Department " + Dept$
'                DeptName$ = " "
'              End If
'            End If
'
'            '--Print Department Header first time thru
'            If Len(LastDeptName$) = 0 Then
'              '--if we're not printing departments on separate pages then
'              '--print a new page header
'              If DeptOnNewPage = False Then
               ' If DoingDetail = True Then
'                If rptopt = 2 Then Print #PRNFile, FF$
'                GoSub PrintBVAPageHeader
                'Else
'                  Print #PRNFile,
'                  Print #PRNFile, "Revenues"
'                End If
'              End If
'              LastDeptName$ = DeptName$
'              LastDept$ = Dept$
             ' If DoingDetail Then

                'GoSub PrintDeptHeader
             ' End If
'            End If

            '--see if we need to subtotal dept
'            If Len(LastDept$) > 0 Then
'              If Dept$ <> LastDept$ Then
'                GoSub PrintDeptTotals
'                If DoingDetail Then
'                  GoSub PrintDeptHeader
'                End If
'              End If
'
'            End If

'          End If
          '===========


          Pct$ = GetPct$(GLAcct.YTD, BudgetAmt#)

      '    If DoingDetail Then
            Variance# = GLAcct.YTD - BudgetAmt# 'Acct.Bgt
            ToPrint$ = Space$(115)
            LSet ToPrint$ = Left$(Account$, 28)
            Mid$(ToPrint$, Col1) = Using$(BgtFmt$, Str$(BudgetAmt#))   'chang
            Select Case WhichReport
            Case 1, 3
              Mid$(ToPrint$, Col2) = Using$(CommaFmt$, Str$(GLAcct.MTD))
              Mid$(ToPrint$, Col3) = Using$(CommaFmt$, Str$(GLAcct.YTD))
              Mid$(ToPrint$, Col4) = Using$(CommaFmt$, Str$(Variance#))
              Mid$(ToPrint$, Col5) = Pct$
            Case 2
              Mid$(ToPrint$, Col2) = Using$(CommaFmt$, Str$(GLAcct.Encumb))
              Mid$(ToPrint$, Col3) = Using$(CommaFmt$, Str$(GLAcct.YTD))
              Mid$(ToPrint$, Col4) = Using$(CommaFmt$, Str$(Variance#))
              Mid$(ToPrint$, Col5) = Pct$
            End Select

            Print #PRNFile, RTrim$(ToPrint$)
           
            Linecnt = Linecnt + 1
            If Linecnt > MaxLines Then
              Print #PRNFile, FF$
              GoSub PrintBVAPageHeader
            End If
        '  End If

          'changed
          If GetMonth Or GetQtr Then
            MTDSum# = MTDSum# + GLAcct.MTD
          End If

          BgtSum# = BgtSum# + BudgetAmt# 'Acct.Bgt
          YTDSum# = YTDSum# + GLAcct.YTD

'          DeptBgtSum# = DeptBgtSum# + BudgetAmt#
'          DeptYTDSum# = DeptYTDSum# + GLAcct.YTD
'          DeptENCSum# = DeptENCSum# + GLAcct.Encumb
'          DeptMTDSum# = DeptMTDSum# + GLAcct.MTD
'
'          LastDept$ = Dept$
'          LastDeptName$ = DeptName$

        End If
      Next      'Revenue Acct

      '--Summarize Revenues
     ' If DoingDetail Then
        GoSub PrintSummaryLines
     ' End If

      ToPrint$ = Space$(115)
      LSet ToPrint$ = "Total Revenues for Function: " & ThisFnct$
      Pct$ = GetPct$(YTDSum#, BgtSum#)
      Variance# = YTDSum# - BgtSum#
      Select Case WhichReport
      Case 1, 3
        Variance# = YTDSum# - BgtSum#
        Mid$(ToPrint$, Col1) = Using$(BgtFmt$, Str$(BgtSum#))
        Mid$(ToPrint$, (Col2 - 2)) = Using$(TotalFmt$, Str$(MTDSum#))
        Mid$(ToPrint$, (Col3 - 2)) = Using$(TotalFmt$, Str$(YTDSum#))
        Mid$(ToPrint$, Col4) = Using$(CommaFmt$, Str$(Variance#))
        Mid$(ToPrint$, Col5) = Pct$
        '--Reset vars
        FnctRevMTD# = MTDSum#
        MTDSum# = 0
      Case 2
        Variance# = YTDSum# - BgtSum#
        Mid$(ToPrint$, Col1) = Using$(BgtFmt$, Str$(BgtSum#))
        Mid$(ToPrint$, (Col2 - 2)) = Using$(TotalFmt$, " 0")
        Mid$(ToPrint$, (Col3 - 2)) = Using$(TotalFmt$, Str$(YTDSum#))
        Mid$(ToPrint$, Col4) = Using$(CommaFmt$, Str$(Variance#))
        Mid$(ToPrint$, Col5) = Pct$
      End Select
      Print #PRNFile, RTrim$(ToPrint$)
      Print #PRNFile, ""
      Linecnt = Linecnt + 2
      
      If Linecnt > MaxLines Then
        Print #PRNFile, FF$
        GoSub PrintBVAPageHeader
      Else
        If SubTotalRevenues Then
          Print #PRNFile, FF$
          GoSub PrintBVAPageHeader
        End If
      End If
      End If
      FnctRevBgt# = BgtSum#
      FnctRevYTD# = YTDSum#
      EncSum# = 0
      BgtSum# = 0
      YTDSum# = 0

      'initialize dept variables
'      DeptBgtSum# = 0
'      DeptMTDSum# = 0
'      DeptYTDSum# = 0
'      DeptENCSum# = 0
'      LastDept$ = ""
'      LastDeptName$ = ""

      '--Search exp accounts list for accounts in this fund
      'RE$ = "Expenditures"
      If DoingExp = True Then
      For cnt = 1 To ECnt
        Rec = ExpAccts%(cnt)
        Get AcctFileNum, Rec, GLAcct

        FnctCode$ = GetFnctCode(GLAcct.FNCTRec) 'Left$(GLAcct.Num, GLFundLen)
        If FnctCode$ = ThisFnct$ Then
          If IncAcct = True Then
            Account$ = QPTrim$(GLAcct.Num) + "  " + QPTrim$(GLAcct.Title)
          Else
            Account$ = QPTrim$(GLAcct.Title)
          End If

          Select Case ActiveYear
          Case 1
            BudgetAmt# = GLAcct.Bgt
          Case 2
            BudgetAmt# = GLAcct.NYApp
          End Select

          '--Extract the Dept$ from the G/L Acct
'          Dept$ = Mid$(GLAcct.Num, GLFundLen + 2, GLAcctLen)

          '--if a new dept, get its name from the Dept name file
'          If Dept$ <> LastDept$ Then
'            DeptRecNum = FindDept(Dept$)
'            If DeptRecNum > 0 Then
'              DeptName$ = QPTrim$(GetDeptTitle$(DeptRecNum))
'            Else
'              'DeptName$ = "Department " + Dept$
'              DeptName$ = " "
'            End If
'          End If
'
          '--Print Department Header first time thru
'          If Len(LastDeptName$) = 0 Then
'            '--if we're not printing departments on separate pages then
'            '--print a new page header
'            If DeptOnNewPage = False Then
'              If DoingDetail = True Then
'                Print #PRNFile, FF$
'                GoSub PrintBVAPageHeader
'              Else
'                Print #PRNFile,
'                Print #PRNFile, "Expenditures"
'              End If
'            End If
''            LastDeptName$ = DeptName$
''            LastDept$ = Dept$
''            If DoingDetail Then
''              GoSub PrintDeptHeader
''            End If
'          End If

          '--see if we need to subtotal dept
'          If Len(LastDept$) > 0 Then
'            If Dept$ <> LastDept$ Then
'              GoSub PrintDeptTotals
'              If DoingDetail Then
'                GoSub PrintDeptHeader
'              End If
'            End If
'          End If

          'If DoingDetail Then   'Print Account Detail
            ToPrint$ = Space$(115)

            LSet ToPrint$ = Left$(Account$, 36)

            Select Case WhichReport
            Case 1, 3
              Pct$ = GetPct$(GLAcct.YTD, BudgetAmt#) 'Acct.Bgt
              Variance# = QPTrim(BudgetAmt# - GLAcct.YTD)
              Mid$(ToPrint$, Col1) = Using$(BgtFmt$, Str$(BudgetAmt#))
              Mid$(ToPrint$, Col2) = Using$(CommaFmt$, Str$(GLAcct.MTD))
              Mid$(ToPrint$, Col3) = Using$(CommaFmt$, Str$(GLAcct.YTD))
              Mid$(ToPrint$, Col4) = Using$(CommaFmt$, Str$(Variance#))
              Mid$(ToPrint$, Col5) = Pct$
            Case 2
              Pct$ = GetPct$(GLAcct.Encumb + GLAcct.YTD, BudgetAmt#) 'Acct.Bgt
              Variance# = QPTrim(BudgetAmt# - GLAcct.Encumb - GLAcct.YTD)
              Mid$(ToPrint$, Col1) = Using$(BgtFmt$, Str$(BudgetAmt#))
              Mid$(ToPrint$, Col2) = Using$(CommaFmt$, Str$(GLAcct.Encumb))
              Mid$(ToPrint$, Col3) = Using$(CommaFmt$, Str$(GLAcct.YTD))
              Mid$(ToPrint$, Col4) = Using$(CommaFmt$, Str$(Variance#))
              Mid$(ToPrint$, Col5) = Pct$
            End Select
            Print #PRNFile, RTrim$(ToPrint$)
            Linecnt = Linecnt + 1
            If Linecnt > MaxLines Then
              Print #PRNFile, FF$
              GoSub PrintBVAPageHeader
            End If
         ' End If

          If GetMonth Or GetQtr Then
            MTDSum# = MTDSum# + GLAcct.MTD
           ' DeptMTDSum# = DeptMTDSum# + GLAcct.MTD
          End If

          BgtSum# = BgtSum# + BudgetAmt#
          YTDSum# = YTDSum# + GLAcct.YTD
          EncSum# = EncSum# + GLAcct.Encumb

'          DeptBgtSum# = DeptBgtSum# + BudgetAmt#
'          DeptYTDSum# = DeptYTDSum# + GLAcct.YTD
'
'          DeptENCSum# = DeptENCSum# + GLAcct.Encumb
'
'          LastDept$ = Dept$
'          LastDeptName$ = DeptName$
'
        End If
      Next      'Exp Acct

      '--Summarize last Dept after loop
     ' GoSub PrintDeptTotals

      '--Now summarize all expenditures
      GoSub PrintSummaryLines   'Print dashed line after last

      '--print total exp for fund
'      If DeptOnNewPage Then
'        Print #PRNFile, FF$
'        GoSub PrintBVAPageHeader
'      End If

      Pct$ = GetPct$(YTDSum#, BgtSum#)
      ToPrint$ = Space$(115)
      LSet ToPrint$ = "Total Expenditures for Function: " & ThisFnct$
      Select Case WhichReport
      Case 1, 3
        Variance# = BgtSum# - YTDSum#
        Mid$(ToPrint$, Col1) = Using$(BgtFmt$, Str$(BgtSum#))
        Mid$(ToPrint$, Col2) = Using$(CommaFmt$, Str$(MTDSum#))
        Mid$(ToPrint$, Col3) = Using$(CommaFmt$, Str$(YTDSum#))
        Mid$(ToPrint$, Col4) = Using$(CommaFmt$, Str$(Variance#))
        Mid$(ToPrint$, Col5) = Pct$
        FnctExpMTD# = MTDSum#
        MTDSum# = 0
      Case 2
        Variance# = BgtSum# - EncSum# - YTDSum#
        Mid$(ToPrint$, Col1) = Using$(BgtFmt$, Str$(BgtSum#))
        Mid$(ToPrint$, Col2) = Using$(CommaFmt$, Str$(EncSum#))
        Mid$(ToPrint$, Col3) = Using$(CommaFmt$, Str$(YTDSum#))
        Mid$(ToPrint$, Col4) = Using$(CommaFmt$, Str$(Variance#))
        Mid$(ToPrint$, Col5) = Pct$
      End Select
      Print #PRNFile, RTrim$(ToPrint$)
      End If
      '--Summarize Exp
      FnctExpBgt# = BgtSum#
      FnctExpYTD# = YTDSum#
      FnctEncYTD# = EncSum#

      BgtSum# = 0
      YTDSum# = 0
      EncSum# = 0

      '--Summarize function
      If GetMonth Or GetQtr Then 'changed
        MTDBal# = Round#(FnctRevMTD# - FnctExpMTD#)
      End If
      BGTBal# = Round#(FnctRevBgt# - FnctExpBgt#)
      YTDBal# = Round#(FnctRevYTD# - FnctExpYTD#)
      EncBal# = Round#(FnctEncYTD#)
      Print #PRNFile,
      '--print the net
      If DoingRevenues And DoingExp Then

      ToPrint$ = Space$(115)
      LSet ToPrint$ = "Revenues Over/(Under) Expenditures"
      'Mid$(ToPrint$, Col1) = Using$(BgtFmt$, Str$(BGTBal#))
      Select Case WhichReport
      Case 1, 3
        Mid$(ToPrint$, Col2) = Using$(CommaFmt$, Str$(MTDBal#))
        Mid$(ToPrint$, Col3) = Using$(CommaFmt$, Str$(YTDBal#))
        '--Reset MTD Variables
        FnctRevMTD# = 0
        FnctExpMTD# = 0
        DeptMTDSum# = 0
      Case 2
        Mid$(ToPrint$, Col3) = Using$(CommaFmt$, Str$(YTDBal#))
      End Select

      'FOR zz = 1 TO 96
        Print #PRNFile, RTrim$(ToPrint$)
      End If
      '--Blank lines between funds
      Print #PRNFile,
      Print #PRNFile,
      Linecnt = Linecnt + 2
      
      '--Reset variables for next fund
      FnctRevBgt# = 0
      FnctRevYTD# = 0
      FnctExpBgt# = 0
      FnctExpYTD# = 0
      FnctEncYTD# = 0
      DeptBgtSum# = 0
      DeptYTDSum# = 0
    End If      'Using fund test
  Next     'process next fund
  If rptopt = 2 Then Print #PRNFile, FF$
  Close
  '====End Report Processing
  If rptopt = 2 Then
    ViewPrint ReportFile$, RptTitle$, True
    KillFile ReportFile$
  Else
'    Load frmLoadingRpt
'    ARptLineRptLand.GetName ReportFile$
'    ARptLineRptLand.startrpt
  End If
  'End Report Printing========================================================
Exit Sub
PrintBVAPageHeader:
  PageNum = PageNum + 1
  If rptopt = 1 Then
    Print #PRNFile, ""
    Print #PRNFile, ""
    Print #PRNFile, ""
  End If
  Print #PRNFile, GLUserName; Tab(60); "Run Date: " + Date$; Tab(103); "      Page: "; PageNum
  Print #PRNFile, FnctName$ + " " + RptTitle$
  Print #PRNFile, "Period Ending: " + txtDate
  Print #PRNFile,
  Print #PRNFile, Desc$(1)
  Print #PRNFile, String$(115, "-")
  Linecnt = 6
  If rptopt = 1 Then
    Linecnt = Linecnt + 3
  End If
Return
PrintSummaryLines:
    '--Print summary lines
    ToPrint$ = Space$(115)
    Mid$(ToPrint$, Col1) = BSumLine$
    Mid$(ToPrint$, (Col2 - 2)) = SumLine$
    Mid$(ToPrint$, (Col3 - 2)) = SumLine$
    Mid$(ToPrint$, (Col4 - 2)) = SumLine$
    Mid$(ToPrint$, Col5) = PSumLine$
    Print #PRNFile, RTrim$(ToPrint$)
    Linecnt = Linecnt + 1
Return


'PrintDeptTotals:
'  If DoingDetail Then
'    GoSub PrintSummaryLines
'    DeptSummary$ = LastDeptName$ + " Totals"
'  Else
'    DeptSummary$ = LastDept$ + " " + LastDeptName$
'  End If
'
'  'IF INSTR(DeptSummary$, "4910") > 0 THEN STOP
'  'IF INSTR(DeptSummary$, "ZONING") > 0 THEN STOP
'
'  ToPrint$ = Space$(96)
'  LSet ToPrint$ = DeptSummary$
'  Select Case WhichReport
'     Case 1, 3
'       Pct$ = GetPct$(DeptYTDSum#, DeptBgtSum#)
'       Variance# = DeptBgtSum# - DeptYTDSum#
'       Mid$(ToPrint$, Col1) = Using$(BgtFmt$, Str$(DeptBgtSum#))
'       Mid$(ToPrint$, Col2) = Using$(CommaFmt$, Str$(DeptMTDSum#))
'       Mid$(ToPrint$, Col3) = Using$(CommaFmt$, Str$(DeptYTDSum#))
'       Mid$(ToPrint$, (Col4 - 2)) = Using$(TotalFmt$, Str$(Variance#))
'       Mid$(ToPrint$, Col5) = Pct$
'       DeptMTDSum# = 0
'     Case 2
'       Pct$ = GetPct$(DeptYTDSum# + DeptENCSum#, DeptBgtSum#)
'       Variance# = DeptBgtSum# - DeptENCSum# - DeptYTDSum#
'       Mid$(ToPrint$, Col1) = Using$(BgtFmt$, Str$(DeptBgtSum#))
'       Mid$(ToPrint$, Col2) = Using$(CommaFmt$, Str$(DeptENCSum#))
'       Mid$(ToPrint$, Col3) = Using$(CommaFmt$, Str$(DeptYTDSum#))
'       Mid$(ToPrint$, (Col4 - 2)) = Using$(TotalFmt$, Str$(Variance#))
'       Mid$(ToPrint$, Col5) = Pct$
'     Case Else
'  End Select
'  Print #PRNFile, RTrim$(ToPrint$)
'  Linecnt = Linecnt + 1
'  If DoingDetail Then
'    '--formfeed if were at max
'    If Linecnt >= MaxLines Then
'      Print #PRNFile, FF$
'      GoSub PrintBVAPageHeader
'    Else
'      '--print a blank line after totals
'      Print #PRNFile,
'      Linecnt = Linecnt + 1
'    End If
'  End If
'  DeptBgtSum# = 0
'  DeptYTDSum# = 0
'  DeptENCSum# = 0
'Return


'PrintDeptHeader:
'  If DeptOnNewPage Then
'    Print #PRNFile, FF$
'    GoSub PrintBVAPageHeader
'  End If
'  ToPrint$ = Space$(80)
'  LSet ToPrint$ = DeptName$
'  Print #PRNFile, RTrim$(ToPrint$)
'  Linecnt = Linecnt + 1
'Return


'''GotErr:
'''  ErrorCode$ = Str$(Err)
'''  Select Case Err
'''    Case 70
'''      Cls
'''      QPrintRC "Access Denied. Try again later.", 12, 1, 12
'''      QPrintRC "Press any key to continue.", 14, 1, 11
'''    Case Else
'''      Cls
'''      QPrintRC "An Error has halted the system, Error Code: " + ErrorCode$, 12
'''      QPrintRC "Press any key exit.", 13, 1, 14
'''   End Select
'''
'''   K$ = INPUT$(1)
'''   Exit Sub
'''
'''Return

CancelExit:
  Exit Sub
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboFunction1.SetFocus
  End If
End Sub
