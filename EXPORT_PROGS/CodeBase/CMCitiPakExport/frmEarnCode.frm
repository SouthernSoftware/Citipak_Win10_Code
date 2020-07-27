VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmEarningsCodeMaint 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Additional Earnings Codes"
   ClientHeight    =   8880
   ClientLeft      =   30
   ClientTop       =   465
   ClientWidth     =   11655
   Icon            =   "frmEarnCode.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8654.619
   ScaleMode       =   0  'User
   ScaleWidth      =   11620
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcomboRET 
      Height          =   405
      Left            =   5670
      TabIndex        =   6
      ToolTipText     =   "Set to 'N' to EXCLUDE Retirement from earnings. Set to ""Y"" to INCLUDE Retirement in earnings."
      Top             =   3450
      Width           =   540
      _Version        =   196608
      _ExtentX        =   952
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
      ColDesigner     =   "frmEarnCode.frx":08CA
   End
   Begin LpLib.fpCombo fpcomboMED 
      Height          =   405
      Left            =   4605
      TabIndex        =   5
      ToolTipText     =   "Set to 'N' to EXCLUDE Medicare from earnings. Set to ""Y"" to INCLUDE Medicare in earnings."
      Top             =   3450
      Width           =   540
      _Version        =   196608
      _ExtentX        =   952
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
      ColDesigner     =   "frmEarnCode.frx":0BC1
   End
   Begin LpLib.fpCombo fpcomboSOC 
      Height          =   405
      Left            =   3510
      TabIndex        =   4
      ToolTipText     =   "Set to 'N' to EXCLUDE Social Security from earnings. Set to ""Y"" to INCLUDE Social Security in earnings."
      Top             =   3450
      Width           =   540
      _Version        =   196608
      _ExtentX        =   952
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
      ColDesigner     =   "frmEarnCode.frx":0EB8
   End
   Begin LpLib.fpCombo fpcomboSWT 
      Height          =   405
      Left            =   2400
      TabIndex        =   3
      ToolTipText     =   "Set to 'N' to EXCLUDE State Withholdings from earnings. Set to ""Y"" to INCLUDE State Withholdings in earnings."
      Top             =   3450
      Width           =   540
      _Version        =   196608
      _ExtentX        =   952
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
      ColDesigner     =   "frmEarnCode.frx":11AF
   End
   Begin LpLib.fpCombo fpcomboFWT 
      Height          =   405
      Left            =   1290
      TabIndex        =   2
      ToolTipText     =   "Set to 'N' to EXCLUDE Federal Withholdings from earnings. Set to ""Y"" to INCLUDE Federal Withholdings in earnings."
      Top             =   3450
      Width           =   540
      _Version        =   196608
      _ExtentX        =   952
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
      ColDesigner     =   "frmEarnCode.frx":14A6
   End
   Begin LpLib.fpCombo fpcmb401KXN 
      Height          =   405
      Left            =   7875
      TabIndex        =   7
      ToolTipText     =   $"frmEarnCode.frx":179D
      Top             =   3450
      Width           =   540
      _Version        =   196608
      _ExtentX        =   952
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
      ColDesigner     =   "frmEarnCode.frx":182B
   End
   Begin EditLib.fpText fpDescription 
      Height          =   372
      Left            =   2640
      TabIndex        =   1
      ToolTipText     =   "Enter a description for this earnings code."
      Top             =   2928
      Width           =   5532
      _Version        =   196608
      _ExtentX        =   9758
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
      ButtonStyle     =   0
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
   Begin FPSpread.vaSpread vaSpreadEarningsCodes 
      Height          =   1305
      Left            =   1740
      TabIndex        =   8
      ToolTipText     =   "Double click any row to bring it up in the Data Entry section."
      Top             =   5955
      Width           =   8265
      _Version        =   196613
      _ExtentX        =   14579
      _ExtentY        =   2302
      _StockProps     =   64
      AutoSize        =   -1  'True
      ButtonDrawMode  =   4
      ColsFrozen      =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   7
      MaxRows         =   3
      ProcessTab      =   -1  'True
      ScrollBars      =   2
      ShadowColor     =   13684944
      SpreadDesigner  =   "frmEarnCode.frx":1B22
      StartingColNumber=   0
      VisibleCols     =   7
      VisibleRows     =   3
      ScrollBarTrack  =   1
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   615
      Left            =   8927
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Press ESC to exit this screen."
      Top             =   2599
      Width           =   2055
      _Version        =   131072
      _ExtentX        =   3625
      _ExtentY        =   1085
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
      ButtonDesigner  =   "frmEarnCode.frx":208E
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdClear 
      Height          =   375
      Left            =   3120
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Press to delete data from the fields above."
      Top             =   4200
      Width           =   3030
      _Version        =   131072
      _ExtentX        =   5345
      _ExtentY        =   661
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
      ButtonDesigner  =   "frmEarnCode.frx":22A2
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSaveContinue 
      Height          =   624
      Left            =   8928
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Press F10 to save the last entry but leave the screen open to allow further editing."
      Top             =   3276
      Width           =   2052
      _Version        =   131072
      _ExtentX        =   3619
      _ExtentY        =   1101
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
      ButtonDesigner  =   "frmEarnCode.frx":248C
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSaveExit 
      Height          =   615
      Left            =   8927
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Press F11 to exit this screen after saving all data on this screen."
      Top             =   3970
      Width           =   2055
      _Version        =   131072
      _ExtentX        =   3625
      _ExtentY        =   1085
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
      ButtonDesigner  =   "frmEarnCode.frx":26AE
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "MATCH 401K"
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
      Left            =   6384
      TabIndex        =   17
      Top             =   3552
      Width           =   1452
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "RET"
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
      Left            =   5184
      TabIndex        =   16
      Top             =   3552
      Width           =   492
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   1092
      Index           =   1
      Left            =   1536
      Top             =   696
      Width           =   8652
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Additional Earnings Codes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2760
      TabIndex        =   13
      Top             =   1056
      Width           =   6012
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Data Entry"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   2592
      TabIndex        =   0
      Top             =   2328
      Width           =   3732
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "FWT"
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
      Left            =   768
      TabIndex        =   15
      Top             =   3552
      Width           =   492
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
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
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   960
      TabIndex        =   14
      Top             =   3024
      Width           =   1572
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1212
      Left            =   1536
      Top             =   576
      Width           =   8652
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Withholding on Earnings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   3984
      TabIndex        =   9
      Top             =   5328
      Width           =   3732
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Height          =   372
      Left            =   3984
      Top             =   5328
      Width           =   3732
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2292
      Left            =   624
      Top             =   2448
      Width           =   10644
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "SWT"
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
      Left            =   1872
      TabIndex        =   12
      Top             =   3552
      Width           =   492
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "SOC"
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
      Left            =   2976
      TabIndex        =   11
      Top             =   3552
      Width           =   492
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H008F8265&
      Caption         =   "MED"
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
      Left            =   4080
      TabIndex        =   10
      Top             =   3552
      Width           =   492
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   8616.048
      X2              =   8616.048
      Y1              =   2401.462
      Y2              =   4606.051
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
Attribute VB_Name = "frmEarningsCodeMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'the deduction codes and the earnings codes control screens
'operate very close to the same way and there are more
'comments in the deduction code code
Option Explicit
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim ClickFlag As Boolean
  Dim RowFlag As Integer
  Dim PriorRowNum As Integer
  Dim ContinueFlag As Boolean
  Dim SaveAndExitFlag As Boolean
  Dim BadDataFlag As Boolean
  Dim NoDataFlag As Boolean
  Dim changeFlag As Boolean
  Dim PriorDesc As String
  Dim PriorLibNum As String
  Dim PriorFWT As String
  Dim PriorSWT As String
  Dim PriorSOC As String
  Dim ClearFieldsFlag As Boolean
  Dim PriorMED As String
  Dim PriorRET As String
  Dim Prior401K As String
  Dim FirstTimeThru As Boolean
  Dim DupFlag As Boolean
  Dim JustExitFlag As Boolean

Private Sub cmdClear_Click()
'This sub was designed to allow the user to clear all fields
'after having been editing existing entries...if the user
'has been editing then in order to enter a brand new entry
'they would have to highlight the last empty row and then
'enter new data...this way the program knows this is a new entry
'and automatically saves new data in the next empty row
  
'before entering new data we must check to see if the last
'entry was saved properly and if not give the user the option
'to save or not to save
  Dim ErnCodeFileHandle As Integer, x As Integer, FileLen As Integer
  Dim ErnCodeFileRec As ErnCodeRecType
  ClearFieldsFlag = True
  PriorDesc = QPTrim$(fpDescription.Text)
  PriorFWT = QPTrim$(fpcomboFWT.Text)
  PriorSWT = QPTrim$(fpcomboSWT.Text)
  PriorSOC = QPTrim$(fpcomboSOC.Text)
  PriorMED = QPTrim$(fpcomboMED.Text)
  PriorRET = QPTrim$(fpcomboRET.Text)
  Prior401K = QPTrim$(fpcmb401KXN.Text)
  If PriorDesc = "" And PriorFWT = "" And PriorSWT = "" And PriorSOC = "" And PriorMED = "" And PriorRET = "" And Prior401K = "" Then Exit Sub
  Call CheckForChanges
  If changeFlag = True Then
    If MsgBox("Your last edit was not saved. Do you want to save it?", vbYesNo) = vbYes Then
      Call cmdSaveContinue_Click
      'exit here if data is bad so the user can make corrections
      If BadDataFlag = True Then GoTo BadData
      ClearFieldsFlag = False
    End If
  End If
  If ClearFieldsFlag = True Then
     fpDescription.Text = ""
     fpcomboFWT.Text = ""
     fpcomboSWT.Text = ""
     fpcomboSOC.Text = ""
     fpcomboMED.Text = ""
     fpcomboRET.Text = ""
     fpcmb401KXN.Text = ""
  End If
  'reset all flags except ClearFieldsFlag to False and
  'reset RowFlag to allow a new entry
BadData:
  BadDataFlag = False
  ContinueFlag = False
  NoDataFlag = False
  SaveAndExitFlag = False
  JustExitFlag = False
  DupFlag = False
  changeFlag = False
  ClickFlag = False
  OpenErnCodeFile ErnCodeFileHandle
  FileLen = LOF(ErnCodeFileHandle) / Len(ErnCodeFileRec)
  Close ErnCodeFileHandle
  RowFlag = FileLen + 1
  fpDescription.SetFocus

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fpcmb401KXN_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmb401KXN.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmb401KXN.ListIndex = -1
  End If
  If fpcmb401KXN.ListDown = False Then
    If KeyCode = vbKeyDown Then
      fpDescription.SetFocus
      KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
      SendKeys "+{Tab}"
      KeyCode = 0
    End If
  End If

End Sub

Private Sub fpcmb401KXN_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fpcomboFWT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboFWT.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboFWT.ListIndex = -1
  End If
  If fpcomboFWT.ListDown = False Then
    If KeyCode = vbKeyUp Then
      SendKeys "+{Tab}"
      KeyCode = 0
    ElseIf KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    End If
  End If
End Sub

Private Sub fpcomboFWT_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fpcomboFWT_LostFocus()
  fpcomboFWT.Action = ActionClearSearchBuffer

End Sub

Private Sub fpcomboMED_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboMED.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboMED.ListIndex = -1
  End If
  If fpcomboMED.ListDown = False Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
      SendKeys "+{Tab}"
      KeyCode = 0
    End If
  End If
End Sub

Private Sub fpcomboMED_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fpcomboMED_LostFocus()
  fpcomboMED.Action = ActionClearSearchBuffer

End Sub

Private Sub fpcomboRET_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboRET.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboRET.ListIndex = -1
  End If
  If fpcomboRET.ListDown = False Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
      SendKeys "+{Tab}"
      KeyCode = 0
    End If
  End If
End Sub

Private Sub fpcomboRET_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fpcomboRET_LostFocus()
  fpcomboRET.Action = ActionClearSearchBuffer

End Sub

Private Sub fpcomboSOC_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboSOC.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboSOC.ListIndex = -1
  End If
  If fpcomboSOC.ListDown = False Then
    If KeyCode = vbKeyUp Then
      SendKeys "+{Tab}"
      KeyCode = 0
    ElseIf KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    End If
  End If

End Sub

Private Sub fpcomboSOC_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub fpcomboSOC_LostFocus()
  fpcomboSOC.Action = ActionClearSearchBuffer

End Sub

Private Sub fpcomboSWT_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboSWT.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboSWT.ListIndex = -1
  End If
  If fpcomboSWT.ListDown = False Then
    If KeyCode = vbKeyUp Then
      SendKeys "+{Tab}"
      KeyCode = 0
    ElseIf KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    End If
  End If

End Sub

Private Sub fpcomboSWT_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%X"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%C"
      Call cmdSaveContinue_Click
      KeyCode = 0
    Case vbKeyF11:
      SendKeys "%E"
      Call cmdSaveExit_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub
Private Sub Form_Load()
  Dim cnt As Integer
  Dim ScrWidth As Long
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  FirstTimeThru = True
  Call FixSpread
  Me.HelpContextID = hlpAdditionalEarning
  LoadECFile
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub cmdExit_Click()
  Dim A%, b%, C%, D%, E%, f%, G%
  A = Len(QPTrim$(fpDescription.Text))
  b = Len(QPTrim$(fpcomboRET.Text))
  C = Len(QPTrim$(fpcomboFWT.Text))
  D = Len(QPTrim$(fpcomboSWT.Text))
  E = Len(QPTrim$(fpcomboSOC.Text))
  f = Len(QPTrim$(fpcomboMED.Text))
  G = Len(QPTrim$(fpcmb401KXN.Text))
  If A + b + C + D + E + f = 0 Then GoTo EmptyFields
  JustExitFlag = True
  Call CheckForChanges
  If changeFlag = True Then
     If MsgBox("A change has been made. Do you want to exit without saving it?", vbYesNo) = vbNo Then
       If FirstTimeThru = True Or ClearFieldsFlag = True Then
         If A <> 0 Then fpDescription.SetFocus
         If b <> 0 Then fpcomboRET.SetFocus
         If C <> 0 Then fpcomboFWT.SetFocus
         If D <> 0 Then fpcomboSWT.SetFocus
         If E <> 0 Then fpcomboSOC.SetFocus
         If f <> 0 Then fpcomboMED.SetFocus
         If G <> 0 Then fpcmb401KXN.SetFocus
       End If
     Exit Sub
     End If
  End If
EmptyFields:
  frmControlFileMaint.Show
  DoEvents
  Unload frmEarningsCodeMaint
End Sub

Private Sub cmdSaveContinue_Click()
   Dim ErnCodeFileHandle As Integer, x As Integer, FileLen As Integer
   Dim ErnCodeFileRec As ErnCodeRecType
   Dim RowCount As Integer
   Dim A$, b$, C$, D$, E$, f$, G$
   Dim TempRowFlag As Integer
   
   If DescInUseCheck(fpDescription.Text, PriorRowNum) = True Then
      MsgBox "This Description is already in use. Please select another Description or press Exit to escape."
      fpDescription.SetFocus
      DupFlag = True
      Exit Sub
   End If
   A = Len(QPTrim$(fpDescription.Text))
   b = Len(QPTrim$(fpcomboRET.Text))
   C = Len(QPTrim$(fpcomboFWT.Text))
   D = Len(QPTrim$(fpcomboSWT.Text))
   E = Len(QPTrim$(fpcomboSOC.Text))
   f = Len(QPTrim$(fpcomboMED.Text))
   G = Len(QPTrim$(fpcmb401KXN.Text))
   'If the user wants to save and then exit the screen
   'we do not turn on the ContinueFlag
   If SaveAndExitFlag = True Then
      ContinueFlag = False
   Else
      ContinueFlag = True
   End If
   'because more than one field needs to have the focus set each
   'If statement below must handle code individually instead of sending
   'error traps to a goto line as done in the next series of If
   'statements
   If A <> 0 Then
      If b = 0 Then
         fpcomboRET.SetFocus
         MsgBox "All fields must be filled out if the Description field is filled out."
         BadDataFlag = True
         'reset the screen to where it was when it was last valid
         GoTo ExitTran
      End If
      If C = 0 Then
         fpcomboFWT.SetFocus
         MsgBox "All fields must be filled out if the Description field is filled out."
         BadDataFlag = True
         GoTo ExitTran
      End If
      If D = 0 Then
         fpcomboSWT.SetFocus
         MsgBox "All fields must be filled out if the Description field is filled out."
         BadDataFlag = True
         GoTo ExitTran
      End If
      If E = 0 Then
         fpcomboSOC.SetFocus
         MsgBox "All fields must be filled out if the Description field is filled out."
         BadDataFlag = True
         GoTo ExitTran
      End If
      If f = 0 Then
         fpcomboMED.SetFocus
         MsgBox "All fields must be filled out if the Description field is filled out."
         BadDataFlag = True
         GoTo ExitTran
      End If
      If G = 0 Then
         fpcmb401KXN.SetFocus
         MsgBox "All fields must be filled out if the Description field is filled out."
         BadDataFlag = True
         GoTo ExitTran
      End If
   End If
   'If nothing has been entered and the user tries to save
   'a message box alerts them to this because we do not want
   'to save empty fields
   If A + b + C + D + E + f + G = 0 Then
      'NoDataFlag is set to True so if the user wanted to
      'exit the screen after this save then that procedure
      'will behave with the save response
      NoDataFlag = True
      If SaveAndExitFlag = False Then
         MsgBox "No new or edited data to save"
      End If
      Exit Sub
   End If
   'if the description field is empty and any other field
   'is not empty this if statement traps this error
   If A = 0 Then
      If b <> 0 Then GoTo BadDataEntry
      If C <> 0 Then GoTo BadDataEntry
      If D <> 0 Then GoTo BadDataEntry
      If E <> 0 Then GoTo BadDataEntry
      If f <> 0 Then GoTo BadDataEntry
      If G <> 0 Then GoTo BadDataEntry
      Else: GoTo EntryDataOK
BadDataEntry:
       MsgBox "Please complete the Description field, double click an existing account to edit or delete all fields to continue."
       BadDataFlag = True
       fpDescription.SetFocus
       GoTo ExitTran
   End If
EntryDataOK:
'   OpenEarnNoMatchFile ErnNoMHandle
   
   OpenErnCodeFile ErnCodeFileHandle
   
   FileLen = LOF(ErnCodeFileHandle) / Len(ErnCodeFileRec)
   If FileLen = 0 Then 'first save
      vaSpreadEarningsCodes.Col = 1
      vaSpreadEarningsCodes.Row = 1
      ErnCodeFileRec.ERNCODE1 = QPTrim$(fpDescription.Text)
      vaSpreadEarningsCodes.Col = 2
      vaSpreadEarningsCodes.Row = 1
      ErnCodeFileRec.ERNFWT1 = QPTrim$(fpcomboFWT.Text)
      vaSpreadEarningsCodes.Col = 3
      vaSpreadEarningsCodes.Row = 1
      ErnCodeFileRec.ERNSWT1 = QPTrim$(fpcomboSWT.Text)
      vaSpreadEarningsCodes.Col = 4
      vaSpreadEarningsCodes.Row = 1
      ErnCodeFileRec.ERNSOC1 = QPTrim$(fpcomboSOC.Text)
      vaSpreadEarningsCodes.Col = 5
      vaSpreadEarningsCodes.Row = 1
      ErnCodeFileRec.ERNMED1 = QPTrim$(fpcomboMED.Text)
      vaSpreadEarningsCodes.Col = 6
      vaSpreadEarningsCodes.Row = 1
      ErnCodeFileRec.ERNRET1 = QPTrim$(fpcomboRET.Text)
      vaSpreadEarningsCodes.Col = 7
      vaSpreadEarningsCodes.Row = 1
      ErnCodeFileRec.EarnYN = QPTrim$(fpcmb401KXN.Text)
      ErnCodeFileRec.Pad = ""
      Put ErnCodeFileHandle, 1, ErnCodeFileRec
      Close ErnCodeFileHandle
      GoTo ClickSave
   End If
   'ClickFlag denotes we are here because the user double clicked a row
   'to edit it
   'If RowFlag is not revalued to the row that was just changed
   'the change takes place but in the row that is now in focus
   'causing data to be saved to the wrong row
   If changeFlag = True Then
      TempRowFlag = RowFlag 'save current row setting
      RowFlag = PriorRowNum 'reset row to the one that was changed
   End If
   'if ClearFieldsFlag is true then we do not want to save anything
   'until we find the first empty row
   If ClearFieldsFlag = True And RowFlag > FileLen Then GoTo NonEditEntry
   If ClickFlag = True Then 'save row that was double clicked for edit
      Get ErnCodeFileHandle, RowFlag, ErnCodeFileRec
      vaSpreadEarningsCodes.Col = 1
      vaSpreadEarningsCodes.Row = RowFlag
      ErnCodeFileRec.ERNCODE1 = QPTrim$(fpDescription.Text)
      vaSpreadEarningsCodes.Col = 2
      vaSpreadEarningsCodes.Row = RowFlag
      ErnCodeFileRec.ERNFWT1 = QPTrim$(fpcomboFWT.Text)
      vaSpreadEarningsCodes.Col = 3
      vaSpreadEarningsCodes.Row = RowFlag
      ErnCodeFileRec.ERNSWT1 = QPTrim$(fpcomboSWT.Text)
      vaSpreadEarningsCodes.Col = 4
      vaSpreadEarningsCodes.Row = RowFlag
      ErnCodeFileRec.ERNSOC1 = QPTrim$(fpcomboSOC.Text)
      vaSpreadEarningsCodes.Col = 5
      vaSpreadEarningsCodes.Row = RowFlag
      ErnCodeFileRec.ERNMED1 = QPTrim$(fpcomboMED.Text)
      vaSpreadEarningsCodes.Col = 6
      vaSpreadEarningsCodes.Row = RowFlag
      ErnCodeFileRec.ERNRET1 = QPTrim$(fpcomboRET.Text)
      vaSpreadEarningsCodes.Col = 7
      vaSpreadEarningsCodes.Row = RowFlag
      ErnCodeFileRec.EarnYN = QPTrim$(fpcmb401KXN.Text)
      ErnCodeFileRec.Pad = ""
      ClickFlag = False
      Put ErnCodeFileHandle, RowFlag, ErnCodeFileRec
      
      Close ErnCodeFileHandle
      
      'change RowFlag back to original value
      If changeFlag = True Then
          changeFlag = False
          RowFlag = TempRowFlag
      Else
          RowFlag = 0
      End If
      GoTo ClickSave
   End If
   'save data from fields at top of form

NonEditEntry:
   For x = 1 To 3
      vaSpreadEarningsCodes.Col = 1
      vaSpreadEarningsCodes.Row = x
      If Len(QPTrim$(vaSpreadEarningsCodes.Value)) = 0 Then
      'save in the next empty row
         RowCount = x
         Exit For
      End If
   Next
  Dim EarnAlert As TempEarnAlertType
  Dim EHandle As Integer
  Dim NumOfEarnAlerts As Integer
  
   If x > 3 Then
     MsgBox "You have reached the maximum allowable deductions."
     Close
     If SaveAndExitFlag = True Then DupFlag = True
     Exit Sub
   End If
   OpenEarnAlertFile EHandle
   NumOfEarnAlerts = LOF(EHandle) / Len(EarnAlert)
   vaSpreadEarningsCodes.Col = 1
   vaSpreadEarningsCodes.Row = RowCount
   
   EarnAlert.ERNCODE1 = QPTrim$(fpDescription.Text)
   EarnAlert.Number = RowCount
   Put EHandle, NumOfEarnAlerts + 1, EarnAlert
   Close EHandle
   
   ErnCodeFileRec.ERNCODE1 = QPTrim$(fpDescription.Text)
   vaSpreadEarningsCodes.Col = 2
   vaSpreadEarningsCodes.Row = RowCount
   ErnCodeFileRec.ERNFWT1 = QPTrim$(fpcomboFWT.Text)
   vaSpreadEarningsCodes.Col = 3
   vaSpreadEarningsCodes.Row = RowCount
   ErnCodeFileRec.ERNSWT1 = QPTrim$(fpcomboSWT.Text)
   vaSpreadEarningsCodes.Col = 4
   vaSpreadEarningsCodes.Row = RowCount
   ErnCodeFileRec.ERNSOC1 = QPTrim$(fpcomboSOC.Text)
   vaSpreadEarningsCodes.Col = 5
   vaSpreadEarningsCodes.Row = RowCount
   ErnCodeFileRec.ERNMED1 = QPTrim$(fpcomboMED.Text)
   vaSpreadEarningsCodes.Col = 6
   vaSpreadEarningsCodes.Row = RowCount
   ErnCodeFileRec.ERNRET1 = QPTrim$(fpcomboRET.Text)
   vaSpreadEarningsCodes.Col = 7
   vaSpreadEarningsCodes.Row = RowCount
   ErnCodeFileRec.EarnYN = QPTrim$(fpcmb401KXN.Text)
   ErnCodeFileRec.Pad = ""
   Put ErnCodeFileHandle, RowCount, ErnCodeFileRec
   Close ErnCodeFileHandle
   'Save And Exit command button used so we don't need anything between here
   'and ExitTran
ClickSave:
   BadDataFlag = False
   FirstTimeThru = False
   changeFlag = False
   If SaveAndExitFlag = True Then GoTo ExitTran ' this save is coming from the
   'the exit and save routine that has already performed everything
   'from here to ExitTran
   MsgBox "Your Information has been saved.", vbOKOnly
   LoadECFile
   fpDescription.SetFocus
ExitTran:
   FirstTimeThru = False
   MainLog ("Earnings Code data was saved.")
End Sub

Private Sub cmdSaveExit_Click()
   SaveAndExitFlag = True
   Call cmdSaveContinue_Click
   If DupFlag = True Then
      DupFlag = False
      Exit Sub
   End If
   If BadDataFlag = True Then
      GoTo ExitTran
   End If
   If NoDataFlag = True Then
      MsgBox "No new or edited data to save"
      GoTo NoData
   End If
   MsgBox "Your Information has been saved.", vbOKOnly
NoData:
   SaveAndExitFlag = False
   frmControlFileMaint.Show
   DoEvents
   Unload frmEarningsCodeMaint
ExitTran:

End Sub

Private Sub LoadECFile()
   Dim ErnCodeFileHandle As Integer, x As Integer, FileLen As Integer
   Dim ErnCodeFileRec As ErnCodeRecType
   
   'all fields in the upper block must be cleared for the
   'ClickFlag to work properly
   NoDataFlag = False
   'not resetting this ClickFlag causes the program to think
   'that the ClickFlag is still on if you exit and then
   'immediately return
   ClickFlag = False
   fpDescription.Text = ""
   fpcomboFWT.Text = ""
   fpcomboSWT.Text = ""
   fpcomboSOC.Text = ""
   fpcomboMED.Text = ""
   fpcomboRET.Text = ""
   fpcmb401KXN.Text = ""
   'load the combo boxes in the upper block..if we are reloading from
   'the Save and Continue button then we don't need to reload the combo
   'boxes because the form was never unloaded
   If FirstTimeThru = True Then
      fpcomboFWT.AddItem "Y"
      fpcomboFWT.AddItem "N"
      fpcomboSWT.AddItem "Y"
      fpcomboSWT.AddItem "N"
      fpcomboSOC.AddItem "Y"
      fpcomboSOC.AddItem "N"
      fpcomboMED.AddItem "Y"
      fpcomboMED.AddItem "N"
      fpcomboRET.AddItem "Y"
      fpcomboRET.AddItem "N"
      fpcmb401KXN.AddItem "Y"
      fpcmb401KXN.AddItem "N"
   End If
   
   OpenErnCodeFile ErnCodeFileHandle
   FileLen = LOF(ErnCodeFileHandle) / Len(ErnCodeFileRec)
   'This for loop loads all data stored on file plus it loads "N" in
   'the FWT, SWT, SOC ,MED and RET fields if no description is on that row
   
   
   For x = 1 To FileLen
      Get ErnCodeFileHandle, x, ErnCodeFileRec
   'load form info
      vaSpreadEarningsCodes.Col = 1
      vaSpreadEarningsCodes.Row = x
      If Len(QPTrim$(ErnCodeFileRec.ERNCODE1)) = 0 Then Exit For '8/19 added
      vaSpreadEarningsCodes.Text = QPTrim$(ErnCodeFileRec.ERNCODE1)
      vaSpreadEarningsCodes.Col = 2
      vaSpreadEarningsCodes.Row = x
      vaSpreadEarningsCodes.Text = QPTrim$(ErnCodeFileRec.ERNFWT1)
      vaSpreadEarningsCodes.Col = 3
      vaSpreadEarningsCodes.Row = x
      vaSpreadEarningsCodes.Text = QPTrim$(ErnCodeFileRec.ERNSWT1)
      vaSpreadEarningsCodes.Col = 4
      vaSpreadEarningsCodes.Row = x
      vaSpreadEarningsCodes.Text = QPTrim$(ErnCodeFileRec.ERNSOC1)
      vaSpreadEarningsCodes.Col = 5
      vaSpreadEarningsCodes.Row = x
      vaSpreadEarningsCodes.Text = QPTrim$(ErnCodeFileRec.ERNMED1)
      vaSpreadEarningsCodes.Col = 6
      vaSpreadEarningsCodes.Row = x
      vaSpreadEarningsCodes.Text = QPTrim$(ErnCodeFileRec.ERNRET1)
      vaSpreadEarningsCodes.Col = 7
      vaSpreadEarningsCodes.Row = x
      vaSpreadEarningsCodes.Text = QPTrim$(ErnCodeFileRec.EarnYN)
   Next
   Close ErnCodeFileHandle
'   Close ErnNoMHandle
   BadDataFlag = False
End Sub

Private Sub fpcomboSWT_LostFocus()
  fpcomboSWT.Action = ActionClearSearchBuffer
End Sub

Private Sub fpDescription_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    fpcomboRET.SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If
  
End Sub

Private Sub mnuExit_Click()
  Call cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
  MainLog ("Earnings Code Maintenance screen printed.")
End Sub

Private Sub vaSpreadEarningsCodes_DblClick(ByVal Col As Long, ByVal Row As Long)
  'save all data before the doubleclick removed them
  
  Dim ErnHandle As Integer
  Dim ErnRec As ErnCodeRecType
  Dim ErnCnt As Integer
  
  OpenErnCodeFile ErnHandle
  ErnCnt = LOF(ErnHandle) / Len(ErnRec)
  Close ErnHandle
  
  If Row > ErnCnt Then
    MsgBox "Empty rows cannot be edited."
    Exit Sub
  End If
  FirstTimeThru = False
  changeFlag = False
  RowFlag = Row
  PriorDesc = QPTrim$(fpDescription.Text)
  PriorFWT = QPTrim$(fpcomboFWT.Text)
  PriorSWT = QPTrim$(fpcomboSWT.Text)
  PriorSOC = QPTrim$(fpcomboSOC.Text)
  PriorMED = QPTrim$(fpcomboMED.Text)
  PriorRET = QPTrim$(fpcomboRET.Text)
  Prior401K = QPTrim$(fpcmb401KXN.Text)
  'if ClearFieldsFlag is true we've already checked for changes
  If ClearFieldsFlag = True Then
     ClearFieldsFlag = False
     GoTo NoChangeCheck
  End If
  If ClickFlag = True Then
    Call CheckForChanges
      If changeFlag = True Then
        If MsgBox("Your last edit was not saved. Do you want to save it?", vbYesNo) = vbYes Then
          Call cmdSaveContinue_Click
          changeFlag = False
        End If
      End If
   End If
'This routine allows the user to double click a specific row
'that places that row's data in the edit fields
NoChangeCheck:
   ClickFlag = True
   'load the fields in the upper block with the data for
   'the file numbered as Row
   vaSpreadEarningsCodes.Col = 1
   vaSpreadEarningsCodes.Row = Row
   fpDescription.Text = QPTrim$(vaSpreadEarningsCodes.Value)
   vaSpreadEarningsCodes.Col = 2
   vaSpreadEarningsCodes.Row = Row
   fpcomboFWT.Text = QPTrim$(vaSpreadEarningsCodes.Text)
   vaSpreadEarningsCodes.Col = 3
   vaSpreadEarningsCodes.Row = Row
   fpcomboSWT.Text = QPTrim$(vaSpreadEarningsCodes.Text)
   vaSpreadEarningsCodes.Col = 4
   vaSpreadEarningsCodes.Row = Row
   fpcomboSOC.Text = QPTrim$(vaSpreadEarningsCodes.Text)
   vaSpreadEarningsCodes.Col = 5
   vaSpreadEarningsCodes.Row = Row
   fpcomboMED.Text = QPTrim$(vaSpreadEarningsCodes.Text)
   vaSpreadEarningsCodes.Col = 6
   vaSpreadEarningsCodes.Row = Row
   fpcomboRET.Text = QPTrim$(vaSpreadEarningsCodes.Text)
   vaSpreadEarningsCodes.Col = 7
   vaSpreadEarningsCodes.Row = Row
   fpcmb401KXN.Text = QPTrim$(vaSpreadEarningsCodes.Text)
   PriorRowNum = RowFlag

End Sub

Private Sub CheckForChanges()
'This routine compares data in the row that just lost focus with the data
'that is in the appropriate row in the spreadsheet...if a change
'has been made it will be detected here
   changeFlag = False
   If FirstTimeThru = True Then PriorRowNum = 1
   If FirstTimeThru = True Or JustExitFlag = True Then PriorDesc = QPTrim$(fpDescription.Text)
   vaSpreadEarningsCodes.Col = 1
   vaSpreadEarningsCodes.Row = PriorRowNum
   If QPTrim$(vaSpreadEarningsCodes.Text) <> QPTrim$(PriorDesc) Then 'QPTrim$(ErnCodeFileRec.ERNCODE1) Then
     changeFlag = True
     fpDescription.SetFocus
   End If
   vaSpreadEarningsCodes.Col = 2
   vaSpreadEarningsCodes.Row = PriorRowNum
   If FirstTimeThru = True Or JustExitFlag = True Then PriorFWT = QPTrim$(fpcomboFWT.Text)
   If QPTrim$(vaSpreadEarningsCodes.Text) <> QPTrim$(PriorFWT) Then
     changeFlag = True
     fpcomboFWT.SetFocus
   End If
   vaSpreadEarningsCodes.Col = 3
   vaSpreadEarningsCodes.Row = PriorRowNum
   If FirstTimeThru = True Or JustExitFlag = True Then PriorSWT = QPTrim$(fpcomboSWT.Text)
   If QPTrim$(vaSpreadEarningsCodes.Text) <> QPTrim$(PriorSWT) Then
     changeFlag = True
     fpcomboSWT.SetFocus
   End If
   vaSpreadEarningsCodes.Col = 4
   vaSpreadEarningsCodes.Row = PriorRowNum
   If FirstTimeThru = True Or JustExitFlag = True Then PriorSOC = QPTrim$(fpcomboSOC.Text)
   If QPTrim$(vaSpreadEarningsCodes.Text) <> QPTrim$(PriorSOC) Then
     changeFlag = True
     fpcomboSOC.SetFocus
   End If
   vaSpreadEarningsCodes.Col = 5
   vaSpreadEarningsCodes.Row = PriorRowNum
   If FirstTimeThru = True Or JustExitFlag = True Then PriorMED = QPTrim$(fpcomboMED.Text)
   If QPTrim$(vaSpreadEarningsCodes.Text) <> QPTrim$(PriorMED) Then
     changeFlag = True
     fpcomboMED.SetFocus
   End If
   vaSpreadEarningsCodes.Col = 6
   vaSpreadEarningsCodes.Row = PriorRowNum
   If FirstTimeThru = True Or JustExitFlag = True Then PriorRET = QPTrim$(fpcomboRET.Text)
   If QPTrim$(vaSpreadEarningsCodes.Text) <> QPTrim$(PriorRET) Then
     changeFlag = True
     fpcomboRET.SetFocus
   End If
   vaSpreadEarningsCodes.Col = 7
   vaSpreadEarningsCodes.Row = PriorRowNum
   If FirstTimeThru = True Or JustExitFlag = True Then Prior401K = QPTrim$(fpcmb401KXN.Text)
   If QPTrim$(vaSpreadEarningsCodes.Text) <> QPTrim$(Prior401K) Then
     changeFlag = True
     fpcmb401KXN.SetFocus
   End If
   JustExitFlag = False
End Sub

Private Function DescInUseCheck(Desc As String, ThisRow As Integer) As Boolean
   Dim ErnCodeFileHandle As Integer, x As Integer, FileLen As Integer
   Dim ErnCodeFileRec As ErnCodeRecType
   
   If QPTrim$(Desc) = "" Then Exit Function
   
   DescInUseCheck = False
   OpenErnCodeFile ErnCodeFileHandle
   FileLen = LOF(ErnCodeFileHandle) / Len(ErnCodeFileRec)
   For x = 1 To FileLen
      If x = ThisRow Then GoTo ThisRowEdit
      Get ErnCodeFileHandle, x, ErnCodeFileRec
      If QPTrim$(Desc) = QPTrim$(ErnCodeFileRec.ERNCODE1) Then
         DescInUseCheck = True
         Exit For
      End If
ThisRowEdit:
   Next x
  Close ErnCodeFileHandle
  
End Function

Private Sub FixSpread()
  Dim COne As Integer
  Dim CTwo As Integer
  Dim CThree As Integer
  Dim CFour As Integer
  Dim CFive As Integer
  Dim CSix As Integer
  '-1 means all rows or all columns....0 means headers
    Select Case ScreenW
      Case 1280
        COne = 10
        coladj = 2
        vaSpreadEarningsCodes.RowHeight(-1) = 18
        vaSpreadEarningsCodes.RowHeight(0) = 18
      Case 1152
        COne = 3
        coladj = 1
        vaSpreadEarningsCodes.RowHeight(0) = 15
        vaSpreadEarningsCodes.RowHeight(-1) = 15
      Case 1024
        COne = 6
        coladj = 4
      Case 800
'        COne = 5
'        coladj = 1
'        vaSpreadEarningsCodes.Font.Size = 12
'        vaSpreadEarningsCodes.RowHeight(-1) = 14
'        vaSpreadEarningsCodes.FontBold = True
      Case Else
       
    End Select
    vaSpreadEarningsCodes.ColWidth(1) = vaSpreadEarningsCodes.ColWidth(1) + COne
    vaSpreadEarningsCodes.ColWidth(2) = vaSpreadEarningsCodes.ColWidth(2) + coladj
    vaSpreadEarningsCodes.ColWidth(3) = vaSpreadEarningsCodes.ColWidth(3) + coladj
    vaSpreadEarningsCodes.ColWidth(4) = vaSpreadEarningsCodes.ColWidth(4) + coladj
    vaSpreadEarningsCodes.ColWidth(5) = vaSpreadEarningsCodes.ColWidth(5) + coladj
    vaSpreadEarningsCodes.ColWidth(6) = vaSpreadEarningsCodes.ColWidth(6) + coladj

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmEarningsCodeMaint.")
      Call Terminate
      End
    End If
  End If
End Sub

