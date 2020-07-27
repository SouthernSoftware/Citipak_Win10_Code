VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmPrnDeptBudAct 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Department Budget vs Actual"
   ClientHeight    =   8640
   ClientLeft      =   30
   ClientTop       =   540
   ClientWidth     =   12195
   Icon            =   "frmPrnDeptBudAct.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboIncAcct 
      Height          =   405
      Left            =   5850
      TabIndex        =   6
      Top             =   5955
      Width           =   1410
      _Version        =   196608
      _ExtentX        =   2487
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
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmPrnDeptBudAct.frx":08CA
   End
   Begin LpLib.fpCombo fpcboRptType 
      Height          =   405
      Left            =   5850
      TabIndex        =   7
      Top             =   6525
      Width           =   1920
      _Version        =   196608
      _ExtentX        =   3387
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
      ColDesigner     =   "frmPrnDeptBudAct.frx":0C31
   End
   Begin LpLib.fpCombo fpcboFund2 
      Height          =   405
      Left            =   5850
      TabIndex        =   2
      Top             =   3660
      Width           =   2760
      _Version        =   196608
      _ExtentX        =   4868
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
      Columns         =   3
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   2
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
      ColDesigner     =   "frmPrnDeptBudAct.frx":0FCF
   End
   Begin LpLib.fpCombo fpcboFund1 
      Height          =   405
      Left            =   5850
      TabIndex        =   1
      Top             =   3090
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
      Columns         =   3
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   2
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
      AutoSearchFillDelay=   100
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmPrnDeptBudAct.frx":13BA
   End
   Begin LpLib.fpCombo fpcboReportOn 
      Height          =   405
      Left            =   5850
      TabIndex        =   5
      Top             =   5370
      Width           =   3615
      _Version        =   196608
      _ExtentX        =   6376
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
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmPrnDeptBudAct.frx":17A5
   End
   Begin LpLib.fpCombo fpcboRepNum 
      Height          =   405
      Left            =   5850
      TabIndex        =   4
      Top             =   4800
      Width           =   3960
      _Version        =   196608
      _ExtentX        =   6985
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
      ColDesigner     =   "frmPrnDeptBudAct.frx":1B0C
   End
   Begin LpLib.fpCombo fpcboDepartment 
      Height          =   405
      Left            =   5850
      TabIndex        =   3
      Top             =   4230
      Width           =   2175
      _Version        =   196608
      _ExtentX        =   3836
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
      Columns         =   3
      Sorted          =   0
      SelDrawFocusRect=   -1  'True
      ColumnSeparatorChar=   9
      ColumnSearch    =   1
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
      AutoSearchFillDelay=   100
      EditMarginLeft  =   1
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmPrnDeptBudAct.frx":1ECB
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
      Top             =   7464
      Width           =   1332
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
      Top             =   7464
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   8280
      Width           =   12192
      _ExtentX        =   21511
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7117
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "10:20 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "3/16/2011"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin EditLib.fpDateTime txtDate 
      Height          =   372
      Left            =   5856
      TabIndex        =   0
      Top             =   2520
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
      Left            =   2472
      TabIndex        =   19
      Top             =   6000
      Width           =   3060
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
      Left            =   3144
      TabIndex        =   18
      Top             =   6576
      Width           =   2388
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Fund:"
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
      Left            =   3936
      TabIndex        =   17
      Top             =   3708
      Width           =   1596
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   4836
      Left            =   1872
      Top             =   2328
      Width           =   8268
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Print Department Budget vs Actual"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3600
      TabIndex        =   16
      Top             =   1272
      Width           =   5004
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   3192
      Top             =   1032
      Width           =   5772
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   3120
      Picture         =   "frmPrnDeptBudAct.frx":22B6
      Top             =   2685
      Width           =   360
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
      Left            =   3960
      TabIndex        =   15
      Top             =   2556
      Width           =   1572
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Fund:"
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
      Left            =   3864
      TabIndex        =   14
      Top             =   3132
      Width           =   1668
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Department Number:"
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
      Left            =   3024
      TabIndex        =   13
      Top             =   4284
      Width           =   2508
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
      Left            =   3504
      TabIndex        =   12
      Top             =   4848
      Width           =   2028
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
      Index           =   3
      Left            =   4128
      TabIndex        =   11
      Top             =   5424
      Width           =   1404
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      Height          =   972
      Left            =   3192
      Top             =   912
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
Attribute VB_Name = "frmPrnDeptBudAct"
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
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim FY1BegDate As Integer, FY1EndDate As Integer, FY2BegDate As Integer, FY2EndDate As Integer
Dim StartFund As String, EndFund As String, FYStartDate As Integer
Dim ActiveYear As Integer


Private Sub cmdExit_Click()
  frmGLReportsMenu.Show
  Unload frmPrnDeptBudAct
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
Private Function ValidFunds()
  If fpcboFund1.Text <> "" And fpcboFund2.Text <> "" Then
    fpcboFund1.Col = 0
    fpcboFund2.Col = 0
    If fpcboFund1.ColText > fpcboFund2.ColText Then
      MsgBox "Invalid Fund Selection, The Beginning Fund Should Be Less or Equal to Ending Fund.", vbOKOnly, "Invalid Selection"
      ValidFunds = False
    Else
      ValidFunds = True
      StartFund = QPTrim(fpcboFund1.ColText)
      EndFund = QPTrim(fpcboFund2.ColText)
    End If
  Else
    MsgBox "Fund Selections May Not Be Left Blank.", vbOKOnly, "Invalid Selection"
  End If
End Function
Private Sub fpcboFund1_GotFocus()
  fpcboFund1.Action = ActionClearSearchBuffer
End Sub
Private Sub fpcboFund2_GotFocus()
  fpcboFund2.Action = ActionClearSearchBuffer
End Sub
Private Sub fpcboDepartment_GotFocus()
  fpcboDepartment.Action = ActionClearSearchBuffer
End Sub

Private Sub fpcboIncAcct_GotFocus()
  fpcboIncAcct.Action = ActionClearSearchBuffer
End Sub

Private Sub fpcboRepNum_GotFocus()
  fpcboRepNum.Action = ActionClearSearchBuffer
End Sub
Private Sub fpcboReportOn_GotFocus()
  fpcboReportOn.Action = ActionClearSearchBuffer
End Sub

Private Sub cmdPrint_Click()
  If ValidDate = True Then
    If ValidFunds = True Then
      If fpcboRptType.ListIndex = 0 Then
        rptopt = 1
      ElseIf fpcboRptType.ListIndex = 1 Then
        rptopt = 2
      End If
      If rptopt = 1 Then
        PrintDeptBgtAct
      ElseIf rptopt = 2 Then
        PrintDeptBgtAct2
      End If
    End If
  End If
End Sub
Private Sub fpcboFund2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboFund2.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboFund2.ListIndex = -1
    fpcboFund2.Action = ActionClearSearchBuffer
  End If
  If fpcboFund2.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboDepartment.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboFund1.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub
Private Sub fpcboFund1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboFund1.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboFund1.ListIndex = -1
    fpcboFund1.Action = ActionClearSearchBuffer
  End If
  If fpcboFund1.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboFund2.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        txtDate.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub
Private Sub fpcboDepartment_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboDepartment.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboDepartment.ListIndex = -1
    fpcboDepartment.Action = ActionClearSearchBuffer
  End If
  If fpcboDepartment.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboRepNum.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboFund2.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub
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
      fpcboReportOn.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboDepartment.SetFocus
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
  If fpcboIncAcct.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      fpcboRptType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcboReportOn.SetFocus
        KeyCode = 0
      End If
    End If
  End If
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
  Me.HelpContextID = hlpDepartmentBudgetVs
  FundstoList fpcboFund1
  FundstoList fpcboFund2
  txtDate.Text = Format(Now, "mm/dd/yyyy")
  DeptList fpcboDepartment
  fpcboDepartment.RemoveItem 0
  fpcboRepNum.InsertRow = "1" & Chr$(9) & "Bgt,Mo,YTD,Var:Req. Wide Carriage"
  fpcboRepNum.InsertRow = "2" & Chr$(9) & "Bgt,Enc,YTD,Var:Req. Wide Carriage"
  fpcboRepNum.InsertRow = "3" & Chr$(9) & "Bgt,QTD,YTD,Var:Req. Wide Carriage"
  fpcboRepNum.ListIndex = 0
  fpcboReportOn.AddItem "Revenues Only"
  fpcboReportOn.AddItem "Expenditures Only"
  fpcboReportOn.AddItem "Both Revenues & Expenditures"
  fpcboReportOn.ListIndex = 2
  fpcboIncAcct.AddItem "Yes"
  fpcboIncAcct.AddItem "No"
  fpcboIncAcct.ListIndex = 0
  fpcboRptType.InsertRow = "Graphics"
  fpcboRptType.InsertRow = "Text"
  fpcboRptType.ListIndex = 0
End Sub

Private Sub Form_Resize()
'  If Me.Visible Then
    Temp_Class.ResizeControls Me
    DoEvents
'  End If
End Sub

Private Sub PrintDeptBgtAct()
  Dim CommaFmt As String, TotalFmt As String, RunBalFmt As String
  Dim SumLine As String, BgtFmt As String, BSumLine As String, PSumLine As String
  Dim DivLine As String, DivLine2 As String, FF As String, RptTitle As String
  Dim MaxLines As Integer, Col1 As Integer, Col2 As Integer, Col3 As Integer
  Dim Col4 As Integer, Col5 As Integer, EndDate As Integer, PRNFile As Integer
  Dim M As String, HollyFlag As Boolean, Pitch12 As String, ThisFund As String
  Dim DoingDetail As Boolean, SubTotalRevenues As Boolean, DeptOnNewPage As Boolean
  Dim WhichReport As Integer, GetMonth As Boolean, GetQtr As Boolean
  Dim RptMonth As String, ReportFile As String, FundCode As String
  Dim FundIdxFile As Integer, NumFunds As Integer, Acct As Integer
  Dim AcctIdxFileNum As Integer, NumGLAccts As Integer, FundName As String
  Dim AcctFileNum As Integer, NumGLAcctRecs As Integer, Rec As Integer
  Dim TransFileNum As Integer, NumTrans As Long, NextTr As Long, AcctNum As String
  Dim BGTBal As Double, YTDBal As Double, MTDBal As Double, UsingFund As Boolean
  Dim ECnt As Integer, RCnt As Integer, d As String, TransMonth As String
  Dim InThisQtr As Boolean, FirstTime As Boolean, Fund As Integer, FundRec As Integer
  Dim Dept As String, LastDept As String, LastDeptName As String, cnt As Integer
  Dim Account As String, BudgetAmt As Double, DeptRecNum As Integer, DeptName As String
  Dim Pct As String, Variance As Double, ToPrint As String, Linecnt As Integer
  Dim MTDSum As Double, BgtSum As Double, YTDSum As Double, DeptBgtSum As Double
  Dim DeptYTDSum As Double, DeptENCSum As Double, DeptMTDSum As Double
  Dim FundRevMTD As Double, FundRevBgt As Double, FundRevYTD As Double
  Dim EncSum As Double, FundExpMTD As Double, FundExpBgt As Double
  Dim FundExpYTD As Double, FundEncYTD As Double, EncBal As Double
  Dim DeptSummary As String, DeptTypeCode As String, DeptType As Integer
  Dim DoingRevenues As Boolean, DoingExp As Boolean, RptDept As String
  Dim DeptExp As String, DeptRev As String, MTDRSum As Double, Newrp As String
  Dim BgtRSum As Double, YTDRSum As Double, MTDESum As Double
  Dim BgtESum As Double, YTDESum As Double, PageNum As Integer, rpt As Integer
  Dim GetEnc As Boolean, IncAcct As Boolean

  '''GetFundCodes FirstFund$, LastFund$
  CommaFmt$ = "###,###,###.##"  'format takes 14 chars
  TotalFmt$ = "#,###,###,###.##" 'format takes 16 chars
  RunBalFmt$ = "##########.##"
  SumLine$ = String$(16, "-")   'column summary line

  BgtFmt$ = "###,###,###.##"         'format takes 11 chars
  BSumLine$ = String$(11, "-")  'summary line for budget columns
  PSumLine$ = "----"            'summary line for Pct columns
  DivLine$ = String$(96, "-")   'dashed line
  DivLine2$ = String$(96, "=")  'Double Line

  'DivLine$ = String$(79, "-")   'dashed line
  'DivLine2$ = String$(79, "=")  'Double Line
  FF$ = Chr$(12)
  ReDim Desc$(1)
  RptTitle$ = "Budget vs. Actual"
  MaxLines = 55
'  If InStr(UCase$(GLUserName), "HOLLY SPR") > 0 Then
'    HollyFlag = True
'    Pitch12$ = Chr$(27) + Chr$(38) + Chr$(107) + Chr$(52) + Chr$(83)
'  End If

  '--Column offsets for printing amounts
'''  Col1 = 38
'''  Col2 = 49
'''  Col3 = 63
'''  Col4 = 77
'''  Col5 = 90
  
  Col1 = 28    'Budget
  Col2 = 44   'Month or Actual YTD
  Col3 = 61    'YTD or Var
  Col4 = 78    'Enc
  Col5 = 93    'Pct
  If fpcboIncAcct.ListIndex = 1 Then
    IncAcct = False
  Else
    IncAcct = True
  End If

  EndDate = DateDiff("d", "12/31/1979", txtDate)

  M$ = Right$(txtDate, 2) + Left$(txtDate, 2)
  FrmShowPctComp.Label1 = "Printing Departmental Budget vs Actual Report"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmPrnDeptBudAct, True


  DeptTypeCode$ = GetDeptOffsets$
  DeptType = Val(DeptTypeCode$)
  If DeptType = 0 Then DeptType = 1

  WhichReport = Val(fpcboRepNum.Text)

  Select Case Mid$(fpcboReportOn.Text, 1, 1)
    Case "B"
      DoingRevenues = True
      DoingExp = True
      rpt = 3
    Case "R"
      DoingRevenues = True
      DoingExp = False
      rpt = 2
    Case "E"
      DoingRevenues = False
      DoingExp = True
      rpt = 1
  End Select
  fpcboDepartment.Col = 1
  RptDept$ = QPTrim$(fpcboDepartment.ColText)

  DeptRecNum = FindDept(RptDept$)
  If DeptRecNum > 0 Then
    DeptName$ = QPTrim$(GetDeptTitle$(DeptRecNum))
  Else
    DeptName$ = "Undefined: " + RptDept$
  End If

  Select Case WhichReport
  Case 1        'Bgt, Month, YTD
    GetMonth = True
    GetQtr = False
    GetEnc = False
    RptMonth = M$
''''''''''''''''123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456''''For 96
    Desc$(1) = "Description                      Budget          Month             YTD           Variance    Pct"
    MTDBal# = 0
  Case 2        'Bgt, Enc, YTD, Variance
    GetMonth = False
    GetEnc = True
    GetQtr = False
    Desc$(1) = "Description                      Budget          Encumb            YTD           Variance    Pct"
  Case 3        'Bgt, QTR, YTD, Variance
    GetMonth = False
    GetQtr = True
    GetEnc = False
    Desc$(1) = "Description                      Budget           QTD              YTD           Variance    Pct"
  End Select

  GetFYDates FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate
  Newrp = "BgtActD"
  GetRPTName Newrp
  ReportFile$ = Newrp
  'ReportFile$ = Unique$(Path$)
  If GetEnc = True Then
        'if need enc totals do this here
    FixPOEncumbRpt EndDate, FYStartDate
  End If

  'P17$ = CHR$(27) + "(s17H"
  PRNFile = FreeFile
  Open ReportFile$ For Output As #PRNFile
  'PRINT #PrnFile, P17$;
'  If HollyFlag Then
'    Print #PRNFile, Pitch12$;
'  End If

  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum, NumGLAcctRecs
  OpenTransFile TransFileNum, NumTrans&
  OpenFundIdx FundIdxFile, NumFunds

  If DoingRevenues Then
    ReDim RevAccts%(1 To NumGLAccts)              'Holds all rev acct record n
  End If
  ReDim ExpAccts%(1 To NumGLAccts)              'Holds all exp acct record num
  ReDim FundList(1) As String                            'List of all active Funds
  GetFundList FundList$(), NumFunds

  '--Build a list of revenue and exp accounts
  For Acct = 1 To NumGLAccts

    Get AcctIdxFileNum, Acct, GLAcctidx
    Get AcctFileNum, GLAcctidx.RecNum, GLAcct

    '--Find what fund this account is in
    FundCode$ = Left$(GLAcct.Num, GLFundLen)

    '--See if the account is in a fund we want to see
    If FundCode$ >= StartFund$ And FundCode$ <= EndFund$ Then

      '--Account is in fund, check to see if its proper type
      '--We want only revenue or expenditure accounts
      If GLAcct.Typ = "R" Or GLAcct.Typ = "E" Then
        Select Case GLAcct.Typ
          Case "E"
            If DoingExp Then
              Select Case DeptType
                Case 1
                  DeptExp$ = Mid$(GLAcct.Num, GLFundLen + 2, GLAcctLen)
                Case 2  'Appalachain Dist Health
                  DeptExp$ = Mid$(GLAcct.Num, GLFundLen + 3, GLAcctLen - 1)
                Case 3  'Oklahoma
                  DeptExp$ = Mid$(GLAcct.Num, GLFundLen + 2, GLAcctLen)
              End Select
              If DeptExp$ = RptDept$ Then
                ECnt = ECnt + 1
                ExpAccts%(ECnt) = GLAcctidx.RecNum
              End If
            End If
          Case "R"
            If DoingRevenues Then
              Select Case DeptType
                Case 1  'Normal
                  DeptRev$ = Mid$(GLAcct.Num, GLFundLen + 3, GLAcctLen - 1)
                Case 2  'Appalachain Dist Health
                  DeptRev$ = Mid$(GLAcct.Num, GLFundLen + 3, GLAcctLen - 1)
                Case 3  'Oklahoma
                  DeptRev$ = Mid$(GLAcct.Num, GLFundLen + 2, GLAcctLen)
              End Select
              If DeptRev$ = RptDept$ Then
                RCnt = RCnt + 1
                RevAccts%(RCnt) = GLAcctidx.RecNum
              End If
            End If
        End Select
      End If    '--test for rev or exp accts
    End If      '--End of acct in fund range test
          FrmShowPctComp.ShowPctComp Acct, NumGLAccts
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnDeptBudAct, True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
  Next          'Process next account


  ActivateControls frmPrnDeptBudAct, True

  '--Now write the report to file.

 ' GoSub PrintDeptPageHdr

  '--Search thru list of revenue accounts
  If DoingRevenues Then
    For cnt = 1 To RCnt
      Rec = RevAccts%(cnt)
      Get AcctFileNum, Rec, GLAcct

      GoSub CalcAcctBal
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

      Pct$ = GetPct$(GLAcct.YTD, BudgetAmt#)
      Variance# = GLAcct.YTD - BudgetAmt#   'Acct.Bgt

      ToPrint$ = Space$(96)
      ToPrint$ = RptDept$ + "~" + DeptName$ + "~R~" + Left$(Account$, 36)
      ToPrint$ = ToPrint$ + "~" + Using$(BgtFmt$, Str$(BudgetAmt#))  'changed
      Select Case WhichReport
      Case 1, 3
        ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(GLAcct.MTD))
        ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(GLAcct.YTD))
        ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(Variance#))
        ToPrint$ = ToPrint$ + "~" + Pct$
      Case 2
        ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(GLAcct.Encumb))
        ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(GLAcct.YTD))
        ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(Variance#))
        ToPrint$ = ToPrint$ + "~" + Pct$
      End Select
      Print #PRNFile, ToPrint$

      If GetMonth Or GetQtr Then 'changed
        MTDRSum# = MTDRSum# + GLAcct.MTD
      End If

      BgtRSum# = BgtRSum# + BudgetAmt# 'Acct.Bgt
      YTDRSum# = YTDRSum# + GLAcct.YTD

    Next      'Revenue Acct

    'GoSub PrintSummaryLines2
    ToPrint$ = Space$(96)
    'LSet ToPrint$ = "Total Revenues"
    Pct$ = GetPct$(YTDRSum#, BgtRSum#)
    Select Case WhichReport
    Case 1, 3
      Variance# = YTDRSum# - BgtRSum#
'      Mid$(ToPrint$, Col1) = Using$(BgtFmt$, Str$(BgtRSum#))
'      Mid$(ToPrint$, Col2 - 2) = Using$(TotalFmt$, Str$(MTDRSum#))
'      Mid$(ToPrint$, Col3 - 2) = Using$(TotalFmt$, Str$(YTDRSum#))
'      Mid$(ToPrint$, Col4) = Using$(CommaFmt$, Str$(Variance#))
'      Mid$(ToPrint$, Col5) = Pct$
      '--Reset vars
      FundRevMTD# = MTDSum#
      MTDSum# = 0
    Case 2
      Variance# = YTDSum# - BgtSum#
'      Mid$(ToPrint$, Col1) = Using$(BgtFmt$, Str$(BgtRSum#))
'      Mid$(ToPrint$, Col2 - 2) = Using$(TotalFmt$, " 0")
'      Mid$(ToPrint$, Col3 - 2) = Using$(TotalFmt$, Str$(YTDRSum#))
'      Mid$(ToPrint$, Col4) = Using$(CommaFmt$, Str$(Variance#))
'      Mid$(ToPrint$, Col5) = Pct$
    End Select
    'Print #PRNFile, RTrim$(ToPrint$)
'    Linecnt = Linecnt + 1
'    If Linecnt > MaxLines Then
'      Print #PRNFile, FF$
'      GoSub PrintDeptPageHdr
'    End If

'    Print #PRNFile,

  End If 'doing revenues

  '--Search exp accounts list for accounts in this fund
  If DoingExp Then
    For cnt = 1 To ECnt
      Rec = ExpAccts%(cnt)
      Get AcctFileNum, Rec, GLAcct
      GoSub CalcAcctBal
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

     ToPrint$ = Space$(96)
     'Pct$ = GetPct$(Acct.Encumb + Acct.YTD, BudgetAmt#) 'Acct.Bgt)
     ToPrint$ = RptDept$ + "~" + DeptName$ + "~E~" + Left$(Account$, 36)
     Select Case WhichReport
     Case 1, 3
       Pct$ = GetPct$(GLAcct.YTD, BudgetAmt#)  'Acct.Bgt)
       Variance# = BudgetAmt# - GLAcct.YTD
       ToPrint$ = ToPrint$ + "~" + Using$(BgtFmt$, Str$(BudgetAmt#))
       ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(GLAcct.MTD))
       ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(GLAcct.YTD))
       ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(Variance#))
       ToPrint$ = ToPrint$ + "~" + Pct$
     Case 2
       Pct$ = GetPct$(GLAcct.Encumb + GLAcct.YTD, BudgetAmt#) 'Acct.Bgt)
       Variance# = BudgetAmt# - GLAcct.Encumb - GLAcct.YTD
       ToPrint$ = ToPrint$ + "~" + Using$(BgtFmt$, Str$(BudgetAmt#))
       ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(GLAcct.Encumb))
       ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(GLAcct.YTD))
       ToPrint$ = ToPrint$ + "~" + Using$(CommaFmt$, Str$(Variance#))
       ToPrint$ = ToPrint$ + "~" + Pct$
      End Select
      Print #PRNFile, ToPrint$
'      Linecnt = Linecnt + 1
'      If Linecnt > MaxLines Then
'        Print #PRNFile, FF$
'        GoSub PrintDeptPageHdr
'      End If

      If GetMonth Or GetQtr Then 'changed
        MTDESum# = MTDESum# + GLAcct.MTD
      End If
      BgtESum# = BgtESum# + BudgetAmt#
      YTDESum# = YTDESum# + GLAcct.YTD
      EncSum# = EncSum# + GLAcct.Encumb

    Next      'Exp Acct

    'GoSub PrintSummaryLines2

    ToPrint$ = Space$(96)
   ' LSet ToPrint$ = "Total Expenditures"
    Pct$ = GetPct$(YTDESum#, BgtESum#)
    'Pct$ = GetPct$(YTDSum#, BgtSum#)
    Select Case WhichReport
    Case 1, 3
      Variance# = YTDESum# - BgtESum#
'      Mid$(ToPrint$, Col1) = Using$(BgtFmt$, Str$(BgtESum#))
'      Mid$(ToPrint$, Col2 - 2) = Using$(TotalFmt$, Str$(MTDESum#))
'      Mid$(ToPrint$, Col3 - 2) = Using$(TotalFmt$, Str$(YTDESum#))
'      Mid$(ToPrint$, Col4) = Using$(CommaFmt$, Str$(Variance#))
'      Mid$(ToPrint$, Col5) = Pct$
      '--Reset vars
      FundRevMTD# = MTDSum#
      MTDSum# = 0
    Case 2
      Variance# = YTDESum# - BgtESum#
'      Mid$(ToPrint$, Col1) = Using$(BgtFmt$, Str$(BgtESum#))
'      Mid$(ToPrint$, Col2 - 2) = Using$(TotalFmt$, Str$(EncSum#)) '**
'      Mid$(ToPrint$, Col3 - 2) = Using$(TotalFmt$, Str$(YTDESum#))
'      Mid$(ToPrint$, Col4) = Using$(CommaFmt$, Str$(Variance#))
'      Mid$(ToPrint$, Col5) = Pct$
    End Select
  '  Print #PRNFile, RTrim$(ToPrint$)

  End If

  '--print summary totals
  If Mid$(fpcboReportOn.Text, 1, 1) = "B" Then
    ToPrint$ = ""
    'Print #PRNFile, ToPrint$
    'Mid$(ToPrint$, 1) = "Revenues over Expenditures"
    Select Case WhichReport
    Case 1, 3
      'Variance# = YTDESum# - BgtESum#
'      Mid$(ToPrint$, Col1) = Using$(BgtFmt$, Str$(BgtRSum# - BgtESum#))
'      Mid$(ToPrint$, Col2 - 2) = Using$(TotalFmt$, Str$(MTDRSum# - MTDESum#))
'      Mid$(ToPrint$, Col3 - 2) = Using$(TotalFmt$, Str$(YTDRSum# - YTDESum#))
      'MID$(ToPrint$, Col4) = FUsing$(STR$(Variance#), CommaFmt$)
      'MID$(ToPrint$, Col5) = Pct$
    Case 2
      'Variance# = YTDESum# - BgtESum#
'      Mid$(ToPrint$, Col1) = Using$(BgtFmt$, Str$(BgtRSum# - BgtESum#))
      'MID$(ToPrint$, Col2) = FUsing$(" 0", TotalFmt$)
'      Mid$(ToPrint$, Col3 - 2) = Using$(TotalFmt$, Str$(YTDRSum# - YTDESum#))
      'MID$(ToPrint$, Col4) = FUsing$(STR$(Variance#), CommaFmt$)
      'MID$(ToPrint$, Col5) = Pct$
    End Select
'    Print #PRNFile, RTrim$(ToPrint$)
  End If
  '=================

'  Print #PRNFile, FF$
  Close
  If RCnt <> 0 Or ECnt <> 0 Then
  '====End Report Processing
   
   ARptDepBudVAct.BgtTot = Using$(BgtFmt$, Str$(BgtRSum# - BgtESum#))
   If rpt = 1 Or rpt = 3 Then
     ARptDepBudVAct.FundTot1 = Using$(TotalFmt$, Str$(MTDRSum# - MTDESum#))
   Else
     ARptDepBudVAct.FundTot1 = 0
   End If
   ARptDepBudVAct.FundTot2 = Using$(TotalFmt$, Str$(YTDRSum# - YTDESum#))
   ARptDepBudVAct.rptopt = rpt
   ARptDepBudVAct.rptnum = WhichReport
   ARptDepBudVAct.labelEnd.Caption = ("Ending Date: " + txtDate)
   ARptDepBudVAct.txtDate = Now
   ARptDepBudVAct.txtTown = GLUserName$
   ARptDepBudVAct.GetName ReportFile$
   ARptDepBudVAct.startrpt
  Else
    MsgBox "No Information To Report From Selections Made.", vbOKOnly, "No Info"
  End If
  'ViewPrint ReportFile$, RptTitle$, True
  Close

  'KillFile ReportFile$
  'End Report Printing========================================================
Exit Sub


'PrintDeptPageHdr:
'  PageNum = PageNum + 1
'  Print #PRNFile, GLUserName; Tab(43); "Run Date: " + Date$; "       Page: "; PageNum
'  Print #PRNFile, DeptName$ + " Department" + " " + RptTitle$
'  Print #PRNFile, "Period Ending: " + txtDate
'  Print #PRNFile,
'  Print #PRNFile, Desc$(1)
'  Print #PRNFile, String$(Len(Desc$(1)), "-")
'  Linecnt = 6
'Return


'PrintSummaryLines2:
'    '--Print summary lines
'    ToPrint$ = Space$(96)
'    Mid$(ToPrint$, Col1) = BSumLine$
'    Mid$(ToPrint$, Col2 - 2) = SumLine$
'    Mid$(ToPrint$, Col3 - 2) = SumLine$
'    Mid$(ToPrint$, Col4 - 2) = SumLine$
'    Mid$(ToPrint$, Col5) = PSumLine$
'    Print #PRNFile, RTrim$(ToPrint$)
'    Linecnt = Linecnt + 1
'Return

''GotDeptErr:
''  ErrorCode$ = Str$(Err)
''  Select Case Err
''    Case 70
''      Cls
''      QPrintRC "Access Denied. Try again later.", 12, 1, 12
''      QPrintRC "Press any key to continue.", 14, 1, 11
''    Case Else
''      Cls
''      QPrintRC "An Error has halted the system, Error Code: " + ErrorCode$, 12
''      QPrintRC "Press any key exit.", 12, 1, 14
''   End Select

   Exit Sub

Return
CalcAcctBal:
  MTDBal# = 0
  YTDBal# = 0
  NextTr& = GLAcct.FrstTran 'get the first trans for this acct

  Do Until NextTr& = 0    'keep going 'til we run out

    Get TransFileNum, NextTr&, GLTrans

    '--Get MTD Account Balance if necessary

    If GLTrans.TRDATE >= FYStartDate And GLTrans.TRDATE <= EndDate Then
    If GetMonth Then
      'Lookhere change num2month to reflect year & month
      d$ = Format(DateAdd("d", GLTrans.TRDATE, "12-31-1979"), "mm/dd/yyyy")
      TransMonth = Right$(d$, 2) + Left$(d$, 2)
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
      'Lookhere change num2month to reflect year & month
      'D$ = Num2Date(Trans.TrDate)
      'TransMonth = Num2Month%(Trans.TrDate)
      'IF TransMonth = RptMonth THEN

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
  'PUT AcctFileNum, AcctIdx.RecNum, Acct 'dupe acct problem

Return
CancelExit:
  Exit Sub
End Sub

Private Function GetDeptOffsets$()
'---------------------------------------------------
'Gets the Dept Code from the setup file.  The
'code determines where to look for the dept code in
'the acct code
'---------------------------------------------------
  Dim SetupFile As Integer
   OpenSetupFile SetupFile
   Get SetupFile, 1, GLSetup

   GetDeptOffsets$ = GLSetup.DeptCode

   Close SetupFile


End Function


Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboFund1.SetFocus
  End If
End Sub
Private Sub PrintDeptBgtAct2()
  Dim CommaFmt As String, TotalFmt As String, RunBalFmt As String
  Dim SumLine As String, BgtFmt As String, BSumLine As String, PSumLine As String
  Dim DivLine As String, DivLine2 As String, FF As String, RptTitle As String
  Dim MaxLines As Integer, Col1 As Integer, Col2 As Integer, Col3 As Integer
  Dim Col4 As Integer, Col5 As Integer, EndDate As Integer, PRNFile As Integer
  Dim M As String, HollyFlag As Boolean, Pitch12 As String, ThisFund As String
  Dim DoingDetail As Boolean, SubTotalRevenues As Boolean, DeptOnNewPage As Boolean
  Dim WhichReport As Integer, GetMonth As Boolean, GetQtr As Boolean
  Dim RptMonth As String, ReportFile As String, FundCode As String
  Dim FundIdxFile As Integer, NumFunds As Integer, Acct As Integer
  Dim AcctIdxFileNum As Integer, NumGLAccts As Integer, FundName As String
  Dim AcctFileNum As Integer, NumGLAcctRecs As Integer, Rec As Integer
  Dim TransFileNum As Integer, NumTrans As Long, NextTr As Long, AcctNum As String
  Dim BGTBal As Double, YTDBal As Double, MTDBal As Double, UsingFund As Boolean
  Dim ECnt As Integer, RCnt As Integer, d As String, TransMonth As String
  Dim InThisQtr As Boolean, FirstTime As Boolean, Fund As Integer, FundRec As Integer
  Dim Dept As String, LastDept As String, LastDeptName As String, cnt As Integer
  Dim Account As String, BudgetAmt As Double, DeptRecNum As Integer, DeptName As String
  Dim Pct As String, Variance As Double, ToPrint As String, Linecnt As Integer
  Dim MTDSum As Double, BgtSum As Double, YTDSum As Double, DeptBgtSum As Double
  Dim DeptYTDSum As Double, DeptENCSum As Double, DeptMTDSum As Double
  Dim FundRevMTD As Double, FundRevBgt As Double, FundRevYTD As Double
  Dim EncSum As Double, FundExpMTD As Double, FundExpBgt As Double
  Dim FundExpYTD As Double, FundEncYTD As Double, EncBal As Double
  Dim DeptSummary As String, DeptTypeCode As String, DeptType As Integer
  Dim DoingRevenues As Boolean, DoingExp As Boolean, RptDept As String
  Dim DeptExp As String, DeptRev As String, MTDRSum As Double, Newrp As String
  Dim BgtRSum As Double, YTDRSum As Double, MTDESum As Double
  Dim BgtESum As Double, YTDESum As Double, PageNum As Integer
  Dim GetEnc As Boolean, IncAcct As Boolean
 
  '''GetFundCodes FirstFund$, LastFund$
  CommaFmt$ = "###,###,###.##"  'format takes 14 chars
  TotalFmt$ = "#,###,###,###.##" 'format takes 16 chars
  RunBalFmt$ = "##########.##"
  SumLine$ = String$(16, "-")   'column summary line

  BgtFmt$ = "###,###,###.##"         'format takes 11 chars
  BSumLine$ = String$(11, "-")  'summary line for budget columns
  PSumLine$ = "----"            'summary line for Pct columns
  DivLine$ = String$(96, "-")   'dashed line
  DivLine2$ = String$(96, "=")  'Double Line

  'DivLine$ = String$(79, "-")   'dashed line
  'DivLine2$ = String$(79, "=")  'Double Line
  FF$ = Chr$(12)
  ReDim Desc$(1)
  RptTitle$ = "Budget vs. Actual"
  MaxLines = 55
  If InStr(UCase$(GLUserName), "HOLLY SPR") > 0 Then
    HollyFlag = True
    Pitch12$ = Chr$(27) + Chr$(38) + Chr$(107) + Chr$(52) + Chr$(83)
  End If

  '--Column offsets for printing amounts
'''  Col1 = 38
'''  Col2 = 49
'''  Col3 = 63
'''  Col4 = 77
'''  Col5 = 90
  
  Col1 = 28    'Budget
  Col2 = 44   'Month or Actual YTD
  Col3 = 61    'YTD or Var
  Col4 = 78    'Enc
  Col5 = 93    'Pct
  If fpcboIncAcct.ListIndex = 1 Then
    IncAcct = False
  Else
    IncAcct = True
  End If

  EndDate = DateDiff("d", "12/31/1979", txtDate)

  M$ = Right$(txtDate, 2) + Left$(txtDate, 2)
  FrmShowPctComp.Label1 = "Printing Departmental Budget vs Actual Report"
  FrmShowPctComp.Show , Me
  DoEvents
  DeActivateControls frmPrnDeptBudAct, True


  DeptTypeCode$ = GetDeptOffsets$
  DeptType = Val(DeptTypeCode$)
  If DeptType = 0 Then DeptType = 1

  WhichReport = Val(fpcboRepNum.Text)

  Select Case Mid$(fpcboReportOn.Text, 1, 1)
    Case "B"
      DoingRevenues = True
      DoingExp = True
    Case "R"
      DoingRevenues = True
      DoingExp = False
    Case "E"
      DoingRevenues = False
      DoingExp = True
  End Select
  fpcboDepartment.Col = 1
  RptDept$ = QPTrim$(fpcboDepartment.ColText)

  DeptRecNum = FindDept(RptDept$)
  If DeptRecNum > 0 Then
    DeptName$ = QPTrim$(GetDeptTitle$(DeptRecNum))
  Else
    DeptName$ = "Undefined: " + RptDept$
  End If

  Select Case WhichReport
  Case 1        'Bgt, Month, YTD
    GetMonth = True
    GetQtr = False
    GetEnc = False
    RptMonth = M$
''''''''''''''''123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456''''For 96
    Desc$(1) = "Description                      Budget          Month             YTD           Variance    Pct"
    MTDBal# = 0
  Case 2        'Bgt, Enc, YTD, Variance
    GetMonth = False
    GetQtr = False
    GetEnc = True
    Desc$(1) = "Description                      Budget          Encumb            YTD           Variance    Pct"
  Case 3        'Bgt, QTR, YTD, Variance
    GetMonth = False
    GetQtr = True
    GetEnc = False
    Desc$(1) = "Description                      Budget           QTD              YTD           Variance    Pct"
  End Select
  If GetEnc = True Then
        'if need enc totals do this here
    FixPOEncumbRpt EndDate, FYStartDate
  End If

  GetFYDates FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate
  Newrp = "BgtActD"
  GetRPTName Newrp
  ReportFile$ = Newrp
  'ReportFile$ = Unique$(Path$)

  'P17$ = CHR$(27) + "(s17H"
  PRNFile = FreeFile
  Open ReportFile$ For Output As #PRNFile
  'PRINT #PrnFile, P17$;
  If HollyFlag Then
    Print #PRNFile, Pitch12$;
  End If

  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum, NumGLAcctRecs
  OpenTransFile TransFileNum, NumTrans&
  OpenFundIdx FundIdxFile, NumFunds

  If DoingRevenues Then
    ReDim RevAccts%(1 To NumGLAccts)              'Holds all rev acct record n
  End If
  ReDim ExpAccts%(1 To NumGLAccts)              'Holds all exp acct record num
  ReDim FundList(1) As String                            'List of all active Funds
  GetFundList FundList$(), NumFunds

  '--Build a list of revenue and exp accounts
  For Acct = 1 To NumGLAccts

    Get AcctIdxFileNum, Acct, GLAcctidx
    Get AcctFileNum, GLAcctidx.RecNum, GLAcct

    '--Find what fund this account is in
    FundCode$ = Left$(GLAcct.Num, GLFundLen)

    '--See if the account is in a fund we want to see
    If FundCode$ >= StartFund$ And FundCode$ <= EndFund$ Then

      '--Account is in fund, check to see if its proper type
      '--We want only revenue or expenditure accounts
      If GLAcct.Typ = "R" Or GLAcct.Typ = "E" Then
        Select Case GLAcct.Typ
          Case "E"
            If DoingExp Then
              Select Case DeptType
                Case 1
                  DeptExp$ = Mid$(GLAcct.Num, GLFundLen + 2, GLAcctLen)
                Case 2  'Appalachain Dist Health
                  DeptExp$ = Mid$(GLAcct.Num, GLFundLen + 3, GLAcctLen - 1)
                Case 3  'Oklahoma
                  DeptExp$ = Mid$(GLAcct.Num, GLFundLen + 2, GLAcctLen)
              End Select
              If DeptExp$ = RptDept$ Then
                ECnt = ECnt + 1
                ExpAccts%(ECnt) = GLAcctidx.RecNum
              End If
            End If
          Case "R"
            If DoingRevenues Then
              Select Case DeptType
                Case 1  'Normal
                  DeptRev$ = Mid$(GLAcct.Num, GLFundLen + 3, GLAcctLen - 1)
                Case 2  'Appalachain Dist Health
                  DeptRev$ = Mid$(GLAcct.Num, GLFundLen + 3, GLAcctLen - 1)
                Case 3  'Oklahoma
                  DeptRev$ = Mid$(GLAcct.Num, GLFundLen + 2, GLAcctLen)
              End Select
              If DeptRev$ = RptDept$ Then
                RCnt = RCnt + 1
                RevAccts%(RCnt) = GLAcctidx.RecNum
              End If
            End If
        End Select
      End If    '--test for rev or exp accts
    End If      '--End of acct in fund range test
          FrmShowPctComp.ShowPctComp Acct, NumGLAccts
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      ActivateControls frmPrnDeptBudAct, True
      Unload FrmShowPctComp
      GoTo CancelExit
    End If
  Next          'Process next account


  ActivateControls frmPrnDeptBudAct, True

  '--Now write the report to file.

  GoSub PrintDeptPageHdr

  '--Search thru list of revenue accounts
  If DoingRevenues Then
    For cnt = 1 To RCnt
      Rec = RevAccts%(cnt)
      Get AcctFileNum, Rec, GLAcct

      GoSub CalcAcctBal
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

      Pct$ = GetPct$(GLAcct.YTD, BudgetAmt#)
      Variance# = GLAcct.YTD - BudgetAmt#   'Acct.Bgt

      ToPrint$ = Space$(96)
      LSet ToPrint$ = Left$(Account$, 36)
      Mid$(ToPrint$, Col1) = Using$(BgtFmt$, Str$(BudgetAmt#))   'changed
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
        GoSub PrintDeptPageHdr
      End If

      If GetMonth Or GetQtr Then 'changed
        MTDRSum# = MTDRSum# + GLAcct.MTD
      End If

      BgtRSum# = BgtRSum# + BudgetAmt# 'Acct.Bgt
      YTDRSum# = YTDRSum# + GLAcct.YTD

    Next      'Revenue Acct

    GoSub PrintSummaryLines2
    ToPrint$ = Space$(96)
    LSet ToPrint$ = "Total Revenues"
    Pct$ = GetPct$(YTDRSum#, BgtRSum#)
    Select Case WhichReport
    Case 1, 3
      Variance# = YTDRSum# - BgtRSum#
      Mid$(ToPrint$, Col1 - 2) = Using$(TotalFmt$, Str$(BgtRSum#))
      Mid$(ToPrint$, Col2 - 2) = Using$(TotalFmt$, Str$(MTDRSum#))
      Mid$(ToPrint$, Col3 - 2) = Using$(TotalFmt$, Str$(YTDRSum#))
      Mid$(ToPrint$, Col4) = Using$(CommaFmt$, Str$(Variance#))
      Mid$(ToPrint$, Col5) = Pct$
      '--Reset vars
      FundRevMTD# = MTDSum#
      MTDSum# = 0
    Case 2
      Variance# = YTDSum# - BgtSum#
      Mid$(ToPrint$, Col1 - 2) = Using$(TotalFmt$, Str$(BgtRSum#))
      Mid$(ToPrint$, Col2 - 2) = Using$(TotalFmt$, " 0")
      Mid$(ToPrint$, Col3 - 2) = Using$(TotalFmt$, Str$(YTDRSum#))
      Mid$(ToPrint$, Col4) = Using$(CommaFmt$, Str$(Variance#))
      Mid$(ToPrint$, Col5) = Pct$
    End Select
    Print #PRNFile, RTrim$(ToPrint$)
    Linecnt = Linecnt + 1
    If Linecnt > MaxLines Then
      Print #PRNFile, FF$
      GoSub PrintDeptPageHdr
    End If

    Print #PRNFile,

  End If 'doing revenues

  '--Search exp accounts list for accounts in this fund
  If DoingExp Then
    For cnt = 1 To ECnt
      Rec = ExpAccts%(cnt)
      Get AcctFileNum, Rec, GLAcct
      GoSub CalcAcctBal
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

     ToPrint$ = Space$(96)
     'Pct$ = GetPct$(Acct.Encumb + Acct.YTD, BudgetAmt#) 'Acct.Bgt)
     LSet ToPrint$ = Left$(Account$, 36)
     Select Case WhichReport
     Case 1, 3
       Pct$ = GetPct$(GLAcct.YTD, BudgetAmt#)  'Acct.Bgt)
       Variance# = BudgetAmt# - GLAcct.YTD
       Mid$(ToPrint$, Col1) = Using$(BgtFmt$, Str$(BudgetAmt#))
       Mid$(ToPrint$, Col2) = Using$(CommaFmt$, Str$(GLAcct.MTD))
       Mid$(ToPrint$, Col3) = Using$(CommaFmt$, Str$(GLAcct.YTD))
       Mid$(ToPrint$, Col4) = Using$(CommaFmt$, Str$(Variance#))
       Mid$(ToPrint$, Col5) = Pct$
     Case 2
       Pct$ = GetPct$(GLAcct.Encumb + GLAcct.YTD, BudgetAmt#) 'Acct.Bgt)
       Variance# = BudgetAmt# - GLAcct.Encumb - GLAcct.YTD
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
        GoSub PrintDeptPageHdr
      End If

      If GetMonth Or GetQtr Then 'changed
        MTDESum# = MTDESum# + GLAcct.MTD
      End If
      BgtESum# = BgtESum# + BudgetAmt#
      YTDESum# = YTDESum# + GLAcct.YTD
      EncSum# = EncSum# + GLAcct.Encumb

    Next      'Exp Acct

    GoSub PrintSummaryLines2

    ToPrint$ = Space$(96)
    LSet ToPrint$ = "Total Expenditures"
    Pct$ = GetPct$(YTDESum#, BgtESum#)
    'Pct$ = GetPct$(YTDSum#, BgtSum#)
    Select Case WhichReport
    Case 1, 3
      Variance# = YTDESum# - BgtESum#
      Mid$(ToPrint$, Col1 - 2) = Using$(TotalFmt$, Str$(BgtESum#))
      Mid$(ToPrint$, Col2 - 2) = Using$(TotalFmt$, Str$(MTDESum#))
      Mid$(ToPrint$, Col3 - 2) = Using$(TotalFmt$, Str$(YTDESum#))
      Mid$(ToPrint$, Col4) = Using$(CommaFmt$, Str$(Variance#))
      Mid$(ToPrint$, Col5) = Pct$
      '--Reset vars
      FundRevMTD# = MTDSum#
      MTDSum# = 0
    Case 2
      Variance# = YTDESum# - BgtESum#
      Mid$(ToPrint$, Col1 - 2) = Using$(TotalFmt$, Str$(BgtESum#))
      Mid$(ToPrint$, Col2 - 2) = Using$(TotalFmt$, Str$(EncSum#)) '**
      Mid$(ToPrint$, Col3 - 2) = Using$(TotalFmt$, Str$(YTDESum#))
      Mid$(ToPrint$, Col4) = Using$(CommaFmt$, Str$(Variance#))
      Mid$(ToPrint$, Col5) = Pct$
    End Select
    Print #PRNFile, RTrim$(ToPrint$)

  End If
  
  '--print summary totals
  If Mid$(fpcboReportOn.Text, 1, 1) = "B" Then
    LSet ToPrint$ = ""
    Print #PRNFile, ToPrint$
    Mid$(ToPrint$, 1) = "Revenues over Expenditures"
    Select Case WhichReport
    Case 1, 3
      'Variance# = YTDESum# - BgtESum#
      Mid$(ToPrint$, Col1 - 2) = Using$(TotalFmt$, Str$(BgtRSum# - BgtESum#))
      Mid$(ToPrint$, Col2 - 2) = Using$(TotalFmt$, Str$(MTDRSum# - MTDESum#))
      Mid$(ToPrint$, Col3 - 2) = Using$(TotalFmt$, Str$(YTDRSum# - YTDESum#))
      'MID$(ToPrint$, Col4) = FUsing$(STR$(Variance#), CommaFmt$)
      'MID$(ToPrint$, Col5) = Pct$
    Case 2
      'Variance# = YTDESum# - BgtESum#
      Mid$(ToPrint$, Col1 - 2) = Using$(TotalFmt$, Str$(BgtRSum# - BgtESum#))
      'MID$(ToPrint$, Col2) = FUsing$(" 0", TotalFmt$)
      Mid$(ToPrint$, Col3 - 2) = Using$(TotalFmt$, Str$(YTDRSum# - YTDESum#))
      'MID$(ToPrint$, Col4) = FUsing$(STR$(Variance#), CommaFmt$)
      'MID$(ToPrint$, Col5) = Pct$
    End Select
    Print #PRNFile, RTrim$(ToPrint$)
  End If
  '=================

  Print #PRNFile, FF$
  Close

  '====End Report Processing

  ViewPrint ReportFile$, RptTitle$, True
  Close

  KillFile ReportFile$
  'End Report Printing========================================================
Exit Sub


PrintDeptPageHdr:
  PageNum = PageNum + 1
  Print #PRNFile, GLUserName; Tab(43); "Run Date: " + Date$; "       Page: "; PageNum
  Print #PRNFile, DeptName$ + " Department" + " " + RptTitle$
  Print #PRNFile, "Period Ending: " + txtDate
  Print #PRNFile,
  Print #PRNFile, Desc$(1)
  Print #PRNFile, String$(Len(Desc$(1)), "-")
  Linecnt = 6
Return


PrintSummaryLines2:
    '--Print summary lines
    ToPrint$ = Space$(96)
    Mid$(ToPrint$, Col1 - 2) = BSumLine$
    Mid$(ToPrint$, Col2 - 2) = SumLine$
    Mid$(ToPrint$, Col3 - 2) = SumLine$
    Mid$(ToPrint$, Col4 - 2) = SumLine$
    Mid$(ToPrint$, Col5) = PSumLine$
    Print #PRNFile, RTrim$(ToPrint$)
    Linecnt = Linecnt + 1
Return

''GotDeptErr:
''  ErrorCode$ = Str$(Err)
''  Select Case Err
''    Case 70
''      Cls
''      QPrintRC "Access Denied. Try again later.", 12, 1, 12
''      QPrintRC "Press any key to continue.", 14, 1, 11
''    Case Else
''      Cls
''      QPrintRC "An Error has halted the system, Error Code: " + ErrorCode$, 12
''      QPrintRC "Press any key exit.", 12, 1, 14
''   End Select

   Exit Sub

Return
CalcAcctBal:
  MTDBal# = 0
  YTDBal# = 0
  NextTr& = GLAcct.FrstTran 'get the first trans for this acct

  Do Until NextTr& = 0    'keep going 'til we run out

    Get TransFileNum, NextTr&, GLTrans

    '--Get MTD Account Balance if necessary

    If GLTrans.TRDATE >= FYStartDate And GLTrans.TRDATE <= EndDate Then
    If GetMonth Then
      'Lookhere change num2month to reflect year & month
      d$ = Format(DateAdd("d", GLTrans.TRDATE, "12-31-1979"), "mm/dd/yyyy")
      TransMonth = Right$(d$, 2) + Left$(d$, 2)
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
      'Lookhere change num2month to reflect year & month
      'D$ = Num2Date(Trans.TrDate)
      'TransMonth = Num2Month%(Trans.TrDate)
      'IF TransMonth = RptMonth THEN

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
  'PUT AcctFileNum, AcctIdx.RecNum, Acct 'dupe acct problem

Return
CancelExit:
  Exit Sub
End Sub

