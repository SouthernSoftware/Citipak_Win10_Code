VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmPRDefaultSet 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Dates / Options"
   ClientHeight    =   8568
   ClientLeft      =   36
   ClientTop       =   588
   ClientWidth     =   11580
   FillColor       =   &H00C0C0C0&
   Icon            =   "frmPRDefaultSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11724
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcomboDefYN 
      Height          =   405
      Left            =   6720
      TabIndex        =   3
      Top             =   1970
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
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
      ColDesigner     =   "frmPRDefaultSet.frx":08CA
   End
   Begin LpLib.fpCombo fpcomboFqMnthly 
      Height          =   405
      Left            =   5910
      TabIndex        =   7
      Top             =   3315
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
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
      ColDesigner     =   "frmPRDefaultSet.frx":0BC1
   End
   Begin LpLib.fpCombo fpcomboFqSAnn 
      Height          =   405
      Left            =   8925
      TabIndex        =   9
      Top             =   2895
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
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
      ColDesigner     =   "frmPRDefaultSet.frx":0EB8
   End
   Begin LpLib.fpCombo fpcomboFqQtrly 
      Height          =   405
      Left            =   5910
      TabIndex        =   8
      Top             =   3750
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
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
      ColDesigner     =   "frmPRDefaultSet.frx":11AF
   End
   Begin LpLib.fpCombo fpcomboFqSMnthly 
      Height          =   405
      Left            =   5910
      TabIndex        =   6
      Top             =   2895
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
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
      ColDesigner     =   "frmPRDefaultSet.frx":14A6
   End
   Begin LpLib.fpCombo fpcomboFqWkly 
      Height          =   405
      Left            =   3030
      TabIndex        =   4
      Top             =   2895
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
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
      ColDesigner     =   "frmPRDefaultSet.frx":179D
   End
   Begin LpLib.fpCombo fpcomboFqBiWkly 
      Height          =   405
      Left            =   3030
      TabIndex        =   5
      Top             =   3315
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
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
      ColDesigner     =   "frmPRDefaultSet.frx":1A94
   End
   Begin LpLib.fpCombo fpcomboFqAnn 
      Height          =   405
      Left            =   8925
      TabIndex        =   10
      Top             =   3315
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
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
      ColDesigner     =   "frmPRDefaultSet.frx":1D8B
   End
   Begin LpLib.fpCombo fpcomboEarnYN 
      Height          =   405
      Index           =   1
      Left            =   8685
      TabIndex        =   14
      Top             =   4950
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
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
      ColDesigner     =   "frmPRDefaultSet.frx":2082
   End
   Begin LpLib.fpCombo fpcomboEarnYN 
      Height          =   405
      Index           =   2
      Left            =   8685
      TabIndex        =   15
      Top             =   5385
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
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
      ColDesigner     =   "frmPRDefaultSet.frx":2379
   End
   Begin LpLib.fpCombo fpcomboEarnYN 
      Height          =   405
      Index           =   3
      Left            =   8685
      TabIndex        =   16
      Top             =   5820
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
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
      ColDesigner     =   "frmPRDefaultSet.frx":2670
   End
   Begin LpLib.fpCombo fpcomboEarnYN 
      Height          =   405
      Index           =   4
      Left            =   8685
      TabIndex        =   17
      Top             =   6255
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
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
      ColDesigner     =   "frmPRDefaultSet.frx":2967
   End
   Begin LpLib.fpCombo fpcomboEarnYN 
      Height          =   405
      Index           =   5
      Left            =   8685
      TabIndex        =   18
      Top             =   6685
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
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
      ColDesigner     =   "frmPRDefaultSet.frx":2C5E
   End
   Begin EditLib.fpDateTime fpDateTimeEnd 
      Height          =   375
      Left            =   8205
      TabIndex        =   2
      ToolTipText     =   "Enter this payroll's ending date."
      Top             =   1545
      Width           =   1740
      _Version        =   196608
      _ExtentX        =   3069
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
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   "10-01-2001"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm-dd-yyyy"
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
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   2070
      Left            =   2010
      TabIndex        =   13
      Top             =   4950
      Width           =   3750
      _Version        =   196613
      _ExtentX        =   6625
      _ExtentY        =   3662
      _StockProps     =   64
      ButtonDrawMode  =   31
      ColHeaderDisplay=   0
      ColsFrozen      =   2
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   2
      MaxRows         =   50
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      ShadowColor     =   13684944
      SpreadDesigner  =   "frmPRDefaultSet.frx":2F55
      VisibleCols     =   2
   End
   Begin EditLib.fpText fptxtEarnDesc 
      Height          =   390
      Index           =   1
      Left            =   6480
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   4950
      Width           =   2070
      _Version        =   196608
      _ExtentX        =   3662
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
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtEarnDesc 
      Height          =   390
      Index           =   2
      Left            =   6480
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   5385
      Width           =   2070
      _Version        =   196608
      _ExtentX        =   3662
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
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtEarnDesc 
      Height          =   390
      Index           =   3
      Left            =   6480
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   5820
      Width           =   2070
      _Version        =   196608
      _ExtentX        =   3662
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
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtEarnDesc 
      Height          =   390
      Index           =   4
      Left            =   6480
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   6255
      Width           =   2070
      _Version        =   196608
      _ExtentX        =   3662
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
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtEarnDesc 
      Height          =   390
      Index           =   5
      Left            =   6480
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   6685
      Width           =   2070
      _Version        =   196608
      _ExtentX        =   3662
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
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpDateTime fpDateTimeBeg 
      Height          =   375
      Left            =   4755
      TabIndex        =   1
      ToolTipText     =   "Enter this payroll's starting date."
      Top             =   1545
      Width           =   1740
      _Version        =   196608
      _ExtentX        =   3069
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
      AutoAdvance     =   0   'False
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
      Text            =   "10-01-2001"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm-dd-yyyy"
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
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   528
      Left            =   8172
      TabIndex        =   36
      TabStop         =   0   'False
      ToolTipText     =   "Press ESC to exit this screen."
      Top             =   7488
      Width           =   1524
      _Version        =   131072
      _ExtentX        =   2688
      _ExtentY        =   931
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
      ButtonDesigner  =   "frmPRDefaultSet.frx":34BC
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   528
      Left            =   6288
      TabIndex        =   37
      TabStop         =   0   'False
      ToolTipText     =   "Press to commit this payroll control data to memory for this payroll."
      Top             =   7488
      Width           =   1524
      _Version        =   131072
      _ExtentX        =   2688
      _ExtentY        =   931
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
      ButtonDesigner  =   "frmPRDefaultSet.frx":3699
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Additional Earning"
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
      Height          =   345
      Left            =   6960
      TabIndex        =   35
      Top             =   4350
      Width           =   2265
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Accounts to Use"
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
      Height          =   345
      Left            =   7110
      TabIndex        =   34
      Top             =   4635
      Width           =   1980
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2910
      Left            =   1395
      Top             =   4325
      Width           =   8895
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1845
      Left            =   1395
      Top             =   2490
      Width           =   8895
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1110
      Left            =   1395
      Top             =   1395
      Width           =   8895
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Deductions to Take"
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
      Height          =   345
      Left            =   2655
      TabIndex        =   28
      Top             =   4545
      Width           =   2370
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Annually"
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
      Height          =   345
      Left            =   7770
      TabIndex        =   27
      Top             =   3465
      Width           =   1020
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Semi-Annually"
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
      Height          =   345
      Left            =   7245
      TabIndex        =   26
      Top             =   3030
      Width           =   1545
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Quarterly"
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
      Height          =   345
      Left            =   4800
      TabIndex        =   25
      Top             =   3900
      Width           =   1020
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Monthly"
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
      Height          =   345
      Left            =   4755
      TabIndex        =   24
      Top             =   3465
      Width           =   1020
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Semi-Monthly"
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
      Height          =   345
      Left            =   4230
      TabIndex        =   23
      Top             =   3030
      Width           =   1545
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bi-Weekly"
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
      Height          =   345
      Left            =   1680
      TabIndex        =   22
      Top             =   3465
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Weekly"
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
      Height          =   345
      Left            =   1965
      TabIndex        =   21
      Top             =   3030
      Width           =   930
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Employees to Pay"
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
      Height          =   345
      Left            =   4680
      TabIndex        =   20
      Top             =   2535
      Width           =   2220
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Set Payroll with Defaults?"
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
      Height          =   345
      Left            =   3600
      TabIndex        =   19
      Top             =   2070
      Width           =   3030
   End
   Begin VB.Label Label3 
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
      Height          =   345
      Left            =   6630
      TabIndex        =   12
      Top             =   1635
      Width           =   1500
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pay Period Beginning Date:"
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
      Height          =   345
      Left            =   1590
      TabIndex        =   11
      Top             =   1635
      Width           =   3030
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   975
      Index           =   1
      Left            =   1515
      Top             =   270
      Width           =   8655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Payroll Dates / Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2805
      TabIndex        =   0
      Top             =   510
      Width           =   6015
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   1515
      Top             =   150
      Width           =   8655
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
Attribute VB_Name = "frmPRDefaultSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Private Sub SavePRDefaults()
  Dim DedCodeHandle As Integer
  Dim DedCodeRec As DedCodeRecType
  Dim ErnCodeHandle As Integer
  Dim ErnCodeRec As ErnCodeRecType
  Dim PHandle As Integer
  Dim PPDRec As PeriodDefaultRecType
  Dim THandle As Integer
  Dim TransRec As TransRecType
  Dim x As Integer
  Dim DedCnt As Integer
  Dim ErnCnt As Integer
  Dim TransCnt As Long
  Dim DedAlert As TempDedAlertType
  Dim DHandle As Integer
  Dim NumOfDedAlerts As Integer
  Dim EHandle As Integer
  Dim EarnAlert As TempEarnAlertType
  Dim NumOfEarnAlerts As Integer
  
  On Error GoTo ERRORSTUFF
  
0:
  OpenDedCodeFile DedCodeHandle
  DedCnt = LOF(DedCodeHandle) \ Len(DedCodeRec)
  Close DedCodeHandle
  If DedCnt = 0 Then GoTo NoDeds
1:
  For x = 1 To DedCnt
'    Get DedCodeHandle, x, DedCodeRec
    vaSpread1.Col = 2
    vaSpread1.Row = x
    PPDRec.UseDed(x) = vaSpread1.Text 'save the
    'Y or N option selected by the user...if it is
    'not equal to Y then by default it is N (in case
    'there is nothing selected)
    If vaSpread1.Text <> "Y" Then PPDRec.UseDed(x) = "N"
  Next x
2:
  If Exist(DedAlertName) Then
    OpenDedAlertFile DHandle
    NumOfDedAlerts = LOF(DHandle) / Len(DedAlert)
    If NumOfDedAlerts = 0 Then
      Close DHandle
    Else
      For x = 1 To NumOfDedAlerts
        Get DHandle, x, DedAlert
          vaSpread1.Col = 2
          vaSpread1.Row = DedAlert.Number
          If QPTrim$(vaSpread1.Text) <> "Y" Then
            If MsgBox("A new deduction, " + QPTrim$(DedAlert.DCDESC1) + ", has been added on row # " + CStr(DedAlert.Number) + " but is currently set to 'N'. Do you wish to set this new deduction to 'Y'?", vbYesNo) = vbYes Then
              vaSpread1.Text = "Y"
              PPDRec.UseDed(DedAlert.Number) = vaSpread1.Text
            Else
              MainLog ("User warned that a new deduction, " + QPTrim$(DedAlert.DCDESC1) + ", has been added but is not set to 'Y' in the current payroll default. The user elected to continue setting the default without resetting the new deduction to 'Y'.")
            End If
          End If
      Next x
    End If
    Close DHandle
    KillFile DedAlertName
  End If
3:
'  For x = 1 To DedCnt
''    Get DedCodeHandle, x, DedCodeRec
'    vaSpread1.Col = 2
'    vaSpread1.Row = x
'    PPDRec.UseDed(x) = vaSpread1.Text 'save the
'    'Y or N option selected by the user...if it is
'    'not equal to Y then by default it is N (in case
'    'there is nothing selected)
'    If vaSpread1.Text <> "Y" Then PPDRec.UseDed(x) = "N"
'  Next x
NoDeds:
4:
  If Exist(EarnAlertName) Then
5:
    OpenEarnAlertFile EHandle
6:
    NumOfEarnAlerts = LOF(EHandle) / Len(EarnAlert)
7:
    If NumOfEarnAlerts = 0 Then
      Close EHandle
    Else
8:
      For x = 1 To NumOfEarnAlerts
9:
        Get EHandle, x, EarnAlert
10:
          If QPTrim$(fpcomboEarnYN(EarnAlert.Number).Text) <> "Y" Then
11:
            If MsgBox("A new earnings code, " + QPTrim$(EarnAlert.ERNCODE1) + ", has been added on row # " + CStr(EarnAlert.Number) + " but is currently set to 'N'. Do you wish to set this new earnings code to 'Y'?", vbYesNo) = vbYes Then
              fpcomboEarnYN(EarnAlert.Number).Text = "Y"
            End If
12:
              MainLog ("User warned that a new earnings code, " + QPTrim$(DedAlert.DCDESC1) + ", has been added but is not set to 'Y' in the current payroll default. The user elected to continue setting the default without resetting the new earnings code to 'Y'.")
          End If
      Next x
    End If
    Close EHandle
13:
    KillFile EarnAlertName
  End If
14:
  PPDRec.USEAE1 = fpcomboEarnYN(1).Text
  PPDRec.USEAE2 = fpcomboEarnYN(2).Text
  PPDRec.USEAE3 = fpcomboEarnYN(3).Text
  PPDRec.PERBEG = Date2Num(fpDateTimeBeg.Text)
  PPDRec.PEREND = Date2Num(fpDateTimeEnd.Text)
  PPDRec.USEDEF = fpcomboDefYN.Text
  If fpcomboDefYN.Text <> "Y" Then PPDRec.USEDEF = "N"
  PPDRec.PAYWK = fpcomboFqWkly.Text
  If fpcomboFqWkly.Text <> "Y" Then PPDRec.PAYWK = "N"
  PPDRec.PAYBIWK = fpcomboFqBiWkly.Text
  If fpcomboFqBiWkly.Text <> "Y" Then PPDRec.PAYBIWK = "N"
  PPDRec.PAYSEMIM = fpcomboFqSMnthly.Text
  If fpcomboFqSMnthly.Text <> "Y" Then PPDRec.PAYSEMIM = "N"
  PPDRec.PAYMO = fpcomboFqMnthly.Text
  If fpcomboFqMnthly.Text <> "Y" Then PPDRec.PAYMO = "N"
  PPDRec.PAYQTR = fpcomboFqQtrly.Text
  If fpcomboFqQtrly.Text <> "Y" Then PPDRec.PAYQTR = "N"
  PPDRec.PAYSEMIA = fpcomboFqSAnn.Text
  If fpcomboFqSAnn.Text <> "Y" Then PPDRec.PAYSEMIA = "N"
  PPDRec.PAYANNL = fpcomboFqAnn.Text
  If fpcomboFqAnn.Text <> "Y" Then PPDRec.PAYANNL = "N"
  PPDRec.PACTIVE = -1 '-1 = true
  PPDRec.MACTIVE = 0 ' 0 = false
15:
  OpenPPDefaultFile PHandle
  Put PHandle, 1, PPDRec
  Close PHandle
16:
  If fpcomboDefYN.Text = "N" Then
  'added progress bar for "N" on 8/22
    FrmShowPctComp.Label1 = "Saving New Payroll Defaults"
    FrmShowPctComp.cmdCancel.Visible = False
    FrmShowPctComp.Show , Me
    DoEvents
    EnableCloseButton Me.hwnd, False
    Me.cmdExit.Enabled = False
    Me.cmdSave.Enabled = False
17:
    OpenTransWorkFile THandle
    TransCnt = LOF(THandle) / Len(TransRec)
    For x = 1 To TransCnt
      Get THandle, x, TransRec
      TransRec.TActive = 0
      Put THandle, x, TransRec
      FrmShowPctComp.ShowPctComp x, TransCnt
    Next x
    Close THandle
    
    Unload FrmShowPctComp
    EnableCloseButton Me.hwnd, True
    Me.cmdExit.Enabled = True
    Me.cmdSave.Enabled = True
  ElseIf fpcomboDefYN.Text = "Y" Then
18:
    Call MakeDefaultTransActs
  End If
19:
  MsgBox "Your information has been saved", vbOKOnly
  frmPayrollProcessingMenu.Show
  DoEvents
  Unload frmPRDefaultSet
  
  Exit Sub
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "modNonSplit", "MakeGLIFFileT", Erl)
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

  Dim PPDefaultHandle As Integer
  Dim PPDefaultRec As PeriodDefaultRecType
  Dim changeFlag As Boolean
  Dim DoWhatFlag As SaveChangeOptions1
  Dim DedCodeHandle As Integer
  Dim DedCodeRec As DedCodeRecType
  Dim DedCnt As Integer
  Dim x As Integer
  Dim NumDateBeg As Long
  Dim NumDateEnd As Long
  changeFlag = False
  
  OpenPPDefaultFile PPDefaultHandle
  Get PPDefaultHandle, 1, PPDefaultRec
  Close
  
  OpenDedCodeFile DedCodeHandle
  DedCnt = LOF(DedCodeHandle) \ Len(DedCodeRec)
  Close
  'upon exit we check all fields for changes and if one is found
  'we give the user an alert and the option to save it...the focus
  'if we return to this screen is on wherever the change was made
  For x = 1 To DedCnt
    vaSpread1.Col = 2
    vaSpread1.Row = x
    If Len(QPTrim$(PPDefaultRec.UseDed(x))) = 0 Then GoTo NoPPDFile1
    If vaSpread1.Text <> QPTrim$(PPDefaultRec.UseDed(x)) Then
      changeFlag = True
      vaSpread1.SetActiveCell 2, x
      GoTo changeFound
    End If
NoPPDFile1:
  Next x
  Close DedCodeHandle
  
  If Len(QPTrim$(PPDefaultRec.USEAE1)) = 0 Then GoTo NoPPDFile2
  If QPTrim$(fpcomboEarnYN(1).Text) <> QPTrim$(PPDefaultRec.USEAE1) Then
    changeFlag = True
    fpcomboEarnYN(1).SetFocus
    GoTo changeFound
  End If
NoPPDFile2:
  
  If Len(QPTrim$(PPDefaultRec.USEAE2)) = 0 Then GoTo NoPPDFile3
  If QPTrim$(fpcomboEarnYN(2).Text) <> QPTrim$(PPDefaultRec.USEAE2) Then
    changeFlag = True
    fpcomboEarnYN(2).SetFocus
    GoTo changeFound
  End If
NoPPDFile3:
    
  If Len(QPTrim$(PPDefaultRec.USEAE3)) = 0 Then GoTo NoPPDFile4
  If QPTrim$(fpcomboEarnYN(3).Text) <> QPTrim$(PPDefaultRec.USEAE3) Then
    changeFlag = True
    fpcomboEarnYN(3).SetFocus
    GoTo changeFound
  End If
NoPPDFile4:
'********NOTE: 2 more fpcomboEarnYN will be needed here
'if we add 2 more earnings fields
  If CheckValDate(fpDateTimeBeg.Text) = False Then
    changeFlag = True
    fpDateTimeBeg.SetFocus
    GoTo changeFound
  End If
  
  If PPDefaultRec.PERBEG = 0 Then GoTo NoDate1
  NumDateBeg = Date2Num(fpDateTimeBeg.Text)
  If NumDateBeg <> PPDefaultRec.PERBEG Then
    changeFlag = True
    fpDateTimeBeg.SetFocus
    GoTo changeFound
  End If
NoDate1:
  If CheckValDate(fpDateTimeEnd.Text) = False Then
    changeFlag = True
    fpDateTimeEnd.SetFocus
    GoTo changeFound
  End If
  
  If PPDefaultRec.PEREND = 0 Then GoTo NoDate2
  NumDateEnd = Date2Num(fpDateTimeEnd.Text)
  If NumDateEnd <> PPDefaultRec.PEREND Then
    changeFlag = True
    fpDateTimeEnd.SetFocus
    GoTo changeFound
  End If
NoDate2:
  If QPTrim$(fpcomboDefYN.Text) = "N" And QPTrim$(PPDefaultRec.USEDEF) = "" Then
    GoTo DEFYNOK
  End If
  If QPTrim$(fpcomboDefYN.Text) <> QPTrim$(PPDefaultRec.USEDEF) Then
    changeFlag = True
    fpcomboDefYN.SetFocus
    GoTo changeFound
  End If
DEFYNOK:

  If QPTrim$(fpcomboFqWkly.Text) = "N" And QPTrim$(PPDefaultRec.PAYWK) = "" Then
    GoTo FQWKLYOK
  End If
  If QPTrim$(fpcomboFqWkly.Text) <> QPTrim$(PPDefaultRec.PAYWK) Then
    changeFlag = True
    fpcomboFqWkly.SetFocus
    GoTo changeFound
  End If
FQWKLYOK:

  If QPTrim$(fpcomboFqBiWkly.Text) = "N" And QPTrim$(PPDefaultRec.PAYBIWK) = "" Then
    GoTo FQBIWKLYOK
  End If
  If QPTrim$(fpcomboFqBiWkly.Text) <> QPTrim$(PPDefaultRec.PAYBIWK) Then
    changeFlag = True
    fpcomboFqBiWkly.SetFocus
    GoTo changeFound
  End If
FQBIWKLYOK:

  If QPTrim$(fpcomboFqSMnthly.Text) = "N" And QPTrim$(PPDefaultRec.PAYSEMIM) = "" Then
    GoTo FQSMNTHLYOK
  End If
  If QPTrim$(fpcomboFqSMnthly.Text) <> QPTrim$(PPDefaultRec.PAYSEMIM) Then
    changeFlag = True
    fpcomboFqSMnthly.SetFocus
    GoTo changeFound
  End If
FQSMNTHLYOK:
  
  If QPTrim$(fpcomboFqMnthly.Text) = "N" And QPTrim$(PPDefaultRec.PAYMO) = "" Then
    GoTo FQMNTHLYOK
  End If
  If QPTrim$(fpcomboFqMnthly.Text) <> QPTrim$(PPDefaultRec.PAYMO) Then
    changeFlag = True
    fpcomboFqMnthly.SetFocus
    GoTo changeFound
  End If
FQMNTHLYOK:
  
  If QPTrim$(fpcomboFqQtrly.Text) = "N" And QPTrim$(PPDefaultRec.PAYQTR) = "" Then
    GoTo FQQTRLYOK
  End If
  If QPTrim$(fpcomboFqQtrly.Text) <> QPTrim$(PPDefaultRec.PAYQTR) Then
    changeFlag = True
    fpcomboFqQtrly.SetFocus
    GoTo changeFound
  End If
FQQTRLYOK:

  If QPTrim$(fpcomboFqSAnn.Text) = "N" And QPTrim$(PPDefaultRec.PAYSEMIA) = "" Then
    GoTo FQSANNOK
  End If
  If QPTrim$(fpcomboFqSAnn.Text) <> QPTrim$(PPDefaultRec.PAYSEMIA) Then
    changeFlag = True
    fpcomboFqSAnn.SetFocus
    GoTo changeFound
  End If
FQSANNOK:
  
  If QPTrim$(fpcomboFqAnn.Text) = "N" And QPTrim$(PPDefaultRec.PAYANNL) = "" Then
    GoTo FQANNOK
  End If
  If QPTrim$(fpcomboFqAnn.Text) <> QPTrim$(PPDefaultRec.PAYANNL) Then
    changeFlag = True
    fpcomboFqAnn.SetFocus
    GoTo changeFound
  End If
FQANNOK:

changeFound:
  If changeFlag = True Then
    DoWhatFlag = PromptSaveChanges(Me)
    Select Case DoWhatFlag
    Case SaveChangeOptions1.scoSaveChanges
      Call cmdSave_Click
    Case SaveChangeOptions1.scoReviewChanges
      Exit Sub
    Case SaveChangeOptions1.scoAbandonChanges
      frmPayrollProcessingMenu.Show
      DoEvents
      Unload frmPRDefaultSet
    End Select
  Else 'no changes found so exit is OK
    frmPayrollProcessingMenu.Show
    DoEvents
    Unload frmPRDefaultSet
  End If

End Sub

Private Sub cmdSave_Click()
  Dim PHandle As Integer
  Dim PRDRec As PeriodDefaultRecType
  Dim DoWhatFlag As PRInProgress
  Dim ThisDate As Integer
  Dim EndDate As Integer
  
  On Error GoTo ERRORSTUFF
  
0:
  If Date2Num(fpDateTimeBeg.Text) > Date2Num(fpDateTimeEnd.Text) Then
    MsgBox "The beginning date entered is before the ending date."
    fpDateTimeBeg.SetFocus
    Exit Sub
  End If
1:
  If CheckValDate(fpDateTimeBeg.Text) = False Then
    MsgBox "Please enter a valid date in the 'Pay Period Beginning Date' field."
    fpDateTimeBeg.SetFocus
    Exit Sub
  End If
2:
  If CheckValDate(fpDateTimeEnd.Text) = False Then
    fpDateTimeEnd.Value = ""
    MsgBox "Please enter a valid date in the 'Ending Date' field."
    fpDateTimeEnd.SetFocus
    Exit Sub
  End If
3:
  ThisDate = Date2Num(Date)
  EndDate = Date2Num(fpDateTimeEnd.Text)
  If Abs(EndDate - ThisDate) >= 60 Then
    If MsgBox("The end date is more than 60 days from today's date. If you wish to edit this date then press Yes.", vbYesNo) = vbYes Then
      Close
      fpDateTimeEnd.SetFocus
      Exit Sub
    Else
      MainLog "User warned that the end date " + fpDateTimeEnd.Text + " is over 60 days away from today's date " + CStr(Date) + " and elected to continue anyway."
    End If
  End If
4:
  If QPTrim$(fpcomboDefYN.Text) = "" Then
    MsgBox "Please make a choice in the 'Set Payroll with Defaults?' field."
    fpcomboDefYN.SetFocus
    Exit Sub
  End If
  
5:
  OpenPPDefaultFile PHandle
  Get PHandle, 1, PRDRec
  Close PHandle
  
6:
  If PRDRec.PACTIVE = -1 Then 'there is already a
  'prdefault in progress so alert the user
  'and force him to decide what to do from here
7:
    DoWhatFlag = PromptPRInProgress(Me)
    Select Case DoWhatFlag
    Case PRInProgress.pripEscape: 'return to screen
      Exit Sub
    Case PRInProgress.pripSave: 'clear current default
    'settings and resave screen settings
8:
      Call MakeTransInactive
      'added ther next line on 8/26
9:
      Call CheckForBadWHPct 'look thru files to see if
      'all salaried employee's withholding percentages
      'equal 100% and if not then alert the user
      'added ther next line on 8/26
10:
      Call DeActivateControls 'don't want the user to
      'be able to Terminate payroll in the middle of
      'saving the defaults
11:
      Call SavePRDefaults
12:
      Call ActivateControls 'added 8/26
    Case Else:
    End Select
  Else
    
13:
    Call CheckForBadWHPct 'look thru files to see if
    'all salaried employee's withholding percentages
    'equal 100% and if not then alert the user
14:
    Call DeActivateControls 'don't want the user to
    'be able to Terminate payroll in the middle of
    'saving the defaults
15:
    Call SavePRDefaults
16:
    Call ActivateControls
  End If
  
17:
  If Exist("TEMPIF.DAT") Then 'saving here means that if registers
  'were already run then it would still be out there and a user could
  'then go straight to print checks without rerunning registers
    KillFile "TEMPIF.DAT"
  End If
  MainLog ("Payroll Defaults saved.")
  Exit Sub
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "modNonSplit", "MakeGLIFFileT", Erl)
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
      SendKeys "%X"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
      Call cmdSave_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Call FixSpread
  Call LoadPRDefaultFile
  Me.HelpContextID = hlpSetPayPeriod
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Sub LoadPRDefaultFile()
  Dim PPDefaultHandle As Integer
  Dim PPDefaultRec As PeriodDefaultRecType
  Dim DedCodeHandle As Integer
  Dim DedCodeRec As DedCodeRecType
  Dim ErnCodeHandle As Integer
  Dim ErnCodeRec As ErnCodeRecType
  Dim x As Integer
  Dim DedCnt As Integer
  Dim ErnCnt As Integer
  Dim Today As String * 11
'  Date$ = FormatDateTime(Date, vbShortDate)
  Today = Date '$
  
  OpenPPDefaultFile PPDefaultHandle
  Get PPDefaultHandle, 1, PPDefaultRec
  Close
  
  OpenDedCodeFile DedCodeHandle
  DedCnt = LOF(DedCodeHandle) \ Len(DedCodeRec)
  vaSpread1.MaxRows = DedCnt
  For x = 1 To DedCnt
    Get DedCodeHandle, x, DedCodeRec
    vaSpread1.Col = 1
    vaSpread1.Row = x
    vaSpread1.Text = DedCodeRec.DCDESC1
    vaSpread1.Col = 2
    vaSpread1.Row = x
    vaSpread1.Text = PPDefaultRec.UseDed(x)
    If vaSpread1.Text <> "Y" Then vaSpread1.Text = "N"
  Next x
  Close DedCodeHandle
  
  fpcomboEarnYN(1).Text = PPDefaultRec.USEAE1
  fpcomboEarnYN(2).Text = PPDefaultRec.USEAE2
  fpcomboEarnYN(3).Text = PPDefaultRec.USEAE3
  If PPDefaultRec.PERBEG > 0 Then
    fpDateTimeBeg = MakeRegDate(PPDefaultRec.PERBEG)
  Else
    fpDateTimeBeg = Today
  End If
  If PPDefaultRec.PEREND > 0 Then
    fpDateTimeEnd = MakeRegDate(PPDefaultRec.PEREND)
  Else
    fpDateTimeEnd = Today
  End If
  fpcomboDefYN.Text = PPDefaultRec.USEDEF
  If Len(QPTrim$(fpcomboDefYN.Text)) = 0 Then fpcomboDefYN.Text = "N"
  fpcomboDefYN.AddItem "Y"
  fpcomboDefYN.AddItem "N"
  fpcomboFqWkly.Text = PPDefaultRec.PAYWK
  If fpcomboFqWkly.Text = "" Then fpcomboFqWkly.Text = "N"
  fpcomboFqWkly.AddItem "Y"
  fpcomboFqWkly.AddItem "N"
  fpcomboFqBiWkly.Text = PPDefaultRec.PAYBIWK
  If fpcomboFqBiWkly.Text = "" Then fpcomboFqBiWkly.Text = "N"
  fpcomboFqBiWkly.AddItem "Y"
  fpcomboFqBiWkly.AddItem "N"
  fpcomboFqSMnthly.Text = PPDefaultRec.PAYSEMIM
  If fpcomboFqSMnthly.Text = "" Then fpcomboFqSMnthly.Text = "N"
  fpcomboFqSMnthly.AddItem "Y"
  fpcomboFqSMnthly.AddItem "N"
  fpcomboFqMnthly.Text = PPDefaultRec.PAYMO
  If fpcomboFqMnthly.Text = "" Then fpcomboFqMnthly.Text = "N"
  fpcomboFqMnthly.AddItem "Y"
  fpcomboFqMnthly.AddItem "N"
  fpcomboFqQtrly.Text = PPDefaultRec.PAYQTR
  If fpcomboFqQtrly.Text = "" Then fpcomboFqQtrly.Text = "N"
  fpcomboFqQtrly.AddItem "Y"
  fpcomboFqQtrly.AddItem "N"
  fpcomboFqSAnn.Text = PPDefaultRec.PAYSEMIA
  If fpcomboFqSAnn.Text = "" Then fpcomboFqSAnn.Text = "N"
  fpcomboFqSAnn.AddItem "Y"
  fpcomboFqSAnn.AddItem "N"
  fpcomboFqAnn.Text = PPDefaultRec.PAYANNL
  If fpcomboFqAnn.Text = "" Then fpcomboFqAnn.Text = "N"
  fpcomboFqAnn.AddItem "Y"
  fpcomboFqAnn.AddItem "N"
  
  OpenErnCodeFile ErnCodeHandle
  ErnCnt = LOF(ErnCodeHandle) \ Len(ErnCodeRec)
  For x = 1 To ErnCnt
    Get ErnCodeHandle, x, ErnCodeRec
    fptxtEarnDesc(x).Text = ErnCodeRec.ERNCODE1
    fpcomboEarnYN(x).AddItem "Y"
    fpcomboEarnYN(x).AddItem "N"
  Next x
  Close ErnCodeHandle
  For x = 1 To ErnCnt
    If Len(QPTrim$(fptxtEarnDesc(x).Text)) > 0 Then
      If fpcomboEarnYN(x).Text <> "Y" Then fpcomboEarnYN(x).Text = "N"
    End If
  Next x
End Sub

Private Sub MakeDefaultTransActs()
  Dim EmpHandle As Integer
  Dim PHandle As Integer
  Dim PPDRec As PeriodDefaultRecType
  Dim NumOfRecs As Long
  Dim RecNo As Long
  Dim PayType$
  Dim x As Integer
  Dim MakeTrans As Boolean
  Dim TransRec As TransRecType
  
  'set up each employee's transaction work file
  'and save it
  
  OpenTransWorkFile TRHandle 'TRHandle is a global
  MakeTrans = False
  OpenPPDefaultFile PHandle
  Get PHandle, 1, PPDRec
  Close PHandle
  OpenEmpData2File EmpHandle
  FrmShowPctComp.Label1 = "Saving New Payroll Defaults"
  FrmShowPctComp.cmdCancel.Visible = False
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdSave.Enabled = False
  NumOfRecs = LOF(EmpHandle) / Len(Emp2Rec(1))
  For x = 1 To NumOfRecs
    Get TRHandle, x, TransRec
    Get EmpHandle, x, Emp2Rec(1)
    PayType$ = UCase$(QPTrim(Emp2Rec(1).EMPPFREQ))
    Select Case PayType$ 'check this employee's pay
    'frequency and if it falls into the "Y" category
    'then we know this employee gets paid in this cycle
    Case "WEEKLY"
      If PPDRec.PAYWK = "Y" Then MakeTrans = True
    Case "BI-WEEKLY"
      If PPDRec.PAYBIWK = "Y" Then MakeTrans = True
    Case "SEMI-MONTHLY"
      If PPDRec.PAYSEMIM = "Y" Then MakeTrans = True
    Case "MONTHLY"
      If PPDRec.PAYMO = "Y" Then MakeTrans = True
    Case "QUARTERLY"
      If PPDRec.PAYQTR = "Y" Then MakeTrans = True
    Case "SEMI-ANNUALLY"
      If PPDRec.PAYSEMIA = "Y" Then MakeTrans = True
    Case "ANNUALLY"
      If PPDRec.PAYANNL = "Y" Then MakeTrans = True
    End Select
    If MakeTrans = True Then
      RecNo = x
      Call CreateEmpTransRecs(RecNo) 'the TRHandle file
      'is saved in CreateEmpTransRecs
      MakeTrans = False
    End If
    FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
  Next x
  Close EmpHandle
  Close TRHandle
  Close
  Unload FrmShowPctComp
  EnableCloseButton Me.hwnd, True
  Me.cmdExit.Enabled = True
  Me.cmdSave.Enabled = True
  
End Sub

Private Function FixSpread()
  Dim COne As Integer
  Dim CTwo As Integer
  Dim CThree As Integer
  Dim CFour As Integer
  Dim CFive As Integer
  Dim CSix As Integer
  '-1 means all rows or all columns....0 means headers
  Select Case ScreenW
    Case 1280
    If Screen.TwipsPerPixelX <> 12 Then
      COne = 16
      coladj = 4
      vaSpread1.FontSize = 18
      vaSpread1.RowHeight(-1) = 22
      vaSpread1.RowHeight(0) = 22
    Else
      COne = 7
      coladj = 3
      vaSpread1.RowHeight(-1) = 21
      vaSpread1.RowHeight(0) = 21
      vaSpread1.FontSize = 16
    End If
    Case 1152
    If Screen.TwipsPerPixelX <> 12 Then
      COne = 9.5
      coladj = 4.5
      vaSpread1.FontSize = 16
      vaSpread1.RowHeight(0) = 18.5
      vaSpread1.RowHeight(-1) = 18.5
    Else
      COne = 5.4
      coladj = 0.8
      vaSpread1.RowHeight(0) = 17
      vaSpread1.RowHeight(-1) = 17
      vaSpread1.FontSize = 14
    End If
    Case 1024
    If Screen.TwipsPerPixelX <> 12 Then
      COne = 7
      coladj = 2#
      vaSpread1.RowHeight(0) = 17.5
      vaSpread1.Font.Size = 12
      vaSpread1.FontBold = True
      vaSpread1.RowHeight(-1) = 17.5
    Else
      COne = 1.4
      coladj = 0.9
    End If
    Case 800
      COne = 1
      coladj = 0.2
      vaSpread1.RowHeight(0) = 15
      vaSpread1.RowHeight(-1) = 15
      vaSpread1.FontSize = 12
    Case Else
       
  End Select
  vaSpread1.ColWidth(1) = vaSpread1.ColWidth(1) + COne
  vaSpread1.ColWidth(2) = vaSpread1.ColWidth(2) + coladj

End Function


Private Sub fpcomboDefYN_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboDefYN.ListDown = True
  End If
  If fpcomboDefYN.ListDown = False Then
    If KeyCode = vbKeyDown Then
'      fpcomboFqWkly.SetFocus
      SendKeys "{Tab}"
      KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
'      fpDateTimeEnd.SetFocus
      SendKeys "+{Tab}"
      KeyCode = 0
    End If
  End If
End Sub

Private Sub fpcomboEarnYN_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  'when a user tabs thru the screen and the focus goes to a
  'combo box then the value can change to a different combo
  'box selection if the user isn't careful...part of this code prevents
  'the selection from changing inadvertantly
  
  If KeyCode = vbKeySpace Then
    fpcomboEarnYN(Index).ListDown = True
  End If
  If fpcomboEarnYN(Index).ListDown = False Then
    If Index = 5 Then
      If KeyCode = vbKeyDown Then
        fpDateTimeBeg.SetFocus
        KeyCode = 0
      ElseIf KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    ElseIf Index <> 5 Then
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      ElseIf KeyCode = vbKeyDown Then
        SendKeys "{Tab}"
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub fpcomboFqAnn_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboFqAnn.ListDown = True
  End If
  If fpcomboFqAnn.ListDown = False Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
      SendKeys "+{Tab}"
      KeyCode = 0
    End If
  End If
End Sub

Private Sub fpcomboFqBiWkly_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboFqBiWkly.ListDown = True
  End If
  If fpcomboFqBiWkly.ListDown = False Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
      SendKeys "+{Tab}"
      KeyCode = 0
    End If
  End If
End Sub

Private Sub fpcomboFqMnthly_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboFqMnthly.ListDown = True
  End If
  If fpcomboFqMnthly.ListDown = False Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
      SendKeys "+{Tab}"
      KeyCode = 0
    End If
  End If
End Sub

Private Sub fpcomboFqQtrly_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboFqQtrly.ListDown = True
  End If
  If fpcomboFqQtrly.ListDown = False Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
      SendKeys "+{Tab}"
      KeyCode = 0
    End If
  End If
End Sub

Private Sub fpcomboFqSAnn_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboFqSAnn.ListDown = True
  End If
  If fpcomboFqSAnn.ListDown = False Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
      SendKeys "+{Tab}"
      KeyCode = 0
    End If
  End If
End Sub

Private Sub fpcomboFqSMnthly_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboFqSMnthly.ListDown = True
  End If
  If fpcomboFqSMnthly.ListDown = False Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
      SendKeys "+{Tab}"
      KeyCode = 0
    End If
  End If
End Sub

Private Sub fpcomboFqWkly_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboFqWkly.ListDown = True
  End If
  If fpcomboFqWkly.ListDown = False Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
      SendKeys "+{Tab}"
      KeyCode = 0
    End If
  End If
End Sub

Private Sub fpDateTimeBeg_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    fpcomboEarnYN(5).SetFocus
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
End Sub

Private Sub vaSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyPageDown Then
    fpcomboEarnYN(1).SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyPageUp Then
    fpcomboFqAnn.SetFocus
    KeyCode = 0
  End If
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub CheckForBadWHPct()
  Dim IdxRecLen As Integer
  Dim IdxFileSize&
  Dim NumOfRecs As Integer
  Dim NHandle As Integer
  Dim TotalPct As Double
  Dim y As Integer
  Dim z As Integer
  Dim x As Integer
  Dim Emp2Rec As EmpData2Type
  Dim EHandle As Integer
  Dim BadFlag As Boolean
  Dim Name As String * 25
  Dim Number As String * 10
  
  'if fpcomboDefYN is "N" then the user will be alerted if he tries
  'to edit the mistaken data so there is no need for an alert at this time
  If QPTrim$(fpcomboDefYN.Text) = "N" Then Exit Sub
  BadFlag = False
  IdxRecLen = 2
  IdxFileSize& = FileSize(PRData + EmpIdxNName)
  NumOfRecs = IdxFileSize& \ IdxRecLen
  If NumOfRecs = 0 Then Exit Sub
  OpenEmpIdxNNameFile NHandle
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  
  For x = 1 To NumOfRecs
    Get NHandle, x, IdxBuff(x) 'load array with employee data
  Next x
  Close NHandle
  
  OpenEmpData2File EHandle

  For x = 1 To 30 'added 9/9/04 because this global array
  'was not being cleared
    EmpInfo(x) = ""
  Next x
  
  z = 1
  
  For x = 1 To NumOfRecs
    TotalPct = 0
    Get EHandle, IdxBuff(x), Emp2Rec
    If Emp2Rec.Deleted = -1 Or Emp2Rec.EMPTDATE <> 0 Then
      GoTo SkipIt
    End If
    If QPTrim$(Emp2Rec.EMPPTYPE) = "Salaried" Then
      For y = 1 To 8 'tally up the w/h totals
        TotalPct = OldRound(TotalPct + Emp2Rec.EDist(y).DAmt) 'added OldRound on 8/20/2003
      Next y
      If TotalPct <> 100 Then 'found a mistake so
      'isolate this employee and load up EmpInfo array
      'with his data
        Name = QPTrim$(Emp2Rec.EmpFName) & " " & QPTrim$(Emp2Rec.EmpLName)
        Number = QPTrim$(Emp2Rec.EmpNo)
        EmpInfo(z) = Name & "  " & Number
        z = z + 1
        BadFlag = True 'tells the program to alert the user
      End If
    End If
SkipIt:
  Next x
  Close EHandle
  If BadFlag = True Then
    frmWarnBadWagePct.Show vbModal
  End If
      
End Sub

'Private Sub MakeTransInactive()
'  Dim TransRec As TransRecType
'  Dim THandle As Integer
'  Dim NumOfRecs As Integer
'  Dim X As Integer
'  '8/20 added progress bar
'  KillFile ("prdata\ChecksPrinted.opn") '10/3/03
'  'without this Killfile the user can resave defaults and
'  'then go directly to post with no warnings.
'  FrmShowPctComp.Label1 = "Clearing Former Payroll Defaults"
'  FrmShowPctComp.cmdCancel.Visible = False
'  FrmShowPctComp.Show , Me
'  DoEvents
'  EnableCloseButton Me.hwnd, False
'  Me.cmdExit.Enabled = False
'  Me.cmdSave.Enabled = False
'  OpenTransWorkFile THandle
'  NumOfRecs = LOF(THandle) \ Len(TransRec)
'  For X = 1 To NumOfRecs
'    Get THandle, X, TransRec
'    TransRec.TActive = False
'    Put THandle, X, TransRec
'    FrmShowPctComp.ShowPctComp X, NumOfRecs
'    If FrmShowPctComp.Out = True Then
'      Close
'      FrmShowPctComp.Out = False
'      EnableCloseButton Me.hwnd, True
'      Me.cmdExit.Enabled = True
'      Me.cmdSave.Enabled = True
'      Exit Sub
'    End If
'  Next X
'  Close THandle
'  Unload FrmShowPctComp
'End Sub
Private Sub DeActivateControls()
  Dim cnt As Integer
  Dim x As Control
  Dim cmdButton As CommandButton

  For cnt = 0 To Me.Count - 1
    Set x = Me.Controls.Item(cnt)
      If TypeOf x Is CommandButton Then
        x.Enabled = False
      End If
  Next cnt
    EnableCloseButton Me.hwnd, False
     
End Sub

Private Sub ActivateControls()
  Dim cmdButton As CommandButton
  Dim x As Control
  Dim cnt As Integer
  
  cmdSave.Enabled = True
  cmdExit.Enabled = True
  
  For cnt = 0 To Me.Count - 1
    Set x = Me.Controls.Item(cnt)
      If TypeOf x Is CommandButton Then
        x.Enabled = True
      End If
  Next cnt
  EnableCloseButton Me.hwnd, True
     
End Sub


