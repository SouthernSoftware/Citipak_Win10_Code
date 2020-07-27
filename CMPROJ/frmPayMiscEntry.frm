VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmPayMiscEntry 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Miscellaneous Payment Entry"
   ClientHeight    =   8865
   ClientLeft      =   3930
   ClientTop       =   1890
   ClientWidth     =   12210
   Icon            =   "frmPayMiscEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboMiscCode 
      Height          =   330
      Index           =   4
      Left            =   5475
      TabIndex        =   16
      Top             =   3480
      Width           =   4575
      _Version        =   196608
      _ExtentX        =   8070
      _ExtentY        =   582
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
      Sorted          =   1
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
      AutoSearchFill  =   0   'False
      AutoSearchFillDelay=   500
      EditMarginLeft  =   20
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmPayMiscEntry.frx":08CA
   End
   Begin LpLib.fpCombo fpcboMiscCode 
      Height          =   330
      Index           =   3
      Left            =   5475
      TabIndex        =   14
      Top             =   3165
      Width           =   4575
      _Version        =   196608
      _ExtentX        =   8070
      _ExtentY        =   582
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
      Sorted          =   1
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
      AutoSearchFill  =   0   'False
      AutoSearchFillDelay=   500
      EditMarginLeft  =   20
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmPayMiscEntry.frx":0CB9
   End
   Begin LpLib.fpCombo fpcboMiscCode 
      Height          =   330
      Index           =   2
      Left            =   5475
      TabIndex        =   12
      Top             =   2850
      Width           =   4575
      _Version        =   196608
      _ExtentX        =   8070
      _ExtentY        =   582
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
      Sorted          =   1
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
      AutoSearchFill  =   0   'False
      AutoSearchFillDelay=   500
      EditMarginLeft  =   20
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmPayMiscEntry.frx":10A8
   End
   Begin LpLib.fpCombo fpcboMiscCode 
      Height          =   330
      Index           =   1
      Left            =   5475
      TabIndex        =   10
      Top             =   2550
      Width           =   4575
      _Version        =   196608
      _ExtentX        =   8070
      _ExtentY        =   582
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
      Sorted          =   1
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
      AutoSearchFill  =   0   'False
      AutoSearchFillDelay=   500
      EditMarginLeft  =   20
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmPayMiscEntry.frx":1497
   End
   Begin LpLib.fpCombo fpcboMiscCode 
      Height          =   330
      Index           =   0
      Left            =   5475
      TabIndex        =   8
      Top             =   2235
      Width           =   4575
      _Version        =   196608
      _ExtentX        =   8070
      _ExtentY        =   582
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
      Sorted          =   1
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
      AutoSearchFill  =   0   'False
      AutoSearchFillDelay=   500
      EditMarginLeft  =   20
      EditMarginTop   =   1
      EditMarginRight =   0
      EditMarginBottom=   3
      ResizeRowToFont =   0   'False
      TextTipMultiLine=   0
      AutoMenu        =   -1  'True
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmPayMiscEntry.frx":1886
   End
   Begin LpLib.fpCombo fpcboTenderType 
      Height          =   345
      Left            =   2910
      TabIndex        =   4
      Top             =   4110
      Width           =   2235
      _Version        =   196608
      _ExtentX        =   3942
      _ExtentY        =   609
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      Text            =   ""
      Columns         =   1
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
      ThreeDInsideStyle=   0
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
      Appearance      =   0
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
      ColDesigner     =   "frmPayMiscEntry.frx":1C75
   End
   Begin EditLib.fpText fptxtCity 
      Height          =   300
      Left            =   1368
      TabIndex        =   2
      Top             =   2808
      Width           =   3924
      _Version        =   196608
      _ExtentX        =   6921
      _ExtentY        =   529
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
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
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
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
   Begin fpBtnAtlLibCtl.fpBtn fpCmdCharge 
      Height          =   375
      Left            =   5445
      TabIndex        =   20
      Top             =   7590
      Width           =   1290
      _Version        =   131072
      _ExtentX        =   2275
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "frmPayMiscEntry.frx":200C
   End
   Begin fpBtnAtlLibCtl.fpBtn fpcmdCheck 
      Height          =   375
      Left            =   4110
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   7590
      Width           =   1260
      _Version        =   131072
      _ExtentX        =   2222
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "frmPayMiscEntry.frx":21E9
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   22
      Top             =   8508
      Width           =   12216
      _ExtentX        =   21537
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7144
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "2:23 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7144
            TextSave        =   "5/14/2018"
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
   Begin EditLib.fpDateTime txtPaymentDate 
      Height          =   324
      Left            =   10080
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1368
      Width           =   1548
      _Version        =   196608
      _ExtentX        =   2730
      _ExtentY        =   572
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
      ControlType     =   1
      Text            =   "10/03/2001"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
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
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtName 
      Height          =   300
      Left            =   1368
      TabIndex        =   0
      Top             =   2220
      Width           =   3924
      _Version        =   196608
      _ExtentX        =   6921
      _ExtentY        =   529
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
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
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
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
   Begin EditLib.fpText fptxtDesc 
      Height          =   348
      Left            =   2160
      TabIndex        =   7
      Top             =   6768
      Width           =   3024
      _Version        =   196608
      _ExtentX        =   5334
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
      AutoCase        =   0
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   19
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
   Begin EditLib.fpText fptxtAddress 
      Height          =   300
      Left            =   1368
      TabIndex        =   1
      Top             =   2520
      Width           =   3924
      _Version        =   196608
      _ExtentX        =   6921
      _ExtentY        =   529
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      ThreeDInsideStyle=   0
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
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
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
   Begin EditLib.fpCurrency fpTotPaid 
      Height          =   312
      Left            =   10080
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   7008
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
      _ExtentY        =   550
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ThreeDInsideStyle=   0
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
      AlignTextH      =   2
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
      ControlType     =   2
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   2
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   "$"
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "999999999.99"
      MinValue        =   "-999999999.99"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   -2147483637
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   0
      Left            =   10080
      TabIndex        =   9
      Top             =   2232
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
      _ExtentY        =   572
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   0   'False
      InvalidColor    =   -2147483637
      InvalidOption   =   2
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   ""
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   1
      Left            =   10080
      TabIndex        =   11
      Top             =   2544
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
      _ExtentY        =   572
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   ""
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   2
      Left            =   10080
      TabIndex        =   13
      Top             =   2856
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
      _ExtentY        =   572
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   ""
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   3
      Left            =   10080
      TabIndex        =   15
      Top             =   3168
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
      _ExtentY        =   572
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   ""
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
   Begin EditLib.fpDoubleSingle fpAmtPaid 
      Height          =   324
      Index           =   4
      Left            =   10080
      TabIndex        =   17
      Top             =   3480
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
      _ExtentY        =   572
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
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483637
      InvalidOption   =   2
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483637
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   ""
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
   Begin EditLib.fpCurrency fpChangeDue 
      Height          =   312
      Left            =   2904
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   5916
      Width           =   2244
      _Version        =   196608
      _ExtentX        =   3958
      _ExtentY        =   550
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ThreeDInsideStyle=   0
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
      AlignTextH      =   2
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
      ControlType     =   2
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   2
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   "$"
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "999999999.99"
      MinValue        =   "-999999999.99"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   -2147483637
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
   Begin EditLib.fpCurrency fpTotReceived 
      Height          =   312
      Left            =   2904
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5340
      Width           =   2244
      _Version        =   196608
      _ExtentX        =   3958
      _ExtentY        =   550
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ThreeDInsideStyle=   0
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
      AlignTextH      =   2
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
      ControlType     =   2
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   2
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   "$"
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "999999999.99"
      MinValue        =   "-999999999.99"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   -2147483637
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
   Begin fpBtnAtlLibCtl.fpBtn fpCmdPost 
      Height          =   375
      Left            =   9390
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   7590
      Width           =   1230
      _Version        =   131072
      _ExtentX        =   2170
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "frmPayMiscEntry.frx":34BB
   End
   Begin fpBtnAtlLibCtl.fpBtn CmdExit 
      Height          =   375
      Left            =   10695
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   7590
      Width           =   1230
      _Version        =   131072
      _ExtentX        =   2170
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "frmPayMiscEntry.frx":3697
   End
   Begin fpBtnAtlLibCtl.fpBtn fpCmdCash 
      Height          =   375
      Left            =   2880
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   7590
      Width           =   1140
      _Version        =   131072
      _ExtentX        =   2011
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "frmPayMiscEntry.frx":3873
   End
   Begin EditLib.fpDoubleSingle fpChkAmt 
      Height          =   324
      Left            =   2904
      TabIndex        =   6
      Top             =   4776
      Width           =   2244
      _Version        =   196608
      _ExtentX        =   3958
      _ExtentY        =   572
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
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
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
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
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
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   ""
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
   Begin EditLib.fpDoubleSingle fpCashAmt 
      Height          =   324
      Left            =   2904
      TabIndex        =   5
      Top             =   4440
      Width           =   2244
      _Version        =   196608
      _ExtentX        =   3958
      _ExtentY        =   572
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
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
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
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
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
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   ""
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
   Begin fpBtnAtlLibCtl.fpBtn fpcmdDrawer 
      Height          =   375
      Left            =   285
      TabIndex        =   28
      Top             =   7590
      Width           =   1245
      _Version        =   131072
      _ExtentX        =   2196
      _ExtentY        =   661
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "frmPayMiscEntry.frx":4B44
   End
   Begin EditLib.fpDoubleSingle fpTAmtOwed 
      Height          =   324
      Left            =   2904
      TabIndex        =   3
      Top             =   3768
      Width           =   2244
      _Version        =   196608
      _ExtentX        =   3958
      _ExtentY        =   572
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
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
      ThreeDInsideShadowColor=   -2147483642
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   0
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
      CaretInsert     =   2
      CaretOverWrite  =   2
      UserEntry       =   1
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
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   "0.00"
      DecimalPlaces   =   2
      DecimalPoint    =   ""
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
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "City,St,Zip:"
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
      Height          =   348
      Left            =   288
      TabIndex        =   52
      Top             =   2832
      Width           =   1008
   End
   Begin VB.Label lblSource 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   10080
      TabIndex        =   51
      Top             =   1080
      Width           =   1560
   End
   Begin VB.Label lblOperator 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   6192
      TabIndex        =   50
      Top             =   1176
      Width           =   732
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
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
      Height          =   324
      Left            =   8640
      TabIndex        =   49
      Top             =   7056
      Width           =   900
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Source:"
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
      Height          =   312
      Left            =   8280
      TabIndex        =   48
      Top             =   1104
      Width           =   1656
   End
   Begin VB.Line Line4 
      X1              =   10056
      X2              =   10056
      Y1              =   1896
      Y2              =   7296
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "Amount Paid"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   10080
      TabIndex        =   47
      Top             =   1920
      Width           =   1788
   End
   Begin VB.Label Label15 
      Caption         =   "Misc Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   312
      Left            =   5472
      TabIndex        =   46
      Top             =   1920
      Width           =   1620
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   2568
      X2              =   5268
      Y1              =   5184
      Y2              =   5184
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
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
      Height          =   372
      Left            =   792
      TabIndex        =   45
      Top             =   6792
      Width           =   1224
   End
   Begin VB.Label Lbl11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Check/Charge Amt Paid:"
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
      Left            =   240
      TabIndex        =   44
      Top             =   4776
      Width           =   2472
   End
   Begin VB.Label lblchange 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Change Due:"
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
      Height          =   372
      Left            =   840
      TabIndex        =   43
      Top             =   5952
      Width           =   1872
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
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
      Height          =   348
      Left            =   288
      TabIndex        =   42
      Top             =   2532
      Width           =   1008
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tender Type:"
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
      Left            =   1128
      TabIndex        =   41
      Top             =   4104
      Width           =   1584
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Amount Paid:"
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
      Left            =   444
      TabIndex        =   40
      Top             =   4440
      Width           =   2268
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Payment Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   216
      TabIndex        =   39
      Top             =   3324
      Width           =   5232
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Received:"
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
      Height          =   312
      Left            =   900
      TabIndex        =   38
      Top             =   5376
      Width           =   1812
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Date:"
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
      Height          =   372
      Index           =   1
      Left            =   8352
      TabIndex        =   37
      Top             =   1440
      Width           =   1584
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Owed:"
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
      Index           =   0
      Left            =   984
      TabIndex        =   36
      Top             =   3792
      Width           =   1728
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
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
      Height          =   348
      Left            =   324
      TabIndex        =   35
      Top             =   2232
      Width           =   972
   End
   Begin VB.Label Label2b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Information:"
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
      Index           =   1
      Left            =   216
      TabIndex        =   34
      Top             =   1920
      Width           =   2160
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operator Number:"
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
      Height          =   312
      Left            =   4296
      TabIndex        =   33
      Top             =   1176
      Width           =   1824
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   456
      Left            =   2580
      Top             =   384
      Width           =   7020
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Miscellaneous Payment Entry"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3558
      TabIndex        =   32
      Top             =   468
      Width           =   5100
   End
   Begin VB.Label Label12 
      Caption         =   "Miscellaneous Desc"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7128
      TabIndex        =   31
      Top             =   1920
      Width           =   2916
   End
   Begin VB.Line Line6 
      X1              =   7104
      X2              =   7104
      Y1              =   1920
      Y2              =   7272
   End
   Begin VB.Line Line7 
      BorderWidth     =   3
      X1              =   5448
      X2              =   11880
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000E&
      BorderWidth     =   3
      FillColor       =   &H8000000E&
      Height          =   5604
      Left            =   216
      Top             =   1800
      Width           =   11772
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000E&
      BorderWidth     =   3
      Height          =   828
      Left            =   210
      Top             =   984
      Width           =   11796
   End
   Begin VB.Shape Shape3 
      Height          =   612
      Left            =   216
      Top             =   7416
      Width           =   11796
   End
   Begin VB.Label lblOperName 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   2280
      TabIndex        =   30
      Top             =   1176
      Width           =   1860
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operator Name:"
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
      Height          =   312
      Left            =   336
      TabIndex        =   29
      Top             =   1176
      Width           =   1824
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   600
      Left            =   2592
      Top             =   264
      Width           =   7020
   End
   Begin VB.Line Line3 
      X1              =   5436
      X2              =   5436
      Y1              =   1824
      Y2              =   7416
   End
End
Attribute VB_Name = "frmPayMiscEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Temp_Class As Resize_Class
Dim Over As clsTextBoxOverRider
Dim EditFlag As Boolean, TempAmtRecv As Double, Answer As Integer
Dim ChkOKFlag As Boolean, BeenDone As Boolean
Dim noreset As Boolean, CmNum As Long, MiscCode As String
Dim Oper As String, PayListRec As Long, RecpPort As String
Dim fromform As Form, toform As Form, codeopt As Integer
Dim DefPayDate As String, RctValidate As Boolean, Biscoe As Boolean
Public Sub Wheretogo(xfrm As Form, tfrm As Form, Optional opt As Integer, Optional DDate As String)
  Set fromform = xfrm
  Set toform = tfrm
  If opt <> 0 Then
    codeopt = opt
  Else
    codeopt = 0
  End If
  If DDate <> "" Then
    DefPayDate = DDate
  End If
End Sub
Private Sub Form_Activate()
  If Not BeenDone Then
    BeenDone = True
    DoEvents
  End If
End Sub

Private Sub cmdExit_Click()
  Chk4Change
  If Answer = 1 Then
    Exit Sub
  ElseIf Answer = 2 Then
    CheckInfo
    If ChkOKFlag Then
      fpCmdPost_Click
    Else
      Exit Sub
    End If
  End If
'  Chk4Change
'  If Answer = 1 Then
'    Exit Sub
'  ElseIf Answer = 2 Then
'    fpCmdPost_Click
'  End If
  Load frmCMPaySource
  CMLog "OUT: CMMisc Payment" + " Oper:" + Oper$
  Unload Me
  frmCMPaySource.Show
  DoEvents
End Sub

Private Sub fpAmtPaid_LostFocus(Index As Integer)
  CalcBALFlds
End Sub

Private Sub fpAmtPaid_ChangeMode(Index As Integer, EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpAmtPaid_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Dim x As Integer
  If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
    If Index < 4 Then
      fpcboMiscCode(Index + 1).SetFocus
    Else
      fpCmdPost.SetFocus
    End If
  ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Then
    If Index > 0 Then
      fpcboMiscCode(Index - 1).SetFocus
    Else
      fpcboMiscCode(Index).SetFocus
    End If
  End If
End Sub
Private Sub fpcboMiscCode_LostFocus(Index As Integer)
  fpcboMiscCode(Index).Action = ActionClearSearchBuffer
End Sub

Private Sub fpcboMiscCode_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboMiscCode(Index).ListDown = True
    KeyCode = 0
  End If
  If KeyCode = vbKeyDelete Then
    fpcboMiscCode(Index).ListIndex = -1
    fpcboMiscCode(Index).Action = ActionClearSearchBuffer
    KeyCode = 0
  End If
  If fpcboMiscCode(Index).ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      If Index < 6 Then
        fpAmtPaid(Index).SetFocus
      Else
        fpCmdPost.SetFocus
      End If
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        If Index > 0 Then
          fpAmtPaid(Index - 1).SetFocus
        Else
          fptxtDesc.SetFocus
        End If
        KeyCode = 0
      End If
    End If
  End If
  DoEvents
End Sub
Private Sub fpCashAmt_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpCashAmt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    If fpChkAmt.Enabled Then
      fpChkAmt.SetFocus
    Else
      fptxtDesc.SetFocus
    End If
  End If
End Sub



Private Sub fpChkAmt_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpChkAmt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtDesc.SetFocus
  End If
End Sub


Private Sub fpTAmtOwed_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fpTAmtOwed_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboTenderType.SetFocus
  ElseIf KeyCode = vbKeyUp Then
    fptxtCity.SetFocus
  End If
End Sub

Private Sub fptxtName_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fptxtName_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtAddress.SetFocus
  End If
End Sub
Private Sub fptxtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fptxtCity.SetFocus
  ElseIf KeyCode = vbKeyUp Then
    fptxtName.SetFocus
  End If
End Sub
Private Sub fptxtAddress_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub fptxtCity_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpTAmtOwed.SetFocus
  ElseIf KeyCode = vbKeyUp Then
    fptxtAddress.SetFocus
  End If
End Sub
Private Sub fptxtCity_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub Chk4Change()
  Answer = 0
  If fpTotReceived <> 0 Or fpTotPaid <> 0 Then
    frmChangedWarning.Show vbModal, Me
    Select Case SaveFlag
    Case False
      Answer = 3
    Case True
      Answer = 2
    Case 1
      Answer = 1
    End Select
  Else
    Answer = 0
  End If
End Sub
Private Sub Chk4OKforNew()
  Dim FntSize As Integer
  Dim cntout As Integer, cnt As Integer
  If Len(fptxtName) > 0 Then
  

  cntout = 0
  Answer = 0
    If fpTotReceived <> 0 Or fpTotPaid <> 0 Then cntout = cntout + 1
  
  If cntout > 0 Then
  
    ReDim MsgText(0 To 5) As String
    FntSize = frmMsgDialog.Label(1).FontSize
    frmMsgDialog.Label(1).FontSize = (FntSize + 2)
    frmMsgDialog.Label(2).FontSize = (FntSize + 2)
    frmMsgDialog.Label(3).FontSize = (FntSize + 2)
    MsgText(0) = "WARNING:Payment In Progress"
    MsgText(1) = ""
    MsgText(2) = "Do You Want to Abandon this Payment?"
    MsgText(3) = "Ok to Abandon,"
    MsgText(4) = "Cancel to Remain on Current Payment."
    MsgText(5) = ""
    If GetOKorNot(MsgText()) Then
     UBLog "USER WANTS TO Abandon"
     CMLog "USER WANTS TO Abandon"
     Answer = 2
    Else
     UBLog "USER Canceled"
     CMLog "USER Canceled"
     Answer = 1
    End If
  Else
    Answer = 0
  End If
  End If
End Sub
Private Sub fpcmdDrawer_Click()
  Dim Port As String, PortFile As Integer ', DPName As String, DefPrinter As String
  On Local Error Resume Next
  If RecpDef = 99 Then Exit Sub
  Port$ = QPTrim$(RecpPort)
  
  CMLog "Oper: " + Oper$ + "CM MiscPay-Open Drawer"
  PortFile = FreeFile
  Open Port$ For Output As #PortFile
  If Biscoe = False Then
    Print #PortFile, Chr$(27); "p"; Chr$(0); Chr$(25); Chr$(250)
    Print #PortFile, Chr$(7)
  Else
    Print #PortFile, Chr$(7)
  End If
  Close PortFile
End Sub

Private Sub fpCmdPost_Click()
On Local Error GoTo ERRORSTUFF
  CalcBALFlds
  CheckInfo
  If ChkOKFlag Then
   'DeActivateControls Me
    If fpcboTenderType.ListIndex = 1 Or fpcboTenderType.ListIndex = 2 Then
      frmPrintReceipt.setvallist = 1
    Else
      frmPrintReceipt.setvallist = 0
    End If
    frmPrintReceipt.Show 1
    CheckInfo
    If ChkOKFlag Then
      If SavePay = True Then
        SaveTransaction
  
        If PrnRecp = True Or PrnVali = True Then
          PrintReceipt
        End If
  
        MsgBox "Transaction Complete.", vbOKOnly, "Complete"
        ClearScn
        fpChkAmt.Enabled = True
        fpCashAmt.Enabled = True
        BeenDone = True
        fptxtName.SetFocus
      End If
    'ActivateControls Me
    End If
  End If
Exit Sub
ERRORSTUFF:
  Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "CMPayMiscEntry", "cmdSave", Erl)
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
   Unload Me
End Sub
'Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Then
'    KeyCode = 0
'    fpAmount(0).SetFocus
'  ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyLeft Then
'    fpCmdSave.SetFocus
'  End If
'End Sub

Private Sub fpcmdCash_Click()
  fpcboTenderType.ListIndex = 0
  fpChkAmt.Enabled = False
  fpCashAmt.Enabled = True
  fpChkAmt = 0
  fpCashAmt.SetFocus
End Sub

Private Sub fpcmdCheck_Click()
  fpcboTenderType.ListIndex = 1
  fpCashAmt.Enabled = False
  fpChkAmt.Enabled = True
  fpCashAmt = 0
  fpChkAmt.SetFocus
End Sub
Private Sub fpCmdCharge_Click()
  fpcboTenderType.ListIndex = 3
  fpCashAmt.Enabled = False
  fpChkAmt.Enabled = True
  fpCashAmt = 0
  fpChkAmt.SetFocus
End Sub

Private Sub fpCashAmt_LostFocus()
fpTotReceived = Round#(fpCashAmt.DoubleValue + fpChkAmt.DoubleValue)
If fpTotReceived <> 0 Then
  fpChangeDue = Round#(fpTotReceived.DoubleValue - fpTotPaid.DoubleValue)
End If
End Sub

Private Sub fpChkAmt_LostFocus()
fpTotReceived = Round#(fpCashAmt.DoubleValue + fpChkAmt.DoubleValue)
If fpTotReceived.DoubleValue <> 0 Then
  fpChangeDue = Round#(fpTotReceived.DoubleValue - fpTotPaid.DoubleValue)
End If
End Sub
Private Sub fpcboTenderType_DropDown()
  ClrAmts
End Sub

Private Sub fpcboTenderType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboTenderType.ListDown = True
   ' ClrAmts
    KeyCode = 0
  End If
  If KeyCode = vbKeyDelete Then
    fpcboTenderType.ListIndex = -1
    fpcboTenderType.Action = ActionClearSearchBuffer
   ' ClrAmts
  End If
  If fpcboTenderType.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      If fpCashAmt.Enabled = True Then
        fpCashAmt.SetFocus
      Else
        fpChkAmt.SetFocus
      End If
        KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpTAmtOwed.SetFocus
        KeyCode = 0
      End If
    End If
  End If
  DoEvents
End Sub
Private Sub ClrAmts()
  Dim cnt As Integer
  fpCashAmt = 0
  fpChkAmt = 0
  fpChangeDue.Enabled = True
  fpChangeDue = 0
  For cnt = 1 To 5
    fpAmtPaid(cnt - 1) = 0
  Next
  fpTotPaid = 0
  fpTotReceived = 0
End Sub
Private Sub fpcboTenderType_SelChange(ItemIndex As Long)
  If BeenDone Then
    fixamts
  End If
End Sub
Private Sub fixamts()
  fpcboTenderType.Action = ActionClearSearchBuffer
  If noreset = False Then
    If fpcboTenderType.ListIndex = 0 Then
      fpCashAmt.Enabled = True
      fpChkAmt = 0
      fpChkAmt.Enabled = False
      fpChangeDue.Enabled = True
      'ClrAmts
     ' fpCashAmt.SetFocus
    ElseIf fpcboTenderType.ListIndex = 1 Then
      fpCashAmt.Enabled = False
      fpCashAmt = 0
      fpChkAmt.Enabled = True
      fpChangeDue.Enabled = True
      'ClrAmts
     ' fpChkAmt.SetFocus
    ElseIf fpcboTenderType.ListIndex = 2 Then
      fpCashAmt.Enabled = True
      fpChkAmt.Enabled = True
      fpChangeDue.Enabled = True
     ' ClrAmts
     'fpCashAmt.SetFocus
    ElseIf fpcboTenderType.ListIndex = 3 Then
      fpCashAmt.Enabled = False
      fpCashAmt = 0
      fpChkAmt.Enabled = True
      fpChangeDue = 0
      fpChangeDue.Enabled = False
     ' ClrAmts
      'fpChkAmt.SetFocus
'    ElseIf fpcboTenderType.ListIndex = -1 Then
'      MsgBox "You Must Select A Tender Type.", vbOKOnly, "Invalid Selection"
'      fpcboTenderType.SetFocus
    End If
  End If
  fpTotReceived = Round#(fpCashAmt.DoubleValue + fpChkAmt.DoubleValue)
  If fpTotReceived > 0 Then
    If fpcboTenderType.ListIndex <> 3 Then
      fpChangeDue = Round#(fpTotReceived.DoubleValue - fpTotPaid.DoubleValue)
    End If
  End If
  DoEvents
  noreset = False
End Sub

'Private Sub fpcboTenderType_LostFocus()
'
'  fpcboTenderType.Action = ActionClearSearchBuffer
'  If noreset = False Then
'
'    If fpcboTenderType.ListIndex = 0 Then
'      fpCashAmt.Enabled = True
'      fpChkAmt = 0
'      fpChkAmt.Enabled = False
'      'fpCashAmt.SetFocus
'    ElseIf fpcboTenderType.ListIndex = 1 Then
'      fpCashAmt = 0
'      fpCashAmt.Enabled = False
'      fpChkAmt.Enabled = True
'      'fpChkAmt.SetFocus
'    ElseIf fpcboTenderType.ListIndex = 2 Then
'      fpCashAmt.Enabled = True
'      fpChkAmt.Enabled = True
'      'fpCashAmt.SetFocus
'    ElseIf fpcboTenderType.ListIndex = 3 Then
'      fpCashAmt = 0
'      fpCashAmt.Enabled = False
'      fpChkAmt.Enabled = True
'      'fpChkAmt.SetFocus
''    ElseIf fpcboTenderType.ListIndex = -1 Then
''      MsgBox "You Must Select A Tender Type.", vbOKOnly, "Invalid Selection"
''      fpcboTenderType.SetFocus
'    End If
'  End If
'
'  noreset = False
'End Sub
'
'Private Sub fptxtDesc_ChangeMode(EditMode As Integer)
'  EditMode = True
'End Sub

Private Sub fptxtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboMiscCode(0).SetFocus
  End If
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    Else
      If MsgBox("Are You Sure You Wish To Close The Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        CMLog "Closed via CMMiscPaymentEntry by " + PWUser$ + " operator-" + Oper$
        CitiTerminate
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
      KeyCode = 0
      DoEvents
      If cmdExit.Enabled Then
        Call cmdExit_Click
      End If
    Case vbKeyF2:
      KeyCode = 0
      DoEvents
      fpcmdDrawer_Click
'    Case vbKeyF4:
'      KeyCode = 0
'      DoEvents
'      Call fpCmdInfo_Click
      If fpCmdCash.Enabled Then
        Call fpcmdCash_Click
      End If
    Case vbKeyF6:
      KeyCode = 0
      DoEvents
      If fpcmdCheck.Enabled Then
        Call fpcmdCheck_Click
      End If
    Case vbKeyF8:
      KeyCode = 0
      DoEvents
      If fpCmdCharge.Enabled Then
        Call fpCmdCharge_Click
      End If
'    Case vbKeyF8:
'      KeyCode = 0
'      DoEvents
'      Call fpcmdFind_Click
'    Case vbKeyF9:
'      KeyCode = 0
'      DoEvents
'      Call fpCmdDist_Click
    Case vbKeyF10:
      KeyCode = 0
      DoEvents
      If fpCmdPost.Enabled Then
        Call fpCmdPost_Click
      End If
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  StatusBar1.Panels.Item(1).Text = TownName$
  If InStr(TownName$, "BISCOE") > 0 Then
    Biscoe = False
  Else
    Biscoe = False
  End If
  txtPaymentDate.Text = DefPayDate
  Misctwo
  noreset = False
  fpcboTenderType.AddItem "Cash"
  fpcboTenderType.AddItem "Check"
  fpcboTenderType.AddItem "Cash & Check"
  fpcboTenderType.AddItem "Charge"
  'fpcboTenderType.ListIndex = -1
  lblOperator = OperNum
  lblOperName.Caption = PWUser
  lblSource.Caption = "Miscellaneous"
  Oper$ = QPTrim(lblOperator.Caption)
  CMLog " IN Oper " + Oper$ + ": CMMisc Payment"
'  LoadPayList
  GetRcpInfo
End Sub
Private Sub GetRcpInfo()
  Dim RP As Integer, lenRP As Integer, RP1 As Integer
  Dim RcptPrnFile As ReceiptPRNType
  RP1 = FreeFile
  lenRP = Len(RcptPrnFile)
  If Exist(RcptFileName$) Then
    Open RcptFileName$ For Random Shared As RP1 Len = lenRP
    Get RP1, 1, RcptPrnFile
    RecpPort = QPTrim(RcptPrnFile.RcpPort)
    If RcptPrnFile.PrnDefYN = 0 Then
      RecpDef = 0
    Else
      On Local Error GoTo nofound
      RP = FreeFile
      Open RecpPort For Output As RP
      Close RP
      RecpDef = 1
    End If
    If RcptPrnFile.CtlDefYN = 0 Then
      CntrlDef = 0
    Else
      CntrlDef = 1
    End If
    If RcptPrnFile.RValidate = 1 Then
      ValiDef = 1
      RctValidate = True
      GetUBBankINfo
    Else
      ValiDef = 0
      RctValidate = False
    End If
  Close RP1
  Else
    RecpDef = 99
    ValiDef = 0
  End If
Exit Sub
nofound:
  RecpDef = 99
  ValiDef = 0
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ''' Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
'
'  If Me.Visible Then
'    Temp_Class.ResizeControls Me
'    DoEvents
'  End If
End Sub
Private Sub ClearScn()
  Dim cnt As Integer
  BeenDone = False
  'LabelDel.Visible = False
  'fpCmdTranHist.Enabled = False
  fptxtName = ""
  fptxtAddress = ""
  fptxtCity = ""
  fpTAmtOwed = 0
  fptxtDesc = ""
  fpcboTenderType.ListIndex = -1
  fpCashAmt = 0
'  fpCashAmt.Enabled = True
'  fpChkAmt.Enabled = True
  fpChkAmt = 0
  fpChangeDue = 0
  For cnt = 1 To 5
    fpAmtPaid(cnt - 1) = 0
    fpcboMiscCode(cnt - 1).ListIndex = -1
  Next
  fpTotPaid = 0
  fpTotReceived = 0
  'fptxtName.SetFocus
End Sub

Private Sub CalcBALFlds()
  Dim cnt As Integer, TPay As Double
  TPay# = 0
  For cnt = 1 To 5
    'TOwd# = Round#(TOwd# + fpAmtOwed(cnt - 1).DoubleValue)
    'TCur# = Round#(TCur# + fpCurrent(cnt - 1).DoubleValue)
    'fpActual(cnt - 1) = Round#(fpCurrent(cnt - 1).DoubleValue - fpAmount(cnt - 1).DoubleValue)
    'fpActual(cnt - 1) = 0
    TPay# = Round#(TPay# + fpAmtPaid(cnt - 1).DoubleValue)
  Next
  'fpTotOwed = TOwd#
  fpTotPaid = TPay#
  If fpTotReceived <> 0 Then
    fpChangeDue = Round#(fpTotReceived.DoubleValue - fpTotPaid.DoubleValue)
  End If
End Sub
Private Sub SaveTransaction()
  Dim ListFile As Integer, CHandle As Integer
  Dim PayFileName As String, UBPayRecLen As Integer
  Dim NumOfRecs As Long, CMTrRecLen As Integer
  Dim cnt As Integer

  ReDim UBPaymentRec(1) As UBPaymentRecType
  Oper$ = QPTrim$(lblOperator.Caption)

  PayFileName$ = "C:\CPWork\CMPAY" + Oper$ + ".DAT"

  UBPayRecLen = Len(UBPaymentRec(1))
  For cnt = 1 To 5
    If fpcboMiscCode(cnt - 1).ListIndex <> -1 Then
      If fpAmtPaid(cnt - 1) < -100000# Then
        UBPaymentRec(1).PaidOwed(cnt).AMTPD1 = 0
      Else
        UBPaymentRec(1).PaidOwed(cnt).AMTPD1 = fpAmtPaid(cnt - 1)
      End If
    Else
      UBPaymentRec(1).PaidOwed(cnt).AMTPD1 = 0
    End If
  Next cnt

  For cnt = 6 To 10     ' Store the Misc Code Record Number in Rev Amt 6-10
    fpcboMiscCode(cnt - 6).col = 2
    If fpcboMiscCode(cnt - 6).ListIndex <> -1 Then
      UBPaymentRec(1).PaidOwed(cnt).AMTPD1 = QPTrim(fpcboMiscCode(cnt - 6).ColText)
    Else
      UBPaymentRec(1).PaidOwed(cnt).AMTPD1 = 0
    End If
  Next cnt
  For cnt = 1 To 5
    UBPaymentRec(1).PaidOwed(cnt).AMTOWE1 = 0
  Next
  UBPaymentRec(1).OperNum = QPTrim(lblOperator.Caption)
  UBPaymentRec(1).payDate = Date2Num(txtPaymentDate)
  UBPaymentRec(1).CustAcct = 99999
  UBPaymentRec(1).CustName = QPTrim(fptxtName)
  UBPaymentRec(1).CustAddr = QPTrim(fptxtAddress)
  UBPaymentRec(1).CUSTCMNT = QPTrim(Label4.Caption)
  'UBPaymentRec(1).TaxExempt = QPTrim(fptaxexmpt)
  UBPaymentRec(1).AmtOwed = fpTAmtOwed
  Select Case fpcboTenderType.ListIndex
    Case 0:
      UBPaymentRec(1).TenderTY = "Cash"
    Case 1:
      UBPaymentRec(1).TenderTY = "Check"
    Case 2:
      UBPaymentRec(1).TenderTY = "Cash & Check"
    Case 3:
      UBPaymentRec(1).TenderTY = "Charge"
    Case Else:
      UBPaymentRec(1).TenderTY = "Unknown"
  End Select
  UBPaymentRec(1).CashAmt = fpCashAmt
  UBPaymentRec(1).ChkAmt = fpChkAmt
  UBPaymentRec(1).AmtRecd = fpTotReceived
  UBPaymentRec(1).Change = fpChangeDue
  UBPaymentRec(1).Desc = QPTrim(fptxtDesc)
  UBPaymentRec(1).TotOwed = fpTAmtOwed
  UBPaymentRec(1).AmtPaid = fpTotPaid
  'UBPaymentRec(1).Status = QPTrim(fpstatus)
  ListFile = FreeFile
  Open PayFileName$ For Random Shared As ListFile Len = UBPayRecLen
    Put #ListFile, 1, UBPaymentRec(1)
    EditFlag = False
  CMLog "Oper:" + Oper$ + " Updated TempFile for Misc Pay"
  

  ReDim CMTrRec(1) As CMTransRecType
  CMTrRecLen = Len(CMTrRec(1))
  CHandle = FreeFile
  Open UBPath$ + "CMTRANS.DAT" For Random Access Read Write Shared As CHandle Len = CMTrRecLen
  CmNum = (LOF(CHandle) \ CMTrRecLen) + 1
  CMTrRec(1).TransDate = UBPaymentRec(1).payDate
  CMTrRec(1).TransAmount = UBPaymentRec(1).AmtPaid
  CMTrRec(1).TransCash = UBPaymentRec(1).CashAmt
  CMTrRec(1).TransCheck = UBPaymentRec(1).ChkAmt
  CMTrRec(1).TransAmtOwed = UBPaymentRec(1).TotOwed
  If Len(QPTrim$(UBPaymentRec(1).Desc)) = 0 Then
    CMTrRec(1).TransDesc = "Miscellaneous Payment"
  Else
    CMTrRec(1).TransDesc = (QPTrim$(UBPaymentRec(1).Desc))
  End If
  CMTrRec(1).TransSource = 1
  CMTrRec(1).TransName = UBPaymentRec(1).CustName
  CMTrRec(1).TransAcctNum = 99999
  CMTrRec(1).TransDetNum = CmNum
  CMTrRec(1).TransOperNum = OperNum
  Select Case QPTrim(UBPaymentRec(1).TenderTY)
    Case "Cash":
      CMTrRec(1).TransTender = 1
    Case "Check":
      CMTrRec(1).TransTender = 2
    Case "Cash & Check":
      CMTrRec(1).TransTender = 3
    Case "Charge":
      CMTrRec(1).TransTender = 4
    Case Else:
      '
  End Select
  CMTrRec(1).ChkByte = Chr$(1)
  CMTrRec(1).TransPad = ""

  For cnt = 1 To 5
    CMTrRec(1).TransRevAmt(cnt) = UBPaymentRec(1).PaidOwed(cnt).AMTPD1
  Next cnt

  For cnt = 1 To 5              ' Store the Misc Code Record Number in Rev Amt 6-10
    CMTrRec(1).TransRevAmt(cnt + 5) = UBPaymentRec(1).PaidOwed(cnt + 5).AMTPD1
  Next cnt

  Put CHandle, (LOF(CHandle) / CMTrRecLen) + 1, CMTrRec(1)
  CMLog "Saved Misc trans " + Str((LOF(CHandle) / CMTrRecLen) + 1)
  Close
  'ClearScn
  
End Sub

Private Sub CheckInfo()
  Dim TestDate As Integer, TestAmt As Double, cnt As Integer
  TestAmt = 0
  ChkOKFlag = True
  TestDate = Date2Num(txtPaymentDate)
  If TestDate < 0 Then
    ChkOKFlag = False
    MsgBox "Invalid Date.", vbOKOnly, "Request Canceled."
    GoTo BadDate
  End If
  If fpcboTenderType.ListIndex = -1 Then
    MsgBox "You Must Select A Tender Type.", vbOKOnly, "Invalid Selection"
    ChkOKFlag = False
    GoTo BadDate
  End If
  If fpcboTenderType.ListIndex = 0 And fpChkAmt.DoubleValue > 0 Then
    ChkOKFlag = False
    MsgBox "Invalid Tender Type. The Check/Charge Amount Should Be ZERO.", vbOKOnly, "Request Canceled."
    GoTo BadDate
  End If
  If (fpcboTenderType.ListIndex = 1 Or fpcboTenderType.ListIndex = 3) And fpCashAmt.DoubleValue > 0 Then
    ChkOKFlag = False
    MsgBox "Invalid Tender Type. The Cash Amount Should Be ZERO.", vbOKOnly, "Request Canceled."
    GoTo BadDate
  End If
  If fpcboTenderType.ListIndex = 2 And (fpChkAmt.DoubleValue <= 0 Or fpCashAmt.DoubleValue <= 0) Then
    ChkOKFlag = False
    MsgBox "Invalid Amounts. The Check and Cash Amount Should Be Greater than ZERO.", vbOKOnly, "Request Canceled."
    GoTo BadDate
  End If
  For cnt = 0 To 4
    If Val(fpAmtPaid(cnt)) <> 0 Then
      If fpcboMiscCode(cnt).ListIndex = -1 Then
        MsgBox "You Must Select A Code.", vbOKOnly, "Invalid Selection"
        ChkOKFlag = False
        GoTo BadDate
      End If
    End If
  Next
  For cnt = 0 To 4
    If fpcboMiscCode(cnt).ListIndex <> -1 Then
      If Val(fpAmtPaid(cnt)) = 0 Then
        MsgBox "You Must Enter An Amount.", vbOKOnly, "Invalid Entry"
        ChkOKFlag = False
        GoTo BadDate
      End If
    End If
  Next

  If fpTotReceived.DoubleValue = 0 Or fpTotPaid.DoubleValue = 0 Then
    ChkOKFlag = False
    MsgBox "Invalid Amount. The Total Received and Total Paid Should NOT Be ZERO.", vbOKOnly, "Request Canceled."
    GoTo BadDate
  End If
  If fpTotPaid.DoubleValue > fpTAmtOwed.DoubleValue Then
    ChkOKFlag = False
    MsgBox "Invalid Amount. The Total Paid Should Not Be Greater than Total Owed.", vbOKOnly, "Request Canceled."
    GoTo BadDate
  End If
  If fpChangeDue.DoubleValue >= 0 Then
    TestAmt = Round#(fpTotReceived.DoubleValue - fpChangeDue.DoubleValue)
    If TestAmt <> fpTotPaid Then '.DoubleValue Then
      ChkOKFlag = False
      MsgBox "The Total Received does NOT equal the Total Paid.", vbOKOnly, "Request Canceled."
      GoTo BadDate
    End If
  Else
    ChkOKFlag = False
    MsgBox "The Amount Distributed May Not Be More Than Amount Received.", vbOKOnly, "Request Canceled."
    GoTo BadDate
  End If
 Exit Sub
BadDate:
  Exit Sub
End Sub
Private Sub PrintReceipt()
  Dim ListFile As Integer, PayFileName As String, UBPayRecLen As Integer
  Dim RecptNum As Long, RHandle As Integer, PayRecpName As String
  Dim CutPaper As String, PostDate As String, RevCnt As Integer
  Dim NumofRevs As Integer, RecpRev As String, RevText As String
  Dim RHandle2 As Integer, PayRecpName2 As String, RptHandle2 As Integer
  ReDim UBPaymentRec(1) As UBPaymentRecType
 ' RecpRev$ = Space$(15)
  CutPaper$ = Chr$(29) + Chr$(86) + Chr$(66) + Chr$(64)
   If InStr(TownName$, "Dobson") > 0 Then CutPaper$ = Chr$(27) + Chr$(100)
  UBPayRecLen = Len(UBPaymentRec(1))
  PayFileName$ = "C:\CPWork\CMPAY" + Oper$ + ".DAT"
  PayRecpName$ = "C:\CPWork\CMRCP" + Oper$ + ".RPT"
  PayRecpName2$ = "C:\CPWork\CMVLD" + Oper$ + ".Rpt"
  PostDate$ = txtPaymentDate
  ListFile = FreeFile
  Open PayFileName$ For Random Shared As ListFile Len = UBPayRecLen
  'RecptNum& = LOF(ListFile) / UBPayRecLen
  Get #ListFile, 1, UBPaymentRec(1)
  Close
  NumofRevs = MaxRevsCnt
  If PrnRecp = False And PrnVali = True Then GoTo Validationthing

  RHandle = FreeFile
  Open PayRecpName$ For Output As RHandle
  If CntrlDef = 1 Then
    If Biscoe = False Then
      fpcmdDrawer_Click
    Else
      Print #RHandle, Chr$(7)
    End If
  End If
  Print #RHandle, TownName$
  Print #RHandle, "CM MISCELLANEOUS PAYMENT"
  Print #RHandle, "Date: "; PostDate$
  Print #RHandle, "Time: "; Time
  Print #RHandle,
  Print #RHandle, "CUSTOMER NAME & DESC. OF PAYMENT"
  Print #RHandle, UBPaymentRec(1).CustName
  Print #RHandle, UBPaymentRec(1).CustAddr
  Print #RHandle, UBPaymentRec(1).Desc
  Print #RHandle, '"Acct. No. "; UBPaymentRec(1).CustAcct
  Print #RHandle,
  Print #RHandle, QPTrim(UBPaymentRec(1).TenderTY)
  Print #RHandle,
  Print #RHandle, "       Cash: "; Using("$##,###,###.##", UBPaymentRec(1).CashAmt)
  If QPTrim$(UBPaymentRec(1).TenderTY) <> "Charge" Then
    Print #RHandle, "      Check: "; Using("$##,###,###.##", UBPaymentRec(1).ChkAmt)
    If Biscoe = False Then Print #RHandle, "     Charge: "; Using("$##,###,###.##", 0)
  Else
    Print #RHandle, "      Check: "; Using("$##,###,###.##", 0)
    If Biscoe = False Then Print #RHandle, "     Charge: "; Using("$##,###,###.##", UBPaymentRec(1).ChkAmt)
  End If
  Print #RHandle, " Total Owed: "; Using("$##,###,###.##", UBPaymentRec(1).TotOwed)
  Print #RHandle, " Total Paid: "; Using("$##,###,###.##", UBPaymentRec(1).AmtRecd)
  Print #RHandle, " Change Due: "; Using("$##,###,###.##", UBPaymentRec(1).Change)
  Print #RHandle, "Amt Applied: "; Using("$##,###,###.##", UBPaymentRec(1).AmtPaid)
  Print #RHandle, "    Balance: "; Using("$##,###,###.##", UBPaymentRec(1).TotOwed - UBPaymentRec(1).AmtPaid)
  Print #RHandle,
  For RevCnt = 1 To 5
    If UBPaymentRec(1).PaidOwed(RevCnt).AMTPD1 <> 0 Or UBPaymentRec(1).PaidOwed(RevCnt).AMTOWE1 <> 0 Then
      MiscCode$ = ""
      GetMiscDesc UBPaymentRec(1).PaidOwed(RevCnt + 5).AMTPD1
      Print #RHandle, MiscCode$; Using("$########.##", UBPaymentRec(1).PaidOwed(RevCnt).AMTPD1)
    End If
  Next
  Print #RHandle,
  Print #RHandle, "Operator: "; OperNum
  Print #RHandle, "Receipt#: "; Using("######", CmNum&)
  Print #RHandle,
  Print #RHandle, "       T H A N K   Y O U !"
  Print #RHandle,
  Print #RHandle,
  Print #RHandle,
'  Print #RHandle,
'  Print #RHandle,
  If CntrlDef = 1 Then
    If Biscoe = False Then Print #RHandle, CutPaper$
  Else
    Print #RHandle,
    Print #RHandle,
    Print #RHandle,
  End If
  Close RHandle

  'Shell$ = "type " + PayRecpName$ + " > com2:"
  'SHELL Shell$
'  If CntrlDef = 1 Then
'    fpcmdDrawer_Click
'  End If
  'PrintRptFile Header$, PayRecpName$, RecpPort, RetCode%, 5
  Dim RptHandle As Integer, LPTHandle As Integer
  Dim RptA As Integer, LPTA As Integer, ToPrintA As String
  Dim ToPrint As String, CopyLoop As Integer, DefPrinter As String
  On Error GoTo Cancel
  'Printer.Print
'''  to strReportFile DefPrinter'[ADDITIVE] | PortName]
10:
  DefPrinter = RecpPort '"LPT" + QPTrim$(Str$(RecpPort)) + ":"
20:
 ' MsgBox "Printer -" + DefPrinter, vbOKOnly
  
  For CopyLoop = 1 To 1 'Copies
    LPTHandle = FreeFile
    Open DefPrinter For Output As LPTHandle
    RptHandle = FreeFile
30:
    Open PayRecpName$ For Input As RptHandle
40:
    Do
      If frmPrint.cmdCancel = False Then
45:
        Line Input #RptHandle, ToPrint$
        
        ToPrint$ = RTrim$(ToPrint$)
        Print #LPTHandle, ToPrint$
      Else
50:
        Exit Do
        'Printer.EndDoc
      End If
    Loop Until eof(RptHandle)
60:
    Close RptHandle
62:
    Close LPTHandle
65:
    Next CopyLoop
68:
 Printer.EndDoc
69:
Validationthing:
  If QPTrim(UBPaymentRec(1).TenderTY) = "Check" Or QPTrim(UBPaymentRec(1).TenderTY) = "Cash & Check" Then
   If RctValidate And PrnVali = True Then
     DefPrinter = RecpPort
     RHandle2 = FreeFile
     Open PayRecpName2$ For Output As RHandle2
     Print #RHandle2, Chr$(27); Chr$(&H63); Chr$(&H30); Chr$(&H4)
     Print #RHandle2, Chr$(13); Chr$(10)
     Print #RHandle2, Tab(12); TownName$
     Print #RHandle2, Tab(12); "Bank- "; BnkAcctNum$
     Print #RHandle2, Tab(12); "FOR DEPOSIT ONLY"
     Print #RHandle2, Tab(12); "Acct. No. "; UBPaymentRec(1).CustAcct
     Print #RHandle2, Tab(12); "Date: "; PostDate$
     Print #RHandle2, Tab(12); "Time: "; Time
     Print #RHandle2,
     Print #RHandle2, Chr$(12)
     Close RHandle2
     LPTHandle = FreeFile
     Open DefPrinter For Output As LPTHandle
     RptHandle2 = FreeFile
     Open PayRecpName2$ For Input As RptHandle2
     Do
       If frmPrint.cmdCancel = False Then
         Line Input #RptHandle2, ToPrint$
         ToPrint$ = RTrim$(ToPrint$)
         Print #LPTHandle, ToPrint$
       Else
         Exit Do
       End If
     Loop Until eof(RptHandle2)
     Close RptHandle2
     Close LPTHandle
    Printer.EndDoc
    CMLog "Oper: " + Oper$ + " Print Validation Acct:" + Str(UBPaymentRec(1).CustAcct)
  End If
 End If
70:
If PrnRecp = True Then
 CMLog "Oper: " + Oper$ + " Print receipt Acct:" + Str(UBPaymentRec(1).CustAcct)
 KillFile PayRecpName$
 KillFile PayFileName$
End If
80:
  Exit Sub
Cancel:
  If Err > 0 Then
    CMLog "Error Code Was " + DefPrinter + Err.Description + Str$(Err) + " (PrintWSet - Line:" & Erl & ")"
    MsgBox "Error Code Was " + DefPrinter + Err.Description + Str$(Err) + " (PrintWSet - Line:" & Erl & ")"
  End If
  Close
  Exit Sub
  
End Sub

Public Sub OpenMiscCodeFile(NumOfMiscRecs, MCFile)
  Dim MiscCodeRecLen As Integer
  ReDim MiscCodeRec(1) As MiscCodeRecType
  MiscCodeRecLen = Len(MiscCodeRec(1))
  MCFile = FreeFile
  Open UBPath$ + "CMMISCCD.DAT" For Random Shared As MCFile Len = MiscCodeRecLen
  NumOfMiscRecs = LOF(MCFile) \ MiscCodeRecLen

End Sub
Public Function Misctwo()
  Dim NumOfMiscRecs As Integer, cnt As Integer, CntA As Integer
  Dim MCFile As Integer
  ReDim MiscCodeRec(1) As MiscCodeRecType

  OpenMiscCodeFile NumOfMiscRecs, MCFile
  Dim TempList As String
  For cnt = 0 To 4
    fpcboMiscCode(cnt).Row = -1
  Next
  For CntA = 1 To NumOfMiscRecs
    Get MCFile, CntA, MiscCodeRec(1)
    If QPTrim$(MiscCodeRec(1).InactiveFlag) <> "Y" Then
      TempList = QPTrim(MiscCodeRec(1).MiscCode) & Chr$(9) & QPTrim$(MiscCodeRec(1).Description) & Chr$(9) & Str(CntA)
      For cnt = 0 To 4
        fpcboMiscCode(cnt).AddItem TempList
      Next
    End If
    TempList = ""
  Next
  Close MCFile
  End Function
Private Function GetMiscDesc(MiscR)
  Dim NumOfMiscRecs As Integer, cnt As Integer, CntA As Integer
  Dim MCFile As Integer
  ReDim MiscCodeRec(1) As MiscCodeRecType
  OpenMiscCodeFile NumOfMiscRecs, MCFile
    Get MCFile, MiscR, MiscCodeRec(1)
      MiscCode$ = QPTrim(MiscCodeRec(1).Description)
  Close MCFile
End Function

