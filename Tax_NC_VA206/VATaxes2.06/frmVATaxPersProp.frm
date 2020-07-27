VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVATaxPersProp 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personal Property Information"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11610
   Icon            =   "frmVATaxPersProp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11610
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbGoToPins 
      Height          =   360
      Left            =   4320
      TabIndex        =   7
      Top             =   3000
      Width           =   1500
      _Version        =   196608
      _ExtentX        =   2646
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
      EditAlignH      =   1
      EditAlignV      =   0
      ColDesigner     =   "frmVATaxPersProp.frx":08CA
   End
   Begin LpLib.fpCombo fpcmbOptRev3 
      Height          =   375
      Left            =   7680
      TabIndex        =   16
      Top             =   5010
      Width           =   2445
      _Version        =   196608
      _ExtentX        =   4313
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
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmVATaxPersProp.frx":0CC1
   End
   Begin LpLib.fpCombo fpcmbOptRev2 
      Height          =   375
      Left            =   7680
      TabIndex        =   15
      Top             =   4590
      Width           =   2445
      _Version        =   196608
      _ExtentX        =   4313
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
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmVATaxPersProp.frx":10B8
   End
   Begin LpLib.fpCombo fpcmbOptRev1 
      Height          =   375
      Left            =   7680
      TabIndex        =   14
      Top             =   4170
      Width           =   2445
      _Version        =   196608
      _ExtentX        =   4313
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
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmVATaxPersProp.frx":14AF
   End
   Begin LpLib.fpCombo fpcmbPPTRAYN 
      Height          =   390
      Left            =   7680
      TabIndex        =   11
      ToolTipText     =   "If (disabled) is displayed then the PPTRA discount feature has been turned off on the System Setup screen."
      Top             =   3330
      Width           =   780
      _Version        =   196608
      _ExtentX        =   1376
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
      ColDesigner     =   "frmVATaxPersProp.frx":18A6
   End
   Begin LpLib.fpCombo fpcmbPRVal 
      Height          =   390
      Left            =   10200
      TabIndex        =   12
      Top             =   2475
      Width           =   780
      _Version        =   196608
      _ExtentX        =   1376
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
      ColDesigner     =   "frmVATaxPersProp.frx":1C45
   End
   Begin LpLib.fpCombo fpcmbProRateYN 
      Height          =   390
      Left            =   7680
      TabIndex        =   10
      Top             =   2895
      Width           =   780
      _Version        =   196608
      _ExtentX        =   1376
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
      ColDesigner     =   "frmVATaxPersProp.frx":1FE4
   End
   Begin LpLib.fpCombo fpcmbDiscoveryYN 
      Height          =   390
      Left            =   7680
      TabIndex        =   8
      Top             =   2475
      Width           =   780
      _Version        =   196608
      _ExtentX        =   1376
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
      ColDesigner     =   "frmVATaxPersProp.frx":2383
   End
   Begin LpLib.fpCombo fpcmbLateListYN 
      Height          =   390
      Left            =   10200
      TabIndex        =   9
      Top             =   2895
      Visible         =   0   'False
      Width           =   780
      _Version        =   196608
      _ExtentX        =   1376
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
      ColDesigner     =   "frmVATaxPersProp.frx":2722
   End
   Begin EditLib.fpDoubleSingle fpDSWEight 
      Height          =   375
      Left            =   6960
      TabIndex        =   19
      Top             =   6810
      Width           =   1815
      _Version        =   196608
      _ExtentX        =   3196
      _ExtentY        =   656
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
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "9000000000"
      MinValue        =   "0"
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
   Begin EditLib.fpCurrency fpCurrPersVal 
      Height          =   375
      Left            =   2325
      TabIndex        =   2
      Top             =   2520
      Width           =   1815
      _Version        =   196608
      _ExtentX        =   3201
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
      Left            =   2858
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   1130
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
      Left            =   9120
      TabIndex        =   1
      Top             =   1800
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
   Begin EditLib.fpText fptxtRecord 
      Height          =   396
      Left            =   2568
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1800
      Width           =   2292
      _Version        =   196608
      _ExtentX        =   4048
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
   Begin EditLib.fpText fptxtPropPin 
      Height          =   396
      Left            =   6408
      TabIndex        =   0
      Top             =   1800
      Width           =   1692
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
   Begin EditLib.fpCurrency fpCurrMobHome 
      Height          =   375
      Left            =   2325
      TabIndex        =   3
      Top             =   3000
      Width           =   1830
      _Version        =   196608
      _ExtentX        =   3228
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
   Begin EditLib.fpCurrency fpCurrMerchCap 
      Height          =   375
      Left            =   2325
      TabIndex        =   4
      Top             =   3480
      Width           =   1815
      _Version        =   196608
      _ExtentX        =   3201
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
   Begin EditLib.fpCurrency fpCurrFarmEq 
      Height          =   375
      Left            =   2325
      TabIndex        =   5
      Top             =   3960
      Width           =   1815
      _Version        =   196608
      _ExtentX        =   3201
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
   Begin EditLib.fpCurrency fpCurrMachTools 
      Height          =   375
      Left            =   2325
      TabIndex        =   6
      Top             =   4440
      Width           =   1815
      _Version        =   196608
      _ExtentX        =   3201
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
      Height          =   390
      Index           =   0
      Left            =   390
      TabIndex        =   22
      Top             =   5280
      Width           =   5415
      _Version        =   196608
      _ExtentX        =   9551
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
   Begin EditLib.fpText fptxtDesc 
      Height          =   390
      Index           =   1
      Left            =   390
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5655
      Width           =   5415
      _Version        =   196608
      _ExtentX        =   9546
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
      MaxLength       =   30
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
      Height          =   390
      Index           =   2
      Left            =   390
      TabIndex        =   24
      Top             =   6030
      Width           =   5415
      _Version        =   196608
      _ExtentX        =   9546
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
   Begin EditLib.fpText fptxtDesc 
      Height          =   390
      Index           =   3
      Left            =   390
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6390
      Width           =   5415
      _Version        =   196608
      _ExtentX        =   9546
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
   Begin EditLib.fpText fptxtDesc 
      Height          =   390
      Index           =   4
      Left            =   390
      TabIndex        =   26
      Top             =   6765
      Width           =   5415
      _Version        =   196608
      _ExtentX        =   9546
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
      AlignTextH      =   0
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
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   420
      Left            =   9456
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1692
      _Version        =   131072
      _ExtentX        =   2984
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
      ButtonDesigner  =   "frmVATaxPersProp.frx":2AC1
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   420
      Left            =   480
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1692
      _Version        =   131072
      _ExtentX        =   2984
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
      ButtonDesigner  =   "frmVATaxPersProp.frx":2C9D
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDelete 
      Height          =   420
      Left            =   2256
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1692
      _Version        =   131072
      _ExtentX        =   2984
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
      ButtonDesigner  =   "frmVATaxPersProp.frx":2E79
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAdd 
      Height          =   420
      Left            =   7656
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1692
      _Version        =   131072
      _ExtentX        =   2984
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
      ButtonDesigner  =   "frmVATaxPersProp.frx":3056
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPageDown 
      Height          =   420
      Left            =   4056
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1692
      _Version        =   131072
      _ExtentX        =   2984
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
      ButtonDesigner  =   "frmVATaxPersProp.frx":3230
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPageUp 
      Height          =   420
      Left            =   5880
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1692
      _Version        =   131072
      _ExtentX        =   2984
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
      ButtonDesigner  =   "frmVATaxPersProp.frx":340C
   End
   Begin EditLib.fpText fptxtVin 
      Height          =   390
      Left            =   7800
      TabIndex        =   17
      Top             =   5760
      Width           =   3015
      _Version        =   196608
      _ExtentX        =   5313
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
   Begin EditLib.fpText fptxtMakeMod 
      Height          =   390
      Left            =   8280
      TabIndex        =   18
      Top             =   6270
      Width           =   2535
      _Version        =   196608
      _ExtentX        =   4466
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
   Begin EditLib.fpDateTime fpModYear 
      Height          =   375
      Left            =   10080
      TabIndex        =   20
      Top             =   6840
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
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
      AllowNull       =   -1  'True
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
      Text            =   ""
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "yyyy"
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
   Begin EditLib.fpDateTime fptxtBillYear 
      Height          =   375
      Left            =   9960
      TabIndex        =   13
      Top             =   3330
      Width           =   1020
      _Version        =   196608
      _ExtentX        =   1799
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
      Text            =   "2018"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "yyyy"
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
   Begin fpBtnAtlLibCtl.fpBtn cmdDetail1 
      Height          =   315
      Left            =   10290
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   4170
      Width           =   630
      _Version        =   131072
      _ExtentX        =   1111
      _ExtentY        =   556
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
      ButtonDesigner  =   "frmVATaxPersProp.frx":35E8
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDetail2 
      Height          =   315
      Left            =   10290
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   4590
      Width           =   630
      _Version        =   131072
      _ExtentX        =   1111
      _ExtentY        =   556
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
      ButtonDesigner  =   "frmVATaxPersProp.frx":37C1
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDetail3 
      Height          =   315
      Left            =   10290
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   5010
      Width           =   630
      _Version        =   131072
      _ExtentX        =   1111
      _ExtentY        =   556
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
      ButtonDesigner  =   "frmVATaxPersProp.frx":399A
   End
   Begin EditLib.fpText fptxtOptSearch 
      Height          =   390
      Left            =   8520
      TabIndex        =   21
      Top             =   7440
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
      Height          =   390
      Left            =   2400
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   7440
      Width           =   3615
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
   Begin fpBtnAtlLibCtl.fpBtn cmdGo 
      Height          =   390
      Left            =   4680
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   3480
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
      ButtonDesigner  =   "frmVATaxPersProp.frx":3B73
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   5055
      Left            =   5925
      Top             =   2280
      Width           =   5295
   End
   Begin VB.Label Label30 
      BackColor       =   &H0080FFFF&
      Caption         =   "Go To Prop #"
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
      Left            =   4240
      TabIndex        =   70
      Top             =   2280
      Width           =   1500
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   4200
      X2              =   6000
      Y1              =   2280
      Y2              =   2280
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
      Left            =   6360
      TabIndex        =   69
      Top             =   7560
      Width           =   1860
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
      Left            =   480
      TabIndex        =   68
      Top             =   7560
      Width           =   1740
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   5910
      X2              =   11180
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Label Label26 
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
      Height          =   255
      Left            =   6120
      TabIndex        =   66
      Top             =   5130
      Width           =   1380
   End
   Begin VB.Label Label25 
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
      Height          =   255
      Left            =   6120
      TabIndex        =   64
      Top             =   4710
      Width           =   1380
   End
   Begin VB.Label Label24 
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
      Height          =   255
      Left            =   6120
      TabIndex        =   62
      Top             =   4290
      Width           =   1380
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
      Left            =   5925
      TabIndex        =   60
      Top             =   3840
      Width           =   1860
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   5910
      X2              =   11180
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Bill Year:"
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
      TabIndex        =   59
      Top             =   3450
      Width           =   1305
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Model Year:"
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
      Left            =   8880
      TabIndex        =   58
      Top             =   6960
      Width           =   1185
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Weight:"
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
      Left            =   6120
      TabIndex        =   57
      Top             =   6930
      Width           =   780
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Make/Model:"
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
      Left            =   6960
      TabIndex        =   56
      Top             =   6390
      Width           =   1260
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "VIN #:"
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
      Left            =   6960
      TabIndex        =   55
      Top             =   5880
      Width           =   660
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Vehicle Desc"
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
      Left            =   5925
      TabIndex        =   54
      Top             =   5520
      Width           =   1620
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PPTRA Y/N?:"
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
      Left            =   6285
      TabIndex        =   53
      Top             =   3450
      Width           =   1260
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Prorate Value:"
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
      Left            =   8685
      TabIndex        =   52
      Top             =   2595
      Width           =   1380
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Prorate Y/N?:"
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
      Left            =   6165
      TabIndex        =   51
      Top             =   3015
      Width           =   1380
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2655
      Left            =   285
      Top             =   2280
      Width           =   3975
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Required Fields = *"
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
      Height          =   612
      Left            =   120
      TabIndex        =   44
      Top             =   1200
      Width           =   1020
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   2415
      Left            =   285
      Top             =   4920
      Width           =   5655
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
      Height          =   270
      Left            =   285
      TabIndex        =   42
      Top             =   4920
      Width           =   2580
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Settings"
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
      Left            =   5925
      TabIndex        =   41
      Top             =   2280
      Width           =   1260
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
      Left            =   8565
      TabIndex        =   40
      Top             =   3015
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mach/Tools Value:"
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
      Left            =   240
      TabIndex        =   39
      Top             =   4560
      Width           =   1980
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Farm Equip Value:"
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
      Left            =   480
      TabIndex        =   38
      Top             =   4080
      Width           =   1740
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Merch Capital Value:"
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
      TabIndex        =   37
      Top             =   3600
      Width           =   1860
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile Home Value:"
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
      Left            =   405
      TabIndex        =   36
      Top             =   3120
      Width           =   1860
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Value:"
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
      Left            =   765
      TabIndex        =   35
      Top             =   2640
      Width           =   1500
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Valuations"
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
      Left            =   285
      TabIndex        =   34
      Top             =   2280
      Width           =   1620
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
      Left            =   648
      TabIndex        =   33
      Top             =   1920
      Width           =   1740
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
      Left            =   4778
      TabIndex        =   32
      Top             =   720
      Width           =   2175
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
      Left            =   6045
      TabIndex        =   30
      Top             =   2595
      Width           =   1500
   End
   Begin VB.Label Label72 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*Pin Number:"
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
      Left            =   4968
      TabIndex        =   29
      Top             =   1908
      Width           =   1380
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
      Height          =   276
      Left            =   8208
      TabIndex        =   28
      Top             =   1920
      Width           =   780
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Property Information"
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
      Left            =   2911
      TabIndex        =   27
      Top             =   360
      Width           =   6015
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   1380
      Index           =   1
      Left            =   1478
      Top             =   300
      Width           =   8655
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1440
      Left            =   1478
      Top             =   240
      Width           =   8655
   End
End
Attribute VB_Name = "frmVATaxPersProp"
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
  Dim PersRecs() As Long
  Public NumOfCustPPRecs As Integer
  Dim TempPROPPIN$
  Dim TempPROPDATE As Integer
  Dim TempPersVal#
  Dim TempMHVALUE#
  Dim TempMCVALUE#
  Dim TempCVALUE#
  Dim TempMTVALUE#
  Dim TempDISCOV$
  Dim TempLATELIST$
  Dim TempOptSearch$
  Dim TempDESC1$
  Dim TempDESC2$
  Dim TempDESC3$
  Dim TempDesc4$
  Dim TempDesc5$
  Dim TempTaxBillYear As Integer
  Dim TempPPTRAYN As String
  Dim TempProrateVal As Integer
  Dim TempProrate As String
  Dim TempVin$
  Dim TempMakeMod$
  Dim TempWeight As Double
  Dim RateDesc() As String * 20
  Dim TempOptRev1Chrg As Integer
  Dim TempOptRev2Chrg As Integer
  Dim TempOptRev3Chrg As Integer
  Dim TempModYear As Integer
  Dim DontExit As Boolean
  Dim Opt1 As Integer
  Dim Opt2 As Integer
  Dim Opt3 As Integer
  
Private Sub cmdAdd_Click()
  On Error GoTo ERRORSTUFF
  
  If Check4Changes(WhichRec) = True Then
    Exit Sub
  End If
  
  If NumOfCustPPRecs = 0 Then
    WhichRec = 0
  Else
    WhichRec = NumOfCustPPRecs + 1
  End If
  
  Call LoadAdd(WhichRec)
  
  cmdAdd.Enabled = False
  cmdPageDown.Enabled = False
  cmdPageUp.Enabled = False
  cmdDelete.Enabled = False
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersProp", "cmdAdd_Click", Erl)
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
  Dim MobVal$
  Dim MerchVal$
  Dim FarmVal$
  Dim MachVal$
  Dim ThisBal As Double
  
  On Error GoTo ERRORSTUFF
  If Check4UnpostedBilling(PersRecs(WhichRec), "P") = True Then Exit Sub
  
  ThisPin$ = QPTrim$(fptxtPropPin.Text)
  If Val(ThisPin) > 0 Then
    ThisBal = GetPersBalance(ThisPin)
    If ThisBal <> 0 Then
      Me.ZOrder 0
      frmVATaxCustAddEdit.Visible = False
      If EditCust = True Then
        frmVATaxCustLookup.Visible = False
      End If
      If AddCust = True Then
        frmVATaxCustMaintMenu.Visible = False
      End If
      Call TaxMsg(900, "This property has an outstanding balance of " + QPTrim$(Using$("$###,###,##0.00", ThisBal)) + ". Please resolve this balance before deleting.")
      If EditCust = True Then
        frmVATaxCustLookup.Visible = True
      End If
      If AddCust = True Then
        frmVATaxCustMaintMenu.Visible = True
      End If
      frmVATaxCustAddEdit.Visible = True
      Me.Show
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
    fptxtPropPin.SetFocus
    Exit Sub
  Else
    Unload frmVATaxMsgWOpts
  End If
  CustName$ = QPTrim$(fptxtThisCust.Text)
  PersVal$ = QPTrim$(fpCurrPersVal.Text)
  MobVal$ = QPTrim$(fpCurrMobHome.Text)
  MerchVal$ = QPTrim$(fpCurrMerchCap.Text)
  FarmVal$ = QPTrim$(fpCurrFarmEq.Text)
  MachVal$ = QPTrim$(fpCurrMachTools.Text)
  
  Call DelPersAbstract(PersRecs(), WhichRec, GCustNum)
'  If PersRecs(0) = 0 Then
'    Unload Me
'    Exit Sub
'  End If
  Call GetPersRecList(PersRecs(), GCustNum, CustName)
  NumOfCustPPRecs = PersRecs(0)
  MainLog ("PERSONAL PROPERTY DELETION: User deleted the following personal property for : " + CustName + " - Pin # " + ThisPin + " - Personal Value: " + PersVal + " - Mobile Value: " + MobVal + " - Merchant Value: " + MerchVal + " - Farm Value: " + FarmVal + " - Machine Value: " + MachVal + ".")
  If PersRecs(0) = 0 Then
    WhichRec = 0
    Call LoadMe
  Else
    WhichRec = 1
    Call LoadAgain(WhichRec)
  End If
  frmVATaxMsg.Label1.Caption = "The personal property was deleted successfully."
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersProp", "cmdDelete_Click", Erl)
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

Private Sub cmdExit_Click()
  
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
    Me.Show
    frmVATaxCustLookup.Show
    frmVATaxCustAddEdit.Show
    frmVATaxCustAddEdit.Refresh '12/19/07
    If frmVATaxMsgWOpts.fptxtChoice.Text = "abort" Then
      Unload frmVATaxMsgWOpts
      Unload Me
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
      Call cmdSave_Click
      If DontExit = True Then
        DontExit = False
        Exit Sub
      Else
        Unload Me
        Exit Sub
      End If
    End If
  End If
  
'  frmVATaxCustLookup.Show
'  frmVATaxCustAddEdit.Show
  If Check4Changes(WhichRec) = True Then
    Exit Sub
  End If
  
  If DontExit = False Then
    Unload Me
  Else
    DontExit = False
  End If
  frmVATaxCustAddEdit.Refresh '12/19/07


End Sub

Private Sub cmdPageUp_Click()
  On Error GoTo ERRORSTUFF
  
  If Check4Changes(WhichRec) = True Then
    Exit Sub
  End If
  
  If WhichRec = NumOfCustPPRecs Then
    frmVATaxMsg.Label1.Caption = "Upper limit reached."
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersProp", "cmdPageUp_Click", Erl)
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
    frmVATaxMsg.Label1.Caption = "Lower limit reached."
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersProp", "cmdPageDown_Click", Erl)
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
  Dim PersPropRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
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
  
'  On Error GoTo ERRORSTUFF
  
  If QPTrim$(fptxtPropPin.Text) = "" Then
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
    fptxtPropPin.SetFocus
    DontExit = True
    Exit Sub
  End If
  
  If fpCurrPersVal.Value = 0 And fpCurrMobHome.Value = 0 And fpCurrMerchCap.Value = 0 And fpCurrFarmEq.Value = 0 And fpCurrMachTools.Value = 0 Then
    frmVATaxMsgWOpts.Label1.Caption = "No property values have been entered. Press F10 to save anyway. Otherwise, press ESC to abort the save procedure."
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
      fpCurrPersVal.SetFocus
      DontExit = True
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
    End If
  End If
  
  If WhichRec <= NumOfCustPPRecs Then
    If Check4DupPinsPers(QPTrim$(fptxtPropPin.Text), PersRecs(WhichRec)) = True Then
      Close
      Exit Sub
    End If
  Else
    If Check4DupPinsPers(QPTrim$(fptxtPropPin.Text), 0) = True Then
      Close
      Exit Sub
    End If
  End If

  If IsCustInPersPreBill(GCustNum) = True Then
    Me.ZOrder 0
    frmVATaxCustAddEdit.Visible = False
    If EditCust = True Then
      frmVATaxCustLookup.Visible = False
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = False
    End If
    If TaxMsgWOpts(600, "This customer is currently included in a personal property prebilling file. Continuing will delete the prebilling file. If you wish to continue then press F10. Otherwise, Press ESC to stop the save procedure.", "F10 Continue", "ESC Abort") = "abort" Then
      Close
      fptxtPropPin.SetFocus
      If EditCust = True Then
        frmVATaxCustLookup.Visible = True
      End If
      If AddCust = True Then
        frmVATaxCustMaintMenu.Visible = True
      End If
      frmVATaxCustAddEdit.Visible = True
      Me.Show
      Exit Sub
    Else
      KillFile PersTaxBillFile
      Me.ZOrder 0
      frmVATaxCustAddEdit.Visible = False
      If EditCust = True Then
        frmVATaxCustLookup.Visible = False
      End If
      If AddCust = True Then
        frmVATaxCustMaintMenu.Visible = False
      End If
      Call TaxMsg(900, "The personal property prebilling file has been deleted.")
      If EditCust = True Then
        frmVATaxCustLookup.Visible = True
      End If
      If AddCust = True Then
        frmVATaxCustMaintMenu.Visible = True
      End If
      frmVATaxCustAddEdit.Visible = True
      Me.Show
      MainLog ("Personal property prebilling file was deleted during save routine for personal property # " + CStr(WhichRec) + " after being warned.")
    End If
    If EditCust = True Then
      frmVATaxCustLookup.Visible = True
    End If
    If AddCust = True Then
      frmVATaxCustMaintMenu.Visible = True
    End If
    frmVATaxCustAddEdit.Visible = True
  End If
    
  OpenTaxCustFile THandle, NumOfCustRecs
  Get THandle, GCustNum, TaxRec
  CustPin& = TaxRec.PIN
  
  PersPropRec.PropPin = QPTrim$(fptxtPropPin.Text)
  PersPropRec.PROPDATE = Date2Num(fptxtDate.Text)
  PersPropRec.PersVal = fpCurrPersVal
  PersPropRec.MHValue = fpCurrMobHome
  PersPropRec.MCValue = fpCurrMerchCap
  PersPropRec.CVALUE = fpCurrFarmEq
  PersPropRec.MTValue = fpCurrMachTools
  PersPropRec.DISCOV = fpcmbDiscoveryYN.Text
  PersPropRec.LateList = fpcmbLateListYN.Text
  PersPropRec.OptSearch = fptxtOptSearch.Text
  PersPropRec.DESC1 = fptxtDesc(0).Text
  PersPropRec.DESC2 = fptxtDesc(1).Text
  PersPropRec.DESC3 = fptxtDesc(2).Text
  PersPropRec.Desc4 = fptxtDesc(3).Text
  PersPropRec.Desc5 = fptxtDesc(4).Text
  PersPropRec.Deleted = 0
  PersPropRec.CustPin = CustPin&
  PersPropRec.TaxBillYear = CInt(fptxtBillYear.Text)
  PersPropRec.PPTRAYN = QPTrim$(fpcmbPPTRAYN.Text)
  PersPropRec.Prorate = QPTrim$(fpcmbProRateYN.Text)
  PersPropRec.ProrateVal = CInt(fpcmbPRVal.Text)
  PersPropRec.Vin = QPTrim$(fptxtVin.Text)
  PersPropRec.MakeMod = QPTrim$(fptxtMakeMod.Text)
  PersPropRec.Weight = CDbl(fpDSWEight.Value)
  If QPTrim$(fpModYear.Text) = "" Then
    PersPropRec.ModYear = 0
  Else
    PersPropRec.ModYear = CInt(fpModYear.Text)
  End If
  If fpcmbOptRev1.Enabled = True Then
    fpcmbOptRev1.Col = 1
    If QPTrim$(fpcmbOptRev1.ColText) = "" Or QPTrim$(fpcmbOptRev1.ColText) = "0" Then 'added = "0" 10/22/07
      PersPropRec.OptRev1Chrg = 0
    Else
      PersPropRec.OptRev1Chrg = 4 'changed 10/22/07'CInt(fpcmbOptRev1.ColText)
    End If
  Else
    PersPropRec.OptRev1Chrg = 0
  End If
  If fpcmbOptRev2.Enabled = True Then
    fpcmbOptRev2.Col = 1
    If QPTrim$(fpcmbOptRev2.ColText) = "" Or QPTrim$(fpcmbOptRev2.ColText) = "0" Then 'added = "0" 10/22/07
      PersPropRec.OptRev2Chrg = 0
    Else
      PersPropRec.OptRev2Chrg = 5 'changed 10/22/07'CInt(fpcmbOptRev2.ColText)
    End If
  Else
    PersPropRec.OptRev2Chrg = 0
  End If
  If fpcmbOptRev3.Enabled = True Then
    fpcmbOptRev3.Col = 1
    If QPTrim$(fpcmbOptRev3.ColText) = "" Or QPTrim$(fpcmbOptRev3.ColText) = "0" Then 'added = "0" 10/22/07
      PersPropRec.OptRev3Chrg = 0
    Else
      PersPropRec.OptRev3Chrg = 6 'changed 10/22/07'CInt(fpcmbOptRev3.ColText)
    End If
  Else
    PersPropRec.OptRev3Chrg = 0
  End If
  
  OpenPersPropFile PHandle, NumOfPersRecs
  
  WhatPers = NumOfPersRecs + 1

  If WhichRec = 0 Then 'first pers prop record for this customer
    PersPropRec.LastYrPrinted = 0
    PersPropRec.VehTaxYear = 0
    PersPropRec.DMVSubmitted = "N"
    PersPropRec.Blank = ""
    TaxRec.FirstPersRec = WhatPers&
    Put THandle, GCustNum, TaxRec
    Close THandle
    ReDim Preserve PersRecs(0 To 1) As Long
    PersRecs(0) = 1 '# of props for this customer
    PersRecs(1) = WhatPers 'record # for this prop
    NumOfCustPPRecs = 1
    PersPropRec.NextRec = 0
    Put PHandle, WhatPers, PersPropRec
    fptxtRecord.Text = CStr(NumOfCustPPRecs) + " of " + CStr(NumOfCustPPRecs)
  ElseIf WhichRec > NumOfCustPPRecs Then 'adding to existing pers prop
    NumOfCustPPRecs = NumOfCustPPRecs + 1
    ReDim Preserve PersRecs(0 To WhichRec) As Long
    PersRecs(0) = PersRecs(0) + 1
    PersRecs(WhichRec) = WhatPers
    PersPropRec.NextRec = 0
    Put PHandle, WhatPers, PersPropRec
    Get PHandle, PersRecs(NumOfCustPPRecs - 1), PersPropRec
    PersPropRec.NextRec = WhatPers
    Put PHandle, PersRecs(NumOfCustPPRecs - 1), PersPropRec
    fptxtRecord.Text = CStr(NumOfCustPPRecs) + " of " + CStr(NumOfCustPPRecs)
  Else 'editing existing data
    Get PHandle, PersRecs(WhichRec), PersPropRec
    PersPropRec.PropPin = QPTrim$(fptxtPropPin.Text)
    PersPropRec.PROPDATE = Date2Num(fptxtDate.Text)
    PersPropRec.PersVal = fpCurrPersVal
    PersPropRec.MHValue = fpCurrMobHome
    PersPropRec.MCValue = fpCurrMerchCap
    PersPropRec.CVALUE = fpCurrFarmEq
    PersPropRec.MTValue = fpCurrMachTools
    PersPropRec.DISCOV = fpcmbDiscoveryYN.Text
    PersPropRec.LateList = fpcmbLateListYN.Text
    PersPropRec.OptSearch = fptxtOptSearch.Text
    PersPropRec.DESC1 = fptxtDesc(0).Text
    PersPropRec.DESC2 = fptxtDesc(1).Text
    PersPropRec.DESC3 = fptxtDesc(2).Text
    PersPropRec.Desc4 = fptxtDesc(3).Text
    PersPropRec.Desc5 = fptxtDesc(4).Text
    PersPropRec.TaxBillYear = CInt(fptxtBillYear.Text)
    PersPropRec.PPTRAYN = QPTrim$(fpcmbPPTRAYN.Text)
    PersPropRec.Prorate = QPTrim$(fpcmbProRateYN.Text)
    PersPropRec.ProrateVal = CInt(fpcmbPRVal.Text)
    PersPropRec.Vin = QPTrim$(fptxtVin.Text)
    PersPropRec.MakeMod = QPTrim$(fptxtMakeMod.Text)
    PersPropRec.Weight = CDbl(fpDSWEight.Value)
    If QPTrim$(fpModYear.Text) = "" Then
      PersPropRec.ModYear = 0
    Else
      PersPropRec.ModYear = CInt(fpModYear.Text)
    End If
    If fpcmbOptRev1.Enabled = True Then
      fpcmbOptRev1.Col = 1
      If QPTrim$(fpcmbOptRev1.ColText) = "" Or QPTrim$(fpcmbOptRev1.ColText) = "0" Then 'added = "0" 10/22/07
        PersPropRec.OptRev1Chrg = 0
      Else
        PersPropRec.OptRev1Chrg = 4 'changed 10/22/07'CInt(fpcmbOptRev1.ColText)
      End If
    Else
      PersPropRec.OptRev1Chrg = 0
    End If
    If fpcmbOptRev2.Enabled = True Then
      fpcmbOptRev2.Col = 1
      If QPTrim$(fpcmbOptRev2.ColText) = "" Or QPTrim$(fpcmbOptRev2.ColText) = "0" Then 'added = "0" 10/22/07
        PersPropRec.OptRev2Chrg = 0
      Else
        PersPropRec.OptRev2Chrg = 5 'changed 10/22/07 CInt(fpcmbOptRev2.ColText)
      End If
    Else
      PersPropRec.OptRev2Chrg = 0
    End If
    If fpcmbOptRev3.Enabled = True Then
      fpcmbOptRev3.Col = 1
      If QPTrim$(fpcmbOptRev3.ColText) = "" Or QPTrim$(fpcmbOptRev3.ColText) = "0" Then 'added = "0" 10/22/07
        PersPropRec.OptRev3Chrg = 0
       Else
        PersPropRec.OptRev3Chrg = 6 'changed 10/22/07 CInt(fpcmbOptRev3.ColText)
      End If
    Else
      PersPropRec.OptRev3Chrg = 0
    End If
    Put PHandle, PersRecs(WhichRec), PersPropRec
    Call LogSaves
  End If
  
  Close PHandle
  Close THandle
  
  ReDim PersRecs(0 To 0) As Long
  Call GetPersRecList(PersRecs(), GCustNum, CustName)
  
  Call MakePersPINFile
  
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
    Call CreateOptPersIdx
  End If
  Call Savemsg(900, "Your personal property data has been saved.")
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersProp", "cmdSave_Click", Erl)
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
  Me.HelpContextID = hlpPersonal
  Call LoadMe
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxPersProp.")
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
  Dim PersPropRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim x As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim TblRec As OptRevRateTablesType
  Dim TRHandle As Integer
  Dim NumOfTRRecs As Integer
  Dim ThisDesc As String * 20
  
'  On Error GoTo ERRORSTUFF
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  fptxtOptSearchDesc.Text = QPTrim$(TaxMasterRec.OptSrchPers)
  
  If TaxMasterRec.PPTRAYN = "N" Then
    Label12.Caption = "PPTRA NA:"
  End If
  fpcmbDiscoveryYN.AddItem "Y"
  fpcmbDiscoveryYN.AddItem "N"
  fpcmbLateListYN.AddItem "Y"
  fpcmbLateListYN.AddItem "N"
  fpcmbProRateYN.AddItem "Y"
  fpcmbProRateYN.AddItem "N"
  fpcmbPRVal.AddItem "0"
  fpcmbPRVal.AddItem "1"
  fpcmbPRVal.AddItem "2"
  fpcmbPRVal.AddItem "3"
  fpcmbPRVal.AddItem "4"
  fpcmbPRVal.AddItem "5"
  fpcmbPRVal.AddItem "6"
  fpcmbPRVal.AddItem "7"
  fpcmbPRVal.AddItem "8"
  fpcmbPRVal.AddItem "9"
  fpcmbPRVal.AddItem "10"
  fpcmbPRVal.AddItem "11"
  fpcmbPRVal.AddItem "12"
  fpcmbPPTRAYN.AddItem "Y"
  fpcmbPPTRAYN.AddItem "N"
  fptxtBillYear.Text = Mid(Date, 7, 4)
'  For x = 1 To 51
'    fptxtBillYear.AddItem CStr(1979 + x)
'  Next x
  ReDim PersRecs(0 To 0) As Long
  Call GetPersRecList(PersRecs(), GCustNum, CustName)
  fptxtThisCust.Text = CustName
  NumOfCustPPRecs = PersRecs(0)
  
  Opt1 = 0
  Opt2 = 0
  Opt3 = 0
  If Exist(TaxRateTableFile) Then
    OpenTaxRateTables TRHandle, NumOfTRRecs
    ReDim RateDesc(0 To NumOfTRRecs) As String * 20
    RateDesc(0) = "NOT IN USE"
    fpcmbOptRev1.InsertRow = QPTrim$(RateDesc(0)) + Chr(9) + CStr(0)
    fpcmbOptRev2.InsertRow = QPTrim$(RateDesc(0)) + Chr(9) + CStr(0)
    fpcmbOptRev3.InsertRow = QPTrim$(RateDesc(0)) + Chr(9) + CStr(0)
    For x = 1 To NumOfTRRecs
      Get TRHandle, x, TblRec
      If TblRec.Deleted = True And TblRec.RevType = "P" Then
        fpcmbOptRev1.ToolTipText = "Please note: If there is no access to the optional revenue in the drop down list then the rate settings for this option have not been set."
        fpcmbOptRev2.ToolTipText = "Please note: If there is no access to the optional revenue in the drop down list then the rate settings for this option have not been set."
        fpcmbOptRev3.ToolTipText = "Please note: If there is no access to the optional revenue in the drop down list then the rate settings for this option have not been set."
      End If
      If TblRec.Deleted = True Or TblRec.RevType = "R" Then GoTo Deleted
      RateDesc(x) = QPTrim$(TblRec.Desc)
      If TblRec.OptRevNum = 4 Then
        fpcmbOptRev1.InsertRow = RateDesc(x) + Chr(9) + CStr(x)
        Opt1 = x
      ElseIf TblRec.OptRevNum = 5 Then
        fpcmbOptRev2.InsertRow = RateDesc(x) + Chr(9) + CStr(x)
        Opt2 = x
      ElseIf TblRec.OptRevNum = 6 Then
        fpcmbOptRev3.InsertRow = RateDesc(x) + Chr(9) + CStr(x)
        Opt3 = x
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
  If NumOfCustPPRecs = 0 Then
    WhichRec = 0
    fptxtPropPin.Text = ""
    fptxtRecord.Text = "None Saved"
    lblMode.Caption = "Mode: Adding"
    fptxtDate.Text = Date
    fpcmbDiscoveryYN.Text = "N"
    fpcmbLateListYN.Text = "N"
    fptxtOptSearch.Text = ""
    fpCurrPersVal = 0
    fpCurrMobHome = 0
    fpCurrMerchCap = 0
    fpCurrFarmEq = 0
    fpCurrMachTools = 0
'    fpCurrSnCitizen = 0
'    fpCurrOther = 0
    fptxtDesc(0).Text = ""
    fptxtDesc(1).Text = ""
    fptxtDesc(2).Text = ""
    fptxtDesc(3).Text = ""
    fptxtDesc(4).Text = ""
    fpcmbProRateYN.Text = "N"
    fpcmbPRVal.Text = "0"
    fpcmbPPTRAYN.Text = "N"
    fptxtVin.Text = ""
    fptxtMakeMod.Text = ""
    fpDSWEight.Value = 0
    fpModYear.Text = Mid(Date, 7, 4)
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
    OpenPersPropFile PHandle, NumOfPersRecs
    Get PHandle, PersRecs(1), PersPropRec
    Close PHandle
    WhichRec = 1
    fptxtRecord.Text = "1 of " + CStr(NumOfCustPPRecs)
    lblMode.Caption = "Mode: Editing"
    fptxtDate.Text = MakeRegDate(PersPropRec.PROPDATE)
    fptxtPropPin.Text = QPTrim$(PersPropRec.PropPin)
    If PersPropRec.DISCOV <> "Y" Then
      fpcmbDiscoveryYN.Text = "N"
    Else
      fpcmbDiscoveryYN.Text = "Y"
    End If
    If PersPropRec.LateList <> "Y" Then
      fpcmbLateListYN.Text = "N"
    Else
      fpcmbLateListYN.Text = "Y"
    End If
    fptxtOptSearch.Text = QPTrim$(PersPropRec.OptSearch)
    fpCurrPersVal = PersPropRec.PersVal
    fpCurrMobHome = PersPropRec.MHValue
    fpCurrMerchCap = PersPropRec.MCValue
    fpCurrFarmEq = PersPropRec.CVALUE
    fpCurrMachTools = PersPropRec.MTValue
'    fpCurrSnCitizen = PersPropRec.EXMPSENI
'    fpCurrOther = PersPropRec.EXMPOTHR
    fptxtDesc(0).Text = PersPropRec.DESC1
    fptxtDesc(1).Text = PersPropRec.DESC2
    fptxtDesc(2).Text = PersPropRec.DESC3
    fptxtDesc(3).Text = PersPropRec.Desc4
    fptxtDesc(4).Text = PersPropRec.Desc5
    fpcmbProRateYN.Text = PersPropRec.Prorate
    fpcmbPRVal.Text = PersPropRec.ProrateVal
    fpcmbPPTRAYN.Text = PersPropRec.PPTRAYN
    fptxtBillYear.Text = PersPropRec.TaxBillYear
    fptxtVin.Text = QPTrim$(PersPropRec.Vin)
    fptxtMakeMod.Text = QPTrim$(PersPropRec.MakeMod)
    fpDSWEight.Value = PersPropRec.Weight
    If PersPropRec.ModYear = 0 Then
      fpModYear.Text = ""
    Else
      fpModYear.Text = CStr(PersPropRec.ModYear)
    End If
    If PersPropRec.OptRev1Chrg > 0 Then
      fpcmbOptRev1.SearchText = RateDesc(Opt1) + Chr(9) + CStr(Opt1)
    Else
      fpcmbOptRev1.SearchText = RateDesc(PersPropRec.OptRev1Chrg) + Chr(9) + CStr(PersPropRec.OptRev1Chrg)
    End If
    fpcmbOptRev1.Action = 0
12:
    If fpcmbOptRev1.SearchIndex <> -1 Then
      fpcmbOptRev1.ListIndex = fpcmbOptRev1.SearchIndex
    Else
      fpcmbOptRev1.ListIndex = 0
    End If
    If PersPropRec.OptRev2Chrg > 0 Then
      fpcmbOptRev2.SearchText = RateDesc(Opt2) + Chr(9) + CStr(Opt2)
    Else
      fpcmbOptRev2.SearchText = RateDesc(PersPropRec.OptRev2Chrg) + Chr(9) + CStr(PersPropRec.OptRev2Chrg)
    End If
    fpcmbOptRev2.Action = 0
    If fpcmbOptRev2.SearchIndex <> -1 Then
      fpcmbOptRev2.ListIndex = fpcmbOptRev2.SearchIndex
    Else
      fpcmbOptRev2.ListIndex = 0
    End If
    If PersPropRec.OptRev3Chrg > 0 Then
      fpcmbOptRev3.SearchText = RateDesc(Opt3) + Chr(9) + CStr(Opt3)
    Else
      fpcmbOptRev3.SearchText = RateDesc(PersPropRec.OptRev3Chrg) + Chr(9) + CStr(PersPropRec.OptRev3Chrg)
    End If
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
    Call AssignTemps
  End If
  Call LoadGoToPinsCmb
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersProp", "LoadMe", Erl)
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

Public Sub LoadAgain(WhichRec)
  Dim PersPropRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  
  On Error GoTo ERRORSTUFF
  
  OpenPersPropFile PHandle, NumOfPersRecs
  Get PHandle, PersRecs(WhichRec), PersPropRec
  Close PHandle
  fptxtRecord.Text = CStr(WhichRec) + " of " + CStr(NumOfCustPPRecs)
  lblMode.Caption = "Mode: Editing"
  fptxtDate.Text = MakeRegDate(PersPropRec.PROPDATE)
  fptxtPropPin.Text = QPTrim$(PersPropRec.PropPin)
  If PersPropRec.DISCOV <> "Y" Then
    fpcmbDiscoveryYN.Text = "N"
  Else
    fpcmbDiscoveryYN.Text = "Y"
  End If
  If PersPropRec.LateList <> "Y" Then
    fpcmbLateListYN.Text = "N"
  Else
    fpcmbLateListYN.Text = "Y"
  End If
  fptxtOptSearch.Text = QPTrim$(PersPropRec.OptSearch)
  fpCurrPersVal = PersPropRec.PersVal
  fpCurrMobHome = PersPropRec.MHValue
  fpCurrMerchCap = PersPropRec.MCValue
  fpCurrFarmEq = PersPropRec.CVALUE
  fpCurrMachTools = PersPropRec.MTValue
'  fpCurrSnCitizen = PersPropRec.EXMPSENI
'  fpCurrOther = PersPropRec.EXMPOTHR
  fptxtDesc(0).Text = PersPropRec.DESC1
  fptxtDesc(1).Text = PersPropRec.DESC2
  fptxtDesc(2).Text = PersPropRec.DESC3
  fptxtDesc(3).Text = PersPropRec.Desc4
  fptxtDesc(4).Text = PersPropRec.Desc5
  fpcmbProRateYN.Text = PersPropRec.Prorate
  fpcmbPRVal.Text = PersPropRec.ProrateVal
  fpcmbPPTRAYN.Text = PersPropRec.PPTRAYN
  fptxtBillYear.Text = PersPropRec.TaxBillYear
  fptxtVin.Text = QPTrim$(PersPropRec.Vin)
  fptxtMakeMod.Text = QPTrim$(PersPropRec.MakeMod)
  fpDSWEight.Value = PersPropRec.Weight
  If PersPropRec.ModYear = 0 Then
    fpModYear.Text = ""
  Else
    fpModYear.Text = CStr(PersPropRec.ModYear)
  End If
  If PersPropRec.OptRev1Chrg > 0 Then
    fpcmbOptRev1.SearchText = RateDesc(Opt1) + Chr(9) + CStr(Opt1)
  Else
    fpcmbOptRev1.SearchText = RateDesc(PersPropRec.OptRev1Chrg) + Chr(9) + CStr(PersPropRec.OptRev1Chrg)
  End If
  fpcmbOptRev1.Action = 0
  If fpcmbOptRev1.SearchIndex <> -1 Then
    fpcmbOptRev1.ListIndex = fpcmbOptRev1.SearchIndex
  Else
    fpcmbOptRev1.ListIndex = 0
  End If
  If PersPropRec.OptRev2Chrg > 0 Then
    fpcmbOptRev2.SearchText = RateDesc(Opt2) + Chr(9) + CStr(Opt2)
  Else
    fpcmbOptRev2.SearchText = RateDesc(PersPropRec.OptRev2Chrg) + Chr(9) + CStr(PersPropRec.OptRev2Chrg)
  End If
  fpcmbOptRev2.Action = 0
  If fpcmbOptRev2.SearchIndex <> -1 Then
    fpcmbOptRev2.ListIndex = fpcmbOptRev2.SearchIndex
  Else
    fpcmbOptRev2.ListIndex = 0
  End If
  If PersPropRec.OptRev3Chrg > 0 Then
    fpcmbOptRev3.SearchText = RateDesc(Opt3) + Chr(9) + CStr(Opt3)
  Else
    fpcmbOptRev3.SearchText = RateDesc(PersPropRec.OptRev3Chrg) + Chr(9) + CStr(PersPropRec.OptRev3Chrg)
  End If
  fpcmbOptRev3.Action = 0
  If fpcmbOptRev3.SearchIndex <> -1 Then
    fpcmbOptRev3.ListIndex = fpcmbOptRev3.SearchIndex
  Else
    fpcmbOptRev3.ListIndex = 0
  End If
  Call AssignTemps
  Call LoadGoToPinsCmb
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersProp", "LoadAgain", Erl)
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
  
  If NumOfCustPPRecs > 0 Then
    fptxtRecord.Text = "Adding Record # " + CStr(WhichRec)
  Else
    fptxtRecord.Text = "Adding 1st Record"
  End If
  lblMode.Caption = "Mode: Adding"
  fptxtDate.Text = Date
  fptxtPropPin.Text = ""
  fpcmbDiscoveryYN.Text = "N"
  fpcmbLateListYN.Text = "N"
  fptxtOptSearch.Text = ""
  fpCurrPersVal = 0
  fpCurrMobHome = 0
  fpCurrMerchCap = 0
  fpCurrFarmEq = 0
  fpCurrMachTools = 0
'  fpCurrSnCitizen = 0
'  fpCurrOther = 0
  fptxtDesc(0).Text = ""
  fptxtDesc(1).Text = ""
  fptxtDesc(2).Text = ""
  fptxtDesc(3).Text = ""
  fptxtDesc(4).Text = ""
  fpcmbProRateYN.Text = "N"
  fpcmbPRVal.Text = "0"
  fpcmbPPTRAYN.Text = "N"
  fptxtBillYear.Text = Mid(Date, 7, 4)
  fptxtVin.Text = ""
  fptxtMakeMod.Text = ""
  fpDSWEight.Value = 0
  fpModYear.Text = Mid(Date, 7, 4)
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
  Call LoadGoToPinsCmb
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersProp", "LoadAdd", Erl)
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

'Private Sub fptxtBillYear_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeySpace Then
'    fptxtBillYear.ListDown = True
'  End If
'  If KeyCode = vbKeyDelete Then
'    fptxtBillYear.ListIndex = -1
'  End If
'  If fptxtBillYear.ListDown <> True Then
'    If KeyCode = vbKeyDown Then
'      fptxtVin.SetFocus
'      KeyCode = 0
'    Else
'      If KeyCode = vbKeyUp Then
'        SendKeys "+{Tab}"
'        KeyCode = 0
'      End If
'    End If
'  End If
'
'End Sub

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

Private Sub fpcmbLateListYN_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbLateListYN.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbLateListYN.ListIndex = -1
  End If
  If fpcmbLateListYN.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbProRateYN.SetFocus
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
      fpcmbOptRev2.SetFocus
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
      fpcmbOptRev3.SetFocus
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
      fptxtVin.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbPPTRAYN_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbPPTRAYN.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPPTRAYN.ListIndex = -1
  End If
  If fpcmbPPTRAYN.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtVin.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbProRateYN_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbProRateYN.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbProRateYN.ListIndex = -1
  End If
  If fpcmbProRateYN.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbPRVal.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If


End Sub

Private Sub fpcmbPRVal_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbPRVal.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPRVal.ListIndex = -1
  End If
  If fpcmbPRVal.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbPPTRAYN.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fptxtDesc_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  If Index = 4 Then
    If KeyCode = vbKeyDown Then
      fptxtPropPin.SetFocus
    ElseIf KeyCode = vbKeyUp Then
      fptxtDesc(3).SetFocus
    End If
  End If
End Sub

Private Sub fptxtPropPin_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fptxtDate.SetFocus
  ElseIf KeyCode = vbKeyUp Then
    fptxtDesc(4).SetFocus
  End If
End Sub

Private Function Check4Changes(WhichRec) As Boolean
  Dim ThisControl As Control
  Dim ThisDesc$
  Dim ThatDesc$
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim ThisDbl#
  Dim ThatDbl#
  Dim choice$
  Dim NoEntry As Boolean
  Dim ThisInt As Integer
  Dim ThatInt As Integer
  
  On Error GoTo ERRORSTUFF
  Check4Changes = False
  NoEntry = True
  If PersRecs(WhichRec) > 0 Then
    OpenPersPropFile PHandle, NumOfPersRecs
    Get PHandle, PersRecs(WhichRec), PersRec
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
    
  Set ThisControl = fptxtPropPin
  ThisDesc = QPTrim$(fptxtPropPin.Text)
  ThatDesc = TempPROPPIN$
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
      PersRec.PropPin = QPTrim$(ThisControl.Text)
      Put PHandle, PersRecs(WhichRec), PersRec
      MsgBox "The Pin Number has been saved successfully."
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
      PersRec.PROPDATE = Date2Num(ThisDesc)
      Put PHandle, PersRecs(WhichRec), PersRec
      MsgBox "The Date has been saved successfully."
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fpCurrPersVal
  ThisDbl = fpCurrPersVal.Value
  ThatDbl = TempPersVal#
  If ThatDbl <> ThisDbl Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Personal Value' field has been changed from " + QPTrim$(Using$("$###,##0.00", ThatDbl)) + " to " + QPTrim$(Using$("$###,##0.00", ThisDbl)) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
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
      PersRec.PersVal = ThisDbl
      Put PHandle, PersRecs(WhichRec), PersRec
      MsgBox "Personal Value has been saved successfully."
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fpCurrMobHome
  ThisDbl = fpCurrMobHome.Value
  ThatDbl = TempMHVALUE#
  If ThatDbl <> ThisDbl Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Mobile Home Value' field has been changed from " + QPTrim$(Using$("$###,##0.00", ThatDbl)) + " to " + QPTrim$(Using$("$###,##0.00", ThisDbl)) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
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
      PersRec.MHValue = ThisDbl
      Put PHandle, PersRecs(WhichRec), PersRec
      MsgBox "Mobile Home Value has been saved successfully."
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fpCurrMerchCap
  ThisDbl = fpCurrMerchCap.Value
  ThatDbl = TempMCVALUE#
  If ThatDbl <> ThisDbl Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Merchant Capital Value' field has been changed from " + QPTrim$(Using$("$###,##0.00", ThatDbl)) + " to " + QPTrim$(Using$("$###,##0.00", ThisDbl)) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
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
      PersRec.MCValue = ThisDbl
      Put PHandle, PersRecs(WhichRec), PersRec
      MsgBox "Merchant Capital Value has been saved successfully."
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fpCurrFarmEq
  ThisDbl = fpCurrFarmEq.Value
  ThatDbl = TempCVALUE#
  If ThatDbl <> ThisDbl Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Farm Equipment Value' field has been changed from " + QPTrim$(Using$("$###,##0.00", ThatDbl)) + " to " + QPTrim$(Using$("$###,##0.00", ThisDbl)) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
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
      PersRec.CVALUE = ThisDbl
      Put PHandle, PersRecs(WhichRec), PersRec
      MsgBox "Farm Equipment Value has been saved successfully."
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fpCurrMachTools
  ThisDbl = fpCurrMachTools.Value
  ThatDbl = TempMTVALUE#
  If ThatDbl <> ThisDbl Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Machine Tools Value' field has been changed from " + QPTrim$(Using$("$###,##0.00", ThatDbl)) + " to " + QPTrim$(Using$("$###,##0.00", ThisDbl)) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
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
      PersRec.MTValue = ThisDbl
      Put PHandle, PersRecs(WhichRec), PersRec
      MsgBox "Machine Tools Value has been saved successfully."
    Else
      GoSub HandleChoice
    End If
  End If
     
'  Set ThisControl = fpCurrSnCitizen
'  ThisDbl = fpCurrSnCitizen.Value
'  ThatDbl = TempEXMPSENI#
'  If ThatDbl <> ThisDbl Then
'    frmVATaxMsgW4Opts.Label1.Caption = "The 'Senior Citizen' field has been changed from " + QPTrim$(Using$("$###,##0.00", ThatDbl)) + " to " + QPTrim$(Using$("$###,##0.00", ThisDbl)) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
'    frmVATaxMsgW4Opts.Label1.Top = 575
'    Me.ZOrder 0
'    frmVATaxCustAddEdit.Visible = False
'    If EditCust = True Then
'      frmVATaxCustLookup.Visible = False
'    End If
'    If AddCust = True Then
'      frmVATaxCustMaintMenu.Visible = False
'    End If
'    frmVATaxMsgW4Opts.Show vbModal
'    If EditCust = True Then
'      frmVATaxCustLookup.Visible = True
'    End If
'    If AddCust = True Then
'      frmVATaxCustMaintMenu.Visible = True
'    End If
'    frmVATaxCustAddEdit.Visible = True
'    Me.Show
'    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
'    Unload frmVATaxMsgW4Opts
'    If choice = "save" Then
'      PersRec.EXMPSENI = ThisDbl
'      Put PHandle, PersRecs(WhichRec), PersRec
'      MsgBox "Senior Citizen has been saved successfully."
'    Else
'      GoSub HandleChoice
'    End If
'  End If
'
'  Set ThisControl = fpCurrOther
'  ThisDbl = fpCurrOther.Value
'  ThatDbl = TempEXMPOTHR#
'  If ThatDbl <> ThisDbl Then
'    frmVATaxMsgW4Opts.Label1.Caption = "The 'Other' field has been changed from " + QPTrim$(Using$("$###,##0.00", ThatDbl)) + " to " + QPTrim$(Using$("$###,##0.00", ThisDbl)) + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
'    frmVATaxMsgW4Opts.Label1.Top = 575
'    Me.ZOrder 0
'    frmVATaxCustAddEdit.Visible = False
'    If EditCust = True Then
'      frmVATaxCustLookup.Visible = False
'    End If
'    If AddCust = True Then
'      frmVATaxCustMaintMenu.Visible = False
'    End If
'    frmVATaxMsgW4Opts.Show vbModal
'    If EditCust = True Then
'      frmVATaxCustLookup.Visible = True
'    End If
'    If AddCust = True Then
'      frmVATaxCustMaintMenu.Visible = True
'    End If
'    frmVATaxCustAddEdit.Visible = True
'    Me.Show
'    choice = frmVATaxMsgW4Opts.fptxtChoice.Text
'    Unload frmVATaxMsgW4Opts
'    If choice = "save" Then
'      PersRec.EXMPOTHR = ThisDbl
'      Put PHandle, PersRecs(WhichRec), PersRec
'      MsgBox "Other has been saved successfully."
'    Else
'      GoSub HandleChoice
'    End If
'  End If
     
  Set ThisControl = fpcmbDiscoveryYN
  ThisDesc = QPTrim$(fpcmbDiscoveryYN.Text)
  ThatDesc = TempDISCOV$
  If ThatDesc <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Discovery (Y/N)' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
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
      PersRec.DISCOV = QPTrim$(ThisControl.Text)
      Put PHandle, PersRecs(WhichRec), PersRec
      MsgBox "Discovery (Y/N) has been saved successfully."
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fpcmbLateListYN
  ThisDesc = QPTrim$(fpcmbLateListYN.Text)
  ThatDesc = TempLATELIST$
  If ThatDesc <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Late List (Y/N)' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
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
      PersRec.LateList = QPTrim$(ThisControl.Text)
      Put PHandle, PersRecs(WhichRec), PersRec
      MsgBox "Late List (Y/N) has been saved successfully."
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fptxtOptSearch
  ThisDesc = QPTrim$(fptxtOptSearch.Text)
  ThatDesc = QPTrim$(TempOptSearch$)
  If ThatDesc <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Optional Search' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
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
      PersRec.OptSearch = QPTrim$(ThisControl.Text)
      Put PHandle, PersRecs(WhichRec), PersRec
      Call CreateOptPersIdx
      Call Savemsg(900, "Optional Search Name has been saved successfully.")
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fptxtDesc(0)
  ThisDesc = fptxtDesc(0).Text
  ThatDesc = TempDESC1$
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
      PersRec.DESC1 = QPTrim$(ThisControl.Text)
      Put PHandle, PersRecs(WhichRec), PersRec
      MsgBox "Description Line #1 has been saved successfully."
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fptxtDesc(1)
  ThisDesc = fptxtDesc(1).Text
  ThatDesc = TempDESC2$
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
      PersRec.DESC2 = QPTrim$(ThisControl.Text)
      Put PHandle, PersRecs(WhichRec), PersRec
      MsgBox "Description Line #2 has been saved successfully."
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fptxtDesc(2)
  ThisDesc = fptxtDesc(2).Text
  ThatDesc = TempDESC3$
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
      PersRec.DESC3 = QPTrim$(ThisControl.Text)
      Put PHandle, PersRecs(WhichRec), PersRec
      MsgBox "Description Line #3 has been saved successfully."
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fptxtDesc(3)
  ThisDesc = fptxtDesc(3).Text
  ThatDesc = TempDesc4$
  If ThatDesc <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Description Line #4' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
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
      PersRec.Desc4 = QPTrim$(ThisControl.Text)
      Put PHandle, PersRecs(WhichRec), PersRec
      MsgBox "Description Line #4 has been saved successfully."
    Else
      GoSub HandleChoice
    End If
  End If
     
  Set ThisControl = fptxtDesc(4)
  ThisDesc = fptxtDesc(4).Text
  ThatDesc = TempDesc5$
  If ThatDesc <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Description Line #5' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
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
      PersRec.Desc5 = QPTrim$(ThisControl.Text)
      Put PHandle, PersRecs(WhichRec), PersRec
      MsgBox "Description Line #5 has been saved successfully."
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpcmbProRateYN
  ThisDesc = QPTrim$(fpcmbProRateYN.Text)
  ThatDesc = QPTrim$(TempProrate$)
  If ThatDesc <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Prorate (Y/N)' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
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
      PersRec.Prorate = QPTrim$(ThisControl.Text)
      Put PHandle, PersRecs(WhichRec), PersRec
      MsgBox "Prorate (Y/N) has been saved successfully."
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpcmbPRVal
  ThisDesc = QPTrim$(fpcmbPRVal.Text)
  ThatDesc = CStr(TempProrateVal)
  If ThatDesc <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Prorate Value' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
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
      PersRec.ProrateVal = CInt(ThisControl.Text)
      Put PHandle, PersRecs(WhichRec), PersRec
      MsgBox "Prorate Value has been saved successfully."
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpcmbPPTRAYN
  ThisDesc = QPTrim$(fpcmbPPTRAYN.Text)
  ThatDesc = TempPPTRAYN
  If ThatDesc <> ThisDesc And Label12 <> "PPTRA NA:" Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'PPTRA (Y/N)' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
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
      PersRec.PPTRAYN = QPTrim$(ThisControl.Text)
      Put PHandle, PersRecs(WhichRec), PersRec
      MsgBox "PPTRA (Y/N) has been saved successfully."
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtBillYear
  ThisDesc = QPTrim$(fptxtBillYear.Text)
  ThatDesc = CStr(TempTaxBillYear)
  If ThatDesc <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Bill Year' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
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
      PersRec.TaxBillYear = CInt(ThisControl.Text)
      Put PHandle, PersRecs(WhichRec), PersRec
      MsgBox "Bill Year has been saved successfully."
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
        PersRec.OptRev1Chrg = ThisInt
        Put PHandle, PersRecs(WhichRec), PersRec
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
        PersRec.OptRev2Chrg = ThisInt
        Put PHandle, PersRecs(WhichRec), PersRec
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
        PersRec.OptRev3Chrg = ThisInt
        Put PHandle, PersRecs(WhichRec), PersRec
        Call Savemsg(900, QPTrim$(Label21.Caption) + " has been saved successfully.")
      Else
        GoSub HandleChoice
      End If
    End If
  End If
  
  
  
  Set ThisControl = fptxtVin
  ThisDesc = QPTrim$(fptxtVin.Text)
  ThatDesc = TempVin
  If ThatDesc <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'VIN #' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
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
      PersRec.Vin = QPTrim$(ThisControl.Text)
      Put PHandle, PersRecs(WhichRec), PersRec
      MsgBox "VIN # has been saved successfully."
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtMakeMod
  ThisDesc = QPTrim$(fptxtMakeMod.Text)
  ThatDesc = TempMakeMod
  If ThatDesc <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Make/Model' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
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
      PersRec.MakeMod = QPTrim$(ThisControl.Text)
      Put PHandle, PersRecs(WhichRec), PersRec
      MsgBox "Make/Model has been saved successfully."
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpDSWEight
  ThisDesc = CStr(fpDSWEight.Value)
  ThatDesc = CStr(TempWeight)
  If Val(ThatDesc) <> Val(ThisDesc) Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Weight' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
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
      PersRec.Weight = CDbl(ThisControl.Text)
      Put PHandle, PersRecs(WhichRec), PersRec
      MsgBox "Weight has been saved successfully."
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpModYear
  ThisDesc = fpModYear.Text
  If TempModYear = 0 Then
    ThatDesc = ""
  Else
    ThatDesc = CStr(TempModYear)
  End If
  If ThatDesc <> ThisDesc Then
    frmVATaxMsgW4Opts.Label1.Caption = "The 'Model Year' field has been changed from " + ThatDesc + " to " + ThisDesc + ". Press F10 to save the change. Press F5 to review the change. Press F3 to abandon this change. Otherwise, press ESC to abandon all changes."
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
      If QPTrim$(ThisControl.Text) = "" Then
        PersRec.ModYear = 0
      Else
        PersRec.ModYear = CInt(ThisControl.Text)
      End If
      Put PHandle, PersRecs(WhichRec), PersRec
      MsgBox "Model Year has been saved successfully."
    Else
      GoSub HandleChoice
    End If
  End If
  
  Exit Function
  
EntryCheck:
  If QPTrim$(fptxtPropPin.Text) <> "" Then
    NoEntry = False
    Return
  ElseIf fpCurrPersVal.Value <> 0 Then
    NoEntry = False
    Return
  ElseIf fpCurrMobHome.Value <> 0 Then
    NoEntry = False
    Return
  ElseIf fpCurrMerchCap.Value <> 0 Then
    NoEntry = False
    Return
  ElseIf fpCurrFarmEq.Value <> 0 Then
    NoEntry = False
    Return
  ElseIf fpCurrMachTools.Value <> 0 Then
    NoEntry = False
    Return
'  ElseIf fpCurrSnCitizen.Value <> 0 Then
'    NoEntry = False
'    Return
'  ElseIf fpCurrOther.Value <> 0 Then
'    NoEntry = False
'    Return
  ElseIf QPTrim$(fptxtDesc(0).Text) <> "" Then
    NoEntry = False
    Return
  ElseIf QPTrim$(fptxtDesc(1).Text) <> "" Then
    NoEntry = False
    Return
  ElseIf QPTrim$(fptxtDesc(2).Text) <> "" Then
    NoEntry = False
    Return
  ElseIf QPTrim$(fptxtDesc(3).Text) <> "" Then
    NoEntry = False
    Return
  ElseIf QPTrim$(fptxtDesc(4).Text) <> "" Then
    NoEntry = False
    Return
  ElseIf QPTrim$(fptxtVin.Text) <> "" Then
    NoEntry = False
    Return
  ElseIf QPTrim$(fptxtMakeMod.Text) <> "" Then
    NoEntry = False
    Return
  ElseIf fpDSWEight.Value <> 0 Then
    NoEntry = False
    Return
  ElseIf fpModYear.Text <> Mid(Date, 7, 4) Then
    NoEntry = False
    Return
  End If
  
  Return
  
HandleChoice:
    Select Case choice
      Case "abandon"
        Close PHandle
        Unload Me
        Exit Function
      Case "dontsave"
      Case "review"
        ThisControl.SetFocus
        Close PHandle
        Check4Changes = True
        Exit Function
      Case Else
    End Select
      
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersProp", "Check4Changes", Erl)
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

Private Sub AssignTemps()
  TempPROPPIN$ = QPTrim$(fptxtPropPin.Text)
  TempPROPDATE% = Date2Num(fptxtDate.Text)
  TempPersVal# = fpCurrPersVal.Value
  TempMHVALUE# = fpCurrMobHome.Value
  TempMCVALUE# = fpCurrMerchCap.Value
  TempCVALUE# = fpCurrFarmEq.Value
  TempMTVALUE# = fpCurrMachTools.Value
'  TempEXMPSENI# = fpCurrSnCitizen.Value
'  TempEXMPOTHR# = fpCurrOther.Value
  TempDISCOV$ = fpcmbDiscoveryYN.Text
  TempLATELIST$ = fpcmbLateListYN.Text
  TempOptSearch = fptxtOptSearch.Text
  TempDESC1$ = fptxtDesc(0).Text
  TempDESC2$ = fptxtDesc(1).Text
  TempDESC3$ = fptxtDesc(2).Text
  TempDesc4$ = fptxtDesc(3).Text
  TempDesc5$ = fptxtDesc(4).Text
  TempTaxBillYear = CInt(fptxtBillYear.Text)
  TempPPTRAYN = fpcmbPPTRAYN.Text
  TempProrateVal = CInt(fpcmbPRVal.Text)
  TempProrate = fpcmbProRateYN.Text
  TempVin$ = QPTrim$(fptxtVin.Text)
  TempMakeMod$ = QPTrim$(fptxtMakeMod.Text)
  TempWeight = CDbl(fpDSWEight.Value)
  If QPTrim$(fpModYear.Text) = "" Then
    TempModYear = 0
  Else
    TempModYear = CInt(fpModYear.Text)
  End If
  fpcmbOptRev1.Col = 1
  TempOptRev1Chrg = CInt(fpcmbOptRev1.ColText)
  fpcmbOptRev2.Col = 1
  TempOptRev2Chrg = CInt(fpcmbOptRev2.ColText)
  fpcmbOptRev3.Col = 1
  TempOptRev3Chrg = CInt(fpcmbOptRev3.ColText)

End Sub

Private Sub LogSaves()
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  
'  On Error GoTo ERRORSTUFF
  
  OpenPersPropFile PHandle, NumOfPersRecs
  Get PHandle, PersRecs(WhichRec), PersRec
  Close PHandle
  
  If QPTrim$(TempPROPPIN$) <> QPTrim$(PersRec.PropPin) Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Property PIN# was changed from " + QPTrim$(TempPROPPIN$) + " to " + QPTrim$(PersRec.PropPin) + " and saved.")
  End If
  
  If TempPROPDATE% <> PersRec.PROPDATE Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the date was changed from " + MakeRegDate(TempPROPDATE) + " to " + MakeRegDate(PersRec.PROPDATE) + " and saved.")
  End If
  
  If TempPersVal# <> PersRec.PersVal Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Personal Value was changed from " + QPTrim$(Using("$###,###,##0.00", TempPersVal)) + " to " + QPTrim$(Using("$###,###,##0.00", PersRec.PersVal)) + " and saved.")
  End If
  
  If TempMHVALUE# <> PersRec.MHValue Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Mobile Home Value was changed from " + QPTrim$(Using("$###,###,##0.00", TempMHVALUE)) + " to " + QPTrim$(Using("$###,###,##0.00", PersRec.MHValue)) + " and saved.")
  End If
  
  If TempMCVALUE# <> PersRec.MCValue Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Merchant Capital Value was changed from " + QPTrim$(Using("$###,###,##0.00", TempMCVALUE)) + " to " + QPTrim$(Using("$###,###,##0.00", PersRec.MCValue)) + " and saved.")
  End If
  
  If TempCVALUE# <> PersRec.CVALUE Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Farm Equipment Value was changed from " + QPTrim$(Using("$###,###,##0.00", TempCVALUE)) + " to " + QPTrim$(Using("$###,###,##0.00", PersRec.CVALUE)) + " and saved.")
  End If
  
  If TempMTVALUE# <> PersRec.MTValue Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Machine Tools Value was changed from " + QPTrim$(Using("$###,###,##0.00", TempMTVALUE)) + " to " + QPTrim$(Using("$###,###,##0.00", PersRec.MTValue)) + " and saved.")
  End If
  
'  If TempEXMPSENI# <> PersRec.EXMPSENI Then
'    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Senior Exemptions Value was changed from " + QPTrim$(Using("$###,###,##0.00", TempEXMPSENI)) + " to " + QPTrim$(Using("$###,###,##0.00", PersRec.EXMPSENI)) + " and saved.")
'  End If
'
'  If TempEXMPOTHR# <> PersRec.EXMPOTHR Then
'    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Other Exemptions Value was changed from " + QPTrim$(Using("$###,###,##0.00", TempEXMPOTHR)) + " to " + QPTrim$(Using("$###,###,##0.00", PersRec.EXMPOTHR)) + " and saved.")
'  End If
  
  If QPTrim$(TempDISCOV$) <> QPTrim$(PersRec.DISCOV) Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Discovery Y/N? was changed from " + QPTrim$(TempDISCOV$) + " to " + QPTrim$(PersRec.DISCOV) + " and saved.")
  End If
  
  If QPTrim$(TempLATELIST$) <> QPTrim$(PersRec.LateList) Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Late List Y/N? was changed from " + QPTrim$(TempLATELIST$) + " to " + QPTrim$(PersRec.LateList) + " and saved.")
  End If
  
  If QPTrim$(TempOptSearch$) <> QPTrim$(PersRec.OptSearch) Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Optional Search was changed from " + QPTrim$(TempOptSearch) + " to " + QPTrim$(PersRec.OptSearch) + " and saved.")
  End If
  
  If QPTrim$(TempDESC1$) <> QPTrim$(PersRec.DESC1) Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Notes Line #1 was changed from " + QPTrim$(TempDESC1$) + " to " + QPTrim$(PersRec.DESC1) + " and saved.")
  End If
  
  If QPTrim$(TempDESC2$) <> QPTrim$(PersRec.DESC2) Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Notes Line #2 was changed from " + QPTrim$(TempDESC2$) + " to " + QPTrim$(PersRec.DESC2) + " and saved.")
  End If
  
  If QPTrim$(TempDESC3$) <> QPTrim$(PersRec.DESC3) Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Notes Line #3 was changed from " + QPTrim$(TempDESC3$) + " to " + QPTrim$(PersRec.DESC3) + " and saved.")
  End If
  
  If QPTrim$(TempDesc4$) <> QPTrim$(PersRec.Desc4) Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Notes Line #4 was changed from " + QPTrim$(TempDesc4$) + " to " + QPTrim$(PersRec.Desc4) + " and saved.")
  End If
  
  If QPTrim$(TempDesc5$) <> QPTrim$(PersRec.Desc5) Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Notes Line #5 was changed from " + QPTrim$(TempDesc5$) + " to " + QPTrim$(PersRec.Desc5) + " and saved.")
  End If
  
'  If TempTaxBillYear <> PersRec.TaxBillYear Then
'    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Tax Bill Year was changed from " + CStr(TempTaxBillYear) + " to " + CStr(PersRec.TaxBillYear) + " and saved.")
'  End If
  
  If QPTrim$(TempPPTRAYN) <> QPTrim$(PersRec.PPTRAYN) Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the PPTRAYN field was changed from " + QPTrim$(TempPPTRAYN) + " to " + QPTrim$(PersRec.PPTRAYN) + " and saved.")
  End If
  
  If TempProrateVal <> PersRec.ProrateVal Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Prorate Value was changed from " + Using$("###0", TempProrateVal) + " to " + Using$("###0", PersRec.ProrateVal) + " and saved.")
  End If
  
  If QPTrim$(TempProrate) <> QPTrim$(PersRec.Prorate) Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Prorate Y/N? field was changed from " + QPTrim$(TempProrate) + " to " + QPTrim$(PersRec.Prorate) + " and saved.")
  End If
  
  If TempOptRev1Chrg% <> PersRec.OptRev1Chrg Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Opt'l Rev 1 Y/N? was changed from " + QPTrim$(RateDesc(TempOptRev1Chrg)) + " to " + QPTrim$(RateDesc(Opt1)) + " and saved.")
  End If
  
  If TempOptRev2Chrg% <> PersRec.OptRev2Chrg Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Opt'l Rev 2 Y/N? was changed from " + QPTrim$(RateDesc(TempOptRev2Chrg)) + " to " + QPTrim$(RateDesc(Opt2)) + " and saved.")
  End If
  
  If TempOptRev3Chrg% <> PersRec.OptRev3Chrg Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Opt'l Rev 3 Y/N? was changed from " + QPTrim$(RateDesc(TempOptRev3Chrg)) + " to " + QPTrim$(RateDesc(Opt3)) + " and saved.")
  End If
  
  If QPTrim$(TempVin) <> QPTrim$(PersRec.Vin) Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Vin # field was changed from " + QPTrim$(TempVin) + " to " + QPTrim$(PersRec.Vin) + " and saved.")
  End If
  
  If QPTrim$(TempMakeMod) <> QPTrim$(PersRec.MakeMod) Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Make/Model field was changed from " + QPTrim$(TempMakeMod) + " to " + QPTrim$(PersRec.MakeMod) + " and saved.")
  End If
  
  If TempWeight <> PersRec.Weight Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Weight field was changed from " + Using$("###,##0.00", TempWeight) + " to " + Using$("###,##0.00", PersRec.Weight) + " and saved.")
  End If
  
  If TempModYear <> PersRec.ModYear Then
    MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": the Model Year field was changed from " + Using$("###0", TempModYear) + " to " + Using$("###0", PersRec.ModYear) + " and saved.")
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersProp", "LogSaves", Erl)
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
  FileName = "C:\CPWork\pdetail1.dat"
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
  
  FileName = "C:\CPWork\pdetail2.dat"
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
  
  FileName = "C:\CPWork\pdetail3.dat"
  ThisFile = FreeFile
  Open FileName For Output As ThisFile
  One = 1
  Print #ThisFile, One
  Close ThisFile
  
  frmVATaxRateDetail.Show vbModal

End Sub

Public Sub LoadGoToPinsCmb()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NextRec As Long
  Dim NumOfTCRecs As Long
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPRecs As Long
  Dim ThisPin As String
  Dim ThatPin As String
  Dim RecCnt As Integer
  
  ThisPin = QPTrim$(fptxtPropPin.Text)
  OpenTaxCustFile TCHandle, NumOfTCRecs
  Get TCHandle, GCustNum, TaxCust
  Close TCHandle
  
  NextRec = TaxCust.FirstPersRec
  If NextRec = 0 Then
    fpcmbGoToPins.Action = ActionClear
    Exit Sub
  End If
  
  fpcmbGoToPins.Action = ActionClear
  
  OpenPersPropFile PHandle, NumOfPRecs
  Do While NextRec > 0
    Get PHandle, NextRec, PersRec
    If PersRec.Deleted <> 0 Then GoTo SkipIt
    RecCnt = RecCnt + 1
    ThatPin = QPTrim$(PersRec.PropPin)
    If ThatPin <> "" And ThatPin = ThisPin Then
      fpcmbGoToPins.Text = ThatPin & Chr(9) & CStr(RecCnt)
    End If
    fpcmbGoToPins.AddItem ThatPin & Chr(9) & CStr(RecCnt)
SkipIt:
    NextRec = PersRec.NextRec
  Loop
     
  Close PHandle
  
End Sub

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

Private Function Check4DupPinsPers(PinNum$, RecNum As Long) As Boolean
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim x As Long
  
  On Error GoTo ERRORSTUFF
  
  Check4DupPinsPers = False
  OpenPersPropFile PHandle, NumOfPersRecs
  For x = 1 To NumOfPersRecs
    Get PHandle, x, PersRec
    If x <> RecNum Then
      If PersRec.Deleted = -1 Then GoTo Deleted
      If PersRec.CustPin = 0 Then GoTo Deleted
      If QPTrim$(PersRec.PropPin) = PinNum Then
        If TaxMsgWOpts(800, "The pin number entered is already in use and should not be used. Do you wish to edit the pin number?", "F10 Continue", "ESC Abort") = "abort" Then
          Check4DupPinsPers = True
          fptxtPropPin.SetFocus
        Else
          MainLog ("For " + QPTrim$(CustName$) + " Personal Property #" + CStr(WhichRec) + ": The user elected to keep the duplicated pin number " + fptxtPropPin.Text + " after being warned of the consequences.")
        End If
        Exit For
      End If
    End If
Deleted:
  Next x
  
  Close PHandle
  Exit Function

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersProp", "Check4DupPinsPers", Erl)
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


