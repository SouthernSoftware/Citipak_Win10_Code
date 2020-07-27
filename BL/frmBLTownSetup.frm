VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLTownSetup 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Town Setup "
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   ForeColor       =   &H00000000&
   Icon            =   "frmBLTownSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbLaserYN 
      Height          =   360
      Left            =   6132
      TabIndex        =   16
      Tag             =   $"frmBLTownSetup.frx":08CA
      Top             =   5988
      Width           =   780
      _Version        =   196608
      _ExtentX        =   1376
      _ExtentY        =   635
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
      ColDesigner     =   "frmBLTownSetup.frx":0A32
   End
   Begin LpLib.fpCombo fpcmbAmtPct 
      Height          =   360
      Left            =   8016
      TabIndex        =   15
      Tag             =   $"frmBLTownSetup.frx":0D29
      Top             =   5472
      Width           =   780
      _Version        =   196608
      _ExtentX        =   1376
      _ExtentY        =   635
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
      ColDesigner     =   "frmBLTownSetup.frx":0EC5
   End
   Begin LpLib.fpCombo fpcmbNewLicNumYN 
      Height          =   360
      Left            =   8016
      TabIndex        =   14
      Tag             =   $"frmBLTownSetup.frx":11BC
      Top             =   4992
      Width           =   780
      _Version        =   196608
      _ExtentX        =   1376
      _ExtentY        =   635
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
      ColDesigner     =   "frmBLTownSetup.frx":1296
   End
   Begin LpLib.fpCombo fpcmbDLQNotice 
      Height          =   360
      Left            =   4356
      TabIndex        =   18
      Tag             =   $"frmBLTownSetup.frx":158D
      Top             =   6852
      Width           =   2556
      _Version        =   196608
      _ExtentX        =   4508
      _ExtentY        =   635
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
      BorderGrayAreaColor=   14737632
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
      ColDesigner     =   "frmBLTownSetup.frx":1791
   End
   Begin LpLib.fpCombo fpcmbAppType 
      Height          =   360
      Left            =   4356
      TabIndex        =   17
      Tag             =   $"frmBLTownSetup.frx":1A88
      Top             =   6420
      Width           =   2556
      _Version        =   196608
      _ExtentX        =   4508
      _ExtentY        =   635
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
      ColDesigner     =   "frmBLTownSetup.frx":1CAE
   End
   Begin LpLib.fpCombo fpcmbAcctMethod 
      Height          =   360
      Left            =   8892
      TabIndex        =   8
      Tag             =   $"frmBLTownSetup.frx":1FA5
      Top             =   1728
      Width           =   1788
      _Version        =   196608
      _ExtentX        =   3154
      _ExtentY        =   635
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
      ColDesigner     =   "frmBLTownSetup.frx":21D7
   End
   Begin LpLib.fpCombo fpcmbGLToCats 
      Height          =   360
      Left            =   9060
      TabIndex        =   13
      Tag             =   $"frmBLTownSetup.frx":24CE
      Top             =   4512
      Width           =   780
      _Version        =   196608
      _ExtentX        =   1376
      _ExtentY        =   635
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
      ColDesigner     =   "frmBLTownSetup.frx":25CD
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   630
      Left            =   4650
      TabIndex        =   21
      TabStop         =   0   'False
      Tag             =   "Press this button to exit back to the main business license menu."
      Top             =   7680
      Width           =   2385
      _Version        =   131072
      _ExtentX        =   4207
      _ExtentY        =   1111
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
      ButtonDesigner  =   "frmBLTownSetup.frx":28C4
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   636
      Left            =   7548
      TabIndex        =   22
      TabStop         =   0   'False
      Tag             =   $"frmBLTownSetup.frx":2AA3
      Top             =   7680
      Width           =   2328
      _Version        =   131072
      _ExtentX        =   4106
      _ExtentY        =   1122
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
      ButtonDesigner  =   "frmBLTownSetup.frx":2C2D
   End
   Begin fpBtnAtlLibCtl.fpBtn fpcmdApps 
      Height          =   360
      Left            =   6960
      TabIndex        =   20
      TabStop         =   0   'False
      Tag             =   $"frmBLTownSetup.frx":2E09
      Top             =   6420
      Width           =   2400
      _Version        =   131072
      _ExtentX        =   4233
      _ExtentY        =   635
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
      ButtonDesigner  =   "frmBLTownSetup.frx":2EE9
   End
   Begin fpBtnAtlLibCtl.fpBtn fpcmdDLQ 
      Height          =   360
      Left            =   6960
      TabIndex        =   19
      TabStop         =   0   'False
      Tag             =   $"frmBLTownSetup.frx":30CA
      ToolTipText     =   "Press to open the delinquent notice number indicated in the drop down box on the left."
      Top             =   6876
      Width           =   2400
      _Version        =   131072
      _ExtentX        =   4233
      _ExtentY        =   635
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
      ButtonDesigner  =   "frmBLTownSetup.frx":31AF
   End
   Begin EditLib.fpText fptxtRevGLAcctNum 
      Height          =   390
      Left            =   8070
      TabIndex        =   9
      Tag             =   $"frmBLTownSetup.frx":3397
      Top             =   2835
      Width           =   2070
      _Version        =   196608
      _ExtentX        =   3662
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
      MaxLength       =   14
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
   Begin EditLib.fpText fptxtCashReceipt 
      Height          =   390
      Left            =   8070
      TabIndex        =   11
      Tag             =   $"frmBLTownSetup.frx":34BB
      Top             =   3795
      Width           =   2070
      _Version        =   196608
      _ExtentX        =   3662
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
      MaxLength       =   14
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
   Begin EditLib.fpText fptxtAcctsRec 
      Height          =   390
      Left            =   8070
      TabIndex        =   10
      Tag             =   $"frmBLTownSetup.frx":35D9
      Top             =   3315
      Width           =   2070
      _Version        =   196608
      _ExtentX        =   3662
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
      MaxLength       =   150
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
   Begin EditLib.fpMask fptxtPhone 
      Height          =   390
      Left            =   4770
      TabIndex        =   7
      Tag             =   "Enter the town's office phone number in this field."
      Top             =   3840
      Width           =   1650
      _Version        =   196608
      _ExtentX        =   2900
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
      Mask            =   "(###)-###-####"
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      RequireFill     =   0   'False
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483637
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      AutoTab         =   0   'False
      ButtonColor     =   -2147483637
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtState 
      Height          =   396
      Left            =   2076
      TabIndex        =   5
      Tag             =   "Enter the state the town is in. Use the generally accepted two character upper case abbreviation (North Carolina = NC)."
      Top             =   3864
      Width           =   588
      _Version        =   196608
      _ExtentX        =   1037
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
      MaxLength       =   2
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
   Begin EditLib.fpMask fptxtZip 
      Height          =   396
      Left            =   2700
      TabIndex        =   6
      Tag             =   "Enter either a five digit or nine digit zip code for the town in this field."
      Top             =   3864
      Width           =   1308
      _Version        =   196608
      _ExtentX        =   2307
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
   Begin fpBtnAtlLibCtl.fpBtn cmdGLList 
      Height          =   1350
      Left            =   10200
      TabIndex        =   40
      TabStop         =   0   'False
      Tag             =   $"frmBLTownSetup.frx":36FF
      Top             =   2835
      Width           =   660
      _Version        =   131072
      _ExtentX        =   1164
      _ExtentY        =   2381
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
      ButtonDesigner  =   "frmBLTownSetup.frx":37FE
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdAdvanceLtr 
      Height          =   360
      Left            =   6960
      TabIndex        =   41
      TabStop         =   0   'False
      Tag             =   $"frmBLTownSetup.frx":39DD
      Top             =   5976
      Width           =   2400
      _Version        =   131072
      _ExtentX        =   4233
      _ExtentY        =   635
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
      ButtonDesigner  =   "frmBLTownSetup.frx":3AAC
   End
   Begin EditLib.fpCurrency fpcurrIssFee 
      Height          =   345
      Left            =   3210
      TabIndex        =   12
      Tag             =   $"frmBLTownSetup.frx":3C90
      Top             =   4515
      Width           =   1410
      _Version        =   196608
      _ExtentX        =   2487
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
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   630
      Left            =   1770
      TabIndex        =   45
      TabStop         =   0   'False
      Tag             =   "Place the cursor over any field and a pop-up balloon will appear containing help information about that field."
      Top             =   7680
      Width           =   2340
      _Version        =   131072
      _ExtentX        =   4128
      _ExtentY        =   1111
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
      ButtonDesigner  =   "frmBLTownSetup.frx":3E7A
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   540
      Left            =   624
      TabIndex        =   46
      Top             =   7680
      Width           =   876
      _Version        =   131072
      _ExtentX        =   1545
      _ExtentY        =   952
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
      ShapeRoundWidth =   180
      ShapeRoundHeight=   180
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
      MaxWidth        =   3160
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
   Begin EditLib.fpText fptxtTownName 
      Height          =   396
      Left            =   2064
      TabIndex        =   0
      Tag             =   "Enter the official name of your town here. For example, 'Town Of Washington'."
      Top             =   1632
      Width           =   4332
      _Version        =   196608
      _ExtentX        =   7641
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
      MaxLength       =   38
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
   Begin EditLib.fpText fptxtContact 
      Height          =   396
      Left            =   2076
      TabIndex        =   1
      Tag             =   "Enter the name of someone in your office that is knowledgeable regarding business licenses."
      Top             =   2076
      Width           =   4332
      _Version        =   196608
      _ExtentX        =   7641
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
   Begin EditLib.fpText fptxtAdd1 
      Height          =   396
      Left            =   2076
      TabIndex        =   2
      Tag             =   "This field requires the primary mailing address. A street address is usually the best entry for this field."
      Top             =   2520
      Width           =   4332
      _Version        =   196608
      _ExtentX        =   7641
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
   Begin EditLib.fpText fptxtAdd2 
      Height          =   396
      Left            =   2076
      TabIndex        =   3
      Tag             =   "Enter the secondary address in this field. Post office box numbers or Suite numbers, etc. usually go in this field."
      ToolTipText     =   "Enter a secondary mailing address here."
      Top             =   2964
      Width           =   4332
      _Version        =   196608
      _ExtentX        =   7641
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
   Begin EditLib.fpText fptxtCity 
      Height          =   396
      Left            =   2076
      TabIndex        =   4
      Tag             =   "Enter the name of the town as it would appear in a return address."
      Top             =   3408
      Width           =   4332
      _Version        =   196608
      _ExtentX        =   7641
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
      MaxLength       =   38
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
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Apply Penalty GL Nums To Categories?:"
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
      Left            =   5220
      TabIndex        =   48
      Top             =   4605
      Width           =   3735
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   5970
      Left            =   600
      Top             =   1455
      Width           =   10470
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
      Left            =   1920
      TabIndex        =   47
      Top             =   8400
      Width           =   2076
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   588
      X2              =   11032
      Y1              =   4410
      Y2              =   4410
   End
   Begin VB.Label Label18 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   576
      TabIndex        =   44
      Top             =   1104
      Width           =   1884
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Issuance Fee:"
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
      Left            =   1815
      TabIndex        =   43
      Top             =   4605
      Width           =   1335
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Create (Laser Only) Advance Letter ?"
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
      Left            =   2424
      TabIndex        =   42
      Top             =   6048
      Width           =   3540
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Accounting Method"
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
      Left            =   6870
      TabIndex        =   39
      Top             =   1770
      Width           =   1890
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   6540
      X2              =   11032
      Y1              =   2355
      Y2              =   2355
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Penalty:  G/L Account Numbers"
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
      Left            =   6990
      TabIndex        =   38
      Top             =   2520
      Width           =   3210
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   6540
      X2              =   6540
      Y1              =   1460
      Y2              =   4416
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Accts Rec:"
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
      Left            =   6975
      TabIndex        =   37
      Top             =   3405
      Width           =   1020
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Receipt:"
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
      Left            =   6630
      TabIndex        =   36
      Top             =   3885
      Width           =   1350
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Revenue:"
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
      Left            =   7110
      TabIndex        =   35
      Top             =   2925
      Width           =   930
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Do You Prefer Amounts Or Percentages For Penalties?*"
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
      Left            =   2640
      TabIndex        =   34
      Top             =   5565
      Width           =   5340
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Do You Wish To Assign Permanent License Numbers?*"
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
      Left            =   2640
      TabIndex        =   33
      Top             =   5115
      Width           =   5295
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Delinquent Notice Type:"
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
      Left            =   1908
      TabIndex        =   32
      Top             =   6948
      Width           =   2364
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Application Type:"
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
      Left            =   2532
      TabIndex        =   31
      Top             =   6516
      Width           =   1932
   End
   Begin VB.Label Label75 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone*"
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
      Left            =   4080
      TabIndex        =   30
      Top             =   4032
      Width           =   684
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "State/Zip*"
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
      Height          =   276
      Left            =   924
      TabIndex        =   29
      Top             =   4000
      Width           =   1020
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "City*"
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
      Height          =   276
      Left            =   1464
      TabIndex        =   28
      Top             =   3534
      Width           =   492
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Height          =   276
      Left            =   1128
      TabIndex        =   27
      Top             =   3112
      Width           =   828
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address*"
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
      Height          =   264
      Left            =   1032
      TabIndex        =   26
      Top             =   2652
      Width           =   924
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Attention"
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
      Left            =   1032
      TabIndex        =   25
      Top             =   2220
      Width           =   924
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Town Name*"
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
      Left            =   696
      TabIndex        =   24
      Top             =   1760
      Width           =   1260
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Town Setup Information"
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
      TabIndex        =   23
      Top             =   528
      Width           =   5292
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   660
      Index           =   1
      Left            =   1500
      Top             =   384
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   1500
      Top             =   264
      Width           =   8652
   End
End
Attribute VB_Name = "frmBLTownSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Dim TempTownName$
  Dim TempTownContact$
  Dim TempTownAdd1$
  Dim TempTownAdd2$
  Dim TempTownCity$
  Dim TempTownState$
  Dim TempTownZip$
  Dim TempTownPhone$
  Dim TempTownAppForm As Integer
  Dim TempTownPenForm As Integer
  Dim TempLicNew$
  Dim TempAmtPct$
  Dim TempAcctMethod$
  Dim TempPenRevNum$
  Dim LastPenRevNum$
  Dim TempPenRecNum$
  Dim LastPenRecNum$
  Dim TempCashRecpt$
  Dim LastCashRecpt$
  Dim TempIssFee As Double
  Dim TempLaserYN$
  Dim TempGLToCats$
  Dim LastAcctMethod$
  Dim JustOpened As Boolean
  Dim AcctMethodChangeFlag As Boolean
  
  Private Temp_Class As Resize_Class

Private Sub cmdAdvanceLtr_Click()

'  If Not Exist("artownsu.dat") Then
'    frmBLMessageBoxJr.Label1.Caption = "Since no town setup data has been saved yet please complete the accounting method data and all the other required fields and save them now before continuing."
'    frmBLMessageBoxJr.Label1.Top = 600
'    frmBLMessageBoxJr.Show vbModal
'    fpcmbAcctMethod.SetFocus
'    Exit Sub
'  End If
  
  If QPTrim$(fpcmbLaserYN.Text) = "1" Then
    frmBLAdvanceLetter.Show
  ElseIf QPTrim$(fpcmbLaserYN.Text) = "2" Then
    frmBLAdvLetter2.Show
  ElseIf QPTrim$(fpcmbLaserYN.Text) = "3" Then
    frmBLAdvanceLtr3.Show
  ElseIf QPTrim$(fpcmbLaserYN.Text) = "4" Then
    frmBLFreeFormatDlnq.Show
  Else
    frmBLMessageBoxJr.Label1.Caption = "Select either a 1, 2 or 3 from the drop down box to bring up either of those forms."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
  End If
End Sub

Private Sub cmdExit_Click()
  Dim TownRec As TownSetUpType
  Dim THandle As Integer
  Dim ChangeFlag As Boolean
  Dim DoWhatFlag As SaveChangeOptions1
  
  On Error GoTo ERRORSTUFF
  
  ChangeFlag = False
  If Exist("artownsu.dat") Then
    OpenTownFile THandle
    Get THandle, 1, TownRec
    Close THandle
  Else
    GoTo BeatIt
  End If
  
  'it might be common for a user to save an application or delinquent form but, since the data for the town
  'is still on the screen and since the user might think since when he saved the
  'form he also saved the town data, it makes it necessary to remind the user
  'to save the town data as a separate process
  If QPTrim$(TownRec.AppTownOf) <> "" And QPTrim$(TownRec.TownName) = "" Then
    frmBLMessageBoxJrWOpts.Label1.Caption = "You have saved a renewal application but no Town data has been saved from this screen. Do you want to save town data before exiting?"
    frmBLMessageBoxJrWOpts.Label1.Top = 700
    frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Save"
    frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Don't Save"
    frmBLMessageBoxJrWOpts.Show vbModal
    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
      Unload frmBLMessageBoxJrWOpts
      Call cmdSave_Click
      Exit Sub
    Else
      Unload frmBLMessageBoxJrWOpts
    End If
  End If
  
  If QPTrim$(TownRec.DlqTownName) <> "" And QPTrim$(TownRec.TownName) = "" Then
    frmBLMessageBoxJrWOpts.Label1.Caption = "You have saved a delinquent notice form but no Town data has been saved from this screen. Do you want to save town data before exiting?"
    frmBLMessageBoxJrWOpts.Label1.Top = 700
    frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Save"
    frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Don't Save"
    frmBLMessageBoxJrWOpts.Show vbModal
    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
      Unload frmBLMessageBoxJrWOpts
      Call cmdSave_Click
      Exit Sub
    Else
      Unload frmBLMessageBoxJrWOpts
    End If
  End If
  
  If Exist("arlaser4.dat") And QPTrim$(TownRec.TownName) = "" Then
    frmBLMessageBoxJrWOpts.Label1.Caption = "You have saved delinquent notice form #4 but no Town data has been saved from this screen. Do you want to save town data before exiting?"
    frmBLMessageBoxJrWOpts.Label1.Top = 700
    frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Save"
    frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Don't Save"
    frmBLMessageBoxJrWOpts.Show vbModal
    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
      Unload frmBLMessageBoxJrWOpts
      Call cmdSave_Click
      Exit Sub
    Else
      Unload frmBLMessageBoxJrWOpts
    End If
  End If
  
  If QPTrim$(fptxtTownName.Text) <> QPTrim$(TownRec.TownName) Then
    ChangeFlag = True
    fptxtTownName.BackColor = &H80FFFF
    fptxtTownName.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(fptxtContact.Text) <> QPTrim$(TownRec.Contact) Then
    ChangeFlag = True
    fptxtContact.BackColor = &H80FFFF
    fptxtContact.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(fptxtAdd1.Text) <> QPTrim$(TownRec.TownAdd1) Then
    ChangeFlag = True
    fptxtAdd1.BackColor = &H80FFFF
    fptxtAdd1.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(fptxtAdd2.Text) <> QPTrim$(TownRec.TownAdd2) Then
    ChangeFlag = True
    fptxtAdd2.BackColor = &H80FFFF
    fptxtAdd2.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(fptxtCity.Text) <> QPTrim$(TownRec.City) Then
    ChangeFlag = True
    fptxtCity.BackColor = &H80FFFF
    fptxtCity.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(fptxtState.Text) <> QPTrim$(TownRec.State) Then
    ChangeFlag = True
    fptxtState.BackColor = &H80FFFF
    fptxtState.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(fptxtZip.Text) <> QPTrim$(TownRec.ZipCode) Then
    ChangeFlag = True
    fptxtZip.BackColor = &H80FFFF
    fptxtZip.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(fptxtPhone.Text) <> QPTrim$(TownRec.TownPhone) Then
    ChangeFlag = True
    fptxtPhone.BackColor = &H80FFFF
    fptxtPhone.SetFocus
    GoTo ChangeFound
  End If
  
  If Exist("artownsu.dat") Then GoTo NoCheckNeeded
  
  If Mid(fpcmbAcctMethod.Text, 1, 1) <> QPTrim$(TownRec.AcctMeth) Then
    ChangeFlag = True
    fpcmbAcctMethod.BackColor = &H80FFFF
    fpcmbAcctMethod.SetFocus
    GoTo ChangeFound
  End If

NoCheckNeeded:
  If AcctMethodChangeFlag = False Then
    If GetGLNum(TownRec.PENREVGLNUM) <> "" Then
      If QPTrim(fptxtRevGLAcctNum.Text) <> GetGLNum(TownRec.PENREVGLNUM) Then
        ChangeFlag = True
        fptxtRevGLAcctNum.BackColor = &H80FFFF
        fptxtRevGLAcctNum.SetFocus
        GoTo ChangeFound
      End If
    End If
    If GetGLNum(TownRec.PENRECGLNUM) <> "" Then
      If QPTrim$(fptxtAcctsRec.Text) <> GetGLNum(TownRec.PENRECGLNUM) Then
        ChangeFlag = True
        fptxtAcctsRec.BackColor = &H80FFFF
        fptxtAcctsRec.SetFocus
        GoTo ChangeFound
      End If
    End If
    If GetGLNum(TownRec.PENCASHACCT) <> "" Then
      If QPTrim$(fptxtCashReceipt.Text) <> GetGLNum(TownRec.PENCASHACCT) Then
        ChangeFlag = True
        fptxtCashReceipt.BackColor = &H80FFFF
        fptxtCashReceipt.SetFocus
        GoTo ChangeFound
      End If
    End If
  End If
  
  If QPTrim$(fpcmbNewLicNumYN.Text) <> QPTrim$(TownRec.LicNumPermYN) Then
    ChangeFlag = True
    fpcmbNewLicNumYN.BackColor = &H80FFFF
    fpcmbNewLicNumYN.SetFocus
    GoTo ChangeFound
  End If
  
  If QPTrim$(fpcmbAmtPct.Text) <> QPTrim$(TownRec.UseAmtPctYN) Then
    ChangeFlag = True
    fpcmbAmtPct.BackColor = &H80FFFF
    fpcmbAmtPct.SetFocus
    GoTo ChangeFound
  End If
  
  If fpcurrIssFee <> TownRec.IssFee Then
    ChangeFlag = True
    fpcurrIssFee.BackColor = &H80FFFF
    fpcurrIssFee.SetFocus
    GoTo ChangeFound
  End If
  
  If Mid(fpcmbLaserYN.Text, 1, 1) <> QPTrim$(TownRec.LaserLtr) Then
    ChangeFlag = True
    fpcmbLaserYN.BackColor = &H80FFFF
    fpcmbLaserYN.SetFocus
    GoTo ChangeFound
  End If
  
  If CInt(Mid(fpcmbDLQNotice.Text, 1, 1)) <> TownRec.DLQNotice Then
    ChangeFlag = True
    fpcmbDLQNotice.BackColor = &H80FFFF
    fpcmbDLQNotice.SetFocus
    GoTo ChangeFound
  End If
  
ChangeFound:
  If ChangeFlag = True Then
    ChangeFlag = False
    DoWhatFlag = PromptSaveChanges(Me)
    Select Case DoWhatFlag
    Case SaveChangeOptions1.scoSaveChanges
      Call cmdSave_Click
      Call MakeBackColorsWhite
      Exit Sub 'don't exit screen
    Case SaveChangeOptions1.scoReviewChanges 'review is just bringing back the current form
      Exit Sub
    Case SaveChangeOptions1.scoAbandonChanges 'abandon
      GoTo BeatIt
      Exit Sub
    Case Else:
    End Select
  End If
  
BeatIt:
  frmBLMainMenu.Show
  
  KillFile "townsetup.dat"
  DoEvents
  MainLog ("Town setup screen exited.")
  Unload frmBLTownSetup
  
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLTownSetup", "cmdExit_Click", Erl)
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

Private Sub cmdGLList_Click()
  frmBLGLList.Show vbModal
End Sub

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "F1 &Turn Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    lblBalloon.Visible = True
    cmdHelp.ToolTipText = ""
    fptxtTownName.ToolTipText = ""
    fptxtContact.ToolTipText = ""
    fptxtAdd1.ToolTipText = ""
    fptxtAdd2.ToolTipText = ""
    fptxtCity.ToolTipText = ""
    fptxtState.ToolTipText = ""
    fptxtZip.ToolTipText = ""
    fptxtPhone.ToolTipText = ""
    fpcmbAcctMethod.ToolTipText = ""
    fptxtRevGLAcctNum.ToolTipText = ""
    fptxtAcctsRec.ToolTipText = ""
    fptxtCashReceipt.ToolTipText = ""
    cmdGLList.ToolTipText = ""
    fpcmbNewLicNumYN.ToolTipText = ""
    fpcmbGLToCats.ToolTipText = ""
    fpcmbAmtPct.ToolTipText = ""
    fpcurrIssFee.ToolTipText = ""
    fpcmbLaserYN.ToolTipText = ""
    cmdAdvanceLtr.ToolTipText = ""
    fpcmbAppType.ToolTipText = ""
    fpcmdApps.ToolTipText = ""
    fpcmbDLQNotice.ToolTipText = ""
    fpcmdDLQ.ToolTipText = ""
    cmdExit.ToolTipText = ""
    cmdSave.ToolTipText = ""
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
'    cmdHelp.ToolTipText = "Click on this button to activate informational balloons for each field."
'    fptxtTownName.ToolTipText = "Enter the town's offical name here. (Ex. Town Of Riverville)"
'    fptxtContact.ToolTipText = "Enter the name of a town office contact."
'    fptxtAdd1.ToolTipText = "Enter the town's street mailing address."
'    fptxtAdd2.ToolTipText = "Enter a secondary mailing address here."
'    fptxtCity.ToolTipText = "Enter the town's mailing name here."
'    fptxtState.ToolTipText = "Enter the town's state here."
'    fptxtZip.ToolTipText = "Enter the town's zip code here."
'    fptxtPhone.ToolTipText = "Enter the town office's phone number."
'    fpcmbAcctMethod.ToolTipText = "Choose one of the three accounting methods supported in business license."
'    fptxtRevGLAcctNum.ToolTipText = "Enter the general ledger revenue number here (for cash and accrual methods.)"
'    fptxtAcctsRec.ToolTipText = "Enter the general ledger accounts receivable number if required (accrual method)."
'    fptxtCashReceipt.ToolTipText = "Enter the cash receipts general ledger number if required (cash and accrual methods)"
'    cmdGLList.ToolTipText = "Press to bring up an interactive list of all general ledger numbers."
'    fpcmbNewLicNumYN.ToolTipText = "Enter Yes here if you wish to have the business license program automatically assign permanent, uneditable license numbers for each new business added."
'    fpcmbGLToCats.ToolTipText = "Select 'Yes' here if you want to display the penalty GL numbers as default when adding a new category."
'    fpcmbAmtPct.ToolTipText = "Enter the penalty assessment method preferred."
'    fpcurrIssFee.ToolTipText = "If you enter a value in this field then the program will include this fee where appropriate."
'    fpcmbLaserYN.ToolTipText = "The advance letter/laser form is a special application form that you design."
'    cmdAdvanceLtr.ToolTipText = "Press to jump to the Advance Letter editing screen."
'    fpcmbAppType.ToolTipText = "You can save an appplication form here or you can go to a specific form and save it there."
'    fpcmdApps.ToolTipText = "Press to bring up the application form indicated in the drop down box at the left."
'    fpcmbDLQNotice.ToolTipText = "You can save a delinquent notice here or you can go to a specific form and save it there."
'    fpcmdDLQ.ToolTipText = "Press to open the delinquent notice number indicated in the drop down box on the left."
'    cmdExit.ToolTipText = "Press to return to the main menu."
'    cmdSave.ToolTipText = "Press to commit data on this screen to memory."
  End If
End Sub

Private Sub cmdSave_Click()
  Dim TownRec As TownSetUpType
  Dim THandle As Integer
  Dim RecNum As Integer
  Dim x As Integer
  
  On Error GoTo ERRORSTUFF
  
  If Exist("artmppst.dat") Then
    If fpcurrIssFee <> TempIssFee Then
      frmBLMessageBoxJrWOpts.Label1.Caption = "The issuance fee has been changed. This change would cause the unposted business license fee file currently saved to be inaccurate. Continuing with this save will delete the unposted business license file. This means the business license register will have to be run again so the new issuance fee will be included. Do you want to continue anyway?"
      frmBLMessageBoxJrWOpts.Label1.Top = 400
      frmBLMessageBoxJrWOpts.Label1.Height = 1600
      frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
      frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Abort"
      frmBLMessageBoxJrWOpts.Show vbModal
      If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
        Close
        Unload frmBLMessageBoxJrWOpts
        fpcurrIssFee.SetFocus
        Exit Sub
      Else
        Unload frmBLMessageBoxJrWOpts
        KillFile "artmppst.dat"
        frmBLMessageBoxJr.Label1.Caption = "The unposted business license file has been deleted."
        frmBLMessageBoxJr.Label1.Top = 800
        frmBLMessageBoxJr.Show vbModal
        MainLog ("User has changed the issuance fee amount while an unposted business license file exists. The user was warned that continuing would delete the 'artmppst.dat' file. The user elected to continue anyway. The file 'artmppst.dat' was deleted.")
      End If
    End If
  End If
    
  'in case any of the backcolors are shaded this sub
  'returns all field back colors to white
  Call MakeBackColorsWhite
  
  'GLNumsOk is a backup procedure (the program goes to great lengths to make
  'sure the GL numbers are entered correctly) just to make sure the GL numbers
  'are filled out properly
  If GLNumsOK = False Then
    Exit Sub
  End If
  
  If GLNumsValid = False Then
    Exit Sub
  End If
  
  If QPTrim$(fptxtTownName.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please add an official town name before saving."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    fptxtTownName.BackColor = &H80FFFF
    fptxtTownName.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtAdd1.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please add a town mailing address before saving."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    fptxtAdd1.BackColor = &H80FFFF
    fptxtAdd1.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtCity.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please add a town name for the address before saving."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    fptxtCity.BackColor = &H80FFFF
    fptxtCity.SetFocus
    Exit Sub
  End If
  
  'the phone mask looks like "(" when it is devoid of a value
  If QPTrim$(fptxtPhone.Text) = "(" Then
    frmBLMessageBoxJr.Label1.Caption = "Please add a town phone number before saving."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    fptxtPhone.BackColor = &H80FFFF
    fptxtPhone.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtState.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please add the town's state before saving."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    fptxtState.BackColor = &H80FFFF
    fptxtState.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtZip.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please add the town's zip code before saving."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    fptxtZip.BackColor = &H80FFFF
    fptxtZip.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fpcmbNewLicNumYN.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please save a Yes or No for whether or not you wish to issue permanent license numbers."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    fpcmbNewLicNumYN.BackColor = &H80FFFF
    fpcmbNewLicNumYN.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fpcmbGLToCats.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please save a Yes or No for whether or not you wish to apply the penalty GL numbers when adding a new category."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    fpcmbGLToCats.BackColor = &H80FFFF
    fpcmbGLToCats.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fpcmbAmtPct.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please save Amt or Pct for penalty charges."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    fpcmbAmtPct.BackColor = &H80FFFF
    fpcmbAmtPct.SetFocus
    Exit Sub
  End If
  
  If Not Exist("artownsu.dat") And fpcmbAcctMethod.Text <> "None" Then
    frmBLMessageBoxJrWOpts.Label1.Caption = "THE ACCOUNTING METHOD FOR BUSINESS LICENSE CANNOT BE CHANGED ONCE IT IS SAVED. PLEASE MAKE SURE THE ACCOUNTING METHOD ABOUT TO BE SAVED NOW IS CORRECT. IF NECESSARY, PRESS ESC TO ABORT THIS SAVE PROCEDURE."
    frmBLMessageBoxJrWOpts.Label1.Top = 600
    frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
    frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
    frmBLMessageBoxJrWOpts.Show vbModal
    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
      Unload frmBLMessageBoxJrWOpts
      Close
      fpcmbAcctMethod.SetFocus
      Exit Sub
    Else
      Unload frmBLMessageBoxJrWOpts
      MainLog ("User warned that once saved the accounting method is permanently locked and cannot be changed. The user elected to continue saving.")
    End If
  End If
  
  If fpcmbAcctMethod.Text = "Accrual" Then
    If fptxtRevGLAcctNum.Text = "" Then
      frmBLMessageBoxJr.Label1.Caption = "Selecting 'Accrual' as the accounting method requires the 'Revenue' field to be filled in."
      frmBLMessageBoxJr.Label1.Top = 800
      frmBLMessageBoxJr.Show vbModal
      fptxtRevGLAcctNum.BackColor = &H80FFFF
      fptxtRevGLAcctNum.SetFocus
      Exit Sub
    ElseIf fptxtAcctsRec.Text = "" Then
      frmBLMessageBoxJr.Label1.Caption = "Selecting 'Accrual' as the accounting method requires the 'Accounts Receivable' field to be filled in."
      frmBLMessageBoxJr.Label1.Top = 800
      frmBLMessageBoxJr.Show vbModal
      fptxtAcctsRec.BackColor = &H80FFFF
      fptxtAcctsRec.SetFocus
      Exit Sub
    ElseIf fptxtCashReceipt.Text = "" Then
      frmBLMessageBoxJr.Label1.Caption = "Selecting 'Accrual' as the accounting method requires the 'Cash Receipt' field to be filled in."
      frmBLMessageBoxJr.Label1.Top = 800
      frmBLMessageBoxJr.Show vbModal
      fptxtCashReceipt.BackColor = &H80FFFF
      fptxtCashReceipt.SetFocus
      Exit Sub
    End If
  ElseIf fpcmbAcctMethod.Text = "Cash" Then
    If fptxtRevGLAcctNum.Text = "" Then
      frmBLMessageBoxJr.Label1.Caption = "Selecting 'Cash' as the accounting method requires the 'Revenue' field to be filled in."
      frmBLMessageBoxJr.Label1.Top = 800
      frmBLMessageBoxJr.Show vbModal
      fptxtRevGLAcctNum.BackColor = &H80FFFF
      fptxtRevGLAcctNum.SetFocus
      Exit Sub
    ElseIf fptxtCashReceipt.Text = "" Then
      frmBLMessageBoxJr.Label1.Caption = "Selecting 'Cash' as the accounting method requires the 'Cash Receipt' field to be filled in."
      frmBLMessageBoxJr.Label1.Top = 800
      frmBLMessageBoxJr.Show vbModal
      fptxtCashReceipt.BackColor = &H80FFFF
      fptxtCashReceipt.SetFocus
      Exit Sub
    End If
  End If

  If QPTrim$(fpcmbLaserYN.Text) = "1" Then
    If Not Exist("arlaser1.dat") Then
      frmBLMessageBoxJrWOpts.Label1.Caption = "You have elected to create advance letter #1 but no advance letter #1 has been created. Do you wish to continue saving anyway?"
      frmBLMessageBoxJrWOpts.Label1.Top = 700
      frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
      frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
      frmBLMessageBoxJrWOpts.Show vbModal
      If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
        Unload frmBLMessageBoxJrWOpts
        Close
        fpcmbLaserYN.SetFocus
        Exit Sub
      Else
        Unload frmBLMessageBoxJrWOpts
      End If
    End If
  End If
  
  If QPTrim$(fpcmbLaserYN.Text) = "2" Then
    If Not Exist("arlaser2.dat") Then
      frmBLMessageBoxJrWOpts.Label1.Caption = "You have elected to create advance letter #2 but no advance letter #2 has been created. Do you wish to continue saving anyway?"
      frmBLMessageBoxJrWOpts.Label1.Top = 700
      frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
      frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
      frmBLMessageBoxJrWOpts.Show vbModal
      If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
        Unload frmBLMessageBoxJrWOpts
        Close
        fpcmbLaserYN.SetFocus
        Exit Sub
      Else
        Unload frmBLMessageBoxJrWOpts
      End If
    End If
  End If
  
  If QPTrim$(fpcmbLaserYN.Text) = "3" Then
    If Not Exist("arlaser3.dat") Then
      frmBLMessageBoxJrWOpts.Label1.Caption = "You have elected to create advance letter #3 but no advance letter #3 has been created. Do you wish to continue saving anyway?"
      frmBLMessageBoxJrWOpts.Label1.Top = 700
      frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
      frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
      frmBLMessageBoxJrWOpts.Show vbModal
      If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
        Unload frmBLMessageBoxJrWOpts
        Close
        fpcmbLaserYN.SetFocus
        Exit Sub
      Else
        Unload frmBLMessageBoxJrWOpts
      End If
    End If
  End If
  
  If Mid(fpcmbDLQNotice.Text, 1, 1) = "4" Then
      If Not Exist("arlaser4.dat") Then
      frmBLMessageBoxJrWOpts.Label1.Caption = "You have elected to create delinquent notice #4 but no delinquent notice #4 has been created. Do you wish to continue saving anyway?"
      frmBLMessageBoxJrWOpts.Label1.Top = 700
      frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
      frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
      frmBLMessageBoxJrWOpts.Show vbModal
      If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
        Unload frmBLMessageBoxJrWOpts
        Close
        fpcmbDLQNotice.SetFocus
        Exit Sub
      Else
        Unload frmBLMessageBoxJrWOpts
      End If
    End If
  End If
  
  'if the town setup file exists then the file is retrieved and
  'the fields that have updates get updates
  If Exist("artownsu.dat") Then
    OpenTownFile THandle
    Get THandle, 1, TownRec
      TownRec.TownName = QPTrim$(fptxtTownName.Text)
      TownRec.Contact = QPTrim$(fptxtContact.Text)
      TownRec.TownAdd1 = QPTrim$(fptxtAdd1.Text)
      TownRec.TownAdd2 = QPTrim$(fptxtAdd2.Text)
      TownRec.City = QPTrim$(fptxtCity.Text)
      TownRec.State = QPTrim$(fptxtState.Text)
      TownRec.ZipCode = QPTrim$(fptxtZip.Text)
      TownRec.TownPhone = QPTrim$(fptxtPhone.Text)
      TownRec.SpareSpace = ""
      TownRec.AppForm = Val(Mid(fpcmbAppType.Text, 1))
      TownRec.DLQNotice = Val(Mid(fpcmbDLQNotice.Text, 1))
      TownRec.LicNumPermYN = QPTrim$(fpcmbNewLicNumYN.Text)
      TownRec.GL2Cats = QPTrim$(fpcmbGLToCats.Text)
      TownRec.UseAmtPctYN = QPTrim$(fpcmbAmtPct.Text)
      TownRec.PENCASHACCT = GetGLRecNum(fptxtCashReceipt.Text)
      If TownRec.PENCASHACCT = 0 Then
        If fpcmbAcctMethod.Text = "Accrual" Or fpcmbAcctMethod.Text = "Cash" Then
          fptxtCashReceipt.SetFocus
          If Not Exist("GLACCT.DAT") Or Not Exist("GLACCT.DAT") Then
            frmBLMessageBoxJr.Label1.Caption = "The 'Cash Receipt' number (" + QPTrim$(fptxtCashReceipt.Text) + ") cannot be verified because General Ledger data is missing. 'Cash Receipt' number not saved."
            frmBLMessageBoxJr.Label1.Top = 700
            frmBLMessageBoxJr.Show vbModal
            fptxtCashReceipt.Text = ""
            MainLog ("User warned that the cash receipt number entered, " + QPTrim$(fptxtCashReceipt.Text) + ", is not being saved because General Ledger data is missing.")
          Else
            frmBLMessageBoxJr.Label1.Caption = "No match for the 'Cash Receipt' number (" + QPTrim$(fptxtCashReceipt.Text) + ") found in the General Ledger number list. 'Cash Receipt' number not saved."
            frmBLMessageBoxJr.Label1.Top = 700
            frmBLMessageBoxJr.Show vbModal
            fptxtCashReceipt.Text = ""
            MainLog ("User warned that the cash receipt number entered, " + QPTrim$(fptxtCashReceipt.Text) + ", is not being saved because a match could not be found in the GL list.")
          End If
        End If
      End If
      
      TownRec.PENRECGLNUM = GetGLRecNum(fptxtAcctsRec.Text)
      If TownRec.PENRECGLNUM = 0 Then
        If fpcmbAcctMethod.Text = "Accrual" Then
          fptxtAcctsRec.SetFocus
          If Not Exist("GLACCT.DAT") Or Not Exist("GLACCT.DAT") Then
            frmBLMessageBoxJr.Label1.Caption = "The 'Accts Rec' number (" + QPTrim$(fptxtAcctsRec.Text) + ") cannot be verified because General Ledger data is missing. 'Accts Rec' number not saved."
            frmBLMessageBoxJr.Label1.Top = 700
            frmBLMessageBoxJr.Show vbModal
            fptxtAcctsRec.Text = ""
            MainLog ("User warned that the accounts receivable number entered, " + QPTrim$(fptxtAcctsRec.Text) + ", is not being saved because General Ledger data is missing.")
          Else
            frmBLMessageBoxJr.Label1.Caption = "No match for the 'Accts Rec' number (" + QPTrim$(fptxtAcctsRec) + ") found in the General Ledger number list. 'Accts Rec' number not saved."
            frmBLMessageBoxJr.Label1.Top = 700
            frmBLMessageBoxJr.Show vbModal
            fptxtAcctsRec.Text = ""
            MainLog ("User warned that the accounts receivable number entered, " + QPTrim$(fptxtAcctsRec.Text) + ", is not being saved because a match could not be found in the GL list.")
          End If
        End If
      End If
      
      TownRec.PENREVGLNUM = GetGLRecNum(fptxtRevGLAcctNum.Text)
      If TownRec.PENREVGLNUM = 0 Then
        If fpcmbAcctMethod.Text = "Accrual" Or fpcmbAcctMethod.Text = "Cash" Then
          fptxtRevGLAcctNum.SetFocus
          If Not Exist("GLACCT.DAT") Or Not Exist("GLACCT.DAT") Then
            frmBLMessageBoxJr.Label1.Caption = "The 'Revenue' number (" + QPTrim$(fptxtRevGLAcctNum.Text) + ") cannot be verified because General Ledger data is missing. 'Revenue' number not saved."
            frmBLMessageBoxJr.Label1.Top = 700
            frmBLMessageBoxJr.Show vbModal
            fptxtRevGLAcctNum.Text = ""
            MainLog ("User warned that the revenue number entered, " + QPTrim$(fptxtRevGLAcctNum.Text) + ", is not being saved because General Ledger data is missing.")
          Else
            frmBLMessageBoxJr.Label1.Caption = "No match for the 'Revenue' number (" + QPTrim$(fptxtRevGLAcctNum.Text) + ") found in the General Ledger number list. 'Revenue' number not saved."
            frmBLMessageBoxJr.Label1.Top = 700
            frmBLMessageBoxJr.Show vbModal
            fptxtRevGLAcctNum.Text = ""
            MainLog ("User warned that the revenue number entered, " + QPTrim$(fptxtRevGLAcctNum.Text) + ", is not being saved because a match could not be found in the GL list.")
          End If
        End If
      End If
      
      TownRec.AcctMeth = Mid(fpcmbAcctMethod.Text, 1, 1)
      TownRec.IssFee = fpcurrIssFee.DoubleValue
      TownRec.LaserLtr = Mid(fpcmbLaserYN.Text, 1, 1)
      GoSub CheckAppsDlq
    Put THandle, 1, TownRec
  Else 'if the town setup file is not already saved then
  'each field has to be saved with a value...if an application
  'or delinquent notice has been saved already then the town setup file
  'already exists
    TownRec.TownName = QPTrim$(fptxtTownName.Text)
    TownRec.Contact = QPTrim$(fptxtContact.Text)
    TownRec.TownAdd1 = QPTrim$(fptxtAdd1.Text)
    TownRec.TownAdd2 = QPTrim$(fptxtAdd2.Text)
    TownRec.City = QPTrim$(fptxtCity.Text)
    TownRec.State = QPTrim$(fptxtState.Text)
    TownRec.ZipCode = QPTrim$(fptxtZip.Text)
    TownRec.TownPhone = QPTrim$(fptxtPhone.Text)
    TownRec.SpareSpace = ""
    TownRec.AppForm = Val(Mid(fpcmbAppType.Text, 1))
    TownRec.DLQNotice = Val(Mid(fpcmbDLQNotice.Text, 1))
    TownRec.LicNumPermYN = QPTrim$(fpcmbNewLicNumYN.Text)
    TownRec.GL2Cats = QPTrim$(fpcmbGLToCats.Text)
    TownRec.UseAmtPctYN = QPTrim$(fpcmbAmtPct.Text)
    TownRec.PENCASHACCT = GetGLRecNum(fptxtCashReceipt.Text)
      If TownRec.PENCASHACCT = 0 Then
        If fpcmbAcctMethod.Text = "Accrual" Or fpcmbAcctMethod.Text = "Cash" Then
          fptxtCashReceipt.SetFocus
          If Not Exist("GLACCT.DAT") Or Not Exist("GLACCT.DAT") Then
            frmBLMessageBoxJr.Label1.Caption = "The 'Cash Receipt' number (" + QPTrim$(fptxtCashReceipt.Text) + ") cannot be verified because General Ledger data is missing. 'Cash Receipt' number not saved."
            frmBLMessageBoxJr.Label1.Top = 700
            frmBLMessageBoxJr.Show vbModal
            fptxtCashReceipt.Text = ""
            MainLog ("User warned that the cash receipt number entered, " + QPTrim$(fptxtCashReceipt.Text) + ", is not being saved because General Ledger data is missing.")
          Else
            frmBLMessageBoxJr.Label1.Caption = "No match for the 'Cash Receipt' number (" + QPTrim$(fptxtCashReceipt.Text) + ") found in the General Ledger number list. 'Cash Receipt' number not saved."
            frmBLMessageBoxJr.Label1.Top = 700
            frmBLMessageBoxJr.Show vbModal
            fptxtCashReceipt.Text = ""
            MainLog ("User warned that the cash receipt number entered, " + QPTrim$(fptxtCashReceipt.Text) + ", is not being saved because a match could not be found in the GL list.")
          End If
        End If
      End If
    TownRec.PENRECGLNUM = GetGLRecNum(fptxtAcctsRec.Text)
      If TownRec.PENRECGLNUM = 0 Then
        If fpcmbAcctMethod.Text = "Accrual" Then
          fptxtAcctsRec.SetFocus
          If Not Exist("GLACCT.DAT") Or Not Exist("GLACCT.DAT") Then
            frmBLMessageBoxJr.Label1.Caption = "The 'Accts Rec' number (" + QPTrim$(fptxtAcctsRec.Text) + ") cannot be verified because General Ledger data is missing. 'Accts Rec' number not saved."
            frmBLMessageBoxJr.Label1.Top = 700
            frmBLMessageBoxJr.Show vbModal
            fptxtAcctsRec.Text = ""
            MainLog ("User warned that the accounts receivable number entered, " + QPTrim$(fptxtAcctsRec.Text) + ", is not being saved because General Ledger data is missing.")
          Else
            frmBLMessageBoxJr.Label1.Caption = "No match for the 'Accts Rec' number (" + QPTrim$(fptxtAcctsRec) + ") found in the General Ledger number list. 'Accts Rec' number not saved."
            frmBLMessageBoxJr.Label1.Top = 700
            frmBLMessageBoxJr.Show vbModal
            fptxtAcctsRec.Text = ""
            MainLog ("User warned that the accounts receivable number entered, " + QPTrim$(fptxtAcctsRec.Text) + ", is not being saved because a match could not be found in the GL list.")
          End If
        End If
      End If
    TownRec.PENREVGLNUM = GetGLRecNum(fptxtRevGLAcctNum.Text)
      If TownRec.PENREVGLNUM = 0 Then
        If fpcmbAcctMethod.Text = "Accrual" Or fpcmbAcctMethod.Text = "Cash" Then
          fptxtRevGLAcctNum.SetFocus
          If Not Exist("GLACCT.DAT") Or Not Exist("GLACCT.DAT") Then
            frmBLMessageBoxJr.Label1.Caption = "The 'Revenue' number (" + QPTrim$(fptxtRevGLAcctNum.Text) + ") cannot be verified because General Ledger data is missing. 'Revenue' number not saved."
            frmBLMessageBoxJr.Label1.Top = 700
            frmBLMessageBoxJr.Show vbModal
            fptxtRevGLAcctNum.Text = ""
            MainLog ("User warned that the revenue number entered, " + QPTrim$(fptxtRevGLAcctNum.Text) + ", is not being saved because General Ledger data is missing.")
          Else
            frmBLMessageBoxJr.Label1.Caption = "No match for the 'Revenue' number (" + QPTrim$(fptxtRevGLAcctNum.Text) + ") found in the General Ledger number list. 'Revenue' number not saved."
            frmBLMessageBoxJr.Label1.Top = 700
            frmBLMessageBoxJr.Show vbModal
            fptxtRevGLAcctNum.Text = ""
            MainLog ("User warned that the revenue number entered, " + QPTrim$(fptxtRevGLAcctNum.Text) + ", is not being saved because a match could not be found in the GL list.")
          End If
        End If
      End If
    TownRec.AcctMeth = Mid(fpcmbAcctMethod.Text, 1, 1)
    TownRec.IssFee = fpcurrIssFee.DoubleValue
    TownRec.LaserLtr = Mid(fpcmbLaserYN.Text, 1, 1)
    TownRec.AppAdd1 = ""
    TownRec.AppBaseFee(1) = 0
    TownRec.AppBaseFee(2) = 0
    TownRec.AppBaseFee(3) = 0
    TownRec.AppBaseFee(4) = 0
    TownRec.AppCentsPer(1) = 0
    TownRec.AppCentsPer(2) = 0
    TownRec.AppCentsPer(3) = 0
    TownRec.AppCentsPer(4) = 0 '20
    TownRec.AppFirstDay = ""
    TownRec.AppLastDay = ""
    TownRec.AppGrsRcpts(1) = 0
    TownRec.AppGrsRcpts(2) = 0
    TownRec.AppGrsRcpts(3) = 0
    TownRec.AppGrsRcpts(4) = 0
    TownRec.AppColFee = 0
    TownRec.AppGrsPct = 0
    TownRec.AppDenom = 0
    TownRec.AppNumer = 0
    TownRec.AppState = ""
    TownRec.AppCity = ""
    TownRec.AppTownOf = ""
    TownRec.AppPayBy = 0
    TownRec.AppZip = "" '30
    TownRec.AppAdminName = ""
    TownRec.AppAdminTitle = ""
    TownRec.AppPhone = ""
    TownRec.AppPct = 0
    TownRec.AppDiscPct = 0
    TownRec.AppDiscMonth = ""
    TownRec.AppDiscDay = 0
    TownRec.AppPenMonth = ""
    TownRec.AppPenDay = 0
    TownRec.AppFiscMonth = ""
    TownRec.AppFiscDay = 0
    TownRec.AppMayorCouncil = ""
    TownRec.AppWholeMonth = 0
    TownRec.AppWholeDay = 0
    TownRec.AppRetailMonth = 0
    TownRec.AppRetailDay = 0
    TownRec.AppFinMonth = 0
    TownRec.AppFinDay = 0
    TownRec.AppContMonth = 0
    TownRec.AppContDay = 0
    TownRec.AppRepairMonth = 0
    TownRec.AppRepairDay = 0
    TownRec.AppStartMonth = ""
    TownRec.AppStartDay = 0
    TownRec.AppLicRetMonth = ""
    TownRec.AppLicRetDay = 0
    TownRec.AppAdoptDate = 0
    TownRec.AppCityOrd = ""
    For x = 1 To 10
     TownRec.AppYrUpDown(x) = "0"
    Next x
    TownRec.DlqAdd1 = ""
    TownRec.DlqAdminName = "" 'used on #1
    TownRec.DlqAdminTitle = "" 'used on #1
    TownRec.DlqCity = ""
    TownRec.DlqPhone = ""
    TownRec.DlqPhone2 = "" 'used on #2
    TownRec.DlqFax = "" 'used on #2
    TownRec.DlqState = ""
    TownRec.DlqTownName = ""
    TownRec.DlqZip = ""
    TownRec.DlqFirstDay = ""
    TownRec.DlqLastDay = ""
    TownRec.DlqFirstHour = ""
    TownRec.DlqLastHour = ""
    TownRec.DlqClerkName = ""
    TownRec.DlqMayorCouncil = ""
    GoSub CheckAppsDlq
    OpenTownFile THandle
    Put THandle, 1, TownRec
  End If
    
  Close THandle
  
  'LogSaves records all saved data to arlog
  Call LogSaves
    
  fpcmbAcctMethod.Enabled = False
  
  frmBLSucSave.Label1.Caption = "Your Town Setup data has been saved successfully."
  frmBLSucSave.Label1.Top = 700
  frmBLSucSave.Show vbModal
  If Exist("pencalc.dat") Then
    frmBLPenProcMenu.Show
    DoEvents
    Unload frmBLTownSetup
  End If
  AcctMethodChangeFlag = False
  Exit Sub
  
CheckAppsDlq: 'this sub looks for which application or delinquent
  'notice is being saved and then checks certain fields to see if it was edited
  'correctly...this is not foolproof but reduces the probability of an edit
  'failure from occurring
  Select Case TownRec.DLQNotice
    Case 1 'If this field is not filled in then the user
    'has not saved delinquent notice #1
      If QPTrim$(TownRec.DlqAdminName) = "" Then
        frmBLMessageBoxJr.Label1.Caption = "Please review Delinquent Form #1 and save it from that screen before exiting."
        frmBLMessageBoxJr.Label1.Top = 700
        frmBLMessageBoxJr.Show vbModal
        fpcmbDLQNotice.SetFocus
        Exit Sub
      End If
    Case 2
      If QPTrim$(TownRec.DlqTownName) = "" Then
        frmBLMessageBoxJr.Label1.Caption = "Please review Delinquent Form #2 and save it from that screen before exiting."
        frmBLMessageBoxJr.Label1.Top = 700
        frmBLMessageBoxJr.Show vbModal
        fpcmbDLQNotice.SetFocus
        Exit Sub
      End If
    Case 3
      If QPTrim$(TownRec.DlqClerkName) = "" Then
        frmBLMessageBoxJr.Label1.Caption = "Please review Delinquent Form #3 and save it from that screen before exiting."
        frmBLMessageBoxJr.Label1.Top = 700
        frmBLMessageBoxJr.Show vbModal
        fpcmbDLQNotice.SetFocus
        Exit Sub
      End If
    Case Else
  End Select
  
  Select Case TownRec.AppForm 'you can save an app type
  'without ever going to the app template...so if an app is saved
  'here we have to ensure that all fields are right if it is to
  'print out correctly...we only have to check one of the
  'required fields that is exclusive to that form to see if
  'it was saved...it is possible to circumvent this check
  'by saving one of the checked fields in one application
  'that is shared by another application
    Case 2
      For x = 1 To 3
        If TownRec.AppGrsRcpts(x) > 0 Then Exit For 'at least 1 of these should be greater than 0
        'it is possible to save all 3 as zero but highly unlikely
      Next x
      If x > 3 Then
        frmBLMessageBoxJr.Label1.Caption = "You have selected Renewal Application #2 but all gross dollar amounts are zero. Please complete Renewal Application #2 before exiting."
        frmBLMessageBoxJr.Label1.Top = 800
        frmBLMessageBoxJr.Show vbModal
        fpcmbAppType.SetFocus
        Exit Sub
      End If
    Case 3
      For x = 1 To 4
        If TownRec.AppBaseFee(x) > 0 Then Exit For 'at least 1 of these should be greater than 0
        'it is possible to save all 4 as zero but highly unlikely
      Next x
      If x > 4 Then
        frmBLMessageBoxJr.Label1.Caption = "You have selected Renewal Application #3 but all Base Fees are zero. Please complete Renewal Application #3 before exiting."
        frmBLMessageBoxJr.Label1.Top = 800
        frmBLMessageBoxJr.Show vbModal
        fpcmbAppType.SetFocus
        Exit Sub
      End If
    Case 4
      If QPTrim$(TownRec.AppCityOrd) = "" Then
        frmBLMessageBoxJr.Label1.Caption = "You have selected Renewal Application #4 but some fields are empty. Please complete Renewal Application #4 before exiting."
        frmBLMessageBoxJr.Label1.Top = 800
        frmBLMessageBoxJr.Show vbModal
        fpcmbAppType.SetFocus
        Exit Sub
      End If
    Case 5
      For x = 1 To 4
        If TownRec.AppCentsPer(x) > 0 Then Exit For
      Next x
      If x > 4 Then
        frmBLMessageBoxJr.Label1.Caption = "You have selected Renewal Application #5 but all Cents Per Gross amounts are zero. Please complete Renewal Application #5 before exiting."
        frmBLMessageBoxJr.Label1.Top = 800
        frmBLMessageBoxJr.Show vbModal
        fpcmbAppType.SetFocus
        Exit Sub
      End If
    Case 6
      If TownRec.AppPct = 0 And TownRec.AppGrsPct = 0 Then
        frmBLMessageBoxJr.Label1.Caption = "You have selected Renewal Application #6 but some fields are empty. Please complete Renewal Application #6 before exiting."
        frmBLMessageBoxJr.Label1.Top = 800
        frmBLMessageBoxJr.Show vbModal
        fpcmbAppType.SetFocus
        Exit Sub
      End If
    Case 7
      If QPTrim$(TownRec.AppFiscMonth) = "" Then
        frmBLMessageBoxJr.Label1.Caption = "You have selected Renewal Application #7 but some fields are empty. Please complete Renewal Application #7 before exiting."
        frmBLMessageBoxJr.Label1.Top = 800
        frmBLMessageBoxJr.Show vbModal
        fpcmbAppType.SetFocus
        Exit Sub
      End If
    Case 8
      If QPTrim$(TownRec.AppFiscMonth) = "" Then
        frmBLMessageBoxJr.Label1.Caption = "You have selected Renewal Application #8 but some fields are empty. Please complete Renewal Application #8 before exiting."
        frmBLMessageBoxJr.Label1.Top = 800
        frmBLMessageBoxJr.Show vbModal
        fpcmbAppType.SetFocus
        Exit Sub
      End If
    Case 9
      If TownRec.AppPct = 0 Then
        frmBLMessageBoxJr.Label1.Caption = "You have selected Renewal Application #9 but some fields are empty. Please complete Renewal Application #9 before exiting."
        frmBLMessageBoxJr.Label1.Top = 800
        frmBLMessageBoxJr.Show vbModal
        fpcmbAppType.SetFocus
        Exit Sub
      End If
    Case Else
  End Select
    
  Return

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLTownSetup", "cmdSave_Click", Erl)
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

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsBLTextBoxOverrider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  'JustOpened is a field that keeps a message from recurring
  'if the amt/pct field is changed...it only happens once no
  'matter how many times that field is changed
  JustOpened = True
  Call LoadMe
  JustOpened = False
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
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
    Case vbKeyF4:
      SendKeys "%P"
      Call fpcmdApps_Click
      KeyCode = 0
    Case vbKeyF2:
      SendKeys "%L"
      Call cmdAdvanceLtr_Click
      KeyCode = 0
    Case vbKeyF6:
      SendKeys "%Q"
      Call fpcmdDLQ_Click
      KeyCode = 0
    Case vbKeyF12:
      SendKeys "%G"
      Call cmdGLList_Click
      KeyCode = 0
    Case vbKeyF1:
      SendKeys "%T"
      Call cmdHelp_Click
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
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLTownSetup.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  Dim TownRec As TownSetUpType
  Dim THandle As Integer
  Dim One As Integer
  Dim DHandle As Integer
  
  On Error GoTo ERRORSTUFF
  
  lblBalloon.Visible = False
  
'  cmdHelp.ToolTipText = "Click on this button to activate informational balloons for each field."
'  fptxtTownName.ToolTipText = "Enter the town's offical name here. (Ex. Town Of Riverville)"
'  fptxtContact.ToolTipText = "Enter the name of a town office contact."
'  fptxtAdd1.ToolTipText = "Enter the town's street mailing address."
'  fptxtAdd2.ToolTipText = "Enter a secondary mailing address here."
'  fptxtCity.ToolTipText = "Enter the town's mailing name here."
'  fptxtState.ToolTipText = "Enter the town's state here."
'  fptxtZip.ToolTipText = "Enter the town's zip code here."
'  fptxtPhone.ToolTipText = "Enter the town office's phone number."
'  fpcmbAcctMethod.ToolTipText = "Choose one of the three accounting methods supported in business license."
  
  AcctMethodChangeFlag = False
  One = 1
  DHandle = FreeFile
  Open "townsetup.dat" For Output As DHandle Len = 2
  Print #DHandle, One
  Close DHandle
  
  If Exist("artownsu.dat") Then
    OpenTownFile THandle
    Get THandle, 1, TownRec
    Close THandle
    fpcmbAcctMethod.Enabled = False
  Else
    GoTo NoFile
  End If
  
  fpcmbGLToCats.AddItem "Yes"
  fpcmbGLToCats.AddItem "No"
  fpcmbAmtPct.AddItem "Amt"
  fpcmbAmtPct.AddItem "Pct"
  fpcmbNewLicNumYN.AddItem "Yes"
  fpcmbNewLicNumYN.AddItem "No"
  fpcmbAcctMethod.AddItem "None"
  fpcmbAcctMethod.AddItem "Cash"
  fpcmbAcctMethod.AddItem "Accrual"
  fpcmbLaserYN.AddItem "0"
  fpcmbLaserYN.AddItem "1"
  fpcmbLaserYN.AddItem "2"
  fpcmbLaserYN.AddItem "3"
  
  'at this point if the town .dat file exists it has
  'been opened and it's ok to read from it
  
  fptxtTownName.Text = QPTrim$(TownRec.TownName)
  'Temp variables are saved for use in the
  'LogSaves sub
  TempTownName$ = QPTrim$(TownRec.TownName)
  fptxtContact.Text = QPTrim$(TownRec.Contact)
  TempTownContact$ = QPTrim$(TownRec.Contact)
  fptxtAdd1.Text = QPTrim$(TownRec.TownAdd1)
  TempTownAdd1$ = QPTrim$(TownRec.TownAdd1)
  fptxtAdd2.Text = QPTrim$(TownRec.TownAdd2)
  TempTownAdd2$ = QPTrim$(TownRec.TownAdd2)
  fptxtCity.Text = QPTrim$(TownRec.City)
  TempTownCity$ = QPTrim$(TownRec.City)
  fptxtState.Text = QPTrim$(TownRec.State)
  TempTownState$ = QPTrim$(TownRec.State)
  fptxtZip.Text = ReplaceString$(TownRec.ZipCode, "-", "")
  TempTownZip$ = ReplaceString$(TownRec.ZipCode, "-", "")
  fptxtPhone.Text = QPTrim$(TownRec.TownPhone)
  TempTownPhone$ = QPTrim$(TownRec.TownPhone)
  fpcmbNewLicNumYN.Text = QPTrim$(TownRec.LicNumPermYN)
  TempLicNew = QPTrim$(TownRec.LicNumPermYN)
  fpcmbAmtPct.Text = QPTrim$(TownRec.UseAmtPctYN)
  TempAmtPct = QPTrim$(TownRec.UseAmtPctYN)
  If TownRec.GL2Cats = "N" Then
    fpcmbGLToCats.Text = "No"
    TempGLToCats = "N"
  ElseIf TownRec.GL2Cats = "Y" Then
    fpcmbGLToCats.Text = "Yes"
    TempGLToCats = "Y"
  Else
    fpcmbGLToCats.Text = "No"
    TempGLToCats = "N"
  End If
  
  Select Case QPTrim$(TownRec.AcctMeth)
    Case "A"
      fpcmbAcctMethod.Text = "Accrual"
      TempAcctMethod = "Accrual"
      LastAcctMethod$ = "Accrual"
    Case "N"
      cmdGLList.Enabled = False
      fpcmbAcctMethod.Text = "None"
      TempAcctMethod = "None"
      LastAcctMethod$ = "None"
    Case "C"
      fpcmbAcctMethod.Text = "Cash"
      TempAcctMethod = "Cash"
      LastAcctMethod$ = "Cash"
  End Select
  
  fptxtRevGLAcctNum.Text = GetGLNum(TownRec.PENREVGLNUM)
  TempPenRevNum = GetGLNum(TownRec.PENREVGLNUM)
  'Last variables are used to reset the GL number fields
  'back to the last valid number if the user makes a mistake
  LastPenRevNum = GetGLNum(TownRec.PENREVGLNUM)
  fptxtAcctsRec.Text = GetGLNum(TownRec.PENRECGLNUM)
  TempPenRecNum = GetGLNum(TownRec.PENRECGLNUM)
  LastPenRecNum = GetGLNum(TownRec.PENRECGLNUM)
  fptxtCashReceipt.Text = GetGLNum(TownRec.PENCASHACCT)
  TempCashRecpt = GetGLNum(TownRec.PENCASHACCT)
  LastCashRecpt = GetGLNum(TownRec.PENCASHACCT)

  Select Case QPTrim$(TownRec.LaserLtr)
    Case "1"
      fpcmbLaserYN.Text = "1"
      TempLaserYN = "1"
    Case "2"
      fpcmbLaserYN.Text = "2"
      TempLaserYN = "2"
    Case "3"
      fpcmbLaserYN.Text = "3"
      TempLaserYN = "3"
    Case Else
      fpcmbLaserYN.Text = "0"
      TempLaserYN = "0"
  End Select
  
  fpcurrIssFee.Text = TownRec.IssFee
  TempIssFee = TownRec.IssFee
  Select Case TownRec.AppForm
    Case 1
      fpcmbAppType.Text = "1. APP STANDARD"
    Case 2
      fpcmbAppType.Text = "2. APP FORM A"
    Case 3
      fpcmbAppType.Text = "3. APP FORM B"
    Case 4
      fpcmbAppType.Text = "4. APP FORM C"
    Case 5
      fpcmbAppType.Text = "5. APP FORM D"
    Case 6
      fpcmbAppType.Text = "6. APP FORM E"
    Case 7
      fpcmbAppType.Text = "7. APP FORM F"
    Case 8
      fpcmbAppType.Text = "8. APP FORM G"
    Case 9
      fpcmbAppType.Text = "9. APP FORM H"
    Case 10
      fpcmbAppType.Text = "10. APP FREE FORMAT"
    Case 11
      fpcmbAppType.Text = "11. NONE"
    Case Else
      fpcmbAppType.Text = "11. NONE"
  End Select
  TempTownAppForm = TownRec.AppForm
  
  Select Case TownRec.DLQNotice
    Case 1
      fpcmbDLQNotice.Text = "1. PENALTY STANDARD"
    Case 2
      fpcmbDLQNotice.Text = "2. PENALTY FORM A"
    Case 3
      fpcmbDLQNotice.Text = "3. PENALTY FORM B"
    Case 4
      fpcmbDLQNotice.Text = "4. FREE FORMAT"
    Case Else
      fpcmbDLQNotice.Text = "5. NONE"
  End Select
  
  TempTownPenForm = TownRec.DLQNotice
  fpcmbAppType.AddItem "1. APP STANDARD"
  fpcmbAppType.AddItem "2. APP FORM A"
  fpcmbAppType.AddItem "3. APP FORM B"
  fpcmbAppType.AddItem "4. APP FORM C"
  fpcmbAppType.AddItem "5. APP FORM D"
  fpcmbAppType.AddItem "6. APP FORM E"
  fpcmbAppType.AddItem "7. APP FORM F"
  fpcmbAppType.AddItem "8. APP FORM G"
  fpcmbAppType.AddItem "9. APP FORM H"
  fpcmbAppType.AddItem "10. APP FREE FORMAT"
  fpcmbAppType.AddItem "11. NONE"
  fpcmbDLQNotice.AddItem "1. PENALTY STANDARD"
  fpcmbDLQNotice.AddItem "2. PENALTY FORM A"
  fpcmbDLQNotice.AddItem "3. PENALTY FORM B"
  fpcmbDLQNotice.AddItem "4. FREE FORMAT"
  fpcmbDLQNotice.AddItem "5. NONE"
  MainLog ("Town setup screen opened.")
  Exit Sub
  
NoFile: 'used if the town file has not yet been saved
  TempTownName$ = ""
  TempTownContact$ = ""
  TempTownAdd1$ = ""
  TempTownAdd2$ = ""
  TempTownCity$ = ""
  TempTownState$ = ""
  TempTownZip$ = ""
  TempTownPhone$ = ""
  TempTownAppForm = 0
  TempTownPenForm = 0
  TempLicNew = "No"
  TempAmtPct = "Pct"
  TempAcctMethod$ = "None"
  cmdGLList.Enabled = False
  TempPenRevNum$ = "0"
  TempPenRecNum$ = "0"
  TempCashRecpt$ = "0"
  TempIssFee = 0
  
  fpcmbAppType.Text = "11. NONE"
  fpcmbAppType.AddItem "1. APP STANDARD"
  fpcmbAppType.AddItem "2. APP FORM A"
  fpcmbAppType.AddItem "3. APP FORM B"
  fpcmbAppType.AddItem "4. APP FORM C"
  fpcmbAppType.AddItem "5. APP FORM D"
  fpcmbAppType.AddItem "6. APP FORM E"
  fpcmbAppType.AddItem "7. APP FORM F"
  fpcmbAppType.AddItem "8. APP FORM G"
  fpcmbAppType.AddItem "9. APP FORM H"
  fpcmbAppType.AddItem "10. APP FREE FORMAT"
  fpcmbAppType.AddItem "11. NONE"
  fpcmbDLQNotice.Text = "5. NONE"
  fpcmbDLQNotice.AddItem "1. PENALTY STANDARD"
  fpcmbDLQNotice.AddItem "2. PENALTY FORM A"
  fpcmbDLQNotice.AddItem "3. PENALTY FORM B"
  fpcmbDLQNotice.AddItem "4. FREE FORMAT"
  fpcmbDLQNotice.AddItem "5. NONE"
  fpcmbGLToCats.Text = "No"
  fpcmbGLToCats.AddItem "No"
  fpcmbGLToCats.AddItem "Yes"
  fpcmbNewLicNumYN.Text = "No"
  fpcmbNewLicNumYN.AddItem "Yes"
  fpcmbNewLicNumYN.AddItem "No"
  fpcmbAmtPct.Text = "Pct"
  fpcmbAmtPct.AddItem "Amt"
  fpcmbAmtPct.AddItem "Pct"
  fpcmbAcctMethod.Text = "None"
  fpcmbAcctMethod.AddItem "None"
  fpcmbAcctMethod.AddItem "Cash"
  fpcmbAcctMethod.AddItem "Accrual"
  fpcmbLaserYN.Text = "0"
  fpcmbLaserYN.AddItem "0"
  fpcmbLaserYN.AddItem "1"
  fpcmbLaserYN.AddItem "2"
  fpcmbLaserYN.AddItem "3"
  fpcurrIssFee.Text = 0
  
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLTownSetup", "LoadMe", Erl)
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
Private Sub Form_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub Check4AmtPct(AppForm As Integer, DlqForm As Integer)
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  
  On Error GoTo ERRORSTUFF
  
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  'if the user decides to change this field there are consequences that will affect
  'the way existing application's and delinquent form's penalty references display...
  'this code locates where the change is an issue and reports it to the user
  Select Case AppForm
    Case 2, 3, 5, 6:
      frmBLMessageBoxJr.Label1.Caption = "Changing this value will affect one field on your saved Application Renewal Form. Please make sure this field reflects the current penalty charge correctly."
      frmBLMessageBoxJr.Label1.Top = 800
      frmBLMessageBoxJr.Show vbModal
      MainLog ("WARNING: Issued to indicate the affect that changing the penalty charge (Amount/Percent) will have on the Application Renewal Form.")
    Case 8, 9:
      frmBLMessageBoxJr.Label1.Caption = "Changing this value will affect two fields on your saved Application Renewal Form. Please make sure these fields reflect the current penalty charge correctly."
      frmBLMessageBoxJr.Label1.Top = 800
      frmBLMessageBoxJr.Show vbModal
      MainLog ("WARNING: Issued to indicate the affect that changing the penalty charge (Amount/Percent) will have on the Application Renewal Form.")
    Case Else:
  
  End Select

  Select Case DlqForm
    Case 1:
      frmBLMessageBoxJr.Label1.Caption = "Changing this value will affect two fields on your saved Delinquent Notice Form. Please make sure these fields reflect the current penalty charge correctly."
      frmBLMessageBoxJr.Label1.Top = 800
      frmBLMessageBoxJr.Show vbModal
      MainLog ("WARNING: Issued to indicate the affect that changing the penalty charge (Amount/Percent) will have on the Delinquent Notice.")
    Case 2:
      frmBLMessageBoxJr.Label1.Caption = "Changing this value will affect one field on your saved Delinquent Notice Form. Please make sure this field reflects the current penalty charge correctly."
      frmBLMessageBoxJr.Label1.Top = 800
      frmBLMessageBoxJr.Show vbModal
      MainLog ("WARNING: Issued to indicate the affect that changing the penalty charge (Amount/Percent) will have on the Delinquent Notice.")
    Case Else:
    
  End Select
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLTownSetup", "Check4AmtPct", Erl)
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

Private Sub fpcmbAcctMethod_Change()
  
  On Error GoTo ERRORSTUFF
  
  If TempAcctMethod$ <> "" Then
    If AcctMethodChangeFlag = False Then
      If QPTrim$(fpcmbAcctMethod.Text) <> TempAcctMethod$ Then
        AcctMethodChangeFlag = True
      End If
    End If
  End If
  
  If QPTrim$(fpcmbAcctMethod.Text) = "" Then fpcmbAcctMethod.Text = "None"
  cmdGLList.Enabled = True
  
  If QPTrim$(fpcmbAcctMethod.Text) = "None" Then
    If QPTrim$(fptxtRevGLAcctNum.Text) <> "" Or QPTrim$(fptxtAcctsRec.Text) <> "" Or QPTrim$(fptxtCashReceipt.Text) <> "" Then
      fptxtRevGLAcctNum.BackColor = &H80FFFF
      fptxtAcctsRec.BackColor = &H80FFFF
      fptxtCashReceipt.BackColor = &H80FFFF
      frmBLMessageBoxJrWOpts.Label1.Caption = "Choosing 'None' as the accounting method eliminates the need for a 'Revenue' number, an 'Accounts Receivable' number and a 'Cash Receipts' number. Continuing will delete these numbers. Do you wish to continue?"
      frmBLMessageBoxJrWOpts.Label1.Top = 700
      frmBLMessageBoxJrWOpts.Show vbModal
      If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
        Unload frmBLMessageBoxJrWOpts
        fptxtRevGLAcctNum.BackColor = &HFFFFFF 'no
        fptxtAcctsRec.BackColor = &HFFFFFF
        fptxtCashReceipt.BackColor = &HFFFFFF
        fpcmbAcctMethod.Text = LastAcctMethod$
        Exit Sub
      Else
        Unload frmBLMessageBoxJrWOpts
        cmdGLList.Enabled = False
        fptxtRevGLAcctNum.Text = ""
        fptxtAcctsRec.Text = ""
        fptxtCashReceipt.Text = ""
        fptxtRevGLAcctNum.BackColor = &HFFFFFF
        fptxtAcctsRec.BackColor = &HFFFFFF
        fptxtCashReceipt.BackColor = &HFFFFFF
        fptxtRevGLAcctNum.Enabled = False
        fptxtAcctsRec.Enabled = False
        fptxtCashReceipt.Enabled = False
      End If
    Else
      fptxtRevGLAcctNum.Enabled = False
      fptxtAcctsRec.Enabled = False
      fptxtCashReceipt.Enabled = False
      cmdGLList.Enabled = False
    End If
  ElseIf fpcmbAcctMethod.Text = "Cash" Then
    If fptxtAcctsRec.Text <> "" Then
      fptxtAcctsRec.BackColor = &H80FFFF
      frmBLMessageBoxJrWOpts.Label1.Caption = "Choosing 'Cash' as the accounting method eliminates the need for an 'Accounts Receivable' number. Continuing will erase this number. Do you wish to continue?"
      frmBLMessageBoxJrWOpts.Label1.Top = 700
      frmBLMessageBoxJrWOpts.Show vbModal
      If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
        Unload frmBLMessageBoxJrWOpts
        fptxtAcctsRec.BackColor = &HFFFFFF
        fpcmbAcctMethod.Text = LastAcctMethod$
        Exit Sub
      Else
        Unload frmBLMessageBoxJrWOpts
        fptxtAcctsRec.Text = ""
        fptxtAcctsRec.BackColor = &HFFFFFF
      End If
    End If
    fptxtAcctsRec.Enabled = False
    fptxtCashReceipt.Enabled = True
    fptxtRevGLAcctNum.Enabled = True
  ElseIf fpcmbAcctMethod.Text = "Accrual" Then
    fptxtAcctsRec.Enabled = True
    fptxtCashReceipt.Enabled = True
    fptxtRevGLAcctNum.Enabled = True
  End If
  LastAcctMethod$ = QPTrim$(fpcmbAcctMethod.Text)
  
  Exit Sub

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLTownSetup", "fpcmbAcctMethod_Change", Erl)
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

Private Sub fpcmbAcctMethod_Click()
  'if this field was highlighted because of a user alert then
  'this code resets the backcolor to white whenever this field
  'is clicked again
  fpcmbAcctMethod.BackColor = &HFFFFFF
End Sub

Private Sub fpcmbAcctMethod_KeyDown(KeyCode As Integer, Shift As Integer)
  fpcmbAcctMethod.BackColor = &HFFFFFF
  If KeyCode = vbKeySpace Then
    fpcmbAcctMethod.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbAcctMethod.ListIndex = -1
  End If
  If fpcmbAcctMethod.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      If fptxtRevGLAcctNum.Enabled = True Then
        fptxtRevGLAcctNum.SetFocus
      Else
        fpcurrIssFee.SetFocus
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

Private Sub fpcmbAmtPct_Change()
  Static Message As Integer
  
  'this sub prevents the warning pop up from displaying every time the user changes the
  'value in the penalty amount/percent field...it allows the pop up after the first
  'change and then blocks it after any other change during this session...if the user
  'leaves this screen and comes back then the session begins again
  If JustOpened = True Then
    Message = 1
  Else
    If Message = 1 Then
      frmBLMessageBoxJr.Label1.Caption = "Changing the penalty charge option may affect the Application Renewal form and the Delinquent Notice. For example, if a $25.00 penalty charge was saved and the penalty charge option was changed to Percent from Amount then your forms will now show a 25% penalty charge."
      frmBLMessageBoxJr.Label1.Height = 1500
      frmBLMessageBoxJr.Show vbModal
      Message = Message + 1
      MainLog ("Penalty charge option was changed and the user was warned of the consequences to the delinquent notices and renewal applications.")
    End If
  End If
  
End Sub

Private Sub fpcmbAmtPct_Click()
  If QPTrim$(fpcmbAmtPct.Text) = "" Then
    fpcmbAmtPct.Text = "Pct"
  End If
End Sub

Private Sub fpcmbAmtPct_DropDown()
  fpcmbAmtPct.BackColor = &HFFFFFF
End Sub

Private Sub fpcmbAmtPct_KeyDown(KeyCode As Integer, Shift As Integer)
  fpcmbAmtPct.BackColor = &HFFFFFF
  If KeyCode = vbKeySpace Then
    fpcmbAmtPct.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbAmtPct.ListIndex = -1
  End If
  If fpcmbAmtPct.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbLaserYN.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbAppType_Change()
  'resets the default to '10. NONE' if the user leaves the
  'field blank
  If QPTrim$(fpcmbAppType.Text) = "" Then
    fpcmbAppType.Text = "11. NONE"
  End If
  'this code resets the command button to coincide with
  'application 10 or whatever application less than 10
  'which is showing in this field...10 had to be singled
  'out because it is two digits
  If fpcmbAppType.Text = "11. NONE" Then
    fpcmdApps.Text = "No Selection"
    fpcmdApps.Enabled = False
  Else
    If InStr(fpcmbAppType.Text, "FREE") Then
      fpcmdApps.Text = "F4 S&how App Type 10"
    Else
      fpcmdApps.Text = "F4 S&how App Type " + Mid(fpcmbAppType.Text, 1, 1)
    End If
    fpcmdApps.Enabled = True
  End If
End Sub

Private Sub fpcmbAppType_DropDown()
  fpcmbAppType.BackColor = &HFFFFFF
End Sub

Private Sub fpcmbAppType_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  fpcmbAppType.BackColor = &HFFFFFF
  If KeyCode = vbKeySpace Then
    fpcmbAppType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbAppType.ListIndex = -1
  End If
  If fpcmbAppType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbDLQNotice.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbDLQNotice_Change()
  'this field defaults to '4. NONE' if the user leaves the
  'field empty
  If QPTrim$(fpcmbDLQNotice.Text) = "" Then
    fpcmbDLQNotice.Text = "5. NONE"
  End If
  'resets the command button to reflect choice in combo box
  If QPTrim$(fpcmbDLQNotice.Text) = "5. NONE" Then
    fpcmdDLQ.Text = "No Selection"
    fpcmdDLQ.Enabled = False
  Else
    fpcmdDLQ.Text = "F6 Show Dl&q Notice " + Mid(fpcmbDLQNotice.Text, 1, 1)
    fpcmdDLQ.Enabled = True
  End If

End Sub

Private Sub fpcmbDLQNotice_DropDown()
  fpcmbDLQNotice.BackColor = &HFFFFFF
End Sub

Private Sub fpcmbDLQNotice_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  fpcmbDLQNotice.BackColor = &HFFFFFF
  If KeyCode = vbKeySpace Then
    fpcmbDLQNotice.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbDLQNotice.ListIndex = -1
  End If
  If fpcmbDLQNotice.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtTownName.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub LogSaves()
  Dim TownRec As TownSetUpType
  Dim THandle As Integer
  Dim AMeth$
  Dim YN$
  
  'this sub records all saves made from this screen in detail...saves
  'to arlog.dat
  On Error GoTo ERRORSTUFF
  
  OpenTownFile THandle
  Get THandle, 1, TownRec
  Close THandle
  
  If QPTrim$(TempTownName$) <> QPTrim$(TownRec.TownName) Then
    MainLog ("The Townname was changed from " + QPTrim$(TempTownName$) + " to " + QPTrim$(TownRec.TownName) + ".")
  End If
  If QPTrim$(TempTownContact$) <> QPTrim$(TownRec.Contact) Then
    MainLog ("The Town's contact was changed from " + QPTrim$(TempTownContact$) + " to " + QPTrim$(TownRec.Contact) + ".")
  End If
  If QPTrim$(TempTownAdd1$) <> QPTrim$(TownRec.TownAdd1) Then
    MainLog ("The Town's address #1 was changed from " + QPTrim$(TempTownAdd1$) + " to " + QPTrim$(TownRec.TownAdd1) + ".")
  End If
  If QPTrim$(TempTownAdd2$) <> QPTrim$(TownRec.TownAdd2) Then
    MainLog ("The Town's address #2 was changed from " + QPTrim$(TempTownAdd2$) + " to " + QPTrim$(TownRec.TownAdd2) + ".")
  End If
  If QPTrim$(TempTownCity$) <> QPTrim$(TownRec.City) Then
    MainLog ("The Town's city name was changed from " + QPTrim$(TempTownCity$) + " to " + QPTrim$(TownRec.City) + ".")
  End If
  If QPTrim$(TempTownState$) <> QPTrim$(TownRec.State) Then
    MainLog ("The Town's state name was changed from " + QPTrim$(TempTownState$) + " to " + QPTrim$(TownRec.State) + ".")
  End If
  If QPTrim$(TempTownZip$) <> ReplaceString(TownRec.ZipCode, "-", "") Then
    MainLog ("The Town's zip code was changed from " + QPTrim$(TempTownZip$) + " to " + QPTrim$(TownRec.ZipCode) + ".")
  End If
  If QPTrim$(TempTownPhone$) <> QPTrim$(TownRec.TownPhone) Then
    MainLog ("The Town's phone number was changed from " + QPTrim$(TempTownPhone$) + " to " + QPTrim$(TownRec.TownPhone) + ".")
  End If
  If TempTownAppForm <> TownRec.AppForm Then
    MainLog ("The Town's application form was changed from " + CStr(TempTownAppForm) + " to " + CStr(TownRec.AppForm) + ".")
  End If
  If TempTownPenForm <> TownRec.DLQNotice Then
    MainLog ("The Town's penalty notice was changed from " + CStr(TempTownPenForm) + " to " + CStr(TownRec.DLQNotice) + ".")
  End If
  If QPTrim$(TempLicNew) <> QPTrim$(TownRec.LicNumPermYN) Then
    MainLog ("The Town's permanent license # option was changed from " + QPTrim$(TempLicNew) + " to " + QPTrim$(TownRec.LicNumPermYN) + ".")
  End If
  If QPTrim$(TempGLToCats) <> QPTrim$(TownRec.GL2Cats) Then
    MainLog ("The Town's default category GL numbers to Penalty GL numbers option was changed from " + QPTrim$(TempGLToCats) + " to " + QPTrim$(TownRec.GL2Cats) + ".")
  End If
  If QPTrim$(TempAmtPct) <> QPTrim$(TownRec.UseAmtPctYN) Then
    MainLog ("The Town's penalty charge was changed from " + QPTrim$(TempAmtPct) + " to " + QPTrim$(TownRec.UseAmtPctYN) + ".")
  End If
  
  If Mid(TempAcctMethod$, 1, 1) <> QPTrim$(TownRec.AcctMeth) Then
    AMeth = TownRec.AcctMeth
    Select Case AMeth
      Case "N"
        AMeth = "None"
      Case "C"
        AMeth = "Cash"
      Case "A"
        AMeth = "Accrual"
    End Select
    MainLog ("The Town's accounting method was changed from " + QPTrim$(TempAcctMethod$) + " to " + AMeth + ".")
  End If
  
  If QPTrim$(TempPenRevNum$) <> GetGLNum(TownRec.PENREVGLNUM) Then
    MainLog ("The Town's penalty revenue number was changed from " + QPTrim$(TempPenRevNum$) + " to " + GetGLNum(TownRec.PENREVGLNUM) + ".")
  End If
  
  If QPTrim$(TempPenRecNum$) <> GetGLNum(TownRec.PENRECGLNUM) Then
    MainLog ("The Town's penalty accounts receivable number was changed from " + QPTrim$(TempPenRecNum$) + " to " + GetGLNum(TownRec.PENRECGLNUM) + ".")
  End If
  
  If QPTrim$(TempCashRecpt$) <> GetGLNum(TownRec.PENCASHACCT) Then
    MainLog ("The Town's penalty cash account number was changed from " + QPTrim$(TempCashRecpt$) + " to " + GetGLNum(TownRec.PENCASHACCT) + ".")
  End If
  
  If TempIssFee <> TownRec.IssFee Then
    MainLog ("The Town's issue fee amount was changed from " + QPTrim$(Using("$#,##0.00", TempIssFee)) + " to " + QPTrim$(Using("$#,##0.00", TownRec.IssFee)) + ".")
  End If
  
  YN$ = QPTrim$(TownRec.LaserLtr)
  If QPTrim$(TempLaserYN$) <> YN$ Then
    GoSub GetYesOrNo
    MainLog ("The Town's laser form option was changed from " + QPTrim$(TempLaserYN$) + " to " + QPTrim$(YN$) + ".")
  End If
  
  Exit Sub
  
GetYesOrNo:
  Select Case YN$
    Case "Y"
      YN$ = "Yes"
    Case "N"
      YN$ = "No"
  End Select
  Return
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLTownSetup", "LogSaves", Erl)
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

Private Sub fpcmbGLToCats_Change()
  If QPTrim$(fpcmbGLToCats.Text) = "" Then
    fpcmbGLToCats.Text = "No"
  End If
End Sub

Private Sub fpcmbGLToCats_DropDown()
  fpcmbGLToCats.BackColor = &HFFFFFF
End Sub

Private Sub fpcmbGLToCats_KeyDown(KeyCode As Integer, Shift As Integer)
  fpcmbGLToCats.BackColor = &HFFFFFF
  If KeyCode = vbKeySpace Then
    fpcmbGLToCats.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbGLToCats.ListIndex = -1
  End If
  If fpcmbGLToCats.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbNewLicNumYN.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbLaserYN_Change()
  If fpcmbLaserYN.Text = "" Then fpcmbLaserYN.Text = "0"
  If fpcmbLaserYN.Text = "1" Then
    cmdAdvanceLtr.Text = "F2 Go To &Letter 1"
    cmdAdvanceLtr.Enabled = True
  ElseIf fpcmbLaserYN.Text = "2" Then
    cmdAdvanceLtr.Text = "F2 Go To &Letter 2"
    cmdAdvanceLtr.Enabled = True
  ElseIf fpcmbLaserYN.Text = "3" Then
    cmdAdvanceLtr.Text = "F2 Go To &Letter 3"
    cmdAdvanceLtr.Enabled = True
  Else
    cmdAdvanceLtr.Text = "No Selection"
    cmdAdvanceLtr.Enabled = False
  End If
End Sub

Private Sub fpcmbLaserYN_DropDown()
  fpcmbLaserYN.BackColor = &HFFFFFF
End Sub

Private Sub fpcmbLaserYN_KeyDown(KeyCode As Integer, Shift As Integer)
  fpcmbLaserYN.BackColor = &HFFFFFF
  If KeyCode = vbKeySpace Then
    fpcmbLaserYN.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbLaserYN.ListIndex = -1
  End If
  If fpcmbLaserYN.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbAppType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbNewLicNumYN_Change()
  If QPTrim$(fpcmbNewLicNumYN.Text) = "" Then
    fpcmbNewLicNumYN.Text = "No"
  End If
End Sub

Private Sub fpcmbNewLicNumYN_DropDown()
  fpcmbNewLicNumYN.BackColor = &HFFFFFF
End Sub

Private Sub fpcmbNewLicNumYN_KeyDown(KeyCode As Integer, Shift As Integer)
  fpcmbNewLicNumYN.BackColor = &HFFFFFF
  If KeyCode = vbKeySpace Then
    fpcmbNewLicNumYN.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbNewLicNumYN.ListIndex = -1
  End If
  If fpcmbNewLicNumYN.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbAmtPct.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbNewLicNumYN_LostFocus()
  If QPTrim$(fpcmbNewLicNumYN.Text) = "" Then
    fpcmbNewLicNumYN.Text = "No"
  End If
End Sub

Private Sub fpcmdApps_Click()
  'this sub directs the program to the correct application
  'notice depending on what the caption is in the command button
  
  On Error Resume Next
  
'  If Not Exist("artownsu.dat") Then
'    frmBLMessageBoxJr.Label1.Caption = "Since no town setup data has been saved yet please complete the accounting method data and all the other required fields and save them now before continuing."
'    frmBLMessageBoxJr.Label1.Top = 600
'    frmBLMessageBoxJr.Show vbModal
'    fpcmbAcctMethod.SetFocus
'    Exit Sub
'  End If
      
  If QPTrim$(fpcmbAppType.Text) = "1. APP STANDARD" Then
    frmBLAppTemplate1.Show
  End If
  If QPTrim$(fpcmbAppType.Text) = "2. APP FORM A" Then
    frmBLAppTemplate2.Show
  End If
  If QPTrim$(fpcmbAppType.Text) = "3. APP FORM B" Then
    frmBLAppTemplate3.Show
  End If
  If QPTrim$(fpcmbAppType.Text) = "4. APP FORM C" Then
    frmBLAppTemplate4.Show
  End If
  If QPTrim$(fpcmbAppType.Text) = "5. APP FORM D" Then
    frmBLAppTemplate5.Show
  End If
  If QPTrim$(fpcmbAppType.Text) = "6. APP FORM E" Then
    frmBLAppTemplate6.Show
  End If
  If QPTrim$(fpcmbAppType.Text) = "7. APP FORM F" Then
    frmBLAppTemplate7.Show
  End If
  If QPTrim$(fpcmbAppType.Text) = "8. APP FORM G" Then
    frmBLAppTemplate8.Show
  End If
  If QPTrim$(fpcmbAppType.Text) = "9. APP FORM H" Then
    frmBLAppTemplate9.Show
  End If
  If QPTrim$(fpcmbAppType.Text) = "10. APP FREE FORMAT" Then
    frmBLFreeFormatApp1.Show
  End If
End Sub

Private Sub fpcmdDLQ_Click()
'  If Not Exist("artownsu.dat") Then
'    frmBLMessageBoxJr.Label1.Caption = "Since no town setup data has been saved yet please complete the accounting method data and all the other required fields and save them now before continuing."
'    frmBLMessageBoxJr.Label1.Top = 600
'    frmBLMessageBoxJr.Show vbModal
'    fpcmbAcctMethod.SetFocus
'    Exit Sub
'  End If
  
  If QPTrim$(fpcmbDLQNotice.Text) = "1. PENALTY STANDARD" Then
    Load frmBLDlqnTemplate1
    frmBLDlqnTemplate1.Show
    Me.Hide
  End If
  If QPTrim$(fpcmbDLQNotice.Text) = "2. PENALTY FORM A" Then
    Load frmBLDlqnTemplate2a
    frmBLDlqnTemplate2a.Show
    Me.Hide
  End If
  If QPTrim$(fpcmbDLQNotice.Text) = "3. PENALTY FORM B" Then
    Load frmBLDlqnTemplate3
    frmBLDlqnTemplate3.Show
    Me.Hide
  End If
  If QPTrim$(fpcmbDLQNotice.Text) = "4. FREE FORMAT" Then
    Load frmBLFreeFormatDlnq
    frmBLFreeFormatDlnq.Show
    Me.Hide
  End If
End Sub

Private Sub fpcurrIssFee_Click(Button As Integer)
  fpcurrIssFee.BackColor = &HFFFFFF
End Sub

Private Sub fpcurrIssFee_KeyDown(KeyCode As Integer, Shift As Integer)
  fpcurrIssFee.BackColor = &HFFFFFF
  If KeyCode = vbKeyDown Then
    fpcmbGLToCats.SetFocus
  End If
  If KeyCode = vbKeyUp Then
    If fptxtCashReceipt.Enabled = True Then
      fptxtCashReceipt.SetFocus
    ElseIf fpcmbAcctMethod.Enabled = True Then
      fpcmbAcctMethod.SetFocus
    Else
      fptxtPhone.SetFocus
    End If
  End If
End Sub

Private Sub fptxtAcctsRec_Change()
  If fptxtAcctsRec.Text <> "" Then
    LastPenRecNum = QPTrim$(fptxtAcctsRec.Text)
  End If
End Sub

Private Sub fptxtAcctsRec_Click(Button As Integer)
  fptxtAcctsRec.BackColor = &HFFFFFF
End Sub

Private Sub fptxtAcctsRec_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtAcctsRec.BackColor = &HFFFFFF
End Sub

Private Sub fptxtAdd1_Click(Button As Integer)
  fptxtAdd1.BackColor = &HFFFFFF
End Sub

Private Sub fptxtAdd1_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtAdd1.BackColor = &HFFFFFF
End Sub

Private Sub fptxtAdd2_Click(Button As Integer)
  fptxtAdd2.BackColor = &HFFFFFF
End Sub

Private Sub fptxtAdd2_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtAdd2.BackColor = &HFFFFFF
End Sub

Private Sub fptxtCashReceipt_Change()
  If fptxtCashReceipt.Text <> "" Then
    LastCashRecpt = QPTrim$(fptxtCashReceipt.Text)
  End If
End Sub

Private Sub fptxtCashReceipt_Click(Button As Integer)
  fptxtCashReceipt.BackColor = &HFFFFFF
End Sub

Private Sub fptxtCashReceipt_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtCashReceipt.BackColor = &HFFFFFF
End Sub

Private Sub fptxtCity_Click(Button As Integer)
  fptxtCity.BackColor = &HFFFFFF
End Sub

Private Sub fptxtCity_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtCity.BackColor = &HFFFFFF
End Sub

Private Sub fptxtContact_Click(Button As Integer)
  fptxtContact.BackColor = &HFFFFFF
End Sub

Private Sub fptxtContact_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtContact.BackColor = &HFFFFFF
End Sub

Private Sub fptxtPhone_Click(Button As Integer)
  fptxtPhone.BackColor = &HFFFFFF
End Sub

Private Sub fptxtPhone_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtPhone.BackColor = &HFFFFFF
End Sub

Private Sub fptxtRevGLAcctNum_Click(Button As Integer)
  fptxtRevGLAcctNum.BackColor = &HFFFFFF
End Sub

Private Sub fptxtRevGLAcctNum_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtRevGLAcctNum.BackColor = &HFFFFFF
End Sub

Private Sub fptxtState_Click(Button As Integer)
  fptxtState.BackColor = &HFFFFFF
End Sub

Private Sub fptxtState_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtState.BackColor = &HFFFFFF
End Sub

Private Sub fptxtTownName_Click(Button As Integer)
  fptxtTownName.BackColor = &HFFFFFF
End Sub

Private Sub fptxtTownName_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtTownName.BackColor = &HFFFFFF
  If KeyCode = vbKeyDown Then
    fptxtContact.SetFocus
  End If
  
  If KeyCode = vbKeyUp Then
    fpcmbDLQNotice.SetFocus
  End If
End Sub

Private Function GLNumsOK() As Boolean
  'looks to make sure the appropriate GL numbers are
  'populated based on the type of Accounting Method
  'selected in the Town Setup
  
  On Error GoTo ERRORSTUFF
  
  GLNumsOK = True
  
'  If fpcmbVerifyYN.Text = "Omit" Then Exit Function
  
  Select Case Mid(fpcmbAcctMethod.Text, 1, 1)
    Case "A"
      If fptxtRevGLAcctNum.Text = "" Then
        fptxtRevGLAcctNum.BackColor = &H80FFFF
        frmBLMessageBoxJr.Label1.Caption = "The 'Accrual' accounting method requires the 'Revenue G/L Account Number' field to be filled in."
        frmBLMessageBoxJr.Label1.Top = 700
        frmBLMessageBoxJr.Show vbModal
        fptxtRevGLAcctNum.BackColor = &HFFFFFF
        fptxtRevGLAcctNum.SetFocus
        GLNumsOK = False
      ElseIf fptxtAcctsRec.Text = "" Then
        fptxtAcctsRec.BackColor = &H80FFFF
        frmBLMessageBoxJr.Label1.Caption = "The 'Accrual' accounting method requires the 'Accounts Receivable Number' field to be filled in."
        frmBLMessageBoxJr.Label1.Top = 700
        frmBLMessageBoxJr.Show vbModal
        fptxtAcctsRec.BackColor = &HFFFFFF
        fptxtAcctsRec.SetFocus
        GLNumsOK = False
      ElseIf fptxtCashReceipt.Text = "" Then
        fptxtCashReceipt.BackColor = &H80FFFF
        frmBLMessageBoxJr.Label1.Caption = "The 'Accrual' accounting method requires the 'Cash Receipt G/L Account Number' field to be filled in."
        frmBLMessageBoxJr.Label1.Top = 700
        frmBLMessageBoxJr.Show vbModal
        fptxtCashReceipt.BackColor = &HFFFFFF
        fptxtCashReceipt.SetFocus
        GLNumsOK = False
      End If
    Case "C"
      If fptxtRevGLAcctNum.Text = "" Then
        fptxtRevGLAcctNum.BackColor = &H80FFFF
        frmBLMessageBoxJr.Label1.Caption = "The 'Cash' accounting method requires the 'Revenue G/L Account Number' field to be filled in."
        frmBLMessageBoxJr.Label1.Top = 700
        frmBLMessageBoxJr.Show vbModal
        fptxtRevGLAcctNum.BackColor = &HFFFFFF
        fptxtRevGLAcctNum.SetFocus
        GLNumsOK = False
      ElseIf fptxtCashReceipt.Text = "" Then
        fptxtCashReceipt.BackColor = &H80FFFF
        frmBLMessageBoxJr.Label1.Caption = "The 'Cash' accounting method requires the 'Cash Receipt G/L Account Number' field to be filled in."
        frmBLMessageBoxJr.Label1.Top = 700
        frmBLMessageBoxJr.Show vbModal
        fptxtCashReceipt.BackColor = &HFFFFFF
        fptxtCashReceipt.SetFocus
        GLNumsOK = False
      ElseIf fptxtAcctsRec.Text <> "" Then
        fptxtAcctsRec.BackColor = &H80FFFF
        frmBLMessageBoxJrWOpts.Label1.Caption = "The 'Cash' accounting method requires the 'Revenue G/L Account Number' field to be empty. Continuing will erase this number. OK to continue?"
        frmBLMessageBoxJrWOpts.Show vbModal
        If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
          Unload frmBLMessageBoxJrWOpts
          fptxtAcctsRec.BackColor = &HFFFFFF
          fptxtAcctsRec.SetFocus
          GLNumsOK = False
        Else
          Unload frmBLMessageBoxJrWOpts
          fptxtAcctsRec.BackColor = &HFFFFFF
          fptxtAcctsRec.Text = ""
        End If
      End If
    Case "N"
      If fptxtRevGLAcctNum <> "" Then
        fptxtRevGLAcctNum.BackColor = &H80FFFF
        frmBLMessageBoxJrWOpts.Label1.Caption = "The 'None' accounting method requires the 'Revenue G/L Account Number' field to be empty. Continuing will erase this number. OK to continue?"
        frmBLMessageBoxJrWOpts.Show vbModal
        If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
          Unload frmBLMessageBoxJrWOpts
          fptxtRevGLAcctNum.BackColor = &HFFFFFF
          fptxtRevGLAcctNum.SetFocus
          GLNumsOK = False
        Else
          Unload frmBLMessageBoxJrWOpts
          fptxtRevGLAcctNum.BackColor = &HFFFFFF
          fptxtRevGLAcctNum.Text = ""
        End If
      ElseIf fptxtAcctsRec.Text <> "" Then
        fptxtAcctsRec.BackColor = &H80FFFF
        frmBLMessageBoxJrWOpts.Label1.Caption = "The 'None' accounting method requires the 'Accounts Receivable Number' field to be empty. Continuing will erase this number. OK to continue?"
        frmBLMessageBoxJrWOpts.Show vbModal
        If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
          Unload frmBLMessageBoxJrWOpts
          fptxtAcctsRec.BackColor = &HFFFFFF
          fptxtAcctsRec.SetFocus
          GLNumsOK = False
        Else
          Unload frmBLMessageBoxJrWOpts
          fptxtAcctsRec.BackColor = &HFFFFFF
          fptxtAcctsRec.Text = ""
        End If
      ElseIf fptxtCashReceipt.Text <> "" Then
        fptxtCashReceipt.BackColor = &H80FFFF
        fptxtAcctsRec.BackColor = &H80FFFF
        frmBLMessageBoxJrWOpts.Label1.Caption = "The 'None' accounting method requires the 'Cash Receipt G/L Account Number' field to be empty. Continuing will erase this number. OK to continue?"
        frmBLMessageBoxJrWOpts.Show vbModal
        If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
          Unload frmBLMessageBoxJrWOpts
          fptxtCashReceipt.BackColor = &HFFFFFF
          fptxtCashReceipt.SetFocus
          GLNumsOK = False
        Else
          Unload frmBLMessageBoxJrWOpts
          fptxtCashReceipt.BackColor = &HFFFFFF
          fptxtCashReceipt.Text = ""
        End If
      End If
    End Select
    
    Exit Function
    
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLTownSetup", "GLNumsOK", Erl)
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

Private Sub fptxtRevGLAcctNum_Change()
  If fptxtRevGLAcctNum.Text <> "" Then
    LastPenRecNum$ = QPTrim$(fptxtRevGLAcctNum.Text)
  End If
End Sub

Private Function GLNumsValid() As Boolean
  Dim GLIdxRec As JGLAcctIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdxRecs As Integer
  Dim x As Integer
  Dim GLAcctRec As GLAcctRecType
  Dim AcctHandle As Integer
  Dim RevNum$, Rev As Integer
  Dim AcctsRecNum$, Acct As Integer
  Dim CashRecNum$, Cash As Integer
  Dim ThisGLNum$
  
  On Error GoTo ERRORSTUFF
  
  'looks to make sure that the GL numbers entered match
  'a GL number available on the GL list
  
  GLNumsValid = True
  
'  If fpcmbVerifyYN.Text = "Omit" Then Exit Function
  If fpcmbAcctMethod.Text <> "None" Then
    If Not Exist("GLACCT.IDX") Or Not Exist("GLACCT.DAT") Then
  '    If fpcmbVerifyYN.Text = "Yes" Then
        If Not Exist("GLACCT.IDX") And Not Exist("GLACCT.DAT") Then
          frmBLMessageBoxJrWOpts.Label1.Caption = "The files 'GLACCT.IDX' and 'GLACCT.DAT' could not be found in the Citipak directory. These files are needed to verify the validity of General Ledger numbers. Numbers must be verified to be saved properly. Press F10 to continue saving anyway or ESC to abort."
          frmBLMessageBoxJrWOpts.Label1.Top = 600
          frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
          frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Abort"
          frmBLMessageBoxJrWOpts.Show vbModal
          If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
            Unload frmBLMessageBoxJrWOpts
            GLNumsValid = False
            Exit Function
          Else
            MainLog ("User warned that 'GLACCT.DAT' and 'GLACCT.IDX' were missing from the Citipak directory and the user elected to continue saving the Town Setup data even though GL numbers could not be verified.")
            Unload frmBLMessageBoxJrWOpts
            GoTo GLCheckCompleted
          End If
        End If
        If Not Exist("GLACCT.IDX") Then
          frmBLMessageBoxJrWOpts.Label1.Caption = "The file 'GLACCT.IDX' could not be found in the Citipak directory. This file is required to verify the validity of General Ledger numbers. Numbers must be verified to be saved properly. Press F10 to continue saving anyway or ESC to abort."
          frmBLMessageBoxJrWOpts.Label1.Top = 600
          frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
          frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Abort"
          frmBLMessageBoxJrWOpts.Show vbModal
          If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
            Unload frmBLMessageBoxJrWOpts
            GLNumsValid = False
            Exit Function
          Else
            MainLog ("User warned that 'GLACCT.IDX' was missing from the Citipak directory and the user elected to continue saving the Town Setup data even though GL numbers could not be verified.")
            Unload frmBLMessageBoxJrWOpts
            GoTo GLCheckCompleted
          End If
        End If
        If Not Exist("GLACCT.DAT") Then
          frmBLMessageBoxJrWOpts.Label1.Caption = "The file 'GLACCT.DAT' could not be found in the Citipak directory. This file is required to verify the validity of General Ledger numbers. Numbers must be verified to be saved properly. Press F10 to continue saving anyway or ESC to abort."
          frmBLMessageBoxJrWOpts.Label1.Top = 600
          frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
          frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Abort"
          frmBLMessageBoxJrWOpts.Show vbModal
          If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
            Unload frmBLMessageBoxJrWOpts
            GLNumsValid = False
            Exit Function
          Else
            MainLog ("User warned that 'GLACCT.DAT' was missing from the Citipak directory and the user elected to continue saving the Town Setup data even though GL numbers could not be verified.")
            Unload frmBLMessageBoxJrWOpts
            GoTo GLCheckCompleted
          End If
        End If
  '    End If
      Exit Function
    End If
  End If
  
GLCheckCompleted:
  OpenGLIdxFile IdxHandle
  NumOfIdxRecs = LOF(IdxHandle) / Len(GLIdxRec)
  If NumOfIdxRecs = 0 Then
'    frmBLMessageBoxJr.Label1.Caption = "There are no General Ledger numbers indexed."
'    frmBLMessageBoxJr.Label1.Top = 900
'    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Function
  End If
  ReDim IdxRec(1 To NumOfIdxRecs) As Integer
  For x = 1 To NumOfIdxRecs
    Get IdxHandle, x, GLIdxRec
    IdxRec(x) = GLIdxRec.RecNo
  Next x
  Close IdxHandle
  
  RevNum$ = QPTrim$(fptxtRevGLAcctNum.Text)
  If RevNum <> "" Then
    Rev = 0
  Else
    Rev = 1
  End If
  
  AcctsRecNum$ = QPTrim$(fptxtAcctsRec.Text)
  If AcctsRecNum$ <> "" Then
    Acct = 0
  Else
    Acct = 1
  End If
  
  CashRecNum$ = QPTrim$(fptxtCashReceipt.Text)
  If CashRecNum$ <> "" Then
    Cash = 0
  Else
    Cash = 1
  End If
  
  OpenGLAcctFile AcctHandle
  For x = 1 To NumOfIdxRecs
    Get AcctHandle, IdxRec(x), GLAcctRec
      If GLAcctRec.Deleted Then GoTo ItsDeleted
      ThisGLNum = QPTrim$(GLAcctRec.Num)
      If Rev = 1 Then GoTo RevIs1
      If ThisGLNum$ = RevNum$ Then
        Rev = 1
      End If
RevIs1:
      If Acct = 1 Then GoTo AcctIs1
      If ThisGLNum$ = AcctsRecNum$ Then
        Acct = 1
      End If
AcctIs1:
      If Cash = 1 Then GoTo CashIs1
      If ThisGLNum$ = CashRecNum$ Then
        Cash = 1
      End If
CashIs1:
      If Rev = 1 And Acct = 1 And Cash = 1 Then
        Exit For
      End If
ItsDeleted:
  Next x
  
  If x > NumOfIdxRecs Then
    If Rev = 0 Then
      fptxtRevGLAcctNum.BackColor = &H80FFFF
      frmBLMessageBoxJr.Label1.Caption = "The GL Number entered for 'Revenue G/L Account Number' does not match any GL Numbers on file."
      frmBLMessageBoxJr.Label1.Top = 800
      frmBLMessageBoxJr.Show vbModal
      fptxtRevGLAcctNum.BackColor = &HFFFFFF
      fptxtRevGLAcctNum.SetFocus
      GLNumsValid = False
      Exit Function
    ElseIf Acct = 0 Then
      fptxtAcctsRec.BackColor = &H80FFFF
      frmBLMessageBoxJr.Label1.Caption = "The GL Number entered for 'Accounts Receivable Number' does not match any GL Numbers on file."
      frmBLMessageBoxJr.Label1.Top = 800
      frmBLMessageBoxJr.Show vbModal
      fptxtAcctsRec.BackColor = &HFFFFFF
      fptxtAcctsRec.SetFocus
      GLNumsValid = False
      Exit Function
    ElseIf Cash = 0 Then
      fptxtCashReceipt.BackColor = &H80FFFF
      frmBLMessageBoxJr.Label1.Caption = "The GL Number entered for 'Cash Receipt G/L Account Number' does not match any GL Numbers on file."
      frmBLMessageBoxJr.Label1.Top = 800
      frmBLMessageBoxJr.Show vbModal
      fptxtCashReceipt.BackColor = &HFFFFFF
      fptxtCashReceipt.SetFocus
      GLNumsValid = False
      Exit Function
    End If
  End If
  
  Exit Function
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLTownSetup", "GLNumsValid", Erl)
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

Private Sub MakeBackColorsWhite()
  Dim cnt As Integer
  Dim x As Control
  Dim cmdButton As CommandButton
  
  'code resets all field's backcolors to white
  On Error Resume Next
  For cnt = 0 To Me.Count - 1
    Set x = Me.Controls.Item(cnt)
      If TypeOf x Is fpText Or TypeOf x Is fpCombo Then
        x.BackColor = &H80000005
      End If
  Next cnt
  EnableCloseButton Me.hwnd, False
  
End Sub

Private Sub fptxtZip_Click(Button As Integer)
  fptxtZip.BackColor = &HFFFFFF
End Sub

Private Sub fptxtZip_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtZip.BackColor = &HFFFFFF
End Sub
