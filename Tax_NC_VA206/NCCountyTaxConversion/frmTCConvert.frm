VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmTCConvert 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Tax Records"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTCConvert.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbChrgInt 
      Height          =   405
      Left            =   2760
      TabIndex        =   26
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
      ColDesigner     =   "frmTCConvert.frx":08CA
   End
   Begin LpLib.fpCombo fpcmbPLateList 
      Height          =   405
      Left            =   10080
      TabIndex        =   8
      Top             =   2760
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
      ColDesigner     =   "frmTCConvert.frx":0B89
   End
   Begin LpLib.fpCombo fpcmbRLateList 
      Height          =   405
      Left            =   6240
      TabIndex        =   6
      Top             =   2760
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
      ColDesigner     =   "frmTCConvert.frx":0E48
   End
   Begin LpLib.fpCombo fpcmbActive 
      Height          =   405
      Left            =   2760
      TabIndex        =   4
      Top             =   4920
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
      ColDesigner     =   "frmTCConvert.frx":1107
   End
   Begin LpLib.fpCombo fpcmbBankrupt 
      Height          =   405
      Left            =   2760
      TabIndex        =   3
      Top             =   4320
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
      ColDesigner     =   "frmTCConvert.frx":13C6
   End
   Begin LpLib.fpCombo fpcmbLateList 
      Height          =   405
      Left            =   2760
      TabIndex        =   2
      Top             =   3720
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
      ColDesigner     =   "frmTCConvert.frx":1685
   End
   Begin LpLib.fpCombo fpcmbTaxEx 
      Height          =   405
      Left            =   2760
      TabIndex        =   1
      Top             =   3120
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
      ColDesigner     =   "frmTCConvert.frx":1944
   End
   Begin LpLib.fpCombo fpcmbPenalty 
      Height          =   405
      Left            =   2760
      TabIndex        =   0
      Top             =   2520
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
      ColDesigner     =   "frmTCConvert.frx":1C03
   End
   Begin VB.CheckBox chkMultipleR 
      BackColor       =   &H008F8265&
      Caption         =   "Use Multiple Real Property"
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   7920
      TabIndex        =   29
      Top             =   6120
      Width           =   3255
   End
   Begin VB.CheckBox chkSingleR 
      BackColor       =   &H008F8265&
      Caption         =   "Use Single Real Property"
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   4560
      TabIndex        =   28
      Top             =   6120
      Width           =   3135
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   636
      Left            =   3216
      TabIndex        =   10
      TabStop         =   0   'False
      Tag             =   "Press the 'Cancel' button to exit this screen and return to the main 'Business License Reports' menu."
      Top             =   7440
      Width           =   1740
      _Version        =   131072
      _ExtentX        =   3069
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
      ButtonDesigner  =   "frmTCConvert.frx":1EC2
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdConvert 
      Height          =   636
      Left            =   6696
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   7440
      Width           =   1740
      _Version        =   131072
      _ExtentX        =   3069
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
      ButtonDesigner  =   "frmTCConvert.frx":20A0
   End
   Begin EditLib.fpDateTime fptxtOpenDate 
      Height          =   372
      Left            =   1920
      TabIndex        =   5
      Top             =   6360
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
   Begin EditLib.fpDateTime fptxtRPropDate 
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   4200
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
      Text            =   "02/22/2006"
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
   Begin EditLib.fpDateTime fptxtPPropDate 
      Height          =   375
      Left            =   8760
      TabIndex        =   9
      Top             =   4200
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
   Begin VB.Label Label16 
      BackColor       =   &H0080FFFF&
      Caption         =   "Conversion Type"
      Height          =   375
      Left            =   4200
      TabIndex        =   30
      Top             =   5520
      Width           =   1935
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   1575
      Left            =   4200
      Top             =   5520
      Width           =   7095
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Charge Interest?:"
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   480
      TabIndex        =   27
      Top             =   5640
      Width           =   2172
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   3255
      Left            =   7920
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "New Pers Property Settings"
      Height          =   372
      Left            =   7920
      TabIndex        =   25
      Top             =   1920
      Width           =   3132
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pers Prop Date:"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   8640
      TabIndex        =   24
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Late List Y/N?:"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   8160
      TabIndex        =   23
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Late List Y/N?:"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4320
      TabIndex        =   22
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Real Prop Date:"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   4800
      TabIndex        =   21
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "New Real Property Settings"
      Height          =   372
      Left            =   4200
      TabIndex        =   20
      Top             =   1920
      Width           =   3132
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   3255
      Left            =   4200
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Active Y/N?:"
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   480
      TabIndex        =   19
      Top             =   5040
      Width           =   2172
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bankrupt Y/N?:"
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   480
      TabIndex        =   18
      Top             =   4440
      Width           =   2172
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Open Date:"
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   480
      TabIndex        =   17
      Top             =   6480
      Width           =   1332
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Late List Y/N?:"
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   480
      TabIndex        =   16
      Top             =   3840
      Width           =   2172
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Exempt Y/N?:"
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   480
      TabIndex        =   15
      Top             =   3240
      Width           =   2172
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "New Customer Settings"
      Height          =   375
      Left            =   420
      TabIndex        =   14
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   5172
      Left            =   420
      Top             =   1920
      Width           =   3372
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Penalty Y/N?:"
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   840
      TabIndex        =   12
      Top             =   2640
      Width           =   1812
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "County Tax Data Conversion"
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
      Left            =   3156
      TabIndex        =   13
      Top             =   636
      Width           =   5292
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   660
      Index           =   1
      Left            =   1500
      Top             =   468
      Width           =   8652
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   1500
      Top             =   360
      Width           =   8652
   End
End
Attribute VB_Name = "frmTCConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim TRealVal As Double
  Dim TPersVal As Double
  Dim TRealX As Double
  Dim TPersX As Double
  Dim TCnt As Long
  Dim FoundRealCnt As Long
  Dim FoundReal() As String
  Dim FoundRealPin() As String
  
Private Sub cmdConvert_Click()
  If chkSingleR.Value = 0 And chkMultipleR.Value = 0 Then
    Call TCMsg(900, "Please select Single or Multiple conversion type.")
    Exit Sub
  End If
  If chkSingleR.Value = 1 Then
    Call ProcessRSingle
  ElseIf chkMultipleR.Value = 1 Then
    Call ProcessRMultiple
  End If
End Sub

Private Sub cmdExit_Click()
  frmTCMainMenuNew.Show
  DoEvents
  Unload Me
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
      SendKeys "%v"
      Call cmdConvert_Click
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
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTCConvert.")
      End
    End If
  End If

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

Private Sub LoadMe()
  
  fpcmbPenalty.Text = "Y"
  fpcmbPenalty.AddItem "N"
  fpcmbPenalty.AddItem "Y"
  fpcmbTaxEx.Text = "N"
  fpcmbTaxEx.AddItem "N"
  fpcmbTaxEx.AddItem "Y"
  fpcmbLateList.Text = "Y"
  fpcmbLateList.AddItem "N"
  fpcmbLateList.AddItem "Y"
  fpcmbBankrupt.Text = "N"
  fpcmbBankrupt.AddItem "N"
  fpcmbBankrupt.AddItem "Y"
  fpcmbActive.Text = "Y"
  fpcmbActive.AddItem "N"
  fpcmbActive.AddItem "Y"
  
  fpcmbChrgInt.Text = "Y"
  fpcmbChrgInt.AddItem "N"
  fpcmbChrgInt.AddItem "Y"
  
  fptxtOpenDate.Text = Date
  
  fpcmbRLateList.Text = "Y"
  fpcmbRLateList.AddItem "Y"
  fpcmbRLateList.AddItem "N"
  fptxtRPropDate.Text = Date
  
  fpcmbPLateList.Text = "Y"
  fpcmbPLateList.AddItem "Y"
  fpcmbPLateList.AddItem "N"
  fptxtPPropDate.Text = Date
  
End Sub

Private Sub ProcessRSingle()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim OldRealRec As PropertyRecType
  Dim OldRHandle As Integer
  Dim NumOfRealRecs As Long
  Dim NumOfNewRealRecs As Long
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim OldPersRec As PersonalRecType
  Dim OldPHandle As Integer
  Dim NumOfPersRecs As Long
  Dim NumOfNewPersRecs As Long
  Dim x As Long, y As Long, z As Long
  Dim TempHandle As Integer
  Dim TempRec As TempConversionData
  Dim NumOfTempRecs As Long
  Dim NewSCustCnt As Long
  Dim NewOldCustCnt As Long
  Dim NextRRec As Long
  Dim NextPRec As Long
  Dim OldRPin As String
  Dim TempRPin As String
  Dim TempPPin As String
  Dim IntPinRec As InternalPinType
  Dim IHandle As Integer
  Dim NumOfIRecs As Long
  Dim IntRecCnt As Long
  Dim PersValue As Double
  Dim AddTotal As Double
  Dim ConvRec As ConvResultsType
  Dim CRHandle As Integer
  Dim NumOfCRRecs As Long
  Dim NumOfErrors As Long
  Dim ErrorCnt As Long
  Dim ErrorRec As ConvErrorType
  Dim EHandle As Integer
  Dim NumOfERecs As Long
  Dim RealValue As Double
  Dim ErrorCode As Integer
  Dim Updated As Boolean
  Dim ThisPPin$
  Dim ThisRPin$
  Dim CurrRealProc As Boolean
  Dim FoundCnt As Long
  Dim FoundIt As Boolean
  Dim ThisCountyString$
  Dim CustCountyString$
  Dim SaveRRec As Long
  Dim CoNumCnt As Long
  
'  On Error GoTo ERRORSTUFF
  
  If Not Exist(ConversionFile) Then
    Call TCMsg(900, "Please process the county data first. Load attempt aborted.")
    Exit Sub
  End If
  
  If TCMsgWOpts(700, "WARNING: Continuing will delete all real property and personal property files so they can be rebuilt with the new data from the county. Press F10 to continue. Otherwise, press ESC to abort.", "F10 Convert", "ESC Abort") = "abort" Then
    Exit Sub
  End If
  
  ReDim NewSCust(1 To 1) As Long 'String 'county number is a string
  ReDim CoNum(1 To 1) As String
  
  If Exist("OLDTAXPROP.DAT") Then
    If TCMsgWOpts(900, "'OLDTAXPROP.DAT' already exists. Do you wish to overwrite?", "F10 Overwrite", "ESC Leave As Is") <> "abort" Then
      KillFile "OLDTAXPROP.DAT"
      Name "TAXPROP.DAT" As "OLDTAXPROP.DAT"
    End If
  Else
    Name "TAXPROP.DAT" As "OLDTAXPROP.DAT"
  End If
  
  If Exist("OLDTAXPERS.DAT") Then
    If TCMsgWOpts(900, "'OLDTAXPERS.DAT' already exists. Do you wish to overwrite?", "F10 Overwrite", "ESC Leave As Is") <> "abort" Then
      KillFile "OLDTAXPERS.DAT"
      Name "TAXPERS.DAT" As "OLDTAXPERS.DAT"
    End If
  Else
    Name "TAXPERS.DAT" As "OLDTAXPERS.DAT"
  End If
  
'  KillFile "TAXPERS.DAT"
'  KillFile "TAXPROP.DAT"
'
  OpenOldRealPropFile OldRHandle, NumOfRealRecs
  Get OldRHandle, 1, OldRealRec
  OpenOldPersPropFile OldPHandle, NumOfPersRecs
  Get OldPHandle, 1, OldPersRec
  
  OpenTempConvFile TempHandle, NumOfTempRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
'    If TaxCust.Cycle > 0 Then GoTo SkipThis 'for Maxton
    TaxCust.FirstPersRec = 0
    TaxCust.FirstPropRec = 0
    Put TCHandle, x, TaxCust
SkipThis:
  Next x
  
  OpenRealPropFile RHandle, NumOfNewRealRecs
  OpenPersPropFile PHandle, NumOfNewPersRecs
  OpenIntPinFile IHandle, NumOfIRecs
  
  IntRecCnt = NumOfIRecs
  TCnt = 0
  KillFile "CNVRSLTS.DAT"
  KillFile "CNVRERRS.DAT"
  
  OpenConvErrorsFile EHandle, NumOfERecs
  OpenConvResultsFile CRHandle, NumOfCRRecs
  
  frmTCShowPctComp.Label1 = "Converting County Data"
  frmTCShowPctComp.Show , Me
  
  For x = 1 To NumOfTempRecs
    Get TempHandle, x, TempRec
    If TempRec.CData.CountyAcct > 0 Then
      ThisCountyString = CStr(TempRec.CData.CountyAcct)
    ElseIf QPTrim$(TempRec.CData.CountyAcctString) <> "" Then
      ThisCountyString = CStr(TempRec.CData.CountyAcctString)
    Else
      ErrorCode = 5
      GoSub SaveErrors
      GoTo CustErrors
    End If
    PersValue = OldRound(TempRec.CData.CVALUE + TempRec.CData.MCVALUE + TempRec.CData.MHVALUE + TempRec.CData.PersVal + TempRec.CData.MTVALUE)
    If PersValue = 0 Then GoTo CustErrors
    RealValue = TempRec.CData.PROPVALU
    For y = 1 To NumOfTCRecs
      Get TCHandle, y, TaxCust
      If TaxCust.CountyAcct > 0 Then
        CustCountyString = CStr(TaxCust.CountyAcct)
      ElseIf QPTrim$(TaxCust.CountyAcctString) <> "" Then
        CustCountyString = QPTrim$(TaxCust.CountyAcctString)
      End If
      Updated = False
      If ThisCountyString <> "" Then 'make sure a valid county number exists
        If IsNumeric(ThisCountyString) = True And IsNumeric(CustCountyString) = True Then 'should take car of leading
        'zeros in one number but not in the other
          If Val(ThisCountyString) = Val(CustCountyString) Then GoTo ApprovedP
        ElseIf CustCountyString = QPTrim$(ThisCountyString) Then 'find existing customer
'          If InStr(TempRec.CData.CustName, "HEAD") Then Stop
ApprovedP:
'          If QPTrim$(ThisCountyString) = "43962" Then Stop
'          If QPTrim$(TempRec.CData.RPinNum) = "055   004" Then Stop
          If PersValue = 0 And OldRound(TempRec.CData.PEXMPOTHR + TempRec.CData.PEXMPSENI) > 0 Then
            ErrorCode = 3
            GoSub SaveErrors
            GoTo CustErrors
          End If
          If Updated = False Then
            Updated = True
            GoSub UpdateCurrCust
          End If
          If PersValue > 0 Then
            If TaxCust.FirstPersRec = 0 Then
              GoSub SavePersProp
              NextPRec = NextPRec + 1
              PersRec.NextRec = 0
              TaxCust.FirstPersRec = NextPRec 'assign the customer link to it
              Put TCHandle, y, TaxCust
              Put PHandle, NextPRec, PersRec
              ConvRec.CountyAcct = 0
              ConvRec.CountyAcctString = QPTrim$(TaxCust.CountyAcctString)
              ConvRec.CustName = TempRec.CData.CustName
              ConvRec.CVALUE = TempRec.CData.CVALUE
              ConvRec.MCVALUE = TempRec.CData.MCVALUE
              ConvRec.MHVALUE = TempRec.CData.MHVALUE
              ConvRec.MTVALUE = TempRec.CData.MTVALUE
              ConvRec.PersVal = TempRec.CData.PersVal
              ConvRec.PEXMPOTHR = TempRec.CData.PEXMPOTHR
              ConvRec.PEXMPSENI = TempRec.CData.PEXMPSENI
              ConvRec.RPinNum = TempRec.CData.RPinNum
              ConvRec.PPinNum = TempRec.CData.PPinNum
              ConvRec.PROPVALU = 0
              ConvRec.REXMPOTHR = 0
              ConvRec.REXMPSENI = 0
              TCnt = TCnt + 1
              Put CRHandle, TCnt, ConvRec
            Else
              GoSub SavePersProp
              NextPRec = NextPRec + 1
              PersRec.NextRec = TaxCust.FirstPersRec
              TaxCust.FirstPersRec = NextPRec
              Put TCHandle, y, TaxCust
              Put PHandle, NextPRec, PersRec
              ConvRec.CountyAcct = 0
              ConvRec.CountyAcctString = QPTrim$(TaxCust.CountyAcctString)
              ConvRec.CustName = TempRec.CData.CustName
              ConvRec.CVALUE = TempRec.CData.CVALUE
              ConvRec.MCVALUE = TempRec.CData.MCVALUE
              ConvRec.MHVALUE = TempRec.CData.MHVALUE
              ConvRec.MTVALUE = TempRec.CData.MTVALUE
              ConvRec.PersVal = TempRec.CData.PersVal
              ConvRec.PEXMPOTHR = TempRec.CData.PEXMPOTHR
              ConvRec.PEXMPSENI = TempRec.CData.PEXMPSENI
              ConvRec.RPinNum = TempRec.CData.RPinNum
              ConvRec.PPinNum = TempRec.CData.PPinNum
              ConvRec.PROPVALU = 0
              ConvRec.REXMPOTHR = 0
              ConvRec.REXMPSENI = 0
              TCnt = TCnt + 1
              Put CRHandle, TCnt, ConvRec
            End If
          End If
          Exit For
        End If 'TempCountyNum <> CustCountyNumber
      End If 'County Number = 0
SkipItP:
    Next y
AddIt:
    If y > NumOfTCRecs And RealValue = 0 Then
      If QPTrim$(TempRec.CData.CountyAcctString) <> "" Then
        For z = 1 To NewSCustCnt
          If QPTrim$(TempRec.CData.CountyAcctString) = CoNum(z) Then
            Exit For
          End If
        Next z
        If z > NewSCustCnt Then
          NewSCustCnt = NewSCustCnt + 1
          ReDim Preserve NewSCust(1 To NewSCustCnt) As Long 'String
          NewSCust(NewSCustCnt) = x
          ReDim Preserve CoNum(1 To NewSCustCnt) As String
          CoNum(NewSCustCnt) = QPTrim$(TempRec.CData.CountyAcctString)
        End If
      Else
         ErrorCode = 5
         TempRec.CData.CountyAcctString = "NOCNTYNUM" + CStr(ErrorCnt + 1)
         TempRec.CData.CountyAcct = -(ErrorCnt + 1)
         GoSub SaveErrors
         GoTo CustErrors
      End If
    End If
CustErrors:
    frmTCShowPctComp.ShowPctComp x, NumOfTempRecs
    If frmTCShowPctComp.Out = True Then
      Close
      frmTCShowPctComp.Out = False
      Unload frmTCShowPctComp
      Exit Sub
    End If
  Next x
  
  Unload frmTCShowPctComp
  frmTCShowPctComp.Label1 = "Converting County Data"
  frmTCShowPctComp.Show , Me

  FoundCnt = 0
  
  For x = 1 To NumOfTempRecs
    Get TempHandle, x, TempRec
    If TempRec.CData.CountyAcct > 0 Then
      ThisCountyString = CStr(TempRec.CData.CountyAcct)
    ElseIf QPTrim$(TempRec.CData.CountyAcctString) <> "" Then
      ThisCountyString = CStr(TempRec.CData.CountyAcctString)
    Else
      ErrorCode = 5
      GoSub SaveErrors
      GoTo CustErrors
    End If
    PersValue = OldRound(TempRec.CData.CVALUE + TempRec.CData.MCVALUE + TempRec.CData.MHVALUE + TempRec.CData.PersVal + TempRec.CData.MTVALUE)
    RealValue = TempRec.CData.PROPVALU
    For y = 1 To NumOfTCRecs
      Get TCHandle, y, TaxCust
      If TaxCust.CountyAcct > 0 Then
        CustCountyString = CStr(TaxCust.CountyAcct)
      ElseIf QPTrim$(TaxCust.CountyAcctString) <> "" Then
        CustCountyString = QPTrim$(TaxCust.CountyAcctString)
      End If
      If TempRec.CData.CountyAcctString <> "" Then 'make sure a valid county number exists
        If IsNumeric(ThisCountyString) = True And IsNumeric(CustCountyString) = True Then 'should take car of leading
        'zeros in one number but not in the other
          If Val(ThisCountyString) = Val(CustCountyString) Then GoTo ApprovedR
        ElseIf CustCountyString = QPTrim$(ThisCountyString) Then 'find existing customer
ApprovedR:
'          If QPTrim$(TaxCust.CountyAcctString) = "3667" Then Stop
'        If QPTrim$(TaxCust.CountyAcctString) = QPTrim$(TempRec.CData.CountyAcctString) Then 'find existing customer
'          If QPTrim$(TempRec.CData.RPinNum) = "055   004" Then Stop
          If RealValue = 0 And OldRound(TempRec.CData.REXMPOTHR + TempRec.CData.REXMPSENI) > 0 Then
            ErrorCode = 4
            GoSub SaveErrors
            GoTo CustErrorsR
          End If
          If TaxCust.FirstPersRec > 0 And TaxCust.FirstPropRec = 0 Then
            GoTo DontUpdate
          End If
            
          If TaxCust.FirstPropRec > 0 Then '8/29/06 start here
            Get RHandle, TaxCust.FirstPropRec, RealRec
            If QPTrim$(RealRec.RealPin) = QPTrim$(TempRec.CData.RPinNum) Or QPTrim$(TempRec.CData.RPinNum) = "" Then
              GoTo CustErrorsR
            Else
              GoTo AddItR
            End If
          End If
          GoSub UpdateCurrCust
DontUpdate:
          If TempRec.CData.PROPVALU > 0 Then 'determine if this is a real property
            If TaxCust.FirstPropRec = 0 Then 'if this is the first property so far
              FoundRealCnt = FoundRealCnt + 1
              ReDim Preserve FoundReal(1 To FoundRealCnt) As String
              FoundReal(FoundRealCnt) = QPTrim$(TaxCust.CountyAcctString)
              ReDim Preserve FoundRealPin(1 To FoundRealCnt) As String
              FoundRealPin(FoundRealCnt) = QPTrim$(TempRec.CData.RPinNum)
              GoSub SaveRealProp
              NextRRec = NextRRec + 1
              RealRec.NextRec = 0
              TaxCust.FirstPropRec = NextRRec 'assign the customer link to it
              Put TCHandle, y, TaxCust
              Put RHandle, NextRRec, RealRec
              ConvRec.CountyAcct = 0
              ConvRec.CountyAcctString = QPTrim$(TaxCust.CountyAcctString)
              ConvRec.CustName = TempRec.CData.CustName
              ConvRec.CVALUE = 0
              ConvRec.MCVALUE = 0
              ConvRec.MHVALUE = 0
              ConvRec.MTVALUE = 0
              ConvRec.PersVal = 0
              ConvRec.PEXMPOTHR = 0
              ConvRec.PEXMPSENI = 0
              ConvRec.RPinNum = TempRec.CData.RPinNum
              ConvRec.PPinNum = TempRec.CData.PPinNum
              ConvRec.PROPVALU = TempRec.CData.PROPVALU
              ConvRec.REXMPOTHR = TempRec.CData.REXMPOTHR
              ConvRec.REXMPSENI = TempRec.CData.REXMPSENI
              TCnt = TCnt + 1
              Put CRHandle, TCnt, ConvRec
            End If
            Exit For
          End If
        End If 'ends search for co num match
      End If 'County Number = 0
SkipItR:
    Next y
    If y > NumOfTCRecs And RealValue > 0 Then
AddItR:
      If QPTrim$(TempRec.CData.CountyAcctString) <> "" Then
        NewSCustCnt = NewSCustCnt + 1
        ReDim Preserve NewSCust(1 To NewSCustCnt) As Long 'String
        NewSCust(NewSCustCnt) = x
      Else
        ErrorCode = 5
        TempRec.CData.CountyAcctString = "NOCNTYNUM" + CStr(ErrorCnt + 1)
        TempRec.CData.CountyAcct = -(ErrorCnt + 1)
        GoSub SaveErrors
        GoTo CustErrorsR
      End If
    End If
CustErrorsR:
    frmTCShowPctComp.ShowPctComp x, NumOfTempRecs
    If frmTCShowPctComp.Out = True Then
      Close
      frmTCShowPctComp.Out = False
      Unload frmTCShowPctComp
      Exit Sub
    End If
  Next x
  Close
  Unload frmTCShowPctComp
  
  If NewSCustCnt > 0 Then
    Call SaveNewSingleRCust(NewSCustCnt, NewSCust())
  End If
  
  frmTCLoadingRpt.Show
  frmTCLoadingRpt.Label1.Caption = "Indexing...please wait"
  DoEvents

  Call CreateCustNameIdx
  Unload frmTCLoadingRpt
  frmTCLoadingRpt.Show
  frmTCLoadingRpt.Label1.Caption = "1 Of 4 completed"
  DoEvents
  Call CreateSrchNameIdx
  Unload frmTCLoadingRpt
  frmTCLoadingRpt.Show
  frmTCLoadingRpt.Label1.Caption = "2 Of 4 completed"
  DoEvents
  Call CreateOptCustIdx
  Unload frmTCLoadingRpt
  frmTCLoadingRpt.Show
  frmTCLoadingRpt.Label1.Caption = "3 Of 4 completed"
  DoEvents
  Call CreateSSIdx
  
  Unload frmTCLoadingRpt
  Call Savemsg(900, "Tax data has converted successfully.")
  
  Exit Sub
  
SaveRealProp:
  TempRPin = QPTrim$(TempRec.CData.RPinNum)
  For z = 1 To NumOfRealRecs
    Get OldRHandle, z, OldRealRec
    If QPTrim$(OldRealRec.RealPin) = TempRPin Then
      Exit For
    End If
  Next z
  
  RealRec.Blank = ""
  RealRec.CustPin = TaxCust.PIN
  RealRec.Deleted = 0
  RealRec.EXMPOTHR = TempRec.CData.REXMPOTHR
  RealRec.EXMPSENI = TempRec.CData.REXMPSENI
  RealRec.Fill1 = ""
  RealRec.LOTACRE = TempRec.CData.LOTACRE
  RealRec.PropSize = TempRec.CData.PropSize
  RealRec.PROPVALU = TempRec.CData.PROPVALU
  RealRec.RealPin = TempRec.CData.RPinNum
  RealRec.BLOCK = TempRec.CData.BLOCK
  RealRec.PROPNOT1 = TempRec.CData.RDESC1
  RealRec.PROPNOT2 = TempRec.CData.RDESC2
  RealRec.PROPNOT3 = TempRec.CData.RDESC3
  RealRec.Map = TempRec.CData.Map
  RealRec.LOTNUMB = TempRec.CData.LOTACRE
  
  If TempRec.CData.OptRev1Chrg <> 0 Then
    RealRec.OptRev1Chrg = 1
  Else
    RealRec.OptRev1Chrg = OldRealRec.OptRev1Chrg
  End If
  If TempRec.CData.OptRev2Chrg <> 0 Then
    RealRec.OptRev2Chrg = 2
  Else
    RealRec.OptRev2Chrg = OldRealRec.OptRev2Chrg
  End If
  If TempRec.CData.OptRev3Chrg <> 0 Then
    RealRec.OptRev3Chrg = 3
  Else
    RealRec.OptRev3Chrg = OldRealRec.OptRev3Chrg
  End If
  If QPTrim$(TempRec.CData.RealOptSearch) <> "" Then
    RealRec.OptSearch = QPTrim$(TempRec.CData.RealOptSearch)
  Else
    RealRec.OptSearch = QPTrim$(OldRealRec.OptSearch)
  End If
  If QPTrim$(TempRec.CData.LateList) <> "" Then
    RealRec.LateList = "Y"
  End If
  If QPTrim$(TempRec.CData.RTownShip) <> "" Then
    RealRec.TownShip = QPTrim$(TempRec.CData.RTownShip)
  End If
  If z <= NumOfRealRecs Then
    RealRec.InternalPin = OldRealRec.InternalPin
    RealRec.GISPOS = OldRealRec.GISPOS
    RealRec.ICPDesc = OldRealRec.ICPDesc
    RealRec.Image = OldRealRec.Image
    RealRec.LastYrPrinted = OldRealRec.LastYrPrinted
    If QPTrim$(TempRec.CData.LateList) = "" Then
      RealRec.LateList = OldRealRec.LateList
    End If
    RealRec.LienDesc = OldRealRec.LienDesc
    RealRec.Mock = OldRealRec.Mock
    If QPTrim$(TempRec.CData.MORTCODE) <> "" Then
      RealRec.MORTCODE = QPTrim$(TempRec.CData.MORTCODE)
    Else
      RealRec.MORTCODE = OldRealRec.MORTCODE
    End If
    RealRec.PropAddr = OldRealRec.PropAddr
    RealRec.PROPDATE = OldRealRec.PROPDATE
    RealRec.PROPDISC = OldRealRec.PROPDISC
'    RealRec.TownShip = OldRealRec.TownShip
    If QPTrim$(TempRec.CData.RTownShip) = "" Then
      RealRec.TownShip = OldRealRec.TownShip
    End If
  Else
    RealRec.InternalPin = IntRecCnt + 1
    Put IHandle, RealRec.InternalPin, IntPinRec
    RealRec.GISPOS = ""
    RealRec.ICPDesc = ""
    RealRec.Image = ""
    RealRec.LastYrPrinted = 0
    If QPTrim$(TempRec.CData.LateList) = "" Then
      RealRec.LateList = fpcmbRLateList.Text
    End If
    RealRec.LienDesc = ""
    RealRec.Mock = "N"
    RealRec.MORTCODE = QPTrim$(TempRec.CData.MORTCODE)
    RealRec.PropAddr = ""
    RealRec.PROPDATE = Date2Num(fptxtRPropDate.Text)
    RealRec.PROPDISC = "N"
    If QPTrim$(TempRec.CData.RTownShip) = "" Then
      RealRec.TownShip = ""
    End If
'    RealRec.TownShip = ""
  End If
  
  Return
  
SavePersProp:
  TempPPin = QPTrim$(TempRec.CData.PPinNum)
  For z = 1 To NumOfPersRecs
    Get OldPHandle, z, OldPersRec
    If QPTrim$(OldPersRec.PropPin) = TempPPin Then
      Exit For
    End If
  Next z
  
  PersRec.Blank = ""
  PersRec.CustPin = TaxCust.PIN
  PersRec.CVALUE = TempRec.CData.CVALUE
  PersRec.Deleted = 0
  PersRec.DESC1 = TempRec.CData.PDESC1
  PersRec.DESC2 = TempRec.CData.PDESC2
  PersRec.DESC3 = TempRec.CData.PDESC3
  PersRec.EXMPOTHR = TempRec.CData.PEXMPOTHR
  PersRec.EXMPSENI = TempRec.CData.PEXMPSENI
  PersRec.MCVALUE = TempRec.CData.MCVALUE
  PersRec.MHVALUE = TempRec.CData.MHVALUE
  PersRec.MTVALUE = TempRec.CData.MTVALUE
  PersRec.PersVal = TempRec.CData.PersVal
  PersRec.PropPin = TempRec.CData.PPinNum
  If QPTrim$(TempRec.CData.LateList) <> "" Then
    PersRec.LateList = "Y"
  End If
  If z <= NumOfPersRecs Then
    PersRec.Desc4 = OldPersRec.Desc4
    PersRec.Desc5 = OldPersRec.Desc5
    PersRec.DISCOV = OldPersRec.DISCOV
    PersRec.DMVSubmitted = OldPersRec.DMVSubmitted
    PersRec.InternalPin = OldPersRec.InternalPin
    PersRec.LastYrPrinted = OldPersRec.LastYrPrinted
    If QPTrim$(TempRec.CData.LateList) = "" Then
      PersRec.LateList = OldPersRec.LateList
    End If
    PersRec.PROPDATE = OldPersRec.PROPDATE
    PersRec.VehTaxYear = OldPersRec.VehTaxYear
  Else
    PersRec.Desc4 = ""
    PersRec.Desc5 = ""
    PersRec.DISCOV = "N"
    PersRec.DMVSubmitted = "N"
    PersRec.InternalPin = IntRecCnt + 1
    Put IHandle, PersRec.InternalPin, IntPinRec
    PersRec.LastYrPrinted = 0
    If QPTrim$(TempRec.CData.LateList) = "" Then
      PersRec.LateList = fpcmbPLateList.Text
    End If
    PersRec.PROPDATE = OldPersRec.PROPDATE
    PersRec.VehTaxYear = 0
  End If
  
  Return
  
UpdateCurrCust:
  Updated = True
  TaxCust.Deleted = 0
  If QPTrim$(TempRec.CData.Addr1) <> "" Then
    TaxCust.Addr1 = TempRec.CData.Addr1
  End If
  If QPTrim$(TempRec.CData.Addr2) <> "" Then
    TaxCust.Addr2 = TempRec.CData.Addr2
  End If
  If QPTrim$(TempRec.CData.City) <> "" Then
    TaxCust.City = TempRec.CData.City
  End If
  If QPTrim$(TempRec.CData.CustName) <> "" Then
    TaxCust.CustName = TempRec.CData.CustName
  End If
  If QPTrim$(TempRec.CData.Zip) <> "" Then
    TaxCust.Zip = TempRec.CData.Zip
  End If
  If QPTrim$(TempRec.CData.State) <> "" Then
    TaxCust.State = TempRec.CData.State
  End If
  If QPTrim$(TempRec.CData.SName) <> "" Then
    TaxCust.SName = TempRec.CData.SName
  End If
  If QPTrim$(TempRec.CData.OptSrchDesc) <> "" Then
    TaxCust.OptSrchDesc = TempRec.CData.OptSrchDesc
  End If
  If QPTrim$(TempRec.CData.County4BillName) <> "" Then
    TaxCust.County4BillName = TempRec.CData.County4BillName
  End If
  TaxCust.Active = "Y"
  If TempRec.CData.Cycle > 0 Then
    TaxCust.Cycle = TempRec.CData.Cycle
  End If
  If QPTrim$(TempRec.CData.CycleName) <> "" Then
    TaxCust.CycleName = QPTrim$(TempRec.CData.CycleName)
  End If
  If QPTrim$(TempRec.CData.CTownShip) <> "" Then
    TaxCust.TownShip = QPTrim$(TempRec.CData.CTownShip)
  End If
  Put TCHandle, y, TaxCust
  Return

SaveErrors:
  ErrorCnt = ErrorCnt + 1
  ErrorRec.CountyAcct = TempRec.CData.CountyAcct
  ErrorRec.CountyAcctString = QPTrim$(TempRec.CData.CountyAcctString)
  ErrorRec.CustName = QPTrim$(TempRec.CData.CustName)
  ErrorRec.ErrorType = ErrorCode
  ErrorRec.PersTot = PersValue
  ErrorRec.PersXTot = OldRound(TempRec.CData.PEXMPOTHR + TempRec.CData.PEXMPSENI)
  ErrorRec.RPinNum = QPTrim$(TempRec.CData.RPinNum)
  ErrorRec.PPinNum = QPTrim$(TempRec.CData.PPinNum)
  ErrorRec.RealTot = RealValue
  ErrorRec.RealXTot = OldRound(TempRec.CData.REXMPOTHR + TempRec.CData.REXMPSENI)
  Put EHandle, ErrorCnt, ErrorRec
  
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTCConvert", "ProcessRSingle", Erl)
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

Private Sub SaveNewSingleRCust(NewSCustCnt As Long, NewSCust() As Long)
  Dim TaxCust As TaxCustType
  Dim NewTaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim x As Long, y As Long, z As Long
  Dim TempHandle As Integer
  Dim TempRec As TempConversionData
  Dim NumOfTempRecs As Long
  Dim IntPinRec As InternalPinType
  Dim IHandle As Integer
  Dim NumOfIRecs As Long
  Dim IntCnt As Long
  Dim ThisCoNum As Long '$
  Dim CustCnt As Long
  Dim RealCnt As Long
  Static PersCnt As Long
  Dim FirstTime As Boolean
  Dim PersValue As Double
  Dim ConvRec As ConvResultsType
  Dim CRHandle As Integer
  Dim NumOfCRRecs As Long
  Dim RealValue As Double
  Dim NumOfErrors As Long
  Dim ErrorCnt As Long
  Dim ErrorRec As ConvErrorType
  Dim EHandle As Integer
  Dim NumOfERecs As Long
  Dim ErrorCode As Integer
  Dim ThisPPin$
  Dim ThisRPin$
  Dim ThisCountyString$
  Dim CustCountyString$
  
'  On Error GoTo ERRORSTUFF
  
  OpenTempConvFile TempHandle, NumOfTempRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenRealPropFile RHandle, NumOfRealRecs
  OpenPersPropFile PHandle, NumOfPersRecs
  OpenIntPinFile IHandle, NumOfIRecs
  OpenConvResultsFile CRHandle, NumOfCRRecs
  OpenConvErrorsFile EHandle, NumOfERecs
  
  ErrorCnt = NumOfERecs
  TCnt = NumOfCRRecs
  
  RealCnt = NumOfRealRecs
  PersCnt = NumOfPersRecs
  CustCnt = NumOfTCRecs
  IntCnt = NumOfIRecs
  frmTCShowPctComp.Label1 = "Adding New Customers"
  frmTCShowPctComp.Show , Me
  
  For x = 1 To NewSCustCnt
    ThisCoNum = NewSCust(x)
    FirstTime = True
      Get TempHandle, ThisCoNum, TempRec
'      If QPTrim$(TempRec.CData.RPinNum) = "055   004" Then Stop
      PersValue = OldRound(TempRec.CData.CVALUE + TempRec.CData.MCVALUE + TempRec.CData.MHVALUE + TempRec.CData.PersVal + TempRec.CData.MTVALUE)
      RealValue = TempRec.CData.PROPVALU
      If PersValue = 0 Then GoTo SkipItP
        ThisPPin = QPTrim$(TempRec.CData.PPinNum)
        If (PersValue > 0 And (ThisPPin = "" Or ThisPPin = "0")) Then
          ErrorCode = 1
          GoSub SaveErrors
          GoTo CustErrors
        ElseIf PersValue = 0 Then
          ErrorCode = 2
          GoSub SaveErrors
          GoTo CustErrors
        ElseIf PersValue = 0 And OldRound(TempRec.CData.PEXMPOTHR + TempRec.CData.PEXMPSENI) > 0 Then
          ErrorCode = 3
          GoSub SaveErrors
          GoTo CustErrors
        End If
        If FirstTime = True Then
          If PersValue > 0 Then
            PersCnt = PersCnt + 1
            GoSub SavePersData
            PersRec.NextRec = 0
            Put PHandle, PersCnt, PersRec
            TaxCust.FirstPropRec = 0
            TaxCust.FirstPersRec = PersCnt
            ConvRec.CountyAcct = TempRec.CData.CountyAcct
            ConvRec.CountyAcctString = QPTrim$(TempRec.CData.CountyAcctString)
            ConvRec.CustName = TempRec.CData.CustName
            ConvRec.CVALUE = TempRec.CData.CVALUE
            ConvRec.MCVALUE = TempRec.CData.MCVALUE
            ConvRec.MHVALUE = TempRec.CData.MHVALUE
            ConvRec.MTVALUE = TempRec.CData.MTVALUE
            ConvRec.PersVal = TempRec.CData.PersVal
            ConvRec.PEXMPOTHR = TempRec.CData.PEXMPOTHR
            ConvRec.PEXMPSENI = TempRec.CData.PEXMPSENI
            ConvRec.PPinNum = TempRec.CData.PPinNum
            ConvRec.RPinNum = TempRec.CData.RPinNum
            ConvRec.PROPVALU = 0
            ConvRec.REXMPOTHR = 0
            ConvRec.REXMPSENI = 0
            TCnt = TCnt + 1
            GoSub SaveCustData 'new 8/29
            Put CRHandle, TCnt, ConvRec
            Put TCHandle, CustCnt, TaxCust
            FirstTime = False
          End If
        Else
          If PersValue > 0 Then
            PersCnt = PersCnt + 1
            GoSub SavePersData
            PersRec.NextRec = TaxCust.FirstPersRec
            Put PHandle, PersCnt, PersRec
            TaxCust.FirstPersRec = PersCnt
            Put TCHandle, CustCnt, TaxCust
            ConvRec.CountyAcct = TempRec.CData.CountyAcct
            ConvRec.CountyAcctString = QPTrim$(TempRec.CData.CountyAcctString)
            ConvRec.CustName = TempRec.CData.CustName
            ConvRec.CVALUE = TempRec.CData.CVALUE
            ConvRec.MCVALUE = TempRec.CData.MCVALUE
            ConvRec.MHVALUE = TempRec.CData.MHVALUE
            ConvRec.MTVALUE = TempRec.CData.MTVALUE
            ConvRec.PersVal = TempRec.CData.PersVal
            ConvRec.PEXMPOTHR = TempRec.CData.PEXMPOTHR
            ConvRec.PEXMPSENI = TempRec.CData.PEXMPSENI
            ConvRec.PPinNum = TempRec.CData.PPinNum
            ConvRec.RPinNum = TempRec.CData.RPinNum
            ConvRec.PROPVALU = 0
            ConvRec.REXMPOTHR = 0
            ConvRec.REXMPSENI = 0
            TCnt = TCnt + 1
            Put CRHandle, TCnt, ConvRec
          End If
      End If
SkipItP:
CustErrors:
    frmTCShowPctComp.ShowPctComp x, NewSCustCnt
    If frmTCShowPctComp.Out = True Then
      Close
      frmTCShowPctComp.Out = False
      Unload frmTCShowPctComp
      Exit Sub
    End If
  Next x
  Unload frmTCShowPctComp
  
  frmTCShowPctComp.Label1 = "Adding New Customers"
  frmTCShowPctComp.Show , Me
  
  For x = 1 To NewSCustCnt
    ThisCoNum = NewSCust(x)
    FirstTime = True
    Get TempHandle, ThisCoNum, TempRec
'      If QPTrim$(TempRec.CData.RPinNum) = "055   004" Then Stop
      RealValue = TempRec.CData.PROPVALU
      If RealValue = 0 Then GoTo SkipItR
        ThisRPin = QPTrim$(TempRec.CData.RPinNum)
        If (RealValue > 0 And (ThisRPin = "" Or ThisRPin = "0")) Then
          ErrorCode = 1
          GoSub SaveErrors
          GoTo SkipItR
        ElseIf RealValue = 0 Then
          ErrorCode = 2
          GoSub SaveErrors
          GoTo SkipItR
        ElseIf RealValue = 0 And OldRound(TempRec.CData.REXMPOTHR + TempRec.CData.REXMPSENI) > 0 Then
          ErrorCode = 4
          GoSub SaveErrors
          GoTo SkipItR
        End If
        If TempRec.CData.PROPVALU > 0 Then
          For z = 1 To FoundRealCnt
            If Val(FoundReal(z)) = ThisCoNum And FoundRealPin(z) = QPTrim$(TempRec.CData.RPinNum) Then
              Exit For
            End If
          Next z
          If z <= FoundRealCnt And FoundRealCnt > 0 Then
            GoTo SkipItR
          End If
          FoundRealCnt = FoundRealCnt + 1
          ReDim Preserve FoundReal(1 To FoundRealCnt) As String
          ReDim Preserve FoundRealPin(1 To FoundRealCnt) As String
          FoundReal(FoundRealCnt) = ThisCoNum
          FoundRealPin(FoundRealCnt) = QPTrim$(TempRec.CData.RPinNum)
          RealCnt = RealCnt + 1
          GoSub SaveRealData
          If TempRec.CData.CountyAcct > 0 Then
            ThisCountyString = CStr(TempRec.CData.CountyAcct)
          ElseIf QPTrim$(TempRec.CData.CountyAcctString) Then
            ThisCountyString = QPTrim$(TempRec.CData.CountyAcctString)
          End If
'          ThisCountyString = QPTrim$(ThisCountyString)
          For z = 1 To NumOfTCRecs
            Get TCHandle, z, TaxCust
            If TaxCust.CountyAcct > 0 Then
              CustCountyString = CStr(TaxCust.CountyAcct)
            ElseIf QPTrim$(TaxCust.CountyAcctString) <> "" Then
              CustCountyString = QPTrim$(TaxCust.CountyAcctString)
            End If
            If CustCountyString = ThisCountyString Then
              If TaxCust.FirstPropRec > 0 Then
                TaxCust = NewTaxCust
                GoSub SaveCustData
                Exit For
              Else
                GoSub SaveCustData
                Exit For
              End If
            End If
          Next z
          If z > NumOfTCRecs Then
            TaxCust = NewTaxCust
            GoSub SaveCustData 'added 8/28/06
          End If
          RealRec.NextRec = 0 'TaxCust.FirstPropRec
          Put RHandle, RealCnt, RealRec
          TaxCust.FirstPropRec = RealCnt
          Put TCHandle, CustCnt, TaxCust
          ConvRec.CountyAcct = TempRec.CData.CountyAcct
          ConvRec.CountyAcctString = QPTrim$(TempRec.CData.CountyAcctString)
          ConvRec.CustName = TempRec.CData.CustName
          ConvRec.CVALUE = 0
          ConvRec.MCVALUE = 0
          ConvRec.MHVALUE = 0
          ConvRec.MTVALUE = 0
          ConvRec.PersVal = 0
          ConvRec.PEXMPOTHR = 0
          ConvRec.PEXMPSENI = 0
          ConvRec.PPinNum = TempRec.CData.PPinNum
          ConvRec.RPinNum = TempRec.CData.RPinNum
          ConvRec.PROPVALU = TempRec.CData.PROPVALU
          ConvRec.REXMPOTHR = TempRec.CData.REXMPOTHR
          ConvRec.REXMPSENI = TempRec.CData.REXMPSENI
          TCnt = TCnt + 1
          Put CRHandle, TCnt, ConvRec
      End If
SkipItR:
    frmTCShowPctComp.ShowPctComp x, NewSCustCnt
    If frmTCShowPctComp.Out = True Then
      Close
      frmTCShowPctComp.Out = False
      Unload frmTCShowPctComp
      Exit Sub
    End If
  Next x
  Unload frmTCShowPctComp
  
  Close
  
  Exit Sub

SaveCustData:
  CustCnt = CustCnt + 1
  TaxCust.Acct = CustCnt
  TaxCust.Active = "Y"
  TaxCust.Addr1 = QPTrim$(TempRec.CData.Addr1)
  TaxCust.Addr2 = QPTrim$(TempRec.CData.Addr2)
  TaxCust.Bankrupt = "N"
  TaxCust.City = QPTrim$(TempRec.CData.City)
  TaxCust.County4BillName = QPTrim$(TempRec.CData.County4BillName)
  TaxCust.County4BillNum = 0
  TaxCust.CountyAcct = TempRec.CData.CountyAcct
  TaxCust.CountyAcctString = QPTrim$(TempRec.CData.CountyAcctString)
  TaxCust.CSSN = QPTrim$(TempRec.CData.CSSN)
  TaxCust.CustName = QPTrim$(TempRec.CData.CustName)
'  TaxCust.Cycle = 0
'  TaxCust.CycleName = ""
  If TempRec.CData.Cycle > 0 Then
    TaxCust.Cycle = TempRec.CData.Cycle
  Else
    TaxCust.Cycle = 0
  End If
  If QPTrim$(TempRec.CData.CycleName) <> "" Then
    TaxCust.CycleName = QPTrim$(TempRec.CData.CycleName)
  Else
    TaxCust.CycleName = ""
  End If
  TaxCust.Deleted = 0
  TaxCust.DeliveryPt = ""
  TaxCust.DrvrsLic = ""
  TaxCust.Employer = ""
  TaxCust.FileVer = FileVers
  If PersValue > 0 Then
    TaxCust.FirstPersRec = PersCnt
  ElseIf TempRec.CData.PROPVALU > 0 Then
    TaxCust.FirstPropRec = RealCnt
  End If
  TaxCust.HPHONE = ""
  TaxCust.Interest = fpcmbChrgInt.Text
  TaxCust.LateNotice = fpcmbLateList.Text
  TaxCust.OPENDATE = Date2Num(fptxtOpenDate.Text)
  TaxCust.OptSrchDesc = TempRec.CData.OptSrchDesc
  TaxCust.OSSN = QPTrim$(TempRec.CData.OSSN)
  TaxCust.Pad1 = ""
  TaxCust.Penalty = fpcmbPenalty.Text
  TaxCust.PIN = CustCnt
  TaxCust.PostalRt = ""
  TaxCust.PrePayBal = 0
  TaxCust.PrePayTrans = 0
  TaxCust.ServiceAdd = ""
  TaxCust.SName = TempRec.CData.SName
  TaxCust.State = TempRec.CData.State
  TaxCust.TaxExempt = fpcmbTaxEx.Text
'  TaxCust.TownShip = ""
  If QPTrim$(TempRec.CData.CTownShip) <> "" Then
    TaxCust.TownShip = QPTrim$(TempRec.CData.CTownShip)
  Else
    TaxCust.TownShip = ""
  End If
  TaxCust.WPHONE = ""
  TaxCust.Zip = TempRec.CData.Zip
  
  Return
  
SaveRealData:
  RealRec.Blank = ""
  RealRec.BLOCK = TempRec.CData.BLOCK
  If FirstTime = True Then
    RealRec.CustPin = CustCnt + 1 'added + 1 7/14/06
  Else
    RealRec.CustPin = CustCnt
  End If
  RealRec.Deleted = 0
  RealRec.EXMPOTHR = TempRec.CData.REXMPOTHR
  RealRec.EXMPSENI = TempRec.CData.REXMPSENI
  RealRec.Fill1 = ""
  RealRec.GISPOS = ""
  RealRec.ICPDesc = ""
  RealRec.Image = ""
  IntCnt = IntCnt + 1
  RealRec.InternalPin = IntCnt
  IntPinRec.PIN = RealCnt
  Put IHandle, IntCnt, IntPinRec
  RealRec.LastYrPrinted = 0
  If QPTrim$(TempRec.CData.LateList) <> "" Then
    RealRec.LateList = "Y"
  Else
    RealRec.LateList = fpcmbRLateList.Text
  End If
  RealRec.LienDesc = ""
  RealRec.LienYN = "N"
  RealRec.LOTACRE = TempRec.CData.LOTACRE
  RealRec.LOTNUMB = ""
  RealRec.Map = TempRec.CData.Map
  RealRec.LOTNUMB = TempRec.CData.LOTACRE
  RealRec.Mock = "N"
  RealRec.MORTCODE = QPTrim$(TempRec.CData.MORTCODE)
  If TempRec.CData.OptRev1Chrg <> 0 Then
    RealRec.OptRev1Chrg = 1
  End If
  If TempRec.CData.OptRev2Chrg <> 0 Then
    RealRec.OptRev1Chrg = 2
  End If
  If TempRec.CData.OptRev3Chrg <> 0 Then
    RealRec.OptRev3Chrg = 3
  End If
  RealRec.OptSearch = QPTrim$(TempRec.CData.RealOptSearch)
  RealRec.PropAddr = ""
  RealRec.PROPDATE = Date2Num(fptxtRPropDate.Text)
  RealRec.PROPDISC = "N"
  RealRec.PROPNOT1 = TempRec.CData.RDESC1
  RealRec.PROPNOT2 = TempRec.CData.RDESC2
  RealRec.PROPNOT3 = TempRec.CData.RDESC3
  RealRec.PropSize = TempRec.CData.PropSize
  RealRec.PROPVALU = TempRec.CData.PROPVALU
  RealRec.RealPin = TempRec.CData.RPinNum
  If QPTrim$(TempRec.CData.RTownShip) <> "" Then
    RealRec.TownShip = QPTrim$(TempRec.CData.RTownShip)
  Else
    RealRec.TownShip = ""
  End If
'  RealRec.TownShip = ""
  
  Return
  
SavePersData:
  PersRec.Blank = ""
  If FirstTime = True Then
    PersRec.CustPin = CustCnt + 1 'added + 1 7/14/06
  Else
    PersRec.CustPin = CustCnt
  End If
  PersRec.CVALUE = TempRec.CData.CVALUE
  PersRec.Deleted = 0
  PersRec.DESC1 = TempRec.CData.PDESC1
  PersRec.DESC2 = TempRec.CData.PDESC2
  PersRec.DESC3 = TempRec.CData.PDESC3
  PersRec.Desc4 = ""
  PersRec.Desc5 = ""
  PersRec.DISCOV = "N"
  PersRec.DMVSubmitted = "N"
  PersRec.EXMPOTHR = TempRec.CData.PEXMPOTHR
  PersRec.EXMPSENI = TempRec.CData.PEXMPSENI
  IntCnt = IntCnt + 1
  PersRec.InternalPin = IntCnt
  IntPinRec.PIN = PersCnt
  Put IHandle, IntCnt, IntPinRec
  PersRec.LastYrPrinted = 0
  If QPTrim$(TempRec.CData.LateList) <> "" Then
    PersRec.LateList = "Y"
  Else
    PersRec.LateList = fpcmbPLateList.Text
  End If
  PersRec.MCVALUE = TempRec.CData.MCVALUE
  PersRec.MHVALUE = TempRec.CData.MHVALUE
  PersRec.MTVALUE = TempRec.CData.MTVALUE
  PersRec.PersVal = TempRec.CData.PersVal
  PersRec.PROPDATE = Date2Num(fptxtPPropDate.Text)
  PersRec.PropPin = TempRec.CData.PPinNum
  PersRec.VehTaxYear = 0
  
  Return
  
SaveErrors:
  ErrorCnt = ErrorCnt + 1
  ErrorRec.CountyAcct = TempRec.CData.CountyAcct
  ErrorRec.CountyAcctString = QPTrim$(TempRec.CData.CountyAcctString)
  ErrorRec.CustName = QPTrim$(TempRec.CData.CustName)
  ErrorRec.ErrorType = ErrorCode
  ErrorRec.PersTot = PersValue
  ErrorRec.PersXTot = OldRound(TempRec.CData.PEXMPOTHR + TempRec.CData.PEXMPSENI)
  ErrorRec.PPinNum = QPTrim$(TempRec.CData.PPinNum)
  ErrorRec.RPinNum = QPTrim$(TempRec.CData.RPinNum)
  ErrorRec.RealTot = RealValue
  ErrorRec.RealXTot = OldRound(TempRec.CData.REXMPOTHR + TempRec.CData.REXMPSENI)
  Put EHandle, ErrorCnt, ErrorRec
  
  Return
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTCConvert", "SaveNewSingleRCust", Erl)
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

Private Sub fpcmbPenalty_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbPenalty.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPenalty.ListIndex = -1
  End If
  If fpcmbPenalty.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbTaxEx.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbPLateList_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbPLateList.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPLateList.ListIndex = -1
  End If
  If fpcmbPLateList.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtPPropDate.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbRLateList_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbRLateList.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbRLateList.ListIndex = -1
  End If
  If fpcmbRLateList.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtRPropDate.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbTaxEx_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbTaxEx.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbTaxEx.ListIndex = -1
  End If
  If fpcmbTaxEx.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbLateList.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbLateList_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbLateList.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbLateList.ListIndex = -1
  End If
  If fpcmbLateList.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbBankrupt.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbBankrupt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbBankrupt.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbBankrupt.ListIndex = -1
  End If
  If fpcmbBankrupt.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbActive.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbActive_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbActive.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbActive.ListIndex = -1
  End If
  If fpcmbActive.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtOpenDate.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub ProcessRMultiple()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim OldRealRec As PropertyRecType
  Dim OldRHandle As Integer
  Dim NumOfRealRecs As Long
  Dim NumOfNewRealRecs As Long
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim OldPersRec As PersonalRecType
  Dim OldPHandle As Integer
  Dim NumOfPersRecs As Long
  Dim NumOfNewPersRecs As Long
  Dim x As Long, y As Long, z As Long
  Dim TempHandle As Integer
  Dim TempRec As TempConversionData
  Dim NumOfTempRecs As Long
  Dim NewSCustCnt As Long
  Dim NewNCustCnt As Long
  Dim NextRRec As Long
  Dim NextPRec As Long
  Dim OldRPin As String
  Dim TempRPin As String
  Dim TempPPin As String
  Dim IntPinRec As InternalPinType
  Dim IHandle As Integer
  Dim NumOfIRecs As Long
  Dim IntRecCnt As Long
  Dim PersValue As Double
  Dim AddTotal As Double
  Dim ConvRec As ConvResultsType
  Dim CRHandle As Integer
  Dim NumOfCRRecs As Long
  Dim NumOfErrors As Long
  Dim ErrorCnt As Long
  Dim ErrorRec As ConvErrorType
  Dim EHandle As Integer
  Dim NumOfERecs As Long
  Dim RealValue As Double
  Dim ErrorCode As Integer
  Dim Updated As Boolean
  Dim ThisPPin$
  Dim ThisRPin$
  Dim ThisCountyString$
  Dim CustCountyString$
  
'  On Error GoTo ERRORSTUFF
  
  If Not Exist(ConversionFile) Then
    Call TCMsg(900, "Please process the county data first. Load attempt aborted.")
    Exit Sub
  End If
  
  If TCMsgWOpts(700, "WARNING: Continuing will delete all real property and personal property files so they can be rebuilt with the new data from the county. Press F10 to continue. Otherwise, press ESC to abort.", "F10 Convert", "ESC Abort") = "abort" Then
    Exit Sub
  End If
  
  ReDim NewSCust(1 To 1) As String 'county number is a string
'  ReDim NewNCust(1 To 1) As Double 'county number is a number
  
  If Exist("OLDTAXPROP.DAT") Then
    If TCMsgWOpts(900, "'OLDTAXPROP.DAT' already exists. Do you wish to overwrite?", "F10 Overwrite", "ESC Leave As Is") <> "abort" Then
      KillFile "OLDTAXPROP.DAT"
      Name "TAXPROP.DAT" As "OLDTAXPROP.DAT"
    End If
  Else
    Name "TAXPROP.DAT" As "OLDTAXPROP.DAT"
  End If
  
  If Exist("OLDTAXPERS.DAT") Then
    If TCMsgWOpts(900, "'OLDTAXPERS.DAT' already exists. Do you wish to overwrite?", "F10 Overwrite", "ESC Leave As Is") <> "abort" Then
      KillFile "OLDTAXPERS.DAT"
      Name "TAXPERS.DAT" As "OLDTAXPERS.DAT"
    End If
  Else
    Name "TAXPERS.DAT" As "OLDTAXPERS.DAT"
  End If
  
'  KillFile "TAXPERS.DAT"
'  KillFile "TAXPROP.DAT"
  
  OpenOldRealPropFile OldRHandle, NumOfRealRecs
  Get OldRHandle, 1, OldRealRec
  OpenOldPersPropFile OldPHandle, NumOfPersRecs
  Get OldPHandle, 1, OldPersRec
  
  OpenTempConvFile TempHandle, NumOfTempRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
'    If TaxCust.Cycle = 1 Or TaxCust.Cycle = 3 Then GoTo SkipThis 'for Maxton
    TaxCust.FirstPersRec = 0
    TaxCust.FirstPropRec = 0
    Put TCHandle, x, TaxCust
SkipThis:
  Next x
  
  OpenRealPropFile RHandle, NumOfNewRealRecs
  OpenPersPropFile PHandle, NumOfNewPersRecs
  OpenIntPinFile IHandle, NumOfIRecs
  
  IntRecCnt = NumOfIRecs
  TCnt = 0
  KillFile "CNVRSLTS.DAT"
  KillFile "CNVRERRS.DAT"
  
  OpenConvErrorsFile EHandle, NumOfERecs
  OpenConvResultsFile CRHandle, NumOfCRRecs
  
  frmTCShowPctComp.Label1 = "Converting County Data"
  frmTCShowPctComp.Show , Me
  
  For x = 1 To NumOfTempRecs
    Get TempHandle, x, TempRec
    If TempRec.CData.CountyAcct > 0 Then
      ThisCountyString = CStr(TempRec.CData.CountyAcct)
    ElseIf QPTrim$(TempRec.CData.CountyAcctString) <> "" Then
      ThisCountyString = CStr(TempRec.CData.CountyAcctString)
    Else
      ErrorCode = 5
      GoSub SaveErrors
      GoTo CustErrors
    End If
    PersValue = OldRound(TempRec.CData.CVALUE + TempRec.CData.MCVALUE + TempRec.CData.MHVALUE + TempRec.CData.PersVal + TempRec.CData.MTVALUE)
    RealValue = TempRec.CData.PROPVALU
    For y = 1 To NumOfTCRecs
      Get TCHandle, y, TaxCust
      If TaxCust.CountyAcct > 0 Then
        CustCountyString = CStr(TaxCust.CountyAcct)
      ElseIf QPTrim$(TaxCust.CountyAcctString) <> "" Then
        CustCountyString = QPTrim$(TaxCust.CountyAcctString)
      End If
      Updated = False
      If ThisCountyString <> "" Then 'make sure a valid county number exists
        If CustCountyString = ThisCountyString Then 'find existing customer
          If PersValue > 0 And RealValue > 0 Then
            ErrorCode = 1
            GoSub SaveErrors
            GoTo CustErrors
          ElseIf PersValue = 0 And RealValue = 0 Then
            ErrorCode = 2
            GoSub SaveErrors
            GoTo CustErrors
          ElseIf PersValue = 0 And OldRound(TempRec.CData.PEXMPOTHR + TempRec.CData.PEXMPSENI) > 0 Then
            ErrorCode = 3
            GoSub SaveErrors
            GoTo CustErrors
          ElseIf RealValue = 0 And OldRound(TempRec.CData.REXMPOTHR + TempRec.CData.REXMPSENI) > 0 Then
            ErrorCode = 4
            GoSub SaveErrors
            GoTo CustErrors
          End If
          If Updated = False Then
            GoSub UpdateCurrCust
          End If
          If TempRec.CData.PROPVALU > 0 Then 'determine if this is a real property
            If TaxCust.FirstPropRec = 0 Then 'if this is the first property so far
              GoSub SaveRealProp
              NextRRec = NextRRec + 1
              RealRec.NextRec = 0
              TaxCust.FirstPropRec = NextRRec 'assign the customer link to it
              Put TCHandle, y, TaxCust
              Put RHandle, NextRRec, RealRec
              ConvRec.CountyAcct = 0
              ConvRec.CountyAcctString = QPTrim$(TaxCust.CountyAcctString)
              ConvRec.CustName = TempRec.CData.CustName
              ConvRec.CVALUE = 0
              ConvRec.MCVALUE = 0
              ConvRec.MHVALUE = 0
              ConvRec.MTVALUE = 0
              ConvRec.PersVal = 0
              ConvRec.PEXMPOTHR = 0
              ConvRec.PEXMPSENI = 0
              ConvRec.RPinNum = TempRec.CData.RPinNum
              ConvRec.PPinNum = TempRec.CData.PPinNum
              ConvRec.PROPVALU = TempRec.CData.PROPVALU
              ConvRec.REXMPOTHR = TempRec.CData.REXMPOTHR
              ConvRec.REXMPSENI = TempRec.CData.REXMPSENI
              TCnt = TCnt + 1
              Put CRHandle, TCnt, ConvRec
            Else
              GoSub SaveRealProp
              NextRRec = NextRRec + 1
              RealRec.NextRec = TaxCust.FirstPropRec
              TaxCust.FirstPropRec = NextRRec
              Put TCHandle, y, TaxCust
              Put RHandle, NextRRec, RealRec
              ConvRec.CountyAcct = 0
              ConvRec.CountyAcctString = QPTrim$(TaxCust.CountyAcctString)
              ConvRec.CustName = TempRec.CData.CustName
              ConvRec.CVALUE = 0
              ConvRec.MCVALUE = 0
              ConvRec.MHVALUE = 0
              ConvRec.MTVALUE = 0
              ConvRec.PersVal = 0
              ConvRec.PEXMPOTHR = 0
              ConvRec.PEXMPSENI = 0
              ConvRec.RPinNum = TempRec.CData.RPinNum
              ConvRec.PPinNum = TempRec.CData.PPinNum
              ConvRec.PROPVALU = TempRec.CData.PROPVALU
              ConvRec.REXMPOTHR = TempRec.CData.REXMPOTHR
              ConvRec.REXMPSENI = TempRec.CData.REXMPSENI
              TCnt = TCnt + 1
              Put CRHandle, TCnt, ConvRec
            End If
          End If
          If PersValue > 0 Then
            If TaxCust.FirstPersRec = 0 Then
              GoSub SavePersProp
              NextPRec = NextPRec + 1
              PersRec.NextRec = 0
              TaxCust.FirstPersRec = NextPRec 'assign the customer link to it
              Put TCHandle, y, TaxCust
              Put PHandle, NextPRec, PersRec
              ConvRec.CountyAcct = 0
              ConvRec.CountyAcctString = QPTrim$(TaxCust.CountyAcctString)
              ConvRec.CustName = TempRec.CData.CustName
              ConvRec.CVALUE = TempRec.CData.CVALUE
              ConvRec.MCVALUE = TempRec.CData.MCVALUE
              ConvRec.MHVALUE = TempRec.CData.MHVALUE
              ConvRec.MTVALUE = TempRec.CData.MTVALUE
              ConvRec.PersVal = TempRec.CData.PersVal
              ConvRec.PEXMPOTHR = TempRec.CData.PEXMPOTHR
              ConvRec.PEXMPSENI = TempRec.CData.PEXMPSENI
              ConvRec.RPinNum = TempRec.CData.RPinNum
              ConvRec.PPinNum = TempRec.CData.PPinNum
              ConvRec.PROPVALU = 0
              ConvRec.REXMPOTHR = 0
              ConvRec.REXMPSENI = 0
              TCnt = TCnt + 1
              Put CRHandle, TCnt, ConvRec
            Else
              GoSub SavePersProp
              NextPRec = NextPRec + 1
              PersRec.NextRec = TaxCust.FirstPersRec
              TaxCust.FirstPersRec = NextPRec
              Put TCHandle, y, TaxCust
              Put PHandle, NextPRec, PersRec
              ConvRec.CountyAcct = 0
              ConvRec.CountyAcctString = QPTrim$(TaxCust.CountyAcctString)
              ConvRec.CustName = TempRec.CData.CustName
              ConvRec.CVALUE = TempRec.CData.CVALUE
              ConvRec.MCVALUE = TempRec.CData.MCVALUE
              ConvRec.MHVALUE = TempRec.CData.MHVALUE
              ConvRec.MTVALUE = TempRec.CData.MTVALUE
              ConvRec.PersVal = TempRec.CData.PersVal
              ConvRec.PEXMPOTHR = TempRec.CData.PEXMPOTHR
              ConvRec.PEXMPSENI = TempRec.CData.PEXMPSENI
              ConvRec.RPinNum = TempRec.CData.RPinNum
              ConvRec.PPinNum = TempRec.CData.PPinNum
              ConvRec.PROPVALU = 0
              ConvRec.REXMPOTHR = 0
              ConvRec.REXMPSENI = 0
              TCnt = TCnt + 1
              Put CRHandle, TCnt, ConvRec
            End If
          End If
          Exit For
        End If 'TempCountyNum <> CustCountyNumber
      End If 'County Number = 0
    Next y
    If y > NumOfTCRecs Then
      If QPTrim$(TempRec.CData.CountyAcctString) <> "" Then
        For z = 1 To NewSCustCnt
          If QPTrim$(TempRec.CData.CountyAcctString) = NewSCust(z) Then
            Exit For
          End If
        Next z
        If z > NewSCustCnt Then
          NewSCustCnt = NewSCustCnt + 1
          ReDim Preserve NewSCust(1 To NewSCustCnt) As String
          NewSCust(NewSCustCnt) = QPTrim$(TempRec.CData.CountyAcctString)
        End If
      Else
         ErrorCode = 5
         TempRec.CData.CountyAcctString = "NOCNTYNUM" + CStr(ErrorCnt + 1)
         TempRec.CData.CountyAcct = -(ErrorCnt + 1)
         GoSub SaveErrors
         GoTo CustErrors
      End If
    End If
CustErrors:
    frmTCShowPctComp.ShowPctComp x, NumOfTempRecs
    If frmTCShowPctComp.Out = True Then
      Close
      frmTCShowPctComp.Out = False
      Unload frmTCShowPctComp
      Exit Sub
    End If
  Next x
  
  Close
  Unload frmTCShowPctComp
  
  If NewSCustCnt > 0 Then
    Call SaveNewSCust(NewSCustCnt, NewSCust())
  End If
  
  frmTCLoadingRpt.Show
  frmTCLoadingRpt.Label1.Caption = "Indexing...please wait"
  DoEvents
  
  Call CreateCustNameIdx
  Unload frmTCLoadingRpt
  frmTCLoadingRpt.Show
  frmTCLoadingRpt.Label1.Caption = "1 Of 4 completed"
  DoEvents
  Call CreateSrchNameIdx
  Unload frmTCLoadingRpt
  frmTCLoadingRpt.Show
  frmTCLoadingRpt.Label1.Caption = "2 Of 4 completed"
  DoEvents
  Call CreateOptCustIdx
  Unload frmTCLoadingRpt
  frmTCLoadingRpt.Show
  frmTCLoadingRpt.Label1.Caption = "3 Of 4 completed"
  DoEvents
  Call CreateSSIdx
  
  Unload frmTCLoadingRpt
  Call Savemsg(900, "Tax data has converted successfully.")
  
  Exit Sub
  
SaveRealProp:
  TempRPin = QPTrim$(TempRec.CData.RPinNum)
  For z = 1 To NumOfRealRecs
    Get OldRHandle, z, OldRealRec
    If QPTrim$(OldRealRec.RealPin) = TempRPin Then
      Exit For
    End If
  Next z
  
  RealRec.Blank = ""
  RealRec.CustPin = TaxCust.PIN
  RealRec.Deleted = 0
  RealRec.EXMPOTHR = TempRec.CData.REXMPOTHR
  RealRec.EXMPSENI = TempRec.CData.REXMPSENI
  RealRec.Fill1 = ""
  RealRec.LOTACRE = TempRec.CData.LOTACRE
  RealRec.PropSize = TempRec.CData.PropSize
  RealRec.PROPVALU = TempRec.CData.PROPVALU
  RealRec.RealPin = TempRec.CData.RPinNum
  RealRec.BLOCK = TempRec.CData.BLOCK
  RealRec.PROPNOT1 = TempRec.CData.RDESC1
  RealRec.PROPNOT2 = TempRec.CData.RDESC2
  RealRec.PROPNOT3 = TempRec.CData.RDESC3
  RealRec.Map = TempRec.CData.Map
  RealRec.LOTNUMB = TempRec.CData.LOTACRE
  
  If TempRec.CData.OptRev1Chrg <> 0 Then
    RealRec.OptRev1Chrg = 1
  Else
    RealRec.OptRev1Chrg = OldRealRec.OptRev1Chrg
  End If
  If TempRec.CData.OptRev2Chrg <> 0 Then
    RealRec.OptRev2Chrg = 2
  Else
    RealRec.OptRev2Chrg = OldRealRec.OptRev2Chrg
  End If
  If TempRec.CData.OptRev3Chrg <> 0 Then
    RealRec.OptRev3Chrg = 3
  Else
    RealRec.OptRev3Chrg = OldRealRec.OptRev3Chrg
  End If
  If QPTrim$(TempRec.CData.RealOptSearch) <> "" Then
    RealRec.OptSearch = QPTrim$(TempRec.CData.RealOptSearch)
  Else
    RealRec.OptSearch = QPTrim$(OldRealRec.OptSearch)
  End If
  If QPTrim$(TempRec.CData.LateList) <> "" Then
    RealRec.LateList = "Y"
  End If
  If QPTrim$(TempRec.CData.RTownShip) <> "" Then
    RealRec.TownShip = QPTrim$(TempRec.CData.RTownShip)
  End If
  If z <= NumOfRealRecs Then
    RealRec.InternalPin = OldRealRec.InternalPin
    RealRec.GISPOS = OldRealRec.GISPOS
    RealRec.ICPDesc = OldRealRec.ICPDesc
    RealRec.Image = OldRealRec.Image
    RealRec.LastYrPrinted = OldRealRec.LastYrPrinted
'    RealRec.OptRev1Chrg = OldRealRec.OptRev1Chrg
'    RealRec.OptRev2Chrg = OldRealRec.OptRev2Chrg
'    RealRec.OptRev3Chrg = OldRealRec.OptRev3Chrg
'    RealRec.OptSearch = OldRealRec.OptSearch
    If QPTrim$(TempRec.CData.LateList) = "" Then
      RealRec.LateList = OldRealRec.LateList
    End If
    RealRec.LienDesc = OldRealRec.LienDesc
    RealRec.Mock = OldRealRec.Mock
    If QPTrim$(TempRec.CData.MORTCODE) <> "" Then
      RealRec.MORTCODE = QPTrim$(TempRec.CData.MORTCODE)
    Else
      RealRec.MORTCODE = OldRealRec.MORTCODE
    End If
'    RealRec.MORTCODE = OldRealRec.MORTCODE
    RealRec.PropAddr = OldRealRec.PropAddr
    RealRec.PROPDATE = OldRealRec.PROPDATE
    RealRec.PROPDISC = OldRealRec.PROPDISC
'    RealRec.TownShip = OldRealRec.TownShip
    If QPTrim$(TempRec.CData.RTownShip) = "" Then
      RealRec.TownShip = OldRealRec.TownShip
    End If
  Else
    RealRec.InternalPin = IntRecCnt + 1
    Put IHandle, RealRec.InternalPin, IntPinRec
    RealRec.GISPOS = ""
    RealRec.ICPDesc = ""
    RealRec.Image = ""
    RealRec.LastYrPrinted = 0
    If QPTrim$(TempRec.CData.LateList) = "" Then
      RealRec.LateList = fpcmbRLateList.Text
    End If
    RealRec.LienDesc = ""
    RealRec.Mock = "N"
    RealRec.MORTCODE = QPTrim$(TempRec.CData.MORTCODE)
'    RealRec.OptRev1Chrg = 0
'    RealRec.OptRev2Chrg = 0
'    RealRec.OptRev3Chrg = 0
'    RealRec.OptSearch = ""
    RealRec.PropAddr = ""
    RealRec.PROPDATE = Date2Num(fptxtRPropDate.Text)
    RealRec.PROPDISC = "N"
'    RealRec.TownShip = ""
    If QPTrim$(TempRec.CData.RTownShip) = "" Then
      RealRec.TownShip = ""
    End If
  End If
  
  Return
  
SavePersProp:
  TempPPin = QPTrim$(TempRec.CData.PPinNum)
  For z = 1 To NumOfPersRecs
    Get OldPHandle, z, OldPersRec
    If QPTrim$(OldPersRec.PropPin) = TempPPin Then
      Exit For
    End If
  Next z
  
  PersRec.Blank = ""
  PersRec.CustPin = TaxCust.PIN
  PersRec.CVALUE = TempRec.CData.CVALUE
  PersRec.Deleted = 0
  PersRec.DESC1 = TempRec.CData.PDESC1
  PersRec.DESC2 = TempRec.CData.PDESC2
  PersRec.DESC3 = TempRec.CData.PDESC3
  PersRec.EXMPOTHR = TempRec.CData.PEXMPOTHR
  PersRec.EXMPSENI = TempRec.CData.PEXMPSENI
  PersRec.MCVALUE = TempRec.CData.MCVALUE
  PersRec.MHVALUE = TempRec.CData.MHVALUE
  PersRec.MTVALUE = TempRec.CData.MTVALUE
  PersRec.PersVal = TempRec.CData.PersVal
  PersRec.PropPin = TempRec.CData.PPinNum
  If QPTrim$(TempRec.CData.LateList) <> "" Then
    PersRec.LateList = "Y"
  End If
  If z <= NumOfPersRecs Then
    PersRec.Desc4 = OldPersRec.Desc4
    PersRec.Desc5 = OldPersRec.Desc5
    PersRec.DISCOV = OldPersRec.DISCOV
    PersRec.DMVSubmitted = OldPersRec.DMVSubmitted
    PersRec.InternalPin = OldPersRec.InternalPin
    PersRec.LastYrPrinted = OldPersRec.LastYrPrinted
    If QPTrim$(TempRec.CData.LateList) = "" Then
      PersRec.LateList = OldPersRec.LateList
    End If
    PersRec.PROPDATE = OldPersRec.PROPDATE
    PersRec.VehTaxYear = OldPersRec.VehTaxYear
  Else
    PersRec.Desc4 = ""
    PersRec.Desc5 = ""
    PersRec.DISCOV = "N"
    PersRec.DMVSubmitted = "N"
    PersRec.InternalPin = IntRecCnt + 1
    Put IHandle, PersRec.InternalPin, IntPinRec
    PersRec.LastYrPrinted = 0
    If QPTrim$(TempRec.CData.LateList) = "" Then
      PersRec.LateList = fpcmbPLateList.Text
    End If
    PersRec.PROPDATE = OldPersRec.PROPDATE
    PersRec.VehTaxYear = 0
  End If
  
  Return
  
UpdateCurrCust:
  Updated = True
  TaxCust.Deleted = 0
  If QPTrim$(TempRec.CData.Addr1) <> "" Then
    TaxCust.Addr1 = TempRec.CData.Addr1
  End If
  If QPTrim$(TempRec.CData.Addr2) <> "" Then
    TaxCust.Addr2 = TempRec.CData.Addr2
  End If
  If QPTrim$(TempRec.CData.City) <> "" Then
    TaxCust.City = TempRec.CData.City
  End If
  If QPTrim$(TempRec.CData.CustName) <> "" Then
    TaxCust.CustName = TempRec.CData.CustName
  End If
  If QPTrim$(TempRec.CData.Zip) <> "" Then
    TaxCust.Zip = TempRec.CData.Zip
  End If
  If QPTrim$(TempRec.CData.State) <> "" Then
    TaxCust.State = TempRec.CData.State
  End If
  If QPTrim$(TempRec.CData.SName) <> "" Then
    TaxCust.SName = TempRec.CData.SName
  End If
  If QPTrim$(TempRec.CData.OptSrchDesc) <> "" Then
    TaxCust.OptSrchDesc = TempRec.CData.OptSrchDesc
  End If
  If QPTrim$(TempRec.CData.County4BillName) <> "" Then
    TaxCust.County4BillName = TempRec.CData.County4BillName
  End If
  TaxCust.Active = "Y"
  If TempRec.CData.Cycle > 0 Then
    TaxCust.Cycle = TempRec.CData.Cycle
  End If
  If QPTrim$(TempRec.CData.CycleName) <> "" Then
    TaxCust.CycleName = QPTrim$(TempRec.CData.CycleName)
  End If
  If QPTrim$(TempRec.CData.CTownShip) <> "" Then
    TaxCust.TownShip = QPTrim$(TempRec.CData.CTownShip)
  End If
  Put TCHandle, y, TaxCust
  Return

SaveErrors:
  ErrorCnt = ErrorCnt + 1
  ErrorRec.CountyAcct = TempRec.CData.CountyAcct
  ErrorRec.CountyAcctString = QPTrim$(TempRec.CData.CountyAcctString)
  ErrorRec.CustName = QPTrim$(TempRec.CData.CustName)
  ErrorRec.ErrorType = ErrorCode
  ErrorRec.PersTot = PersValue
  ErrorRec.PersXTot = OldRound(TempRec.CData.PEXMPOTHR + TempRec.CData.PEXMPSENI)
  ErrorRec.RPinNum = QPTrim$(TempRec.CData.RPinNum)
  ErrorRec.PPinNum = QPTrim$(TempRec.CData.PPinNum)
  ErrorRec.RealTot = RealValue
  ErrorRec.RealXTot = OldRound(TempRec.CData.REXMPOTHR + TempRec.CData.REXMPSENI)
  Put EHandle, ErrorCnt, ErrorRec
  
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTCConvert", "ProcessRMultiple", Erl)
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

Private Sub SaveNewSCust(NewSCustCnt As Long, NewSCust() As String)
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim x As Long, y As Long, z As Long
  Dim TempHandle As Integer
  Dim TempRec As TempConversionData
  Dim NumOfTempRecs As Long
  Dim IntPinRec As InternalPinType
  Dim IHandle As Integer
  Dim NumOfIRecs As Long
  Dim IntCnt As Long
  Dim ThisCoNum$
  Dim CustCnt As Long
  Dim RealCnt As Long
  Dim PersCnt As Long
  Dim FirstTime As Boolean
  Dim PersValue As Double
  Dim ConvRec As ConvResultsType
  Dim CRHandle As Integer
  Dim NumOfCRRecs As Long
  Dim RealValue As Double
  Dim NumOfErrors As Long
  Dim ErrorCnt As Long
  Dim ErrorRec As ConvErrorType
  Dim EHandle As Integer
  Dim NumOfERecs As Long
  Dim ErrorCode As Integer
  Dim ThisPPin$
  Dim ThisRPin$
  
  On Error GoTo ERRORSTUFF
  
  OpenTempConvFile TempHandle, NumOfTempRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenRealPropFile RHandle, NumOfRealRecs
  OpenPersPropFile PHandle, NumOfPersRecs
  OpenIntPinFile IHandle, NumOfIRecs
  OpenConvResultsFile CRHandle, NumOfCRRecs
  OpenConvErrorsFile EHandle, NumOfERecs
  
  ErrorCnt = NumOfERecs
  TCnt = NumOfCRRecs
  
  RealCnt = NumOfRealRecs
  PersCnt = NumOfPersRecs
  CustCnt = NumOfTCRecs
  IntCnt = NumOfIRecs
  frmTCShowPctComp.Label1 = "Adding New Customers"
  frmTCShowPctComp.Show , Me
  
  For x = 1 To NewSCustCnt
    ThisCoNum = QPTrim$(NewSCust(x))
    FirstTime = True
    For y = 1 To NumOfTempRecs
      Get TempHandle, y, TempRec
      PersValue = OldRound(TempRec.CData.CVALUE + TempRec.CData.MCVALUE + TempRec.CData.MHVALUE + TempRec.CData.PersVal + TempRec.CData.MTVALUE)
      RealValue = TempRec.CData.PROPVALU
      If QPTrim$(TempRec.CData.CountyAcctString) = ThisCoNum Then
        ThisPPin = QPTrim$(TempRec.CData.PPinNum)
        ThisRPin = QPTrim$(TempRec.CData.RPinNum)
        If (PersValue > 0 And (ThisPPin = "" Or ThisPPin = "0")) And (RealValue > 0 And (ThisRPin = "" Or ThisRPin = "0")) Then
          ErrorCode = 1
          GoSub SaveErrors
          GoTo CustErrors
        ElseIf PersValue = 0 And RealValue = 0 Then
          ErrorCode = 2
          GoSub SaveErrors
          GoTo CustErrors
        ElseIf PersValue = 0 And OldRound(TempRec.CData.PEXMPOTHR + TempRec.CData.PEXMPSENI) > 0 Then
          ErrorCode = 3
          GoSub SaveErrors
          GoTo CustErrors
        ElseIf RealValue = 0 And OldRound(TempRec.CData.REXMPOTHR + TempRec.CData.REXMPSENI) > 0 Then
          ErrorCode = 4
          GoSub SaveErrors
          GoTo CustErrors
        End If
        If FirstTime = True Then
          If PersValue > 0 Then
            PersCnt = PersCnt + 1
            GoSub SavePersData
            PersRec.NextRec = 0
            Put PHandle, PersCnt, PersRec
            TaxCust.FirstPropRec = 0
            TaxCust.FirstPersRec = PersCnt
            ConvRec.CountyAcct = TempRec.CData.CountyAcct
            ConvRec.CountyAcctString = QPTrim$(TempRec.CData.CountyAcctString)
            ConvRec.CustName = TempRec.CData.CustName
            ConvRec.CVALUE = TempRec.CData.CVALUE
            ConvRec.MCVALUE = TempRec.CData.MCVALUE
            ConvRec.MHVALUE = TempRec.CData.MHVALUE
            ConvRec.MTVALUE = TempRec.CData.MTVALUE
            ConvRec.PersVal = TempRec.CData.PersVal
            ConvRec.PEXMPOTHR = TempRec.CData.PEXMPOTHR
            ConvRec.PEXMPSENI = TempRec.CData.PEXMPSENI
            ConvRec.PPinNum = TempRec.CData.PPinNum
            ConvRec.RPinNum = TempRec.CData.RPinNum
            ConvRec.PROPVALU = 0
            ConvRec.REXMPOTHR = 0
            ConvRec.REXMPSENI = 0
            TCnt = TCnt + 1
            Put CRHandle, TCnt, ConvRec
          End If
          If TempRec.CData.PROPVALU > 0 Then
            RealCnt = RealCnt + 1
            GoSub SaveRealData
            RealRec.NextRec = 0
            Put RHandle, RealCnt, RealRec
            TaxCust.FirstPersRec = 0
            TaxCust.FirstPropRec = RealCnt
            ConvRec.CountyAcct = TempRec.CData.CountyAcct
            ConvRec.CountyAcctString = QPTrim$(TempRec.CData.CountyAcctString)
            ConvRec.CustName = TempRec.CData.CustName
            ConvRec.CVALUE = 0
            ConvRec.MCVALUE = 0
            ConvRec.MHVALUE = 0
            ConvRec.MTVALUE = 0
            ConvRec.PersVal = 0
            ConvRec.PEXMPOTHR = 0
            ConvRec.PEXMPSENI = 0
            ConvRec.PPinNum = TempRec.CData.PPinNum
            ConvRec.RPinNum = TempRec.CData.RPinNum
            ConvRec.PROPVALU = TempRec.CData.PROPVALU
            ConvRec.REXMPOTHR = TempRec.CData.REXMPOTHR
            ConvRec.REXMPSENI = TempRec.CData.REXMPSENI
            TCnt = TCnt + 1
            Put CRHandle, TCnt, ConvRec
          End If
          GoSub SaveCustData
          Put TCHandle, CustCnt, TaxCust
          FirstTime = False
        Else
          If PersValue > 0 Then
            PersCnt = PersCnt + 1
            GoSub SavePersData
            PersRec.NextRec = TaxCust.FirstPersRec
            Put PHandle, PersCnt, PersRec
            TaxCust.FirstPersRec = PersCnt
            Put TCHandle, CustCnt, TaxCust
            ConvRec.CountyAcct = TempRec.CData.CountyAcct
            ConvRec.CountyAcctString = QPTrim$(TempRec.CData.CountyAcctString)
            ConvRec.CustName = TempRec.CData.CustName
            ConvRec.CVALUE = TempRec.CData.CVALUE
            ConvRec.MCVALUE = TempRec.CData.MCVALUE
            ConvRec.MHVALUE = TempRec.CData.MHVALUE
            ConvRec.MTVALUE = TempRec.CData.MTVALUE
            ConvRec.PersVal = TempRec.CData.PersVal
            ConvRec.PEXMPOTHR = TempRec.CData.PEXMPOTHR
            ConvRec.PEXMPSENI = TempRec.CData.PEXMPSENI
            ConvRec.PPinNum = TempRec.CData.PPinNum
            ConvRec.RPinNum = TempRec.CData.RPinNum
            ConvRec.PROPVALU = 0
            ConvRec.REXMPOTHR = 0
            ConvRec.REXMPSENI = 0
            TCnt = TCnt + 1
            Put CRHandle, TCnt, ConvRec
          End If
          If TempRec.CData.PROPVALU > 0 Then
            RealCnt = RealCnt + 1
            GoSub SaveRealData
            RealRec.NextRec = TaxCust.FirstPropRec
            Put RHandle, RealCnt, RealRec
            TaxCust.FirstPropRec = RealCnt
            Put TCHandle, CustCnt, TaxCust
            ConvRec.CountyAcct = TempRec.CData.CountyAcct
            ConvRec.CountyAcctString = QPTrim$(TempRec.CData.CountyAcctString)
            ConvRec.CustName = TempRec.CData.CustName
            ConvRec.CVALUE = 0
            ConvRec.MCVALUE = 0
            ConvRec.MHVALUE = 0
            ConvRec.MTVALUE = 0
            ConvRec.PersVal = 0
            ConvRec.PEXMPOTHR = 0
            ConvRec.PEXMPSENI = 0
            ConvRec.PPinNum = TempRec.CData.PPinNum
            ConvRec.RPinNum = TempRec.CData.RPinNum
            ConvRec.PROPVALU = TempRec.CData.PROPVALU
            ConvRec.REXMPOTHR = TempRec.CData.REXMPOTHR
            ConvRec.REXMPSENI = TempRec.CData.REXMPSENI
            TCnt = TCnt + 1
            Put CRHandle, TCnt, ConvRec
          End If
        End If
      End If
    Next y
CustErrors:
    frmTCShowPctComp.ShowPctComp x, NewSCustCnt
    If frmTCShowPctComp.Out = True Then
      Close
      frmTCShowPctComp.Out = False
      Unload frmTCShowPctComp
      Exit Sub
    End If
  Next x
  Unload frmTCShowPctComp
  
  Close
  
  Exit Sub

SaveCustData:
  CustCnt = CustCnt + 1
  TaxCust.Acct = CustCnt
  TaxCust.Active = "Y"
  TaxCust.Addr1 = QPTrim$(TempRec.CData.Addr1)
  TaxCust.Addr2 = QPTrim$(TempRec.CData.Addr2)
  TaxCust.Bankrupt = "N"
  TaxCust.City = QPTrim$(TempRec.CData.City)
  TaxCust.County4BillName = QPTrim$(TempRec.CData.County4BillName)
  TaxCust.County4BillNum = 0
  TaxCust.CountyAcct = TempRec.CData.CountyAcct
  TaxCust.CountyAcctString = QPTrim$(TempRec.CData.CountyAcctString)
  TaxCust.CSSN = QPTrim$(TempRec.CData.CSSN)
  TaxCust.CustName = QPTrim$(TempRec.CData.CustName)
  If TempRec.CData.Cycle > 0 Then
    TaxCust.Cycle = TempRec.CData.Cycle
  Else
    TaxCust.Cycle = 0
  End If
  If QPTrim$(TempRec.CData.CycleName) <> "" Then
    TaxCust.CycleName = QPTrim$(TempRec.CData.CycleName)
  Else
    TaxCust.CycleName = ""
  End If
  TaxCust.Deleted = 0
  TaxCust.DeliveryPt = ""
  TaxCust.DrvrsLic = ""
  TaxCust.Employer = ""
  TaxCust.FileVer = FileVers
  If PersValue > 0 Then
    TaxCust.FirstPersRec = PersCnt
  ElseIf TempRec.CData.PROPVALU > 0 Then
    TaxCust.FirstPropRec = RealCnt
  End If
  TaxCust.HPHONE = ""
  TaxCust.Interest = fpcmbChrgInt.Text
  TaxCust.LateNotice = fpcmbLateList.Text
  TaxCust.OPENDATE = Date2Num(fptxtOpenDate.Text)
  TaxCust.OptSrchDesc = TempRec.CData.OptSrchDesc
  TaxCust.OSSN = QPTrim$(TempRec.CData.OSSN)
  TaxCust.Pad1 = ""
  TaxCust.Penalty = fpcmbPenalty.Text
  TaxCust.PIN = CustCnt
  TaxCust.PostalRt = ""
  TaxCust.PrePayBal = 0
  TaxCust.PrePayTrans = 0
  TaxCust.ServiceAdd = ""
  TaxCust.SName = TempRec.CData.SName
  TaxCust.State = TempRec.CData.State
  TaxCust.TaxExempt = fpcmbTaxEx.Text
  If QPTrim$(TempRec.CData.CTownShip) <> "" Then
    TaxCust.TownShip = QPTrim$(TempRec.CData.CTownShip)
  Else
    TaxCust.TownShip = ""
  End If
'  TaxCust.TownShip = ""
  TaxCust.WPHONE = ""
  TaxCust.Zip = TempRec.CData.Zip
  
  Return
  
SaveRealData:
  RealRec.Blank = ""
  RealRec.BLOCK = TempRec.CData.BLOCK
'  RealRec.CustPin = CustCnt
  If FirstTime = True Then
    RealRec.CustPin = CustCnt + 1 'added + 1 7/14/06
  Else
    RealRec.CustPin = CustCnt
  End If
  RealRec.Deleted = 0
  RealRec.EXMPOTHR = TempRec.CData.REXMPOTHR
  RealRec.EXMPSENI = TempRec.CData.REXMPSENI
  RealRec.Fill1 = ""
  RealRec.GISPOS = ""
  RealRec.ICPDesc = ""
  RealRec.Image = ""
  IntCnt = IntCnt + 1
  RealRec.InternalPin = IntCnt
  IntPinRec.PIN = RealCnt
  Put IHandle, IntCnt, IntPinRec
  RealRec.LastYrPrinted = 0
  If QPTrim$(TempRec.CData.LateList) <> "" Then
    RealRec.LateList = "Y"
  Else
    RealRec.LateList = fpcmbRLateList.Text
  End If
  RealRec.LienDesc = ""
  RealRec.LienYN = "N"
  RealRec.LOTACRE = TempRec.CData.LOTACRE
  RealRec.LOTNUMB = ""
  RealRec.Map = TempRec.CData.Map
  RealRec.LOTNUMB = TempRec.CData.LOTACRE
  RealRec.Mock = "N"
  RealRec.MORTCODE = QPTrim$(TempRec.CData.MORTCODE)
  If TempRec.CData.OptRev1Chrg <> 0 Then
    RealRec.OptRev1Chrg = 1
  End If
  If TempRec.CData.OptRev2Chrg <> 0 Then
    RealRec.OptRev1Chrg = 2
  End If
  If TempRec.CData.OptRev3Chrg <> 0 Then
    RealRec.OptRev3Chrg = 3
  End If
  RealRec.OptSearch = QPTrim$(TempRec.CData.RealOptSearch)
  RealRec.PropAddr = ""
  RealRec.PROPDATE = Date2Num(fptxtRPropDate.Text)
  RealRec.PROPDISC = "N"
  RealRec.PROPNOT1 = TempRec.CData.RDESC1
  RealRec.PROPNOT2 = TempRec.CData.RDESC2
  RealRec.PROPNOT3 = TempRec.CData.RDESC3
  RealRec.PropSize = TempRec.CData.PropSize
  RealRec.PROPVALU = TempRec.CData.PROPVALU
  RealRec.RealPin = TempRec.CData.RPinNum
  If QPTrim$(TempRec.CData.RTownShip) <> "" Then
    RealRec.TownShip = QPTrim$(TempRec.CData.RTownShip)
  Else
    RealRec.TownShip = ""
  End If
'  RealRec.TownShip = ""
  
  Return
  
SavePersData:
  PersRec.Blank = ""
'  PersRec.CustPin = CustCnt
  If FirstTime = True Then
    PersRec.CustPin = CustCnt + 1 'added + 1 7/14/06
  Else
    PersRec.CustPin = CustCnt
  End If
  PersRec.CVALUE = TempRec.CData.CVALUE
  PersRec.Deleted = 0
  PersRec.DESC1 = TempRec.CData.PDESC1
  PersRec.DESC2 = TempRec.CData.PDESC2
  PersRec.DESC3 = TempRec.CData.PDESC3
  PersRec.Desc4 = ""
  PersRec.Desc5 = ""
  PersRec.DISCOV = "N"
  PersRec.DMVSubmitted = "N"
  PersRec.EXMPOTHR = TempRec.CData.PEXMPOTHR
  PersRec.EXMPSENI = TempRec.CData.PEXMPSENI
  IntCnt = IntCnt + 1
  PersRec.InternalPin = IntCnt
  IntPinRec.PIN = PersCnt
  Put IHandle, IntCnt, IntPinRec
  PersRec.LastYrPrinted = 0
  If QPTrim$(TempRec.CData.LateList) <> "" Then
    PersRec.LateList = "Y"
  Else
    PersRec.LateList = fpcmbPLateList.Text
  End If
  PersRec.MCVALUE = TempRec.CData.MCVALUE
  PersRec.MHVALUE = TempRec.CData.MHVALUE
  PersRec.MTVALUE = TempRec.CData.MTVALUE
  PersRec.PersVal = TempRec.CData.PersVal
  PersRec.PROPDATE = Date2Num(fptxtPPropDate.Text)
  PersRec.PropPin = TempRec.CData.PPinNum
  PersRec.VehTaxYear = 0
  
  Return
  
SaveErrors:
  ErrorCnt = ErrorCnt + 1
  ErrorRec.CountyAcct = TempRec.CData.CountyAcct
  ErrorRec.CountyAcctString = QPTrim$(TempRec.CData.CountyAcctString)
  ErrorRec.CustName = QPTrim$(TempRec.CData.CustName)
  ErrorRec.ErrorType = ErrorCode
  ErrorRec.PersTot = PersValue
  ErrorRec.PersXTot = OldRound(TempRec.CData.PEXMPOTHR + TempRec.CData.PEXMPSENI)
  ErrorRec.PPinNum = QPTrim$(TempRec.CData.PPinNum)
  ErrorRec.RPinNum = QPTrim$(TempRec.CData.RPinNum)
  ErrorRec.RealTot = RealValue
  ErrorRec.RealXTot = OldRound(TempRec.CData.REXMPOTHR + TempRec.CData.REXMPSENI)
  Put EHandle, ErrorCnt, ErrorRec
  
  Return
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTCConvert", "SaveNewSCust", Erl)
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

