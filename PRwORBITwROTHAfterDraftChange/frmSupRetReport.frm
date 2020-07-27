VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmSupRetReport 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supplemental Retirement Report"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmSupRetReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint2 
      Height          =   7452
      Left            =   2160
      TabIndex        =   8
      Top             =   696
      Width           =   7356
      _Version        =   196609
      _ExtentX        =   12975
      _ExtentY        =   13144
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmSupRetReport.frx":08CA
      Begin LpLib.fpCombo fpcomboDiskFile 
         Height          =   405
         Left            =   6240
         TabIndex        =   6
         ToolTipText     =   "This file will appear in the 401K directory in the Citipak directory"
         Top             =   5250
         Width           =   780
         _Version        =   196608
         _ExtentX        =   1376
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
         ColDesigner     =   "frmSupRetReport.frx":08E6
      End
      Begin LpLib.fpCombo fpcomboLPDed 
         Height          =   405
         Left            =   3750
         TabIndex        =   3
         Top             =   2925
         Width           =   2985
         _Version        =   196608
         _ExtentX        =   5265
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
         ColDesigner     =   "frmSupRetReport.frx":0BDD
      End
      Begin LpLib.fpCombo fpcomboVolDed 
         Height          =   405
         Left            =   3750
         TabIndex        =   2
         Top             =   2355
         Width           =   2985
         _Version        =   196608
         _ExtentX        =   5265
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
         ColDesigner     =   "frmSupRetReport.frx":0ED4
      End
      Begin LpLib.fpCombo fpcomboPrintOpt 
         Height          =   405
         Left            =   3750
         TabIndex        =   7
         Top             =   5835
         Width           =   2355
         _Version        =   196608
         _ExtentX        =   4154
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
         ColDesigner     =   "frmSupRetReport.frx":11CB
      End
      Begin LpLib.fpCombo fpcmbRoth 
         Height          =   405
         Left            =   3735
         TabIndex        =   20
         Top             =   4680
         Width           =   2970
         _Version        =   196608
         _ExtentX        =   5239
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
         ColDesigner     =   "frmSupRetReport.frx":14C2
      End
      Begin EditLib.fpDateTime fptxtStart 
         Height          =   368
         Left            =   3744
         TabIndex        =   0
         Top             =   1296
         Width           =   1836
         _Version        =   196608
         _ExtentX        =   3238
         _ExtentY        =   649
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
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDateTime fptxtEnd 
         Height          =   348
         Left            =   3744
         TabIndex        =   1
         Top             =   1824
         Width           =   1836
         _Version        =   196608
         _ExtentX        =   3238
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
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   696
         Left            =   4368
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Press F10 to generate the Supplemental Retirement report."
         Top             =   6504
         Width           =   1884
         _Version        =   131072
         _ExtentX        =   3323
         _ExtentY        =   1228
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
         ButtonDesigner  =   "frmSupRetReport.frx":17B9
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
         Height          =   696
         Left            =   1344
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Press ESC to exit this screen."
         Top             =   6504
         Width           =   1896
         _Version        =   131072
         _ExtentX        =   3344
         _ExtentY        =   1228
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
         ButtonDesigner  =   "frmSupRetReport.frx":1998
      End
      Begin EditLib.fpDoubleSingle fptxtCGMRate 
         Height          =   375
         Left            =   3720
         TabIndex        =   4
         Top             =   3480
         Width           =   1095
         _Version        =   196608
         _ExtentX        =   1931
         _ExtentY        =   661
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
         DecimalPlaces   =   -1
         DecimalPoint    =   "."
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
      Begin EditLib.fpDoubleSingle fptxtCLMRate 
         Height          =   375
         Left            =   3720
         TabIndex        =   5
         Top             =   4080
         Width           =   1095
         _Version        =   196608
         _ExtentX        =   1931
         _ExtentY        =   661
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
         DecimalPlaces   =   -1
         DecimalPoint    =   "."
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
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Roth Deduction:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   21
         Top             =   4728
         Width           =   3036
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Print Option:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   1632
         TabIndex        =   17
         Top             =   5928
         Width           =   1500
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Prepare Supplemental Retirement Submission Report?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   16
         Top             =   5352
         Width           =   5916
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Code L Matching Rate:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   768
         TabIndex        =   15
         Top             =   4176
         Width           =   2652
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Code G Matching Rate:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   720
         TabIndex        =   14
         Top             =   3600
         Width           =   2700
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Voluntary Deduction:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   432
         TabIndex        =   13
         Top             =   2448
         Width           =   2988
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Loan Payment Deduction:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   384
         TabIndex        =   12
         Top             =   2976
         Width           =   3036
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   732
         Left            =   1104
         Top             =   240
         Width           =   5340
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "End Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   1392
         TabIndex        =   11
         Top             =   1920
         Width           =   2028
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Start Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   1392
         TabIndex        =   10
         Top             =   1440
         Width           =   2028
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Supplemental Retirement Report"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   444
         Left            =   1344
         TabIndex        =   9
         Top             =   432
         Width           =   4908
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   4
      Height          =   7812
      Left            =   1980
      Top             =   528
      Width           =   7692
   End
End
Attribute VB_Name = "frmSupRetReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class
Dim HaveFlag As Boolean

Private Sub cmdCreateMediaFile_Click()
  Call VertMenu401
End Sub

Private Sub cmdEscape_Click()
   frmReportsProcessing.Show
   DoEvents
   Unload frmSupRetReport
End Sub
 
Private Sub cmdProcess_Click()
  
'  On Error GoTo ErrorHandler1
  If fpcomboPrintOpt.Text = "Graphical" Then
    Call PrintGraphics
    Exit Sub
  ElseIf fpcomboPrintOpt.Text = "Text" Then
    Call PrintText
    Exit Sub
  Else
    Exit Sub
  End If
  
ErrorHandler1:
  Close
  MsgBox "ERROR: If this problem persists then call Southern Software."
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      Call cmdEscape_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdProcess_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Call LoadSupRetReportScreen
  Me.HelpContextID = hlpSupplemental
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub

Private Sub LoadSupRetReportScreen()
  Dim Today As String * 10
  Dim UnitHandle As Integer, Unit(1) As UnitFileRecType
  Dim DedCodeFileHandle As Integer, x As Integer, DedRecCnt As Integer
  Dim DedCodeFileRec As DedCodeRecType
  Dim City As String
  Dim NumOfRecs As Long
  Dim EmpData2FileHandle As Integer
  Dim EmpRecSize As EmpData2Type
  Dim K401Rec As K401DedType
  Dim KHandle As Integer
  
  If Exist(K401DedName) Then
    Open401KDedFile KHandle
    Get KHandle, 1, K401Rec
    Close
    'employer edit screen where the user can enter a permanent deduction
    If QPTrim$(K401Rec.VolDed) = "Unsaved" Or QPTrim$(K401Rec.VolDed) = "" Then GoTo UnSavedVol 'Unsaved comes from the
'    If CInt(Mid(K401Rec.VolDed, 1, 1)) < 10 Then
'      fpcomboVolDed.Text = " " + QPTrim$(K401Rec.VolDed)
'    Else
'      fpcomboVolDed.Text = QPTrim$(K401Rec.VolDed)
'    End If
    If Mid(K401Rec.VolDed, 2, 1) = ")" Then
      fpcomboVolDed.Text = " " + QPTrim$(K401Rec.VolDed)
    ElseIf Mid(K401Rec.VolDed, 3, 1) = ")" Then
      fpcomboVolDed.Text = QPTrim$(K401Rec.VolDed)
    End If
UnSavedVol:
    If QPTrim$(K401Rec.LPDed) = "Unsaved" Or QPTrim$(K401Rec.LPDed) = "" Then GoTo UnSavedLP 'Unsaved comes from the
'    If CInt(Mid(K401Rec.LPDed, 1, 1)) < 10 Then
'      fpcomboLPDed.Text = " " + QPTrim$(K401Rec.LPDed)
'    Else
'      fpcomboLPDed.Text = QPTrim$(K401Rec.LPDed)
'    End If
    If Mid(K401Rec.LPDed, 2, 1) = ")" Then
      fpcomboLPDed.Text = " " + QPTrim$(K401Rec.LPDed)
    ElseIf Mid(K401Rec.LPDed, 3, 1) = ")" Then
      fpcomboLPDed.Text = QPTrim$(K401Rec.LPDed)
    End If
UnSavedLP:
    If QPTrim$(K401Rec.RothDed) = "Unsaved" Or QPTrim$(K401Rec.RothDed) = "" Then
      fpcmbRoth.Text = " 0) No Roth"
      GoTo UnSavedRoth
    End If
    If Mid(K401Rec.RothDed, 2, 1) = ")" Then
      fpcmbRoth.Text = " " + QPTrim$(K401Rec.RothDed)
    ElseIf Mid(K401Rec.RothDed, 3, 1) = ")" Then
      fpcmbRoth.Text = QPTrim$(K401Rec.RothDed)
    End If
UnSavedRoth:
  End If
  
  Today = Date '$
  fptxtStart.Text = Mid(Today, 1, 2) + "-01-" + Mid(Today, 7, 4) '8/21
  'changed from Today
  fptxtEnd.Text = Today
   
  fpcomboPrintOpt.AddItem "Graphical"
  fpcomboPrintOpt.AddItem "Text"
  fpcomboPrintOpt.Text = "Graphical"
  HaveFlag = False
  OpenDedCodeFile DedCodeFileHandle
  DedRecCnt = LOF(DedCodeFileHandle) / Len(DedCodeFileRec)
  If DedRecCnt = 0 Then
    MsgBox "No deduction records on file"
    fpcomboVolDed.Text = "No deduction records on file."
    fpcomboLPDed.Text = "No deduction records on file."
    fpcmbRoth.Text = "No deduction records on file."
    cmdProcess.Enabled = False
    Close
    Exit Sub
  End If
  fpcomboDiskFile.Text = "N"
  fpcomboDiskFile.AddItem "Y"
  fpcomboDiskFile.AddItem "N"
  For x = 1 To DedRecCnt
    Get DedCodeFileHandle, x, DedCodeFileRec
      If x < 10 Then
        fpcomboVolDed.AddItem " " + CStr(x) + ") " + (DedCodeFileRec.DCDESC1)
        fpcomboLPDed.AddItem " " + CStr(x) + ") " + (DedCodeFileRec.DCDESC1)
        fpcmbRoth.AddItem " " + CStr(x) + ") " + (DedCodeFileRec.DCDESC1)
      Else
        fpcomboVolDed.AddItem CStr(x) + ") " + (DedCodeFileRec.DCDESC1)
        fpcomboLPDed.AddItem CStr(x) + ") " + (DedCodeFileRec.DCDESC1)
        fpcmbRoth.AddItem CStr(x) + ") " + (DedCodeFileRec.DCDESC1)
      End If
  Next x
  Close DedCodeFileHandle
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, Unit(1)
  Close UnitHandle
  City = QPTrim(Unit(1).UFEMPR)
  
  fptxtCGMRate = Unit(1).GMatch401K
  fptxtCLMRate = Unit(1).LMatch401K
  
  OpenEmpData2File EmpData2FileHandle
  NumOfRecs = LOF(EmpData2FileHandle) \ Len(EmpRecSize)
  Close EmpData2FileHandle
  If NumOfRecs = 0 Then
    MsgBox "No records on file"
    fpcomboVolDed.Text = "No records on file. Please exit."
    fpcomboLPDed.Text = "No records on file. Please exit."
    fpcmbRoth.Text = "No records on file. Please exit."
    Close
    Exit Sub
  End If

End Sub
Private Sub PrintGraphics()
  Dim EmpRecSize As Long, TRecSize As Long, PrnDef$
  Dim RptName$, DedCodeHandle As Integer, x As Integer
  Dim DebugFlag As Boolean, Command$, NumOfRecs&
  Dim Image1$, Image2$, Image3$, DepSize As Integer
  Dim UnitHandle As Integer, IdxRecLen As Integer
  Dim Text1Len%, Text2Len%, RecNo As Long, IdxFileSize&
  Dim Text3Len%, Text4Len%, Choice$(), EPRTotal#
  Dim cnt As Long, GCntNum As Long, UsingThisOne As Boolean
  Dim GCntName As Long, LCntName As Long
  Dim EPct#
  Dim LowDate&, HighDate&, GPct#, LPct#, TempDed$
  Dim VCodeNum As Integer, LCodeNum As Integer, RCodeNum As Integer
  Dim VCalcAmt#, LCalcAmt#, GCalcAmt#
  Dim EmpFile As Integer, NumOfRec&, EmpRet$, RHandle As Integer, THandle As Integer
  Dim DHandle As Integer, EmpRType$, LCntNum&, TransRecNum&
  Dim TMatchAmt#, TotalMatchAmt#, TotalVAmt#, TotalLAmt#, TotalRoth#
  Dim EPrinted As Integer, TempDate$, SubTotalMatchAmt#, SubTotalVAmt#, SubTotalRAmt#
  Dim SubTotalLAmt#, CurrEmpNo$, PntCnt As Integer, TotalRAmt#
  Dim Totals$, Ret$, EmpGBuffNum() As EmpSortType, GVTotal#, GLTotal#
  Dim TVLawAmt#, TLLawAmt#, LawMatAmt#, TotLawMatAmt#, TRothAmt#
  Dim TVGenAmt#, TLGenAmt#, GenMatAmt#, TotGenMatAmt#
  Dim offset As Integer, TotalGross#, RptTitle$
  Dim ThisSort() As Integer
  Dim TempEmpNo As Long, TempRecNo As Long, y As Long
  Dim Largest As Long, SortedRecNo As Long, swapThis As Long
  Dim Smallest As Long, DedRecCnt As Long, AlphaFlag As Boolean
  Dim LCnt As Long, GCnt As Long, EmpIdxLNameHandle As Integer
  Dim Unit(1) As UnitFileRecType, Text1 As String
  Dim Text2 As String, Text3 As String, Text4 As String
  Dim LNum As Long, GNum As Long, GLArray() As K401RptType
  Dim z As Long, five As String, bigName As String, smallName As String
  Dim q As Long, smallIdx As Long, LFlag As Boolean, GFlag As Boolean
  Dim tempEmpNum$, TempEmpName$, tempSSN$, TempVAmt$
  Dim TempLAmt$, TempMAmt$, tempBatch$, tempHDate$
  Dim TempGross$, tempRetType$, thisNum As String, Temp401K As K401RptType
  Dim ThisName As String, ThisEmp As String, TempNo As String
  Dim TempIdx As Long, Justx As Long, oldFive As String, Lfive$
  Dim EmpInDept As Integer
  Dim EmpData2FileHandle As Integer
  Dim Emp2Rec As EmpData2Type
  Dim Month$, FirstTime As Boolean
  Dim AllCnt As Integer, RunCnt As Integer
  Dim newFive As Long, oldOffset As Long
  Dim dlm As String
  Dim UnitRec As UnitFileRecType
  Dim UHandle As Integer
  Dim GenLaw$, TEPrinted As Integer
  ReDim TransHRec(1) As TransRecType
  ReDim K401Rec(1 To 2) As K401RptType
  Dim Name19Lgth As String * 19
  Dim LimitFlag As Boolean
  Dim ECnt As Integer
  Dim NumOfErns As Integer
  Dim ErnHandle As Integer
  Dim ErnRec As ErnCodeRecType
  Dim LowDateString$, EndDateString$
  Dim ThisCnt As Integer
  Dim RCalcAmt#
  Dim RCnt As Integer
  Dim RothAmt#
  
  OpenUnitFile UHandle 'added 8/28/03
  Get UHandle, 1, UnitRec
  Close UHandle
  LimitFlag = False
  If QPTrim$(UnitRec.LMT401YN) = "Y" Then
    LimitFlag = True
  End If
  LowDateString = fptxtStart.Text
  EndDateString = fptxtEnd.Text
  EmpRecSize = Len(Emp2Rec)
  TRecSize = Len(TransHRec(1))
  ReDim DedCodes(1 To 50) As DedCodeRecType
  FirstTime = True
  dlm = "~"
'-------------------Entry Error Checking-------------------------
  
  If Date2Num(fptxtStart.Text) > Date2Num(fptxtEnd.Text) Then
     MsgBox "Error: The Start date is later than the End date."
     fptxtStart.SetFocus
     Exit Sub
  End If
  
  If QPTrim$(fptxtStart.Text) = "" Then
     MsgBox "Please select a Start date."
     fptxtStart.SetFocus
     Exit Sub
  End If

  If QPTrim$(fptxtEnd.Text) = "" Then
     MsgBox "Please select an End date."
     fptxtEnd.SetFocus
     Exit Sub
  End If
  
  If Len(fpcomboVolDed.Text) <= 0 Then
     MsgBox "Please select a Voluntary Deduction from the drop down box."
     fpcomboVolDed.SetFocus
     Exit Sub
  End If
  
  If Len(fpcomboLPDed.Text) <= 0 Then
     MsgBox "Please select a Loan Payment Deduction from the drop down box."
     fpcomboLPDed.SetFocus
     Exit Sub
  End If
 
  If fptxtCGMRate.Value < 0 Then
     MsgBox "Please enter a valid Code G Matching Rate."
     fptxtCGMRate.SetFocus
     Exit Sub
  End If

  If fptxtCLMRate.Value < 0 Then
     MsgBox "Please enter a valid Code L Matching Rate."
     fptxtCLMRate.SetFocus
     Exit Sub
  End If

  If Len(fpcmbRoth.Text) <= 0 Then
     MsgBox "Please select a Roth Deduction from the drop down box."
     fpcmbRoth.SetFocus
     Exit Sub
  End If
 
'-----------Entry Error Checking ^-------------------------
  If fpcomboDiskFile.Text = "Y" Then
    Call ElectronicFile 'VertMenu401
  End If
  RptName$ = "PRRPTS\401KG.RPT"

  Image1$ = "#,##0.00"
  Image2$ = "###,##0.00"
  Image3$ = "###"
  
  ReDim K401Totals(1) As K401RptType
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, Unit(1)
  Close UnitHandle

  Text1 = "Supplemental Retirement Income Plan"
  If Unit(1).UFSTATE = "NC" Then
    Text2 = Unit(1).UFSTATE + " 401K Center Plan # 002003"
  Else
    Text2 = Unit(1).UFSTATE + " 401K Center Plan"
  End If
  Text3 = "401K Sub Plan Name: " + QPTrim$(Unit(1).UFEMPR)
  Text4 = "401K Sub Plan Number: " + QPTrim$(Unit(1).BBTCNTNO)
  
  Text1Len = Len(Text1)
  Text2Len = Len(Text2)
  Text3Len = Len(Text3)
  Text4Len = Len(Text4)
  
  LowDate = Date2Num(fptxtStart.Text)
  HighDate = Date2Num(fptxtEnd.Text)
  
  GPct# = fptxtCGMRate.Value
  LPct# = fptxtCLMRate.Value
  
  OpenDedCodeFile DedCodeHandle
  DedRecCnt = LOF(DedCodeHandle) / Len(DedCodes(1))

  ReDim DedCodes(1 To DedRecCnt) As DedCodeRecType
  
  For x = 1 To DedRecCnt
    Get DedCodeHandle, x, DedCodes(x)  'changed alot
  Next x
  Close DedCodeHandle
  
  VCodeNum = Mid(fpcomboVolDed.Text, 1, 2)
  LCodeNum = Mid(fpcomboLPDed.Text, 1, 2)
  RCodeNum = Mid(fpcmbRoth.Text, 1, 2)
  
  VCalcAmt# = 0
  LCalcAmt# = 0
  GCalcAmt# = 0
  RCalcAmt# = 0
  
  GoSub DisplayRptTitle
  
  OpenEmpData2File EmpFile
  
  NumOfRec = LOF(EmpFile) / EmpRecSize
  If NumOfRec = 0 Then
    MsgBox "No employee transaction records on file"
    Close
    Exit Sub
  End If
  
  FrmShowPctComp.Out = False
  
  FrmShowPctComp.Label1 = "Supplemental Retirement Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
  
  AllCnt = NumOfRec
  
  ReDim EmpLBuff(1 To NumOfRec) As EmpSortType
  ReDim EmpGBuff(1 To NumOfRec) As EmpSortType
  
  For cnt = 1 To NumOfRec
    Get #EmpFile, cnt, Emp2Rec
    EmpRet$ = UCase$(Left$(LTrim$(Emp2Rec.EMPRETTP), 1))
    If EmpRet = "" Then EmpRet = "G" 'added 10/8/03
    Select Case EmpRet$
    Case "L"
      LCnt = LCnt + 1
      ReDim Preserve EmpLBuff(1 To LCnt) As EmpSortType
      EmpLBuff(LCnt).EmpNo = Emp2Rec.EmpNo
      EmpLBuff(LCnt).RecNo = cnt
      AllCnt = AllCnt + 2
    Case "G"
      GCnt = GCnt + 1
      ReDim Preserve EmpGBuff(1 To GCnt) As EmpSortType
      EmpGBuff(GCnt).EmpNo = Emp2Rec.EmpNo
      EmpGBuff(GCnt).RecNo = cnt
      AllCnt = AllCnt + 2
    End Select
    FrmShowPctComp.ShowPctComp cnt, AllCnt 'added 5/2, designed
    'to synchronize the progress bar with the number of
    'files processed
    RunCnt = RunCnt + 1
    If RunCnt > AllCnt Then AllCnt = AllCnt + (RunCnt - AllCnt)
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next
  
  offset = 0
  five = 0
  bigName = ""
  For x = 1 To LCnt
  Get #EmpFile, EmpLBuff(x).RecNo, Emp2Rec
     If QPTrim$(Emp2Rec.EmpLName + Emp2Rec.EmpFName) > bigName Then
        bigName = QPTrim$(Emp2Rec.EmpLName + Emp2Rec.EmpFName)
     End If
  Next x
  RunCnt = RunCnt + 1
  If RunCnt > AllCnt Then AllCnt = AllCnt + (RunCnt - AllCnt)
  FrmShowPctComp.ShowPctComp RunCnt, AllCnt '5/2
  If FrmShowPctComp.Out = True Then
    Close
    FrmShowPctComp.Out = False
    Me.cmdEscape.Enabled = True
    Me.cmdProcess.Enabled = True
    EnableCloseButton Me.hwnd, True
    Unload FrmShowPctComp
    Exit Sub
  End If
  'now sort this series
  smallName = QPTrim$(bigName + "z") '"z" is used to make the
  'bigName large enough to include the largest name
  'in the list in the sort
  smallIdx = 1
  y = 1
  Do
    If LCnt = 0 Then Exit Do
    For x = y To LCnt
       Get #EmpFile, EmpLBuff(x).RecNo, Emp2Rec
       If QPTrim$(Emp2Rec.EmpLName + Emp2Rec.EmpFName) < smallName Then
         smallName = QPTrim$(Emp2Rec.EmpLName + Emp2Rec.EmpFName)
         ThisEmp = Emp2Rec.EmpNo
         smallIdx = EmpLBuff(x).RecNo
         Justx = x
       End If
    Next x
    RunCnt = RunCnt + 1
    If RunCnt > AllCnt Then AllCnt = AllCnt + (RunCnt - AllCnt)
    FrmShowPctComp.ShowPctComp RunCnt, AllCnt
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      Exit Sub
    End If
    If y = LCnt Then Exit Do
    'swap to index properly
    TempNo = EmpLBuff(y).EmpNo
    TempIdx = EmpLBuff(y).RecNo
    EmpLBuff(y).EmpNo = ThisEmp
    EmpLBuff(y).RecNo = smallIdx
    EmpLBuff(Justx).EmpNo = TempNo
    EmpLBuff(Justx).RecNo = TempIdx
    y = y + 1
    smallName = QPTrim$(bigName + "z")
  Loop
  
  offset = 0
  five = 0
  bigName = ""
  For x = 1 To GCnt
    Get #EmpFile, EmpGBuff(x).RecNo, Emp2Rec
      If QPTrim$(Emp2Rec.EmpLName + Emp2Rec.EmpFName) > bigName Then
        bigName = QPTrim$(Emp2Rec.EmpLName + Emp2Rec.EmpFName)
      End If
  Next x
    RunCnt = RunCnt + 1
    If RunCnt > AllCnt Then AllCnt = AllCnt + (RunCnt - AllCnt)
    FrmShowPctComp.ShowPctComp RunCnt, AllCnt '5/2
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      Exit Sub
    End If
  'now sort this series
  smallName = QPTrim$(bigName + "z")
  smallIdx = 1
  y = 1
  Do
    If GCnt = 0 Then Exit Do
    For x = y To GCnt
      Get #EmpFile, EmpGBuff(x).RecNo, Emp2Rec
      If QPTrim$(Emp2Rec.EmpLName + Emp2Rec.EmpFName) < smallName Then
        smallName = QPTrim$(Emp2Rec.EmpLName + Emp2Rec.EmpFName)
        ThisEmp = Emp2Rec.EmpNo
        smallIdx = EmpGBuff(x).RecNo
        Justx = x
      End If
    Next x
    RunCnt = RunCnt + 1
    If RunCnt > AllCnt Then AllCnt = AllCnt + (RunCnt - AllCnt)
    FrmShowPctComp.ShowPctComp RunCnt, AllCnt '5/2
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      Exit Sub
    End If
    If y = GCnt Then Exit Do
    'swap to index properly
    TempNo = EmpGBuff(y).EmpNo
    TempIdx = EmpGBuff(y).RecNo
    EmpGBuff(y).EmpNo = ThisEmp
    EmpGBuff(y).RecNo = smallIdx
    EmpGBuff(Justx).EmpNo = TempNo
    EmpGBuff(Justx).RecNo = TempIdx
    y = y + 1
    smallName = QPTrim$(bigName + "z")
  Loop
  Close EmpFile
  
  RHandle = FreeFile
'  On Error GoTo ErrorHandler
  Open RptName$ For Output As RHandle
  THandle = FreeFile
  Open PRData + TransHistFileName For Random As THandle Len = TRecSize
  
  OpenEmpData2File DHandle
  
  FirstTime = True
  EmpRType$ = "G"
  
  For RecNo = 1 To GCnt
    UsingThisOne = False
    VCalcAmt# = 0
    LCalcAmt# = 0
    GCalcAmt# = 0
    RCalcAmt# = 0
    Get DHandle, CLng(EmpGBuff(RecNo).RecNo), Emp2Rec
'    If QPTrim$(Emp2Rec.EmpNo) = "81110" Then Stop
    'If CLng(EmpGBuff(RecNo).RecNo) = 64 Then Stop
    If Emp2Rec.EMPTDATE <> 0 Then
      If LowDate > Emp2Rec.EMPTDATE Then
        GoTo GSkipEm
      End If
    End If
    If Emp2Rec.LastTransRec <= 0 Then
      GoTo GSkipEm
    End If
    TransRecNum& = CLng(Emp2Rec.LastTransRec)
    Do
      Get THandle, TransRecNum&, TransHRec(1)
      Select Case TransHRec(1).CheckDate
      Case LowDate To HighDate
        If VCodeNum > 0 Then
          If TransHRec(1).DAmt(VCodeNum) <> 0 Then
            VCalcAmt# = OldRound(VCalcAmt# + TransHRec(1).DAmt(VCodeNum))
            UsingThisOne = True
          End If
        End If
        If LCodeNum > 0 Then
          If TransHRec(1).DAmt(LCodeNum) <> 0 Then
            LCalcAmt# = OldRound(LCalcAmt# + TransHRec(1).DAmt(LCodeNum))
            UsingThisOne = True
          End If
        End If
        If RCodeNum > 0 Then
          If TransHRec(1).DAmt(RCodeNum) <> 0 Then
            RCalcAmt# = OldRound(RCalcAmt# + TransHRec(1).DAmt(RCodeNum))
            UsingThisOne = True
          End If
        End If
        GCalcAmt# = OldRound(GCalcAmt# + TransHRec(1).GrossPay) '9/26/03
        'this code traps the program if the .Less401k is set to true which means that
        'during the payroll processing the program spotted an alternate earnings
        'code that had been earmarked for exclusion with employer matching funds
        For ECnt = 1 To 3
          If TransHRec(1).Less401k(ECnt) = True Then 'true means don't match this ae amount
            GCalcAmt# = OldRound(GCalcAmt# - TransHRec(1).EAmt(ECnt))
          End If
        Next ECnt
        UsingThisOne = True
      Case Else
      End Select
      If TransHRec(1).PrevTransRec <= 0 Then
        If UsingThisOne Then
          EPct# = GPct#
          GoSub PrintThisOne
        End If
        Exit Do
      Else
        TransRecNum& = CLng(TransHRec(1).PrevTransRec)
      End If
    Loop
GSkipEm:
     RunCnt = RunCnt + 1
     If RunCnt > AllCnt Then AllCnt = AllCnt + (RunCnt - AllCnt)
     FrmShowPctComp.ShowPctComp RunCnt, AllCnt '5/2
     If FrmShowPctComp.Out = True Then
       Close
       FrmShowPctComp.Out = False
       Me.cmdEscape.Enabled = True
       Me.cmdProcess.Enabled = True
       EnableCloseButton Me.hwnd, True
       Unload FrmShowPctComp
       Exit Sub
     End If
  Next
  
  FirstTime = True
  SubTotalVAmt# = 0
  SubTotalLAmt# = 0
  SubTotalMatchAmt# = 0
  SubTotalRAmt# = 0
  
  EmpRType$ = "L"
  EPrinted = 0
  For RecNo = 1 To LCnt
    UsingThisOne = False
    VCalcAmt# = 0
    LCalcAmt# = 0
    RCalcAmt# = 0
    GCalcAmt# = 0
    Get DHandle, CLng(EmpLBuff(RecNo).RecNo), Emp2Rec
    If Emp2Rec.EMPTDATE <> 0 Then
      If LowDate > Emp2Rec.EMPTDATE Then
        GoTo LSkipEm
      End If
    End If
    If Emp2Rec.LastTransRec <= 0 Then
      GoTo LSkipEm
    End If
    TransRecNum& = CLng(Emp2Rec.LastTransRec)
    Do
      Get THandle, TransRecNum&, TransHRec(1)
      Select Case TransHRec(1).CheckDate
      Case LowDate To HighDate
        If VCodeNum > 0 Then
          If TransHRec(1).DAmt(VCodeNum) <> 0 Then
            VCalcAmt# = OldRound(VCalcAmt# + TransHRec(1).DAmt(VCodeNum))
            UsingThisOne = True
          End If
        End If
        If LCodeNum > 0 Then
          If TransHRec(1).DAmt(LCodeNum) <> 0 Then
            LCalcAmt# = OldRound(LCalcAmt# + TransHRec(1).DAmt(LCodeNum))
            UsingThisOne = True
          End If
        End If
        If RCodeNum > 0 Then
          If TransHRec(1).DAmt(RCodeNum) <> 0 Then
            RCalcAmt# = OldRound(RCalcAmt# + TransHRec(1).DAmt(RCodeNum))
            UsingThisOne = True
          End If
        End If
        GCalcAmt# = OldRound(GCalcAmt# + TransHRec(1).GrossPay) '9/26/03
        'this code traps the program if the .Less401k is set to true which means that
        'during the payroll processing the program spotted an alternate earnings
        'code that had been earmarked for exclusion with employer matching funds
        For ECnt = 1 To 3
            If TransHRec(1).Less401k(ECnt) = True Then
              GCalcAmt# = OldRound(GCalcAmt# - TransHRec(1).EAmt(ECnt))
            End If
        Next ECnt
        UsingThisOne = True
      Case Else
      End Select
      If TransHRec(1).PrevTransRec <= 0 Then
        If UsingThisOne Then
          EPct# = LPct#
          GoSub PrintThisOne
        End If
        Exit Do
      Else
        TransRecNum& = CLng(TransHRec(1).PrevTransRec)
      End If
    Loop
LSkipEm:
      RunCnt = RunCnt + 1
      If RunCnt > AllCnt Then AllCnt = AllCnt + (RunCnt - AllCnt)
      FrmShowPctComp.ShowPctComp RunCnt, AllCnt '5/2
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        Me.cmdEscape.Enabled = True
        Me.cmdProcess.Enabled = True
        EnableCloseButton Me.hwnd, True
        Unload FrmShowPctComp
        Exit Sub
    End If
  Next

  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  
  GoSub PrintGrandTots
  Close DHandle
  Close THandle
  Close RHandle
  Close
  Unload FrmShowPctComp
  
  If ThisCnt = 0 Then
    MsgBox "There are no records for the parameters entered."
    Exit Sub
  End If
  
  arSuppRet.Show
  frmLoadingRpt.Show
  
  MainLog ("Supplemental Retirement report processed.")
  Exit Sub

PrintThisOne:
  If QPTrim$(Emp2Rec.YN401K) <> "Y" Then 'added 8/28/03
    TMatchAmt# = 0 'filter out non participators only when
    'matching amounts are calculated
'    If VCalcAmt# > 0 Or LCalcAmt# > 0 Or TMatchAmt# > 0 Then
    If RCalcAmt > 0 Or VCalcAmt# > 0 Or LCalcAmt# > 0 Or TMatchAmt# > 0 Then 'added for ROTH
      GoTo No401KMatch
    Else
      GoTo SkipEMBubba
    End If
  Else
    TMatchAmt# = OldRound((GCalcAmt# * EPct#) * 0.01)
  End If
  
'  If EPct# >= 0 And (VCalcAmt# > 0 Or LCalcAmt# > 0 Or TMatchAmt# > 0) Then
  If EPct# >= 0 And (RCalcAmt > 0 Or VCalcAmt# > 0 Or LCalcAmt# > 0 Or TMatchAmt# > 0) Then 'added for Roth
    If LimitFlag = True And EmpRType$ = "G" Then 'added 8/28/03
      If VCalcAmt# = 0 And RCalcAmt# = 0 Then 'added RCalcAmt for Roth
        If LCalcAmt# = 0 Then 'If a "G" employee does not have
        'a deduction for either voluntary or loan then nothing
        'is printed
          GoTo SkipEMBubba
        Else
          TMatchAmt# = 0
        End If
      Else  '"G" employee with LCalAmt > Or RCalcAmt > 0
      'if employee contributes as much as
      'or more than the % entered on the screen then the employer
      'also matches the % amount on the screen (but no more)...however,
      'if the employee does not contribute as much or more than the
      '% on the screen then the employer only matches what the employee
      'contributes...ex. (employee earns $1000...employer willing to contribute
      '5% (on screen %)...employee contributes $100 (10%) then the employer
      'contributes $50 (5%)...however if the employee contributes $20(2% which is
      'less than 5%) then the employer only contributes 2% ($20)
'        If TMatchAmt# > VCalcAmt# Then
'          TMatchAmt# = VCalcAmt#
'        End If
        If TMatchAmt# > OldRound(VCalcAmt# + RCalcAmt#) Then
          TMatchAmt# = OldRound(VCalcAmt# + RCalcAmt#)
        End If
      End If
    End If
No401KMatch:
    TotalGross# = OldRound(TotalGross# + GCalcAmt#)
    TotalMatchAmt# = OldRound(TotalMatchAmt# + TMatchAmt#)
    TotalVAmt# = OldRound(TotalVAmt# + VCalcAmt#)
    TotalLAmt# = OldRound(TotalLAmt# + LCalcAmt#)
    TotalRoth# = OldRound(TotalRoth# + RCalcAmt#)
    EPrinted = EPrinted + 1
    TEPrinted = TEPrinted + 1
    ReDim K401Rec(1 To 2) As K401RptType
    K401Rec(1).EmpName = LTrim$(RTrim$(Emp2Rec.EmpLName)) + ", " + LTrim$(RTrim$(Emp2Rec.EmpFName))
    Name19Lgth = K401Rec(1).EmpName
    K401Rec(1).SSN = Left$(Emp2Rec.EmpSSN, 3) + "-" + Mid$(Emp2Rec.EmpSSN, 4, 2) + "-" + Mid$(Emp2Rec.EmpSSN, 6, 4)
    RSet K401Rec(1).VAmt = LTrim$(RTrim$(Using(Image1$, VCalcAmt#)))

    RSet K401Rec(1).LAmt = LTrim$(RTrim$(Using(Image1$, LCalcAmt#)))

    RSet K401Rec(1).MAmt = LTrim$(RTrim$(Using(Image1$, TMatchAmt#)))
    
    RSet K401Rec(1).RAmt = LTrim$(RTrim$(Using(Image1$, RCalcAmt#)))
    RSet K401Rec(1).Batch = QPTrim$(Unit(1).BBTBATCH)
    TempDate$ = MakeRegDate(HighDate)
    K401Rec(1).HDate = Left$(TempDate$, 6) + Right$(TempDate$, 2)
    
    SubTotalMatchAmt# = OldRound(SubTotalMatchAmt# + TMatchAmt#)
    SubTotalVAmt# = OldRound(SubTotalVAmt# + VCalcAmt#)
    SubTotalLAmt# = OldRound(SubTotalLAmt# + LCalcAmt#)
    SubTotalRAmt# = OldRound(SubTotalRAmt# + RCalcAmt#)
    
    RSet K401Rec(2).VAmt = LTrim$(RTrim$(Using(Image2$, SubTotalVAmt#)))
    RSet K401Rec(2).LAmt = LTrim$(RTrim$(Using(Image2$, SubTotalLAmt#)))
    RSet K401Rec(2).MAmt = LTrim$(RTrim$(Using(Image2$, SubTotalMatchAmt#)))
    RSet K401Rec(2).RAmt = LTrim$(RTrim$(Using(Image2$, SubTotalRAmt#)))

    If EmpRType$ = "L" Then
       TVLawAmt# = K401Rec(2).VAmt
       TLLawAmt# = K401Rec(2).LAmt
       TotLawMatAmt# = K401Rec(2).MAmt
    ElseIf EmpRType$ = "G" Then
       TVGenAmt# = K401Rec(2).VAmt
       TLGenAmt# = K401Rec(2).LAmt
       TotGenMatAmt# = K401Rec(2).MAmt
    End If
'    TRothAmt# = K401Rec(2).RAmt
  Else
    GoTo SkipEMBubba
  End If
    
  PntCnt = PntCnt + 1
  RSet K401Totals(1).VAmt = LTrim$(RTrim$(Using(Image2$, TotalVAmt#)))
  RSet K401Totals(1).LAmt = LTrim$(RTrim$(Using(Image2$, TotalLAmt#)))
  RSet K401Totals(1).MAmt = LTrim$(RTrim$(Using(Image2$, TotalMatchAmt#)))
  RSet K401Totals(1).RAmt = LTrim$(RTrim$(Using(Image2$, TotalRoth#)))

  If EmpRType = "L" Then
    GenLaw = "Law Enforcement"
  ElseIf EmpRType = "G" Then
    GenLaw = "General "
  End If
  ThisCnt = ThisCnt + 1
  '            0              1           2                 3
  Print #1, EmpRType$; dlm; Text1; dlm; Text2; dlm; Unit(1).UFEMPR; dlm;
  '                4                  5                        6                  7                  8                      9                   10
  Print #1, Unit(1).BBTCNTNO; dlm; GenLaw; dlm; Name19Lgth; dlm; K401Rec(1).SSN; dlm; K401Rec(1).VAmt; dlm; K401Rec(1).LAmt; dlm; K401Rec(1).MAmt; dlm;
  '                11                 12                13                          14                  15                16
  Print #1, K401Rec(1).HDate; dlm; EmpRType; dlm; GenLaw & " SubTotals"; dlm; Str$(EPrinted); dlm; SubTotalVAmt#; dlm; SubTotalLAmt#; dlm;
  '                  17                                      18
  Print #1, SubTotalMatchAmt#; dlm; LTrim$(RTrim$(Using(Image2$, OldRound#(SubTotalVAmt# + SubTotalLAmt# + SubTotalMatchAmt#)))); dlm;
  '                     19                                20                                21                               22
  Print #1, Using("####0.00", TVLawAmt#); dlm; Using("####0.00", TVGenAmt#); dlm; Using("####0.00", TLLawAmt#); dlm; Using("####0.00", TLGenAmt#); dlm;
  '                      23                                 24                               25                             26                           27         28                    29
  Print #1, Using("####0.00", TotLawMatAmt#); dlm; Using("####0.00", TotGenMatAmt#); dlm; Str$(EPrinted); dlm; Using("#,###,##0.00", TotalGross#); dlm; "Y"; dlm; LowDateString; dlm; EndDateString; dlm;
  '              30                   31                    32
  Print #1, K401Rec(1).RAmt; dlm; SubTotalRAmt#; dlm; K401Totals(1).RAmt
  
SkipEMBubba:

Return

PrintGrandTots:
  '          0          1           2               3
  Print #1, "L"; dlm; Text1; dlm; Text2; dlm; Unit(1).UFEMPR; dlm;
  '               4                5        6        7         8        9       10
  Print #1, Unit(1).BBTCNTNO; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm;
  '         11       12                13                            14                   15                  16
  Print #1, ""; dlm; ""; dlm; "Law Enforcement SubTotals"; dlm; Str$(EPrinted); dlm; SubTotalVAmt#; dlm; SubTotalLAmt#; dlm;
  '               17                                                                  18
  Print #1, SubTotalMatchAmt#; dlm; LTrim$(RTrim$(Using(Image2$, OldRound#(SubTotalVAmt# + SubTotalLAmt# + SubTotalMatchAmt#)))); dlm;
  '                     19                                20                                21                               22
  Print #1, Using("####0.00", TVLawAmt#); dlm; Using("####0.00", TVGenAmt#); dlm; Using("####0.00", TLLawAmt#); dlm; Using("####0.00", TLGenAmt#); dlm;
  '                      23                                 24                               25                             26                           27              28                29
  Print #1, Using("####0.00", TotLawMatAmt#); dlm; Using("####0.00", TotGenMatAmt#); dlm; Str$(TEPrinted); dlm; Using("#,###,##0.00", TotalGross#); dlm; "N"; dlm; LowDateString; dlm; EndDateString; dlm;
  '         30            31                     32
  Print #1, ""; dlm; SubTotalRAmt#; dlm; K401Totals(1).RAmt
Return
  
PrintOver:

  Close DHandle
  Close THandle
  Close

SupRetFinish:
  If FrmShowPctComp.Out = False Then
    FrmShowPctComp.Out = True
    Unload FrmShowPctComp
  End If
  
  arSuppRet.Show
  
  MainLog ("Supplemental Retirement report processed.")
  Exit Sub


DisplayRptTitle:
  RptTitle$ = "Supplemental Retirement Report."
Return

ErrorHandler:
  Close
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  Unload FrmShowPctComp
  MsgBox "ERROR: If this problem persists please consult Southern Software."

End Sub

Private Sub fpcomboDiskFile_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboDiskFile.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboDiskFile.ListIndex = -1
  End If
  If fpcomboDiskFile.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcomboLPDed.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcomboLPDed_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboLPDed.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboLPDed.ListIndex = -1
  End If
  If fpcomboLPDed.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fpcomboVolDed.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcomboLPDed_LostFocus()
  fpcomboLPDed.Action = ActionClearSearchBuffer

End Sub

Private Sub fpcomboVolDed_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboVolDed.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboVolDed.ListIndex = -1
  End If
  If fpcomboVolDed.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fptxtEnd.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcomboVolDed_LostFocus()
  fpcomboVolDed.Action = ActionClearSearchBuffer
End Sub

Private Sub VertMenu401()
  Dim TRec(1) As TransRecType
  Dim E2Rec(1) As EmpData2Type
  Dim Unit(1) As UnitFileRecType
  Dim UHandle As Integer
  Dim MaxLen As Integer
  Dim Image1$, Image2$
  Dim EMonth$
  Dim Year$
  Dim VCodeNum As Integer
  Dim LCodeNum As Integer
  Dim GPct#, LPct#
  Dim LowDate As Long
  Dim HiDate As Long
  Dim IdxRecLen As Integer
  Dim IdxFileSize&
  Dim NumOfRecs As Integer
  Dim IdxNHandle As Integer
  Dim IdxRec(1) As NumbSortIdxType
  Dim x As Integer
  Dim RptFile As Integer
  Dim UsingThisOne As Boolean
  Dim CrLf$
  Dim IdxLHandle As Integer
  Dim IdxLRec(1) As NameSortIdxType
  Dim hFile As Integer
  Dim EFile As Integer
  Dim D401Len As Integer
  Dim T401Len As Integer
  Dim VCalcAmt#, LCalAmt#, GCalcAmt#
  Dim RecNo As Integer
  Dim TransRecNum&
  Dim EmpRType$
  Dim EPrinted As Integer
  Dim TMatchAmt#
  Dim TotalVAmt#, TotalLAmt#
  Dim TotalMatchAmt#
  Dim VolDed$
  Dim LoanDed$
  Dim ContDed$
  Dim TVolDed$
  Dim TLoanDed$
  Dim TContDed$
  Dim TDetRecs$
  Dim RptName$
  Dim LCalcAmt#
  Dim EPct#
  Dim DedCode(1) As DedCodeRecType
  Dim DedFile As Integer
  Dim NumOfDed As Integer
  Dim cnt As Integer
  Dim StringLen As Integer
  Dim LimitFlag As Boolean
  Dim ECnt As Integer
  Dim ThisFile As String
  Dim NumOfErns As Integer
  Dim ErnHandle As Integer
  Dim ErnRec As ErnCodeRecType
  
  frmSave401K2Dir.Show
  DoEvents
  OpenUnitFile UHandle 'added 8/28/03
  Get UHandle, 1, Unit(1)
  Close UHandle
'401k here
'  On Error GoTo ErrorHandler1
  ThisFile = "\401K\nk0" + QPTrim$(Unit(1).BBTCNTNO) + ".txt"
  If DirExists(StartPath + "\401K") Then
    If Exist(StartPath + ThisFile) Then
      KillFile (StartPath + ThisFile)
    End If
  Else
    MkDir StartPath + "\401K"
  End If
'  On Error GoTo ErrorHandler
  LimitFlag = False
  If QPTrim$(Unit(1).LMT401YN) = "Y" Then
    LimitFlag = True
  End If
  
  GoSub LoadDedCodes

  MaxLen = 15

  Image1$ = "#####0.00"
  Image2$ = "######"

  EMonth = Mid(fptxtEnd.Text, 1, 2)
  Year = Mid(fptxtEnd.Text, 9, 2)

  If fpcomboVolDed.Text = "" Then
    MsgBox "Please select a Voluntary Deduction from the pick list."
    fpcomboVolDed.SetFocus
    Close
    Exit Sub
  Else
    VCodeNum = Mid(fpcomboVolDed.Text, 1, 2)
  End If
  
  If fpcomboLPDed.Text = "" Then
    MsgBox "Please select a Loan Payment Deduction from the pick list."
    fpcomboLPDed.SetFocus
    Close
    Exit Sub
  Else
    LCodeNum = Mid(fpcomboLPDed.Text, 1, 2)
  End If
  If fptxtCGMRate.Value < 0 Then
    MsgBox "Please enter a valid figure in the Code G Matching Rate field."
    fptxtCGMRate.SetFocus
    Close
    Exit Sub
  Else
    GPct# = fptxtCGMRate.Value
  End If
  If fptxtCLMRate.Value < 0 Then
    MsgBox "Please enter a valid figure in the Code L Matching Rate field."
    fptxtCLMRate.SetFocus
    Close
    Exit Sub
  Else
    LPct# = fptxtCLMRate.Value
  End If

  LowDate = Date2Num(fptxtStart.Text)
  HiDate = Date2Num(fptxtEnd.Text)
  
  OpenEmpIdxNNameFile IdxNHandle
  NumOfRecs = LOF(IdxNHandle) \ 2

  ReDim IdxBuff(1 To NumOfRecs) As Integer
  For x = 1 To NumOfRecs
    Get IdxNHandle, x, IdxBuff(x)
  Next x
  Close IdxNHandle
  
  ReDim D401kRec(1) As DetailRecType
  ReDim T401kRec(1) As TrailerRecType
  D401Len = Len(D401kRec(1))
  T401Len = Len(T401kRec(1))
  
  RptName$ = StartPath + ThisFile
  
  'make disk report here

  CrLf$ = Chr$(13) + Chr$(10)

  ReDim TransHRec(1) As TransRecType
  
  OpenEmpIdxLNameFile IdxLHandle
  NumOfRecs = LOF(IdxLHandle) / 2
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  For x = 1 To NumOfRecs
    Get IdxLHandle, x, IdxBuff(x)
  Next x
  'got input here

  RptFile = FreeFile
  'this errorhandler line traps if there is no
  'diskette is in the A: drive
  On Local Error GoTo ErrorHandler
  Open RptName$ For Output As #RptFile
  
  Close RptFile

  RptFile = FreeFile
  Open RptName$ For Random As RptFile Len = D401Len
  OpenTransHistFile hFile
  OpenEmpData2File EFile
  For RecNo = 1 To NumOfRecs
    UsingThisOne = False
    VCalcAmt# = 0
    LCalcAmt# = 0
    GCalcAmt# = 0
    Get EFile, IdxBuff(RecNo), Emp2Rec(1)
    If Emp2Rec(1).EMPTDATE <> 0 Then
      If LowDate > Emp2Rec(1).EMPTDATE Then
        GoTo SkipEm
      End If
    End If
    If Emp2Rec(1).LastTransRec <= 0 Then
      GoTo SkipEm
    End If

    TransRecNum& = Emp2Rec(1).LastTransRec
    Do
      Get hFile, TransRecNum&, TransHRec(1)
      Select Case TransHRec(1).CheckDate

      Case LowDate To HiDate

        If VCodeNum > 0 Then
          If TransHRec(1).DAmt(VCodeNum) <> 0 Then
            VCalcAmt# = OldRound#(VCalcAmt# + TransHRec(1).DAmt(VCodeNum))
            UsingThisOne = True
          End If
        End If
        If LCodeNum > 0 Then
          If TransHRec(1).DAmt(LCodeNum) <> 0 Then
            LCalcAmt# = OldRound#(LCalcAmt# + TransHRec(1).DAmt(LCodeNum))
            UsingThisOne = True
          End If
        End If
        EmpRType$ = UCase$(Left$(LTrim$(Emp2Rec(1).EMPRETTP), 1))
        If EmpRType$ = "" Then EmpRType$ = "G" 'added 10/8/03
        If EmpRType$ = "L" Or EmpRType$ = "G" Then
          GCalcAmt# = OldRound#(GCalcAmt# + TransHRec(1).GrossPay)
          UsingThisOne = True
            For ECnt = 1 To 3 'True means subtract out this earnings code from matching
              If TransHRec(1).Less401k(ECnt) = True Then
                GCalcAmt# = OldRound(GCalcAmt# - TransHRec(1).EAmt(ECnt))
              End If
            Next ECnt
        End If
      Case Else
      End Select

      If TransHRec(1).PrevTransRec <= 0 Then
        If UsingThisOne Then
          If EmpRType$ = "L" Then
            EPct# = LPct#
          Else
            EPct# = GPct#
          End If
          GoSub PrintThisOne
        End If
        Exit Do
      Else
        TransRecNum& = CLng(TransHRec(1).PrevTransRec)
      End If

    Loop

SkipEm:
  Next

  GoSub DoTrailerRec

EndTheProg:
  Close
  Unload frmSave401K2Dir
  MainLog ("Supplemental Retirement data saved to Citipak Directory.")
  Exit Sub
  
  
PrintThisOne:
  If QPTrim$(Emp2Rec(1).YN401K) <> "Y" Then 'added 8/28/03
    TMatchAmt# = 0 'filter out non participators only when
    'matching amounts are calculated
    If VCalcAmt# + LCalcAmt# = 0 Then
      GoTo SkipEMBubba
    Else
      GoTo No401KMatchDisk
    End If
  Else
    TMatchAmt# = OldRound((GCalcAmt# * EPct#) * 0.01)
  End If

  If EPct# > 0 Or VCalcAmt# > 0 Or LCalcAmt# > 0 Then
    ReDim D401kRec(1) As DetailRecType
    If LimitFlag = True And EmpRType$ = "G" Then 'added 8/28/03
      If VCalcAmt# = 0 Then
        If LCalcAmt# = 0 Then 'If a "G" employee does not have
        'a deduction for either voluntary or loan then nothing
        'is printed to disk
          GoTo SkipEMBubba
        Else
          TMatchAmt# = 0
        End If
      Else '"G" employee with VCalcAmt > 0
      'if employee contributes as much as
      'or more than the % entered on the screen then the employer
      'also matches the % amount on the screen (but no more)...however,
      'if the employee does not contribute as much or more than the
      '% on the screen then the employer only matches what the employee
      'contributes...ex. (employee earns $1000...employer willing to contribute
      '5% (on screen %)...employee contributes $100 (10%) then the employer
      'contributes $50 (5%)...however if the employee contributes $20(2% which is
      'less than 5%) then the employer only contributes 2% ($20)
        If TMatchAmt# > VCalcAmt# Then
          TMatchAmt# = VCalcAmt#
        End If
      End If
    End If
    
No401KMatchDisk:
    EPrinted = EPrinted + 1
    TotalVAmt# = OldRound#(TotalVAmt# + VCalcAmt#)
    TotalLAmt# = OldRound#(TotalLAmt# + LCalcAmt#)
    'TotalMatchAmt only collects TMatchAmt values if
    '   1) EPct or VCalcAmt or LCalcAmt > 0 (EPct is either
    '      the top % on screen if employee is G or bottom if L
    '   2) Employee is G and VCalcAmt > 0
    '   3) Employee is G and VCalcAmt = 0 but LCalcAmt > 0 (TMatch is 0 in this case)
    '
    TotalMatchAmt# = OldRound#(TotalMatchAmt# + TMatchAmt#)

    LSet D401kRec(1).ID = "D"
    LSet D401kRec(1).Batch = "01001"
    LSet D401kRec(1).PCN = QPTrim$(Unit(1).BBTCNTNO)
'    LSet D401kRec(1).ProcDate = EMonth$ + "31" + Year 'Mid(Year, 3, 2)
    LSet D401kRec(1).ProcDate = EMonth$ + LastDayOfMonth(EMonth, Year) + Year 'Mid(Year, 3, 2) 'added 7/7/2004
    LSet D401kRec(1).SSN = Emp2Rec(1).EmpSSN
    LSet D401kRec(1).EmpName = QPTrim$(Emp2Rec(1).EmpFName) + " " + QPTrim$(Emp2Rec(1).EmpLName)
    VolDed$ = CStr(VCalcAmt# * 100)
    StringLen = Len(VolDed$)
    VolDed$ = String(7 - StringLen, "0") + VolDed$
    LSet D401kRec(1).EmpVolDed = VolDed$
    LoanDed$ = CStr(LCalcAmt# * 100)
    StringLen = Len(LoanDed$)
    LoanDed$ = String(7 - StringLen, "0") + LoanDed$
    LSet D401kRec(1).EmpLoanPay = LoanDed$      ''AS STRING * 8
    TMatchAmt# = TMatchAmt# * 100
    StringLen = Len(CStr(TMatchAmt#))
    ContDed$ = String(7 - StringLen, "0") + CStr(TMatchAmt#)
    LSet D401kRec(1).EmpContAmt = ContDed$      ''AS STRING * 8
    D401kRec(1).CrLf = Chr(13) + Chr(10)

    Put #RptFile, , D401kRec(1)
  End If

SkipEMBubba:
  Return

DoTrailerRec:

  LSet T401kRec(1).ID = "T"
  TVolDed$ = CStr(TotalVAmt# * 100)
  StringLen = Len(TVolDed$)
  TVolDed$ = String(10 - StringLen, "0") + TVolDed$
  LSet T401kRec(1).TotVolDED = TVolDed$

  TLoanDed$ = CStr(TotalLAmt# * 100)
  StringLen = Len(TLoanDed$)
  TLoanDed$ = String(10 - StringLen, "0") + TLoanDed$
  LSet T401kRec(1).TotLoanAmt = TLoanDed$

  TContDed$ = CStr(TotalMatchAmt# * 100)
  StringLen = Len(TContDed$)
  TContDed$ = String(10 - StringLen, "0") + TContDed$
  LSet T401kRec(1).TotContAmt = TContDed$       ''AS STRING * 11
  LSet T401kRec(1).Filler = ""

  TDetRecs$ = Using$("###", EPrinted)
  TDetRecs$ = "000000" + QPTrim$(TDetRecs$)
  
  T401kRec(1).TotDRecs = Right$(TDetRecs$, 6)
  LSet T401kRec(1).CrLf = Chr(13) + Chr(10) '"" 'QPTrim$(CrLf$)
  Put #RptFile, , T401kRec(1)

  Return

LoadDedCodes:
  OpenDedCodeFile DedFile
  NumOfDed = LOF(DedFile) / Len(DedCode(1))
  ReDim Item$(1 To NumOfDed)
  For cnt = 1 To NumOfDed
    Get DedFile, cnt, DedCode(1)
    Item$(cnt) = Str$(cnt) + ") " + DedCode(1).DCDESC1
  Next
  Close
Return

ErrorHandler1:
  Unload frmSave401K2Dir
  Close
  MsgBox "ERROR: A problem has occurred in writing the file nk0" + QPTrim$(Unit(1).BBTCNTNO) + ".txt to the 401K directory. Delete or rename the 401K directory and try again. If this doesn't solve the problem call Southern Software Support @ 1-800-842-8190."
'  MsgBox "ERROR: The 401K directory only exists to contain the file nk0" + QPTrim$(Unit(1).BBTCNTNO) + ".txt. If the 401K directory exists but does not contain this file or is empty then please delete the entire directory. Payroll will create the correct file while processing the Supplemental Retirement report. File not saved."
  Exit Sub
 
ErrorHandler:
  Unload frmSave401K2Dir
  MsgBox "ERROR: If this problem persists please consult Southern Software."
  Close
  Unload frmProcessing
End Sub

'Private Function LeapYearCheck(ThisYear As String) As Integer
'  Dim YearDif As Integer
'  If ThisYear = "2000" Then
'    LeapYearCheck = Date2Num("02-29-" + ThisYear)
'    Exit Function
'  End If
'  YearDif = CInt(ThisYear) - 2000
'  If YearDif Mod 4 = 0 Then
'    LeapYearCheck = Date2Num("02-29-" + ThisYear)
'  Else
'    LeapYearCheck = Date2Num("02-28-" + ThisYear)
'  End If
'
'End Function

'Private Sub LowHighDates(ThisMonth As String, LowDate As Long, HighDate As Long)
'    Select Case ThisMonth
'      Case "January":
'        LowDate = Date2Num("01-01-" + fpDateTimeYear.Text)
'        HighDate = Date2Num("01-31-" + fpDateTimeYear.Text)
'      Case "February":
'        LowDate = Date2Num("02-01-" + fpDateTimeYear.Text)
'        HighDate = LeapYearCheck(fpDateTimeYear.Text)
'      Case "March":
'        LowDate = Date2Num("03-01-" + fpDateTimeYear.Text)
'        HighDate = Date2Num("03-31-" + fpDateTimeYear.Text)
'      Case "April":
'        LowDate = Date2Num("04-01-" + fpDateTimeYear.Text)
'        HighDate = Date2Num("04-30-" + fpDateTimeYear.Text)
'      Case "May":
'        LowDate = Date2Num("05-01-" + fpDateTimeYear.Text)
'        HighDate = Date2Num("05-31-" + fpDateTimeYear.Text)
'      Case "June":
'        LowDate = Date2Num("06-01-" + fpDateTimeYear.Text)
'        HighDate = Date2Num("06-30-" + fpDateTimeYear.Text)
'      Case "July":
'        LowDate = Date2Num("07-01-" + fpDateTimeYear.Text)
'        HighDate = Date2Num("07-31-" + fpDateTimeYear.Text)
'      Case "August":
'        LowDate = Date2Num("08-01-" + fpDateTimeYear.Text)
'        HighDate = Date2Num("08-31-" + fpDateTimeYear.Text)
'      Case "September":
'        LowDate = Date2Num("09-01-" + fpDateTimeYear.Text)
'        HighDate = Date2Num("09-30-" + fpDateTimeYear.Text)
'      Case "October":
'        LowDate = Date2Num("10-01-" + fpDateTimeYear.Text)
'        HighDate = Date2Num("10-31-" + fpDateTimeYear.Text)
'      Case "November":
'        LowDate = Date2Num("11-01-" + fpDateTimeYear.Text)
'        HighDate = Date2Num("11-30-" + fpDateTimeYear.Text)
'      Case "December":
'        LowDate = Date2Num("12-01-" + fpDateTimeYear.Text)
'        HighDate = Date2Num("12-31-" + fpDateTimeYear.Text)
'      Case Else:
'        MsgBox "Error: Invalid date entered."
'        Exit Sub
'      End Select
'
'End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdEscape.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmSupRetReport.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub PrintText()
  Dim EmpRecSize As Long, TRecSize As Long, PrnDef$
  Dim RptName$, DedCodeHandle As Integer, x As Integer
  Dim DebugFlag As Boolean, Command$, FF$, NumOfRecs&
  Dim Image1$, Image2$, Image3$, DepSize As Integer
  Dim UnitHandle As Integer, IdxRecLen As Integer
  Dim Text1Len%, Text2Len%, RecNo As Long, IdxFileSize&
  Dim Text3Len%, Text4Len%, Choice$(), EPRTotal#
  Dim cnt As Long, GCntNum As Long, UsingThisOne As Boolean
  Dim GCntName As Long, LCntName As Long
  Dim EPct#
  Dim LowDate&, HighDate&, GPct#, LPct#, TempDed$
  Dim VCodeNum As Integer, LCodeNum As Integer
  Dim VCalcAmt#, LCalcAmt#, GCalcAmt#
  Dim EmpFile As Integer, NumOfRec&, EmpRet$, RHandle As Integer, THandle As Integer
  Dim DHandle As Integer, EmpRType$, LCntNum&, TransRecNum&
  Dim TMatchAmt#, TotalMatchAmt#, TotalVAmt#, TotalLAmt#
  Dim EPrinted As Integer, TempDate$, SubTotalMatchAmt#, SubTotalVAmt#
  Dim SubTotalLAmt#, CurrEmpNo$, PntCnt As Integer
  Dim Totals$, Ret$, EmpGBuffNum() As EmpSortType, GVTotal#, GLTotal#
  Dim TVLawAmt#, TLLawAmt#, LawMatAmt#, TotLawMatAmt#
  Dim TVGenAmt#, TLGenAmt#, GenMatAmt#, TotGenMatAmt#
  Dim offset As Integer, TotalGross#, RptTitle$
  Dim ThisSort() As Integer
  Dim TempEmpNo As Long, TempRecNo As Long, y As Long
  Dim Largest As Long, SortedRecNo As Long, swapThis As Long
  Dim Smallest As Long, DedRecCnt As Long, AlphaFlag As Boolean
  Dim LCnt As Long, GCnt As Long, EmpIdxLNameHandle As Integer
  Dim Unit(1) As UnitFileRecType, Text1 As String
  Dim Text2 As String, Text3 As String, Text4 As String
  Dim LNum As Long, GNum As Long, GLArray() As K401RptType
  Dim z As Long, five As String, bigName As String, smallName As String
  Dim q As Long, smallIdx As Long, LFlag As Boolean, GFlag As Boolean
  Dim tempEmpNum$, TempEmpName$, tempSSN$, TempVAmt$
  Dim TempLAmt$, TempMAmt$, tempBatch$, tempHDate$
  Dim TempGross$, tempRetType$, thisNum As String, Temp401K As K401RptType
  Dim ThisName As String, ThisEmp As String, TempNo As String
  Dim TempIdx As Long, Justx As Long, oldFive As String, Lfive$
  Dim EmpInDept As Integer
  Dim EmpData2FileHandle As Integer
  Dim Emp2Rec As EmpData2Type
  Dim Month$, FirstTime As Boolean
  Dim AllCnt As Integer, RunCnt As Integer
  Dim newFive As Long, oldOffset As Long
  Dim UnitRec As UnitFileRecType
  Dim UHandle As Integer
  Dim LimitFlag As Boolean
  Dim ECnt As Integer, TotalRoth#
  Dim NumOfErns As Integer, SubTotalRAmt#
  Dim ErnHandle As Integer
  Dim ErnRec As ErnCodeRecType
  Dim ThisCnt As Integer, SubTotalRoth#
  Dim GPCnt As Integer
  
  ReDim TransHRec(1) As TransRecType
  ReDim K401Rec(1 To 2) As K401RptType
  ReDim GrossPay(1 To 1) As Long
  
  OpenUnitFile UHandle 'added limit flag on 8/28/03
  Get UHandle, 1, UnitRec
  Close UHandle
  LimitFlag = False
  If QPTrim$(UnitRec.LMT401YN) = "Y" Then
    LimitFlag = True
  End If
  
  EmpRecSize = Len(Emp2Rec)
  TRecSize = Len(TransHRec(1))
  ReDim DedCodes(1 To 50) As DedCodeRecType
  FirstTime = True
'-------------------Entry Error Checking-------------------------
  
  If Date2Num(fptxtStart.Text) > Date2Num(fptxtEnd.Text) Then
     MsgBox "Error: The Start date is later than the End date."
     fptxtStart.SetFocus
     Exit Sub
  End If
  
  If QPTrim$(fptxtStart.Text) = "" Then
     MsgBox "Please select a Start date."
     fptxtStart.SetFocus
     Exit Sub
  End If

  If QPTrim$(fptxtEnd.Text) = "" Then
     MsgBox "Please select an End date."
     fptxtEnd.SetFocus
     Exit Sub
  End If
  
  If Len(fpcomboVolDed.Text) <= 0 Then
     MsgBox "Please select a Voluntary Deduction from the drop down box."
     fpcomboVolDed.SetFocus
     Exit Sub
  End If
  
  If Len(fpcomboLPDed.Text) <= 0 Then
     MsgBox "Please select a Loan Payment Deduction from the drop down box."
     fpcomboLPDed.SetFocus
     Exit Sub
  End If
 
  If fptxtCGMRate.Value < 0 Then
     MsgBox "Please enter valid a Code G Matching Rate."
     fptxtCGMRate.SetFocus
     Exit Sub
  End If

  If fptxtCLMRate.Value < 0 Then
     MsgBox "Please enter a valid Code L Matching Rate."
     fptxtCLMRate.SetFocus
     Exit Sub
  End If

  If Len(fpcmbRoth.Text) <= 0 Then
     MsgBox "Please select a Roth Deduction from the drop down box."
     fpcmbRoth.SetFocus
     Exit Sub
  End If
  
  '-----------Entry Error Checking ^-------------------------
  If fpcomboDiskFile.Text = "Y" Then
    Call ElectronicFile ' VertMenu401
  End If
  RptName$ = "PRRPTS\401K.RPT"

  FF$ = Chr$(12)

  Image1$ = "#,##0.00"
  Image2$ = "###,##0.00"
  Image3$ = "###"
  
  ReDim Choice$(0 To 50, 0 To 2)
  DepSize = 5
  ReDim K401Totals(1) As K401RptType
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, Unit(1)
  Close UnitHandle

  Text1 = "Supplemental Retirement Income Plan"
  If Unit(1).UFSTATE = "NC" Then
    Text2 = Unit(1).UFSTATE + " 401K Center Plan # 002003"
  Else
    Text2 = Unit(1).UFSTATE + " 401K Center Plan"
  End If
  Text3 = "401K Sub Plan Name: " + QPTrim$(Unit(1).UFEMPR)
  Text4 = "401K Sub Plan Number: " + QPTrim$(Unit(1).BBTCNTNO)
  
  Text1Len = Len(Text1)
  Text2Len = Len(Text2)
  Text3Len = Len(Text3)
  Text4Len = Len(Text4)
  
  Choice$(0, 1) = "3,4"
  Choice$(0, 2) = "7"
  Choice$(1, 2) = "1) Department"
  Choice$(2, 2) = "2) Employee Name"

  
  LowDate = Date2Num(fptxtStart.Text)
  HighDate = Date2Num(fptxtEnd.Text)
  GPct# = fptxtCGMRate.Value
  LPct# = fptxtCLMRate.Value
  
  OpenDedCodeFile DedCodeHandle
  DedRecCnt = LOF(DedCodeHandle) / Len(DedCodes(1))

  ReDim DedCodes(1 To DedRecCnt) As DedCodeRecType
  
  For x = 1 To DedRecCnt
    Get DedCodeHandle, x, DedCodes(x)  'changed alot
    Choice$(x, 1) = QPTrim$(DedCodes(x).DCDESC1)
  Next x
  Close DedCodeHandle
  Dim RCodeNum As Integer
  Dim RCalcAmt#
  
  VCodeNum = Mid(fpcomboVolDed.Text, 1, 2)
  LCodeNum = Mid(fpcomboLPDed.Text, 1, 2)
  RCodeNum = Mid(fpcmbRoth.Text, 1, 2)
  
  VCalcAmt# = 0
  LCalcAmt# = 0
  GCalcAmt# = 0
  RCalcAmt# = 0
  
  GoSub DisplayRptTitle
  
  OpenEmpData2File EmpFile
  
  NumOfRec = LOF(EmpFile) / EmpRecSize
  If NumOfRec = 0 Then
    MsgBox "No employee transaction records on file"
    Close
    Exit Sub
  End If
  
  FrmShowPctComp.Out = False
  
  FrmShowPctComp.Label1 = "Supplemental Retirement Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
  
  AllCnt = NumOfRec
  
  ReDim EmpLBuff(1 To NumOfRec) As EmpSortType
  ReDim EmpGBuff(1 To NumOfRec) As EmpSortType
  For cnt = 1 To NumOfRec
    Get #EmpFile, cnt, Emp2Rec
    EmpRet$ = UCase$(Left$(LTrim$(Emp2Rec.EMPRETTP), 1))
    If EmpRet = "" Then EmpRet = "G" 'added 10/8/03
    Select Case EmpRet$
    Case "L"
      LCnt = LCnt + 1
      ReDim Preserve EmpLBuff(1 To LCnt) As EmpSortType
      EmpLBuff(LCnt).EmpNo = Emp2Rec.EmpNo
      EmpLBuff(LCnt).RecNo = cnt
      AllCnt = AllCnt + 2
    Case "G"
      GCnt = GCnt + 1
      ReDim Preserve EmpGBuff(1 To GCnt) As EmpSortType
      EmpGBuff(GCnt).EmpNo = Emp2Rec.EmpNo
      EmpGBuff(GCnt).RecNo = cnt
      AllCnt = AllCnt + 2
    End Select
    FrmShowPctComp.ShowPctComp cnt, AllCnt 'added 5/2, designed
    'to synchronize the progress bar with the number of
    'files processed
    RunCnt = RunCnt + 1
    If RunCnt > AllCnt Then AllCnt = AllCnt + (RunCnt - AllCnt)
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      Exit Sub
    End If
    
  Next
  
  offset = 0
  five = 0
  bigName = ""
  For x = 1 To LCnt
  Get #EmpFile, EmpLBuff(x).RecNo, Emp2Rec
     If QPTrim$(Emp2Rec.EmpLName + Emp2Rec.EmpFName) > bigName Then
        bigName = QPTrim$(Emp2Rec.EmpLName + Emp2Rec.EmpFName)
     End If
  
  Next x
  RunCnt = RunCnt + 1
  If RunCnt > AllCnt Then AllCnt = AllCnt + (RunCnt - AllCnt)
  FrmShowPctComp.ShowPctComp RunCnt, AllCnt '5/2
  If FrmShowPctComp.Out = True Then
    Close
    FrmShowPctComp.Out = False
    Me.cmdEscape.Enabled = True
    Me.cmdProcess.Enabled = True
    EnableCloseButton Me.hwnd, True
    Unload FrmShowPctComp
    Exit Sub
  End If
  'now sort this series
  smallName = QPTrim$(bigName + "z") '"z" is used to make the
  'bigName large enough to include the largest name
  'in the list in the sort
  smallIdx = 1
  y = 1
  Do
    If LCnt = 0 Then Exit Do
    For x = y To LCnt
       Get #EmpFile, EmpLBuff(x).RecNo, Emp2Rec
       If QPTrim$(Emp2Rec.EmpLName + Emp2Rec.EmpFName) < smallName Then
         smallName = QPTrim$(Emp2Rec.EmpLName + Emp2Rec.EmpFName)
         ThisEmp = Emp2Rec.EmpNo
         smallIdx = EmpLBuff(x).RecNo
         Justx = x
       End If
    
    Next x
    RunCnt = RunCnt + 1
    If RunCnt > AllCnt Then AllCnt = AllCnt + (RunCnt - AllCnt)
    FrmShowPctComp.ShowPctComp RunCnt, AllCnt
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      Exit Sub
    End If
    If y = LCnt Then Exit Do
    'swap to index properly
    TempNo = EmpLBuff(y).EmpNo
    TempIdx = EmpLBuff(y).RecNo
    EmpLBuff(y).EmpNo = ThisEmp
    EmpLBuff(y).RecNo = smallIdx
    EmpLBuff(Justx).EmpNo = TempNo
    EmpLBuff(Justx).RecNo = TempIdx
    y = y + 1
    smallName = QPTrim$(bigName + "z")
  Loop
  
  offset = 0
  five = 0
  bigName = ""
  For x = 1 To GCnt
    Get #EmpFile, EmpGBuff(x).RecNo, Emp2Rec
      If QPTrim$(Emp2Rec.EmpLName + Emp2Rec.EmpFName) > bigName Then
        bigName = QPTrim$(Emp2Rec.EmpLName + Emp2Rec.EmpFName)
      End If
    
  Next x
    RunCnt = RunCnt + 1
    If RunCnt > AllCnt Then AllCnt = AllCnt + (RunCnt - AllCnt)
    FrmShowPctComp.ShowPctComp RunCnt, AllCnt '5/2
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      Exit Sub
    End If
  'now sort this series
  smallName = QPTrim$(bigName + "z")
  smallIdx = 1
  y = 1
  Do
    If GCnt = 0 Then Exit Do
    For x = y To GCnt
      Get #EmpFile, EmpGBuff(x).RecNo, Emp2Rec
      If QPTrim$(Emp2Rec.EmpLName + Emp2Rec.EmpFName) < smallName Then
        smallName = QPTrim$(Emp2Rec.EmpLName + Emp2Rec.EmpFName)
        ThisEmp = Emp2Rec.EmpNo
        smallIdx = EmpGBuff(x).RecNo
        Justx = x
      End If
    Next x
    RunCnt = RunCnt + 1
    If RunCnt > AllCnt Then AllCnt = AllCnt + (RunCnt - AllCnt)
    FrmShowPctComp.ShowPctComp RunCnt, AllCnt '5/2
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      Exit Sub
    End If
    If y = GCnt Then Exit Do
    'swap to index properly
    TempNo = EmpGBuff(y).EmpNo
    TempIdx = EmpGBuff(y).RecNo
    EmpGBuff(y).EmpNo = ThisEmp
    EmpGBuff(y).RecNo = smallIdx
    EmpGBuff(Justx).EmpNo = TempNo
    EmpGBuff(Justx).RecNo = TempIdx
    y = y + 1
    smallName = QPTrim$(bigName + "z")
  Loop
  Close EmpFile
  
  RHandle = FreeFile
  Open RptName$ For Output As RHandle
  RPTSetupPRN 12, RHandle
  THandle = FreeFile
  Open PRData + TransHistFileName For Random As THandle Len = TRecSize
  
  OpenEmpData2File DHandle
  GoSub K401Header
  
  'set FirstTime to true so header for "general" prints
  'only once
  FirstTime = True
  EmpRType$ = "G"
  
  For RecNo = 1 To GCnt
    UsingThisOne = False
    TVLawAmt# = 0
    TLLawAmt# = 0
    TLGenAmt# = 0
    TVGenAmt# = 0
    VCalcAmt# = 0
    LCalcAmt# = 0
    GCalcAmt# = 0
    RCalcAmt# = 0
    Get DHandle, CLng(EmpGBuff(RecNo).RecNo), Emp2Rec
    'If CLng(EmpGBuff(RecNo).RecNo) = 64 Then Stop
    If Emp2Rec.EMPTDATE <> 0 Then
      If LowDate > Emp2Rec.EMPTDATE Then
        GoTo GSkipEm
      End If
    End If
    If Emp2Rec.LastTransRec <= 0 Then
      GoTo GSkipEm
    End If
    TransRecNum& = CLng(Emp2Rec.LastTransRec)
    Do
      Get THandle, TransRecNum&, TransHRec(1)
      Select Case TransHRec(1).CheckDate
      Case LowDate To HighDate
        If VCodeNum > 0 Then
          If TransHRec(1).DAmt(VCodeNum) <> 0 Then
            VCalcAmt# = OldRound(VCalcAmt# + TransHRec(1).DAmt(VCodeNum))
            UsingThisOne = True
          End If
        End If
        If LCodeNum > 0 Then
          If TransHRec(1).DAmt(LCodeNum) <> 0 Then
            LCalcAmt# = OldRound(LCalcAmt# + TransHRec(1).DAmt(LCodeNum))
            UsingThisOne = True
          End If
        End If
        If RCodeNum > 0 Then
          If TransHRec(1).DAmt(RCodeNum) <> 0 Then
            RCalcAmt# = OldRound(RCalcAmt# + TransHRec(1).DAmt(RCodeNum))
            UsingThisOne = True
          End If
        End If
'        For z = 1 To GPCnt
'          If TransHRec(1).EmpPin = GrossPay(z) Then
'            GoTo GPDone
'          End If
'        Next z
'        GPCnt = GPCnt + 1
'        ReDim Preserve GrossPay(1 To GPCnt) As Long
'        GrossPay(GPCnt) = TransHRec(1).EmpPin
        GCalcAmt# = OldRound(GCalcAmt# + TransHRec(1).GrossPay) '9/26/03
        'this code traps the program if the .Less401k is set to true which means that
        'during the payroll processing the program spotted an alternate earnings
        'code that had been earmarked for exclusion with employer matching funds
        For ECnt = 1 To 3
          If TransHRec(1).Less401k(ECnt) = True Then
            GCalcAmt# = OldRound(GCalcAmt# - TransHRec(1).EAmt(ECnt))
          End If
        Next ECnt
GPDone:
        UsingThisOne = True
      Case Else
      End Select
      If TransHRec(1).PrevTransRec <= 0 Then
        If UsingThisOne Then
          EPct# = GPct#
          GoSub PrintThisOne
        End If
        Exit Do
      Else
        TransRecNum& = CLng(TransHRec(1).PrevTransRec)
      End If
    Loop
GSkipEm:
     RunCnt = RunCnt + 1
     If RunCnt > AllCnt Then AllCnt = AllCnt + (RunCnt - AllCnt)
     FrmShowPctComp.ShowPctComp RunCnt, AllCnt '5/2
     If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        Me.cmdEscape.Enabled = True
        Me.cmdProcess.Enabled = True
        EnableCloseButton Me.hwnd, True
        Unload FrmShowPctComp
        Exit Sub
      End If
  Next
  
  GoSub PrintSubTotals
  'reset FirstTime so the heading for law enforcement
  'prints only once
  FirstTime = True
  
  EmpRType$ = "L"

  For RecNo = 1 To LCnt
    UsingThisOne = False
    VCalcAmt# = 0
    LCalcAmt# = 0
    RCalcAmt# = 0
    GCalcAmt# = 0
    Get DHandle, CLng(EmpLBuff(RecNo).RecNo), Emp2Rec
    If Emp2Rec.EMPTDATE <> 0 Then
      If LowDate > Emp2Rec.EMPTDATE Then
        GoTo LSkipEm
      End If
    End If
    If Emp2Rec.LastTransRec <= 0 Then
      GoTo LSkipEm
    End If
    TransRecNum& = CLng(Emp2Rec.LastTransRec)
    Do
      Get THandle, TransRecNum&, TransHRec(1)
      Select Case TransHRec(1).CheckDate
      Case LowDate To HighDate
        If VCodeNum > 0 Then
          If TransHRec(1).DAmt(VCodeNum) <> 0 Then
            VCalcAmt# = OldRound(VCalcAmt# + TransHRec(1).DAmt(VCodeNum))
           UsingThisOne = True
          End If
        End If
        If LCodeNum > 0 Then
          If TransHRec(1).DAmt(LCodeNum) <> 0 Then 'added < on 11/3/2008
            LCalcAmt# = OldRound(LCalcAmt# + TransHRec(1).DAmt(LCodeNum))
            UsingThisOne = True
          End If
        End If
        If RCodeNum > 0 Then
          If TransHRec(1).DAmt(RCodeNum) <> 0 Then
            RCalcAmt# = OldRound(RCalcAmt# + TransHRec(1).DAmt(RCodeNum))
            UsingThisOne = True
          End If
        End If
'        For z = 1 To GPCnt
'          If TransHRec(1).EmpPin = GrossPay(z) Then
'            GoTo GPDone2
'          End If
'        Next z
        GCalcAmt# = OldRound(GCalcAmt# + TransHRec(1).GrossPay) '9/26/03
        'this code traps the program if the .Less401k is set to true which means that
        'during the payroll processing the program spotted an alternate earnings
        'code that had been earmarked for exclusion with employer matching funds
        For ECnt = 1 To 3
            If TransHRec(1).Less401k(ECnt) = True Then
              GCalcAmt# = OldRound(GCalcAmt# - TransHRec(1).EAmt(ECnt))
            End If
        Next ECnt
GPDone2:
        UsingThisOne = True
      Case Else
      End Select
      If TransHRec(1).PrevTransRec <= 0 Then
        If UsingThisOne Then
          EPct# = LPct#
          GoSub PrintThisOne
        End If
        Exit Do
      Else
        TransRecNum& = CLng(TransHRec(1).PrevTransRec)
      End If
    Loop
LSkipEm:

      RunCnt = RunCnt + 1
      If RunCnt > AllCnt Then AllCnt = AllCnt + (RunCnt - AllCnt)
      FrmShowPctComp.ShowPctComp RunCnt, AllCnt '5/2
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        Me.cmdEscape.Enabled = True
        Me.cmdProcess.Enabled = True
        EnableCloseButton Me.hwnd, True
        Unload FrmShowPctComp
        Exit Sub
      End If
  Next

  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  
  GoSub PrintSubTotals
  LFlag = True
  GFlag = True
  GoSub PrintGrandTots
  Close DHandle
  Close THandle
  RPTSetupPRN 123, RHandle '7/24
  Close RHandle
  Close

  GoTo SupRetFinish

PrintThisOne:
  If QPTrim$(Emp2Rec.YN401K) <> "Y" Then 'added 8/28/03
    TMatchAmt# = 0 'filter out non participators only when
    'matching amounts are calculated
   'If RCalcAmt > 0 Or VCalcAmt# > 0 Or LCalcAmt# > 0 Or TMatchAmt# > 0 Then 'added for ROTH
    If RCalcAmt > 0 Or VCalcAmt# > 0 Or LCalcAmt# > 0 Or TMatchAmt# > 0 Then 'added 9/26/03
      GoTo No401KMatch
    Else
      GoTo SkipEMBubba
    End If
  Else
    TMatchAmt# = OldRound((GCalcAmt# * EPct#) * 0.01)
  End If
  
  If EPct# >= 0 And (RCalcAmt# > 0 Or VCalcAmt# > 0 Or LCalcAmt# > 0 Or TMatchAmt# > 0) Then
    If LimitFlag = True And EmpRType$ = "G" Then 'added 8/28/03
      If VCalcAmt# = 0 And RCalcAmt# = 0 Then
        If LCalcAmt# = 0 Then 'If a "G" employee does not have
        'a deduction for either voluntary or loan then nothing
        'is printed to disk
          GoTo SkipEMBubba
        Else
          TMatchAmt# = 0
        End If
      Else  '"G" employee with LCalAmt > 0
      'if employee contributes as much as
      'or more than the % entered on the screen then the employer
      'also matches the % amount on the screen (but no more)...however,
      'if the employee does not contribute as much or more than the
      '% on the screen then the employer only matches what the employee
      'contributes...ex. (employee earns $1000...employer willing to contribute
      '5% (on screen %)...employee contributes $100 (10%) then the employer
      'contributes $50 (5%)...however if the employee contributes $20(2% which is
      'less than 5%) then the employer only contributes 2% ($20)
        If TMatchAmt# > OldRound(VCalcAmt# + RCalcAmt#) Then
          TMatchAmt# = OldRound(VCalcAmt# + RCalcAmt#)
        End If
      End If
    End If
No401KMatch:
    TotalGross# = OldRound(TotalGross# + GCalcAmt#)
    TotalMatchAmt# = OldRound(TotalMatchAmt# + TMatchAmt#)
    TotalVAmt# = OldRound(TotalVAmt# + VCalcAmt#)
    TotalLAmt# = OldRound(TotalLAmt# + LCalcAmt#)
    TotalRoth# = OldRound(TotalRoth# + RCalcAmt#)
    EPrinted = EPrinted + 1
    ReDim K401Rec(1 To 2) As K401RptType
    K401Rec(1).EmpName = LTrim$(RTrim$(Emp2Rec.EmpLName)) + " " + LTrim$(RTrim$(Emp2Rec.EmpFName))
    K401Rec(1).SSN = Left$(Emp2Rec.EmpSSN, 3) + "-" + Mid$(Emp2Rec.EmpSSN, 4, 2) + "-" + Mid$(Emp2Rec.EmpSSN, 6, 4)
    RSet K401Rec(1).VAmt = LTrim$(RTrim$(Using(Image1$, VCalcAmt#)))
    If VCalcAmt# = 0 Then
      RSet K401Rec(1).VAmt = ".00"
    End If

    RSet K401Rec(1).LAmt = LTrim$(RTrim$(Using(Image1$, LCalcAmt#)))
    If LCalcAmt# = 0 Then
      RSet K401Rec(1).LAmt = ".00"
    End If

    RSet K401Rec(1).MAmt = LTrim$(RTrim$(Using(Image1$, TMatchAmt#)))
    If TMatchAmt# = 0 Then
      RSet K401Rec(1).MAmt = ".00"
    End If
    RSet K401Rec(1).RAmt = LTrim$(RTrim$(Using(Image1$, RCalcAmt#)))
    RSet K401Rec(1).Batch = QPTrim$(Unit(1).BBTBATCH)
    TempDate$ = MakeRegDate(HighDate)
    K401Rec(1).HDate = Left$(TempDate$, 6) + Right$(TempDate$, 2)
      If FirstTime Then
        Select Case EmpRType$
        Case "L"
          Print #1,
          Print #1, "***Law Enforcement***";
          Print #1,
          Print #1,
        Case "G"
          Print #1,
          Print #1, "***General***";
          Print #1,
          Print #1,
        End Select
        FirstTime = False
      End If
      SubTotalMatchAmt# = OldRound(SubTotalMatchAmt# + TMatchAmt#)
      SubTotalVAmt# = OldRound(SubTotalVAmt# + VCalcAmt#)
      SubTotalLAmt# = OldRound(SubTotalLAmt# + LCalcAmt#)
      SubTotalRAmt# = OldRound(SubTotalRAmt# + RCalcAmt#)
    Else
      GoTo SkipEMBubba:
    
    End If
    ThisCnt = ThisCnt + 1
    Print #1, K401Rec(1).EmpName; K401Rec(1).SSN; "   "; K401Rec(1).VAmt; " "; K401Rec(1).LAmt;
    Print #1, "  "; K401Rec(1).MAmt; "    "; K401Rec(1).HDate; " "; EmpRType; K401Rec(1).RAmt
    PntCnt = PntCnt + 1
  
SkipEMBubba:
Return
  
K401Header:
  TempDate$ = Date$
  TempDate$ = Left$(TempDate$, 6) + Right$(TempDate$, 2)
  Print #1, Space$(48 - (Text1Len \ 2)); Text1 '; Tab(81); "Date: "; TempDate$
  Print #1, Space$(48 - (Text2Len \ 2)); Text2 '; Tab(81); "Page: 1 "
  Print #1, Space$(48 - (Text3Len \ 2)); Text3
  Print #1, Space$(48 - (Text4Len \ 2)); Text4
  Print #1, Tab(30); "From: "; Tab(36); fptxtStart.Text; Tab(50); "To: "; Tab(56); fptxtEnd.Text
  Print #1, ""
  Print #1, "                                                            Post-Tax                   Pay Prd               "
  Print #1, "Employee Name                       SS#          Source A   Loan Pmts    Source P        End         Source U"
Return
  
PrintSubTotals:
  Totals$ = Space$(121)
  RSet K401Rec(2).VAmt = LTrim$(RTrim$(Using(Image2$, SubTotalVAmt#)))
  RSet K401Rec(2).LAmt = LTrim$(RTrim$(Using(Image2$, SubTotalLAmt#)))
  RSet K401Rec(2).MAmt = LTrim$(RTrim$(Using(Image2$, SubTotalMatchAmt#)))
  RSet K401Rec(2).RAmt = LTrim$(RTrim$(Using(Image2$, SubTotalRAmt#)))
  If EmpRType$ = "L" Then
     TVLawAmt# = K401Rec(2).VAmt
     TLLawAmt# = K401Rec(2).LAmt
     TotLawMatAmt# = K401Rec(2).MAmt
  ElseIf EmpRType$ = "G" Then
     TVGenAmt# = K401Rec(2).VAmt
     TLGenAmt# = K401Rec(2).LAmt
     TotGenMatAmt# = K401Rec(2).MAmt
  End If
  
  If SubTotalVAmt# = 0 Then
    RSet K401Rec(2).VAmt = ".00"
  End If
  If SubTotalLAmt# = 0 Then
    RSet K401Rec(2).LAmt = ".00"
  End If
  If SubTotalMatchAmt# = 0 Then
    RSet K401Rec(2).MAmt = ".00"
  End If
  If SubTotalRAmt# = 0 Then
    RSet K401Rec(2).RAmt = ".00"
  End If
  
  LSet Totals$ = "SubTotals"
  Mid$(Totals$, 37) = Using(Image3$, PntCnt)
  Mid$(Totals$, 47) = K401Rec(2).VAmt
  Mid$(Totals$, 59) = K401Rec(2).LAmt
  Mid$(Totals$, 72) = K401Rec(2).MAmt
  Mid$(Totals$, 87) = LTrim$(RTrim$(Using(Image2$, OldRound#(SubTotalVAmt# + SubTotalLAmt# + SubTotalMatchAmt#))))
  Mid$(Totals$, 99) = K401Rec(2).RAmt
  Print #1,
  Print #1, Totals$
  SubTotalVAmt# = 0
  SubTotalLAmt# = 0
  SubTotalRAmt# = 0
  SubTotalMatchAmt# = 0
  PntCnt = 0
Return

PrintOver:

  Close DHandle
  Close THandle
  Print #RHandle, PrnDef$;
  Close

SupRetFinish:
  If FrmShowPctComp.Out = False Then
    FrmShowPctComp.Out = True
    Unload FrmShowPctComp
  End If
  
  If ThisCnt = 0 Then
    MsgBox "There are no records for the parameters entered."
    Exit Sub
  End If
  
  ViewPrint RptName$, RptTitle$, True
  
  MainLog ("Supplemental Retirement report processed.")
  Exit Sub

PrintGrandTots:
  Totals$ = Space$(95)
  RSet K401Totals(1).VAmt = LTrim$(RTrim$(Using(Image2$, TotalVAmt#)))
  RSet K401Totals(1).LAmt = LTrim$(RTrim$(Using(Image2$, TotalLAmt#)))
  RSet K401Totals(1).MAmt = LTrim$(RTrim$(Using(Image2$, TotalMatchAmt#)))
  RSet K401Totals(1).RAmt = LTrim$(RTrim$(Using(Image2$, TotalRoth#)))
  
  If TotalVAmt# = 0 Then
    RSet K401Totals(1).VAmt = ".00"
  End If
  If TotalLAmt# = 0 Then
    RSet K401Totals(1).LAmt = ".00"
  End If
  If TotalMatchAmt# = 0 Then
    RSet K401Totals(1).MAmt = ".00"
  End If
  If TotalRoth# = 0 Then
    RSet K401Totals(1).RAmt = ".00"
  End If

  Totals$ = Space$(151)
  LSet Totals$ = "Totals"
  Mid$(Totals$, 22) = Using(Image2$, TotalGross#)
  Mid$(Totals$, 37) = Str$(EPrinted)
  Mid$(Totals$, 47) = K401Totals(1).VAmt
  Mid$(Totals$, 59) = K401Totals(1).LAmt
  Mid$(Totals$, 72) = K401Totals(1).MAmt
  Mid$(Totals$, 87) = LTrim$(RTrim$(Using("#,###,##0.00", OldRound#(TotalVAmt# + TotalLAmt# + TotalMatchAmt#))))
  Mid$(Totals$, 99) = K401Totals(1).RAmt

  Print #RHandle,
  Print #RHandle, Totals$
  If LFlag = False Or GFlag = False Then GoTo PrintOver
  Print #RHandle,
  Print #RHandle, "Subtotals"
  Print #RHandle,
  Print #RHandle, " Total Source A Law "; Using("####0.00", TVLawAmt#)
  Print #RHandle, " Total Source A Gen "; Using("####0.00", TVGenAmt#)
  Print #RHandle,
  Print #RHandle, " Total Loan Law     "; Using("####0.00", TLLawAmt#)
  Print #RHandle, " Total Loan Gen     "; Using("####0.00", TLGenAmt#)
  Print #RHandle,
  Print #RHandle, "Total Source P Law  "; Using("####0.00", TotLawMatAmt#)
  Print #RHandle, "Total Source P Gen  "; Using("####0.00", TotGenMatAmt#)
  Print #RHandle,
  Print #RHandle, "Total Gross Pay"; " "; Using("#,###,##0.00", TotalGross#) ' + "  *excludes ROTH"
  Print #RHandle,
  Print #RHandle, " Total Source U"; " "; Using("#,###,##0.00", TotalRoth#)
  Print #RHandle, FF$

  Return
  
DisplayRptTitle:
  RptTitle$ = "Supplemental Retirement Report."
Return

End Sub
Private Sub fpcomboPrintOpt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboPrintOpt.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboPrintOpt.ListIndex = -1
  End If
  If fpcomboPrintOpt.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      cmdEscape.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub ElectronicFile()
  Dim TRec(1) As TransRecType
  Dim E2Rec(1) As EmpData2Type
  Dim Unit(1) As UnitFileRecType
  Dim UHandle As Integer
  Dim MaxLen As Integer
  Dim Image1$, Image2$
  Dim EMonth$
  Dim Year$
  Dim VCodeNum As Integer
  Dim LCodeNum As Integer
  Dim GPct#, LPct#
  Dim LowDate As Long
  Dim HiDate As Long
  Dim IdxRecLen As Integer
  Dim IdxFileSize&
  Dim NumOfRecs As Integer
  Dim IdxNHandle As Integer
  Dim IdxRec(1) As NumbSortIdxType
  Dim x As Integer, z As Integer
  Dim RptFile As Integer
  Dim UsingThisOne As Boolean
  Dim IdxLHandle As Integer
  Dim IdxLRec(1) As NameSortIdxType
  Dim hFile As Integer
  Dim EFile As Integer
  Dim VCalcAmt#, LCalAmt#, GCalcAmt#, RothAmt#
  Dim RecNo As Integer
  Dim TransRecNum&
  Dim EmpRType$
  Dim EPrinted As Integer
  Dim TMatchAmt#
  Dim TotalVAmt#, TotalLAmt#, TotalRAmt#
  Dim TotalMatchAmt#
  Dim VolDed$
  Dim LoanDed$
  Dim ContDed$
  Dim TVolDed$
  Dim TLoanDed$
  Dim TContDed$
  Dim TDetRecs$
  Dim RptName$
  Dim LCalcAmt#
  Dim EPct#
  Dim DedCode(1) As DedCodeRecType
  Dim DedFile As Integer
  Dim NumOfDed As Integer
  Dim cnt As Integer
  Dim StringLen As Integer
  Dim LimitFlag As Boolean
  Dim ECnt As Integer
  Dim ThisFile As String
  Dim NumOfErns As Integer
  Dim ErnHandle As Integer
  Dim ErnRec As ErnCodeRecType
  Dim NumOfDetRecs As Integer
  Dim NumOfLoanRecs As Integer
  Dim SupHeader As SupHeaderRecType
  Dim SupTrailer As SupTrailerRecType
  Dim RCodeNum As Integer
  Dim CreateYear$, CreateMonth$, CreateDay$
  Dim PrintLine$
  Dim ThisAmtS$
  Dim ThisAmtD#
  Dim PrintCnt As Integer
  Dim NegChar1$
  Dim NegChar2$
  Dim TruncName As String * 30
  Dim ThisLen As Integer
  Dim WorkName$
  Dim BuildName$
  Dim Thisch As String * 1
  Dim FinalName$
  Dim PrudTmp$
  Dim PrudAddInfo$
  
  OpenUnitFile UHandle 'added 8/28/03
  Get UHandle, 1, Unit(1)
  Close UHandle
  
  If Len(QPTrim$(CStr(Unit(1).BBTCNTNO))) <> 6 Then
    MsgBox ("The Plan Number should be a 6 digit number. Then number stored is " + QPTrim$(Unit(1).BBTCNTNO) + " and will not be accepted by the receiver of this file.")
    Close
    Exit Sub
  End If
  
  frmSave401K2Dir.Show
  DoEvents
  
'  On Error GoTo ErrorHandler1
  
  ThisFile = "\401K\NK" + QPTrim$(Unit(1).BBTCNTNO) + ".txt"
  If DirExists(StartPath + "\401K") Then
    If Exist(StartPath + ThisFile) Then
      KillFile (StartPath + ThisFile)
    End If
  Else
    MkDir StartPath + "\401K"
  End If
  
'  On Error GoTo ErrorHandler

  LimitFlag = False
  If QPTrim$(Unit(1).LMT401YN) = "Y" Then
    LimitFlag = True
  End If
  
  GoSub LoadDedCodes
  
  MaxLen = 15
  If fpcomboVolDed.Text = "" Then
    MsgBox "Please select a Voluntary Deduction from the pick list."
    fpcomboVolDed.SetFocus
    Close
    Exit Sub
  Else
    VCodeNum = Mid(fpcomboVolDed.Text, 1, 2)
  End If
  
  If fpcomboLPDed.Text = "" Then
    MsgBox "Please select a Loan Payment Deduction from the pick list."
    fpcomboLPDed.SetFocus
    Close
    Exit Sub
  Else
    LCodeNum = Mid(fpcomboLPDed.Text, 1, 2)
  End If
  
  If fptxtCGMRate.Value < 0 Then
    MsgBox "Please enter a valid figure in the Code G Matching Rate field."
    fptxtCGMRate.SetFocus
    Close
    Exit Sub
  Else
    GPct# = fptxtCGMRate.Value
  End If
  
  If fptxtCLMRate.Value < 0 Then
    MsgBox "Please enter a valid figure in the Code L Matching Rate field."
    fptxtCLMRate.SetFocus
    Close
    Exit Sub
  Else
    LPct# = fptxtCLMRate.Value
  End If

  If fpcmbRoth.Text = "" Then
    MsgBox "Please select a Roth Payment Deduction from the pick list."
    fpcmbRoth.SetFocus
    Close
    Exit Sub
  Else
    RCodeNum = Mid(fpcmbRoth.Text, 1, 2)
  End If
  
  LowDate = Date2Num(fptxtStart.Text)
  HiDate = Date2Num(fptxtEnd.Text)
  
  OpenEmpIdxNNameFile IdxNHandle
  NumOfRecs = LOF(IdxNHandle) \ 2

  ReDim IdxBuff(1 To NumOfRecs) As Integer
  For x = 1 To NumOfRecs
    Get IdxNHandle, x, IdxBuff(x)
  Next x
  Close IdxNHandle
  
  ReDim SupDet(1 To 2) As SupDetailRecType
  ReDim SupLoans(1 To 2) As SupLoanRecType
  
  RptName$ = StartPath + ThisFile
  
  ReDim TransHRec(1) As TransRecType
  
  OpenEmpIdxLNameFile IdxLHandle
  NumOfRecs = LOF(IdxLHandle) / 2
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  For x = 1 To NumOfRecs
    Get IdxLHandle, x, IdxBuff(x)
  Next x
  'got input here

  RptFile = FreeFile
  'this errorhandler line traps if there is no
  'diskette is in the A: drive
  On Local Error GoTo ErrorHandler
  Open RptName$ For Output As #RptFile
  
  Close RptFile

  RptFile = FreeFile
  Open RptName$ For Output As #RptFile
  
  PrintCnt = 0
  GoSub DoHeaderRec
  
  OpenTransHistFile hFile
  OpenEmpData2File EFile
  For RecNo = 1 To NumOfRecs
    UsingThisOne = False
    VCalcAmt# = 0
    GCalcAmt# = 0
    RothAmt# = 0
    Get EFile, IdxBuff(RecNo), Emp2Rec(1)
    If Emp2Rec(1).EMPTDATE <> 0 Then
      If LowDate > Emp2Rec(1).EMPTDATE Then
        GoTo SkipEm
      End If
    End If
    If Emp2Rec(1).LastTransRec <= 0 Then
      GoTo SkipEm
    End If

    TransRecNum& = Emp2Rec(1).LastTransRec
    Do
      Get hFile, TransRecNum&, TransHRec(1)
      Select Case TransHRec(1).CheckDate

      Case LowDate To HiDate
        If RCodeNum > 0 Then
          If TransHRec(1).DAmt(RCodeNum) <> 0 Then
            RothAmt# = OldRound#(RothAmt# + TransHRec(1).DAmt(RCodeNum))
            UsingThisOne = True
          End If
        End If
        If VCodeNum > 0 Then
          If TransHRec(1).DAmt(VCodeNum) <> 0 Then
            VCalcAmt# = OldRound#(VCalcAmt# + TransHRec(1).DAmt(VCodeNum))
            UsingThisOne = True
          End If
        End If
        EmpRType$ = UCase$(Left$(LTrim$(Emp2Rec(1).EMPRETTP), 1))
        If EmpRType$ = "" Then EmpRType$ = "G" 'added 10/8/03
        If EmpRType$ = "L" Or EmpRType$ = "G" Then 'L = Law G = General
          GCalcAmt# = OldRound#(GCalcAmt# + TransHRec(1).GrossPay)
          UsingThisOne = True
            For ECnt = 1 To 3 'True means subtract out this earnings code from matching
              If TransHRec(1).Less401k(ECnt) = True Then
                GCalcAmt# = OldRound(GCalcAmt# - TransHRec(1).EAmt(ECnt))
              End If
            Next ECnt
        End If
      Case Else
      End Select

      If TransHRec(1).PrevTransRec <= 0 Then
        If UsingThisOne Then
          If EmpRType$ = "L" Then
            EPct# = LPct#
          Else
            EPct# = GPct#
          End If
          GoSub PrintThisOneV
        End If
        Exit Do
      Else
        TransRecNum& = CLng(TransHRec(1).PrevTransRec)
      End If

    Loop

SkipEm:
  Next
  
  'now do loans
  For RecNo = 1 To NumOfRecs
    UsingThisOne = False
    LCalcAmt# = 0
    Get EFile, IdxBuff(RecNo), Emp2Rec(1)
    If Emp2Rec(1).EMPTDATE <> 0 Then
      If LowDate > Emp2Rec(1).EMPTDATE Then
        GoTo SkipEmL
      End If
    End If
    If Emp2Rec(1).LastTransRec <= 0 Then
      GoTo SkipEmL
    End If

    TransRecNum& = Emp2Rec(1).LastTransRec
    Do
      Get hFile, TransRecNum&, TransHRec(1)
      Select Case TransHRec(1).CheckDate
      Case LowDate To HiDate
        If LCodeNum > 0 Then
          If TransHRec(1).DAmt(LCodeNum) <> 0 Then
            LCalcAmt# = OldRound#(LCalcAmt# + TransHRec(1).DAmt(LCodeNum))
            UsingThisOne = True
          End If
        End If
      Case Else
      End Select
      If TransHRec(1).PrevTransRec <= 0 Then
        If UsingThisOne Then
          GoSub PrintThisOneL
        End If
        Exit Do
      Else
        TransRecNum& = CLng(TransHRec(1).PrevTransRec)
      End If
    Loop
SkipEmL:
  Next

  GoSub DoTrailerRec

EndTheProg:
  Close
  Unload frmSave401K2Dir
  MainLog ("Supplemental/Roth Retirement data saved to Citipak Directory.")
  Exit Sub
  
PrintThisOneV:
  SupDet(1) = SupDet(2)
  If QPTrim$(Emp2Rec(1).YN401K) <> "Y" Then 'added 8/28/03
    TMatchAmt# = 0 'filter out non participators only when
    'matching amounts are calculated
    If VCalcAmt# = 0 And RothAmt# = 0 Then 'added RCalcAmt for Roth
      GoTo SkipEMBubba
    Else
      GoSub No401KMatchDisk
      Return 'added this return on 9/28/06 brought to my attention by Calabash
    End If
  Else
    TMatchAmt# = OldRound((GCalcAmt# * EPct#) * 0.01)
  End If
  
  If EPct# > 0 Or VCalcAmt# > 0 Or RothAmt > 0 Then 'Or LCalcAmt# > 0 Then
    If LimitFlag = True And EmpRType$ = "G" Then 'added 8/28/03
      If VCalcAmt# = 0 And RothAmt# = 0 Then
        If LCalcAmt# = 0 Then 'If a "G" employee does not have
        'a deduction for either voluntary or loan then nothing
        'is printed to disk
          GoTo SkipEMBubba
        Else
          TMatchAmt# = 0
        End If
      Else '"G" employee with VCalcAmt > 0
      'if employee contributes as much as
      'or more than the % entered on the screen then the employer
      'also matches the % amount on the screen (but no more)...however,
      'if the employee does not contribute as much or more than the
      '% on the screen then the employer only matches what the employee
      'contributes...ex. (employee earns $1000...employer willing to contribute
      '5% (on screen %)...employee contributes $100 (10%) then the employer
      'contributes $50 (5%)...however if the employee contributes $20(2% which is
      'less than 5%) then the employer only contributes 2% ($20)
        If TMatchAmt# > OldRound(VCalcAmt# + RothAmt#) Then
          TMatchAmt# = OldRound(VCalcAmt# + RothAmt#)
        End If
      End If
    End If
  End If
  
No401KMatchDisk:
  EPrinted = EPrinted + 1
  TotalVAmt# = OldRound#(TotalVAmt# + VCalcAmt#)
  TotalRAmt# = OldRound#(TotalRAmt + RothAmt#)
    'TotalMatchAmt only collects TMatchAmt values if
    '   1) EPct or VCalcAmt > 0 (EPct is either
    '      the top % on screen if employee is G or bottom if L
    '   2) Employee is G and VCalcAmt > 0
    '   3) Employee is G and VCalcAmt = 0 but LCalcAmt > 0 (TMatch is 0 in this case)
    '
  TotalMatchAmt# = OldRound#(TotalMatchAmt# + TMatchAmt#)
  GoSub DoDetailRecs
Return

PrintThisOneL:
  TotalLAmt# = OldRound#(TotalLAmt# + LCalcAmt#)
  SupLoans(1) = SupLoans(2)
  If LCalcAmt# > 0 Then
    GoSub DoLoanRecs
      Return
    End If
SkipEMBubba:
Return
  
DoHeaderRec:
  PrintLine = ""
  SupHeader.RecType = "001"
  PrintLine = SupHeader.RecType
  SupHeader.FileType = "COMBINED  "
  PrintLine = PrintLine + SupHeader.FileType
  SupHeader.ClientID = "002003" 'QPTrim$(Unit(1).BBTCNTNO)
  PrintLine = PrintLine + SupHeader.ClientID
  SupHeader.CreateDate = Format(Date, "yyyymmdd")
  PrintLine = PrintLine + SupHeader.CreateDate
  SupHeader.CreateTime = Format(Time, "hhmmss")
  PrintLine = PrintLine + SupHeader.CreateTime
  SupHeader.Sender = "P"
  PrintLine = PrintLine + SupHeader.Sender
  'SupHeader.Filler260 = " "
  'PrintLine = PrintLine + SupHeader.Filler260
  PrintLine = PrintLine + Space(283)
  Print #RptFile, PrintLine
Return
  
DoTrailerRec:
  PrintLine = ""
  SupTrailer.RecordType = "999"
  PrintLine = SupTrailer.RecordType
  SupTrailer.FileType = "COMBINED  "
  PrintLine = PrintLine + SupTrailer.FileType
  SupTrailer.ClientID = "002003" ' QPTrim$(Unit(1).BBTCNTNO)
  PrintLine = PrintLine + SupTrailer.ClientID
  SupTrailer.CreateDate = SupHeader.CreateDate
  PrintLine = PrintLine + SupTrailer.CreateDate
  SupTrailer.CreateTime = SupHeader.CreateTime
  PrintLine = PrintLine + SupTrailer.CreateTime
  ThisAmtD = PrintCnt
  GoSub HandleAmtRecCnt
  SupTrailer.DetRecCnt = ThisAmtS$
  PrintLine = PrintLine + ThisAmtS$
  SupTrailer.Filler1 = ""
  PrintLine = PrintLine + SupTrailer.Filler1
  ThisAmtD = OldRound#(TotalVAmt# + TotalRAmt# + TotalMatchAmt#)
  GoSub HandleAmt11
  SupTrailer.TotContrib = ThisAmtS$
  PrintLine = PrintLine + SupTrailer.TotContrib
  PrintLine = PrintLine + SupTrailer.Filler1
  ThisAmtD = TotalLAmt#
  GoSub HandleAmt11
  SupTrailer.TotLnRePay = ThisAmtS$
  PrintLine = PrintLine + SupTrailer.TotLnRePay
'  SupTrailer.Filler228 = ""
'  PrintLine = PrintLine + SupTrailer.Filler228
  PrintLine = PrintLine + Space(251)
  Print #RptFile, PrintLine
Return
  
DoDetailRecs:
  PrintLine = ""
  SupDet(1).TransCode = "114"
  PrintLine = SupDet(1).TransCode
  SupDet(1).SeqCode = "01"
  PrintLine = PrintLine + SupDet(1).SeqCode
  SupDet(1).PlanNum = "002003"
  PrintLine = PrintLine + SupDet(1).PlanNum
  SupDet(1).Filler3 = ""
  PrintLine = PrintLine + SupDet(1).Filler3
  SupDet(1).ParticID = QPTrim$(Emp2Rec(1).EmpSSN)
  SupDet(1).ParticID = ReplaceString(SupDet(1).ParticID, "-", "")
  PrintLine = PrintLine + SupDet(1).ParticID
  SupDet(1).SubPlanNum = QPTrim$(Unit(1).BBTCNTNO) '"      "
  PrintLine = PrintLine + SupDet(1).SubPlanNum
  SupDet(1).Filler6 = ""
  PrintLine = PrintLine + SupDet(1).Filler6
  SupDet(1).Investment1 = "**"
  PrintLine = PrintLine + SupDet(1).Investment1
  SupDet(1).SourceEmpe = "A"
  PrintLine = PrintLine + SupDet(1).SourceEmpe
  ThisAmtD = VCalcAmt#
  GoSub HandleAmt9
  SupDet(1).ConEmpeAmt1 = ThisAmtS$
  PrintLine = PrintLine + SupDet(1).ConEmpeAmt1
  SupDet(1).Investment2 = "**"
  PrintLine = PrintLine + SupDet(1).Investment2
  SupDet(1).SourceEmpr = "P"
  PrintLine = PrintLine + SupDet(1).SourceEmpr
  ThisAmtD = TMatchAmt#
  GoSub HandleAmt9
  SupDet(1).ConEmprAmt2 = ThisAmtS$
  PrintLine = PrintLine + SupDet(1).ConEmprAmt2
  SupDet(1).Investment3 = "**"
  PrintLine = PrintLine + SupDet(1).Investment3
  SupDet(1).SourceRoth = "U"
  PrintLine = PrintLine + SupDet(1).SourceRoth
  ThisAmtD = RothAmt#
  GoSub HandleAmt9
  SupDet(1).ConRothAmt3 = ThisAmtS$
  PrintLine = PrintLine + SupDet(1).ConRothAmt3
  SupDet(1).Filler36 = ""
  PrintLine = PrintLine + SupDet(1).Filler36
  SupDet(1).PayPrdEndDt = Format(fptxtEnd.Text, "YYYYMMDD")
  PrintLine = PrintLine + SupDet(1).PayPrdEndDt
  If Emp2Rec(1).EMPFEDS = "M" Then
    SupDet(1).MarStatus = "2"
  Else
    SupDet(1).MarStatus = "1"
  End If
  PrintLine = PrintLine + SupDet(1).MarStatus
  SupDet(1).Filler4 = ""
  PrintLine = PrintLine + SupDet(1).Filler4
  GoSub WorkTheName
  SupDet(1).Name = TruncName 'QPTrim$(Emp2Rec(1).EmpLName) + " " + QPTrim$(Emp2Rec(1).EmpFName) + "   "
  PrintLine = PrintLine + SupDet(1).Name
  
  GoSub MakePrudInfo
  PrintLine = PrintLine + PrudAddInfo$
  'EmpAddr$ PrudInfo1$

  PrintCnt = PrintCnt + 1
  Print #RptFile, PrintLine
  
Return

'082113 Changed for Prudential format
MakePrudInfo:
    PrudAddInfo$ = Space(171)
    'PrudInfo1$ = Space(82)
    'EmpAddr$ = Space(89)
    Mid$(PrudAddInfo$, 1, 30) = QPTrim(Emp2Rec(1).EmpAddr1)
  '151-180 30 Addr1  '181-210 30 addr2
    Mid$(PrudAddInfo$, 61, 18) = QPTrim(Emp2Rec(1).EmpCity)
  '211 - 228 X(18) CITY  YES
    Mid$(PrudAddInfo$, 79, 2) = QPTrim(Emp2Rec(1).EmpState)
  '229 - 230 X(2) STATE  YES
    Mid$(PrudAddInfo$, 81, 9) = QPTrim(Emp2Rec(1).EmpZip)
  '231 - 239 X(9) ZIPCODE  YES
    Mid$(PrudAddInfo$, 90, 8) = Format(MakeRegDate(Emp2Rec(1).EMPBDAY), "YYYYMMDD")
  '240 - 247 X(8) DATE OF BIRTH CCYYMMDD YES YES
    Mid$(PrudAddInfo$, 98, 8) = Format(MakeRegDate(Emp2Rec(1).EMPHDATE), "YYYYMMDD")
  '248 - 255 X(8) ORIGINAL DATE OF HIRE CCYYMMDD (DATE FIRST HIRED)
    Mid$(PrudAddInfo$, 106, 8) = "        "    'Format(MakeRegDate(Emp2Rec(1).EMPHDATE), "YYYYMMDD")
  '256 - 263 X(8) ADJ. HIRE DATE CCYYMMDD OPTIONAL
    Mid$(PrudAddInfo$, 114, 8) = "00000000" 'Format(MakeRegDate(Emp2Rec(1).EMPHDATE), "YYYYMMDD")
  '264 - 271 9(8) YEARS OF SERVICE ‘00000000’  OPTIONAL (ZERO FILL IF NOT SENDING)
    Mid$(PrudAddInfo$, 122, 6) = "000000" 'Format(MakeRegDate(Emp2Rec(1).EMPHDATE), "YYYYMMDD")
  '272 - 277 S9(4)V99 YTD HOURS OF SERVICE ‘000000’ OPTIONAL TOTAL HOURS WORKED FOR THE PLAN YEAR. (ZERO FILL IF NOT SENDING)
    If Left$(QPTrim$(Emp2Rec(1).EMPGENDR), 1) = "F" Then
      Mid$(PrudAddInfo$, 128, 1) = "2"
    Else
      Mid$(PrudAddInfo$, 128, 1) = "1"
    End If
  '278 X GENDER CODE ‘1’ = MALE, ‘2’ = FEMALE YES
    Mid$(PrudAddInfo$, 129, 1) = " "
  '279 X FILLER SPACES/BLANK YES
    Select Case UCase(QPTrim$(Emp2Rec(1).EMPPFREQ))
    Case "WEEKLY"
      PrudTmp$ = "7"
    Case "BI-WEEKLY"
      PrudTmp$ = "6"
    Case "SEMI-MONTHLY"
      PrudTmp$ = "5"
    Case "MONTHLY"
      PrudTmp$ = "4"
    Case "QUARTERLY"
      PrudTmp$ = "3"
    Case "SEMI-ANNUALLY"
      PrudTmp$ = "2"
    Case "ANNUALLY"
      PrudTmp$ = "1"
    Case Else
      PrudTmp$ = "X"
    End Select
    Mid$(PrudAddInfo$, 130, 1) = PrudTmp$
  '280 X PAYROLL FREQUENCY    YES (LOAN FREQUENCY) ‘3’ = QUARTERLY  ‘4’ = MONTHLY ‘5’ = SEMI-MONTHLY ‘6’ = BI-WEEKLY ‘7’ = WEEKLY
    PrudTmp$ = MakeRegDate(Emp2Rec(1).EMPTDATE)
    If PrudTmp$ <> "12/31/1979" Then
      Mid$(PrudAddInfo$, 131, 2) = "32"
    Else
      Mid$(PrudAddInfo$, 131, 2) = "00"
    End If
'Comment this vvvvvv
  '281 - 282 X(2) EE STATUS ‘00’ = ACTIVE  YES ‘32’ = TERMINATED – PAYMENT DEFERRED ‘3D’ = DEATH   YES
  '283-289 7 Filler
    Mid$(PrudAddInfo$, 133, 7) = " "
    If InStr(Emp2Rec(1).EMPRETTP, "LAW") > 0 Then
      Mid$(PrudAddInfo$, 140, 1) = "Y"
    Else
      Mid$(PrudAddInfo$, 140, 1) = "N"
    End If
    
  '290 X(1) LEO INDICTATOR “N” = NOT LEO MUST BE VALUED IF LEFT BLANK WILL REMOVE LEO INDICTATOR
    'Mid$(PrudInfo1$, 52, 8) = "DATEDATE"
  '291 - 298 X(8) STATUS DATE CCYYMMDD YES
    'Mid$(PrudInfo1$, 60, 13) = "EMPIDEMPID..."
  '299-311 X(13) EMPLOYEE ID  OPTIONAL
    'Mid$(PrudInfo1$, 73, 2) = ""
  '312-313 X(2) SUB STATUS UL – Leave of absence  YES nnRL – Return from Leave
    'Mid$(PrudInfo1$, 75, 8) = "87654321"
  '314-321 X(8) SUB STATUS DATE CCYYMMDD YES

Return

DoLoanRecs:
  PrintLine = ""
  SupLoans(1).TransCode = "385"
  PrintLine = SupLoans(1).TransCode
  SupLoans(1).Filler2 = ""
  PrintLine = PrintLine + SupLoans(1).Filler2
  SupLoans(1).PlanNum = "002003"
  PrintLine = PrintLine + SupLoans(1).PlanNum
  SupLoans(1).Filler3A = ""
  PrintLine = PrintLine + SupLoans(1).Filler3A
  SupLoans(1).ParticID = QPTrim$(Emp2Rec(1).EmpSSN)
  SupLoans(1).ParticID = ReplaceString(SupLoans(1).ParticID, "-", "")
  PrintLine = PrintLine + SupLoans(1).ParticID
  SupLoans(1).SubPlanNum = QPTrim$(Unit(1).BBTCNTNO) '"      "
  PrintLine = PrintLine + SupLoans(1).SubPlanNum
  SupLoans(1).DateXOverRide = "1"
  PrintLine = PrintLine + SupLoans(1).DateXOverRide
  SupLoans(1).Filler4 = ""
  PrintLine = PrintLine + SupLoans(1).Filler4
  ThisAmtD = LCalcAmt#
  GoSub HandleAmt9
  SupLoans(1).LoanRepay = ThisAmtS$
  PrintLine = PrintLine + SupLoans(1).LoanRepay
  SupLoans(1).Filler10 = ""
  PrintLine = PrintLine + SupLoans(1).Filler10
  SupLoans(1).LoanPayOverRide = "2"
  PrintLine = PrintLine + SupLoans(1).LoanPayOverRide
  SupLoans(1).Filler3B = ""
  PrintLine = PrintLine + SupLoans(1).Filler3B
  SupLoans(1).LoanPayRollCode = "PDED"
  PrintLine = PrintLine + SupLoans(1).LoanPayRollCode
  SupLoans(1).Filler237 = ""
  PrintLine = PrintLine + Space(260)
  'PrintLine = PrintLine + SupLoans(1).Filler237
  PrintCnt = PrintCnt + 1
  Print #RptFile, PrintLine
 
Return
  
LoadDedCodes:
  OpenDedCodeFile DedFile
  NumOfDed = LOF(DedFile) / Len(DedCode(1))
  ReDim Item$(1 To NumOfDed)
  For cnt = 1 To NumOfDed
    Get DedFile, cnt, DedCode(1)
    Item$(cnt) = Str$(cnt) + ") " + DedCode(1).DCDESC1
  Next
  Close
Return

HandleAmtRecCnt:
  ThisAmtS$ = "0"
  NegChar2$ = "X"
  If ThisAmtD# < 0 Then GoSub GetNegChar
  ThisAmtS = CStr(ThisAmtD#)
  ThisAmtS = ReplaceString(ThisAmtS, "$", "")
  ThisAmtS = ReplaceString$(ThisAmtS, ",", "")
  ThisAmtS = ReplaceString$(ThisAmtS, ".", "")
  If NegChar2$ <> "X" Then
    ThisAmtS$ = Mid(ThisAmtS$, 1, StringLen - 1) + NegChar2$
  End If
  StringLen = Len(ThisAmtS$)
  ThisAmtS = String(9 - StringLen, "0") + ThisAmtS$
  
Return
  
HandleAmt9:
  ThisAmtS$ = "0"
  NegChar2$ = "X"
  If ThisAmtD# < 0 Then GoSub GetNegChar
  ThisAmtS = CStr(ThisAmtD#)
  If InStr(ThisAmtS, ".") = 0 Then
    ThisAmtS = ThisAmtS + "00"
  End If
  If InStr(ThisAmtS$, ".") = Len(ThisAmtS$) - 1 Then
    ThisAmtS$ = ThisAmtS$ + "0"
  End If
  ThisAmtS = ReplaceString(ThisAmtS, "$", "")
  ThisAmtS = ReplaceString$(ThisAmtS, ",", "")
  ThisAmtS = ReplaceString$(ThisAmtS, ".", "")
  If NegChar2$ <> "X" Then
    ThisAmtS$ = Mid(ThisAmtS$, 1, StringLen - 1) + NegChar2$
  End If
  StringLen = Len(ThisAmtS$)
  ThisAmtS = String(9 - StringLen, "0") + ThisAmtS$
  
  Return
  
HandleAmt11:
  ThisAmtS$ = "0"
  NegChar2$ = "X"
  If ThisAmtD# < 0 Then GoSub GetNegChar
  ThisAmtS = CStr(ThisAmtD#)
  If InStr(ThisAmtS, ".") = 0 Then
    ThisAmtS = ThisAmtS + "00"
  End If
  If InStr(ThisAmtS$, ".") = Len(ThisAmtS$) - 1 Then
    ThisAmtS$ = ThisAmtS$ + "0"
  End If
  ThisAmtS = ReplaceString(ThisAmtS, "$", "")
  ThisAmtS = ReplaceString$(ThisAmtS, ",", "")
  ThisAmtS = ReplaceString$(ThisAmtS, ".", "")
  If NegChar2$ <> "X" Then
    ThisAmtS$ = Mid(ThisAmtS$, 1, StringLen - 1) + NegChar2$
  End If
  StringLen = Len(ThisAmtS$)
  ThisAmtS = String(11 - StringLen, "0") + ThisAmtS$
  
  Return
  
GetNegChar:
  NegChar1$ = Mid(ThisAmtS, Len(ThisAmtS$), 1)
  Select Case NegChar1$
    Case "0"
      NegChar2$ = "}"
    Case "1"
      NegChar2 = "J"
    Case "2"
      NegChar2 = "K"
    Case "3"
      NegChar2 = "L"
    Case "4"
      NegChar2 = "M"
    Case "5"
      NegChar2 = "N"
    Case "6"
      NegChar2 = "O"
    Case "7"
      NegChar2 = "P"
    Case "8"
      NegChar2 = "Q"
    Case "9"
      NegChar2 = "R"
    Case Else
  End Select
  
  Return
  
WorkTheName:
  BuildName = ""
  TruncName = ""
  WorkName = ""
  FinalName$ = ""
  
  ThisLen = 0
  WorkName = QPTrim$(Emp2Rec(1).EmpLName) + ","
  ThisLen = Len(WorkName)
  If ThisLen > 0 Then
    For z = 1 To ThisLen
      Thisch = Mid(WorkName, z, 1)
      If Thisch <> " " Then
        BuildName = BuildName + Thisch
      End If
    Next z
  End If
  FinalName = BuildName + " "
  BuildName = ""
  WorkName = QPTrim$(Emp2Rec(1).EmpFName)
  ThisLen = Len(WorkName)
  If ThisLen > 0 Then
    For z = 1 To ThisLen
      Thisch = Mid(WorkName, z, 1)
      If Thisch <> " " Then
        BuildName = BuildName + Thisch
      ElseIf Thisch = " " Then
        If Mid(WorkName, z + 1, 1) <> " " Then
          BuildName = BuildName + " " + Mid(WorkName, z + 1, 1)
          Exit For
        End If
      End If
    Next z
  End If
  FinalName = FinalName + BuildName
  TruncName = FinalName

  Return
  
ErrorHandler1:
  Unload frmSave401K2Dir
  Close
  MsgBox "ERROR: A problem has occurred in writing the file nk0" + QPTrim$(Unit(1).BBTCNTNO) + ".txt to the 401K directory. Delete or rename the 401K directory and try again. If this doesn't solve the problem call Southern Software Support @ 1-800-842-8190."
'  MsgBox "ERROR: The 401K directory only exists to contain the file nk0" + QPTrim$(Unit(1).BBTCNTNO) + ".txt. If the 401K directory exists but does not contain this file or is empty then please delete the entire directory. Payroll will create the correct file while processing the Supplemental Retirement report. File not saved."
  Exit Sub
 
ErrorHandler:
  Unload frmSave401K2Dir
  MsgBox "ERROR: If this problem persists please consult Southern Software."
  Close
  Unload frmProcessing

End Sub
