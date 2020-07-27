VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmTaxRealClassRpt 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Real Property Classification Report"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   Icon            =   "frmTaxRealClassRpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   6396
      Left            =   1920
      TabIndex        =   0
      Top             =   1140
      Width           =   7788
      _Version        =   196609
      _ExtentX        =   13737
      _ExtentY        =   11282
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmTaxRealClassRpt.frx":08CA
      Begin LpLib.fpCombo fpcmbIncInactive 
         Height          =   384
         Left            =   3048
         TabIndex        =   4
         Top             =   3330
         Width           =   3444
         _Version        =   196608
         _ExtentX        =   6075
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
         ColDesigner     =   "frmTaxRealClassRpt.frx":08E6
      End
      Begin LpLib.fpCombo fpcmbClass 
         Height          =   384
         Left            =   3048
         TabIndex        =   5
         Top             =   3912
         Width           =   3444
         _Version        =   196608
         _ExtentX        =   6075
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
         BackColor       =   16777215
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
         ColDesigner     =   "frmTaxRealClassRpt.frx":0C51
      End
      Begin LpLib.fpCombo fpcmbTownship 
         Height          =   384
         Left            =   3048
         TabIndex        =   1
         Top             =   1632
         Width           =   3444
         _Version        =   196608
         _ExtentX        =   6075
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
         ColDesigner     =   "frmTaxRealClassRpt.frx":0FBC
      End
      Begin LpLib.fpCombo fpcmbPrintOpt 
         Height          =   384
         Left            =   3048
         TabIndex        =   2
         Top             =   2208
         Width           =   3444
         _Version        =   196608
         _ExtentX        =   6075
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
         ColDesigner     =   "frmTaxRealClassRpt.frx":1327
      End
      Begin LpLib.fpCombo fpcmbPrintOrder 
         Height          =   384
         Left            =   3048
         TabIndex        =   6
         Top             =   4464
         Width           =   3444
         _Version        =   196608
         _ExtentX        =   6075
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
         BackColor       =   16777215
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
         ColDesigner     =   "frmTaxRealClassRpt.frx":1692
      End
      Begin LpLib.fpCombo fpcmbRptOpt 
         Height          =   384
         Left            =   3048
         TabIndex        =   3
         Top             =   2760
         Width           =   3444
         _Version        =   196608
         _ExtentX        =   6075
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
         BackColor       =   16777215
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
         ColDesigner     =   "frmTaxRealClassRpt.frx":19FD
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   645
         Left            =   2040
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "Press the 'Cancel' button to exit this screen and return to the main 'Business License Reports' menu."
         Top             =   5370
         Width           =   1740
         _Version        =   131072
         _ExtentX        =   3069
         _ExtentY        =   1138
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
         ButtonDesigner  =   "frmTaxRealClassRpt.frx":1D68
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   645
         Left            =   4275
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   $"frmTaxRealClassRpt.frx":1F46
         Top             =   5370
         Width           =   1740
         _Version        =   131072
         _ExtentX        =   3069
         _ExtentY        =   1138
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
         ButtonDesigner  =   "frmTaxRealClassRpt.frx":1FF1
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Account Status:"
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
         Left            =   1176
         TabIndex        =   15
         Top             =   3426
         Width           =   1716
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Class:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1464
         TabIndex        =   14
         Top             =   4008
         Width           =   1428
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Real Township:"
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
         Left            =   1152
         TabIndex        =   13
         Top             =   1704
         Width           =   1740
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   3900
         Left            =   888
         Top             =   1248
         Width           =   5976
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Print Order:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1584
         TabIndex        =   12
         Top             =   4548
         Width           =   1308
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Real Property Classification Report"
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
         Height          =   396
         Left            =   1440
         TabIndex        =   11
         Top             =   456
         Width           =   4932
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   696
         Left            =   1176
         Top             =   312
         Width           =   5388
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Report Type:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1392
         TabIndex        =   10
         Top             =   2280
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Report Option:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1224
         TabIndex        =   9
         Top             =   2844
         Width           =   1668
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   6660
      Left            =   1800
      Top             =   996
      Width           =   8052
   End
End
Attribute VB_Name = "frmTaxRealClassRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim UseOpt As String * 1
  Dim ThisOpt$
  Dim ClassType$
  Dim TownName$

Private Sub cmdExit_Click()
  frmTaxReportsMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdProcess_Click()
  If fpcmbPrintOpt.Text = "Graphical" Then
    Call PrintGraphics
  Else
    frmTaxMsg.Label1.Caption = "Pitch 10 is recommended for this printout."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Call PrintText
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
      SendKeys "%P"
      Call cmdProcess_Click
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
  Me.HelpContextID = hlpRealPropClass
  Call LoadMe

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxRealClassRpt.")
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

Private Sub fpcmbClass_Change()
  ClassType$ = QPTrim$(fpcmbClass.Text)
End Sub

Private Sub fpcmbClass_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbClass.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbClass.ListIndex = -1
  End If
  If fpcmbClass.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbPrintOrder.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbIncInactive_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbIncInactive.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbIncInactive.ListIndex = -1
  End If
  If fpcmbIncInactive.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbClass.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbPrintOpt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbPrintOpt.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPrintOpt.ListIndex = -1
  End If
  If fpcmbPrintOpt.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbRptOpt.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbPrintOrder_Change()
  If ThisOpt <> "" Then
    If InStr(fpcmbPrintOrder.Text, ThisOpt) Then
      UseOpt = "Y"
    Else
      UseOpt = "N"
    End If
  End If
End Sub

Private Sub fpcmbPrintOrder_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbPrintOrder.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPrintOrder.ListIndex = -1
  End If
  If fpcmbPrintOrder.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbTownship.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbRptOpt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbRptOpt.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbRptOpt.ListIndex = -1
  End If
  If fpcmbRptOpt.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbIncInactive.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbTownship_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbTownship.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbTownship.ListIndex = -1
  End If
  If fpcmbTownship.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbPrintOpt.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub LoadMe()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim TSRec As TownshipType
  Dim TSCnt As Integer
  Dim TSHandle As Integer
  Dim x As Integer
  
  If Exist(TaxTownships) Then
    fpcmbTownship.Text = "ALL"
    fpcmbTownship.AddItem "ALL"
    OpenTownshipFile TSHandle, TSCnt
    For x = 1 To TSCnt
      Get TSHandle, x, TSRec
      fpcmbTownship.AddItem QPTrim$(TSRec.TownShip)
    Next x
    Close TSHandle
  Else
    fpcmbTownship.Text = "No Townships Saved"
  End If
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  TownName = QPTrim$(TaxMasterRec.Name)
  
  fpcmbPrintOpt.Text = "Graphical"
  fpcmbPrintOpt.AddItem "Graphical"
  fpcmbPrintOpt.AddItem "Text"
  fpcmbRptOpt.Text = "Address of Property"
  fpcmbRptOpt.AddItem "Address of Property"
  fpcmbRptOpt.AddItem "First Line of Notes"
  fpcmbPrintOrder.Text = "Name Order"
  fpcmbPrintOrder.AddItem "Name Order"
  fpcmbPrintOrder.AddItem "Acct Number Order"
  fpcmbPrintOrder.AddItem "Search Name"
  fpcmbClass.Text = "ALL"
  fpcmbClass.AddItem "ALL"
  fpcmbClass.AddItem "UNCLASSIFIED"
  fpcmbClass.AddItem "INDUSTRIAL"
  fpcmbClass.AddItem "COMMERCIAL"
  fpcmbClass.AddItem "PRIVATE"
  fpcmbIncInactive.Text = "Both"
  fpcmbIncInactive.AddItem "Both"
  fpcmbIncInactive.AddItem "Active Only"
  fpcmbIncInactive.AddItem "Inactive Only"
  For x = 1 To 6
    If QPTrim$(TaxMasterRec.ClassName(x)) <> "" Then
      fpcmbClass.AddItem QPTrim$(TaxMasterRec.ClassName(x))
    End If
  Next x
  
  ThisOpt = QPTrim$(TaxMasterRec.OptSrchCust)
  UseOpt = "N"
  If ThisOpt <> "" Then
    fpcmbPrintOrder.AddItem ThisOpt + " Order"
  End If
  
  ClassType$ = QPTrim$(fpcmbClass.Text)
End Sub

Private Sub PrintGraphics()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim IdxFlag As Boolean
  Dim RptFile$
  Dim RptHandle As Integer
  Dim Sub1RptFile$
  Dim Sub1RptHandle As Integer
  Dim Sub2RptFile$
  Dim Sub2RptHandle As Integer
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim dlm$
  Dim InactiveFlag As Boolean
  Dim x As Long, y As Long, z As Long
  Dim NextRec As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim PropAdd$, PropTownShip$
  Dim CustCnt As Long
  Dim CustRec As Long
  Dim CustName$
  Dim ThisTownship$
  Dim RealTotVal As Double
  Dim PersTotVal As Double
  Dim TotVal As Double
  Dim TotLLCnt As Long
  Dim TotRealLLCnt As Long
  Dim TotPersLLCnt As Long
  Dim ThisPersVal As Double
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim OptFlag As Boolean
  Dim RptOptFlag As Integer
  Dim Class$
  Dim RealVal As Double
  Dim RealDisc As Double
  Dim RealCustCnt As Long
  Dim GTotal As Double
  Dim ThisClass$
  Dim ThisClassType$
  Dim PrnCnt As Long
  Dim TownshipCnt As Integer
  Dim ThatTownship$
  Dim TSRec As TownshipType
  Dim TSCnt As Integer
  Dim TSHandle As Integer
  Dim TMHandle As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TSYN As Boolean
  Dim PrintClass$
  Dim TSAmtTot As Double
  Dim TSDiscTot As Double
  Dim TSNetTot As Double
  Dim ClassAmtTot As Double
  Dim ClassDiscTot As Double
  Dim ClassNetTot As Double
  Dim ClassNet As Double
  Dim ClassCntTot As Long
  Dim TSCntTot As Long
  Dim GAmtTot As Double
  Dim GDiscTot As Double
  Dim GNetTot As Double
  Dim GCntTot As Long
  Dim GTSAmtTot As Double
  Dim GTSDiscTot As Double
  Dim GTSNetTot As Double
  Dim GTSCntTot As Long
  
  'on error goto ERRORSTUFF
  
  InactiveFlag = False
  If fpcmbIncInactive.Text <> "Active Only" Then
    InactiveFlag = True
  End If
  
  TSYN = False
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  Class$ = fpcmbClass.Text
  ThisTownship = fpcmbTownship.Text
  
  If fpcmbRptOpt.Text = "Address of Property" Then
    RptOptFlag = 1
  Else
    RptOptFlag = 2
  End If
  
  IdxFlag = False
  OptFlag = False
  dlm$ = "~"
  If QPTrim$(fpcmbPrintOrder.Text) = "Name Order" Then
    OpenNameIdxFile IdxHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmTaxMsg.Label1.Caption = "There are no customers saved."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show
      fpcmbPrintOrder.SetFocus
      Close IdxHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get IdxHandle, x, IdxRec
      IdxArray(x) = IdxRec.CustRec
    Next x
    Close IdxHandle
    IdxFlag = True
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "Search Name" Then
    OpenSrchNameIdxFile SHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmTaxMsg.Label1.Caption = "There are no search names indexed."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show
      fpcmbPrintOrder.SetFocus
      Close SHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get SHandle, x, SearchRec
      IdxArray(x) = SearchRec.CustRec
    Next x
    Close SHandle
    IdxFlag = True
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = ThisOpt + " Order" Then
    OpenCustOptSearchFile OHandle, NumOfIdx
    If NumOfIdx = 0 Then
      Call TaxMsg(900, "There are no " + ThisOpt + "descriptions indexed.")
      Close OHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long
    
    For x = 1 To NumOfIdx
      Get OHandle, x, OptRec
      IdxArray(x) = OptRec.CustRec
    Next x
    Close OHandle
    IdxFlag = True
    OptFlag = True
  End If

  RptFile$ = "TAXRPTS\REALCLAS.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenRealPropFile RHandle, NumOfRealRecs
  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  End If
  frmTaxShowPctComp.Label1 = "Gathering Late Listing Data"
  frmTaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  
  ReDim ClassAmt(1 To 10) As Double
  ReDim ClassDisc(1 To 10) As Double
  ReDim ClassCnt(1 To 10) As Long
  
  If fpcmbTownship.Text <> "No Townships Saved" Then
    OpenTownshipFile TSHandle, TSCnt
    ReDim TownshipName(1 To TSCnt + 1, 1 To 10) As String
    ReDim TownshipAmt(1 To TSCnt + 1, 1 To 10) As Double
    ReDim TownshipDisc(1 To TSCnt + 1, 1 To 10) As Double
    ReDim TownshipNet(1 To TSCnt + 1, 1 To 10) As Double
    ReDim TownshipCount(1 To TSCnt + 1, 1 To 10) As Long
    For x = 1 To TSCnt
      Get TSHandle, x, TSRec
      For y = 1 To 10
        TownshipName(x, y) = QPTrim$(TSRec.TownShip)
      Next y
    Next x
    For y = 1 To 10
      TownshipName(TSCnt + 1, y) = "NOT SAVED"
    Next y
    Close TSHandle
  End If
    
  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
      CustRec = IdxArray(x)
    Else
      Get TCHandle, x, TaxCust
      CustRec = x
    End If
    If TaxCust.Deleted <> 0 Then GoTo SkipIt
    
    If InactiveFlag = False Then
      If TaxCust.Active = "N" Then GoTo SkipIt
    End If
    If TaxCust.FirstPropRec > 0 Then
      NextRec = TaxCust.FirstPropRec
      Do While NextRec > 0
        Get RHandle, NextRec, RealRec
        If RealRec.Deleted <> 0 Then GoTo MoveOn
        ThatTownship = QPTrim$(RealRec.TownShip)
        If ThatTownship = "" Then ThatTownship = "Not Saved"
        If ThisTownship <> "No Townships Saved" Then
          TSYN = True
          If ThisTownship <> "ALL" Then
            If ThisTownship <> ThatTownship Then
              GoTo MoveOn
            End If
          End If
        End If
        
        If ClassType <> "ALL" Then
          If ClassType <> QPTrim$(RealRec.ICPDesc) Then GoTo MoveOn
        End If
        ThisClassType = QPTrim$(RealRec.ICPDesc)
        If ThisClassType = "" Then
          ThisClassType = "UNCLASSIFIED"
        End If
        PrnCnt = PrnCnt + 1
        If InactiveFlag = True Then
          If TaxCust.Active = "N" Then
            '                     0                       1                        2
            Print #RptHandle, TownName; dlm; QPTrim$(TaxCust.CustName) + " **A**"; dlm; TaxCust.Acct; dlm;
          Else
            '                     0                       1                        2
            Print #RptHandle, TownName; dlm; QPTrim$(TaxCust.CustName); dlm; TaxCust.Acct; dlm;
          End If
        Else
          '                     0                       1                        2
          Print #RptHandle, TownName; dlm; QPTrim$(TaxCust.CustName); dlm; TaxCust.Acct; dlm;
        End If
        
        If ThisClassType = "INDUSTRIAL" Then
          If ClassType = "ALL" Then
            '                   3
            Print #RptHandle, "ALL"; dlm;
          Else
            '                      3
            Print #RptHandle, "INDUSTRIAL"; dlm;
          End If
          ThisClass = "INDUSTRIAL"
          ClassAmt(1) = OldRound(ClassAmt(1) + RealRec.PROPVALU)
          ClassDisc(1) = OldRound(ClassDisc(1) + (RealRec.EXMPOTHR + RealRec.EXMPSENI))
          ClassCnt(1) = ClassCnt(1) + 1
          If fpcmbTownship.Text = "No Townships Saved" Then GoTo SkipIt1
          If ThatTownship = "Not Saved" Then
            TownshipAmt(TSCnt + 1, 1) = OldRound(TownshipAmt(TSCnt + 1, 1) + RealRec.PROPVALU)
            TownshipDisc(TSCnt + 1, 1) = OldRound(TownshipDisc(TSCnt + 1, 1) + (RealRec.EXMPOTHR + RealRec.EXMPSENI))
            TownshipNet(TSCnt + 1, 1) = OldRound(TownshipNet(TSCnt + 1, 1) + (RealRec.PROPVALU - (RealRec.EXMPOTHR + RealRec.EXMPSENI)))
            TownshipCount(TSCnt + 1, 1) = OldRound(TownshipCount(TSCnt + 1, 1) + 1)
          Else
            For z = 1 To TSCnt + 1
              If TownshipName(z, 1) = ThatTownship Then
                TownshipAmt(z, 1) = OldRound(TownshipAmt(z, 1) + RealRec.PROPVALU)
                TownshipDisc(z, 1) = OldRound(TownshipDisc(z, 1) + (RealRec.EXMPOTHR + RealRec.EXMPSENI))
                TownshipNet(z, 1) = OldRound(TownshipNet(z, 1) + (RealRec.PROPVALU - (RealRec.EXMPOTHR + RealRec.EXMPSENI)))
                TownshipCount(z, 1) = OldRound(TownshipCount(z, 1) + 1)
              End If
            Next z
          End If
SkipIt1:
        ElseIf ThisClassType = "COMMERCIAL" Then
          If ClassType = "ALL" Then
            '                   3
            Print #RptHandle, "ALL"; dlm;
          Else
            '                       3
            Print #RptHandle, "COMMERCIAL"; dlm;
          End If
          ThisClass = "COMMERCIAL"
          ClassAmt(2) = OldRound(ClassAmt(2) + RealRec.PROPVALU)
          ClassDisc(2) = OldRound(ClassDisc(2) + (RealRec.EXMPOTHR + RealRec.EXMPSENI))
          ClassCnt(2) = ClassCnt(2) + 1
          If fpcmbTownship.Text = "No Townships Saved" Then GoTo SkipIt2
          If ThatTownship = "Not Saved" Then
            TownshipAmt(TSCnt + 1, 2) = OldRound(TownshipAmt(TSCnt + 1, 2) + RealRec.PROPVALU)
            TownshipDisc(TSCnt + 1, 2) = OldRound(TownshipDisc(TSCnt + 1, 2) + (RealRec.EXMPOTHR + RealRec.EXMPSENI))
            TownshipNet(TSCnt + 1, 2) = OldRound(TownshipNet(TSCnt + 1, 2) + (RealRec.PROPVALU - (RealRec.EXMPOTHR + RealRec.EXMPSENI)))
            TownshipCount(TSCnt + 1, 2) = OldRound(TownshipCount(TSCnt + 1, 2) + 1)
          Else
            For z = 1 To TSCnt
              If TownshipName(z, 2) = ThatTownship Then
                TownshipAmt(z, 2) = OldRound(TownshipAmt(z, 2) + RealRec.PROPVALU)
                TownshipDisc(z, 2) = OldRound(TownshipDisc(z, 2) + (RealRec.EXMPOTHR + RealRec.EXMPSENI))
                TownshipNet(z, 2) = OldRound(TownshipNet(z, 2) + (RealRec.PROPVALU - (RealRec.EXMPOTHR + RealRec.EXMPSENI)))
                TownshipCount(z, 2) = OldRound(TownshipCount(z, 2) + 1)
              End If
            Next z
          End If
SkipIt2:
        ElseIf ThisClassType = "PRIVATE" Then
          If ClassType = "ALL" Then
            '                   3
            Print #RptHandle, "ALL"; dlm;
          Else
            '                     3
            Print #RptHandle, "PRIVATE"; dlm;
          End If
          ThisClass = "PRIVATE"
          ClassAmt(3) = OldRound(ClassAmt(3) + RealRec.PROPVALU)
          ClassDisc(3) = OldRound(ClassDisc(3) + (RealRec.EXMPOTHR + RealRec.EXMPSENI))
          ClassCnt(3) = ClassCnt(3) + 1
          If fpcmbTownship.Text = "No Townships Saved" Then GoTo SkipIt3
          If ThatTownship = "Not Saved" Then
            TownshipAmt(TSCnt + 1, 3) = OldRound(TownshipAmt(TSCnt + 1, 3) + RealRec.PROPVALU)
            TownshipDisc(TSCnt + 1, 3) = OldRound(TownshipDisc(TSCnt + 1, 3) + (RealRec.EXMPOTHR + RealRec.EXMPSENI))
            TownshipNet(TSCnt + 1, 3) = OldRound(TownshipNet(TSCnt + 1, 3) + (RealRec.PROPVALU - (RealRec.EXMPOTHR + RealRec.EXMPSENI)))
            TownshipCount(TSCnt + 1, 3) = OldRound(TownshipCount(TSCnt + 1, 3) + 1)
          Else
            For z = 1 To TSCnt + 1
              If TownshipName(z, 3) = ThatTownship Then
                TownshipAmt(z, 3) = OldRound(TownshipAmt(z, 3) + RealRec.PROPVALU)
                TownshipDisc(z, 3) = OldRound(TownshipDisc(z, 3) + (RealRec.EXMPOTHR + RealRec.EXMPSENI))
                TownshipNet(z, 3) = OldRound(TownshipNet(z, 3) + (RealRec.PROPVALU - (RealRec.EXMPOTHR + RealRec.EXMPSENI)))
                TownshipCount(z, 3) = OldRound(TownshipCount(z, 3) + 1)
              End If
            Next z
          End If
SkipIt3:
        ElseIf ThisClassType = "UNCLASSIFIED" Then
          If ClassType = "ALL" Then
            '                   3
            Print #RptHandle, "ALL"; dlm;
          Else
            '                       3
            Print #RptHandle, "UNCLASSIFIED"; dlm;
          End If
          ThisClass = "UNCLASSIFIED"
          ClassAmt(4) = OldRound(ClassAmt(4) + RealRec.PROPVALU)
          ClassDisc(4) = OldRound(ClassDisc(4) + (RealRec.EXMPOTHR + RealRec.EXMPSENI))
          ClassCnt(4) = ClassCnt(4) + 1
          If fpcmbTownship.Text = "No Townships Saved" Then GoTo SkipIt4
          If ThatTownship = "Not Saved" Then
            TownshipAmt(TSCnt + 1, 4) = OldRound(TownshipAmt(TSCnt + 1, 4) + RealRec.PROPVALU)
            TownshipDisc(TSCnt + 1, 4) = OldRound(TownshipDisc(TSCnt + 1, 4) + (RealRec.EXMPOTHR + RealRec.EXMPSENI))
            TownshipNet(TSCnt + 1, 4) = OldRound(TownshipNet(TSCnt + 1, 4) + (RealRec.PROPVALU - (RealRec.EXMPOTHR + RealRec.EXMPSENI)))
            TownshipCount(TSCnt + 1, 4) = OldRound(TownshipCount(TSCnt + 1, 4) + 1)
          Else
            For z = 1 To TSCnt + 1
              If TownshipName(z, 4) = ThatTownship Then
                TownshipAmt(z, 4) = OldRound(TownshipAmt(z, 4) + RealRec.PROPVALU)
                TownshipDisc(z, 4) = OldRound(TownshipDisc(z, 4) + (RealRec.EXMPOTHR + RealRec.EXMPSENI))
                TownshipNet(z, 4) = OldRound(TownshipNet(z, 4) + (RealRec.PROPVALU - (RealRec.EXMPOTHR + RealRec.EXMPSENI)))
                TownshipCount(z, 4) = OldRound(TownshipCount(z, 4) + 1)
              End If
            Next z
          End If
SkipIt4:
        Else
          For y = 1 To 6
            If QPTrim$(TaxMasterRec.ClassName(y)) = ThisClassType Then
              If ClassType = "ALL" Then
                Print #RptHandle, "ALL"; dlm;
              Else
                Print #RptHandle, ThisClassType; dlm;
              End If
              ThisClass = ClassType
              ClassAmt(y + 4) = OldRound(ClassAmt(y + 4) + RealRec.PROPVALU)
              ClassDisc(y + 4) = OldRound(ClassDisc(y + 4) + (RealRec.EXMPOTHR + RealRec.EXMPSENI))
              ClassCnt(y + 4) = ClassCnt(y + 4) + 1
              If fpcmbTownship.Text = "No Townships Saved" Then GoTo SkipIt5
              If ThatTownship = "Not Saved" Then
                TownshipAmt(TSCnt + 1, y + 4) = OldRound(TownshipAmt(TSCnt + 1, y + 4) + RealRec.PROPVALU)
                TownshipDisc(TSCnt + 1, y + 4) = OldRound(TownshipDisc(TSCnt + 1, y + 4) + (RealRec.EXMPOTHR + RealRec.EXMPSENI))
                TownshipNet(TSCnt + 1, y + 4) = OldRound(TownshipNet(TSCnt + 1, y + 4) + (RealRec.PROPVALU - (RealRec.EXMPOTHR + RealRec.EXMPSENI)))
                TownshipCount(TSCnt + 1, y + 4) = OldRound(TownshipCount(TSCnt + 1, y + 4) + 1)
              Else
                For z = 1 To TSCnt + 1
                  If TownshipName(z, y + 4) = ThatTownship Then
                    TownshipAmt(z, y + 4) = OldRound(TownshipAmt(z, y + 4) + RealRec.PROPVALU)
                    TownshipDisc(z, y + 4) = OldRound(TownshipDisc(z, y + 4) + (RealRec.EXMPOTHR + RealRec.EXMPSENI))
                    TownshipNet(z, y + 4) = OldRound(TownshipNet(z, y + 4) + (RealRec.PROPVALU - (RealRec.EXMPOTHR + RealRec.EXMPSENI)))
                    TownshipCount(z, y + 4) = OldRound(TownshipCount(z, y + 4) + 1)
                  End If
                Next z
              End If
              Exit For
            End If
SkipIt5:
          Next y
          If y > 6 Then
            Print #RptHandle, "UNKNOWN"; dlm;
          End If
        End If
NoClass:
        GAmtTot = OldRound(GAmtTot + RealRec.PROPVALU)
        GDiscTot = OldRound(GDiscTot + RealRec.EXMPOTHR + RealRec.EXMPSENI)
        GNetTot = OldRound(GNetTot + RealRec.PROPVALU - (RealRec.EXMPOTHR + RealRec.EXMPSENI))
        GCntTot = GCntTot + 1
          '                      4                                      5                                        6
        Print #RptHandle, RealRec.PROPVALU; dlm; OldRound(RealRec.EXMPOTHR + RealRec.EXMPSENI); dlm; QPTrim$(RealRec.RealPin); dlm;
        If fpcmbRptOpt.Text = "Address of Property" Then
          If QPTrim$(RealRec.PropAddr) <> "" Then
            '                           7
            Print #RptHandle, QPTrim$(RealRec.PropAddr); dlm;
          Else
            '                      7
            Print #RptHandle, "Not Saved"; dlm;
          End If
        Else
          If QPTrim$(RealRec.PROPNOT1) <> "" Then
            '                             7
            Print #RptHandle, QPTrim$(RealRec.PROPNOT1); dlm;
          Else
            '                      7
            Print #RptHandle, "Not Saved"; dlm;
          End If
        End If
        '                                               8                                                 9                  10
        Print #RptHandle, OldRound(RealRec.PROPVALU - (RealRec.EXMPOTHR + RealRec.EXMPSENI)); dlm; fpcmbRptOpt.Text; dlm; ThisClass; dlm;
        '                      11                      12
        Print #RptHandle, ThatTownship; dlm; fpcmbTownship.Text; dlm;
        If UseOpt = "Y" Then
          '                  13                        14
          Print #RptHandle, ThisOpt; dlm; QPTrim$(TaxCust.OptSrchDesc); dlm;
        Else
          '                 13       14
          Print #RptHandle, ""; dlm; ""; dlm;
        End If
        '                      15              16             17
        Print #RptHandle, ThisClassType; dlm; TSYN; dlm; InactiveFlag
        
MoveOn:
        NextRec = RealRec.NextRec
      Loop
    End If
SkipIt:
    frmTaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      Exit Sub
    End If
  Next x
  If ClassType = "ALL" Then
    '                 0         1        2         3          4              5              6             7           8        9       10       11
    Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; "ALL"; dlm; GAmtTot; dlm; GDiscTot; dlm; GNetTot; dlm; GCntTot; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm;
  Else
    '                 0         1        2          3               4             5              6             7           8        9       10       11
    Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ThisClass; dlm; GAmtTot; dlm; GDiscTot; dlm; GNetTot; dlm; GCntTot; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm;
  End If
  
  '                         12               13       14             15             16             17
  Print #RptHandle, fpcmbTownship.Text; dlm; ""; dlm; ""; dlm; ThisClassType; dlm; TSYN; dlm; InactiveFlag
  
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  
  Close
  If PrnCnt = 0 Then
    Call TaxMsg(900, "There are no records to display using the criteria entered.")
    Exit Sub
  End If
  
  If fpcmbTownship.Text = "No Townships Saved" Then GoTo NoTownships
  Sub1RptFile$ = "TAXRPTS\SUB1REALCLAS.RPT"
  Sub1RptHandle = FreeFile
  Open Sub1RptFile For Output As #Sub1RptHandle
  GTSAmtTot = 0
  GTSDiscTot = 0
  GTSNetTot = 0
  GTSCntTot = 0
  
  For x = 1 To TSCnt + 1
    TSAmtTot = 0
    TSDiscTot = 0
    TSNetTot = 0
    TSCntTot = 0
    For y = 1 To 10
      If TownshipAmt(x, y) = 0 Then GoTo JumpIt
      TSAmtTot = OldRound(TSAmtTot + TownshipAmt(x, y))
      TSDiscTot = OldRound(TSDiscTot + TownshipDisc(x, y))
      TSNetTot = OldRound(TSNetTot + TownshipNet(x, y))
      TSCntTot = OldRound(TSCntTot + TownshipCount(x, y))
      GTSAmtTot = OldRound(GTSAmtTot + TownshipAmt(x, y))
      GTSDiscTot = OldRound(GTSDiscTot + TownshipDisc(x, y))
      GTSNetTot = OldRound(GTSNetTot + TownshipNet(x, y))
      GTSCntTot = OldRound(GTSCntTot + TownshipCount(x, y))
      If y >= 5 Then
        If QPTrim$(TaxMasterRec.ClassName(y - 4)) = "" Then GoTo JumpIt
        PrintClass = QPTrim$(TaxMasterRec.ClassName(y - 4))
        Print #Sub1RptHandle, TownshipName(x, y); dlm; TownshipAmt(x, y); dlm; TownshipDisc(x, y); dlm; TownshipNet(x, y); dlm; PrintClass; dlm;
        '                        5               6              7                  8                   9
        Print #Sub1RptHandle, TSAmtTot; dlm; TSDiscTot; dlm; TSNetTot; dlm; TownshipCount(x, y); dlm; TSCntTot; dlm;
        '                        10              11               12              13
        Print #Sub1RptHandle, GTSAmtTot; dlm; GTSDiscTot; dlm; GTSNetTot; dlm; GTSCntTot
        GoTo JumpIt
      End If
      Select Case y
        Case 1
          PrintClass = "INDUSTRIAL"
        Case 2
          PrintClass = "COMMERCIAL"
        Case 3
          PrintClass = "PRIVATE"
        Case 4
          PrintClass = "UNCLASSIFIED"
      End Select
      '                           0                         1                      2                         3                    4
      Print #Sub1RptHandle, TownshipName(x, y); dlm; TownshipAmt(x, y); dlm; TownshipDisc(x, y); dlm; TownshipNet(x, y); dlm; PrintClass; dlm;
      '                        5               6              7                  8                   9
      Print #Sub1RptHandle, TSAmtTot; dlm; TSDiscTot; dlm; TSNetTot; dlm; TownshipCount(x, y); dlm; TSCntTot; dlm;
      '                        10              11               12              13
      Print #Sub1RptHandle, GTSAmtTot; dlm; GTSDiscTot; dlm; GTSNetTot; dlm; GTSCntTot

JumpIt:
    Next y
  Next x
  Close
NoTownships:
  Sub2RptFile$ = "TAXRPTS\SUB2REALCLAS.RPT"
  Sub2RptHandle = FreeFile
  Open Sub2RptFile For Output As #Sub2RptHandle
  ClassAmtTot = 0
  ClassDiscTot = 0
  ClassNetTot = 0
  ClassCntTot = 0
  For y = 1 To 10
    If ClassAmt(y) = 0 Then GoTo JumpIt2
    ClassAmtTot = OldRound(ClassAmtTot + ClassAmt(y))
    ClassDiscTot = OldRound(ClassDiscTot + ClassDisc(y))
    ClassNet = OldRound(ClassAmt(y) - ClassDisc(y))
    ClassNetTot = OldRound(ClassNetTot + ClassNet)
    ClassCntTot = OldRound(ClassCntTot + ClassCnt(y))
    If y >= 5 Then
      If QPTrim$(TaxMasterRec.ClassName(y - 4)) = "" Then GoTo JumpIt2
      '                         0                  1                 2                           3                                 4
      Print #Sub2RptHandle, ClassAmt(y); dlm; ClassDisc(y); dlm; ClassCnt(y); dlm; QPTrim$(TaxMasterRec.ClassName(y - 4)); dlm; ClassNet; dlm;
      '                         5                  6                 7                 8
      Print #Sub2RptHandle, ClassAmtTot; dlm; ClassDiscTot; dlm; ClassNetTot; dlm; ClassCntTot
   Else
      Select Case y
        Case 1
          PrintClass = "INDUSTRIAL"
        Case 2
          PrintClass = "COMMERCIAL"
        Case 3
          PrintClass = "PRIVATE"
        Case 4
          PrintClass = "UNCLASSIFIED"
      End Select
    '                         0                  1                 2                 3               4
    Print #Sub2RptHandle, ClassAmt(y); dlm; ClassDisc(y); dlm; ClassCnt(y); dlm; PrintClass; dlm; ClassNet; dlm;
    '                         5                  6                 7                 8
    Print #Sub2RptHandle, ClassAmtTot; dlm; ClassDiscTot; dlm; ClassNetTot; dlm; ClassCntTot
   End If
JumpIt2:
  Next y
  Close
  
  arTaxRealClassRpt.Show
  frmTaxLoadingRpt.Show
  Exit Sub
  
ERRORSTUFF:
   Unload frmTaxShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxRealClassRpt", "PrintGraphics", Erl)
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

Private Sub PrintText()
  Dim FF$
  Dim Page As Integer
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim IdxFlag As Boolean
  Dim RptFile$
  Dim RptHandle As Integer
  Dim Sub1RptFile$
  Dim Sub1RptHandle As Integer
  Dim Sub2RptFile$
  Dim Sub2RptHandle As Integer
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim InactiveFlag As Boolean
  Dim x As Long, y As Long, z As Long
  Dim NextRec As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim PropAdd$, PropTownShip$
  Dim CustCnt As Long
  Dim CustRec As Long
  Dim CustName$
  Dim ThisTownship$
  Dim RealTotVal As Double
  Dim PersTotVal As Double
  Dim TotVal As Double
  Dim TotLLCnt As Long
  Dim TotRealLLCnt As Long
  Dim TotPersLLCnt As Long
  Dim ThisPersVal As Double
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim OptFlag As Boolean
  Dim RptOptFlag As Integer
  Dim Class$
  Dim RealVal As Double
  Dim RealDisc As Double
  Dim RealCustCnt As Long
  Dim GTotal As Double
  Dim ThisClass$
  Dim ThisClassType$
  Dim PrnCnt As Long
  Dim TownshipCnt As Integer
  Dim ThatTownship$
  Dim TSRec As TownshipType
  Dim TSCnt As Integer
  Dim TSHandle As Integer
  Dim TMHandle As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TSYN As Boolean
  Dim PrintClass$
  Dim TSAmtTot As Double
  Dim TSDiscTot As Double
  Dim TSNetTot As Double
  Dim ClassAmtTot As Double
  Dim ClassDiscTot As Double
  Dim ClassNetTot As Double
  Dim ClassNet As Double
  Dim ClassCntTot As Long
  Dim TSCntTot As Long
  Dim GAmtTot As Double
  Dim GDiscTot As Double
  Dim GNetTot As Double
  Dim GCntTot As Long
  Dim GTSAmtTot As Double
  Dim GTSDiscTot As Double
  Dim GTSNetTot As Double
  Dim GTSCntTot As Long
  Dim ThisName$
  Dim ThatName$
  Dim Description$
  Dim ThisTab As Integer
  Dim ThisLen As Integer
  
  'on error goto ERRORSTUFF
  
  InactiveFlag = False
  If fpcmbIncInactive.Text <> "Active Only" Then
    InactiveFlag = True
  End If
  FF$ = Chr(12)
  MaxLines = 58
  TSYN = False
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  Class$ = fpcmbClass.Text
  ThisTownship = fpcmbTownship.Text
  
  If fpcmbRptOpt.Text = "Address of Property" Then
    RptOptFlag = 1
  Else
    RptOptFlag = 2
  End If
  
  IdxFlag = False
  OptFlag = False
  If QPTrim$(fpcmbPrintOrder.Text) = "Name Order" Then
    OpenNameIdxFile IdxHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmTaxMsg.Label1.Caption = "There are no customers saved."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show
      fpcmbPrintOrder.SetFocus
      Close IdxHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get IdxHandle, x, IdxRec
      IdxArray(x) = IdxRec.CustRec
    Next x
    Close IdxHandle
    IdxFlag = True
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "Search Name" Then
    OpenSrchNameIdxFile SHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmTaxMsg.Label1.Caption = "There are no search names indexed."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show
      fpcmbPrintOrder.SetFocus
      Close SHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get SHandle, x, SearchRec
      IdxArray(x) = SearchRec.CustRec
    Next x
    Close SHandle
    IdxFlag = True
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = ThisOpt + " Order" Then
    OpenCustOptSearchFile OHandle, NumOfIdx
    If NumOfIdx = 0 Then
      Call TaxMsg(900, "There are no " + ThisOpt + "descriptions indexed.")
      Close OHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long
    
    For x = 1 To NumOfIdx
      Get OHandle, x, OptRec
      IdxArray(x) = OptRec.CustRec
    Next x
    Close OHandle
    IdxFlag = True
    OptFlag = True
  End If

  RptFile$ = "TAXRPTS\REALCLAS.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  GoSub PrintHeader
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenRealPropFile RHandle, NumOfRealRecs
  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  End If
  frmTaxShowPctComp.Label1 = "Gathering Late Listing Data"
  frmTaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  
  ReDim ClassAmt(1 To 10) As Double
  ReDim ClassDisc(1 To 10) As Double
  ReDim ClassCnt(1 To 10) As Long
  
  If fpcmbTownship.Text <> "No Townships Saved" Then
    OpenTownshipFile TSHandle, TSCnt
    ReDim TownshipName(1 To TSCnt + 1, 1 To 10) As String
    ReDim TownshipAmt(1 To TSCnt + 1, 1 To 10) As Double
    ReDim TownshipDisc(1 To TSCnt + 1, 1 To 10) As Double
    ReDim TownshipNet(1 To TSCnt + 1, 1 To 10) As Double
    ReDim TownshipCount(1 To TSCnt + 1, 1 To 10) As Long
    For x = 1 To TSCnt
      Get TSHandle, x, TSRec
      For y = 1 To 10
        TownshipName(x, y) = QPTrim$(TSRec.TownShip)
      Next y
    Next x
    For y = 1 To 10
      TownshipName(TSCnt + 1, y) = "NOT SAVED"
    Next y
    Close TSHandle
  End If
    
  ThatName = ""
  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
      CustRec = IdxArray(x)
    Else
      Get TCHandle, x, TaxCust
      CustRec = x
    End If
    If TaxCust.Deleted <> 0 Then GoTo SkipIt
    If InactiveFlag = False Then
      If TaxCust.Active = "N" Then GoTo SkipIt
    End If
    ThisName = QPTrim$(TaxCust.CustName)
    If TaxCust.FirstPropRec > 0 Then
      NextRec = TaxCust.FirstPropRec
      Do While NextRec > 0
        Get RHandle, NextRec, RealRec
        If RealRec.Deleted <> 0 Then GoTo MoveOn
        ThatTownship = QPTrim$(RealRec.TownShip)
        If ThatTownship = "" Then ThatTownship = "Not Saved"
        If ThisTownship <> "No Townships Saved" Then
          TSYN = True
          If ThisTownship <> "ALL" Then
            If ThisTownship <> ThatTownship Then
              GoTo MoveOn
            End If
          End If
        End If
        
        If ClassType <> "ALL" Then
          If ClassType <> QPTrim$(RealRec.ICPDesc) Then GoTo MoveOn
        End If
        ThisClassType = QPTrim$(RealRec.ICPDesc)
        If ThisClassType = "" Then
          ThisClassType = "UNCLASSIFIED"
        End If
        PrnCnt = PrnCnt + 1
        If ThisClassType = "INDUSTRIAL" Then
          ThisClass = "INDUSTRIAL"
          ClassAmt(1) = OldRound(ClassAmt(1) + RealRec.PROPVALU)
          ClassDisc(1) = OldRound(ClassDisc(1) + (RealRec.EXMPOTHR + RealRec.EXMPSENI))
          ClassCnt(1) = ClassCnt(1) + 1
          If fpcmbTownship.Text = "No Townships Saved" Then GoTo SkipIt1
          If ThatTownship = "Not Saved" Then
            TownshipAmt(TSCnt + 1, 1) = OldRound(TownshipAmt(TSCnt + 1, 1) + RealRec.PROPVALU)
            TownshipDisc(TSCnt + 1, 1) = OldRound(TownshipDisc(TSCnt + 1, 1) + (RealRec.EXMPOTHR + RealRec.EXMPSENI))
            TownshipNet(TSCnt + 1, 1) = OldRound(TownshipNet(TSCnt + 1, 1) + (RealRec.PROPVALU - (RealRec.EXMPOTHR + RealRec.EXMPSENI)))
            TownshipCount(TSCnt + 1, 1) = OldRound(TownshipCount(TSCnt + 1, 1) + 1)
          Else
            For z = 1 To TSCnt + 1
              If TownshipName(z, 1) = ThatTownship Then
                TownshipAmt(z, 1) = OldRound(TownshipAmt(z, 1) + RealRec.PROPVALU)
                TownshipDisc(z, 1) = OldRound(TownshipDisc(z, 1) + (RealRec.EXMPOTHR + RealRec.EXMPSENI))
                TownshipNet(z, 1) = OldRound(TownshipNet(z, 1) + (RealRec.PROPVALU - (RealRec.EXMPOTHR + RealRec.EXMPSENI)))
                TownshipCount(z, 1) = OldRound(TownshipCount(z, 1) + 1)
              End If
            Next z
          End If
SkipIt1:
        ElseIf ThisClassType = "COMMERCIAL" Then
          ThisClass = "COMMERCIAL"
          ClassAmt(2) = OldRound(ClassAmt(2) + RealRec.PROPVALU)
          ClassDisc(2) = OldRound(ClassDisc(2) + (RealRec.EXMPOTHR + RealRec.EXMPSENI))
          ClassCnt(2) = ClassCnt(2) + 1
          If fpcmbTownship.Text = "No Townships Saved" Then GoTo SkipIt2
          If ThatTownship = "Not Saved" Then
            TownshipAmt(TSCnt + 1, 2) = OldRound(TownshipAmt(TSCnt + 1, 2) + RealRec.PROPVALU)
            TownshipDisc(TSCnt + 1, 2) = OldRound(TownshipDisc(TSCnt + 1, 2) + (RealRec.EXMPOTHR + RealRec.EXMPSENI))
            TownshipNet(TSCnt + 1, 2) = OldRound(TownshipNet(TSCnt + 1, 2) + (RealRec.PROPVALU - (RealRec.EXMPOTHR + RealRec.EXMPSENI)))
            TownshipCount(TSCnt + 1, 2) = OldRound(TownshipCount(TSCnt + 1, 2) + 1)
          Else
            For z = 1 To TSCnt
              If TownshipName(z, 2) = ThatTownship Then
                TownshipAmt(z, 2) = OldRound(TownshipAmt(z, 2) + RealRec.PROPVALU)
                TownshipDisc(z, 2) = OldRound(TownshipDisc(z, 2) + (RealRec.EXMPOTHR + RealRec.EXMPSENI))
                TownshipNet(z, 2) = OldRound(TownshipNet(z, 2) + (RealRec.PROPVALU - (RealRec.EXMPOTHR + RealRec.EXMPSENI)))
                TownshipCount(z, 2) = OldRound(TownshipCount(z, 2) + 1)
              End If
            Next z
          End If
SkipIt2:
        ElseIf ThisClassType = "PRIVATE" Then
          ThisClass = "PRIVATE"
          ClassAmt(3) = OldRound(ClassAmt(3) + RealRec.PROPVALU)
          ClassDisc(3) = OldRound(ClassDisc(3) + (RealRec.EXMPOTHR + RealRec.EXMPSENI))
          ClassCnt(3) = ClassCnt(3) + 1
          If fpcmbTownship.Text = "No Townships Saved" Then GoTo SkipIt3
          If ThatTownship = "Not Saved" Then
            TownshipAmt(TSCnt + 1, 3) = OldRound(TownshipAmt(TSCnt + 1, 3) + RealRec.PROPVALU)
            TownshipDisc(TSCnt + 1, 3) = OldRound(TownshipDisc(TSCnt + 1, 3) + (RealRec.EXMPOTHR + RealRec.EXMPSENI))
            TownshipNet(TSCnt + 1, 3) = OldRound(TownshipNet(TSCnt + 1, 3) + (RealRec.PROPVALU - (RealRec.EXMPOTHR + RealRec.EXMPSENI)))
            TownshipCount(TSCnt + 1, 3) = OldRound(TownshipCount(TSCnt + 1, 3) + 1)
          Else
            For z = 1 To TSCnt + 1
              If TownshipName(z, 3) = ThatTownship Then
                TownshipAmt(z, 3) = OldRound(TownshipAmt(z, 3) + RealRec.PROPVALU)
                TownshipDisc(z, 3) = OldRound(TownshipDisc(z, 3) + (RealRec.EXMPOTHR + RealRec.EXMPSENI))
                TownshipNet(z, 3) = OldRound(TownshipNet(z, 3) + (RealRec.PROPVALU - (RealRec.EXMPOTHR + RealRec.EXMPSENI)))
                TownshipCount(z, 3) = OldRound(TownshipCount(z, 3) + 1)
              End If
            Next z
          End If
SkipIt3:
        ElseIf ThisClassType = "UNCLASSIFIED" Then
          ThisClass = "UNCLASSIFIED"
          ClassAmt(4) = OldRound(ClassAmt(4) + RealRec.PROPVALU)
          ClassDisc(4) = OldRound(ClassDisc(4) + (RealRec.EXMPOTHR + RealRec.EXMPSENI))
          ClassCnt(4) = ClassCnt(4) + 1
          If fpcmbTownship.Text = "No Townships Saved" Then GoTo SkipIt4
          If ThatTownship = "Not Saved" Then
            TownshipAmt(TSCnt + 1, 4) = OldRound(TownshipAmt(TSCnt + 1, 4) + RealRec.PROPVALU)
            TownshipDisc(TSCnt + 1, 4) = OldRound(TownshipDisc(TSCnt + 1, 4) + (RealRec.EXMPOTHR + RealRec.EXMPSENI))
            TownshipNet(TSCnt + 1, 4) = OldRound(TownshipNet(TSCnt + 1, 4) + (RealRec.PROPVALU - (RealRec.EXMPOTHR + RealRec.EXMPSENI)))
            TownshipCount(TSCnt + 1, 4) = OldRound(TownshipCount(TSCnt + 1, 4) + 1)
          Else
            For z = 1 To TSCnt + 1
              If TownshipName(z, 4) = ThatTownship Then
                TownshipAmt(z, 4) = OldRound(TownshipAmt(z, 4) + RealRec.PROPVALU)
                TownshipDisc(z, 4) = OldRound(TownshipDisc(z, 4) + (RealRec.EXMPOTHR + RealRec.EXMPSENI))
                TownshipNet(z, 4) = OldRound(TownshipNet(z, 4) + (RealRec.PROPVALU - (RealRec.EXMPOTHR + RealRec.EXMPSENI)))
                TownshipCount(z, 4) = OldRound(TownshipCount(z, 4) + 1)
              End If
            Next z
          End If
SkipIt4:
        Else
          For y = 1 To 6
            If QPTrim$(TaxMasterRec.ClassName(y)) = ThisClassType Then
              ThisClass = ClassType
              ClassAmt(y + 4) = OldRound(ClassAmt(y + 4) + RealRec.PROPVALU)
              ClassDisc(y + 4) = OldRound(ClassDisc(y + 4) + (RealRec.EXMPOTHR + RealRec.EXMPSENI))
              ClassCnt(y + 4) = ClassCnt(y + 4) + 1
              If fpcmbTownship.Text = "No Townships Saved" Then GoTo SkipIt5
              If ThatTownship = "Not Saved" Then
                TownshipAmt(TSCnt + 1, y + 4) = OldRound(TownshipAmt(TSCnt + 1, y + 4) + RealRec.PROPVALU)
                TownshipDisc(TSCnt + 1, y + 4) = OldRound(TownshipDisc(TSCnt + 1, y + 4) + (RealRec.EXMPOTHR + RealRec.EXMPSENI))
                TownshipNet(TSCnt + 1, y + 4) = OldRound(TownshipNet(TSCnt + 1, y + 4) + (RealRec.PROPVALU - (RealRec.EXMPOTHR + RealRec.EXMPSENI)))
                TownshipCount(TSCnt + 1, y + 4) = OldRound(TownshipCount(TSCnt + 1, y + 4) + 1)
              Else
                For z = 1 To TSCnt + 1
                  If TownshipName(z, y + 4) = ThatTownship Then
                    TownshipAmt(z, y + 4) = OldRound(TownshipAmt(z, y + 4) + RealRec.PROPVALU)
                    TownshipDisc(z, y + 4) = OldRound(TownshipDisc(z, y + 4) + (RealRec.EXMPOTHR + RealRec.EXMPSENI))
                    TownshipNet(z, y + 4) = OldRound(TownshipNet(z, y + 4) + (RealRec.PROPVALU - (RealRec.EXMPOTHR + RealRec.EXMPSENI)))
                    TownshipCount(z, y + 4) = OldRound(TownshipCount(z, y + 4) + 1)
                  End If
                Next z
              End If
              Exit For
            End If
SkipIt5:
          Next y
          If y > 6 Then
            Print #RptHandle, "UNKNOWN"
          End If
        End If
NoClass:
        GAmtTot = OldRound(GAmtTot + RealRec.PROPVALU)
        GDiscTot = OldRound(GDiscTot + RealRec.EXMPOTHR + RealRec.EXMPSENI)
        GNetTot = OldRound(GNetTot + RealRec.PROPVALU - (RealRec.EXMPOTHR + RealRec.EXMPSENI))
        GCntTot = GCntTot + 1
        If LineCnt > MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintHeader
          GoSub PrintCustHeader
        End If
        If ThisName <> ThatName Then
          GoSub PrintCustHeader
          ThatName = ThisName
        End If
          
        Print #RptHandle, Tab(2); QPTrim$(RealRec.RealPin); Tab(17); Using$("$###,###,##0.00", RealRec.PROPVALU); Tab(40); Using$("##,###,##0.00", OldRound(RealRec.EXMPOTHR + RealRec.EXMPSENI)); Tab(63); Using$("$###,###,##0.00", OldRound(RealRec.PROPVALU - (RealRec.EXMPOTHR + RealRec.EXMPSENI)))
        Print #RptHandle, Tab(2); ThisClassType; Tab(20);
        If fpcmbRptOpt.Text = "Address of Property" Then
          If QPTrim$(RealRec.PropAddr) <> "" Then
            Print #RptHandle, QPTrim$(RealRec.PropAddr); Tab(45);
          Else
            Print #RptHandle, "Not Saved"; Tab(45);
          End If
        Else
          If QPTrim$(RealRec.PROPNOT1) <> "" Then
            Print #RptHandle, QPTrim$(RealRec.PROPNOT1); Tab(45);
          Else
            Print #RptHandle, "Not Saved"; Tab(45);
          End If
        End If
        Print #RptHandle, ThatTownship
        
MoveOn:
        NextRec = RealRec.NextRec
      Loop
    End If
SkipIt:
    frmTaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      Exit Sub
    End If
  Next x
  Print #RptHandle, String(77, "-")
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  
  If PrnCnt = 0 Then
    Call TaxMsg(900, "There are no records to display using the criteria entered.")
    Close
    Exit Sub
  End If
  
  Print #RptHandle, FF$
  GoSub PrintSummaryHeader
  Print #RptHandle, "Report Totals: "; Tab(20); "Class Count"; Tab(35); "Prop Value"; Tab(51); "Discounts"; Tab(65); "Net Tax Value"
  Print #RptHandle, String(77, "-")
  Print #RptHandle, Tab(22); Using$("##,##0", GCntTot); Tab(30); Using$("$###,###,##0.00", GAmtTot); Tab(46); Using$("$##,###,##0.00", GDiscTot); Tab(63); Using$("$###,###,##0.00", GNetTot)
  Print #RptHandle,
  Print #RptHandle,
  LineCnt = LineCnt + 5
  GoSub PrintClassHeader
  
  ClassAmtTot = 0
  ClassDiscTot = 0
  ClassNetTot = 0
  ClassCntTot = 0
  For y = 1 To 10
    If ClassAmt(y) = 0 Then GoTo JumpIt2
    ClassAmtTot = OldRound(ClassAmtTot + ClassAmt(y))
    ClassDiscTot = OldRound(ClassDiscTot + ClassDisc(y))
    ClassNet = OldRound(ClassAmt(y) - ClassDisc(y))
    ClassNetTot = OldRound(ClassNetTot + ClassNet)
    ClassCntTot = OldRound(ClassCntTot + ClassCnt(y))
    If LineCnt > MaxLines - 3 Then
      Print #RptHandle, FF$
      GoSub PrintSummaryHeader
      GoSub PrintClassHeader
    End If
    If y >= 5 Then
      If QPTrim$(TaxMasterRec.ClassName(y - 4)) = "" Then GoTo JumpIt2
      Print #RptHandle, Tab(2); QPTrim$(TaxMasterRec.ClassName(y - 4)); Tab(22); Using$("##,##0", ClassCnt(y));
      Print #RptHandle, Tab(30); Using$("$###,###,##0.00", ClassAmt(y)); Tab(46); Using$("$##,###,##0.00", ClassDisc(y));
      Print #RptHandle, Tab(63); Using$("$###,###,##0.00", ClassNet)
      LineCnt = LineCnt + 1
    Else
      Select Case y
        Case 1
          PrintClass = "INDUSTRIAL"
        Case 2
          PrintClass = "COMMERCIAL"
        Case 3
          PrintClass = "PRIVATE"
        Case 4
          PrintClass = "UNCLASSIFIED"
      End Select
      If LineCnt > MaxLines - 3 Then
        Print #RptHandle, FF$
        GoSub PrintSummaryHeader
        GoSub PrintClassHeader
      End If
      Print #RptHandle, Tab(2); PrintClass; Tab(22); Using$("##,##0", ClassCnt(y));
      Print #RptHandle, Tab(30); Using$("$###,###,##0.00", ClassAmt(y)); Tab(46); Using$("$##,###,##0.00", ClassDisc(y));
      Print #RptHandle, Tab(63); Using$("$###,###,##0.00", ClassNet)
      LineCnt = LineCnt + 1
    End If
JumpIt2:
  Next y
  If LineCnt > MaxLines - 3 Then
    Print #RptHandle, FF$
    GoSub PrintSummaryHeader
    GoSub PrintClassHeader
  End If
  Print #RptHandle, String(77, "-")
  Print #RptHandle, "Class Totals: "; Tab(22); Using$("##,##0", ClassCntTot); Tab(30); Using$("$###,###,##0.00", ClassAmtTot); Tab(46);
  Print #RptHandle, Using$("$##,###,##0.00", ClassDiscTot); Tab(63); Using$("$###,###,##0.00", ClassNetTot)
  
  If fpcmbTownship.Text = "No Townships Saved" Then GoTo NoTownships
  Print #RptHandle,
  Print #RptHandle,
  LineCnt = LineCnt + 2
  GoSub PrintTownshipHeader
  
  GTSAmtTot = 0
  GTSDiscTot = 0
  GTSNetTot = 0
  GTSCntTot = 0
  
  For x = 1 To TSCnt + 1
    TSAmtTot = 0
    TSDiscTot = 0
    TSNetTot = 0
    TSCntTot = 0
    For y = 1 To 10
      If TownshipAmt(x, y) > 0 Then
        Print #RptHandle, Tab(2); QPTrim$(TownshipName(x, y))
        LineCnt = LineCnt + 1
        Exit For
      End If
    Next y
    For y = 1 To 10
      If TownshipAmt(x, y) = 0 Then GoTo JumpIt
      TSAmtTot = OldRound(TSAmtTot + TownshipAmt(x, y))
      TSDiscTot = OldRound(TSDiscTot + TownshipDisc(x, y))
      TSNetTot = OldRound(TSNetTot + TownshipNet(x, y))
      TSCntTot = OldRound(TSCntTot + TownshipCount(x, y))
      GTSAmtTot = OldRound(GTSAmtTot + TownshipAmt(x, y))
      GTSDiscTot = OldRound(GTSDiscTot + TownshipDisc(x, y))
      GTSNetTot = OldRound(GTSNetTot + TownshipNet(x, y))
      GTSCntTot = OldRound(GTSCntTot + TownshipCount(x, y))
      If LineCnt > MaxLines + 3 Then
        Print #RptHandle, FF$
        GoSub PrintSummaryHeader
        GoSub PrintTownshipHeader
      End If
      If y >= 5 Then
        If QPTrim$(TaxMasterRec.ClassName(y - 4)) = "" Then GoTo JumpIt
        PrintClass = QPTrim$(TaxMasterRec.ClassName(y - 4))
        Print #RptHandle, Tab(6); PrintClass; Tab(22); Using$("##,##0", TownshipCount(x, y)); Tab(30); Using$("$###,###,##0.00", TownshipAmt(x, y)); Tab(46); Using$("$##,###,##0.00", TownshipDisc(x, y)); Tab(63); Using$("$###,###,##0.00", TownshipNet(x, y))
        LineCnt = LineCnt + 1
        GoTo JumpIt
      End If
      Select Case y
        Case 1
          PrintClass = "INDUSTRIAL"
        Case 2
          PrintClass = "COMMERCIAL"
        Case 3
          PrintClass = "PRIVATE"
        Case 4
          PrintClass = "UNCLASSIFIED"
      End Select
      If LineCnt > MaxLines - 3 Then
        Print #RptHandle, FF$
        GoSub PrintSummaryHeader
        GoSub PrintTownshipHeader
      End If
      Print #RptHandle, Tab(6); PrintClass; Tab(22); Using$("##,##0", TownshipCount(x, y)); Tab(30); Using$("$###,###,##0.00", TownshipAmt(x, y)); Tab(46); Using$("$##,###,##0.00", TownshipDisc(x, y)); Tab(63); Using$("$###,###,##0.00", TownshipNet(x, y))
      LineCnt = LineCnt + 1
JumpIt:
    Next y
    If LineCnt > MaxLines - 4 Then
      Print #RptHandle, FF$
      GoSub PrintSummaryHeader
      GoSub PrintTownshipHeader
    End If
    If TSAmtTot > 0 Then
      Print #RptHandle, String(77, "-")
      Print #RptHandle, "Township Totals:"; Tab(22); Using$("##,##0", TSCntTot); Tab(30); Using$("$###,###,##0.00", TSAmtTot); Tab(46); Using$("$##,###,##0.00", TSDiscTot); Tab(63); Using$("$###,###,##0.00", TSNetTot)
      Print #RptHandle,
      LineCnt = LineCnt + 3
    End If
  Next x
  If LineCnt > MaxLines - 4 Then
    Print #RptHandle, FF$
    GoSub PrintSummaryHeader
  End If
  Print #RptHandle, String(77, "-")
  Print #RptHandle, "Grand T'Ship Ttls:"; Tab(22); Using$("##,##0", GTSCntTot); Tab(30); Using$("$###,###,##0.00", GTSAmtTot); Tab(46); Using$("$##,###,##0.00", GTSDiscTot); Tab(63); Using$("$###,###,##0.00", GTSNetTot)
  Print #RptHandle,
  LineCnt = LineCnt + 3
  
NoTownships:
  Print #RptHandle, FF$
  Close
  
  ViewPrint RptFile, "Gathering Classification Data", True
  Exit Sub
  
PrintHeader:
  If RptOptFlag = 1 Then
    Description = "Address of Property"
  Else
    Description = "First Line of Notes"
  End If
  Page = Page + 1
  Print #RptHandle, Tab(20); "Tax Real Property Classification Report"
  Print #RptHandle, TownName; Tab(65); "Page #: " + CStr(Page)
  If InactiveFlag = True Then
    Print #RptHandle, "Report Date: " + CStr(Date); Tab(62); "**I** = Inactive"
  Else
    Print #RptHandle, "Report Date: " + CStr(Date)
  End If
  Print #RptHandle, "Township:"; Tab(11); ThisTownship;
  
  ThisLen = Len(ClassType)
  ThisTab = 16 + ThisLen
  ThisTab = 78 - ThisTab
  Print #RptHandle, Tab(ThisTab); "Classification: " + ClassType
  
  Print #RptHandle,
  Print #RptHandle, Tab(2); "Cust Acct#"; Tab(18); "Customer Name"
  Print #RptHandle, Tab(2); "Property Pin"; Tab(18); "Property Value"; Tab(35); "Property Discounts"; Tab(65); "Net Tax Value"
  Print #RptHandle, Tab(2); "Property Class"; Tab(20); Description; Tab(45); "Property Township"
  Print #RptHandle, String(77, "-")
  LineCnt = 9
  
  Return
  
PrintCustHeader:
  If LineCnt >= MaxLines - 4 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  If LineCnt <> 9 Then
    Print #RptHandle, String(77, "-")
  End If
  If InactiveFlag = True Then
    If TaxCust.Active = "N" Then
      Print #RptHandle, Tab(2); Using$("########0", TaxCust.Acct); Tab(18); QPTrim$(TaxCust.CustName) + " **I**"
    Else
      Print #RptHandle, Tab(2); Using$("########0", TaxCust.Acct); Tab(18); QPTrim$(TaxCust.CustName)
    End If
  Else
    Print #RptHandle, Tab(2); Using$("########0", TaxCust.Acct); Tab(18); QPTrim$(TaxCust.CustName)
  End If
  If UseOpt = "Y" Then
    Print #RptHandle, Tab(8); ThisOpt + ": " + QPTrim$(TaxCust.OptSrchDesc)
    LineCnt = LineCnt + 1
  End If
  Print #RptHandle, String(77, ".")
  LineCnt = LineCnt + 1
  
  Return
  
PrintSummaryHeader:
  Page = Page + 1
  Print #RptHandle, Tab(15); "Tax Real Property Classification Report Summary"
  Print #RptHandle, TownName; Tab(65); "Page #: " + CStr(Page)
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle, "Township:"; Tab(11); ThisTownship;
  ThisLen = Len(ClassType)
  ThisTab = 16 + ThisLen
  ThisTab = 78 - ThisTab
  Print #RptHandle, Tab(ThisTab); "Classification: " + ClassType
  Print #RptHandle, String(77, "-")
  LineCnt = 7
  
  Return

PrintTownshipHeader:
  Print #RptHandle, Tab(2); "Township Summary"
  Print #RptHandle, Tab(6); "Class"; Tab(20); "Class Count"; Tab(35); "Prop Value"; Tab(51); "Discounts"; Tab(65); "Net Tax Value"
  Print #RptHandle, String(77, "-")
  LineCnt = LineCnt + 3
  
  Return
  
PrintClassHeader:
  Print #RptHandle, Tab(2); "Classification Summary";
  Print #RptHandle, Tab(20); "Class Count"; Tab(35); "Prop Value"; Tab(51); "Discounts"; Tab(65); "Net Tax Value"
  Print #RptHandle, String(77, "-")
  LineCnt = LineCnt + 2

  Return
  
ERRORSTUFF:
   Unload frmTaxShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxRealClassRpt", "PrintText", Erl)
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

