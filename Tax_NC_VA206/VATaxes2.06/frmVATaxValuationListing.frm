VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmVATaxValuationListing 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Valuation Listing"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxValuationListing.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   7230
      Left            =   1920
      TabIndex        =   7
      Top             =   750
      Width           =   7785
      _Version        =   196609
      _ExtentX        =   13732
      _ExtentY        =   12753
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmVATaxValuationListing.frx":08CA
      Begin LpLib.fpCombo fpcmbPrintOpt 
         Height          =   384
         Left            =   3048
         TabIndex        =   5
         Top             =   4692
         Width           =   3564
         _Version        =   196608
         _ExtentX        =   6286
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
         ColDesigner     =   "frmVATaxValuationListing.frx":08E6
      End
      Begin LpLib.fpCombo fpcmbDetSum 
         Height          =   384
         Left            =   3528
         TabIndex        =   4
         Top             =   4116
         Width           =   2844
         _Version        =   196608
         _ExtentX        =   5016
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
         ColDesigner     =   "frmVATaxValuationListing.frx":0BDD
      End
      Begin LpLib.fpCombo fpcmbPrintOrder 
         Height          =   384
         Left            =   3048
         TabIndex        =   6
         Top             =   5280
         Width           =   3564
         _Version        =   196608
         _ExtentX        =   6286
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
         ColDesigner     =   "frmVATaxValuationListing.frx":0ED4
      End
      Begin LpLib.fpCombo fpcmbIncInactive 
         Height          =   384
         Left            =   3540
         TabIndex        =   3
         Top             =   3528
         Width           =   2820
         _Version        =   196608
         _ExtentX        =   4974
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
         ColDesigner     =   "frmVATaxValuationListing.frx":11CB
      End
      Begin LpLib.fpCombo fpcmbPropType 
         Height          =   384
         Left            =   3540
         TabIndex        =   2
         Top             =   2952
         Width           =   2820
         _Version        =   196608
         _ExtentX        =   4974
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
         ColDesigner     =   "frmVATaxValuationListing.frx":14C2
      End
      Begin LpLib.fpCombo fpcmbTownship 
         Height          =   384
         Left            =   3540
         TabIndex        =   1
         Top             =   2364
         Width           =   2820
         _Version        =   196608
         _ExtentX        =   4974
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
         ColDesigner     =   "frmVATaxValuationListing.frx":17B9
      End
      Begin EditLib.fpText fptxtAddress 
         Height          =   405
         Left            =   3480
         TabIndex        =   0
         Top             =   1800
         Width           =   2820
         _Version        =   196608
         _ExtentX        =   4974
         _ExtentY        =   714
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
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   636
         Left            =   2040
         TabIndex        =   16
         TabStop         =   0   'False
         Tag             =   "Press the 'Cancel' button to exit this screen and return to the main 'Business License Reports' menu."
         Top             =   6240
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
         ButtonDesigner  =   "frmVATaxValuationListing.frx":1AB0
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   636
         Left            =   4272
         TabIndex        =   17
         TabStop         =   0   'False
         Tag             =   $"frmVATaxValuationListing.frx":1C8E
         Top             =   6240
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
         ButtonDesigner  =   "frmVATaxValuationListing.frx":1D39
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Real Address:"
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
         Left            =   1680
         TabIndex        =   15
         Top             =   1890
         Width           =   1545
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Township:"
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
         Left            =   1920
         TabIndex        =   14
         Top             =   2490
         Width           =   1305
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Property Type:"
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
         Left            =   1320
         TabIndex        =   13
         Top             =   3090
         Width           =   1905
      End
      Begin VB.Label Label1 
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
         Height          =   345
         Left            =   1275
         TabIndex        =   12
         Top             =   3630
         Width           =   1950
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
         Left            =   1560
         TabIndex        =   11
         Top             =   5370
         Width           =   1305
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Detail/Summary:"
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
         Left            =   1320
         TabIndex        =   10
         Top             =   4200
         Width           =   1905
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   4500
         Left            =   1005
         Top             =   1485
         Width           =   5970
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Valuation Listing"
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
         Height          =   390
         Left            =   1560
         TabIndex        =   9
         Top             =   570
         Width           =   4935
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   1410
         Top             =   435
         Width           =   5265
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
         Left            =   1395
         TabIndex        =   8
         Top             =   4800
         Width           =   1500
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   7500
      Left            =   1800
      Top             =   615
      Width           =   8055
   End
End
Attribute VB_Name = "frmVATaxValuationListing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Dim Town$
  Dim UseOpt As String * 1
  Dim ThisOpt$
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  frmVATaxReportsMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdProcess_Click()
  If fpcmbDetSum.Text = "Summary" Then
    If fpcmbPrintOpt.Text = "Graphical" Then
      Call PrintGraphics
    Else
      frmVATaxMsg.Label1.Caption = "Pitch 17 is recommended for this printout."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show vbModal
      Call PrintText
    End If
  Else
    If fpcmbPrintOpt.Text = "Graphical" Then
      Call PrintGraphicsDet
    Else
      frmVATaxMsg.Label1.Caption = "Pitch 17 is recommended for this printout."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show vbModal
      Call PrintTextDet
    End If
  End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%E"
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
  Me.HelpContextID = hlpMasterValuation
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxValuationListing.")
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

Private Sub LoadMe()
  Dim TSRec As TownshipType
  Dim TSCnt As Integer
  Dim TSHandle As Integer
  Dim x As Long
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  
  On Error GoTo ERRORSTUFF
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  Town$ = QPTrim$(TaxMasterRec.Name)
  fpcmbTownship.Text = "All"
  fpcmbTownship.AddItem "All"
  If Exist(TaxTownships) Then
    OpenTownshipFile TSHandle, TSCnt
    For x = 1 To TSCnt
      Get TSHandle, x, TSRec
      fpcmbTownship.AddItem QPTrim$(TSRec.TownShip)
    Next x
    Close TSHandle
  End If
  fpcmbPropType.Text = "Both"
  fpcmbPropType.AddItem "Both"
  fpcmbPropType.AddItem "Real Only"
  fpcmbPropType.AddItem "Personal Only"
  
  fpcmbPrintOpt.Text = "Graphical"
  fpcmbPrintOpt.AddItem "Graphical"
  fpcmbPrintOpt.AddItem "Text"
  
  fpcmbPrintOrder.Text = "Name Order"
  fpcmbPrintOrder.AddItem "Name Order"
  fpcmbPrintOrder.AddItem "Acct Number Order"
  fpcmbPrintOrder.AddItem "Search Name"
  ThisOpt = QPTrim$(TaxMasterRec.OptSrchCust)
  If ThisOpt <> "" Then
    fpcmbPrintOrder.AddItem ThisOpt + " Order"
  End If
  
  fpcmbDetSum.Text = "Summary"
  fpcmbDetSum.AddItem "Detail"
  fpcmbDetSum.AddItem "Summary"
  
  fpcmbIncInactive.Text = "Both"
  fpcmbIncInactive.AddItem "Both"
  fpcmbIncInactive.AddItem "Active"
  fpcmbIncInactive.AddItem "Inactive"
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxValuationListing", "LoadMe", Erl)
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

Private Sub PrintGraphicsDet()
  Dim x As Long, y As Integer
  Dim dlm$
  Dim PropTypeFlag As Boolean
  Dim InactiveFlag As Boolean
  Dim ThisTownship$
  Dim ThisAdd$
  Dim RptFile$
  Dim RptHandle As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim WhatPers&
  Dim WhatReal&
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim ThisTot As Double
  Dim ThisLtr As String * 1
  Dim NewLtr As String * 1
  Dim NameFlag As Boolean
  Dim AlphaCnt As Integer
  Dim GTotCnt As Long
  Dim NewAlphaFlag As Boolean
  Dim SearchFlag As Boolean
  Dim GRealVal As Double
  Dim AlphaRealVal As Double
  Dim GBldgVal As Double
  Dim AlphaBldgVal As Double
  Dim GPersVal As Double
  Dim GTPersVal As Double
  Dim GTMTVal As Double
  Dim GTMCVal As Double
  Dim GTMHVal As Double
  Dim GTFarmVal As Double
  Dim AlphaPersVal As Double
  Dim GDscntVal As Double
  Dim AlphaDscntVal As Double
  Dim AlphaTotVal As Double
  Dim GTotVal As Double
  Dim ThisRealVal As Double
  Dim ThisBldgVal As Double
  Dim ThisDscntVal As Double
  Dim ThisPersVal As Double
  Dim ThisNet As Double
  Dim InactiveYN As String
  Dim AlphaNet As Double
  Dim GNet As Double
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim CustCnt As Long
  Dim AlphaCustCnt As Integer
  Dim ThisName$
  Dim CustRealTot As Double
  Dim CustBldgTot As Double
  Dim CustPersTot As Double
  Dim CustDscntTot As Double
  Dim CustNetTot As Double
  Dim ThisCnt As Integer
  Dim ThisCustCnt As Integer
  Dim PrintThis As Boolean
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim OptFlag As Boolean
  Dim ActiveFlag As String * 1
  Dim PropType$
  Dim PrintOrder$
  
  On Error GoTo ERRORSTUFF
  
  If fpcmbPrintOrder.Text = "Acct Number Order" Then
    PrintOrder = "N"
  Else
    PrintOrder = "Y"
  End If
  PropType$ = fpcmbPropType.Text
  PrintThis = False
  OptFlag = False
  dlm = "~"
  NewAlphaFlag = False
  SearchFlag = False
  If InStr(fpcmbPrintOrder.Text, "Name") Then
    NameFlag = True
  Else
    NameFlag = False
  End If

  ThisLtr = ""
  NewLtr = ""
  ThisAdd$ = QPTrim$(fptxtAddress.Text)
  ThisTownship = QPTrim$(fpcmbTownship.Text)
  ActiveFlag = "B"
  If QPTrim$(fpcmbIncInactive.Text) = "Active" Then
    ActiveFlag = "A"
  ElseIf QPTrim$(fpcmbIncInactive.Text) = "Inactive" Then
    ActiveFlag = "I"
  End If

  If QPTrim$(fpcmbPrintOrder.Text) = "Name Order" Then
    OpenNameIdxFile IdxHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmVATaxMsg.Label1.Caption = "There are no customers saved."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show
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
      frmVATaxMsg.Label1.Caption = "There are no search names indexed."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show
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
    SearchFlag = True
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

  RptFile$ = "TAXRPTS\TXVALLST.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle

  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenRealPropFile RHandle, NumOfRealRecs
  OpenPersPropFile PHandle, NumOfPersRecs

  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  End If

  frmVATaxShowPctComp.Label1 = "Gathering Valuation Data"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
    Else
      Get TCHandle, x, TaxCust
    End If
    If TaxCust.Deleted <> 0 Then GoTo SkipIt
    If ActiveFlag = "A" And TaxCust.Active = "N" Then
      GoTo SkipIt
    ElseIf ActiveFlag = "I" And TaxCust.Active = "Y" Then
      GoTo SkipIt
    End If
    If IdxFlag = True Then
      If SearchFlag = True Then
        ThisLtr = Mid(QPTrim$(TaxCust.SName), 1, 1)
      ElseIf OptFlag = True Then
        ThisLtr = Mid(QPTrim$(TaxCust.OptSrchDesc), 1, 1)
      Else
        ThisLtr = Mid(QPTrim$(TaxCust.CustName), 1, 1)
      End If
    End If
      WhatPers = TaxCust.FirstPropRec
      If GTotCnt > 0 And (WhatPers > 0 Or TaxCust.FirstPersRec > 0) Then
        If NewLtr <> ThisLtr Then
          NewLtr = ThisLtr
          NewAlphaFlag = True
          AlphaRealVal = 0
          AlphaBldgVal = 0
          AlphaPersVal = 0
          AlphaDscntVal = 0
          AlphaTotVal = 0
          AlphaNet = 0
          AlphaCnt = 0
          AlphaCustCnt = 0
        End If
      Else
        If WhatPers > 0 Then
          NewLtr = ThisLtr
        End If
      End If
'    End If
    ThisCustCnt = 0
    ReDim CustRealVal(1 To 1) As Double
    ReDim CustBldgVal(1 To 1) As Double
    ReDim CustDscntVal(1 To 1) As Double
    ReDim CustPersVal(1 To 1) As Double
    ReDim CustNet(1 To 1) As Double
    ReDim ThisPin(1 To 1) As String
    CustRealTot = 0
    CustBldgTot = 0
    CustDscntTot = 0
    CustPersTot = 0
    CustNetTot = 0
    ThisCustCnt = 0
    
    If QPTrim$(fpcmbPropType.Text) = "Personal Only" Then
      If QPTrim$(ThisTownship) = "All" And QPTrim$(fptxtAddress.Text) = "" Then
        GoTo PersOnly
      Else
        GoTo SkipIt
      End If
    End If

    Do While WhatPers > 0
      Get RHandle, WhatPers&, RealRec
      If RealRec.Deleted = True Then GoTo NextRec
      If ThisTownship <> "All" Then
        If QPTrim$(UCase(RealRec.TownShip)) <> ThisTownship Then GoTo NextRec
      End If
      If ThisAdd$ <> "" Then
        If InStr(RealRec.PropAddr, ThisAdd) = 0 Then
          GoTo NextRec
        End If
      End If
      GTotCnt = GTotCnt + 1
      ThisRealVal = RealRec.PROPVALU
      ThisBldgVal = RealRec.BldgVal
'      ThisDscntVal = OldRound(RealRec.EXMPSENI + RealRec.EXMPOTHR)
      ThisDscntVal = OldRound(RealRec.EXMPOTHR) '6/14/06
      ThisPersVal = 0
      ThisNet = OldRound(ThisRealVal + RealRec.BldgVal - ThisDscntVal)
      ThisCustCnt = ThisCustCnt + 1
      ReDim Preserve CustRealVal(1 To ThisCustCnt) As Double
      CustRealVal(ThisCustCnt) = OldRound(CustRealVal(ThisCustCnt) + RealRec.PROPVALU)
      CustRealTot = OldRound(CustRealTot + ThisRealVal)

      ReDim Preserve CustBldgVal(1 To ThisCustCnt) As Double
      CustBldgVal(ThisCustCnt) = OldRound(CustBldgVal(ThisCustCnt) + RealRec.BldgVal)
      CustBldgTot = OldRound(CustBldgTot + ThisBldgVal)

      ReDim Preserve CustDscntVal(1 To ThisCustCnt) As Double
      CustDscntVal(ThisCustCnt) = OldRound(CustDscntVal(ThisCustCnt) + ThisDscntVal)
      CustDscntTot = OldRound(CustDscntTot + ThisDscntVal)

      ReDim Preserve CustPersVal(1 To ThisCustCnt) As Double
      CustPersVal(ThisCustCnt) = OldRound(CustPersVal(ThisCustCnt) + 0)
      CustPersTot = OldRound(CustPersTot + ThisPersVal)

      ReDim Preserve CustNet(1 To ThisCustCnt) As Double
      CustNet(ThisCustCnt) = OldRound(CustNet(ThisCustCnt) + ThisNet)
      CustNetTot = OldRound(CustNetTot + ThisNet)
      
      ReDim Preserve ThisPin(1 To ThisCustCnt) As String
      ThisPin(ThisCustCnt) = "Real Pin: " + QPTrim$(RealRec.RealPin)
      GRealVal = OldRound(GRealVal + RealRec.PROPVALU)
      GBldgVal = OldRound(GBldgVal + RealRec.BldgVal)
      GPersVal = GPersVal + 0
'      GDscntVal = OldRound(GDscntVal + RealRec.EXMPSENI + RealRec.EXMPOTHR)
      GDscntVal = OldRound(GDscntVal + RealRec.EXMPOTHR) '6/14/06
'      GTotVal = OldRound(GTotVal + RealRec.PROPVALU + RealRec.EXMPSENI + RealRec.EXMPOTHR)
      GTotVal = OldRound(GTotVal + RealRec.PROPVALU + RealRec.EXMPOTHR) '6/14/06
      AlphaCnt = AlphaCnt + 1
      AlphaRealVal = OldRound(AlphaRealVal + RealRec.PROPVALU)
      AlphaBldgVal = OldRound(AlphaBldgVal + RealRec.BldgVal)
      AlphaPersVal = AlphaPersVal + 0
'      AlphaDscntVal = OldRound(AlphaDscntVal + RealRec.EXMPSENI + RealRec.EXMPOTHR)
      AlphaDscntVal = OldRound(AlphaDscntVal + RealRec.EXMPOTHR) '6/14/06
'      AlphaTotVal = OldRound(AlphaTotVal + RealRec.PROPVALU + RealRec.BldgVal + RealRec.EXMPSENI + RealRec.EXMPOTHR)
      AlphaTotVal = OldRound(AlphaTotVal + RealRec.PROPVALU + RealRec.BldgVal + RealRec.EXMPOTHR) '6/14/06
      AlphaNet = OldRound(AlphaNet + ThisNet)
      GNet = OldRound(GNet + ThisNet)
      PrintThis = True
NextRec:
      WhatPers& = RealRec.NextRec
    Loop

    If QPTrim$(fpcmbPropType.Text) = "Real Only" Then GoTo UseAdd 'SkipIt
    If ThisAdd$ <> "" Then GoTo UseAdd

PersOnly:
    If QPTrim$(fpcmbPropType.Text) = "Both" Then
      If QPTrim$(ThisTownship) <> "All" Or QPTrim$(fptxtAddress.Text) <> "" Then
        GoTo UseAdd 'SkipIt
      End If
    End If
    
    WhatPers = TaxCust.FirstPersRec
    Do While WhatPers > 0
      Get PHandle, WhatPers, PersRec
      GTotCnt = GTotCnt + 1
      ThisRealVal = 0
      ThisBldgVal = 0
      ThisDscntVal = 0 ' 6/14/06 no more pers exemptions OldRound(PersRec.EXMPSENI + PersRec.EXMPOTHR)
      ThisPersVal = OldRound(PersRec.PersVal + PersRec.MCValue + PersRec.CVALUE + PersRec.MHValue + PersRec.MTValue)
      ThisNet = OldRound(ThisPersVal - ThisDscntVal)
      ThisCustCnt = ThisCustCnt + 1
      
      ReDim Preserve CustRealVal(1 To ThisCustCnt) As Double
      CustRealVal(ThisCustCnt) = OldRound(CustRealVal(ThisCustCnt) + 0)
      CustRealTot = OldRound(CustRealTot + ThisRealVal)
      
      ReDim Preserve CustBldgVal(1 To ThisCustCnt) As Double
      CustBldgVal(ThisCustCnt) = OldRound(CustBldgVal(ThisCustCnt) + 0)
      CustBldgTot = OldRound(CustBldgTot + ThisBldgVal)
      
      ReDim Preserve CustDscntVal(1 To ThisCustCnt) As Double
      CustDscntVal(ThisCustCnt) = 0 '6/14/06 no more pers exemptions OldRound(CustDscntVal(ThisCustCnt) + PersRec.EXMPSENI + PersRec.EXMPOTHR)
      CustDscntTot = OldRound(CustDscntTot + ThisDscntVal)
      
      ReDim Preserve CustPersVal(1 To ThisCustCnt) As Double
      CustPersVal(ThisCustCnt) = OldRound(CustPersVal(ThisCustCnt) + ThisPersVal)
      CustPersTot = OldRound(CustPersTot + ThisPersVal)
      
      ReDim Preserve CustNet(1 To ThisCustCnt) As Double
      CustNet(ThisCustCnt) = OldRound(CustNet(ThisCustCnt) + ThisPersVal + ThisRealVal - ThisDscntVal)
      CustNetTot = OldRound(CustNetTot + ThisNet)
 
      ReDim Preserve ThisPin(1 To ThisCustCnt) As String
      ThisPin(ThisCustCnt) = "Pers Pin: " + QPTrim$(PersRec.PropPin)
      GRealVal = GRealVal + 0
      GBldgVal = GBldgVal + 0
      GPersVal = OldRound(GPersVal + ThisPersVal)
      GTPersVal = OldRound(GTPersVal + PersRec.PersVal) 'added 12/15/06
      GTMTVal = OldRound(GTMTVal + PersRec.MTValue) 'added 12/15/06
      GTMCVal = OldRound(GTMCVal + PersRec.MCValue) 'added 12/15/06
      GTMHVal = OldRound(GTMHVal + PersRec.MHValue) 'added 12/15/06
      GTFarmVal = OldRound(GTFarmVal + PersRec.CVALUE) 'added 12/15/06
'      GDscntVal = 0 ' 6/14/06 no more personal exemptions OldRound(GDscntVal + PersRec.EXMPOTHR + PersRec.EXMPSENI)
      GTotVal = OldRound(GTotVal + PersRec.PersVal) '6/14/06 no more pers exemptions  + PersRec.EXMPSENI + PersRec.EXMPOTHR)
      AlphaCnt = AlphaCnt + 1
      AlphaRealVal = AlphaRealVal + 0
      AlphaBldgVal = AlphaBldgVal + 0
      AlphaPersVal = OldRound(AlphaPersVal + ThisPersVal)
'      AlphaDscntVal = 0 ' 6/14/06 no more personal exemptions OldRound(AlphaDscntVal + PersRec.EXMPSENI + PersRec.EXMPOTHR)
      AlphaTotVal = OldRound(AlphaTotVal + PersRec.PersVal) ' 6/14/06 no more personal exemptions   + PersRec.EXMPSENI + PersRec.EXMPOTHR)
      AlphaNet = OldRound(AlphaNet + ThisNet)
      GNet = OldRound(GNet + ThisNet)
      PrintThis = True
'    End If
UseAdd:
      WhatPers = PersRec.NextRec
    Loop
    If PrintThis = True Then
      GoSub PrintIt
      PrintThis = False
    End If
SkipIt:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  
  Close
  If GTotCnt = 0 Then
    Call TaxMsg(900, "There are no property valuations for the parameters entered.")
    Close
    Exit Sub
  End If
  arVATaxValDet.Show
  frmVATaxLoadReport.Show
  DoEvents

  Exit Sub

PrintIt:
  CustCnt = CustCnt + 1
  AlphaCustCnt = AlphaCustCnt + 1
  For y = 1 To ThisCustCnt
    ThisName = QPTrim$(TaxCust.CustName)
    '                   0              1                 2
    Print #RptHandle, Town$; dlm; TaxCust.Acct; dlm; ThisName; dlm;
    '                    3                4                5                 6
    Print #RptHandle, ThisLtr; dlm; CustRealTot; dlm; CustPersTot; dlm; CustDscntTot; dlm;
    '                    7              8               9                  10
    Print #RptHandle, CustNetTot; dlm; AlphaCnt; dlm; AlphaRealVal; dlm; AlphaPersVal; dlm;
    '                      11                12             13            14
    Print #RptHandle, AlphaDscntVal; dlm; AlphaNet; dlm; GTotCnt; dlm; GRealVal; dlm;
    '                    15             16            17           18           19
    Print #RptHandle, GPersVal; dlm; GDscntVal; dlm; GNet; dlm; ThisPin(y); dlm; CustCnt; dlm;
    '                     20                  21                  22                      23                 24
    Print #RptHandle, AlphaCustCnt; dlm; CustRealVal(y); dlm; CustPersVal(y); dlm; CustDscntVal(y); dlm; CustNet(y); dlm;
    '                    25              26
    Print #RptHandle, ThisAdd; dlm; ActiveFlag; dlm;
    If UseOpt = "Y" Then
      '                    27                    28                         29             30                 31              32               33                34
      Print #RptHandle, ThisOpt; dlm; QPTrim$(TaxCust.OptSrchDesc); dlm; PropType; dlm; PrintOrder; dlm; CustBldgTot; dlm; GBldgVal; dlm; AlphaBldgVal; dlm; CustBldgVal(y); dlm;
    Else
      '                 27       28          29              30               31               32              33                  34
      Print #RptHandle, ""; dlm; ""; dlm; PropType; dlm; PrintOrder; dlm; CustBldgTot; dlm; GBldgVal; dlm; AlphaBldgVal; dlm; CustBldgVal(y); dlm;
    End If
    '                    35              36            37            38             39
    Print #RptHandle, GTPersVal; dlm; GTMTVal; dlm; GTMCVal; dlm; GTMHVal; dlm; GTFarmVal
  Next y
  
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxValuationListing", "PrintGraphicsDet", Erl)
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

Private Sub PrintGraphics()
  Dim x As Long, y As Integer
  Dim dlm$
  Dim PropTypeFlag As Boolean
  Dim InactiveFlag As Boolean
  Dim ThisTownship$
  Dim ThisAdd$
  Dim RptFile$
  Dim RptHandle As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim WhatPers&
  Dim WhatReal&
  Dim InactiveYN As String
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim ThisTot As Double
  Dim ThisLtr As String * 1
  Dim NewLtr As String * 1
  Dim NameFlag As Boolean
  Dim AlphaCnt As Integer
  Dim GTotCnt As Long
  Dim NewAlphaFlag As Boolean
  Dim SearchFlag As Boolean
  Dim GRealVal As Double
  Dim AlphaRealVal As Double
  Dim GBldgVal As Double
  Dim AlphaBldgVal As Double
  Dim GPersVal As Double
  Dim AlphaPersVal As Double
  Dim GDscntVal As Double
  Dim AlphaDscntVal As Double
  Dim AlphaTotVal As Double
  Dim GTotVal As Double
  Dim ThisRealVal As Double
  Dim ThisBldgVal As Double
  Dim ThisDscntVal As Double
  Dim ThisPersVal As Double
  Dim GTPersVal As Double
  Dim GTMTVal As Double
  Dim GTMCVal As Double
  Dim GTMHVal As Double
  Dim GTFarmVal As Double
  Dim ThisNet As Double
  Dim ThisPin As String
  Dim AlphaNet As Double
  Dim GNet As Double
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim CustCnt As Long
  Dim AlphaCustCnt As Integer
  Dim CustRealVal As Double
  Dim CustBldgVal As Double
  Dim CustPersVal As Double
  Dim CustDscntVal As Double
  Dim CustNet As Double
  Dim PrintThis As Boolean
  Dim ThisName$
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim OptFlag As Boolean
  Dim ActiveFlag$
  Dim PropType$
  Dim PrintOrder$

'  Dim AHandle As Integer
  
'  On Error GoTo ERRORSTUFF
'  AHandle = FreeFile
'  Open "value.dat" For Output As AHandle
  
  If fpcmbPrintOrder.Text = "Acct Number Order" Then
    PrintOrder = "N"
  Else
    PrintOrder = "Y"
  End If
  
  PropType = fpcmbPropType.Text
  
  PrintThis = False
  dlm = "~"
  NewAlphaFlag = False
  SearchFlag = False
  OptFlag = False
  
  If InStr(fpcmbPrintOrder.Text, "Name") Then
    NameFlag = True
  Else
    NameFlag = False
  End If
  
  ThisLtr = ""
  NewLtr = ""
  ThisAdd$ = QPTrim$(fptxtAddress.Text)
  ThisTownship = QPTrim$(fpcmbTownship.Text)
  ActiveFlag = "B"
  If QPTrim$(fpcmbIncInactive.Text) = "Active" Then
    ActiveFlag = "A"
  ElseIf QPTrim$(fpcmbIncInactive.Text) = "Inactive" Then
    ActiveFlag = "I"
  End If
  
  If QPTrim$(fpcmbPrintOrder.Text) = "Name Order" Then
    OpenNameIdxFile IdxHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmVATaxMsg.Label1.Caption = "There are no customers saved."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show
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
      frmVATaxMsg.Label1.Caption = "There are no search names indexed."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show
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
    SearchFlag = True
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
  
  RptFile$ = "TAXRPTS\TXVALLSTSUM.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenRealPropFile RHandle, NumOfRealRecs
  OpenPersPropFile PHandle, NumOfPersRecs
  
  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  End If
  
  frmVATaxShowPctComp.Label1 = "Gathering Valuation Data"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
    Else
      Get TCHandle, x, TaxCust
    End If
    If TaxCust.Deleted <> 0 Then GoTo SkipIt
    If ActiveFlag = "A" And TaxCust.Active = "N" Then
      GoTo SkipIt
    ElseIf ActiveFlag = "I" And TaxCust.Active = "Y" Then
      GoTo SkipIt
    End If
    If IdxFlag = True Then
      If SearchFlag = True Then
        ThisLtr = Mid(QPTrim$(TaxCust.SName), 1, 1)
      ElseIf OptFlag = True Then
        ThisLtr = Mid(QPTrim$(TaxCust.OptSrchDesc), 1, 1)
      Else
        ThisLtr = Mid(QPTrim$(TaxCust.CustName), 1, 1)
      End If
    End If
      WhatPers = TaxCust.FirstPropRec
      If GTotCnt > 0 And (WhatPers > 0 Or TaxCust.FirstPersRec > 0) Then
        If NewLtr <> ThisLtr Then
          NewLtr = ThisLtr
          NewAlphaFlag = True
          AlphaRealVal = 0
          AlphaBldgVal = 0
          AlphaPersVal = 0
          AlphaDscntVal = 0
          AlphaTotVal = 0
          AlphaNet = 0
          AlphaCnt = 0
          AlphaCustCnt = 0
        End If
      Else
        If WhatPers > 0 Then
          NewLtr = ThisLtr
        End If
      End If
'    End If
    
    CustRealVal = 0
    CustBldgVal = 0
    CustDscntVal = 0
    CustPersVal = 0
    CustNet = 0
    
    If QPTrim$(fpcmbPropType.Text) = "Personal Only" Then
      If QPTrim$(ThisTownship) = "All" And QPTrim$(fptxtAddress.Text) = "" Then
        GoTo PersOnly
      Else
        GoTo SkipIt
      End If
    End If
    
    Do While WhatPers > 0
      Get RHandle, WhatPers&, RealRec
      If RealRec.Deleted = True Then GoTo NextRec
      If ThisTownship <> "All" Then
        If QPTrim$(UCase(RealRec.TownShip)) <> ThisTownship Then GoTo NextRec
      End If
      If ThisAdd$ <> "" Then
        If InStr(RealRec.PropAddr, ThisAdd) = 0 Then
          GoTo NextRec
        End If
      End If
      GTotCnt = GTotCnt + 1
      ThisRealVal = RealRec.PROPVALU
      ThisBldgVal = RealRec.BldgVal
      ThisPersVal = 0
      ThisDscntVal = OldRound(RealRec.EXMPOTHR) '6/14/06
      ThisNet = OldRound(ThisRealVal + ThisBldgVal - ThisDscntVal)
      CustRealVal = OldRound(CustRealVal + RealRec.PROPVALU)
      
      CustBldgVal = OldRound(CustBldgVal + RealRec.BldgVal)
      CustDscntVal = OldRound(CustDscntVal + RealRec.EXMPOTHR) '6/14/06
      CustPersVal = CustPersVal + 0
      CustNet = OldRound(CustNet + ThisPersVal + ThisRealVal + ThisBldgVal - ThisDscntVal)
      GRealVal = OldRound(GRealVal + RealRec.PROPVALU)
      GBldgVal = OldRound(GBldgVal + RealRec.BldgVal)
      GPersVal = GPersVal + 0
      GDscntVal = OldRound(GDscntVal + RealRec.EXMPOTHR) '6/14/06
      GTotVal = OldRound(GTotVal + RealRec.PROPVALU + RealRec.EXMPOTHR) '6/14/06
      AlphaCnt = AlphaCnt + 1
      AlphaRealVal = OldRound(AlphaRealVal + RealRec.PROPVALU)
      AlphaBldgVal = OldRound(AlphaBldgVal + RealRec.BldgVal)
      AlphaPersVal = AlphaPersVal + 0
      AlphaDscntVal = OldRound(AlphaDscntVal + RealRec.EXMPOTHR) '6/14/06
      AlphaTotVal = OldRound(AlphaTotVal + RealRec.PROPVALU + RealRec.BldgVal + RealRec.EXMPOTHR) '6/14/06
      ThisPin = "Real Estate: " + QPTrim$(RealRec.RealPin)
      AlphaNet = OldRound(AlphaNet + ThisNet)
      GNet = OldRound(GNet + ThisNet)
      PrintThis = True
NextRec:
      WhatPers& = RealRec.NextRec
    Loop
    
    If QPTrim$(fpcmbPropType.Text) = "Real Only" Then GoTo AddUse 'SkipIt
    If ThisAdd$ <> "" Then GoTo AddUse
PersOnly:
    If QPTrim$(fpcmbPropType.Text) = "Both" Then
      If QPTrim$(ThisTownship) <> "All" Or QPTrim$(fptxtAddress.Text) <> "" Then
        GoTo AddUse 'SkipIt
      End If
    End If
    
    WhatPers = TaxCust.FirstPersRec
    Do While WhatPers > 0
      Get PHandle, WhatPers, PersRec
      GTotCnt = GTotCnt + 1
      ThisRealVal = 0
      ThisBldgVal = 0
      ThisDscntVal = 0 '6/14/06 no more pers exemptions OldRound(PersRec.EXMPSENI + PersRec.EXMPOTHR)
      ThisPersVal = OldRound(PersRec.PersVal + PersRec.MCValue + PersRec.CVALUE + PersRec.MHValue + PersRec.MTValue)
'      If PersRec.MHValue > 0 Then
'        Print #AHandle, CStr(TaxCust.Acct) + "~" + QPTrim$(TaxCust.CustName) + "~" + Using$("########.00", RealRec.PROPVALU)
'      End If
      ThisNet = OldRound(ThisPersVal - ThisDscntVal)
      ThisPin = "Property Pin: " + QPTrim$(PersRec.PropPin)
      CustRealVal = CustRealVal + 0
      CustBldgVal = CustBldgVal + 0
      CustDscntVal = 0 '6/14/06 no more per exemptions OldRound(CustDscntVal + PersRec.EXMPSENI + PersRec.EXMPOTHR)
      CustPersVal = OldRound(CustPersVal + ThisPersVal)
      CustNet = OldRound(CustNet + ThisPersVal + ThisRealVal + ThisBldgVal - ThisDscntVal)
      GRealVal = GRealVal + 0
      GBldgVal = GBldgVal + 0
      GPersVal = OldRound(GPersVal + ThisPersVal)
      GTPersVal = OldRound(GTPersVal + PersRec.PersVal) 'added 12/15/06
      GTMTVal = OldRound(GTMTVal + PersRec.MTValue) 'added 12/15/06
      GTMCVal = OldRound(GTMCVal + PersRec.MCValue) 'added 12/15/06
      GTMHVal = OldRound(GTMHVal + PersRec.MHValue) 'added 12/15/06
      GTFarmVal = OldRound(GTFarmVal + PersRec.CVALUE) 'added 12/15/06
'      GDscntVal = 0 'no more pers exemptions OldRound(GDscntVal + PersRec.EXMPOTHR + PersRec.EXMPSENI)
      GTotVal = OldRound(GTotVal + PersRec.PersVal) '6/14/06 no more pers exemptions  + PersRec.EXMPSENI + PersRec.EXMPOTHR)
      AlphaCnt = AlphaCnt + 1
      AlphaRealVal = AlphaRealVal + 0
      AlphaBldgVal = AlphaBldgVal + 0
      AlphaPersVal = OldRound(AlphaPersVal + ThisPersVal)
'      AlphaDscntVal = 0 '6/14/06 no more pers exemptions OldRound(AlphaDscntVal + PersRec.EXMPSENI + PersRec.EXMPOTHR)
      AlphaTotVal = OldRound(AlphaTotVal + PersRec.PersVal) '6/14/06 no more pers exemptions  + PersRec.EXMPSENI + PersRec.EXMPOTHR)
      AlphaNet = OldRound(AlphaNet + ThisNet)
      GNet = OldRound(GNet + ThisNet)
      PrintThis = True
'    End If
    WhatPers = PersRec.NextRec
    Loop
AddUse:
    If PrintThis = True Then
      GoSub PrintIt
    End If
    PrintThis = False
SkipIt:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  
  Close
  If GTotCnt = 0 Then
    Call TaxMsg(900, "There are no property valuations for the parameters entered.")
    Close
    Exit Sub
  End If
  
  arVATaxValSum.Show
  frmVATaxLoadReport.Show
  DoEvents
  
  Exit Sub

PrintIt:
  CustCnt = CustCnt + 1
  AlphaCustCnt = AlphaCustCnt + 1
  ThisName = QPTrim$(TaxCust.CustName)
  '                   0              1                           2
  Print #RptHandle, Town$; dlm; TaxCust.Acct; dlm; ThisName; dlm;
  '                    3                4                5                 6
  Print #RptHandle, ThisLtr; dlm; CustRealVal; dlm; CustPersVal; dlm; CustDscntVal; dlm;
  '                    7              8               9                  10
  Print #RptHandle, CustNet; dlm; AlphaCnt; dlm; AlphaRealVal; dlm; AlphaPersVal; dlm;
  '                      11                12             13            14
  Print #RptHandle, AlphaDscntVal; dlm; AlphaNet; dlm; GTotCnt; dlm; GRealVal; dlm;
  '                    15             16            17           18           19
  Print #RptHandle, GPersVal; dlm; GDscntVal; dlm; GNet; dlm; ThisPin; dlm; CustCnt; dlm;
  '                     20                21             22
  Print #RptHandle, AlphaCustCnt; dlm; ThisAdd; dlm; ActiveFlag; dlm;
  If UseOpt = "Y" Then
    '                    23                      24                       25              26               27                 28                29
    Print #RptHandle, ThisOpt; dlm; QPTrim$(TaxCust.OptSrchDesc); dlm; PropType; dlm; PrintOrder; dlm; CustBldgVal; dlm; AlphaBldgVal; dlm; GBldgVal; dlm;
  Else
    '                 23       24          25              26               27                 28               29
    Print #RptHandle, ""; dlm; ""; dlm; PropType; dlm; PrintOrder; dlm; CustBldgVal; dlm; AlphaBldgVal; dlm; GBldgVal; dlm;
  End If
  '                    30              31            32            33             34
  Print #RptHandle, GTPersVal; dlm; GTMTVal; dlm; GTMCVal; dlm; GTMHVal; dlm; GTFarmVal
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxValuationListing", "PrintGraphics", Erl)
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

Private Sub PrintTextDet()
  Dim x As Long, y As Integer
  Dim PropTypeFlag As Boolean
  Dim InactiveFlag As Boolean
  Dim ThisTownship$
  Dim ThisAdd$
  Dim RptFile$
  Dim RptHandle As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim WhatPers&
  Dim WhatReal&
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim ThisTot As Double
  Dim ThisLtr As String * 1
  Dim NewLtr As String * 1
  Dim NameFlag As Boolean
  Dim AlphaCnt As Integer
  Dim GTotCnt As Long
  Dim NewAlphaFlag As Boolean
  Dim SearchFlag As Boolean
  Dim GRealVal As Double
  Dim AlphaRealVal As Double
  Dim GBldgVal As Double
  Dim AlphaBldgVal As Double
  Dim GPersVal As Double
  Dim GTPersVal As Double
  Dim GTMTVal As Double
  Dim GTMCVal As Double
  Dim GTMHVal As Double
  Dim GTFarmVal As Double
  Dim AlphaPersVal As Double
  Dim GDscntVal As Double
  Dim AlphaDscntVal As Double
  Dim AlphaTotVal As Double
  Dim GTotVal As Double
  Dim ThisRealVal As Double
  Dim ThisBldgVal As Double
  Dim ThisDscntVal As Double
  Dim ThisPersVal As Double
  Dim ThisNet As Double
'  Dim ThisPin As String
  Dim AlphaNet As Double
  Dim GNet As Double
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim CustCnt As Long
  Dim AlphaCustCnt As Integer
  Dim Page As Integer
  Dim LineCnt As Integer
  Dim MaxLines As Integer
  Dim FF$
  Dim ThisName$
  Dim CustRealTot As Double
  Dim CustBldgTot As Double
  Dim CustPersTot As Double
  Dim CustDscntTot As Double
  Dim CustNetTot As Double
  Dim ThisCnt As Integer
  Dim ThisCustCnt As Integer
  Dim PrintThis As Boolean
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim OptFlag As Boolean
  Dim ActiveFlag As String * 1
  Dim PropType$
  Dim PrintOrder$
  
  On Error GoTo ERRORSTUFF
  
  If fpcmbPrintOrder.Text = "Acct Number Order" Then
    PrintOrder = "N"
  Else
    PrintOrder = "Y"
  End If
  
  PropType = fpcmbPropType.Text
  
  FF$ = Chr(12)
  MaxLines = 58
  OptFlag = False
  NewAlphaFlag = False
  SearchFlag = False
  If InStr(fpcmbPrintOrder.Text, "Name") Then
    NameFlag = True
  Else
    NameFlag = False
  End If

  ThisLtr = ""
  NewLtr = ""
  ThisAdd$ = QPTrim$(fptxtAddress.Text)
  ThisTownship = QPTrim$(fpcmbTownship.Text)
  ActiveFlag = "B"
  If QPTrim$(fpcmbIncInactive.Text) = "Active" Then
    ActiveFlag = "A"
  ElseIf QPTrim$(fpcmbIncInactive.Text) = "Inactive" Then
    ActiveFlag = "I"
  End If

  If QPTrim$(fpcmbPrintOrder.Text) = "Name Order" Then
    OpenNameIdxFile IdxHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmVATaxMsg.Label1.Caption = "There are no customers saved."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show
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
      frmVATaxMsg.Label1.Caption = "There are no search names indexed."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show
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
    SearchFlag = True
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

  RptFile$ = "TAXRPTS\TXVALLST.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle

  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenRealPropFile RHandle, NumOfRealRecs
  OpenPersPropFile PHandle, NumOfPersRecs
  GoSub PrintHeader
  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  End If
  frmVATaxShowPctComp.Label1 = "Gathering Valuation Data"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False

  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
    Else
      Get TCHandle, x, TaxCust
    End If
    If TaxCust.Deleted <> 0 Then GoTo SkipIt
    If ActiveFlag = "A" And TaxCust.Active = "N" Then
      GoTo SkipIt
    ElseIf ActiveFlag = "I" And TaxCust.Active = "Y" Then
      GoTo SkipIt
    End If
    If IdxFlag = True Then
      If SearchFlag = True Then
        ThisLtr = Mid(QPTrim$(TaxCust.SName), 1, 1)
      ElseIf OptFlag = True Then
        ThisLtr = Mid(QPTrim$(TaxCust.OptSrchDesc), 1, 1)
      Else
        ThisLtr = Mid(QPTrim$(TaxCust.CustName), 1, 1)
      End If
    End If
      WhatPers = TaxCust.FirstPropRec
      If GTotCnt > 0 And (WhatPers > 0 Or TaxCust.FirstPersRec > 0) Then
        If NewLtr <> ThisLtr Then
          If LineCnt >= MaxLines - 5 Then
            Print #RptHandle, FF$
            GoSub PrintHeader
          End If
          If AlphaNet > 0 Then
            GoSub PrintAlphaHeader
          End If
          NewLtr = ThisLtr
          NewAlphaFlag = True
          AlphaRealVal = 0
          AlphaBldgVal = 0
          AlphaPersVal = 0
          AlphaDscntVal = 0
          AlphaTotVal = 0
          AlphaNet = 0
          AlphaCnt = 0
          AlphaCustCnt = 0
        End If
      Else
        If WhatPers > 0 Then
          NewLtr = ThisLtr
        End If
      End If
'    End If
    ThisCustCnt = 0
    ReDim CustRealVal(1 To 1) As Double
    ReDim CustBldgVal(1 To 1) As Double
    ReDim CustDscntVal(1 To 1) As Double
    ReDim CustPersVal(1 To 1) As Double
    ReDim CustNet(1 To 1) As Double
    ReDim ThisPin(1 To 1) As String
    CustRealTot = 0
    CustBldgTot = 0
    CustDscntTot = 0
    CustPersTot = 0
    CustNetTot = 0
    ThisCustCnt = 0
    If QPTrim$(fpcmbPropType.Text) = "Personal Only" Then
      If QPTrim$(ThisTownship) = "All" And QPTrim$(fptxtAddress.Text) = "" Then
        GoTo PersOnly
      Else
        GoTo SkipIt
      End If
    End If

    Do While WhatPers > 0
      Get RHandle, WhatPers&, RealRec
      If RealRec.Deleted = True Then GoTo NextRec
      If ThisTownship <> "All" Then
        If QPTrim$(UCase(RealRec.TownShip)) <> ThisTownship Then GoTo NextRec
      End If
      If ThisAdd$ <> "" Then
        If InStr(RealRec.PropAddr, ThisAdd) = 0 Then
          GoTo NextRec
        End If
      End If
      GTotCnt = GTotCnt + 1
      ThisRealVal = RealRec.PROPVALU
      ThisBldgVal = RealRec.BldgVal
'      ThisDscntVal = OldRound(RealRec.EXMPSENI + RealRec.EXMPOTHR)
      ThisDscntVal = RealRec.EXMPOTHR '6/14/06
      ThisPersVal = 0
      ThisNet = OldRound(ThisRealVal + ThisBldgVal - ThisDscntVal)
      ThisCustCnt = ThisCustCnt + 1
      ReDim Preserve CustRealVal(1 To ThisCustCnt) As Double
      CustRealVal(ThisCustCnt) = OldRound(CustRealVal(ThisCustCnt) + RealRec.PROPVALU)
      CustRealTot = OldRound(CustRealTot + ThisRealVal)

      ReDim Preserve CustBldgVal(1 To ThisCustCnt) As Double
      CustBldgVal(ThisCustCnt) = OldRound(CustBldgVal(ThisCustCnt) + RealRec.BldgVal)
      CustBldgTot = OldRound(CustBldgTot + ThisBldgVal)

      ReDim Preserve CustDscntVal(1 To ThisCustCnt) As Double
      CustDscntVal(ThisCustCnt) = OldRound(CustDscntVal(ThisCustCnt) + ThisDscntVal)
      CustDscntTot = OldRound(CustDscntTot + ThisDscntVal)

      ReDim Preserve CustPersVal(1 To ThisCustCnt) As Double
      CustPersVal(ThisCustCnt) = OldRound(CustPersVal(ThisCustCnt) + 0)
      CustPersTot = OldRound(CustPersTot + ThisPersVal)

      ReDim Preserve CustNet(1 To ThisCustCnt) As Double
      CustNet(ThisCustCnt) = OldRound(CustNet(ThisCustCnt) + ThisNet)
      CustNetTot = OldRound(CustNetTot + ThisNet)
      
      ReDim Preserve ThisPin(1 To ThisCustCnt) As String
      ThisPin(ThisCustCnt) = "Real Pin: " + QPTrim$(RealRec.RealPin)
      GRealVal = OldRound(GRealVal + RealRec.PROPVALU)
      GBldgVal = OldRound(GBldgVal + RealRec.BldgVal)
      GPersVal = GPersVal + 0
'      GDscntVal = OldRound(GDscntVal + RealRec.EXMPSENI + RealRec.EXMPOTHR)
      GDscntVal = OldRound(GDscntVal + RealRec.EXMPOTHR) '6/14/06
'      GTotVal = OldRound(GTotVal + RealRec.PROPVALU + RealRec.EXMPSENI + RealRec.EXMPOTHR)
      GTotVal = OldRound(GTotVal + RealRec.PROPVALU + RealRec.EXMPOTHR) '6/14/06
      AlphaCnt = AlphaCnt + 1
      AlphaRealVal = OldRound(AlphaRealVal + RealRec.PROPVALU)
      AlphaBldgVal = OldRound(AlphaBldgVal + RealRec.BldgVal)
      AlphaPersVal = AlphaPersVal + 0
'      AlphaDscntVal = OldRound(AlphaDscntVal + RealRec.EXMPSENI + RealRec.EXMPOTHR)
      AlphaDscntVal = OldRound(AlphaDscntVal + RealRec.EXMPOTHR) '6/14/06
'      AlphaTotVal = OldRound(AlphaTotVal + RealRec.PROPVALU + RealRec.BldgVal + RealRec.EXMPSENI + RealRec.EXMPOTHR)
      AlphaTotVal = OldRound(AlphaTotVal + RealRec.PROPVALU + RealRec.BldgVal + RealRec.EXMPOTHR) '6/14/06
      AlphaNet = OldRound(AlphaNet + ThisNet)
      GNet = OldRound(GNet + ThisNet)
      PrintThis = True
NextRec:
      WhatPers& = RealRec.NextRec
    Loop
    
    If ThisAdd$ <> "" Then
      GoTo UseAdd
    End If

    If QPTrim$(fpcmbPropType.Text) = "Real Only" Then GoTo UseAdd ' SkipIt
PersOnly:
    If QPTrim$(fpcmbPropType.Text) = "Both" Then
      If QPTrim$(ThisTownship) <> "All" Or QPTrim$(fptxtAddress.Text) <> "" Then
        GoTo UseAdd 'SkipIt
      End If
    End If
    
    WhatPers = TaxCust.FirstPersRec
    Do While WhatPers > 0
      Get PHandle, WhatPers, PersRec
      GTotCnt = GTotCnt + 1
      ThisRealVal = 0
      ThisBldgVal = 0
      ThisDscntVal = 0 '6/14/06 no more pers exemptions OldRound(PersRec.EXMPSENI + PersRec.EXMPOTHR)
      ThisPersVal = OldRound(PersRec.PersVal + PersRec.MCValue + PersRec.CVALUE + PersRec.MHValue + PersRec.MTValue)
      ThisNet = OldRound(ThisPersVal - ThisDscntVal)
      ThisCustCnt = ThisCustCnt + 1
      
      ReDim Preserve CustRealVal(1 To ThisCustCnt) As Double
      CustRealVal(ThisCustCnt) = OldRound(CustRealVal(ThisCustCnt) + 0)
      CustRealTot = OldRound(CustRealTot + ThisRealVal)
      
      ReDim Preserve CustBldgVal(1 To ThisCustCnt) As Double
      CustBldgVal(ThisCustCnt) = OldRound(CustBldgVal(ThisCustCnt) + 0)
      CustBldgTot = OldRound(CustBldgTot + ThisBldgVal)
      
      ReDim Preserve CustDscntVal(1 To ThisCustCnt) As Double
      CustDscntVal(ThisCustCnt) = 0 '6/14/06 no more pers exemptions OldRound(CustDscntVal(ThisCustCnt) + PersRec.EXMPSENI + PersRec.EXMPOTHR)
      CustDscntTot = OldRound(CustDscntTot + ThisDscntVal)
      
      ReDim Preserve CustPersVal(1 To ThisCustCnt) As Double
      CustPersVal(ThisCustCnt) = OldRound(CustPersVal(ThisCustCnt) + ThisPersVal)
      CustPersTot = OldRound(CustPersTot + ThisPersVal)
      
      ReDim Preserve CustNet(1 To ThisCustCnt) As Double
      CustNet(ThisCustCnt) = OldRound(CustNet(ThisCustCnt) + ThisPersVal + ThisRealVal - ThisDscntVal)
      CustNetTot = OldRound(CustNetTot + ThisNet)
 
      ReDim Preserve ThisPin(1 To ThisCustCnt) As String
      ThisPin(ThisCustCnt) = "Pers Pin: " + QPTrim$(PersRec.PropPin)
      GRealVal = GRealVal + 0
      GBldgVal = GBldgVal + 0
      GPersVal = OldRound(GPersVal + ThisPersVal)
'      GDscntVal = 0 '6/14/06 no more pers exemptions OldRound(GDscntVal + PersRec.EXMPOTHR + PersRec.EXMPSENI)
      GTPersVal = OldRound(GTPersVal + PersRec.PersVal) 'added 12/15/06
      GTMTVal = OldRound(GTMTVal + PersRec.MTValue) 'added 12/15/06
      GTMCVal = OldRound(GTMCVal + PersRec.MCValue) 'added 12/15/06
      GTMHVal = OldRound(GTMHVal + PersRec.MHValue) 'added 12/15/06
      GTFarmVal = OldRound(GTFarmVal + PersRec.CVALUE) 'added 12/15/06
      GTotVal = OldRound(GTotVal + PersRec.PersVal) '6/14/06 no more pers exemptions + PersRec.EXMPSENI + PersRec.EXMPOTHR)
      AlphaCnt = AlphaCnt + 1
      AlphaRealVal = AlphaRealVal + 0
      AlphaBldgVal = AlphaBldgVal + 0
      AlphaPersVal = OldRound(AlphaPersVal + ThisPersVal)
'      AlphaDscntVal = 0 '6/14/06 no more pers exemptions OldRound(AlphaDscntVal + PersRec.EXMPSENI + PersRec.EXMPOTHR)
      AlphaTotVal = OldRound(AlphaTotVal + PersRec.PersVal) '6/14/06 no more pers exemptions + PersRec.EXMPSENI + PersRec.EXMPOTHR)
      AlphaNet = OldRound(AlphaNet + ThisNet)
      GNet = OldRound(GNet + ThisNet)
      PrintThis = True
      WhatPers& = PersRec.NextRec
    Loop
    
UseAdd:
    If PrintThis = True Then
      GoSub PrintIt
    End If
    PrintThis = False
SkipIt:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  If LineCnt >= MaxLines - 4 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  
  If AlphaNet > 0 Then
    GoSub PrintAlphaHeader
  End If
  If GTotCnt = 0 Then
    Call TaxMsg(900, "There are no property valuations for the parameters entered.")
    Close
    Exit Sub
  End If
  
  GoSub PrintSummary
  Print #RptHandle, FF$
  Close

  ViewPrint RptFile, "Tax Valuation Report", True

  Exit Sub

PrintIt:
  GoSub PrintCustHeader
  CustCnt = CustCnt + 1
  AlphaCustCnt = AlphaCustCnt + 1
  
  For y = 1 To ThisCustCnt
    Print #RptHandle, Tab(5); ThisPin(y); Tab(31); Using$("##,###,##0.00", CustRealVal(y)); Tab(45); Using$("##,###,##0.00", CustBldgVal(y)); Tab(59); Using$("##,###,##0.00", CustPersVal(y)); Tab(73); Using$("##,###,##0.00", CustDscntVal(y)); Tab(87); Using$("##,###,##0.00", CustNet(y))
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintHeader
      GoSub PrintCustHeader
    End If
  Next y
  
  Return

PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(30); "Tax Valuations Listing Detail"
  Print #RptHandle, Town; Tab(75); "Page #: " + CStr(Page)
  Print #RptHandle, "Report Date: " + CStr(Date); Tab(65); "Property Type: " + PropType
  If QPTrim$(fptxtAddress.Text) <> "" Then
    Print #RptHandle, "Address Key: " + QPTrim$(fptxtAddress.Text)
  Else
    Print #RptHandle, "Address Key: All Addresses Included"
  End If
  If ActiveFlag = "B" Then
    Print #RptHandle, "Active and Inactive"
  ElseIf ActiveFlag = "A" Then
    Print #RptHandle, "Active Only"
  ElseIf ActiveFlag = "I" Then
    Print #RptHandle, "Inactive Only"
  End If
  Print #RptHandle, Tab(31); "[-----------------------------Valuations----------------------------]"
  Print #RptHandle, "Acct# "; Tab(7); "Customer Name"; Tab(40); "Real"; Tab(50); "Building"; Tab(68); "Pers"; Tab(80); "Discnt"; Tab(97); "Net"
  Print #RptHandle, String(99, "-")
  If ActiveFlag = "I" Then
    LineCnt = 8
  Else
    LineCnt = 7
  End If

  Return

PrintCustHeader:
  If UseOpt = "Y" Then
    If LineCnt >= MaxLines - 5 Then
      Print #RptHandle, FF$
      GoSub PrintHeader
    End If
  Else
    If LineCnt >= MaxLines - 4 Then
      Print #RptHandle, FF$
      GoSub PrintHeader
    End If
  End If
  If ActiveFlag <> "I" And LineCnt <> 7 Then
    Print #RptHandle, String(99, "-")
    LineCnt = LineCnt + 1
  ElseIf ActiveFlag = "I" And LineCnt <> 8 Then
    Print #RptHandle, String(99, "-")
    LineCnt = LineCnt + 1
  End If
    ThisName = QPTrim$(TaxCust.CustName)
  If Len(ThisName) > 24 Then
  LineCnt = LineCnt + 1
  End If
  Print #RptHandle, Using$("####0", TaxCust.Acct); Tab(7); ThisName; Tab(31); Using$("##,###,##0.00", CustRealTot); Tab(45); Using$("##,###,##0.00", CustBldgTot); Tab(59); Using$("##,###,##0.00", CustPersTot); Tab(73); Using$("##,###,##0.00", CustDscntTot); Tab(87); Using$("##,###,##0.00", CustNetTot)
  If UseOpt = "Y" Then
    Print #RptHandle, Tab(7); ThisOpt + ": " + QPTrim$(TaxCust.OptSrchDesc)
    LineCnt = LineCnt + 1
  End If
  Print #RptHandle, Tab(5); String(81, ".")
  LineCnt = LineCnt + 2
  
  Return
  
PrintAlphaHeader:
  If PrintOrder = "N" Then Return
  If ActiveFlag <> "I" And LineCnt <> 7 Then
    Print #RptHandle, String(99, "-")
    LineCnt = LineCnt + 1
  ElseIf ActiveFlag = "I" And LineCnt <> 8 Then
    Print #RptHandle, String(99, "-")
    LineCnt = LineCnt + 1
  End If
  Print #RptHandle, Tab(2); "Summary for '" + NewLtr + "'"; Tab(18); "Cnt: " + Using$("####0", AlphaCustCnt); Tab(31); Using$("##,###,##0.00", AlphaRealVal); Tab(45); Using$("##,###,##0.00", AlphaBldgVal); Tab(59); Using$("##,###,##0.00", AlphaPersVal); Tab(73); Using$("##,###,##0.00", AlphaDscntVal); Tab(87); Using$("##,###,##0.00", AlphaNet)
  Print #RptHandle, String(99, "-")
  Print #RptHandle,
  Print #RptHandle,
  LineCnt = LineCnt + 4
  
  Return
  
PrintSummary:
  If LineCnt >= MaxLines - 10 Then
    Print #RptHandle, FF$
    Page = Page + 1
    Print #RptHandle, Tab(30); "Tax Valuations Listing Detail"
    Print #RptHandle, Town; Tab(75); "Page #: " + CStr(Page)
    Print #RptHandle, "Report Date: " + CStr(Date)
    If QPTrim$(fptxtAddress.Text) <> "" Then
      Print #RptHandle, "Address Key: " + QPTrim$(fptxtAddress.Text)
    Else
      Print #RptHandle, "Address Key: All Addresses Included"
    End If
  End If
  Print #RptHandle, String(99, "_")
  Print #RptHandle,
  Print #RptHandle, "Total Customers Printed:                   " + Using$("####0", CustCnt)
  Print #RptHandle, "Total Transactions:                       " + Using$("#####0", GTotCnt)
  Print #RptHandle,
  Print #RptHandle, "Total Real Value:              " + Using$("$#,###,###,##0.00", GRealVal)
  Print #RptHandle, "Total Building Value:          " + Using$("$#,###,###,##0.00", GBldgVal)
  Print #RptHandle,
  Print #RptHandle, "Total Personal Only Value:     " + Using$("$#,###,###,##0.00", GTPersVal)
  Print #RptHandle, "Total Machine Tools Value:     " + Using$("$#,###,###,##0.00", GTMTVal)
  Print #RptHandle, "Total Merchant Capital Value:  " + Using$("$#,###,###,##0.00", GTMCVal)
  Print #RptHandle, "Total Mobile Homes Value:      " + Using$("$#,###,###,##0.00", GTMHVal)
  Print #RptHandle, "Total Farm Equipment Value:    " + Using$("$#,###,###,##0.00", GTFarmVal)
  Print #RptHandle, String(48, "-")
  Print #RptHandle, "Total All Personal Value:      " + Using$("$#,###,###,##0.00", GPersVal)
  Print #RptHandle, "Total Discount Value:          " + Using$("$#,###,###,##0.00", GDscntVal)
  Print #RptHandle,
  Print #RptHandle, "Total Net Value:               " + Using$("$#,###,###,##0.00", GNet)
  
  
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxValuationListing", "PrintTextDet", Erl)
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
  Dim x As Long, y As Integer
  Dim PropTypeFlag As Boolean
  Dim InactiveFlag As Boolean
  Dim ThisTownship$
  Dim ThisAdd$
  Dim RptFile$
  Dim RptHandle As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim WhatPers&
  Dim WhatReal&
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim ThisTot As Double
  Dim ThisLtr As String * 1
  Dim NewLtr As String * 1
  Dim NameFlag As Boolean
  Dim AlphaCnt As Integer
  Dim GTotCnt As Long
  Dim NewAlphaFlag As Boolean
  Dim SearchFlag As Boolean
  Dim GRealVal As Double
  Dim AlphaRealVal As Double
  Dim GBldgVal As Double
  Dim AlphaBldgVal As Double
  Dim GPersVal As Double
  Dim GTPersVal As Double
  Dim GTMTVal As Double
  Dim GTMCVal As Double
  Dim GTMHVal As Double
  Dim GTFarmVal As Double
  Dim AlphaPersVal As Double
  Dim GDscntVal As Double
  Dim AlphaDscntVal As Double
  Dim AlphaTotVal As Double
  Dim GTotVal As Double
  Dim ThisRealVal As Double
  Dim ThisBldgVal As Double
  Dim ThisDscntVal As Double
  Dim ThisPersVal As Double
  Dim ThisNet As Double
'  Dim ThisPin As String
  Dim AlphaNet As Double
  Dim GNet As Double
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim CustCnt As Long
  Dim AlphaCustCnt As Integer
  Dim Page As Integer
  Dim LineCnt As Integer
  Dim MaxLines As Integer
  Dim FF$
  Dim ThisName$
  Dim CustRealTot As Double
  Dim CustBldgTot As Double
  Dim CustPersTot As Double
  Dim CustDscntTot As Double
  Dim CustNetTot As Double
  Dim ThisCnt As Integer
  Dim ThisCustCnt As Integer
  Dim PrintThis As Boolean
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim OptFlag As Boolean
  Dim ActiveFlag As String * 1
  Dim PropType$
  Dim PrintOrder$
  
  On Error GoTo ERRORSTUFF
  
  If fpcmbPrintOrder.Text = "Acct Number Order" Then
    PrintOrder = "N"
  Else
    PrintOrder = "Y"
  End If
  
  PropType = fpcmbPropType.Text
  
  FF$ = Chr(12)
  MaxLines = 58
  NewAlphaFlag = False
  OptFlag = False
  SearchFlag = False
  If InStr(fpcmbPrintOrder.Text, "Name") Then
    NameFlag = True
  Else
    NameFlag = False
  End If

  ThisLtr = ""
  NewLtr = ""
  ThisAdd$ = QPTrim$(fptxtAddress.Text)
  ThisTownship = QPTrim$(fpcmbTownship.Text)
  ActiveFlag = "B"
  If QPTrim$(fpcmbIncInactive.Text) = "Active" Then
    ActiveFlag = "A"
  ElseIf QPTrim$(fpcmbIncInactive.Text) = "Inactive" Then
    ActiveFlag = "I"
  End If

  If QPTrim$(fpcmbPrintOrder.Text) = "Name Order" Then
    OpenNameIdxFile IdxHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmVATaxMsg.Label1.Caption = "There are no customers saved."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show
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
      frmVATaxMsg.Label1.Caption = "There are no search names indexed."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show
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
    SearchFlag = True
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

  RptFile$ = "TAXRPTS\TXVALLST.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle

  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenRealPropFile RHandle, NumOfRealRecs
  OpenPersPropFile PHandle, NumOfPersRecs
  GoSub PrintHeader
  
  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  End If
  
  frmVATaxShowPctComp.Label1 = "Gathering Valuation Data"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False

  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
    Else
      Get TCHandle, x, TaxCust
    End If
    If TaxCust.Deleted <> 0 Then GoTo SkipIt
    If ActiveFlag = "A" And TaxCust.Active = "N" Then
      GoTo SkipIt
    ElseIf ActiveFlag = "I" And TaxCust.Active = "Y" Then
      GoTo SkipIt
    End If
    If IdxFlag = True Then
      If SearchFlag = True Then
        ThisLtr = Mid(QPTrim$(TaxCust.SName), 1, 1)
      ElseIf OptFlag = True Then
        ThisLtr = Mid(QPTrim$(TaxCust.OptSrchDesc), 1, 1)
      Else
        ThisLtr = Mid(QPTrim$(TaxCust.CustName), 1, 1)
      End If
    End If
      WhatPers = TaxCust.FirstPropRec
      If GTotCnt > 0 And (WhatPers > 0 Or TaxCust.FirstPersRec > 0) Then
        If NewLtr <> ThisLtr Then
          If LineCnt >= MaxLines - 5 Then
            Print #RptHandle, FF$
            GoSub PrintHeader
          End If
           
          If AlphaNet > 0 Then
            GoSub PrintAlphaHeader
          End If
          NewLtr = ThisLtr
          NewAlphaFlag = True
          AlphaRealVal = 0
          AlphaBldgVal = 0
          AlphaPersVal = 0
          AlphaDscntVal = 0
          AlphaTotVal = 0
          AlphaNet = 0
          AlphaCnt = 0
          AlphaCustCnt = 0
        End If
      Else
        If WhatPers > 0 Then
          NewLtr = ThisLtr
        End If
      End If
'    End If
    ThisCustCnt = 0
    ReDim CustRealVal(1 To 1) As Double
    ReDim CustBldgVal(1 To 1) As Double
    ReDim CustDscntVal(1 To 1) As Double
    ReDim CustPersVal(1 To 1) As Double
    ReDim CustNet(1 To 1) As Double
    ReDim ThisPin(1 To 1) As String
    CustRealTot = 0
    CustBldgTot = 0
    CustDscntTot = 0
    CustPersTot = 0
    CustNetTot = 0
    
    If QPTrim$(fpcmbPropType.Text) = "Personal Only" Then
      If QPTrim$(ThisTownship) = "All" And QPTrim$(fptxtAddress.Text) = "" Then
        GoTo PersOnly
      Else
        GoTo SkipIt
      End If
    End If

    Do While WhatPers > 0
      Get RHandle, WhatPers&, RealRec
      If RealRec.Deleted = True Then GoTo NextRec
      If ThisTownship <> "All" Then
        If QPTrim$(UCase(RealRec.TownShip)) <> ThisTownship Then GoTo NextRec
      End If
      If ThisAdd$ <> "" Then
        If InStr(RealRec.PropAddr, ThisAdd) = 0 Then
          GoTo NextRec
        End If
      End If
      GTotCnt = GTotCnt + 1
      ThisRealVal = RealRec.PROPVALU
      ThisBldgVal = RealRec.BldgVal
'      ThisDscntVal = OldRound(RealRec.EXMPSENI + RealRec.EXMPOTHR)
      ThisDscntVal = RealRec.EXMPOTHR '6/14/06
      ThisPersVal = 0
      ThisNet = OldRound(ThisRealVal + ThisBldgVal - ThisDscntVal)
      ThisCustCnt = 1
      ReDim Preserve CustRealVal(1 To ThisCustCnt) As Double
      CustRealVal(ThisCustCnt) = OldRound(CustRealVal(ThisCustCnt) + RealRec.PROPVALU)
      CustRealTot = OldRound(CustRealTot + ThisRealVal)

      ReDim Preserve CustBldgVal(1 To ThisCustCnt) As Double
      CustBldgVal(ThisCustCnt) = OldRound(CustBldgVal(ThisCustCnt) + RealRec.BldgVal)
      CustBldgTot = OldRound(CustBldgTot + ThisBldgVal)

      ReDim Preserve CustDscntVal(1 To ThisCustCnt) As Double
      CustDscntVal(ThisCustCnt) = OldRound(CustDscntVal(ThisCustCnt) + ThisDscntVal)
      CustDscntTot = OldRound(CustDscntTot + ThisDscntVal)

      ReDim Preserve CustPersVal(1 To ThisCustCnt) As Double
      CustPersVal(ThisCustCnt) = OldRound(CustPersVal(ThisCustCnt) + 0)
      CustPersTot = OldRound(CustPersTot + ThisPersVal)

      ReDim Preserve CustNet(1 To ThisCustCnt) As Double
      CustNet(ThisCustCnt) = OldRound(CustNet(ThisCustCnt) + ThisNet)
      CustNetTot = OldRound(CustNetTot + ThisNet)
      
      ReDim Preserve ThisPin(1 To ThisCustCnt) As String
      ThisPin(ThisCustCnt) = "Real Estate: " + QPTrim$(RealRec.RealPin)
      GRealVal = OldRound(GRealVal + RealRec.PROPVALU)
      GBldgVal = OldRound(GBldgVal + RealRec.BldgVal)
      GPersVal = GPersVal + 0
'      GDscntVal = OldRound(GDscntVal + RealRec.EXMPSENI + RealRec.EXMPOTHR)
      GDscntVal = OldRound(GDscntVal + RealRec.EXMPOTHR) '6/14/06
'      GTotVal = OldRound(GTotVal + RealRec.PROPVALU + RealRec.BldgVal + RealRec.EXMPSENI + RealRec.EXMPOTHR)
      GTotVal = OldRound(GTotVal + RealRec.PROPVALU + RealRec.BldgVal + RealRec.EXMPOTHR) '6/14/06
      AlphaCnt = AlphaCnt + 1
      AlphaRealVal = OldRound(AlphaRealVal + RealRec.PROPVALU)
      AlphaBldgVal = OldRound(AlphaBldgVal + RealRec.BldgVal)
      AlphaPersVal = AlphaPersVal + 0
'      AlphaDscntVal = OldRound(AlphaDscntVal + RealRec.EXMPSENI + RealRec.EXMPOTHR)
      AlphaDscntVal = OldRound(AlphaDscntVal + RealRec.EXMPOTHR) '6/14/06
'      AlphaTotVal = OldRound(AlphaTotVal + RealRec.PROPVALU + RealRec.BldgVal + RealRec.EXMPSENI + RealRec.EXMPOTHR)
      AlphaTotVal = OldRound(AlphaTotVal + RealRec.PROPVALU + RealRec.BldgVal + RealRec.EXMPOTHR) '6/14/06
      AlphaNet = OldRound(AlphaNet + ThisNet)
      GNet = OldRound(GNet + ThisNet)
      PrintThis = True
NextRec:
      WhatPers& = RealRec.NextRec
    Loop
    
    If ThisAdd$ <> "" Then
      GoTo UseAdd
    End If

    If QPTrim$(fpcmbPropType.Text) = "Real Only" Then GoTo UseAdd 'SkipIt
PersOnly:
    If QPTrim$(fpcmbPropType.Text) = "Both" Then
      If QPTrim$(ThisTownship) <> "All" Or QPTrim$(fptxtAddress.Text) <> "" Then
        GoTo UseAdd ' SkipIt
      End If
    End If
    
   WhatPers = TaxCust.FirstPersRec
    Do While WhatPers > 0
      Get PHandle, WhatPers, PersRec
      GTotCnt = GTotCnt + 1
      ThisRealVal = 0
      ThisBldgVal = 0
      ThisDscntVal = 0 ' 6/14/06 no more personal exemptions OldRound(PersRec.EXMPSENI + PersRec.EXMPOTHR)
      ThisPersVal = OldRound(PersRec.PersVal + PersRec.MCValue + PersRec.CVALUE + PersRec.MHValue + PersRec.MTValue)
      ThisNet = OldRound(ThisPersVal - ThisDscntVal)
      ThisCustCnt = ThisCustCnt + 1
      
      ReDim Preserve CustRealVal(1 To ThisCustCnt) As Double
      CustRealVal(ThisCustCnt) = OldRound(CustRealVal(ThisCustCnt) + 0)
      CustRealTot = OldRound(CustRealTot + ThisRealVal)
      
      ReDim Preserve CustBldgVal(1 To ThisCustCnt) As Double
      CustBldgVal(ThisCustCnt) = OldRound(CustBldgVal(ThisCustCnt) + 0)
      CustBldgTot = OldRound(CustBldgTot + ThisBldgVal)
      
      ReDim Preserve CustDscntVal(1 To ThisCustCnt) As Double
      CustDscntVal(ThisCustCnt) = 0 ' 6/14/06 no more pers exemptions OldRound(CustDscntVal(ThisCustCnt) + PersRec.EXMPSENI + PersRec.EXMPOTHR)
      CustDscntTot = OldRound(CustDscntTot + ThisDscntVal)
      
      ReDim Preserve CustPersVal(1 To ThisCustCnt) As Double
      CustPersVal(ThisCustCnt) = OldRound(CustPersVal(ThisCustCnt) + ThisPersVal)
      CustPersTot = OldRound(CustPersTot + ThisPersVal)
      
      ReDim Preserve CustNet(1 To ThisCustCnt) As Double
      CustNet(ThisCustCnt) = OldRound(CustNet(ThisCustCnt) + ThisPersVal + ThisRealVal + ThisBldgVal - ThisDscntVal)
      CustNetTot = OldRound(CustNetTot + ThisNet)
 
      ReDim Preserve ThisPin(1 To ThisCustCnt) As String
      ThisPin(ThisCustCnt) = "Property Pin: " + QPTrim$(PersRec.PropPin)
      GRealVal = GRealVal + 0
      GBldgVal = GBldgVal + 0
      GPersVal = OldRound(GPersVal + ThisPersVal)
'      GDscntVal = 0 '6/14/06 no more personal exemptions OldRound(GDscntVal + PersRec.EXMPOTHR + PersRec.EXMPSENI)
      GTPersVal = OldRound(GTPersVal + PersRec.PersVal) 'added 12/15/06
      GTMTVal = OldRound(GTMTVal + PersRec.MTValue) 'added 12/15/06
      GTMCVal = OldRound(GTMCVal + PersRec.MCValue) 'added 12/15/06
      GTMHVal = OldRound(GTMHVal + PersRec.MHValue) 'added 12/15/06
      GTFarmVal = OldRound(GTFarmVal + PersRec.CVALUE) 'added 12/15/06
      GTotVal = OldRound(GTotVal + PersRec.PersVal) '6/14/06 no more pers exemptions + PersRec.EXMPSENI + PersRec.EXMPOTHR)
      AlphaCnt = AlphaCnt + 1
      AlphaRealVal = AlphaRealVal + 0
      AlphaPersVal = OldRound(AlphaPersVal + ThisPersVal)
'      AlphaDscntVal = 0 '6/14/06 no more personal exemptions OldRound(AlphaDscntVal + PersRec.EXMPSENI + PersRec.EXMPOTHR)
      AlphaTotVal = OldRound(AlphaTotVal + PersRec.PersVal) '6/14/06 no more pers exemptions + PersRec.EXMPSENI + PersRec.EXMPOTHR)
      AlphaNet = OldRound(AlphaNet + ThisNet)
      GNet = OldRound(GNet + ThisNet)
      PrintThis = True
'    End If
      WhatPers = PersRec.NextRec
    Loop
UseAdd:
    If PrintThis = True Then
      GoSub PrintIt
    End If
    PrintThis = False
SkipIt:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  If LineCnt >= MaxLines - 4 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  
  If AlphaNet > 0 Then
    GoSub PrintAlphaHeader
  End If
  If GTotCnt = 0 Then
    Call TaxMsg(900, "There are no property valuations for the parameters entered.")
    Close
    Exit Sub
  End If
  
  GoSub PrintSummary
  Print #RptHandle, FF$
  Close

  ViewPrint RptFile, "Tax Valuation Report", True

  Exit Sub

PrintIt:
  GoSub PrintCustHeader
  CustCnt = CustCnt + 1
  AlphaCustCnt = AlphaCustCnt + 1
  
  Return
  
PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(30); "Tax Valuations Listing Summary"
  Print #RptHandle, Town; Tab(75); "Page #: " + CStr(Page)
  Print #RptHandle, "Report Date: " + CStr(Date); Tab(58); "Property Type: " + PropType
  If QPTrim$(fptxtAddress.Text) <> "" Then
    Print #RptHandle, "Address Key: " + QPTrim$(fptxtAddress.Text)
  Else
    Print #RptHandle, "Address Key: All Addresses Included"
  End If
  If ActiveFlag = "B" Then
    Print #RptHandle, "Active and Inactive"
  ElseIf ActiveFlag = "A" Then
    Print #RptHandle, "Active Only"
  ElseIf ActiveFlag = "I" Then
    Print #RptHandle, "Inactive Only"
  End If
  Print #RptHandle, Tab(31); "[-----------------------------Valuations----------------------------]"
  Print #RptHandle, "Acct# "; Tab(7); "Customer Name"; Tab(40); "Real"; Tab(54); "Bldg"; Tab(68); "Pers"; Tab(80); "Discnt"; Tab(97); "Net"
  Print #RptHandle, String(99, "-")
  If ActiveFlag = "I" Then
    LineCnt = 8
  Else
    LineCnt = 7
  End If

  Return

PrintCustHeader:
  If LineCnt >= MaxLines - 4 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  If ActiveFlag = "I" Then
    If LineCnt <> 8 Then
      Print #RptHandle, String(99, "-")
      LineCnt = LineCnt + 1
    End If
  Else
    If LineCnt <> 7 Then
      Print #RptHandle, String(99, "-")
      LineCnt = LineCnt + 1
    End If
  End If
  ThisName = QPTrim$(TaxCust.CustName)
  If Len(ThisName) > 24 Then
    LineCnt = LineCnt + 1
  End If
  Print #RptHandle, Using$("####0", TaxCust.Acct); Tab(7); ThisName; Tab(31); Using$("##,###,##0.00", CustRealTot); Tab(45); Using$("##,###,##0.00", CustBldgTot); Tab(59); Using$("##,###,##0.00", CustPersTot); Tab(73); Using$("##,###,##0.00", CustDscntTot); Tab(87); Using$("##,###,##0.00", CustNetTot)
  LineCnt = LineCnt + 1
  If UseOpt = "Y" Then
    Print #RptHandle, Tab(7); ThisOpt + ":" + QPTrim$(TaxCust.OptSrchDesc)
    LineCnt = LineCnt + 1
  End If

  Return
  
PrintAlphaHeader:
  If PrintOrder = "N" Then Return
  If LineCnt >= MaxLines - 4 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  If ActiveFlag <> "I" And LineCnt <> 7 Then
    Print #RptHandle, String(99, "-")
    LineCnt = LineCnt + 1
  ElseIf ActiveFlag = "I" And LineCnt <> 8 Then
    Print #RptHandle, String(99, "-")
    LineCnt = LineCnt + 1
  End If
  Print #RptHandle, Tab(2); "Summary for '" + NewLtr + "'"; Tab(18); "Cnt: " + Using$("####0", AlphaCustCnt); Tab(31); Using$("##,###,##0.00", AlphaRealVal); Tab(45); Using$("##,###,##0.00", AlphaBldgVal); Tab(59); Using$("##,###,##0.00", AlphaPersVal); Tab(73); Using$("##,###,##0.00", AlphaDscntVal); Tab(87); Using$("##,###,##0.00", AlphaNet)
  Print #RptHandle, String(99, "-")
  Print #RptHandle,
  Print #RptHandle,
  LineCnt = LineCnt + 4
  
  Return
  
PrintSummary:
  If LineCnt >= MaxLines - 10 Then
    Print #RptHandle, FF$
    Page = Page + 1
    Print #RptHandle, Tab(30); "Tax Valuations Listing Summary"
    Print #RptHandle, Town; Tab(75); "Page #: " + CStr(Page)
    Print #RptHandle, "Report Date: " + CStr(Date)
    If QPTrim$(fptxtAddress.Text) <> "" Then
      Print #RptHandle, "Address Key: " + QPTrim$(fptxtAddress.Text)
    Else
      Print #RptHandle, "Address Key: All Addresses Included"
    End If
  End If
  Print #RptHandle, String(99, "_")
  Print #RptHandle,
  Print #RptHandle, "Total Customers Printed:                   " + Using$("####0", CustCnt)
  Print #RptHandle, "Total Transactions:                       " + Using$("#####0", GTotCnt)
  Print #RptHandle,
  Print #RptHandle, "Total Real Value:              " + Using$("$#,###,###,##0.00", GRealVal)
  Print #RptHandle, "Total Building Value:          " + Using$("$#,###,###,##0.00", GBldgVal)
  Print #RptHandle,
  Print #RptHandle, "Total Personal Only Value:     " + Using$("$#,###,###,##0.00", GTPersVal)
  Print #RptHandle, "Total Machine Tools Value:     " + Using$("$#,###,###,##0.00", GTMTVal)
  Print #RptHandle, "Total Merchant Capital Value:  " + Using$("$#,###,###,##0.00", GTMCVal)
  Print #RptHandle, "Total Mobile Homes Value:      " + Using$("$#,###,###,##0.00", GTMHVal)
  Print #RptHandle, "Total Farm Equipment Value:    " + Using$("$#,###,###,##0.00", GTFarmVal)
  Print #RptHandle, String(48, "-")
  Print #RptHandle, "Total All Personal Value:      " + Using$("$#,###,###,##0.00", GPersVal)
  Print #RptHandle, "Total Discount Value:          " + Using$("$#,###,###,##0.00", GDscntVal)
  Print #RptHandle,
  Print #RptHandle, "Total Net Value:               " + Using$("$#,###,###,##0.00", GNet)
  
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxValuationListing", "PrintText", Erl)
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

Private Sub fpcmbDetSum_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbDetSum.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbDetSum.ListIndex = -1
  End If
  If fpcmbPrintOpt.ListDown <> True Then
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

Private Sub fpcmbIncInactive_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbIncInactive.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbIncInactive.ListIndex = -1
  End If
  If fpcmbIncInactive.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbDetSum.SetFocus
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
      If fptxtAddress.Enabled = True Then
        fptxtAddress.SetFocus
      Else
        fpcmbTownship.SetFocus
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

Private Sub fpcmbPropType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbPropType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPropType.ListIndex = -1
  End If
  If fpcmbPropType.ListDown <> True Then
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
      fpcmbPropType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
End Sub
