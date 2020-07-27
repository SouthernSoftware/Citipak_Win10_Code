VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmVATaxCustListRpt 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Customer Listing"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxCustListRpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   7116
      Left            =   1920
      TabIndex        =   3
      Top             =   822
      Width           =   7788
      _Version        =   196609
      _ExtentX        =   13737
      _ExtentY        =   12552
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmVATaxCustListRpt.frx":08CA
      Begin LpLib.fpCombo fpcmbFlags 
         Height          =   384
         Left            =   3528
         TabIndex        =   2
         Top             =   2520
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
         ColDesigner     =   "frmVATaxCustListRpt.frx":08E6
      End
      Begin LpLib.fpCombo fpcmbIncInactive 
         Height          =   384
         Left            =   3528
         TabIndex        =   1
         Top             =   1956
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
         ColDesigner     =   "frmVATaxCustListRpt.frx":0BDD
      End
      Begin LpLib.fpCombo fpcmbDetSum 
         Height          =   384
         Left            =   3528
         TabIndex        =   0
         Top             =   1416
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
         ColDesigner     =   "frmVATaxCustListRpt.frx":0ED4
      End
      Begin LpLib.fpCombo fpcmbPrintOrder 
         Height          =   384
         Left            =   2928
         TabIndex        =   6
         Top             =   4236
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
         ColDesigner     =   "frmVATaxCustListRpt.frx":11CB
      End
      Begin LpLib.fpCombo fpcmbPrintOpt 
         Height          =   384
         Left            =   2928
         TabIndex        =   5
         Top             =   3648
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
         ColDesigner     =   "frmVATaxCustListRpt.frx":14C2
      End
      Begin LpLib.fpCombo fpcmbTownship 
         Height          =   384
         Left            =   2928
         TabIndex        =   4
         Top             =   3072
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
         ColDesigner     =   "frmVATaxCustListRpt.frx":17B9
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00D0D0D0&
         Height          =   852
         Left            =   1440
         TabIndex        =   16
         Top             =   4920
         Width           =   5172
         Begin VB.OptionButton OptBothW 
            BackColor       =   &H00D0D0D0&
            Caption         =   "Both With Prop Only"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   240
            TabIndex        =   21
            Top             =   480
            Width           =   2532
         End
         Begin VB.OptionButton OptBothWO 
            BackColor       =   &H00D0D0D0&
            Caption         =   "Both Even With No Prop"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   240
            TabIndex        =   19
            Top             =   120
            Width           =   2532
         End
         Begin VB.OptionButton OptReal 
            BackColor       =   &H00D0D0D0&
            Caption         =   "Real Only"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   3480
            TabIndex        =   18
            Top             =   120
            Width           =   1572
         End
         Begin VB.OptionButton OptPers 
            BackColor       =   &H00D0D0D0&
            Caption         =   "Personal Only"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   3480
            TabIndex        =   17
            Top             =   480
            Width           =   1572
         End
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   636
         Left            =   2040
         TabIndex        =   14
         TabStop         =   0   'False
         Tag             =   "Press the 'Cancel' button to exit this screen and return to the main 'Business License Reports' menu."
         Top             =   6120
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
         ButtonDesigner  =   "frmVATaxCustListRpt.frx":1AB0
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   636
         Left            =   4272
         TabIndex        =   15
         TabStop         =   0   'False
         Tag             =   $"frmVATaxCustListRpt.frx":1C8E
         Top             =   6120
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
         ButtonDesigner  =   "frmVATaxCustListRpt.frx":1D39
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Select Cust Type:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   348
         Left            =   2760
         TabIndex        =   20
         Top             =   4680
         Width           =   1980
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Customer Flags:"
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
         Left            =   1488
         TabIndex        =   13
         Top             =   2628
         Width           =   1860
      End
      Begin VB.Label Label2 
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
         Left            =   1392
         TabIndex        =   12
         Top             =   2040
         Width           =   1956
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
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
         Height          =   348
         Left            =   1272
         TabIndex        =   11
         Top             =   3180
         Width           =   1500
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
         Left            =   1440
         TabIndex        =   10
         Top             =   1524
         Width           =   1908
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   4620
         Left            =   1008
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
         Left            =   1476
         TabIndex        =   9
         Top             =   4320
         Width           =   1308
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Master Customer Listing"
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
         Left            =   1800
         TabIndex        =   8
         Top             =   450
         Width           =   4335
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   1530
         Top             =   315
         Width           =   4905
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
         Height          =   348
         Left            =   1272
         TabIndex        =   7
         Top             =   3756
         Width           =   1500
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   7380
      Left            =   1800
      Top             =   678
      Width           =   8052
   End
End
Attribute VB_Name = "frmVATaxCustListRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  'Private Temp_Class As Resize_Class
  Dim Town$
  Dim ThisOpt$
  Dim UseOpt As String * 1
  Dim Opt1Desc$
  Dim Opt2Desc$
  Dim Opt3Desc$
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
      frmVATaxMsg.Label1.Caption = "Pitch 12 is recommended for this printout."
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
  'Set Temp_Class = New Resize_Class
  'Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpMasterCustomer
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxCustListRpt.")
      Call Terminate
      End
    End If
  End If

End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    'Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub LoadMe()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim TSRec As TownshipType
  Dim TSCnt As Integer
  Dim TSHandle As Integer
  Dim x As Integer
  
  On Error GoTo ERRORSTUFF
  
  OptBothWO.Value = True
  UseOpt = "N"
  If Exist(TaxTownships) Then
    fpcmbTownship.Text = "All"
    fpcmbTownship.AddItem "All"
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
  Town = QPTrim$(TaxMasterRec.Name)
  
  Opt1Desc = QPTrim$(TaxMasterRec.OptRev1)
  Opt2Desc = QPTrim$(TaxMasterRec.OptRev2)
  Opt3Desc = QPTrim$(TaxMasterRec.OptRev3)
  fpcmbIncInactive.Text = "Both"
  fpcmbIncInactive.AddItem "Both"
  fpcmbIncInactive.AddItem "Active"
  fpcmbIncInactive.AddItem "Inactive"
  
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
  
  fpcmbFlags.Text = "None"
  fpcmbFlags.AddItem "None"
  fpcmbFlags.AddItem "Tax Exempt"
  fpcmbFlags.AddItem "Charge Interest"
  fpcmbFlags.AddItem "Allow Late Notice"
  fpcmbFlags.AddItem "Bankrupt"
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxCustListRpt", "LoadMe", Erl)
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

Private Sub fpcmbDetSum_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbDetSum.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbDetSum.ListIndex = -1
  End If
  If fpcmbDetSum.ListDown <> True Then
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

Private Sub fpcmbFlags_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbFlags.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbFlags.ListIndex = -1
  End If
  If fpcmbFlags.ListDown <> True Then
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

Private Sub fpcmbIncInactive_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbIncInactive.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbIncInactive.ListIndex = -1
  End If
  If fpcmbIncInactive.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbFlags.SetFocus
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

Private Sub PrintGraphicsDet()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim RptFile$
  Dim RptHandle As Integer
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim dlm$
  Dim InactiveFlag As Boolean
  Dim x As Long, y As Long
  Dim NextRec As Long
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim PropAdd$, PropTownShip$
  Dim CustCnt As Long
  Dim Count As Boolean
  Dim CustRec As Long
  Dim CustName$
  Dim ThisTownship$
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim ActiveFlag As String * 1
  Dim PCnt As Long
  Dim FlagFlag As String * 1
  Dim FlagType$
  Dim BillType$
  
  On Error GoTo ERRORSTUFF
  
  If OptBothW.Value = True Then
    BillType$ = "Both With Property"
  ElseIf OptBothWO.Value = True Then
    BillType$ = "Both With/Without Property"
  ElseIf OptReal.Value = True Then
    BillType$ = "Real Only"
  ElseIf OptPers.Value = True Then
    BillType$ = "Personal Only"
  End If
  
  Select Case fpcmbFlags.Text
    Case "None"
      FlagFlag = "N"
      FlagType = "None"
    Case "Tax Exempt"
      FlagFlag = "X"
      FlagType = "Tax Exempt"
    Case "Charge Interest"
      FlagFlag = "I"
      FlagType = "Charge Interest"
    Case "Allow Late Notice"
      FlagFlag = "L"
      FlagType = "Allow Late Notice"
    Case "Bankrupt"
      FlagFlag = "B"
      FlagType = "Bankrupt"
    Case Else
      FlagFlag = "N"
      FlagType = "None"
  End Select
    
  ThisTownship = fpcmbTownship.Text
  IdxFlag = False
  ActiveFlag = "B"
  If QPTrim$(fpcmbIncInactive.Text) = "Active" Then
    ActiveFlag = "A"
  ElseIf QPTrim$(fpcmbIncInactive.Text) = "Inactive" Then
    ActiveFlag = "I"
  End If
  Count = False
  dlm$ = "~"
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
  End If

  RptFile$ = "TAXRPTS\CSTLSTDT.RPT"
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
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  
  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
      CustRec = IdxArray(x)
    Else
      Get TCHandle, x, TaxCust
      CustRec = x
    End If
    If TaxCust.Deleted <> 0 Then GoTo SkipIt
    Select Case FlagFlag
      Case "X"
        If TaxCust.TaxExempt = "N" Then
          GoTo SkipIt
        End If
      Case "I"
        If TaxCust.Interest = "N" Then
          GoTo SkipIt
        End If
      Case "L"
        If TaxCust.LateNotice = "N" Then
          GoTo SkipIt
        End If
      Case "B"
        If TaxCust.Bankrupt = "N" Then
          GoTo SkipIt
        End If
      Case Else
    End Select
      
    If ActiveFlag = "A" And TaxCust.Active = "N" Then
      GoTo SkipIt
    ElseIf ActiveFlag = "I" And TaxCust.Active = "Y" Then
      GoTo SkipIt
    End If
    If ThisTownship <> "No Townships Saved" Then
      If ThisTownship <> "All" Then
        If ThisTownship <> QPTrim$(UCase(TaxCust.TownShip)) Then
          GoTo SkipIt
        End If
      End If
    End If
    
    Count = False
    
    CustName = QPTrim$(TaxCust.CustName)
    If BillType$ = "Both With Property" Then
      If TaxCust.FirstPropRec = 0 And TaxCust.FirstPersRec = 0 Then GoTo SkipIt
    End If
    
    If BillType = "Personal Only" Then GoTo PersOnly
    NextRec = TaxCust.FirstPropRec
    If NextRec > 0 Then
      Do While NextRec > 0
        Get RHandle, NextRec, RealRec
        If RealRec.Deleted = True Then GoTo Deleted
        PropAdd = QPTrim$(RealRec.PropAddr)
        If PropAdd = "" Then PropAdd = "No Address Saved"
        PropTownShip = QPTrim$(RealRec.TownShip)
        If PropTownShip = "" Then PropTownShip = "No Township Saved"
        If Count = False Then
          CustCnt = CustCnt + 1
          Count = True
        End If
        TaxCust.Zip = InsertZipDash(TaxCust.Zip)
        '                   0             1                 2
        Print #RptHandle, Town; dlm; TaxCust.Acct; dlm; CustName; dlm;
        '                            3                              4
        Print #RptHandle, QPTrim$(TaxCust.Addr1); dlm; QPTrim$(TaxCust.Addr2); dlm;
        '                                                   5
        Print #RptHandle, QPTrim$(TaxCust.City) + ", " + TaxCust.State + "  " + QPTrim$(TaxCust.Zip); dlm;
        '                         6                      7
        Print #RptHandle, TaxCust.Interest; dlm; TaxCust.Bankrupt; dlm;
        '                            8                          9
        Print #RptHandle, QPTrim$(TaxCust.DrvrsLic); dlm; TaxCust.HPHONE; dlm;
        '                       10                     11                      12
        Print #RptHandle, TaxCust.WPHONE; dlm; TaxCust.LateNotice; dlm; TaxCust.TaxExempt; dlm;
        '                            13                          14                    15                     16
        Print #RptHandle, QPTrim$(RealRec.RealPin); dlm; RealRec.PROPNOT1; dlm; RealRec.PROPNOT2; dlm; RealRec.PROPNOT3; dlm;
        '                 17       18                   19
        Print #RptHandle, ""; dlm; ""; dlm; QPTrim$(RealRec.GISPOS); dlm;
        '                                                       20                                                         21                     22
'        Print #RptHandle, QPTrim$(RealRec.Map) + "/" + QPTrim(RealRec.BLOCK) + "/" + QPTrim$(RealRec.LOTNUMB); dlm; RealRec.EXMPOTHR; dlm; RealRec.EXMPSENI; dlm;
        Print #RptHandle, QPTrim$(RealRec.Map) + "/" + QPTrim(RealRec.BLOCK) + "/" + QPTrim$(RealRec.LOTNUMB); dlm; RealRec.EXMPOTHR; dlm; 0; dlm; '6/14/06
        '                         23                     24
        Print #RptHandle, RealRec.LateList; dlm; RealRec.LienDesc; dlm;
        '                            25                      26                 27             28
        Print #RptHandle, QPTrim$(RealRec.MORTCODE); dlm; PropAdd; dlm; RealRec.PROPVALU; dlm; 0; dlm;
        '                 29      30      31        32           33               34                     35
        Print #RptHandle, 0; dlm; 0; dlm; 0; dlm; "Real"; dlm; CustCnt; dlm; TaxCust.Active; dlm; RealRec.LienDesc; dlm;
        '                        36                   37
        Print #RptHandle, TaxCust.TownShip; dlm; ThisTownship; dlm;
        If UseOpt = "Y" Then
          '                            38                         39             40              41                  42                   43
          Print #RptHandle, QPTrim$(TaxCust.OptSrchDesc); dlm; ThisOpt; dlm; ActiveFlag; dlm; FlagType; dlm; RealRec.PROPVALU; dlm; RealRec.BldgVal
        Else
          '                 38       39           40             41                  42                   43
          Print #RptHandle, ""; dlm; ""; dlm; ActiveFlag; dlm; FlagType; dlm; RealRec.PROPVALU; dlm; RealRec.BldgVal
        End If
        PCnt = PCnt + 1
Deleted:
        NextRec = RealRec.NextRec
      Loop
    End If
    
    If BillType = "Real Only" Then GoTo SkipIt
PersOnly:
    NextRec = TaxCust.FirstPersRec
    Do While NextRec > 0
      Get PHandle, NextRec, PersRec
      If PersRec.Deleted = True Then GoTo NotThisOne
      If Count = False Then
        CustCnt = CustCnt + 1
        Count = True
      End If
      TaxCust.Zip = InsertZipDash(TaxCust.Zip)
      '                   0             1                 2
      Print #RptHandle, Town; dlm; TaxCust.Acct; dlm; CustName; dlm;
      '                            3                              4
      Print #RptHandle, QPTrim$(TaxCust.Addr1); dlm; QPTrim$(TaxCust.Addr2); dlm;
      '                                                   5
      Print #RptHandle, QPTrim$(TaxCust.City) + ", " + TaxCust.State + " " + QPTrim$(TaxCust.Zip); dlm;
      '                         6                      7
      Print #RptHandle, TaxCust.Interest; dlm; TaxCust.Bankrupt; dlm;
      '                            8                          9
      Print #RptHandle, QPTrim$(TaxCust.DrvrsLic); dlm; TaxCust.HPHONE; dlm;
      '                       10                     11                      12
      Print #RptHandle, TaxCust.WPHONE; dlm; TaxCust.LateNotice; dlm; TaxCust.TaxExempt; dlm;
      '                            13                       14                    15                 16
      Print #RptHandle, QPTrim$(PersRec.PropPin); dlm; PersRec.DESC1; dlm; PersRec.DESC2; dlm; PersRec.DESC3; dlm;
      '                      17                  18             19
      Print #RptHandle, PersRec.Desc4; dlm; PersRec.Desc5; dlm; ""; dlm;
      '                 20               21                     22
'      Print #RptHandle, ""; dlm; PersRec.EXMPOTHR; dlm; PersRec.EXMPSENI; dlm;
      Print #RptHandle, ""; dlm; 0; dlm; 0; dlm; '6/14/06 no more pers exemptions
      '                 23                     24
      Print #RptHandle, PersRec.LateList; dlm; ""; dlm;
      '                 25      26                 27               28                      29
      Print #RptHandle, ""; dlm; ""; dlm; PersRec.CVALUE; dlm; PersRec.MCValue; dlm; PersRec.MHValue; dlm;
      '                       30                     31                 32             33                34                   35
      Print #RptHandle, PersRec.MTValue; dlm; PersRec.PersVal; dlm; "Personal"; dlm; CustCnt; dlm; TaxCust.Active; dlm; RealRec.LienDesc; dlm;
      '                         36                  37
      Print #RptHandle, TaxCust.TownShip; dlm; ThisTownship; dlm;
      If UseOpt = "Y" Then
        '                            38                         39             40              41
        Print #RptHandle, QPTrim$(TaxCust.OptSrchDesc); dlm; ThisOpt; dlm; ActiveFlag; dlm; FlagType; dlm; 0; dlm; 0
      Else
        '                 38       39           40              41
        Print #RptHandle, ""; dlm; ""; dlm; ActiveFlag; dlm; FlagType; dlm; 0; dlm; 0
      End If
      PCnt = PCnt + 1
NotThisOne:
      NextRec = PersRec.NextRec
    Loop
'    End If
SkipIt:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      Exit Sub
    End If
  Next x
  
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  
  Close
  
  If PCnt = 0 Then
    Call TaxMsg(900, "There are no customers to report using the parameters entered.")
    Exit Sub
  End If
  
  arVATaxCustListDet.Show
  frmVATaxLoadReport.Show

  Exit Sub
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxCustListRpt", "PrintGraphicsDet", Erl)
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

Private Sub PrintTextDet()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim RptFile$
  Dim RptHandle As Integer
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim InactiveFlag As Boolean
  Dim x As Long, y As Long
  Dim NextRec As Long
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim PropAdd$, PropTownShip$
  Dim CustCnt As Long
  Dim Count As Boolean
  Dim CustRec As Long
  Dim CustName$
  Dim Page As Integer
  Dim FF$
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim ThisTownship$
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim ActiveFlag As String * 1
  Dim PCnt As Long
  Dim FlagFlag As String * 1
  Dim FlagType$
  Dim BillType$
  
  On Error GoTo ERRORSTUFF
  
  If OptBothW.Value = True Then
    BillType$ = "Both With Property"
  ElseIf OptBothWO.Value = True Then
    BillType$ = "Both With/Without Property"
  ElseIf OptReal.Value = True Then
    BillType$ = "Real Only"
  ElseIf OptPers.Value = True Then
    BillType$ = "Personal Only"
  End If
  
  Select Case fpcmbFlags.Text
    Case "None"
      FlagFlag = "N"
      FlagType = "None"
    Case "Tax Exempt"
      FlagFlag = "X"
      FlagType = "Tax Exempt"
    Case "Charge Interest"
      FlagFlag = "I"
      FlagType = "Charge Interest"
    Case "Allow Late Notice"
      FlagFlag = "L"
      FlagType = "Allow Late Notice"
    Case "Bankrupt"
      FlagFlag = "B"
      FlagType = "Bankrupt"
    Case Else
      FlagFlag = "N"
      FlagType = "None"
  End Select
    
  ThisTownship = fpcmbTownship.Text
  FF$ = Chr(12)
  MaxLines = 58
  
  IdxFlag = False
  ActiveFlag = "B"
  If QPTrim$(fpcmbIncInactive.Text) = "Active" Then
    ActiveFlag = "A"
  ElseIf QPTrim$(fpcmbIncInactive.Text) = "Inactive" Then
    ActiveFlag = "I"
  End If
  Count = False
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
  End If

  RptFile$ = "TAXRPTS\CSTLSTDT.PRN"
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
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  GoSub PrintHeader
  
  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
      CustRec = IdxArray(x)
    Else
      Get TCHandle, x, TaxCust
      CustRec = x
    End If
    If TaxCust.Deleted <> 0 Then GoTo SkipIt
    Select Case FlagFlag
      Case "X"
        If TaxCust.TaxExempt = "N" Then
          GoTo SkipIt
        End If
      Case "I"
        If TaxCust.Interest = "N" Then
          GoTo SkipIt
        End If
      Case "L"
        If TaxCust.LateNotice = "N" Then
          GoTo SkipIt
        End If
      Case "B"
        If TaxCust.Bankrupt = "N" Then
          GoTo SkipIt
        End If
      Case Else
    End Select
    If ActiveFlag = "A" And TaxCust.Active = "N" Then
      GoTo SkipIt
    ElseIf ActiveFlag = "I" And TaxCust.Active = "Y" Then
      GoTo SkipIt
    End If
    If ThisTownship <> "No Townships Saved" Then
      If ThisTownship <> "All" Then
        If ThisTownship <> QPTrim$(UCase(TaxCust.TownShip)) Then
          GoTo SkipIt
        End If
      End If
    End If
    If LineCnt > MaxLines - 10 Then
      Print #RptHandle, FF$
      GoSub PrintHeader
    End If
    CustName = QPTrim$(TaxCust.CustName)
    
    If BillType$ = "Both With Property" Then
      If TaxCust.FirstPropRec = 0 And TaxCust.FirstPersRec = 0 Then GoTo SkipIt
    End If
    
    If BillType = "Personal Only" Then
      If TaxCust.FirstPersRec > 0 Then
        GoSub PrintCustHeader
        GoTo PersOnly
      Else
        GoTo SkipIt
      End If
    ElseIf BillType = "Real Only" Then
      If TaxCust.FirstPropRec > 0 Then
        GoSub PrintCustHeader
      Else
        GoTo SkipIt
      End If
    End If
    NextRec = TaxCust.FirstPropRec
    Count = False
    If NextRec > 0 Then
      Print #RptHandle, Tab(2); "REAL PROPERTY"
      Print #RptHandle, Tab(2); String(81, "-")
      LineCnt = LineCnt + 2
    End If
    If NextRec > 0 Then
      Do While NextRec > 0
        Get RHandle, NextRec, RealRec
        If RealRec.Deleted = True Then GoTo Deleted
        PropAdd = QPTrim$(RealRec.PropAddr)
        If PropAdd = "" Then PropAdd = "No Address Saved"
        PropTownShip = QPTrim$(RealRec.TownShip)
        If PropTownShip = "" Then PropTownShip = "No Township Saved"
        If Count = False Then
          CustCnt = CustCnt + 1
          Count = True
        End If
        GoSub PrintReal
Deleted:
        NextRec = RealRec.NextRec
      Loop
    End If
    If BillType = "Real Only" Then GoTo SkipIt
PersOnly:
    NextRec = TaxCust.FirstPersRec
    If NextRec > 0 Then
      If LineCnt > MaxLines - 9 Then
        Print #RptHandle, FF$
        GoSub PrintHeader
        GoSub PrintCustHeader
      End If
      Print #RptHandle, Tab(2); "PERSONAL PROPERTY"
      Print #RptHandle, Tab(2); String(81, "-")
      Do While NextRec > 0
        Get PHandle, NextRec, PersRec
        If PersRec.Deleted = True Then GoTo SkipItDel
        If Count = False Then
          CustCnt = CustCnt + 1
          Count = True
        End If
        GoSub PrintPers
SkipItDel:
        NextRec = PersRec.NextRec
      Loop
    End If
SkipIt:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      Exit Sub
    End If
  Next x
  
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  
  If LineCnt > MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  Print #RptHandle, "Total Customers Printed: " + Using("####0", CustCnt)
  Print #RptHandle, FF$
  
  Close
  
  If PCnt = 0 Then
    Call TaxMsg(900, "There are no customers to report using the parameters entered.")
    Exit Sub
  End If
  
  ViewPrint RptFile, "Tax Master Customer List", True
  
  Exit Sub
  
PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(30); "Tax Master Customer Listing Detail"
  Print #RptHandle, Town; Tab(75); "Page #: " + CStr(Page)
  If FlagFlag <> "N" Then
    Print #RptHandle, "Flag Type: " + FlagType
  End If
  Print #RptHandle, "Report Date: " + CStr(Date); Tab(60); "Customer Status: " + fpcmbIncInactive.Text
  Print #RptHandle, String(85, "-")
  If FlagFlag <> "N" Then
    LineCnt = 5
  Else
    LineCnt = 4
  End If
  
  Return

PrintCustHeader:
  If LineCnt >= MaxLines - 8 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  TaxCust.Zip = InsertZipDash(TaxCust.Zip)
  If LineCnt > 7 Then
    Print #RptHandle,
    LineCnt = LineCnt + 1
  End If
  Print #RptHandle, "Acct #: "; Tab(10); QPTrim$(TaxCust.CustName); Tab(68); "Active Y/N:"; Tab(85); TaxCust.Active
  Print #RptHandle, Tab(2); Using$("#####0", TaxCust.Acct); Tab(10); QPTrim$(TaxCust.Addr1); Tab(68); "Interest Y/N:"; Tab(85); TaxCust.Interest
  Print #RptHandle, Tab(10); QPTrim$(TaxCust.Addr2); Tab(68); "Tax Exempt Y/N:"; Tab(85); TaxCust.TaxExempt
  Print #RptHandle, Tab(10); QPTrim$(TaxCust.City) + ", " + QPTrim$(TaxCust.State) + " " + QPTrim$(TaxCust.Zip); Tab(68); "Bankrupt Y/N:"; Tab(85); TaxCust.Bankrupt
  Print #RptHandle, Tab(10); "Home Phone:"; Tab(22); QPTrim$(TaxCust.HPHONE); Tab(39); "Work Phone:"; Tab(51); QPTrim$(TaxCust.WPHONE); Tab(68); "Late Notice Y/N:"; Tab(85); TaxCust.LateNotice
  If UseOpt = "N" Then
    Print #RptHandle, Tab(10); "Township:"; Tab(20); QPTrim$(TaxCust.TownShip)
  Else
    Print #RptHandle, Tab(10); "Township:"; Tab(20); QPTrim$(TaxCust.TownShip); Tab(40); ThisOpt + ": " + QPTrim$(TaxCust.OptSrchDesc)
  End If
  Print #RptHandle, Tab(2); String$(81, "-")
  LineCnt = LineCnt + 7
  Return

PrintReal:
  If LineCnt >= MaxLines - 9 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
    GoSub PrintCustHeader
  End If
  Print #RptHandle, Tab(5); "PIN #:"; Tab(15); QPTrim$(RealRec.RealPin); Tab(50); "Property Value:"; Tab(66); Using$("$###,###,##0.00", RealRec.PROPVALU)
  LineCnt = LineCnt + 1
  If LineCnt > MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintHeader
    GoSub PrintCustHeader
  End If
  Print #RptHandle, Tab(50); "Building Value:"; Tab(66); Using$("$###,###,##0.00", RealRec.BldgVal)
  LineCnt = LineCnt + 1
  If LineCnt > MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintHeader
    GoSub PrintCustHeader
  End If
  Print #RptHandle, Tab(5); "Property Address:"; Tab(25); QPTrim$(RealRec.PropAddr)
  LineCnt = LineCnt + 1
  If LineCnt > MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintHeader
    GoSub PrintCustHeader
  End If
  Print #RptHandle, Tab(5); "Notes:"; Tab(13); RealRec.PROPNOT1; Tab(50); "Other Exemptions:"; Tab(70); Using$("$###,##0.00", RealRec.EXMPOTHR)
  LineCnt = LineCnt + 1
  If LineCnt > MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintHeader
    GoSub PrintCustHeader
  End If
'  Print #RptHandle, Tab(13); RealRec.PROPNOT2; Tab(45); "Senior Exemption:"; Tab(62); Using$("$##,##0.00", 0) '6/14/06 RealRec.EXMPSENI)
'  LineCnt = LineCnt + 1
'  If LineCnt > MaxLines Then
'    Print #RptHandle, FF$
'    GoSub PrintHeader
'    GoSub PrintCustHeader
'  End If
  Print #RptHandle, Tab(13); RealRec.PROPNOT3; Tab(45); "Map/Block/Lot:"; Tab(62); QPTrim$(RealRec.Map) + "/" + QPTrim$(RealRec.BLOCK) + "/" + QPTrim$(RealRec.LOTNUMB)
  LineCnt = LineCnt + 1
  If LineCnt > MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintHeader
    GoSub PrintCustHeader
  End If
  Print #RptHandle, Tab(5); "Mortgage Code:"; Tab(20); QPTrim$(RealRec.MORTCODE); Tab(40); "Late Listing:"; Tab(55); RealRec.LateList; Tab(58); "Land/GIS Key:"; Tab(73); QPTrim$(RealRec.GISPOS)
  LineCnt = LineCnt + 1
  If LineCnt > MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintHeader
    GoSub PrintCustHeader
  End If
  Print #RptHandle, Tab(5); "Lien Y/N:"; Tab(15); RealRec.LienYN; Tab(20); "Lien Description:"; Tab(38); QPTrim$(RealRec.LienDesc)
  LineCnt = LineCnt + 1
  If LineCnt > MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintHeader
    GoSub PrintCustHeader
  End If
  If TaxCust.FirstPersRec = 0 Then
    Print #RptHandle, String(85, "-")
  Else
    Print #RptHandle, Tab(5); String(81, ".")
  End If
  LineCnt = LineCnt + 1
  If LineCnt > MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintHeader
    GoSub PrintCustHeader
  End If
  
  PCnt = PCnt + 1
  Return

PrintPers:
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintHeader
    GoSub PrintCustHeader
  End If
  Print #RptHandle, Tab(5); "PIN #:"; Tab(15); QPTrim$(PersRec.PropPin);
  LineCnt = LineCnt + 1
  If LineCnt > MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintHeader
    GoSub PrintCustHeader
  End If
  Print #RptHandle, Tab(5); "Personal Value:"; Tab(25); Using$("$###,###,##0.00", PersRec.PersVal); Tab(41); "Notes:"; Tab(48); PersRec.DESC1
  LineCnt = LineCnt + 1
  If LineCnt > MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintHeader
    GoSub PrintCustHeader
  End If
  Print #RptHandle, Tab(5); "Mobile Home:"; Tab(25); Using$("$###,###,##0.00", PersRec.MHValue); Tab(48); PersRec.DESC2
  LineCnt = LineCnt + 1
  If LineCnt > MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintHeader
    GoSub PrintCustHeader
  End If
  Print #RptHandle, Tab(5); "Farm Equipment:"; Tab(25); Using$("$###,###,##0.00", PersRec.CVALUE); Tab(48); PersRec.DESC3
  LineCnt = LineCnt + 1
  If LineCnt > MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintHeader
    GoSub PrintCustHeader
  End If
  Print #RptHandle, Tab(5); "Merchant Capital:"; Tab(25); Using$("$###,###,##0.00", PersRec.MCValue); Tab(48); PersRec.Desc4
  LineCnt = LineCnt + 1
  If LineCnt > MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintHeader
    GoSub PrintCustHeader
  End If
  Print #RptHandle, Tab(5); "Machine Tools:"; Tab(25); Using$("$###,###,##0.00", PersRec.MTValue); Tab(48); PersRec.Desc5
  LineCnt = LineCnt + 1
  If LineCnt > MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintHeader
    GoSub PrintCustHeader
  End If
    Print #RptHandle, Tab(5); String(81, ".")
'  Print #RptHandle, String(85, "-")
  LineCnt = LineCnt + 1
  If LineCnt > MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintHeader
    GoSub PrintCustHeader
  End If
  
  PCnt = PCnt + 1
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxCustListRpt", "PrintTextDet", Erl)
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

Private Sub PrintGraphics()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim RptFile$
  Dim RptHandle As Integer
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim dlm$
  Dim InactiveFlag As Boolean
  Dim x As Long, y As Long
  Dim NextRec As Long
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim PropAdd$, PropTownShip$
  Dim CustCnt As Long
  Dim CustRec As Long
  Dim CustName$
  Dim ThisTownship$
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim ActiveFlag As String * 1
  Dim PCnt As Long
  Dim FlagFlag As String * 1
  Dim FlagType$
  Dim BillType$
  
  On Error GoTo ERRORSTUFF
  If OptBothW.Value = True Then
    BillType$ = "Both With Property"
  ElseIf OptBothWO.Value = True Then
    BillType$ = "Both With/Without Property"
  ElseIf OptReal.Value = True Then
    BillType$ = "Real Only"
  ElseIf OptPers.Value = True Then
    BillType$ = "Personal Only"
  End If
  
  Select Case fpcmbFlags.Text
    Case "None"
      FlagFlag = "N"
      FlagType = "None"
    Case "Tax Exempt"
      FlagFlag = "X"
      FlagType = "Tax Exempt"
    Case "Charge Interest"
      FlagFlag = "I"
      FlagType = "Charge Interest"
    Case "Allow Late Notice"
      FlagFlag = "L"
      FlagType = "Allow Late Notice"
    Case "Bankrupt"
      FlagFlag = "B"
      FlagType = "Bankrupt"
    Case Else
      FlagFlag = "N"
      FlagType = "None"
  End Select
    
  ThisTownship = fpcmbTownship.Text
  
  IdxFlag = False
  ActiveFlag = "B"
  If QPTrim$(fpcmbIncInactive.Text) = "Active" Then
    ActiveFlag = "A"
  ElseIf QPTrim$(fpcmbIncInactive.Text) = "Inactive" Then
    ActiveFlag = "I"
  End If
  dlm$ = "~"
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
  End If

  RptFile$ = "TAXRPTS\CSTLSTSM.RPT"
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
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  
  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
      CustRec = IdxArray(x)
    Else
      Get TCHandle, x, TaxCust
      CustRec = x
    End If
    If TaxCust.Deleted <> 0 Then GoTo SkipIt
    Select Case FlagFlag
      Case "X"
        If TaxCust.TaxExempt = "N" Then
          GoTo SkipIt
        End If
      Case "I"
        If TaxCust.Interest = "N" Then
          GoTo SkipIt
        End If
      Case "L"
        If TaxCust.LateNotice = "N" Then
          GoTo SkipIt
        End If
      Case "B"
        If TaxCust.Bankrupt = "N" Then
          GoTo SkipIt
        End If
      Case Else
    End Select
      
    If ActiveFlag = "A" And TaxCust.Active = "N" Then
      GoTo SkipIt
    ElseIf ActiveFlag = "I" And TaxCust.Active = "Y" Then
      GoTo SkipIt
    End If
    If ThisTownship <> "No Townships Saved" Then
      If ThisTownship <> "All" Then
        If ThisTownship <> QPTrim$(UCase(TaxCust.TownShip)) Then
          GoTo SkipIt
        End If
      End If
    End If
    
    If ActiveFlag <> "A" Then
      If TaxCust.Active = "N" Then
        CustName = QPTrim$(TaxCust.CustName) + " (I)"
      Else
        CustName = QPTrim$(TaxCust.CustName)
      End If
    Else
      CustName = QPTrim$(TaxCust.CustName)
    End If

    If BillType = "Both With/Without Property" Then GoTo GotCount
    If BillType = "Personal Only" Or BillType = "Both With Property" Then
      If TaxCust.FirstPersRec > 0 Then
        NextRec = TaxCust.FirstPersRec
        Do While NextRec > 0
          Get PHandle, NextRec, PersRec
          If PersRec.Deleted = 0 Then 'got a valid prop
            NextRec = 0
            GoTo GotCount
          Else
            NextRec = PersRec.NextRec
          End If
        Loop
      End If
      If BillType = "Personal Only" Then GoTo SkipIt
    End If
    If BillType = "Real Only" Or BillType = "Both With Property" Then
      If TaxCust.FirstPropRec > 0 Then
        NextRec = TaxCust.FirstPropRec
        Do While NextRec > 0
          Get RHandle, NextRec, RealRec
          If RealRec.Deleted = 0 Then
            NextRec = 0
            GoTo GotCount
          Else
            NextRec = RealRec.NextRec
          End If
        Loop
      End If
      GoTo SkipIt
    End If
GotCount:
    CustCnt = CustCnt + 1
    TaxCust.Zip = InsertZipDash(TaxCust.Zip)
    '                   0             1                 2                      3
    Print #RptHandle, Town; dlm; TaxCust.Acct; dlm; CustName; dlm; QPTrim$(TaxCust.Addr1); dlm;
    '                                                   4
    Print #RptHandle, QPTrim$(TaxCust.City) + ", " + TaxCust.State + "  " + QPTrim$(TaxCust.Zip); dlm;
    '                             5                          6              7                  8
    Print #RptHandle, QPTrim$(TaxCust.TownShip); dlm; ThisTownship; dlm; CustCnt; dlm; fpcmbIncInactive.Text; dlm;
    '
    If UseOpt = "Y" Then
      '                    9                     10                         11               12             13
      Print #RptHandle, ThisOpt; dlm; QPTrim$(TaxCust.OptSrchDesc); dlm; ActiveFlag; dlm; FlagType; dlm; BillType
    Else
      '                  9       10           11              12            13
      Print #RptHandle, ""; dlm; ""; dlm; ActiveFlag; dlm; FlagType; dlm; BillType
    End If
    PCnt = PCnt + 1
SkipIt:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      Exit Sub
    End If
  Next x
  
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  
  Close
  
  If PCnt = 0 Then
    Call TaxMsg(900, "There are no customers to report using the parameters entered.")
    Exit Sub
  End If
  
  arVATaxCustListSum.Show
  frmVATaxLoadReport.Show
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxCustListRpt", "PrintGraphics", Erl)
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

Private Sub PrintText()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim RptFile$
  Dim RptHandle As Integer
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim InactiveFlag As Boolean
  Dim x As Long, y As Long
  Dim NextRec As Long
  Dim CustCnt As Long
  Dim CustRec As Long
  Dim CustName As String * 40
  Dim Page As Integer
  Dim FF$
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim ThisTownship$
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim ActiveFlag As String * 1
  Dim PCnt As Long
  Dim FlagFlag As String * 1
  Dim FlagType$
  Dim BillType$
  
  On Error GoTo ERRORSTUFF
  
  If OptBothW.Value = True Then
    BillType$ = "Both With Property"
  ElseIf OptBothWO.Value = True Then
    BillType$ = "Both With/Without Property"
  ElseIf OptReal.Value = True Then
    BillType$ = "Real Only"
  ElseIf OptPers.Value = True Then
    BillType$ = "Personal Only"
  End If
  
  Select Case fpcmbFlags.Text
    Case "None"
      FlagFlag = "N"
      FlagType = "None"
    Case "Tax Exempt"
      FlagFlag = "X"
      FlagType = "Tax Exempt"
    Case "Charge Interest"
      FlagFlag = "I"
      FlagType = "Charge Interest"
    Case "Allow Late Notice"
      FlagFlag = "L"
      FlagType = "Allow Late Notice"
    Case "Bankrupt"
      FlagFlag = "B"
      FlagType = "Bankrupt"
    Case Else
      FlagFlag = "N"
      FlagType = "None"
  End Select
    
  ThisTownship = fpcmbTownship.Text
  ActiveFlag = "B"
  If QPTrim$(fpcmbIncInactive.Text) = "Active" Then
    ActiveFlag = "A"
  ElseIf QPTrim$(fpcmbIncInactive.Text) = "Inactive" Then
    ActiveFlag = "I"
  End If
  FF$ = Chr(12)
  MaxLines = 58
  
  IdxFlag = False
  If fpcmbIncInactive.Text = "No" Then
    InactiveFlag = False
  Else
    InactiveFlag = True
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
  End If

  RptFile$ = "TAXRPTS\CSTLSTSM.PRN"
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
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  
  GoSub PrintHeader
  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
      CustRec = IdxArray(x)
    Else
      Get TCHandle, x, TaxCust
      CustRec = x
    End If
    If TaxCust.Deleted <> 0 Then GoTo SkipIt
    Select Case FlagFlag
      Case "X"
        If TaxCust.TaxExempt = "N" Then
          GoTo SkipIt
        End If
      Case "I"
        If TaxCust.Interest = "N" Then
          GoTo SkipIt
        End If
      Case "L"
        If TaxCust.LateNotice = "N" Then
          GoTo SkipIt
        End If
      Case "B"
        If TaxCust.Bankrupt = "N" Then
          GoTo SkipIt
        End If
      Case Else
    End Select
    If ActiveFlag = "A" And TaxCust.Active = "N" Then
      GoTo SkipIt
    ElseIf ActiveFlag = "I" And TaxCust.Active = "Y" Then
      GoTo SkipIt
    End If
    If ThisTownship <> "No Townships Saved" Then
      If ThisTownship <> "All" Then
        If ThisTownship <> QPTrim$(UCase(TaxCust.TownShip)) Then
          GoTo SkipIt
        End If
      End If
    End If
    If ActiveFlag <> "A" Then
      If TaxCust.Active = "N" Then
        CustName = QPTrim$(TaxCust.CustName) + " (I)"
      Else
        CustName = QPTrim$(TaxCust.CustName)
      End If
    Else
      CustName = QPTrim$(TaxCust.CustName)
    End If
      
    If LineCnt > MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintHeader
    End If
    
    If BillType = "Both With/Without Property" Then GoTo GotCount
    If BillType = "Personal Only" Or BillType = "Both With Property" Then
      If TaxCust.FirstPersRec > 0 Then
        NextRec = TaxCust.FirstPersRec
        Do While NextRec > 0
          Get PHandle, NextRec, PersRec
          If PersRec.Deleted = 0 Then 'got a valid prop
            NextRec = 0
            GoTo GotCount
          Else
            NextRec = PersRec.NextRec
          End If
        Loop
      End If
      If BillType = "Personal Only" Then GoTo SkipIt
    End If
    If BillType = "Real Only" Or BillType = "Both With Property" Then
      If TaxCust.FirstPropRec > 0 Then
        NextRec = TaxCust.FirstPropRec
        Do While NextRec > 0
          Get RHandle, NextRec, RealRec
          If RealRec.Deleted = 0 Then
            NextRec = 0
            GoTo GotCount
          Else
            NextRec = RealRec.NextRec
          End If
        Loop
      End If
      GoTo SkipIt
    End If
GotCount:
    CustCnt = CustCnt + 1
    TaxCust.Zip = InsertZipDash(TaxCust.Zip)
    Print #RptHandle, Using$("####0", TaxCust.Acct); Tab(8); CustName; Tab(52); QPTrim$(TaxCust.Addr1);
    Print #RptHandle, Tab(83); QPTrim$(TaxCust.City) + ", " + QPTrim$(TaxCust.State) + "  " + QPTrim$(TaxCust.Zip)
    PCnt = PCnt + 1
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintHeader
    End If
    If UseOpt = "Y" Then
      Print #RptHandle, Tab(8); ThisOpt + ": " + QPTrim$(TaxCust.OptSrchDesc)
      If LineCnt >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintHeader
      End If
    End If
      
SkipIt:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      Exit Sub
    End If
  Next x
  
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  
  If LineCnt > MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  Print #RptHandle,
  Print #RptHandle, "Total Customers Printed: " + Using("####0", CustCnt)
  Print #RptHandle, FF$
  
  Close
  If PCnt = 0 Then
    Call TaxMsg(900, "There are no customers to report using the parameters entered.")
    Exit Sub
  End If
  
  ViewPrint RptFile, "Tax Master Customer List", True
  
  Exit Sub
  
PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(30); "Tax Master Customer Listing Summary"
  Print #RptHandle, Town; Tab(75); "Page #: " + CStr(Page)
  If FlagFlag <> "N" Then
    Print #RptHandle, "Flag Type: " + FlagType
  End If
  Print #RptHandle, "Report Date: " + CStr(Date)
  If ActiveFlag = "A" Then
    Print #RptHandle, "Customer Status: " + fpcmbIncInactive.Text
  Else
    Print #RptHandle, "Customer Status: " + fpcmbIncInactive.Text; Tab(70); "(I) = Inactive"
  End If
  Print #RptHandle, "Property Type: " + BillType$
  Print #RptHandle, "Acct#"; Tab(8); "Customer Name"; Tab(52); "Address"
  Print #RptHandle, String(118, "-")
  If FlagFlag <> "N" Then
    LineCnt = 8
  Else
    LineCnt = 7
  End If
  
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxCustListRpt", "PrintText", Erl)
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
