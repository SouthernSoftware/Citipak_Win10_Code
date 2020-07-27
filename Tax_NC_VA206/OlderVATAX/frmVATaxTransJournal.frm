VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmVATaxTransJournal 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Transaction Journal"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxTransJournal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   7236
      Left            =   1560
      TabIndex        =   10
      Top             =   516
      Width           =   8508
      _Version        =   196609
      _ExtentX        =   15007
      _ExtentY        =   12763
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmVATaxTransJournal.frx":08CA
      Begin LpLib.fpCombo fpcmbIncInactive 
         Height          =   405
         Left            =   4950
         TabIndex        =   5
         Top             =   3570
         Width           =   1140
         _Version        =   196608
         _ExtentX        =   2011
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
         ColDesigner     =   "frmVATaxTransJournal.frx":08E6
      End
      Begin LpLib.fpCombo fpcmbPrintOpt 
         Height          =   405
         Left            =   2925
         TabIndex        =   7
         Top             =   4485
         Width           =   3570
         _Version        =   196608
         _ExtentX        =   6297
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
         ColDesigner     =   "frmVATaxTransJournal.frx":0BDD
      End
      Begin LpLib.fpCombo fpcmbPrintOrder 
         Height          =   405
         Left            =   2925
         TabIndex        =   8
         Top             =   4950
         Width           =   3570
         _Version        =   196608
         _ExtentX        =   6297
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
         ColDesigner     =   "frmVATaxTransJournal.frx":0ED4
      End
      Begin LpLib.fpCombo fpcmbTransType 
         Height          =   405
         Left            =   3645
         TabIndex        =   0
         Top             =   1410
         Width           =   3060
         _Version        =   196608
         _ExtentX        =   5397
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
         ColDesigner     =   "frmVATaxTransJournal.frx":11CB
      End
      Begin LpLib.fpCombo fpcmbTaxType 
         Height          =   405
         Left            =   2925
         TabIndex        =   6
         Top             =   4020
         Width           =   3570
         _Version        =   196608
         _ExtentX        =   6297
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
         ColDesigner     =   "frmVATaxTransJournal.frx":14C2
      End
      Begin LpLib.fpCombo fpcmbDetSum 
         Height          =   405
         Left            =   3525
         TabIndex        =   3
         Top             =   2640
         Width           =   2850
         _Version        =   196608
         _ExtentX        =   5027
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
         ColDesigner     =   "frmVATaxTransJournal.frx":17B9
      End
      Begin LpLib.fpCombo fpcmbOperNum 
         Height          =   405
         Left            =   4470
         TabIndex        =   4
         ToolTipText     =   "This option is only available using the summery format."
         Top             =   3120
         Width           =   1140
         _Version        =   196608
         _ExtentX        =   2011
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
         ColDesigner     =   "frmVATaxTransJournal.frx":1AB0
      End
      Begin VB.CheckBox chkQuick 
         BackColor       =   &H00D0D0D0&
         Caption         =   "Quick Search"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   3480
         TabIndex        =   24
         Top             =   6120
         Width           =   1812
      End
      Begin EditLib.fpText fptxtDesc 
         Height          =   372
         Left            =   2928
         TabIndex        =   9
         Top             =   5400
         Width           =   4524
         _Version        =   196608
         _ExtentX        =   7980
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
      Begin EditLib.fpDateTime fptxtBegDate 
         Height          =   372
         Left            =   2520
         TabIndex        =   1
         Top             =   2040
         Width           =   1740
         _Version        =   196608
         _ExtentX        =   3069
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
      Begin EditLib.fpDateTime fptxtEndDate 
         Height          =   372
         Left            =   6000
         TabIndex        =   2
         Top             =   2040
         Width           =   1740
         _Version        =   196608
         _ExtentX        =   3069
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
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   636
         Left            =   960
         TabIndex        =   22
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
         ButtonDesigner  =   "frmVATaxTransJournal.frx":1DA7
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   636
         Left            =   5832
         TabIndex        =   23
         TabStop         =   0   'False
         Tag             =   $"frmVATaxTransJournal.frx":1F85
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
         ButtonDesigner  =   "frmVATaxTransJournal.frx":2030
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdQSHelp 
         Height          =   372
         Left            =   3240
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   6600
         Width           =   2100
         _Version        =   131072
         _ExtentX        =   3704
         _ExtentY        =   656
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
         ButtonDesigner  =   "frmVATaxTransJournal.frx":220F
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Desc:"
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
         Left            =   756
         TabIndex        =   21
         Top             =   5472
         Width           =   2028
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000004&
         X1              =   400
         X2              =   8040
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000004&
         X1              =   400
         X2              =   8040
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Operator Number:"
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
         Left            =   2256
         TabIndex        =   20
         Top             =   3216
         Width           =   2076
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
         TabIndex        =   19
         Top             =   2760
         Width           =   1908
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Type:"
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
         Left            =   1560
         TabIndex        =   18
         Top             =   4128
         Width           =   1212
      End
      Begin VB.Label Label4 
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
         Height          =   300
         Left            =   4320
         TabIndex        =   17
         Top             =   2136
         Width           =   1572
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Beginning Date:"
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
         Left            =   600
         TabIndex        =   16
         Top             =   2136
         Width           =   1812
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   4740
         Left            =   408
         Top             =   1248
         Width           =   7644
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
         TabIndex        =   15
         Top             =   5040
         Width           =   1308
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Transaction Journal"
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
         Left            =   2016
         TabIndex        =   14
         Top             =   456
         Width           =   4332
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   696
         Left            =   1776
         Top             =   312
         Width           =   4908
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type:"
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
         Left            =   1440
         TabIndex        =   13
         Top             =   1524
         Width           =   2052
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
         TabIndex        =   12
         Top             =   4566
         Width           =   1500
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Include Inactive Accounts?:"
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
         Left            =   1776
         TabIndex        =   11
         Top             =   3660
         Width           =   3036
      End
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmVATaxTransJournal.frx":23F7
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
      Height          =   492
      Left            =   120
      TabIndex        =   26
      Top             =   8076
      Visible         =   0   'False
      Width           =   11400
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   7500
      Left            =   1440
      Top             =   372
      Width           =   8772
   End
End
Attribute VB_Name = "frmVATaxTransJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  'Private Temp_Class As Resize_Class
  Dim UseOpt As String * 1
  Dim ThisOpt$
  Dim Town$
  Dim Opt1Desc$
  Dim Opt2Desc$
  Dim Opt3Desc$
  Dim POpt1Desc$
  Dim POpt2Desc$
  Dim POpt3Desc$

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
      If fpcmbTaxType.Text = "Real Only" Then
        Call PrintRGraphicsDet
      ElseIf fpcmbTaxType.Text = "Personal Only" Then
        Call PrintPGraphicsDet
      End If
    Else
      If fpcmbTaxType.Text = "Real Only" Then
        frmVATaxMsg.Label1.Caption = "Pitch 12 is recommended for this printout."
        frmVATaxMsg.Label1.Top = 900
        frmVATaxMsg.Show vbModal
        Call PrintRTextDet
      ElseIf fpcmbTaxType.Text = "Personal Only" Then
        frmVATaxMsg.Label1.Caption = "Pitch 12 is recommended for this printout."
        frmVATaxMsg.Label1.Top = 900
        frmVATaxMsg.Show vbModal
        Call PrintPTextDet
      End If
    End If
  End If
End Sub

Private Sub cmdQSHelp_Click()
  Call TaxMsg(700, "Quick Search cannot guarantee all qualifying transactions will be reported. (ex. A manual transaction posted on 5/10/2006 but dated as 5/10/2002.")
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
  Me.HelpContextID = hlpTransaction
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxTransJournal.")
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
  Dim CitiPassFile As Integer, NumPassRecs As Integer
  Dim CitiPass As CitiPassType
  Dim x As Integer
  
  On Error GoTo ERRORSTUFF
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  Opt1Desc = QPTrim$(TaxMasterRec.OptRev1)
  Opt2Desc = QPTrim$(TaxMasterRec.OptRev2)
  Opt3Desc = QPTrim$(TaxMasterRec.OptRev3)
  POpt1Desc = QPTrim$(TaxMasterRec.POptRev1)
  POpt2Desc = QPTrim$(TaxMasterRec.POptRev2)
  POpt3Desc = QPTrim$(TaxMasterRec.POptRev3)
  Town$ = QPTrim$(TaxMasterRec.Name)
  fpcmbTransType.Text = " 0) All"
  fpcmbTransType.AddItem " 0) All"
  fpcmbTransType.AddItem " 1) Billing" 'include #9
  fpcmbTransType.AddItem " 2) Payment"
  fpcmbTransType.AddItem " 3) Release"
  fpcmbTransType.AddItem " 4) Interest"
  fpcmbTransType.AddItem " 5) Penalty"
  fpcmbTransType.AddItem " 6) Advertising"
  fpcmbTransType.AddItem " 7) Adjust Pay Down" 'include #10
  fpcmbTransType.AddItem " 9) Credit at Billing"
  fpcmbTransType.AddItem "11) Adjust Prepay Down"
  fpcmbTransType.AddItem "12) Refund Prepay"
  fpcmbTransType.AddItem "13) Adjust Bill Down" 'include #23
  fpcmbTransType.AddItem "14) Adjust Bill Up" 'include #24
  fpcmbTransType.AddItem "21) Payment w/Overpay" 'include #22
  fpcmbTransType.AddItem "22) Overpayment" 'include #22
  fpcmbTransType.AddItem "30) PPTRA Removal"
  
  fptxtBegDate = Date
  fptxtEndDate = Date
  fpcmbIncInactive.Text = "No"
  fpcmbIncInactive.AddItem "No"
  fpcmbIncInactive.AddItem "Yes"
  
  fpcmbTaxType.Text = "Real Only"
  fpcmbTaxType.AddItem "Real Only"
  fpcmbTaxType.AddItem "Personal Only"
  
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
  
  fpcmbOperNum.Text = "All"
  fpcmbOperNum.AddItem "All"
  OpenCitiPassFile CitiPassFile, NumPassRecs
  For x = 1 To NumPassRecs
    Get CitiPassFile, x, CitiPass
    fpcmbOperNum.AddItem CStr(CitiPass.PassNum)
  Next x
  Close
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxTransJournal", "LoadMe", Erl)
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

Private Sub fpcmbDetSum_Change()
  If fpcmbDetSum.Text = "Detail" Then
    fpcmbOperNum.Text = "All"
    fpcmbOperNum.Enabled = False
  Else
    fpcmbOperNum.Enabled = True
  End If
  
  If fpcmbTransType.Text = " 2) Payment" Then Exit Sub
  If fpcmbDetSum.Text = "Detail" And InStr(fpcmbTransType.Text, "All") Then
    Label10.Caption = "The report 'All/Detail' lists all tax bills posted within the date range entered and includes only transactions affecting each bill regardless of date."
    Label10.Visible = True
  Else
    Label10.Visible = False
  End If
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
      If fpcmbOperNum.Enabled = True Then
        fpcmbOperNum.SetFocus
      Else
        fpcmbTaxType.SetFocus
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

Private Sub fpcmbIncInactive_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbIncInactive.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbIncInactive.ListIndex = -1
  End If
  If fpcmbIncInactive.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbTaxType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbOperNum_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbOperNum.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbOperNum.ListIndex = -1
  End If
  If fpcmbOperNum.ListDown <> True Then
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
      fptxtDesc.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbTaxType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbTaxType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbTaxType.ListIndex = -1
  End If
  If fpcmbTaxType.ListDown <> True Then
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

Private Sub fpcmbTransType_Change()
  If fpcmbTransType.Text = " 2) Payment" Then
    Label10.Caption = "The 'Payments' classification and the 'Payments w/OverPay' classification must be processed separately."
    Label10.Visible = True
    Exit Sub
  Else
    Label10.Visible = False
  End If
  
  If fpcmbDetSum.Text = "Detail" And InStr(fpcmbTransType.Text, "All") Then
    Label10.Caption = "The report 'All/Detail' lists all tax bills posted within the date range entered and includes only transactions affecting each bill regardless of date."
    Label10.Visible = True
  Else
    Label10.Visible = False
  End If
End Sub

Private Sub fpcmbTransType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbTransType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbTransType.ListIndex = -1
  End If
  If fpcmbTransType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtBegDate.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub PrintGraphics()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim x As Long, y As Integer
  Dim BegDate As Integer
  Dim EndDate As Integer
  Dim dlm$
  Dim ThisRec As Long
  Dim ThisClass As Integer
  Dim ThisType As String
  Dim RptFile$
  Dim RptHandle As Integer
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim InactiveFlag As Boolean
  Dim ThisName$, ThisBillType$
  Dim TCnt As Long
  Dim TotAmt As Double
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim NumOfSrchRecs As Long
  Dim ThisTransType As String
  Dim YrCnt As Integer, ThisYear As Integer
  Dim SubRptFile$
  Dim SubRptHandle As Integer
  Dim Sub2RptFile$
  Dim Sub2RptHandle As Integer
  Dim BigYr As Integer
  Dim HoldBigYr As Integer
  Dim HoldYr As Integer
  Dim Nextx As Integer
  Dim Thisx As Integer
  Dim Nexty As Integer
  Dim Thisy As Integer
  Dim z As Integer
  Dim ThisOperNum As Integer
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim HoldPrinc As Double
  Dim HoldPers As Double
  Dim HoldMT As Double
  Dim HoldMC As Double
  Dim HoldFE As Double
  Dim HoldMH As Double
  Dim HoldInt As Double
  Dim HoldAdv As Double
  Dim HoldLateList As Double
  Dim HoldPen As Double
  Dim HoldOpt1 As Double
  Dim HoldOpt2 As Double
  Dim HoldOpt3 As Double
  Dim OptFlag As Boolean
  Dim TransDesc$
  Dim GCntTot As Long
  Dim GAmtTot As Double
  Dim BadCnt As Integer
  Dim QFlag As Boolean
  Dim GPrincTot As Double
  Dim GPersTot As Double
  Dim GMTTot As Double
  Dim GMCTot As Double
  Dim GFETot As Double
  Dim GMHTot As Double
  Dim GIntTot As Double
  Dim GAdvTot As Double
  Dim GLateListTot As Double
  Dim GPenTot As Double
  Dim GOpt1Tot As Double
  Dim GOpt2Tot As Double
  Dim GOpt3Tot As Double
  Dim ManFlag As Boolean
  Dim Disc1 As Double '1/16/2007
  Dim Disc2 As Double '1/16/2007
  Dim Disc3 As Double '1/16/2007
  Dim Disc4 As Double '1/16/2007
  Dim Disc5 As Double '1/16/2007
  Dim Disc6 As Double '1/16/2007
  Dim Disc7 As Double '1/16/2007
  Dim Disc8 As Double '1/16/2007
  Dim DiscApplied As Boolean '1/16/2007
  Dim SaveAmt As Double '1/16/2007
  
  On Error GoTo ERRORSTUFF
  
  QFlag = False
  If chkQuick.Value = 1 Then QFlag = True
  DiscApplied = False '1/16/07
  
  TransDesc$ = QPTrim$(fptxtDesc.Text)
  
  If fpcmbOperNum.Text = "All" Then
    ThisOperNum = 0
  Else
    ThisOperNum = CInt(fpcmbOperNum.Text)
  End If
  OptFlag = False
  IdxFlag = False
  If CheckB4Printing = False Then
    Exit Sub
  End If
  ThisBillType = QPTrim$(fpcmbTaxType.Text)
  If fpcmbIncInactive.Text = "No" Then
    InactiveFlag = False
  Else
    InactiveFlag = True
  End If
  
  dlm$ = "~"
  If Mid(fpcmbTransType.Text, 1, 1) = "" Then
    ThisClass = CInt(Mid(fpcmbTransType.Text, 2, 1))
  Else
    ThisClass = CInt(Mid(fpcmbTransType.Text, 1, 2))
  End If
    
  Select Case ThisClass
    Case 1
      ThisType = "Billing"
    Case 2
      ThisType = "Payment"
    Case 3
      ThisType = "Release"
    Case 4
      ThisType = "Interest"
    Case 5
      ThisType = "Penalty"
    Case 6
      ThisType = "Advertising Charge"
    Case 7, 10
      ThisType = "Adjust Pay Down"
    Case 9
      ThisType = "Credit at Billing"
    Case 11
      ThisType = "Adjust Prepay Down"
    Case 12
      ThisType = "Refund Prepay"
    Case 13
      ThisType = "Adjust Bill Down"
    Case 14, 24
      ThisType = "Adjust Bill Up"
    Case 21
      ThisType = "Payment w/Overpay"
    Case 22
      ThisType = "Overpayment Only"
    Case 30
      ThisType = "PPTRA Removal"
    Case Else
      ThisType = "All"
  End Select
    
  BegDate = Date2Num(fptxtBegDate.Text)
  EndDate = Date2Num(fptxtEndDate.Text)
  
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
    '  If IdxRec.CustRec = 2047 Then Stop
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
    OptFlag = True
  End If

  RptFile$ = "TAXRPTS\TAXJRNL.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  
  ReDim TotByYrAndType(1 To 18, 1 To 1) As Double
  ReDim CntByYrAndType(1 To 18, 1 To 1) As Integer
  ReDim TotByType(1 To 18) As Double
  ReDim TotCntByType(1 To 18) As Long
  ReDim ThEYear(1 To 1) As Integer
  
  ReDim TotByYrAndPrinc(1 To 1) As Double
  ReDim TotByYrAndPers(1 To 1) As Double
  ReDim TotByYrAndMT(1 To 1) As Double
  ReDim TotByYrAndMC(1 To 1) As Double
  ReDim TotByYrAndFE(1 To 1) As Double
  ReDim TotByYrAndMH(1 To 1) As Double
  ReDim TotByYrAndInt(1 To 1) As Double
  ReDim TotByYrAndAdv(1 To 1) As Double
  ReDim TotByYrAndLateList(1 To 1) As Double
  ReDim TotByYrAndPen(1 To 1) As Double
  ReDim TotByYrAndOpt1(1 To 1) As Double
  ReDim TotByYrAndOpt2(1 To 1) As Double
  ReDim TotByYrAndOpt3(1 To 1) As Double
  
  frmVATaxShowPctComp.Label1 = "Gathering Tax Transaction Data"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  
  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  End If
  
  For x = 1 To NumOfTCRecs
    If IdxFlag = False Then
      Get TCHandle, x, TaxCust
    Else
      Get TCHandle, IdxArray(x), TaxCust
    End If
    
    If TaxCust.Active = "N" And InactiveFlag = False Then
      GoTo SkipIt
    End If
    ThisName = QPTrim$(TaxCust.CustName)
    ThisRec = TaxCust.LastTrans
    BadCnt = 0
    Do While ThisRec > 0
      Get TTHandle, ThisRec, TaxTrans
      DiscApplied = False '1/16/07
      If QFlag = True Then
        If TaxTrans.TransDate < BegDate Then
          BadCnt = BadCnt + 1
          If BadCnt > 3 Then Exit Do
        End If
      End If
      If TransDesc <> "" Then
        If InStr(1, TaxTrans.Description, TransDesc) = 0 Then GoTo SkipIt
      End If
      If TaxTrans.TranType = 10 Then TaxTrans.TranType = 7
      If ThisOperNum <> 0 Then
        If TaxTrans.OperNum <> ThisOperNum Then GoTo SkipIt
      End If
      If TaxTrans.BillType <> "R" And TaxTrans.BillType <> "P" Then TaxTrans.BillType = ""
      If TaxTrans.BillType <> Mid(fpcmbTaxType.Text, 1, 1) And QPTrim$(TaxTrans.BillType) <> "" Then GoTo SkipIt
      If ThisClass = 7 And TaxTrans.TranType = 10 Then GoTo ItsOK
      If ThisClass = 14 And TaxTrans.TranType = 24 Then GoTo ItsOK
'      If ThisClass = 2 And TaxTrans.TranType = 21 Then GoTo ItsOK '7/6/06
      If ThisClass <> 0 And TaxTrans.TranType <> ThisClass Then GoTo SkipIt
ItsOK:
      If TaxTrans.TransDate >= BegDate And TaxTrans.TransDate <= EndDate Then
        If TaxTrans.BelongTo > 0 Then
          Get TTHandle, TaxTrans.BelongTo, TaxTrans
        End If
        
        If YrCnt = 0 Then
           YrCnt = YrCnt + 1
           ThisYear = YrCnt
           ReDim Preserve ThEYear(1 To YrCnt) As Integer
           ThEYear(YrCnt) = TaxTrans.TaxYear
           ReDim Preserve TotByYrAndType(1 To 18, 1 To YrCnt) As Double
           ReDim Preserve CntByYrAndType(1 To 18, 1 To YrCnt) As Integer
           ReDim Preserve TotByYrAndPrinc(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndPers(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndMT(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndMC(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndFE(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndMH(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndInt(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndAdv(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndLateList(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndPen(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndOpt1(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndOpt2(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndOpt3(1 To YrCnt) As Double
            
           TotByYrAndPrinc(YrCnt) = 0
           TotByYrAndPers(YrCnt) = 0
           TotByYrAndMT(YrCnt) = 0
           TotByYrAndMC(YrCnt) = 0
           TotByYrAndFE(YrCnt) = 0
           TotByYrAndMH(YrCnt) = 0
           TotByYrAndInt(YrCnt) = 0
           TotByYrAndAdv(YrCnt) = 0
           TotByYrAndLateList(YrCnt) = 0
           TotByYrAndPen(YrCnt) = 0
           TotByYrAndOpt1(YrCnt) = 0
           TotByYrAndOpt2(YrCnt) = 0
           TotByYrAndOpt3(YrCnt) = 0
           For y = 1 To 18
             TotByYrAndType(y, YrCnt) = 0
             CntByYrAndType(y, YrCnt) = 0
           Next y
         Else
           For y = 1 To YrCnt
             If TaxTrans.TaxYear = ThEYear(y) Then
               ThisYear = y
               Exit For
             End If
           Next y
           If y > YrCnt Then
             YrCnt = YrCnt + 1
             ThisYear = YrCnt
             ReDim Preserve ThEYear(1 To YrCnt) As Integer
             ThEYear(YrCnt) = TaxTrans.TaxYear
             ReDim Preserve TotByYrAndType(1 To 18, 1 To YrCnt) As Double
             ReDim Preserve CntByYrAndType(1 To 18, 1 To YrCnt) As Integer
             ReDim Preserve TotByYrAndPrinc(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndPers(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndMT(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndMC(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndFE(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndMH(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndInt(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndAdv(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndLateList(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndPen(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndOpt1(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndOpt2(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndOpt3(1 To YrCnt) As Double
             TotByYrAndPrinc(YrCnt) = 0
             TotByYrAndPers(YrCnt) = 0
             TotByYrAndMT(YrCnt) = 0
             TotByYrAndMC(YrCnt) = 0
             TotByYrAndFE(YrCnt) = 0
             TotByYrAndMH(YrCnt) = 0
             TotByYrAndInt(YrCnt) = 0
             TotByYrAndAdv(YrCnt) = 0
             TotByYrAndLateList(YrCnt) = 0
             TotByYrAndPen(YrCnt) = 0
             TotByYrAndOpt1(YrCnt) = 0
             TotByYrAndOpt2(YrCnt) = 0
             TotByYrAndOpt3(YrCnt) = 0
             For y = 1 To 18
               TotByYrAndType(y, YrCnt) = 0
               CntByYrAndType(y, YrCnt) = 0
             Next y
           End If
         End If
         Get TTHandle, ThisRec, TaxTrans
'         If TaxTrans.CustomerRec = 653 Then Stop
         Select Case TaxTrans.TranType
           Case 1
            ThisTransType = "Billing"
            TotByYrAndType(1, ThisYear) = OldRound(TotByYrAndType(1, ThisYear) + TaxTrans.Amount)
            TotByType(1) = OldRound(TotByType(1) + TaxTrans.Amount)
            CntByYrAndType(1, ThisYear) = OldRound(CntByYrAndType(1, ThisYear) + 1)
            TotCntByType(1) = OldRound(TotCntByType(1) + 1)
            TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1)
            TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1)
            TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2)
            TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3)
            TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4)
            TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5)
            '7/11/06 added back interest, advertising and penalty to accommodate manual bills
            TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.Interest)
            TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.Collection)
            TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.Penalty)
            
            TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateList)
            TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1)
            TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2)
            TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3)
          Case 2
            ThisTransType = "Payment"
            If TaxTrans.DiscAmt > 0 Then 'added 1/16/07
              If TaxTrans.BillType = "R" Then
                GoSub ApplyDiscR
              ElseIf TaxTrans.BillType = "P" Then
                GoSub ApplyDiscP
              End If
            End If
            TotByYrAndType(2, ThisYear) = OldRound(TotByYrAndType(2, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(2, ThisYear) = OldRound(CntByYrAndType(2, ThisYear) + 1)
            TotByType(2) = OldRound(TotByType(2) + TaxTrans.Amount)
            TotCntByType(2) = OldRound(TotCntByType(2) + 1)
            TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1Pd)
            TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1Pd)
            TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2Pd)
            TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3Pd)
            TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4Pd)
            TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5Pd)
            TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
            TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
            TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
            TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
            TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
            TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
            TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
          Case 3
            '7/12/06 changed revenues for Release to paid from charged
            ThisTransType = "Release"
            TotByYrAndType(3, ThisYear) = OldRound(TotByYrAndType(3, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(3, ThisYear) = OldRound(CntByYrAndType(3, ThisYear) + 1)
            TotByType(3) = OldRound(TotByType(3) + TaxTrans.Amount)
            TotCntByType(3) = OldRound(TotCntByType(3) + 1)
            TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1Pd)
            TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1Pd)
            TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2Pd)
            TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3Pd)
            TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4Pd)
            TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5Pd)
            TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
            TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
            TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
            TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
            TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
            TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
            TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
          Case 4
            ThisTransType = "Interest"
            TotByYrAndType(4, ThisYear) = OldRound(TotByYrAndType(4, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(4, ThisYear) = OldRound(CntByYrAndType(4, ThisYear) + 1)
            TotByType(4) = OldRound(TotByType(4) + TaxTrans.Amount)
            TotCntByType(4) = OldRound(TotCntByType(4) + 1)
          Case 5
            ThisTransType = "Penalty"
            TotByYrAndType(5, ThisYear) = OldRound(TotByYrAndType(5, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(5, ThisYear) = OldRound(CntByYrAndType(5, ThisYear) + 1)
            TotByType(5) = OldRound(TotByType(5) + TaxTrans.Amount)
            TotCntByType(5) = OldRound(TotCntByType(5) + 1)
          Case 6
            ThisTransType = "Advertising Charge"
            TotByYrAndType(6, ThisYear) = OldRound(TotByYrAndType(6, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(6, ThisYear) = OldRound(CntByYrAndType(6, ThisYear) + 1)
            TotByType(6) = OldRound(TotByType(6) + TaxTrans.Amount)
            TotCntByType(6) = OldRound(TotCntByType(6) + 1)
          Case 7
            ThisTransType = "Adjust Pay Down"
            TotByYrAndType(7, ThisYear) = OldRound(TotByYrAndType(7, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(7, ThisYear) = OldRound(CntByYrAndType(7, ThisYear) + 1)
            TotByType(7) = OldRound(TotByType(7) + TaxTrans.Amount)
            TotCntByType(7) = OldRound(TotCntByType(7) + 1)
            TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1Pd)
            TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1Pd)
            TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2Pd)
            TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3Pd)
            TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4Pd)
            TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5Pd)
            TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
            TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
            TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
            TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
            TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
            TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
            TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
          Case 9
            ThisTransType = "Credit Applied at Billing"
            TotByYrAndType(8, ThisYear) = OldRound(TotByYrAndType(8, ThisYear) + TaxTrans.Revenue.PrePaidUsed)
            CntByYrAndType(8, ThisYear) = OldRound(CntByYrAndType(8, ThisYear) + 1)
            TotByType(8) = OldRound(TotByType(8) + TaxTrans.Revenue.PrePaidUsed)
            TotCntByType(8) = OldRound(TotCntByType(8) + 1)
          Case 13
            ThisTransType = "Adjust Bill Down"
            TotByYrAndType(9, ThisYear) = OldRound(TotByYrAndType(9, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(9, ThisYear) = OldRound(CntByYrAndType(9, ThisYear) + 1)
            TotByType(9) = OldRound(TotByType(9) + TaxTrans.Amount)
            TotCntByType(9) = OldRound(TotCntByType(9) + 1)
            TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1)
            TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1)
            TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2)
            TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3)
            TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4)
            TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5)
            TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.Interest)
            TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.Collection)
            TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateList)
            TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.Penalty)
            TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1)
            TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2)
            TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3)
          Case 14
            ThisTransType = "Adjust Bill Up"
            TotByYrAndType(10, ThisYear) = OldRound(TotByYrAndType(10, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(10, ThisYear) = OldRound(CntByYrAndType(10, ThisYear) + 1)
            TotByType(10) = OldRound(TotByType(10) + TaxTrans.Amount)
            TotCntByType(10) = OldRound(TotCntByType(10) + 1)
            TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1)
            TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1)
            TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2)
            TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3)
            TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4)
            TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5)
            TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.Interest)
            TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.Collection)
            TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateList)
            TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.Penalty)
            TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1)
            TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2)
            TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3)
          Case 21
            ThisTransType = "Bill Pay\Overpay" '7/6/06 added revenues and changed Amount to PrePaidAmt
            If TaxTrans.DiscAmt > 0 Then 'added 1/16/07
              If TaxTrans.BillType = "R" Then
                GoSub ApplyDiscR
              ElseIf TaxTrans.BillType = "P" Then
                GoSub ApplyDiscP
              End If
            End If
            If fpcmbTransType.Text <> " 0) All" Then 'added the All if statement on 7/7/06
              TotByYrAndType(11, ThisYear) = OldRound(TotByYrAndType(11, ThisYear) + TaxTrans.Revenue.PrePaidAmt) '  .Amount)
            Else
              TotByYrAndType(11, ThisYear) = OldRound(TotByYrAndType(11, ThisYear) + TaxTrans.Amount)
            End If
'            TotByYrAndType(11, ThisYear) = OldRound(TotByYrAndType(11, ThisYear) + TaxTrans.Revenue.PrePaidAmt) '.Amount)
            CntByYrAndType(11, ThisYear) = OldRound(CntByYrAndType(11, ThisYear) + 1)
            TotByType(11) = OldRound(TotByType(11) + TaxTrans.Amount)
            TotCntByType(11) = OldRound(TotCntByType(11) + 1)
            TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1Pd)
            TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1Pd)
            TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2Pd)
            TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3Pd)
            TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4Pd)
            TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5Pd)
            TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
            TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
            TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
            TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
            TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
            TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
            TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
          Case 22
            ThisTransType = "Overpayment"
            TotByYrAndType(12, ThisYear) = OldRound(TotByYrAndType(12, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(12, ThisYear) = OldRound(CntByYrAndType(12, ThisYear) + 1)
            TotByType(12) = OldRound(TotByType(12) + TaxTrans.Amount)
            TotCntByType(12) = OldRound(TotCntByType(12) + 1)
          Case 24
            ThisTransType = "Adjust Bill Up Affecting Credit Balance"
            TotByYrAndType(13, ThisYear) = OldRound(TotByYrAndType(13, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(13, ThisYear) = OldRound(CntByYrAndType(13, ThisYear) + 1)
            TotByType(13) = OldRound(TotByType(13) + TaxTrans.Amount)
            TotCntByType(13) = OldRound(TotCntByType(13) + 1)
            TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1Pd)
            TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1Pd)
            TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2Pd)
            TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3Pd)
            TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4Pd)
            TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5Pd)
            TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
            TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
            TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
            TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
            TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
            TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
            TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)

'            TotByYrAndType(13, ThisYear) = OldRound(TotByYrAndType(13, ThisYear) + TaxTrans.Amount)
'            CntByYrAndType(13, ThisYear) = OldRound(CntByYrAndType(13, ThisYear) + 1)
'            TotByType(13) = OldRound(TotByType(13) + TaxTrans.Amount)
'            TotCntByType(13) = OldRound(TotCntByType(13) + 1)
'            TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1)
'            TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1)
'            TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2)
'            TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3)
'            TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4)
'            TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5)
'            TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.Interest)
'            TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.Collection)
'            TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateList)
'            TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.Penalty)
'            TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1)
'            TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2)
'            TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3)
          Case 10  '7/11/06 added Pd on revenues
            
            ThisTransType = "Adjust Pay Down Affecting Credit Balance"
            TotByYrAndType(14, ThisYear) = OldRound(TotByYrAndType(14, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(14, ThisYear) = OldRound(CntByYrAndType(14, ThisYear) + 1)
            TotByType(14) = OldRound(TotByType(14) + TaxTrans.Amount)
            TotCntByType(14) = OldRound(TotCntByType(14) + 1)
            TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1Pd)
            TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1Pd)
            TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2Pd)
            TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3Pd)
            TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4Pd)
            TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5Pd)
            TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
            TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
            TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
            TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
            TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
            TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
            TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
          Case 11
            ThisTransType = "Adjust Prepay Down"
            TotByYrAndType(15, ThisYear) = OldRound(TotByYrAndType(15, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(15, ThisYear) = OldRound(CntByYrAndType(15, ThisYear) + 1)
            TotByType(15) = OldRound(TotByType(15) + TaxTrans.Amount)
            TotCntByType(15) = OldRound(TotCntByType(15) + 1)
          Case 12
            ThisTransType = "Refund Prepay"
            TotByYrAndType(16, ThisYear) = OldRound(TotByYrAndType(16, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(16, ThisYear) = OldRound(CntByYrAndType(16, ThisYear) + 1)
            TotByType(16) = OldRound(TotByType(16) + TaxTrans.Amount)
            TotCntByType(16) = OldRound(TotCntByType(16) + 1)
          Case 30
            ThisTransType = "PPTRA Removal"
            TotByYrAndType(17, ThisYear) = OldRound(TotByYrAndType(17, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(17, ThisYear) = OldRound(CntByYrAndType(17, ThisYear) + 1)
            TotByType(17) = OldRound(TotByType(17) + TaxTrans.Amount)
            TotCntByType(17) = OldRound(TotCntByType(17) + 1)
          Case Else
            ThisTransType = "Unknown"
            TotByYrAndType(18, ThisYear) = OldRound(TotByYrAndType(18, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(18, ThisYear) = OldRound(CntByYrAndType(18, ThisYear) + 1)
            TotByType(18) = OldRound(TotByType(18) + TaxTrans.Amount)
            TotCntByType(18) = OldRound(TotCntByType(18) + 1)
        End Select
        TCnt = TCnt + 1
        If TaxTrans.TranType = 2 Or TaxTrans.TranType = 21 Then 'added 1/16/07
          If TaxTrans.DiscAmt > 0 Then
            If TaxTrans.BillType = "R" Then
              GoSub ApplyDiscR
            ElseIf TaxTrans.BillType = "P" Then
              GoSub ApplyDiscP
            End If
          End If
        End If
        TotAmt = OldRound(TotAmt + TaxTrans.Amount)
        '                   0            1                 2                   3
        Print #RptHandle, Town$; dlm; ThisName; dlm; TaxCust.Acct; dlm; TaxCust.Active; dlm;
        '                                 4                           5                6
        Print #RptHandle, MakeRegDate(TaxTrans.TransDate); dlm; ThisBillType; dlm; ThisType; dlm;
        If TaxTrans.BelongTo > 0 Then
          Get TTHandle, TaxTrans.BelongTo, TaxTrans
            '                          7                         8                          9
            Print #RptHandle, MakeRegDate(BegDate); dlm; MakeRegDate(EndDate); dlm; TaxTrans.TaxYear; dlm;
            Get TTHandle, ThisRec, TaxTrans
            If DiscApplied = True Then 'added 1/16/07
              If TaxTrans.BillType = "R" Then
                GoSub ApplyDiscR
              ElseIf TaxTrans.BillType = "P" Then
                GoSub ApplyDiscP
              End If
            End If
        Else
          '                          7                         8                          9
          Print #RptHandle, MakeRegDate(BegDate); dlm; MakeRegDate(EndDate); dlm; TaxTrans.TaxYear; dlm;
        End If
        If TaxTrans.TranType <> 9 Then
          '                      10                11          12                       13
          Print #RptHandle, TaxTrans.Amount; dlm; TCnt; dlm; TotAmt; dlm; TaxTrans.Revenue.PrePaidAmt; dlm;
        Else
          '                      10                              11          12                       13
          Print #RptHandle, TaxTrans.Revenue.PrePaidUsed; dlm; TCnt; dlm; TotAmt; dlm; TaxTrans.Revenue.PrePaidAmt; dlm;
        End If
        If TaxTrans.BelongTo > 0 Then
          Get TTHandle, TaxTrans.BelongTo, TaxTrans
          '                             14
          Print #RptHandle, ParseBillNum(TaxTrans.Description); dlm;
          Get TTHandle, ThisRec, TaxTrans
        Else
          '                 14
          Print #RptHandle, 0; dlm;
        End If
        If ThisOperNum = 0 Then
          '                                15                        16               17
          Print #RptHandle, QPTrim$(TaxTrans.Description); dlm; ThisTransType; dlm; "All"; dlm;
        Else
          '                                15                        16                   17
          Print #RptHandle, QPTrim$(TaxTrans.Description); dlm; ThisTransType; dlm; ThisOperNum; dlm;
        End If
        If UseOpt = "Y" Then
          '                   18                      19                           20
          Print #RptHandle, ThisOpt; dlm; QPTrim$(TaxCust.OptSrchDesc); dlm; TaxTrans.OperNum
        Else
          '                 18       19             20
          Print #RptHandle, ""; dlm; ""; dlm; TaxTrans.OperNum
        End If
      End If
      
SkipIt:
      ThisRec = TaxTrans.LastTrans
    Loop
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
MoveOn:
  Next x
  
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  
  Close
  If TCnt = 0 Then
    Call TaxMsg(900, "No transactions were found that fit the parameters entered.")
    Close
    Exit Sub
  End If
  
  If YrCnt > 0 And fpcmbTaxType.Text = "Real Only" Then
    GoSub PrintSub
    GoSub PrintTotals
  ElseIf YrCnt > 0 And fpcmbTaxType.Text = "Personal Only" Then
    GoSub PrintSubPers
    GoSub PrintTotalsPers
  End If
  
  arVATaxTransJournal.Show
  
  Exit Sub
  
PrintSub:
  SubRptFile$ = "TAXRPTS\SUBTAXJRNL.RPT"
  SubRptHandle = FreeFile
  Open SubRptFile For Output As #SubRptHandle
  BigYr = 0
  For x = 1 To YrCnt
    If ThEYear(x) > BigYr Then
      BigYr = ThEYear(x)
    End If
  Next x
  ReDim HoldAmt(1 To 18, 1 To YrCnt) As Double
  ReDim HoldCnt(1 To 18, 1 To YrCnt) As Integer
  Nexty = 1
  Nextx = 1
  HoldBigYr = 0
    For x = 1 To 18
      For y = Nexty To YrCnt
        If ThEYear(y) >= HoldBigYr Then
          HoldBigYr = ThEYear(y)
          Thisx = x
          Thisy = y
        End If
      Next y
      For z = 1 To 18
        HoldAmt(z, Thisy) = TotByYrAndType(z, Nexty)
        HoldCnt(z, Thisy) = CntByYrAndType(z, Nexty)
      Next z
      HoldYr = ThEYear(Nexty)
      HoldPrinc = TotByYrAndPrinc(Nexty)
      HoldInt = TotByYrAndInt(Nexty)
      HoldAdv = TotByYrAndAdv(Nexty)
      HoldLateList = TotByYrAndLateList(Nexty)
      HoldPen = TotByYrAndPen(Nexty)
      HoldOpt1 = TotByYrAndOpt1(Nexty)
      HoldOpt2 = TotByYrAndOpt2(Nexty)
      HoldOpt3 = TotByYrAndOpt3(Nexty)
      For z = 1 To 18
        TotByYrAndType(z, Nexty) = TotByYrAndType(z, Thisy)
        CntByYrAndType(z, Nexty) = CntByYrAndType(z, Thisy)
      Next z
      ThEYear(Nexty) = ThEYear(Thisy)
      TotByYrAndPrinc(Nexty) = TotByYrAndPrinc(Thisy)
      TotByYrAndInt(Nexty) = TotByYrAndInt(Thisy)
      TotByYrAndAdv(Nexty) = TotByYrAndAdv(Thisy)
      TotByYrAndLateList(Nexty) = TotByYrAndLateList(Thisy)
      TotByYrAndPen(Nexty) = TotByYrAndPen(Thisy)
      TotByYrAndOpt1(Nexty) = TotByYrAndOpt1(Thisy)
      TotByYrAndOpt2(Nexty) = TotByYrAndOpt2(Thisy)
      TotByYrAndOpt3(Nexty) = TotByYrAndOpt3(Thisy)
      For z = 1 To 18
        TotByYrAndType(z, Thisy) = HoldAmt(z, Thisy)
        CntByYrAndType(z, Thisy) = HoldCnt(z, Thisy)
      Next z
      ThEYear(Thisy) = HoldYr
      TotByYrAndPrinc(Thisy) = HoldPrinc
      TotByYrAndInt(Thisy) = HoldInt
      TotByYrAndAdv(Thisy) = HoldAdv
      TotByYrAndLateList(Thisy) = HoldLateList
      TotByYrAndPen(Thisy) = HoldPen
      TotByYrAndOpt1(Thisy) = HoldOpt1
      TotByYrAndOpt2(Thisy) = HoldOpt2
      TotByYrAndOpt3(Thisy) = HoldOpt3
      If Nexty >= YrCnt Then Exit For
      HoldBigYr = 0 'BigYr + 1
      Nexty = Nexty + 1
    Next x
  
  For y = 1 To YrCnt
    For x = 1 To 18
      If TotByYrAndType(x, y) > 0 Then
        Select Case x
          Case 1
            Print #SubRptHandle, "Billing"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
              Print #SubRptHandle, TotByYrAndPrinc(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10                 11         12             13             14               15
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm; TotByYrAndPen(y)
            Else
              '                    4       5       6       7       8       9       10      11      12        13       14       15
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 2
            Print #SubRptHandle, "Payment"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
              Print #SubRptHandle, TotByYrAndPrinc(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm; TotByYrAndPen(y)
            Else
              '                    4       5       6       7       8       9       10      11      12        13       14       15
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 3
            Print #SubRptHandle, "Release"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
              Print #SubRptHandle, TotByYrAndPrinc(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm; TotByYrAndPen(y)
            Else
              '                    4       5       6       7       8       9       10      11      12        13       14       15
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 4
            Print #SubRptHandle, "Interest"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 5
            Print #SubRptHandle, "Penalty"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 6
            Print #SubRptHandle, "Advertising"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 7
            Print #SubRptHandle, "Adjust Pay Down"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
              Print #SubRptHandle, TotByYrAndPrinc(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm; TotByYrAndPen(y)
            Else
              '                    4       5       6       7       8       9       10      11      12        13       14       15
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 8
            Print #SubRptHandle, "Credit at Billing"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 9
            Print #SubRptHandle, "Adjust Bill Down"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
              Print #SubRptHandle, TotByYrAndPrinc(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm; TotByYrAndPen(y)
            Else
              '                    4       5       6       7       8       9       10      11      12        13       14       15
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 10
            Print #SubRptHandle, "Adjust Bill Up"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
              Print #SubRptHandle, TotByYrAndPrinc(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm; TotByYrAndPen(y)
            Else
              '                    4       5       6       7       8       9       10      11      12        13       14       15
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 11
            If fpcmbTransType.Text <> " 0) All" Then
              Print #SubRptHandle, "Bill OverPay"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
              '                            4                      5                      6                         7
              Print #SubRptHandle, TotByYrAndPrinc(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm; TotByYrAndPen(y)
            Else
              Print #SubRptHandle, "Bill Pay/OverPay"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
              '                    4       5       6       7       8       9       10      11      12        13       14       15
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 12
            Print #SubRptHandle, "OverPayment"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 13
            Print #SubRptHandle, "Adjust Bill Up Affecting Credit Balance"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
              Print #SubRptHandle, TotByYrAndPrinc(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm; TotByYrAndPen(y)
            Else
              '                    4       5       6       7       8       9       10      11      12        13       14       15
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 14
            Print #SubRptHandle, "Adjust Pay Dwn Affecting Credit Balance"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
              Print #SubRptHandle, TotByYrAndPrinc(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm; TotByYrAndPen(y)
            Else
              '                    4       5       6       7       8       9       10      11      12        13       14       15
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 15
            Print #SubRptHandle, "Adjust Prepay Down"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 16
            Print #SubRptHandle, "Refund Prepay"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 17
            Print #SubRptHandle, "PPTRA Removal"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 18
            Print #SubRptHandle, "Unknown"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
        End Select
      End If
    Next x
  Next y
  Close SubRptHandle
  
  Return
  
PrintTotals:
  Sub2RptFile$ = "TAXRPTS\SUB2TAXJRNL.RPT"
  Sub2RptHandle = FreeFile
  Open Sub2RptFile For Output As #Sub2RptHandle
  GCntTot = 0
  GAmtTot = 0
  
  For x = 1 To YrCnt
    GPrincTot = GPrincTot + TotByYrAndPrinc(x)     'Total is allways zero
    GIntTot = GIntTot + TotByYrAndInt(x)
    GAdvTot = GAdvTot + TotByYrAndAdv(x)
    GLateListTot = GLateListTot + TotByYrAndLateList(x)
    GPenTot = GPenTot + TotByYrAndPen(x)
    GOpt1Tot = GOpt1Tot + TotByYrAndOpt1(x)
    GOpt2Tot = GOpt2Tot + TotByYrAndOpt2(x)
    GOpt3Tot = GOpt3Tot + TotByYrAndOpt3(x)
  Next x
  Dim CaseR11Tot As Double 'added 7/6/06
  For x = 1 To 18
    GCntTot = GCntTot + TotCntByType(x)
    GAmtTot = GAmtTot + TotByType(x)
    If TotByType(x) > 0 Then
      If x <> 11 Or (x = 11 And fpcmbTransType.Text = " 0) All") Then 'added the Or part on 7/7/06
        Print #Sub2RptHandle, TotByType(x); dlm; TotCntByType(x); dlm; GAmtTot; dlm; GCntTot; dlm;
      Else 'inserted on 7/6/06
        CaseR11Tot = OldRound(TotByType(11) - (GPrincTot + GIntTot + GAdvTot + GLateListTot + GOpt1Tot + GOpt2Tot + GOpt3Tot + GPenTot))
        Print #Sub2RptHandle, CaseR11Tot; dlm; TotCntByType(x); dlm; GAmtTot; dlm; GCntTot; dlm;
      End If
      '                         4              5             6               7
      Print #Sub2RptHandle, GPrincTot; dlm; GIntTot; dlm; GAdvTot; dlm; GLateListTot; dlm;
      '                        8              9             10              11             12             13
      Print #Sub2RptHandle, GOpt1Tot; dlm; GOpt2Tot; dlm; GOpt3Tot; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm;
      
      If fpcmbTransType.Text <> " 0) All" Then
        '                      14
        Print #Sub2RptHandle, "1"; dlm;
      Else
        '                      14
        Print #Sub2RptHandle, "2"; dlm;
      End If
      '                        15
      Print #Sub2RptHandle, GPenTot; dlm;
      Select Case x
        Case 1 '                 16          17
          Print #Sub2RptHandle, "Billing"; dlm; 3
        Case 2
          Print #Sub2RptHandle, "Payment"; dlm; 3
        Case 3
          Print #Sub2RptHandle, "Release"; dlm; 3
        Case 4
          Print #Sub2RptHandle, "Interest"; dlm; 44
        Case 5
          Print #Sub2RptHandle, "Penalty"; dlm; 44
        Case 6
          Print #Sub2RptHandle, "Advertising"; dlm; 44
        Case 7
          Print #Sub2RptHandle, "Adjust Pay Down"; dlm; 3
        Case 8
          Print #Sub2RptHandle, "Credit at Billing"; dlm; 44
        Case 9
          Print #Sub2RptHandle, "Adjust Bill Down"; dlm; 3
        Case 10
          Print #Sub2RptHandle, "Adjust Bill Up"; dlm; 3
        Case 11
          If fpcmbTransType.Text <> " 0) All" Then
            Print #Sub2RptHandle, "Bill OverPay"; dlm; 3 'changed from 44 to 3 on 7/6/06
          Else
            Print #Sub2RptHandle, "Bill Pay/OverPay"; dlm; 3 'changed from 44 to 3 on 7/6/06
          End If
        Case 12
          Print #Sub2RptHandle, "OverPayment"; dlm; 44
        Case 13
          Print #Sub2RptHandle, "Adjust Bill Up Affecting Credit Balance"; dlm; 3
        Case 14
          Print #Sub2RptHandle, "Adjust Pay Dwn Affecting Credit Balance"; dlm; 3
        Case 15
          Print #Sub2RptHandle, "Adjust Prepay Down"; dlm; 44
        Case 16
          Print #Sub2RptHandle, "Refund Prepay"; dlm; 44
        Case 17
          Print #Sub2RptHandle, "PPTRA Removal"; dlm; 44
        Case 18
          Print #Sub2RptHandle, "Unknown"; dlm; 44
      End Select
    End If
  Next x
  Close Sub2RptHandle
  Return
  
PrintSubPers:
  SubRptFile$ = "TAXRPTS\SUBTAXJRNLP.RPT"
  SubRptHandle = FreeFile
  Open SubRptFile For Output As #SubRptHandle
  BigYr = 0
  For x = 1 To YrCnt
    If ThEYear(x) > BigYr Then
      BigYr = ThEYear(x)
    End If
  Next x
  ReDim HoldAmt(1 To 18, 1 To YrCnt) As Double
  ReDim HoldCnt(1 To 18, 1 To YrCnt) As Integer
  Nexty = 1
  Nextx = 1
  HoldBigYr = 0
    For x = 1 To 18
      For y = Nexty To YrCnt
        If ThEYear(y) >= HoldBigYr Then
          HoldBigYr = ThEYear(y)
          Thisx = x
          Thisy = y
        End If
      Next y
      For z = 1 To 18
        HoldAmt(z, Thisy) = TotByYrAndType(z, Nexty)
        HoldCnt(z, Thisy) = CntByYrAndType(z, Nexty)
      Next z
      HoldYr = ThEYear(Nexty)
      HoldPers = TotByYrAndPers(Nexty)
      HoldMT = TotByYrAndMT(Nexty)
      HoldMC = TotByYrAndMC(Nexty)
      HoldFE = TotByYrAndFE(Nexty)
      HoldMH = TotByYrAndMH(Nexty)
      HoldInt = TotByYrAndInt(Nexty)
      HoldAdv = TotByYrAndAdv(Nexty)
      HoldLateList = TotByYrAndLateList(Nexty)
      HoldPen = TotByYrAndPen(Nexty)
      HoldOpt1 = TotByYrAndOpt1(Nexty)
      HoldOpt2 = TotByYrAndOpt2(Nexty)
      HoldOpt3 = TotByYrAndOpt3(Nexty)
      For z = 1 To 18
        TotByYrAndType(z, Nexty) = TotByYrAndType(z, Thisy)
        CntByYrAndType(z, Nexty) = CntByYrAndType(z, Thisy)
      Next z
      ThEYear(Nexty) = ThEYear(Thisy)
      TotByYrAndPers(Nexty) = TotByYrAndPers(Thisy)
      TotByYrAndMT(Nexty) = TotByYrAndMT(Thisy)
      TotByYrAndMC(Nexty) = TotByYrAndMC(Thisy)
      TotByYrAndFE(Nexty) = TotByYrAndFE(Thisy)
      TotByYrAndMH(Nexty) = TotByYrAndMH(Thisy)
      TotByYrAndInt(Nexty) = TotByYrAndInt(Thisy)
      TotByYrAndAdv(Nexty) = TotByYrAndAdv(Thisy)
      TotByYrAndLateList(Nexty) = TotByYrAndLateList(Thisy)
      TotByYrAndPen(Nexty) = TotByYrAndPen(Thisy)
      TotByYrAndOpt1(Nexty) = TotByYrAndOpt1(Thisy)
      TotByYrAndOpt2(Nexty) = TotByYrAndOpt2(Thisy)
      TotByYrAndOpt3(Nexty) = TotByYrAndOpt3(Thisy)
      For z = 1 To 18
        TotByYrAndType(z, Thisy) = HoldAmt(z, Thisy)
        CntByYrAndType(z, Thisy) = HoldCnt(z, Thisy)
      Next z
      ThEYear(Thisy) = HoldYr
      TotByYrAndPers(Thisy) = HoldPers
      TotByYrAndMT(Thisy) = HoldMT
      TotByYrAndMC(Thisy) = HoldMC
      TotByYrAndFE(Thisy) = HoldFE
      TotByYrAndMH(Thisy) = HoldMH
      TotByYrAndInt(Thisy) = HoldInt
      TotByYrAndAdv(Thisy) = HoldAdv
      TotByYrAndLateList(Thisy) = HoldLateList
      TotByYrAndPen(Thisy) = HoldPen
      TotByYrAndOpt1(Thisy) = HoldOpt1
      TotByYrAndOpt2(Thisy) = HoldOpt2
      TotByYrAndOpt3(Thisy) = HoldOpt3
      If Nexty >= YrCnt Then Exit For
      HoldBigYr = 0 'BigYr + 1
      Nexty = Nexty + 1
    Next x

  For y = 1 To YrCnt
    For x = 1 To 18
      If TotByYrAndType(x, y) > 0 Then
        Select Case x
          Case 1
            Print #SubRptHandle, "Billing"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
              Print #SubRptHandle, TotByYrAndPers(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10                 11         12             13             14
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm;
              '                           15                    16                    17                   18                    19
              Print #SubRptHandle, TotByYrAndPen(y); dlm; TotByYrAndMT(y); dlm; TotByYrAndMC(y); dlm; TotByYrAndFE(y); dlm; TotByYrAndMH(y)
            Else
              '                    4       5       6       7       8       9       10      11      12        13       14       15       16       17       18       19
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 2
            Print #SubRptHandle, "Payment"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
'              Print #SubRptHandle, TotByYrAndPrinc(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              Print #SubRptHandle, TotByYrAndPers(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10                 11         12             13             14
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm;
              '                           15                    16                    17                   18                    19
              Print #SubRptHandle, TotByYrAndPen(y); dlm; TotByYrAndMT(y); dlm; TotByYrAndMC(y); dlm; TotByYrAndFE(y); dlm; TotByYrAndMH(y)
            Else
              '                    4       5       6       7       8       9       10      11       12       13       14       15
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 3
            Print #SubRptHandle, "Release"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
'              Print #SubRptHandle, TotByYrAndPrinc(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              Print #SubRptHandle, TotByYrAndPers(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10                 11         12             13             14
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm;
              '                           15                    16                    17                   18                    19
              Print #SubRptHandle, TotByYrAndPen(y); dlm; TotByYrAndMT(y); dlm; TotByYrAndMC(y); dlm; TotByYrAndFE(y); dlm; TotByYrAndMH(y)
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 4 '
            Print #SubRptHandle, "Interest"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 5
            Print #SubRptHandle, "Penalty"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 6
            Print #SubRptHandle, "Advertising"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 7
            Print #SubRptHandle, "Adjust Pay Down"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
'              Print #SubRptHandle, TotByYrAndPrinc(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              Print #SubRptHandle, TotByYrAndPers(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10                 11         12             13             14
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm;
              '                           15                    16                    17                   18                    19
              Print #SubRptHandle, TotByYrAndPen(y); dlm; TotByYrAndMT(y); dlm; TotByYrAndMC(y); dlm; TotByYrAndFE(y); dlm; TotByYrAndMH(y)
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 8
            Print #SubRptHandle, "Credit at Billing"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 9
            Print #SubRptHandle, "Adjust Bill Down"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
'              Print #SubRptHandle, TotByYrAndPrinc(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              Print #SubRptHandle, TotByYrAndPers(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10                 11         12             13             14
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm;
              '                           15                    16                    17                   18                    19
              Print #SubRptHandle, TotByYrAndPen(y); dlm; TotByYrAndMT(y); dlm; TotByYrAndMC(y); dlm; TotByYrAndFE(y); dlm; TotByYrAndMH(y)
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 10
            Print #SubRptHandle, "Adjust Bill Up"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
'              Print #SubRptHandle, TotByYrAndPrinc(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              Print #SubRptHandle, TotByYrAndPers(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10                 11         12             13             14
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; POpt1Desc; dlm; POpt2Desc; dlm; POpt3Desc; dlm;
              '                           15                    16                    17                   18                    19
              Print #SubRptHandle, TotByYrAndPen(y); dlm; TotByYrAndMT(y); dlm; TotByYrAndMC(y); dlm; TotByYrAndFE(y); dlm; TotByYrAndMH(y)
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 11
            If fpcmbTransType.Text <> " 0) All" Then
              Print #SubRptHandle, "Bill OverPay"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
              '                            4                      5                      6                         7
'              Print #SubRptHandle, TotByYrAndPrinc(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              Print #SubRptHandle, TotByYrAndPers(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10                 11         12             13             14
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm;
              '                           15                    16                    17                   18                    19
              Print #SubRptHandle, TotByYrAndPen(y); dlm; TotByYrAndMT(y); dlm; TotByYrAndMC(y); dlm; TotByYrAndFE(y); dlm; TotByYrAndMH(y)
            Else
              Print #SubRptHandle, "Bill Pay/OverPay"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 12
            Print #SubRptHandle, "OverPayment"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 13
            Print #SubRptHandle, "Adjust Bill Up Affecting Credit Balance"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
'              Print #SubRptHandle, TotByYrAndPrinc(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              Print #SubRptHandle, TotByYrAndPers(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10                 11         12             13             14
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm;
              '                           15                    16                    17                   18                    19
              Print #SubRptHandle, TotByYrAndPen(y); dlm; TotByYrAndMT(y); dlm; TotByYrAndMC(y); dlm; TotByYrAndFE(y); dlm; TotByYrAndMH(y)
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 14
            Print #SubRptHandle, "Adjust Pay Dwn Affecting Credit Balance"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
'              Print #SubRptHandle, TotByYrAndPrinc(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              Print #SubRptHandle, TotByYrAndPers(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10                 11         12             13             14
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm;
              '                           15                    16                    17                   18                    19
              Print #SubRptHandle, TotByYrAndPen(y); dlm; TotByYrAndMT(y); dlm; TotByYrAndMC(y); dlm; TotByYrAndFE(y); dlm; TotByYrAndMH(y)
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 15
            Print #SubRptHandle, "Adjust Prepay Down"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 16
            Print #SubRptHandle, "Refund Prepay"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 17
            Print #SubRptHandle, "PPTRA Removal"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
          Case 18
            Print #SubRptHandle, "Unknown"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""
            End If
        End Select
      End If
    Next x
  Next y
  Close SubRptHandle
  
  Return
  
PrintTotalsPers:
  Sub2RptFile$ = "TAXRPTS\SUB2TAXJRNLP.RPT"
  Sub2RptHandle = FreeFile
  Open Sub2RptFile For Output As #Sub2RptHandle
  GCntTot = 0
  GAmtTot = 0
  For x = 1 To YrCnt
    GPersTot = GPersTot + TotByYrAndPers(x)
    GMTTot = GMTTot + TotByYrAndMT(x)
    GMCTot = GMCTot + TotByYrAndMC(x)
    GFETot = GFETot + TotByYrAndFE(x)
    GMHTot = GMHTot + TotByYrAndMH(x)
    GIntTot = GIntTot + TotByYrAndInt(x)
    GAdvTot = GAdvTot + TotByYrAndAdv(x)
    GLateListTot = GLateListTot + TotByYrAndLateList(x)
    GPenTot = GPenTot + TotByYrAndPen(x)
    GOpt1Tot = GOpt1Tot + TotByYrAndOpt1(x)
    GOpt2Tot = GOpt2Tot + TotByYrAndOpt2(x)
    GOpt3Tot = GOpt3Tot + TotByYrAndOpt3(x)
  Next x
  Dim CaseP11Tot As Double 'added 7/6/06
  For x = 1 To 18
    GCntTot = GCntTot + TotCntByType(x)
    GAmtTot = GAmtTot + TotByType(x)
    If TotByType(x) > 0 Then
      If x <> 11 Or (x = 11 And fpcmbTransType.Text = " 0) All") Then '7/7/06 added Or part
        Print #Sub2RptHandle, TotByType(x); dlm; TotCntByType(x); dlm; GAmtTot; dlm; GCntTot; dlm;
      Else 'inserted 7/6/06
        CaseP11Tot = OldRound(TotByType(11) - (GPersTot + GIntTot + GAdvTot + GLateListTot + GOpt1Tot + GOpt2Tot + GOpt3Tot + GPenTot))
        CaseP11Tot = OldRound(CaseP11Tot - (GMTTot + GMCTot + GFETot + GMHTot))
        Print #Sub2RptHandle, CaseP11Tot; dlm; TotCntByType(x); dlm; GAmtTot; dlm; GCntTot; dlm;
      End If
      '                        4              5             6               7
      Print #Sub2RptHandle, GPersTot; dlm; GIntTot; dlm; GAdvTot; dlm; GLateListTot; dlm;
      '                        8              9             10              11             12             13
      Print #Sub2RptHandle, GOpt1Tot; dlm; GOpt2Tot; dlm; GOpt3Tot; dlm; POpt1Desc; dlm; POpt2Desc; dlm; POpt3Desc; dlm;
      '                       14            15           16           17           18
      Print #Sub2RptHandle, GPenTot; dlm; GMTTot; dlm; GMCTot; dlm; GFETot; dlm; GMHTot; dlm;
      If fpcmbTransType.Text <> " 0) All" Then
        '                      19
        Print #Sub2RptHandle, "1"; dlm;
      Else
        '                      19
        Print #Sub2RptHandle, "2"; dlm;
      End If
      Select Case x
        Case 1
          '                         20         '21
          Print #Sub2RptHandle, "Billing"; dlm; 3
        Case 2
          Print #Sub2RptHandle, "Payment"; dlm; 3
        Case 3
          Print #Sub2RptHandle, "Release"; dlm; 3
        Case 4
          Print #Sub2RptHandle, "Interest"; dlm; 44
        Case 5
          Print #Sub2RptHandle, "Penalty"; dlm; 44
        Case 6
          Print #Sub2RptHandle, "Advertising"; dlm; 44
        Case 7
          Print #Sub2RptHandle, "Adjust Pay Down"; dlm; 3
        Case 8
          Print #Sub2RptHandle, "Credit at Billing"; dlm; 44
        Case 9
          Print #Sub2RptHandle, "Adjust Bill Down"; dlm; 3
        Case 10
          Print #Sub2RptHandle, "Adjust Bill Up"; dlm; 3
        Case 11
          If fpcmbTransType.Text <> " 0) All" Then
            Print #Sub2RptHandle, "Bill OverPay"; dlm; 3 'changed from 44 on 7/6/06
          Else
            Print #Sub2RptHandle, "Bill Pay/OverPay"; dlm; 3 'changed from 44 on 7/6/06
          End If
        Case 12
          Print #Sub2RptHandle, "OverPayment"; dlm; 44
        Case 13
          Print #Sub2RptHandle, "Adjust Bill Up Affecting Credit Balance"; dlm; 3
        Case 14
          Print #Sub2RptHandle, "Adjust Pay Dwn Affecting Credit Balance"; dlm; 3
        Case 15
          Print #Sub2RptHandle, "Adjust Prepay Down"; dlm; 44
        Case 16
          Print #Sub2RptHandle, "Refund Prepay"; dlm; 44
        Case 17
          Print #Sub2RptHandle, "PPTRA Removal"; dlm; 44
        Case 18
          Print #Sub2RptHandle, "Unknown"; dlm; 44
      End Select
    End If
  Next x
  Close Sub2RptHandle
  Return
  
ApplyDiscP:
  Disc1 = 0
  Disc2 = 0
  Disc3 = 0
  Disc4 = 0
  Disc5 = 0
  Disc6 = 0
  Disc7 = 0
  Disc8 = 0
  If TaxTrans.Amount = 0 Then Return
  If TaxTrans.TranType = 1 Then
    SaveAmt = OldRound(TaxTrans.Amount - TaxTrans.DiscAmt)
  Else
    SaveAmt = TaxTrans.Amount
    TaxTrans.Amount = OldRound(TaxTrans.Amount + TaxTrans.DiscAmt)
  End If
  Disc1 = OldRound(TaxTrans.Revenue.Principle1Pd / SaveAmt)
  Disc1 = OldRound(Disc1 * TaxTrans.DiscAmt)
  Disc2 = OldRound(TaxTrans.Revenue.Principle2Pd / SaveAmt)
  Disc2 = OldRound(Disc2 * TaxTrans.DiscAmt)
  Disc3 = OldRound(TaxTrans.Revenue.Principle3Pd / SaveAmt)
  Disc3 = OldRound(Disc3 * TaxTrans.DiscAmt)
  Disc4 = OldRound(TaxTrans.Revenue.Principle4Pd / SaveAmt)
  Disc4 = OldRound(Disc4 * TaxTrans.DiscAmt)
  Disc5 = OldRound(TaxTrans.Revenue.Principle5Pd / SaveAmt)
  Disc5 = OldRound(Disc5 * TaxTrans.DiscAmt)
  Disc6 = OldRound(TaxTrans.Revenue.RevOpt1Pd / SaveAmt)
  Disc6 = OldRound(Disc6 * TaxTrans.DiscAmt)
  Disc7 = OldRound(TaxTrans.Revenue.RevOpt2Pd / SaveAmt)
  Disc7 = OldRound(Disc7 * TaxTrans.DiscAmt)
  Disc8 = OldRound(TaxTrans.Revenue.RevOpt3Pd / SaveAmt)
  Disc8 = OldRound(Disc8 * TaxTrans.DiscAmt)
  TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd + Disc1)
  TaxTrans.Revenue.Principle2Pd = OldRound(TaxTrans.Revenue.Principle2Pd + Disc2)
  TaxTrans.Revenue.Principle3Pd = OldRound(TaxTrans.Revenue.Principle3Pd + Disc3)
  TaxTrans.Revenue.Principle4Pd = OldRound(TaxTrans.Revenue.Principle4Pd + Disc4)
  TaxTrans.Revenue.Principle5Pd = OldRound(TaxTrans.Revenue.Principle5Pd + Disc5)
  TaxTrans.Revenue.RevOpt1Pd = OldRound(TaxTrans.Revenue.RevOpt1Pd + Disc6)
  TaxTrans.Revenue.RevOpt2Pd = OldRound(TaxTrans.Revenue.RevOpt2Pd + Disc7)
  TaxTrans.Revenue.RevOpt3Pd = OldRound(TaxTrans.Revenue.RevOpt3Pd + Disc8)
  DiscApplied = True
  
  Return
  
ApplyDiscR:
  Disc1 = 0
  Disc2 = 0
  Disc3 = 0
  Disc4 = 0
  If TaxTrans.Amount = 0 Then Return
  If TaxTrans.TranType = 1 Then
    SaveAmt = OldRound(TaxTrans.Amount - TaxTrans.DiscAmt)
  Else
    SaveAmt = TaxTrans.Amount
    TaxTrans.Amount = OldRound(TaxTrans.Amount + TaxTrans.DiscAmt)
  End If
  Disc1 = OldRound(TaxTrans.Revenue.Principle1Pd / SaveAmt)
  Disc1 = OldRound(Disc1 * TaxTrans.DiscAmt)
  Disc2 = OldRound(TaxTrans.Revenue.RevOpt1Pd / SaveAmt)
  Disc2 = OldRound(Disc2 * TaxTrans.DiscAmt)
  Disc3 = OldRound(TaxTrans.Revenue.RevOpt2Pd / SaveAmt)
  Disc3 = OldRound(Disc3 * TaxTrans.DiscAmt)
  Disc4 = OldRound(TaxTrans.Revenue.RevOpt3Pd / SaveAmt)
  Disc4 = OldRound(Disc4 * TaxTrans.DiscAmt)
  TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd + Disc1)
  TaxTrans.Revenue.RevOpt1Pd = OldRound(TaxTrans.Revenue.RevOpt1Pd + Disc2)
  TaxTrans.Revenue.RevOpt2Pd = OldRound(TaxTrans.Revenue.RevOpt2Pd + Disc3)
  TaxTrans.Revenue.RevOpt3Pd = OldRound(TaxTrans.Revenue.RevOpt3Pd + Disc4)
  DiscApplied = True
  
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxTransJournal", "PrintGraphics", Erl)
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

Private Function CheckB4Printing() As Boolean
  Dim ThisDate As Integer
  Dim EndDate As Integer
  Dim BegDate As Integer
  
  On Error GoTo ERRORSTUFF
  
  CheckB4Printing = True
  If QPTrim$(fpcmbTransType.Text) = "" Then
    Call TaxMsg(900, "Please make a selection for 'Transaction Type'.")
    fpcmbTransType.SetFocus
    Exit Function
  End If
  
  BegDate = Date2Num(fptxtBegDate.Text)
  EndDate = Date2Num(fptxtEndDate.Text)
  If BegDate > EndDate Then
    CheckB4Printing = False
    Call TaxMsg(800, "The beginning date comes after the ending date. Please correct this oversight.")
    fptxtBegDate.SetFocus
    Exit Function
  End If
  
  If QPTrim$(fpcmbIncInactive.Text) = "" Then
    CheckB4Printing = False
    Call TaxMsg(900, "Please make a selection for 'Include Inactive Accounts'.")
    fpcmbIncInactive.SetFocus
    Exit Function
  End If
  
  If QPTrim$(fpcmbTaxType.Text) = "" Then
    CheckB4Printing = False
    Call TaxMsg(900, "Please make a selection for 'Tax Type'.")
    fpcmbTaxType.SetFocus
    Exit Function
  End If
  
  If QPTrim$(fpcmbPrintOpt.Text) = "" Then
    CheckB4Printing = False
    Call TaxMsg(900, "Please make a selection for 'Report Type'.")
    fpcmbPrintOpt.SetFocus
    Exit Function
  End If
  
  If QPTrim$(fpcmbPrintOrder.Text) = "" Then
    CheckB4Printing = False
    Call TaxMsg(900, "Please make a selection for 'Print Order'.")
    fpcmbPrintOrder.SetFocus
    Exit Function
  End If
  
  Exit Function
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxTransJournal", "CheckB4Printing", Erl)
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

Private Sub PrintText()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim x As Long, y As Integer
  Dim BegDate As Integer
  Dim EndDate As Integer
  Dim ThisRec As Long
  Dim ThisClass As Integer
  Dim ThisType As String
  Dim RptFile$
  Dim RptHandle As Integer
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim InactiveFlag As Boolean
  Dim ThisName$, ThisBillType$
  Dim TCnt As Long, NewName$
  Dim TotAmt As Double
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim NumOfSrchRecs As Long
  Dim ThisTransType As String
  Dim YrCnt As Integer, ThisYear As Integer
  Dim BigYr As Integer
  Dim HoldBigYr As Integer
  Dim HoldYr As Integer
  Dim HoldPrinc As Double
  Dim HoldPers As Double
  Dim HoldMT As Double
  Dim HoldMC As Double
  Dim HoldFE As Double
  Dim HoldMH As Double
  Dim HoldInt As Double
  Dim HoldAdv As Double
  Dim HoldLateList As Double
  Dim HoldPen As Double
  Dim HoldOpt1 As Double
  Dim HoldOpt2 As Double
  Dim HoldOpt3 As Double
  Dim Nextx As Integer
  Dim Thisx As Integer
  Dim Nexty As Integer
  Dim Thisy As Integer
  Dim z As Integer
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim FF$, Page As Integer
  Dim CustName$, PrintCnt As Integer
  Dim ThisBillNum As String * 8
  Dim ThisOperNum As Integer
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim OptFlag As Boolean
  Dim TransDesc$
  Dim GAmtTot As Double
  Dim GCntTot As Long
  Dim BadCnt As Integer
  Dim QFlag As Boolean
  Dim GPrincTot As Double
  Dim GPersTot As Double
  Dim GMTTot As Double
  Dim GMCTot As Double
  Dim GFETot As Double
  Dim GMHTot As Double
  Dim GIntTot As Double
  Dim GAdvTot As Double
  Dim GLateListTot As Double
  Dim GPenTot As Double
  Dim GOpt1Tot As Double
  Dim GOpt2Tot As Double
  Dim GOpt3Tot As Double
  Dim Disc1 As Double '1/16/2007
  Dim Disc2 As Double '1/16/2007
  Dim Disc3 As Double '1/16/2007
  Dim Disc4 As Double '1/16/2007
  Dim Disc5 As Double '1/16/2007
  Dim Disc6 As Double '1/16/2007
  Dim Disc7 As Double '1/16/2007
  Dim Disc8 As Double '1/16/2007
  Dim DiscApplied As Boolean '1/16/2007
  Dim SaveAmt As Double '1/16/2007
  
  On Error GoTo ERRORSTUFF
  
  QFlag = False
  If chkQuick.Value = 1 Then QFlag = True
  DiscApplied = False '1/16/07
  
  TransDesc = QPTrim$(fptxtDesc.Text)
  
  If fpcmbOperNum.Text = "All" Then
    ThisOperNum = 0
  Else
    ThisOperNum = CInt(fpcmbOperNum.Text)
  End If
  
  CustName = ""
  OptFlag = False
  IdxFlag = False
  If CheckB4Printing = False Then
    Exit Sub
  End If
  FF$ = Chr$(12)
  MaxLines = 58
  LineCnt = 0
  ThisBillType = QPTrim$(fpcmbTaxType.Text)
  If fpcmbIncInactive.Text = "No" Then
    InactiveFlag = False
  Else
    InactiveFlag = True
  End If
  
  If Mid(fpcmbTransType.Text, 1, 1) = "" Then
    ThisClass = CInt(Mid(fpcmbTransType.Text, 2, 1))
  Else
    ThisClass = CInt(Mid(fpcmbTransType.Text, 1, 2))
  End If
    
  Select Case ThisClass
    Case 1
      ThisType = "Billing"
    Case 2
      ThisType = "Payment"
    Case 3
      ThisType = "Release"
    Case 4
      ThisType = "Interest"
    Case 5
      ThisType = "Penalty"
    Case 6
      ThisType = "Advertising Charge"
    Case 7, 10
      ThisType = "Adjust Pay Down"
    Case 9
      ThisType = "Credit at Billing"
    Case 11
      ThisType = "Adjust Prepay Down"
    Case 12
      ThisType = "Refund Prepay"
    Case 14, 24
      ThisType = "Adjust Bill Up"
    Case 21
      ThisType = "Payment w/Overpay"
    Case 22
      ThisType = "Overpayment Only"
    Case 30
      ThisType = "PPTRA Removal"
    Case Else
      ThisType = "All"
  End Select
    
  BegDate = Date2Num(fptxtBegDate.Text)
  EndDate = Date2Num(fptxtEndDate.Text)
  
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
    OptFlag = True
  End If

  RptFile$ = "TAXRPTS\TAXJRNL.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  GoSub PrintHeader
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  
  frmVATaxShowPctComp.Label1 = "Gathering Tax Transaction Data"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  
  ReDim TotByYrAndType(1 To 18, 1 To 1) As Double
  ReDim CntByYrAndType(1 To 18, 1 To 1) As Integer
  ReDim TotByType(1 To 18) As Double
  ReDim TotCntByType(1 To 18) As Long
  ReDim ThEYear(1 To 1) As Integer
  
  ReDim TotByYrAndPrinc(1 To 1) As Double
  ReDim TotByYrAndPers(1 To 1) As Double
  ReDim TotByYrAndMT(1 To 1) As Double
  ReDim TotByYrAndMC(1 To 1) As Double
  ReDim TotByYrAndFE(1 To 1) As Double
  ReDim TotByYrAndMH(1 To 1) As Double
  ReDim TotByYrAndInt(1 To 1) As Double
  ReDim TotByYrAndAdv(1 To 1) As Double
  ReDim TotByYrAndLateList(1 To 1) As Double
  ReDim TotByYrAndPen(1 To 1) As Double
  ReDim TotByYrAndOpt1(1 To 1) As Double
  ReDim TotByYrAndOpt2(1 To 1) As Double
  ReDim TotByYrAndOpt3(1 To 1) As Double
  
  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  End If
  
  For x = 1 To NumOfTCRecs
    If IdxFlag = False Then
      Get TCHandle, x, TaxCust
    Else
      Get TCHandle, IdxArray(x), TaxCust
    End If
    If TaxCust.Active = "N" And InactiveFlag = False Then
      GoTo SkipIt
    End If
    ThisName = QPTrim$(TaxCust.CustName)
    ThisRec = TaxCust.LastTrans
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintHeader
    End If
    PrintCnt = 0
    BadCnt = 0
    Do While ThisRec > 0
      Get TTHandle, ThisRec, TaxTrans
      DiscApplied = False '1/16/07
      If QFlag = True Then
        If TaxTrans.TransDate < BegDate Then
          BadCnt = BadCnt + 1
          If BadCnt > 3 Then Exit Do
        End If
      End If
      If TransDesc <> "" Then
        If InStr(1, TaxTrans.Description, TransDesc) = 0 Then GoTo SkipIt
      End If
      If ThisOperNum <> 0 Then
        If ThisOperNum <> TaxTrans.OperNum Then GoTo SkipIt
      End If
      If TaxTrans.BillType <> "R" And TaxTrans.BillType <> "P" Then TaxTrans.BillType = ""
      If TaxTrans.BillType <> Mid(fpcmbTaxType.Text, 1, 1) And QPTrim$(TaxTrans.BillType) <> "" Then GoTo SkipIt
      If ThisClass = 7 And TaxTrans.TranType = 10 Then GoTo ItsOK
      If ThisClass = 14 And TaxTrans.TranType = 24 Then GoTo ItsOK
'      If ThisClass = 2 And TaxTrans.TranType = 21 Then GoTo ItsOK 'commented out on 7/6/06
      If ThisClass <> 0 And TaxTrans.TranType <> ThisClass Then GoTo SkipIt
ItsOK:
      If TaxTrans.TransDate >= BegDate And TaxTrans.TransDate <= EndDate Then
        If TaxTrans.BelongTo > 0 Then
          Get TTHandle, TaxTrans.BelongTo, TaxTrans
        End If
        If PrintCnt = 0 Then
          If LineCnt <> 9 Then
            Print #RptHandle,
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintHeader
            End If
          End If
          GoSub PrintCustHeader
        End If
        PrintCnt = PrintCnt + 1
        If YrCnt = 0 Then
           YrCnt = YrCnt + 1
           ThisYear = YrCnt
           ReDim Preserve ThEYear(1 To YrCnt) As Integer
           ThEYear(YrCnt) = TaxTrans.TaxYear
           ReDim Preserve TotByYrAndType(1 To 18, 1 To YrCnt) As Double
           ReDim Preserve CntByYrAndType(1 To 18, 1 To YrCnt) As Integer
           ReDim Preserve TotByYrAndPrinc(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndPers(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndMT(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndMC(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndFE(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndMH(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndInt(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndAdv(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndLateList(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndPen(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndOpt1(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndOpt2(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndOpt3(1 To YrCnt) As Double
           For y = 1 To 18
             TotByYrAndType(y, YrCnt) = 0
             CntByYrAndType(y, YrCnt) = 0
           Next y
         Else
           For y = 1 To YrCnt
             If TaxTrans.TaxYear = ThEYear(y) Then
               ThisYear = y
               Exit For
             End If
           Next y
           If y > YrCnt Then
             YrCnt = YrCnt + 1
             ThisYear = YrCnt
             ReDim Preserve ThEYear(1 To YrCnt) As Integer
             ThEYear(YrCnt) = TaxTrans.TaxYear
             ReDim Preserve TotByYrAndType(1 To 18, 1 To YrCnt) As Double
             ReDim Preserve CntByYrAndType(1 To 18, 1 To YrCnt) As Integer
             ReDim Preserve TotByYrAndPrinc(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndPers(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndMT(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndMC(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndFE(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndMH(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndInt(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndAdv(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndLateList(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndPen(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndOpt1(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndOpt2(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndOpt3(1 To YrCnt) As Double
             For y = 1 To 18
               TotByYrAndType(y, YrCnt) = 0
               CntByYrAndType(y, YrCnt) = 0
             Next y
           End If
         End If
         Get TTHandle, ThisRec, TaxTrans
          
         Select Case TaxTrans.TranType
           Case 1
             ThisTransType = "Billing"
             TotByYrAndType(1, ThisYear) = OldRound(TotByYrAndType(1, ThisYear) + TaxTrans.Amount)
             TotByType(1) = OldRound(TotByType(1) + TaxTrans.Amount)
             CntByYrAndType(1, ThisYear) = OldRound(CntByYrAndType(1, ThisYear) + 1)
             TotCntByType(1) = OldRound(TotCntByType(1) + 1)
             TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1)
             TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1)
             TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2)
             TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3)
             TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4)
             TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5)
             '7/11/06 added back interest, advertising and penalty to accommodate manual bills
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.Interest)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.Collection)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.Penalty)
             
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateList)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3)
           Case 2
             ThisTransType = "Payment"
             If TaxTrans.DiscAmt > 0 Then 'added 1/16/07
               If TaxTrans.BillType = "R" Then
                 GoSub ApplyDiscR
               ElseIf TaxTrans.BillType = "P" Then
                 GoSub ApplyDiscP
               End If
             End If
             TotByYrAndType(2, ThisYear) = OldRound(TotByYrAndType(2, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(2, ThisYear) = OldRound(CntByYrAndType(2, ThisYear) + 1)
             TotByType(2) = OldRound(TotByType(2) + TaxTrans.Amount)
             TotCntByType(2) = OldRound(TotCntByType(2) + 1)
             TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2Pd)
             TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3Pd)
             TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4Pd)
             TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5Pd)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
           Case 3
            '7/12/06 changed revenues for Release to paid from charged
             ThisTransType = "Release"
             TotByYrAndType(3, ThisYear) = OldRound(TotByYrAndType(3, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(3, ThisYear) = OldRound(CntByYrAndType(3, ThisYear) + 1)
             TotByType(3) = OldRound(TotByType(3) + TaxTrans.Amount)
             TotCntByType(3) = OldRound(TotCntByType(3) + 1)
             TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2Pd)
             TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3Pd)
             TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4Pd)
             TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5Pd)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
           Case 4
             ThisTransType = "Interest"
             TotByYrAndType(4, ThisYear) = OldRound(TotByYrAndType(4, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(4, ThisYear) = OldRound(CntByYrAndType(4, ThisYear) + 1)
             TotByType(4) = OldRound(TotByType(4) + TaxTrans.Amount)
             TotCntByType(4) = OldRound(TotCntByType(4) + 1)
           Case 5
             ThisTransType = "Penalty"
             TotByYrAndType(5, ThisYear) = OldRound(TotByYrAndType(5, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(5, ThisYear) = OldRound(CntByYrAndType(5, ThisYear) + 1)
             TotByType(5) = OldRound(TotByType(5) + TaxTrans.Amount)
             TotCntByType(5) = OldRound(TotCntByType(5) + 1)
           Case 6
             ThisTransType = "Advertising Charge"
             TotByYrAndType(6, ThisYear) = OldRound(TotByYrAndType(6, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(6, ThisYear) = OldRound(CntByYrAndType(6, ThisYear) + 1)
             TotByType(6) = OldRound(TotByType(6) + TaxTrans.Amount)
             TotCntByType(6) = OldRound(TotCntByType(6) + 1)
           Case 7
             ThisTransType = "Adjust Pay Down"
             TotByYrAndType(7, ThisYear) = OldRound(TotByYrAndType(7, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(7, ThisYear) = OldRound(CntByYrAndType(7, ThisYear) + 1)
             TotByType(7) = OldRound(TotByType(7) + TaxTrans.Amount)
             TotCntByType(7) = OldRound(TotCntByType(7) + 1) '7/11/06 added Pd to revenues
             TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2Pd)
             TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3Pd)
             TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4Pd)
             TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5Pd)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
           Case 9
             ThisTransType = "Cred at Billing"
             TotByYrAndType(8, ThisYear) = OldRound(TotByYrAndType(8, ThisYear) + TaxTrans.Revenue.PrePaidUsed)
             CntByYrAndType(8, ThisYear) = OldRound(CntByYrAndType(8, ThisYear) + 1)
             TotByType(8) = OldRound(TotByType(8) + TaxTrans.Revenue.PrePaidUsed)
             TotCntByType(8) = OldRound(TotCntByType(8) + 1)
           Case 13
             ThisTransType = "Adjust Bill Down"
             TotByYrAndType(9, ThisYear) = OldRound(TotByYrAndType(9, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(9, ThisYear) = OldRound(CntByYrAndType(9, ThisYear) + 1)
             TotByType(9) = OldRound(TotByType(9) + TaxTrans.Amount)
             TotCntByType(9) = OldRound(TotCntByType(9) + 1)
             TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1)
             TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1)
             TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2)
             TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3)
             TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4)
             TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.Interest)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.Collection)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateList)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.Penalty)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3)
           Case 14
             ThisTransType = "Adjust Bill Up"
             TotByYrAndType(10, ThisYear) = OldRound(TotByYrAndType(10, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(10, ThisYear) = OldRound(CntByYrAndType(10, ThisYear) + 1)
             TotByType(10) = OldRound(TotByType(10) + TaxTrans.Amount)
             TotCntByType(10) = OldRound(TotCntByType(10) + 1)
             TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1)
             TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1)
             TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2)
             TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3)
             TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4)
             TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.Interest)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.Collection)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateList)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.Penalty)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3)
           Case 21 'added revenue detail and changed Amount to PrePaidAmt on 7/6/06
             ThisTransType = "Billpay/Overpay" '7/7/06 added the All if statement
             If TaxTrans.DiscAmt > 0 Then 'added 1/16/07
               If TaxTrans.BillType = "R" Then
                 GoSub ApplyDiscR
               ElseIf TaxTrans.BillType = "P" Then
                 GoSub ApplyDiscP
               End If
             End If
             If fpcmbTransType.Text <> " 0) All" Then
               TotByYrAndType(11, ThisYear) = OldRound(TotByYrAndType(11, ThisYear) + TaxTrans.Revenue.PrePaidAmt) '  .Amount)
             Else
               TotByYrAndType(11, ThisYear) = OldRound(TotByYrAndType(11, ThisYear) + TaxTrans.Amount)
             End If
             CntByYrAndType(11, ThisYear) = OldRound(CntByYrAndType(11, ThisYear) + 1)
             TotByType(11) = OldRound(TotByType(11) + TaxTrans.Amount)
             TotCntByType(11) = OldRound(TotCntByType(11) + 1)
             TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2Pd)
             TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3Pd)
             TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4Pd)
             TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5Pd)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
           Case 22
             ThisTransType = "Overpayment"
             TotByYrAndType(12, ThisYear) = OldRound(TotByYrAndType(12, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(12, ThisYear) = OldRound(CntByYrAndType(12, ThisYear) + 1)
             TotByType(12) = OldRound(TotByType(12) + TaxTrans.Amount)
             TotCntByType(12) = OldRound(TotCntByType(12) + 1)
           Case 24
             ThisTransType = "Adj Bill Up -Cre"
             TotByYrAndType(13, ThisYear) = OldRound(TotByYrAndType(13, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(13, ThisYear) = OldRound(CntByYrAndType(13, ThisYear) + 1)
             TotByType(13) = OldRound(TotByType(13) + TaxTrans.Amount)
             TotCntByType(13) = OldRound(TotCntByType(13) + 1)
             TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2Pd)
             TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3Pd)
             TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4Pd)
             TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5Pd)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
             
'             TotByYrAndType(13, ThisYear) = OldRound(TotByYrAndType(13, ThisYear) + TaxTrans.Amount)
'             CntByYrAndType(13, ThisYear) = OldRound(CntByYrAndType(13, ThisYear) + 1)
'             TotByType(13) = OldRound(TotByType(13) + TaxTrans.Amount)
'             TotCntByType(13) = OldRound(TotCntByType(13) + 1)
'             TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1)
'             TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1)
'             TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2)
'             TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3)
'             TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4)
'             TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5)
'             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.Interest)
'             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.Collection)
'             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateList)
'             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.Penalty)
'             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1)
'             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2)
'             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3)
           Case 10 ', 24 '7/11/06 added Pd on revenues
             ThisTransType = "Adj Pay Dwn -Cre"

             TotByYrAndType(14, ThisYear) = OldRound(TotByYrAndType(14, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(14, ThisYear) = OldRound(CntByYrAndType(14, ThisYear) + 1)
             TotByType(14) = OldRound(TotByType(14) + TaxTrans.Amount)
             TotCntByType(14) = OldRound(TotCntByType(14) + 1)
             TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2Pd)
             TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3Pd)
             TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4Pd)
             TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5Pd)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
           Case 11
             ThisTransType = "Adj Prepay  -Cre"
             TotByYrAndType(15, ThisYear) = OldRound(TotByYrAndType(15, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(15, ThisYear) = OldRound(CntByYrAndType(15, ThisYear) + 1)
             TotByType(15) = OldRound(TotByType(15) + TaxTrans.Amount)
             TotCntByType(15) = OldRound(TotCntByType(15) + 1)
           Case 12
             ThisTransType = "Ref Prepay  -Cre"
             TotByYrAndType(16, ThisYear) = OldRound(TotByYrAndType(16, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(16, ThisYear) = OldRound(CntByYrAndType(16, ThisYear) + 1)
             TotByType(16) = OldRound(TotByType(16) + TaxTrans.Amount)
             TotCntByType(16) = OldRound(TotCntByType(16) + 1)
           Case 30
             ThisTransType = "PPTRA Removal"
             TotByYrAndType(17, ThisYear) = OldRound(TotByYrAndType(17, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(17, ThisYear) = OldRound(CntByYrAndType(17, ThisYear) + 1)
             TotByType(17) = OldRound(TotByType(17) + TaxTrans.Amount)
             TotCntByType(17) = OldRound(TotCntByType(17) + 1)
           Case Else
             ThisTransType = "Unknown"
             TotByYrAndType(18, ThisYear) = OldRound(TotByYrAndType(18, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(18, ThisYear) = OldRound(CntByYrAndType(18, ThisYear) + 1)
             TotByType(18) = OldRound(TotByType(18) + TaxTrans.Amount)
             TotCntByType(18) = OldRound(TotCntByType(18) + 1)
         End Select
         TCnt = TCnt + 1
         If TaxTrans.TranType = 2 Or TaxTrans.TranType = 21 Then 'added 1/16/07
           If TaxTrans.DiscAmt > 0 Then 'added 1/16/07
             If TaxTrans.BillType = "R" Then
               GoSub ApplyDiscR
             ElseIf TaxTrans.BillType = "P" Then
               GoSub ApplyDiscP
             End If
           End If
         End If
         TotAmt = OldRound(TotAmt + TaxTrans.Amount)
      
         Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); QPTrim$(TaxTrans.Description);
         If TaxTrans.BelongTo > 0 Then
           Get TTHandle, TaxTrans.BelongTo, TaxTrans
           Print #RptHandle, Tab(37); Using$("###0", TaxTrans.TaxYear);
           Get TTHandle, ThisRec, TaxTrans
           If DiscApplied = True Then 'added 1/16/07
             If TaxTrans.BillType = "R" Then
               GoSub ApplyDiscR
             ElseIf TaxTrans.BillType = "P" Then
               GoSub ApplyDiscP
             End If
           End If
         Else
           Print #RptHandle, Tab(37); Using$("###0", TaxTrans.TaxYear);
         End If
         Print #RptHandle, Tab(45); Using$("$##,##0.00", TaxTrans.Revenue.PrePaidAmt);
         If TaxTrans.TranType <> 9 Then
           Print #RptHandle, Tab(56); Using$("$##,##0.00", TaxTrans.Amount); Tab(69);
         Else
           Print #RptHandle, Tab(56); Using$("$##,##0.00", TaxTrans.Revenue.PrePaidUsed); Tab(69);
         End If
        
        LineCnt = LineCnt + 1
        If TaxTrans.BelongTo > 0 Then
          Get TTHandle, TaxTrans.BelongTo, TaxTrans
          ThisBillNum = ParseBillNum(TaxTrans.Description)
          If IsNumeric(ThisBillNum) Then
            Print #RptHandle, Using$("######", CDbl(ThisBillNum));
          Else
            Print #RptHandle, "   " + ThisBillNum;
          End If
        Else
          Print #RptHandle, "     0";
        End If
      
        Get TTHandle, ThisRec, TaxTrans
        Print #RptHandle, Tab(79); ThisTransType; Tab(98); CStr(TaxTrans.OperNum)
        If LineCnt >= MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintHeader
        End If
      End If
      
SkipIt:
      ThisRec = TaxTrans.LastTrans
    Loop
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
  
  If TCnt = 0 Then
    Call TaxMsg(900, "No transactions were found that fit The parameters entered.")
    Close
    Exit Sub
  End If
  
  If YrCnt > 0 And fpcmbTaxType.Text = "Real Only" Then
    GoSub SortIt
    GoSub PrintTotals
  ElseIf YrCnt > 0 And fpcmbTaxType.Text = "Personal Only" Then
    GoSub SortItPers
    GoSub PrintTotalsPers
  End If
'  If YrCnt > 0 Then
'    GoSub SortIt
'    GoSub PrintTotals
'  End If
  Print #RptHandle, FF$
  Close
  ViewPrint RptFile, "Tax Transactions Report", True
  
  Exit Sub
  
SortIt:
  
  BigYr = 0
  For x = 1 To YrCnt
    If ThEYear(x) > BigYr Then
      BigYr = ThEYear(x)
    End If
  Next x
  ReDim HoldAmt(1 To 18, 1 To YrCnt) As Double
  ReDim HoldCnt(1 To 18, 1 To YrCnt) As Integer
  
  Nexty = 1
  Nextx = 1
  HoldBigYr = 0
  For x = 1 To 18
    For y = Nexty To YrCnt
      If ThEYear(y) >= HoldBigYr Then
        HoldBigYr = ThEYear(y)
        Thisx = x
        Thisy = y
      End If
    Next y
    For z = 1 To 18
      HoldAmt(z, Thisy) = TotByYrAndType(z, Nexty)
      HoldCnt(z, Thisy) = CntByYrAndType(z, Nexty)
    Next z
    HoldYr = ThEYear(Nexty)
    HoldPrinc = TotByYrAndPrinc(Nexty)
    HoldInt = TotByYrAndInt(Nexty)
    HoldAdv = TotByYrAndAdv(Nexty)
    HoldLateList = TotByYrAndLateList(Nexty)
    HoldPen = TotByYrAndPen(Nexty)
    HoldOpt1 = TotByYrAndOpt1(Nexty)
    HoldOpt2 = TotByYrAndOpt2(Nexty)
    HoldOpt3 = TotByYrAndOpt3(Nexty)
    For z = 1 To 18
      TotByYrAndType(z, Nexty) = TotByYrAndType(z, Thisy)
      CntByYrAndType(z, Nexty) = CntByYrAndType(z, Thisy)
    Next z
    ThEYear(Nexty) = ThEYear(Thisy)
    TotByYrAndPrinc(Nexty) = TotByYrAndPrinc(Thisy)
    TotByYrAndInt(Nexty) = TotByYrAndInt(Thisy)
    TotByYrAndAdv(Nexty) = TotByYrAndAdv(Thisy)
    TotByYrAndLateList(Nexty) = TotByYrAndLateList(Thisy)
    TotByYrAndPen(Nexty) = TotByYrAndPen(Thisy)
    TotByYrAndOpt1(Nexty) = TotByYrAndOpt1(Thisy)
    TotByYrAndOpt2(Nexty) = TotByYrAndOpt2(Thisy)
    TotByYrAndOpt3(Nexty) = TotByYrAndOpt3(Thisy)
    For z = 1 To 18
      TotByYrAndType(z, Thisy) = HoldAmt(z, Thisy)
      CntByYrAndType(z, Thisy) = HoldCnt(z, Thisy)
    Next z
    ThEYear(Thisy) = HoldYr
    TotByYrAndPrinc(Thisy) = HoldPrinc
    TotByYrAndInt(Thisy) = HoldInt
    TotByYrAndAdv(Thisy) = HoldAdv
    TotByYrAndLateList(Thisy) = HoldLateList
    TotByYrAndPen(Thisy) = HoldPen
    TotByYrAndOpt1(Thisy) = HoldOpt1
    TotByYrAndOpt2(Thisy) = HoldOpt2
    TotByYrAndOpt3(Thisy) = HoldOpt3
    If Nexty >= YrCnt Then Exit For
    HoldBigYr = 0
    Nexty = Nexty + 1
  Next x
  Print #RptHandle, FF$
  GoSub PrintSortHeader
  LineCnt = LineCnt + 2
  For y = 1 To YrCnt
   If LineCnt >= MaxLines - 4 Then
     Print #RptHandle, FF$
     GoSub PrintSortHeader
     Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
   Else
     Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
   End If
   LineCnt = LineCnt + 1
   For x = 1 To 18
      If TotByYrAndType(x, y) > 0 Then
        Select Case x
          Case 1
            Print #RptHandle, "  Billing"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOne
            Print #RptHandle, Tab(5); "Principal:    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPrinc(y))
            '7/11/06 added back Int, Adv and Pen to accommodate manual bills
            Print #RptHandle, Tab(5); "Interest:     "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Advertising:  "; Tab(55); Using$("$###,###,##0.00", TotByYrAndAdv(y))
            Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndLateList(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 5 '4
            If LineCnt >= MaxLines - 3 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 2
            Print #RptHandle, "  Payment"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOne
            Print #RptHandle, Tab(5); "Principal:    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPrinc(y))
            Print #RptHandle, Tab(5); "Interest :    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Advertising : "; Tab(55); Using$("$###,###,##0.00", TotByYrAndAdv(y))
            Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndLateList(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 5
            If LineCnt >= MaxLines - 3 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 3
            Print #RptHandle, "  Release"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOne
            Print #RptHandle, Tab(5); "Principal:    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPrinc(y))
            Print #RptHandle, Tab(5); "Interest :    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Advertising : "; Tab(55); Using$("$###,###,##0.00", TotByYrAndAdv(y))
            Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndLateList(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 5
            If LineCnt >= MaxLines - 3 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 4
            Print #RptHandle, "  Interest"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
          Case 5
            Print #RptHandle, "  Penalty"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
          Case 6
            Print #RptHandle, "  Advertising"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
          Case 7
            Print #RptHandle, "  Adjust Pay Down"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOne
            Print #RptHandle, Tab(5); "Principal:    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPrinc(y))
            Print #RptHandle, Tab(5); "Interest :    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Advertising : "; Tab(55); Using$("$###,###,##0.00", TotByYrAndAdv(y))
            Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndLateList(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 5
            If LineCnt >= MaxLines - 3 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 8
            Print #RptHandle, "  Credit at Billing"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
          Case 9
            Print #RptHandle, "  Adjust Bill Down"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOne
            Print #RptHandle, Tab(5); "Principal:    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPrinc(y))
            Print #RptHandle, Tab(5); "Interest :    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Advertising : "; Tab(55); Using$("$###,###,##0.00", TotByYrAndAdv(y))
            Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndLateList(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 5
            If LineCnt >= MaxLines - 3 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 10
            Print #RptHandle, "  Adjust Bill Up"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOne
            Print #RptHandle, Tab(5); "Principal:    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPrinc(y))
            Print #RptHandle, Tab(5); "Interest :    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Advertising : "; Tab(55); Using$("$###,###,##0.00", TotByYrAndAdv(y))
            Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndLateList(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 5
            If LineCnt >= MaxLines - 3 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 11
            If fpcmbTransType.Text <> " 0) All" Then '7/7/06 added All if
              Print #RptHandle, "  Bill OverPay"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            Else
              Print #RptHandle, "  Bill Pay/OverPay"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            End If
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOne
            'added revenue detail on 7/6/06
            Print #RptHandle, Tab(5); "Principal:    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPrinc(y))
            Print #RptHandle, Tab(5); "Interest :    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Advertising : "; Tab(55); Using$("$###,###,##0.00", TotByYrAndAdv(y))
            Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndLateList(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 5
            If LineCnt >= MaxLines - 3 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 12
            Print #RptHandle, "  OverPayment"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
          Case 13
            Print #RptHandle, "  Adj Bill Up Affecting Credit Bal"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOne
            Print #RptHandle, Tab(5); "Principal:    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPrinc(y))
            Print #RptHandle, Tab(5); "Interest :    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Advertising : "; Tab(55); Using$("$###,###,##0.00", TotByYrAndAdv(y))
            Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndLateList(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 5
            If LineCnt >= MaxLines - 3 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 14
            Print #RptHandle, "  Adj Pay Dn Affecting Credit Bal"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOne
            Print #RptHandle, Tab(5); "Principal:    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPrinc(y))
            Print #RptHandle, Tab(5); "Interest :    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Advertising : "; Tab(55); Using$("$###,###,##0.00", TotByYrAndAdv(y))
            Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndLateList(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 5
            If LineCnt >= MaxLines - 3 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 15
            Print #RptHandle, "  Adj Prepay Down"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
          Case 16
            Print #RptHandle, "  Refund Prepay"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
          Case 17
            Print #RptHandle, "  PPTRA Removal"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
          Case 18
            Print #RptHandle, "  Unknown"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
        End Select
      End If
NextOne:
    Next x
    Print #RptHandle, String$(100, "-")
    Print #RptHandle,
    LineCnt = LineCnt + 2
  Next y
  
  Return

PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(38); "Tax Transactions Journal"
  Print #RptHandle, Town; Tab(75); "Page #: " + CStr(Page)
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle, "Tax Type: " + ThisType
  Print #RptHandle, "Transaction Type: " + ThisBillType
  If ThisOperNum <> 0 Then
    Print #RptHandle, "Date Range: " + fptxtBegDate.Text + " to " + fptxtEndDate.Text; Tab(65); "Operator Number: " + CStr(ThisOperNum)
  Else
    Print #RptHandle, "Date Range: " + fptxtBegDate.Text + " to " + fptxtEndDate.Text; Tab(65); "Operator Number: " + "All"
  End If
  Print #RptHandle,
  Print #RptHandle, "Trans Date"; Tab(12); "Description"; Tab(35); "Tax Year"; Tab(44); "Overpay Amt"; Tab(57); "Trans Amt"; Tab(67); "Belongs To"; Tab(78); "Trans Type"; Tab(95); "Oper #"
  Print #RptHandle, String(100, "-")
  LineCnt = 9
  
  Return
  
PrintCustHeader:
  If LineCnt <> 9 Then
    Print #RptHandle, String(100, "-")
    LineCnt = LineCnt + 1
  End If
  If LineCnt >= MaxLines - 3 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  Print #RptHandle, "Cust Num: " + Using$("#######0", TaxCust.Acct); Tab(21); "Customer Name: "; Tab(37); QPTrim$(TaxCust.CustName); Tab(80); "Active: "; Tab(89); TaxCust.Active
  If UseOpt = "Y" Then
    Print #RptHandle, Tab(21); ThisOpt + ": " + QPTrim$(TaxCust.OptSrchDesc)
    LineCnt = LineCnt + 1
  End If
  Print #RptHandle, String(100, ".")
  LineCnt = LineCnt + 2
  
  Return
  
PrintSortHeader:
  Page = Page + 1
  Print #RptHandle, Tab(30); "Tax Transactions Journal Summary"
  Print #RptHandle, Town; Tab(75); "Page #: " + CStr(Page)
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle, "Tax Type: " + ThisType
  Print #RptHandle, "Transaction Type: " + ThisBillType
  Print #RptHandle, "Date Range: " + fptxtBegDate.Text + " to " + fptxtEndDate.Text
  Print #RptHandle, "Note: Adjustment transaction amounts are reflected in revenues and "
  Print #RptHandle, "      adjustment transaction totals exclusively. "
  Print #RptHandle, "Description"; Tab(35); "Trans Cnt"; Tab(64); "Amount"
  Print #RptHandle, String$(100, "-")
  LineCnt = 10
  
  Return
  
PrintTotalsHeader:
  Page = Page + 1
  Print #RptHandle, Tab(30); "Tax Transactions Journal Summary"
  Print #RptHandle, "Grand Totals"
  Print #RptHandle, Town; Tab(75); "Page #: " + CStr(Page)
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle, "Tax Type: " + ThisType
  Print #RptHandle, "Transaction Type: " + ThisBillType
  Print #RptHandle, "Date Range: " + fptxtBegDate.Text + " to " + fptxtEndDate.Text
  Print #RptHandle, "Note: Adjustment transaction amounts are reflected in revenues and "
  Print #RptHandle, "      adjustment transaction totals exclusively. "
  Print #RptHandle, Tab(35); "Trans Cnt"; Tab(64); "Amount"
  Print #RptHandle, String$(100, "-")
  LineCnt = 10

  Return
  
PrintTotals:
  GCntTot = 0
  GAmtTot = 0
  If LineCnt > MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintTotalsHeader
  Else
    Print #RptHandle,
    Print #RptHandle, "Grand Totals"
    Print #RptHandle, Tab(35); "Trans Cnt"; Tab(64); "Amount"
    Print #RptHandle, String$(100, "-")
    LineCnt = LineCnt + 4
  End If
  For x = 1 To YrCnt
    GPrincTot = GPrincTot + TotByYrAndPrinc(x)
    GIntTot = GIntTot + TotByYrAndInt(x)
    GAdvTot = GAdvTot + TotByYrAndAdv(x)
    GLateListTot = GLateListTot + TotByYrAndLateList(x)
    GPenTot = GPenTot + TotByYrAndPen(x)
    GOpt1Tot = GOpt1Tot + TotByYrAndOpt1(x)
    GOpt2Tot = GOpt2Tot + TotByYrAndOpt2(x)
    GOpt3Tot = GOpt3Tot + TotByYrAndOpt3(x)
  Next x
  For x = 1 To 18
    GCntTot = GCntTot + TotCntByType(x)
    GAmtTot = GAmtTot + TotByType(x)
    If TotByType(x) > 0 Then
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintTotalsHeader
      End If
      Select Case x
        Case 1
          Print #RptHandle, "Billing";
        Case 2
          Print #RptHandle, "Payment";
        Case 3
          Print #RptHandle, "Release";
        Case 4
          Print #RptHandle, "Interest";
        Case 5
          Print #RptHandle, "Penalty";
        Case 6
          Print #RptHandle, "Advertising";
        Case 7
          Print #RptHandle, "Adjust Pay Down";
        Case 8
          Print #RptHandle, "Credit at Billing";
        Case 9
          Print #RptHandle, "Adjust Bill Down";
        Case 10
          Print #RptHandle, "Adjust Bill Up";
        Case 11
          If fpcmbTransType.Text = " 0) All" Then '7/7/06 add Bill Pay/Overpay
            Print #RptHandle, "Bill Pay/OverPay";
          Else
            Print #RptHandle, "Bill OverPay";
          End If
        Case 12
          Print #RptHandle, "OverPayment";
        Case 13
          Print #RptHandle, "Adjust Bill Up Affecting Credit Balance";
        Case 14
          Print #RptHandle, "Adjust Pay Dwn Affecting Credit Balance";
        Case 15
          Print #RptHandle, "Adjust Prepay Down";
        Case 16
          Print #RptHandle, "Refund Prepay";
        Case 17
          Print #RptHandle, "PPTRA Removal";
        Case 18
          Print #RptHandle, "Unknown";
      End Select
      Dim CaseR11Tot As Double 'added on 7/6/06
      If x <> 11 Or (x = 11 And fpcmbTransType.Text = " 0) All") Then 'added All if statement on 7/7/06
        Print #RptHandle, Tab(38); Using$("##,##0", TotCntByType(x)); Tab(55); Using$("$###,###,##0.00", TotByType(x))
      Else 'added on 7/6/06
        CaseR11Tot = OldRound(TotByType(11) - (GPrincTot + GIntTot + GAdvTot + GLateListTot + GPenTot + GOpt1Tot + GOpt2Tot + GOpt3Tot))
        Print #RptHandle, Tab(38); Using$("##,##0", TotCntByType(11)); Tab(55); Using$("$###,###,##0.00", CaseR11Tot)
      End If
      LineCnt = LineCnt + 1
      If fpcmbTransType.Text = " 0) All" Then GoTo All
      Select Case x
        Case 1
          If LineCnt >= MaxLines - 5 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Principal: "; Tab(55); Using$("$###,###,##0.00", GPrincTot)
          '7/11/06 added back int, adv and pen to accommodate manual bills
          Print #RptHandle, Tab(5); "Interest: "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising: "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", GLateListTot)
          Print #RptHandle, Tab(5); "Penalty: "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 2
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
        Case 2
          If LineCnt >= MaxLines - 7 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Principal:    "; Tab(55); Using$("$###,###,##0.00", GPrincTot)
          Print #RptHandle, Tab(5); "Interest:     "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising:  "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", GLateListTot)
          Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 5
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
        Case 3
          If LineCnt >= MaxLines - 7 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Principal: "; Tab(55); Using$("$###,###,##0.00", GPrincTot)
          Print #RptHandle, Tab(5); "Interest: "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising: "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", GLateListTot)
          Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 5
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
        Case 4
          GoTo All
        Case 5
          GoTo All
        Case 6
          GoTo All
        Case 7
          If LineCnt >= MaxLines - 7 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Principal:    "; Tab(55); Using$("$###,###,##0.00", GPrincTot)
          Print #RptHandle, Tab(5); "Interest:     "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising:  "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", GLateListTot)
          Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 5
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
        Case 8
          GoTo All
        Case 9
          If LineCnt >= MaxLines - 7 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Principal:    "; Tab(55); Using$("$###,###,##0.00", GPrincTot)
          Print #RptHandle, Tab(5); "Interest:     "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising:  "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", GLateListTot)
          Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 5
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
        Case 10
          If LineCnt >= MaxLines - 7 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Principal:    "; Tab(55); Using$("$###,###,##0.00", GPrincTot)
          Print #RptHandle, Tab(5); "Interest:     "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising:  "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", GLateListTot)
          Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 5
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
       Case 11 'added detail and took out GoTo All on 7/6/06
'         GoTo All
          If LineCnt >= MaxLines - 7 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Principal:    "; Tab(55); Using$("$###,###,##0.00", GPrincTot)
          Print #RptHandle, Tab(5); "Interest:     "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising:  "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", GLateListTot)
          Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 5
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
       Case 12
         GoTo All
       Case 13
          If LineCnt >= MaxLines - 7 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Principal:    "; Tab(55); Using$("$###,###,##0.00", GPrincTot)
          Print #RptHandle, Tab(5); "Interest:     "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising:  "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", GLateListTot)
          Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 5
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
        Case 14
          If LineCnt >= MaxLines - 7 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Principal:    "; Tab(55); Using$("$###,###,##0.00", GPrincTot)
          Print #RptHandle, Tab(5); "Interest:     "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising:  "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", GLateListTot)
          Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 5
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
        Case 15
          GoTo All
        Case 16
          GoTo All
        Case 17
          GoTo All
        Case Else
          GoTo All
      End Select
All:
    End If
  Next x
  Print #RptHandle, String(100, "-")
  Print #RptHandle, "Grand Totals: "; Tab(38); Using$("##,##0", GCntTot); Tab(55); Using$("$###,###,##0.00", GAmtTot)
  
  Return
  
SortItPers:
  BigYr = 0
  For x = 1 To YrCnt
    If ThEYear(x) > BigYr Then
      BigYr = ThEYear(x)
    End If
  Next x
  ReDim HoldAmt(1 To 18, 1 To YrCnt) As Double
  ReDim HoldCnt(1 To 18, 1 To YrCnt) As Integer
  
  Nexty = 1
  Nextx = 1
  HoldBigYr = 0
  For x = 1 To 18
    For y = Nexty To YrCnt
      If ThEYear(y) >= HoldBigYr Then
        HoldBigYr = ThEYear(y)
        Thisx = x
        Thisy = y
      End If
    Next y
    For z = 1 To 18
      HoldAmt(z, Thisy) = TotByYrAndType(z, Nexty)
      HoldCnt(z, Thisy) = CntByYrAndType(z, Nexty)
    Next z
    HoldYr = ThEYear(Nexty)
    HoldPers = TotByYrAndPers(Nexty)
    HoldMT = TotByYrAndMT(Nexty)
    HoldMC = TotByYrAndMC(Nexty)
    HoldFE = TotByYrAndFE(Nexty)
    HoldMH = TotByYrAndMH(Nexty)
    HoldInt = TotByYrAndInt(Nexty)
    HoldAdv = TotByYrAndAdv(Nexty)
    HoldLateList = TotByYrAndLateList(Nexty)
    HoldPen = TotByYrAndPen(Nexty)
    HoldOpt1 = TotByYrAndOpt1(Nexty)
    HoldOpt2 = TotByYrAndOpt2(Nexty)
    HoldOpt3 = TotByYrAndOpt3(Nexty)
    For z = 1 To 18
      TotByYrAndType(z, Nexty) = TotByYrAndType(z, Thisy)
      CntByYrAndType(z, Nexty) = CntByYrAndType(z, Thisy)
    Next z
    ThEYear(Nexty) = ThEYear(Thisy)
    TotByYrAndPers(Nexty) = TotByYrAndPers(Thisy)
    TotByYrAndMT(Nexty) = TotByYrAndMT(Thisy)
    TotByYrAndMC(Nexty) = TotByYrAndMC(Thisy)
    TotByYrAndFE(Nexty) = TotByYrAndFE(Thisy)
    TotByYrAndMH(Nexty) = TotByYrAndMH(Thisy)
    TotByYrAndInt(Nexty) = TotByYrAndInt(Thisy)
    TotByYrAndAdv(Nexty) = TotByYrAndAdv(Thisy)
    TotByYrAndLateList(Nexty) = TotByYrAndLateList(Thisy)
    TotByYrAndPen(Nexty) = TotByYrAndPen(Thisy)
    TotByYrAndOpt1(Nexty) = TotByYrAndOpt1(Thisy)
    TotByYrAndOpt2(Nexty) = TotByYrAndOpt2(Thisy)
    TotByYrAndOpt3(Nexty) = TotByYrAndOpt3(Thisy)
    For z = 1 To 18
      TotByYrAndType(z, Thisy) = HoldAmt(z, Thisy)
      CntByYrAndType(z, Thisy) = HoldCnt(z, Thisy)
    Next z
    ThEYear(Thisy) = HoldYr
    TotByYrAndPers(Thisy) = HoldPers
    TotByYrAndMT(Thisy) = HoldMT
    TotByYrAndMC(Thisy) = HoldMC
    TotByYrAndFE(Thisy) = HoldFE
    TotByYrAndMH(Thisy) = HoldMH
    TotByYrAndInt(Thisy) = HoldInt
    TotByYrAndAdv(Thisy) = HoldAdv
    TotByYrAndLateList(Thisy) = HoldLateList
    TotByYrAndPen(Thisy) = HoldPen
    TotByYrAndOpt1(Thisy) = HoldOpt1
    TotByYrAndOpt2(Thisy) = HoldOpt2
    TotByYrAndOpt3(Thisy) = HoldOpt3
    If Nexty >= YrCnt Then Exit For
    HoldBigYr = 0
    Nexty = Nexty + 1
  Next x
  Print #RptHandle, FF$
  GoSub PrintSortHeader
  LineCnt = LineCnt + 2
  For y = 1 To YrCnt
   If LineCnt >= MaxLines - 4 Then
     Print #RptHandle, FF$
     GoSub PrintSortHeader
     Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
   Else
     Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
   End If
   LineCnt = LineCnt + 1
   For x = 1 To 18
      If TotByYrAndType(x, y) > 0 Then
        Select Case x
          Case 1
            Print #RptHandle, "  Billing"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOneP
            If LineCnt > MaxLines - 10 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            Print #RptHandle, Tab(5); "Personal:            "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPers(y))
            Print #RptHandle, Tab(5); "Machine Tools:       "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMT(y))
            Print #RptHandle, Tab(5); "Merchant Capital:    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMC(y))
            Print #RptHandle, Tab(5); "Farm Equipment:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndFE(y))
            Print #RptHandle, Tab(5); "Mobile Homes:        "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMH(y))
            '7/11/06 added back int, adv and pen to accommodate manual bills...took out late listing
            Print #RptHandle, Tab(5); "Interest :    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Advertising : "; Tab(55); Using$("$###,###,##0.00", TotByYrAndAdv(y))
'            Print #RptHandle, Tab(5); "Late Listing:        "; Tab(55); Using$("$###,###,##0.00", TotByYrAndLateList(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 8
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 2
            Print #RptHandle, "  Payment"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt > MaxLines - 13 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOneP
            Print #RptHandle, Tab(5); "Personal:    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPers(y))
            Print #RptHandle, Tab(5); "Machine Tools:       "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMT(y))
            Print #RptHandle, Tab(5); "Merchant Capital:    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMC(y))
            Print #RptHandle, Tab(5); "Farm Equipment:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndFE(y))
            Print #RptHandle, Tab(5); "Mobile Homes:        "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMH(y))
            Print #RptHandle, Tab(5); "Interest :    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Advertising : "; Tab(55); Using$("$###,###,##0.00", TotByYrAndAdv(y))
'            Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndLateList(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 9
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 3
            Print #RptHandle, "  Release"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt > MaxLines - 13 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOneP
            Print #RptHandle, Tab(5); "Personal:    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPers(y))
            Print #RptHandle, Tab(5); "Machine Tools:       "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMT(y))
            Print #RptHandle, Tab(5); "Merchant Capital:    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMC(y))
            Print #RptHandle, Tab(5); "Farm Equipment:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndFE(y))
            Print #RptHandle, Tab(5); "Mobile Homes:        "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMH(y))
            Print #RptHandle, Tab(5); "Interest :    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Advertising : "; Tab(55); Using$("$###,###,##0.00", TotByYrAndAdv(y))
'            Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndLateList(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 9
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 4
            Print #RptHandle, "  Interest"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
          Case 5
            Print #RptHandle, "  Penalty"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
          Case 6
            Print #RptHandle, "  Advertising"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
          Case 7
            Print #RptHandle, "  Adjust Pay Down"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt > MaxLines - 13 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOneP
            Print #RptHandle, Tab(5); "Personal:            "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPers(y))
            Print #RptHandle, Tab(5); "Machine Tools:       "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMT(y))
            Print #RptHandle, Tab(5); "Merchant Capital:    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMC(y))
            Print #RptHandle, Tab(5); "Farm Equipment:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndFE(y))
            Print #RptHandle, Tab(5); "Mobile Homes:        "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMH(y))
            Print #RptHandle, Tab(5); "Interest :    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Advertising : "; Tab(55); Using$("$###,###,##0.00", TotByYrAndAdv(y))
'            Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndLateList(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 9
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 8
            Print #RptHandle, "  Credit at Billing"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
          Case 9
            Print #RptHandle, "  Adjust Bill Down"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt > MaxLines - 13 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOneP
            Print #RptHandle, Tab(5); "Personal:            "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPers(y))
            Print #RptHandle, Tab(5); "Machine Tools:       "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMT(y))
            Print #RptHandle, Tab(5); "Merchant Capital:    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMC(y))
            Print #RptHandle, Tab(5); "Farm Equipment:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndFE(y))
            Print #RptHandle, Tab(5); "Mobile Homes:        "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMH(y))
            Print #RptHandle, Tab(5); "Interest :    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Advertising : "; Tab(55); Using$("$###,###,##0.00", TotByYrAndAdv(y))
'            Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndLateList(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 9
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 10
            Print #RptHandle, "  Adjust Bill Up"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt > MaxLines - 13 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOneP
            Print #RptHandle, Tab(5); "Personal:            "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPers(y))
            Print #RptHandle, Tab(5); "Machine Tools:       "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMT(y))
            Print #RptHandle, Tab(5); "Merchant Capital:    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMC(y))
            Print #RptHandle, Tab(5); "Farm Equipment:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndFE(y))
            Print #RptHandle, Tab(5); "Mobile Homes:        "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMH(y))
            Print #RptHandle, Tab(5); "Interest :    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Advertising : "; Tab(55); Using$("$###,###,##0.00", TotByYrAndAdv(y))
'            Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndLateList(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 9
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 11
            If fpcmbTransType.Text <> " 0) All" Then
              Print #RptHandle, "  Bill OverPay"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            Else
              Print #RptHandle, "  Bill Pay/OverPay"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            End If
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            '7/6/06-----------------------------------------------------
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOneP
            Print #RptHandle, Tab(5); "Personal:            "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPers(y))
            Print #RptHandle, Tab(5); "Machine Tools:       "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMT(y))
            Print #RptHandle, Tab(5); "Merchant Capital:    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMC(y))
            Print #RptHandle, Tab(5); "Farm Equipment:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndFE(y))
            Print #RptHandle, Tab(5); "Mobile Homes:        "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMH(y))
            Print #RptHandle, Tab(5); "Interest :    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Advertising : "; Tab(55); Using$("$###,###,##0.00", TotByYrAndAdv(y))
'            Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndLateList(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 9
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
            '-------------------------------------------7/6/06
          Case 12
            Print #RptHandle, "  OverPayment"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
          Case 13
            Print #RptHandle, "  Adj Bill Up Affecting Credit Bal"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt > MaxLines - 13 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOneP
            Print #RptHandle, Tab(5); "Personal:            "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPers(y))
            Print #RptHandle, Tab(5); "Machine Tools:       "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMT(y))
            Print #RptHandle, Tab(5); "Merchant Capital:    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMC(y))
            Print #RptHandle, Tab(5); "Farm Equipment:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndFE(y))
            Print #RptHandle, Tab(5); "Mobile Homes:        "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMH(y))
            Print #RptHandle, Tab(5); "Interest :    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Advertising : "; Tab(55); Using$("$###,###,##0.00", TotByYrAndAdv(y))
'            Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndLateList(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 9
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 14
            Print #RptHandle, "  Adj Pay Dn Affecting Credit Bal"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt > MaxLines - 13 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOneP
            Print #RptHandle, Tab(5); "Personal:            "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPers(y))
            Print #RptHandle, Tab(5); "Machine Tools:       "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMT(y))
            Print #RptHandle, Tab(5); "Merchant Capital:    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMC(y))
            Print #RptHandle, Tab(5); "Farm Equipment:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndFE(y))
            Print #RptHandle, Tab(5); "Mobile Homes:        "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMH(y))
            Print #RptHandle, Tab(5); "Interest :    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Advertising : "; Tab(55); Using$("$###,###,##0.00", TotByYrAndAdv(y))
'            Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndLateList(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 9
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 15
            Print #RptHandle, "  Adj Prepay Down"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
          Case 16
            Print #RptHandle, "  Refund Prepay"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
          Case 17
            Print #RptHandle, "  PPTRA Removal"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
          Case 18
            Print #RptHandle, "  Unknown"; Tab(38); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
        End Select
      End If
NextOneP:
    Next x
    Print #RptHandle, String$(100, "-")
    Print #RptHandle,
    LineCnt = LineCnt + 2
  Next y

  Return
  
PrintTotalsPers:
  GCntTot = 0
  GAmtTot = 0
  If LineCnt > MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintTotalsHeader
  Else
    Print #RptHandle,
    Print #RptHandle, "Grand Totals"
    Print #RptHandle, Tab(35); "Trans Cnt"; Tab(64); "Amount"
    Print #RptHandle, String$(100, "-")
    LineCnt = LineCnt + 4
  End If
  For x = 1 To YrCnt
    GPersTot = GPersTot + TotByYrAndPers(x)
    GMTTot = GMTTot + TotByYrAndMT(x)
    GMCTot = GMCTot + TotByYrAndMC(x)
    GFETot = GFETot + TotByYrAndFE(x)
    GMHTot = GMHTot + TotByYrAndMH(x)
    GIntTot = GIntTot + TotByYrAndInt(x)
    GAdvTot = GAdvTot + TotByYrAndAdv(x)
    GLateListTot = GLateListTot + TotByYrAndLateList(x)
    GPenTot = GPenTot + TotByYrAndPen(x)
    GOpt1Tot = GOpt1Tot + TotByYrAndOpt1(x)
    GOpt2Tot = GOpt2Tot + TotByYrAndOpt2(x)
    GOpt3Tot = GOpt3Tot + TotByYrAndOpt3(x)
  Next x
  For x = 1 To 18
    GCntTot = GCntTot + TotCntByType(x)
    GAmtTot = GAmtTot + TotByType(x)
    If TotByType(x) > 0 Then
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintTotalsHeader
      End If
      Select Case x
        Case 1
          Print #RptHandle, "Billing";
        Case 2
          Print #RptHandle, "Payment";
        Case 3
          Print #RptHandle, "Release";
        Case 4
          Print #RptHandle, "Interest";
        Case 5
          Print #RptHandle, "Penalty";
        Case 6
          Print #RptHandle, "Advertising";
        Case 7
          Print #RptHandle, "Adjust Pay Down";
        Case 8
          Print #RptHandle, "Credit at Billing";
        Case 9
          Print #RptHandle, "Adjust Bill Down";
        Case 10
          Print #RptHandle, "Adjust Bill Up";
        Case 11
          If fpcmbTransType.Text <> " 0) All" Then 'added If statement 7/7/06
            Print #RptHandle, "Bill OverPay";
          Else
            Print #RptHandle, "Bill Pay/OverPay";
          End If
        Case 12
          Print #RptHandle, "OverPayment";
        Case 13
          Print #RptHandle, "Adjust Bill Up Affecting Credit Balance";
        Case 14
          Print #RptHandle, "Adjust Pay Dwn Affecting Credit Balance";
        Case 15
          Print #RptHandle, "Adjust Prepay Down";
        Case 16
          Print #RptHandle, "Refund Prepay";
        Case 17
          Print #RptHandle, "PPTRA Removal";
        Case 18
          Print #RptHandle, "Unknown";
      End Select
      Dim CaseP11Tot As Double 'added on 7/6/06
      If x <> 11 Or (x = 11 And fpcmbTransType.Text = " 0) All") Then '7/7/06 added the Or part
        Print #RptHandle, Tab(38); Using$("##,##0", TotCntByType(x)); Tab(55); Using$("$###,###,##0.00", TotByType(x))
      Else 'added on 7/6/06
        CaseP11Tot = OldRound(TotByType(11) - (GPersTot + GIntTot + GMTTot + GMCTot + GFETot + GMHTot + GLateListTot + GAdvTot))
        CaseP11Tot = OldRound(CaseP11Tot - (GOpt1Tot + GOpt2Tot + GOpt3Tot + GPenTot))
        Print #RptHandle, Tab(38); Using$("##,##0", TotCntByType(11)); Tab(55); Using$("$###,###,##0.00", CaseP11Tot)
      End If
      LineCnt = LineCnt + 1
      If fpcmbTransType.Text = " 0) All" Then GoTo AllP
      Select Case x
        Case 1
          If LineCnt >= MaxLines - 10 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Personal:         "; Tab(55); Using$("$###,###,##0.00", GPersTot)
          Print #RptHandle, Tab(5); "Machine Tools:    "; Tab(55); Using$("$###,###,##0.00", GMTTot)
          Print #RptHandle, Tab(5); "Merchant Capital: "; Tab(55); Using$("$###,###,##0.00", GMCTot)
          Print #RptHandle, Tab(5); "Farm Equipment:   "; Tab(55); Using$("$###,###,##0.00", GFETot)
          Print #RptHandle, Tab(5); "Mobile Homes:     "; Tab(55); Using$("$###,###,##0.00", GMHTot)
          '7/11/06 added back int, adv and pen to accommodate manual bills
          Print #RptHandle, Tab(5); "Interest:         "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising:      "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Late Listing:     "; Tab(55); Using$("$###,###,##0.00", GLateListTot)
          Print #RptHandle, Tab(5); "Penalty:          "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 9
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
        Case 2
          If LineCnt >= MaxLines - 13 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Personal:         "; Tab(55); Using$("$###,###,##0.00", GPersTot)
          Print #RptHandle, Tab(5); "Machine Tools:    "; Tab(55); Using$("$###,###,##0.00", GMTTot)
          Print #RptHandle, Tab(5); "Merchant Capital: "; Tab(55); Using$("$###,###,##0.00", GMCTot)
          Print #RptHandle, Tab(5); "Farm Equipment:   "; Tab(55); Using$("$###,###,##0.00", GFETot)
          Print #RptHandle, Tab(5); "Mobile Homes:     "; Tab(55); Using$("$###,###,##0.00", GMHTot)
          Print #RptHandle, Tab(5); "Interest:         "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising:      "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Late Listing:     "; Tab(55); Using$("$###,###,##0.00", GLateListTot)
          Print #RptHandle, Tab(5); "Penalty:          "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 9
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
        Case 3
          If LineCnt >= MaxLines - 13 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Personal:         "; Tab(55); Using$("$###,###,##0.00", GPersTot)
          Print #RptHandle, Tab(5); "Machine Tools:    "; Tab(55); Using$("$###,###,##0.00", GMTTot)
          Print #RptHandle, Tab(5); "Merchant Capital: "; Tab(55); Using$("$###,###,##0.00", GMCTot)
          Print #RptHandle, Tab(5); "Farm Equipment:   "; Tab(55); Using$("$###,###,##0.00", GFETot)
          Print #RptHandle, Tab(5); "Mobile Homes:     "; Tab(55); Using$("$###,###,##0.00", GMHTot)
          Print #RptHandle, Tab(5); "Interest:         "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising:      "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Late Listing:     "; Tab(55); Using$("$###,###,##0.00", GLateListTot)
          Print #RptHandle, Tab(5); "Penalty:          "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 9
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
        Case 4
          GoTo AllP
        Case 5
          GoTo AllP
        Case 6
          GoTo AllP
        Case 7
          If LineCnt >= MaxLines - 13 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Personal:         "; Tab(55); Using$("$###,###,##0.00", GPersTot)
          Print #RptHandle, Tab(5); "Machine Tools:    "; Tab(55); Using$("$###,###,##0.00", GMTTot)
          Print #RptHandle, Tab(5); "Merchant Capital: "; Tab(55); Using$("$###,###,##0.00", GMCTot)
          Print #RptHandle, Tab(5); "Farm Equipment:   "; Tab(55); Using$("$###,###,##0.00", GFETot)
          Print #RptHandle, Tab(5); "Mobile Homes:     "; Tab(55); Using$("$###,###,##0.00", GMHTot)
          Print #RptHandle, Tab(5); "Interest:         "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising:      "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Late Listing:     "; Tab(55); Using$("$###,###,##0.00", GLateListTot)
          Print #RptHandle, Tab(5); "Penalty:          "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 9
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
        Case 8
          GoTo AllP
        Case 9
          If LineCnt >= MaxLines - 13 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Personal:         "; Tab(55); Using$("$###,###,##0.00", GPersTot)
          Print #RptHandle, Tab(5); "Machine Tools:    "; Tab(55); Using$("$###,###,##0.00", GMTTot)
          Print #RptHandle, Tab(5); "Merchant Capital: "; Tab(55); Using$("$###,###,##0.00", GMCTot)
          Print #RptHandle, Tab(5); "Farm Equipment:   "; Tab(55); Using$("$###,###,##0.00", GFETot)
          Print #RptHandle, Tab(5); "Mobile Homes:     "; Tab(55); Using$("$###,###,##0.00", GMHTot)
          Print #RptHandle, Tab(5); "Interest:         "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising:      "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Late Listing:     "; Tab(55); Using$("$###,###,##0.00", GLateListTot)
          Print #RptHandle, Tab(5); "Penalty:          "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 9
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
        Case 10
          If LineCnt >= MaxLines - 13 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Personal:         "; Tab(55); Using$("$###,###,##0.00", GPersTot)
          Print #RptHandle, Tab(5); "Machine Tools:    "; Tab(55); Using$("$###,###,##0.00", GMTTot)
          Print #RptHandle, Tab(5); "Merchant Capital: "; Tab(55); Using$("$###,###,##0.00", GMCTot)
          Print #RptHandle, Tab(5); "Farm Equipment:   "; Tab(55); Using$("$###,###,##0.00", GFETot)
          Print #RptHandle, Tab(5); "Mobile Homes:     "; Tab(55); Using$("$###,###,##0.00", GMHTot)
          Print #RptHandle, Tab(5); "Interest:         "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising:      "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Late Listing:     "; Tab(55); Using$("$###,###,##0.00", GLateListTot)
          Print #RptHandle, Tab(5); "Penalty:          "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 9
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
       Case 11 'added detail and took out GoTo AllP on 7/6/06
'         GoTo AllP
          If LineCnt >= MaxLines - 13 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Personal:         "; Tab(55); Using$("$###,###,##0.00", GPersTot)
          Print #RptHandle, Tab(5); "Machine Tools:    "; Tab(55); Using$("$###,###,##0.00", GMTTot)
          Print #RptHandle, Tab(5); "Merchant Capital: "; Tab(55); Using$("$###,###,##0.00", GMCTot)
          Print #RptHandle, Tab(5); "Farm Equipment:   "; Tab(55); Using$("$###,###,##0.00", GFETot)
          Print #RptHandle, Tab(5); "Mobile Homes:     "; Tab(55); Using$("$###,###,##0.00", GMHTot)
          Print #RptHandle, Tab(5); "Interest:         "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising:      "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Late Listing:     "; Tab(55); Using$("$###,###,##0.00", GLateListTot)
          Print #RptHandle, Tab(5); "Penalty:          "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 9
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
       Case 12
         GoTo AllP
       Case 13
          If LineCnt >= MaxLines - 13 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Personal:         "; Tab(55); Using$("$###,###,##0.00", GPersTot)
          Print #RptHandle, Tab(5); "Machine Tools:    "; Tab(55); Using$("$###,###,##0.00", GMTTot)
          Print #RptHandle, Tab(5); "Merchant Capital: "; Tab(55); Using$("$###,###,##0.00", GMCTot)
          Print #RptHandle, Tab(5); "Farm Equipment:   "; Tab(55); Using$("$###,###,##0.00", GFETot)
          Print #RptHandle, Tab(5); "Mobile Homes:     "; Tab(55); Using$("$###,###,##0.00", GMHTot)
          Print #RptHandle, Tab(5); "Interest:         "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising:      "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Late Listing:     "; Tab(55); Using$("$###,###,##0.00", GLateListTot)
          Print #RptHandle, Tab(5); "Penalty:          "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 9
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
        Case 14
          If LineCnt >= MaxLines - 13 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Personal:         "; Tab(55); Using$("$###,###,##0.00", GPersTot)
          Print #RptHandle, Tab(5); "Machine Tools:    "; Tab(55); Using$("$###,###,##0.00", GMTTot)
          Print #RptHandle, Tab(5); "Merchant Capital: "; Tab(55); Using$("$###,###,##0.00", GMCTot)
          Print #RptHandle, Tab(5); "Farm Equipment:   "; Tab(55); Using$("$###,###,##0.00", GFETot)
          Print #RptHandle, Tab(5); "Mobile Homes:     "; Tab(55); Using$("$###,###,##0.00", GMHTot)
          Print #RptHandle, Tab(5); "Interest:         "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising:      "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Late Listing:     "; Tab(55); Using$("$###,###,##0.00", GLateListTot)
          Print #RptHandle, Tab(5); "Penalty:          "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 9
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
        Case 15
          GoTo AllP
        Case 16
          GoTo AllP
        Case 17
          GoTo AllP
        Case Else
          GoTo AllP
      End Select
AllP:
    End If
  Next x
  Print #RptHandle, String(100, "-")
  Print #RptHandle, "Grand Totals: "; Tab(38); Using$("##,##0", GCntTot); Tab(55); Using$("$###,###,##0.00", GAmtTot)
  
  Return
  
ApplyDiscP: '1/16/07
  Disc1 = 0
  Disc2 = 0
  Disc3 = 0
  Disc4 = 0
  Disc5 = 0
  Disc6 = 0
  Disc7 = 0
  Disc8 = 0
  If TaxTrans.Amount = 0 Then Return
  If TaxTrans.TranType = 1 Then
    SaveAmt = OldRound(TaxTrans.Amount - TaxTrans.DiscAmt)
  Else
    SaveAmt = TaxTrans.Amount
    TaxTrans.Amount = OldRound(TaxTrans.Amount + TaxTrans.DiscAmt)
  End If
  Disc1 = OldRound(TaxTrans.Revenue.Principle1Pd / SaveAmt)
  Disc1 = OldRound(Disc1 * TaxTrans.DiscAmt)
  Disc2 = OldRound(TaxTrans.Revenue.Principle2Pd / SaveAmt)
  Disc2 = OldRound(Disc2 * TaxTrans.DiscAmt)
  Disc3 = OldRound(TaxTrans.Revenue.Principle3Pd / SaveAmt)
  Disc3 = OldRound(Disc3 * TaxTrans.DiscAmt)
  Disc4 = OldRound(TaxTrans.Revenue.Principle4Pd / SaveAmt)
  Disc4 = OldRound(Disc4 * TaxTrans.DiscAmt)
  Disc5 = OldRound(TaxTrans.Revenue.Principle5Pd / SaveAmt)
  Disc5 = OldRound(Disc5 * TaxTrans.DiscAmt)
  Disc6 = OldRound(TaxTrans.Revenue.RevOpt1Pd / SaveAmt)
  Disc6 = OldRound(Disc6 * TaxTrans.DiscAmt)
  Disc7 = OldRound(TaxTrans.Revenue.RevOpt2Pd / SaveAmt)
  Disc7 = OldRound(Disc7 * TaxTrans.DiscAmt)
  Disc8 = OldRound(TaxTrans.Revenue.RevOpt3Pd / SaveAmt)
  Disc8 = OldRound(Disc8 * TaxTrans.DiscAmt)
  TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd + Disc1)
  TaxTrans.Revenue.Principle2Pd = OldRound(TaxTrans.Revenue.Principle2Pd + Disc2)
  TaxTrans.Revenue.Principle3Pd = OldRound(TaxTrans.Revenue.Principle3Pd + Disc3)
  TaxTrans.Revenue.Principle4Pd = OldRound(TaxTrans.Revenue.Principle4Pd + Disc4)
  TaxTrans.Revenue.Principle5Pd = OldRound(TaxTrans.Revenue.Principle5Pd + Disc5)
  TaxTrans.Revenue.RevOpt1Pd = OldRound(TaxTrans.Revenue.RevOpt1Pd + Disc6)
  TaxTrans.Revenue.RevOpt2Pd = OldRound(TaxTrans.Revenue.RevOpt2Pd + Disc7)
  TaxTrans.Revenue.RevOpt3Pd = OldRound(TaxTrans.Revenue.RevOpt3Pd + Disc8)
  DiscApplied = True
  
  Return
  
ApplyDiscR: ' 1/16/07
  Disc1 = 0
  Disc2 = 0
  Disc3 = 0
  Disc4 = 0
  If TaxTrans.Amount = 0 Then Return
  If TaxTrans.TranType = 1 Then
    SaveAmt = OldRound(TaxTrans.Amount - TaxTrans.DiscAmt)
  Else
    SaveAmt = TaxTrans.Amount
    TaxTrans.Amount = OldRound(TaxTrans.Amount + TaxTrans.DiscAmt)
  End If
  Disc1 = OldRound(TaxTrans.Revenue.Principle1Pd / SaveAmt)
  Disc1 = OldRound(Disc1 * TaxTrans.DiscAmt)
  Disc2 = OldRound(TaxTrans.Revenue.RevOpt1Pd / SaveAmt)
  Disc2 = OldRound(Disc2 * TaxTrans.DiscAmt)
  Disc3 = OldRound(TaxTrans.Revenue.RevOpt2Pd / SaveAmt)
  Disc3 = OldRound(Disc3 * TaxTrans.DiscAmt)
  Disc4 = OldRound(TaxTrans.Revenue.RevOpt3Pd / SaveAmt)
  Disc4 = OldRound(Disc4 * TaxTrans.DiscAmt)
  TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd + Disc1)
  TaxTrans.Revenue.RevOpt1Pd = OldRound(TaxTrans.Revenue.RevOpt1Pd + Disc2)
  TaxTrans.Revenue.RevOpt2Pd = OldRound(TaxTrans.Revenue.RevOpt2Pd + Disc3)
  TaxTrans.Revenue.RevOpt3Pd = OldRound(TaxTrans.Revenue.RevOpt3Pd + Disc4)
  DiscApplied = True
  
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxTransJournal", "PrintText", Erl)
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

Private Sub PrintRGraphicsDet()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim x As Long, y As Integer
  Dim BegDate As Integer
  Dim EndDate As Integer
  Dim dlm$
  Dim ThisRec As Long
  Dim ThisClass As Integer
  Dim ThisType As String
  Dim RptFile$
  Dim RptHandle As Integer
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim InactiveFlag As Boolean
  Dim ThisName$, ThisBillType$
  Dim TCnt As Long
  Dim TotAmt As Double
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim NumOfSrchRecs As Long
  Dim ThisTransType As String
  Dim YrCnt As Integer, ThisYear As Integer
  Dim SubRptFile$
  Dim SubRptHandle As Integer
  Dim Sub2RptFile$
  Dim Sub2RptHandle As Integer
  Dim BigYr As Integer
  Dim HoldBigYr As Integer
  Dim HoldYr As Integer
  Dim HoldPrinc As Double
  Dim HoldInt As Double
  Dim HoldAdv As Double
  Dim HoldLateList As Double
  Dim HoldPen As Double
  Dim HoldOpt1 As Double
  Dim HoldOpt2 As Double
  Dim HoldOpt3 As Double
  Dim Nextx As Integer
  Dim Thisx As Integer
  Dim Nexty As Integer
  Dim Thisy As Integer
  Dim z As Integer
  Dim PrincDif As Double
  Dim IntDif As Double
  Dim AdvDif As Double
  Dim LateListDif As Double
  Dim PenDif As Double
  Dim Opt1Dif As Double, BillCnt As Integer
  Dim Opt2Dif As Double, ThisBillRec As Long
  Dim Opt3Dif As Double
'  Dim Opt1Desc$
'  Dim Opt2Desc$
'  Dim Opt3Desc$
  Dim POpt1Desc$
  Dim POpt2Desc$
  Dim POpt3Desc$
  Dim CustBal As Double
  Dim BillBal As Double
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim OptFlag As Boolean
  Dim TransDesc$
  Dim GAmtTot As Double
  Dim GCntTot As Long
  Dim GBillCnt As Long
  Dim BadCnt As Integer
  Dim QFlag As Boolean
  Dim GPrincTot As Double
  Dim GIntTot As Double
  Dim GAdvTot As Double
  Dim GLateListTot As Double
  Dim GPenTot As Double
  Dim GOpt1Tot As Double
  Dim GOpt2Tot As Double
  Dim GOpt3Tot As Double
  Dim TotYearCnt As Long
  Dim TotYearAmt As Double
  Dim PrePayDone As Boolean 'added 7/10/06
  Dim Disc1 As Double '1/16/2007
  Dim Disc2 As Double '1/16/2007
  Dim Disc3 As Double '1/16/2007
  Dim Disc4 As Double '1/16/2007
  Dim DiscApplied As Boolean '1/16/2007
  Dim SaveAmt As Double '1/16/2007
  
  On Error GoTo ERRORSTUFF
  
  PrePayDone = False '7/10/06
  QFlag = False
  If chkQuick.Value = 1 Then QFlag = True
  DiscApplied = False '1/16/07
  TransDesc = QPTrim$(fptxtDesc.Text)
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  Opt1Desc = QPTrim$(TaxMasterRec.OptRev1)
  Opt2Desc = QPTrim$(TaxMasterRec.OptRev2)
  Opt3Desc = QPTrim$(TaxMasterRec.OptRev3)
  POpt1Desc = QPTrim$(TaxMasterRec.POptRev1)
  POpt2Desc = QPTrim$(TaxMasterRec.POptRev2)
  POpt3Desc = QPTrim$(TaxMasterRec.POptRev3)
  IdxFlag = False
  OptFlag = False
  
  If CheckB4Printing = False Then
    Exit Sub
  End If
  
  ThisBillType = QPTrim$(fpcmbTaxType.Text)
  If fpcmbIncInactive.Text = "No" Then
    InactiveFlag = False
  Else
    InactiveFlag = True
  End If
  
  dlm$ = "~"
  If Mid(fpcmbTransType.Text, 1, 1) = "" Then
    ThisClass = CInt(Mid(fpcmbTransType.Text, 2, 1))
  Else
    ThisClass = CInt(Mid(fpcmbTransType.Text, 1, 2))
  End If
    
  Select Case ThisClass
    Case 1
      ThisType = "Billing"
    Case 2
      ThisType = "Payment"
    Case 3
      ThisType = "Release"
    Case 4
      ThisType = "Interest"
    Case 5
      ThisType = "Penalty"
    Case 6
      ThisType = "Advertising Charge"
    Case 7, 10
      ThisType = "Adjust Pay Down"
    Case 9
      ThisType = "Credit At Billing"
    Case 11
      ThisType = "Adjust Prepay Down"
    Case 12
      ThisType = "Refund Prepay"
    Case 13
      ThisType = "Adjust Bill Down"
    Case 14, 24
      ThisType = "Adjust Bill Up"
    Case 21
      ThisType = "Payment w/Overpay"
    Case 22
      ThisType = "Overpayment Only"
    Case 30
      ThisType = "PPTRA Removal"
    Case Else
      ThisType = "All"
  End Select
    
  BegDate = Date2Num(fptxtBegDate.Text)
  EndDate = Date2Num(fptxtEndDate.Text)
  
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
    OptFlag = True
  End If
  
  RptFile$ = "TAXRPTS\TXJRLDT.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  
  frmVATaxShowPctComp.Label1 = "Gathering Tax Transaction Data"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  
  ReDim TotByYrAndType(1 To 18, 1 To 1) As Double
  ReDim CntByYrAndType(1 To 18, 1 To 1) As Integer
  ReDim TotByType(1 To 18) As Double
  ReDim TotCntByType(1 To 18) As Long
  ReDim ThEYear(1 To 1) As Integer
  
  ReDim TotByYrAndPrinc(1 To 1) As Double
  ReDim TotByYrAndInt(1 To 1) As Double
  ReDim TotByYrAndAdv(1 To 1) As Double
  ReDim TotByYrAndLateList(1 To 1) As Double
  ReDim TotByYrAndPen(1 To 1) As Double
  ReDim TotByYrAndOpt1(1 To 1) As Double
  ReDim TotByYrAndOpt2(1 To 1) As Double
  ReDim TotByYrAndOpt3(1 To 1) As Double
  
  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  End If
  If InStr(fpcmbTransType.Text, "All") Then GoTo PrintAll
  If fpcmbTransType.Text <> "All" Then
    For x = 1 To NumOfTCRecs
      If IdxFlag = False Then
        Get TCHandle, x, TaxCust
      Else
        Get TCHandle, IdxArray(x), TaxCust
      End If
      If TaxCust.Active = "N" And InactiveFlag = False Then
        GoTo SkipIt
      End If
      ThisName = QPTrim$(TaxCust.CustName)
      ThisRec = TaxCust.LastTrans
      BadCnt = 0
      Do While ThisRec > 0
        Get TTHandle, ThisRec, TaxTrans
        If QFlag = True Then
          If TaxTrans.TransDate < BegDate Then
            BadCnt = BadCnt + 1
            If BadCnt > 3 Then Exit Do
          End If
        End If
        If TransDesc <> "" Then
          If InStr(1, TaxTrans.Description, TransDesc) = 0 Then GoTo SkipIt
        End If
        If TaxTrans.BillType <> "R" And TaxTrans.BillType <> "P" Then TaxTrans.BillType = ""
        If TaxTrans.BillType <> Mid(fpcmbTaxType.Text, 1, 1) And QPTrim$(TaxTrans.BillType) <> "" Then
          GoTo SkipIt
        End If
        If ThisClass = 7 And TaxTrans.TranType = 10 Then GoTo ItsOK
        If ThisClass = 14 And TaxTrans.TranType = 24 Then GoTo ItsOK
'        If ThisClass = 2 And TaxTrans.TranType = 21 Then GoTo ItsOK '7/6/06 commented out
        If ThisClass <> 0 And TaxTrans.TranType <> ThisClass Then GoTo SkipIt
ItsOK:
        If TaxTrans.TransDate >= BegDate And TaxTrans.TransDate <= EndDate Then
          If TaxTrans.BelongTo > 0 Then
            Get TTHandle, TaxTrans.BelongTo, TaxTrans
          End If
        
          If YrCnt = 0 Then
            YrCnt = YrCnt + 1
            ThisYear = YrCnt
            ReDim Preserve ThEYear(1 To YrCnt) As Integer
            ThEYear(YrCnt) = TaxTrans.TaxYear
            ReDim Preserve TotByYrAndType(1 To 18, 1 To YrCnt) As Double
            ReDim Preserve CntByYrAndType(1 To 18, 1 To YrCnt) As Integer
            ReDim Preserve TotByYrAndPrinc(1 To YrCnt) As Double
            ReDim Preserve TotByYrAndInt(1 To YrCnt) As Double
            ReDim Preserve TotByYrAndAdv(1 To YrCnt) As Double
            ReDim Preserve TotByYrAndLateList(1 To YrCnt) As Double
            ReDim Preserve TotByYrAndPen(1 To YrCnt) As Double
            ReDim Preserve TotByYrAndOpt1(1 To YrCnt) As Double
            ReDim Preserve TotByYrAndOpt2(1 To YrCnt) As Double
            ReDim Preserve TotByYrAndOpt3(1 To YrCnt) As Double
            
            TotByYrAndPrinc(YrCnt) = 0
            TotByYrAndInt(YrCnt) = 0
            TotByYrAndAdv(YrCnt) = 0
            TotByYrAndLateList(YrCnt) = 0
            TotByYrAndPen(YrCnt) = 0
            TotByYrAndOpt1(YrCnt) = 0
            TotByYrAndOpt2(YrCnt) = 0
            TotByYrAndOpt3(YrCnt) = 0
           
            For y = 1 To 18
              TotByYrAndType(y, YrCnt) = 0
              CntByYrAndType(y, YrCnt) = 0
            Next y
          Else
            For y = 1 To YrCnt
              If TaxTrans.TaxYear = ThEYear(y) Then
                ThisYear = y
                Exit For
              End If
            Next y
            If y > YrCnt Then
              YrCnt = YrCnt + 1
              ThisYear = YrCnt
              ReDim Preserve ThEYear(1 To YrCnt) As Integer
              ThEYear(YrCnt) = TaxTrans.TaxYear
              ReDim Preserve TotByYrAndType(1 To 18, 1 To YrCnt) As Double
              ReDim Preserve CntByYrAndType(1 To 18, 1 To YrCnt) As Integer
              ReDim Preserve TotByYrAndPrinc(1 To YrCnt) As Double
              ReDim Preserve TotByYrAndInt(1 To YrCnt) As Double
              ReDim Preserve TotByYrAndAdv(1 To YrCnt) As Double
              ReDim Preserve TotByYrAndLateList(1 To YrCnt) As Double
              ReDim Preserve TotByYrAndPen(1 To YrCnt) As Double
              ReDim Preserve TotByYrAndOpt1(1 To YrCnt) As Double
              ReDim Preserve TotByYrAndOpt2(1 To YrCnt) As Double
              ReDim Preserve TotByYrAndOpt3(1 To YrCnt) As Double
            
              TotByYrAndPrinc(YrCnt) = 0
              TotByYrAndInt(YrCnt) = 0
              TotByYrAndAdv(YrCnt) = 0
              TotByYrAndLateList(YrCnt) = 0
              TotByYrAndPen(YrCnt) = 0
              TotByYrAndOpt1(YrCnt) = 0
              TotByYrAndOpt2(YrCnt) = 0
              TotByYrAndOpt3(YrCnt) = 0
              For y = 1 To 18
                TotByYrAndType(y, YrCnt) = 0
                CntByYrAndType(y, YrCnt) = 0
              Next y
           End If
         End If
         Get TTHandle, ThisRec, TaxTrans
          
         Select Case TaxTrans.TranType
           Case 1
             ThisTransType = "Billing"
             TotByYrAndType(1, ThisYear) = OldRound(TotByYrAndType(1, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(1, ThisYear) = OldRound(CntByYrAndType(1, ThisYear) + 1)
             TotByType(1) = OldRound(TotByType(1) + TaxTrans.Amount)
             TotCntByType(1) = OldRound(TotCntByType(1) + 1)
             TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1)
             '7/11/06 added back interest, advertising and penalty to accommodate manual bills
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.Interest)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.Collection)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.Penalty)
             
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateList)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3)
           Case 2
             ThisTransType = "Payment"
             If TaxTrans.DiscAmt > 0 Then 'added 1/16/07
               GoSub ApplyDisc
             End If
             TotByYrAndType(2, ThisYear) = OldRound(TotByYrAndType(2, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(2, ThisYear) = OldRound(CntByYrAndType(2, ThisYear) + 1)
             TotByType(2) = OldRound(TotByType(2) + TaxTrans.Amount)
             TotCntByType(2) = OldRound(TotCntByType(2) + 1)
             TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
           Case 3
            '7/12/06 changed revenues for Release to paid from charged
             ThisTransType = "Release"
             TotByYrAndType(3, ThisYear) = OldRound(TotByYrAndType(3, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(3, ThisYear) = OldRound(CntByYrAndType(3, ThisYear) + 1)
             TotByType(3) = OldRound(TotByType(3) + TaxTrans.Amount)
             TotCntByType(3) = OldRound(TotCntByType(3) + 1)
             TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
           Case 4
             ThisTransType = "Interest"
             TotByYrAndType(4, ThisYear) = OldRound(TotByYrAndType(4, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(4, ThisYear) = OldRound(CntByYrAndType(4, ThisYear) + 1)
             TotByType(4) = OldRound(TotByType(4) + TaxTrans.Amount)
             TotCntByType(4) = OldRound(TotCntByType(4) + 1)
           Case 5
             ThisTransType = "Penalty"
             TotByYrAndType(5, ThisYear) = OldRound(TotByYrAndType(5, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(5, ThisYear) = OldRound(CntByYrAndType(5, ThisYear) + 1)
             TotByType(5) = OldRound(TotByType(5) + TaxTrans.Amount)
             TotCntByType(5) = OldRound(TotCntByType(5) + 1)
           Case 6
             ThisTransType = "Advertising Charge"
             TotByYrAndType(6, ThisYear) = OldRound(TotByYrAndType(6, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(6, ThisYear) = OldRound(CntByYrAndType(6, ThisYear) + 1)
             TotByType(6) = OldRound(TotByType(6) + TaxTrans.Amount)
             TotCntByType(6) = OldRound(TotCntByType(6) + 1)
           Case 7
             ThisTransType = "Adjust Pay Down"
             TotByYrAndType(7, ThisYear) = OldRound(TotByYrAndType(7, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(7, ThisYear) = OldRound(CntByYrAndType(7, ThisYear) + 1)
             TotByType(7) = OldRound(TotByType(7) + TaxTrans.Amount)
             TotCntByType(7) = OldRound(TotCntByType(7) + 1)
             TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
           Case 9
             ThisTransType = "Credit Applied at Billing"
             TotByYrAndType(8, ThisYear) = OldRound(TotByYrAndType(8, ThisYear) + TaxTrans.Revenue.PrePaidUsed)
             CntByYrAndType(8, ThisYear) = OldRound(CntByYrAndType(8, ThisYear) + 1)
             TotByType(8) = OldRound(TotByType(8) + TaxTrans.Revenue.PrePaidUsed)
             TotCntByType(8) = OldRound(TotCntByType(8) + 1)
           Case 13
             ThisTransType = "Adjust Bill Down"
             TotByYrAndType(9, ThisYear) = OldRound(TotByYrAndType(9, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(9, ThisYear) = OldRound(CntByYrAndType(9, ThisYear) + 1)
             TotByType(9) = OldRound(TotByType(9) + TaxTrans.Amount)
             TotCntByType(9) = OldRound(TotCntByType(9) + 1)
             TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.Interest)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.Collection)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateList)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.Penalty)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3)
           Case 14
             ThisTransType = "Adjust Bill Up"
             TotByYrAndType(10, ThisYear) = OldRound(TotByYrAndType(10, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(10, ThisYear) = OldRound(CntByYrAndType(10, ThisYear) + 1)
             TotByType(10) = OldRound(TotByType(10) + TaxTrans.Amount)
             TotCntByType(10) = OldRound(TotCntByType(10) + 1)
             TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.Interest)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.Collection)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateList)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3)
           Case 21
             ThisTransType = "Billpay/Overpay" '7/6/06 changed Amount to PrePaidAmt
             If TaxTrans.DiscAmt > 0 Then 'added 1/16/07
               GoSub ApplyDisc
             End If
             If fpcmbTransType.Text <> " 0) All" Then 'added the All if statement on 7/7/06
               TotByYrAndType(11, ThisYear) = OldRound(TotByYrAndType(11, ThisYear) + TaxTrans.Revenue.PrePaidAmt)
             Else
               TotByYrAndType(11, ThisYear) = OldRound(TotByYrAndType(11, ThisYear) + TaxTrans.Amount)
             End If
'             TotByYrAndType(11, ThisYear) = OldRound(TotByYrAndType(11, ThisYear) + TaxTrans.Revenue.PrePaidAmt) '  .Amount)
             CntByYrAndType(11, ThisYear) = OldRound(CntByYrAndType(11, ThisYear) + 1)
             TotByType(11) = OldRound(TotByType(11) + TaxTrans.Amount)
             TotCntByType(11) = OldRound(TotCntByType(11) + 1)
             TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
           Case 22
             ThisTransType = "Overpayment"
             TotByYrAndType(12, ThisYear) = OldRound(TotByYrAndType(12, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(12, ThisYear) = OldRound(CntByYrAndType(12, ThisYear) + 1)
             TotByType(12) = OldRound(TotByType(12) + TaxTrans.Amount)
             TotCntByType(12) = OldRound(TotCntByType(12) + 1)
           Case 24 'go to here
             TotByYrAndType(13, ThisYear) = OldRound(TotByYrAndType(13, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(13, ThisYear) = OldRound(CntByYrAndType(13, ThisYear) + 1)
             TotByType(13) = OldRound(TotByType(13) + TaxTrans.Amount)
             TotCntByType(13) = OldRound(TotCntByType(13) + 1)
             TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
'             ThisTransType = "Adjust Bill Up Affecting Credit Balance"
'             TotByYrAndType(13, ThisYear) = OldRound(TotByYrAndType(13, ThisYear) + TaxTrans.Amount)
'             CntByYrAndType(13, ThisYear) = OldRound(CntByYrAndType(13, ThisYear) + 1)
'             TotByType(13) = OldRound(TotByType(13) + TaxTrans.Amount)
'             TotCntByType(13) = OldRound(TotCntByType(13) + 1)
'             TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1)
'             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.Interest)
'             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.Collection)
'             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateList)
'             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
'             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1)
'             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2)
'             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3)
           Case 10 ', 24 '7/11/06 added Pd on revenues
             ThisTransType = "Adjust Pay Down Affecting Credit Balance"
             TotByYrAndType(14, ThisYear) = OldRound(TotByYrAndType(14, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(14, ThisYear) = OldRound(CntByYrAndType(14, ThisYear) + 1)
             TotByType(14) = OldRound(TotByType(14) + TaxTrans.Amount)
             TotCntByType(14) = OldRound(TotCntByType(14) + 1)
             TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
           Case 11
             ThisTransType = "Adjust Prepay Down"
             TotByYrAndType(15, ThisYear) = OldRound(TotByYrAndType(15, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(15, ThisYear) = OldRound(CntByYrAndType(15, ThisYear) + 1)
             TotByType(15) = OldRound(TotByType(15) + TaxTrans.Amount)
             TotCntByType(15) = OldRound(TotCntByType(15) + 1)
           Case 12
             ThisTransType = "Refund Prepay"
             TotByYrAndType(16, ThisYear) = OldRound(TotByYrAndType(16, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(16, ThisYear) = OldRound(CntByYrAndType(16, ThisYear) + 1)
             TotByType(16) = OldRound(TotByType(16) + TaxTrans.Amount)
             TotCntByType(16) = OldRound(TotCntByType(16) + 1)
           Case 30
             ThisTransType = "PPTRA Removal"
             TotByYrAndType(17, ThisYear) = OldRound(TotByYrAndType(17, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(17, ThisYear) = OldRound(CntByYrAndType(17, ThisYear) + 1)
             TotByType(17) = OldRound(TotByType(17) + TaxTrans.Amount)
             TotCntByType(17) = OldRound(TotCntByType(17) + 1)
           Case Else
             ThisTransType = "Unknown"
             TotByYrAndType(18, ThisYear) = OldRound(TotByYrAndType(18, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(18, ThisYear) = OldRound(CntByYrAndType(18, ThisYear) + 1)
             TotByType(18) = OldRound(TotByType(18) + TaxTrans.Amount)
             TotCntByType(18) = OldRound(TotCntByType(18) + 1)
          End Select
          TCnt = TCnt + 1
          If TaxTrans.DiscAmt > 0 Then GoSub ApplyDisc 'added 1/16/07
          TotAmt = OldRound(TotAmt + TaxTrans.Amount)
          
          '------------------------------------------------------------------
          PrincDif = OldRound(TaxTrans.Revenue.Principle1 - TaxTrans.Revenue.Principle1Pd)
          IntDif = OldRound(TaxTrans.Revenue.Interest - TaxTrans.Revenue.InterestPd)
          AdvDif = OldRound(TaxTrans.Revenue.Collection - TaxTrans.Revenue.CollectionPd)
          LateListDif = OldRound(TaxTrans.Revenue.LateList - TaxTrans.Revenue.LateListPd)
          PenDif = OldRound(TaxTrans.Revenue.Penalty - TaxTrans.Revenue.PenaltyPd)
          Opt1Dif = OldRound(TaxTrans.Revenue.RevOpt1 - TaxTrans.Revenue.RevOpt1Pd)
          Opt2Dif = OldRound(TaxTrans.Revenue.RevOpt2 - TaxTrans.Revenue.RevOpt2Pd)
          Opt3Dif = OldRound(TaxTrans.Revenue.RevOpt3 - TaxTrans.Revenue.RevOpt3Pd)
      
          '                   0            1                 2                   3
          Print #RptHandle, Town$; dlm; ThisName; dlm; TaxCust.Acct; dlm; TaxCust.Active; dlm;
          '                                 4                           5                6
          Print #RptHandle, MakeRegDate(TaxTrans.TransDate); dlm; ThisBillType; dlm; ThisType; dlm;
          If TaxTrans.BelongTo > 0 Then
            Get TTHandle, TaxTrans.BelongTo, TaxTrans
              '                          7                         8                          9
              Print #RptHandle, MakeRegDate(BegDate); dlm; MakeRegDate(EndDate); dlm; TaxTrans.TaxYear; dlm;
              Get TTHandle, ThisRec, TaxTrans
              If TaxTrans.DiscAmt > 0 Then GoSub ApplyDisc 'added 1/16/07
          Else
            '                          7                         8                          9
            Print #RptHandle, MakeRegDate(BegDate); dlm; MakeRegDate(EndDate); dlm; TaxTrans.TaxYear; dlm;
          End If
          If TaxTrans.TranType <> 9 Then
            '                      10                11          12                       13
            Print #RptHandle, TaxTrans.Amount; dlm; TCnt; dlm; TotAmt; dlm; TaxTrans.Revenue.PrePaidAmt; dlm;
          Else
            '                      10                11          12                       13
            Print #RptHandle, TaxTrans.Revenue.PrePaidUsed; dlm; TCnt; dlm; TotAmt; dlm; TaxTrans.Revenue.PrePaidAmt; dlm;
          End If
          If TaxTrans.BelongTo > 0 Then
            Get TTHandle, TaxTrans.BelongTo, TaxTrans
            '                             14
            Print #RptHandle, ParseBillNum(TaxTrans.Description); dlm;
          Else
            '                 14
            Print #RptHandle, 0; dlm;
          End If
          Get TTHandle, ThisRec, TaxTrans
          If TaxTrans.DiscAmt > 0 Then GoSub ApplyDisc 'added 1/16/07
          '                                15                        16
          Print #RptHandle, QPTrim$(TaxTrans.Description); dlm; ThisTransType; dlm;
          '                               17                              18                           19
          Print #RptHandle, TaxTrans.Revenue.Principle1; dlm; TaxTrans.Revenue.Principle1Pd; ; dlm; PrincDif; dlm;
          '                               20                              21                      22
          Print #RptHandle, TaxTrans.Revenue.Interest; dlm; TaxTrans.Revenue.InterestPd; ; dlm; IntDif; dlm;
          '                               23                              24                        25
          Print #RptHandle, TaxTrans.Revenue.Collection; dlm; TaxTrans.Revenue.CollectionPd; dlm; AdvDif; dlm;
          '                               26                              27                       28
          Print #RptHandle, TaxTrans.Revenue.LateList; dlm; TaxTrans.Revenue.LateListPd; dlm; LateListDif; dlm;
          '                               29                              30                   31
          Print #RptHandle, TaxTrans.Revenue.RevOpt1; dlm; TaxTrans.Revenue.RevOpt1Pd; dlm; Opt1Dif; dlm;
          '                               32                              33                   34
          Print #RptHandle, TaxTrans.Revenue.RevOpt2; dlm; TaxTrans.Revenue.RevOpt2Pd; dlm; Opt2Dif; dlm;
          '                               35                              36                   37
          Print #RptHandle, TaxTrans.Revenue.RevOpt3; dlm; TaxTrans.Revenue.RevOpt3Pd; dlm; Opt3Dif; dlm;
          '                    38             39             40
          Print #RptHandle, Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm;
          If UseOpt = "Y" Then
            '                    41                     42
            Print #RptHandle, ThisOpt; dlm; QPTrim$(TaxCust.OptSrchDesc); dlm;
          Else
            '                 41       42
            Print #RptHandle, ""; dlm; ""; dlm;
          End If
          '                             43                                 44                       45
          Print #RptHandle, TaxTrans.Revenue.Principle4; dlm; TaxTrans.Revenue.Principle4Pd; dlm; 0; dlm;
          '                             46                              47                         48
          Print #RptHandle, TaxTrans.Revenue.Principle5; dlm; TaxTrans.Revenue.Principle5Pd; dlm; 0; dlm;
          '                             49                              50                    51               52
          Print #RptHandle, TaxTrans.Revenue.Penalty; dlm; TaxTrans.Revenue.PenaltyPd; dlm; PenDif; dlm; TaxTrans.OperNum
        End If
SkipIt:
        ThisRec = TaxTrans.LastTrans
      Loop
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
  End If
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  
  Close
  If TCnt = 0 Then
    Call TaxMsg(900, "No transactions were found that fit the parameters entered.")
    Close
    Exit Sub
  End If
  
  If YrCnt > 0 Then
    GoSub PrintSub
    GoSub PrintTotals
  End If
  
  arVATaxTransJournalDet.Show
  
  Exit Sub
  
PrintSub:
  SubRptFile$ = "TAXRPTS\SUBTXJRLDETREAL.RPT"
  SubRptHandle = FreeFile
  Open SubRptFile For Output As #SubRptHandle
  BigYr = 0
  For x = 1 To YrCnt
    If ThEYear(x) > BigYr Then
      BigYr = ThEYear(x)
    End If
  Next x
  ReDim HoldAmt(1 To 18, 1 To YrCnt) As Double
  ReDim HoldCnt(1 To 18, 1 To YrCnt) As Integer
  Nexty = 1
  Nextx = 1
  HoldBigYr = 0
  For x = 1 To 18
    For y = Nexty To YrCnt
      If ThEYear(y) >= HoldBigYr Then
        HoldBigYr = ThEYear(y)
        Thisx = x
        Thisy = y
      End If
    Next y
    For z = 1 To 18
      HoldAmt(z, Thisy) = TotByYrAndType(z, Nexty)
      HoldCnt(z, Thisy) = CntByYrAndType(z, Nexty)
    Next z
    HoldYr = ThEYear(Nexty)
    If fpcmbTransType.Text = " 0) All" Then GoTo All1
    HoldPrinc = TotByYrAndPrinc(Nexty)
    HoldInt = TotByYrAndInt(Nexty)
    HoldAdv = TotByYrAndAdv(Nexty)
    HoldLateList = TotByYrAndLateList(Nexty)
    HoldPen = TotByYrAndPen(Nexty)
    HoldOpt1 = TotByYrAndOpt1(Nexty)
    HoldOpt2 = TotByYrAndOpt2(Nexty)
    HoldOpt3 = TotByYrAndOpt3(Nexty)
All1:
    For z = 1 To 18
      TotByYrAndType(z, Nexty) = TotByYrAndType(z, Thisy)
      CntByYrAndType(z, Nexty) = CntByYrAndType(z, Thisy)
    Next z
    ThEYear(Nexty) = ThEYear(Thisy)
    If fpcmbTransType.Text = " 0) All" Then GoTo All2
    TotByYrAndPrinc(Nexty) = TotByYrAndPrinc(Thisy)
    TotByYrAndInt(Nexty) = TotByYrAndInt(Thisy)
    TotByYrAndAdv(Nexty) = TotByYrAndAdv(Thisy)
    TotByYrAndLateList(Nexty) = TotByYrAndLateList(Thisy)
    TotByYrAndPen(Nexty) = TotByYrAndPen(Thisy)
    TotByYrAndOpt1(Nexty) = TotByYrAndOpt1(Thisy)
    TotByYrAndOpt2(Nexty) = TotByYrAndOpt2(Thisy)
    TotByYrAndOpt3(Nexty) = TotByYrAndOpt3(Thisy)
All2:
    For z = 1 To 18
      TotByYrAndType(z, Thisy) = HoldAmt(z, Thisy)
      CntByYrAndType(z, Thisy) = HoldCnt(z, Thisy)
    Next z
    ThEYear(Thisy) = HoldYr
    If fpcmbTransType.Text = " 0) All" Then GoTo All3
    TotByYrAndPrinc(Thisy) = HoldPrinc
    TotByYrAndInt(Thisy) = HoldInt
    TotByYrAndAdv(Thisy) = HoldAdv
    TotByYrAndLateList(Thisy) = HoldLateList
    TotByYrAndPen(Thisy) = HoldPen
    TotByYrAndOpt1(Thisy) = HoldOpt1
    TotByYrAndOpt2(Thisy) = HoldOpt2
    TotByYrAndOpt3(Thisy) = HoldOpt3
All3:
    If Nexty >= YrCnt Then Exit For
    HoldBigYr = 0
    Nexty = Nexty + 1
  Next x
  
  For y = 1 To YrCnt
    TotYearCnt = 0
    TotYearAmt = 0
    For x = 1 To 18
      If TotByYrAndType(x, y) > 0 Then
        TotYearCnt = TotYearCnt + CntByYrAndType(x, y)
        TotYearAmt = TotYearAmt + TotByYrAndType(x, y)
        Select Case x
          Case 1
            Print #SubRptHandle, "Billing"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
              Print #SubRptHandle, TotByYrAndPrinc(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10                 11         12             13             14
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm;
              '                          15
              Print #SubRptHandle, TotByYrAndPen(y); dlm; 0; dlm; 0; dlm; "N"
            Else
              '                    4       5       6       7       8       9       10      11      12        13       14       15           16               17            18
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 2
            '                        0                1                    2                         3
            Print #SubRptHandle, "Payment"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
              Print #SubRptHandle, TotByYrAndPrinc(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10                 11         12             13              14
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm;
              '                           15
              Print #SubRptHandle, TotByYrAndPen(y); dlm; 0; dlm; 0; dlm; "N"
            Else
              '                    4       5       6       7       8       9       10      11       12       13       14       15
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 3
            Print #SubRptHandle, "Release"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
              Print #SubRptHandle, TotByYrAndPrinc(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10                 11         12             13              14
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm;
              '                           15
              Print #SubRptHandle, TotByYrAndPen(y); dlm; 0; dlm; 0; dlm; "N"
            Else
              '                    4       5       6       7       8       9       10      11       12       13       14       15
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 4
            Print #SubRptHandle, "Interest"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                    4       5       6       7       8       9       10      11      12       13       14       15
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 5
            Print #SubRptHandle, "Penalty"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                    4       5       6       7       8       9       10      11      12       13       14       15
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 6
            Print #SubRptHandle, "Advertising"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                    4       5       6       7       8       9       10      11      12       13       14       15
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 7
            Print #SubRptHandle, "Adjust Pay Down"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
              Print #SubRptHandle, TotByYrAndPrinc(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10                 11         12             13              14
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm;
              '                           15
              Print #SubRptHandle, TotByYrAndPen(y); dlm; 0; dlm; 0; dlm; "N"
            Else
              '                    4       5       6       7       8       9       10      11       12       13       14       15
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 8
            Print #SubRptHandle, "Credit at Billing"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                    4       5       6       7       8       9       10      11      12       13       14       15
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 9
            Print #SubRptHandle, "Adjust Bill Down"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
              Print #SubRptHandle, TotByYrAndPrinc(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10                 11         12             13              14
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm;
              '                           15
              Print #SubRptHandle, TotByYrAndPen(y); dlm; 0; dlm; 0; dlm; "N"
              'Print #SubRptHandle, TotByYrAndPen(y); dlm; 0; dlm; 0; dlm; "N"
            Else
              '                    4       5       6       7       8       9       10      11       12       13       14       15
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 10
            Print #SubRptHandle, "Adjust Bill Up"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
              Print #SubRptHandle, TotByYrAndPrinc(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10                 11         12             13              14
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm;
              '                           15
              Print #SubRptHandle, TotByYrAndPen(y); dlm; 0; dlm; 0; dlm; "N"
            Else
              '                    4       5       6       7       8       9       10      11       12       13       14       15
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 11
            If fpcmbTransType.Text <> " 0) All" Then 'added Bill Pay/OverPay on 7/7/06
              Print #SubRptHandle, "Bill OverPay"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
              '                            4                      5                      6                         7
              Print #SubRptHandle, TotByYrAndPrinc(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10                 11         12             13              14
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm;
              '                           15
              Print #SubRptHandle, TotByYrAndPen(y); dlm; 0; dlm; 0; dlm; "N"
            Else
              Print #SubRptHandle, "Bill Pay/OverPay"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
              '                    4       5       6       7       8       9       10      11       12       13       14       15
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 12
            Print #SubRptHandle, "OverPayment"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                    4       5       6       7       8       9       10      11      12       13       14       15
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 13
            Print #SubRptHandle, "Adjust Bill Up Affecting Credit Balance"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
              Print #SubRptHandle, TotByYrAndPrinc(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10                 11         12             13             14
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm;
              '                          15
              Print #SubRptHandle, TotByYrAndPen(y); dlm; 0; dlm; 0; dlm; "N"
            Else
              '                    4       5       6       7       8       9       10      11      12        13       14       15
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 14
            Print #SubRptHandle, "Adjust Pay Dn Affecting Credit Balance"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
              Print #SubRptHandle, TotByYrAndPrinc(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10                 11         12             13             14
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm;
              '                          15
              Print #SubRptHandle, TotByYrAndPen(y); dlm; 0; dlm; 0; dlm; "N"
            Else
              '                    4       5       6       7       8       9       10      11      12        13       14       15
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 15
            Print #SubRptHandle, "Adjust Prepay Down"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                    4       5       6       7       8       9       10      11      12       13       14       15
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 16
            Print #SubRptHandle, "Refund Prepay"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                    4       5       6       7       8       9       10      11      12       13       14       15
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 17
            Print #SubRptHandle, "PPTRA Removal"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                    4       5       6       7       8       9       10      11      12       13       14       15
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 18
            Print #SubRptHandle, "Unknown"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                    4       5       6       7       8       9       10      11      12       13       14       15
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
        End Select
      End If
    Next x
  Next y
  Close SubRptHandle
  
  Return
  
PrintTotals:
  Sub2RptFile$ = "TAXRPTS\SUB2TXJRLDETREAL.RPT"
  Sub2RptHandle = FreeFile
  Open Sub2RptFile For Output As #Sub2RptHandle
  GCntTot = 0
  GAmtTot = 0
  If fpcmbTransType.Text = " 0) All" Then GoTo All4
  
  For x = 1 To YrCnt
    GPrincTot = GPrincTot + TotByYrAndPrinc(x)
    GIntTot = GIntTot + TotByYrAndInt(x)
    GAdvTot = GAdvTot + TotByYrAndAdv(x)
    GLateListTot = GLateListTot + TotByYrAndLateList(x)
    GPenTot = GPenTot + TotByYrAndPen(x)
    GOpt1Tot = GOpt1Tot + TotByYrAndOpt1(x)
    GOpt2Tot = GOpt2Tot + TotByYrAndOpt2(x)
    GOpt3Tot = GOpt3Tot + TotByYrAndOpt3(x)
  Next x
All4:
  Dim Case11Tot As Double '7/6/06 added
  For x = 1 To 18
    GCntTot = GCntTot + TotCntByType(x)
    GAmtTot = GAmtTot + TotByType(x)
    If TotByType(x) > 0 Then
      If x <> 11 Then
        '                         0                    1                  2             3
        Print #Sub2RptHandle, TotByType(x); dlm; TotCntByType(x); dlm; GAmtTot; dlm; GCntTot; dlm;
      Else 'added 7/6/06
        Case11Tot = OldRound(TotByType(11) - (GPrincTot + GIntTot + GAdvTot + GLateListTot + GOpt1Tot))
        Case11Tot = OldRound(Case11Tot - (GOpt2Tot + GOpt3Tot + GPenTot))
        '                         0                 1                  2             3
        Print #Sub2RptHandle, Case11Tot; dlm; TotCntByType(x); dlm; GAmtTot; dlm; GCntTot; dlm;
      End If
      '                         4              5             6               7
      Print #Sub2RptHandle, GPrincTot; dlm; GIntTot; dlm; GAdvTot; dlm; GLateListTot; dlm;
      '                        8              9             10              11             12             13
      Print #Sub2RptHandle, GOpt1Tot; dlm; GOpt2Tot; dlm; GOpt3Tot; dlm; Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm;
      '                        14
      Print #Sub2RptHandle, GPenTot; dlm;
      If fpcmbTransType.Text <> " 0) All" Then
        '                      15
        Print #Sub2RptHandle, "1"; dlm;
      Else
        '                      15
        Print #Sub2RptHandle, "2"; dlm;
      End If
      Select Case x
        Case 1
          '                         16
          Print #Sub2RptHandle, "Billing"; dlm; 3
        Case 2
          Print #Sub2RptHandle, "Payment"; dlm; 3
        Case 3
          Print #Sub2RptHandle, "Release"; dlm; 3
        Case 4
          Print #Sub2RptHandle, "Interest"; dlm; 44
        Case 5
          Print #Sub2RptHandle, "Penalty"; dlm; 44
        Case 6
          Print #Sub2RptHandle, "Advertising"; dlm; 44
        Case 7
          Print #Sub2RptHandle, "Adjust Pay Down"; dlm; 3
        Case 8
          Print #Sub2RptHandle, "Credit at Billing"; dlm; 44
        Case 9
          Print #Sub2RptHandle, "Adjust Bill Down"; dlm; 3
        Case 10
          Print #Sub2RptHandle, "Adjust Bill Up"; dlm; 3
        Case 11
          If InStr(fpcmbTransType.Text, "All") Then 'added Bill Pay/OverPay on 7/7/06
            Print #Sub2RptHandle, "Bill Pay/OverPay"; dlm; 3 'changed from 44 on 7/6/06
          Else
            Print #Sub2RptHandle, "Bill OverPay"; dlm; 3 'changed from 44 on 7/6/06
          End If
        Case 12
          Print #Sub2RptHandle, "OverPayment"; dlm; 44
        Case 13
          Print #Sub2RptHandle, "Adjust Bill Up Affecting Credit Balance"; dlm; 3
        Case 14
          Print #Sub2RptHandle, "Adjust Pay Dwn Affecting Credit Balance"; dlm; 3
        Case 15
          Print #Sub2RptHandle, "Adjust Prepay Down"; dlm; 44
        Case 16
          Print #Sub2RptHandle, "Refund Prepay"; dlm; 33
        Case 17
          Print #Sub2RptHandle, "PPTRA Removal"; dlm; 44
        Case 18
          Print #Sub2RptHandle, "Unknown"; dlm; 44
      End Select
    End If
  Next x
  Close Sub2RptHandle
  
  Return

PrintAll:
  For x = 1 To NumOfTCRecs
    PrePayDone = False
    If IdxFlag = False Then
      Get TCHandle, x, TaxCust
      CustBal = GetCustBalance(x, -1)
    Else
      Get TCHandle, IdxArray(x), TaxCust
      CustBal = GetCustBalance(IdxArray(x), -1)
    End If
    If TaxCust.Active = "N" And InactiveFlag = False Then
      GoTo SkipThisOne
    End If
    ThisName = QPTrim$(TaxCust.CustName)
    ReDim BillRec(1 To 1) As Long
    GoSub PrintPrepay
    ThisRec = TaxCust.LastTrans
    
    BillCnt = 0
    BadCnt = 0
    Do While ThisRec > 0
      Get TTHandle, ThisRec, TaxTrans
      If QFlag = True Then
        If TaxTrans.TransDate < BegDate Then
          BadCnt = BadCnt + 1
          If BadCnt > 3 Then Exit Do
        End If
      End If
      If TransDesc <> "" Then
        If InStr(1, TaxTrans.Description, TransDesc) = 0 Then
          GoTo Nope
        End If
      End If
'      If TaxTrans.TranType = 1 Then Stop
      If TaxTrans.BillType <> "R" And TaxTrans.BillType <> "P" Then TaxTrans.BillType = ""
      If TaxTrans.BillType <> Mid(fpcmbTaxType.Text, 1, 1) And QPTrim$(TaxTrans.BillType) <> "" Then GoTo Nope
      If TaxTrans.TransDate < BegDate Or TaxTrans.TransDate > EndDate Then GoTo Nope
      If TaxTrans.TranType = 1 Then
        BillCnt = BillCnt + 1
        ReDim Preserve BillRec(1 To BillCnt) As Long
        BillRec(BillCnt) = ThisRec
      ElseIf TaxTrans.TranType = 11 Or TaxTrans.TranType = 12 Then
        GoTo Nope
'        BillCnt = BillCnt + 1
'        ReDim Preserve BillRec(1 To BillCnt) As Long
'        BillRec(BillCnt) = ThisRec
      End If
Nope:
      ThisRec = TaxTrans.LastTrans
    Loop
  
    For z = 1 To BillCnt
      Get TTHandle, BillRec(z), TaxTrans
      If YrCnt = 0 Then
         YrCnt = YrCnt + 1
         ThisYear = YrCnt
         ReDim Preserve ThEYear(1 To YrCnt) As Integer
         ThEYear(YrCnt) = TaxTrans.TaxYear
         ReDim Preserve TotByYrAndType(1 To 18, 1 To YrCnt) As Double
         ReDim Preserve CntByYrAndType(1 To 18, 1 To YrCnt) As Integer
         For y = 1 To 18
           TotByYrAndType(y, YrCnt) = 0
           CntByYrAndType(y, YrCnt) = 0
         Next y
       Else
         For y = 1 To YrCnt
           If TaxTrans.TaxYear = ThEYear(y) Then
             ThisYear = y
             Exit For
           End If
         Next y
         If y > YrCnt Then
           YrCnt = YrCnt + 1
           ThisYear = YrCnt
           ReDim Preserve ThEYear(1 To YrCnt) As Integer
           ThEYear(YrCnt) = TaxTrans.TaxYear
           ReDim Preserve TotByYrAndType(1 To 18, 1 To YrCnt) As Double
           ReDim Preserve CntByYrAndType(1 To 18, 1 To YrCnt) As Integer
           For y = 1 To 18
             TotByYrAndType(y, YrCnt) = 0
             CntByYrAndType(y, YrCnt) = 0
           Next y
        End If
      End If
'      If TaxTrans.TranType = 11 Then
'        ThisRec = BillRec(z)
'        GoTo PrePay
'      End If
      ThisTransType = "Billing"
      TCnt = TCnt + 1
      If TaxTrans.DiscAmt > 0 Then GoSub ApplyDisc '1/16/07
      TotAmt = OldRound(TotAmt + TaxTrans.Amount)
      
      '-----------------------------------------------------------------------
      PrincDif = OldRound(TaxTrans.Revenue.Principle1 - TaxTrans.Revenue.Principle1Pd)
      IntDif = OldRound(TaxTrans.Revenue.Interest - TaxTrans.Revenue.InterestPd)
      AdvDif = OldRound(TaxTrans.Revenue.Collection - TaxTrans.Revenue.CollectionPd)
      LateListDif = OldRound(TaxTrans.Revenue.LateList - TaxTrans.Revenue.LateListPd)
      PenDif = OldRound(TaxTrans.Revenue.Penalty - TaxTrans.Revenue.PenaltyPd)
      Opt1Dif = OldRound(TaxTrans.Revenue.RevOpt1 - TaxTrans.Revenue.RevOpt1Pd)
      Opt2Dif = OldRound(TaxTrans.Revenue.RevOpt2 - TaxTrans.Revenue.RevOpt2Pd)
      Opt3Dif = OldRound(TaxTrans.Revenue.RevOpt3 - TaxTrans.Revenue.RevOpt3Pd)
      BillBal = OldRound(PenDif + PrincDif + IntDif + AdvDif + LateListDif + Opt1Dif + Opt2Dif + Opt3Dif)
      '                   0            1                 2                   3
      Print #RptHandle, Town$; dlm; ThisName; dlm; TaxCust.Acct; dlm; TaxCust.Active; dlm;
      '                                 4                           5                6
      Print #RptHandle, MakeRegDate(TaxTrans.TransDate); dlm; ThisBillType; dlm; ThisType; dlm;
      '                          7                         8                          9
      Print #RptHandle, MakeRegDate(BegDate); dlm; MakeRegDate(EndDate); dlm; TaxTrans.TaxYear; dlm;
      '                      10                11          12                       13
      Print #RptHandle, TaxTrans.Amount; dlm; TCnt; dlm; TotAmt; dlm; TaxTrans.Revenue.PrePaidAmt; dlm;
      '                14
      Print #RptHandle, 0; dlm;
      '                                15                        16
      Print #RptHandle, QPTrim$(TaxTrans.Description); dlm; ThisTransType; dlm;
      '                               17                              18                           19
      Print #RptHandle, TaxTrans.Revenue.Principle1; dlm; TaxTrans.Revenue.Principle1Pd; ; dlm; PrincDif; dlm;
      '                               20                              21                      22
      Print #RptHandle, TaxTrans.Revenue.Interest; dlm; TaxTrans.Revenue.InterestPd; ; dlm; IntDif; dlm;
      '                               23                              24                        25
      Print #RptHandle, TaxTrans.Revenue.Collection; dlm; TaxTrans.Revenue.CollectionPd; dlm; AdvDif; dlm;
      '                               26                              27                       28
      Print #RptHandle, TaxTrans.Revenue.LateList; dlm; TaxTrans.Revenue.LateListPd; dlm; LateListDif; dlm;
      '                               29                              30                   31
      Print #RptHandle, TaxTrans.Revenue.RevOpt1; dlm; TaxTrans.Revenue.RevOpt1Pd; dlm; Opt1Dif; dlm;
      '                               32                              33                   34
      Print #RptHandle, TaxTrans.Revenue.RevOpt2; dlm; TaxTrans.Revenue.RevOpt2Pd; dlm; Opt2Dif; dlm;
      '                               35                              36                   37
      Print #RptHandle, TaxTrans.Revenue.RevOpt3; dlm; TaxTrans.Revenue.RevOpt3Pd; dlm; Opt3Dif; dlm;
      '                    38             39             40             41              42            43                 44
      Print #RptHandle, Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm; BillRec(z); dlm; CustBal; dlm; BillBal; dlm; TaxTrans.TranType; dlm;
      '                           45                              46                      47               48
      Print #RptHandle, TaxTrans.Revenue.Penalty; dlm; TaxTrans.Revenue.PenaltyPd; dlm; PenDif; dlm; TaxTrans.OperNum
      
      ThisRec = TaxCust.LastTrans
      TotByYrAndType(1, ThisYear) = OldRound(TotByYrAndType(1, ThisYear) + TaxTrans.Amount)
      CntByYrAndType(1, ThisYear) = OldRound(CntByYrAndType(1, ThisYear) + 1)
      TotByType(1) = OldRound(TotByType(1) + TaxTrans.Amount)
      TotCntByType(1) = OldRound(TotCntByType(1) + 1)

'PrePay:
      Do While ThisRec > 0
        Get TTHandle, ThisRec, TaxTrans
        If TaxTrans.TranType = 22 Then 'Prepay transactions can only be printed one time, not
        'for every iteration
          If PrePayDone = False Then
            GoTo PrepaySkip
          Else
            GoTo Nope2
          End If
        End If
        If TaxTrans.TranType = 11 Or TaxTrans.TranType = 12 Then GoTo Nope2
        If TaxTrans.BelongTo <> BillRec(z) Then GoTo Nope2
PrepaySkip:
        Select Case TaxTrans.TranType
          Case 2
            ThisTransType = "Payment"
            If TaxTrans.DiscAmt > 0 Then 'added 1/16/07
              GoSub ApplyDisc
            End If
            TotByYrAndType(2, ThisYear) = OldRound(TotByYrAndType(2, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(2, ThisYear) = OldRound(CntByYrAndType(2, ThisYear) + 1)
            TotByType(2) = OldRound(TotByType(2) + TaxTrans.Amount)
            TotCntByType(2) = OldRound(TotCntByType(2) + 1)
          Case 3
            ThisTransType = "Release"
            TotByYrAndType(3, ThisYear) = OldRound(TotByYrAndType(3, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(3, ThisYear) = OldRound(CntByYrAndType(3, ThisYear) + 1)
            TotByType(3) = OldRound(TotByType(3) + TaxTrans.Amount)
            TotCntByType(3) = OldRound(TotCntByType(3) + 1)
          Case 4
            ThisTransType = "Interest"
            TotByYrAndType(4, ThisYear) = OldRound(TotByYrAndType(4, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(4, ThisYear) = OldRound(CntByYrAndType(4, ThisYear) + 1)
            TotByType(4) = OldRound(TotByType(4) + TaxTrans.Amount)
            TotCntByType(4) = OldRound(TotCntByType(4) + 1)
          Case 5
            ThisTransType = "Penalty"
            TotByYrAndType(5, ThisYear) = OldRound(TotByYrAndType(5, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(5, ThisYear) = OldRound(CntByYrAndType(5, ThisYear) + 1)
            TotByType(5) = OldRound(TotByType(5) + TaxTrans.Amount)
            TotCntByType(5) = OldRound(TotCntByType(5) + 1)
          Case 6
            ThisTransType = "Advertising Charge"
            TotByYrAndType(6, ThisYear) = OldRound(TotByYrAndType(6, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(6, ThisYear) = OldRound(CntByYrAndType(6, ThisYear) + 1)
            TotByType(6) = OldRound(TotByType(6) + TaxTrans.Amount)
            TotCntByType(6) = OldRound(TotCntByType(6) + 1)
          Case 7
            ThisTransType = "Adjust Pay Down"
            TotByYrAndType(7, ThisYear) = OldRound(TotByYrAndType(7, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(7, ThisYear) = OldRound(CntByYrAndType(7, ThisYear) + 1)
            TotByType(7) = OldRound(TotByType(7) + TaxTrans.Amount)
            TotCntByType(7) = OldRound(TotCntByType(7) + 1)
          Case 9
            ThisTransType = "Credit Applied at Billing"
            TotByYrAndType(8, ThisYear) = OldRound(TotByYrAndType(8, ThisYear) + TaxTrans.Revenue.PrePaidUsed)
            CntByYrAndType(8, ThisYear) = OldRound(CntByYrAndType(8, ThisYear) + 1)
            TotByType(8) = OldRound(TotByType(8) + TaxTrans.Revenue.PrePaidUsed)
            TotCntByType(8) = OldRound(TotCntByType(8) + 1)
          Case 13
            ThisTransType = "Adjust Bill Down"
            TotByYrAndType(9, ThisYear) = OldRound(TotByYrAndType(9, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(9, ThisYear) = OldRound(CntByYrAndType(9, ThisYear) + 1)
            TotByType(9) = OldRound(TotByType(9) + TaxTrans.Amount)
            TotCntByType(9) = OldRound(TotCntByType(9) + 1)
          Case 14
            ThisTransType = "Adjust Bill Up"
            TotByYrAndType(10, ThisYear) = OldRound(TotByYrAndType(10, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(10, ThisYear) = OldRound(CntByYrAndType(10, ThisYear) + 1)
            TotByType(10) = OldRound(TotByType(10) + TaxTrans.Amount)
            TotCntByType(10) = OldRound(TotCntByType(10) + 1)
          Case 21
            ThisTransType = "Billpay/Overpay"
            If TaxTrans.DiscAmt > 0 Then 'added 1/16/07
              GoSub ApplyDisc
            End If
            TotByYrAndType(11, ThisYear) = OldRound(TotByYrAndType(11, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(11, ThisYear) = OldRound(CntByYrAndType(11, ThisYear) + 1)
            TotByType(11) = OldRound(TotByType(11) + TaxTrans.Amount)
            TotCntByType(11) = OldRound(TotCntByType(11) + 1)
          Case 22
            ThisTransType = "Overpayment"
            TotByYrAndType(12, ThisYear) = OldRound(TotByYrAndType(12, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(12, ThisYear) = OldRound(CntByYrAndType(12, ThisYear) + 1)
            TotByType(12) = OldRound(TotByType(12) + TaxTrans.Amount)
            TotCntByType(12) = OldRound(TotCntByType(12) + 1)
          Case 24
            ThisTransType = "Adjust Bill Up Affecting Credit Balance"
            TotByYrAndType(13, ThisYear) = OldRound(TotByYrAndType(13, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(13, ThisYear) = OldRound(CntByYrAndType(13, ThisYear) + 1)
            TotByType(13) = OldRound(TotByType(13) + TaxTrans.Amount)
            TotCntByType(13) = OldRound(TotCntByType(13) + 1)
          Case 10
            ThisTransType = "Adjust Pay Dwn Affecting Credit Balance"
            TotByYrAndType(14, ThisYear) = OldRound(TotByYrAndType(14, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(14, ThisYear) = OldRound(CntByYrAndType(14, ThisYear) + 1)
            TotByType(14) = OldRound(TotByType(14) + TaxTrans.Amount)
            TotCntByType(14) = OldRound(TotCntByType(14) + 1)
          Case 30
            ThisTransType = "PPTRA Removal"
            TotByYrAndType(17, ThisYear) = OldRound(TotByYrAndType(17, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(17, ThisYear) = OldRound(CntByYrAndType(17, ThisYear) + 1)
            TotByType(17) = OldRound(TotByType(17) + TaxTrans.Amount)
            TotCntByType(17) = OldRound(TotCntByType(17) + 1)
          Case Else
            ThisTransType = "Unknown"
            TotByYrAndType(18, ThisYear) = OldRound(TotByYrAndType(18, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(18, ThisYear) = OldRound(CntByYrAndType(18, ThisYear) + 1)
            TotByType(18) = OldRound(TotByType(18) + TaxTrans.Amount)
            TotCntByType(18) = OldRound(TotCntByType(18) + 1)
        End Select
        TCnt = TCnt + 1

        '                   0            1                 2                   3
        Print #RptHandle, Town$; dlm; ThisName; dlm; TaxCust.Acct; dlm; TaxCust.Active; dlm;
        '                                 4                           5                6
        Print #RptHandle, MakeRegDate(TaxTrans.TransDate); dlm; ThisBillType; dlm; ThisType; dlm;
        '                          7                         8                          9
        Print #RptHandle, MakeRegDate(BegDate); dlm; MakeRegDate(EndDate); dlm; TaxTrans.TaxYear; dlm;
        If TaxTrans.TranType <> 9 Then
          '                      10                11          12                       13
          Print #RptHandle, TaxTrans.Amount; dlm; TCnt; dlm; TotAmt; dlm; TaxTrans.Revenue.PrePaidAmt; dlm;
        Else
          '                      10                11          12                       13
          Print #RptHandle, TaxTrans.Revenue.PrePaidUsed; dlm; TCnt; dlm; TotAmt; dlm; TaxTrans.Revenue.PrePaidAmt; dlm;
        End If
        '                      14
        If TaxTrans.BelongTo > 0 Then
          Get TTHandle, TaxTrans.BelongTo, TaxTrans
          '                             14
          Print #RptHandle, ParseBillNum(TaxTrans.Description); dlm;
        Else
          '                 14
          Print #RptHandle, 0; dlm;
        End If
        Get TTHandle, ThisRec, TaxTrans
        If TaxTrans.DiscAmt > 0 Then GoSub ApplyDisc 'added 1/16/07
        '                                15                        16
        Print #RptHandle, QPTrim$(TaxTrans.Description); dlm; ThisTransType; dlm;
        '                               17                              18                           19
        Print #RptHandle, TaxTrans.Revenue.Principle1; dlm; TaxTrans.Revenue.Principle1Pd; ; dlm; 0; dlm;
        '                               20                              21                      22
        Print #RptHandle, TaxTrans.Revenue.Interest; dlm; TaxTrans.Revenue.InterestPd; ; dlm; 0; dlm;
        '                               23                              24                        25
        Print #RptHandle, TaxTrans.Revenue.Collection; dlm; TaxTrans.Revenue.CollectionPd; dlm; 0; dlm;
        '                               26                              27                       28
        Print #RptHandle, TaxTrans.Revenue.LateList; dlm; TaxTrans.Revenue.LateListPd; dlm; 0; dlm;
        '                               29                              30                   31
        Print #RptHandle, TaxTrans.Revenue.RevOpt1; dlm; TaxTrans.Revenue.RevOpt1Pd; dlm; 0; dlm;
        '                               32                              33                   34
        Print #RptHandle, TaxTrans.Revenue.RevOpt2; dlm; TaxTrans.Revenue.RevOpt2Pd; dlm; 0; dlm;
        '                               35                              36                   37
        Print #RptHandle, TaxTrans.Revenue.RevOpt3; dlm; TaxTrans.Revenue.RevOpt3Pd; dlm; 0; dlm;
        '                    38             39             40             41              42            43                44
        Print #RptHandle, Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm; BillRec(z); dlm; CustBal; dlm; BillBal; dlm; TaxTrans.TranType; dlm;
        '                               45                           46                     47               48
        Print #RptHandle, TaxTrans.Revenue.Penalty; dlm; TaxTrans.Revenue.PenaltyPd; dlm; PenDif; dlm; TaxTrans.OperNum
    
Nope2:
        ThisRec = TaxTrans.LastTrans
      Loop
      PrePayDone = True 'added 7/10/06
    Next z
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
SkipThisOne:
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  Close
  
  If TCnt = 0 Then
    Call TaxMsg(900, "No transactions were found that fit the parameters entered.")
    Close
    Exit Sub
  End If
  
  If YrCnt > 0 Then
    GoSub PrintSub
    GoSub PrintTotals
  End If
'  start here 5/16
  arVATaxJrnlAllDetail.Show
  
  Exit Sub

PrintPrepay:
  ThisRec = TaxCust.LastTrans
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
    If TaxTrans.TranType <> 11 And TaxTrans.TranType <> 12 Then GoTo NotThisTrans
    If TaxTrans.TransDate < BegDate Or TaxTrans.TransDate > EndDate Then GoTo NotThisTrans
    If TaxTrans.TranType = 11 Then
      ThisTransType = "Adj Prepay Down"
    ElseIf TaxTrans.TranType = 12 Then
      ThisTransType = "Refund Prepay"
    End If
    If YrCnt = 0 Then
      YrCnt = YrCnt + 1
      ThisYear = YrCnt
      ReDim Preserve ThEYear(1 To YrCnt) As Integer
      ThEYear(YrCnt) = TaxTrans.TaxYear
      ReDim Preserve TotByYrAndType(1 To 18, 1 To YrCnt) As Double
      ReDim Preserve CntByYrAndType(1 To 18, 1 To YrCnt) As Integer
      For y = 1 To 18
        TotByYrAndType(y, YrCnt) = 0
        CntByYrAndType(y, YrCnt) = 0
      Next y
    Else
      For y = 1 To YrCnt
        If TaxTrans.TaxYear = ThEYear(y) Then
          ThisYear = y
          Exit For
          End If
      Next y
      If y > YrCnt Then
        YrCnt = YrCnt + 1
        ThisYear = YrCnt
        ReDim Preserve ThEYear(1 To YrCnt) As Integer
        ThEYear(YrCnt) = TaxTrans.TaxYear
        ReDim Preserve TotByYrAndType(1 To 18, 1 To YrCnt) As Double
        ReDim Preserve CntByYrAndType(1 To 18, 1 To YrCnt) As Integer
        For y = 1 To 18
          TotByYrAndType(y, YrCnt) = 0
          CntByYrAndType(y, YrCnt) = 0
        Next y
      End If
    End If
    If TaxTrans.TranType = 11 Then
      TotByYrAndType(15, ThisYear) = OldRound(TotByYrAndType(15, ThisYear) + TaxTrans.Amount)
      CntByYrAndType(15, ThisYear) = OldRound(CntByYrAndType(15, ThisYear) + 1)
    ElseIf TaxTrans.TranType = 12 Then
      TotByYrAndType(16, ThisYear) = OldRound(TotByYrAndType(16, ThisYear) + TaxTrans.Amount)
      CntByYrAndType(16, ThisYear) = OldRound(CntByYrAndType(16, ThisYear) + 1)
    End If
    TCnt = TCnt + 1
    TotAmt = OldRound(TotAmt + TaxTrans.Amount)
    PrincDif = 0
    IntDif = 0
    AdvDif = 0
    LateListDif = 0
    PenDif = 0
    Opt1Dif = 0
    Opt2Dif = 0
    Opt3Dif = 0
    BillBal = OldRound(PenDif + PrincDif + IntDif + AdvDif + LateListDif + Opt1Dif + Opt2Dif + Opt3Dif)
    TotAmt = OldRound(TotAmt + TaxTrans.Amount)
    PrincDif = 0
    IntDif = 0
    AdvDif = 0
    LateListDif = 0
    PenDif = 0
    Opt1Dif = 0
    Opt2Dif = 0
    Opt3Dif = 0
    BillBal = 0
    '                   0            1                 2                   3
    Print #RptHandle, Town$; dlm; ThisName; dlm; TaxCust.Acct; dlm; TaxCust.Active; dlm;
    '                                 4                           5                6
    Print #RptHandle, MakeRegDate(TaxTrans.TransDate); dlm; ThisBillType; dlm; ThisType; dlm;
    '                          7                         8                          9
    Print #RptHandle, MakeRegDate(BegDate); dlm; MakeRegDate(EndDate); dlm; TaxTrans.TaxYear; dlm;
    '                      10                11          12                       13
    Print #RptHandle, TaxTrans.Amount; dlm; TCnt; dlm; TotAmt; dlm; TaxTrans.Revenue.PrePaidAmt; dlm;
    '                14
    Print #RptHandle, 0; dlm;
    '                                15                        16
    Print #RptHandle, QPTrim$(TaxTrans.Description); dlm; ThisTransType; dlm;
    '                               17                              18                           19
    Print #RptHandle, TaxTrans.Revenue.Principle1; dlm; TaxTrans.Revenue.Principle1Pd; ; dlm; PrincDif; dlm;
    '                               20                              21                      22
    Print #RptHandle, TaxTrans.Revenue.Interest; dlm; TaxTrans.Revenue.InterestPd; ; dlm; IntDif; dlm;
    '                               23                              24                        25
    Print #RptHandle, TaxTrans.Revenue.Collection; dlm; TaxTrans.Revenue.CollectionPd; dlm; AdvDif; dlm;
    '                               26                              27                       28
    Print #RptHandle, TaxTrans.Revenue.LateList; dlm; TaxTrans.Revenue.LateListPd; dlm; LateListDif; dlm;
    '                               29                              30                   31
    Print #RptHandle, TaxTrans.Revenue.RevOpt1; dlm; TaxTrans.Revenue.RevOpt1Pd; dlm; Opt1Dif; dlm;
    '                               32                              33                   34
    Print #RptHandle, TaxTrans.Revenue.RevOpt2; dlm; TaxTrans.Revenue.RevOpt2Pd; dlm; Opt2Dif; dlm;
    '                               35                              36                   37
    Print #RptHandle, TaxTrans.Revenue.RevOpt3; dlm; TaxTrans.Revenue.RevOpt3Pd; dlm; Opt3Dif; dlm;
    '                    38             39             40          41        42            43                 44
    Print #RptHandle, Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm; 0; dlm; CustBal; dlm; BillBal; dlm; TaxTrans.TranType; dlm;
    '                           45                              46                      47               48
    Print #RptHandle, TaxTrans.Revenue.Penalty; dlm; TaxTrans.Revenue.PenaltyPd; dlm; PenDif; dlm; TaxTrans.OperNum 'added this line on 1/30/07
NotThisTrans:
    ThisRec = TaxTrans.LastTrans
  Loop
  
  Return

ApplyDisc: '1/16/07
  Disc1 = 0
  Disc2 = 0
  Disc3 = 0
  Disc4 = 0
  If TaxTrans.Amount = 0 Then Return
  If TaxTrans.TranType = 1 Then
    SaveAmt = OldRound(TaxTrans.Amount - TaxTrans.DiscAmt)
  Else
    SaveAmt = TaxTrans.Amount
    TaxTrans.Amount = OldRound(TaxTrans.Amount + TaxTrans.DiscAmt)
  End If
  Disc1 = OldRound(TaxTrans.Revenue.Principle1Pd / SaveAmt)
  Disc1 = OldRound(Disc1 * TaxTrans.DiscAmt)
  Disc2 = OldRound(TaxTrans.Revenue.RevOpt1Pd / SaveAmt)
  Disc2 = OldRound(Disc2 * TaxTrans.DiscAmt)
  Disc3 = OldRound(TaxTrans.Revenue.RevOpt2Pd / SaveAmt)
  Disc3 = OldRound(Disc3 * TaxTrans.DiscAmt)
  Disc4 = OldRound(TaxTrans.Revenue.RevOpt3Pd / SaveAmt)
  Disc4 = OldRound(Disc4 * TaxTrans.DiscAmt)
  TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd + Disc1)
  TaxTrans.Revenue.RevOpt1Pd = OldRound(TaxTrans.Revenue.RevOpt1Pd + Disc2)
  TaxTrans.Revenue.RevOpt2Pd = OldRound(TaxTrans.Revenue.RevOpt2Pd + Disc3)
  TaxTrans.Revenue.RevOpt3Pd = OldRound(TaxTrans.Revenue.RevOpt3Pd + Disc4)
  DiscApplied = True
  
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxTransJournal", "PrintRGraphicsDet", Erl)
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

Private Sub PrintRTextDet()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim x As Long, y As Integer
  Dim BegDate As Integer
  Dim EndDate As Integer
  Dim ThisRec As Long
  Dim ThisClass As Integer
  Dim ThisType As String
  Dim RptFile$
  Dim RptHandle As Integer
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim InactiveFlag As Boolean
  Dim ThisName$, ThisBillType$
  Dim TCnt As Long, NewName$
  Dim TotAmt As Double
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim NumOfSrchRecs As Long
  Dim ThisTransType As String
  Dim YrCnt As Integer, ThisYear As Integer
  Dim BigYr As Integer
  Dim HoldBigYr As Integer
  Dim HoldYr As Integer
  Dim HoldPrinc As Double
  Dim HoldInt As Double
  Dim HoldAdv As Double
  Dim HoldLateList As Double
  Dim HoldOpt1 As Double
  Dim HoldOpt2 As Double
  Dim HoldOpt3 As Double
  Dim HoldPen As Double
  Dim Nextx As Integer
  Dim Thisx As Integer
  Dim Nexty As Integer
  Dim Thisy As Integer
  Dim z As Integer
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim FF$, Page As Integer
  Dim CustName$, PrintCnt As Integer
  Dim ThisBillNum As String * 8
  Dim Opt1Desc$
  Dim Opt2Desc$
  Dim Opt3Desc$
  Dim PrincDif As Double
  Dim IntDif As Double
  Dim AdvDif As Double
  Dim LateListDif As Double
  Dim PenDif As Double
  Dim Opt1Dif As Double
  Dim Opt2Dif As Double
  Dim Opt3Dif As Double
  Dim ThisBal As Double
  Dim ThisCustRec As Long
  Dim CustBal As Double
  Dim BillBal As Double
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim OptFlag As Boolean
  Dim ThisYrCnt As Long
  Dim ThatName$
  Dim TransDesc$
  Dim GAmtTot As Double
  Dim GCntTot As Long
  Dim BadCnt As Integer
  Dim QFlag As Boolean
  Dim GPrincTot As Double
  Dim GIntTot As Double
  Dim GAdvTot As Double
  Dim GLateListTot As Double
  Dim GPenTot As Double
  Dim GOpt1Tot As Double
  Dim GOpt2Tot As Double
  Dim GOpt3Tot As Double
  Dim PrePayDone As Boolean 'added 7/10/06
  Dim Disc1 As Double '1/16/2007
  Dim Disc2 As Double '1/16/2007
  Dim Disc3 As Double '1/16/2007
  Dim Disc4 As Double '1/16/2007
  Dim DiscApplied As Boolean '1/16/2007
  Dim SaveAmt As Double '1/16/2007
  
  On Error GoTo ERRORSTUFF
  
  PrePayDone = False
  QFlag = False
  If chkQuick.Value = 1 Then QFlag = True
  DiscApplied = False '1/16/07
  
  TransDesc = QPTrim$(fptxtDesc.Text)
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  Opt1Desc = QPTrim$(TaxMasterRec.OptRev1)
  Opt2Desc = QPTrim$(TaxMasterRec.OptRev2)
  Opt3Desc = QPTrim$(TaxMasterRec.OptRev3)
  CustName = ""
  IdxFlag = False
  OptFlag = False
  If CheckB4Printing = False Then
    Exit Sub
  End If
  FF$ = Chr$(12)
  MaxLines = 56
  LineCnt = 0
  ThisBillType = QPTrim$(fpcmbTaxType.Text)
  If fpcmbIncInactive.Text = "No" Then
    InactiveFlag = False
  Else
    InactiveFlag = True
  End If
  
  If Mid(fpcmbTransType.Text, 1, 1) = "" Then
    ThisClass = CInt(Mid(fpcmbTransType.Text, 2, 1))
  Else
    ThisClass = CInt(Mid(fpcmbTransType.Text, 1, 2))
  End If
    
  Select Case ThisClass
    Case 1
      ThisType = "Billing"
    Case 2
      ThisType = "Payment"
    Case 3
      ThisType = "Release"
    Case 4
      ThisType = "Interest"
    Case 5
      ThisType = "Penalty"
    Case 6
      ThisType = "Advertising Charge"
    Case 7, 10
      ThisType = "Adjust Pay Down"
    Case 9
      ThisType = "Credit Applied at Billing"
    Case 11
      ThisType = "Adjust Prepay Down"
    Case 12
      ThisType = "Refund Prepay"
    Case 13
      ThisType = "Adjust Bill Down"
    Case 14, 24
      ThisType = "Adjust Bill Up"
    Case 21
      ThisType = "Payment w/Overpay"
    Case 22
      ThisType = "Overpayment Only"
    Case 30
      ThisType = "PPTRA Removal"
    Case Else
      ThisType = "All"
  End Select
    
  BegDate = Date2Num(fptxtBegDate.Text)
  EndDate = Date2Num(fptxtEndDate.Text)
  
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
    OptFlag = True
  End If

  RptFile$ = "TAXRPTS\TAXRDJRNL.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  GoSub PrintHeader
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  
  frmVATaxShowPctComp.Label1 = "Gathering Tax Transaction Data"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  
  ReDim TotByYrAndType(1 To 18, 1 To 1) As Double
  ReDim CntByYrAndType(1 To 18, 1 To 1) As Integer
  ReDim TotByType(1 To 18) As Double
  ReDim TotCntByType(1 To 18) As Long
  
  ReDim ThEYear(1 To 1) As Integer
  ReDim TotByYrAndPrinc(1 To 1) As Double
  ReDim TotByYrAndInt(1 To 1) As Double
  ReDim TotByYrAndAdv(1 To 1) As Double
  ReDim TotByYrAndLateList(1 To 1) As Double
  ReDim TotByYrAndLateList(1 To 1) As Double
  ReDim TotByYrAndPen(1 To 1) As Double
  ReDim TotByYrAndOpt1(1 To 1) As Double
  ReDim TotByYrAndOpt2(1 To 1) As Double
  ReDim TotByYrAndOpt3(1 To 1) As Double
  
  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  End If
  If InStr(fpcmbTransType.Text, "All") Then GoTo PrintAll
  
  For x = 1 To NumOfTCRecs
    If IdxFlag = False Then
      Get TCHandle, x, TaxCust
      ThisCustRec = x
    Else
      Get TCHandle, IdxArray(x), TaxCust
      ThisCustRec = IdxArray(x)
    End If
    
    CustBal = GetCustBalance(ThisCustRec, -1)
    
    If TaxCust.Active = "N" And InactiveFlag = False Then
      GoTo SkipIt
    End If
    ThisName = QPTrim$(TaxCust.CustName)
    
    ThisRec = TaxCust.LastTrans
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintHeader
    End If
    PrintCnt = 0
    BadCnt = 0
    Do While ThisRec > 0
      Get TTHandle, ThisRec, TaxTrans
      If QFlag = True Then
        If TaxTrans.TransDate < BegDate Then
          BadCnt = BadCnt + 1
          If BadCnt > 3 Then Exit Do
        End If
      End If
      If TransDesc <> "" Then
        If InStr(1, TaxTrans.Description, TransDesc) = 0 Then GoTo SkipIt
      End If
      If TaxTrans.BillType <> "R" And TaxTrans.BillType <> "P" Then TaxTrans.BillType = ""
      If TaxTrans.BillType <> Mid(fpcmbTaxType.Text, 1, 1) And QPTrim$(TaxTrans.BillType) <> "" Then GoTo SkipIt
      If ThisClass = 7 And TaxTrans.TranType = 10 Then GoTo ItsOK
      If ThisClass = 14 And TaxTrans.TranType = 24 Then GoTo ItsOK
'      If ThisClass = 2 And TaxTrans.TranType = 21 Then GoTo ItsOK 'commented out 7/6/06
      If ThisClass <> 0 And TaxTrans.TranType <> ThisClass Then GoTo SkipIt
ItsOK:
      If TaxTrans.TransDate >= BegDate And TaxTrans.TransDate <= EndDate Then
        If TaxTrans.BelongTo > 0 Then
          Get TTHandle, TaxTrans.BelongTo, TaxTrans
        End If
        If PrintCnt = 0 Then
          If LineCnt <> 9 Then
            Print #RptHandle,
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintHeader
            End If
          End If
          GoSub PrintCustHeader
        End If
        PrintCnt = PrintCnt + 1
        If YrCnt = 0 Then
           YrCnt = YrCnt + 1
           ThisYear = YrCnt
           ReDim Preserve ThEYear(1 To YrCnt) As Integer
           ThEYear(YrCnt) = TaxTrans.TaxYear
           ReDim Preserve TotByYrAndType(1 To 18, 1 To YrCnt) As Double
           ReDim Preserve CntByYrAndType(1 To 18, 1 To YrCnt) As Integer
           ReDim Preserve TotByYrAndPrinc(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndInt(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndAdv(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndLateList(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndPen(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndOpt1(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndOpt2(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndOpt3(1 To YrCnt) As Double
           TotByYrAndPrinc(YrCnt) = 0
           TotByYrAndInt(YrCnt) = 0
           TotByYrAndAdv(YrCnt) = 0
           TotByYrAndLateList(YrCnt) = 0
           TotByYrAndPen(YrCnt) = 0
           TotByYrAndOpt1(YrCnt) = 0
           TotByYrAndOpt2(YrCnt) = 0
           TotByYrAndOpt3(YrCnt) = 0
           For y = 1 To 18
             TotByYrAndType(y, YrCnt) = 0
             CntByYrAndType(y, YrCnt) = 0
           Next y
         Else
           For y = 1 To YrCnt
             If TaxTrans.TaxYear = ThEYear(y) Then
               ThisYear = y
               Exit For
             End If
           Next y
           If y > YrCnt Then
             YrCnt = YrCnt + 1
             ThisYear = YrCnt
             ReDim Preserve ThEYear(1 To YrCnt) As Integer
             ThEYear(YrCnt) = TaxTrans.TaxYear
             ReDim Preserve TotByYrAndType(1 To 18, 1 To YrCnt) As Double
             ReDim Preserve CntByYrAndType(1 To 18, 1 To YrCnt) As Integer
             ReDim Preserve TotByYrAndPrinc(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndInt(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndAdv(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndLateList(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndPen(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndOpt1(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndOpt2(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndOpt3(1 To YrCnt) As Double
             TotByYrAndPrinc(YrCnt) = 0
             TotByYrAndInt(YrCnt) = 0
             TotByYrAndAdv(YrCnt) = 0
             TotByYrAndLateList(YrCnt) = 0
             TotByYrAndPen(YrCnt) = 0
             TotByYrAndOpt1(YrCnt) = 0
             TotByYrAndOpt2(YrCnt) = 0
             TotByYrAndOpt3(YrCnt) = 0
             For y = 1 To 18
               TotByYrAndType(y, YrCnt) = 0
               CntByYrAndType(y, YrCnt) = 0
             Next y
           End If
         End If
         Get TTHandle, ThisRec, TaxTrans
         DiscApplied = False 'added 1/16/07
         Select Case TaxTrans.TranType
           Case 1
             ThisTransType = "Billing"
             TotByYrAndType(1, ThisYear) = OldRound(TotByYrAndType(1, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(1, ThisYear) = OldRound(CntByYrAndType(1, ThisYear) + 1)
             TotByType(1) = OldRound(TotByType(1) + TaxTrans.Amount)
             TotCntByType(1) = OldRound(TotCntByType(1) + 1)
             TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1)
             '7/11/06 added back interest, advertising and penalty to accommodate manual bills
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.Interest)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.Collection)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.Penalty)
             
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateList)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3)
           Case 2
             ThisTransType = "Payment"
             If TaxTrans.DiscAmt > 0 Then 'added 1/16/07
               GoSub ApplyDisc
             End If
             TotByYrAndType(2, ThisYear) = OldRound(TotByYrAndType(2, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(2, ThisYear) = OldRound(CntByYrAndType(2, ThisYear) + 1)
             TotByType(2) = OldRound(TotByType(2) + TaxTrans.Amount)
             TotCntByType(2) = OldRound(TotCntByType(2) + 1)
             TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
           Case 3
            '7/12/06 changed revenues for Release to paid from charged
             ThisTransType = "Release"
             TotByYrAndType(3, ThisYear) = OldRound(TotByYrAndType(3, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(3, ThisYear) = OldRound(CntByYrAndType(3, ThisYear) + 1)
             TotByType(3) = OldRound(TotByType(3) + TaxTrans.Amount)
             TotCntByType(3) = OldRound(TotCntByType(3) + 1)
             TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
           Case 4
             ThisTransType = "Interest"
             TotByYrAndType(4, ThisYear) = OldRound(TotByYrAndType(4, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(4, ThisYear) = OldRound(CntByYrAndType(4, ThisYear) + 1)
             TotByType(4) = OldRound(TotByType(4) + TaxTrans.Amount)
             TotCntByType(4) = OldRound(TotCntByType(4) + 1)
           Case 5
             ThisTransType = "Penalty"
             TotByYrAndType(5, ThisYear) = OldRound(TotByYrAndType(5, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(5, ThisYear) = OldRound(CntByYrAndType(5, ThisYear) + 1)
             TotByType(5) = OldRound(TotByType(5) + TaxTrans.Amount)
             TotCntByType(5) = OldRound(TotCntByType(5) + 1)
           Case 6
             ThisTransType = "Advertising Charge"
             TotByYrAndType(6, ThisYear) = OldRound(TotByYrAndType(6, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(6, ThisYear) = OldRound(CntByYrAndType(6, ThisYear) + 1)
             TotByType(6) = OldRound(TotByType(6) + TaxTrans.Amount)
             TotCntByType(6) = OldRound(TotCntByType(6) + 1)
           Case 7
             ThisTransType = "Adjust Pay Down"
             TotByYrAndType(7, ThisYear) = OldRound(TotByYrAndType(7, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(7, ThisYear) = OldRound(CntByYrAndType(7, ThisYear) + 1)
             TotByType(7) = OldRound(TotByType(7) + TaxTrans.Amount)
             TotCntByType(7) = OldRound(TotCntByType(7) + 1)
             TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
           Case 9
             ThisTransType = "Credit Applied at Billing"
             TotByYrAndType(8, ThisYear) = OldRound(TotByYrAndType(8, ThisYear) + TaxTrans.Revenue.PrePaidUsed)
             CntByYrAndType(8, ThisYear) = OldRound(CntByYrAndType(8, ThisYear) + 1)
             TotByType(8) = OldRound(TotByType(8) + TaxTrans.Revenue.PrePaidUsed)
             TotCntByType(8) = OldRound(TotCntByType(8) + 1)
           Case 13
             ThisTransType = "Adjust Bill Down"
             TotByYrAndType(9, ThisYear) = OldRound(TotByYrAndType(9, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(9, ThisYear) = OldRound(CntByYrAndType(9, ThisYear) + 1)
             TotByType(9) = OldRound(TotByType(9) + TaxTrans.Amount)
             TotCntByType(9) = OldRound(TotCntByType(9) + 1)
             TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.Interest)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.Collection)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateList)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.Penalty)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3)
           Case 14
             ThisTransType = "Adjust Bill Up"
             TotByYrAndType(10, ThisYear) = OldRound(TotByYrAndType(10, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(10, ThisYear) = OldRound(CntByYrAndType(10, ThisYear) + 1)
             TotByType(10) = OldRound(TotByType(10) + TaxTrans.Amount)
             TotCntByType(10) = OldRound(TotCntByType(10) + 1)
             TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.Interest)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.Collection)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateList)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.Penalty)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3)
           Case 21
             ThisTransType = "Billpay/Overpay" '7/6/06 changed Amount to PrePaidAmt
             If TaxTrans.DiscAmt > 0 Then 'added 1/16/07
               GoSub ApplyDisc
             End If
             If fpcmbTransType.Text <> " 0) All" Then 'added the All if statement on 7/7/06
               TotByYrAndType(11, ThisYear) = OldRound(TotByYrAndType(11, ThisYear) + TaxTrans.Revenue.PrePaidAmt)
             Else
               TotByYrAndType(11, ThisYear) = OldRound(TotByYrAndType(11, ThisYear) + TaxTrans.Amount)
             End If
'             TotByYrAndType(11, ThisYear) = OldRound(TotByYrAndType(11, ThisYear) + TaxTrans.Revenue.PrePaidAmt) '  .Amount)
             CntByYrAndType(11, ThisYear) = OldRound(CntByYrAndType(11, ThisYear) + 1)
             TotByType(11) = OldRound(TotByType(11) + TaxTrans.Amount)
             TotCntByType(11) = OldRound(TotCntByType(11) + 1)
             TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
           Case 22
             ThisTransType = "Overpayment"
             TotByYrAndType(12, ThisYear) = OldRound(TotByYrAndType(12, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(12, ThisYear) = OldRound(CntByYrAndType(12, ThisYear) + 1)
             TotByType(12) = OldRound(TotByType(12) + TaxTrans.Amount)
             TotCntByType(12) = OldRound(TotCntByType(12) + 1)
           Case 24
             ThisTransType = "Adj Bill Up -Cre"
             TotByYrAndType(13, ThisYear) = OldRound(TotByYrAndType(13, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(13, ThisYear) = OldRound(CntByYrAndType(13, ThisYear) + 1)
             TotByType(13) = OldRound(TotByType(13) + TaxTrans.Amount)
             TotCntByType(13) = OldRound(TotCntByType(13) + 1)
             TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
'             TotByYrAndType(13, ThisYear) = OldRound(TotByYrAndType(13, ThisYear) + TaxTrans.Amount)
'             CntByYrAndType(13, ThisYear) = OldRound(CntByYrAndType(13, ThisYear) + 1)
'             TotByType(13) = OldRound(TotByType(13) + TaxTrans.Amount)
'             TotCntByType(13) = OldRound(TotCntByType(13) + 1)
'             TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1)
'             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.Interest)
'             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.Collection)
'             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateList)
'             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.Penalty)
'             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1)
'             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2)
'             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3)
           Case 10 ', 24 '7/11/06 added Pd on revenues
             ThisTransType = "Adj Pay Dwn -Cre"
             TotByYrAndType(14, ThisYear) = OldRound(TotByYrAndType(14, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(14, ThisYear) = OldRound(CntByYrAndType(14, ThisYear) + 1)
             TotByType(14) = OldRound(TotByType(14) + TaxTrans.Amount)
             TotCntByType(14) = OldRound(TotCntByType(14) + 1)
             TotByYrAndPrinc(ThisYear) = OldRound(TotByYrAndPrinc(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
           Case 11
             ThisTransType = "Adj Prepay -Cre"
             TotByYrAndType(15, ThisYear) = OldRound(TotByYrAndType(15, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(15, ThisYear) = OldRound(CntByYrAndType(15, ThisYear) + 1)
             TotByType(15) = OldRound(TotByType(15) + TaxTrans.Amount)
             TotCntByType(15) = OldRound(TotCntByType(15) + 1)
           Case 12
             ThisTransType = "Ref Prepay -Cre"
             TotByYrAndType(16, ThisYear) = OldRound(TotByYrAndType(16, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(16, ThisYear) = OldRound(CntByYrAndType(16, ThisYear) + 1)
             TotByType(16) = OldRound(TotByType(16) + TaxTrans.Amount)
             TotCntByType(16) = OldRound(TotCntByType(16) + 1)
           Case 30
             ThisTransType = "PPTRA Removal"
             TotByYrAndType(17, ThisYear) = OldRound(TotByYrAndType(17, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(17, ThisYear) = OldRound(CntByYrAndType(17, ThisYear) + 1)
             TotByType(17) = OldRound(TotByType(17) + TaxTrans.Amount)
             TotCntByType(17) = OldRound(TotCntByType(17) + 1)
           Case Else
             ThisTransType = "Unknown"
             TotByYrAndType(18, ThisYear) = OldRound(TotByYrAndType(18, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(18, ThisYear) = OldRound(CntByYrAndType(18, ThisYear) + 1)
             TotByType(18) = OldRound(TotByType(18) + TaxTrans.Amount)
             TotCntByType(18) = OldRound(TotCntByType(18) + 1)
        End Select
        TCnt = TCnt + 1
        If TaxTrans.DiscAmt > 0 Then GoSub ApplyDisc 'added 1/16/07
        TotAmt = OldRound(TotAmt + TaxTrans.Amount)
        PrincDif = OldRound(TaxTrans.Revenue.Principle1 - TaxTrans.Revenue.Principle1Pd)
        IntDif = OldRound(TaxTrans.Revenue.Interest - TaxTrans.Revenue.InterestPd)
        AdvDif = OldRound(TaxTrans.Revenue.Collection - TaxTrans.Revenue.CollectionPd)
        LateListDif = OldRound(TaxTrans.Revenue.LateList - TaxTrans.Revenue.LateListPd)
        PenDif = OldRound(TaxTrans.Revenue.Penalty - TaxTrans.Revenue.PenaltyPd)
        Opt1Dif = OldRound(TaxTrans.Revenue.RevOpt1 - TaxTrans.Revenue.RevOpt1Pd)
        Opt2Dif = OldRound(TaxTrans.Revenue.RevOpt2 - TaxTrans.Revenue.RevOpt2Pd)
        Opt3Dif = OldRound(TaxTrans.Revenue.RevOpt3 - TaxTrans.Revenue.RevOpt3Pd)
        ThisBal = OldRound(PrincDif + IntDif + AdvDif + LateListDif + Opt1Dif + Opt2Dif + Opt3Dif)
        Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); QPTrim$(TaxTrans.Description);
        If TaxTrans.BelongTo > 0 Then
          Get TTHandle, TaxTrans.BelongTo, TaxTrans
          Print #RptHandle, Tab(37); Using$("###0", TaxTrans.TaxYear);
          Get TTHandle, ThisRec, TaxTrans
          If TaxTrans.DiscAmt > 0 Then GoSub ApplyDisc 'added 1/16/07
        Else
          Print #RptHandle, Tab(37); Using$("###0", TaxTrans.TaxYear);
        End If
        Print #RptHandle, Tab(45); Using$("$##,##0.00", TaxTrans.Revenue.PrePaidAmt);
        If TaxTrans.TranType <> 9 Then
          Print #RptHandle, Tab(56); Using$("$##,##0.00", TaxTrans.Amount); Tab(69);
        Else
          Print #RptHandle, Tab(56); Using$("$##,##0.00", TaxTrans.Revenue.PrePaidUsed); Tab(69);
        End If
        LineCnt = LineCnt + 1
        If TaxTrans.BelongTo > 0 Then
          Get TTHandle, TaxTrans.BelongTo, TaxTrans
          ThisBillNum = ParseBillNum(TaxTrans.Description)
          If IsNumeric(ThisBillNum) Then
            Print #RptHandle, Using$("######", CDbl(ThisBillNum));
          Else
            Print #RptHandle, "   " + ThisBillNum;
          End If
        Else
          Print #RptHandle, "     0";
        End If
      
        Get TTHandle, ThisRec, TaxTrans
        If TaxTrans.DiscAmt > 0 Then GoSub ApplyDisc 'added 1/16/07
        Print #RptHandle, Tab(79); ThisTransType
        If LineCnt >= MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintHeader
        End If
        If ThisType <> "Billing" Then
          Print #RptHandle, Tab(3); "Oper #: " + CStr(TaxTrans.OperNum);
          Print #RptHandle, Tab(15); "Principle         "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle1); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle1Pd)
          Print #RptHandle, Tab(15); "Interest          "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Interest); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.InterestPd)
          Print #RptHandle, Tab(15); "Advertising       "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Collection); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.CollectionPd)
          Print #RptHandle, Tab(15); "Late Listing      "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.LateList); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.LateListPd)
          Print #RptHandle, Tab(15); "Penalty           "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Penalty); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.PenaltyPd)
          LineCnt = LineCnt + 6
          If Len(Opt1Desc) > 0 Then
            Print #RptHandle, Tab(15); Opt1Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1Pd)
            LineCnt = LineCnt + 1
          End If
          If Len(Opt2Desc) > 0 Then
            Print #RptHandle, Tab(15); Opt2Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2Pd)
            LineCnt = LineCnt + 1
          End If
            If Len(Opt3Desc) > 0 Then
            Print #RptHandle, Tab(15); Opt3Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3Pd)
            LineCnt = LineCnt + 1
          End If
        Else
          Print #RptHandle, Tab(3); "Oper #: " + CStr(TaxTrans.OperNum);
          Print #RptHandle, Tab(15); "Principle         "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle1); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle1Pd); Tab(80); Using$("$##,##0.00", PrincDif)
          '7/11/06 added back int, adv and pen to accommodate manual bills
          Print #RptHandle, Tab(15); "Interest          "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Interest); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.InterestPd); Tab(80); Using$("$##,##0.00", IntDif)
          Print #RptHandle, Tab(15); "Advertising       "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Collection); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.CollectionPd); Tab(80); Using$("$##,##0.00", AdvDif)
          Print #RptHandle, Tab(15); "Late Listing      "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.LateList); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.LateListPd); Tab(80); Using$("$##,##0.00", LateListDif)
          Print #RptHandle, Tab(15); "Penalty           "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Penalty); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.PenaltyPd); Tab(80); Using$("$##,##0.00", PenDif)
          LineCnt = LineCnt + 6
          If Len(Opt1Desc) > 0 Then
            Print #RptHandle, Tab(15); Opt1Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1Pd); Tab(80); Using$("$##,##0.00", Opt1Dif)
            LineCnt = LineCnt + 1
          End If
          If Len(Opt2Desc) > 0 Then
            Print #RptHandle, Tab(15); Opt2Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2Pd); Tab(80); Using$("$##,##0.00", Opt2Dif)
            LineCnt = LineCnt + 1
          End If
            If Len(Opt3Desc) > 0 Then
            Print #RptHandle, Tab(15); Opt3Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3Pd); Tab(80); Using$("$##,##0.00", Opt3Dif)
            LineCnt = LineCnt + 1
          End If
          Print #RptHandle, Tab(15); "Bill Balance:"; Tab(80); Using$("$##,##0.00", ThisBal)
        End If
    End If
SkipIt:
    ThisRec = TaxTrans.LastTrans
    Loop
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
  
  If TCnt = 0 Then
    Call TaxMsg(900, "No transactions were found that fit the parameters entered.")
    Close
    Exit Sub
  End If
  If YrCnt > 0 Then
    GoSub SortIt
    GoSub PrintTotals
  End If
  Print #RptHandle, FF$
  Close
  ViewPrint RptFile, "Tax Transactions Report", True
  
  Exit Sub
  
SortIt:
  BigYr = 0
  For x = 1 To YrCnt
    If ThEYear(x) > BigYr Then
      BigYr = ThEYear(x)
    End If
  Next x
  ReDim HoldAmt(1 To 18, 1 To YrCnt) As Double
  ReDim HoldCnt(1 To 18, 1 To YrCnt) As Integer
  Nexty = 1
  Nextx = 1
  HoldBigYr = 0
  For x = 1 To 18
    For y = Nexty To YrCnt
      If ThEYear(y) >= HoldBigYr Then
        HoldBigYr = ThEYear(y)
        Thisx = x
        Thisy = y
      End If
    Next y
    For z = 1 To 18
      HoldAmt(z, Thisy) = TotByYrAndType(z, Nexty)
      HoldCnt(z, Thisy) = CntByYrAndType(z, Nexty)
    Next z
    HoldYr = ThEYear(Nexty)
    If fpcmbTransType.Text = " 0) All" Then GoTo All1
    HoldPrinc = TotByYrAndPrinc(Nexty)
    HoldInt = TotByYrAndInt(Nexty)
    HoldAdv = TotByYrAndAdv(Nexty)
    HoldLateList = TotByYrAndLateList(Nexty)
    HoldPen = TotByYrAndPen(Nexty)
    HoldOpt1 = TotByYrAndOpt1(Nexty)
    HoldOpt2 = TotByYrAndOpt2(Nexty)
    HoldOpt3 = TotByYrAndOpt3(Nexty)
All1:
    For z = 1 To 18
      TotByYrAndType(z, Nexty) = TotByYrAndType(z, Thisy)
      CntByYrAndType(z, Nexty) = CntByYrAndType(z, Thisy)
    Next z
    ThEYear(Nexty) = ThEYear(Thisy)
    If fpcmbTransType.Text = " 0) All" Then GoTo All2
    TotByYrAndPrinc(Nexty) = TotByYrAndPrinc(Thisy)
    TotByYrAndInt(Nexty) = TotByYrAndInt(Thisy)
    TotByYrAndAdv(Nexty) = TotByYrAndAdv(Thisy)
    TotByYrAndLateList(Nexty) = TotByYrAndLateList(Thisy)
    TotByYrAndPen(Nexty) = TotByYrAndPen(Thisy)
    TotByYrAndOpt1(Nexty) = TotByYrAndOpt1(Thisy)
    TotByYrAndOpt2(Nexty) = TotByYrAndOpt2(Thisy)
    TotByYrAndOpt3(Nexty) = TotByYrAndOpt3(Thisy)
All2:
    For z = 1 To 18
      TotByYrAndType(z, Thisy) = HoldAmt(z, Thisy)
      CntByYrAndType(z, Thisy) = HoldCnt(z, Thisy)
    Next z
    ThEYear(Thisy) = HoldYr
    If fpcmbTransType.Text = " 0) All" Then GoTo All3
    TotByYrAndPrinc(Thisy) = HoldPrinc
    TotByYrAndInt(Thisy) = HoldInt
    TotByYrAndAdv(Thisy) = HoldAdv
    TotByYrAndLateList(Thisy) = HoldLateList
    TotByYrAndOpt1(Thisy) = HoldOpt1
    TotByYrAndOpt2(Thisy) = HoldOpt2
    TotByYrAndOpt3(Thisy) = HoldOpt3
All3:
    If Nexty >= YrCnt Then Exit For
    HoldBigYr = 0 'BigYr + 1
    Nexty = Nexty + 1
  Next x
  Print #RptHandle, FF$
  GoSub PrintSortHeader
  Print #RptHandle, "Total Transaction Count: " + Using$("#####0", TCnt)
  Print #RptHandle, String(94, "-")
  LineCnt = LineCnt + 2
  For y = 1 To YrCnt
    If LineCnt >= MaxLines - 4 Then
      Print #RptHandle, FF$
      GoSub PrintSortHeader
      Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
    Else
      Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
    End If
    LineCnt = LineCnt + 1
    ThisYrCnt = 0
    For x = 1 To 18
      If TotByYrAndType(x, y) > 0 Then
        ThisYrCnt = OldRound(ThisYrCnt + CntByYrAndType(x, y))
        Select Case x
          Case 1
            Print #RptHandle, "  Billing"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOne
            Print #RptHandle, Tab(5); "Principal:    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPrinc(y))
            '7/11/06 added back Int, Adv and Pen to accommodate manual bills
            Print #RptHandle, Tab(5); "Interest:     "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Advertising:  "; Tab(55); Using$("$###,###,##0.00", TotByYrAndAdv(y))
            Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndLateList(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 5
            If LineCnt >= MaxLines - 3 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 2
            Print #RptHandle, "  Payment"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOne
            Print #RptHandle, Tab(5); "Principal:    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPrinc(y))
            Print #RptHandle, Tab(5); "Interest :    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Advertising : "; Tab(55); Using$("$###,###,##0.00", TotByYrAndAdv(y))
            Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndLateList(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 7
            If LineCnt >= MaxLines - 3 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 3
            Print #RptHandle, "  Release"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOne
            Print #RptHandle, Tab(5); "Principal:    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPrinc(y))
            Print #RptHandle, Tab(5); "Interest :    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Advertising : "; Tab(55); Using$("$###,###,##0.00", TotByYrAndAdv(y))
            Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndLateList(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 7
            If LineCnt >= MaxLines - 3 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 4
            Print #RptHandle, "  Interest"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
          Case 5
            Print #RptHandle, "  Penalty"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
          Case 6
            Print #RptHandle, "  Advertising"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
          Case 7
            Print #RptHandle, "  Adjust Pay Down"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOne
            Print #RptHandle, Tab(5); "Principal:    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPrinc(y))
            Print #RptHandle, Tab(5); "Interest :    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Advertising : "; Tab(55); Using$("$###,###,##0.00", TotByYrAndAdv(y))
            Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndLateList(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 7
            If LineCnt >= MaxLines - 3 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 8
            Print #RptHandle, "  Credit at Billing"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
          Case 9
            Print #RptHandle, "  Adjust Bill Down"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOne
            Print #RptHandle, Tab(5); "Principal:    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPrinc(y))
            Print #RptHandle, Tab(5); "Interest :    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Advertising : "; Tab(55); Using$("$###,###,##0.00", TotByYrAndAdv(y))
            Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndLateList(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 7
            If LineCnt >= MaxLines - 3 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
         Case 10
            Print #RptHandle, "  Adjust Bill Up"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOne
            Print #RptHandle, Tab(5); "Principal:    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPrinc(y))
            Print #RptHandle, Tab(5); "Interest :    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Advertising : "; Tab(55); Using$("$###,###,##0.00", TotByYrAndAdv(y))
            Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndLateList(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 7
            If LineCnt >= MaxLines - 3 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 11
            If fpcmbTransType.Text = " 0) All" Then 'added Bill Pay/OverPay on 7/7/06
              Print #RptHandle, "  Bill Pay/OverPay"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            Else
              Print #RptHandle, "  Bill OverPay"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            End If
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            'added detail on 7/6/06
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOne
            Print #RptHandle, Tab(5); "Principal:    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPrinc(y))
            Print #RptHandle, Tab(5); "Interest :    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Advertising : "; Tab(55); Using$("$###,###,##0.00", TotByYrAndAdv(y))
            Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndLateList(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 7
            If LineCnt >= MaxLines - 3 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 12
            Print #RptHandle, "  OverPayment"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
          Case 13
            Print #RptHandle, "  Adj Bill Up: -Credit "; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOne
            Print #RptHandle, Tab(5); "Principal:    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPrinc(y))
            Print #RptHandle, Tab(5); "Interest :    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Advertising : "; Tab(55); Using$("$###,###,##0.00", TotByYrAndAdv(y))
            Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndLateList(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 7
            If LineCnt >= MaxLines - 3 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 14
            Print #RptHandle, "  Adj Pay Dn: -Credit"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOne
            Print #RptHandle, Tab(5); "Principal:    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPrinc(y))
            Print #RptHandle, Tab(5); "Interest :    "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Advertising : "; Tab(55); Using$("$###,###,##0.00", TotByYrAndAdv(y))
            Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndLateList(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 7
            If LineCnt >= MaxLines - 3 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 15
            Print #RptHandle, "  Adjust Prepay Down"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
          Case 16
            Print #RptHandle, "  Refund Prepay"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
          Case 17
            Print #RptHandle, "  PPTRA Removal"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
            End If
          Case 18
            Print #RptHandle, "  Unknown"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
            End If
        End Select
      End If
NextOne:
    Next x
    Print #RptHandle, "  Total Year Count: "; Tab(30); Using$("###,###,##0", ThisYrCnt)
    Print #RptHandle, String$(94, "-")
    Print #RptHandle,
    LineCnt = LineCnt + 3
  Next y
  
  Return

PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(30); "Tax Transactions Journal"
  Print #RptHandle, Town; Tab(75); "Page #: " + CStr(Page)
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle, "Tax Type: " + ThisType
  Print #RptHandle, "Transaction Type: " + ThisBillType
  Print #RptHandle, "Date Range: " + fptxtBegDate.Text + " to " + fptxtEndDate.Text
  Print #RptHandle,
  Print #RptHandle, "Trans Date"; Tab(12); "Description"; Tab(35); "Tax Year"; Tab(44); "Overpay Amt"; Tab(57); "Trans Amt"; Tab(78); "Trans Type" 'Tab(67); "Belongs To"; Tab(78); "Trans Type"
  Print #RptHandle, String(94, "-")
  LineCnt = 9
  
  Return
  
PrintCustHeader:
  If LineCnt <> 9 Then
    Print #RptHandle, String(94, "-")
    LineCnt = LineCnt + 1
  End If
  If LineCnt >= MaxLines - 5 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  Print #RptHandle, "Cust Num: " + Using$("#######0", TaxCust.Acct); Tab(21); "Customer Name: "; Tab(37); QPTrim$(TaxCust.CustName); Tab(80); "Active: "; Tab(89); TaxCust.Active
  If UseOpt = "Y" Then
    Print #RptHandle, Tab(21); ThisOpt + ":" + QPTrim$(TaxCust.OptSrchDesc)
    LineCnt = LineCnt + 1
  End If
  Print #RptHandle, "Total Outstanding Customer Balance: " + Using$("$##,##0.00", CustBal)
  Print #RptHandle, Tab(15); "Revenue"; Tab(44); "Amount"; Tab(59); "Amount Paid"; Tab(83); "Balance"
  Print #RptHandle, String(94, ".")
  LineCnt = LineCnt + 4
  
  Return
  
PrintSortHeader:
  Page = Page + 1
  Print #RptHandle, Tab(25); "Tax Transactions Journal Summary"
  Print #RptHandle, Town; Tab(75); "Page #: " + CStr(Page)
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle, "Tax Type: " + ThisType
  Print #RptHandle, "Transaction Type: " + ThisBillType
  Print #RptHandle, "Date Range: " + fptxtBegDate.Text + " to " + fptxtEndDate.Text
  Print #RptHandle, "Note: Adjustment transaction amounts are reflected in revenues and "
  Print #RptHandle, "      adjustment transaction totals exclusively. "
  Print #RptHandle, "Description"; Tab(35); "Trans Cnt"; Tab(64); "Amount"
  Print #RptHandle, String$(94, "-")
  LineCnt = 10
  
  Return
  
PrintTotalsHeader:
  Page = Page + 1
  Print #RptHandle, Tab(30); "Tax Transactions Journal Summary"
  Print #RptHandle, "Grand Totals"
  Print #RptHandle, Town; Tab(75); "Page #: " + CStr(Page)
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle, "Tax Type: " + ThisType
  Print #RptHandle, "Transaction Type: " + ThisBillType
  Print #RptHandle, "Date Range: " + fptxtBegDate.Text + " to " + fptxtEndDate.Text
  Print #RptHandle, "Note: Adjustment transaction amounts are reflected in revenues and "
  Print #RptHandle, "      adjustment transaction totals exclusively. "
  Print #RptHandle, Tab(35); "Trans Cnt"; Tab(64); "Amount"
  Print #RptHandle, String$(94, "-")
  LineCnt = 10

  Return
  
PrintTotals:
  GCntTot = 0
  GAmtTot = 0
  If LineCnt > MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintTotalsHeader
  Else
    Print #RptHandle,
    Print #RptHandle, "Grand Totals"
    Print #RptHandle, Tab(35); "Trans Cnt"; Tab(64); "Amount"
    Print #RptHandle, String$(94, "-")
    LineCnt = LineCnt + 4
  End If
  If fpcmbTransType.Text = " 0) All" Then GoTo All4
  For x = 1 To YrCnt
    GPrincTot = GPrincTot + TotByYrAndPrinc(x)
    GIntTot = GIntTot + TotByYrAndInt(x)
    GAdvTot = GAdvTot + TotByYrAndAdv(x)
    GLateListTot = GLateListTot + TotByYrAndLateList(x)
    GPenTot = GPenTot + TotByYrAndPen(x)
    GOpt1Tot = GOpt1Tot + TotByYrAndOpt1(x)
    GOpt2Tot = GOpt2Tot + TotByYrAndOpt2(x)
    GOpt3Tot = GOpt3Tot + TotByYrAndOpt3(x)
  Next x
All4:
  For x = 1 To 18
    GCntTot = GCntTot + TotCntByType(x)
    GAmtTot = GAmtTot + TotByType(x)
    If TotByType(x) > 0 Then
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintTotalsHeader
      End If
      Select Case x
        Case 1
          Print #RptHandle, "Billing";
        Case 2
          Print #RptHandle, "Payment";
        Case 3
          Print #RptHandle, "Release";
        Case 4
          Print #RptHandle, "Interest";
        Case 5
          Print #RptHandle, "Penalty";
        Case 6
          Print #RptHandle, "Advertising";
        Case 7
          Print #RptHandle, "Adjust Pay Down";
        Case 8
          Print #RptHandle, "Credit at Billing";
        Case 9
          Print #RptHandle, "Adjust Bill Down";
        Case 10
          Print #RptHandle, "Adjust Bill Up";
        Case 11
          If fpcmbTransType.Text = " 0) All" Then
            Print #RptHandle, "Bill Pay/OverPay";
          Else
            Print #RptHandle, "Bill OverPay";
          End If
        Case 12
          Print #RptHandle, "OverPayment";
        Case 13
          Print #RptHandle, "Adjust Bill Up Affecting Credit Balance";
        Case 14
          Print #RptHandle, "Adjust Pay Dwn Affecting Credit Balance";
        Case 15
          Print #RptHandle, "Adjust Prepay Down";
        Case 16
          Print #RptHandle, "Refund Prepay";
        Case 17
          Print #RptHandle, "PPTRA Removal";
        Case 18
          Print #RptHandle, "Unknown";
      End Select
      Dim Case11Tot As Double '7/6/06 added
      If x <> 11 Then
        Print #RptHandle, Tab(38); Using$("##,##0", TotCntByType(x)); Tab(55); Using$("$###,###,##0.00", TotByType(x))
      Else
        Case11Tot = OldRound(TotByType(11) - (GPrincTot + GIntTot + GAdvTot + GPenTot + GLateListTot))
        Case11Tot = OldRound(Case11Tot - (GOpt1Tot + GOpt2Tot + GOpt3Tot))
        Print #RptHandle, Tab(38); Using$("##,##0", TotCntByType(x)); Tab(55); Using$("$###,###,##0.00", Case11Tot)
      End If
      LineCnt = LineCnt + 1
      If fpcmbTransType.Text = " 0) All" Then GoTo All
      Select Case x
        Case 1
          If LineCnt >= MaxLines - 5 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Principal: "; Tab(55); Using$("$###,###,##0.00", GPrincTot)
          '7/11/06 added back int, adv and pen to accommodate manual bills
          Print #RptHandle, Tab(5); "Interest: "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising: "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", GLateListTot)
          Print #RptHandle, Tab(5); "Penalty: "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 2 '4
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
        Case 2
          If LineCnt >= MaxLines - 8 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Principal: "; Tab(55); Using$("$###,###,##0.00", GPrincTot)
          Print #RptHandle, Tab(5); "Interest: "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising: "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", GLateListTot)
          Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 5
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
        Case 3
          If LineCnt >= MaxLines - 8 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Principal: "; Tab(55); Using$("$###,###,##0.00", GPrincTot)
          Print #RptHandle, Tab(5); "Interest: "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising: "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", GLateListTot)
          Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 5
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
        Case 4
          GoTo All
        Case 5
          GoTo All
        Case 6
          GoTo All
        Case 7
          If LineCnt >= MaxLines - 8 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Principal: "; Tab(55); Using$("$###,###,##0.00", GPrincTot)
          Print #RptHandle, Tab(5); "Interest: "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising: "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", GLateListTot)
          Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 5
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
        Case 8
          GoTo All
        Case 9
          If LineCnt >= MaxLines - 8 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Principal: "; Tab(55); Using$("$###,###,##0.00", GPrincTot)
          Print #RptHandle, Tab(5); "Interest: "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising: "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", GLateListTot)
          Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 5
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
        Case 10
          If LineCnt >= MaxLines - 8 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Principal: "; Tab(55); Using$("$###,###,##0.00", GPrincTot)
          Print #RptHandle, Tab(5); "Interest: "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising: "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", GLateListTot)
          Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 5
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
       Case 11 '7/6/06 commented out GoTo All and added detail
'         GoTo All
          If LineCnt >= MaxLines - 8 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Principal: "; Tab(55); Using$("$###,###,##0.00", GPrincTot)
          Print #RptHandle, Tab(5); "Interest: "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising: "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", GLateListTot)
          Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 5
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
       Case 12
         GoTo All
       Case 13
          If LineCnt >= MaxLines - 8 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Principal: "; Tab(55); Using$("$###,###,##0.00", GPrincTot)
          Print #RptHandle, Tab(5); "Interest: "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising: "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", GLateListTot)
          Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 5
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
        Case 14
          If LineCnt >= MaxLines - 8 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Principal: "; Tab(55); Using$("$###,###,##0.00", GPrincTot)
          Print #RptHandle, Tab(5); "Interest: "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising: "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Late Listing: "; Tab(55); Using$("$###,###,##0.00", GLateListTot)
          Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 5
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
        Case 15
          GoTo All
        Case 16
          GoTo All
        Case 17
          GoTo All
        Case Else
          GoTo All
      End Select
    End If
All:
  Next x
  Print #RptHandle, String(94, "-")
  Print #RptHandle, "Grand Totals: "; Tab(38); Using$("##,##0", GCntTot); Tab(55); Using$("$###,###,##0.00", GAmtTot)
  
  Return
  
PrintAll:
  ThatName = ""
  PrePayDone = False
  For x = 1 To NumOfTCRecs
    If IdxFlag = False Then
      Get TCHandle, x, TaxCust
      CustBal = GetCustBalance(x, -1)
    Else
      Get TCHandle, IdxArray(x), TaxCust
      CustBal = GetCustBalance(IdxArray(x), -1)
    End If
    If TaxCust.Active = "N" And InactiveFlag = False Then
      GoTo SkipThisOne
    End If
    ThisName = QPTrim$(TaxCust.CustName)
    ReDim BillRec(1 To 1) As Long
    BillCnt = 0
    BadCnt = 0
    GoSub PrintPrepay
    ThisRec = TaxCust.LastTrans
    Do While ThisRec > 0
      Get TTHandle, ThisRec, TaxTrans
      If QFlag = True Then
        If TaxTrans.TransDate < BegDate Then
          BadCnt = BadCnt + 1
          If BadCnt > 3 Then Exit Do
        End If
      End If
      If TransDesc <> "" Then
        If InStr(1, TaxTrans.Description, TransDesc) = 0 Then GoTo Nope
      End If
      If TaxTrans.TransDate < BegDate Or TaxTrans.TransDate > EndDate Then GoTo Nope
      If TaxTrans.BillType <> "R" And TaxTrans.BillType <> "P" Then TaxTrans.BillType = ""
      If TaxTrans.BillType <> Mid(fpcmbTaxType.Text, 1, 1) And QPTrim$(TaxTrans.BillType) <> "" Then GoTo Nope
      If TaxTrans.TranType = 1 Then
        BillCnt = BillCnt + 1
        ReDim Preserve BillRec(1 To BillCnt) As Long
        BillRec(BillCnt) = ThisRec
      End If
Nope:
      ThisRec = TaxTrans.LastTrans
    Loop

    For z = 1 To BillCnt
      If ThisName <> ThatName Then
        GoSub PrintCustHeader
        ThatName = ThisName
      End If
      Get TTHandle, BillRec(z), TaxTrans
      If YrCnt = 0 Then
        YrCnt = YrCnt + 1
        ThisYear = YrCnt
        ReDim Preserve ThEYear(1 To YrCnt) As Integer
        ThEYear(YrCnt) = TaxTrans.TaxYear
        ReDim Preserve TotByYrAndType(1 To 18, 1 To YrCnt) As Double
        ReDim Preserve CntByYrAndType(1 To 18, 1 To YrCnt) As Integer
        For y = 1 To 18
          TotByYrAndType(y, YrCnt) = 0
          CntByYrAndType(y, YrCnt) = 0
        Next y
      Else
        For y = 1 To YrCnt
          If TaxTrans.TaxYear = ThEYear(y) Then
            ThisYear = y
            Exit For
          End If
        Next y
        If y > YrCnt Then
          YrCnt = YrCnt + 1
          ThisYear = YrCnt
          ReDim Preserve ThEYear(1 To YrCnt) As Integer
          ThEYear(YrCnt) = TaxTrans.TaxYear
          ReDim Preserve TotByYrAndType(1 To 18, 1 To YrCnt) As Double
          ReDim Preserve CntByYrAndType(1 To 18, 1 To YrCnt) As Integer
          For y = 1 To 18
            TotByYrAndType(y, YrCnt) = 0
            CntByYrAndType(y, YrCnt) = 0
          Next y
        End If
      End If
      ThisTransType = "Billing"
      TCnt = TCnt + 1
      If TaxTrans.DiscAmt > 0 Then GoSub ApplyDisc 'added 1/16/07
      TotAmt = OldRound(TotAmt + TaxTrans.Amount)
      PrincDif = OldRound(TaxTrans.Revenue.Principle1 - TaxTrans.Revenue.Principle1Pd)
      IntDif = OldRound(TaxTrans.Revenue.Interest - TaxTrans.Revenue.InterestPd)
      AdvDif = OldRound(TaxTrans.Revenue.Collection - TaxTrans.Revenue.CollectionPd)
      LateListDif = OldRound(TaxTrans.Revenue.LateList - TaxTrans.Revenue.LateListPd)
      PenDif = OldRound(TaxTrans.Revenue.Penalty - TaxTrans.Revenue.PenaltyPd)
      Opt1Dif = OldRound(TaxTrans.Revenue.RevOpt1 - TaxTrans.Revenue.RevOpt1Pd)
      Opt2Dif = OldRound(TaxTrans.Revenue.RevOpt2 - TaxTrans.Revenue.RevOpt2Pd)
      Opt3Dif = OldRound(TaxTrans.Revenue.RevOpt3 - TaxTrans.Revenue.RevOpt3Pd)
      BillBal = OldRound(PrincDif + IntDif + AdvDif + LateListDif + Opt1Dif + Opt2Dif + Opt3Dif)
      ThisBillNum = ParseBillNum(TaxTrans.Description)
      
      Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); QPTrim$(TaxTrans.Description);
      Print #RptHandle, Tab(37); Using$("###0", TaxTrans.TaxYear); Tab(45); Using$("$##,##0.00", TaxTrans.Revenue.PrePaidAmt);
      Print #RptHandle, Tab(56); Using$("$##,##0.00", TaxTrans.Amount);
      Print #RptHandle, Tab(79); ThisTransType
      LineCnt = LineCnt + 1
      If LineCnt >= MaxLines - 4 Then
        Print #RptHandle, FF$
        GoSub PrintHeader
        GoSub PrintCustHeader
      End If
      Print #RptHandle, Tab(3); "Oper #: " + CStr(TaxTrans.OperNum);
      Print #RptHandle, Tab(15); "Principle         "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle1); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle1Pd); Tab(80); Using$("$##,##0.00", PrincDif)
      Print #RptHandle, Tab(15); "Interest          "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Interest); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.InterestPd); Tab(80); Using$("$##,##0.00", IntDif)
      Print #RptHandle, Tab(15); "Advertising       "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Collection); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.CollectionPd); Tab(80); Using$("$##,##0.00", AdvDif)
      Print #RptHandle, Tab(15); "Late Listing      "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.LateList); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.LateListPd); Tab(80); Using$("$##,##0.00", LateListDif)
      Print #RptHandle, Tab(15); "Penalty           "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Penalty); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.PenaltyPd); Tab(80); Using$("$##,##0.00", PenDif)
      LineCnt = LineCnt + 5
      If LineCnt >= MaxLines - 3 Then
        Print #RptHandle, FF$
        GoSub PrintHeader
        GoSub PrintCustHeader
      End If
      If Len(Opt1Desc) > 0 Then
        Print #RptHandle, Tab(15); Opt1Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1Pd); Tab(80); Using$("$##,##0.00", Opt1Dif)
        LineCnt = LineCnt + 1
      End If
      If Len(Opt2Desc) > 0 Then
        Print #RptHandle, Tab(15); Opt2Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2Pd); Tab(80); Using$("$##,##0.00", Opt2Dif)
        LineCnt = LineCnt + 1
      End If
      If Len(Opt3Desc) > 0 Then
        Print #RptHandle, Tab(15); Opt3Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3Pd); Tab(80); Using$("$##,##0.00", Opt3Dif)
        LineCnt = LineCnt + 1
      End If
      Print #RptHandle, Tab(15); "Bill Balance:"; Tab(80); Using$("$##,##0.00", BillBal)
      
      ThisRec = TaxCust.LastTrans
      TotByYrAndType(1, ThisYear) = OldRound(TotByYrAndType(1, ThisYear) + TaxTrans.Amount)
      CntByYrAndType(1, ThisYear) = OldRound(CntByYrAndType(1, ThisYear) + 1)
      TotByType(1) = OldRound(TotByType(1) + TaxTrans.Amount)
      TotCntByType(1) = OldRound(TotCntByType(1) + 1)

      Do While ThisRec > 0
        Get TTHandle, ThisRec, TaxTrans
        If TaxTrans.TranType = 22 Then 'Prepay transactions can only be printed one time, not
        'for every iteration
          If PrePayDone = False Then
            GoTo PrepaySkip
          Else
            GoTo Nope2
          End If
        End If
        If TaxTrans.TranType = 11 Or TaxTrans.TranType = 12 Then GoTo Nope2
        If TaxTrans.BelongTo <> BillRec(z) Then GoTo Nope2
PrepaySkip:
        Select Case TaxTrans.TranType
          Case 2
            ThisTransType = "Payment"
            If TaxTrans.DiscAmt > 0 Then 'added 1/16/07
              GoSub ApplyDisc
            End If
            TotByYrAndType(2, ThisYear) = OldRound(TotByYrAndType(2, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(2, ThisYear) = OldRound(CntByYrAndType(2, ThisYear) + 1)
            TotByType(2) = OldRound(TotByType(2) + TaxTrans.Amount)
            TotCntByType(2) = OldRound(TotCntByType(2) + 1)
          Case 3
            ThisTransType = "Release"
            TotByYrAndType(3, ThisYear) = OldRound(TotByYrAndType(3, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(3, ThisYear) = OldRound(CntByYrAndType(3, ThisYear) + 1)
            TotByType(3) = OldRound(TotByType(3) + TaxTrans.Amount)
            TotCntByType(3) = OldRound(TotCntByType(3) + 1)
          Case 4
            ThisTransType = "Interest"
            TotByYrAndType(4, ThisYear) = OldRound(TotByYrAndType(4, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(4, ThisYear) = OldRound(CntByYrAndType(4, ThisYear) + 1)
            TotByType(4) = OldRound(TotByType(4) + TaxTrans.Amount)
            TotCntByType(4) = OldRound(TotCntByType(4) + 1)
          Case 5
            ThisTransType = "Penalty"
            TotByYrAndType(5, ThisYear) = OldRound(TotByYrAndType(5, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(5, ThisYear) = OldRound(CntByYrAndType(5, ThisYear) + 1)
            TotByType(5) = OldRound(TotByType(5) + TaxTrans.Amount)
            TotCntByType(5) = OldRound(TotCntByType(5) + 1)
          Case 6
            ThisTransType = "Advertising Charge"
            TotByYrAndType(6, ThisYear) = OldRound(TotByYrAndType(6, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(6, ThisYear) = OldRound(CntByYrAndType(6, ThisYear) + 1)
            TotByType(6) = OldRound(TotByType(6) + TaxTrans.Amount)
            TotCntByType(6) = OldRound(TotCntByType(6) + 1)
          Case 7
            ThisTransType = "Adjust Pay Down"
            TotByYrAndType(7, ThisYear) = OldRound(TotByYrAndType(7, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(7, ThisYear) = OldRound(CntByYrAndType(7, ThisYear) + 1)
            TotByType(7) = OldRound(TotByType(7) + TaxTrans.Amount)
            TotCntByType(7) = OldRound(TotCntByType(7) + 1)
          Case 9
            ThisTransType = "Credit Applied at Billing"
            TotByYrAndType(8, ThisYear) = OldRound(TotByYrAndType(8, ThisYear) + TaxTrans.Revenue.PrePaidUsed)
            CntByYrAndType(8, ThisYear) = OldRound(CntByYrAndType(8, ThisYear) + 1)
            TotByType(8) = OldRound(TotByType(8) + TaxTrans.Revenue.PrePaidUsed)
            TotCntByType(8) = OldRound(TotCntByType(8) + 1)
          Case 13
            ThisTransType = "Adjust Bill Down"
            TotByYrAndType(9, ThisYear) = OldRound(TotByYrAndType(9, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(9, ThisYear) = OldRound(CntByYrAndType(9, ThisYear) + 1)
            TotByType(9) = OldRound(TotByType(9) + TaxTrans.Amount)
            TotCntByType(9) = OldRound(TotCntByType(9) + 1)
          Case 14
            ThisTransType = "Adjust Bill Up"
            TotByYrAndType(10, ThisYear) = OldRound(TotByYrAndType(10, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(10, ThisYear) = OldRound(CntByYrAndType(10, ThisYear) + 1)
            TotByType(10) = OldRound(TotByType(10) + TaxTrans.Amount)
            TotCntByType(10) = OldRound(TotCntByType(10) + 1)
          Case 21
            ThisTransType = "Billpay/Overpay"
            If TaxTrans.DiscAmt > 0 Then 'added 1/16/07
              GoSub ApplyDisc
            End If
            TotByYrAndType(11, ThisYear) = OldRound(TotByYrAndType(11, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(11, ThisYear) = OldRound(CntByYrAndType(11, ThisYear) + 1)
            TotByType(11) = OldRound(TotByType(11) + TaxTrans.Amount)
            TotCntByType(11) = OldRound(TotCntByType(11) + 1)
          Case 22
            ThisTransType = "Overpayment"
            TotByYrAndType(12, ThisYear) = OldRound(TotByYrAndType(12, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(12, ThisYear) = OldRound(CntByYrAndType(12, ThisYear) + 1)
            TotByType(12) = OldRound(TotByType(12) + TaxTrans.Amount)
            TotCntByType(12) = OldRound(TotCntByType(12) + 1)
          Case 24
            ThisTransType = "Adj Bill Up: -Credit Bal"
            TotByYrAndType(13, ThisYear) = OldRound(TotByYrAndType(13, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(13, ThisYear) = OldRound(CntByYrAndType(13, ThisYear) + 1)
            TotByType(13) = OldRound(TotByType(13) + TaxTrans.Amount)
            TotCntByType(13) = OldRound(TotCntByType(13) + 1)
          Case 10
            ThisTransType = "Adj Pay Dwn: -Credit Bal"
            TotByYrAndType(14, ThisYear) = OldRound(TotByYrAndType(14, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(14, ThisYear) = OldRound(CntByYrAndType(14, ThisYear) + 1)
            TotByType(14) = OldRound(TotByType(14) + TaxTrans.Amount)
            TotCntByType(14) = OldRound(TotCntByType(14) + 1)
          Case 30
            ThisTransType = "PPTRA Removal"
            TotByYrAndType(15, ThisYear) = OldRound(TotByYrAndType(15, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(15, ThisYear) = OldRound(CntByYrAndType(15, ThisYear) + 1)
            TotByType(15) = OldRound(TotByType(15) + TaxTrans.Amount)
            TotCntByType(15) = OldRound(TotCntByType(15) + 1)
'          Case 11
'            ThisTransType = "Adjust Prepay Down"
'            TotByYrAndType(15, ThisYear) = OldRound(TotByYrAndType(15, ThisYear) + TaxTrans.Amount)
'            CntByYrAndType(15, ThisYear) = OldRound(CntByYrAndType(15, ThisYear) + 1)
'          Case 12
'            ThisTransType = "Refund Prepay"
'            TotByYrAndType(16, ThisYear) = OldRound(TotByYrAndType(16, ThisYear) + TaxTrans.Amount)
'            CntByYrAndType(16, ThisYear) = OldRound(CntByYrAndType(16, ThisYear) + 1)
          Case Else
            ThisTransType = "Unknown"
            TotByYrAndType(17, ThisYear) = OldRound(TotByYrAndType(17, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(17, ThisYear) = OldRound(CntByYrAndType(17, ThisYear) + 1)
            TotByType(17) = OldRound(TotByType(17) + TaxTrans.Amount)
            TotCntByType(17) = OldRound(TotCntByType(17) + 1)
        End Select
        TCnt = TCnt + 1
        If LineCnt >= MaxLines - 2 Then
          Print #RptHandle, FF$
          GoSub PrintHeader
          GoSub PrintCustHeader
        End If
        If TaxTrans.TranType <> 11 Then
          Print #RptHandle, Tab(15); String(80, "^")
        Else
          Print #RptHandle, Tab(15); String(80, "-")
        End If
        Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); QPTrim$(TaxTrans.Description);
        Print #RptHandle, Tab(37); Using$("###0", TaxTrans.TaxYear); Tab(45); Using$("$##,##0.00", TaxTrans.Revenue.PrePaidAmt);
        If TaxTrans.TranType <> 9 Then
          Print #RptHandle, Tab(56); Using$("$##,##0.00", TaxTrans.Amount); Tab(79); ThisTransType
        Else
          Print #RptHandle, Tab(56); Using$("$##,##0.00", TaxTrans.Revenue.PrePaidUsed); Tab(79); ThisTransType
        End If
        
        LineCnt = LineCnt + 2
        If LineCnt >= MaxLines - 4 Then
          Print #RptHandle, FF$
          GoSub PrintHeader
          GoSub PrintCustHeader
        End If
        Print #RptHandle, Tab(3); "Oper #: " + CStr(TaxTrans.OperNum);
        Print #RptHandle, Tab(15); "Principle         "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle1); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle1Pd); Tab(75); "Belongs to Bill#"
        If TaxTrans.TranType = 11 Then
          Print #RptHandle, Tab(15); "Interest          "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Interest); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.InterestPd); Tab(75); "NA"
        Else
          Print #RptHandle, Tab(15); "Interest          "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Interest); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.InterestPd); Tab(75); QPTrim$(ThisBillNum)
        End If
        Print #RptHandle, Tab(15); "Advertising       "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Collection); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.CollectionPd)
        Print #RptHandle, Tab(15); "Late Listing      "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.LateList); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.LateListPd)
        Print #RptHandle, Tab(15); "Penalty           "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Penalty); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.PenaltyPd)
        LineCnt = LineCnt + 5
        If LineCnt >= MaxLines - 3 Then
          Print #RptHandle, FF$
          GoSub PrintHeader
          GoSub PrintCustHeader
        End If
        If Len(Opt1Desc) > 0 Then
          Print #RptHandle, Tab(15); Opt1Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1Pd)
          LineCnt = LineCnt + 1
        End If
        If Len(Opt2Desc) > 0 Then
          Print #RptHandle, Tab(15); Opt2Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2Pd)
          LineCnt = LineCnt + 1
        End If
        If Len(Opt3Desc) > 0 Then
          Print #RptHandle, Tab(15); Opt3Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3Pd)
          LineCnt = LineCnt + 1
        End If
Nope2:
        ThisRec = TaxTrans.LastTrans
      Loop
      PrePayDone = True 'added 7/10/06
    Next z
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
SkipThisOne:
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  If TCnt = 0 Then
    Call TaxMsg(900, "No transactions were found that fit the parameters entered.")
    Close
    Exit Sub
  End If
  
  If YrCnt > 0 Then
    GoSub SortIt
    GoSub PrintTotals
  End If
  Print #RptHandle, FF$
  Close
  ViewPrint RptFile, "Tax Transactions Report", True
  Exit Sub
  
PrintPrepay:
  ThisRec = TaxCust.LastTrans
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
    If TaxTrans.TranType <> 11 And TaxTrans.TranType <> 12 Then GoTo NotThisTrans
    If TaxTrans.TransDate < BegDate Or TaxTrans.TransDate > EndDate Then GoTo NotThisTrans
    If ThisName <> ThatName Then
      GoSub PrintCustHeader
      ThatName = ThisName
    End If
    If TaxTrans.TranType = 11 Then
      ThisTransType = "Adj Prepay Down"
    ElseIf TaxTrans.TranType = 12 Then
      ThisTransType = "Refund Prepay"
    End If
    If YrCnt = 0 Then
      YrCnt = YrCnt + 1
      ThisYear = YrCnt
      ReDim Preserve ThEYear(1 To YrCnt) As Integer
      ThEYear(YrCnt) = TaxTrans.TaxYear
      ReDim Preserve TotByYrAndType(1 To 18, 1 To YrCnt) As Double
      ReDim Preserve CntByYrAndType(1 To 18, 1 To YrCnt) As Integer
      For y = 1 To 18
        TotByYrAndType(y, YrCnt) = 0
        CntByYrAndType(y, YrCnt) = 0
      Next y
    Else
      For y = 1 To YrCnt
        If TaxTrans.TaxYear = ThEYear(y) Then
          ThisYear = y
          Exit For
          End If
      Next y
      If y > YrCnt Then
        YrCnt = YrCnt + 1
        ThisYear = YrCnt
        ReDim Preserve ThEYear(1 To YrCnt) As Integer
        ThEYear(YrCnt) = TaxTrans.TaxYear
        ReDim Preserve TotByYrAndType(1 To 18, 1 To YrCnt) As Double
        ReDim Preserve CntByYrAndType(1 To 18, 1 To YrCnt) As Integer
        For y = 1 To 18
          TotByYrAndType(y, YrCnt) = 0
          CntByYrAndType(y, YrCnt) = 0
        Next y
      End If
    End If
    If TaxTrans.TranType = 11 Then
      TotByYrAndType(15, ThisYear) = OldRound(TotByYrAndType(15, ThisYear) + TaxTrans.Amount)
      CntByYrAndType(15, ThisYear) = OldRound(CntByYrAndType(15, ThisYear) + 1)
    ElseIf TaxTrans.TranType = 12 Then
      TotByYrAndType(16, ThisYear) = OldRound(TotByYrAndType(16, ThisYear) + TaxTrans.Amount)
      CntByYrAndType(16, ThisYear) = OldRound(CntByYrAndType(16, ThisYear) + 1)
    End If
    TCnt = TCnt + 1
    TotAmt = OldRound(TotAmt + TaxTrans.Amount)
    PrincDif = 0
    IntDif = 0
    AdvDif = 0
    LateListDif = 0
    Opt1Dif = 0
    Opt2Dif = 0
    Opt3Dif = 0
    BillBal = OldRound(PrincDif + IntDif + AdvDif + LateListDif + Opt1Dif + Opt2Dif + Opt3Dif)
    ThisBillNum = "NA"
    Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); QPTrim$(TaxTrans.Description);
    Print #RptHandle, Tab(37); Using$("###0", TaxTrans.TaxYear); Tab(45); Using$("$##,##0.00", TaxTrans.Revenue.PrePaidAmt);
    Print #RptHandle, Tab(56); Using$("$##,##0.00", TaxTrans.Amount); 'Tab(69);
    Print #RptHandle, Tab(79); ThisTransType
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines - 4 Then
      Print #RptHandle, FF$
      GoSub PrintHeader
      GoSub PrintCustHeader
    End If
    Print #RptHandle, Tab(15); "Principle         "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle1); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle1Pd); Tab(75); "Belongs to Bill#"
    Print #RptHandle, Tab(15); "Interest          "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Interest); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.InterestPd); Tab(75); "NA"
    Print #RptHandle, Tab(15); "Advertising       "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Collection); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.CollectionPd)
    Print #RptHandle, Tab(15); "Late Listing      "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.LateList); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.LateListPd)
    Print #RptHandle, Tab(15); "Penalty           "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Penalty); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.PenaltyPd)
    LineCnt = LineCnt + 5
    If LineCnt >= MaxLines - 3 Then
      Print #RptHandle, FF$
      GoSub PrintHeader
      GoSub PrintCustHeader
    End If
    If Len(Opt1Desc) > 0 Then
      Print #RptHandle, Tab(15); Opt1Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1Pd)
      LineCnt = LineCnt + 1
    End If
    If Len(Opt2Desc) > 0 Then
      Print #RptHandle, Tab(15); Opt2Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2Pd)
      LineCnt = LineCnt + 1
    End If
    If Len(Opt3Desc) > 0 Then
      Print #RptHandle, Tab(15); Opt3Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3Pd)
      LineCnt = LineCnt + 1
    End If
    Print #RptHandle, Tab(15); String(80, "x")
    If LineCnt >= MaxLines - 2 Then
      Print #RptHandle, FF$
      GoSub PrintHeader
      GoSub PrintCustHeader
    End If
NotThisTrans:
    ThisRec = TaxTrans.LastTrans
  Loop
  Return
  
ApplyDisc: '1/16/07
  Disc1 = 0
  Disc2 = 0
  Disc3 = 0
  Disc4 = 0
  If TaxTrans.Amount = 0 Then Return
  If TaxTrans.TranType = 1 Then
    SaveAmt = OldRound(TaxTrans.Amount - TaxTrans.DiscAmt)
  Else
    SaveAmt = TaxTrans.Amount
    TaxTrans.Amount = OldRound(TaxTrans.Amount + TaxTrans.DiscAmt)
  End If
  Disc1 = OldRound(TaxTrans.Revenue.Principle1Pd / SaveAmt)
  Disc1 = OldRound(Disc1 * TaxTrans.DiscAmt)
  Disc2 = OldRound(TaxTrans.Revenue.RevOpt1Pd / SaveAmt)
  Disc2 = OldRound(Disc2 * TaxTrans.DiscAmt)
  Disc3 = OldRound(TaxTrans.Revenue.RevOpt2Pd / SaveAmt)
  Disc3 = OldRound(Disc3 * TaxTrans.DiscAmt)
  Disc4 = OldRound(TaxTrans.Revenue.RevOpt3Pd / SaveAmt)
  Disc4 = OldRound(Disc4 * TaxTrans.DiscAmt)
  TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd + Disc1)
  TaxTrans.Revenue.RevOpt1Pd = OldRound(TaxTrans.Revenue.RevOpt1Pd + Disc2)
  TaxTrans.Revenue.RevOpt2Pd = OldRound(TaxTrans.Revenue.RevOpt2Pd + Disc3)
  TaxTrans.Revenue.RevOpt3Pd = OldRound(TaxTrans.Revenue.RevOpt3Pd + Disc4)
  DiscApplied = True
  
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxTransJournal", "PrintRTextDet", Erl)
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

Private Sub fptxtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fpcmbTransType.SetFocus
  ElseIf KeyCode = vbKeyUp Then
    fpcmbPrintOrder.SetFocus
  End If
End Sub

Private Sub PrintPGraphicsDet()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim x As Long, y As Integer
  Dim BegDate As Integer
  Dim EndDate As Integer
  Dim dlm$
  Dim ThisRec As Long
  Dim ThisClass As Integer
  Dim ThisType As String
  Dim RptFile$
  Dim RptHandle As Integer
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim InactiveFlag As Boolean
  Dim ThisName$, ThisBillType$
  Dim TCnt As Long
  Dim TotAmt As Double
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim NumOfSrchRecs As Long
  Dim ThisTransType As String
  Dim YrCnt As Integer, ThisYear As Integer
  Dim SubRptFile$
  Dim SubRptHandle As Integer
  Dim Sub2RptFile$
  Dim Sub2RptHandle As Integer
  Dim BigYr As Integer
  Dim HoldBigYr As Integer
  Dim HoldYr As Integer
  Dim HoldPers As Double
  Dim HoldMT As Double
  Dim HoldMC As Double
  Dim HoldFE As Double
  Dim HoldMH As Double
  Dim HoldInt As Double
  Dim HoldAdv As Double
  Dim HoldLateList As Double
  Dim HoldPen As Double
  Dim HoldOpt1 As Double
  Dim HoldOpt2 As Double
  Dim HoldOpt3 As Double
  Dim Nextx As Integer
  Dim Thisx As Integer
  Dim Nexty As Integer
  Dim Thisy As Integer
  Dim z As Integer
  Dim PersDif As Double
  Dim IntDif As Double
  Dim MTDif As Double
  Dim MCDif As Double
  Dim FEDif As Double
  Dim MHDif As Double
  Dim PenDif As Double
  Dim Opt1Dif As Double
  Dim Opt2Dif As Double
  Dim Opt3Dif As Double
  Dim Opt1Desc$
  Dim Opt2Desc$, BillCnt As Integer
  Dim Opt3Desc$, ThisBillRec As Long
  Dim POpt1Desc$
  Dim POpt2Desc$
  Dim POpt3Desc$
  Dim CustBal As Double
  Dim BillBal As Double
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim OptFlag As Boolean
  Dim TransDesc$
  Dim GAmtTot As Double
  Dim GCntTot As Long
  Dim GBillCnt As Long
  Dim BadCnt As Integer
  Dim QFlag As Boolean
  Dim GPersTot As Double
  Dim GMTTot As Double
  Dim GMCTot As Double
  Dim GFETot As Double
  Dim GMHTot As Double
  Dim GPenTot As Double
  Dim GIntTot As Double
  Dim GAdvTot As Double
  Dim GLateListTot As Double
  Dim GOpt1Tot As Double
  Dim GOpt2Tot As Double
  Dim GOpt3Tot As Double
  Dim TotRev As Double
  Dim TotYearCnt As Long
  Dim TotYearAmt As Double
  Dim PrePayDone As Boolean 'added 7/10/06
  Dim WhatsLeft#
  Dim ThisDif#
  Dim PersPd#
  Dim MTPd#
  Dim MCPd#
  Dim FEPd#
  Dim MHPd#
  Dim RevOpt1Pd#
  Dim RevOpt2Pd#
  Dim RevOpt3Pd#
  Dim Disc1 As Double '1/16/2007
  Dim Disc2 As Double '1/16/2007
  Dim Disc3 As Double '1/16/2007
  Dim Disc4 As Double '1/16/2007
  Dim Disc5 As Double '1/16/2007
  Dim Disc6 As Double '1/16/2007
  Dim Disc7 As Double '1/16/2007
  Dim Disc8 As Double '1/16/2007
  Dim DiscApplied As Boolean '1/16/2007
  Dim SaveAmt As Double '1/16/2007
  
  On Error GoTo ERRORSTUFF
  
  PrePayDone = False '7/10/06
  DiscApplied = False 'added 1/16/07
  QFlag = False
  If chkQuick.Value = 1 Then QFlag = True
  
  TransDesc = QPTrim$(fptxtDesc.Text)
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  POpt1Desc = QPTrim$(TaxMasterRec.POptRev1)
  POpt2Desc = QPTrim$(TaxMasterRec.POptRev2)
  POpt3Desc = QPTrim$(TaxMasterRec.POptRev3)
  IdxFlag = False
  OptFlag = False
  
  If CheckB4Printing = False Then
    Exit Sub
  End If
  ThisBillType = QPTrim$(fpcmbTaxType.Text)
  If fpcmbIncInactive.Text = "No" Then
    InactiveFlag = False
  Else
    InactiveFlag = True
  End If
  
  dlm$ = "~"
  If Mid(fpcmbTransType.Text, 1, 1) = "" Then
    ThisClass = CInt(Mid(fpcmbTransType.Text, 2, 1))
  Else
    ThisClass = CInt(Mid(fpcmbTransType.Text, 1, 2))
  End If
    
  Select Case ThisClass
    Case 1
      ThisType = "Billing"
    Case 2
      ThisType = "Payment"
    Case 3
      ThisType = "Release"
    Case 4
      ThisType = "Interest"
    Case 5
      ThisType = "Penalty"
    Case 6
      ThisType = "Advertising Charge"
    Case 7, 10
      ThisType = "Adjust Pay Down"
    Case 9
      ThisType = "Credit At Billing"
    Case 11
      ThisType = "Adjust Prepay Down"
    Case 12
      ThisType = "Refund Prepay"
    Case 13
      ThisType = "Adjust Bill Down"
    Case 14, 24
      ThisType = "Adjust Bill Up"
    Case 21
      ThisType = "Payment w/Overpay"
    Case 22
      ThisType = "Overpayment Only"
    Case 30
      ThisType = "PPTRA Removal"
    Case Else
      ThisType = "All"
  End Select
    
  BegDate = Date2Num(fptxtBegDate.Text)
  EndDate = Date2Num(fptxtEndDate.Text)
  
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
    OptFlag = True
  End If
  
  RptFile$ = "TAXRPTS\TXJRLDT.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  
  frmVATaxShowPctComp.Label1 = "Gathering Tax Transaction Data"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  
  ReDim TotByYrAndType(1 To 18, 1 To 1) As Double
  ReDim CntByYrAndType(1 To 18, 1 To 1) As Integer
  ReDim TotByType(1 To 18) As Double
  ReDim TotCntByType(1 To 18) As Long
  ReDim ThEYear(1 To 1) As Integer
  
  ReDim TotByYrAndPers(1 To 1) As Double
  ReDim TotByYrAndMT(1 To 1) As Double
  ReDim TotByYrAndMC(1 To 1) As Double
  ReDim TotByYrAndFE(1 To 1) As Double
  ReDim TotByYrAndMH(1 To 1) As Double
  ReDim TotByYrAndInt(1 To 1) As Double
  ReDim TotByYrAndAdv(1 To 1) As Double
  ReDim TotByYrAndLateList(1 To 1) As Double
  ReDim TotByYrAndPen(1 To 1) As Double
  ReDim TotByYrAndOpt1(1 To 1) As Double
  ReDim TotByYrAndOpt2(1 To 1) As Double
  ReDim TotByYrAndOpt3(1 To 1) As Double
  
  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  End If
  
  If InStr(fpcmbTransType.Text, "All") Then GoTo PrintAll
  If fpcmbTransType.Text <> "All" Then
    For x = 1 To NumOfTCRecs
      If IdxFlag = False Then
        Get TCHandle, x, TaxCust
      Else
        Get TCHandle, IdxArray(x), TaxCust
      End If
      If TaxCust.Active = "N" And InactiveFlag = False Then
        GoTo SkipIt
      End If
      ThisName = QPTrim$(TaxCust.CustName)
      ThisRec = TaxCust.LastTrans
      BadCnt = 0
      Do While ThisRec > 0
        Get TTHandle, ThisRec, TaxTrans
        If QFlag = True Then
          If TaxTrans.TransDate < BegDate Then
            BadCnt = BadCnt + 1
            If BadCnt > 3 Then Exit Do
          End If
        End If
        If TransDesc <> "" Then
          If InStr(1, TaxTrans.Description, TransDesc) = 0 Then GoTo SkipIt
        End If
        If TaxTrans.BillType <> "R" And TaxTrans.BillType <> "P" Then TaxTrans.BillType = ""
        If TaxTrans.BillType <> Mid(fpcmbTaxType.Text, 1, 1) And QPTrim$(TaxTrans.BillType) <> "" Then
          GoTo SkipIt
        End If
        If ThisClass = 7 And TaxTrans.TranType = 10 Then GoTo ItsOK
        If ThisClass = 14 And TaxTrans.TranType = 24 Then GoTo ItsOK
'        If ThisClass = 2 And TaxTrans.TranType = 21 Then GoTo ItsOK 'commented out on 7/6/06
        If ThisClass <> 0 And TaxTrans.TranType <> ThisClass Then GoTo SkipIt
ItsOK:
        If TaxTrans.TransDate >= BegDate And TaxTrans.TransDate <= EndDate Then
          If TaxTrans.BelongTo > 0 Then
            Get TTHandle, TaxTrans.BelongTo, TaxTrans
          End If
        
          If YrCnt = 0 Then
            YrCnt = YrCnt + 1
            ThisYear = YrCnt
            ReDim Preserve ThEYear(1 To YrCnt) As Integer
            ThEYear(YrCnt) = TaxTrans.TaxYear
            ReDim Preserve TotByYrAndType(1 To 18, 1 To YrCnt) As Double
            ReDim Preserve CntByYrAndType(1 To 18, 1 To YrCnt) As Integer
            ReDim Preserve TotByYrAndPers(1 To YrCnt) As Double
            ReDim Preserve TotByYrAndMT(1 To YrCnt) As Double
            ReDim Preserve TotByYrAndMC(1 To YrCnt) As Double
            ReDim Preserve TotByYrAndFE(1 To YrCnt) As Double
            ReDim Preserve TotByYrAndMH(1 To YrCnt) As Double
            ReDim Preserve TotByYrAndInt(1 To YrCnt) As Double
            ReDim Preserve TotByYrAndAdv(1 To YrCnt) As Double
            ReDim Preserve TotByYrAndLateList(1 To YrCnt) As Double
            ReDim Preserve TotByYrAndPen(1 To YrCnt) As Double
            ReDim Preserve TotByYrAndOpt1(1 To YrCnt) As Double
            ReDim Preserve TotByYrAndOpt2(1 To YrCnt) As Double
            ReDim Preserve TotByYrAndOpt3(1 To YrCnt) As Double
            
            TotByYrAndPers(YrCnt) = 0
            TotByYrAndMT(YrCnt) = 0
            TotByYrAndMC(YrCnt) = 0
            TotByYrAndFE(YrCnt) = 0
            TotByYrAndMH(YrCnt) = 0
            TotByYrAndInt(YrCnt) = 0
            TotByYrAndAdv(YrCnt) = 0
            TotByYrAndLateList(YrCnt) = 0
            TotByYrAndPen(YrCnt) = 0
            TotByYrAndOpt1(YrCnt) = 0
            TotByYrAndOpt2(YrCnt) = 0
            TotByYrAndOpt3(YrCnt) = 0
            
            For y = 1 To 18
              TotByYrAndType(y, YrCnt) = 0
              CntByYrAndType(y, YrCnt) = 0
            Next y
          Else
            For y = 1 To YrCnt
              If TaxTrans.TaxYear = ThEYear(y) Then
                ThisYear = y
                Exit For
              End If
            Next y
            If y > YrCnt Then
              YrCnt = YrCnt + 1
              ThisYear = YrCnt
              ReDim Preserve ThEYear(1 To YrCnt) As Integer
              ThEYear(YrCnt) = TaxTrans.TaxYear
              ReDim Preserve TotByYrAndType(1 To 18, 1 To YrCnt) As Double
              ReDim Preserve CntByYrAndType(1 To 18, 1 To YrCnt) As Integer
              ReDim Preserve TotByYrAndPers(1 To YrCnt) As Double
              ReDim Preserve TotByYrAndMT(1 To YrCnt) As Double
              ReDim Preserve TotByYrAndMC(1 To YrCnt) As Double
              ReDim Preserve TotByYrAndFE(1 To YrCnt) As Double
              ReDim Preserve TotByYrAndMH(1 To YrCnt) As Double
              ReDim Preserve TotByYrAndInt(1 To YrCnt) As Double
              ReDim Preserve TotByYrAndAdv(1 To YrCnt) As Double
              ReDim Preserve TotByYrAndLateList(1 To YrCnt) As Double
              ReDim Preserve TotByYrAndPen(1 To YrCnt) As Double
              ReDim Preserve TotByYrAndPers(1 To YrCnt) As Double
              ReDim Preserve TotByYrAndOpt1(1 To YrCnt) As Double
              ReDim Preserve TotByYrAndOpt2(1 To YrCnt) As Double
              ReDim Preserve TotByYrAndOpt3(1 To YrCnt) As Double
              TotByYrAndPers(YrCnt) = 0
              TotByYrAndMT(YrCnt) = 0
              TotByYrAndMC(YrCnt) = 0
              TotByYrAndFE(YrCnt) = 0
              TotByYrAndMH(YrCnt) = 0
              TotByYrAndInt(YrCnt) = 0
              TotByYrAndAdv(YrCnt) = 0
              TotByYrAndLateList(YrCnt) = 0
              TotByYrAndPen(YrCnt) = 0
              TotByYrAndOpt1(YrCnt) = 0
              TotByYrAndOpt2(YrCnt) = 0
              TotByYrAndOpt3(YrCnt) = 0
              For y = 1 To 18
                TotByYrAndType(y, YrCnt) = 0
                CntByYrAndType(y, YrCnt) = 0
              Next y
           End If
         End If
         Get TTHandle, ThisRec, TaxTrans
         DiscApplied = False 'added 1/16/07
         Select Case TaxTrans.TranType
           Case 1
             ThisTransType = "Billing"
             TotByYrAndType(1, ThisYear) = OldRound(TotByYrAndType(1, ThisYear) + TaxTrans.Amount)
             TotByType(1) = OldRound(TotByType(1) + TaxTrans.Amount)
             CntByYrAndType(1, ThisYear) = OldRound(CntByYrAndType(1, ThisYear) + 1)
             TotCntByType(1) = OldRound(TotCntByType(1) + 1)
             TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1)
             TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2)
             TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3)
             TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4)
             TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5)
             '7/11/06 added back interest, advertising and penalty to accommodate manual bills
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.Interest)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.Collection)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.Penalty)
             
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateList)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3)
           Case 2
             ThisTransType = "Payment"
             If TaxTrans.DiscAmt > 0 Then 'added 1/16/07
               GoSub ApplyDisc
             End If
             TotByYrAndType(2, ThisYear) = OldRound(TotByYrAndType(2, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(2, ThisYear) = OldRound(CntByYrAndType(2, ThisYear) + 1)
             TotByType(2) = OldRound(TotByType(2) + TaxTrans.Amount)
             TotCntByType(2) = OldRound(TotCntByType(2) + 1)
             '                                                                                             ________________________1/30/07____________
             TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + (TaxTrans.Revenue.Principle1Pd + TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl)) '1/30/07
             TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2Pd)
             TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3Pd)
             TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4Pd)
             TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5Pd)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
           Case 3
            '7/12/06 changed revenues for Release to paid from charged
             ThisTransType = "Release"
             TotByYrAndType(3, ThisYear) = OldRound(TotByYrAndType(3, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(3, ThisYear) = OldRound(CntByYrAndType(3, ThisYear) + 1)
             TotByType(3) = OldRound(TotByType(3) + TaxTrans.Amount)
             TotCntByType(3) = OldRound(TotCntByType(3) + 1)
             TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2Pd)
             TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3Pd)
             TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4Pd)
             TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5Pd)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
           Case 4
             ThisTransType = "Interest"
             TotByYrAndType(4, ThisYear) = OldRound(TotByYrAndType(4, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(4, ThisYear) = OldRound(CntByYrAndType(4, ThisYear) + 1)
             TotByType(4) = OldRound(TotByType(4) + TaxTrans.Amount)
             TotCntByType(4) = OldRound(TotCntByType(4) + 1)
           Case 5
             ThisTransType = "Penalty"
             TotByYrAndType(5, ThisYear) = OldRound(TotByYrAndType(5, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(5, ThisYear) = OldRound(CntByYrAndType(5, ThisYear) + 1)
             TotByType(5) = OldRound(TotByType(5) + TaxTrans.Amount)
             TotCntByType(5) = OldRound(TotCntByType(5) + 1)
           Case 6
             ThisTransType = "Advertising Charge"
             TotByYrAndType(6, ThisYear) = OldRound(TotByYrAndType(6, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(6, ThisYear) = OldRound(CntByYrAndType(6, ThisYear) + 1)
             TotByType(6) = OldRound(TotByType(6) + TaxTrans.Amount)
             TotCntByType(6) = OldRound(TotCntByType(6) + 1)
           Case 7
             ThisTransType = "Adjust Pay Down"
             TotByYrAndType(7, ThisYear) = OldRound(TotByYrAndType(7, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(7, ThisYear) = OldRound(CntByYrAndType(7, ThisYear) + 1)
             TotByType(7) = OldRound(TotByType(7) + TaxTrans.Amount)
             TotCntByType(7) = OldRound(TotCntByType(7) + 1)
             TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2Pd)
             TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3Pd)
             TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4Pd)
             TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5Pd)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
           Case 9
             ThisTransType = "Credit Applied at Billing"
             TotByYrAndType(8, ThisYear) = OldRound(TotByYrAndType(8, ThisYear) + TaxTrans.Revenue.PrePaidUsed)
             CntByYrAndType(8, ThisYear) = OldRound(CntByYrAndType(8, ThisYear) + 1)
'             TotByType(8) = OldRound(TotByType(8) + TaxTrans.Amount)
             TotByType(8) = OldRound(TotByType(8) + TaxTrans.Revenue.PrePaidUsed)
             TotCntByType(8) = OldRound(TotCntByType(8) + 1)
           Case 13
             ThisTransType = "Adjust Bill Down"
             TotByYrAndType(9, ThisYear) = OldRound(TotByYrAndType(9, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(9, ThisYear) = OldRound(CntByYrAndType(9, ThisYear) + 1)
             TotByType(9) = OldRound(TotByType(9) + TaxTrans.Amount)
             TotCntByType(9) = OldRound(TotCntByType(9) + 1)
             TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1)
             TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2)
             TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3)
             TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4)
             TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.Interest)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.Collection)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateList)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.Penalty)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3)
           Case 14
             ThisTransType = "Adjust Bill Up"
             TotByYrAndType(10, ThisYear) = OldRound(TotByYrAndType(10, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(10, ThisYear) = OldRound(CntByYrAndType(10, ThisYear) + 1)
             TotByType(10) = OldRound(TotByType(10) + TaxTrans.Amount)
             TotCntByType(10) = OldRound(TotCntByType(10) + 1)
             TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1)
             TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2)
             TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3)
             TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4)
             TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.Interest)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.Collection)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateList)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.Penalty)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3)
           Case 21
             ThisTransType = "Bill Pay/Overpay" '7/6/06  changed Amount to PrePaidAmt
             If TaxTrans.DiscAmt > 0 Then 'added 1/16/07
               GoSub ApplyDisc
             End If
             If fpcmbTransType.Text <> " 0) All" Then 'added the All if statement on 7/7/06
               TotByYrAndType(11, ThisYear) = OldRound(TotByYrAndType(11, ThisYear) + TaxTrans.Revenue.PrePaidAmt)
             Else
               TotByYrAndType(11, ThisYear) = OldRound(TotByYrAndType(11, ThisYear) + TaxTrans.Amount)
             End If
'             TotByYrAndType(11, ThisYear) = OldRound(TotByYrAndType(11, ThisYear) + TaxTrans.Revenue.PrePaidAmt) '  Amount)
             CntByYrAndType(11, ThisYear) = OldRound(CntByYrAndType(11, ThisYear) + 1)
             TotByType(11) = OldRound(TotByType(11) + TaxTrans.Amount)
             TotCntByType(11) = OldRound(TotCntByType(11) + 1)
             '                                                                                               _______________added 1/30/07_____________
             TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + (TaxTrans.Revenue.Principle1Pd + TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
             TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2Pd)
             TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3Pd)
             TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4Pd)
             TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5Pd)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
           Case 22
             ThisTransType = "Overpayment"
             TotByYrAndType(12, ThisYear) = OldRound(TotByYrAndType(12, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(12, ThisYear) = OldRound(CntByYrAndType(12, ThisYear) + 1)
             TotByType(12) = OldRound(TotByType(12) + TaxTrans.Amount)
             TotCntByType(12) = OldRound(TotCntByType(12) + 1)
           Case 24
             ThisTransType = "Adjust Bill Up Affecting Credit Balance"
             TotByYrAndType(13, ThisYear) = OldRound(TotByYrAndType(13, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(13, ThisYear) = OldRound(CntByYrAndType(13, ThisYear) + 1)
             TotByType(13) = OldRound(TotByType(13) + TaxTrans.Amount)
             TotCntByType(13) = OldRound(TotCntByType(13) + 1)
             TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2Pd)
             TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3Pd)
             TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4Pd)
             TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5Pd)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
'             TotByYrAndType(13, ThisYear) = OldRound(TotByYrAndType(13, ThisYear) + TaxTrans.Amount)
'             CntByYrAndType(13, ThisYear) = OldRound(CntByYrAndType(13, ThisYear) + 1)
'             TotByType(13) = OldRound(TotByType(13) + TaxTrans.Amount)
'             TotCntByType(13) = OldRound(TotCntByType(13) + 1)
'             TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1)
'             TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2)
'             TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3)
'             TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4)
'             TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5)
'             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.Interest)
'             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.Collection)
'             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateList)
'             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.Penalty)
'             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1)
'             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2)
'             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3)
           Case 10 ', 24 '7/11/06 added Pd on revenues
             ThisTransType = "Adjust Pay Down Affecting Credit Balance"
             TotByYrAndType(14, ThisYear) = OldRound(TotByYrAndType(14, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(14, ThisYear) = OldRound(CntByYrAndType(14, ThisYear) + 1)
             TotByType(14) = OldRound(TotByType(14) + TaxTrans.Amount)
             TotCntByType(14) = OldRound(TotCntByType(14) + 1)
             TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2Pd)
             TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3Pd)
             TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4Pd)
             TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5Pd)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
           Case 11
             ThisTransType = "Adjust Prepay Down"
             TotByYrAndType(15, ThisYear) = OldRound(TotByYrAndType(15, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(15, ThisYear) = OldRound(CntByYrAndType(15, ThisYear) + 1)
             TotByType(15) = OldRound(TotByType(15) + TaxTrans.Amount)
             TotCntByType(15) = OldRound(TotCntByType(15) + 1)
           Case 12
             ThisTransType = "Refund Prepay"
             TotByYrAndType(16, ThisYear) = OldRound(TotByYrAndType(16, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(16, ThisYear) = OldRound(CntByYrAndType(16, ThisYear) + 1)
             TotByType(16) = OldRound(TotByType(16) + TaxTrans.Amount)
             TotCntByType(16) = OldRound(TotCntByType(16) + 1)
           Case 30
             ThisTransType = "PPTRA Removal"
             TotByYrAndType(17, ThisYear) = OldRound(TotByYrAndType(17, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(17, ThisYear) = OldRound(CntByYrAndType(17, ThisYear) + 1)
             TotByType(17) = OldRound(TotByType(17) + TaxTrans.Amount)
             TotCntByType(17) = OldRound(TotCntByType(17) + 1)
           Case Else
             ThisTransType = "Unknown"
             TotByYrAndType(18, ThisYear) = OldRound(TotByYrAndType(18, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(18, ThisYear) = OldRound(CntByYrAndType(18, ThisYear) + 1)
             TotByType(18) = OldRound(TotByType(18) + TaxTrans.Amount)
             TotCntByType(18) = OldRound(TotCntByType(18) + 1)
          End Select
          TCnt = TCnt + 1
          TotAmt = OldRound(TotAmt + TaxTrans.Amount)
          
          '------------------------------------------------------------------
          '                                                                               __________________1/30/07__________________
          PersDif = OldRound(TaxTrans.Revenue.Principle1 - (TaxTrans.Revenue.Principle1Pd + TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
          MTDif = OldRound(TaxTrans.Revenue.Principle2 - TaxTrans.Revenue.Principle2Pd)
          MCDif = OldRound(TaxTrans.Revenue.Principle3 - TaxTrans.Revenue.Principle3Pd)
          FEDif = OldRound(TaxTrans.Revenue.Principle4 - TaxTrans.Revenue.Principle4Pd)
          MHDif = OldRound(TaxTrans.Revenue.Principle5 - TaxTrans.Revenue.Principle5Pd)
          IntDif = OldRound(TaxTrans.Revenue.Interest - TaxTrans.Revenue.InterestPd)
          PenDif = OldRound(TaxTrans.Revenue.Penalty - TaxTrans.Revenue.PenaltyPd)
          Opt1Dif = OldRound(TaxTrans.Revenue.RevOpt1 - TaxTrans.Revenue.RevOpt1Pd)
          Opt2Dif = OldRound(TaxTrans.Revenue.RevOpt2 - TaxTrans.Revenue.RevOpt2Pd)
          Opt3Dif = OldRound(TaxTrans.Revenue.RevOpt3 - TaxTrans.Revenue.RevOpt3Pd)
      
          '                   0            1                 2                   3
          Print #RptHandle, Town$; dlm; ThisName; dlm; TaxCust.Acct; dlm; TaxCust.Active; dlm;
          '                                 4                           5                6
          Print #RptHandle, MakeRegDate(TaxTrans.TransDate); dlm; ThisBillType; dlm; ThisType; dlm;
          If TaxTrans.BelongTo > 0 Then
            Get TTHandle, TaxTrans.BelongTo, TaxTrans
              '                          7                         8                          9
              Print #RptHandle, MakeRegDate(BegDate); dlm; MakeRegDate(EndDate); dlm; TaxTrans.TaxYear; dlm;
              Get TTHandle, ThisRec, TaxTrans
              If TaxTrans.DiscAmt > 0 Then GoSub ApplyDisc 'added 1/16/07
          Else
            '                          7                         8                          9
            Print #RptHandle, MakeRegDate(BegDate); dlm; MakeRegDate(EndDate); dlm; TaxTrans.TaxYear; dlm;
          End If
          If TaxTrans.TranType <> 9 Then
            '                      10                11          12                       13
            Print #RptHandle, TaxTrans.Amount; dlm; TCnt; dlm; TotAmt; dlm; TaxTrans.Revenue.PrePaidAmt; dlm;
          Else
            '                      10                11          12                       13
            Print #RptHandle, TaxTrans.Revenue.PrePaidUsed; dlm; TCnt; dlm; TotAmt; dlm; TaxTrans.Revenue.PrePaidAmt; dlm;
          End If
          If TaxTrans.BelongTo > 0 Then
            Get TTHandle, TaxTrans.BelongTo, TaxTrans
            '                             14
            Print #RptHandle, ParseBillNum(TaxTrans.Description); dlm;
          Else
            '                 14
            Print #RptHandle, 0; dlm;
          End If
          Get TTHandle, ThisRec, TaxTrans
          If DiscApplied = True Then GoSub ApplyDisc 'added 1/16/07
          '                                15                        16
          Print #RptHandle, QPTrim$(TaxTrans.Description); dlm; ThisTransType; dlm;
          '                               17                              18                           19
          Print #RptHandle, TaxTrans.Revenue.Principle1; dlm; TaxTrans.Revenue.Principle1Pd; ; dlm; PersDif; dlm;
          '                               20                              21                      22
          Print #RptHandle, TaxTrans.Revenue.Interest; dlm; TaxTrans.Revenue.InterestPd; ; dlm; IntDif; dlm;
          '                               23                              24                        25
          Print #RptHandle, TaxTrans.Revenue.Principle2; dlm; TaxTrans.Revenue.Principle2Pd; dlm; MTDif; dlm;
          '                               26                              27                       28
          Print #RptHandle, TaxTrans.Revenue.Principle3; dlm; TaxTrans.Revenue.Principle3Pd; dlm; MCDif; dlm;
          '                               29                              30                   31
          Print #RptHandle, TaxTrans.Revenue.RevOpt1; dlm; TaxTrans.Revenue.RevOpt1Pd; dlm; Opt1Dif; dlm;
          '                               32                              33                   34
          Print #RptHandle, TaxTrans.Revenue.RevOpt2; dlm; TaxTrans.Revenue.RevOpt2Pd; dlm; Opt2Dif; dlm;
          '                               35                              36                   37
          Print #RptHandle, TaxTrans.Revenue.RevOpt3; dlm; TaxTrans.Revenue.RevOpt3Pd; dlm; Opt3Dif; dlm;
          '                    38              39              40
          Print #RptHandle, POpt1Desc; dlm; POpt2Desc; dlm; POpt3Desc; dlm;
          If UseOpt = "Y" Then
            '                    41                     42
            Print #RptHandle, ThisOpt; dlm; QPTrim$(TaxCust.OptSrchDesc); dlm;
          Else
            '                 41       42
            Print #RptHandle, ""; dlm; ""; dlm;
          End If
          '                             43                                 44                       45
          Print #RptHandle, TaxTrans.Revenue.Principle4; dlm; TaxTrans.Revenue.Principle4Pd; dlm; FEDif; dlm;
          '                             46                              47                         48
          Print #RptHandle, TaxTrans.Revenue.Principle5; dlm; TaxTrans.Revenue.Principle5Pd; dlm; MHDif; dlm;
          '                             49                              50                    51               52
          Print #RptHandle, TaxTrans.Revenue.Penalty; dlm; TaxTrans.Revenue.PenaltyPd; dlm; PenDif; dlm; TaxTrans.OperNum
        End If
SkipIt:
        ThisRec = TaxTrans.LastTrans
      Loop
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
  End If
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  
  Close
  If TCnt = 0 Then
    Call TaxMsg(900, "No transactions were found that fit the parameters entered.")
    Close
    Exit Sub
  End If
  'start here on May 1, 2006
  If YrCnt > 0 Then
    GoSub PrintSub
    GoSub PrintTotals
  End If
  
  arVATaxTransJournalDet.Show
  
  Exit Sub
  
PrintSub:
  SubRptFile$ = "TAXRPTS\SUBTXJRLDET.RPT"
  SubRptHandle = FreeFile
  Open SubRptFile For Output As #SubRptHandle
  BigYr = 0
  For x = 1 To YrCnt
    If ThEYear(x) > BigYr Then
      BigYr = ThEYear(x)
    End If
  Next x
  ReDim HoldAmt(1 To 18, 1 To YrCnt) As Double
  ReDim HoldCnt(1 To 18, 1 To YrCnt) As Integer
  Nexty = 1
  Nextx = 1
  HoldBigYr = 0
  For x = 1 To 18
    For y = Nexty To YrCnt
      If ThEYear(y) >= HoldBigYr Then
        HoldBigYr = ThEYear(y)
        Thisx = x
        Thisy = y
      End If
    Next y
    For z = 1 To 18
      HoldAmt(z, Thisy) = TotByYrAndType(z, Nexty)
      HoldCnt(z, Thisy) = CntByYrAndType(z, Nexty)
    Next z
    HoldYr = ThEYear(Nexty)
    If fpcmbTransType.Text = " 0) All" Then GoTo All1
    HoldPers = TotByYrAndPers(Nexty)
    HoldMT = TotByYrAndMT(Nexty)
    HoldMC = TotByYrAndMC(Nexty)
    HoldFE = TotByYrAndFE(Nexty)
    HoldMH = TotByYrAndMH(Nexty)
    HoldInt = TotByYrAndInt(Nexty)
    HoldAdv = TotByYrAndAdv(Nexty)
    HoldPen = TotByYrAndPen(Nexty)
    HoldLateList = TotByYrAndLateList(Nexty)
    HoldOpt1 = TotByYrAndOpt1(Nexty)
    HoldOpt2 = TotByYrAndOpt2(Nexty)
    HoldOpt3 = TotByYrAndOpt3(Nexty)
All1:
    For z = 1 To 18
      TotByYrAndType(z, Nexty) = TotByYrAndType(z, Thisy)
      CntByYrAndType(z, Nexty) = CntByYrAndType(z, Thisy)
    Next z
    ThEYear(Nexty) = ThEYear(Thisy)
    If fpcmbTransType.Text = " 0) All" Then GoTo All2
    TotByYrAndPers(Nexty) = TotByYrAndPers(Thisy)
    TotByYrAndMT(Nexty) = TotByYrAndMT(Thisy)
    TotByYrAndMC(Nexty) = TotByYrAndMC(Thisy)
    TotByYrAndFE(Nexty) = TotByYrAndFE(Thisy)
    TotByYrAndMH(Nexty) = TotByYrAndMH(Thisy)
    TotByYrAndInt(Nexty) = TotByYrAndInt(Thisy)
    TotByYrAndAdv(Nexty) = TotByYrAndAdv(Thisy)
    TotByYrAndPen(Nexty) = TotByYrAndPen(Thisy)
    TotByYrAndLateList(Nexty) = TotByYrAndLateList(Thisy)
    TotByYrAndOpt1(Nexty) = TotByYrAndOpt1(Thisy)
    TotByYrAndOpt2(Nexty) = TotByYrAndOpt2(Thisy)
    TotByYrAndOpt3(Nexty) = TotByYrAndOpt3(Thisy)
All2:
    For z = 1 To 18
      TotByYrAndType(z, Thisy) = HoldAmt(z, Thisy)
      CntByYrAndType(z, Thisy) = HoldCnt(z, Thisy)
    Next z
    ThEYear(Thisy) = HoldYr
    If fpcmbTransType.Text = " 0) All" Then GoTo All3
    TotByYrAndPers(Thisy) = HoldPers
    TotByYrAndMT(Thisy) = HoldMT
    TotByYrAndMC(Thisy) = HoldMC
    TotByYrAndFE(Thisy) = HoldFE
    TotByYrAndMH(Thisy) = HoldMH
    TotByYrAndInt(Thisy) = HoldInt
    TotByYrAndAdv(Thisy) = HoldAdv
    TotByYrAndPen(Thisy) = HoldPen
    TotByYrAndLateList(Thisy) = HoldLateList
    TotByYrAndOpt1(Thisy) = HoldOpt1
    TotByYrAndOpt2(Thisy) = HoldOpt2
    TotByYrAndOpt3(Thisy) = HoldOpt3
All3:
    If Nexty >= YrCnt Then Exit For
    HoldBigYr = 0
    Nexty = Nexty + 1
  Next x
  For y = 1 To YrCnt
    TotYearCnt = 0
    TotYearAmt = 0
    For x = 1 To 18
      If TotByYrAndType(x, y) > 0 Then
        TotYearCnt = TotYearCnt + CntByYrAndType(x, y)
        TotYearAmt = TotYearAmt + TotByYrAndType(x, y)
        Select Case x
          Case 1
            Print #SubRptHandle, "Billing"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
              Print #SubRptHandle, TotByYrAndPers(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10                 11         12              13              14
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; POpt1Desc; dlm; POpt2Desc; dlm; POpt3Desc; dlm;
              '                          15                    16                    17                    18                    19
              Print #SubRptHandle, TotByYrAndMT(y); dlm; TotByYrAndMC(y); dlm; TotByYrAndFE(y); dlm; TotByYrAndMH(y); dlm; TotByYrAndPen(y); dlm; 0; dlm; 0; dlm; "N"
            Else
              '                    4       5       6       7       8       9       10      11       12       13       14      15      16      17      18      19           20               21            22
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 2
            Print #SubRptHandle, "Payment"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
              Print #SubRptHandle, TotByYrAndPers(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10                 11         12             13             14
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; POpt1Desc; dlm; POpt2Desc; dlm; POpt3Desc; dlm;
              '                          15                    16                    17                    18                     19
              Print #SubRptHandle, TotByYrAndMT(y); dlm; TotByYrAndMC(y); dlm; TotByYrAndFE(y); dlm; TotByYrAndMH(y); dlm; TotByYrAndPen(y); dlm; 0; dlm; 0; dlm; "N"
            Else
              '                    4       5       6       7       8       9       10      11       12       13       14      15      16      17      18       19
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 3
            Print #SubRptHandle, "Release"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
              Print #SubRptHandle, TotByYrAndPers(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10                 11         12             13             14
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; POpt1Desc; dlm; POpt2Desc; dlm; POpt3Desc; dlm;
              '                          15                    16                    17                    18                   19
              Print #SubRptHandle, TotByYrAndMT(y); dlm; TotByYrAndMC(y); dlm; TotByYrAndFE(y); dlm; TotByYrAndMH(y); dlm; TotByYrAndPen(y); dlm; 0; dlm; 0; dlm; "N"
            Else
              '                    4       5       6       7       8       9       10      11       12       13       14      15      16      17      18      19
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 4
            Print #SubRptHandle, "Interest"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 5
            Print #SubRptHandle, "Penalty"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 6
            Print #SubRptHandle, "Advertising"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 7
            Print #SubRptHandle, "Adjust Pay Down"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
              Print #SubRptHandle, TotByYrAndPers(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10                 11         12             13             14
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; POpt1Desc; dlm; POpt2Desc; dlm; POpt3Desc; dlm;
              '                          15                    16                    17                    18                    19
              Print #SubRptHandle, TotByYrAndMT(y); dlm; TotByYrAndMC(y); dlm; TotByYrAndFE(y); dlm; TotByYrAndMH(y); dlm; TotByYrAndPen(y); dlm; 0; dlm; 0; dlm; "N"
            Else
              '                    4       5       6       7       8       9       10      11       12       13       14      15      16      17      18       19
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 8
            Print #SubRptHandle, "Credit at Billing"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 9
            Print #SubRptHandle, "Adjust Bill Down"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
              Print #SubRptHandle, TotByYrAndPers(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10                 11         12             13             14
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; POpt1Desc; dlm; POpt2Desc; dlm; POpt3Desc; dlm;
              '                          15                    16                    17                    18                    19
              Print #SubRptHandle, TotByYrAndMT(y); dlm; TotByYrAndMC(y); dlm; TotByYrAndFE(y); dlm; TotByYrAndMH(y); dlm; TotByYrAndPen(y); dlm; 0; dlm; 0; dlm; "N"
            Else
              '                    4       5       6       7       8       9       10      11       12       13       14      15      16      17      18
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 10
            Print #SubRptHandle, "Adjust Bill Up"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
              Print #SubRptHandle, TotByYrAndPers(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10                 11         12             13             14
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; POpt1Desc; dlm; POpt2Desc; dlm; POpt3Desc; dlm;
              '                          15                    16                    17                    18                   19
              Print #SubRptHandle, TotByYrAndMT(y); dlm; TotByYrAndMC(y); dlm; TotByYrAndFE(y); dlm; TotByYrAndMH(y); dlm; TotByYrAndPen(y); dlm; 0; dlm; 0; dlm; "N"
            Else
              '                    4       5       6       7       8       9       10      11       12       13       14      15      16      17      18
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 11
            If fpcmbTransType.Text <> " 0) All" Then 'added Bill Pay/OverPay on 7/7/06
              Print #SubRptHandle, "Bill OverPay"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
              '                            4                      5                      6                         7
              Print #SubRptHandle, TotByYrAndPers(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10                 11         12             13             14
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; POpt1Desc; dlm; POpt2Desc; dlm; POpt3Desc; dlm;
              '                          15                    16                    17                    18                    19
              Print #SubRptHandle, TotByYrAndMT(y); dlm; TotByYrAndMC(y); dlm; TotByYrAndFE(y); dlm; TotByYrAndMH(y); dlm; TotByYrAndPen(y); dlm; 0; dlm; 0; dlm; "N"
            Else
              Print #SubRptHandle, "Bill Pay/OverPay"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
              '                    4       5       6       7       8       9       10      11       12       13       14      15      16      17      18
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 12
            Print #SubRptHandle, "OverPayment"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 13
            Print #SubRptHandle, "Adjust Bill Up Affecting Credit Balance"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
              Print #SubRptHandle, TotByYrAndPers(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10                 11         12             13             14
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; POpt1Desc; dlm; POpt2Desc; dlm; POpt3Desc; dlm;
              '                          15                    16                    17                    18                   19
              Print #SubRptHandle, TotByYrAndMT(y); dlm; TotByYrAndMC(y); dlm; TotByYrAndFE(y); dlm; TotByYrAndMH(y); dlm; TotByYrAndPen(y); dlm; 0; dlm; 0; dlm; "N"
            Else
              '                    4       5       6       7       8       9       10      11       12       13       14      15      16      17      18       19
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 14
            Print #SubRptHandle, "Adjust Pay Dn Affecting Credit Balance"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              '                            4                      5                      6                         7
              Print #SubRptHandle, TotByYrAndPers(y); dlm; TotByYrAndInt(y); dlm; TotByYrAndAdv(y); dlm; TotByYrAndLateList(y); dlm;
              '                            8                       9                    10                 11         12             13             14
              Print #SubRptHandle, TotByYrAndOpt1(y); dlm; TotByYrAndOpt2(y); dlm; TotByYrAndOpt3(y); dlm; x; dlm; POpt1Desc; dlm; POpt2Desc; dlm; POpt3Desc; dlm;
              '                          15                    16                    17                    18                    19
              Print #SubRptHandle, TotByYrAndMT(y); dlm; TotByYrAndMC(y); dlm; TotByYrAndFE(y); dlm; TotByYrAndMH(y); dlm; TotByYrAndPen(y); dlm; 0; dlm; 0; dlm; "N"
            Else
              '                    4       5       6       7       8       9       10      11       12       13       14      15      16      17      18
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 15
            Print #SubRptHandle, "Adjust Prepay Down"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 16
            Print #SubRptHandle, "Refund Prepay"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 17
            Print #SubRptHandle, "PPTRA Removal"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
          Case 18
            Print #SubRptHandle, "Unknown"; dlm; ThEYear(y); dlm; TotByYrAndType(x, y); dlm; CntByYrAndType(x, y); dlm;
            If fpcmbTransType.Text <> " 0) All" Then
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; x; dlm; ""; dlm; ""; dlm; ""; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            Else
              Print #SubRptHandle, 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 44; dlm; ""; dlm; ""; dlm; ""; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; 0; dlm; TotYearCnt; dlm; TotYearAmt; dlm; "Y"
            End If
        End Select
      End If
    Next x
  Next y
  Close SubRptHandle
  
  Return
  
PrintTotals:
  Sub2RptFile$ = "TAXRPTS\SUB2TXJRLDET.RPT"
  Sub2RptHandle = FreeFile
  Open Sub2RptFile For Output As #Sub2RptHandle
  GCntTot = 0
  GAmtTot = 0
  If fpcmbTransType.Text = " 0) All" Then GoTo All4
  For x = 1 To YrCnt
    GPersTot = GPersTot + TotByYrAndPers(x)
    GMTTot = GMTTot + TotByYrAndMT(x)
    GMCTot = GMCTot + TotByYrAndMC(x)
    GFETot = GFETot + TotByYrAndFE(x)
    GMHTot = GMHTot + TotByYrAndMH(x)
    GIntTot = GIntTot + TotByYrAndInt(x)
    GAdvTot = GAdvTot + TotByYrAndAdv(x)
    GLateListTot = GLateListTot + TotByYrAndLateList(x)
    GPenTot = GPenTot + TotByYrAndPen(x)
    GOpt1Tot = GOpt1Tot + TotByYrAndOpt1(x)
    GOpt2Tot = GOpt2Tot + TotByYrAndOpt2(x)
    GOpt3Tot = GOpt3Tot + TotByYrAndOpt3(x)
  Next x
All4:
  Dim Case11Tot As Double 'added 7/6/06
  For x = 1 To 18
    GCntTot = GCntTot + TotCntByType(x)
    GAmtTot = GAmtTot + TotByType(x)
    If TotByType(x) > 0 Then
      If x <> 11 Then
        '                          0                   1                  2             3
        Print #Sub2RptHandle, TotByType(x); dlm; TotCntByType(x); dlm; GAmtTot; dlm; GCntTot; dlm;
      Else 'added 7/6/06
        Case11Tot = OldRound(TotByType(x) - (GPersTot + GIntTot + GAdvTot + GLateListTot + GPenTot + GOpt1Tot))
        Case11Tot = OldRound(Case11Tot - (GOpt2Tot + GOpt3Tot + GMTTot + GMCTot + GFETot + GMHTot))
        '                         0                   1                  2             3
        Print #Sub2RptHandle, Case11Tot; dlm; TotCntByType(x); dlm; GAmtTot; dlm; GCntTot; dlm;
      End If
      '                         4              5             6               7
      Print #Sub2RptHandle, GPersTot; dlm; GIntTot; dlm; GAdvTot; dlm; GLateListTot; dlm;
      '                        8              9             10              11              12               13
      Print #Sub2RptHandle, GOpt1Tot; dlm; GOpt2Tot; dlm; GOpt3Tot; dlm; POpt1Desc; dlm; POpt2Desc; dlm; POpt3Desc; dlm;
      '                       14           15           16           17           18
      Print #Sub2RptHandle, GMTTot; dlm; GMCTot; dlm; GFETot; dlm; GMHTot; dlm; GPenTot; dlm;
      If fpcmbTransType.Text <> " 0) All" Then
        '                      19
        Print #Sub2RptHandle, "1"; dlm;
      Else
        '                      19
        Print #Sub2RptHandle, "2"; dlm;
      End If
      Select Case x
        Case 1
          '                         20          21
          Print #Sub2RptHandle, "Billing"; dlm; 3
        Case 2
          Print #Sub2RptHandle, "Payment"; dlm; 3
        Case 3
          Print #Sub2RptHandle, "Release"; dlm; 3
        Case 4
          Print #Sub2RptHandle, "Interest"; dlm; 44
        Case 5
          Print #Sub2RptHandle, "Penalty"; dlm; 44
        Case 6
          Print #Sub2RptHandle, "Advertising"; dlm; 44
        Case 7
          Print #Sub2RptHandle, "Adjust Pay Down"; dlm; 3
        Case 8
          Print #Sub2RptHandle, "Credit at Billing"; dlm; 44
        Case 9
          Print #Sub2RptHandle, "Adjust Bill Down"; dlm; 3
        Case 10
          Print #Sub2RptHandle, "Adjust Bill Up"; dlm; 44
        Case 11
          If fpcmbTransType.Text = " 0) All" Then 'added All part on 7/7/06
            Print #Sub2RptHandle, "Bill Pay/OverPay"; dlm; 3 'changed from 44 on 7/6/06
          Else
            Print #Sub2RptHandle, "Bill OverPay"; dlm; 3 'changed from 44 on 7/6/06
          End If
        Case 12
          Print #Sub2RptHandle, "OverPayment"; dlm; 44
        Case 13
          Print #Sub2RptHandle, "Adjust Bill Up Affecting Credit Balance"; dlm; 3
        Case 14
          Print #Sub2RptHandle, "Adjust Pay Dwn Affecting Credit Balance"; dlm; 3
        Case 15
          Print #Sub2RptHandle, "Adjust Prepay Down"; dlm; 44
        Case 16
          Print #Sub2RptHandle, "Refund Prepay"; dlm; 44
        Case 17
          Print #Sub2RptHandle, "PPTRA Removal"; dlm; 44
        Case 18
          Print #Sub2RptHandle, "Unknown"; dlm; 44
      End Select
    End If
  Next x
  Close Sub2RptHandle
  
  Return

PrintAll:
  For x = 1 To NumOfTCRecs
    PrePayDone = False
    If IdxFlag = False Then
      Get TCHandle, x, TaxCust
      CustBal = GetCustBalance(x, -1)
    Else
      Get TCHandle, IdxArray(x), TaxCust
      CustBal = GetCustBalance(IdxArray(x), -1)
    End If
    If TaxCust.Active = "N" And InactiveFlag = False Then
      GoTo SkipThisOne
    End If
    ThisName = QPTrim$(TaxCust.CustName)
    ReDim BillRec(1 To 1) As Long
    GoSub PrintPrepay
    ThisRec = TaxCust.LastTrans
    BillCnt = 0
    BadCnt = 0
    Do While ThisRec > 0
      Get TTHandle, ThisRec, TaxTrans
      If QFlag = True Then
        If TaxTrans.TransDate < BegDate Then
          BadCnt = BadCnt + 1
          If BadCnt > 3 Then Exit Do
        End If
      End If
      If TransDesc <> "" Then
        If InStr(1, TaxTrans.Description, TransDesc) = 0 Then
          GoTo Nope
        End If
      End If
      If TaxTrans.BillType <> "R" And TaxTrans.BillType <> "P" Then TaxTrans.BillType = ""
      If TaxTrans.BillType <> Mid(fpcmbTaxType.Text, 1, 1) And QPTrim$(TaxTrans.BillType) <> "" Then GoTo Nope
      If TaxTrans.TransDate < BegDate Or TaxTrans.TransDate > EndDate Then GoTo Nope
      If TaxTrans.TranType = 1 Then
        BillCnt = BillCnt + 1
        ReDim Preserve BillRec(1 To BillCnt) As Long
        BillRec(BillCnt) = ThisRec
      ElseIf TaxTrans.TranType = 11 Or TaxTrans.TranType = 12 Then
        GoTo Nope
      End If
Nope:
      ThisRec = TaxTrans.LastTrans
    Loop
  
    For z = 1 To BillCnt
      Get TTHandle, BillRec(z), TaxTrans
      DiscApplied = False 'added 1/16/07
      If TaxTrans.DiscAmt > 0 Then
        If TaxTrans.TranType = 1 Or TaxTrans.TranType = 2 Or TaxTrans.TranType = 21 Then 'added 1/16/07
          GoSub ApplyDisc
        End If
      End If
      If YrCnt = 0 Then
         YrCnt = YrCnt + 1
         ThisYear = YrCnt
         ReDim Preserve ThEYear(1 To YrCnt) As Integer
         ThEYear(YrCnt) = TaxTrans.TaxYear
         ReDim Preserve TotByYrAndType(1 To 18, 1 To YrCnt) As Double
         ReDim Preserve CntByYrAndType(1 To 18, 1 To YrCnt) As Integer
         For y = 1 To 18
           TotByYrAndType(y, YrCnt) = 0
           CntByYrAndType(y, YrCnt) = 0
         Next y
       Else
         For y = 1 To YrCnt
           If TaxTrans.TaxYear = ThEYear(y) Then
             ThisYear = y
             Exit For
           End If
         Next y
         If y > YrCnt Then
           YrCnt = YrCnt + 1
           ThisYear = YrCnt
           ReDim Preserve ThEYear(1 To YrCnt) As Integer
           ThEYear(YrCnt) = TaxTrans.TaxYear
           ReDim Preserve TotByYrAndType(1 To 18, 1 To YrCnt) As Double
           ReDim Preserve CntByYrAndType(1 To 18, 1 To YrCnt) As Integer
           For y = 1 To 18
             TotByYrAndType(y, YrCnt) = 0
             CntByYrAndType(y, YrCnt) = 0
           Next y
        End If
      End If
      ThisTransType = "Billing"
      TCnt = TCnt + 1
      TotAmt = OldRound(TotAmt + TaxTrans.Amount)
      
      '-----------------------------------------------------------------------
      '                                                                               ______________1/30/07______________________
      PersDif = OldRound(TaxTrans.Revenue.Principle1 - (TaxTrans.Revenue.Principle1Pd + TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl)) 'added 1/30/07
      MTDif = OldRound(TaxTrans.Revenue.Principle2 - TaxTrans.Revenue.Principle2Pd)
      MCDif = OldRound(TaxTrans.Revenue.Principle3 - TaxTrans.Revenue.Principle3Pd)
      FEDif = OldRound(TaxTrans.Revenue.Principle4 - TaxTrans.Revenue.Principle4Pd)
      MHDif = OldRound(TaxTrans.Revenue.Principle5 - TaxTrans.Revenue.Principle5Pd)
      IntDif = OldRound(TaxTrans.Revenue.Interest - TaxTrans.Revenue.InterestPd)
      PenDif = OldRound(TaxTrans.Revenue.Penalty - TaxTrans.Revenue.PenaltyPd)
      Opt1Dif = OldRound(TaxTrans.Revenue.RevOpt1 - TaxTrans.Revenue.RevOpt1Pd)
      Opt2Dif = OldRound(TaxTrans.Revenue.RevOpt2 - TaxTrans.Revenue.RevOpt2Pd)
      Opt3Dif = OldRound(TaxTrans.Revenue.RevOpt3 - TaxTrans.Revenue.RevOpt3Pd)
      BillBal = OldRound(PenDif + PersDif + IntDif + MTDif + MCDif + FEDif + MHDif + Opt1Dif + Opt2Dif + Opt3Dif)
'      If TaxCust.Acct = 1651 Then Stop
      '                   0            1                 2                   3
      Print #RptHandle, Town$; dlm; ThisName; dlm; TaxCust.Acct; dlm; TaxCust.Active; dlm;
      '                                 4                           5                6
      Print #RptHandle, MakeRegDate(TaxTrans.TransDate); dlm; ThisBillType; dlm; ThisType; dlm;
      '                          7                         8                          9
      Print #RptHandle, MakeRegDate(BegDate); dlm; MakeRegDate(EndDate); dlm; TaxTrans.TaxYear; dlm;
      '                      10                11          12                       13
      Print #RptHandle, TaxTrans.Amount; dlm; TCnt; dlm; TotAmt; dlm; TaxTrans.Revenue.PrePaidAmt; dlm;
      '                14
      Print #RptHandle, 0; dlm;
      '                                15                        16
      Print #RptHandle, QPTrim$(TaxTrans.Description); dlm; ThisTransType; dlm;
      '                               17                              18                        19
      Print #RptHandle, TaxTrans.Revenue.Principle1; dlm; TaxTrans.Revenue.Principle1Pd; dlm; PersDif; dlm;
      '                               20                              21                    22
      Print #RptHandle, TaxTrans.Revenue.Interest; dlm; TaxTrans.Revenue.InterestPd; dlm; IntDif; dlm;
      '                               23                              24                        25
      Print #RptHandle, TaxTrans.Revenue.Principle2; dlm; TaxTrans.Revenue.Principle2Pd; dlm; MTDif; dlm;
      '                               26                              27                       28
      Print #RptHandle, TaxTrans.Revenue.Principle3; dlm; TaxTrans.Revenue.Principle3Pd; dlm; MCDif; dlm;
      '                               29                              30                   31
      Print #RptHandle, TaxTrans.Revenue.RevOpt1; dlm; TaxTrans.Revenue.RevOpt1Pd; dlm; Opt1Dif; dlm;
      '                               32                              33                   34
      Print #RptHandle, TaxTrans.Revenue.RevOpt2; dlm; TaxTrans.Revenue.RevOpt2Pd; dlm; Opt2Dif; dlm;
      '                               35                              36                   37
      Print #RptHandle, TaxTrans.Revenue.RevOpt3; dlm; TaxTrans.Revenue.RevOpt3Pd; dlm; Opt3Dif; dlm;
      '                    38             39             40                41              42            43                 44
      Print #RptHandle, POpt1Desc; dlm; POpt2Desc; dlm; POpt3Desc; dlm; BillRec(z); dlm; CustBal; dlm; BillBal; dlm; TaxTrans.TranType; dlm;
      '                              45                                46                       47
      Print #RptHandle, TaxTrans.Revenue.Principle4; dlm; TaxTrans.Revenue.Principle4Pd; dlm; FEDif; dlm;
      '                               48                              49                       50
      Print #RptHandle, TaxTrans.Revenue.Principle5; dlm; TaxTrans.Revenue.Principle5Pd; dlm; MHDif; dlm;
      '                              51                                52                 53              54
      Print #RptHandle, TaxTrans.Revenue.Penalty; dlm; TaxTrans.Revenue.PenaltyPd; dlm; PenDif; dlm; TaxTrans.OperNum
      
      ThisRec = TaxCust.LastTrans
      TotByYrAndType(1, ThisYear) = OldRound(TotByYrAndType(1, ThisYear) + TaxTrans.Amount)
      CntByYrAndType(1, ThisYear) = OldRound(CntByYrAndType(1, ThisYear) + 1)
      TotByType(1) = OldRound(TotByType(1) + TaxTrans.Amount)
      TotCntByType(1) = OldRound(TotCntByType(1) + 1)

      Do While ThisRec > 0
        Get TTHandle, ThisRec, TaxTrans
        If TaxTrans.TranType = 22 Then 'Prepay transactions can only be printed one time, not
        'for every iteration
          If PrePayDone = False Then
            GoTo PrepaySkip
          Else
            GoTo Nope2
          End If
        End If
        If TaxTrans.TranType = 11 Or TaxTrans.TranType = 12 Then GoTo Nope2
        If TaxTrans.BelongTo <> BillRec(z) Then GoTo Nope2
PrepaySkip:
        Select Case TaxTrans.TranType
          Case 2
            ThisTransType = "Payment"
            If TaxTrans.DiscAmt > 0 Then '1/16/07
              GoSub ApplyDisc
            End If
            TotByYrAndType(2, ThisYear) = OldRound(TotByYrAndType(2, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(2, ThisYear) = OldRound(CntByYrAndType(2, ThisYear) + 1)
            TotByType(2) = OldRound(TotByType(2) + TaxTrans.Amount)
            TotCntByType(2) = OldRound(TotCntByType(2) + 1)
          Case 3
            ThisTransType = "Release"
            TotByYrAndType(3, ThisYear) = OldRound(TotByYrAndType(3, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(3, ThisYear) = OldRound(CntByYrAndType(3, ThisYear) + 1)
            TotByType(3) = OldRound(TotByType(3) + TaxTrans.Amount)
            TotCntByType(3) = OldRound(TotCntByType(3) + 1)
          Case 4
            ThisTransType = "Interest"
            TotByYrAndType(4, ThisYear) = OldRound(TotByYrAndType(4, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(4, ThisYear) = OldRound(CntByYrAndType(4, ThisYear) + 1)
            TotByType(4) = OldRound(TotByType(4) + TaxTrans.Amount)
            TotCntByType(4) = OldRound(TotCntByType(4) + 1)
          Case 5
            ThisTransType = "Penalty"
            TotByYrAndType(5, ThisYear) = OldRound(TotByYrAndType(5, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(5, ThisYear) = OldRound(CntByYrAndType(5, ThisYear) + 1)
            TotByType(5) = OldRound(TotByType(5) + TaxTrans.Amount)
            TotCntByType(5) = OldRound(TotCntByType(5) + 1)
          Case 6
            ThisTransType = "Advertising Charge"
            TotByYrAndType(6, ThisYear) = OldRound(TotByYrAndType(6, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(6, ThisYear) = OldRound(CntByYrAndType(6, ThisYear) + 1)
            TotByType(6) = OldRound(TotByType(6) + TaxTrans.Amount)
            TotCntByType(6) = OldRound(TotCntByType(6) + 1)
          Case 7
            ThisTransType = "Adjust Pay Down"
            TotByYrAndType(7, ThisYear) = OldRound(TotByYrAndType(7, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(7, ThisYear) = OldRound(CntByYrAndType(7, ThisYear) + 1)
            TotByType(7) = OldRound(TotByType(7) + TaxTrans.Amount)
            TotCntByType(7) = OldRound(TotCntByType(7) + 1)
          Case 9
            ThisTransType = "Credit Applied at Billing"
            TotByYrAndType(8, ThisYear) = OldRound(TotByYrAndType(8, ThisYear) + TaxTrans.Revenue.PrePaidUsed)
            CntByYrAndType(8, ThisYear) = OldRound(CntByYrAndType(8, ThisYear) + 1)
            TotByType(8) = OldRound(TotByType(8) + TaxTrans.Revenue.PrePaidUsed)
            TotCntByType(8) = OldRound(TotCntByType(8) + 1)
          Case 13
            ThisTransType = "Adjust Bill Down"
            TotByYrAndType(9, ThisYear) = OldRound(TotByYrAndType(9, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(9, ThisYear) = OldRound(CntByYrAndType(9, ThisYear) + 1)
            TotByType(9) = OldRound(TotByType(9) + TaxTrans.Amount)
            TotCntByType(9) = OldRound(TotCntByType(9) + 1)
          Case 14
            ThisTransType = "Adjust Bill Up"
            TotByYrAndType(10, ThisYear) = OldRound(TotByYrAndType(10, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(10, ThisYear) = OldRound(CntByYrAndType(10, ThisYear) + 1)
            TotByType(10) = OldRound(TotByType(10) + TaxTrans.Amount)
            TotCntByType(10) = OldRound(TotCntByType(10) + 1)
          Case 21
            ThisTransType = "Billpay/Overpay"
            If TaxTrans.DiscAmt > 0 Then
              GoSub ApplyDisc
            End If
            TotByYrAndType(11, ThisYear) = OldRound(TotByYrAndType(11, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(11, ThisYear) = OldRound(CntByYrAndType(11, ThisYear) + 1)
            TotByType(11) = OldRound(TotByType(11) + TaxTrans.Amount)
            TotCntByType(11) = OldRound(TotCntByType(11) + 1)
          Case 22
            ThisTransType = "Overpayment"
            TotByYrAndType(12, ThisYear) = OldRound(TotByYrAndType(12, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(12, ThisYear) = OldRound(CntByYrAndType(12, ThisYear) + 1)
            TotByType(12) = OldRound(TotByType(12) + TaxTrans.Amount)
            TotCntByType(12) = OldRound(TotCntByType(12) + 1)
          Case 24
            ThisTransType = "Adjust Bill Up Affecting Credit Balance"
            TotByYrAndType(13, ThisYear) = OldRound(TotByYrAndType(13, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(13, ThisYear) = OldRound(CntByYrAndType(13, ThisYear) + 1)
            TotByType(13) = OldRound(TotByType(13) + TaxTrans.Amount)
            TotCntByType(13) = OldRound(TotCntByType(13) + 1)
          Case 10
            ThisTransType = "Adjust Pay Dwn Affecting Credit Balance"
            TotByYrAndType(14, ThisYear) = OldRound(TotByYrAndType(14, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(14, ThisYear) = OldRound(CntByYrAndType(14, ThisYear) + 1)
            TotByType(14) = OldRound(TotByType(14) + TaxTrans.Amount)
            TotCntByType(14) = OldRound(TotCntByType(14) + 1)
          Case 30
            ThisTransType = "PPTRA Removal"
            TotByYrAndType(17, ThisYear) = OldRound(TotByYrAndType(17, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(17, ThisYear) = OldRound(CntByYrAndType(17, ThisYear) + 1)
            TotByType(17) = OldRound(TotByType(17) + TaxTrans.Amount)
            TotCntByType(17) = OldRound(TotCntByType(17) + 1)
          Case Else
            ThisTransType = "Unknown"
            TotByYrAndType(18, ThisYear) = OldRound(TotByYrAndType(18, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(18, ThisYear) = OldRound(CntByYrAndType(18, ThisYear) + 1)
            TotByType(18) = OldRound(TotByType(18) + TaxTrans.Amount)
            TotCntByType(18) = OldRound(TotCntByType(18) + 1)
        End Select
        TCnt = TCnt + 1

        '                   0            1                 2                   3
        Print #RptHandle, Town$; dlm; ThisName; dlm; TaxCust.Acct; dlm; TaxCust.Active; dlm;
        '                                 4                           5                6
        Print #RptHandle, MakeRegDate(TaxTrans.TransDate); dlm; ThisBillType; dlm; ThisType; dlm;
        '                          7                         8                          9
        Print #RptHandle, MakeRegDate(BegDate); dlm; MakeRegDate(EndDate); dlm; TaxTrans.TaxYear; dlm;
        If TaxTrans.TranType <> 9 Then
          '                      10                11          12                       13
          Print #RptHandle, TaxTrans.Amount; dlm; TCnt; dlm; TotAmt; dlm; TaxTrans.Revenue.PrePaidAmt; dlm;
        Else
          '                      10                11          12                       13
          Print #RptHandle, TaxTrans.Revenue.PrePaidUsed; dlm; TCnt; dlm; TotAmt; dlm; TaxTrans.Revenue.PrePaidAmt; dlm;
        End If
        '                      14
        If TaxTrans.BelongTo > 0 Then
          Get TTHandle, TaxTrans.BelongTo, TaxTrans
          '                             14
          Print #RptHandle, ParseBillNum(TaxTrans.Description); dlm;
        Else
          '                 14
          Print #RptHandle, 0; dlm;
        End If
        Get TTHandle, ThisRec, TaxTrans
        If DiscApplied = True Then GoSub ApplyDisc 'added 1/16/07
        '                                15                        16
        Print #RptHandle, QPTrim$(TaxTrans.Description); dlm; ThisTransType; dlm;
        '                               17                              18                           19
        Print #RptHandle, TaxTrans.Revenue.Principle1; dlm; TaxTrans.Revenue.Principle1Pd; ; dlm; 0; dlm;
        '                               20                              21                      22
        Print #RptHandle, TaxTrans.Revenue.Interest; dlm; TaxTrans.Revenue.InterestPd; ; dlm; 0; dlm;
        '                               23                              24                        25
        Print #RptHandle, TaxTrans.Revenue.Principle2; dlm; TaxTrans.Revenue.Principle2Pd; dlm; 0; dlm;
        '                               26                              27                       28
        Print #RptHandle, TaxTrans.Revenue.Principle3; dlm; TaxTrans.Revenue.Principle3Pd; dlm; 0; dlm;
        '                               29                              30                31
        Print #RptHandle, TaxTrans.Revenue.RevOpt1; dlm; TaxTrans.Revenue.RevOpt1Pd; dlm; 0; dlm;
        '                               32                              33                34
        Print #RptHandle, TaxTrans.Revenue.RevOpt2; dlm; TaxTrans.Revenue.RevOpt2Pd; dlm; 0; dlm;
        '                               35                              36                37
        Print #RptHandle, TaxTrans.Revenue.RevOpt3; dlm; TaxTrans.Revenue.RevOpt3Pd; dlm; 0; dlm;
        '                    38             39               40             41               42            43                44
        Print #RptHandle, POpt1Desc; dlm; POpt2Desc; dlm; POpt3Desc; dlm; BillRec(z); dlm; CustBal; dlm; BillBal; dlm; TaxTrans.TranType; dlm;
        '                               45                              46                      47
        Print #RptHandle, TaxTrans.Revenue.Principle4; dlm; TaxTrans.Revenue.Principle4Pd; dlm; 0; dlm;
        '                               48                              49                      50
        Print #RptHandle, TaxTrans.Revenue.Principle5; dlm; TaxTrans.Revenue.Principle5Pd; dlm; 0; dlm;
        '                              51                             52                 53              54
        Print #RptHandle, TaxTrans.Revenue.Penalty; dlm; TaxTrans.Revenue.PenaltyPd; dlm; 0; dlm; TaxTrans.OperNum
    
Nope2:
        ThisRec = TaxTrans.LastTrans
      Loop
      PrePayDone = True 'added 7/10/06
    Next z
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
SkipThisOne:
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  Close
  
  If TCnt = 0 Then
    Call TaxMsg(900, "No transactions were found that fit the parameters entered.")
    Close
    Exit Sub
  End If
  
  If YrCnt > 0 Then
    GoSub PrintSub
    GoSub PrintTotals
  End If
  
  arVATaxPJrnlAllDetail.Show
  
  Exit Sub

PrintPrepay:
  ThisRec = TaxCust.LastTrans
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
    If TaxTrans.TranType <> 11 And TaxTrans.TranType <> 12 Then GoTo NotThisTrans
    If TaxTrans.TransDate < BegDate Or TaxTrans.TransDate > EndDate Then GoTo NotThisTrans
    If TaxTrans.TranType = 11 Then
      ThisTransType = "Adj Prepay Down"
    ElseIf TaxTrans.TranType = 12 Then
      ThisTransType = "Refund Prepay"
    End If
    If YrCnt = 0 Then
      YrCnt = YrCnt + 1
      ThisYear = YrCnt
      ReDim Preserve ThEYear(1 To YrCnt) As Integer
      ThEYear(YrCnt) = TaxTrans.TaxYear
      ReDim Preserve TotByYrAndType(1 To 18, 1 To YrCnt) As Double
      ReDim Preserve CntByYrAndType(1 To 18, 1 To YrCnt) As Integer
      For y = 1 To 18
        TotByYrAndType(y, YrCnt) = 0
        CntByYrAndType(y, YrCnt) = 0
      Next y
    Else
      For y = 1 To YrCnt
        If TaxTrans.TaxYear = ThEYear(y) Then
          ThisYear = y
          Exit For
          End If
      Next y
      If y > YrCnt Then
        YrCnt = YrCnt + 1
        ThisYear = YrCnt
        ReDim Preserve ThEYear(1 To YrCnt) As Integer
        ThEYear(YrCnt) = TaxTrans.TaxYear
        ReDim Preserve TotByYrAndType(1 To 18, 1 To YrCnt) As Double
        ReDim Preserve CntByYrAndType(1 To 18, 1 To YrCnt) As Integer
        For y = 1 To 18
          TotByYrAndType(y, YrCnt) = 0
          CntByYrAndType(y, YrCnt) = 0
        Next y
      End If
    End If
    If TaxTrans.TranType = 11 Then
      TotByYrAndType(15, ThisYear) = OldRound(TotByYrAndType(15, ThisYear) + TaxTrans.Amount)
      CntByYrAndType(15, ThisYear) = OldRound(CntByYrAndType(15, ThisYear) + 1)
    ElseIf TaxTrans.TranType = 12 Then
      TotByYrAndType(16, ThisYear) = OldRound(TotByYrAndType(16, ThisYear) + TaxTrans.Amount)
      CntByYrAndType(16, ThisYear) = OldRound(CntByYrAndType(16, ThisYear) + 1)
    End If
    TCnt = TCnt + 1
    TotAmt = OldRound(TotAmt + TaxTrans.Amount)
    PersDif = 0
    IntDif = 0
    MTDif = 0
    MCDif = 0
    PenDif = 0
    Opt1Dif = 0
    Opt2Dif = 0
    Opt3Dif = 0
    BillBal = OldRound(PenDif + PersDif + IntDif + MTDif + MCDif + Opt1Dif + Opt2Dif + Opt3Dif)
    TotAmt = OldRound(TotAmt + TaxTrans.Amount)
    PersDif = 0
    IntDif = 0
    MTDif = 0
    MCDif = 0
    PenDif = 0
    Opt1Dif = 0
    Opt2Dif = 0
    Opt3Dif = 0
    BillBal = 0
    '                   0            1                 2                   3
    Print #RptHandle, Town$; dlm; ThisName; dlm; TaxCust.Acct; dlm; TaxCust.Active; dlm;
    '                                 4                           5                6
    Print #RptHandle, MakeRegDate(TaxTrans.TransDate); dlm; ThisBillType; dlm; ThisType; dlm;
    '                          7                         8                          9
    Print #RptHandle, MakeRegDate(BegDate); dlm; MakeRegDate(EndDate); dlm; TaxTrans.TaxYear; dlm;
    '                      10                11          12                       13
    Print #RptHandle, TaxTrans.Amount; dlm; TCnt; dlm; TotAmt; dlm; TaxTrans.Revenue.PrePaidAmt; dlm;
    '                14
    Print #RptHandle, 0; dlm;
    '                                15                        16
    Print #RptHandle, QPTrim$(TaxTrans.Description); dlm; ThisTransType; dlm;
    '                               17                              18                           19
    Print #RptHandle, TaxTrans.Revenue.Principle1; dlm; TaxTrans.Revenue.Principle1Pd; ; dlm; PersDif; dlm;
    '                               20                              21                      22
    Print #RptHandle, TaxTrans.Revenue.Interest; dlm; TaxTrans.Revenue.InterestPd; ; dlm; IntDif; dlm;
    '                               23                              24                        25
    Print #RptHandle, TaxTrans.Revenue.Collection; dlm; TaxTrans.Revenue.CollectionPd; dlm; MTDif; dlm;
    '                               26                              27                       28
    Print #RptHandle, TaxTrans.Revenue.LateList; dlm; TaxTrans.Revenue.LateListPd; dlm; MCDif; dlm;
    '                               29                              30                   31
    Print #RptHandle, TaxTrans.Revenue.RevOpt1; dlm; TaxTrans.Revenue.RevOpt1Pd; dlm; Opt1Dif; dlm;
    '                               32                              33                   34
    Print #RptHandle, TaxTrans.Revenue.RevOpt2; dlm; TaxTrans.Revenue.RevOpt2Pd; dlm; Opt2Dif; dlm;
    '                               35                              36                   37
    Print #RptHandle, TaxTrans.Revenue.RevOpt3; dlm; TaxTrans.Revenue.RevOpt3Pd; dlm; Opt3Dif; dlm;
    '                    38             39             40          41        42            43                 44
    Print #RptHandle, Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm; 0; dlm; CustBal; dlm; BillBal; dlm; TaxTrans.TranType
NotThisTrans:
    ThisRec = TaxTrans.LastTrans
  Loop
  
  Return

ApplyDisc:
  Disc1 = 0
  Disc2 = 0
  Disc3 = 0
  Disc4 = 0
  Disc5 = 0
  Disc6 = 0
  Disc7 = 0
  Disc8 = 0
  If TaxTrans.Amount = 0 Then Return
  If TaxTrans.TranType = 1 Then
    SaveAmt = OldRound(TaxTrans.Amount - TaxTrans.DiscAmt)
  Else
    SaveAmt = TaxTrans.Amount
    TaxTrans.Amount = OldRound(TaxTrans.Amount + TaxTrans.DiscAmt)
  End If
  Disc1 = OldRound(TaxTrans.Revenue.Principle1Pd / SaveAmt)
  Disc1 = OldRound(Disc1 * TaxTrans.DiscAmt)
  Disc2 = OldRound(TaxTrans.Revenue.Principle2Pd / SaveAmt)
  Disc2 = OldRound(Disc2 * TaxTrans.DiscAmt)
  Disc3 = OldRound(TaxTrans.Revenue.Principle3Pd / SaveAmt)
  Disc3 = OldRound(Disc3 * TaxTrans.DiscAmt)
  Disc4 = OldRound(TaxTrans.Revenue.Principle4Pd / SaveAmt)
  Disc4 = OldRound(Disc4 * TaxTrans.DiscAmt)
  Disc5 = OldRound(TaxTrans.Revenue.Principle5Pd / SaveAmt)
  Disc5 = OldRound(Disc5 * TaxTrans.DiscAmt)
  Disc6 = OldRound(TaxTrans.Revenue.RevOpt1Pd / SaveAmt)
  Disc6 = OldRound(Disc6 * TaxTrans.DiscAmt)
  Disc7 = OldRound(TaxTrans.Revenue.RevOpt2Pd / SaveAmt)
  Disc7 = OldRound(Disc7 * TaxTrans.DiscAmt)
  Disc8 = OldRound(TaxTrans.Revenue.RevOpt3Pd / SaveAmt)
  Disc8 = OldRound(Disc8 * TaxTrans.DiscAmt)
  TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd + Disc1)
  TaxTrans.Revenue.Principle2Pd = OldRound(TaxTrans.Revenue.Principle2Pd + Disc2)
  TaxTrans.Revenue.Principle3Pd = OldRound(TaxTrans.Revenue.Principle3Pd + Disc3)
  TaxTrans.Revenue.Principle4Pd = OldRound(TaxTrans.Revenue.Principle4Pd + Disc4)
  TaxTrans.Revenue.Principle5Pd = OldRound(TaxTrans.Revenue.Principle5Pd + Disc5)
  TaxTrans.Revenue.RevOpt1Pd = OldRound(TaxTrans.Revenue.RevOpt1Pd + Disc6)
  TaxTrans.Revenue.RevOpt2Pd = OldRound(TaxTrans.Revenue.RevOpt2Pd + Disc7)
  TaxTrans.Revenue.RevOpt3Pd = OldRound(TaxTrans.Revenue.RevOpt3Pd + Disc8)
  DiscApplied = True
  
  Return


ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxTransJournal", "PrintPGraphicsDet", Erl)
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

Private Sub PrintPTextDet()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim x As Long, y As Integer
  Dim BegDate As Integer
  Dim EndDate As Integer
  Dim ThisRec As Long
  Dim ThisClass As Integer
  Dim ThisType As String
  Dim RptFile$
  Dim RptHandle As Integer
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim InactiveFlag As Boolean
  Dim ThisName$, ThisBillType$
  Dim TCnt As Long, NewName$
  Dim TotAmt As Double
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim NumOfSrchRecs As Long
  Dim ThisTransType As String
  Dim YrCnt As Integer, ThisYear As Integer
  Dim BigYr As Integer
  Dim HoldBigYr As Integer
  Dim HoldYr As Integer
  Dim HoldPers As Double
  Dim HoldMT As Double
  Dim HoldMC As Double
  Dim HoldFE As Double
  Dim HoldMH As Double
  Dim HoldInt As Double
  Dim HoldAdv As Double
  Dim HoldLateList As Double
  Dim HoldPen As Double
  Dim HoldOpt1 As Double
  Dim HoldOpt2 As Double
  Dim HoldOpt3 As Double
  Dim Nextx As Integer
  Dim Thisx As Integer
  Dim Nexty As Integer
  Dim Thisy As Integer
  Dim z As Integer
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim FF$, Page As Integer
  Dim CustName$, PrintCnt As Integer
  Dim ThisBillNum As String * 8
  Dim Opt1Desc$
  Dim Opt2Desc$
  Dim Opt3Desc$
  Dim PersDif As Double
  Dim IntDif As Double
  Dim MTDif As Double
  Dim MCDif As Double
  Dim FEDif As Double
  Dim MHDif As Double
  Dim LateListDif As Double
  Dim PenDif As Double
  Dim Opt1Dif As Double
  Dim Opt2Dif As Double
  Dim Opt3Dif As Double
  Dim ThisBal As Double
  Dim ThisCustRec As Long
  Dim CustBal As Double
  Dim BillBal As Double
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim OptFlag As Boolean
  Dim ThisYrCnt As Long
  Dim ThatName$
  Dim TransDesc$
  Dim GAmtTot As Double
  Dim GCntTot As Long
  Dim GPersTot As Double
  Dim GMTTot As Double
  Dim GMCTot As Double
  Dim GFETot As Double
  Dim GMHTot As Double
  Dim GPenTot As Double
  Dim GIntTot As Double
  Dim GAdvTot As Double
  Dim GLateListTot As Double
  Dim GOpt1Tot As Double
  Dim GOpt2Tot As Double
  Dim GOpt3Tot As Double
  Dim BadCnt As Integer
  Dim QFlag As Boolean
  Dim PrePayDone As Boolean 'added 7/10/06
  Dim Disc1 As Double '1/16/2007
  Dim Disc2 As Double '1/16/2007
  Dim Disc3 As Double '1/16/2007
  Dim Disc4 As Double '1/16/2007
  Dim Disc5 As Double '1/16/2007
  Dim Disc6 As Double '1/16/2007
  Dim Disc7 As Double '1/16/2007
  Dim Disc8 As Double '1/16/2007
  Dim DiscApplied As Boolean '1/16/2007
  Dim SaveAmt As Double '1/16/2007
  
  On Error GoTo ERRORSTUFF
  
  PrePayDone = False '7/10/06
  
  QFlag = False
  If chkQuick.Value = 1 Then QFlag = True
  DiscApplied = False '1/16/07
  
  TransDesc = QPTrim$(fptxtDesc.Text)
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  Opt1Desc = QPTrim$(TaxMasterRec.POptRev1)
  Opt2Desc = QPTrim$(TaxMasterRec.POptRev2)
  Opt3Desc = QPTrim$(TaxMasterRec.POptRev3)
  CustName = ""
  IdxFlag = False
  OptFlag = False
  If CheckB4Printing = False Then
    Exit Sub
  End If
  FF$ = Chr$(12)
  MaxLines = 56
  LineCnt = 0
  ThisBillType = QPTrim$(fpcmbTaxType.Text)
  If fpcmbIncInactive.Text = "No" Then
    InactiveFlag = False
  Else
    InactiveFlag = True
  End If
  
  If Mid(fpcmbTransType.Text, 1, 1) = "" Then
    ThisClass = CInt(Mid(fpcmbTransType.Text, 2, 1))
  Else
    ThisClass = CInt(Mid(fpcmbTransType.Text, 1, 2))
  End If
    
  Select Case ThisClass
    Case 1
      ThisType = "Billing"
    Case 2
      ThisType = "Payment"
    Case 3
      ThisType = "Release"
    Case 4
      ThisType = "Interest"
    Case 5
      ThisType = "Penalty"
    Case 6
      ThisType = "Advertising Charge"
    Case 7, 10
      ThisType = "Adjust Pay Down"
    Case 9
      ThisType = "Credit Applied at Billing"
    Case 11
      ThisType = "Adjust Prepay Down"
    Case 12
      ThisType = "Refund Prepay"
    Case 13
      ThisType = "Adjust Bill Down"
    Case 14, 24
      ThisType = "Adjust Bill Up"
    Case 21
      ThisType = "Payment w/Overpay"
    Case 22
      ThisType = "Overpayment Only"
    Case 30
      ThisType = "PPTRA Removal"
    Case Else
      ThisType = "All"
  End Select
    
  BegDate = Date2Num(fptxtBegDate.Text)
  EndDate = Date2Num(fptxtEndDate.Text)
  
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
    OptFlag = True
  End If

  RptFile$ = "TAXRPTS\TAXPDJRNL.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  GoSub PrintHeader
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  
  frmVATaxShowPctComp.Label1 = "Gathering Tax Transaction Data"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  
  ReDim TotByYrAndType(1 To 18, 1 To 1) As Double
  ReDim CntByYrAndType(1 To 18, 1 To 1) As Integer
  ReDim TotByType(1 To 18) As Double
  ReDim TotCntByType(1 To 18) As Long
  ReDim ThEYear(1 To 1) As Integer
  
  ReDim TotByYrAndPers(1 To 1) As Double
  ReDim TotByYrAndMT(1 To 1) As Double
  ReDim TotByYrAndMC(1 To 1) As Double
  ReDim TotByYrAndFE(1 To 1) As Double
  ReDim TotByYrAndMH(1 To 1) As Double
  ReDim TotByYrAndInt(1 To 1) As Double
  ReDim TotByYrAndAdv(1 To 1) As Double
  ReDim TotByYrAndLateList(1 To 1) As Double
  ReDim TotByYrAndPen(1 To 1) As Double
  ReDim TotByYrAndOpt1(1 To 1) As Double
  ReDim TotByYrAndOpt2(1 To 1) As Double
  ReDim TotByYrAndOpt3(1 To 1) As Double
  
  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  End If
  If InStr(fpcmbTransType.Text, "All") Then GoTo PrintAll
  
  For x = 1 To NumOfTCRecs
    If IdxFlag = False Then
      Get TCHandle, x, TaxCust
      ThisCustRec = x
    Else
      Get TCHandle, IdxArray(x), TaxCust
      ThisCustRec = IdxArray(x)
    End If
    If TaxCust.Active = "N" And InactiveFlag = False Then
      GoTo SkipIt
    End If
    
    CustBal = GetCustBalance(ThisCustRec, -1)
    
    ThisName = QPTrim$(TaxCust.CustName)
    
    ThisRec = TaxCust.LastTrans
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintHeader
    End If
    PrintCnt = 0
    BadCnt = 0
    Do While ThisRec > 0
      Get TTHandle, ThisRec, TaxTrans
      If QFlag = True Then
        If TaxTrans.TransDate < BegDate Then
          BadCnt = BadCnt + 1
          If BadCnt > 3 Then Exit Do
        End If
      End If
      If TransDesc <> "" Then
        If InStr(1, TaxTrans.Description, TransDesc) = 0 Then GoTo SkipIt
      End If
      If TaxTrans.BillType <> "R" And TaxTrans.BillType <> "P" Then TaxTrans.BillType = ""
      If TaxTrans.BillType <> Mid(fpcmbTaxType.Text, 1, 1) And QPTrim$(TaxTrans.BillType) <> "" Then GoTo SkipIt
      If ThisClass = 7 And TaxTrans.TranType = 10 Then GoTo ItsOK
      If ThisClass = 14 And TaxTrans.TranType = 24 Then GoTo ItsOK
'      If ThisClass = 2 And TaxTrans.TranType = 21 Then GoTo ItsOK 'commented out 7/6/06
      If ThisClass <> 0 And TaxTrans.TranType <> ThisClass Then GoTo SkipIt
ItsOK:
      If TaxTrans.TransDate >= BegDate And TaxTrans.TransDate <= EndDate Then
        If TaxTrans.BelongTo > 0 Then
          Get TTHandle, TaxTrans.BelongTo, TaxTrans
        End If
        If PrintCnt = 0 Then
          If LineCnt <> 9 Then
            Print #RptHandle,
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintHeader
            End If
          End If
          GoSub PrintCustHeader
        End If
        PrintCnt = PrintCnt + 1
        If YrCnt = 0 Then
           YrCnt = YrCnt + 1
           ThisYear = YrCnt
           ReDim Preserve ThEYear(1 To YrCnt) As Integer
           ThEYear(YrCnt) = TaxTrans.TaxYear
           ReDim Preserve TotByYrAndType(1 To 18, 1 To YrCnt) As Double
           ReDim Preserve CntByYrAndType(1 To 18, 1 To YrCnt) As Integer
           ReDim Preserve TotByYrAndPers(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndMT(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndMC(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndFE(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndMH(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndInt(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndAdv(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndLateList(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndPen(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndOpt1(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndOpt2(1 To YrCnt) As Double
           ReDim Preserve TotByYrAndOpt3(1 To YrCnt) As Double
            
           TotByYrAndPers(YrCnt) = 0
           TotByYrAndMT(YrCnt) = 0
           TotByYrAndMC(YrCnt) = 0
           TotByYrAndFE(YrCnt) = 0
           TotByYrAndMH(YrCnt) = 0
           TotByYrAndInt(YrCnt) = 0
           TotByYrAndAdv(YrCnt) = 0
           TotByYrAndLateList(YrCnt) = 0
           TotByYrAndPen(YrCnt) = 0
           TotByYrAndOpt1(YrCnt) = 0
           TotByYrAndOpt2(YrCnt) = 0
           TotByYrAndOpt3(YrCnt) = 0
           ReDim Preserve TotByYrAndType(1 To 18, 1 To YrCnt) As Double
           ReDim Preserve CntByYrAndType(1 To 18, 1 To YrCnt) As Integer
           For y = 1 To 18
             TotByYrAndType(y, YrCnt) = 0
             CntByYrAndType(y, YrCnt) = 0
           Next y
         Else
           For y = 1 To YrCnt
             If TaxTrans.TaxYear = ThEYear(y) Then
               ThisYear = y
               Exit For
             End If
           Next y
           If y > YrCnt Then
             YrCnt = YrCnt + 1
             ThisYear = YrCnt
             ReDim Preserve ThEYear(1 To YrCnt) As Integer
             ThEYear(YrCnt) = TaxTrans.TaxYear
             ReDim Preserve TotByYrAndType(1 To 18, 1 To YrCnt) As Double
             ReDim Preserve CntByYrAndType(1 To 18, 1 To YrCnt) As Integer
             ReDim Preserve TotByYrAndPers(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndMT(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndMC(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndFE(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndMH(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndInt(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndAdv(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndLateList(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndPen(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndPers(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndOpt1(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndOpt2(1 To YrCnt) As Double
             ReDim Preserve TotByYrAndOpt3(1 To YrCnt) As Double
             TotByYrAndPers(YrCnt) = 0
             TotByYrAndMT(YrCnt) = 0
             TotByYrAndMC(YrCnt) = 0
             TotByYrAndFE(YrCnt) = 0
             TotByYrAndMH(YrCnt) = 0
             TotByYrAndInt(YrCnt) = 0
             TotByYrAndAdv(YrCnt) = 0
             TotByYrAndLateList(YrCnt) = 0
             TotByYrAndPen(YrCnt) = 0
             TotByYrAndOpt1(YrCnt) = 0
             TotByYrAndOpt2(YrCnt) = 0
             TotByYrAndOpt3(YrCnt) = 0
             ReDim Preserve TotByYrAndType(1 To 18, 1 To YrCnt) As Double
             ReDim Preserve CntByYrAndType(1 To 18, 1 To YrCnt) As Integer
             For y = 1 To 18
               TotByYrAndType(y, YrCnt) = 0
               CntByYrAndType(y, YrCnt) = 0
             Next y
           End If
         End If
         Get TTHandle, ThisRec, TaxTrans
          
         Select Case TaxTrans.TranType
           Case 1
             ThisTransType = "Billing"
             TotByYrAndType(1, ThisYear) = OldRound(TotByYrAndType(1, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(1, ThisYear) = OldRound(CntByYrAndType(1, ThisYear) + 1)
             TotByType(1) = OldRound(TotByType(1) + TaxTrans.Amount)
             TotCntByType(1) = OldRound(TotCntByType(1) + 1)
             TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1)
             TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2)
             TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3)
             TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4)
             TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5)
             '7/11/06 added back interest, advertising and penalty to accommodate manual bills
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.Interest)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.Collection)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.Penalty)
             
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateList)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3)
           Case 2
             ThisTransType = "Payment"
             If TaxTrans.DiscAmt > 0 Then 'added 1/16/07
               GoSub ApplyDisc
             End If
             TotByYrAndType(2, ThisYear) = OldRound(TotByYrAndType(2, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(2, ThisYear) = OldRound(CntByYrAndType(2, ThisYear) + 1)
             TotByType(2) = OldRound(TotByType(2) + TaxTrans.Amount)
             TotCntByType(2) = OldRound(TotCntByType(2) + 1)
             TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2Pd)
             TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3Pd)
             TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4Pd)
             TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5Pd)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
           Case 3
            '7/12/06 changed revenues for Release to paid from charged
             ThisTransType = "Release"
             TotByYrAndType(3, ThisYear) = OldRound(TotByYrAndType(3, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(3, ThisYear) = OldRound(CntByYrAndType(3, ThisYear) + 1)
             TotByType(3) = OldRound(TotByType(3) + TaxTrans.Amount)
             TotCntByType(3) = OldRound(TotCntByType(3) + 1)
             TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2Pd)
             TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3Pd)
             TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4Pd)
             TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5Pd)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
           Case 4
             ThisTransType = "Interest"
             TotByYrAndType(4, ThisYear) = OldRound(TotByYrAndType(4, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(4, ThisYear) = OldRound(CntByYrAndType(4, ThisYear) + 1)
             TotByType(4) = OldRound(TotByType(4) + TaxTrans.Amount)
             TotCntByType(4) = OldRound(TotCntByType(4) + 1)
           Case 5
             ThisTransType = "Penalty"
             TotByYrAndType(5, ThisYear) = OldRound(TotByYrAndType(5, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(5, ThisYear) = OldRound(CntByYrAndType(5, ThisYear) + 1)
             TotByType(5) = OldRound(TotByType(5) + TaxTrans.Amount)
             TotCntByType(5) = OldRound(TotCntByType(5) + 1)
           Case 6
             ThisTransType = "Advertising Charge"
             TotByYrAndType(6, ThisYear) = OldRound(TotByYrAndType(6, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(6, ThisYear) = OldRound(CntByYrAndType(6, ThisYear) + 1)
             TotByType(6) = OldRound(TotByType(6) + TaxTrans.Amount)
             TotCntByType(6) = OldRound(TotCntByType(6) + 1)
           Case 7
             ThisTransType = "Adjust Pay Down"
             TotByYrAndType(7, ThisYear) = OldRound(TotByYrAndType(7, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(7, ThisYear) = OldRound(CntByYrAndType(7, ThisYear) + 1)
             TotByType(7) = OldRound(TotByType(7) + TaxTrans.Amount)
             TotCntByType(7) = OldRound(TotCntByType(7) + 1)
             TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2Pd)
             TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3Pd)
             TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4Pd)
             TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5Pd)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
           Case 9
             ThisTransType = "Credit Applied at Billing"
             TotByYrAndType(8, ThisYear) = OldRound(TotByYrAndType(8, ThisYear) + TaxTrans.Revenue.PrePaidUsed)
             CntByYrAndType(8, ThisYear) = OldRound(CntByYrAndType(8, ThisYear) + 1)
             TotByType(8) = OldRound(TotByType(8) + TaxTrans.Revenue.PrePaidUsed)
             TotCntByType(8) = OldRound(TotCntByType(8) + 1)
           Case 13
             ThisTransType = "Adjust Bill Down"
             TotByYrAndType(9, ThisYear) = OldRound(TotByYrAndType(9, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(9, ThisYear) = OldRound(CntByYrAndType(9, ThisYear) + 1)
             TotByType(9) = OldRound(TotByType(9) + TaxTrans.Amount)
             TotCntByType(9) = OldRound(TotCntByType(9) + 1)
             TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1)
             TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2)
             TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3)
             TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4)
             TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.Interest)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.Collection)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateList)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.Penalty)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3)
           Case 14
             ThisTransType = "Adjust Bill Up"
             TotByYrAndType(10, ThisYear) = OldRound(TotByYrAndType(10, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(10, ThisYear) = OldRound(CntByYrAndType(10, ThisYear) + 1)
             TotByType(10) = OldRound(TotByType(10) + TaxTrans.Amount)
             TotCntByType(10) = OldRound(TotCntByType(10) + 1)
             TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1)
             TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2)
             TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3)
             TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4)
             TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.Interest)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.Collection)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateList)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.Penalty)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3)
           Case 21
             ThisTransType = "Billpay/Overpay" 'changed Amount to PrePaidAmt
             If TaxTrans.DiscAmt > 0 Then 'added 1/16/07
               GoSub ApplyDisc
             End If
             If fpcmbTransType.Text <> " 0) All" Then 'added the All if statement on 7/7/06
               TotByYrAndType(11, ThisYear) = OldRound(TotByYrAndType(11, ThisYear) + TaxTrans.Revenue.PrePaidAmt)
             Else
               TotByYrAndType(11, ThisYear) = OldRound(TotByYrAndType(11, ThisYear) + TaxTrans.Amount)
             End If
'             TotByYrAndType(11, ThisYear) = OldRound(TotByYrAndType(11, ThisYear) + TaxTrans.Revenue.PrePaidAmt) '.Amount)
             CntByYrAndType(11, ThisYear) = OldRound(CntByYrAndType(11, ThisYear) + 1)
             TotByType(11) = OldRound(TotByType(11) + TaxTrans.Amount)
             TotCntByType(11) = OldRound(TotCntByType(11) + 1)
             TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2Pd)
             TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3Pd)
             TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4Pd)
             TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5Pd)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
           Case 22
             ThisTransType = "Overpayment"
             TotByYrAndType(12, ThisYear) = OldRound(TotByYrAndType(12, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(12, ThisYear) = OldRound(CntByYrAndType(12, ThisYear) + 1)
             TotByType(12) = OldRound(TotByType(12) + TaxTrans.Amount)
             TotCntByType(12) = OldRound(TotCntByType(12) + 1)
           Case 24
             ThisTransType = "Adj Bill Up -Cre"
             TotByYrAndType(13, ThisYear) = OldRound(TotByYrAndType(13, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(13, ThisYear) = OldRound(CntByYrAndType(13, ThisYear) + 1)
             TotByType(13) = OldRound(TotByType(13) + TaxTrans.Amount)
             TotCntByType(13) = OldRound(TotCntByType(13) + 1)
             TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2Pd)
             TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3Pd)
             TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4Pd)
             TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5Pd)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
'             TotByYrAndType(13, ThisYear) = OldRound(TotByYrAndType(13, ThisYear) + TaxTrans.Amount)
'             CntByYrAndType(13, ThisYear) = OldRound(CntByYrAndType(13, ThisYear) + 1)
'             TotByType(13) = OldRound(TotByType(13) + TaxTrans.Amount)
'             TotCntByType(13) = OldRound(TotCntByType(13) + 1)
'             TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1)
'             TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2)
'             TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3)
'             TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4)
'             TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5)
'             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.Interest)
'             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.Collection)
'             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateList)
'             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.Penalty)
'             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1)
'             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2)
'             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3)
           Case 10 ', 24 '7/11/06 added Pd on revenues
             ThisTransType = "Adj Pay Dwn -Cre"
             TotByYrAndType(14, ThisYear) = OldRound(TotByYrAndType(14, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(14, ThisYear) = OldRound(CntByYrAndType(14, ThisYear) + 1)
             TotByType(14) = OldRound(TotByType(14) + TaxTrans.Amount)
             TotCntByType(14) = OldRound(TotCntByType(14) + 1)
             TotByYrAndPers(ThisYear) = OldRound(TotByYrAndPers(ThisYear) + TaxTrans.Revenue.Principle1Pd)
             TotByYrAndMT(ThisYear) = OldRound(TotByYrAndMT(ThisYear) + TaxTrans.Revenue.Principle2Pd)
             TotByYrAndMC(ThisYear) = OldRound(TotByYrAndMC(ThisYear) + TaxTrans.Revenue.Principle3Pd)
             TotByYrAndFE(ThisYear) = OldRound(TotByYrAndFE(ThisYear) + TaxTrans.Revenue.Principle4Pd)
             TotByYrAndMH(ThisYear) = OldRound(TotByYrAndMH(ThisYear) + TaxTrans.Revenue.Principle5Pd)
             TotByYrAndInt(ThisYear) = OldRound(TotByYrAndInt(ThisYear) + TaxTrans.Revenue.InterestPd)
             TotByYrAndAdv(ThisYear) = OldRound(TotByYrAndAdv(ThisYear) + TaxTrans.Revenue.CollectionPd)
             TotByYrAndLateList(ThisYear) = OldRound(TotByYrAndLateList(ThisYear) + TaxTrans.Revenue.LateListPd)
             TotByYrAndPen(ThisYear) = OldRound(TotByYrAndPen(ThisYear) + TaxTrans.Revenue.PenaltyPd)
             TotByYrAndOpt1(ThisYear) = OldRound(TotByYrAndOpt1(ThisYear) + TaxTrans.Revenue.RevOpt1Pd)
             TotByYrAndOpt2(ThisYear) = OldRound(TotByYrAndOpt2(ThisYear) + TaxTrans.Revenue.RevOpt2Pd)
             TotByYrAndOpt3(ThisYear) = OldRound(TotByYrAndOpt3(ThisYear) + TaxTrans.Revenue.RevOpt3Pd)
           Case 11
             ThisTransType = "Adj Prepay -Cre"
             TotByYrAndType(15, ThisYear) = OldRound(TotByYrAndType(15, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(15, ThisYear) = OldRound(CntByYrAndType(15, ThisYear) + 1)
             TotByType(15) = OldRound(TotByType(15) + TaxTrans.Amount)
             TotCntByType(15) = OldRound(TotCntByType(15) + 1)
           Case 12
             ThisTransType = "Ref Prepay -Cre"
             TotByYrAndType(16, ThisYear) = OldRound(TotByYrAndType(16, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(16, ThisYear) = OldRound(CntByYrAndType(16, ThisYear) + 1)
             TotByType(16) = OldRound(TotByType(16) + TaxTrans.Amount)
             TotCntByType(16) = OldRound(TotCntByType(16) + 1)
           Case 30
             ThisTransType = "PPTRA Removal"
             TotByYrAndType(17, ThisYear) = OldRound(TotByYrAndType(17, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(17, ThisYear) = OldRound(CntByYrAndType(17, ThisYear) + 1)
             TotByType(17) = OldRound(TotByType(17) + TaxTrans.Amount)
             TotCntByType(17) = OldRound(TotCntByType(17) + 1)
           Case Else
             ThisTransType = "Unknown"
             TotByYrAndType(18, ThisYear) = OldRound(TotByYrAndType(18, ThisYear) + TaxTrans.Amount)
             CntByYrAndType(18, ThisYear) = OldRound(CntByYrAndType(18, ThisYear) + 1)
             TotByType(18) = OldRound(TotByType(18) + TaxTrans.Amount)
             TotCntByType(18) = OldRound(TotCntByType(18) + 1)
        End Select
        TCnt = TCnt + 1
        TotAmt = OldRound(TotAmt + TaxTrans.Amount)
        PersDif = OldRound(TaxTrans.Revenue.Principle1 - TaxTrans.Revenue.Principle1Pd)
        IntDif = OldRound(TaxTrans.Revenue.Interest - TaxTrans.Revenue.InterestPd)
        MTDif = OldRound(TaxTrans.Revenue.Principle2 - TaxTrans.Revenue.Principle2Pd)
        MCDif = OldRound(TaxTrans.Revenue.Principle3 - TaxTrans.Revenue.Principle3Pd)
        FEDif = OldRound(TaxTrans.Revenue.Principle4 - TaxTrans.Revenue.Principle4Pd)
        MHDif = OldRound(TaxTrans.Revenue.Principle5 - TaxTrans.Revenue.Principle5Pd)
        LateListDif = OldRound(TaxTrans.Revenue.LateList - TaxTrans.Revenue.LateListPd)
        PenDif = OldRound(TaxTrans.Revenue.Penalty - TaxTrans.Revenue.PenaltyPd)
        Opt1Dif = OldRound(TaxTrans.Revenue.RevOpt1 - TaxTrans.Revenue.RevOpt1Pd)
        Opt2Dif = OldRound(TaxTrans.Revenue.RevOpt2 - TaxTrans.Revenue.RevOpt2Pd)
        Opt3Dif = OldRound(TaxTrans.Revenue.RevOpt3 - TaxTrans.Revenue.RevOpt3Pd)
        ThisBal = OldRound(PersDif + IntDif + MTDif + MCDif + FEDif + MHDif + PenDif + Opt1Dif + Opt2Dif + Opt3Dif)
        Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); QPTrim$(TaxTrans.Description);
        If TaxTrans.BelongTo > 0 Then
          Get TTHandle, TaxTrans.BelongTo, TaxTrans
          Print #RptHandle, Tab(37); Using$("###0", TaxTrans.TaxYear);
          Get TTHandle, ThisRec, TaxTrans
          If TaxTrans.DiscAmt > 0 Then 'added 1/16/07
            GoSub ApplyDisc
          End If
        Else
          Print #RptHandle, Tab(37); Using$("###0", TaxTrans.TaxYear);
        End If
        Print #RptHandle, Tab(45); Using$("$##,##0.00", TaxTrans.Revenue.PrePaidAmt);
        If TaxTrans.TranType <> 9 Then
          Print #RptHandle, Tab(56); Using$("$##,##0.00", TaxTrans.Amount); Tab(69);
        Else
          Print #RptHandle, Tab(56); Using$("$##,##0.00", TaxTrans.Revenue.PrePaidUsed); Tab(69);
        End If
        LineCnt = LineCnt + 1
        If TaxTrans.BelongTo > 0 Then
          Get TTHandle, TaxTrans.BelongTo, TaxTrans
          ThisBillNum = ParseBillNum(TaxTrans.Description)
          If IsNumeric(ThisBillNum) Then
            Print #RptHandle, Using$("######", CDbl(ThisBillNum));
          Else
            Print #RptHandle, "   " + ThisBillNum;
          End If
        Else
          Print #RptHandle, "     ";
        End If
      
        Get TTHandle, ThisRec, TaxTrans
        If TaxTrans.DiscAmt > 0 Then 'added 1/16/07
          GoSub ApplyDisc
        End If
        Print #RptHandle, Tab(79); ThisTransType
        If LineCnt >= MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintHeader
        End If
        If ThisType <> "Billing" Then
          Print #RptHandle, Tab(3); "Oper #: " + CStr(TaxTrans.OperNum);
          Print #RptHandle, Tab(15); "Personal          "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle1); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle1Pd)
          Print #RptHandle, Tab(15); "Machine Tools     "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle2); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle2Pd)
          Print #RptHandle, Tab(15); "Merchant Capital  "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle3); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle3Pd)
          Print #RptHandle, Tab(15); "Farm Equipment    "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle4); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle4Pd)
          Print #RptHandle, Tab(15); "Mobile Homes      "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle5); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle5Pd)
          '7/11/06 added back interest and advertising to accommodate manual bills
          Print #RptHandle, Tab(15); "Interest          "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Interest); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.InterestPd)
          Print #RptHandle, Tab(15); "Late Listing      "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.LateList); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.LateListPd) '; Tab(80); Using$("$##,##0.00", LateListDif)
          Print #RptHandle, Tab(15); "Penalty           "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Penalty); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.PenaltyPd)
          LineCnt = LineCnt + 8
          If Len(Opt1Desc) > 0 Then
            Print #RptHandle, Tab(15); Opt1Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1Pd)
            LineCnt = LineCnt + 1
          End If
          If Len(Opt2Desc) > 0 Then
            Print #RptHandle, Tab(15); Opt2Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2Pd)
            LineCnt = LineCnt + 1
          End If
            If Len(Opt3Desc) > 0 Then
            Print #RptHandle, Tab(15); Opt3Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3Pd)
            LineCnt = LineCnt + 1
          End If
        Else
          Print #RptHandle, Tab(3); "Oper #: " + CStr(TaxTrans.OperNum);
          Print #RptHandle, Tab(15); "Personal          "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle1); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle1Pd); Tab(80); Using$("$##,##0.00", PersDif)
          Print #RptHandle, Tab(15); "Machine Tools     "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle2); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle2Pd); Tab(80); Using$("$##,##0.00", MTDif)
          Print #RptHandle, Tab(15); "Merchant Capital  "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle3); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle3Pd); Tab(80); Using$("$##,##0.00", MCDif)
          Print #RptHandle, Tab(15); "Farm Equipment    "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle4); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle4Pd); Tab(80); Using$("$##,##0.00", FEDif)
          Print #RptHandle, Tab(15); "Mobile Homes      "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle5); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle5Pd); Tab(80); Using$("$##,##0.00", MHDif)
          Print #RptHandle, Tab(15); "Interest          "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Interest); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.InterestPd); Tab(80); Using$("$##,##0.00", IntDif)
          Print #RptHandle, Tab(15); "Late Listing      "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.LateList); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.LateListPd); Tab(80); Using$("$##,##0.00", LateListDif)
          Print #RptHandle, Tab(15); "Penalty           "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Penalty); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.PenaltyPd); Tab(80); Using$("$##,##0.00", PenDif)
          LineCnt = LineCnt + 8
          If Len(Opt1Desc) > 0 Then
            Print #RptHandle, Tab(15); Opt1Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1Pd); Tab(80); Using$("$##,##0.00", Opt1Dif)
            LineCnt = LineCnt + 1
          End If
          If Len(Opt2Desc) > 0 Then
            Print #RptHandle, Tab(15); Opt2Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2Pd); Tab(80); Using$("$##,##0.00", Opt2Dif)
            LineCnt = LineCnt + 1
          End If
            If Len(Opt3Desc) > 0 Then
            Print #RptHandle, Tab(15); Opt3Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3Pd); Tab(80); Using$("$##,##0.00", Opt3Dif)
            LineCnt = LineCnt + 1
          End If
          Print #RptHandle, Tab(15); "Bill Balance:"; Tab(80); Using$("$##,##0.00", ThisBal)
        End If
    End If
SkipIt:
    ThisRec = TaxTrans.LastTrans
    Loop
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
  
  If TCnt = 0 Then
    Call TaxMsg(900, "No transactions were found that fit the parameters entered.")
    Close
    Exit Sub
  End If
  If YrCnt > 0 Then
    GoSub SortIt
    GoSub PrintTotals
  End If
  Print #RptHandle, FF$
  Close
  ViewPrint RptFile, "Tax Transactions Report", True
  
  Exit Sub
  
SortIt:
  BigYr = 0
  For x = 1 To YrCnt
    If ThEYear(x) > BigYr Then
      BigYr = ThEYear(x)
    End If
  Next x
  ReDim HoldAmt(1 To 18, 1 To YrCnt) As Double
  ReDim HoldCnt(1 To 18, 1 To YrCnt) As Integer
  Nexty = 1
  Nextx = 1
  HoldBigYr = 0
  For x = 1 To 18
    For y = Nexty To YrCnt
      If ThEYear(y) >= HoldBigYr Then
        HoldBigYr = ThEYear(y)
        Thisx = x
        Thisy = y
      End If
    Next y
    For z = 1 To 18
      HoldAmt(z, Thisy) = TotByYrAndType(z, Nexty)
      HoldCnt(z, Thisy) = CntByYrAndType(z, Nexty)
    Next z
    HoldYr = ThEYear(Nexty)
    If fpcmbTransType.Text = " 0) All" Then GoTo All1
    HoldPers = TotByYrAndPers(Nexty)
    HoldMT = TotByYrAndMT(Nexty)
    HoldMC = TotByYrAndMC(Nexty)
    HoldFE = TotByYrAndFE(Nexty)
    HoldMH = TotByYrAndMH(Nexty)
    HoldInt = TotByYrAndInt(Nexty)
    HoldAdv = TotByYrAndAdv(Nexty)
    HoldPen = TotByYrAndPen(Nexty)
    HoldLateList = TotByYrAndLateList(Nexty)
    HoldOpt1 = TotByYrAndOpt1(Nexty)
    HoldOpt2 = TotByYrAndOpt2(Nexty)
    HoldOpt3 = TotByYrAndOpt3(Nexty)
All1:
    For z = 1 To 18
      TotByYrAndType(z, Nexty) = TotByYrAndType(z, Thisy)
      CntByYrAndType(z, Nexty) = CntByYrAndType(z, Thisy)
    Next z
    ThEYear(Nexty) = ThEYear(Thisy)
    If fpcmbTransType.Text = " 0) All" Then GoTo All2
    TotByYrAndPers(Nexty) = TotByYrAndPers(Thisy)
    TotByYrAndMT(Nexty) = TotByYrAndMT(Thisy)
    TotByYrAndMC(Nexty) = TotByYrAndMC(Thisy)
    TotByYrAndFE(Nexty) = TotByYrAndFE(Thisy)
    TotByYrAndMH(Nexty) = TotByYrAndMH(Thisy)
    TotByYrAndInt(Nexty) = TotByYrAndInt(Thisy)
    TotByYrAndAdv(Nexty) = TotByYrAndAdv(Thisy)
    TotByYrAndPen(Nexty) = TotByYrAndPen(Thisy)
    TotByYrAndLateList(Nexty) = TotByYrAndLateList(Thisy)
    TotByYrAndOpt1(Nexty) = TotByYrAndOpt1(Thisy)
    TotByYrAndOpt2(Nexty) = TotByYrAndOpt2(Thisy)
    TotByYrAndOpt3(Nexty) = TotByYrAndOpt3(Thisy)
All2:
    For z = 1 To 18
      TotByYrAndType(z, Thisy) = HoldAmt(z, Thisy)
      CntByYrAndType(z, Thisy) = HoldCnt(z, Thisy)
    Next z
    ThEYear(Thisy) = HoldYr
    If fpcmbTransType.Text = " 0) All" Then GoTo All3
    TotByYrAndPers(Thisy) = HoldPers
    TotByYrAndMT(Thisy) = HoldMT
    TotByYrAndMC(Thisy) = HoldMC
    TotByYrAndFE(Thisy) = HoldFE
    TotByYrAndMH(Thisy) = HoldMH
    TotByYrAndInt(Thisy) = HoldInt
    TotByYrAndAdv(Thisy) = HoldAdv
    TotByYrAndPen(Thisy) = HoldPen
    TotByYrAndLateList(Thisy) = HoldLateList
    TotByYrAndOpt1(Thisy) = HoldOpt1
    TotByYrAndOpt2(Thisy) = HoldOpt2
    TotByYrAndOpt3(Thisy) = HoldOpt3
All3:
    If Nexty >= YrCnt Then Exit For
    HoldBigYr = 0 'BigYr + 1
    Nexty = Nexty + 1
  Next x
  Print #RptHandle, FF$
  GoSub PrintSortHeader
  Print #RptHandle, "Total Transaction Count: " + Using$("#####0", TCnt)
  Print #RptHandle, String(94, "-")
  LineCnt = LineCnt + 2
  For y = 1 To YrCnt
    If LineCnt >= MaxLines - 4 Then
      Print #RptHandle, FF$
      GoSub PrintSortHeader
      Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
    Else
      Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
    End If
    LineCnt = LineCnt + 1
    ThisYrCnt = 0
    For x = 1 To 18
      If TotByYrAndType(x, y) > 0 Then
        ThisYrCnt = OldRound(ThisYrCnt + CntByYrAndType(x, y))
        Select Case x
          Case 1
            Print #RptHandle, "  Billing"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOne
            Print #RptHandle, Tab(5); "Personal Prop:"; Tab(55); Using$("$###,###,##0.00", TotByYrAndPers(y))
            Print #RptHandle, Tab(5); "Machine Tools:"; Tab(55); Using$("$###,###,##0.00", TotByYrAndMT(y))
            Print #RptHandle, Tab(5); "Merchant Cap: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMC(y))
            Print #RptHandle, Tab(5); "Farm Equip:   "; Tab(55); Using$("$###,###,##0.00", TotByYrAndFE(y))
            Print #RptHandle, Tab(5); "Mobile Homes: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMH(y))
            LineCnt = LineCnt + 7
            If LineCnt >= MaxLines - 3 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 2
            Print #RptHandle, "  Payment"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOne
            Print #RptHandle, Tab(5); "Personal Prop:"; Tab(55); Using$("$###,###,##0.00", TotByYrAndPers(y))
            Print #RptHandle, Tab(5); "Machine Tools:"; Tab(55); Using$("$###,###,##0.00", TotByYrAndMT(y))
            Print #RptHandle, Tab(5); "Merchant Cap: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMC(y))
            Print #RptHandle, Tab(5); "Farm Equip:   "; Tab(55); Using$("$###,###,##0.00", TotByYrAndFE(y))
            Print #RptHandle, Tab(5); "Mobile Homes: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMH(y))
            Print #RptHandle, Tab(5); "Interest:     "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 9
            If LineCnt >= MaxLines - 3 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 3
            Print #RptHandle, "  Release"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOne
            Print #RptHandle, Tab(5); "Personal Prop:"; Tab(55); Using$("$###,###,##0.00", TotByYrAndPers(y))
            Print #RptHandle, Tab(5); "Machine Tools:"; Tab(55); Using$("$###,###,##0.00", TotByYrAndMT(y))
            Print #RptHandle, Tab(5); "Merchant Cap: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMC(y))
            Print #RptHandle, Tab(5); "Farm Equip:   "; Tab(55); Using$("$###,###,##0.00", TotByYrAndFE(y))
            Print #RptHandle, Tab(5); "Mobile Homes: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMH(y))
            Print #RptHandle, Tab(5); "Interest:     "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 9
            If LineCnt >= MaxLines - 3 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 4
            Print #RptHandle, "  Interest"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
          Case 5
            Print #RptHandle, "  Penalty"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
          Case 6
            Print #RptHandle, "  Advertising"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
          Case 7
            Print #RptHandle, "  Adjust Pay Down"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOne
            Print #RptHandle, Tab(5); "Personal Prop:"; Tab(55); Using$("$###,###,##0.00", TotByYrAndPers(y))
            Print #RptHandle, Tab(5); "Machine Tools:"; Tab(55); Using$("$###,###,##0.00", TotByYrAndMT(y))
            Print #RptHandle, Tab(5); "Merchant Cap: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMC(y))
            Print #RptHandle, Tab(5); "Farm Equip:   "; Tab(55); Using$("$###,###,##0.00", TotByYrAndFE(y))
            Print #RptHandle, Tab(5); "Mobile Homes: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMH(y))
            Print #RptHandle, Tab(5); "Interest:     "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 9
            If LineCnt >= MaxLines - 3 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 8
            Print #RptHandle, "  Credit at Billing"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
          Case 9
            Print #RptHandle, "  Adjust Bill Down"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOne
            Print #RptHandle, Tab(5); "Personal Prop:"; Tab(55); Using$("$###,###,##0.00", TotByYrAndPers(y))
            Print #RptHandle, Tab(5); "Machine Tools:"; Tab(55); Using$("$###,###,##0.00", TotByYrAndMT(y))
            Print #RptHandle, Tab(5); "Merchant Cap: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMC(y))
            Print #RptHandle, Tab(5); "Farm Equip:   "; Tab(55); Using$("$###,###,##0.00", TotByYrAndFE(y))
            Print #RptHandle, Tab(5); "Mobile Homes: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMH(y))
            Print #RptHandle, Tab(5); "Interest:     "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 9
            If LineCnt >= MaxLines - 3 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 10
            Print #RptHandle, "  Adjust Bill Up"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOne
            Print #RptHandle, Tab(5); "Personal Prop:"; Tab(55); Using$("$###,###,##0.00", TotByYrAndPers(y))
            Print #RptHandle, Tab(5); "Machine Tools:"; Tab(55); Using$("$###,###,##0.00", TotByYrAndMT(y))
            Print #RptHandle, Tab(5); "Merchant Cap: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMC(y))
            Print #RptHandle, Tab(5); "Farm Equip:   "; Tab(55); Using$("$###,###,##0.00", TotByYrAndFE(y))
            Print #RptHandle, Tab(5); "Mobile Homes: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMH(y))
            Print #RptHandle, Tab(5); "Interest:     "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 9
            If LineCnt >= MaxLines - 3 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 11 'added detail on 7/6/06
            If fpcmbTransType.Text = " 0) All" Then 'added Bill Pay/OverPay on 7/7/06
              Print #RptHandle, "  Bill Pay/OverPay"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            Else
              Print #RptHandle, "  Bill OverPay"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            End If
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOne
            Print #RptHandle, Tab(5); "Personal Prop:"; Tab(55); Using$("$###,###,##0.00", TotByYrAndPers(y))
            Print #RptHandle, Tab(5); "Machine Tools:"; Tab(55); Using$("$###,###,##0.00", TotByYrAndMT(y))
            Print #RptHandle, Tab(5); "Merchant Cap: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMC(y))
            Print #RptHandle, Tab(5); "Farm Equip:   "; Tab(55); Using$("$###,###,##0.00", TotByYrAndFE(y))
            Print #RptHandle, Tab(5); "Mobile Homes: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMH(y))
            Print #RptHandle, Tab(5); "Interest:     "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 9
            If LineCnt >= MaxLines - 3 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 12
            Print #RptHandle, "  OverPayment"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
          Case 13
            Print #RptHandle, "  Adj Bill Up: -Credit "; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOne
            Print #RptHandle, Tab(5); "Personal Prop:"; Tab(55); Using$("$###,###,##0.00", TotByYrAndPers(y))
            Print #RptHandle, Tab(5); "Machine Tools:"; Tab(55); Using$("$###,###,##0.00", TotByYrAndMT(y))
            Print #RptHandle, Tab(5); "Merchant Cap: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMC(y))
            Print #RptHandle, Tab(5); "Farm Equip:   "; Tab(55); Using$("$###,###,##0.00", TotByYrAndFE(y))
            Print #RptHandle, Tab(5); "Mobile Homes: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMH(y))
            Print #RptHandle, Tab(5); "Interest:     "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 9
            If LineCnt >= MaxLines - 3 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 14
            Print #RptHandle, "  Adj Pay Dn: -Credit"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If fpcmbTransType.Text = " 0) All" Then GoTo NextOne
            Print #RptHandle, Tab(5); "Personal Prop:"; Tab(55); Using$("$###,###,##0.00", TotByYrAndPers(y))
            Print #RptHandle, Tab(5); "Machine Tools:"; Tab(55); Using$("$###,###,##0.00", TotByYrAndMT(y))
            Print #RptHandle, Tab(5); "Merchant Cap: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMC(y))
            Print #RptHandle, Tab(5); "Farm Equip:   "; Tab(55); Using$("$###,###,##0.00", TotByYrAndFE(y))
            Print #RptHandle, Tab(5); "Mobile Homes: "; Tab(55); Using$("$###,###,##0.00", TotByYrAndMH(y))
            Print #RptHandle, Tab(5); "Interest:     "; Tab(55); Using$("$###,###,##0.00", TotByYrAndInt(y))
            Print #RptHandle, Tab(5); "Penalty:      "; Tab(55); Using$("$###,###,##0.00", TotByYrAndPen(y))
            LineCnt = LineCnt + 9
            If LineCnt >= MaxLines - 3 Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt1(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt2(y))
              LineCnt = LineCnt + 1
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", TotByYrAndOpt3(y))
              LineCnt = LineCnt + 1
            End If
          Case 15
            Print #RptHandle, "  Adjust Prepay Down"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
          Case 16
            Print #RptHandle, "  Refund Prepay"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
              Print #RptHandle, "Year: " + Using$("###0", ThEYear(y))
            End If
          Case 17
            Print #RptHandle, "  PPTRA Removal"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
            End If
          Case 18
            Print #RptHandle, "  Unknown"; Tab(35); Using$("##,##0", CntByYrAndType(x, y)); Tab(55); Using$("$###,###,##0.00", TotByYrAndType(x, y)) 'dlm; TheYear(y); dlm;
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintSortHeader
            End If
        End Select
      End If
NextOne:
    Next x
    Print #RptHandle, "  Total Year Count: "; Tab(30); Using$("###,###,##0", ThisYrCnt)
    Print #RptHandle, String$(94, "-")
    Print #RptHandle,
    LineCnt = LineCnt + 3
  Next y
  
  Return

PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(30); "Tax Transactions Journal"
  Print #RptHandle, Town; Tab(75); "Page #: " + CStr(Page)
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle, "Tax Type: " + ThisType
  Print #RptHandle, "Transaction Type: " + ThisBillType
  Print #RptHandle, "Date Range: " + fptxtBegDate.Text + " to " + fptxtEndDate.Text
  Print #RptHandle,
  Print #RptHandle, "Trans Date"; Tab(12); "Description"; Tab(35); "Tax Year"; Tab(44); "Overpay Amt"; Tab(57); "Trans Amt"; Tab(78); "Trans Type" 'Tab(67); "Belongs To"; Tab(78); "Trans Type"
  Print #RptHandle, String(94, "-")
  LineCnt = 9
  
  Return
  
PrintCustHeader:
  If LineCnt <> 9 Then
    Print #RptHandle, String(94, "-")
    LineCnt = LineCnt + 1
  End If
  If LineCnt >= MaxLines - 5 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  Print #RptHandle, "Cust Num: " + Using$("#######0", TaxCust.Acct); Tab(21); "Customer Name: "; Tab(37); QPTrim$(TaxCust.CustName); Tab(80); "Active: "; Tab(89); TaxCust.Active
  If UseOpt = "Y" Then
    Print #RptHandle, Tab(21); ThisOpt + ":" + QPTrim$(TaxCust.OptSrchDesc)
    LineCnt = LineCnt + 1
  End If
  Print #RptHandle, "Total Outstanding Customer Balance: " + Using$("$##,##0.00", CustBal)
  Print #RptHandle, Tab(15); "Revenue"; Tab(44); "Amount"; Tab(59); "Amount Paid"; Tab(83); "Balance"
  Print #RptHandle, String(94, ".")
  LineCnt = LineCnt + 4
  
  Return
  
PrintSortHeader:
  Page = Page + 1
  Print #RptHandle, Tab(25); "Tax Transactions Journal Summary"
  Print #RptHandle, Town; Tab(75); "Page #: " + CStr(Page)
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle, "Tax Type: " + ThisType
  Print #RptHandle, "Transaction Type: " + ThisBillType
  Print #RptHandle, "Date Range: " + fptxtBegDate.Text + " to " + fptxtEndDate.Text
  Print #RptHandle, "Note: Adjustment transaction amounts are reflected in revenues and "
  Print #RptHandle, "      adjustment transaction totals exclusively. "
  Print #RptHandle, "Description"; Tab(35); "Trans Cnt"; Tab(64); "Amount"
  Print #RptHandle, String$(94, "-")
  LineCnt = 10
  
  Return
  
PrintTotalsHeader:
  Page = Page + 1
  Print #RptHandle, Tab(30); "Tax Transactions Journal Summary"
  Print #RptHandle, "Grand Totals"
  Print #RptHandle, Town; Tab(75); "Page #: " + CStr(Page)
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle, "Tax Type: " + ThisType
  Print #RptHandle, "Transaction Type: " + ThisBillType
  Print #RptHandle, "Date Range: " + fptxtBegDate.Text + " to " + fptxtEndDate.Text
  Print #RptHandle, "Note: Adjustment transaction amounts are reflected in revenues and "
  Print #RptHandle, "      adjustment transaction totals exclusively. "
  Print #RptHandle, Tab(35); "Trans Cnt"; Tab(64); "Amount"
  Print #RptHandle, String$(94, "-")
  LineCnt = 10

  Return
  
PrintTotals:
  GCntTot = 0
  GAmtTot = 0
  If LineCnt > MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintTotalsHeader
  Else
    Print #RptHandle,
    Print #RptHandle, "Grand Totals"
    Print #RptHandle, Tab(35); "Trans Cnt"; Tab(64); "Amount"
    Print #RptHandle, String$(94, "-")
    LineCnt = LineCnt + 4
  End If
  If fpcmbTransType.Text = " 0) All" Then GoTo All4
  For x = 1 To YrCnt
    GPersTot = GPersTot + TotByYrAndPers(x)
    GMTTot = GMTTot + TotByYrAndMT(x)
    GMCTot = GMCTot + TotByYrAndMC(x)
    GFETot = GFETot + TotByYrAndFE(x)
    GMHTot = GMHTot + TotByYrAndMH(x)
    GIntTot = GIntTot + TotByYrAndInt(x)
    GAdvTot = GAdvTot + TotByYrAndAdv(x)
    GLateListTot = GLateListTot + TotByYrAndLateList(x)
    GPenTot = GPenTot + TotByYrAndPen(x)
    GOpt1Tot = GOpt1Tot + TotByYrAndOpt1(x)
    GOpt2Tot = GOpt2Tot + TotByYrAndOpt2(x)
    GOpt3Tot = GOpt3Tot + TotByYrAndOpt3(x)
  Next x
All4:
  For x = 1 To 18
    GCntTot = GCntTot + TotCntByType(x)
    GAmtTot = GAmtTot + TotByType(x)
    If TotByType(x) > 0 Then
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintTotalsHeader
      End If
      Select Case x
        Case 1
          Print #RptHandle, "Billing";
        Case 2
          Print #RptHandle, "Payment";
        Case 3
          Print #RptHandle, "Release";
        Case 4
          Print #RptHandle, "Interest";
        Case 5
          Print #RptHandle, "Penalty";
        Case 6
          Print #RptHandle, "Advertising";
        Case 7
          Print #RptHandle, "Adjust Pay Down";
        Case 8
          Print #RptHandle, "Credit at Billing";
        Case 9
          Print #RptHandle, "Adjust Bill Down";
        Case 10
          Print #RptHandle, "Adjust Bill Up";
        Case 11
          If fpcmbTransType.Text = " 0) All" Then 'added Bill Pay/OverPay on 7/7/06
            Print #RptHandle, "Bill Pay/OverPay";
          Else
            Print #RptHandle, "Bill OverPay";
          End If
        Case 12
          Print #RptHandle, "OverPayment";
        Case 13
          Print #RptHandle, "Adjust Bill Up Affecting Credit Balance";
        Case 14
          Print #RptHandle, "Adjust Pay Dwn Affecting Credit Balance";
        Case 15
          Print #RptHandle, "Adjust Prepay Down";
        Case 16
          Print #RptHandle, "Refund Prepay";
        Case 17
          Print #RptHandle, "PPTRA Removal";
        Case 18
          Print #RptHandle, "Unknown";
      End Select
      Dim Case11Tot As Double 'added 7/6/06
      If x <> 11 Then
        Print #RptHandle, Tab(35); Using$("##,##0", TotCntByType(x)); Tab(55); Using$("$###,###,##0.00", TotByType(x))
      Else
        Case11Tot = OldRound(TotByType(11) - (GPersTot + GMTTot + GMCTot + GFETot + GMHTot + GOpt1Tot + GOpt2Tot + GOpt3Tot))
        Case11Tot = OldRound(Case11Tot - (GIntTot + GAdvTot + GPenTot))
        Print #RptHandle, Tab(35); Using$("##,##0", TotCntByType(11)); Tab(55); Using$("$###,###,##0.00", Case11Tot)
      End If
      LineCnt = LineCnt + 1
      If fpcmbTransType.Text = " 0) All" Then GoTo All
      Select Case x
        Case 1
          If LineCnt >= MaxLines - 8 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Personal Prop: "; Tab(55); Using$("$###,###,##0.00", GPersTot)
          Print #RptHandle, Tab(5); "Machine Tools: "; Tab(55); Using$("$###,###,##0.00", GMTTot)
          Print #RptHandle, Tab(5); "Merchant Cap:  "; Tab(55); Using$("$###,###,##0.00", GMCTot)
          Print #RptHandle, Tab(5); "Farm Equip:    "; Tab(55); Using$("$###,###,##0.00", GFETot)
          Print #RptHandle, Tab(5); "Mobile Homes:  "; Tab(55); Using$("$###,###,##0.00", GMHTot)
          '7/11/06 added back int, adv and pen to accommodate manual bills
          Print #RptHandle, Tab(5); "Interest:      "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising:   "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Penalty:       "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 7
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
        Case 2
          If LineCnt >= MaxLines - 11 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Personal Prop: "; Tab(55); Using$("$###,###,##0.00", GPersTot)
          Print #RptHandle, Tab(5); "Machine Tools: "; Tab(55); Using$("$###,###,##0.00", GMTTot)
          Print #RptHandle, Tab(5); "Merchant Cap:  "; Tab(55); Using$("$###,###,##0.00", GMCTot)
          Print #RptHandle, Tab(5); "Farm Equip:    "; Tab(55); Using$("$###,###,##0.00", GFETot)
          Print #RptHandle, Tab(5); "Mobile Homes:  "; Tab(55); Using$("$###,###,##0.00", GMHTot)
          Print #RptHandle, Tab(5); "Interest:      "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising:   "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Penalty:       "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 9
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
        Case 3
          If LineCnt >= MaxLines - 11 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Personal Prop: "; Tab(55); Using$("$###,###,##0.00", GPersTot)
          Print #RptHandle, Tab(5); "Machine Tools: "; Tab(55); Using$("$###,###,##0.00", GMTTot)
          Print #RptHandle, Tab(5); "Merchant Cap:  "; Tab(55); Using$("$###,###,##0.00", GMCTot)
          Print #RptHandle, Tab(5); "Farm Equip:    "; Tab(55); Using$("$###,###,##0.00", GFETot)
          Print #RptHandle, Tab(5); "Mobile Homes:  "; Tab(55); Using$("$###,###,##0.00", GMHTot)
          Print #RptHandle, Tab(5); "Interest:      "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising:   "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Penalty:       "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 9
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
        Case 4
          GoTo All
        Case 5
          GoTo All
        Case 6
          GoTo All
        Case 7
          If LineCnt >= MaxLines - 11 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Personal Prop: "; Tab(55); Using$("$###,###,##0.00", GPersTot)
          Print #RptHandle, Tab(5); "Machine Tools: "; Tab(55); Using$("$###,###,##0.00", GMTTot)
          Print #RptHandle, Tab(5); "Merchant Cap:  "; Tab(55); Using$("$###,###,##0.00", GMCTot)
          Print #RptHandle, Tab(5); "Farm Equip:    "; Tab(55); Using$("$###,###,##0.00", GFETot)
          Print #RptHandle, Tab(5); "Mobile Homes:  "; Tab(55); Using$("$###,###,##0.00", GMHTot)
          Print #RptHandle, Tab(5); "Interest:      "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising:   "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Penalty:       "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 9
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
        Case 8
          GoTo All
        Case 9
          If LineCnt >= MaxLines - 11 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Personal Prop: "; Tab(55); Using$("$###,###,##0.00", GPersTot)
          Print #RptHandle, Tab(5); "Machine Tools: "; Tab(55); Using$("$###,###,##0.00", GMTTot)
          Print #RptHandle, Tab(5); "Merchant Cap:  "; Tab(55); Using$("$###,###,##0.00", GMCTot)
          Print #RptHandle, Tab(5); "Farm Equip:    "; Tab(55); Using$("$###,###,##0.00", GFETot)
          Print #RptHandle, Tab(5); "Mobile Homes:  "; Tab(55); Using$("$###,###,##0.00", GMHTot)
          Print #RptHandle, Tab(5); "Interest:      "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising:   "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Penalty:       "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 9
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
        Case 10
          If LineCnt >= MaxLines - 11 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Personal Prop: "; Tab(55); Using$("$###,###,##0.00", GPersTot)
          Print #RptHandle, Tab(5); "Machine Tools: "; Tab(55); Using$("$###,###,##0.00", GMTTot)
          Print #RptHandle, Tab(5); "Merchant Cap:  "; Tab(55); Using$("$###,###,##0.00", GMCTot)
          Print #RptHandle, Tab(5); "Farm Equip:    "; Tab(55); Using$("$###,###,##0.00", GFETot)
          Print #RptHandle, Tab(5); "Mobile Homes:  "; Tab(55); Using$("$###,###,##0.00", GMHTot)
          Print #RptHandle, Tab(5); "Interest:      "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising:   "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Penalty:       "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 9
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
       Case 11 '7/6/06 commented out GoTo All and added detail
'         GoTo All
          If LineCnt >= MaxLines - 11 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Personal Prop: "; Tab(55); Using$("$###,###,##0.00", GPersTot)
          Print #RptHandle, Tab(5); "Machine Tools: "; Tab(55); Using$("$###,###,##0.00", GMTTot)
          Print #RptHandle, Tab(5); "Merchant Cap:  "; Tab(55); Using$("$###,###,##0.00", GMCTot)
          Print #RptHandle, Tab(5); "Farm Equip:    "; Tab(55); Using$("$###,###,##0.00", GFETot)
          Print #RptHandle, Tab(5); "Mobile Homes:  "; Tab(55); Using$("$###,###,##0.00", GMHTot)
          Print #RptHandle, Tab(5); "Interest:      "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising:   "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Penalty:       "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 9
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
       Case 12
         GoTo All
       Case 13
          If LineCnt >= MaxLines - 11 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Personal Prop: "; Tab(55); Using$("$###,###,##0.00", GPersTot)
          Print #RptHandle, Tab(5); "Machine Tools: "; Tab(55); Using$("$###,###,##0.00", GMTTot)
          Print #RptHandle, Tab(5); "Merchant Cap:  "; Tab(55); Using$("$###,###,##0.00", GMCTot)
          Print #RptHandle, Tab(5); "Farm Equip:    "; Tab(55); Using$("$###,###,##0.00", GFETot)
          Print #RptHandle, Tab(5); "Mobile Homes:  "; Tab(55); Using$("$###,###,##0.00", GMHTot)
          Print #RptHandle, Tab(5); "Interest:      "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising:   "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Penalty:       "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 9
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
        Case 14
          If LineCnt >= MaxLines - 11 Then
            Print #RptHandle, FF$
            GoSub PrintTotalsHeader
          End If
          Print #RptHandle, Tab(5); "Personal Prop: "; Tab(55); Using$("$###,###,##0.00", GPersTot)
          Print #RptHandle, Tab(5); "Machine Tools: "; Tab(55); Using$("$###,###,##0.00", GMTTot)
          Print #RptHandle, Tab(5); "Merchant Cap:  "; Tab(55); Using$("$###,###,##0.00", GMCTot)
          Print #RptHandle, Tab(5); "Farm Equip:    "; Tab(55); Using$("$###,###,##0.00", GFETot)
          Print #RptHandle, Tab(5); "Mobile Homes:  "; Tab(55); Using$("$###,###,##0.00", GMHTot)
          Print #RptHandle, Tab(5); "Interest:      "; Tab(55); Using$("$###,###,##0.00", GIntTot)
          Print #RptHandle, Tab(5); "Advertising:   "; Tab(55); Using$("$###,###,##0.00", GAdvTot)
          Print #RptHandle, Tab(5); "Penalty:       "; Tab(55); Using$("$###,###,##0.00", GPenTot)
          LineCnt = LineCnt + 9
          If QPTrim$(Opt1Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt1Desc); Tab(55); Using$("$###,###,##0.00", GOpt1Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt2Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt2Desc); Tab(55); Using$("$###,###,##0.00", GOpt2Tot)
          End If
          LineCnt = LineCnt + 1
          If QPTrim$(Opt3Desc) <> "" Then
            Print #RptHandle, Tab(5); QPTrim$(Opt3Desc); Tab(55); Using$("$###,###,##0.00", GOpt3Tot)
          End If
          LineCnt = LineCnt + 1
        Case 15
          GoTo All
        Case 16
          GoTo All
        Case 17
          GoTo All
        Case Else
          GoTo All
      End Select
    End If
All:
  Next x
  Print #RptHandle, String(94, "-")
  Print #RptHandle, "Grand Totals: "; Tab(35); Using$("##,##0", GCntTot); Tab(55); Using$("$###,###,##0.00", GAmtTot)
  
  Return
  
PrintAll:
  ThatName = ""
  For x = 1 To NumOfTCRecs
    PrePayDone = False
    If IdxFlag = False Then
      Get TCHandle, x, TaxCust
      CustBal = GetCustBalance(x, -1)
    Else
      Get TCHandle, IdxArray(x), TaxCust
      CustBal = GetCustBalance(IdxArray(x), -1)
    End If
    If TaxCust.Active = "N" And InactiveFlag = False Then
      GoTo SkipThisOne
    End If
    ThisName = QPTrim$(TaxCust.CustName)
    ReDim BillRec(1 To 1) As Long
    BillCnt = 0
    BadCnt = 0
    GoSub PrintPrepay
    ThisRec = TaxCust.LastTrans
    Do While ThisRec > 0
      Get TTHandle, ThisRec, TaxTrans
      If QFlag = True Then
        If TaxTrans.TransDate < BegDate Then
          BadCnt = BadCnt + 1
          If BadCnt > 3 Then Exit Do
        End If
      End If
      If TransDesc <> "" Then
        If InStr(1, TaxTrans.Description, TransDesc) = 0 Then GoTo Nope
      End If
      If TaxTrans.TransDate < BegDate Or TaxTrans.TransDate > EndDate Then GoTo Nope
      If TaxTrans.BillType <> "R" And TaxTrans.BillType <> "P" Then TaxTrans.BillType = ""
      If TaxTrans.BillType <> Mid(fpcmbTaxType.Text, 1, 1) And QPTrim$(TaxTrans.BillType) <> "" Then GoTo Nope
      If TaxTrans.TranType = 1 Then
        BillCnt = BillCnt + 1
        ReDim Preserve BillRec(1 To BillCnt) As Long
        BillRec(BillCnt) = ThisRec
      End If
Nope:
      ThisRec = TaxTrans.LastTrans
    Loop

    For z = 1 To BillCnt
      If ThisName <> ThatName Then
        GoSub PrintCustHeader
        ThatName = ThisName
      End If
      Get TTHandle, BillRec(z), TaxTrans
      If YrCnt = 0 Then
        YrCnt = YrCnt + 1
        ThisYear = YrCnt
        ReDim Preserve ThEYear(1 To YrCnt) As Integer
        ThEYear(YrCnt) = TaxTrans.TaxYear
        ReDim Preserve TotByYrAndType(1 To 18, 1 To YrCnt) As Double
        ReDim Preserve CntByYrAndType(1 To 18, 1 To YrCnt) As Integer
        For y = 1 To 18
          TotByYrAndType(y, YrCnt) = 0
          CntByYrAndType(y, YrCnt) = 0
        Next y
      Else
        For y = 1 To YrCnt
          If TaxTrans.TaxYear = ThEYear(y) Then
            ThisYear = y
            Exit For
          End If
        Next y
        If y > YrCnt Then
          YrCnt = YrCnt + 1
          ThisYear = YrCnt
          ReDim Preserve ThEYear(1 To YrCnt) As Integer
          ThEYear(YrCnt) = TaxTrans.TaxYear
          ReDim Preserve TotByYrAndType(1 To 18, 1 To YrCnt) As Double
          ReDim Preserve CntByYrAndType(1 To 18, 1 To YrCnt) As Integer
          For y = 1 To 18
            TotByYrAndType(y, YrCnt) = 0
            CntByYrAndType(y, YrCnt) = 0
          Next y
        End If
      End If
      ThisTransType = "Billing"
      TCnt = TCnt + 1
      If TaxTrans.DiscAmt > 0 Then 'added 1/16/07
        GoSub ApplyDisc
      End If
      TotAmt = OldRound(TotAmt + TaxTrans.Amount)
      PersDif = OldRound(TaxTrans.Revenue.Principle1 - TaxTrans.Revenue.Principle1Pd)
      MTDif = OldRound(TaxTrans.Revenue.Principle2 - TaxTrans.Revenue.Principle2Pd)
      MCDif = OldRound(TaxTrans.Revenue.Principle3 - TaxTrans.Revenue.Principle3Pd)
      FEDif = OldRound(TaxTrans.Revenue.Principle4 - TaxTrans.Revenue.Principle4Pd)
      MHDif = OldRound(TaxTrans.Revenue.Principle5 - TaxTrans.Revenue.Principle5Pd)
      IntDif = OldRound(TaxTrans.Revenue.Interest - TaxTrans.Revenue.InterestPd)
      PenDif = OldRound(TaxTrans.Revenue.Penalty - TaxTrans.Revenue.PenaltyPd)
      Opt1Dif = OldRound(TaxTrans.Revenue.RevOpt1 - TaxTrans.Revenue.RevOpt1Pd)
      Opt2Dif = OldRound(TaxTrans.Revenue.RevOpt2 - TaxTrans.Revenue.RevOpt2Pd)
      Opt3Dif = OldRound(TaxTrans.Revenue.RevOpt3 - TaxTrans.Revenue.RevOpt3Pd)
      BillBal = OldRound(PersDif + IntDif + MTDif + MCDif + FEDif + MHDif + PenDif + Opt1Dif + Opt2Dif + Opt3Dif)
      ThisBillNum = ParseBillNum(TaxTrans.Description)
      
      Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); QPTrim$(TaxTrans.Description);
      Print #RptHandle, Tab(37); Using$("###0", TaxTrans.TaxYear); Tab(45); Using$("$##,##0.00", TaxTrans.Revenue.PrePaidAmt);
      Print #RptHandle, Tab(56); Using$("$##,##0.00", TaxTrans.Amount);
      Print #RptHandle, Tab(79); ThisTransType
      LineCnt = LineCnt + 1
      If LineCnt >= MaxLines - 4 Then
        Print #RptHandle, FF$
        GoSub PrintHeader
        GoSub PrintCustHeader
      End If
      Print #RptHandle, Tab(3); "Oper #: " + CStr(TaxTrans.OperNum);
      Print #RptHandle, Tab(15); "Principle         "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle1); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle1Pd); Tab(80); Using$("$##,##0.00", PersDif)
      Print #RptHandle, Tab(15); "Machine Tools     "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle2); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle2Pd); Tab(80); Using$("$##,##0.00", MTDif)
      Print #RptHandle, Tab(15); "Merchant Capital  "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle3); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle3Pd); Tab(80); Using$("$##,##0.00", MCDif)
      Print #RptHandle, Tab(15); "Farm Equipment    "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle4); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle4Pd); Tab(80); Using$("$##,##0.00", FEDif)
      Print #RptHandle, Tab(15); "Merchant Capital  "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle5); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle5Pd); Tab(80); Using$("$##,##0.00", MHDif)
      Print #RptHandle, Tab(15); "Interest          "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Interest); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.InterestPd); Tab(80); Using$("$##,##0.00", IntDif)
      Print #RptHandle, Tab(15); "Penalty           "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Penalty); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.PenaltyPd); Tab(80); Using$("$##,##0.00", PenDif)
      LineCnt = LineCnt + 7
      If LineCnt >= MaxLines - 3 Then
        Print #RptHandle, FF$
        GoSub PrintHeader
        GoSub PrintCustHeader
      End If
      If Len(Opt1Desc) > 0 Then
        Print #RptHandle, Tab(15); Opt1Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1Pd); Tab(80); Using$("$##,##0.00", Opt1Dif)
        LineCnt = LineCnt + 1
      End If
      If Len(Opt2Desc) > 0 Then
        Print #RptHandle, Tab(15); Opt2Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2Pd); Tab(80); Using$("$##,##0.00", Opt2Dif)
        LineCnt = LineCnt + 1
      End If
      If Len(Opt3Desc) > 0 Then
        Print #RptHandle, Tab(15); Opt3Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3Pd); Tab(80); Using$("$##,##0.00", Opt3Dif)
        LineCnt = LineCnt + 1
      End If
      Print #RptHandle, Tab(15); "Bill Balance:"; Tab(80); Using$("$##,##0.00", BillBal)
      
      ThisRec = TaxCust.LastTrans
      TotByYrAndType(1, ThisYear) = OldRound(TotByYrAndType(1, ThisYear) + TaxTrans.Amount)
      CntByYrAndType(1, ThisYear) = OldRound(CntByYrAndType(1, ThisYear) + 1)
      TotByType(1) = OldRound(TotByType(1) + TaxTrans.Amount)
      TotCntByType(1) = OldRound(TotCntByType(1) + 1)

      Do While ThisRec > 0
        Get TTHandle, ThisRec, TaxTrans
        If TaxTrans.TranType = 22 Then 'Prepay transactions can only be printed one time, not
        'for every iteration
          If PrePayDone = False Then
            GoTo PrepaySkip
          Else
            GoTo Nope2
          End If
        End If
        If TaxTrans.TranType = 11 Or TaxTrans.TranType = 12 Then GoTo Nope2
        If TaxTrans.BelongTo <> BillRec(z) Then GoTo Nope2
PrepaySkip:
        Select Case TaxTrans.TranType
          Case 2
            ThisTransType = "Payment"
             If TaxTrans.DiscAmt > 0 Then 'added 1/16/07
               GoSub ApplyDisc
             End If
            TotByYrAndType(2, ThisYear) = OldRound(TotByYrAndType(2, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(2, ThisYear) = OldRound(CntByYrAndType(2, ThisYear) + 1)
            TotByType(2) = OldRound(TotByType(2) + TaxTrans.Amount)
            TotCntByType(2) = OldRound(TotCntByType(2) + 1)
          Case 3
            ThisTransType = "Release"
            TotByYrAndType(3, ThisYear) = OldRound(TotByYrAndType(3, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(3, ThisYear) = OldRound(CntByYrAndType(3, ThisYear) + 1)
            TotByType(3) = OldRound(TotByType(3) + TaxTrans.Amount)
            TotCntByType(3) = OldRound(TotCntByType(3) + 1)
          Case 4
            ThisTransType = "Interest"
            TotByYrAndType(4, ThisYear) = OldRound(TotByYrAndType(4, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(4, ThisYear) = OldRound(CntByYrAndType(4, ThisYear) + 1)
            TotByType(4) = OldRound(TotByType(4) + TaxTrans.Amount)
            TotCntByType(4) = OldRound(TotCntByType(4) + 1)
          Case 5
            ThisTransType = "Penalty"
            TotByYrAndType(5, ThisYear) = OldRound(TotByYrAndType(5, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(5, ThisYear) = OldRound(CntByYrAndType(5, ThisYear) + 1)
            TotByType(5) = OldRound(TotByType(5) + TaxTrans.Amount)
            TotCntByType(5) = OldRound(TotCntByType(5) + 1)
          Case 6
            ThisTransType = "Advertising Charge"
            TotByYrAndType(6, ThisYear) = OldRound(TotByYrAndType(6, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(6, ThisYear) = OldRound(CntByYrAndType(6, ThisYear) + 1)
            TotByType(6) = OldRound(TotByType(6) + TaxTrans.Amount)
            TotCntByType(6) = OldRound(TotCntByType(6) + 1)
          Case 7
            ThisTransType = "Adjust Pay Down"
            TotByYrAndType(7, ThisYear) = OldRound(TotByYrAndType(7, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(7, ThisYear) = OldRound(CntByYrAndType(7, ThisYear) + 1)
            TotByType(7) = OldRound(TotByType(7) + TaxTrans.Amount)
            TotCntByType(7) = OldRound(TotCntByType(7) + 1)
          Case 9
            ThisTransType = "Credit Applied at Billing"
            TotByYrAndType(8, ThisYear) = OldRound(TotByYrAndType(8, ThisYear) + TaxTrans.Revenue.PrePaidUsed)
            CntByYrAndType(8, ThisYear) = OldRound(CntByYrAndType(8, ThisYear) + 1)
            TotByType(8) = OldRound(TotByType(8) + TaxTrans.Revenue.PrePaidUsed)
            TotCntByType(8) = OldRound(TotCntByType(8) + 1)
          Case 13
            ThisTransType = "Adjust Bill Down"
            TotByYrAndType(9, ThisYear) = OldRound(TotByYrAndType(9, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(9, ThisYear) = OldRound(CntByYrAndType(9, ThisYear) + 1)
            TotByType(9) = OldRound(TotByType(9) + TaxTrans.Amount)
            TotCntByType(9) = OldRound(TotCntByType(9) + 1)
          Case 14
            ThisTransType = "Adjust Bill Up"
            TotByYrAndType(10, ThisYear) = OldRound(TotByYrAndType(10, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(10, ThisYear) = OldRound(CntByYrAndType(10, ThisYear) + 1)
            TotByType(10) = OldRound(TotByType(10) + TaxTrans.Amount)
            TotCntByType(10) = OldRound(TotCntByType(10) + 1)
          Case 21
            ThisTransType = "Billpay/Overpay"
            If TaxTrans.DiscAmt > 0 Then 'added 1/16/07
              GoSub ApplyDisc
            End If
            TotByYrAndType(11, ThisYear) = OldRound(TotByYrAndType(11, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(11, ThisYear) = OldRound(CntByYrAndType(11, ThisYear) + 1)
            TotByType(11) = OldRound(TotByType(11) + TaxTrans.Amount)
            TotCntByType(11) = OldRound(TotCntByType(11) + 1)
          Case 22
            ThisTransType = "Overpayment"
            TotByYrAndType(12, ThisYear) = OldRound(TotByYrAndType(12, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(12, ThisYear) = OldRound(CntByYrAndType(12, ThisYear) + 1)
            TotByType(12) = OldRound(TotByType(12) + TaxTrans.Amount)
            TotCntByType(12) = OldRound(TotCntByType(12) + 1)
          Case 24
            ThisTransType = "Adj Bill Up: -Credit Bal"
            TotByYrAndType(13, ThisYear) = OldRound(TotByYrAndType(13, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(13, ThisYear) = OldRound(CntByYrAndType(13, ThisYear) + 1)
            TotByType(13) = OldRound(TotByType(13) + TaxTrans.Amount)
            TotCntByType(13) = OldRound(TotCntByType(13) + 1)
          Case 10
            ThisTransType = "Adj Pay Dwn: -Credit Bal"
            TotByYrAndType(14, ThisYear) = OldRound(TotByYrAndType(14, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(14, ThisYear) = OldRound(CntByYrAndType(14, ThisYear) + 1)
            TotByType(14) = OldRound(TotByType(14) + TaxTrans.Amount)
            TotCntByType(14) = OldRound(TotCntByType(14) + 1)
          Case 30
            ThisTransType = "PPTRA Removal"
            TotByYrAndType(15, ThisYear) = OldRound(TotByYrAndType(15, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(15, ThisYear) = OldRound(CntByYrAndType(15, ThisYear) + 1)
            TotByType(15) = OldRound(TotByType(15) + TaxTrans.Amount)
            TotCntByType(15) = OldRound(TotCntByType(15) + 1)
'          Case 11
'            ThisTransType = "Adjust Prepay Down"
'            TotByYrAndType(15, ThisYear) = OldRound(TotByYrAndType(15, ThisYear) + TaxTrans.Amount)
'            CntByYrAndType(15, ThisYear) = OldRound(CntByYrAndType(15, ThisYear) + 1)
'          Case 12
'            ThisTransType = "Refund Prepay"
'            TotByYrAndType(16, ThisYear) = OldRound(TotByYrAndType(16, ThisYear) + TaxTrans.Amount)
'            CntByYrAndType(16, ThisYear) = OldRound(CntByYrAndType(16, ThisYear) + 1)
          Case Else
            ThisTransType = "Unknown"
            TotByYrAndType(17, ThisYear) = OldRound(TotByYrAndType(17, ThisYear) + TaxTrans.Amount)
            CntByYrAndType(17, ThisYear) = OldRound(CntByYrAndType(17, ThisYear) + 1)
            TotByType(17) = OldRound(TotByType(17) + TaxTrans.Amount)
            TotCntByType(17) = OldRound(TotCntByType(17) + 1)
        End Select
        TCnt = TCnt + 1
        If LineCnt >= MaxLines - 2 Then
          Print #RptHandle, FF$
          GoSub PrintHeader
          GoSub PrintCustHeader
        End If
        If TaxTrans.TranType <> 11 Then
          Print #RptHandle, Tab(15); String(80, "^")
        Else
          Print #RptHandle, Tab(15); String(80, "-")
        End If
        Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); QPTrim$(TaxTrans.Description);
        Print #RptHandle, Tab(37); Using$("###0", TaxTrans.TaxYear); Tab(45); Using$("$##,##0.00", TaxTrans.Revenue.PrePaidAmt);
        If TaxTrans.TranType <> 9 Then
          Print #RptHandle, Tab(56); Using$("$##,##0.00", TaxTrans.Amount); Tab(79); ThisTransType
        Else
          Print #RptHandle, Tab(56); Using$("$##,##0.00", TaxTrans.Revenue.PrePaidUsed); Tab(79); ThisTransType
        End If
        
        LineCnt = LineCnt + 2
        If LineCnt >= MaxLines - 4 Then
          Print #RptHandle, FF$
          GoSub PrintHeader
          GoSub PrintCustHeader
        End If
        Print #RptHandle, Tab(3); "Oper #: " + CStr(TaxTrans.OperNum);
        Print #RptHandle, Tab(15); "Personal          "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle1); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle1Pd); Tab(75); "Belongs to Bill#"
        If TaxTrans.TranType = 11 Then
          Print #RptHandle, Tab(15); "Machine Tools     "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle2); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle2Pd); Tab(75); "NA"
        Else
          Print #RptHandle, Tab(15); "Machine Tools     "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle2); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle2Pd); Tab(75); QPTrim$(ThisBillNum)
        End If
        Print #RptHandle, Tab(15); "Merchant Capital  "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle3); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle3Pd)
        Print #RptHandle, Tab(15); "Farm Equipment    "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle4); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle4Pd)
        Print #RptHandle, Tab(15); "Mobile Homes      "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle5); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle5Pd)
        Print #RptHandle, Tab(15); "Interest          "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Interest); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.InterestPd)
        Print #RptHandle, Tab(15); "Penalty           "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Penalty); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.PenaltyPd)
        LineCnt = LineCnt + 7
        If LineCnt >= MaxLines - 3 Then
          Print #RptHandle, FF$
          GoSub PrintHeader
          GoSub PrintCustHeader
        End If
        If Len(Opt1Desc) > 0 Then
          Print #RptHandle, Tab(15); Opt1Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1Pd)
          LineCnt = LineCnt + 1
        End If
        If Len(Opt2Desc) > 0 Then
          Print #RptHandle, Tab(15); Opt2Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2Pd)
          LineCnt = LineCnt + 1
        End If
        If Len(Opt3Desc) > 0 Then
          Print #RptHandle, Tab(15); Opt3Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3Pd)
          LineCnt = LineCnt + 1
        End If
Nope2:
        ThisRec = TaxTrans.LastTrans
      Loop
      PrePayDone = True 'added 7/10/06
    Next z
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
SkipThisOne:
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  If TCnt = 0 Then
    Call TaxMsg(900, "No transactions were found that fit the parameters entered.")
    Close
    Exit Sub
  End If
  
  If YrCnt > 0 Then
    GoSub SortIt
    GoSub PrintTotals
  End If
  Print #RptHandle, FF$
  Close
  ViewPrint RptFile, "Tax Transactions Report", True
  Exit Sub
  
PrintPrepay:
  ThisRec = TaxCust.LastTrans
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
    If TaxTrans.TranType <> 11 And TaxTrans.TranType <> 12 Then GoTo NotThisTrans
    If TaxTrans.TransDate < BegDate Or TaxTrans.TransDate > EndDate Then GoTo NotThisTrans
    If ThisName <> ThatName Then
      GoSub PrintCustHeader
      ThatName = ThisName
    End If
    If TaxTrans.TranType = 11 Then
      ThisTransType = "Adj Prepay Down"
    ElseIf TaxTrans.TranType = 12 Then
      ThisTransType = "Refund Prepay"
    End If
    If YrCnt = 0 Then
      YrCnt = YrCnt + 1
      ThisYear = YrCnt
      ReDim Preserve ThEYear(1 To YrCnt) As Integer
      ThEYear(YrCnt) = TaxTrans.TaxYear
      ReDim Preserve TotByYrAndType(1 To 18, 1 To YrCnt) As Double
      ReDim Preserve CntByYrAndType(1 To 18, 1 To YrCnt) As Integer
      For y = 1 To 18
        TotByYrAndType(y, YrCnt) = 0
        CntByYrAndType(y, YrCnt) = 0
      Next y
    Else
      For y = 1 To YrCnt
        If TaxTrans.TaxYear = ThEYear(y) Then
          ThisYear = y
          Exit For
          End If
      Next y
      If y > YrCnt Then
        YrCnt = YrCnt + 1
        ThisYear = YrCnt
        ReDim Preserve ThEYear(1 To YrCnt) As Integer
        ThEYear(YrCnt) = TaxTrans.TaxYear
        ReDim Preserve TotByYrAndType(1 To 18, 1 To YrCnt) As Double
        ReDim Preserve CntByYrAndType(1 To 18, 1 To YrCnt) As Integer
        For y = 1 To 18
          TotByYrAndType(y, YrCnt) = 0
          CntByYrAndType(y, YrCnt) = 0
        Next y
      End If
    End If
    If TaxTrans.TranType = 11 Then
      TotByYrAndType(15, ThisYear) = OldRound(TotByYrAndType(15, ThisYear) + TaxTrans.Amount)
      CntByYrAndType(15, ThisYear) = OldRound(CntByYrAndType(15, ThisYear) + 1)
    ElseIf TaxTrans.TranType = 12 Then
      TotByYrAndType(16, ThisYear) = OldRound(TotByYrAndType(16, ThisYear) + TaxTrans.Amount)
      CntByYrAndType(16, ThisYear) = OldRound(CntByYrAndType(16, ThisYear) + 1)
    End If
    TCnt = TCnt + 1
    TotAmt = OldRound(TotAmt + TaxTrans.Amount)
    PersDif = 0
    MTDif = 0
    MCDif = 0
    FEDif = 0
    MHDif = 0
    IntDif = 0
    PenDif = 0
    Opt1Dif = 0
    Opt2Dif = 0
    Opt3Dif = 0
    BillBal = OldRound(PersDif + IntDif + MTDif + MCDif + FEDif + MHDif + PenDif + Opt1Dif + Opt2Dif + Opt3Dif)
    ThisBillNum = "NA"
    Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); QPTrim$(TaxTrans.Description);
    Print #RptHandle, Tab(37); Using$("###0", TaxTrans.TaxYear); Tab(45); Using$("$##,##0.00", TaxTrans.Revenue.PrePaidAmt);
    Print #RptHandle, Tab(56); Using$("$##,##0.00", TaxTrans.Amount); 'Tab(69);
    Print #RptHandle, Tab(79); ThisTransType
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines - 4 Then
      Print #RptHandle, FF$
      GoSub PrintHeader
      GoSub PrintCustHeader
    End If
    Print #RptHandle, Tab(15); "Personal          "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle1); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle1Pd); Tab(75); "Belongs to Bill#"
    Print #RptHandle, Tab(15); "Machine Tools     "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle2); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle2Pd)
    Print #RptHandle, Tab(15); "Merchant Capital  "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle3); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle3Pd)
    Print #RptHandle, Tab(15); "Farm Equipment    "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle4); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle4Pd)
    Print #RptHandle, Tab(15); "Mobile Homes      "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle5); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle5Pd)
    Print #RptHandle, Tab(15); "Interest          "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Interest); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.InterestPd); Tab(75); "NA"
    Print #RptHandle, Tab(15); "Penalty           "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Penalty); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.PenaltyPd); Tab(75); "NA"
    LineCnt = LineCnt + 7
    If LineCnt >= MaxLines - 3 Then
      Print #RptHandle, FF$
      GoSub PrintHeader
      GoSub PrintCustHeader
    End If
    If Len(Opt1Desc) > 0 Then
      Print #RptHandle, Tab(15); Opt1Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1Pd)
      LineCnt = LineCnt + 1
    End If
    If Len(Opt2Desc) > 0 Then
      Print #RptHandle, Tab(15); Opt2Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2Pd)
      LineCnt = LineCnt + 1
    End If
    If Len(Opt3Desc) > 0 Then
      Print #RptHandle, Tab(15); Opt3Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3Pd)
      LineCnt = LineCnt + 1
    End If
    Print #RptHandle, Tab(15); String(80, "x")
    If LineCnt >= MaxLines - 2 Then
      Print #RptHandle, FF$
      GoSub PrintHeader
      GoSub PrintCustHeader
    End If
NotThisTrans:
    ThisRec = TaxTrans.LastTrans
  Loop
  Return
  
ApplyDisc:
  Disc1 = 0
  Disc2 = 0
  Disc3 = 0
  Disc4 = 0
  Disc5 = 0
  Disc6 = 0
  Disc7 = 0
  Disc8 = 0
  If TaxTrans.Amount = 0 Then Return
  If TaxTrans.TranType = 1 Then
    SaveAmt = OldRound(TaxTrans.Amount - TaxTrans.DiscAmt)
  Else
    SaveAmt = TaxTrans.Amount
    TaxTrans.Amount = OldRound(TaxTrans.Amount + TaxTrans.DiscAmt)
  End If
  Disc1 = OldRound(TaxTrans.Revenue.Principle1Pd / SaveAmt)
  Disc1 = OldRound(Disc1 * TaxTrans.DiscAmt)
  Disc2 = OldRound(TaxTrans.Revenue.Principle2Pd / SaveAmt)
  Disc2 = OldRound(Disc2 * TaxTrans.DiscAmt)
  Disc3 = OldRound(TaxTrans.Revenue.Principle3Pd / SaveAmt)
  Disc3 = OldRound(Disc3 * TaxTrans.DiscAmt)
  Disc4 = OldRound(TaxTrans.Revenue.Principle4Pd / SaveAmt)
  Disc4 = OldRound(Disc4 * TaxTrans.DiscAmt)
  Disc5 = OldRound(TaxTrans.Revenue.Principle5Pd / SaveAmt)
  Disc5 = OldRound(Disc5 * TaxTrans.DiscAmt)
  Disc6 = OldRound(TaxTrans.Revenue.RevOpt1Pd / SaveAmt)
  Disc6 = OldRound(Disc6 * TaxTrans.DiscAmt)
  Disc7 = OldRound(TaxTrans.Revenue.RevOpt2Pd / SaveAmt)
  Disc7 = OldRound(Disc7 * TaxTrans.DiscAmt)
  Disc8 = OldRound(TaxTrans.Revenue.RevOpt3Pd / SaveAmt)
  Disc8 = OldRound(Disc8 * TaxTrans.DiscAmt)
  TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd + Disc1)
  TaxTrans.Revenue.Principle2Pd = OldRound(TaxTrans.Revenue.Principle2Pd + Disc2)
  TaxTrans.Revenue.Principle3Pd = OldRound(TaxTrans.Revenue.Principle3Pd + Disc3)
  TaxTrans.Revenue.Principle4Pd = OldRound(TaxTrans.Revenue.Principle4Pd + Disc4)
  TaxTrans.Revenue.Principle5Pd = OldRound(TaxTrans.Revenue.Principle5Pd + Disc5)
  TaxTrans.Revenue.RevOpt1Pd = OldRound(TaxTrans.Revenue.RevOpt1Pd + Disc6)
  TaxTrans.Revenue.RevOpt2Pd = OldRound(TaxTrans.Revenue.RevOpt2Pd + Disc7)
  TaxTrans.Revenue.RevOpt3Pd = OldRound(TaxTrans.Revenue.RevOpt3Pd + Disc8)
  DiscApplied = True
  
  Return
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxTransJournal", "PrintRTextDet", Erl)
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
