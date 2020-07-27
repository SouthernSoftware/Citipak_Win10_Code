VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmVATaxMasterBalList 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Master Balance Listing"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxMasterBalList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   6864
      Left            =   1908
      TabIndex        =   6
      Top             =   948
      Width           =   7788
      _Version        =   196609
      _ExtentX        =   13737
      _ExtentY        =   12107
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmVATaxMasterBalList.frx":08CA
      Begin LpLib.fpCombo fpcmbPropType 
         Height          =   405
         Left            =   2925
         TabIndex        =   5
         Top             =   4440
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
         ColDesigner     =   "frmVATaxMasterBalList.frx":08E6
      End
      Begin LpLib.fpCombo fpcmbIncInactive 
         Height          =   405
         Left            =   3525
         TabIndex        =   2
         Top             =   2730
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
         ColDesigner     =   "frmVATaxMasterBalList.frx":0BDD
      End
      Begin LpLib.fpCombo fpcmbPrintOpt 
         Height          =   405
         Left            =   2925
         TabIndex        =   3
         Top             =   3285
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
         ColDesigner     =   "frmVATaxMasterBalList.frx":0ED4
      End
      Begin LpLib.fpCombo fpcmbPrintOrder 
         Height          =   405
         Left            =   2925
         TabIndex        =   4
         Top             =   3870
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
         ColDesigner     =   "frmVATaxMasterBalList.frx":11CB
      End
      Begin LpLib.fpCombo fpcmbDetSum 
         Height          =   405
         Left            =   3525
         TabIndex        =   1
         Top             =   2160
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
         ColDesigner     =   "frmVATaxMasterBalList.frx":14C2
      End
      Begin LpLib.fpCombo fpcmbTaxYear 
         Height          =   405
         Left            =   3720
         TabIndex        =   0
         Top             =   1560
         Width           =   1380
         _Version        =   196608
         _ExtentX        =   2434
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
         ColDesigner     =   "frmVATaxMasterBalList.frx":17B9
      End
      Begin VB.CheckBox chkZeroBal 
         BackColor       =   &H00D0D0D0&
         Caption         =   "Check to Include Zero Balances"
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
         Left            =   2280
         TabIndex        =   16
         Top             =   5040
         Width           =   3732
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   636
         Left            =   2040
         TabIndex        =   13
         TabStop         =   0   'False
         Tag             =   "Press the 'Cancel' button to exit this screen and return to the main 'Business License Reports' menu."
         Top             =   5880
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
         ButtonDesigner  =   "frmVATaxMasterBalList.frx":1AB0
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   636
         Left            =   4272
         TabIndex        =   14
         TabStop         =   0   'False
         Tag             =   $"frmVATaxMasterBalList.frx":1C8E
         Top             =   5880
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
         ButtonDesigner  =   "frmVATaxMasterBalList.frx":1D39
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
         Left            =   1116
         TabIndex        =   15
         Top             =   4524
         Width           =   1668
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
         Left            =   1395
         TabIndex        =   12
         Top             =   2835
         Width           =   1950
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Year:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   11
         Top             =   1635
         Width           =   1095
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
         Left            =   1275
         TabIndex        =   10
         Top             =   3390
         Width           =   1500
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
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Master Balance Listing"
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
         TabIndex        =   9
         Top             =   450
         Width           =   4335
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
         Left            =   1470
         TabIndex        =   8
         Top             =   3960
         Width           =   1305
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   4356
         Left            =   1008
         Top             =   1248
         Width           =   5976
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
         TabIndex        =   7
         Top             =   2280
         Width           =   1905
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   7128
      Left            =   1788
      Top             =   804
      Width           =   8052
   End
End
Attribute VB_Name = "frmVATaxMasterBalList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  'Private Temp_Class As Resize_Class
  Dim Town$
  Dim UseOpt As String * 1
  Dim ThisOpt$
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
      frmVATaxMsg.Label1.Caption = "Pitch 10 is recommended for this printout."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show vbModal
      Call PrintText
    End If
  Else
    If fpcmbPrintOpt.Text = "Graphical" Then
      If fpcmbPropType.Text = "Real Only" Then
        Call PrintRGraphicsDet
      Else
        Call PrintPGraphicsDet
      End If
    Else
      frmVATaxMsg.Label1.Caption = "Pitch 10 is recommended for this printout."
      frmVATaxMsg.Label1.Top = 900
      frmVATaxMsg.Show vbModal
      If fpcmbPropType.Text = "Real Only" Then
        Call PrintRTextDet
      Else
        Call PrintPTextDet
      End If
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
  Me.HelpContextID = hlpMasterBalance
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxMasterBalList.")
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
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long, y As Integer
  Dim YrCnt As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim BigYr As Integer
  Dim ThisBigYr As Integer
  Dim HoldYr As Integer
  Dim Nextx As Integer
  Dim Thisx As Integer
  
  On Error GoTo ERRORSTUFF
  
  UseOpt = "N"
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  Town = QPTrim$(TaxMasterRec.Name)
  Opt1Desc = QPTrim$(TaxMasterRec.OptRev1)
  Opt2Desc = QPTrim$(TaxMasterRec.OptRev2)
  Opt3Desc = QPTrim$(TaxMasterRec.OptRev3)
  POpt1Desc = QPTrim$(TaxMasterRec.POptRev1)
  POpt2Desc = QPTrim$(TaxMasterRec.POptRev2)
  POpt3Desc = QPTrim$(TaxMasterRec.POptRev3)
  
  frmVATaxLoadReport.Label1.Caption = "Loading Years"
  frmVATaxLoadReport.Show
  DoEvents
  ReDim Years(1 To 1) As Integer
  YrCnt = 0
  OpenTaxTransFile TTHandle, NumOfTTRecs
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
    If YrCnt = 0 Then
      If TaxTrans.TaxYear > 0 Then
        YrCnt = YrCnt + 1
        ReDim Preserve Years(1 To YrCnt) As Integer
        Years(YrCnt) = TaxTrans.TaxYear
      End If
    Else
      For y = 1 To YrCnt
        If TaxTrans.TaxYear = Years(y) Then
          Exit For
        End If
      Next y
      If y > YrCnt Then
        YrCnt = YrCnt + 1
        ReDim Preserve Years(1 To YrCnt) As Integer
        Years(YrCnt) = TaxTrans.TaxYear
      End If
    End If
  Next x
  Close TTHandle
  
  BigYr = 0
  For x = 1 To YrCnt
    If Years(x) > BigYr Then
      BigYr = Years(x)
    End If
  Next x
  
  Nextx = 1
  ThisBigYr = BigYr + 1
  Do While Nextx <= YrCnt
    For x = Nextx To YrCnt
      If Years(x) < ThisBigYr Then
        ThisBigYr = Years(x)
        Thisx = x
      End If
    Next x
    HoldYr = Years(Nextx)
    Years(Nextx) = Years(Thisx)
    Years(Thisx) = HoldYr
    Nextx = Nextx + 1
    ThisBigYr = BigYr + 1
  Loop
    
  fpcmbTaxYear.Text = "All"
  fpcmbTaxYear.AddItem "All"
  
  For x = YrCnt To 1 Step -1
    fpcmbTaxYear.AddItem CStr(Years(x))
  Next x
  
  Unload frmVATaxLoadReport
  fpcmbIncInactive.Text = "Both"
  fpcmbIncInactive.AddItem "Both"
  fpcmbIncInactive.AddItem "Active Only"
  fpcmbIncInactive.AddItem "Inactive Only"
  
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
  
  fpcmbPropType.Action = ActionClear
  fpcmbPropType.Text = "Both"
  fpcmbPropType.AddItem "Both"
  fpcmbPropType.AddItem "Real Only"
  fpcmbPropType.AddItem "Personal Only"
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxMasterBalList", "LoadMe", Erl)
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
  fpcmbPropType.Action = ActionClear
  If fpcmbDetSum.Text = "Detail" Then
    fpcmbPropType.Text = "Real Only"
    fpcmbPropType.AddItem "Real Only"
    fpcmbPropType.AddItem "Personal Only"
  Else
    fpcmbPropType.Text = "Both"
    fpcmbPropType.AddItem "Both"
    fpcmbPropType.AddItem "Real Only"
    fpcmbPropType.AddItem "Personal Only"
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

Private Sub fpcmbIncInactive_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbIncInactive.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbIncInactive.ListIndex = -1
  End If
  If fpcmbIncInactive.ListDown <> True Then
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

Private Sub fpcmbPropType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbPropType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPropType.ListIndex = -1
  End If
  If fpcmbPropType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbTaxYear.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbTaxYear_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbTaxYear.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbTaxYear.ListIndex = -1
  End If
  If fpcmbTaxYear.ListDown <> True Then
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

Private Sub PrintGraphics()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long, y As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim Balance As Double
  Dim ThisRec As Long
  Dim ThisYear As Integer
  Dim GTCnt As Long
  Dim TCnt As Integer
  Dim CustName$
  Dim YrCnt As Integer
  Dim GYrCnt As Integer
  Dim OverPay As Double
  Dim HoldYr As Integer
  Dim HoldBal As Double
  Dim Nextz As Integer
  Dim z As Integer
  Dim Thisz As Integer
  Dim BigYr As Integer
  Dim ThisBigYr As Integer
  Dim GBal As Double
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim RptFile$
  Dim RptHandle As Integer
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim dlm$
  Dim CustRec As Long
  Dim InactiveFlag As Boolean
  Dim SubRptFile$
  Dim SubRptHandle As Integer
  Dim TransCnt As Long
  Dim OP As Double
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim ActiveFlag$
  Dim CustTotBal As Double
  Dim PropFlag$
  Dim ThisOP$
  Dim TestBal#
  Dim CreditUsed As Double '2/2/07
  Dim OverPaid As Double '2/2/07
'  Dim AHandle As Integer

  On Error GoTo ERRORSTUFF
'  AHandle = FreeFile
'  Open "masterbal.dat" For Output As AHandle
  
  PropFlag = "B"
  If fpcmbPropType.Text = "Real Only" Then
    PropFlag = "R"
  ElseIf fpcmbPropType.Text = "Personal Only" Then
    PropFlag = "P"
  End If
  
  IdxFlag = False
  ActiveFlag = "B"
  If QPTrim$(fpcmbIncInactive.Text) = "Active Only" Then
    ActiveFlag = "A"
  ElseIf QPTrim$(fpcmbIncInactive.Text) = "Inactive Only" Then
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
  
  RptFile$ = "TAXRPTS\TXMSTBAL.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  ReDim GYearBal(1 To 1) As Double
  ReDim GYears(1 To 1) As Integer
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  GYrCnt = 0
  
  frmVATaxShowPctComp.Label1 = "Gathering Tax Transaction Data"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  
  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  End If
  
  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
      CustRec = IdxArray(x)
    Else
      Get TCHandle, x, TaxCust
      CustRec = x
    End If
    Balance = 0 '7/19/06
    If TaxCust.Deleted <> 0 Then GoTo Inactive
    If ActiveFlag = "A" And TaxCust.Active = "N" Then
      GoTo Inactive
    ElseIf ActiveFlag = "I" And TaxCust.Active = "Y" Then
      GoTo Inactive
    End If
    
    If ActiveFlag = "B" Then
      If TaxCust.Active = "N" Then
        CustName = QPTrim$(TaxCust.CustName) + "(I)"
      Else
        CustName = QPTrim$(TaxCust.CustName)
      End If
    Else
      CustName = QPTrim$(TaxCust.CustName)
    End If
'    If TaxCust.Acct = 1284 Then Stop
    If fpcmbTaxYear.Text = "All" Then
      GoSub GetAllBalance
    Else
      GoSub GetYearBalance
    End If
    CustTotBal = TestBal#
    If TaxCust.LastTrans > 0 Then
      ReDim YearBal(1 To 1) As Double
      YrCnt = 0
      ReDim Years(1 To 1) As Integer
      ThisRec = TaxCust.LastTrans
      If fpcmbTaxYear.Text = "All" Then
        OP = CustTotBal
        If OP < 0 Then
          OverPay = OldRound(OverPay + Abs(OP))
        End If
      ElseIf fpcmbTaxYear.Text <> "All" Then
        OP = CustTotBal
        If OP < 0 Then
          OverPay = OldRound(OverPay + Abs(OP))
        End If
      End If
      Do While ThisRec > 0
        Get TTHandle, ThisRec, TaxTrans
        If OP < 0 Then
          Balance = OP
          TaxTrans.LastTrans = 0
          If fpcmbTaxYear.Text <> "All" Then
            TaxTrans.TaxYear = CInt(fpcmbTaxYear.Text) 'tax year balance will be the balance
            'for the selected year
          End If
          GoTo GoWithZero
        End If
        If TaxTrans.TranType = 1 Then
          If fpcmbTaxYear.Text <> "All" Then
            If CInt(fpcmbTaxYear.Text) <> TaxTrans.TaxYear Then GoTo SkipIt
          End If
          If PropFlag <> "B" Then
            If PropFlag <> TaxTrans.BillType Then GoTo SkipIt
          End If
          Balance# = OldRound#(TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
          Balance# = OldRound#(Balance# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
          Balance# = OldRound#(Balance# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
          Balance# = OldRound#(Balance# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
          Balance# = OldRound#(Balance# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt + TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
          If Balance = 0 And chkZeroBal.Value = 1 Then
            GoTo GoWithZero
          End If
          If Balance <> 0 Then
GoWithZero:
            GBal = OldRound(GBal + Balance#)
            If YrCnt = 0 Then
              YrCnt = YrCnt + 1
              ReDim Preserve Years(1 To YrCnt) As Integer
              Years(YrCnt) = TaxTrans.TaxYear
              ReDim Preserve YearBal(1 To YrCnt) As Double
              YearBal(YrCnt) = Balance#
            Else
              For y = 1 To YrCnt
                If Years(y) = TaxTrans.TaxYear Then '
                  YearBal(y) = OldRound(YearBal(y) + Balance#)
                  Exit For
                End If '
              Next y
              If y > YrCnt Then '
                YrCnt = YrCnt + 1
                ReDim Preserve Years(1 To YrCnt) As Integer
                Years(YrCnt) = TaxTrans.TaxYear
                ReDim Preserve YearBal(1 To YrCnt) As Double
                YearBal(YrCnt) = Balance#
              End If '
            End If '
            
            If GYrCnt = 0 Then '
              GYrCnt = GYrCnt + 1
              ReDim Preserve GYears(1 To GYrCnt) As Integer
              GYears(GYrCnt) = TaxTrans.TaxYear
              ReDim Preserve GYearBal(1 To GYrCnt) As Double
              GYearBal(GYrCnt) = Balance#
            Else
              For y = 1 To GYrCnt
                If GYears(y) = TaxTrans.TaxYear Then '
                  GYearBal(y) = OldRound(GYearBal(y) + Balance#)
                  Exit For
                End If '
              Next y
              If y > GYrCnt Then '
                GYrCnt = GYrCnt + 1
                ReDim Preserve GYears(1 To GYrCnt) As Integer
                GYears(GYrCnt) = TaxTrans.TaxYear
                ReDim Preserve GYearBal(1 To GYrCnt) As Double
                GYearBal(GYrCnt) = Balance#
              End If '
            End If '
            
          BigYr = 0
          Nextz = 1
          For z = 1 To YrCnt
            If Years(z) > BigYr Then
              BigYr = Years(z)
            End If
          Next z
          
          ThisBigYr = BigYr + 1
          Nextz = 1
          Do While Nextz <= YrCnt
            For z = Nextz To YrCnt
              If Years(z) < ThisBigYr Then
                ThisBigYr = Years(z)
                Thisz = z
              End If
            Next z
            HoldYr = Years(Nextz)
            HoldBal = YearBal(Nextz)
            Years(Nextz) = Years(Thisz)
            YearBal(Nextz) = YearBal(Thisz)
            Years(Thisz) = HoldYr
            YearBal(Thisz) = HoldBal
            Nextz = Nextz + 1
            ThisBigYr = BigYr + 1
          Loop
        End If
      End If
SkipIt:
        ThisRec = TaxTrans.LastTrans
      Loop
'      Print #AHandle, CStr(TaxCust.Acct) & "~" & QPTrim$(TaxCust.CustName) & "~" & Using$("$##,###.##", TestBal)
      For z = YrCnt To 1 Step -1
        TransCnt = TransCnt + 1
        '                   0            1             2            3               4                5            6
        Print #RptHandle, Town; dlm; CustName; dlm; CustRec; dlm; Years(z); dlm; YearBal(z); dlm; TransCnt; dlm; GBal; dlm;
        If UseOpt = "Y" Then
          If fpcmbTaxYear.Text = "All" Then 'added 12/5/06
            '                    7                     8                           9                10              11             12               13
            Print #RptHandle, ThisOpt; dlm; QPTrim$(TaxCust.OptSrchDesc); dlm; ActiveFlag; dlm; CustTotBal; dlm; PropFlag; dlm; OverPay; dlm; fpcmbTaxYear.Text
          Else
            '                    7                     8                           9                10              11          12            13
            Print #RptHandle, ThisOpt; dlm; QPTrim$(TaxCust.OptSrchDesc); dlm; ActiveFlag; dlm; CustTotBal; dlm; PropFlag; dlm; 0; dlm; fpcmbTaxYear.Text
          End If
        Else
          If fpcmbTaxYear.Text = "All" Then 'added 12/5/06
            '                  7        8           9                10              11            12                13
            Print #RptHandle, ""; dlm; ""; dlm; ActiveFlag; dlm; CustTotBal; dlm; PropFlag; dlm; OverPay; dlm; fpcmbTaxYear.Text
          Else
            '                  7        8           9                10              11          12             13
            Print #RptHandle, ""; dlm; ""; dlm; ActiveFlag; dlm; CustTotBal; dlm; PropFlag; dlm; 0; dlm; fpcmbTaxYear.Text
          End If
        End If
      Next z
    End If
Inactive:
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
  
  BigYr = 0
  For x = 1 To GYrCnt
    If GYears(x) > BigYr Then
      BigYr = GYears(x)
    End If
  Next x
  Close
  
  ThisBigYr = BigYr + 1
  Nextz = 1
  Do While Nextz <= GYrCnt
    For z = Nextz To GYrCnt
      If GYears(z) < ThisBigYr Then
         ThisBigYr = GYears(z)
         Thisz = z
      End If
    Next z
    HoldYr = GYears(Nextz)
    HoldBal = GYearBal(Nextz)
    GYears(Nextz) = GYears(Thisz)
    GYearBal(Nextz) = GYearBal(Thisz)
    GYears(Thisz) = HoldYr
    GYearBal(Thisz) = HoldBal
    Nextz = Nextz + 1
    ThisBigYr = BigYr + 1
  Loop
  
  SubRptFile$ = "TAXRPTS\TXMSTBALSUB.RPT"
  SubRptHandle = FreeFile
  Open SubRptFile For Output As #SubRptHandle
  
  If InStr(CStr(OverPay), "E") Or fpcmbTaxYear.Text <> "All" Then OverPay = 0 'added All 12/5/06
  For x = 1 To GYrCnt
    If x = GYrCnt Then
      '                        0               1                2
      Print #SubRptHandle, GYears(x); dlm; GYearBal(x); dlm; OverPay; dlm; 1
    Else
      '                        0               1                2
      Print #SubRptHandle, GYears(x); dlm; GYearBal(x); dlm; OverPay; dlm; 0
    End If
  Next x
  
  Close
  If GYrCnt = 0 Then
    Call TaxMsg(900, "There are no balances to report for the parameters entered.")
    Exit Sub
  End If
        
  arVATaxMasterBalSum.Show
  
  Exit Sub
  
GetYearBalance:
  ThisRec = TaxCust.LastTrans
  TestBal = 0
  CreditUsed = 0
  OverPaid = 0
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
    If CInt(fpcmbTaxYear.Text) <> TaxTrans.TaxYear Then GoTo SkipItYear
    If PropFlag = "R" Then
      If TaxTrans.BillType <> "R" Then GoTo SkipItYear
    ElseIf PropFlag = "P" Then
      If TaxTrans.BillType <> "P" Then GoTo SkipItYear
    End If
    If TaxTrans.TranType = 10 Then 'added 10/3/06 adjust bill down affecting credit
       TestBal = TestBal + TaxTrans.Revenue.PrePaidUsed
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 11 Then 'added 10/3/06 prepay adjust down
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 24 Then 'added 10/3/06 adjust bill up affecting credit
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 21 Or TaxTrans.TranType = 22 Then
'      TestBal# = OldRound#(TestBal# - TaxTrans.Revenue.PrePaidAmt)'remmed on 2/2/07
      OverPaid = OldRound(OverPaid + TaxTrans.Revenue.PrePaidAmt) 'added 2/2/07
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 12 Then 'refund on prepay 7/17/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 9 Then 'credit applied at billing  'added 2/2/07
      CreditUsed = OldRound(CreditUsed + TaxTrans.Revenue.PrePaidUsed)
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 1 Then
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt + TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
    End If
SkipItYear:
    ThisRec = TaxTrans.LastTrans
  Loop
  If OverPaid = 0 Then CreditUsed = 0 'added 2/20/07
  
  TestBal = OldRound(TestBal - (OverPaid - CreditUsed)) 'added 2/2/07
  
  Return
  
GetAllBalance:
  ThisRec = TaxCust.LastTrans
  TestBal = 0
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
    If PropFlag = "R" Then
      If TaxTrans.BillType <> "R" Then GoTo DoAgain
    ElseIf PropFlag = "P" Then
      If TaxTrans.BillType <> "P" Then GoTo DoAgain
    End If
    If TaxTrans.TranType = 10 Then 'added 10/3/06 adjust bill down affecting credit
       TestBal = TestBal + TaxTrans.Revenue.PrePaidUsed
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 11 Then 'added 10/3/06 prepay adjust down
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 24 Then 'added 10/3/06 adjust bill up affecting credit
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 21 Or TaxTrans.TranType = 22 Then
      TestBal# = OldRound#(TestBal# - TaxTrans.Revenue.PrePaidAmt)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 12 Then 'refund on prepay 7/17/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 9 Then 'added 8/11/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 1 Then
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt + TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
    End If
DoAgain:
    ThisRec = TaxTrans.LastTrans
  Loop
  
  Return

ERRORSTUFF:
   Unload frmVATaxShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxMasterBalList", "PrintGraphics", Erl)
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
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long, y As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim Balance As Double
  Dim ThisRec As Long
  Dim ThisYear As Integer
  Dim GTCnt As Long
  Dim TCnt As Integer
  Dim CustName$
  Dim YrCnt As Integer
  Dim GYrCnt As Integer
  Dim OverPay As Double
  Dim HoldYr As Integer
  Dim HoldBal As Double
  Dim Nextz As Integer
  Dim z As Integer
  Dim Thisz As Integer
  Dim BigYr As Integer
  Dim ThisBigYr As Integer
  Dim GBal As Double
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim RptFile$
  Dim RptHandle As Integer
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim CustRec As Long
  Dim InactiveFlag As Boolean
  Dim SubRptFile$
  Dim SubRptHandle As Integer
  Dim TransCnt As Long
  Dim OP As Double
  Dim LineCnt As Integer
  Dim MaxLines As Integer
  Dim Page As Integer
  Dim FF$
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim ActiveFlag$
  Dim ThisAcct$
  Dim ThatAcct$
  Dim CustTotBal As Double
  Dim PropFlag$
  Dim TestBal As Double
  Dim ThisOP$
  Dim PrintCHeader As Boolean
  Dim CreditUsed As Double '2/2/07
  Dim OverPaid As Double '2/2/07
  
  On Error GoTo ERRORSTUFF
  
  PrintCHeader = False
  PropFlag = "B"
  If fpcmbPropType.Text = "Real Only" Then
    PropFlag = "R"
  ElseIf fpcmbPropType.Text = "Personal Only" Then
    PropFlag = "P"
  End If
  
  FF$ = Chr(12)
  MaxLines = 58
  
  IdxFlag = False
  ActiveFlag = "B"
  If QPTrim$(fpcmbIncInactive.Text) = "Active Only" Then
    ActiveFlag = "A"
  ElseIf QPTrim$(fpcmbIncInactive.Text) = "Inactive Only" Then
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

  RptFile$ = "TAXRPTS\TXMSTBAL.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  ReDim GYearBal(1 To 1) As Double
  ReDim GYears(1 To 1) As Integer
  GoSub PrintHeader
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  GYrCnt = 0
  frmVATaxShowPctComp.Label1 = "Gathering Tax Transaction Data"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  
  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  End If
  
  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
      CustRec = IdxArray(x)
    Else
      Get TCHandle, x, TaxCust
      CustRec = x
    End If
    If TaxCust.Deleted <> 0 Then GoTo Inactive 'SkipIt
    If ActiveFlag = "A" And TaxCust.Active = "N" Then
      GoTo Inactive
    ElseIf ActiveFlag = "I" And TaxCust.Active = "Y" Then
      GoTo Inactive
    End If
    PrintCHeader = True
    If ActiveFlag = "B" Then
      If TaxCust.Active = "N" Then
        CustName = QPTrim$(TaxCust.CustName) + "(I)"
      Else
        CustName = QPTrim$(TaxCust.CustName)
      End If
    Else
      CustName = QPTrim$(TaxCust.CustName)
    End If
   
    ThisAcct = CStr(TaxCust.Acct)
    TestBal = 0
    
    If fpcmbTaxYear.Text <> "All" Then
      If TaxCust.LastTrans > 0 Then
        GoSub GetYearBalance
        CustTotBal = TestBal
      End If
      If CustTotBal = 0 Then
        If chkZeroBal.Value = 0 Then
          GoTo Inactive
        End If
      End If
    Else
      GoSub GetAllBalance
      CustTotBal = TestBal
      If CustTotBal = 0 Then
        If chkZeroBal.Value = 0 Then
          GoTo Inactive
        End If
      End If
    End If
    
    If TaxCust.LastTrans > 0 Then
      ReDim YearBal(1 To 1) As Double
      YrCnt = 0
      ReDim Years(1 To 1) As Integer
      If fpcmbTaxYear.Text = "All" Then
        OP = CustTotBal
        If OP < 0 Then
          OverPay = OldRound(OverPay + Abs(OP))
        End If
      ElseIf fpcmbTaxYear.Text <> "All" Then
        OP = CustTotBal
        If OP < 0 Then
          OverPay = OldRound(OverPay + Abs(OP))
        End If
      End If
      If ThatAcct <> ThisAcct Then
        If LineCnt > MaxLines - 4 Then
          Print #RptHandle, FF$
          GoSub PrintHeader
          If LineCnt <> 6 Then
            Print #RptHandle,
            LineCnt = LineCnt + 1
          End If
        End If
        If TestBal <> 0 Or CustTotBal <> 0 Or chkZeroBal.Value = 1 And PrintCHeader = True Then
          GoSub PrintCustHeader
          PrintCHeader = False
        End If
        ThatAcct = ThisAcct
      End If
      ThisRec = TaxCust.LastTrans
    
      Do While ThisRec > 0
        Get TTHandle, ThisRec, TaxTrans
        If OP < 0 Then
          Balance# = OP
          TaxTrans.LastTrans = 0
          If fpcmbTaxYear.Text <> "All" Then
            TaxTrans.TaxYear = CInt(fpcmbTaxYear.Text) 'tax year balance will be the balance
            'for the selected year
          End If
          GoTo GoWithZero
        End If
        If TaxTrans.TranType = 1 Then
          If fpcmbTaxYear.Text <> "All" Then
            If CInt(fpcmbTaxYear.Text) <> TaxTrans.TaxYear Then GoTo SkipIt
          End If
          If PropFlag <> "B" Then
            If PropFlag <> TaxTrans.BillType Then GoTo SkipIt
          End If
          Balance# = OldRound#(TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
          Balance# = OldRound#(Balance# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
          Balance# = OldRound#(Balance# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
          Balance# = OldRound#(Balance# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
          Balance# = OldRound#(Balance# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt + TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
          If Balance = 0 And chkZeroBal.Value = 1 Then
            GoTo GoWithZero
          End If
          If Balance <> 0 Then
GoWithZero:
            GBal = OldRound(GBal + Balance#)
            If YrCnt = 0 Then '
              YrCnt = YrCnt + 1
              ReDim Preserve Years(1 To YrCnt) As Integer
              Years(YrCnt) = TaxTrans.TaxYear
              ReDim Preserve YearBal(1 To YrCnt) As Double
              YearBal(YrCnt) = Balance#
            Else
              For y = 1 To YrCnt
                If Years(y) = TaxTrans.TaxYear Then '
                  YearBal(y) = OldRound(YearBal(y) + Balance#)
                  Exit For
                End If '
              Next y
              If y > YrCnt Then '
                YrCnt = YrCnt + 1
                ReDim Preserve Years(1 To YrCnt) As Integer
                Years(YrCnt) = TaxTrans.TaxYear
                ReDim Preserve YearBal(1 To YrCnt) As Double
                YearBal(YrCnt) = Balance#
              End If '
            End If '
            
            If GYrCnt = 0 Then '
              GYrCnt = GYrCnt + 1
              ReDim Preserve GYears(1 To GYrCnt) As Integer
              GYears(GYrCnt) = TaxTrans.TaxYear
              ReDim Preserve GYearBal(1 To GYrCnt) As Double
              GYearBal(GYrCnt) = Balance#
            Else
              For y = 1 To GYrCnt
                If GYears(y) = TaxTrans.TaxYear Then '
                  GYearBal(y) = OldRound(GYearBal(y) + Balance#)
                  Exit For
                End If '
              Next y
              If y > GYrCnt Then '
                GYrCnt = GYrCnt + 1
                ReDim Preserve GYears(1 To GYrCnt) As Integer
                GYears(GYrCnt) = TaxTrans.TaxYear
                ReDim Preserve GYearBal(1 To GYrCnt) As Double
                GYearBal(GYrCnt) = Balance#
              End If '
            End If '
            
            BigYr = 0
            Nextz = 1
            For z = 1 To YrCnt
              If Years(z) > BigYr Then
                BigYr = Years(z)
              End If
            Next z
          
            ThisBigYr = BigYr + 1
            Nextz = 1
            Do While Nextz <= YrCnt
              For z = Nextz To YrCnt
                If Years(z) < ThisBigYr Then
                  ThisBigYr = Years(z)
                  Thisz = z
                End If
              Next z
              HoldYr = Years(Nextz)
              HoldBal = YearBal(Nextz)
              Years(Nextz) = Years(Thisz)
              YearBal(Nextz) = YearBal(Thisz)
              Years(Thisz) = HoldYr
              YearBal(Thisz) = HoldBal
              Nextz = Nextz + 1
              ThisBigYr = BigYr + 1
            Loop
          End If
        End If
SkipIt:
          ThisRec = TaxTrans.LastTrans
        Loop
        'NOTE:added solely to match up with graphics report...take out when data is not corrupted
'        If TestBal = 0 And CustTotBal < 0 Then
'          If YrCnt > 0 Then
'            TransCnt = TransCnt + YrCnt
'          End If
'          GoTo Inactive
'        End If
        'NOTE:added solely to match up with graphics report...take out when data is not corrupted
        For z = YrCnt To 1 Step -1
          TransCnt = TransCnt + 1
          Print #RptHandle, Tab(5); "Tax Year: " + Using$("###0", Years(z)); Tab(53); "Year Total:   " + Using$("$#,###,##0.00", YearBal(z))
          LineCnt = LineCnt + 1
          If LineCnt >= MaxLines Then
            Print #RptHandle, FF$
            GoSub PrintHeader
          End If
          If UseOpt = "Y" Then
            Print #RptHandle, Tab(10); ThisOpt + ": " + QPTrim$(TaxCust.OptSrchDesc)
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintHeader
            End If
          End If
        Next z
        If YrCnt > 0 Then
          Print #RptHandle, String(79, "-")
          LineCnt = LineCnt + 1
        End If
      End If
Inactive:
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
  
  BigYr = 0
  For x = 1 To GYrCnt
    If GYears(x) > BigYr Then
      BigYr = GYears(x)
    End If
  Next x
  
  ThisBigYr = BigYr + 1
  Nextz = 1
  Do While Nextz <= GYrCnt
    For z = Nextz To GYrCnt
      If GYears(z) < ThisBigYr Then
         ThisBigYr = GYears(z)
         Thisz = z
      End If
    Next z
    HoldYr = GYears(Nextz)
    HoldBal = GYearBal(Nextz)
    GYears(Nextz) = GYears(Thisz)
    GYearBal(Nextz) = GYearBal(Thisz)
    GYears(Thisz) = HoldYr
    GYearBal(Thisz) = HoldBal
    Nextz = Nextz + 1
    ThisBigYr = BigYr + 1
  Loop
  
  Print #RptHandle, FF$
  
  If InStr(CStr(OverPay), "E") Then OverPay = 0
  Page = Page + 1
  Print #RptHandle, Tab(25); "Tax Master Balance Listing Summary"
  If PropFlag = "B" Then
    Print #RptHandle, "Real And Personal"
  ElseIf PropFlag = "P" Then
    Print #RptHandle, "Personal Only"
  Else
    Print #RptHandle, "Real Only"
  End If
  Print #RptHandle, Town
  Print #RptHandle, "Report Date: " + CStr(Now); Tab(71); "Page # " + CStr(Page)
  Print #RptHandle, String(79, "-")
  
  Print #RptHandle, "Total Entries: " + Using$("##,##0", TransCnt)
  Print #RptHandle, "Total Tax Balance:  " + Using$("$###,###,##0.00", GBal)
  If fpcmbTaxYear.Text = "All" Then
    Print #RptHandle, "Total Over Payment: " + Using$("$###,###,##0.00", OverPay)
  Else
    Print #RptHandle,
  End If
  Print #RptHandle,
  Print #RptHandle, "Tax Totals By Year"
  Print #RptHandle, Tab(12); "Tax Year"; Tab(26); "Amount Owed"
  For x = 1 To GYrCnt
    If x = GYrCnt Then
      Print #RptHandle, Tab(15); Using$("###0", GYears(x)); Tab(24); Using$("$#,###,##0.00", GYearBal(x)) '; Tab(44); Using$("$##,##0.00", OverPay)
    Else
      Print #RptHandle, Tab(15); Using$("###0", GYears(x)); Tab(24); Using$("$#,###,##0.00", GYearBal(x))
    End If
  Next x
  Print #RptHandle, FF$
  Close
  If GYrCnt = 0 Then
    Call TaxMsg(900, "There are no balances to report for the parameters entered.")
    Exit Sub
  End If
        
  ViewPrint RptFile, "Tax Master Balance Report", True
  
  Exit Sub
  
PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(25); "Tax Master Balance Listing Summary"
  Print #RptHandle, Town
  If ActiveFlag = "B" Then
    Print #RptHandle, "Customer Status: " + "Active And Inactive"; Tab(65); "(I) = Inactive"
  Else
    Print #RptHandle, "Customer Status: " + fpcmbIncInactive.Text
  End If
  Print #RptHandle, "Report Date: " + CStr(Now); Tab(71); "Page # " + CStr(Page)
  Print #RptHandle, "Property Type: ";
  If PropFlag = "B" Then
    Print #RptHandle, "Both Real and Personal"
  ElseIf PropFlag = "P" Then
    Print #RptHandle, "Personal Only"
  ElseIf PropFlag = "R" Then
    Print #RptHandle, "Real Only"
  End If
  Print #RptHandle, "Acct #"; Tab(8); "Customer Name" '; Tab(59); "Tax Year"; Tab(73); "Balance"
  Print #RptHandle, String(79, "-")
  LineCnt = 7
  
  Return
  
PrintCustHeader:
  Print #RptHandle, CustRec; Tab(8); CustName; Tab(50); "Total Balance: "; Tab(67); Using$("$#,###,##0.00", CustTotBal)
  Print #RptHandle, String(79, ".")
  LineCnt = LineCnt + 2
  Return

GetYearBalance:
  ThisRec = TaxCust.LastTrans
  TestBal = 0
  CreditUsed = 0
  OverPaid = 0
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
    If CInt(fpcmbTaxYear.Text) <> TaxTrans.TaxYear Then GoTo SkipItYear
    If PropFlag = "R" Then
      If TaxTrans.BillType <> "R" Then GoTo SkipItYear
    ElseIf PropFlag = "P" Then
      If TaxTrans.BillType <> "P" Then GoTo SkipItYear
    End If
    If TaxTrans.TranType = 10 Then 'added 10/3/06 adjust bill down affecting credit
       TestBal = TestBal + TaxTrans.Revenue.PrePaidUsed
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 11 Then 'added 10/3/06 prepay adjust down
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 24 Then 'added 10/3/06 adjust bill up affecting credit
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 21 Or TaxTrans.TranType = 22 Then
'      TestBal# = OldRound#(TestBal# - TaxTrans.Revenue.PrePaidAmt)'remmed on 2/2/07
      OverPaid = OldRound(OverPaid + TaxTrans.Revenue.PrePaidAmt) 'added 2/2/07
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 12 Then 'refund on prepay 7/17/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 9 Then 'credit applied at billing  'added 2/2/07
      CreditUsed = OldRound(CreditUsed + TaxTrans.Revenue.PrePaidUsed)
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 1 Then
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt + TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
    End If
SkipItYear:
    ThisRec = TaxTrans.LastTrans
  Loop
  If OverPaid = 0 Then CreditUsed = 0 'added 2/20/07
  
  TestBal = OldRound(TestBal - (OverPaid - CreditUsed)) 'added 2/2/07
  
  Return
  
GetAllBalance:
  ThisRec = TaxCust.LastTrans
  TestBal = 0
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
    If PropFlag = "R" Then
      If TaxTrans.BillType <> "R" Then GoTo DoAgain
    ElseIf PropFlag = "P" Then
      If TaxTrans.BillType <> "P" Then GoTo DoAgain
    End If
    If TaxTrans.TranType = 10 Then 'added 10/3/06 adjust bill down affecting credit
       TestBal = TestBal + TaxTrans.Revenue.PrePaidUsed
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 11 Then 'added 10/3/06 prepay adjust down
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 24 Then 'added 10/3/06 adjust bill up affecting credit
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 21 Or TaxTrans.TranType = 22 Then
      TestBal# = OldRound#(TestBal# - TaxTrans.Revenue.PrePaidAmt)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 12 Then 'refund on prepay 7/17/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 9 Then 'added 8/11/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 1 Then
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt + TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
    End If
DoAgain:
    ThisRec = TaxTrans.LastTrans
  Loop
  
  Return
  
ERRORSTUFF:
   Unload frmVATaxShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxMasterBalList", "PrintText", Erl)
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
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long, y As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim Balance As Double
  Dim ThisRec As Long
  Dim ThisYear As Integer
  Dim GTCnt As Long
  Dim TCnt As Integer
  Dim CustName$
  Dim YrCnt As Integer
  Dim GYrCnt As Integer
  Dim OverPay As Double
  Dim HoldYr As Integer
  Dim HoldBal As Double
  Dim Nextz As Integer
  Dim z As Integer
  Dim Thisz As Integer
  Dim BigYr As Integer
  Dim ThisBigYr As Integer
  Dim GBal As Double
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim RptFile$
  Dim RptHandle As Integer
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim dlm$
  Dim CustRec As Long
  Dim InactiveFlag As Boolean
  Dim SubRptFile$
  Dim SubRptHandle As Integer
  Dim TransCnt As Long
  Dim OP As Double
  Dim ThisPrincBal As Double
  Dim ThisIntBal As Double
  Dim ThisPenBal As Double
  Dim ThisAdvBal As Double
  Dim ThisLateListBal As Double
  Dim ThisOpt1Bal As Double
  Dim ThisOpt2Bal As Double
  Dim ThisOpt3Bal As Double
  Dim HoldPrinc As Double
  Dim HoldInt As Double
  Dim HoldPen As Double
  Dim HoldAdv As Double
  Dim HoldLateList As Double
  Dim HoldOpt1 As Double
  Dim HoldOpt2 As Double
  Dim HoldOpt3 As Double
  Dim GPrincTot As Double
  Dim GIntTot As Double
  Dim GPenTot As Double
  Dim GAdvTot As Double
  Dim GLateListTot As Double
  Dim GOpt1Tot As Double
  Dim GOpt2Tot As Double
  Dim GOpt3Tot As Double
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim ActiveFlag$
  Dim PinCnt As Integer
  Dim ThisPin$
  Dim PropType$
  Dim CustTotBal As Double
  Dim BillNum$
  Dim ThisOP$
  Dim TestBal#
  Dim Disc1 As Double '1/16/2007
  Dim Disc2 As Double '1/16/2007
  Dim Disc3 As Double '1/16/2007
  Dim Disc4 As Double '1/16/2007
  Dim Disc5 As Double '9/17/2007
  Dim Dif As Double '9/19/07
  Dim DiscApplied As Boolean '1/16/2007
  Dim SaveAmt As Double '1/16/2007
  Dim CreditUsed As Double '2/2/07
  Dim OverPaid As Double '2/2/07
  
  On Error GoTo ERRORSTUFF
  
  IdxFlag = False
  ActiveFlag = "B"
  If QPTrim$(fpcmbIncInactive.Text) = "Active Only" Then
    ActiveFlag = "A"
  ElseIf QPTrim$(fpcmbIncInactive.Text) = "Inactive Only" Then
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

  RptFile$ = "TAXRPTS\TXMSTBALDET.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  ReDim GYearBal(1 To 1) As Double
  ReDim GYears(1 To 1) As Integer
  ReDim GPrincBal(1 To 1) As Double
  ReDim GIntBal(1 To 1) As Double
  ReDim GPenBal(1 To 1) As Double
  ReDim GAdvBal(1 To 1) As Double
  ReDim GLateListBal(1 To 1) As Double
  ReDim GOPt1Bal(1 To 1) As Double
  ReDim GOPt2Bal(1 To 1) As Double
  ReDim GOPt3Bal(1 To 1) As Double
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  GYrCnt = 0
  frmVATaxShowPctComp.Label1 = "Gathering Tax Transaction Data"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  
  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  End If
  
  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
      CustRec = IdxArray(x)
    Else
      Get TCHandle, x, TaxCust
      CustRec = x
    End If
    Balance = 0 'added 7/19/06
    If TaxCust.Deleted <> 0 Then GoTo Inactive
    If ActiveFlag = "A" And TaxCust.Active = "N" Then
      GoTo Inactive
    ElseIf ActiveFlag = "I" And TaxCust.Active = "Y" Then
      GoTo Inactive
    End If
    
    If ActiveFlag = "B" Then
      If TaxCust.Active = "N" Then
        CustName = QPTrim$(TaxCust.CustName) + "(I)"
      Else
        CustName = QPTrim$(TaxCust.CustName)
      End If
    Else
      CustName = QPTrim$(TaxCust.CustName)
    End If
    If fpcmbTaxYear.Text = "All" Then
      GoSub GetAllBalance
    Else
      GoSub GetYearBalance
    End If
    CustTotBal = TestBal#
    OP = TestBal#
    
    If InStr(CStr(CustTotBal), "E") Then CustTotBal = 0
    If TaxCust.LastTrans > 0 Then
      YrCnt = 0
      PinCnt = 0
      ReDim YearBal(1 To 1) As Double
      ReDim Years(1 To 1) As Integer
      ReDim PrincBal(1 To 1) As Double
      ReDim IntBal(1 To 1) As Double
      ReDim PenBal(1 To 1) As Double
      ReDim AdvBal(1 To 1) As Double
      ReDim LateListBal(1 To 1) As Double
      ReDim Opt1Bal(1 To 1) As Double
      ReDim Opt2Bal(1 To 1) As Double
      ReDim Opt3Bal(1 To 1) As Double
      
      ThisRec = TaxCust.LastTrans
      ThisOP = CStr(OP)
      If InStr(ThisOP, "E") Then OP = 0
      If OP < 0 Then
        OverPay = OldRound(OverPay + Abs(OP))
      End If
      
      Do While ThisRec > 0
        Get TTHandle, ThisRec, TaxTrans
        ThisPrincBal = 0 '9/1/06
        ThisIntBal = 0 '9/1/06
        ThisPenBal = 0 '9/1/06
        ThisAdvBal = 0 '9/1/06
        ThisLateListBal = 0 '9/1/06
        ThisOpt1Bal = 0 '9/1/06
        ThisOpt2Bal = 0 '9/1/06
        ThisOpt3Bal = 0 '9/1/06
        Balance = 0 '9/1/06
        If CustTotBal < 0 Then
          Balance# = CustTotBal
          If fpcmbTaxYear.Text <> "All" Then
            TaxTrans.TaxYear = CInt(fpcmbTaxYear.Text)
          End If
          TaxTrans.LastTrans = 0
          GoTo GoWithZero
        End If
        If TaxTrans.TranType = 1 And TaxTrans.BillType = "R" Then
          If fpcmbTaxYear.Text <> "All" Then
            If CInt(fpcmbTaxYear.Text) <> TaxTrans.TaxYear Then GoTo SkipIt
          End If
          If TaxTrans.DiscAmt > 0 Then GoSub ApplyDisc 'added 1/16/07
          ThisPrincBal = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
          ThisPrincBal = OldRound(ThisPrincBal - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
          ThisIntBal = OldRound(TaxTrans.Revenue.Interest - TaxTrans.Revenue.InterestPd)
          ThisPenBal = OldRound(TaxTrans.Revenue.Penalty - TaxTrans.Revenue.PenaltyPd)
          ThisAdvBal = OldRound(TaxTrans.Revenue.Collection - TaxTrans.Revenue.CollectionPd)
          ThisLateListBal = OldRound(TaxTrans.Revenue.LateList - TaxTrans.Revenue.LateListPd)
          ThisOpt1Bal = OldRound(TaxTrans.Revenue.RevOpt1 - TaxTrans.Revenue.RevOpt1Pd)
          ThisOpt2Bal = OldRound(TaxTrans.Revenue.RevOpt2 - TaxTrans.Revenue.RevOpt2Pd)
          ThisOpt3Bal = OldRound(TaxTrans.Revenue.RevOpt3 - TaxTrans.Revenue.RevOpt3Pd)
          Balance# = OldRound#(TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
          Balance# = OldRound#(Balance# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
          Balance# = OldRound#(Balance# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
          Balance# = OldRound#(Balance# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
          Balance# = OldRound#(Balance# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd)) 'changed on 1/16/07 + TaxTrans.DiscAmt))
          If Balance = 0 And chkZeroBal.Value = 1 Then
            GoTo GoWithZero
          End If
          If Balance <> 0 Then
GoWithZero:
            GBal = OldRound(GBal + Balance#)
            If QPTrim$(TaxTrans.RealPin) = "0" And QPTrim$(TaxTrans.PersPin) = "0" Then
              ThisPin = "0"
              PropType = "UNATTACHED"
            ElseIf QPTrim$(TaxTrans.RealPin) = "-1" Then
              ThisPin = "-1"
              PropType = "MOCK"
            ElseIf QPTrim$(TaxTrans.RealPin) <> "0" And QPTrim$(TaxTrans.RealPin) <> "" Then
              ThisPin = QPTrim$(TaxTrans.RealPin)
              PropType = "REAL"
            Else
              ThisPin = "0"
              PropType = "UNATTACHED"
            End If
            
            If GYrCnt = 0 Then '
              GYrCnt = GYrCnt + 1
              ReDim Preserve GYears(1 To GYrCnt) As Integer
              GYears(GYrCnt) = TaxTrans.TaxYear
              ReDim Preserve GYearBal(1 To GYrCnt) As Double
              GYearBal(GYrCnt) = Balance#
              ReDim Preserve GPrincBal(1 To GYrCnt) As Double
              GPrincBal(GYrCnt) = ThisPrincBal#
              GPrincTot = ThisPrincBal#
              ReDim Preserve GIntBal(1 To GYrCnt) As Double
              GIntBal(GYrCnt) = ThisIntBal#
              GIntTot = ThisIntBal#
              ReDim Preserve GPenBal(1 To GYrCnt) As Double
              GPenBal(GYrCnt) = ThisPenBal#
              GPenTot = ThisPenBal#
              ReDim Preserve GAdvBal(1 To GYrCnt) As Double
              GAdvBal(GYrCnt) = ThisAdvBal#
              GAdvTot = ThisAdvBal#
              ReDim Preserve GLateListBal(1 To GYrCnt) As Double
              GLateListBal(GYrCnt) = ThisLateListBal#
              GLateListTot = ThisLateListBal#
              ReDim Preserve GOPt1Bal(1 To GYrCnt) As Double
              GOPt1Bal(GYrCnt) = ThisOpt1Bal#
              GOpt1Tot = ThisOpt1Bal#
              ReDim Preserve GOPt2Bal(1 To GYrCnt) As Double
              GOPt2Bal(GYrCnt) = ThisOpt2Bal#
              GOpt2Tot = ThisOpt2Bal#
              ReDim Preserve GOPt3Bal(1 To GYrCnt) As Double
              GOPt3Bal(GYrCnt) = ThisOpt3Bal#
              GOpt3Tot = ThisOpt3Bal#
            Else
              For y = 1 To GYrCnt
                If GYears(y) = TaxTrans.TaxYear Then '
                  GYearBal(y) = OldRound(GYearBal(y) + Balance#)
                  GPrincBal(y) = OldRound(GPrincBal(y) + ThisPrincBal#)
                  GPrincTot = OldRound(GPrincTot# + ThisPrincBal#)
                  GIntBal(y) = OldRound(GIntBal(y) + ThisIntBal#)
                  GIntTot = OldRound(GIntTot# + ThisIntBal#)
                  GPenBal(y) = OldRound(GPenBal(y) + ThisPenBal#)
                  GPenTot = OldRound(GPenTot# + ThisPenBal#)
                  GAdvBal(y) = OldRound(GAdvBal(y) + ThisAdvBal#)
                  GAdvTot = OldRound(GAdvTot# + ThisAdvBal#)
                  GLateListBal(y) = OldRound(GLateListBal(y) + ThisLateListBal#)
                  GLateListTot = OldRound(GLateListTot# + ThisLateListBal#)
                  GOPt1Bal(y) = OldRound(GOPt1Bal(y) + ThisOpt1Bal#)
                  GOpt1Tot = OldRound(GOpt1Tot# + ThisOpt1Bal#)
                  GOPt2Bal(y) = OldRound(GOPt2Bal(y) + ThisOpt2Bal#)
                  GOpt2Tot = OldRound(GOpt2Tot# + ThisOpt2Bal#)
                  GOPt3Bal(y) = OldRound(GOPt3Bal(y) + ThisOpt3Bal#)
                  GOpt3Tot = OldRound(GOpt3Tot# + ThisOpt3Bal#)
                  Exit For
                End If '
              Next y
              If y > GYrCnt Then '
                GYrCnt = GYrCnt + 1
                ReDim Preserve GYears(1 To GYrCnt) As Integer
                GYears(GYrCnt) = TaxTrans.TaxYear
                ReDim Preserve GYearBal(1 To GYrCnt) As Double
                GYearBal(GYrCnt) = Balance#
                ReDim Preserve GPrincBal(1 To GYrCnt) As Double
                GPrincBal(GYrCnt) = ThisPrincBal#
                GPrincTot# = OldRound(GPrincTot# + ThisPrincBal#)
                
                ReDim Preserve GIntBal(1 To GYrCnt) As Double
                GIntBal(GYrCnt) = ThisIntBal#
                GIntTot# = OldRound(GIntTot# + ThisIntBal#)
                
                ReDim Preserve GPenBal(1 To GYrCnt) As Double
                GPenBal(GYrCnt) = ThisPenBal#
                GPenTot# = OldRound(GPenTot# + ThisPenBal#)
                
                ReDim Preserve GAdvBal(1 To GYrCnt) As Double
                GAdvBal(GYrCnt) = ThisAdvBal#
                GAdvTot# = OldRound(GAdvTot# + ThisAdvBal#)
                ReDim Preserve GLateListBal(1 To GYrCnt) As Double
                GLateListBal(GYrCnt) = ThisLateListBal#
                GLateListTot# = OldRound(GLateListTot + ThisLateListBal#)
                ReDim Preserve GOPt1Bal(1 To GYrCnt) As Double
                GOPt1Bal(GYrCnt) = ThisOpt1Bal#
                GOpt1Tot# = OldRound(GOpt1Tot# + ThisOpt1Bal#)
                ReDim Preserve GOPt2Bal(1 To GYrCnt) As Double
                GOPt2Bal(GYrCnt) = ThisOpt2Bal#
                GOpt2Tot# = OldRound(GOpt2Tot# + ThisOpt2Bal#)
                ReDim Preserve GOPt3Bal(1 To GYrCnt) As Double
                GOPt3Bal(GYrCnt) = ThisOpt3Bal#
                GOpt3Tot# = OldRound(GOpt3Tot# + ThisOpt3Bal#)
              End If '
            End If '
            BillNum = ParseBillNum(TaxTrans.Description)
            TransCnt = TransCnt + 1
            '                   0            1             2                 3                  4              5            6
            Print #RptHandle, Town; dlm; CustName; dlm; CustRec; dlm; TaxTrans.TaxYear; dlm; Balance#; dlm; TransCnt; dlm; GBal; dlm;
            '                         7                8                 9                   10                  11                 12                 13
            Print #RptHandle, ThisPrincBal#; dlm; ThisIntBal#; dlm; ThisAdvBal#; dlm; ThisLateListBal#; dlm; ThisOpt1Bal#; dlm; ThisOpt2Bal#; dlm; ThisOpt3Bal#; dlm;
            '                     14             15            16              17                18            19             20
            Print #RptHandle, GPrincTot; dlm; GIntTot; dlm; GAdvTot; dlm; GLateListTot; dlm; GOpt1Tot; dlm; GOpt2Tot; dlm; GOpt3Tot; dlm;
            '                    21              22            23
            Print #RptHandle, Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm;
            If UseOpt = "Y" Then
              '                    24                      25                        26
              Print #RptHandle, ThisOpt; dlm; QPTrim$(TaxCust.OptSrchDesc); dlm; ActiveFlag; dlm;
            Else
              '                 24       25           26
              Print #RptHandle, ""; dlm; ""; dlm; ActiveFlag; dlm;
            End If
            If fpcmbTaxYear.Text = "All" Then
              '                    27            28              29                30              31             32           33                 34
              Print #RptHandle, PropType; dlm; ThisPin; dlm; CustTotBal; dlm; ThisPenBal#; dlm; GPenTot#; dlm; BillNum; dlm; OverPay; dlm; fpcmbTaxYear.Text
            Else
              '                    27            28              29                30              31             32        33              34
              Print #RptHandle, PropType; dlm; ThisPin; dlm; CustTotBal; dlm; ThisPenBal#; dlm; GPenTot#; dlm; BillNum; dlm; 0; dlm; fpcmbTaxYear.Text
            End If
          End If
        End If
SkipIt:
        ThisRec = TaxTrans.LastTrans
      Loop
    End If
Inactive:
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
  
  BigYr = 0
  For x = 1 To GYrCnt
    If GYears(x) > BigYr Then
      BigYr = GYears(x)
    End If
  Next x
  Close
  
  ThisBigYr = BigYr + 1
  Nextz = 1
  Do While Nextz <= GYrCnt
    For z = Nextz To GYrCnt
      If GYears(z) < ThisBigYr Then
         ThisBigYr = GYears(z)
         Thisz = z
      End If
    Next z
    HoldYr = GYears(Nextz)
    HoldBal = GYearBal(Nextz)
    HoldPrinc = GPrincBal(Nextz)
    HoldInt = GIntBal(Nextz)
    HoldPen = GPenBal(Nextz)
    HoldAdv = GAdvBal(Nextz)
    HoldLateList = GLateListBal(Nextz)
    HoldOpt1 = GOPt1Bal(Nextz)
    HoldOpt2 = GOPt2Bal(Nextz)
    HoldOpt3 = GOPt3Bal(Nextz)
    GYears(Nextz) = GYears(Thisz)
    GYearBal(Nextz) = GYearBal(Thisz)
    GPrincBal(Nextz) = GPrincBal(Thisz)
    GIntBal(Nextz) = GIntBal(Thisz)
    GPenBal(Nextz) = GPenBal(Thisz)
    GAdvBal(Nextz) = GAdvBal(Thisz)
    GLateListBal(Nextz) = GLateListBal(Thisz)
    GOPt1Bal(Nextz) = GOPt1Bal(Thisz)
    GOPt2Bal(Nextz) = GOPt2Bal(Thisz)
    GOPt3Bal(Nextz) = GOPt3Bal(Thisz)
    GYears(Thisz) = HoldYr
    GYearBal(Thisz) = HoldBal
    GPrincBal(Thisz) = HoldPrinc
    GIntBal(Thisz) = HoldInt
    GPenBal(Thisz) = HoldPen
    GAdvBal(Thisz) = HoldAdv
    GLateListBal(Thisz) = HoldLateList
    GOPt1Bal(Thisz) = HoldOpt1
    GOPt2Bal(Thisz) = HoldOpt2
    GOPt3Bal(Thisz) = HoldOpt3
    Nextz = Nextz + 1
    ThisBigYr = BigYr + 1
  Loop
  
  SubRptFile$ = "TAXRPTS\TXMSTBALSUBDET.RPT"
  SubRptHandle = FreeFile
  Open SubRptFile For Output As #SubRptHandle
 
  If InStr(CStr(OverPay), "E") Or fpcmbTaxYear.Text <> "All" Then OverPay = 0
  For x = 1 To GYrCnt
    If x = GYrCnt Then
      '                        0               1                2          3
      Print #SubRptHandle, GYears(x); dlm; GYearBal(x); dlm; OverPay; dlm; 1; dlm;
      '                         4                 5                6                   7
      Print #SubRptHandle, GPrincBal(x); dlm; GIntBal(x); dlm; GAdvBal(x); dlm; GLateListBal(x); dlm;
      '                         8                 9                10               11
      Print #SubRptHandle, GOPt1Bal(x); dlm; GOPt2Bal(x); dlm; GOPt3Bal(x); dlm; Opt1Desc; dlm;
      '                        12            13             14
      Print #SubRptHandle, Opt2Desc; dlm; Opt3Desc; dlm; GPenBal(x)
    Else
      '                        0               1                2
      Print #SubRptHandle, GYears(x); dlm; GYearBal(x); dlm; OverPay; dlm; 0; dlm;
      '                         4                 5                6                   7
      Print #SubRptHandle, GPrincBal(x); dlm; GIntBal(x); dlm; GAdvBal(x); dlm; GLateListBal(x); dlm;
      '                         8                 9                10                11
      Print #SubRptHandle, GOPt1Bal(x); dlm; GOPt2Bal(x); dlm; GOPt3Bal(x); dlm; Opt1Desc; dlm;
      '                        12            13              14
      Print #SubRptHandle, Opt2Desc; dlm; Opt3Desc; dlm; GPenBal(x)
    End If
  Next x
  
  Close
        
  If GYrCnt = 0 Then
    Call TaxMsg(900, "There are no balances to report for the parameters entered.")
    Exit Sub
  End If
  
  arVATaxMasterBalDet.Show
  
  Exit Sub

GetYearBalance:
  ThisRec = TaxCust.LastTrans
  TestBal = 0
  CreditUsed = 0
  OverPaid = 0
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
    If CInt(fpcmbTaxYear.Text) <> TaxTrans.TaxYear Then GoTo SkipItYear
    If TaxTrans.BillType <> "R" Then GoTo SkipItYear
    If TaxTrans.TranType = 10 Then 'added 10/3/06 adjust bill down affecting credit
      TestBal = TestBal + TaxTrans.Revenue.PrePaidUsed
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 11 Then 'added 10/3/06 prepay adjust down
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 24 Then 'added 10/3/06 adjust bill up affecting credit
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 21 Or TaxTrans.TranType = 22 Then
'      TestBal# = OldRound#(TestBal# - TaxTrans.Revenue.PrePaidAmt)'remmed on 2/2/07
      OverPaid = OldRound(OverPaid + TaxTrans.Revenue.PrePaidAmt) 'added 2/2/07
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 12 Then 'refund on prepay 7/17/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 9 Then 'credit applied at billing  'added 2/2/07
      CreditUsed = OldRound(CreditUsed + TaxTrans.Revenue.PrePaidUsed)
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 1 Then
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt + TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
    End If
SkipItYear:
    ThisRec = TaxTrans.LastTrans
  Loop
  If OverPaid = 0 Then CreditUsed = 0 'added 2/20/07
  
  TestBal = OldRound(TestBal - (OverPaid - CreditUsed)) 'added 2/2/07
  
  Return
  
GetAllBalance:
  ThisRec = TaxCust.LastTrans
  TestBal = 0
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
    If TaxTrans.BillType <> "R" Then GoTo DoAgain
    If TaxTrans.TranType = 10 Then 'added 10/3/06 adjust bill down affecting credit
       TestBal = TestBal + TaxTrans.Revenue.PrePaidUsed
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 11 Then 'added 10/3/06 prepay adjust down
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 24 Then 'added 10/3/06 adjust bill up affecting credit
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 21 Or TaxTrans.TranType = 22 Then
      TestBal# = OldRound#(TestBal# - TaxTrans.Revenue.PrePaidAmt)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 12 Then 'refund on prepay 7/17/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 9 Then 'added 8/11/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 1 Then
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt + TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
    End If
DoAgain:
    ThisRec = TaxTrans.LastTrans
  Loop
  
  Return

ApplyDisc:
  Disc1 = 0
  Disc2 = 0
  Disc3 = 0
  Disc4 = 0
  Disc5 = 0
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
  Disc5 = OldRound(TaxTrans.Revenue.LateListPd / SaveAmt)
  Disc5 = OldRound(Disc5 * TaxTrans.DiscAmt)
  TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd + Disc1 + Disc5)
  TaxTrans.Revenue.RevOpt1Pd = OldRound(TaxTrans.Revenue.RevOpt1Pd + Disc2)
  TaxTrans.Revenue.RevOpt2Pd = OldRound(TaxTrans.Revenue.RevOpt2Pd + Disc3)
  TaxTrans.Revenue.RevOpt3Pd = OldRound(TaxTrans.Revenue.RevOpt3Pd + Disc4)
  Dif = OldRound(Disc1 + Disc2 + Disc3 + Disc4 + Disc5) 'added 9/19/07
  If Dif <> TaxTrans.DiscAmt Then
   If Disc1 > 0 Or Disc5 > 0 Then
     TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd + Dif)
   ElseIf Disc2 > 0 Then
     TaxTrans.Revenue.RevOpt1Pd = OldRound(TaxTrans.Revenue.RevOpt1Pd + Dif)
   ElseIf Disc3 > 0 Then
     TaxTrans.Revenue.RevOpt2Pd = OldRound(TaxTrans.Revenue.RevOpt2Pd + Dif)
   ElseIf Disc4 > 0 Then
     TaxTrans.Revenue.RevOpt3Pd = OldRound(TaxTrans.Revenue.RevOpt3Pd + Dif)
   End If
  End If
  DiscApplied = True
  
  Return

ERRORSTUFF:
   Unload frmVATaxShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxMasterBalList", "PrintRGraphicsDet", Erl)
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
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long, y As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim Balance As Double
  Dim ThisRec As Long
  Dim ThisYear As Integer
  Dim GTCnt As Long
  Dim TCnt As Integer
  Dim CustName$
  Dim YrCnt As Integer
  Dim GYrCnt As Integer
  Dim OverPay As Double
  Dim HoldYr As Integer
  Dim HoldBal As Double
  Dim Nextz As Integer
  Dim z As Integer
  Dim Thisz As Integer
  Dim BigYr As Integer
  Dim ThisBigYr As Integer
  Dim GBal As Double
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim RptFile$
  Dim RptHandle As Integer
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim CustRec As Long
  Dim InactiveFlag As Boolean
  Dim SubRptFile$
  Dim SubRptHandle As Integer
  Dim TransCnt As Long
  Dim OP As Double
  Dim ThisPrincBal As Double
  Dim ThisIntBal As Double
  Dim ThisAdvBal As Double
  Dim ThisLateListBal As Double
  Dim ThisPenBal As Double
  Dim ThisOpt1Bal As Double
  Dim ThisOpt2Bal As Double
  Dim ThisOpt3Bal As Double
  Dim HoldPrinc As Double
  Dim HoldInt As Double
  Dim HoldAdv As Double
  Dim HoldLateList As Double
  Dim HoldPen As Double
  Dim HoldOpt1 As Double
  Dim HoldOpt2 As Double
  Dim HoldOpt3 As Double
  Dim GPrincTot As Double
  Dim GIntTot As Double
  Dim GAdvTot As Double
  Dim GLateListTot As Double
  Dim GPenTot As Double
  Dim GOpt1Tot As Double
  Dim GOpt2Tot As Double
  Dim GOpt3Tot As Double
  Dim FF$
  Dim Page As Integer
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim ThisAcct$, ThatAcct$
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim ActiveFlag$
  Dim ThisPin$
  Dim PropType$
  Dim CustTotBal As Double
  Dim BillNum$
  Dim ThisOP$
  Dim TestBal#
  Dim Disc1 As Double '1/16/2007
  Dim Disc2 As Double '1/16/2007
  Dim Disc3 As Double '1/16/2007
  Dim Disc4 As Double '1/16/2007
  Dim Disc5 As Double '9/17/2007
  Dim Dif As Double '9/17/07
  Dim DiscApplied As Boolean '1/16/2007
  Dim SaveAmt As Double '1/16/2007
  Dim CreditUsed As Double '2/2/07
  Dim OverPaid As Double '2/2/07
  
  On Error GoTo ERRORSTUFF
  
  MaxLines = 58
  FF$ = Chr(12)
  IdxFlag = False
  ActiveFlag = "B"
  If QPTrim$(fpcmbIncInactive.Text) = "Active Only" Then
    ActiveFlag = "A"
  ElseIf QPTrim$(fpcmbIncInactive.Text) = "Inactive Only" Then
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

  RptFile$ = "TAXRPTS\TXMSTBALDET.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  GoSub PrintHeader
  ReDim GYearBal(1 To 1) As Double
  ReDim GYears(1 To 1) As Integer
  ReDim GPrincBal(1 To 1) As Double
  ReDim GIntBal(1 To 1) As Double
  ReDim GAdvBal(1 To 1) As Double
  ReDim GLateListBal(1 To 1) As Double
  ReDim GPenBal(1 To 1) As Double
  ReDim GOPt1Bal(1 To 1) As Double
  ReDim GOPt2Bal(1 To 1) As Double
  ReDim GOPt3Bal(1 To 1) As Double
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  GYrCnt = 0
  frmVATaxShowPctComp.Label1 = "Gathering Tax Transaction Data"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  
  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  End If
  
  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
      CustRec = IdxArray(x)
    Else
      Get TCHandle, x, TaxCust
      CustRec = x
    End If
    Balance = 0 'added 7/19/06
    If TaxCust.Deleted <> 0 Then GoTo Inactive 'SkipIt
    If ActiveFlag = "A" And TaxCust.Active = "N" Then
      GoTo Inactive
    ElseIf ActiveFlag = "I" And TaxCust.Active = "Y" Then
      GoTo Inactive
    End If
    
    If ActiveFlag = "B" Then
      If TaxCust.Active = "N" Then
        CustName = QPTrim$(TaxCust.CustName) + "(I)"
      Else
        CustName = QPTrim$(TaxCust.CustName)
      End If
    Else
      CustName = QPTrim$(TaxCust.CustName)
    End If
    ThatAcct = CStr(TaxCust.Acct)
    
    If fpcmbTaxYear.Text = "All" Then
      GoSub GetAllBalance
    Else
      GoSub GetYearBalance
    End If
    
    CustTotBal = TestBal 'GetCustRealBalance(CustRec, -1)
    OP = TestBal
    If TaxCust.LastTrans > 0 Then
      ReDim YearBal(1 To 1) As Double
      YrCnt = 0
      ReDim Years(1 To 1) As Integer
      ReDim PrincBal(1 To 1) As Double
      ReDim IntBal(1 To 1) As Double
      ReDim AdvBal(1 To 1) As Double
      ReDim LateListBal(1 To 1) As Double
      ReDim PenBal(1 To 1) As Double
      ReDim Opt1Bal(1 To 1) As Double
      ReDim Opt2Bal(1 To 1) As Double
      ReDim Opt3Bal(1 To 1) As Double
      ThisAcct = ""
      
      ThisRec = TaxCust.LastTrans
      ThisOP = CStr(OP)
      If InStr(ThisOP, "E") Then OP = 0
      If OP < 0 Then
        OverPay = OldRound(OverPay + Abs(OP))
      End If
      Do While ThisRec > 0
        ThisPrincBal = 0 '9/1/06
        ThisIntBal = 0 '9/1/06
        ThisPenBal = 0 '9/1/06
        ThisAdvBal = 0 '9/1/06
        ThisLateListBal = 0 '9/1/06
        ThisOpt1Bal = 0 '9/1/06
        ThisOpt2Bal = 0 '9/1/06
        ThisOpt3Bal = 0 '9/1/06
        Balance = 0 '9/1/06
        Get TTHandle, ThisRec, TaxTrans
        If CustTotBal < 0 Then
          Balance# = CustTotBal
          If fpcmbTaxYear.Text <> "All" Then
            TaxTrans.TaxYear = CInt(fpcmbTaxYear.Text)
          End If
          TaxTrans.LastTrans = 0
          GoTo GoWithZero
        End If
        If TaxTrans.TranType = 1 And TaxTrans.BillType = "R" Then
          If fpcmbTaxYear.Text <> "All" Then
            If CInt(fpcmbTaxYear.Text) <> TaxTrans.TaxYear Then GoTo SkipIt
          End If
          If TaxTrans.DiscAmt > 0 Then GoSub ApplyDisc '1/16/07
          ThisPrincBal = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
          ThisPrincBal = OldRound(ThisPrincBal - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
          ThisIntBal = OldRound(TaxTrans.Revenue.Interest - TaxTrans.Revenue.InterestPd)
          ThisAdvBal = OldRound(TaxTrans.Revenue.Collection - TaxTrans.Revenue.CollectionPd)
          ThisLateListBal = OldRound(TaxTrans.Revenue.LateList - TaxTrans.Revenue.LateListPd)
          ThisPenBal = OldRound(TaxTrans.Revenue.Penalty - TaxTrans.Revenue.PenaltyPd)
          ThisOpt1Bal = OldRound(TaxTrans.Revenue.RevOpt1 - TaxTrans.Revenue.RevOpt1Pd)
          ThisOpt2Bal = OldRound(TaxTrans.Revenue.RevOpt2 - TaxTrans.Revenue.RevOpt2Pd)
          ThisOpt3Bal = OldRound(TaxTrans.Revenue.RevOpt3 - TaxTrans.Revenue.RevOpt3Pd)
          Balance# = OldRound#(TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
          Balance# = OldRound#(Balance# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
          Balance# = OldRound#(Balance# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
          Balance# = OldRound#(Balance# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
          Balance# = OldRound#(Balance# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd)) 'changed on 1/16/07 + TaxTrans.DiscAmt))
          If Balance = 0 And chkZeroBal.Value = 1 Then
            GoTo GoWithZero
          End If
          If Balance <> 0 Then
GoWithZero:
            If ThisAcct <> ThatAcct Then
              GoSub PrintCustHeader
              ThisAcct = ThatAcct
            End If
            GBal = OldRound(GBal + Balance#)
            If QPTrim$(TaxTrans.RealPin) = "0" And QPTrim$(TaxTrans.PersPin) = "0" Then
              ThisPin = "0"
              PropType = "UNATTACHED"
            ElseIf QPTrim$(TaxTrans.RealPin) = "-1" Then
              ThisPin = "-1"
              PropType = "MOCK"
            ElseIf QPTrim$(TaxTrans.RealPin) <> "0" And QPTrim$(TaxTrans.RealPin) <> "" Then
              ThisPin = QPTrim$(TaxTrans.RealPin)
              PropType = "REAL"
            Else
              ThisPin = "0"
              PropType = "UNATTACHED"
            End If
            TransCnt = TransCnt + 1
            If LineCnt >= MaxLines - 1 Then
              Print #RptHandle, FF$
              GoSub PrintHeader
              GoSub PrintCustHeader
            End If
            Print #RptHandle, "Prop Type: " + PropType; Tab(23); "Pin #: " + ThisPin; Tab(43); "Tax Year: " + Using$("###0", TaxTrans.TaxYear); Tab(59); "Balance: " + Using$("$###,##0.00", Balance#)
            Print #RptHandle, "Bill Number: " + ParseBillNum(TaxTrans.Description)
            LineCnt = LineCnt + 2
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintHeader
              GoSub PrintCustHeader
            End If
            If QPTrim$(Opt1Desc) <> "" Then
              Print #RptHandle, Tab(5); "Principle:    " + Using$("$##,##0.00", ThisPrincBal);
              Print #RptHandle, Tab(40); QPTrim$(Opt1Desc); Tab(69); Using$("$##,##0.00", ThisOpt1Bal)
              LineCnt = LineCnt + 1
              If LineCnt >= MaxLines Then
                Print #RptHandle, FF$
                GoSub PrintHeader
                GoSub PrintCustHeader
              End If
            Else
              Print #RptHandle, Tab(5); "Principle:    " + Using$("$##,##0.00", ThisPrincBal)
              LineCnt = LineCnt + 1
              If LineCnt >= MaxLines Then
                 Print #RptHandle, FF$
                 GoSub PrintHeader
                 GoSub PrintCustHeader
              End If
            End If
            If QPTrim$(Opt2Desc) <> "" Then
              Print #RptHandle, Tab(5); "Interest:     " + Using$("$##,##0.00", ThisIntBal);
              Print #RptHandle, Tab(40); QPTrim$(Opt2Desc); Tab(69); Using$("$##,##0.00", ThisOpt2Bal)
              LineCnt = LineCnt + 1
              If LineCnt >= MaxLines Then
                Print #RptHandle, FF$
                GoSub PrintHeader
                GoSub PrintCustHeader
              End If
            Else
              Print #RptHandle, Tab(5); "Interest:     " + Using$("$##,##0.00", ThisIntBal)
              LineCnt = LineCnt + 1
              If LineCnt >= MaxLines Then
                Print #RptHandle, FF$
                GoSub PrintHeader
                GoSub PrintCustHeader
              End If
            End If
            If QPTrim$(Opt3Desc) <> "" Then
              Print #RptHandle, Tab(5); "Advertising:  " + Using$("$##,##0.00", ThisAdvBal);
              Print #RptHandle, Tab(40); QPTrim$(Opt3Desc); Tab(69); Using$("$##.##0.00", ThisOpt3Bal)
              LineCnt = LineCnt + 1
              If LineCnt >= MaxLines Then
                Print #RptHandle, FF$
                GoSub PrintHeader
                GoSub PrintCustHeader
              End If
            Else
              Print #RptHandle, Tab(5); "Advertising:  " + Using$("$##,##0.00", ThisAdvBal)
              LineCnt = LineCnt + 1
              If LineCnt >= MaxLines Then
                Print #RptHandle, FF$
                GoSub PrintHeader
                GoSub PrintCustHeader
              End If
            End If
            Print #RptHandle, Tab(5); "Late Listing: " + Using$("$##,##0.00", ThisLateListBal)
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintHeader
              GoSub PrintCustHeader
            End If
            Print #RptHandle, Tab(5); "Penalty:      " + Using$("$##,##0.00", ThisPenBal)
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintHeader
              GoSub PrintCustHeader
            End If
            
            If GYrCnt = 0 Then '
              GYrCnt = GYrCnt + 1
              ReDim Preserve GYears(1 To GYrCnt) As Integer
              GYears(GYrCnt) = TaxTrans.TaxYear
              ReDim Preserve GYearBal(1 To GYrCnt) As Double
              GYearBal(GYrCnt) = Balance#
              ReDim Preserve GPrincBal(1 To GYrCnt) As Double
              GPrincBal(GYrCnt) = ThisPrincBal#
              GPrincTot = ThisPrincBal#
              ReDim Preserve GIntBal(1 To GYrCnt) As Double
              GIntBal(GYrCnt) = ThisIntBal#
              GIntTot = ThisIntBal#
              ReDim Preserve GAdvBal(1 To GYrCnt) As Double
              GAdvBal(GYrCnt) = ThisAdvBal#
              GAdvTot = ThisAdvBal#
              ReDim Preserve GLateListBal(1 To GYrCnt) As Double
              GLateListBal(GYrCnt) = ThisLateListBal#
              GLateListTot = ThisLateListBal#
              ReDim Preserve GPenBal(1 To GYrCnt) As Double
              GPenBal(GYrCnt) = ThisPenBal#
              GPenTot = ThisPenBal#
              ReDim Preserve GOPt1Bal(1 To GYrCnt) As Double
              GOPt1Bal(GYrCnt) = ThisOpt1Bal#
              GOpt1Tot = ThisOpt1Bal#
              ReDim Preserve GOPt2Bal(1 To GYrCnt) As Double
              GOPt2Bal(GYrCnt) = ThisOpt2Bal#
              GOpt2Tot = ThisOpt2Bal#
              ReDim Preserve GOPt3Bal(1 To GYrCnt) As Double
              GOPt3Bal(GYrCnt) = ThisOpt3Bal#
              GOpt3Tot = ThisOpt3Bal#
            Else
              For y = 1 To GYrCnt
                If GYears(y) = TaxTrans.TaxYear Then '
                  GYearBal(y) = OldRound(GYearBal(y) + Balance#)
                  GPrincBal(y) = OldRound(GPrincBal(y) + ThisPrincBal#)
                  GPrincTot = OldRound(GPrincTot# + ThisPrincBal#)
                  GIntBal(y) = OldRound(GIntBal(y) + ThisIntBal#)
                  GIntTot = OldRound(GIntTot# + ThisIntBal#)
                  GAdvBal(y) = OldRound(GAdvBal(y) + ThisAdvBal#)
                  GAdvTot = OldRound(GAdvTot# + ThisAdvBal#)
                  GLateListBal(y) = OldRound(GLateListBal(y) + ThisLateListBal#)
                  GLateListTot = OldRound(GLateListTot# + ThisLateListBal#)
                  GPenBal(y) = OldRound(GPenBal(y) + ThisPenBal#)
                  GPenTot = OldRound(GPenTot# + ThisPenBal#)
                  GOPt1Bal(y) = OldRound(GOPt1Bal(y) + ThisOpt1Bal#)
                  GOpt1Tot = OldRound(GOpt1Tot# + ThisOpt1Bal#)
                  GOPt2Bal(y) = OldRound(GOPt2Bal(y) + ThisOpt2Bal#)
                  GOpt2Tot = OldRound(GOpt2Tot# + ThisOpt2Bal#)
                  GOPt3Bal(y) = OldRound(GOPt3Bal(y) + ThisOpt3Bal#)
                  GOpt3Tot = OldRound(GOpt3Tot# + ThisOpt3Bal#)
                  Exit For
                End If '
              Next y
              If y > GYrCnt Then '
                GYrCnt = GYrCnt + 1
                ReDim Preserve GYears(1 To GYrCnt) As Integer
                GYears(GYrCnt) = TaxTrans.TaxYear
                ReDim Preserve GYearBal(1 To GYrCnt) As Double
                GYearBal(GYrCnt) = Balance#
                ReDim Preserve GPrincBal(1 To GYrCnt) As Double
                GPrincBal(GYrCnt) = ThisPrincBal#
                GPrincTot# = OldRound(GPrincTot# + ThisPrincBal#)
                ReDim Preserve GIntBal(1 To GYrCnt) As Double
                GIntBal(GYrCnt) = ThisIntBal#
                GIntTot# = OldRound(GIntTot# + ThisIntBal#)
                ReDim Preserve GAdvBal(1 To GYrCnt) As Double
                GAdvBal(GYrCnt) = ThisAdvBal#
                GAdvTot# = OldRound(GAdvTot# + ThisAdvBal#)
                ReDim Preserve GLateListBal(1 To GYrCnt) As Double
                GLateListBal(GYrCnt) = ThisLateListBal#
                GLateListTot# = OldRound(GLateListTot + ThisLateListBal#)
                ReDim Preserve GPenBal(1 To GYrCnt) As Double
                GPenBal(GYrCnt) = ThisPenBal#
                GPenTot# = OldRound(GPenTot + ThisPenBal#)
                ReDim Preserve GOPt1Bal(1 To GYrCnt) As Double
                GOPt1Bal(GYrCnt) = ThisOpt1Bal#
                GOpt1Tot# = OldRound(GOpt1Tot# + ThisOpt1Bal#)
                ReDim Preserve GOPt2Bal(1 To GYrCnt) As Double
                GOPt2Bal(GYrCnt) = ThisOpt2Bal#
                GOpt2Tot# = OldRound(GOpt2Tot# + ThisOpt2Bal#)
                ReDim Preserve GOPt3Bal(1 To GYrCnt) As Double
                GOPt3Bal(GYrCnt) = ThisOpt3Bal#
                GOpt3Tot# = OldRound(GOpt3Tot# + ThisOpt3Bal#)
              End If '
            End If '
          End If
        End If
SkipIt:
        ThisRec = TaxTrans.LastTrans
      Loop
    End If
Inactive:
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
  
  BigYr = 0
  For x = 1 To GYrCnt
    If GYears(x) > BigYr Then
      BigYr = GYears(x)
    End If
  Next x
  
  ThisBigYr = BigYr + 1
  Nextz = 1
  Do While Nextz <= GYrCnt
    For z = Nextz To GYrCnt
      If GYears(z) < ThisBigYr Then
         ThisBigYr = GYears(z)
         Thisz = z
      End If
    Next z
    HoldYr = GYears(Nextz)
    HoldBal = GYearBal(Nextz)
    HoldPrinc = GPrincBal(Nextz)
    HoldInt = GIntBal(Nextz)
    HoldAdv = GAdvBal(Nextz)
    HoldLateList = GLateListBal(Nextz)
    HoldPen = GPenBal(Nextz)
    HoldOpt1 = GOPt1Bal(Nextz)
    HoldOpt2 = GOPt2Bal(Nextz)
    HoldOpt3 = GOPt3Bal(Nextz)
    GYears(Nextz) = GYears(Thisz)
    GYearBal(Nextz) = GYearBal(Thisz)
    GPrincBal(Nextz) = GPrincBal(Thisz)
    GIntBal(Nextz) = GIntBal(Thisz)
    GAdvBal(Nextz) = GAdvBal(Thisz)
    GLateListBal(Nextz) = GLateListBal(Thisz)
    GPenBal(Nextz) = GPenBal(Thisz)
    GOPt1Bal(Nextz) = GOPt1Bal(Thisz)
    GOPt2Bal(Nextz) = GOPt2Bal(Thisz)
    GOPt3Bal(Nextz) = GOPt3Bal(Thisz)
    GYears(Thisz) = HoldYr
    GYearBal(Thisz) = HoldBal
    GPrincBal(Thisz) = HoldPrinc
    GIntBal(Thisz) = HoldInt
    GAdvBal(Thisz) = HoldAdv
    GLateListBal(Thisz) = HoldLateList
    GPenBal(Thisz) = HoldPen
    GOPt1Bal(Thisz) = HoldOpt1
    GOPt2Bal(Thisz) = HoldOpt2
    GOPt3Bal(Thisz) = HoldOpt3
    
    Nextz = Nextz + 1
    ThisBigYr = BigYr + 1
  Loop
  
  If InStr(CStr(OverPay), "E") Or fpcmbTaxYear.Text <> "All" Then OverPay = 0
  Print #RptHandle, FF$
  GoSub PrintEndHeader
  Print #RptHandle, "Total Entries: " + Using$("####0", TransCnt)
  Print #RptHandle, "Total Tax Balance: "; Tab(30); Using$("$###,###,##0.00", GBal)
  If fpcmbTaxYear.Text = "All" Then
    Print #RptHandle, "Over Payment:      "; Tab(30); Using$("$###,###,##0.00", OverPay)
  Else
    Print #RptHandle,
  End If
  Print #RptHandle,
  Print #RptHandle, "Principle Total:          "; Tab(30); Using$("$###,###,##0.00", GPrincTot)
  Print #RptHandle, "Interest Total:           "; Tab(30); Using$("$###,###,##0.00", GIntTot)
  Print #RptHandle, "Advertising Total:        "; Tab(30); Using$("$###,###,##0.00", GAdvTot)
  Print #RptHandle, "Late Listing Total:       "; Tab(30); Using$("$###,###,##0.00", GLateListTot)
  Print #RptHandle, "Penalty Total:            "; Tab(30); Using$("$###,###,##0.00", GPenTot)
  LineCnt = LineCnt + 8
  If QPTrim$(Opt1Desc) <> "" Then
    Print #RptHandle, QPTrim$(Opt1Desc) + ": "; Tab(30); Using$("$###,###,##0.00", GOpt1Tot)
    LineCnt = LineCnt + 1
  End If
  If QPTrim$(Opt2Desc) <> "" Then
    Print #RptHandle, QPTrim$(Opt2Desc) + ": "; Tab(30); Using$("$###,###,##0.00", GOpt2Tot)
    LineCnt = LineCnt + 1
  End If
  If QPTrim$(Opt3Desc) <> "" Then
    Print #RptHandle, QPTrim$(Opt3Desc) + ": "; Tab(30); Using$("$###,###,##0.00", GOpt3Tot)
    LineCnt = LineCnt + 1
  End If
  Print #RptHandle,
  Print #RptHandle,
  LineCnt = LineCnt + 2
  
  For x = 1 To GYrCnt
    Print #RptHandle, "Tax Year"; Tab(34); "Amount Owed"
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintEndHeader
      Print #RptHandle, "Tax Year"; Tab(34); "Amount Owed" '; Tab(60); "Over Payments"
    End If
    Print #RptHandle, Using$("###0", GYears(x)); Tab(30); Using$("$###,###,##0.00", GYearBal(x))
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintEndHeader
      Print #RptHandle, "Tax Year" '; Tab(30); "Amount Owed" '; Tab(60); "Over Payments"
    End If
'    End If
    Print #RptHandle, Tab(5); "Principle:"; Tab(30); Using$("$###,###,##0.00", GPrincBal(x))
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintEndHeader
      Print #RptHandle, "Tax Year" '; Tab(30); "Amount Owed" '; Tab(60); "Over Payments"
    End If
    Print #RptHandle, Tab(5); "Interest: "; Tab(30); Using$("$###,###,##0.00", GIntBal(x))
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintEndHeader
      Print #RptHandle, "Tax Year" '; Tab(30); "Amount Owed" '; Tab(60); "Over Payments"
    End If
    Print #RptHandle, Tab(5); "Advertising:"; Tab(30); Using$("$###,###,##0.00", GAdvBal(x))
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintEndHeader
      Print #RptHandle, "Tax Year" '; Tab(30); "Amount Owed" '; Tab(60); "Over Payments"
    End If
    Print #RptHandle, Tab(5); "Late Listing:"; Tab(30); Using$("$###,###,##0.00", GLateListBal(x))
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintEndHeader
      Print #RptHandle, "Tax Year" '; Tab(30); "Amount Owed" '; Tab(60); "Over Payments"
    End If
    Print #RptHandle, Tab(5); "Penalty:"; Tab(30); Using$("$###,###,##0.00", GPenBal(x))
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintEndHeader
      Print #RptHandle, "Tax Year" '; Tab(30); "Amount Owed" '; Tab(60); "Over Payments"
    End If
    If QPTrim$(Opt1Desc) <> "" Then
      Print #RptHandle, Tab(5); Opt1Desc + ":"; Tab(30); Using$("$###,###,##0.00", GOPt1Bal(x))
      LineCnt = LineCnt + 1
      If LineCnt >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintEndHeader
        Print #RptHandle, "Tax Year" '; Tab(30); "Amount Owed" '; Tab(60); "Over Payments"
      End If
    End If
    If QPTrim$(Opt2Desc) <> "" Then
      Print #RptHandle, Tab(5); Opt2Desc + ":"; Tab(30); Using$("$###,###,##0.00", GOPt2Bal(x))
      LineCnt = LineCnt + 1
      If LineCnt >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintEndHeader
        Print #RptHandle, "Tax Year" '; Tab(30); "Amount Owed" '; Tab(60); "Over Payments"
      End If
    End If
    If QPTrim$(Opt3Desc) <> "" Then
      Print #RptHandle, Tab(5); Opt3Desc + ":"; Tab(30); Using$("$###,###,##0.00", GOPt3Bal(x))
      LineCnt = LineCnt + 1
      If LineCnt >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintEndHeader
        Print #RptHandle, "Tax Year" '; Tab(30); "Amount Owed" '; Tab(60); "Over Payments"
      End If
    End If
    Print #RptHandle,
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintEndHeader
      Print #RptHandle, "Tax Year" '; Tab(30); "Amount Owed"; Tab(60); "Over Payments"
    End If
  Next x
  
  Print #RptHandle, FF$
  Close
        
  If GYrCnt = 0 Then
    Call TaxMsg(900, "There are no balances to report for the parameters entered.")
    Exit Sub
  End If
  ViewPrint RptFile, "Tax Master Balance Listing", True
  
  Exit Sub

PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(22); "Tax Master Balance Listing Detail - Real"
  Print #RptHandle, Town
  If Len(CStr(Page)) = 2 Then
    Print #RptHandle, "Report Date: " + CStr(Now); Tab(71); "Page # " + CStr(Page)
  ElseIf Len(CStr(Page)) = 3 Then
    Print #RptHandle, "Report Date: " + CStr(Now); Tab(70); "Page # " + CStr(Page)
  ElseIf Len(CStr(Page)) = 4 Then
    Print #RptHandle, "Report Date: " + CStr(Now); Tab(69); "Page # " + CStr(Page)
  Else
    Print #RptHandle, "Report Date: " + CStr(Now); Tab(72); "Page # " + CStr(Page)
  End If
  If ActiveFlag = "B" Then
    Print #RptHandle, "Customer Status: " + "Active And Inactive"; Tab(65); "(I) = Inactive"
  Else
    Print #RptHandle, "Customer Status: " + fpcmbIncInactive.Text
  End If
  Print #RptHandle, "Acct #"; Tab(8); "Customer Name"; Tab(59); "Tax Year"; Tab(73); "Balance"
  Print #RptHandle, String(79, "-")
  LineCnt = 6
  
  Return
  
PrintCustHeader:
  If LineCnt <> 6 Then
    Print #RptHandle, String(79, "-")
    LineCnt = LineCnt + 1
  End If
  Print #RptHandle, Using$("####0", CustRec); Tab(8); CustName; Tab(55); "Cust Balance: " + QPTrim$(Using$("$###,###,##0.00", CustTotBal))
  Print #RptHandle, String(79, "-")
  LineCnt = LineCnt + 2
  Return

PrintEndHeader:
  Page = Page + 1
  Print #RptHandle, Tab(25); "Tax Master Balance Listing Detail"
  Print #RptHandle, "Real Only"
  Print #RptHandle, Town
  If Len(CStr(Page)) = 2 Then
    Print #RptHandle, "Report Date: " + CStr(Now); Tab(71); "Page # " + CStr(Page)
  ElseIf Len(CStr(Page)) = 3 Then
    Print #RptHandle, "Report Date: " + CStr(Now); Tab(70); "Page # " + CStr(Page)
  ElseIf Len(CStr(Page)) = 4 Then
    Print #RptHandle, "Report Date: " + CStr(Now); Tab(69); "Page # " + CStr(Page)
  Else
    Print #RptHandle, "Report Date: " + CStr(Now); Tab(72); "Page # " + CStr(Page)
  End If

  Print #RptHandle, String(79, "-")
  LineCnt = 5
  
  Return
  
GetYearBalance:
  ThisRec = TaxCust.LastTrans
  TestBal = 0
  CreditUsed = 0
  OverPaid = 0
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
    If CInt(fpcmbTaxYear.Text) <> TaxTrans.TaxYear Then GoTo SkipItYear
    If TaxTrans.BillType <> "R" Then GoTo SkipItYear
    If TaxTrans.TranType = 10 Then 'added 10/3/06 adjust bill down affecting credit
       TestBal = TestBal + TaxTrans.Revenue.PrePaidUsed
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 11 Then 'added 10/3/06 prepay adjust down
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 24 Then 'added 10/3/06 adjust bill up affecting credit
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 21 Or TaxTrans.TranType = 22 Then
'      TestBal# = OldRound#(TestBal# - TaxTrans.Revenue.PrePaidAmt)'remmed on 2/2/07
      OverPaid = OldRound(OverPaid + TaxTrans.Revenue.PrePaidAmt) 'added 2/2/07
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 12 Then 'refund on prepay 7/17/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 9 Then 'credit applied at billing  'added 2/2/07
      CreditUsed = OldRound(CreditUsed + TaxTrans.Revenue.PrePaidUsed)
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 1 Then
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt + TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
    End If
SkipItYear:
    ThisRec = TaxTrans.LastTrans
  Loop
  If OverPaid = 0 Then CreditUsed = 0 'added 2/20/07
  
  TestBal = OldRound(TestBal - (OverPaid - CreditUsed)) 'added 2/2/07
  
  Return
  
GetAllBalance:
  ThisRec = TaxCust.LastTrans
  TestBal = 0
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
    If TaxTrans.BillType <> "R" Then GoTo DoAgain
    If TaxTrans.TranType = 10 Then 'added 10/3/06 adjust bill down affecting credit
       TestBal = TestBal + TaxTrans.Revenue.PrePaidUsed
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 11 Then 'added 10/3/06 prepay adjust down
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 24 Then 'added 10/3/06 adjust bill up affecting credit
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 21 Or TaxTrans.TranType = 22 Then
      TestBal# = OldRound#(TestBal# - TaxTrans.Revenue.PrePaidAmt)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 12 Then 'refund on prepay 7/17/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 9 Then 'added 8/11/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 1 Then
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt + TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
    End If
DoAgain:
    ThisRec = TaxTrans.LastTrans
  Loop
  
  Return

ApplyDisc:
  Disc1 = 0
  Disc2 = 0
  Disc3 = 0
  Disc4 = 0
  Disc5 = 5
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
  Disc5 = OldRound(TaxTrans.Revenue.LateListPd / SaveAmt)
  Disc5 = OldRound(Disc5 * TaxTrans.DiscAmt)
  TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd + Disc1 + Disc5)
  TaxTrans.Revenue.RevOpt1Pd = OldRound(TaxTrans.Revenue.RevOpt1Pd + Disc2)
  TaxTrans.Revenue.RevOpt2Pd = OldRound(TaxTrans.Revenue.RevOpt2Pd + Disc3)
  TaxTrans.Revenue.RevOpt3Pd = OldRound(TaxTrans.Revenue.RevOpt3Pd + Disc4)
  Dif = OldRound(Disc1 + Disc2 + Disc3 + Disc4 + Disc5) 'added 9/19/07
  If Dif <> TaxTrans.DiscAmt Then
   If Disc1 > 0 Or Disc5 > 0 Then
     TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd + Dif)
   ElseIf Disc2 > 0 Then
     TaxTrans.Revenue.RevOpt1Pd = OldRound(TaxTrans.Revenue.RevOpt1Pd + Dif)
   ElseIf Disc3 > 0 Then
     TaxTrans.Revenue.RevOpt2Pd = OldRound(TaxTrans.Revenue.RevOpt2Pd + Dif)
   ElseIf Disc4 > 0 Then
     TaxTrans.Revenue.RevOpt3Pd = OldRound(TaxTrans.Revenue.RevOpt3Pd + Dif)
   End If
  End If
  DiscApplied = True
  
  Return

ERRORSTUFF:
   Unload frmVATaxShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxMasterBalList", "PrintRTextDet", Erl)
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

Private Sub PrintPGraphicsDet()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long, y As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim Balance As Double
  Dim ThisRec As Long
  Dim ThisYear As Integer
  Dim GTCnt As Long
  Dim TCnt As Integer
  Dim CustName$
  Dim YrCnt As Integer
  Dim GYrCnt As Integer
  Dim OverPay As Double
  Dim HoldYr As Integer
  Dim HoldBal As Double
  Dim Nextz As Integer
  Dim z As Integer
  Dim Thisz As Integer
  Dim BigYr As Integer
  Dim ThisBigYr As Integer
  Dim GBal As Double
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim RptFile$
  Dim RptHandle As Integer
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim dlm$
  Dim CustRec As Long
  Dim InactiveFlag As Boolean
  Dim SubRptFile$
  Dim SubRptHandle As Integer
  Dim TransCnt As Long
  Dim OP As Double
  Dim ThisPersBal As Double
  Dim ThisIntBal As Double
  Dim ThisMTBal As Double
  Dim ThisMCBal As Double
  Dim ThisFEBal As Double
  Dim ThisMHBal As Double
  Dim ThisPenBal As Double
  Dim ThisOpt1Bal As Double
  Dim ThisOpt2Bal As Double
  Dim ThisOpt3Bal As Double
  Dim HoldPers As Double
  Dim HoldInt As Double
  Dim HoldMT As Double
  Dim HoldMC As Double
  Dim HoldFE As Double
  Dim HoldMH As Double
  Dim HoldPen As Double
  Dim HoldOpt1 As Double
  Dim HoldOpt2 As Double
  Dim HoldOpt3 As Double
  Dim GPersTot As Double
  Dim GIntTot As Double
  Dim GMTTot As Double
  Dim GMCTot As Double
  Dim GFETot As Double
  Dim GMHTot As Double
  Dim GPenTot As Double
  Dim GOpt1Tot As Double
  Dim GOpt2Tot As Double
  Dim GOpt3Tot As Double
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim ActiveFlag$
  Dim PinCnt As Integer
  Dim ThisPin$
  Dim PropType$
  Dim CustTotBal As Double
  Dim ThisBillNum$
  Dim ThisOP$
  Dim TestBal#
  Dim Disc1 As Double '1/16/2007
  Dim Disc2 As Double '1/16/2007
  Dim Disc3 As Double '1/16/2007
  Dim Disc4 As Double '1/16/2007
  Dim Disc5 As Double '1/16/2007
  Dim Disc6 As Double '1/16/2007
  Dim Disc7 As Double '1/16/2007
  Dim Disc8 As Double '1/16/2007
  Dim Disc9 As Double '9/17/2007
  Dim Dif As Double '9/19/07
  Dim DiscApplied As Boolean '1/16/2007
  Dim SaveAmt As Double '1/16/2007
  Dim CreditUsed As Double '2/2/07
  Dim OverPaid As Double '2/2/07
  
  On Error GoTo ERRORSTUFF
  
  IdxFlag = False
  ActiveFlag = "B"
  If QPTrim$(fpcmbIncInactive.Text) = "Active Only" Then
    ActiveFlag = "A"
  ElseIf QPTrim$(fpcmbIncInactive.Text) = "Inactive Only" Then
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

  RptFile$ = "TAXRPTS\TXPMSTBALDET.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  ReDim GYearBal(1 To 1) As Double
  ReDim GYears(1 To 1) As Integer
  ReDim GPersBal(1 To 1) As Double
  ReDim GMTBal(1 To 1) As Double
  ReDim GMCBal(1 To 1) As Double
  ReDim GFEBal(1 To 1) As Double
  ReDim GMHBal(1 To 1) As Double
  ReDim GIntBal(1 To 1) As Double
  ReDim GPenBal(1 To 1) As Double
  ReDim GOPt1Bal(1 To 1) As Double
  ReDim GOPt2Bal(1 To 1) As Double
  ReDim GOPt3Bal(1 To 1) As Double
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  GYrCnt = 0
  frmVATaxShowPctComp.Label1 = "Gathering Tax Transaction Data"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  
'  If fpcmbTaxYear.Text = "All" Then
'    OverPay = GetOverPayAmount(-1, "P")
'  Else
'    OverPay = GetOverPayAmount(CInt(fpcmbTaxYear.Text), "P")
'  End If
  
  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  End If
  
  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
      CustRec = IdxArray(x)
    Else
      Get TCHandle, x, TaxCust
      CustRec = x
    End If
    Balance = 0
    If TaxCust.Deleted <> 0 Then GoTo Inactive
    If ActiveFlag = "A" And TaxCust.Active = "N" Then
      GoTo Inactive
    ElseIf ActiveFlag = "I" And TaxCust.Active = "Y" Then
      GoTo Inactive
    End If
    
    If ActiveFlag = "B" Then
      If TaxCust.Active = "N" Then
        CustName = QPTrim$(TaxCust.CustName) + "(I)"
      Else
        CustName = QPTrim$(TaxCust.CustName)
      End If
    Else
      CustName = QPTrim$(TaxCust.CustName)
    End If
    If fpcmbTaxYear.Text = "All" Then
      GoSub GetAllBalance
    Else
      GoSub GetYearBalance
    End If
    CustTotBal = TestBal#
    OP = TestBal#
    If InStr(CStr(CustTotBal), "E") Then CustTotBal = 0
    
'      CustTotBal = GetCustPersBalance(CustRec, -1)
    If TaxCust.LastTrans > 0 Then
      YrCnt = 0
      PinCnt = 0
      ReDim YearBal(1 To 1) As Double
      ReDim Years(1 To 1) As Integer
      ReDim PersBal(1 To 1) As Double
      ReDim MTBal(1 To 1) As Double
      ReDim MCBal(1 To 1) As Double
      ReDim FEBal(1 To 1) As Double
      ReDim MHBal(1 To 1) As Double
      ReDim IntBal(1 To 1) As Double
      ReDim PenBal(1 To 1) As Double
      ReDim Opt1Bal(1 To 1) As Double
      ReDim Opt2Bal(1 To 1) As Double
      ReDim Opt3Bal(1 To 1) As Double
      
      ThisRec = TaxCust.LastTrans
      ThisOP = CStr(OP)
      If InStr(ThisOP, "E") Then OP = 0
      If OP < 0 Then
        OverPay = OldRound(OverPay + Abs(OP))
      End If
      Do While ThisRec > 0
        Get TTHandle, ThisRec, TaxTrans
        ThisPersBal = 0 '9/1/06
        ThisMTBal = 0 '9/1/06
        ThisMCBal = 0 '9/1/06
        ThisFEBal = 0 '9/1/06
        ThisMHBal = 0 '9/1/06
        ThisIntBal = 0 '9/1/06
        ThisPenBal = 0 '9/1/06
        ThisOpt1Bal = 0 '9/1/06
        ThisOpt2Bal = 0 '9/1/06
        ThisOpt3Bal = 0 '9/1/06
        Balance# = 0 '9/1/06
        If CustTotBal < 0 Then
          Balance = CustTotBal
          If fpcmbTaxYear.Text <> "All" Then
            TaxTrans.TaxYear = CInt(fpcmbTaxYear.Text)
          End If
          TaxTrans.LastTrans = 0
          GoTo GoWithZero
        End If
        If TaxTrans.TranType = 1 And TaxTrans.BillType = "P" Then
          If fpcmbTaxYear.Text <> "All" Then
            If CInt(fpcmbTaxYear.Text) <> TaxTrans.TaxYear Then GoTo SkipIt
          End If
          If TaxTrans.DiscAmt > 0 Then GoSub ApplyDisc 'added 1/16/07
          ThisPersBal = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.PPTRARmvl - TaxTrans.PPTRADisc - TaxTrans.Revenue.Principle1Pd)
          ThisMTBal = OldRound(TaxTrans.Revenue.Principle2 - TaxTrans.Revenue.Principle2Pd)
          ThisMCBal = OldRound(TaxTrans.Revenue.Principle3 - TaxTrans.Revenue.Principle3Pd)
          ThisFEBal = OldRound(TaxTrans.Revenue.Principle4 - TaxTrans.Revenue.Principle4Pd)
          ThisMHBal = OldRound(TaxTrans.Revenue.Principle5 - TaxTrans.Revenue.Principle5Pd)
          ThisIntBal = OldRound(TaxTrans.Revenue.Interest - TaxTrans.Revenue.InterestPd)
          ThisPenBal = OldRound(TaxTrans.Revenue.Penalty - TaxTrans.Revenue.PenaltyPd)
          ThisOpt1Bal = OldRound(TaxTrans.Revenue.RevOpt1 - TaxTrans.Revenue.RevOpt1Pd)
          ThisOpt2Bal = OldRound(TaxTrans.Revenue.RevOpt2 - TaxTrans.Revenue.RevOpt2Pd)
          ThisOpt3Bal = OldRound(TaxTrans.Revenue.RevOpt3 - TaxTrans.Revenue.RevOpt3Pd)
          Balance# = OldRound#(TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
          Balance# = OldRound#(Balance# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3 + TaxTrans.PPTRARmvl)
          Balance# = OldRound#(Balance# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
          Balance# = OldRound#(Balance# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
'          Balance# = OldRound#(Balance# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt + TaxTrans.PPTRADisc))
          Balance# = OldRound#(Balance# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.PPTRADisc))  '1/16/07 took out .DiscAmt
          If Balance = 0 And chkZeroBal.Value = 1 Then
            GoTo GoWithZero
          End If
          If Balance <> 0 Then
GoWithZero:
            GBal = OldRound(GBal + Balance#)
            If QPTrim$(TaxTrans.RealPin) = "0" And QPTrim$(TaxTrans.PersPin) = "0" Then
              ThisPin = "0"
              PropType = "UNATTACHED"
            ElseIf QPTrim$(TaxTrans.RealPin) = "-1" Then
              ThisPin = "-1"
              PropType = "MOCK"
            ElseIf QPTrim$(TaxTrans.PersPin) <> "0" And QPTrim$(TaxTrans.PersPin) <> "" Then
              ThisPin = QPTrim$(TaxTrans.PersPin)
              PropType = "PERSONAL"
            Else
              ThisPin = "0"
              PropType = "UNATTACHED"
            End If
            
            If GYrCnt = 0 Then
              GYrCnt = GYrCnt + 1
              ReDim Preserve GYears(1 To GYrCnt) As Integer
              GYears(GYrCnt) = TaxTrans.TaxYear
              ReDim Preserve GYearBal(1 To GYrCnt) As Double
              GYearBal(GYrCnt) = Balance#
              ReDim Preserve GPersBal(1 To GYrCnt) As Double
              GPersBal(GYrCnt) = ThisPersBal#
              GPersTot = ThisPersBal#
              ReDim Preserve GIntBal(1 To GYrCnt) As Double
              GIntBal(GYrCnt) = ThisIntBal#
              GIntTot = ThisIntBal#
              
              ReDim Preserve GMTBal(1 To GYrCnt) As Double
              GMTBal(GYrCnt) = ThisMTBal#
              GMTTot = ThisMTBal#
              ReDim Preserve GMCBal(1 To GYrCnt) As Double
              GMCBal(GYrCnt) = ThisMCBal#
              GMCTot = ThisMCBal#
              
              ReDim Preserve GFEBal(1 To GYrCnt) As Double
              GFEBal(GYrCnt) = ThisFEBal#
              GFETot = ThisFEBal#
              ReDim Preserve GMHBal(1 To GYrCnt) As Double
              GMHBal(GYrCnt) = ThisMHBal#
              GMHTot = ThisMHBal#
              ReDim Preserve GPenBal(1 To GYrCnt) As Double
              GPenBal(GYrCnt) = ThisPenBal#
              GPenTot = ThisPenBal#
              
              ReDim Preserve GOPt1Bal(1 To GYrCnt) As Double
              GOPt1Bal(GYrCnt) = ThisOpt1Bal#
              GOpt1Tot = ThisOpt1Bal#
              ReDim Preserve GOPt2Bal(1 To GYrCnt) As Double
              GOPt2Bal(GYrCnt) = ThisOpt2Bal#
              GOpt2Tot = ThisOpt2Bal#
              ReDim Preserve GOPt3Bal(1 To GYrCnt) As Double
              GOPt3Bal(GYrCnt) = ThisOpt3Bal#
              GOpt3Tot = ThisOpt3Bal#
            Else
              For y = 1 To GYrCnt
                If GYears(y) = TaxTrans.TaxYear Then '
                  GYearBal(y) = OldRound(GYearBal(y) + Balance#)
                  GPersBal(y) = OldRound(GPersBal(y) + ThisPersBal#)
                  GPersTot = OldRound(GPersTot# + ThisPersBal#)
                  GIntBal(y) = OldRound(GIntBal(y) + ThisIntBal#)
                  GIntTot = OldRound(GIntTot# + ThisIntBal#)
                  GMTBal(y) = OldRound(GMTBal(y) + ThisMTBal#)
                  GMTTot = OldRound(GMTTot# + ThisMTBal#)
                  GMCBal(y) = OldRound(GMCBal(y) + ThisMCBal#)
                  GMCTot = OldRound(GMCTot# + ThisMCBal#)
                  
                  GFEBal(y) = OldRound(GFEBal(y) + ThisFEBal#)
                  GFETot = OldRound(GFETot# + ThisFEBal#)
                  GMHBal(y) = OldRound(GMHBal(y) + ThisMHBal#)
                  GMHTot = OldRound(GMHTot# + ThisMHBal#)
                  GPenBal(y) = OldRound(GPenBal(y) + ThisPenBal#)
                  GPenTot = OldRound(GPenTot# + ThisPenBal#)
                  
                  GOPt1Bal(y) = OldRound(GOPt1Bal(y) + ThisOpt1Bal#)
                  GOpt1Tot = OldRound(GOpt1Tot# + ThisOpt1Bal#)
                  GOPt2Bal(y) = OldRound(GOPt2Bal(y) + ThisOpt2Bal#)
                  GOpt2Tot = OldRound(GOpt2Tot# + ThisOpt2Bal#)
                  GOPt3Bal(y) = OldRound(GOPt3Bal(y) + ThisOpt3Bal#)
                  GOpt3Tot = OldRound(GOpt3Tot# + ThisOpt3Bal#)
                  Exit For
                End If '
              Next y
              If y > GYrCnt Then '
                GYrCnt = GYrCnt + 1
                ReDim Preserve GYears(1 To GYrCnt) As Integer
                GYears(GYrCnt) = TaxTrans.TaxYear
                ReDim Preserve GYearBal(1 To GYrCnt) As Double
                GYearBal(GYrCnt) = Balance#
                ReDim Preserve GPersBal(1 To GYrCnt) As Double
                GPersBal(GYrCnt) = ThisPersBal#
                GPersTot# = OldRound(GPersTot# + ThisPersBal#)
                
                ReDim Preserve GIntBal(1 To GYrCnt) As Double
                GIntBal(GYrCnt) = ThisIntBal#
                GIntTot# = OldRound(GIntTot# + ThisIntBal#)
                ReDim Preserve GMTBal(1 To GYrCnt) As Double
                GMTBal(GYrCnt) = ThisMTBal#
                GMTTot# = OldRound(GMTTot# + ThisMTBal#)
                ReDim Preserve GMCBal(1 To GYrCnt) As Double
                GMCBal(GYrCnt) = ThisMCBal#
                GMCTot# = OldRound(GMCTot + ThisMCBal#)
                
                ReDim Preserve GFEBal(1 To GYrCnt) As Double
                GFEBal(GYrCnt) = ThisFEBal#
                GFETot# = OldRound(GFETot# + ThisFEBal#)
                ReDim Preserve GMHBal(1 To GYrCnt) As Double
                GMHBal(GYrCnt) = ThisMHBal#
                GMHTot# = OldRound(GMHTot + ThisMHBal#)
                ReDim Preserve GPenBal(1 To GYrCnt) As Double
                GPenBal(GYrCnt) = ThisPenBal#
                GPenTot# = OldRound(GPenTot# + ThisPenBal#)
                
                ReDim Preserve GOPt1Bal(1 To GYrCnt) As Double
                GOPt1Bal(GYrCnt) = ThisOpt1Bal#
                GOpt1Tot# = OldRound(GOpt1Tot# + ThisOpt1Bal#)
                ReDim Preserve GOPt2Bal(1 To GYrCnt) As Double
                GOPt2Bal(GYrCnt) = ThisOpt2Bal#
                GOpt2Tot# = OldRound(GOpt2Tot# + ThisOpt2Bal#)
                ReDim Preserve GOPt3Bal(1 To GYrCnt) As Double
                GOPt3Bal(GYrCnt) = ThisOpt3Bal#
                GOpt3Tot# = OldRound(GOpt3Tot# + ThisOpt3Bal#)
              End If '
            End If '
            TransCnt = TransCnt + 1
            ThisRec = ThisRec
            ThisBillNum = ParseBillNum(TaxTrans.Description)
            '                   0            1             2                 3                  4              5            6
            Print #RptHandle, Town; dlm; CustName; dlm; CustRec; dlm; TaxTrans.TaxYear; dlm; Balance#; dlm; TransCnt; dlm; GBal; dlm;
            '                         7                8              9                10               11                   12                13
            Print #RptHandle, ThisPersBal#; dlm; ThisIntBal#; dlm; ThisMTBal#; dlm; ThisMCBal#; dlm; ThisOpt1Bal#; dlm; ThisOpt2Bal#; dlm; ThisOpt3Bal#; dlm;
            '                     14             15            16              17                18            19             20
            Print #RptHandle, GPersTot; dlm; GIntTot; dlm; GMTTot; dlm; GMCTot; dlm; GOpt1Tot; dlm; GOpt2Tot; dlm; GOpt3Tot; dlm;
            '                    21              22              23
            Print #RptHandle, POpt1Desc; dlm; POpt2Desc; dlm; POpt3Desc; dlm;
            If UseOpt = "Y" Then
              '                    24                      25                        26
              Print #RptHandle, ThisOpt; dlm; QPTrim$(TaxCust.OptSrchDesc); dlm; ActiveFlag; dlm;
            Else
              '                 24       25           26
              Print #RptHandle, ""; dlm; ""; dlm; ActiveFlag; dlm;
            End If
            '                    27            28              29               30             31              32             33              34              35
            Print #RptHandle, PropType; dlm; ThisPin; dlm; CustTotBal; dlm; ThisFEBal#; dlm; GFETot#; dlm; ThisMHBal#; dlm; GMHTot#; dlm; ThisPenBal#; dlm; GPenTot#; dlm;
            If fpcmbTaxYear.Text = "All" Then 'added All 12/5/06
              '                     36               37                38
              Print #RptHandle, ThisBillNum; dlm; OverPay; dlm; fpcmbTaxYear.Text
            Else
               '                     36           37            38
              Print #RptHandle, ThisBillNum; dlm; 0; dlm; fpcmbTaxYear.Text
           End If
          End If
        End If
SkipIt:
        ThisRec = TaxTrans.LastTrans
      Loop
    End If
Inactive:
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
  
  BigYr = 0
  For x = 1 To GYrCnt
    If GYears(x) > BigYr Then
      BigYr = GYears(x)
    End If
  Next x
  Close
  
  ThisBigYr = BigYr + 1
  Nextz = 1
  Do While Nextz <= GYrCnt
    For z = Nextz To GYrCnt
      If GYears(z) < ThisBigYr Then
         ThisBigYr = GYears(z)
         Thisz = z
      End If
    Next z
    HoldYr = GYears(Nextz)
    HoldBal = GYearBal(Nextz)
    HoldPers = GPersBal(Nextz)
    HoldMT = GMTBal(Nextz)
    HoldMC = GMCBal(Nextz)
    HoldFE = GFEBal(Nextz)
    HoldMH = GMHBal(Nextz)
    HoldInt = GIntBal(Nextz)
    HoldPen = GPenBal(Nextz)
    HoldOpt1 = GOPt1Bal(Nextz)
    HoldOpt2 = GOPt2Bal(Nextz)
    HoldOpt3 = GOPt3Bal(Nextz)
    GYears(Nextz) = GYears(Thisz)
    GYearBal(Nextz) = GYearBal(Thisz)
    GPersBal(Nextz) = GPersBal(Thisz)
    GMTBal(Nextz) = GMTBal(Thisz)
    GMCBal(Nextz) = GMCBal(Thisz)
    GFEBal(Nextz) = GFEBal(Thisz)
    GMHBal(Nextz) = GMHBal(Thisz)
    GIntBal(Nextz) = GIntBal(Thisz)
    GPenBal(Nextz) = GPenBal(Thisz)
    GOPt1Bal(Nextz) = GOPt1Bal(Thisz)
    GOPt2Bal(Nextz) = GOPt2Bal(Thisz)
    GOPt3Bal(Nextz) = GOPt3Bal(Thisz)
    GYears(Thisz) = HoldYr
    GYearBal(Thisz) = HoldBal
    GPersBal(Thisz) = HoldPers
    GMTBal(Thisz) = HoldMT
    GMCBal(Thisz) = HoldMC
    GFEBal(Thisz) = HoldFE
    GMHBal(Thisz) = HoldMH
    GIntBal(Thisz) = HoldInt
    GPenBal(Thisz) = HoldPen
    GOPt1Bal(Thisz) = HoldOpt1
    GOPt2Bal(Thisz) = HoldOpt2
    GOPt3Bal(Thisz) = HoldOpt3
    Nextz = Nextz + 1
    ThisBigYr = BigYr + 1
  Loop
  
  SubRptFile$ = "TAXRPTS\TXPMSTBALSUBDET.RPT"
  SubRptHandle = FreeFile
  Open SubRptFile For Output As #SubRptHandle
  
  
  If InStr(CStr(OverPay), "E") Or fpcmbTaxYear.Text <> "All" Then OverPay = 0 'added All on 12/5/06
  For x = 1 To GYrCnt
    If x = GYrCnt Then
      '                        0               1                2          3
      Print #SubRptHandle, GYears(x); dlm; GYearBal(x); dlm; OverPay; dlm; 1; dlm;
      '                         4                 5                6                   7
      Print #SubRptHandle, GPersBal(x); dlm; GIntBal(x); dlm; GMTBal(x); dlm; GMCBal(x); dlm;
      '                         8                 9                10               11
      Print #SubRptHandle, GOPt1Bal(x); dlm; GOPt2Bal(x); dlm; GOPt3Bal(x); dlm; POpt1Desc; dlm;
      '                        12            13              14              15              16
      Print #SubRptHandle, POpt2Desc; dlm; POpt3Desc; dlm; GFEBal(x); dlm; GMHBal(x); dlm; GPenBal(x)
    Else
      '                        0               1                2
      Print #SubRptHandle, GYears(x); dlm; GYearBal(x); dlm; OverPay; dlm; 0; dlm;
      '                         4                 5                6                   7
      Print #SubRptHandle, GPersBal(x); dlm; GIntBal(x); dlm; GMTBal(x); dlm; GMCBal(x); dlm;
      '                         8                 9                10                11
      Print #SubRptHandle, GOPt1Bal(x); dlm; GOPt2Bal(x); dlm; GOPt3Bal(x); dlm; POpt1Desc; dlm;
      '                        12            13               14             15              16
      Print #SubRptHandle, POpt2Desc; dlm; POpt3Desc; dlm; GFEBal(x); dlm; GMHBal(x); dlm; GPenBal(x)
    End If
  Next x
  
  Close
        
  If GYrCnt = 0 Then
    Call TaxMsg(900, "There are no balances to report for the parameters entered.")
    Exit Sub
  End If
  
  arVATaxPMasterBalDet.Show
  
  Exit Sub

GetYearBalance:
  ThisRec = TaxCust.LastTrans
  TestBal = 0
  CreditUsed = 0
  OverPaid = 0
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
    If CInt(fpcmbTaxYear.Text) <> TaxTrans.TaxYear Then GoTo SkipItYear
    If TaxTrans.BillType <> "P" Then GoTo SkipItYear
    If TaxTrans.TranType = 10 Then 'added 10/3/06 adjust bill down affecting credit
       TestBal = TestBal + TaxTrans.Revenue.PrePaidUsed
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 11 Then 'added 10/3/06 prepay adjust down
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 24 Then 'added 10/3/06 adjust bill up affecting credit
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 21 Or TaxTrans.TranType = 22 Then
'      TestBal# = OldRound#(TestBal# - TaxTrans.Revenue.PrePaidAmt)'remmed on 2/2/07
      OverPaid = OldRound(OverPaid + TaxTrans.Revenue.PrePaidAmt) 'added 2/2/07
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 12 Then 'refund on prepay 7/17/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 9 Then 'credit applied at billing  'added 2/2/07
      CreditUsed = OldRound(CreditUsed + TaxTrans.Revenue.PrePaidUsed)
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 1 Then
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt + TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
    End If
SkipItYear:
    ThisRec = TaxTrans.LastTrans
  Loop
  If OverPaid = 0 Then CreditUsed = 0 'added 2/20/07
  
  TestBal = OldRound(TestBal - (OverPaid - CreditUsed)) 'added 2/2/07
  
  Return
  
GetAllBalance:
  ThisRec = TaxCust.LastTrans
  TestBal = 0
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
    If TaxTrans.BillType <> "P" Then GoTo DoAgain
    If TaxTrans.TranType = 10 Then 'added 10/3/06 adjust bill down affecting credit
       TestBal = TestBal + TaxTrans.Revenue.PrePaidUsed
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 11 Then 'added 10/3/06 prepay adjust down
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 24 Then 'added 10/3/06 adjust bill up affecting credit
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 21 Or TaxTrans.TranType = 22 Then
      TestBal# = OldRound#(TestBal# - TaxTrans.Revenue.PrePaidAmt)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 12 Then 'refund on prepay 7/17/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 9 Then 'added 8/11/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 1 Then
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt + TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
    End If
DoAgain:
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
  Disc9 = 0
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
  Disc9 = OldRound(TaxTrans.Revenue.LateListPd / SaveAmt)
  Disc9 = OldRound(Disc9 * TaxTrans.DiscAmt)
  TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd + Disc1 + Disc9)
  TaxTrans.Revenue.Principle2Pd = OldRound(TaxTrans.Revenue.Principle2Pd + Disc2)
  TaxTrans.Revenue.Principle3Pd = OldRound(TaxTrans.Revenue.Principle3Pd + Disc3)
  TaxTrans.Revenue.Principle4Pd = OldRound(TaxTrans.Revenue.Principle4Pd + Disc4)
  TaxTrans.Revenue.Principle5Pd = OldRound(TaxTrans.Revenue.Principle5Pd + Disc5)
  TaxTrans.Revenue.RevOpt1Pd = OldRound(TaxTrans.Revenue.RevOpt1Pd + Disc6)
  TaxTrans.Revenue.RevOpt2Pd = OldRound(TaxTrans.Revenue.RevOpt2Pd + Disc7)
  TaxTrans.Revenue.RevOpt3Pd = OldRound(TaxTrans.Revenue.RevOpt3Pd + Disc8)
  Dif = OldRound(Disc1 + Disc2 + Disc3 + Disc4 + Disc5 + Disc6 + Disc7 + Disc8 + Disc9)
  If Dif <> 0 Then
    If Disc1 > 0 Or Disc9 > 0 Then
      TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd + Dif)
    ElseIf Disc2 > 0 Then
      TaxTrans.Revenue.Principle2Pd = OldRound(TaxTrans.Revenue.Principle2Pd + Dif)
    ElseIf Disc3 > 0 Then
      TaxTrans.Revenue.Principle3Pd = OldRound(TaxTrans.Revenue.Principle3Pd + Dif)
    ElseIf Disc4 > 0 Then
      TaxTrans.Revenue.Principle4Pd = OldRound(TaxTrans.Revenue.Principle4Pd + Dif)
    ElseIf Disc5 > 0 Then
      TaxTrans.Revenue.Principle5Pd = OldRound(TaxTrans.Revenue.Principle5Pd + Dif)
    ElseIf Disc6 > 0 Then
      TaxTrans.Revenue.RevOpt1Pd = OldRound(TaxTrans.Revenue.RevOpt1Pd + Dif)
    ElseIf Disc7 > 0 Then
      TaxTrans.Revenue.RevOpt2Pd = OldRound(TaxTrans.Revenue.RevOpt2Pd + Dif)
    ElseIf Disc8 > 0 Then
      TaxTrans.Revenue.RevOpt3Pd = OldRound(TaxTrans.Revenue.RevOpt3Pd + Dif)
    End If
  End If
  DiscApplied = True
  
  Return

ERRORSTUFF:
   Unload frmVATaxShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxMasterBalList", "PrintRGraphicsDet", Erl)
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
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long, y As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim Balance As Double
  Dim ThisRec As Long
  Dim ThisYear As Integer
  Dim GTCnt As Long
  Dim TCnt As Integer
  Dim CustName$
  Dim YrCnt As Integer
  Dim GYrCnt As Integer
  Dim OverPay As Double
  Dim HoldYr As Integer
  Dim HoldBal As Double
  Dim Nextz As Integer
  Dim z As Integer
  Dim Thisz As Integer
  Dim BigYr As Integer
  Dim ThisBigYr As Integer
  Dim GBal As Double
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim RptFile$
  Dim RptHandle As Integer
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim CustRec As Long
  Dim InactiveFlag As Boolean
  Dim SubRptFile$
  Dim SubRptHandle As Integer
  Dim TransCnt As Long
  Dim OP As Double
  Dim ThisPersBal As Double
  Dim ThisIntBal As Double
  Dim ThisMTBal As Double
  Dim ThisMCBal As Double
  Dim ThisFEBal As Double
  Dim ThisMHBal As Double
  Dim ThisPenBal As Double
  Dim ThisOpt1Bal As Double
  Dim ThisOpt2Bal As Double
  Dim ThisOpt3Bal As Double
  Dim HoldPers As Double
  Dim HoldInt As Double
  Dim HoldMT As Double
  Dim HoldMC As Double
  Dim HoldFE As Double
  Dim HoldMH As Double
  Dim HoldPen As Double
  Dim HoldOpt1 As Double
  Dim HoldOpt2 As Double
  Dim HoldOpt3 As Double
  Dim GPersTot As Double
  Dim GIntTot As Double
  Dim GMTTot As Double
  Dim GMCTot As Double
  Dim GFETot As Double
  Dim GMHTot As Double
  Dim GPenTot As Double
  Dim GOpt1Tot As Double
  Dim GOpt2Tot As Double
  Dim GOpt3Tot As Double
  Dim FF$
  Dim Page As Integer
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim ThisAcct$, ThatAcct$
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim ActiveFlag$
  Dim ThisPin$
  Dim PropType$
  Dim CustTotBal As Double
  Dim BillNum$
  Dim ThisOP$
  Dim TestBal#
  Dim Disc1 As Double '1/16/2007
  Dim Disc2 As Double '1/16/2007
  Dim Disc3 As Double '1/16/2007
  Dim Disc4 As Double '1/16/2007
  Dim Disc5 As Double '1/16/2007
  Dim Disc6 As Double '1/16/2007
  Dim Disc7 As Double '1/16/2007
  Dim Disc8 As Double '1/16/2007
  Dim Disc9 As Double '9/17/2007
  Dim Dif As Double '9/19/07
  Dim DiscApplied As Boolean '1/16/2007
  Dim SaveAmt As Double '1/16/2007
  Dim CreditUsed As Double '2/2/07
  Dim OverPaid As Double '2/2/07
  
  On Error GoTo ERRORSTUFF
  
  MaxLines = 58
  FF$ = Chr(12)
  IdxFlag = False
  ActiveFlag = "B"
  If QPTrim$(fpcmbIncInactive.Text) = "Active Only" Then
    ActiveFlag = "A"
  ElseIf QPTrim$(fpcmbIncInactive.Text) = "Inactive Only" Then
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

  RptFile$ = "TAXRPTS\TXMSTBALDET.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  GoSub PrintHeader
  ReDim GYearBal(1 To 1) As Double
  ReDim GYears(1 To 1) As Integer
  ReDim GPersBal(1 To 1) As Double
  ReDim GIntBal(1 To 1) As Double
  ReDim GMTBal(1 To 1) As Double
  ReDim GMCBal(1 To 1) As Double
  ReDim GFEBal(1 To 1) As Double
  ReDim GMHBal(1 To 1) As Double
  ReDim GPenBal(1 To 1) As Double
  ReDim GOPt1Bal(1 To 1) As Double
  ReDim GOPt2Bal(1 To 1) As Double
  ReDim GOPt3Bal(1 To 1) As Double
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenTaxTransFile TTHandle, NumOfTTRecs
  GYrCnt = 0
  frmVATaxShowPctComp.Label1 = "Gathering Tax Transaction Data"
  frmVATaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  
  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  End If
  
  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
      CustRec = IdxArray(x)
    Else
      Get TCHandle, x, TaxCust
      CustRec = x
    End If
    Balance = 0 'added 7/19/06
    If TaxCust.Deleted <> 0 Then GoTo Inactive 'SkipIt
    If ActiveFlag = "A" And TaxCust.Active = "N" Then
      GoTo Inactive
    ElseIf ActiveFlag = "I" And TaxCust.Active = "Y" Then
      GoTo Inactive
    End If
    
    If ActiveFlag = "B" Then
      If TaxCust.Active = "N" Then
        CustName = QPTrim$(TaxCust.CustName) + "(I)"
      Else
        CustName = QPTrim$(TaxCust.CustName)
      End If
    Else
      CustName = QPTrim$(TaxCust.CustName)
    End If
    
    If fpcmbTaxYear.Text = "All" Then
      GoSub GetAllBalance
    Else
      GoSub GetYearBalance
    End If
    
    CustTotBal = TestBal ' GetCustPersBalance(CustRec, -1)
    OP = TestBal
    ThatAcct = CStr(TaxCust.Acct)
    
    If TaxCust.LastTrans > 0 Then
      ReDim YearBal(1 To 1) As Double
      YrCnt = 0
      ReDim Years(1 To 1) As Integer
      ReDim PersBal(1 To 1) As Double
      ReDim IntBal(1 To 1) As Double
      ReDim MTBal(1 To 1) As Double
      ReDim MCBal(1 To 1) As Double
      ReDim FEBal(1 To 1) As Double
      ReDim MHBal(1 To 1) As Double
      ReDim PenBal(1 To 1) As Double
      ReDim Opt1Bal(1 To 1) As Double
      ReDim Opt2Bal(1 To 1) As Double
      ReDim Opt3Bal(1 To 1) As Double
      ThisAcct = ""
      
      ThisRec = TaxCust.LastTrans
      ThisOP = CStr(OP)
      If InStr(ThisOP, "E") Then OP = 0
      If OP < 0 Then
        OverPay = OldRound(OverPay + Abs(OP))
      End If
      Do While ThisRec > 0
        Get TTHandle, ThisRec, TaxTrans
        ThisPersBal = 0 '9/1/06
        ThisMTBal = 0 '9/1/06
        ThisMCBal = 0 '9/1/06
        ThisFEBal = 0 '9/1/06
        ThisMHBal = 0 '9/1/06
        ThisIntBal = 0 '9/1/06
        ThisPenBal = 0 '9/1/06
        ThisOpt1Bal = 0 '9/1/06
        ThisOpt2Bal = 0 '9/1/06
        ThisOpt3Bal = 0 '9/1/06
        Balance# = 0 '9/1/06
        If CustTotBal < 0 Then
          Balance# = CustTotBal
          If fpcmbTaxYear.Text <> "All" Then
            TaxTrans.TaxYear = CInt(fpcmbTaxYear.Text)
          End If
          TaxTrans.LastTrans = 0
          GoTo GoWithZero
        End If
        If TaxTrans.TranType = 1 And TaxTrans.BillType = "P" Then
          If fpcmbTaxYear.Text <> "All" Then
            If CInt(fpcmbTaxYear.Text) <> TaxTrans.TaxYear Then GoTo SkipIt
          End If
          If TaxTrans.DiscAmt > 0 Then GoSub ApplyDisc 'added 1/16/07
          ThisPersBal = OldRound(TaxTrans.Revenue.Principle1 + TaxTrans.PPTRARmvl - TaxTrans.PPTRADisc - TaxTrans.Revenue.Principle1Pd)
          ThisIntBal = OldRound(TaxTrans.Revenue.Interest - TaxTrans.Revenue.InterestPd)
          ThisMTBal = OldRound(TaxTrans.Revenue.Principle2 - TaxTrans.Revenue.Principle2Pd)
          ThisMCBal = OldRound(TaxTrans.Revenue.Principle3 - TaxTrans.Revenue.Principle3Pd)
          ThisFEBal = OldRound(TaxTrans.Revenue.Principle4 - TaxTrans.Revenue.Principle4Pd)
          ThisMHBal = OldRound(TaxTrans.Revenue.Principle5 - TaxTrans.Revenue.Principle5Pd)
          ThisPenBal = OldRound(TaxTrans.Revenue.Penalty - TaxTrans.Revenue.PenaltyPd)
          ThisOpt1Bal = OldRound(TaxTrans.Revenue.RevOpt1 - TaxTrans.Revenue.RevOpt1Pd)
          ThisOpt2Bal = OldRound(TaxTrans.Revenue.RevOpt2 - TaxTrans.Revenue.RevOpt2Pd)
          ThisOpt3Bal = OldRound(TaxTrans.Revenue.RevOpt3 - TaxTrans.Revenue.RevOpt3Pd)
          Balance# = OldRound#(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
          Balance# = OldRound#(Balance# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
          Balance# = OldRound#(Balance# + TaxTrans.Revenue.Collection + TaxTrans.PPTRARmvl)
          Balance# = OldRound#(Balance# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
          Balance# = OldRound#(Balance# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
          Balance# = OldRound#(Balance# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd)) 'changed on 1/16/07 + TaxTrans.DiscAmt))
          Balance# = OldRound#(Balance# - TaxTrans.PPTRADisc)
          If Balance = 0 And chkZeroBal.Value = 1 Then
            GoTo GoWithZero
          End If
          If Balance <> 0 Then
GoWithZero:
            If ThisAcct <> ThatAcct Then
              GoSub PrintCustHeader
              ThisAcct = ThatAcct
            End If
            GBal = OldRound(GBal + Balance#)
            If QPTrim$(TaxTrans.RealPin) = "0" And QPTrim$(TaxTrans.PersPin) = "0" Then
              ThisPin = "0"
              PropType = "UNATTACHED"
            ElseIf QPTrim$(TaxTrans.PersPin) <> "0" And QPTrim$(TaxTrans.PersPin) <> "" Then
              ThisPin = QPTrim$(TaxTrans.PersPin)
              PropType = "PERSONAL"
            Else
              ThisPin = "0"
              PropType = "UNATTACHED"
            End If
            TransCnt = TransCnt + 1
            If LineCnt >= MaxLines - 1 Then
              Print #RptHandle, FF$
              GoSub PrintHeader
              GoSub PrintCustHeader
            End If
            Print #RptHandle, "Prop Type: " + PropType; Tab(23); "Pin #: " + ThisPin; Tab(43); "Tax Year: " + Using$("###0", TaxTrans.TaxYear); Tab(59); "Balance: " + Using$("$###,##0.00", Balance#) '; dlm; TransCnt; dlm; GBal; dlm;
            Print #RptHandle, "BillNum: " + ParseBillNum(TaxTrans.Description)
            LineCnt = LineCnt + 2
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintHeader
              GoSub PrintCustHeader
            End If
            If QPTrim$(POpt1Desc) <> "" Then
              Print #RptHandle, Tab(5); "Personal:   " + Using$("$##,##0.00", ThisPersBal);
              Print #RptHandle, Tab(40); QPTrim$(POpt1Desc); Tab(69); Using$("$##,##0.00", ThisOpt1Bal)
              LineCnt = LineCnt + 1
              If LineCnt >= MaxLines Then
                Print #RptHandle, FF$
                GoSub PrintHeader
                GoSub PrintCustHeader
              End If
            Else
              Print #RptHandle, Tab(5); "Personal:   " + Using$("$##,##0.00", ThisPersBal)
              LineCnt = LineCnt + 1
              If LineCnt >= MaxLines Then
                 Print #RptHandle, FF$
                 GoSub PrintHeader
                 GoSub PrintCustHeader
              End If
            End If
            If QPTrim$(POpt2Desc) <> "" Then
              Print #RptHandle, Tab(5); "Mach Tools: " + Using$("$##,##0.00", ThisMTBal);
              Print #RptHandle, Tab(40); QPTrim$(POpt2Desc); Tab(69); Using$("$##,##0.00", ThisOpt2Bal)
              LineCnt = LineCnt + 1
              If LineCnt >= MaxLines Then
                Print #RptHandle, FF$
                GoSub PrintHeader
                GoSub PrintCustHeader
              End If
            Else
              Print #RptHandle, Tab(5); "Mach Tools: " + Using$("$##,##0.00", ThisMTBal)
              LineCnt = LineCnt + 1
              If LineCnt >= MaxLines Then
                Print #RptHandle, FF$
                GoSub PrintHeader
                GoSub PrintCustHeader
              End If
            End If
            If QPTrim$(POpt3Desc) <> "" Then
              Print #RptHandle, Tab(5); "Merch Cap:  " + Using$("$##,##0.00", ThisMCBal);
              Print #RptHandle, Tab(40); QPTrim$(POpt3Desc); Tab(69); Using$("$##,##0.00", ThisOpt3Bal)
              LineCnt = LineCnt + 1
              If LineCnt >= MaxLines Then
                Print #RptHandle, FF$
                GoSub PrintHeader
                GoSub PrintCustHeader
              End If
            Else
              Print #RptHandle, Tab(5); "Merch Cap:  " + Using$("$##,##0.00", ThisMCBal)
              LineCnt = LineCnt + 1
              If LineCnt >= MaxLines Then
                Print #RptHandle, FF$
                GoSub PrintHeader
                GoSub PrintCustHeader
              End If
            End If
            Print #RptHandle, Tab(5); "Farm Equip: " + Using$("$##,##0.00", ThisFEBal)
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintHeader
              GoSub PrintCustHeader
            End If
            Print #RptHandle, Tab(5); "Mbl Homes:  " + Using$("$##,##0.00", ThisMHBal)
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintHeader
              GoSub PrintCustHeader
            End If
            Print #RptHandle, Tab(5); "Interest:   " + Using$("$##,##0.00", ThisIntBal)
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintHeader
              GoSub PrintCustHeader
            End If
            Print #RptHandle, Tab(5); "Penalty:    " + Using$("$##,##0.00", ThisPenBal)
            LineCnt = LineCnt + 1
            If LineCnt >= MaxLines Then
              Print #RptHandle, FF$
              GoSub PrintHeader
              GoSub PrintCustHeader
            End If
            
            If GYrCnt = 0 Then '
              GYrCnt = GYrCnt + 1
              ReDim Preserve GYears(1 To GYrCnt) As Integer
              GYears(GYrCnt) = TaxTrans.TaxYear
              ReDim Preserve GYearBal(1 To GYrCnt) As Double
              GYearBal(GYrCnt) = Balance#
              ReDim Preserve GPersBal(1 To GYrCnt) As Double
              GPersBal(GYrCnt) = ThisPersBal#
              GPersTot = ThisPersBal#
              ReDim Preserve GIntBal(1 To GYrCnt) As Double
              GIntBal(GYrCnt) = ThisIntBal#
              GIntTot = ThisIntBal#
              ReDim Preserve GMTBal(1 To GYrCnt) As Double
              GMTBal(GYrCnt) = ThisMTBal#
              GMTTot = ThisMTBal#
              ReDim Preserve GMCBal(1 To GYrCnt) As Double
              GMCBal(GYrCnt) = ThisMCBal#
              GMCTot = ThisMCBal#
              ReDim Preserve GFEBal(1 To GYrCnt) As Double
              GFEBal(GYrCnt) = ThisFEBal#
              GFETot = ThisFEBal#
              ReDim Preserve GMHBal(1 To GYrCnt) As Double
              GMHBal(GYrCnt) = ThisMHBal#
              GMHTot = ThisMHBal#
              ReDim Preserve GPenBal(1 To GYrCnt) As Double
              GPenBal(GYrCnt) = ThisPenBal#
              GPenTot = ThisPenBal#
              ReDim Preserve GOPt1Bal(1 To GYrCnt) As Double
              GOPt1Bal(GYrCnt) = ThisOpt1Bal#
              GOpt1Tot = ThisOpt1Bal#
              ReDim Preserve GOPt2Bal(1 To GYrCnt) As Double
              GOPt2Bal(GYrCnt) = ThisOpt2Bal#
              GOpt2Tot = ThisOpt2Bal#
              ReDim Preserve GOPt3Bal(1 To GYrCnt) As Double
              GOPt3Bal(GYrCnt) = ThisOpt3Bal#
              GOpt3Tot = ThisOpt3Bal#
            Else
              For y = 1 To GYrCnt
                If GYears(y) = TaxTrans.TaxYear Then '
                  GYearBal(y) = OldRound(GYearBal(y) + Balance#)
                  GPersBal(y) = OldRound(GPersBal(y) + ThisPersBal#)
                  GPersTot = OldRound(GPersTot# + ThisPersBal#)
                  GIntBal(y) = OldRound(GIntBal(y) + ThisIntBal#)
                  GIntTot = OldRound(GIntTot# + ThisIntBal#)
                  GMTBal(y) = OldRound(GMTBal(y) + ThisMTBal#)
                  GMTTot = OldRound(GMTTot# + ThisMTBal#)
                  GMCBal(y) = OldRound(GMCBal(y) + ThisMCBal#)
                  GMCTot = OldRound(GMCTot# + ThisMCBal#)
                  GFEBal(y) = OldRound(GFEBal(y) + ThisFEBal#)
                  GFETot = OldRound(GFETot# + ThisFEBal#)
                  GMHBal(y) = OldRound(GMHBal(y) + ThisMHBal#)
                  GMHTot = OldRound(GMHTot# + ThisMHBal#)
                  GPenBal(y) = OldRound(GPenBal(y) + ThisPenBal#)
                  GPenTot = OldRound(GPenTot# + ThisPenBal#)
                  GOPt1Bal(y) = OldRound(GOPt1Bal(y) + ThisOpt1Bal#)
                  GOpt1Tot = OldRound(GOpt1Tot# + ThisOpt1Bal#)
                  GOPt2Bal(y) = OldRound(GOPt2Bal(y) + ThisOpt2Bal#)
                  GOpt2Tot = OldRound(GOpt2Tot# + ThisOpt2Bal#)
                  GOPt3Bal(y) = OldRound(GOPt3Bal(y) + ThisOpt3Bal#)
                  GOpt3Tot = OldRound(GOpt3Tot# + ThisOpt3Bal#)
                  Exit For
                End If '
              Next y
              If y > GYrCnt Then '
                GYrCnt = GYrCnt + 1
                ReDim Preserve GYears(1 To GYrCnt) As Integer
                GYears(GYrCnt) = TaxTrans.TaxYear
                ReDim Preserve GYearBal(1 To GYrCnt) As Double
                GYearBal(GYrCnt) = Balance#
                ReDim Preserve GPersBal(1 To GYrCnt) As Double
                GPersBal(GYrCnt) = ThisPersBal#
                GPersTot# = OldRound(GPersTot# + ThisPersBal#)
                ReDim Preserve GIntBal(1 To GYrCnt) As Double
                GIntBal(GYrCnt) = ThisIntBal#
                GIntTot# = OldRound(GIntTot# + ThisIntBal#)
                ReDim Preserve GMTBal(1 To GYrCnt) As Double
                GMTBal(GYrCnt) = ThisMTBal#
                GMTTot# = OldRound(GMTTot# + ThisMTBal#)
                ReDim Preserve GMCBal(1 To GYrCnt) As Double
                GMCBal(GYrCnt) = ThisMCBal#
                GMCTot# = OldRound(GMCTot + ThisMCBal#)
                ReDim Preserve GFEBal(1 To GYrCnt) As Double
                GFEBal(GYrCnt) = ThisFEBal#
                GFETot# = OldRound(GFETot# + ThisFEBal#)
                ReDim Preserve GMHBal(1 To GYrCnt) As Double
                GMHBal(GYrCnt) = ThisMHBal#
                GMHTot# = OldRound(GMHTot + ThisMHBal#)
                ReDim Preserve GPenBal(1 To GYrCnt) As Double
                GPenBal(GYrCnt) = ThisPenBal#
                GPenTot# = OldRound(GPenTot# + ThisPenBal#)
                ReDim Preserve GOPt1Bal(1 To GYrCnt) As Double
                GOPt1Bal(GYrCnt) = ThisOpt1Bal#
                GOpt1Tot# = OldRound(GOpt1Tot# + ThisOpt1Bal#)
                ReDim Preserve GOPt2Bal(1 To GYrCnt) As Double
                GOPt2Bal(GYrCnt) = ThisOpt2Bal#
                GOpt2Tot# = OldRound(GOpt2Tot# + ThisOpt2Bal#)
                ReDim Preserve GOPt3Bal(1 To GYrCnt) As Double
                GOPt3Bal(GYrCnt) = ThisOpt3Bal#
                GOpt3Tot# = OldRound(GOpt3Tot# + ThisOpt3Bal#)
              End If '
            End If '
          End If
        End If
SkipIt:
        ThisRec = TaxTrans.LastTrans
      Loop
    End If
Inactive:
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
  
  BigYr = 0
  For x = 1 To GYrCnt
    If GYears(x) > BigYr Then
      BigYr = GYears(x)
    End If
  Next x
  
  ThisBigYr = BigYr + 1
  Nextz = 1
  Do While Nextz <= GYrCnt
    For z = Nextz To GYrCnt
      If GYears(z) < ThisBigYr Then
         ThisBigYr = GYears(z)
         Thisz = z
      End If
    Next z
    HoldYr = GYears(Nextz)
    HoldBal = GYearBal(Nextz)
    HoldPers = GPersBal(Nextz)
    HoldInt = GIntBal(Nextz)
    HoldMT = GMTBal(Nextz)
    HoldMC = GMCBal(Nextz)
    HoldFE = GFEBal(Nextz)
    HoldMH = GMHBal(Nextz)
    HoldPen = GPenBal(Nextz)
    HoldOpt1 = GOPt1Bal(Nextz)
    HoldOpt2 = GOPt2Bal(Nextz)
    HoldOpt3 = GOPt3Bal(Nextz)
    GYears(Nextz) = GYears(Thisz)
    GYearBal(Nextz) = GYearBal(Thisz)
    GPersBal(Nextz) = GPersBal(Thisz)
    GIntBal(Nextz) = GIntBal(Thisz)
    GMTBal(Nextz) = GMTBal(Thisz)
    GMCBal(Nextz) = GMCBal(Thisz)
    GFEBal(Nextz) = GFEBal(Thisz)
    GMHBal(Nextz) = GMHBal(Thisz)
    GPenBal(Nextz) = GPenBal(Thisz)
    GOPt1Bal(Nextz) = GOPt1Bal(Thisz)
    GOPt2Bal(Nextz) = GOPt2Bal(Thisz)
    GOPt3Bal(Nextz) = GOPt3Bal(Thisz)
    GYears(Thisz) = HoldYr
    GYearBal(Thisz) = HoldBal
    GPersBal(Thisz) = HoldPers
    GIntBal(Thisz) = HoldInt
    GMTBal(Thisz) = HoldMT
    GMCBal(Thisz) = HoldMC
    GFEBal(Thisz) = HoldFE
    GMHBal(Thisz) = HoldMH
    GPenBal(Thisz) = HoldPen
    GOPt1Bal(Thisz) = HoldOpt1
    GOPt2Bal(Thisz) = HoldOpt2
    GOPt3Bal(Thisz) = HoldOpt3
    
    Nextz = Nextz + 1
    ThisBigYr = BigYr + 1
  Loop
  
  If InStr(CStr(OverPay), "E") Or fpcmbTaxYear.Text <> "All" Then OverPay = 0 'added All 12/5/06
  Print #RptHandle, FF$
  GoSub PrintEndHeader
  Print #RptHandle, "Total Entries: " + Using$("####0", TransCnt)
  Print #RptHandle, "Total Tax Balance: "; Tab(30); Using$("$###,###,##0.00", GBal)
  If fpcmbTaxYear.Text = "All" Then
    Print #RptHandle, "OverPay:           "; Tab(30); Using$("$###,###,##0.00", OverPay)
  Else
    Print #RptHandle,
  End If
  Print #RptHandle,
  Print #RptHandle, "Personal Total:          "; Tab(30); Using$("$###,###,##0.00", GPersTot)
  Print #RptHandle, "Mach Tools Total:        "; Tab(30); Using$("$###,###,##0.00", GMTTot)
  Print #RptHandle, "Merch Cap Total:         "; Tab(30); Using$("$###,###,##0.00", GMCTot)
  Print #RptHandle, "Farm Equip Total:        "; Tab(30); Using$("$###,###,##0.00", GFETot)
  Print #RptHandle, "Mbl Homes Total:         "; Tab(30); Using$("$###,###,##0.00", GMHTot)
  Print #RptHandle, "Interest Total:          "; Tab(30); Using$("$###,###,##0.00", GIntTot)
  Print #RptHandle, "Penalty  Total:          "; Tab(30); Using$("$###,###,##0.00", GPenTot)
  LineCnt = LineCnt + 10
  If QPTrim$(POpt1Desc) <> "" Then
    Print #RptHandle, QPTrim$(POpt1Desc) + ": "; Tab(30); Using$("$###,###,##0.00", GOpt1Tot)
    LineCnt = LineCnt + 1
  End If
  If QPTrim$(POpt2Desc) <> "" Then
    Print #RptHandle, QPTrim$(POpt2Desc) + ": "; Tab(30); Using$("$###,###,##0.00", GOpt2Tot)
    LineCnt = LineCnt + 1
  End If
  If QPTrim$(POpt3Desc) <> "" Then
    Print #RptHandle, QPTrim$(POpt3Desc) + ": "; Tab(30); Using$("$###,###,##0.00", GOpt3Tot)
    LineCnt = LineCnt + 1
  End If
  Print #RptHandle,
  Print #RptHandle,
  LineCnt = LineCnt + 2
  
  For x = 1 To GYrCnt
    Print #RptHandle, "Tax Year"; Tab(34); "Amount Owed"
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintEndHeader
      Print #RptHandle, "Tax Year"; Tab(34); "Amount Owed" '; Tab(60); "Over Payments"
    End If
    Print #RptHandle, Using$("###0", GYears(x)); Tab(30); Using$("$###,###,##0.00", GYearBal(x))
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintEndHeader
      Print #RptHandle, "Tax Year" '; Tab(30); "Amount Owed" '; Tab(60); "Over Payments"
    End If
    Print #RptHandle, Tab(5); "Personal: "; Tab(30); Using$("$###,###,##0.00", GPersBal(x))
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintEndHeader
      Print #RptHandle, "Tax Year" '; Tab(30); "Amount Owed" '; Tab(60); "Over Payments"
    End If
    Print #RptHandle, Tab(5); "Mach Tools: "; Tab(30); Using$("$###,###,##0.00", GMTBal(x))
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintEndHeader
      Print #RptHandle, "Tax Year" '; Tab(30); "Amount Owed" '; Tab(60); "Over Payments"
    End If
    Print #RptHandle, Tab(5); "Merch Cap: "; Tab(30); Using$("$###,###,##0.00", GMCBal(x))
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintEndHeader
      Print #RptHandle, "Tax Year" '; Tab(30); "Amount Owed" '; Tab(60); "Over Payments"
    End If
    Print #RptHandle, Tab(5); "Farm Equip: "; Tab(30); Using$("$###,###,##0.00", GFEBal(x))
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintEndHeader
      Print #RptHandle, "Tax Year" '; Tab(30); "Amount Owed" '; Tab(60); "Over Payments"
    End If
    Print #RptHandle, Tab(5); "Mbl Homes: "; Tab(30); Using$("$###,###,##0.00", GMHBal(x))
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintEndHeader
      Print #RptHandle, "Tax Year" '; Tab(30); "Amount Owed" '; Tab(60); "Over Payments"
    End If
    Print #RptHandle, Tab(5); "Interest: "; Tab(30); Using$("$###,###,##0.00", GIntBal(x))
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintEndHeader
      Print #RptHandle, "Tax Year" '; Tab(30); "Amount Owed" '; Tab(60); "Over Payments"
    End If
    Print #RptHandle, Tab(5); "Penalty: "; Tab(30); Using$("$###,###,##0.00", GPenBal(x))
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintEndHeader
      Print #RptHandle, "Tax Year" '; Tab(30); "Amount Owed" '; Tab(60); "Over Payments"
    End If
    If QPTrim$(POpt1Desc) <> "" Then
      Print #RptHandle, Tab(5); POpt1Desc + ":"; Tab(30); Using$("$###,###,##0.00", GOPt1Bal(x))
      LineCnt = LineCnt + 1
      If LineCnt >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintEndHeader
        Print #RptHandle, "Tax Year" '; Tab(30); "Amount Owed" '; Tab(60); "Over Payments"
      End If
    End If
    If QPTrim$(POpt2Desc) <> "" Then
      Print #RptHandle, Tab(5); POpt2Desc + ":"; Tab(30); Using$("$###,###,##0.00", GOPt2Bal(x))
      LineCnt = LineCnt + 1
      If LineCnt >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintEndHeader
        Print #RptHandle, "Tax Year" '; Tab(30); "Amount Owed" '; Tab(60); "Over Payments"
      End If
    End If
    If QPTrim$(POpt3Desc) <> "" Then
      Print #RptHandle, Tab(5); POpt3Desc + ":"; Tab(30); Using$("$###,###,##0.00", GOPt3Bal(x))
      LineCnt = LineCnt + 1
      If LineCnt >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintEndHeader
        Print #RptHandle, "Tax Year" '; Tab(30); "Amount Owed" '; Tab(60); "Over Payments"
      End If
    End If
    Print #RptHandle,
    LineCnt = LineCnt + 1
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintEndHeader
      Print #RptHandle, "Tax Year" '; Tab(30); "Amount Owed"; Tab(60); "Over Payments"
    End If
  Next x
  
  Print #RptHandle, FF$
  Close
        
  If GYrCnt = 0 Then
    Call TaxMsg(900, "There are no balances to report for the parameters entered.")
    Exit Sub
  End If
  ViewPrint RptFile, "Tax Master Balance Listing", True
  
  Exit Sub

PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(20); "Tax Master Balance Listing Detail - Personal"
  Print #RptHandle, Town
  If Len(CStr(Page)) = 2 Then
    Print #RptHandle, "Report Date: " + CStr(Now); Tab(71); "Page # " + CStr(Page)
  ElseIf Len(CStr(Page)) = 3 Then
    Print #RptHandle, "Report Date: " + CStr(Now); Tab(70); "Page # " + CStr(Page)
  ElseIf Len(CStr(Page)) = 4 Then
    Print #RptHandle, "Report Date: " + CStr(Now); Tab(69); "Page # " + CStr(Page)
  Else
    Print #RptHandle, "Report Date: " + CStr(Now); Tab(72); "Page # " + CStr(Page)
  End If
  If ActiveFlag = "B" Then
    Print #RptHandle, "Customer Status: " + "Active And Inactive"; Tab(65); "(I) = Inactive"
  Else
    Print #RptHandle, "Customer Status: " + fpcmbIncInactive.Text
  End If
  Print #RptHandle, "Acct #"; Tab(8); "Customer Name"; Tab(59); "Tax Year"; Tab(73); "Balance"
  Print #RptHandle, String(79, "-")
  LineCnt = 6
  
  Return
  
PrintCustHeader:
  If LineCnt <> 6 Then
    Print #RptHandle, String(79, "-")
    LineCnt = LineCnt + 1
  End If
  Print #RptHandle, Using$("####0", CustRec); Tab(8); CustName; Tab(55); "Cust Balance: " + QPTrim$(Using$("$###,###,##0.00", CustTotBal))
  Print #RptHandle, String(79, "-")
  LineCnt = LineCnt + 2
  Return

PrintEndHeader:
  Page = Page + 1
  Print #RptHandle, Tab(25); "Tax Master Balance Listing Detail"
  Print #RptHandle, "Personal Only"
  Print #RptHandle, Town
  If Len(CStr(Page)) = 2 Then
    Print #RptHandle, "Report Date: " + CStr(Now); Tab(71); "Page # " + CStr(Page)
  ElseIf Len(CStr(Page)) = 3 Then
    Print #RptHandle, "Report Date: " + CStr(Now); Tab(70); "Page # " + CStr(Page)
  ElseIf Len(CStr(Page)) = 4 Then
    Print #RptHandle, "Report Date: " + CStr(Now); Tab(69); "Page # " + CStr(Page)
  Else
    Print #RptHandle, "Report Date: " + CStr(Now); Tab(72); "Page # " + CStr(Page)
  End If

  Print #RptHandle, String(79, "-")
  LineCnt = 5
  
  Return
  
GetYearBalance:
  ThisRec = TaxCust.LastTrans
  TestBal = 0
  CreditUsed = 0
  OverPaid = 0
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
    If CInt(fpcmbTaxYear.Text) <> TaxTrans.TaxYear Then GoTo SkipItYear
    If TaxTrans.BillType <> "P" Then GoTo SkipItYear
    If TaxTrans.TranType = 10 Then 'added 10/3/06 adjust bill down affecting credit
       TestBal = TestBal + TaxTrans.Revenue.PrePaidUsed
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 11 Then 'added 10/3/06 prepay adjust down
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 24 Then 'added 10/3/06 adjust bill up affecting credit
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 21 Or TaxTrans.TranType = 22 Then
'      TestBal# = OldRound#(TestBal# - TaxTrans.Revenue.PrePaidAmt)'remmed on 2/2/07
      OverPaid = OldRound(OverPaid + TaxTrans.Revenue.PrePaidAmt) 'added 2/2/07
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 12 Then 'refund on prepay 7/17/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 9 Then 'credit applied at billing  'added 2/2/07
      CreditUsed = OldRound(CreditUsed + TaxTrans.Revenue.PrePaidUsed)
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 1 Then
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt + TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
    End If
SkipItYear:
    ThisRec = TaxTrans.LastTrans
  Loop
  If OverPaid = 0 Then CreditUsed = 0 'added 2/20/07
  
  TestBal = OldRound(TestBal - (OverPaid - CreditUsed)) 'added 2/2/07
  
  Return
  
GetAllBalance:
  ThisRec = TaxCust.LastTrans
  TestBal = 0
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
    If TaxTrans.BillType <> "P" Then GoTo DoAgain
    If TaxTrans.TranType = 10 Then 'added 10/3/06 adjust bill down affecting credit
       TestBal = TestBal + TaxTrans.Revenue.PrePaidUsed
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 11 Then 'added 10/3/06 prepay adjust down
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 24 Then 'added 10/3/06 adjust bill up affecting credit
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 21 Or TaxTrans.TranType = 22 Then
      TestBal# = OldRound#(TestBal# - TaxTrans.Revenue.PrePaidAmt)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 12 Then 'refund on prepay 7/17/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 9 Then 'added 8/11/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 1 Then
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt + TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
    End If
DoAgain:
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
  Disc9 = 0
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
  Disc9 = OldRound(TaxTrans.Revenue.LateListPd / SaveAmt)
  Disc9 = OldRound(Disc9 * TaxTrans.DiscAmt)
  TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd + Disc1 + Disc9)
  TaxTrans.Revenue.Principle2Pd = OldRound(TaxTrans.Revenue.Principle2Pd + Disc2)
  TaxTrans.Revenue.Principle3Pd = OldRound(TaxTrans.Revenue.Principle3Pd + Disc3)
  TaxTrans.Revenue.Principle4Pd = OldRound(TaxTrans.Revenue.Principle4Pd + Disc4)
  TaxTrans.Revenue.Principle5Pd = OldRound(TaxTrans.Revenue.Principle5Pd + Disc5)
  TaxTrans.Revenue.RevOpt1Pd = OldRound(TaxTrans.Revenue.RevOpt1Pd + Disc6)
  TaxTrans.Revenue.RevOpt2Pd = OldRound(TaxTrans.Revenue.RevOpt2Pd + Disc7)
  TaxTrans.Revenue.RevOpt3Pd = OldRound(TaxTrans.Revenue.RevOpt3Pd + Disc8)
  Dif = OldRound(Disc1 + Disc2 + Disc3 + Disc4 + Disc5 + Disc6 + Disc7 + Disc8 + Disc9)
  If Dif <> 0 Then
    If Disc1 > 0 Or Disc9 > 0 Then
      TaxTrans.Revenue.Principle1Pd = OldRound(TaxTrans.Revenue.Principle1Pd + Dif)
    ElseIf Disc2 > 0 Then
      TaxTrans.Revenue.Principle2Pd = OldRound(TaxTrans.Revenue.Principle2Pd + Dif)
    ElseIf Disc3 > 0 Then
      TaxTrans.Revenue.Principle3Pd = OldRound(TaxTrans.Revenue.Principle3Pd + Dif)
    ElseIf Disc4 > 0 Then
      TaxTrans.Revenue.Principle4Pd = OldRound(TaxTrans.Revenue.Principle4Pd + Dif)
    ElseIf Disc5 > 0 Then
      TaxTrans.Revenue.Principle5Pd = OldRound(TaxTrans.Revenue.Principle5Pd + Dif)
    ElseIf Disc6 > 0 Then
      TaxTrans.Revenue.RevOpt1Pd = OldRound(TaxTrans.Revenue.RevOpt1Pd + Dif)
    ElseIf Disc7 > 0 Then
      TaxTrans.Revenue.RevOpt2Pd = OldRound(TaxTrans.Revenue.RevOpt2Pd + Dif)
    ElseIf Disc8 > 0 Then
      TaxTrans.Revenue.RevOpt3Pd = OldRound(TaxTrans.Revenue.RevOpt3Pd + Dif)
    End If
  End If
  DiscApplied = True
  
  Return
  
ERRORSTUFF:
   Unload frmVATaxShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxMasterBalList", "PrintRTextDet", Erl)
     Case emrExitProc:
       Resume Proc_Exit
     Case emrResume:
       Resume
     Case emrResumeNext:
       Resume Next
     Case Else
      '--- Technically, this should never hapPers.
       Resume Proc_Exit
   End Select
  
Proc_Exit:
  '--- Cleanup code goes here...
    Close
    ClearInUse PWcnt
    Terminate
  

End Sub

Private Function GetOverPayAmount(ThEYear As Integer, ThisType As String) As Double
  Dim x As Long
  Dim TaxCust As TaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim OP As Double
  Dim TotOP
  Dim ActiveFlag$
  Dim TestBal#
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim ThisRec As Long
  Dim PropFlag$
  
  PropFlag = ThisType
  ActiveFlag = "B"
  If QPTrim$(fpcmbIncInactive.Text) = "Active Only" Then
    ActiveFlag = "A"
  ElseIf QPTrim$(fpcmbIncInactive.Text) = "Inactive Only" Then
    ActiveFlag = "I"
  End If
  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenTaxCustFile CHandle, NumOfCRecs
  frmVATaxShowPctComp.Label1 = "Gathering Tax Transaction Data"
  frmVATaxShowPctComp.Show , Me
  frmVATaxShowPctComp.cmdCancel.Visible = False
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  For x = 1 To NumOfCRecs
    Get CHandle, x, TaxCust
    If TaxCust.Deleted <> 0 Then GoTo Inactive
    If ActiveFlag = "A" And TaxCust.Active = "N" Then
      GoTo Inactive
    ElseIf ActiveFlag = "I" And TaxCust.Active = "Y" Then
      GoTo Inactive
    End If
    If ThEYear < 0 Then
      GoSub GetAllBalance
    Else
      GoSub GetYearBalance
    End If
    OP = TestBal#
'      If ThisType = "P" Then
'        OP = OldRound(GetCustPersBalance(x, -1))
'      ElseIf ThisType = "R" Then
'        OP = OldRound(GetCustRealBalance(x, -1))
'      Else
'        OP = OldRound(GetCustBalance(x, -1))
'      End If
'      If OP < 0 Then
'        TotOP = OldRound(TotOP + Abs(OP))
'      End If
'    Else
'      OP = OldRound(GetCustBalanceForYear(x, ThEYear, ThisType))
      If OP < 0 Then
        TotOP = OldRound(TotOP + Abs(OP))
      End If
'    End If
Inactive:
    frmVATaxShowPctComp.ShowPctComp x, NumOfCRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      Exit Function
    End If
  Next x
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  
  Close CHandle
  Close TTHandle
  GetOverPayAmount = TotOP
  Exit Function
  
GetYearBalance:
  ThisRec = TaxCust.LastTrans
  TestBal = 0
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
    If CInt(fpcmbTaxYear.Text) <> TaxTrans.TaxYear Then GoTo SkipItYear
    If PropFlag = "R" Then
      If TaxTrans.BillType <> "R" Then GoTo SkipItYear
    ElseIf PropFlag = "P" Then
      If TaxTrans.BillType <> "P" Then GoTo SkipItYear
    End If
    If TaxTrans.TranType = 21 Or TaxTrans.TranType = 22 Then
      TestBal# = OldRound#(TestBal# - TaxTrans.Revenue.PrePaidAmt)
      GoTo SkipItYear
    End If
    If TaxTrans.TranType = 1 Then
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt))
    End If
SkipItYear:
    ThisRec = TaxTrans.LastTrans
  Loop
  
  Return
  
GetAllBalance:
  ThisRec = TaxCust.LastTrans
  TestBal = 0
  Do While ThisRec > 0
    Get TTHandle, ThisRec, TaxTrans
    If PropFlag = "R" Then
      If TaxTrans.BillType <> "R" Then GoTo DoAgain
    ElseIf PropFlag = "P" Then
      If TaxTrans.BillType <> "P" Then GoTo DoAgain
    End If
    If TaxTrans.TranType = 21 Or TaxTrans.TranType = 22 Then
      TestBal# = OldRound#(TestBal# - TaxTrans.Revenue.PrePaidAmt)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 12 Then 'refund on prepay 7/17/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 9 Then 'added 8/11/06
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.PrePaidUsed)
      GoTo DoAgain
    End If
    If TaxTrans.TranType = 1 Then
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.LateList + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
      TestBal# = OldRound#(TestBal# + TaxTrans.Revenue.Interest + TaxTrans.Revenue.Penalty + TaxTrans.Revenue.Collection + TaxTrans.Revenue.RevOpt1 + TaxTrans.Revenue.RevOpt2 + TaxTrans.Revenue.RevOpt3)
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.InterestPd + TaxTrans.Revenue.PenaltyPd + TaxTrans.Revenue.CollectionPd + TaxTrans.Revenue.LateListPd))
      TestBal# = OldRound#(TestBal# - (TaxTrans.Revenue.RevOpt1Pd + TaxTrans.Revenue.RevOpt2Pd + TaxTrans.Revenue.RevOpt3Pd + TaxTrans.DiscAmt))
    End If
DoAgain:
    ThisRec = TaxTrans.LastTrans
  Loop
  
  Return
  
End Function

