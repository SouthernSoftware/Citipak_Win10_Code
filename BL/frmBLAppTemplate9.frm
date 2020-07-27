VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLAppTemplate9 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Application Renewal Template #9"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   585
   ClientWidth     =   11655
   Icon            =   "frmBLAppTemplate9.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   8415
      Left            =   1980
      TabIndex        =   15
      Top             =   45
      Width           =   7110
      _Version        =   196609
      _ExtentX        =   12541
      _ExtentY        =   14843
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483639
      Caption         =   ""
      Picture         =   "frmBLAppTemplate9.frx":08CA
      Begin LpLib.fpCombo fpcmbDiscMonth 
         Height          =   288
         Left            =   3132
         TabIndex        =   3
         Tag             =   $"frmBLAppTemplate9.frx":08E6
         Top             =   3360
         Width           =   1176
         _Version        =   196608
         _ExtentX        =   2074
         _ExtentY        =   508
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   7.5
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
         ColDesigner     =   "frmBLAppTemplate9.frx":0999
      End
      Begin LpLib.fpCombo fpcmbPenMonth 
         Height          =   288
         Left            =   3468
         TabIndex        =   6
         Tag             =   "Select from the drop down box the last month a customer can pay the business license renewal fee without penalty."
         Top             =   3696
         Width           =   1164
         _Version        =   196608
         _ExtentX        =   2053
         _ExtentY        =   508
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   7.5
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
         ColDesigner     =   "frmBLAppTemplate9.frx":0CC8
      End
      Begin LpLib.fpCombo fpcmbFiscMonth 
         Height          =   288
         Left            =   432
         TabIndex        =   8
         Tag             =   $"frmBLAppTemplate9.frx":0FF7
         Top             =   4800
         Width           =   1176
         _Version        =   196608
         _ExtentX        =   2074
         _ExtentY        =   508
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   7.5
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
         ColDesigner     =   "frmBLAppTemplate9.frx":1083
      End
      Begin LpLib.fpCombo fpcmbDiscDay 
         Height          =   288
         Left            =   4296
         TabIndex        =   4
         Tag             =   $"frmBLAppTemplate9.frx":13B2
         Top             =   3360
         Width           =   576
         _Version        =   196608
         _ExtentX        =   1016
         _ExtentY        =   508
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   7.5
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
         ColDesigner     =   "frmBLAppTemplate9.frx":146A
      End
      Begin LpLib.fpCombo fpcmbPenDay 
         Height          =   288
         Left            =   4608
         TabIndex        =   7
         Tag             =   $"frmBLAppTemplate9.frx":1799
         Top             =   3696
         Width           =   588
         _Version        =   196608
         _ExtentX        =   1037
         _ExtentY        =   508
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   7.5
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
         ColDesigner     =   "frmBLAppTemplate9.frx":1820
      End
      Begin LpLib.fpCombo fpcmbFiscDay 
         Height          =   288
         Left            =   1596
         TabIndex        =   9
         Tag             =   $"frmBLAppTemplate9.frx":1B4F
         Top             =   4800
         Width           =   564
         _Version        =   196608
         _ExtentX        =   995
         _ExtentY        =   508
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   7.5
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
         ColDesigner     =   "frmBLAppTemplate9.frx":1BD6
      End
      Begin LpLib.fpCombo fpcmbYear1 
         Height          =   288
         Left            =   3648
         TabIndex        =   1
         Tag             =   $"frmBLAppTemplate9.frx":1F05
         Top             =   720
         Width           =   540
         _Version        =   196608
         _ExtentX        =   952
         _ExtentY        =   508
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   7.5
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
         ColDesigner     =   "frmBLAppTemplate9.frx":221C
      End
      Begin LpLib.fpCombo fpcmbYear2 
         Height          =   288
         Left            =   2208
         TabIndex        =   10
         Tag             =   $"frmBLAppTemplate9.frx":254B
         Top             =   4800
         Width           =   540
         _Version        =   196608
         _ExtentX        =   952
         _ExtentY        =   508
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   7.5
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
         ColDesigner     =   "frmBLAppTemplate9.frx":2862
      End
      Begin EditLib.fpText fptxtTownOf 
         Height          =   252
         Left            =   2352
         TabIndex        =   0
         Tag             =   $"frmBLAppTemplate9.frx":2B91
         Top             =   288
         Width           =   2412
         _Version        =   196608
         _ExtentX        =   4254
         _ExtentY        =   444
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   7.5
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
      Begin EditLib.fpText fptxtDiscPct 
         Height          =   276
         Left            =   624
         TabIndex        =   2
         Tag             =   $"frmBLAppTemplate9.frx":2C43
         Top             =   3348
         Width           =   540
         _Version        =   196608
         _ExtentX        =   952
         _ExtentY        =   487
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   7.5
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
         AutoCase        =   0
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483643
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
         Text            =   "1"
         CharValidationText=   "1 2 3 4 5 6 7 8 9 0 ."
         MaxLength       =   6
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
      Begin EditLib.fpText fptxtPenPct 
         Height          =   276
         Left            =   624
         TabIndex        =   5
         Tag             =   "In this field enter the percentage amount that will be assesed to delinquent customers."
         Top             =   3672
         Width           =   540
         _Version        =   196608
         _ExtentX        =   952
         _ExtentY        =   487
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   7.5
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
         AutoCase        =   0
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483643
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
         Text            =   "1"
         CharValidationText=   "1 2 3 4 5 6 7 8 9 0 ."
         MaxLength       =   6
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
      Begin EditLib.fpText fptxtAdd 
         Height          =   252
         Left            =   3648
         TabIndex        =   11
         Tag             =   "Enter the town's mailing address in this field."
         Top             =   7584
         Width           =   2940
         _Version        =   196608
         _ExtentX        =   5186
         _ExtentY        =   444
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   7.5
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
         Height          =   252
         Left            =   3648
         TabIndex        =   12
         Tag             =   "Enter the town's mailing name here."
         Top             =   7824
         Width           =   1788
         _Version        =   196608
         _ExtentX        =   3154
         _ExtentY        =   444
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   7.5
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
      Begin EditLib.fpText fptxtState 
         Height          =   252
         Left            =   5424
         TabIndex        =   13
         Tag             =   "Enter the town's state here (NC = North Carolina)."
         Top             =   7824
         Width           =   300
         _Version        =   196608
         _ExtentX        =   529
         _ExtentY        =   444
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   7.5
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
         Height          =   252
         Left            =   5712
         TabIndex        =   14
         Tag             =   "Enter the town's postal code in this field."
         Top             =   7824
         Width           =   876
         _Version        =   196608
         _ExtentX        =   1545
         _ExtentY        =   444
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   7.5
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
      Begin VB.Label Label34 
         BackColor       =   &H80000009&
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   432
         TabIndex        =   58
         Top             =   3360
         Width           =   156
      End
      Begin VB.Label Label33 
         BackColor       =   &H80000009&
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   432
         TabIndex        =   57
         Top             =   3696
         Width           =   156
      End
      Begin VB.Label Label31 
         BackColor       =   &H80000009&
         Caption         =   ","
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   2160
         TabIndex        =   56
         Top             =   4896
         Width           =   75
      End
      Begin VB.Label Label30 
         BackColor       =   &H80000009&
         Caption         =   "Town Of"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   3648
         TabIndex        =   51
         Top             =   7344
         Width           =   2988
      End
      Begin VB.Label Label28 
         BackColor       =   &H80000009&
         Caption         =   "LICENSE NUMBER_____________"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   432
         TabIndex        =   50
         Top             =   7824
         Width           =   2748
      End
      Begin VB.Label Label27 
         BackColor       =   &H80000009&
         Caption         =   "RATE CLASS_________________"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   432
         TabIndex        =   49
         Top             =   7584
         Width           =   2748
      End
      Begin VB.Label Label26 
         BackColor       =   &H80000009&
         Caption         =   "SIC CLASS___________________"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   432
         TabIndex        =   48
         Top             =   7344
         Width           =   2796
      End
      Begin VB.Label Label25 
         BackColor       =   &H80000009&
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   5760
         TabIndex        =   47
         Top             =   6624
         Width           =   396
      End
      Begin VB.Label Label22 
         BackColor       =   &H80000009&
         Caption         =   "Title"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   3264
         TabIndex        =   46
         Top             =   6624
         Width           =   348
      End
      Begin VB.Label Label15 
         BackColor       =   &H80000009&
         Caption         =   "all information requested will result in an audit from all required sources."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   432
         TabIndex        =   45
         Top             =   6048
         Width           =   6348
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000009&
         Caption         =   "application. All licenses are subject to being audited. Failure to provide"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   432
         TabIndex        =   44
         Top             =   5808
         Width           =   6108
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000009&
         Caption         =   "provides for penalties of making false or fraudulent statements in this"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   432
         TabIndex        =   43
         Top             =   5568
         Width           =   5532
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000009&
         Caption         =   "Internal Revenue of the United States. I understand that the Town Ordinance"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   432
         TabIndex        =   42
         Top             =   5328
         Width           =   6204
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000009&
         Caption         =   "the S.C. Tax Commission of Insurance Commissioner and with the Collector of"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   432
         TabIndex        =   41
         Top             =   5088
         Width           =   6156
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000009&
         Caption         =   "transacted at or through the above location for the calendar year ending"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   432
         TabIndex        =   40
         Top             =   4608
         Width           =   5772
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000009&
         Caption         =   "and that the report corresponds with the records with "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   2784
         TabIndex        =   39
         Top             =   4848
         Width           =   4188
      End
      Begin VB.Label Label20 
         BackColor       =   &H80000009&
         Caption         =   ":_____________"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   216
         Left            =   5232
         TabIndex        =   38
         Top             =   3792
         Width           =   1524
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000009&
         Caption         =   "% Penalty Per Month After"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   1248
         TabIndex        =   37
         Top             =   3744
         Width           =   2124
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         Caption         =   "Social Security Number:____________________________________________"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   435
         TabIndex        =   35
         Top             =   2400
         Width           =   6615
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         Caption         =   "Business Name: __________________________________________________"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   450
         TabIndex        =   34
         Top             =   960
         Width           =   6600
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000009&
         Caption         =   "Mailing Address:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   432
         TabIndex        =   33
         Top             =   1200
         Width           =   6300
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000009&
         Caption         =   "                            _________________________________________________________"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   432
         TabIndex        =   32
         Top             =   1440
         Width           =   6396
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000009&
         Caption         =   "Telephone Number:_______________________________________________"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   435
         TabIndex        =   31
         Top             =   1920
         Width           =   6600
      End
      Begin VB.Label Label23 
         BackColor       =   &H80000009&
         Caption         =   "Type of Business:_________________________________________________"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   435
         TabIndex        =   30
         Top             =   2160
         Width           =   6615
      End
      Begin VB.Label Label18 
         BackColor       =   &H80000009&
         Caption         =   "BUSINESS LICENSE APPLICATION"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   2400
         TabIndex        =   29
         Top             =   528
         Width           =   2364
      End
      Begin VB.Label Label19 
         BackColor       =   &H80000009&
         Caption         =   "For Year:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   2976
         TabIndex        =   28
         Top             =   768
         Width           =   684
      End
      Begin VB.Label Label21 
         BackColor       =   &H80000009&
         Caption         =   "Business Address:_________________________________________________"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   435
         TabIndex        =   27
         Top             =   1680
         Width           =   6630
      End
      Begin VB.Label Label24 
         BackColor       =   &H80000009&
         Caption         =   "Federal Identification Number:_______________________________________"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   435
         TabIndex        =   26
         Top             =   2640
         Width           =   6570
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000009&
         Caption         =   "Gross Income Previous Year:________________________________________"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   435
         TabIndex        =   25
         Top             =   2880
         Width           =   6615
      End
      Begin VB.Label Label32 
         BackColor       =   &H80000009&
         Caption         =   "License as Calculated:______________________________________________"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   435
         TabIndex        =   24
         Top             =   3120
         Width           =   6615
      End
      Begin VB.Label Label36 
         BackColor       =   &H80000009&
         Caption         =   "% Discount, If Paid by "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   1248
         TabIndex        =   23
         Top             =   3408
         Width           =   1788
      End
      Begin VB.Label Label40 
         BackColor       =   &H80000009&
         Caption         =   ":__________________"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   216
         Left            =   4992
         TabIndex        =   22
         Top             =   3456
         Width           =   1644
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000009&
         Caption         =   "TOTAL AMOUNT DUE:_____________________________________________"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   435
         TabIndex        =   21
         Top             =   4080
         Width           =   6645
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000009&
         Caption         =   "FOR OFFICE USE ONLY                               PLEASE REMIT TO:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   432
         TabIndex        =   20
         Top             =   7104
         Width           =   5916
      End
      Begin VB.Label Label17 
         BackColor       =   &H80000009&
         Caption         =   "__________________________________________________________________________"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   432
         TabIndex        =   19
         Top             =   6432
         Width           =   6300
      End
      Begin VB.Label Label43 
         BackColor       =   &H80000009&
         Caption         =   "      This is to certify that the above is a true statement of the business "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   480
         TabIndex        =   18
         Top             =   4368
         Width           =   6108
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label44 
         BackColor       =   &H80000009&
         Caption         =   "Signature "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   480
         TabIndex        =   17
         Top             =   6624
         Width           =   1980
      End
      Begin VB.Label Label49 
         BackColor       =   &H80000009&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   6960
         TabIndex        =   16
         Top             =   7248
         Width           =   60
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   696
      Left            =   9468
      TabIndex        =   52
      TabStop         =   0   'False
      Tag             =   "Press the 'Cancel' button to close this screen and return to the Town Setup screen."
      Top             =   6420
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
      ButtonDesigner  =   "frmBLAppTemplate9.frx":2CDC
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdNext 
      Height          =   684
      Left            =   9468
      TabIndex        =   53
      TabStop         =   0   'False
      Tag             =   "Press this 'Next App' button to close this application screen and open up the screen for application #1."
      Top             =   4536
      Width           =   1884
      _Version        =   131072
      _ExtentX        =   3323
      _ExtentY        =   1206
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
      ButtonDesigner  =   "frmBLAppTemplate9.frx":2EBA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   684
      Left            =   9468
      TabIndex        =   54
      TabStop         =   0   'False
      Tag             =   "Press 'Save' to save the currently active application as application #9. All fields will be committed to memory."
      Top             =   7368
      Width           =   1884
      _Version        =   131072
      _ExtentX        =   3323
      _ExtentY        =   1206
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
      ButtonDesigner  =   "frmBLAppTemplate9.frx":3099
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdLast 
      Height          =   672
      Left            =   9468
      TabIndex        =   55
      TabStop         =   0   'False
      Tag             =   "Press this 'Last App' to close this screen and open the screen for application #8."
      Top             =   5496
      Width           =   1884
      _Version        =   131072
      _ExtentX        =   3323
      _ExtentY        =   1185
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
      ButtonDesigner  =   "frmBLAppTemplate9.frx":3275
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   492
      Left            =   9456
      TabIndex        =   59
      Tag             =   $"frmBLAppTemplate9.frx":3454
      ToolTipText     =   "Press to bring up a brief help screen."
      Top             =   3360
      Width           =   1884
      _Version        =   131072
      _ExtentX        =   3323
      _ExtentY        =   868
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
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
      ButtonDesigner  =   "frmBLAppTemplate9.frx":351E
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   444
      Left            =   10032
      TabIndex        =   60
      Top             =   1152
      Width           =   780
      _Version        =   131072
      _ExtentX        =   1376
      _ExtentY        =   783
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
      ShapeRoundWidth =   195
      ShapeRoundHeight=   195
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
      MaxWidth        =   4000
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
      Left            =   9360
      TabIndex        =   61
      Top             =   4128
      Width           =   2052
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   876
      Left            =   9264
      Top             =   3156
      Width           =   2268
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Renewal Application #9 South Carolina"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   972
      Left            =   9540
      TabIndex        =   36
      Top             =   1776
      Width           =   1740
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   972
      Left            =   9396
      Top             =   1764
      Width           =   1980
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuPrnScn 
         Caption         =   "Prin&t Screen"
      End
      Begin VB.Menu mnuEXit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmBLAppTemplate9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  Unload frmBLAppTemplate9
  frmBLTownSetup.fpcmbAppType.SetFocus
End Sub

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    lblBalloon.Visible = True
    cmdHelp.Text = "F1 &Turn Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    cmdHelp.ToolTipText = ""
    fptxtTownOf.ToolTipText = ""
    fpcmbYear1.ToolTipText = ""
    fptxtDiscPct.ToolTipText = ""
    fpcmbDiscMonth.ToolTipText = ""
    fpcmbDiscDay.ToolTipText = ""
    fptxtPenPct.ToolTipText = ""
    fpcmbPenMonth.ToolTipText = ""
    fpcmbPenDay.ToolTipText = ""
    fpcmbFiscMonth.ToolTipText = ""
    fpcmbFiscDay.ToolTipText = ""
    fptxtAdd.ToolTipText = ""
    fptxtCity.ToolTipText = ""
    fptxtState.ToolTipText = ""
    fptxtZip.ToolTipText = ""
    cmdNext.ToolTipText = ""
    cmdLast.ToolTipText = ""
    cmdExit.ToolTipText = ""
    cmdSave.ToolTipText = ""
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Label1.Caption = "This application is designed for use in South Carolina. The South Carolina Insuance Commissioner is referenced in the last paragraph."
    frmBLMessageBoxJr.Show vbModal
    frmBLMessageBoxJr.Label1.Top = 450
    frmBLMessageBoxJr.Label1.Height = 1300
    frmBLMessageBoxJr.Label1.Caption = "Some of the discretionary values appearing on this page are initially supplied from the Town Setup screen. If other application templates have been used then some of the values here may have carried over from them. PLEASE REVIEW ALL values to make sure they reflect the CURRENT situation."
    frmBLMessageBoxJr.Show vbModal
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
'    cmdHelp.ToolTipText = "If Help is turned on then click to deactivate the informational balloons. If turned off then press to activate instructional balloons."
'    fptxtTownOf.ToolTipText = "Enter 'Town Of  Your Town' here."
'    fpcmbYear1.ToolTipText = "Select 'Curr' if you want the current year displayed here. Select  '+1' if you want the next year displayed here or select '-1' if you want the prior year displayed here."
'    fptxtDiscPct.ToolTipText = "Enter the early pay discount here."
'    fpcmbDiscMonth.ToolTipText = "Select the pay early month."
'    fpcmbDiscDay.ToolTipText = "Select the early pay date."
'    fptxtPenPct.ToolTipText = "Enter the penalty percentage here."
'    fpcmbPenMonth.ToolTipText = "Select the late pay month."
'    fpcmbPenDay.ToolTipText = "Select the late pay date."
'    fpcmbFiscMonth.ToolTipText = "Select the fiscal ending month."
'    fpcmbFiscDay.ToolTipText = "Select the fiscal ending date."
'    fptxtAdd.ToolTipText = "Enter your town's mailing address here."
'    fptxtCity.ToolTipText = "Enter your town's mailing name here."
'    fptxtState.ToolTipText = "Enter your town's state (ex. NC) here."
'    fptxtZip.ToolTipText = "Enter your town's zip code here."
'    cmdNext.ToolTipText = "Press to move to application template #1."
'    cmdLast.ToolTipText = "Press to move to business application #8."
'    cmdExit.ToolTipText = "Press to return to the Town Setup screen."
'    cmdSave.ToolTipText = ""
    
  End If
    

End Sub

Private Sub cmdLast_Click()
  frmBLAppTemplate8.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdSave_Click()
  Dim TownRec As TownSetUpType
  Dim THandle As Integer
  Dim x As Integer
  Dim TempCustRec As TempCustRecType
  Dim TempHandle As Integer
  Dim TempCnt As Integer
  
  On Error GoTo ERRORSTUFF
  
  If QPTrim$(fptxtAdd.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter a valid mailing address for your town."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtAdd.BackColor = &H80FFFF
    fptxtAdd.SetFocus
    Exit Sub
  End If
  
  If Val(fptxtPenPct.Text) = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter the penalty percentage."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtPenPct.BackColor = &H80FFFF
    fptxtPenPct.SetFocus
    Exit Sub
  End If

  If QPTrim$(fptxtState.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter the town's state."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtState.BackColor = &H80FFFF
    fptxtState.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtCity.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter the town's mailing name."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtCity.BackColor = &H80FFFF
    fptxtCity.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtTownOf.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter the town's official name."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtTownOf.BackColor = &H80FFFF
    fptxtTownOf.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtZip.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter the town's zip code."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtZip.BackColor = &H80FFFF
    fptxtZip.SetFocus
    Exit Sub
  End If

  If Exist("artownsu.dat") Then
    OpenTownFile THandle
    Get THandle, 1, TownRec
      TownRec.AppAdd1 = QPTrim(fptxtAdd.Text)
      TownRec.AppDiscPct = CDbl(fptxtDiscPct.Text)
      TownRec.AppDiscMonth = QPTrim$(fpcmbDiscMonth.Text)
      TownRec.AppDiscDay = CInt(fpcmbDiscDay.Text)
      TownRec.AppPenMonth = QPTrim$(fpcmbPenMonth.Text)
      TownRec.AppPenDay = CInt(fpcmbPenDay.Text)
      TownRec.AppFiscMonth = QPTrim$(fpcmbFiscMonth.Text)
      TownRec.AppFiscDay = CInt(fpcmbFiscDay.Text)
      TownRec.AppPct = CDbl(fptxtPenPct.Text)
      TownRec.AppState = QPTrim$(fptxtState.Text)
      TownRec.AppCity = QPTrim$(fptxtCity.Text)
      TownRec.AppTownOf = QPTrim$(fptxtTownOf.Text)
      TownRec.AppZip = QPTrim$(fptxtZip.Text)
      TownRec.AppYrUpDown(1) = fpcmbYear1.Text
      TownRec.AppYrUpDown(2) = fpcmbYear2.Text
      TownRec.AppForm = 9
    Put THandle, 1, TownRec
  Else
    TownRec.TownName = ""
    TownRec.Contact = ""
    TownRec.TownAdd1 = ""
    TownRec.TownAdd2 = ""
    TownRec.City = ""
    TownRec.State = ""
    TownRec.ZipCode = ""
    TownRec.TownPhone = ""
    TownRec.SpareSpace = ""
    TownRec.AppForm = 9
    TownRec.DLQNotice = 0
    TownRec.AppAdd1 = QPTrim(fptxtAdd.Text)
    TownRec.AppBaseFee(1) = 0
    TownRec.AppBaseFee(2) = 0
    TownRec.AppBaseFee(3) = 0
    TownRec.AppBaseFee(4) = 0
    TownRec.AppCentsPer(1) = 0
    TownRec.AppCentsPer(2) = 0
    TownRec.AppCentsPer(3) = 0
    TownRec.AppCentsPer(4) = 0
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
    TownRec.AppState = QPTrim$(fptxtState.Text)
    TownRec.AppCity = QPTrim$(fptxtCity.Text)
    TownRec.AppTownOf = QPTrim$(fptxtTownOf.Text)
    TownRec.AppZip = QPTrim$(fptxtZip.Text) '30
    TownRec.AppPct = CDbl(fptxtPenPct.Text)
    TownRec.AppAdminName = ""
    TownRec.AppAdminTitle = ""
    TownRec.AppPhone = ""
    TownRec.AppDiscPct = CDbl(fptxtDiscPct.Text)
    TownRec.AppDiscMonth = QPTrim$(fpcmbDiscMonth.Text)
    TownRec.AppDiscDay = CInt(fpcmbDiscDay.Text)
    TownRec.AppPenMonth = QPTrim$(fpcmbPenMonth.Text)
    TownRec.AppPenDay = CInt(fpcmbPenDay.Text)
    TownRec.AppFiscMonth = QPTrim$(fpcmbFiscMonth.Text)
    TownRec.AppFiscDay = CInt(fpcmbFiscDay.Text)
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
    TownRec.AppPayBy = 0
    TownRec.AppCityOrd = ""
    TownRec.AppYrUpDown(1) = fpcmbYear1.Text
    TownRec.AppYrUpDown(2) = fpcmbYear2.Text
    For x = 3 To 10
     TownRec.AppYrUpDown(x) = "0"
    Next x
    TownRec.DlqAdd1 = ""
    TownRec.DlqAdminName = ""
    TownRec.DlqAdminTitle = ""
    TownRec.DlqCity = ""
    TownRec.DlqPhone = ""
    TownRec.DlqPhone2 = ""
    TownRec.DlqFax = "" '40
    TownRec.DlqState = ""
    TownRec.DlqTownName = ""
    TownRec.DlqZip = ""
    TownRec.DlqFirstDay = ""
    TownRec.DlqLastDay = ""
    TownRec.DlqFirstHour = ""
    TownRec.DlqLastHour = ""
    TownRec.DlqClerkName = ""
    TownRec.DlqMayorCouncil = "" '49
    TownRec.LicNumPermYN = "No"
    TownRec.UseAmtPctYN = "Pct"
    TownRec.PENCASHACCT = 0
    TownRec.PENRECGLNUM = 0
    TownRec.PENREVGLNUM = 0
    TownRec.IssFee = 0
    TownRec.AcctMeth = ""
    TownRec.LaserLtr = "N"
    TownRec.GL2Cats = "N"
    OpenTownFile THandle
    Put THandle, 1, TownRec
  End If
  Close THandle
  
  'added as a precaution to prevent the user from running application
  'renewal form #9 then coming here to save different data and then
  'trying to run application renewal reprints which will use this
  'latest saved data while the originals have the old data...now the
  'user will have to print applications over
  If Exist("artmpcus.dat") Then
    OpenTempCustRec TempHandle
    TempCnt = LOF(TempHandle) / Len(TempCustRec)
    If TempCnt > 0 Then
      Get TempHandle, 1, TempCustRec
      Close TempHandle
      If TempCustRec.AppType = 9 Then
        KillFile "artmpcus.dat"
      End If
    Else
      Close TempHandle
    End If
  End If
  
  frmBLSucSave.Label1.Caption = "Your renewal application notice #9 data has been saved successfully."
  frmBLSucSave.Label1.Top = 700
  frmBLSucSave.Show vbModal
  Call cmdExit_Click
  frmBLTownSetup.fpcmbAppType.Text = "9. APP FORM H"
  frmBLTownSetup.fpcmdApps.Text = "F3 S&how App Type 9"
  
  MainLog ("Application #9 saved.")
  
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLAppTemplate9", "cmdSave_Click", Erl)
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
  Call LoadMe
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
      SendKeys "%N"
      Call cmdNext_Click
      KeyCode = 0
    Case vbKeyF1:
      SendKeys "%L"
      Call cmdLast_Click
      KeyCode = 0
    Case vbKeyF3:
      SendKeys "%H"
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
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLAppTemplate9.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  Dim TownRec As TownSetUpType
  Dim THandle As Integer
  Dim x As Integer
  
  On Error GoTo ERRORSTUFF
  
  lblBalloon.Visible = False
'  fptxtTownOf.ToolTipText = "Enter 'Town Of  Your Town' here."
'  fpcmbYear1.ToolTipText = "Select 'Curr' if you want the current year displayed here. Select  '+1' if you want the next year displayed here or select '-1' if you want the prior year displayed here."
'  fptxtDiscPct.ToolTipText = "Enter the early pay discount here."
'  fpcmbDiscMonth.ToolTipText = "Select the pay early month."
'  fpcmbDiscDay.ToolTipText = "Select the early pay date."
'  fptxtPenPct.ToolTipText = "Enter the penalty percentage here."
'  fpcmbPenMonth.ToolTipText = "Select the late pay month."
'  fpcmbPenDay.ToolTipText = "Select the late pay date."
'  fpcmbFiscMonth.ToolTipText = "Select the fiscal ending month."
'  fpcmbFiscDay.ToolTipText = "Select the fiscal ending date."
'  fptxtAdd.ToolTipText = "Enter your town's mailing address here."
'  fptxtCity.ToolTipText = "Enter your town's mailing name here."
'  fptxtState.ToolTipText = "Enter your town's state (ex. NC) here."
'  fptxtZip.ToolTipText = "Enter your town's zip code here."
'  cmdNext.ToolTipText = "Press to move to application template #1."
'  cmdLast.ToolTipText = "Press to move to business application #8."
'  cmdExit.ToolTipText = "Press to return to the Town Setup screen."
'  cmdSave.ToolTipText = ""
  If QPTrim$(frmBLTownSetup.fpcmbAmtPct.Text) = "Amt" Then
    Label34.Caption = "$"
    Label36.Caption = " Discount, If Paid by "
    Label33.Caption = "$"
    Label7.Caption = " Penalty Per Month After"
    fptxtPenPct.ToolTipText = "Enter the penalty amount here."
  Else
    Label34.Caption = ""
    Label36.Caption = "% Discount, If Paid by "
    Label33.Caption = ""
    Label7.Caption = "% Penalty Per Month After "
    fptxtPenPct.ToolTipText = "Enter the penalty percentage here."
  End If
  
  If Exist("artownsu.dat") Then
    OpenTownFile THandle
    Get THandle, 1, TownRec
    Close THandle
    If QPTrim(TownRec.AppTownOf) = "" Then
      If QPTrim$(frmBLTownSetup.fptxtTownName.Text) <> "" Then
        fptxtTownOf.Text = QPTrim$(frmBLTownSetup.fptxtTownName.Text)
      Else
        fptxtTownOf.Text = "Town Of 'Your Town'"
      End If
    Else
      fptxtTownOf.Text = QPTrim(TownRec.AppTownOf)
    End If
    
    Label30.Caption = QPTrim$(fptxtTownOf.Text)
    
    If QPTrim(TownRec.AppAdd1) = "" Then
      If QPTrim$(frmBLTownSetup.fptxtAdd1.Text) <> "" Then
        fptxtAdd.Text = QPTrim$(frmBLTownSetup.fptxtAdd1.Text)
      Else
        fptxtAdd.Text = "Street Address"
      End If
    Else
      fptxtAdd.Text = QPTrim(TownRec.AppAdd1)
    End If
    
    If QPTrim(TownRec.AppCity) = "" Then
      If QPTrim$(frmBLTownSetup.fptxtCity.Text) <> "" Then
        fptxtCity.Text = QPTrim$(frmBLTownSetup.fptxtCity.Text)
      Else
        fptxtCity.Text = "Your Town"
      End If
    Else
      fptxtCity.Text = QPTrim(TownRec.AppCity)
    End If

    If QPTrim(TownRec.AppState) = "" Then
      If QPTrim$(frmBLTownSetup.fptxtState.Text) <> "" Then
        fptxtState.Text = QPTrim$(frmBLTownSetup.fptxtState.Text)
      Else
        fptxtState.Text = "ST"
      End If
    Else
      fptxtState.Text = QPTrim(TownRec.AppState)
    End If
    
    If QPTrim(TownRec.AppZip) = "" Then
      If QPTrim$(frmBLTownSetup.fptxtZip.Text) <> "" Then
        fptxtZip.Text = QPTrim$(frmBLTownSetup.fptxtZip.Text)
      Else
        fptxtZip.Text = "11111-1111"
      End If
    Else
      fptxtZip.Text = QPTrim(TownRec.AppZip)
    End If
    
    fptxtDiscPct.Text = TownRec.AppDiscPct
    
    If QPTrim$(TownRec.AppDiscMonth) = "" Then
      fpcmbDiscMonth.Text = "January"
    Else
      fpcmbDiscMonth.Text = QPTrim$(TownRec.AppDiscMonth)
    End If
    
    If TownRec.AppDiscDay = 0 Then
      fpcmbDiscDay.Text = 1
    Else
      fpcmbDiscDay.Text = TownRec.AppDiscDay
    End If
    
    fptxtPenPct.Text = TownRec.AppPct
    
    If QPTrim$(TownRec.AppPenMonth) = "" Then
      fpcmbPenMonth.Text = "January"
    Else
      fpcmbPenMonth.Text = QPTrim$(TownRec.AppPenMonth)
    End If
    
    If TownRec.AppPenDay = 0 Then
      fpcmbPenDay.Text = 1
    Else
      fpcmbPenDay.Text = TownRec.AppPenDay
    End If
    
    If QPTrim$(TownRec.AppFiscMonth) = "" Then
      fpcmbFiscMonth.Text = "January"
    Else
      fpcmbFiscMonth.Text = QPTrim$(TownRec.AppFiscMonth)
    End If
    
    If TownRec.AppFiscDay = 0 Then
      fpcmbFiscDay.Text = 1
    Else
      fpcmbFiscDay.Text = TownRec.AppFiscDay
    End If
    
    For x = 1 To 2
      If QPTrim$(TownRec.AppYrUpDown(x)) = "0" Then TownRec.AppYrUpDown(x) = "Curr"
    Next x
    
    fpcmbYear1.Text = TownRec.AppYrUpDown(1)
    fpcmbYear2.Text = TownRec.AppYrUpDown(2)
  Else
    If QPTrim$(frmBLTownSetup.fptxtTownName.Text) <> "" Then
      fptxtTownOf.Text = QPTrim$(frmBLTownSetup.fptxtTownName.Text)
    Else
      fptxtTownOf.Text = "Town Of 'Your Town'"
    End If
    
    Label30.Caption = QPTrim$(fptxtTownOf.Text)
    
    If QPTrim$(frmBLTownSetup.fptxtAdd1.Text) <> "" Then
      fptxtAdd.Text = QPTrim$(frmBLTownSetup.fptxtAdd1.Text)
    Else
      fptxtAdd.Text = "Street Address"
    End If
    
    If QPTrim$(frmBLTownSetup.fptxtCity.Text) <> "" Then
      fptxtCity.Text = QPTrim$(frmBLTownSetup.fptxtCity.Text)
    Else
      fptxtCity.Text = "Your Town"
    End If

    If QPTrim$(frmBLTownSetup.fptxtState.Text) <> "" Then
      fptxtState.Text = QPTrim$(frmBLTownSetup.fptxtState.Text)
    Else
      fptxtState.Text = "ST"
    End If
    
    If QPTrim$(frmBLTownSetup.fptxtZip.Text) <> "" Then
      fptxtZip.Text = QPTrim$(frmBLTownSetup.fptxtZip.Text)
    Else
      fptxtZip.Text = "11111-1111"
    End If
    fptxtDiscPct.Text = "?"
    fpcmbDiscMonth.Text = "January"
    fpcmbDiscDay.Text = "1"
    fptxtPenPct.Text = "?"
    fpcmbPenMonth.Text = "December"
    fpcmbPenDay.Text = "31"
    fpcmbFiscMonth.Text = "December"
    fpcmbFiscDay.Text = "31"
    fpcmbYear1.Text = "Curr"
    fpcmbYear2.Text = "Curr"
  End If
  
  For x = 1 To 12
    Select Case x
      Case 1
        fpcmbDiscMonth.AddItem "January"
        fpcmbPenMonth.AddItem "January"
        fpcmbFiscMonth.AddItem "January"
      Case 2
        fpcmbDiscMonth.AddItem "February"
        fpcmbPenMonth.AddItem "February"
        fpcmbFiscMonth.AddItem "February"
      Case 3
        fpcmbDiscMonth.AddItem "March"
        fpcmbPenMonth.AddItem "March"
        fpcmbFiscMonth.AddItem "March"
      Case 4
        fpcmbDiscMonth.AddItem "April"
        fpcmbPenMonth.AddItem "April"
        fpcmbFiscMonth.AddItem "April"
      Case 5
        fpcmbDiscMonth.AddItem "May"
        fpcmbPenMonth.AddItem "May"
        fpcmbFiscMonth.AddItem "May"
      Case 6
        fpcmbDiscMonth.AddItem "June"
        fpcmbPenMonth.AddItem "June"
        fpcmbFiscMonth.AddItem "June"
      Case 7
        fpcmbDiscMonth.AddItem "July"
        fpcmbPenMonth.AddItem "July"
        fpcmbFiscMonth.AddItem "July"
      Case 8
        fpcmbDiscMonth.AddItem "August"
        fpcmbPenMonth.AddItem "August"
        fpcmbFiscMonth.AddItem "August"
      Case 9
        fpcmbDiscMonth.AddItem "September"
        fpcmbPenMonth.AddItem "September"
        fpcmbFiscMonth.AddItem "September"
      Case 10
        fpcmbDiscMonth.AddItem "October"
        fpcmbPenMonth.AddItem "October"
        fpcmbFiscMonth.AddItem "October"
      Case 11
        fpcmbDiscMonth.AddItem "November"
        fpcmbPenMonth.AddItem "November"
        fpcmbFiscMonth.AddItem "November"
      Case 12
        fpcmbDiscMonth.AddItem "December"
        fpcmbPenMonth.AddItem "December"
        fpcmbFiscMonth.AddItem "December"
    End Select
  Next x

  For x = 1 To 31
    fpcmbDiscDay.AddItem CStr(x)
    fpcmbPenDay.AddItem CStr(x)
    fpcmbFiscDay.AddItem CStr(x)
  Next x
  
  
  fpcmbYear1.AddItem "Curr"
  fpcmbYear1.AddItem "+1"
  fpcmbYear1.AddItem "-1"
  fpcmbYear2.AddItem "Curr"
  fpcmbYear2.AddItem "+1"
  fpcmbYear2.AddItem "-1"

  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLAppTemplate9", "LoadMe", Erl)
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

Private Sub fpcmbDiscDay_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  fpcmbDiscDay.BackColor = &H80000005
  If KeyCode = vbKeySpace Then
    fpcmbDiscDay.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbDiscDay.ListIndex = -1
  End If
  If fpcmbDiscDay.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtPenPct.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbDiscMonth_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  fpcmbDiscMonth.BackColor = &H80000005
  If KeyCode = vbKeySpace Then
    fpcmbDiscMonth.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbDiscMonth.ListIndex = -1
  End If
  If fpcmbDiscMonth.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbDiscDay.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbFiscDay_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  fpcmbFiscDay.BackColor = &H80000005
  If KeyCode = vbKeySpace Then
    fpcmbFiscDay.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbFiscDay.ListIndex = -1
  End If
  If fpcmbFiscDay.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbYear2.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbFiscMonth_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  fpcmbFiscMonth.BackColor = &H80000005
  If KeyCode = vbKeySpace Then
    fpcmbFiscMonth.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbFiscMonth.ListIndex = -1
  End If
  If fpcmbFiscMonth.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbFiscDay.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbPenDay_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  fpcmbPenDay.BackColor = &H80000005
  If KeyCode = vbKeySpace Then
    fpcmbPenDay.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPenDay.ListIndex = -1
  End If
  If fpcmbPenDay.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbFiscDay.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbPenMonth_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  fpcmbPenMonth.BackColor = &H80000005
  If KeyCode = vbKeySpace Then
    fpcmbPenMonth.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPenMonth.ListIndex = -1
  End If
  If fpcmbPenMonth.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbPenDay.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fptxtAdd_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtAdd.BackColor = &H80000005

End Sub

Private Sub fptxtCity_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtCity.BackColor = &H80000005

End Sub

Private Sub fptxtDiscPct_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtDiscPct.BackColor = &H80000005

End Sub

Private Sub fptxtPenPct_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtPenPct.BackColor = &H80000005

End Sub

Private Sub fptxtState_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtState.BackColor = &H80000005

End Sub

Private Sub fptxtTownOf_Change()
  Label30.Caption = QPTrim$(fptxtTownOf.Text)
End Sub

Private Sub fptxtTownOf_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtTownOf.BackColor = &H80000005

End Sub

Private Sub fptxtZip_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtZip.BackColor = &H80000005

End Sub

Private Sub cmdNext_Click()
  frmBLFreeFormatApp1.Show
  DoEvents
  Unload Me
End Sub

Private Sub fpcmbYear1_KeyDown(KeyCode As Integer, Shift As Integer)
  fpcmbYear1.BackColor = -2147483643
  If KeyCode = vbKeySpace Then
    fpcmbYear1.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbYear1.ListIndex = -1
  End If
  If fpcmbYear1.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtDiscPct.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbYear2_KeyDown(KeyCode As Integer, Shift As Integer)
  fpcmbYear2.BackColor = -2147483643
  If KeyCode = vbKeySpace Then
    fpcmbYear2.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbYear2.ListIndex = -1
  End If
  If fpcmbYear2.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtAdd.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

