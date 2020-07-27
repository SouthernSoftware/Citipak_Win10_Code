VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLAppTemplate3 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Application Renewal Template #3"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   585
   ClientWidth     =   11655
   Icon            =   "frmBLAppTemplate3.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   8535
      Left            =   1965
      TabIndex        =   23
      Top             =   45
      Width           =   7110
      _Version        =   196609
      _ExtentX        =   12541
      _ExtentY        =   15055
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
      Picture         =   "frmBLAppTemplate3.frx":08CA
      Begin LpLib.fpCombo fpcmbEndDay 
         Height          =   300
         Left            =   4995
         TabIndex        =   13
         Tag             =   "From this drop down box select the last day after which license renewal becomes invalid."
         Top             =   5355
         Width           =   540
         _Version        =   196608
         _ExtentX        =   952
         _ExtentY        =   529
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
         ColDesigner     =   "frmBLAppTemplate3.frx":08E6
      End
      Begin LpLib.fpCombo fpcmbEndMonth 
         Height          =   300
         Left            =   4470
         TabIndex        =   12
         Tag             =   "From this drop down box select the month after which the license renewal becomes invalid."
         Top             =   5355
         Width           =   540
         _Version        =   196608
         _ExtentX        =   952
         _ExtentY        =   529
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
         ColDesigner     =   "frmBLAppTemplate3.frx":0C15
      End
      Begin LpLib.fpCombo fpcmbStartDay 
         Height          =   300
         Left            =   3120
         TabIndex        =   11
         Tag             =   "From this drop down box select the first day the license renewal becomes valid."
         Top             =   5355
         Width           =   540
         _Version        =   196608
         _ExtentX        =   952
         _ExtentY        =   529
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
         ColDesigner     =   "frmBLAppTemplate3.frx":0F44
      End
      Begin LpLib.fpCombo fpcmbStartMonth 
         Height          =   300
         Left            =   2595
         TabIndex        =   10
         Tag             =   "From this drop down box select the month that the license renewal becomes valid."
         Top             =   5355
         Width           =   540
         _Version        =   196608
         _ExtentX        =   952
         _ExtentY        =   529
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
         ColDesigner     =   "frmBLAppTemplate3.frx":1273
      End
      Begin LpLib.fpCombo fpcmbPayByMonth 
         Height          =   300
         Left            =   4995
         TabIndex        =   16
         Tag             =   "Select a month from this drop down list that will be the final month that the license renewal can be paid without a penalty."
         Top             =   7035
         Width           =   540
         _Version        =   196608
         _ExtentX        =   952
         _ExtentY        =   529
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
         ColDesigner     =   "frmBLAppTemplate3.frx":15A2
      End
      Begin LpLib.fpCombo fpcmbPayByDay 
         Height          =   300
         Left            =   5520
         TabIndex        =   17
         Tag             =   "Select a day from the drop down list that is the final day the license renewal fee can be paid without a penalty."
         Top             =   7035
         Width           =   540
         _Version        =   196608
         _ExtentX        =   952
         _ExtentY        =   529
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
         ColDesigner     =   "frmBLAppTemplate3.frx":18D1
      End
      Begin LpLib.fpCombo fpcmbYear1 
         Height          =   300
         Left            =   3795
         TabIndex        =   1
         Tag             =   $"frmBLAppTemplate3.frx":1C00
         Top             =   450
         Width           =   540
         _Version        =   196608
         _ExtentX        =   952
         _ExtentY        =   529
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
         ColDesigner     =   "frmBLAppTemplate3.frx":1EC9
      End
      Begin LpLib.fpCombo fpcmbYear2 
         Height          =   300
         Left            =   5610
         TabIndex        =   14
         Tag             =   $"frmBLAppTemplate3.frx":21F8
         Top             =   5355
         Width           =   540
         _Version        =   196608
         _ExtentX        =   952
         _ExtentY        =   529
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
         ColDesigner     =   "frmBLAppTemplate3.frx":24C1
      End
      Begin LpLib.fpCombo fpcmbYear3 
         Height          =   300
         Left            =   6150
         TabIndex        =   18
         Tag             =   $"frmBLAppTemplate3.frx":27F0
         Top             =   7035
         Width           =   540
         _Version        =   196608
         _ExtentX        =   952
         _ExtentY        =   529
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
         ColDesigner     =   "frmBLAppTemplate3.frx":2AB9
      End
      Begin EditLib.fpCurrency fptxtBaseFee1 
         Height          =   255
         Left            =   3030
         TabIndex        =   2
         Tag             =   "Enter an amount here that corresponds to a 'Flat Rate' license fee calculations method for a category code."
         Top             =   2805
         Width           =   690
         _Version        =   196608
         _ExtentX        =   1206
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
      Begin EditLib.fpText fptxtTownOf 
         Height          =   255
         Left            =   2490
         TabIndex        =   0
         Tag             =   $"frmBLAppTemplate3.frx":2DE8
         Top             =   30
         Width           =   2415
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
      Begin EditLib.fpText fptxtAdd 
         Height          =   255
         Left            =   435
         TabIndex        =   19
         Tag             =   "Enter the address where the town receives it's mail in this field."
         Top             =   7800
         Width           =   3030
         _Version        =   196608
         _ExtentX        =   5355
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
      Begin EditLib.fpText fptxtTownName 
         Height          =   255
         Left            =   435
         TabIndex        =   20
         Tag             =   "Enter the town's mailing name in this field."
         Top             =   8040
         Width           =   2310
         _Version        =   196608
         _ExtentX        =   4085
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
      Begin EditLib.fpText fptxtState 
         Height          =   255
         Left            =   2880
         TabIndex        =   21
         Tag             =   "Enter the town's state here (ex. VA)."
         Top             =   8040
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
         Height          =   255
         Left            =   3315
         TabIndex        =   22
         Tag             =   "Enter the town's postal code in this field."
         Top             =   8040
         Width           =   870
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
      Begin EditLib.fpText fptxtPct 
         Height          =   255
         Left            =   2685
         TabIndex        =   15
         Tag             =   $"frmBLAppTemplate3.frx":2EBB
         Top             =   7080
         Width           =   450
         _Version        =   196608
         _ExtentX        =   783
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
         Text            =   ""
         CharValidationText=   "0 1 2 3 4 5 6 7 8 9 ."
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
      Begin EditLib.fpText fptxtNumerator 
         Height          =   255
         Left            =   2970
         TabIndex        =   4
         Tag             =   $"frmBLAppTemplate3.frx":2FAB
         Top             =   3090
         Width           =   255
         _Version        =   196608
         _ExtentX        =   444
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
         Text            =   ""
         CharValidationText=   "0 1 2 3 4 5 6 7 8 9 "
         MaxLength       =   3
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
      Begin EditLib.fpText fptxtDenominator 
         Height          =   255
         Left            =   3360
         TabIndex        =   5
         Tag             =   $"frmBLAppTemplate3.frx":30B9
         Top             =   3090
         Width           =   255
         _Version        =   196608
         _ExtentX        =   444
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
         Text            =   ""
         CharValidationText=   "0 1 2 3 4 5 6 7 8 9 "
         MaxLength       =   5
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
      Begin EditLib.fpText fptxtGrsPct 
         Height          =   255
         Left            =   3840
         TabIndex        =   6
         Tag             =   $"frmBLAppTemplate3.frx":31D7
         Top             =   3090
         Width           =   255
         _Version        =   196608
         _ExtentX        =   444
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
         Text            =   ""
         CharValidationText=   "0 1 2 3 4 5 6 7 8 9 "
         MaxLength       =   3
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
      Begin EditLib.fpCurrency fptxtBaseFee2 
         Height          =   255
         Left            =   1965
         TabIndex        =   3
         Tag             =   $"frmBLAppTemplate3.frx":32EA
         Top             =   3090
         Width           =   690
         _Version        =   196608
         _ExtentX        =   1206
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
      Begin EditLib.fpCurrency fptxtBaseFee3 
         Height          =   255
         Left            =   4275
         TabIndex        =   8
         Tag             =   "Enter an amount here that corresponds to a 'Flat Rate' license fee calculations method for a category code."
         Top             =   3570
         Width           =   690
         _Version        =   196608
         _ExtentX        =   1206
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
      Begin EditLib.fpCurrency fptxtBaseFee4 
         Height          =   255
         Left            =   4320
         TabIndex        =   9
         Tag             =   "Enter an amount here that corresponds to a 'Flat Rate' license fee calculations method for a category code."
         Top             =   4005
         Width           =   690
         _Version        =   196608
         _ExtentX        =   1206
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
      Begin EditLib.fpCurrency fptxtGross1 
         Height          =   255
         Left            =   1485
         TabIndex        =   7
         Tag             =   $"frmBLAppTemplate3.frx":337A
         Top             =   3330
         Width           =   930
         _Version        =   196608
         _ExtentX        =   1630
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
      Begin VB.Label fplblIssfee4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Caption         =   "$XX.XX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5370
         TabIndex        =   81
         Top             =   4050
         Width           =   540
      End
      Begin VB.Label fplblIssFee3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Caption         =   "$XX.XX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5280
         TabIndex        =   80
         Top             =   3630
         Width           =   585
      End
      Begin VB.Label fplblIssFee2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Caption         =   "$XX.XX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2790
         TabIndex        =   79
         Top             =   3390
         Width           =   585
      End
      Begin VB.Label fplblIssFee1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Caption         =   "$XX.XX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4125
         TabIndex        =   78
         Top             =   2850
         Width           =   585
      End
      Begin VB.Label Label14 
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
         Height          =   210
         Left            =   6045
         TabIndex        =   77
         Top             =   7125
         Width           =   105
      End
      Begin VB.Label Label12 
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
         Height          =   210
         Left            =   5520
         TabIndex        =   76
         Top             =   5445
         Width           =   105
      End
      Begin VB.Label Label51 
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
         Height          =   210
         Left            =   6720
         TabIndex        =   71
         Top             =   7125
         Width           =   105
      End
      Begin VB.Label Label49 
         BackColor       =   &H80000009&
         Caption         =   "over"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   1155
         TabIndex        =   70
         Top             =   3390
         Width           =   300
      End
      Begin VB.Label Label38 
         BackColor       =   &H80000009&
         Caption         =   "Issuance Fee."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1155
         TabIndex        =   69
         Top             =   4245
         Width           =   1065
      End
      Begin VB.Label Label35 
         BackColor       =   &H80000009&
         Caption         =   "plus"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5040
         TabIndex        =   68
         Top             =   4050
         Width           =   300
      End
      Begin VB.Label Label37 
         BackColor       =   &H80000009&
         Caption         =   "plus"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4995
         TabIndex        =   67
         Top             =   3630
         Width           =   345
      End
      Begin VB.Label Label34 
         BackColor       =   &H80000009&
         Caption         =   "Issuance Fee."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1155
         TabIndex        =   66
         Top             =   3810
         Width           =   1065
      End
      Begin VB.Label Label33 
         BackColor       =   &H80000009&
         Caption         =   "plus"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2445
         TabIndex        =   65
         Top             =   3390
         Width           =   300
      End
      Begin VB.Label Label30 
         BackColor       =   &H80000009&
         Caption         =   "Issuance Fee."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3510
         TabIndex        =   64
         Top             =   3390
         Width           =   1065
      End
      Begin VB.Label Label28 
         BackColor       =   &H80000009&
         Caption         =   "of"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3645
         TabIndex        =   63
         Top             =   3150
         Width           =   150
      End
      Begin VB.Label Label26 
         BackColor       =   &H80000009&
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3270
         TabIndex        =   62
         Top             =   3150
         Width           =   105
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000009&
         Caption         =   "plus "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2685
         TabIndex        =   61
         Top             =   3150
         Width           =   300
      End
      Begin VB.Label Label25 
         BackColor       =   &H80000009&
         Caption         =   "plus "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3750
         TabIndex        =   60
         Top             =   2850
         Width           =   300
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000009&
         Caption         =   "Estimate of ______________ gross receipts or preceding year's gross"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   435
         TabIndex        =   58
         Top             =   4725
         Width           =   6300
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         Caption         =   "________________________________________________Phone:(___)___-___________"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   435
         TabIndex        =   57
         Top             =   2370
         Width           =   6210
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         Caption         =   "Business Name: ___________________________________________________________"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   435
         TabIndex        =   56
         Top             =   690
         Width           =   6150
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000009&
         Caption         =   "Street Address of Business:__________________________________________________"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   435
         TabIndex        =   55
         Top             =   930
         Width           =   6300
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000009&
         Caption         =   "Zoning of Business Location:_________________________________________________"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   435
         TabIndex        =   54
         Top             =   1170
         Width           =   6390
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000009&
         Caption         =   "Telephone Number:________________________________________________________"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   435
         TabIndex        =   53
         Top             =   1410
         Width           =   6150
      End
      Begin VB.Label Label23 
         BackColor       =   &H80000009&
         Caption         =   "Applicant's Name:__________________________________________________________"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   435
         TabIndex        =   52
         Top             =   1890
         Width           =   6210
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
         Height          =   210
         Left            =   2550
         TabIndex        =   51
         Top             =   270
         Width           =   2370
      End
      Begin VB.Label Label19 
         BackColor       =   &H80000009&
         Caption         =   "For Year: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3120
         TabIndex        =   50
         Top             =   510
         Width           =   690
      End
      Begin VB.Label Label20 
         BackColor       =   &H80000009&
         Caption         =   "_________________________________________________________________________"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   435
         TabIndex        =   49
         Top             =   1650
         Width           =   6105
      End
      Begin VB.Label Label21 
         BackColor       =   &H80000009&
         Caption         =   "Applicant's Address:________________________________________________________"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   435
         TabIndex        =   48
         Top             =   2130
         Width           =   6540
      End
      Begin VB.Label Label22 
         BackColor       =   &H80000009&
         Caption         =   "TYPE OF BUSINESS LICENSE APPLYING FOR:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   435
         TabIndex        =   47
         Top             =   2610
         Width           =   6210
      End
      Begin VB.Label Label24 
         BackColor       =   &H80000009&
         Caption         =   "________ Contracting or Construction"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   435
         TabIndex        =   46
         Top             =   2850
         Width           =   2610
      End
      Begin VB.Label Label27 
         BackColor       =   &H80000009&
         Caption         =   "Issuance Fee."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4800
         TabIndex        =   45
         Top             =   2850
         Width           =   1065
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000009&
         Caption         =   "________ Retail Sales "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   435
         TabIndex        =   44
         Top             =   3150
         Width           =   1545
      End
      Begin VB.Label Label31 
         BackColor       =   &H80000009&
         Caption         =   "% of gross receipts"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4200
         TabIndex        =   43
         Top             =   3150
         Width           =   1410
      End
      Begin VB.Label Label32 
         BackColor       =   &H80000009&
         Caption         =   "________ Financial, Real Estate or Professional Service"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   435
         TabIndex        =   42
         Top             =   3630
         Width           =   3810
      End
      Begin VB.Label Label36 
         BackColor       =   &H80000009&
         Caption         =   "________ Repair, Personal, Business or Delivery Service"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   435
         TabIndex        =   41
         Top             =   4050
         Width           =   3900
      End
      Begin VB.Label Label40 
         BackColor       =   &H80000009&
         Caption         =   "________ Other (Specify) ____________________________________________________"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   435
         TabIndex        =   40
         Top             =   4485
         Width           =   6390
      End
      Begin VB.Label Label41 
         BackColor       =   &H80000009&
         Caption         =   "receipts _____________________. Enclose copy of most recent schedule C"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   435
         TabIndex        =   39
         Top             =   4920
         Width           =   6300
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000009&
         Caption         =   "or other comparable federal document."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   435
         TabIndex        =   38
         Top             =   5160
         Width           =   6300
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000009&
         Caption         =   "AMOUNT OF LICENSE TAX FOR "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   435
         TabIndex        =   37
         Top             =   5445
         Width           =   2175
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000009&
         Caption         =   ", THROUGH"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3645
         TabIndex        =   36
         Top             =   5445
         Width           =   780
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000009&
         Caption         =   "IS:______"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6195
         TabIndex        =   35
         Top             =   5445
         Width           =   780
      End
      Begin VB.Label Label15 
         BackColor       =   &H80000009&
         Caption         =   "ANY SPECIAL CONDITION OR REQUIREMENTS, IF ANY, UNDER WHICH LICENSED"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   435
         TabIndex        =   34
         Top             =   5685
         Width           =   6210
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000009&
         Caption         =   "ACTIVITY SHALL BE CONDUCTED: ______________________________________________"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   435
         TabIndex        =   33
         Top             =   5925
         Width           =   6300
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
         Height          =   210
         Left            =   435
         TabIndex        =   32
         Top             =   6120
         Width           =   6300
      End
      Begin VB.Label Label42 
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
         Height          =   210
         Left            =   435
         TabIndex        =   31
         Top             =   6315
         Width           =   6300
      End
      Begin VB.Label Label43 
         BackColor       =   &H80000009&
         Caption         =   $"frmBLAppTemplate3.frx":3418
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   480
         TabIndex        =   30
         Top             =   6510
         Width           =   6210
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label44 
         BackColor       =   &H80000009&
         Caption         =   "Signature of Applicant"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3450
         TabIndex        =   29
         Top             =   6885
         Width           =   1545
      End
      Begin VB.Label Label45 
         BackColor       =   &H80000009&
         Caption         =   "%, Renew Your License By"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   28
         Top             =   7125
         Width           =   1830
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label46 
         BackColor       =   &H80000009&
         Caption         =   "Return Application and Fee to:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   435
         TabIndex        =   27
         Top             =   7365
         Width           =   2220
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTownOf 
         BackColor       =   &H80000009&
         Caption         =   "Town Of XXXXXXXXXXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   435
         TabIndex        =   26
         Top             =   7560
         Width           =   3180
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label47 
         BackColor       =   &H80000009&
         Caption         =   ", "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2790
         TabIndex        =   25
         Top             =   8040
         Width           =   105
      End
      Begin VB.Label Label48 
         BackColor       =   &H80000009&
         Caption         =   "To Avoid Late Penalty Charge of "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   435
         TabIndex        =   24
         Top             =   7125
         Width           =   2220
         WordWrap        =   -1  'True
      End
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   690
      Left            =   9465
      TabIndex        =   72
      TabStop         =   0   'False
      Tag             =   "Press the 'Cancel' button to close this screen and return to the Town Setup screen."
      Top             =   6420
      Width           =   1890
      _Version        =   131072
      _ExtentX        =   3334
      _ExtentY        =   1217
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
      ButtonDesigner  =   "frmBLAppTemplate3.frx":34C8
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdNext 
      Height          =   690
      Left            =   9465
      TabIndex        =   73
      TabStop         =   0   'False
      Tag             =   "Press this 'Next App' button to close this application screen and open up the screen for application #4."
      Top             =   4530
      Width           =   1890
      _Version        =   131072
      _ExtentX        =   3334
      _ExtentY        =   1217
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
      ButtonDesigner  =   "frmBLAppTemplate3.frx":36A6
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   690
      Left            =   9465
      TabIndex        =   74
      TabStop         =   0   'False
      Tag             =   "Press 'Save' to save the currently active application as application #3. All fields will be committed to memory."
      Top             =   7365
      Width           =   1890
      _Version        =   131072
      _ExtentX        =   3334
      _ExtentY        =   1217
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
      ButtonDesigner  =   "frmBLAppTemplate3.frx":3885
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdLast 
      Height          =   675
      Left            =   9465
      TabIndex        =   75
      TabStop         =   0   'False
      Tag             =   "Press this 'Last App' to close this screen and open the screen for application #2."
      Top             =   5490
      Width           =   1890
      _Version        =   131072
      _ExtentX        =   3334
      _ExtentY        =   1191
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
      ButtonDesigner  =   "frmBLAppTemplate3.frx":3A61
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
      Height          =   492
      Left            =   9456
      TabIndex        =   82
      Tag             =   $"frmBLAppTemplate3.frx":3C40
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
      ButtonDesigner  =   "frmBLAppTemplate3.frx":3D0A
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   444
      Left            =   9984
      TabIndex        =   83
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
      TabIndex        =   84
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
      Caption         =   "Renewal Application #3"
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
      Height          =   732
      Left            =   9492
      TabIndex        =   59
      Top             =   1920
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
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmBLAppTemplate3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  Unload frmBLAppTemplate3
  frmBLTownSetup.fpcmbAppType.SetFocus
End Sub

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    lblBalloon.Visible = True
    frmBLMessageBox.Label1.Height = 300
    frmBLMessageBox.Label2.Top = 800
    frmBLMessageBox.Label2.Height = 1500
    frmBLMessageBox.Label2.Caption = "Some of the discretionary values appearing on this page are supplied from the Town Setup screen. If other application templates have been used then some of the values here may have carried over from them. PLEASE REVIEW ALL values to make sure they reflect the CURRENT situation."
    frmBLMessageBox.Label3.Top = 2700
    frmBLMessageBox.Label3.Caption = "Issuance Fee is furnished solely from the Issuance Fee value appearing on the Town Setup screen. If no Issuance Fee is assessed then there will be no reference to that fee on the final printed letter."
    frmBLMessageBox.Show vbModal
    cmdHelp.Text = "F1 &Turn Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    fptxtTownOf.ToolTipText = ""
    fpcmbYear1.ToolTipText = ""
    fptxtBaseFee1.ToolTipText = ""
    fptxtBaseFee2.ToolTipText = ""
    fptxtNumerator.ToolTipText = ""
    fptxtDenominator.ToolTipText = ""
    fptxtGrsPct.ToolTipText = ""
    fptxtGross1.ToolTipText = ""
    fptxtBaseFee3.ToolTipText = ""
    fptxtBaseFee4.ToolTipText = ""
    fpcmbStartMonth.ToolTipText = ""
    fpcmbStartDay.ToolTipText = ""
    fpcmbEndMonth.ToolTipText = ""
    fpcmbEndDay.ToolTipText = ""
    fpcmbYear2.ToolTipText = ""
    fptxtPct.ToolTipText = ""
    fpcmbPayByMonth.ToolTipText = ""
    fpcmbPayByDay.ToolTipText = ""
    fpcmbYear3.ToolTipText = ""
    fptxtAdd.ToolTipText = ""
    fptxtTownName.ToolTipText = ""
    fptxtState.ToolTipText = ""
    fptxtZip.ToolTipText = ""
    cmdHelp.ToolTipText = ""
    cmdNext.ToolTipText = ""
    cmdLast.ToolTipText = ""
    cmdExit.ToolTipText = ""
    cmdSave.ToolTipText = ""
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
'    cmdHelp.ToolTipText = "Press this button to activate/deactivate instructional balloons."
'    fptxtTownOf.ToolTipText = "Enter 'Town Of  Your Town' here."
'    fpcmbYear1.ToolTipText = "Select 'Curr' if you want the current year displayed here. Select  '+1' if you want the next year displayed here or select '-1' if you want the prior year displayed here."
'    fptxtBaseFee1.ToolTipText = "Enter the base fee for Contracting or Construction here."
'    fptxtBaseFee2.ToolTipText = "Enter the base fee for Retail Sales here."
'    fptxtNumerator.ToolTipText = "Enter the denominator of the fraction of percentage of Retail Sales gross receipts here."
'    fptxtDenominator.ToolTipText = "Enter the numerator of the fraction of percentage of Retail Sales gross receipts here."
'    fptxtGrsPct.ToolTipText = "Enter the percentage of Retail Sales gross receipts here."
'    fptxtGross1.ToolTipText = "Enter Retail Sales gross amount here."
'    fptxtBaseFee3.ToolTipText = "Enter the base fee for Financial, Real Estate or Professional Service here."
'    fptxtBaseFee4.ToolTipText = "Enter the base fee for Repair, Personal, Business or Delivery Service here."
'    fpcmbStartMonth.ToolTipText = "Select the first valid month for this year's business license."
'    fpcmbStartDay.ToolTipText = "Select the first valid day for this year's business license."
'    fpcmbEndMonth.ToolTipText = "Select the last valid month for this year's business license."
'    fpcmbEndDay.ToolTipText = "Select the last valid month for this year's business license."
'    fpcmbYear2.ToolTipText = "Select 'Curr' if you want the current year displayed here. Select  '+1' if you want the next year displayed here or select '-1' if you want the prior year displayed here."
'    fptxtPct.ToolTipText = "Enter the penalty amount here."
'    fpcmbPayByMonth.ToolTipText = "This is the last month to remit for this year's business license."
'    fpcmbPayByDay.ToolTipText = "This is the last day to remit for this year's business license."
'    fpcmbYear3.ToolTipText = "Select 'Curr' if you want the current year displayed here. Select  '+1' if you want the next year displayed here or select '-1' if you want the prior year displayed here."
'    fptxtAdd.ToolTipText = "Enter your town's street address here."
'    fptxtTownName.ToolTipText = "Enter your town's mailing name here."
'    fptxtState.ToolTipText = "Enter your town's state here (ex. VA) here."
'    fptxtZip.ToolTipText = "Enter your town's zip code here."
'    cmdNext.ToolTipText = "Press to move to application template #4."
'    cmdLast.ToolTipText = "Press to move to business application #2."
'    cmdExit.ToolTipText = "Press to return to the Town Setup screen."
'    cmdSave.ToolTipText = "Press to save the data on this screen."
  End If
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
    Case vbKeyF2:
      SendKeys "%L"
      Call cmdLast_Click
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
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLAppTemplate3.")
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
'  fptxtBaseFee1.ToolTipText = "Enter the base fee for Contracting or Construction here."
'  fptxtBaseFee2.ToolTipText = "Enter the base fee for Retail Sales here."
'  fptxtNumerator.ToolTipText = "Enter the denominator of the fraction of percentage of Retail Sales gross receipts here."
'  fptxtDenominator.ToolTipText = "Enter the numerator of the fraction of percentage of Retail Sales gross receipts here."
'  fptxtGrsPct.ToolTipText = "Enter the percentage of Retail Sales gross receipts here."
'  fptxtGross1.ToolTipText = "Enter Retail Sales gross amount here."
'  fptxtBaseFee3.ToolTipText = "Enter the base fee for Financial, Real Estate or Professional Service here."
'  fptxtBaseFee4.ToolTipText = "Enter the base fee for Repair, Personal, Business or Delivery Service here."
'  fpcmbStartMonth.ToolTipText = "Select the first valid month for this year's business license."
'  fpcmbStartDay.ToolTipText = "Select the first valid day for this year's business license."
'  fpcmbEndMonth.ToolTipText = "Select the last valid month for this year's business license."
'  fpcmbEndDay.ToolTipText = "Select the last valid month for this year's business license."
'  fpcmbYear2.ToolTipText = "Select 'Curr' if you want the current year displayed here. Select  '+1' if you want the next year displayed here or select '-1' if you want the prior year displayed here."
'  fptxtPct.ToolTipText = "Enter the penalty amount here."
'  fpcmbPayByMonth.ToolTipText = "This is the last month to remit for this year's business license."
'  fpcmbPayByDay.ToolTipText = "This is the last day to remit for this year's business license."
'  fpcmbYear3.ToolTipText = "Select 'Curr' if you want the current year displayed here. Select  '+1' if you want the next year displayed here or select '-1' if you want the prior year displayed here."
'  fptxtAdd.ToolTipText = "Enter your town's street address here."
'  fptxtTownName.ToolTipText = "Enter your town's mailing name here."
'  fptxtState.ToolTipText = "Enter your town's state here (ex. VA) here."
'  fptxtZip.ToolTipText = "Enter your town's zip code here."
'  cmdNext.ToolTipText = "Press to move to application template #4."
'  cmdLast.ToolTipText = "Press to move to business application #2."
'  cmdExit.ToolTipText = "Press to return to the Town Setup screen."
'  cmdSave.ToolTipText = "Press to save the data on this screen."
  
  If QPTrim$(frmBLTownSetup.fpcmbAmtPct.Text) = "Amt" Then
    fptxtPct.ToolTipText = "Enter the penalty amount here."
    Label48.Caption = "To Avoid Late Penalty Charge of $"
    Label45.Caption = ", Renew Your License By"
  Else
    fptxtPct.ToolTipText = "Enter the penalty percentage here."
    Label48.Caption = "To Avoid Late Penalty Charge of "
    Label45.Caption = "%, Renew Your License By"
  End If
  
  If Exist("artownsu.dat") Then
    OpenTownFile THandle
    Get THandle, 1, TownRec
    Close THandle
    If QPTrim(TownRec.AppAdd1) = "" Then
      fptxtAdd.Text = QPTrim$(frmBLTownSetup.fptxtAdd1.Text)
    Else
      fptxtAdd.Text = QPTrim(TownRec.AppAdd1)
    End If
    fptxtBaseFee1.Text = TownRec.AppBaseFee(1)
    fptxtBaseFee2.Text = TownRec.AppBaseFee(2)
    fptxtBaseFee3.Text = TownRec.AppBaseFee(3)
    fptxtBaseFee4.Text = TownRec.AppBaseFee(4)
    
    If QPTrim$(TownRec.AppStartMonth) = "" Then
      fpcmbStartMonth.Text = "JAN"
    Else
      fpcmbStartMonth.Text = QPTrim$(TownRec.AppStartMonth)
    End If
    
    If TownRec.AppStartDay = 0 Then
      fpcmbStartDay.Text = "1"
    Else
      fpcmbStartDay.Text = TownRec.AppStartDay
    End If
    
    If QPTrim$(TownRec.AppLicRetMonth) = "" Then
      fpcmbEndMonth.Text = "DEC"
    Else
      fpcmbEndMonth.Text = QPTrim$(TownRec.AppLicRetMonth)
    End If
    
    If TownRec.AppLicRetDay = 0 Then
      fpcmbEndDay.Text = "31"
    Else
      fpcmbEndDay.Text = TownRec.AppLicRetDay
    End If
    
    fptxtGross1.Text = TownRec.AppGrsRcpts(1)
    fptxtPct = TownRec.AppPct
    
    If QPTrim$(TownRec.AppState) = "" Then
      fptxtState.Text = QPTrim$(frmBLTownSetup.fptxtState.Text)
    Else
      fptxtState.Text = QPTrim$(TownRec.AppState)
    End If
    
    If QPTrim$(TownRec.AppCity) = "" Then
      fptxtTownName.Text = QPTrim$(frmBLTownSetup.fptxtCity.Text)
    Else
      fptxtTownName.Text = QPTrim$(TownRec.AppCity)
    End If
    
    If QPTrim$(TownRec.AppTownOf) = "" Then
      fptxtTownOf.Text = QPTrim$(frmBLTownSetup.fptxtTownName.Text)
    Else
      fptxtTownOf.Text = QPTrim$(TownRec.AppTownOf)
    End If
    
    If QPTrim$(TownRec.AppZip) = "" Then
      fptxtZip.Text = QPTrim$(frmBLTownSetup.fptxtZip.Text)
    Else
      fptxtZip.Text = QPTrim$(TownRec.AppZip)
    End If
    
    fptxtGrsPct.Text = TownRec.AppGrsPct
    
    fplblIssFee1.Caption = QPTrim$(Using$("$#,##0.00", TownRec.IssFee)) 'frmBLTownSetup.fpcurrIssFee
    fplblIssFee2.Caption = QPTrim$(Using$("$#,##0.00", TownRec.IssFee)) 'frmBLTownSetup.fpcurrIssFee
    fplblIssFee3.Caption = QPTrim$(Using$("$#,##0.00", TownRec.IssFee)) 'frmBLTownSetup.fpcurrIssFee
    fplblIssfee4.Caption = QPTrim$(Using$("$#,##0.00", TownRec.IssFee)) 'frmBLTownSetup.fpcurrIssFee
    
    fptxtDenominator.Text = TownRec.AppDenom
    fptxtNumerator.Text = TownRec.AppNumer
    
    If QPTrim$(TownRec.AppPenMonth) = "" Then
      fpcmbPayByMonth.Text = "JAN"
    Else
      fpcmbPayByMonth.Text = QPTrim(TownRec.AppPenMonth)
    End If
    If TownRec.AppPenDay = 0 Then
      fpcmbPayByDay.Text = "1"
    Else
      fpcmbPayByDay.Text = TownRec.AppPenDay
    End If
    
    For x = 1 To 3
      If QPTrim$(TownRec.AppYrUpDown(x)) = "0" Then TownRec.AppYrUpDown(x) = "Curr"
    Next x
    
    fpcmbYear1.Text = TownRec.AppYrUpDown(1)
    fpcmbYear2.Text = TownRec.AppYrUpDown(2)
    fpcmbYear3.Text = TownRec.AppYrUpDown(3)
  Else 'no town setup saved
    If QPTrim$(frmBLTownSetup.fptxtAdd1.Text) <> "" Then
      fptxtAdd.Text = QPTrim(frmBLTownSetup.fptxtAdd1.Text)
    Else
      fptxtAdd.Text = "Street address"
    End If
    
    fptxtBaseFee1.Text = "0.00"
    fptxtBaseFee2.Text = "0.00"
    fptxtBaseFee3.Text = "0.00"
    fptxtBaseFee4.Text = "0.00"
    fpcmbStartMonth.Text = "JAN"
    fpcmbStartDay.Text = "1"
    fpcmbEndMonth.Text = "DEC"
    fpcmbEndDay.Text = "31"
    fptxtGross1.Text = "0.00"
    fplblIssFee1.Caption = frmBLTownSetup.fpcurrIssFee '"0.00"
    fplblIssFee2.Caption = frmBLTownSetup.fpcurrIssFee '"0.00"
    fplblIssFee3.Caption = frmBLTownSetup.fpcurrIssFee '"0.00"
    fplblIssfee4.Caption = frmBLTownSetup.fpcurrIssFee '"0.00"
    
    If QPTrim$(frmBLTownSetup.fptxtState.Text) <> "" Then
      fptxtState.Text = QPTrim$(frmBLTownSetup.fptxtState.Text)
    Else
      fptxtState.Text = "NC"
    End If
    
    If QPTrim$(frmBLTownSetup.fptxtCity.Text) <> "" Then
      fptxtTownName.Text = QPTrim$(frmBLTownSetup.fptxtCity.Text)
    Else
      fptxtTownName.Text = "Town mailing name"
    End If
    
    If QPTrim$(frmBLTownSetup.fptxtTownName) <> "" Then
      fptxtTownOf.Text = QPTrim$(frmBLTownSetup.fptxtTownName.Text)
    Else
      fptxtTownOf.Text = "Town Of 'Your Town'"
    End If
    
    If QPTrim$(frmBLTownSetup.fptxtZip.Text) <> "" Then
      fptxtZip.Text = QPTrim$(frmBLTownSetup.fptxtZip.Text)
    Else
      fptxtZip.Text = "11111-1111"
    End If
    fpcmbPayByMonth.Text = "JAN"
    fpcmbPayByDay.Text = "1"
    fpcmbYear1.Text = "Curr"
    fpcmbYear2.Text = "Curr"
    fpcmbYear3.Text = "Curr"
  End If
  
  lblTownOf.Caption = QPTrim$(fptxtTownOf.Text)
  For x = 1 To 12
    Select Case x
      Case 1
        fpcmbStartMonth.AddItem "JAN"
        fpcmbEndMonth.AddItem "JAN"
        fpcmbPayByMonth.AddItem "JAN"
      Case 2
        fpcmbStartMonth.AddItem "FEB"
        fpcmbEndMonth.AddItem "FEB"
        fpcmbPayByMonth.AddItem "FEB"
      Case 3
        fpcmbStartMonth.AddItem "MAR"
        fpcmbEndMonth.AddItem "MAR"
        fpcmbPayByMonth.AddItem "MAR"
      Case 4
        fpcmbStartMonth.AddItem "APR"
        fpcmbEndMonth.AddItem "APR"
        fpcmbPayByMonth.AddItem "APR"
      Case 5
        fpcmbStartMonth.AddItem "MAY"
        fpcmbEndMonth.AddItem "MAY"
        fpcmbPayByMonth.AddItem "MAY"
      Case 6
        fpcmbStartMonth.AddItem "JUN"
        fpcmbEndMonth.AddItem "JUN"
        fpcmbPayByMonth.AddItem "JUN"
      Case 7
        fpcmbStartMonth.AddItem "JUL"
        fpcmbEndMonth.AddItem "JUL"
        fpcmbPayByMonth.AddItem "JUL"
      Case 8
        fpcmbStartMonth.AddItem "AUG"
        fpcmbEndMonth.AddItem "AUG"
        fpcmbPayByMonth.AddItem "AUG"
      Case 9
        fpcmbStartMonth.AddItem "SEP"
        fpcmbEndMonth.AddItem "SEP"
        fpcmbPayByMonth.AddItem "SEP"
      Case 10
        fpcmbStartMonth.AddItem "OCT"
        fpcmbEndMonth.AddItem "OCT"
        fpcmbPayByMonth.AddItem "OCT"
      Case 11
        fpcmbStartMonth.AddItem "NOV"
        fpcmbEndMonth.AddItem "NOV"
        fpcmbPayByMonth.AddItem "NOV"
      Case 12
        fpcmbStartMonth.AddItem "DEC"
        fpcmbEndMonth.AddItem "DEC"
        fpcmbPayByMonth.AddItem "DEC"
    End Select
  Next x
      
  For x = 1 To 31
    fpcmbStartDay.AddItem CStr(x)
    fpcmbEndDay.AddItem CStr(x)
    fpcmbPayByDay.AddItem CStr(x)
  Next x
  fpcmbYear1.AddItem "Curr"
  fpcmbYear1.AddItem "+1"
  fpcmbYear1.AddItem "-1"
  fpcmbYear2.AddItem "Curr"
  fpcmbYear2.AddItem "+1"
  fpcmbYear2.AddItem "-1"
  fpcmbYear3.AddItem "Curr"
  fpcmbYear3.AddItem "+1"
  fpcmbYear3.AddItem "-1"
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLAppTemplate3", "LoadMe", Erl)
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
  
  If QPTrim$(fptxtDenominator.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter the denominator for the fraction of percentage of gross receipts for Retail Sales."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtDenominator.BackColor = &H80FFFF
    fptxtDenominator.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtNumerator.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter the numerator for the fraction of percentage of gross receipts for Retail Sales."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtNumerator.BackColor = &H80FFFF
    fptxtNumerator.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtGrsPct.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter the percentage of gross receipts for Retail Sales."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtPct.BackColor = &H80FFFF
    fptxtPct.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fpcmbStartMonth.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter the first valid month of the new business license."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fpcmbStartMonth.BackColor = &H80FFFF
    fpcmbStartMonth.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fpcmbStartDay.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter the first valid day of the new business license."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fpcmbStartDay.BackColor = &H80FFFF
    fpcmbStartDay.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fpcmbEndMonth.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter the last valid month of the new business license."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fpcmbEndMonth.BackColor = &H80FFFF
    fpcmbEndMonth.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fpcmbEndDay.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter the last valid day of the new business license."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fpcmbEndDay.BackColor = &H80FFFF
    fpcmbEndDay.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fpcmbPayByMonth.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter the last month the new business license fee is to be paid."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fpcmbPayByMonth.BackColor = &H80FFFF
    fpcmbPayByMonth.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fpcmbPayByDay.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter the last day the new business license fee is to be paid."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fpcmbPayByDay.BackColor = &H80FFFF
    fpcmbPayByDay.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtPct.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter the penalty percentage."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtPct.BackColor = &H80FFFF
    fptxtPct.SetFocus
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
  
  If QPTrim$(fptxtTownName.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter the town's mailing name."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fptxtTownName.BackColor = &H80FFFF
    fptxtTownName.SetFocus
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
  
  If QPTrim$(fpcmbYear1.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter the year desired."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fpcmbYear1.BackColor = &H80FFFF
    fpcmbYear1.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fpcmbYear2.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter the year desired."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fpcmbYear2.BackColor = &H80FFFF
    fpcmbYear2.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fpcmbYear3.Text) = "" Then
    frmBLMessageBoxJr.Label1.Caption = "Please enter the year desired."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fpcmbYear3.BackColor = &H80FFFF
    fpcmbYear3.SetFocus
    Exit Sub
  End If
  
  If fptxtBaseFee1 = 0 Then
    If fptxtBaseFee2 = 0 Then
      If fptxtBaseFee3 = 0 Then
        If fptxtBaseFee4 = 0 Then
          fptxtBaseFee1.BackColor = &H80FFFF
          fptxtBaseFee2.BackColor = &H80FFFF
          fptxtBaseFee3.BackColor = &H80FFFF
          fptxtBaseFee4.BackColor = &H80FFFF
          frmBLMessageBoxJr.Label1.Caption = "Please enter an amount for at least one of the base fees."
          frmBLMessageBoxJr.Label1.Top = 800
          frmBLMessageBoxJr.Show vbModal
          fptxtBaseFee1.SetFocus
          Exit Sub
        End If
      End If
    End If
  End If
  
  If Exist("artownsu.dat") Then
    OpenTownFile THandle
    Get THandle, 1, TownRec
      TownRec.AppAdd1 = QPTrim(fptxtAdd.Text)
      TownRec.AppBaseFee(1) = fptxtBaseFee1.Text
      TownRec.AppBaseFee(2) = fptxtBaseFee2.Text
      TownRec.AppBaseFee(3) = fptxtBaseFee3.Text
      TownRec.AppBaseFee(4) = fptxtBaseFee4.Text
      TownRec.AppStartMonth = QPTrim(fpcmbStartMonth.Text)
      TownRec.AppStartDay = fpcmbStartDay.Text
      TownRec.AppLicRetMonth = QPTrim$(fpcmbEndMonth.Text)
      TownRec.AppLicRetDay = fpcmbEndDay.Text
      TownRec.AppGrsRcpts(1) = fptxtGross1.Text
      TownRec.AppPct = fptxtPct.Text
      TownRec.AppState = QPTrim$(fptxtState.Text)
      TownRec.AppCity = QPTrim$(fptxtTownName.Text)
      TownRec.AppTownOf = QPTrim$(fptxtTownOf.Text)
      TownRec.AppZip = QPTrim$(fptxtZip.Text)
      TownRec.AppGrsPct = fptxtGrsPct.Text
      TownRec.AppDenom = fptxtDenominator.Text
      TownRec.AppNumer = fptxtNumerator.Text
      TownRec.AppPenMonth = QPTrim$(fpcmbPayByMonth.Text)
      TownRec.AppPenDay = fpcmbPayByDay.Text
      TownRec.AppYrUpDown(1) = fpcmbYear1.Text
      TownRec.AppYrUpDown(2) = fpcmbYear2.Text
      TownRec.AppYrUpDown(3) = fpcmbYear3.Text
      TownRec.AppForm = 3
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
    TownRec.AppForm = 3
    TownRec.DLQNotice = 0
    TownRec.AppAdd1 = QPTrim(fptxtAdd.Text)
    TownRec.AppBaseFee(1) = fptxtBaseFee1.Text
    TownRec.AppBaseFee(2) = fptxtBaseFee2.Text
    TownRec.AppBaseFee(3) = fptxtBaseFee3.Text
    TownRec.AppBaseFee(4) = fptxtBaseFee4.Text
    TownRec.AppCentsPer(1) = 0
    TownRec.AppCentsPer(2) = 0
    TownRec.AppCentsPer(3) = 0
    TownRec.AppCentsPer(4) = 0
    TownRec.AppFirstDay = ""
    TownRec.AppLastDay = ""
    TownRec.AppGrsRcpts(1) = fptxtGross1.Text
    TownRec.AppGrsRcpts(2) = 0
    TownRec.AppGrsRcpts(3) = 0
    TownRec.AppGrsRcpts(4) = 0
    TownRec.AppColFee = 0
    TownRec.AppGrsPct = fptxtGrsPct.Text
    TownRec.AppDenom = fptxtDenominator.Text
    TownRec.AppNumer = fptxtNumerator.Text
    TownRec.AppState = QPTrim$(fptxtState.Text)
    TownRec.AppCity = QPTrim$(fptxtTownName.Text)
    TownRec.AppTownOf = QPTrim$(fptxtTownOf.Text)
    TownRec.AppZip = QPTrim$(fptxtZip.Text)
    TownRec.AppPct = fptxtPct.Text
    TownRec.AppPayBy = 0
    TownRec.AppPhone = ""
    TownRec.AppAdminName = ""
    TownRec.AppAdminTitle = ""
    TownRec.AppDiscPct = 0
    TownRec.AppDiscMonth = ""
    TownRec.AppDiscDay = 0
    TownRec.AppPenMonth = QPTrim$(fpcmbPayByMonth.Text)
    TownRec.AppPenDay = fpcmbPayByDay.Text
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
    TownRec.AppStartMonth = QPTrim$(fpcmbStartMonth.Text)
    TownRec.AppStartDay = fpcmbStartDay.Text
    TownRec.AppLicRetMonth = QPTrim$(fpcmbEndMonth.Text)
    TownRec.AppLicRetDay = fpcmbEndDay.Text
    TownRec.AppAdoptDate = 0
    TownRec.AppCityOrd = ""
    TownRec.AppYrUpDown(1) = fpcmbYear1.Text
    TownRec.AppYrUpDown(2) = fpcmbYear2.Text
    TownRec.AppYrUpDown(3) = fpcmbYear3.Text
    For x = 4 To 10
      TownRec.AppYrUpDown(x) = "0"
    Next x
    TownRec.DlqAdd1 = ""
    TownRec.DlqAdminName = ""
    TownRec.DlqAdminTitle = ""
    TownRec.DlqCity = ""
    TownRec.DlqPhone = ""
    TownRec.DlqPhone2 = ""
    TownRec.DlqFax = ""
    TownRec.DlqState = ""
    TownRec.DlqTownName = ""
    TownRec.DlqZip = ""
    TownRec.DlqFirstDay = ""
    TownRec.DlqLastDay = ""
    TownRec.DlqFirstHour = ""
    TownRec.DlqLastHour = ""
    TownRec.DlqClerkName = ""
    TownRec.DlqMayorCouncil = ""
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
  'renewal form #3 then coming here to save different data and then
  'trying to run application renewal reprints which will use this
  'latest saved data while the originals have the old data...now the
  'user will have to print applications over
  If Exist("artmpcus.dat") Then
    OpenTempCustRec TempHandle
    TempCnt = LOF(TempHandle) / Len(TempCustRec)
    If TempCnt > 0 Then
      Get TempHandle, 1, TempCustRec
      Close TempHandle
      If TempCustRec.AppType = 3 Then
        KillFile "artmpcus.dat"
      End If
    Else
      Close TempHandle
    End If
  End If
  
  frmBLSucSave.Label1.Caption = "Your renewal application notice #3 data has been saved successfully."
  frmBLSucSave.Label1.Top = 700
  frmBLSucSave.Show vbModal
  Call cmdExit_Click
  frmBLTownSetup.fpcmbAppType.Text = "3. APP FORM B"
  frmBLTownSetup.fpcmdApps.Text = "F3 S&how App Type 3"
  
  MainLog ("Application #3 saved.")
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLAppTemplate3", "cmdSave_Click", Erl)
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

Private Sub fpcmbEndDay_KeyDown(KeyCode As Integer, Shift As Integer)
  fpcmbEndDay.BackColor = -2147483643
  If KeyCode = vbKeySpace Then
    fpcmbEndDay.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbEndDay.ListIndex = -1
  End If
  If fpcmbEndDay.ListDown <> True Then
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

Private Sub fpcmbEndMonth_KeyDown(KeyCode As Integer, Shift As Integer)
  fpcmbEndMonth.BackColor = -2147483643
  If KeyCode = vbKeySpace Then
    fpcmbEndMonth.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbEndMonth.ListIndex = -1
  End If
  If fpcmbEndMonth.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbEndDay.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbPayByDay_KeyDown(KeyCode As Integer, Shift As Integer)
  fpcmbPayByDay.BackColor = -2147483643
  If KeyCode = vbKeySpace Then
    fpcmbPayByDay.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPayByDay.ListIndex = -1
  End If
  If fpcmbPayByDay.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbYear3.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbPayByMonth_KeyDown(KeyCode As Integer, Shift As Integer)
  fpcmbPayByMonth.BackColor = -2147483643
  If KeyCode = vbKeySpace Then
    fpcmbPayByMonth.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPayByMonth.ListIndex = -1
  End If
  If fpcmbPayByMonth.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbPayByDay.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbStartDay_KeyDown(KeyCode As Integer, Shift As Integer)
  fpcmbStartDay.BackColor = -2147483643
  If KeyCode = vbKeySpace Then
    fpcmbStartDay.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbStartDay.ListIndex = -1
  End If
  If fpcmbStartDay.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbEndMonth.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbStartMonth_KeyDown(KeyCode As Integer, Shift As Integer)
  fpcmbStartMonth.BackColor = -2147483643
  If KeyCode = vbKeySpace Then
    fpcmbStartMonth.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbStartMonth.ListIndex = -1
  End If
  If fpcmbStartMonth.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbStartDay.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

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
      fptxtBaseFee1.SetFocus
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
      fptxtPct.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbYear3_KeyDown(KeyCode As Integer, Shift As Integer)
  fpcmbYear3.BackColor = -2147483643
  If KeyCode = vbKeySpace Then
    fpcmbYear3.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbYear3.ListIndex = -1
  End If
  If fpcmbYear3.ListDown <> True Then
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

Private Sub fptxtAdd_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtAdd.BackColor = -2147483643

End Sub

Private Sub fptxtBaseFee1_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtBaseFee1.BackColor = -2147483643
  fptxtBaseFee2.BackColor = -2147483643
  fptxtBaseFee3.BackColor = -2147483643
  fptxtBaseFee4.BackColor = -2147483643

End Sub

Private Sub fptxtBaseFee2_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtBaseFee1.BackColor = -2147483643
  fptxtBaseFee2.BackColor = -2147483643
  fptxtBaseFee3.BackColor = -2147483643
  fptxtBaseFee4.BackColor = -2147483643
End Sub

Private Sub fptxtBaseFee3_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtBaseFee1.BackColor = -2147483643
  fptxtBaseFee2.BackColor = -2147483643
  fptxtBaseFee3.BackColor = -2147483643
  fptxtBaseFee4.BackColor = -2147483643

End Sub

Private Sub fptxtBaseFee4_Change()
  fptxtBaseFee1.BackColor = -2147483643
  fptxtBaseFee2.BackColor = -2147483643
  fptxtBaseFee3.BackColor = -2147483643
  fptxtBaseFee4.BackColor = -2147483643

End Sub

Private Sub fptxtGross1_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtGross1.BackColor = -2147483643

End Sub

Private Sub fptxtDenominator_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtDenominator.BackColor = -2147483643

End Sub

Private Sub fptxtNumerator_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtNumerator.BackColor = -2147483643

End Sub

Private Sub fptxtGrsPct_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtGrsPct.BackColor = -2147483643

End Sub

Private Sub fptxtPct_Change()
  fptxtPct.BackColor = -2147483643

End Sub

Private Sub fptxtState_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtState.BackColor = -2147483643

End Sub

Private Sub fptxtTownName_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtTownName.BackColor = -2147483643

End Sub

Private Sub fptxtTownOf_Change()
  lblTownOf.Caption = fptxtTownOf.Text
End Sub

Private Sub fptxtTownOf_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtTownOf.BackColor = -2147483643

End Sub

Private Sub fptxtZip_KeyDown(KeyCode As Integer, Shift As Integer)
  fptxtZip.BackColor = -2147483643

End Sub

Private Sub cmdNext_Click()
  frmBLAppTemplate4.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdLast_Click()
  frmBLAppTemplate2.Show
  DoEvents
  Unload Me
End Sub

