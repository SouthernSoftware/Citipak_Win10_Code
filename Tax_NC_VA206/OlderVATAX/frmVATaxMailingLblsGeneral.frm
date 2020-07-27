VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmVATaxMailingLblsGeneral 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Mailing Labels"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxMailingLblsGeneral.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   7836
      Left            =   360
      TabIndex        =   6
      Top             =   468
      Width           =   10668
      _Version        =   196609
      _ExtentX        =   18817
      _ExtentY        =   13822
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmVATaxMailingLblsGeneral.frx":08CA
      Begin LpLib.fpCombo fpcmbPrintOrder 
         Height          =   405
         Left            =   3540
         TabIndex        =   5
         Top             =   4560
         Width           =   3015
         _Version        =   196608
         _ExtentX        =   5318
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
         ColDesigner     =   "frmVATaxMailingLblsGeneral.frx":08E6
      End
      Begin LpLib.fpCombo fpcmbRange 
         Height          =   405
         Left            =   3300
         TabIndex        =   1
         Top             =   2400
         Width           =   4095
         _Version        =   196608
         _ExtentX        =   7223
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
         ColDesigner     =   "frmVATaxMailingLblsGeneral.frx":0C15
      End
      Begin LpLib.fpCombo fpcmbLabel 
         Height          =   405
         Left            =   3900
         TabIndex        =   0
         Top             =   1365
         Width           =   4470
         _Version        =   196608
         _ExtentX        =   7885
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
         ColDesigner     =   "frmVATaxMailingLblsGeneral.frx":0F44
      End
      Begin LpLib.fpCombo fpcmbTownship 
         Height          =   405
         Left            =   3540
         TabIndex        =   4
         Top             =   4035
         Width           =   3015
         _Version        =   196608
         _ExtentX        =   5318
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
         ColDesigner     =   "frmVATaxMailingLblsGeneral.frx":1273
      End
      Begin LpLib.fpList fpList 
         Height          =   1740
         Left            =   480
         TabIndex        =   12
         Tag             =   $"frmVATaxMailingLblsGeneral.frx":15A2
         Top             =   5280
         Width           =   9735
         _Version        =   196608
         _ExtentX        =   17171
         _ExtentY        =   3069
         TextAlias       =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
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
         Columns         =   5
         Sorted          =   0
         LineWidth       =   1
         SelDrawFocusRect=   -1  'True
         ColumnSeparatorChar=   9
         ColumnSearch    =   1
         ColumnWidthScale=   2
         RowHeight       =   -1
         MultiSelect     =   1
         WrapList        =   0   'False
         WrapWidth       =   0
         SelMax          =   -1
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
         ColumnHeaderShow=   -1  'True
         ColumnHeaderHeight=   -1
         GrpsFrozen      =   0
         BorderGrayAreaColor=   -2147483637
         ExtendRow       =   0
         DataField       =   ""
         OLEDragMode     =   0
         OLEDropMode     =   0
         Redraw          =   -1  'True
         ResizeRowToFont =   0   'False
         TextTipMultiLine=   0
         ColDesigner     =   "frmVATaxMailingLblsGeneral.frx":171B
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00D0D0D0&
         Height          =   1092
         Left            =   6900
         TabIndex        =   19
         Top             =   3900
         Width           =   1812
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
            Left            =   120
            TabIndex        =   22
            Top             =   750
            Width           =   1572
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
            Left            =   120
            TabIndex        =   21
            Top             =   450
            Width           =   1572
         End
         Begin VB.OptionButton OptBoth 
            BackColor       =   &H00D0D0D0&
            Caption         =   "Use Both"
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
            Left            =   120
            TabIndex        =   20
            Top             =   150
            Width           =   1572
         End
      End
      Begin EditLib.fpLongInteger fpLIFrom 
         Height          =   372
         Left            =   3900
         TabIndex        =   2
         Top             =   3240
         Width           =   1092
         _Version        =   196608
         _ExtentX        =   1926
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
         Text            =   "0"
         MaxValue        =   "2147483647"
         MinValue        =   "0"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
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
      Begin fpBtnAtlLibCtl.fpBtn cmdAlign 
         Height          =   492
         Left            =   5652
         TabIndex        =   13
         TabStop         =   0   'False
         Tag             =   $"frmVATaxMailingLblsGeneral.frx":1AF2
         Top             =   7200
         Width           =   1548
         _Version        =   131072
         _ExtentX        =   2730
         _ExtentY        =   868
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
         ButtonDesigner  =   "frmVATaxMailingLblsGeneral.frx":1BD1
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   492
         Left            =   1560
         TabIndex        =   14
         TabStop         =   0   'False
         Tag             =   "Press 'Exit' to return to the main Customer Maintenance menu."
         Top             =   7200
         Width           =   1692
         _Version        =   131072
         _ExtentX        =   2984
         _ExtentY        =   868
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
         ButtonDesigner  =   "frmVATaxMailingLblsGeneral.frx":1DAD
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   492
         Left            =   7536
         TabIndex        =   15
         TabStop         =   0   'False
         Tag             =   $"frmVATaxMailingLblsGeneral.frx":1F8B
         Top             =   7200
         Width           =   1692
         _Version        =   131072
         _ExtentX        =   2984
         _ExtentY        =   868
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
         ButtonDesigner  =   "frmVATaxMailingLblsGeneral.frx":2026
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdClear 
         Height          =   492
         Left            =   3600
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   7200
         Width           =   1692
         _Version        =   131072
         _ExtentX        =   2984
         _ExtentY        =   868
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
         ButtonDesigner  =   "frmVATaxMailingLblsGeneral.frx":2205
      End
      Begin EditLib.fpLongInteger fpLITo 
         Height          =   372
         Left            =   6420
         TabIndex        =   3
         Top             =   3240
         Width           =   1092
         _Version        =   196608
         _ExtentX        =   1926
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
         Text            =   "0"
         MaxValue        =   "2147483647"
         MinValue        =   "0"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
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
      Begin VB.Line Line1 
         BorderColor     =   &H80000004&
         BorderWidth     =   2
         X1              =   2220
         X2              =   8460
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
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
         Left            =   5100
         TabIndex        =   18
         Tag             =   "`"
         Top             =   3360
         Width           =   1212
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "From:"
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
         Left            =   2580
         TabIndex        =   17
         Tag             =   "`"
         Top             =   3360
         Width           =   1212
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   1896
         Left            =   2220
         Top             =   1920
         Width           =   6228
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
         Height          =   300
         Left            =   2220
         TabIndex        =   11
         Tag             =   "`"
         Top             =   4128
         Width           =   1212
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
         Left            =   2100
         TabIndex        =   10
         Top             =   4680
         Width           =   1308
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   3948
         Left            =   1140
         Top             =   1152
         Width           =   8412
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label Type:"
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
         Left            =   2412
         TabIndex        =   9
         Top             =   1476
         Width           =   1356
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Mailing Labels"
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
         Left            =   2460
         TabIndex        =   8
         Top             =   336
         Width           =   5868
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   696
         Left            =   2256
         Top             =   192
         Width           =   6228
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Select A Print Option:"
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
         Left            =   4020
         TabIndex        =   7
         Tag             =   "`"
         Top             =   2040
         Width           =   2412
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   8100
      Left            =   240
      Top             =   324
      Width           =   10932
   End
End
Attribute VB_Name = "frmVATaxMailingLblsGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  'Private Temp_Class As Resize_Class
  Dim CustRecs() As Long
  Dim CustCnt As Long
  Dim ThisOpt$
  Dim ZipRecs() As Long
  Dim ZipCnt As Long
  
Private Sub SortZipCodes()
 Dim Nextx As Long
 Dim HoldZip As Long
 Dim CustRec As TaxCustType
 Dim CustHandle As Integer
 Dim NumOfCustRecs As Long
 Dim BigNum As Long
 Dim SmallNum As Long
 Dim ThisZip As Long
 Dim ThisStringZip As String
 Dim x As Long, y As Integer
 Dim LocalCustRec() As Long
 Dim LocalZipRec() As Long
 Dim LocalZipGroup() As Long
 Dim LocalCnt As Long
 Dim Thisx As Long
 Dim ThisCustCnt As Long
 
 BigNum = 0
 SmallNum = 0
 OpenTaxCustFile CustHandle, NumOfCustRecs
 For x = 1 To NumOfCustRecs
 'make a list of all unique zip codes
    Get CustHandle, x, CustRec
    If CustRec.Deleted <> 0 Then GoTo SkipIt
    ThisCustCnt = ThisCustCnt + 1
    ThisStringZip = ""
     If QPTrim$(ReplaceString(CustRec.Zip, "-", "")) = "" Then
       ThisZip = 999999999
     Else
       ThisStringZip = QPTrim$(ReplaceString(CustRec.Zip, "-", ""))
       ThisZip = CLng(ThisStringZip)
     End If
    If x = 1 Then
       LocalCnt = LocalCnt + 1
       ReDim LocalZipRec(1 To LocalCnt) As Long
       LocalZipRec(LocalCnt) = ThisZip
    Else
       For y = 1 To LocalCnt
         If ThisZip = LocalZipRec(y) Then
           Exit For
         End If
       Next y
       If y > LocalCnt Then
         LocalCnt = LocalCnt + 1
         ReDim Preserve LocalZipRec(1 To LocalCnt) As Long
         LocalZipRec(LocalCnt) = ThisZip
       End If
    End If

SkipIt:
 Next x

 ReDim LocalZipGroup(1 To ThisCustCnt) As Long
 ReDim LocalCustRec(1 To ThisCustCnt) As Long
 ZipCnt = ThisCustCnt
 'categorize each customer's zip code
 For x = 1 To NumOfCustRecs
   Get CustHandle, x, CustRec
    If CustRec.Deleted <> 0 Then GoTo SkipThis
    Nextx = Nextx + 1
   ThisStringZip = ""
     If QPTrim$(ReplaceString(CustRec.Zip, "-", "")) = "" Then
       ThisZip = 999999999
     Else
       ThisStringZip = QPTrim$(ReplaceString(CustRec.Zip, "-", ""))
       ThisZip = CLng(ThisStringZip)
     End If
     For y = 1 To LocalCnt
       If ThisZip = LocalZipRec(y) Then
         LocalZipGroup(Nextx) = ThisZip
         LocalCustRec(Nextx) = x
         Exit For
       End If
     Next y
SkipThis:
 Next x
 
 'locate the biggest zip code
  For x = 1 To LocalCnt
    If LocalZipRec(x) > BigNum Then
      BigNum = LocalZipRec(x)
    End If
  Next x
  
  Nextx = 1
  ThisZip = 0
  SmallNum = BigNum + 1
   
  'sort in ascending order
  Do
     For x = Nextx To LocalCnt
       ThisZip = LocalZipRec(x)
       If ThisZip < SmallNum Then
         SmallNum = ThisZip
         Thisx = x
       End If
     Next x
     HoldZip = LocalZipRec(Nextx)
     LocalZipRec(Nextx) = LocalZipRec(Thisx)
     LocalZipRec(Thisx) = HoldZip
     SmallNum = BigNum + 1
     Nextx = Nextx + 1
     If Nextx = LocalCnt Then Exit Do
   Loop
   
   Dim Nexty As Long
   Nextx = 0
   ReDim ZipRecs(1 To ZipCnt) As Long
   Nexty = 1
   'load final customer list
   For y = Nexty To LocalCnt
     For x = 1 To ZipCnt
       If LocalZipGroup(x) = LocalZipRec(Nexty) Then
         Nextx = Nextx + 1
         ZipRecs(Nextx) = LocalCustRec(x)
        End If
     Next x
     Nexty = Nexty + 1
   Next y
   Close CustHandle
End Sub

Private Sub cmdAlign_Click()
  Dim TaxRpt As Integer
  Dim LType As Integer
  Dim cnt As Integer
  Dim Align$
  Dim ReportFile$
  
  On Error GoTo ERRORSTUFF
  
  ReDim OSet(1 To 4) As Integer
  
  Align$ = String$(34, "X")
  OSet(1) = 1
  OSet(2) = 37
  OSet(3) = 74
  OSet(4) = 110
  
  If fpcmbLabel.Text = "2) 1 X 2 5/8 1 Wide Text" Then
    LType = 1
  ElseIf fpcmbLabel.Text = "3) 1 X 2 5/8 3 Wide Text" Then
    LType = 2
  ElseIf fpcmbLabel.Text = "4) 1 X 2 5/8 4 Wide Text" Then
    LType = 3
  Else
    Call TaxMsg(900, "Please make a Label Type selection.")
    fpcmbLabel.SetFocus
    Exit Sub
  End If
  
  TaxRpt = FreeFile
  Open "TAXLABEL.RPT" For Output As TaxRpt
  Select Case LType
  Case 1
    For cnt = 1 To 5
      Print #TaxRpt, Align$
    Next
    Print #TaxRpt,
  Case 2
    For cnt = 1 To 5
      Print #TaxRpt, Align$; Tab(OSet(2)); Align$; Tab(OSet(3)); Align$
    Next
    Print #TaxRpt,
  Case 3
    For cnt = 1 To 5
      Print #TaxRpt, Align$; Tab(OSet(2)); Align$; Tab(OSet(3)); Align$; Tab(OSet(4)); Align$
    Next
    Print #TaxRpt,
  End Select

  Close TaxRpt

  ViewPrint "TAXLABEL.RPT", "Mailing Labels Alignment", True
  
  MainLog ("Mailing labels 'Align' feature for advertising charges used.")
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxMailLblsGeneral", "cmdAlign_Click", Erl)
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

Private Sub cmdClear_Click()
  fpList.Action = ActionDeselectAll
End Sub

Private Sub cmdExit_Click()
  frmVATaxReportsMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdProcess_Click()
  If InStr(fpcmbLabel.Text, "Graphical") Then
    Call PrintGraphics
  ElseIf InStr(fpcmbLabel.Text, "Text") Then
    Call TaxMsg(900, "Pitch 10 is recommended for this report.")
    Call PrintText
  Else
    Exit Sub
  End If
End Sub

Private Sub PrintText()
  Dim LType As Integer
  Dim ReportFile$
  Dim RptHandle As Integer
  Dim x As Long
  Dim TaxCustRec As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim Zip$
  Dim LabelCnt As Integer
  Dim PCnt As Integer
  Dim AcctNumber&
  Dim cnt As Integer
  Dim Address As String * 23
  Dim AllFlag As Boolean
  Dim ValidCnt As Integer
  Dim SpreadCnt As Long
  Dim NumOfCustIdx As Long
  Dim RowCnt As Long
  Dim ThisRec As Long
  On Error GoTo ERRORSTUFF

  AllFlag = True
  If fpcmbRange.Text <> "ALL" Then
    AllFlag = False
    If fpList.ListCount = 0 Then
      Call TaxMsg(900, "Please make a selection from the list")
      Exit Sub
    End If
  End If

  If fpcmbLabel.Text = "2) 1 X 2 5/8 1 Wide Text" Then
    LType = 1
  ElseIf fpcmbLabel.Text = "3) 1 X 2 5/8 3 Wide Text" Then
    LType = 2
  ElseIf fpcmbLabel.Text = "4) 1 X 2 5/8 4 Wide Text" Then
    LType = 3
  Else
    Call TaxMsg(900, "Please make a Label Type selection.")
    fpcmbLabel.SetFocus
    Exit Sub
  End If

  RowCnt = fpList.ListCount
  ReDim SpreadIdx(1 To 1) As Long
  SpreadCnt = 0
  If AllFlag = False Then
    If fpLIFrom.Enabled = True And fpLITo.Enabled = True Then
      If fpLIFrom.Value > fpLITo.Value Then
        Call TaxMsg(800, "Please make sure the 'From' field is less than or equal to the 'To' field.")
        fpLIFrom.SetFocus
        Close
        Exit Sub
      End If
      For x = 0 To RowCnt - 1
        fpList.Row = x
        fpList.Col = 3
        ThisRec = CLng(fpList.ColList)
        If ThisRec >= fpLIFrom.Value And ThisRec <= fpLITo.Value Then
          SpreadCnt = SpreadCnt + 1
          ReDim Preserve SpreadIdx(1 To SpreadCnt) As Long
          SpreadIdx(SpreadCnt) = ThisRec
        End If
      Next x
    Else
      For x = 0 To RowCnt - 1
        fpList.Row = x
        If fpList.Selected = True Then
          fpList.ListIndex = x
          fpList.Col = 3
          SpreadCnt = SpreadCnt + 1
          ReDim Preserve SpreadIdx(1 To SpreadCnt) As Long
          SpreadIdx(SpreadCnt) = CInt(fpList.ColText)
        End If
      Next x
    End If
  Else
    ReDim SpreadIdx(1 To CustCnt) As Long
    SpreadCnt = CustCnt
    For x = 1 To CustCnt
      SpreadIdx(x) = CustRecs(x)
    Next x
  End If
  
  ReportFile$ = "TAXLABEL.PRN"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  ReDim ToPrint(1 To 5, 1 To 5) As String

  If SpreadCnt > 0 Then
    NumOfCustIdx = SpreadCnt
    If SpreadCnt < 11 Then
      If TaxMsgWOpts(800, "Printing so few labels wastes mailing label paper. Press F10 if you wish to continue anyway. Otherwise, press ESC to stop the printing process.", "F10 Continue", "ESC Abort") = "abort" Then
        Unload frmVATaxMsgWOpts
        fpcmbRange.SetFocus
        Exit Sub
      Else
        Unload frmVATaxMsgWOpts
      End If
    End If
  End If

  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  For cnt = 1 To NumOfCustIdx
    If SpreadCnt > 0 Then
      Get TCHandle, SpreadIdx(cnt), TaxCustRec
    End If

    If Mid(TaxCustRec.Zip, 7, 1) <> " " Then
      Zip$ = TaxCustRec.Zip
      Zip$ = QPTrim$(Zip$)
    Else
      Zip$ = Left$(TaxCustRec.Zip, 5)
      Zip$ = QPTrim$(Zip$)
    End If

    Select Case LType
    Case 1
      Print #RptHandle, "Cust # " + CStr(SpreadIdx(cnt))
      Print #RptHandle, Left$(QPTrim$(TaxCustRec.CustName), 23)
      Print #RptHandle, Left$(QPTrim$(TaxCustRec.Addr1), 23)
      If Len(QPTrim$(TaxCustRec.Addr2)) > 0 Then
        Print #RptHandle, Left$(QPTrim$(TaxCustRec.Addr2), 23)
        Print #RptHandle, Left$(QPTrim$(TaxCustRec.City), 18) + ", " + QPTrim$(TaxCustRec.State) + " " + QPTrim(Zip$)
      Else
        Print #RptHandle, Left$(QPTrim$(TaxCustRec.City), 18) + ", " + QPTrim$(TaxCustRec.State) + " " + QPTrim$(Zip$)
        Print #RptHandle,
      End If
      Print #RptHandle,
      ValidCnt = ValidCnt + 1
    Case 2
      LabelCnt = LabelCnt + 1 'this requires a line to be printed
      'in columns of 3 with each column containing data gathered
      'from different customers...it also must limit the size of some
      'variables to accommodate the limitations of a mailing label's
      'size
      ToPrint(1, LabelCnt) = "Cust # " + CStr(SpreadIdx(cnt))
      ToPrint(2, LabelCnt) = Left(QPTrim$(TaxCustRec.CustName), 23)
      ToPrint(3, LabelCnt) = Left(QPTrim$(TaxCustRec.Addr1), 23)

      If Len(QPTrim$(TaxCustRec.Addr2)) > 0 Then
          ToPrint(4, LabelCnt) = Left(QPTrim$(TaxCustRec.Addr2), 23)
          ToPrint(5, LabelCnt) = Left(QPTrim$(TaxCustRec.City), 18) + ", " + QPTrim$(TaxCustRec.State) + " " + QPTrim$(Zip$)
      Else
          ToPrint(4, LabelCnt) = Left(QPTrim$(TaxCustRec.City), 18) + ", " + QPTrim$(TaxCustRec.State) + " " + QPTrim$(Zip$)
      End If

      If LabelCnt = 3 Then
        For PCnt = 1 To 5
          Print #RptHandle, ToPrint(PCnt, 1); Tab(37); ToPrint(PCnt, 2); Tab(74); ToPrint(PCnt, 3)
          ToPrint(PCnt, 1) = ""
          ToPrint(PCnt, 2) = ""
          ToPrint(PCnt, 3) = ""
          ToPrint(PCnt, 4) = ""
          ToPrint(PCnt, 5) = ""
        Next
        Print #RptHandle,
        LabelCnt = 0
      End If
      ValidCnt = ValidCnt + 1

    Case 3
      LabelCnt = LabelCnt + 1
      ToPrint(1, LabelCnt) = "Cust # " + CStr(SpreadIdx(cnt))
      ToPrint(2, LabelCnt) = Left(QPTrim$(TaxCustRec.CustName), 23)
      ToPrint(3, LabelCnt) = Left(QPTrim$(TaxCustRec.Addr1), 23)

      If Len(QPTrim$(TaxCustRec.Addr2)) > 0 Then
        ToPrint(4, LabelCnt) = Left(QPTrim$(TaxCustRec.Addr2), 23)
        ToPrint(5, LabelCnt) = Left(QPTrim$(TaxCustRec.City), 18) + ", " + QPTrim$(TaxCustRec.State) + " " + QPTrim(Zip$)
      Else
        ToPrint(4, LabelCnt) = Left(QPTrim$(TaxCustRec.City), 18) + ", " + QPTrim$(TaxCustRec.State) + " " + QPTrim$(Zip$)
      End If

      If LabelCnt = 4 Then
        For PCnt = 1 To 5
          Print #RptHandle, ToPrint(PCnt, 1); Tab(37); ToPrint(PCnt, 2); Tab(74); ToPrint(PCnt, 3); Tab(110); ToPrint(PCnt, 4)
          ToPrint(PCnt, 1) = ""
          ToPrint(PCnt, 2) = ""
          ToPrint(PCnt, 3) = ""
          ToPrint(PCnt, 4) = ""
          ToPrint(PCnt, 5) = ""
        Next
        Print #RptHandle,
        LabelCnt = 0
      End If
      ValidCnt = ValidCnt + 1
    End Select

NextLabel:
  Next

  If LType = 2 Or LType = 3 Then 'this catches the last
  'line of a print job if the last line contains fewer than
  'the number required to trigger another print line
    If LabelCnt > 0 Then
      For PCnt = 1 To 5
        Print #RptHandle, ToPrint(PCnt, 1); Tab(37); ToPrint(PCnt, 2); Tab(74); ToPrint(PCnt, 3)
      Next
      Print #RptHandle,
    End If
  End If
  PCnt = 0

  Close

  If ValidCnt = 0 Then
    Call TaxMsg(900, "There are no customers on record with the parameters entered.")
    Exit Sub
  End If
  
  ViewPrint ReportFile$, "Mailing Labels for Advertising Charges", True

  KillFile ReportFile$

  MainLog ("General mailing labels for advertising charges processed in text format.")

  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxMailLblsGeneral", "PrintText", Erl)
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
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  'Set Temp_Class = New Resize_Class
  'Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Me.HelpContextID = hlpMailingLabels
  Call LoadMe
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    'Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
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
    Case vbKeyF3:
      SendKeys "%L"
      Call cmdClear_Click
      KeyCode = 0
    Case vbKeyF5:
      SendKeys "%A"
      Call cmdAlign_Click
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
      KillFile "C:\CPWork\mllbls.dat"
      ClearInUse PWcnt
      MainLog ("CitiTax.exe terminated via menu bar on frmVATaxMailLblsAdv.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim CustRec As Long
  Dim TSRec As TownshipType
  Dim TSCnt As Integer
  Dim TSHandle As Integer
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim PropType$
  
  On Error Resume Next
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
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
  
  ReDim CustRecs(1 To 1) As Long
  CustCnt = 0
  OpenTaxCustFile TCHandle, NumOfTCRecs
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    CustRec = x
    If TaxCust.Deleted <> 0 Then GoTo SkipIt
    If TaxCust.FirstPersRec > 0 Then
      PropType = "Personal"
    End If
    If TaxCust.FirstPropRec > 0 Then
      If PropType = "Personal" Then
        PropType = "Both"
      Else
        PropType = "Real"
      End If
    End If
    If PropType = "" Then PropType = "Unknown"
    CustCnt = CustCnt + 1
    ReDim Preserve CustRecs(1 To CustCnt) As Long
    CustRecs(CustCnt) = CustRec
    fpList.InsertRow = "  " & Using$("#####", TaxCust.Acct) & Chr$(9) & "  " & QPTrim$(TaxCust.CustName) & Chr$(9) & "  " & QPTrim(TaxCust.Addr1) & Chr$(9) & CStr(CustRec) & Chr$(9) & PropType

SkipIt:
  Next x
  
  fpList.ListIndex = 0
  
  Close
  
  fpcmbLabel.Text = "1) 1 X 2 5/8 3 Wide Graphical"
  fpcmbLabel.AddItem "1) 1 X 2 5/8 3 Wide Graphical"
  fpcmbLabel.AddItem "2) 1 X 2 5/8 1 Wide Text"
  fpcmbLabel.AddItem "3) 1 X 2 5/8 3 Wide Text"
  fpcmbLabel.AddItem "4) 1 X 2 5/8 4 Wide Text"
  
  fpcmbRange.Text = "ALL"
  fpcmbRange.AddItem "ALL"
  fpcmbRange.AddItem "SELECT INDIVIDUALS FROM LIST"
  fpcmbRange.AddItem "USE FROM/TO OPTION"
  
  fpcmbPrintOrder.Text = "Name Order"
  fpcmbPrintOrder.AddItem "Name Order"
  fpcmbPrintOrder.AddItem "Acct Number Order"
  fpcmbPrintOrder.AddItem "Search Name"
  fpcmbPrintOrder.AddItem "Zip Code Order"
  ThisOpt = QPTrim$(TaxMasterRec.OptSrchCust)
  If ThisOpt <> "" Then
    fpcmbPrintOrder.AddItem ThisOpt + " Order"
  End If
  OptBoth.Value = True
End Sub

Private Sub fpcmbLabel_Change()
  If QPTrim$(fpcmbLabel.Text) = "" Then
    fpcmbLabel.Text = "1) 1 X 2 5/8 3 Wide Graphical"
  End If
  
  If InStr(fpcmbLabel.Text, "Text") Then
    cmdAlign.Enabled = True
  Else
    cmdAlign.Enabled = False
  End If
End Sub

Private Sub fpcmbLabel_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbLabel.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbLabel.ListIndex = -1
  End If
  If fpcmbLabel.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbRange.SetFocus
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
  If fpcmbRange.Text = "USE FROM/TO OPTION" And fpcmbPrintOrder.Text <> "Acct Number Order" Then
    Call TaxMsg(800, "If the option 'USE FROM/TO OPTION' is in use then " + fpcmbPrintOrder.Text + " cannot be used.")
    fpcmbPrintOrder.Text = "Acct Number Order"
    fpcmbRange.SetFocus
    Exit Sub
  End If
  
  Call LoadList
End Sub

Private Sub fpcmbRange_Change()
  If fpcmbRange.Text = "ALL" Then
    fpList.Action = ActionDeselectAll
    fpList.Enabled = False
  Else
    fpList.Enabled = True
  End If
  
  If fpcmbRange.Text <> "USE FROM/TO OPTION" Then
    fpList.MultiSelect = MultiSelectSimple
    fpLIFrom.Enabled = False
    fpLITo.Enabled = False
  ElseIf fpcmbRange.Text = "USE FROM/TO OPTION" Then
    fpList.Action = ActionDeselectAll
    fpList.MultiSelect = MultiSelectNone
    fpcmbPrintOrder.Text = "Acct Number Order"
    fpLIFrom.Enabled = True
    fpLITo.Enabled = True
  End If
End Sub

Private Sub fpcmbRange_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbRange.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbRange.ListIndex = -1
  End If
  If fpcmbRange.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      If fpLIFrom.Enabled = True Then
        fpLIFrom.SetFocus
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

Private Sub PrintGraphics()
  Dim LType As Integer
  Dim ReportFile$
  Dim RptHandle As Integer
  Dim AdvTrans As InterestRecType
  Dim ATHandle As Integer
  Dim NumOfATRecs As Long
  Dim x As Long
  Dim TaxCustRec As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim Zip$
  Dim DidCnt As Integer
  Dim LabelCnt As Integer
  Dim PCnt As Integer
  Dim AcctNumber&
  Dim cnt As Integer
  Dim AllFlag As Boolean
  Dim dlm$
  Dim ValidCnt As Integer
  Dim SpreadCnt As Long
  Dim NumOfCustIdx As Long
  Dim RowCnt As Long
  Dim ThisRec As Long
  On Error GoTo ERRORSTUFF

  AllFlag = True
  If fpcmbRange.Text <> "ALL" Then
    AllFlag = False
    If fpList.ListCount = 0 Then
      Call TaxMsg(900, "Please make a selection from the list")
      Exit Sub
    End If
  End If
  
  dlm = "~"
  
  ReDim SpreadIdx(1 To 1) As Long
  SpreadCnt = 0
  RowCnt = fpList.ListCount
  If AllFlag = False Then
    If fpLIFrom.Enabled = True And fpLITo.Enabled = True Then
      If fpLIFrom.Value > fpLITo.Value Then
        Call TaxMsg(800, "Please make sure the 'From' field is less than or equal to the 'To' field.")
        fpLIFrom.SetFocus
        Close
        Exit Sub
      End If
      For x = 0 To RowCnt - 1
        fpList.Row = x
        fpList.Col = 3
        ThisRec = CLng(fpList.ColList)
        If ThisRec >= fpLIFrom.Value And ThisRec <= fpLITo.Value Then
          SpreadCnt = SpreadCnt + 1
          ReDim Preserve SpreadIdx(1 To SpreadCnt) As Long
          SpreadIdx(SpreadCnt) = ThisRec
        End If
      Next x
    Else
      For x = 0 To RowCnt - 1
        fpList.Row = x
        If fpList.Selected = True Then
          fpList.ListIndex = x
          fpList.Col = 3
          SpreadCnt = SpreadCnt + 1
          ReDim Preserve SpreadIdx(1 To SpreadCnt) As Long
          SpreadIdx(SpreadCnt) = CLng(fpList.ColText)
        End If
      Next x
    End If
  Else
    ReDim SpreadIdx(1 To CustCnt) As Long
    SpreadCnt = CustCnt
    For x = 1 To CustCnt
      SpreadIdx(x) = CustRecs(x)
    Next x
  End If
  
  ReportFile$ = "TAXRPTS\TAXLABEL.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle

  ReDim ToPrint(1 To 5, 1 To 5) As String

  If SpreadCnt > 0 Then
    NumOfCustIdx = SpreadCnt
    If SpreadCnt < 11 Then
      If TaxMsgWOpts(800, "Printing so few labels wastes mailing label paper. Press F10 if you wish to continue anyway. Otherwise, press ESC to stop the printing process.", "F10 Continue", "ESC Abort") = "abort" Then
        Unload frmVATaxMsgWOpts
        fpcmbRange.SetFocus
        Exit Sub
      Else
        Unload frmVATaxMsgWOpts
      End If
    End If
  End If

  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  For cnt = 1 To NumOfCustIdx
    If SpreadCnt > 0 Then
      Get TCHandle, SpreadIdx(cnt), TaxCustRec
    End If
    ValidCnt = ValidCnt + 1

    If Mid(TaxCustRec.Zip, 7, 1) <> " " Then
      Zip$ = TaxCustRec.Zip
      Zip$ = QPTrim$(Zip$)
    Else
      Zip$ = Left$(TaxCustRec.Zip, 5)
      Zip$ = QPTrim$(Zip$)
    End If
    LabelCnt = LabelCnt + 1 'this requires a line to be printed
    'in columns of 3 with each column containing data gathered
    'from different customers...it also must limit the size of some
    'variables to accommodate the limitations of a mailing label's
    'size
    ToPrint(1, LabelCnt) = QPTrim$(TaxCustRec.CustName)
    ToPrint(2, LabelCnt) = Left(QPTrim$(TaxCustRec.Addr1), 23)
    ToPrint(3, LabelCnt) = Left(QPTrim$(TaxCustRec.Addr2), 23)
    ToPrint(4, LabelCnt) = Left(QPTrim$(TaxCustRec.City), 18) + ", " + QPTrim$(TaxCustRec.State) + " " + QPTrim$(Zip$)

    If LabelCnt = 3 Then 'got a complete line
      For PCnt = 1 To 4
        '                       0                    1                      2
        Print #RptHandle, ToPrint(PCnt, 1); dlm; ToPrint(PCnt, 2); dlm; ToPrint(PCnt, 3); dlm; ' ToPrint(PCnt, 4); dlm; ToPrint(PCnt, 5); dlm;
        ToPrint(PCnt, 1) = ""
        ToPrint(PCnt, 2) = ""
        ToPrint(PCnt, 3) = ""
        ToPrint(PCnt, 4) = ""
        ToPrint(PCnt, 5) = ""
      Next
      Print #RptHandle, ToPrint(PCnt, 1); dlm; ToPrint(PCnt, 2); dlm; ToPrint(PCnt, 3) '; dlm; ToPrint(PCnt, 4); dlm; ToPrint(PCnt, 5)
      ToPrint(PCnt, 1) = ""
      ToPrint(PCnt, 2) = ""
      ToPrint(PCnt, 3) = ""
      ToPrint(PCnt, 4) = ""
      ToPrint(PCnt, 5) = ""
      LabelCnt = 0
    End If

NextLabel:
  Next

  'this catches the last
  'line of a print job if the last line contains fewer than
  'the number required to trigger another print line
  For PCnt = 1 To 4
    '                   0,3,6,9,12            1,4,7,10,13            2,5,8,11,14
    Print #RptHandle, ToPrint(PCnt, 1); dlm; ToPrint(PCnt, 2); dlm; ToPrint(PCnt, 3); dlm; ' ToPrint(PCnt, 4); dlm; ToPrint(PCnt, 5); dlm;
  Next
  Print #RptHandle, ToPrint(PCnt, 1); dlm; ToPrint(PCnt, 2); dlm; ToPrint(PCnt, 3) '; dlm; ToPrint(PCnt, 4); dlm; ToPrint(PCnt, 5)

  PCnt = 0

  Close
  If ValidCnt = 0 Then
    Call TaxMsg(900, "There are no customers on record with the parameters entered.")
    Exit Sub
  End If

  arVATaxMailLabels.Show

  MainLog ("General mailing labels processed in graphics format for advertising charges.")

  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxMailLblsGeneral", "PrintGraphics", Erl)
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

Private Sub fpcmbPrintOrder_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbPrintOrder.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPrintOrder.ListIndex = -1
  End If
  If fpcmbPrintOrder.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbLabel.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub LoadList()
  Dim x As Long
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim IdxFlag As Boolean
  Dim ZipFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim CustRec As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim ThisTownship$
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim OptFlag As Boolean
  Dim PropType$
  
  fpList.Clear
  CustCnt = 0
  ReDim CustRecs(1 To 1) As Long
  OptFlag = False
  
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
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "Zip Code Order" Then
    Call SortZipCodes
    ReDim IdxArray(1 To ZipCnt) As Long
    For x = 1 To ZipCnt
      IdxArray(x) = ZipRecs(x)
    Next x
    ZipFlag = True
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
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  ElseIf ZipFlag = True Then
    NumOfTCRecs = ZipCnt
  End If
  
  If fpcmbTownship.Text <> "ALL" Then
    ThisTownship = fpcmbTownship.Text
  Else
    ThisTownship = "All"
  End If
  
  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
      CustRec = IdxArray(x)
    ElseIf ZipFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
      CustRec = IdxArray(x)
    Else
      Get TCHandle, x, TaxCust
      CustRec = x
    End If
    If TaxCust.Deleted <> 0 Then GoTo SkipIt
    If ThisTownship = "No Townships Saved" Then GoTo NoTownshipsSaved
    If ThisTownship <> "All" Then
      If ThisTownship <> QPTrim$(UCase(TaxCust.TownShip)) Then GoTo SkipIt
    End If
NoTownshipsSaved:
    If TaxCust.FirstPersRec > 0 Then
      PropType = "Personal"
    End If
    If TaxCust.FirstPropRec > 0 Then
      If PropType = "Personal" Then
        PropType = "Both"
      Else
        PropType = "Real"
      End If
    End If
    If PropType = "" Then PropType = "Unknown"
    If OptReal.Value = True Then If PropType = "Personal" Then GoTo SkipIt
    If OptPers.Value = True Then If PropType = "Real" Then GoTo SkipIt
    CustCnt = CustCnt + 1
    ReDim Preserve CustRecs(1 To CustCnt) As Long
    CustRecs(CustCnt) = CustRec
    fpList.InsertRow = "  " & Using$("#####", TaxCust.Acct) & Chr$(9) & "  " & QPTrim$(TaxCust.CustName) & Chr$(9) & "  " & QPTrim(TaxCust.Addr1) & Chr$(9) & CStr(CustRec) & Chr$(9) & PropType
SkipIt:
  Next x
  
  fpList.ListIndex = 0
  
  Close

End Sub

Private Sub fpcmbTownship_Change()
  Call LoadList
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

Private Sub OptBoth_Click()
  fpLIFrom.Value = 0
  fpLITo.Value = 0
  fpList.Action = ActionDeselectAll
  Call LoadList
End Sub

Private Sub OptPers_Click()
  fpLIFrom.Value = 0
  fpLITo.Value = 0
  fpList.Action = ActionDeselectAll
  Call LoadList
End Sub

Private Sub OptReal_Click()
  fpLIFrom.Value = 0
  fpLITo.Value = 0
  fpList.Action = ActionDeselectAll
  Call LoadList
End Sub
