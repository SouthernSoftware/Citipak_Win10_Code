VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmVATaxPrintLateNotice 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Print Late Notices"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmVATaxPrintLateNotice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   7350
      Left            =   1920
      TabIndex        =   7
      Top             =   750
      Width           =   7785
      _Version        =   196609
      _ExtentX        =   13732
      _ExtentY        =   12965
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmVATaxPrintLateNotice.frx":08CA
      Begin LpLib.fpCombo fpcmbPrintOpt 
         Height          =   405
         Left            =   2925
         TabIndex        =   6
         Top             =   5160
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
         ColDesigner     =   "frmVATaxPrintLateNotice.frx":08E6
      End
      Begin LpLib.fpCombo fpcmbTaxYear 
         Height          =   405
         Left            =   4005
         TabIndex        =   0
         Top             =   2100
         Width           =   1740
         _Version        =   196608
         _ExtentX        =   3069
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
         ColDesigner     =   "frmVATaxPrintLateNotice.frx":0BDD
      End
      Begin LpLib.fpCombo fpcmbPrintOrder 
         Height          =   405
         Left            =   2925
         TabIndex        =   5
         Top             =   4650
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
         ColDesigner     =   "frmVATaxPrintLateNotice.frx":0ED4
      End
      Begin LpLib.fpCombo fpcmbBillType 
         Height          =   405
         Left            =   2925
         TabIndex        =   4
         Top             =   4155
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
         ColDesigner     =   "frmVATaxPrintLateNotice.frx":11CB
      End
      Begin EditLib.fpText fptxtBegAcct 
         Height          =   390
         Left            =   4365
         TabIndex        =   2
         Top             =   3135
         Width           =   1410
         _Version        =   196608
         _ExtentX        =   2487
         _ExtentY        =   688
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
         CharValidationText=   "1 2 3 4 5 6 7 8 9 0"
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
      Begin EditLib.fpText fptxtEndAcct 
         Height          =   390
         Left            =   4365
         TabIndex        =   3
         Top             =   3645
         Width           =   1410
         _Version        =   196608
         _ExtentX        =   2487
         _ExtentY        =   688
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
         CharValidationText=   "1 2 3 4 5 6 7 8 9 0"
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
      Begin EditLib.fpText fptxtCurrForm 
         Height          =   390
         Left            =   3480
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Late notices are selected on the System Setup screen."
         Top             =   1275
         Width           =   2850
         _Version        =   196608
         _ExtentX        =   5027
         _ExtentY        =   688
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
         ControlType     =   1
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   50
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
      Begin EditLib.fpDateTime fptxtLtrDate 
         Height          =   375
         Left            =   4005
         TabIndex        =   1
         Top             =   2640
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
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   630
         Left            =   2040
         TabIndex        =   18
         TabStop         =   0   'False
         Tag             =   "Press the 'Cancel' button to exit this screen and return to the main 'Business License Reports' menu."
         Top             =   6330
         Width           =   1740
         _Version        =   131072
         _ExtentX        =   3069
         _ExtentY        =   1111
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
         ButtonDesigner  =   "frmVATaxPrintLateNotice.frx":14C2
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   630
         Left            =   4275
         TabIndex        =   19
         TabStop         =   0   'False
         Tag             =   $"frmVATaxPrintLateNotice.frx":16A0
         Top             =   6330
         Width           =   1740
         _Version        =   131072
         _ExtentX        =   3069
         _ExtentY        =   1111
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
         ButtonDesigner  =   "frmVATaxPrintLateNotice.frx":174B
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Print Type:"
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
         TabIndex        =   17
         Top             =   5250
         Width           =   1305
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Letter Date:"
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
         Left            =   1800
         TabIndex        =   16
         Top             =   2700
         Width           =   2025
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bill Type:"
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
         TabIndex        =   15
         Top             =   4245
         Width           =   1305
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Form In Use:"
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
         TabIndex        =   14
         Top             =   1335
         Width           =   1665
      End
      Begin VB.Label Label8 
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2640
         TabIndex        =   12
         Top             =   2205
         Width           =   1185
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ending Acct #:"
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
         Left            =   2040
         TabIndex        =   11
         Top             =   3735
         Width           =   2145
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   1440
         Top             =   435
         Width           =   5265
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Print Late Notices"
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
         TabIndex        =   10
         Top             =   570
         Width           =   4935
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   4005
         Left            =   1005
         Top             =   1890
         Width           =   5970
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
         Left            =   1440
         TabIndex        =   9
         Top             =   4740
         Width           =   1305
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Beginning Acct #:"
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
         Left            =   2040
         TabIndex        =   8
         Top             =   3240
         Width           =   2145
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   7620
      Left            =   1800
      Top             =   615
      Width           =   8055
   End
End
Attribute VB_Name = "frmVATaxPrintLateNotice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim TownName$
  Dim TownAdd1$
  Dim TownAdd2$
  Dim TownCSZ$
  Dim LtrDate$
  Dim RGTaxYear As Integer
  Dim PGTaxYear As Integer
  Dim NegOpt As String * 1

Private Sub cmdExit_Click()
  frmVATaxLateNoticeMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdProcess_Click()
  Call GetPrintData
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
  Me.HelpContextID = hlpPrintLate
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmVATaxPrintLateNotice.")
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
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim x As Long
  Dim FirstNum$
  Dim LastNum$
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim y As Integer
  Dim BigYr As Integer
  Dim ThisBigYr As Integer
  Dim HoldYr As Integer
  Dim Nextx As Integer
  Dim Thisx As Integer
  Dim YrCnt As Integer
  Dim FakeDate$
  Dim IntDate As Integer
  Dim ThisMonth As Integer
  
  On Error GoTo ERRORSTUFF
  
  NegOpt = "N"
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
  fpcmbTaxYear.AddItem "ALL YEARS"
  
  For x = YrCnt To 1 Step -1
    fpcmbTaxYear.AddItem CStr(Years(x))
  Next x
  Unload frmVATaxLoadReport
 
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  Select Case TaxMasterRec.LateForm
    Case 0:
      fptxtCurrForm.Text = "None Saved"
    Case 1:
      fptxtCurrForm.Text = "SELF EDIT #1"
    Case Else
  End Select
  
  TownName$ = QPTrim$(TaxMasterRec.Name)
  TownAdd1$ = QPTrim$(TaxMasterRec.Add1)
  TownAdd2$ = QPTrim$(TaxMasterRec.Add2)
  TownCSZ$ = QPTrim$(TaxMasterRec.City) + ", " + QPTrim$(TaxMasterRec.TaxSt) + "  " + QPTrim$(TaxMasterRec.Zip)
  fptxtLtrDate = Date
  LtrDate = fptxtLtrDate.Text
  RGTaxYear = CInt(TaxMasterRec.RTaxYear)
  PGTaxYear = CInt(TaxMasterRec.PTaxYear)
'  fpcmbTaxYear.Text = CStr(RGTaxYear)
  fpcmbTaxYear.ListIndex = 0
  
  IntDate = Date2Num(Date) 'converting date this way eliminates system date setting issues
  FakeDate = MakeRegDate(IntDate)
'  fptxtAdvDate = Mid(FakeDate, 1, 3) + "01" + Mid(FakeDate, 6, 10)
  ThisMonth = CInt(Mid(FakeDate, 1, 2))
  
'  Select Case ThisMonth
'    Case 1, 3, 5, 7, 8, 10, 12
'      fptxtPayByDate = Mid(FakeDate, 1, 3) + "31" + Mid(FakeDate, 6, 10)
'    Case 2
'      fptxtPayByDate = Mid(FakeDate, 1, 3) + "28" + Mid(FakeDate, 6, 10)
'    Case 4, 6, 9, 11
'      fptxtPayByDate = Mid(FakeDate, 1, 3) + "30" + Mid(FakeDate, 6, 10)
'    Case Else
'      fptxtPayByDate = Mid(FakeDate, 1, 3) + "28" + Mid(FakeDate, 6, 10)
'  End Select
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  For x = 1 To NumOfTCRecs
    Get TCHandle, x, TaxCust
    If TaxCust.Deleted <> 0 Then GoTo NotThisOne
    FirstNum = CStr(x)
    Exit For
NotThisOne:
  Next x
  
  For x = NumOfTCRecs To 1 Step -1
    Get TCHandle, x, TaxCust
    If TaxCust.Deleted <> 0 Then GoTo NotThisOne2
    LastNum = CStr(x)
    Exit For
NotThisOne2:
  Next x
  
  Close
  fptxtBegAcct.Text = FirstNum
  fptxtEndAcct.Text = LastNum
  
  fpcmbPrintOrder.Text = "Name Order"
  fpcmbPrintOrder.AddItem "Name Order"
  fpcmbPrintOrder.AddItem "Acct Number Order"
  
  fpcmbBillType.Text = "REAL ONLY"
  fpcmbBillType.AddItem "REAL ONLY"
  fpcmbBillType.AddItem "PERSONAL ONLY"
  
  If TaxMasterRec.LateForm = 1 Then
    fpcmbPrintOpt.Text = "Graphical"
    fpcmbPrintOpt.AddItem "Graphical"
    fpcmbPrintOpt.AddItem "Text"
  Else
    fpcmbPrintOpt.Enabled = False
    fpcmbPrintOpt.Text = "No Option"
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPrintLateNotice", "LoadMe", Erl)
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

Private Sub fpcmbBillType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbBillType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbBillType.ListIndex = -1
  End If
  If fpcmbBillType.ListDown <> True Then
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

Private Sub fpcmbPrintOpt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbPrintOrder.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPrintOrder.ListIndex = -1
  End If
  If fpcmbPrintOrder.ListDown <> True Then
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

Private Sub fpcmbPrintOrder_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbPrintOrder.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPrintOrder.ListIndex = -1
  End If
  If fpcmbPrintOrder.ListDown <> True Then
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

Private Sub GetPrintData()
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim LLRec As LateListPrintType
  Dim LLHandle As Integer
  Dim NumOfLLRecs As Long
  Dim LLCnt As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim NextRec As Long
  Dim PrintIt As Boolean
  Dim Balance#
  Dim PrincBal As Double
  Dim IntBal As Double
  Dim AdvBal As Double
  Dim LateListBal As Double
  Dim Opt1Bal As Double
  Dim Opt2Bal As Double
  Dim Opt3Bal As Double
  Dim PenBal As Double
  Dim PersBal As Double
  Dim MTBal As Double
  Dim MCBal As Double
  Dim FEBal As Double
  Dim MHBal As Double
  Dim x As Long
  Dim ThisRealRec As Long
  Dim ThisPersRec As Long
  Dim BillType$
  Dim ThisCustRec As Long
  Dim CurrBal As Double
  Dim PrevBal As Double
  Dim TotBal As Double
  Dim ThisTaxYear As Integer
  Dim Notified As Boolean
  Dim ThisBal$
  On Error GoTo ERRORSTUFF
  
  If fpcmbTaxYear.Text = "ALL YEARS" Then
    ThisTaxYear = -1
  Else
    ThisTaxYear = CInt(fpcmbTaxYear.Text)
  End If
  BillType = Mid(fpcmbBillType.Text, 1, 1)
  If QPTrim$(fpcmbPrintOrder.Text) = "Name Order" Then
    OpenNameIdxFile IdxHandle, NumOfIdx
    ReDim IdxArray(1 To NumOfIdx) As Long
    For x = 1 To NumOfIdx
      Get IdxHandle, x, IdxRec
      IdxArray(x) = IdxRec.CustRec
    Next x
    Close IdxHandle
    IdxFlag = True
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "Search Name" Then
    OpenSrchNameIdxFile SHandle, NumOfIdx
    ReDim IdxArray(1 To NumOfIdx) As Long
    For x = 1 To NumOfIdx
      Get SHandle, x, SearchRec
      IdxArray(x) = SearchRec.CustRec
    Next x
    Close SHandle
    IdxFlag = True
  End If
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  OpenRealPropFile RHandle, NumOfRealRecs
  OpenPersPropFile PHandle, NumOfPersRecs
  
  If IdxFlag = True Then
    NumOfTCRecs = NumOfIdx
  End If
  
  If Exist("TXLLPRN.DAT") Then
    KillFile "TXLLPRN.DAT"
  End If
  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenLatePrnFile LLHandle, NumOfLLRecs
  LLCnt = 0
  frmVATaxShowPctComp.Label1 = "Gathering Tax Transaction Data"
  frmVATaxShowPctComp.Show , Me
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  EnableCloseButton Me.hwnd, False
  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
      ThisCustRec = IdxArray(x)
    Else
      Get TCHandle, x, TaxCust
      ThisCustRec = x
    End If
    If TaxCust.Deleted <> 0 Then GoTo SkipIt
    If TaxCust.LateNotice <> "Y" Then GoTo SkipIt
    If ThisCustRec < CLng(fptxtBegAcct.Text) Or ThisCustRec > CLng(fptxtEndAcct.Text) Then GoTo SkipIt
    PrintIt = False
    NextRec = TaxCust.LastTrans
    If NextRec = 0 Then GoTo SkipIt
    Balance# = 0
    If fpcmbBillType.Text <> "REAL ONLY" Then
      GoTo PersOnly
    End If
    TotBal = GetCustRealBalance(ThisCustRec, -1)
    If ThisTaxYear <> -1 Then
      CurrBal = GetCustBalanceForYear(ThisCustRec, ThisTaxYear, "R")
    Else
      CurrBal = GetCustBalanceForYear(ThisCustRec, RGTaxYear, "R")
    End If
    PrevBal = OldRound(TotBal - CurrBal)
    ThisBal = CStr(PrevBal)
    If InStr(ThisBal, "E") Then PrevBal = 0
    If Notified = True Then GoTo Flagged
    If PrevBal < 0 Then
      Notified = True
      frmVATaxShowPctComp.Hide
      If TaxMsgWOpts(600, "Some of the previous balances are negatives but the overall balances are accurate. If you wish to have these late notices exclude negative previous balances press F10. Otherwise, press ESC to continue with no modifications.", "F10 No Negatives", "ESC Print As Is") = "abort" Then
        NegOpt = "N"
        MainLog ("User warned that late notices will contain some negative previous balances. The user elected not to modify these negative balances.")
      Else
        MainLog ("User warned that late notices will contain some negative previous balances. The user elected to modify these negative balances.")
        NegOpt = "Y"
      End If
      frmVATaxShowPctComp.Show
    End If
Flagged:
    Do While NextRec > 0
      Get TTHandle, NextRec, TaxTrans
      If BillType <> QPTrim$(TaxTrans.BillType) Then GoTo NextLoop
      If TaxTrans.TranType = 1 Then 'And TaxTrans.TaxYear = ThisTaxYear And TaxTrans.BillType = "R" Then
        If ThisTaxYear = -1 And TaxTrans.TaxYear = RGTaxYear Then
          PrincBal = OldRound#(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
          PrincBal = OldRound#(PrincBal - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
          IntBal = OldRound(TaxTrans.Revenue.Interest - TaxTrans.Revenue.InterestPd)
          AdvBal = OldRound#(TaxTrans.Revenue.Collection - TaxTrans.Revenue.CollectionPd)
          LateListBal = OldRound#(TaxTrans.Revenue.LateList - TaxTrans.Revenue.LateListPd)
          PenBal = OldRound#(TaxTrans.Revenue.Penalty - TaxTrans.Revenue.PenaltyPd)
          Opt1Bal = OldRound#(TaxTrans.Revenue.RevOpt1 - TaxTrans.Revenue.RevOpt1Pd)
          Opt2Bal = OldRound#(TaxTrans.Revenue.RevOpt2 - TaxTrans.Revenue.RevOpt2Pd)
          Opt3Bal = OldRound#(TaxTrans.Revenue.RevOpt3 - TaxTrans.Revenue.RevOpt3Pd)
          Balance = OldRound#(PrincBal + IntBal + AdvBal + LateListBal + Opt1Bal + Opt2Bal + Opt3Bal + PenBal)
        ElseIf ThisTaxYear <> -1 And TaxTrans.TaxYear = ThisTaxYear Then
          PrincBal = OldRound#(TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3 + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
          PrincBal = OldRound#(PrincBal - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd + TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd))
          IntBal = OldRound(TaxTrans.Revenue.Interest - TaxTrans.Revenue.InterestPd)
          AdvBal = OldRound#(TaxTrans.Revenue.Collection - TaxTrans.Revenue.CollectionPd)
          LateListBal = OldRound#(TaxTrans.Revenue.LateList - TaxTrans.Revenue.LateListPd)
          PenBal = OldRound#(TaxTrans.Revenue.Penalty - TaxTrans.Revenue.PenaltyPd)
          Opt1Bal = OldRound#(TaxTrans.Revenue.RevOpt1 - TaxTrans.Revenue.RevOpt1Pd)
          Opt2Bal = OldRound#(TaxTrans.Revenue.RevOpt2 - TaxTrans.Revenue.RevOpt2Pd)
          Opt3Bal = OldRound#(TaxTrans.Revenue.RevOpt3 - TaxTrans.Revenue.RevOpt3Pd)
          Balance = OldRound#(PrincBal + IntBal + AdvBal + LateListBal + Opt1Bal + Opt2Bal + Opt3Bal + PenBal)
        End If
      Else
        GoTo NextLoop
      End If
      
      If ThisTaxYear = -1 And Balance = 0 And PrevBal > 0 Then GoTo RGoAnyway
      If Balance > 0 Then
RGoAnyway:
        LLCnt = LLCnt + 1
        LLRec.TownName = TownName
        LLRec.LateSeqNum = LLCnt
        LLRec.CustName = QPTrim$(TaxCust.CustName)
        LLRec.Addr1 = TaxCust.Addr1
        LLRec.Addr2 = TaxCust.Addr2
        LLRec.City = TaxCust.City
        LLRec.State = TaxCust.State
        LLRec.Zip = TaxCust.Zip
'        LLRec.AdvDate = Date2Num(fptxtAdvDate.Text)
        LLRec.AdvDate = Date2Num(Date)
'        LLRec.PayDate = Date2Num(fptxtPayByDate.Text)
        LLRec.PayDate = Date2Num(Date)
        LLRec.TaxYear = ThisTaxYear
        If ThisTaxYear = -1 Then LLRec.TaxYear = RGTaxYear
        LLRec.PrincBal = PrincBal
        LLRec.IntBal = IntBal
        LLRec.AdvBal = AdvBal
        LLRec.LateListBal = LateListBal
        LLRec.Opt1Bal = Opt1Bal
        LLRec.Opt2Bal = Opt2Bal
        LLRec.Opt3Bal = Opt3Bal
        LLRec.CustAcct = TaxCust.Acct
        LLRec.PenBal = PenBal
        LLRec.MTBal = 0
        LLRec.MCBal = 0
        LLRec.FEBal = 0
        LLRec.MHBal = 0
        LLRec.PersBal = 0
        If TaxTrans.CustPin = 0 Then 'Dos transaction
          LLRec.RealValue = 0
          LLRec.RealExemp = 0
          LLRec.PersValue = 0
          LLRec.PersExemp = 0
        ElseIf TaxTrans.CustPin > 0 Then 'Windows transaction if <> 0
          If QPTrim$(TaxTrans.RealPin) = "0" Then ' = "0" comes from conversion so if it
          'is an empty string then it is a Windows transaction for personal property
            LLRec.RealValue = 0
            LLRec.RealExemp = 0
          Else
            ThisRealRec = GetRealRec(TaxTrans.RealPin)
            If ThisRealRec > 0 Then
              Get RHandle, ThisRealRec, RealRec
              LLRec.RealValue = OldRound(RealRec.PROPVALU)
              LLRec.RealExemp = OldRound(RealRec.EXMPOTHR) '6/14/06  + RealRec.EXMPSENI)
            Else
              LLRec.RealValue = 0
              LLRec.RealExemp = 0
            End If
          End If
          If QPTrim$(TaxTrans.PersPin) = "0" Then
            LLRec.PersValue = 0
            LLRec.PersExemp = 0
          Else
            ThisPersRec = GetPersRec(TaxTrans.PersPin)
            If ThisPersRec > 0 Then
              Get PHandle, ThisPersRec, PersRec
              LLRec.PersValue = OldRound(PersRec.CVALUE + PersRec.MCValue + PersRec.MHValue + PersRec.MTValue + PersRec.PersVal)
              LLRec.PersExemp = 0 ' 6/14/06 no more pers exemptionsOldRound(PersRec.EXMPOTHR + PersRec.EXMPSENI)
            Else
              LLRec.PersValue = 0
              LLRec.PersExemp = 0
            End If
          End If
        End If
        LLRec.TotBal = TotBal
        LLRec.CurrBal = CurrBal
        LLRec.PrevBal = PrevBal
        LLRec.LtrDate = Date2Num(LtrDate)
        If fpcmbPrintOpt.Text = "Graphical" Then
          LLRec.LtrType = "G"
        Else
          LLRec.LtrType = "T"
        End If
        Put LLHandle, LLCnt, LLRec
      End If
      PrintIt = True
NextLoop:
      NextRec = TaxTrans.LastTrans
    Loop
    GoTo SkipIt
    
PersOnly:
    TotBal = GetCustPersBalance(ThisCustRec, -1)
    If ThisTaxYear <> -1 Then
      CurrBal = GetCustBalanceForYear(ThisCustRec, ThisTaxYear, "P")
    Else
      CurrBal = GetCustBalanceForYear(ThisCustRec, PGTaxYear, "P")
    End If
    PrevBal = OldRound(TotBal - CurrBal)
    ThisBal = CStr(PrevBal)
    If InStr(ThisBal, "E") Then PrevBal = 0
    Do While NextRec > 0
      Get TTHandle, NextRec, TaxTrans
      If BillType <> QPTrim$(TaxTrans.BillType) Then GoTo NextLoopP
      If TaxTrans.TranType = 1 Then 'And TaxTrans.TaxYear = ThisTaxYear And TaxTrans.BillType = "P" Then
        If ThisTaxYear = -1 And TaxTrans.TaxYear = PGTaxYear Then
          PersBal = OldRound#(TaxTrans.Revenue.Principle1 - TaxTrans.PPTRADisc + TaxTrans.PPTRARmvl - TaxTrans.Revenue.Principle1Pd)
          MTBal = OldRound#(TaxTrans.Revenue.Principle2 - TaxTrans.Revenue.Principle2Pd)
          MCBal = OldRound#(TaxTrans.Revenue.Principle3 - TaxTrans.Revenue.Principle3Pd)
          FEBal = OldRound#(TaxTrans.Revenue.Principle4 - TaxTrans.Revenue.Principle4Pd)
          MHBal = OldRound#(TaxTrans.Revenue.Principle5 - TaxTrans.Revenue.Principle5Pd)
          IntBal = OldRound(TaxTrans.Revenue.Interest - TaxTrans.Revenue.InterestPd)
          AdvBal = 0
          LateListBal = 0
          PenBal = OldRound#(TaxTrans.Revenue.Penalty - TaxTrans.Revenue.PenaltyPd)
          Opt1Bal = OldRound#(TaxTrans.Revenue.RevOpt1 - TaxTrans.Revenue.RevOpt1Pd)
          Opt2Bal = OldRound#(TaxTrans.Revenue.RevOpt2 - TaxTrans.Revenue.RevOpt2Pd)
          Opt3Bal = OldRound#(TaxTrans.Revenue.RevOpt3 - TaxTrans.Revenue.RevOpt3Pd)
          Balance = OldRound#(PersBal + IntBal + MTBal + MCBal + FEBal + MHBal + Opt1Bal + Opt2Bal + Opt3Bal + PenBal)
        ElseIf ThisTaxYear <> -1 And TaxTrans.TaxYear = ThisTaxYear Then
          PersBal = OldRound#(TaxTrans.Revenue.Principle1 - TaxTrans.PPTRADisc + TaxTrans.PPTRARmvl - TaxTrans.Revenue.Principle1Pd)
          MTBal = OldRound#(TaxTrans.Revenue.Principle2 - TaxTrans.Revenue.Principle2Pd)
          MCBal = OldRound#(TaxTrans.Revenue.Principle3 - TaxTrans.Revenue.Principle3Pd)
          FEBal = OldRound#(TaxTrans.Revenue.Principle4 - TaxTrans.Revenue.Principle4Pd)
          MHBal = OldRound#(TaxTrans.Revenue.Principle5 - TaxTrans.Revenue.Principle5Pd)
          IntBal = OldRound(TaxTrans.Revenue.Interest - TaxTrans.Revenue.InterestPd)
          AdvBal = 0
          LateListBal = 0
          PenBal = OldRound#(TaxTrans.Revenue.Penalty - TaxTrans.Revenue.PenaltyPd)
          Opt1Bal = OldRound#(TaxTrans.Revenue.RevOpt1 - TaxTrans.Revenue.RevOpt1Pd)
          Opt2Bal = OldRound#(TaxTrans.Revenue.RevOpt2 - TaxTrans.Revenue.RevOpt2Pd)
          Opt3Bal = OldRound#(TaxTrans.Revenue.RevOpt3 - TaxTrans.Revenue.RevOpt3Pd)
          Balance = OldRound#(PersBal + IntBal + MTBal + MCBal + FEBal + MHBal + Opt1Bal + Opt2Bal + Opt3Bal + PenBal)
        End If
      Else
        GoTo NextLoopP
      End If
      
      If ThisTaxYear = -1 And Balance = 0 And PrevBal > 0 Then GoTo PGoAnyway
      If Balance > 0 Then
PGoAnyway:
        LLCnt = LLCnt + 1
        LLRec.TownName = TownName
        LLRec.LateSeqNum = LLCnt
        LLRec.CustName = QPTrim$(TaxCust.CustName)
        LLRec.Addr1 = TaxCust.Addr1
        LLRec.Addr2 = TaxCust.Addr2
        LLRec.City = TaxCust.City
        LLRec.State = TaxCust.State
        LLRec.Zip = TaxCust.Zip
'        LLRec.AdvDate = Date2Num(fptxtAdvDate.Text)
        LLRec.AdvDate = Date2Num(Date)
'        LLRec.PayDate = Date2Num(fptxtPayByDate.Text)
        LLRec.PayDate = Date2Num(Date)
        LLRec.TaxYear = ThisTaxYear
        If ThisTaxYear = -1 Then LLRec.TaxYear = PGTaxYear
        LLRec.PrincBal = 0
        LLRec.IntBal = IntBal
        LLRec.AdvBal = 0
        LLRec.LateListBal = 0
        LLRec.Opt1Bal = Opt1Bal
        LLRec.Opt2Bal = Opt2Bal
        LLRec.Opt3Bal = Opt3Bal
        LLRec.PersBal = PersBal
        LLRec.MTBal = MTBal
        LLRec.MCBal = MCBal '
        LLRec.FEBal = FEBal
        LLRec.MHBal = MHBal
        LLRec.CustAcct = TaxCust.Acct
        LLRec.PenBal = PenBal
        If TaxTrans.CustPin = 0 Then 'Dos transaction
          LLRec.RealValue = 0
          LLRec.RealExemp = 0
          LLRec.PersValue = 0
          LLRec.PersExemp = 0
        ElseIf TaxTrans.CustPin > 0 Then 'Windows transaction if <> 0
          If QPTrim$(TaxTrans.PersPin) = "0" Then
            LLRec.PersValue = 0
            LLRec.PersExemp = 0
          Else
            ThisPersRec = GetPersRec(TaxTrans.PersPin)
            If ThisPersRec > 0 Then
              Get PHandle, ThisPersRec, PersRec
              LLRec.PersValue = OldRound(PersRec.CVALUE + PersRec.MCValue + PersRec.MHValue + PersRec.MTValue + PersRec.PersVal)
              LLRec.PersExemp = 0 '6/14/06 no more pers exemptions OldRound(PersRec.EXMPOTHR + PersRec.EXMPSENI)
            Else
              LLRec.PersValue = 0
              LLRec.PersExemp = 0
            End If
          End If
        End If
        LLRec.TotBal = TotBal
        LLRec.CurrBal = CurrBal
        LLRec.PrevBal = PrevBal
        LLRec.LtrDate = Date2Num(LtrDate)
        LLRec.NegYN = NegOpt
        If fpcmbPrintOpt.Text = "Graphical" Then
          LLRec.LtrType = "G"
        Else
          LLRec.LtrType = "T"
        End If
        Put LLHandle, LLCnt, LLRec
      End If
      PrintIt = True
NextLoopP:
      NextRec = TaxTrans.LastTrans
    Loop
SkipIt:
    frmVATaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmVATaxShowPctComp.Out = True Then
      Close
      frmVATaxShowPctComp.Out = False
      Unload frmVATaxShowPctComp
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Close
      Exit Sub
    End If
  Next x
 
  Unload frmVATaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  Close
  
  If fptxtCurrForm.Text = "SELF EDIT #1" Then
    If fpcmbPrintOpt.Text = "Graphical" Then
      Call PrintGraphicsSelfEdit1
    Else
      Call PrintTextSelfEdit1
    End If
  End If
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPrintLateNotice", "GetPrintData", Erl)
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

Private Function GetRealRec(PIN$) As Long
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim x As Long
  
  GetRealRec = 0
  PIN$ = QPTrim$(PIN$)
  OpenRealPropFile RHandle, NumOfRealRecs
  For x = 1 To NumOfRealRecs
    Get RHandle, x, RealRec
    If QPTrim$(RealRec.RealPin) = PIN$ Then
      GetRealRec = x
      Exit For
    End If
  Next x
  Close RHandle
End Function

Private Function GetPersRec(PIN$) As Long
  Dim PersRec As PersonalRecType
  Dim PHandle As Integer
  Dim NumOfPersRecs As Long
  Dim x As Long
  
  GetPersRec = 0
  PIN$ = QPTrim$(PIN$)
  OpenPersPropFile PHandle, NumOfPersRecs
  For x = 1 To NumOfPersRecs
    Get PHandle, x, PersRec
    If QPTrim$(PersRec.PropPin) = PIN$ Then
      GetPersRec = x
      Exit For
    End If
  Next x
  Close PHandle
  
End Function

Private Sub fpcmbTaxYear_Change()
'  TaxYear = CInt(fpcmbTaxYear.Text)
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
      fptxtLtrDate.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub PrintGraphicsSelfEdit1()
  Dim dlm$
  Dim RptFile$
  Dim RptHandle As Integer
  Dim LLRec As LateListPrintType
  Dim LLHandle As Integer
  Dim NumOfLLRecs As Long
  Dim x As Long, y As Integer
  Dim LLtrRec As TAXLateLetterType
  Dim YrCnt As Integer
  Dim ThisRec As Integer
  Dim ThatRec As Integer
  
  On Error GoTo ERRORSTUFF
  
  dlm$ = "~"
  OpenLateLtrFile LLHandle
  Get LLHandle, 1, LLtrRec
  Close LLHandle
  RptFile$ = "TAXRPTS\LATENOTICE.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  OpenLatePrnFile LLHandle, NumOfLLRecs
  If NumOfLLRecs = 0 Then
    Call TaxMsg(900, "There are no late notices necessary for the parameters entered.")
    Close
    Exit Sub
  End If
  YrCnt = 0
  
  ReDim Years(1 To 1) As Integer
  For x = 1 To NumOfLLRecs
    Get LLHandle, x, LLRec
    If x = 1 Then
      ThatRec = LLRec.CustAcct
      ThisRec = LLRec.CustAcct
      YrCnt = YrCnt + 1
      ReDim Preserve Years(1 To YrCnt) As Integer
      Years(YrCnt) = LLRec.TaxYear
    Else
      ThisRec = LLRec.CustAcct
      If ThatRec <> ThisRec Then
        YrCnt = 1
        ReDim Years(1 To 1) As Integer
        Years(YrCnt) = LLRec.TaxYear
        ThatRec = ThisRec
      Else
        For y = 1 To YrCnt
          If LLRec.TaxYear = Years(y) Then
            GoTo NextYear
          End If
        Next y
        If y > YrCnt Then
          YrCnt = YrCnt + 1
          ReDim Preserve Years(1 To YrCnt) As Integer
          Years(YrCnt) = LLRec.TaxYear
        End If
      End If
    End If
    '                          0                           1                       2
    Print #RptHandle, QPTrim$(LLRec.Addr1); dlm; QPTrim$(LLRec.Addr2); dlm; LLRec.AdvBal; dlm;
    '                            3                             4                            5
    Print #RptHandle, MakeRegDate(LLRec.AdvDate); dlm; QPTrim$(LLRec.City); dlm; QPTrim$(LLRec.CustName); dlm;
    '                       6                    7                      8                     9
    Print #RptHandle, LLRec.IntBal; dlm; LLRec.LateListBal; dlm; LLRec.LateSeqNum; dlm; LLRec.Opt1Bal; dlm;
    '                       10                  11                       12                      13
    Print #RptHandle, LLRec.Opt2Bal; dlm; LLRec.Opt3Bal; dlm; MakeRegDate(LLRec.PayDate); dlm; LLRec.PersExemp; dlm;
    '                       14                  15                     16                     17
    Print #RptHandle, LLRec.PersValue; dlm; LLRec.PrincBal; dlm; LLRec.RealExemp; dlm; LLRec.RealValue; dlm;
    '                       18                       19                           20                         21
    Print #RptHandle, QPTrim$(LLRec.State); dlm; LLRec.TaxYear; dlm; QPTrim$(LLRec.TownName); dlm; QPTrim$(LLRec.Zip); dlm;
    '                       22                   23                           24                25
    Print #RptHandle, QPTrim$(TownAdd1); dlm; QPTrim$(TownAdd2); dlm; QPTrim$(TownCSZ); dlm; LtrDate; dlm;
    '                       26                   27                28              29                30
    Print #RptHandle, LLtrRec.Head1; dlm; LLtrRec.Head2; dlm; LLtrRec.Head3; dlm; LLtrRec.Head4; dlm; LLtrRec.Head5; dlm;
    
    For y = 1 To 20
      '31 - 50
      Print #RptHandle, LLtrRec.Body(y); dlm;
    Next y
    If fpcmbBillType.Text = "REAL ONLY" Then
      '                     51                  52                   53                  54                55             56
      Print #RptHandle, LLRec.TotBal; dlm; LLRec.CurrBal; dlm; LLRec.PrevBal; dlm; LLRec.CustAcct; dlm; RGTaxYear; dlm; NegOpt
    ElseIf fpcmbBillType.Text = "PERSONAL ONLY" Then
      '                     51                  52                   53                  54                55             56
      Print #RptHandle, LLRec.TotBal; dlm; LLRec.CurrBal; dlm; LLRec.PrevBal; dlm; LLRec.CustAcct; dlm; PGTaxYear; dlm; NegOpt
    End If

NextYear:
  Next x
  
  Close
  
  arVATaxLateLetter.Show
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPrintLateNotice", "PrintGraphicsSelfEdit1", Erl)
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

Private Sub fptxtLtrDate_Change()
  LtrDate = fptxtLtrDate.Text
End Sub

Private Sub PrintTextSelfEdit1()
  Dim RptFile$
  Dim RptHandle As Integer
  Dim LLRec As LateListPrintType
  Dim LLHandle As Integer
  Dim NumOfLLRecs As Long
  Dim x As Long, y As Integer
  Dim LLtrRec As TAXLateLetterType
  Dim FF$
  Dim HdrLen As Integer
  Dim Start1 As Integer
  Dim Start2 As Integer
  Dim Start3 As Integer
  Dim Start4 As Integer
  Dim Start5 As Integer
  Dim YrCnt As Integer
  Dim ThisRec As Integer
  Dim ThatRec As Integer
  
  On Error GoTo ERRORSTUFF
  
  FF$ = Chr(12)
  OpenLateLtrFile LLHandle
  Get LLHandle, 1, LLtrRec
  Close LLHandle
  HdrLen = Len(QPTrim$(LLtrRec.Head1))
  HdrLen = HdrLen / 2
  Start1 = 40 - HdrLen
  HdrLen = Len(QPTrim$(LLtrRec.Head2))
  HdrLen = HdrLen / 2
  Start2 = 40 - HdrLen
  HdrLen = Len(QPTrim$(LLtrRec.Head3))
  HdrLen = HdrLen / 2
  Start3 = 40 - HdrLen
  HdrLen = Len(QPTrim$(LLtrRec.Head4))
  HdrLen = HdrLen / 2
  Start4 = 40 - HdrLen
  HdrLen = Len(QPTrim$(LLtrRec.Head5))
  HdrLen = HdrLen / 2
  Start5 = 40 - HdrLen
  
  RptFile$ = "TAXRPTS\LATENOTICE.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  OpenLatePrnFile LLHandle, NumOfLLRecs
  If NumOfLLRecs = 0 Then
    Call TaxMsg(900, "There are no late notices necessary for the parameters entered.")
    Close
    Exit Sub
  End If
  YrCnt = 0
  
  ReDim Years(1 To 1) As Integer
  For x = 1 To NumOfLLRecs
    Get LLHandle, x, LLRec
    If x = 1 Then
      ThatRec = LLRec.CustAcct
      ThisRec = LLRec.CustAcct
      YrCnt = YrCnt + 1
      ReDim Preserve Years(1 To YrCnt) As Integer
      Years(YrCnt) = LLRec.TaxYear
    Else
      ThisRec = LLRec.CustAcct
      If ThatRec <> ThisRec Then
        YrCnt = 1
        ReDim Years(1 To 1) As Integer
        Years(YrCnt) = LLRec.TaxYear
        ThatRec = ThisRec
      Else
        For y = 1 To YrCnt
          If LLRec.TaxYear = Years(y) Then
            GoTo NextYear
          End If
        Next y
        If y > YrCnt Then
          YrCnt = YrCnt + 1
          ReDim Preserve Years(1 To YrCnt) As Integer
          Years(YrCnt) = LLRec.TaxYear
        End If
      End If
    End If
    Print #RptHandle,
    Print #RptHandle,
    Print #RptHandle,
    Print #RptHandle,
    Print #RptHandle,
    Print #RptHandle,
    Print #RptHandle, Tab(Start1); QPTrim$(LLtrRec.Head1)
    Print #RptHandle, Tab(Start2); QPTrim$(LLtrRec.Head2)
    Print #RptHandle, Tab(Start3); QPTrim$(LLtrRec.Head3)
    Print #RptHandle, Tab(Start4); QPTrim$(LLtrRec.Head4)
    Print #RptHandle, Tab(Start5); QPTrim$(LLtrRec.Head5)
    Print #RptHandle,
    Print #RptHandle, MakeRegDate(LLRec.LtrDate)
    Print #RptHandle,
    Print #RptHandle, QPTrim$(LLRec.CustName)
    Print #RptHandle, QPTrim$(LLRec.Addr1)
    Print #RptHandle, QPTrim$(LLRec.Addr2)
    Print #RptHandle, QPTrim$(LLRec.City) + ", " + QPTrim$(LLRec.State) + "  " + QPTrim$(LLRec.Zip)
    Print #RptHandle,
    
    For y = 1 To 10
      Print #RptHandle, LLtrRec.Body(y)
    Next y
    
    Print #RptHandle,
    If NegOpt = "N" Then
      If fpcmbBillType.Text = "REAL ONLY" Then
        If LLRec.TaxYear = RGTaxYear Then
          Print #RptHandle, Tab(5); "Customer Account: "; Tab(23); Using$("#####0", LLRec.CustAcct); Tab(33); "Prev Tax Years: "; Tab(55); Using$("$###,###,##0.00", LLRec.PrevBal)
        Else
          Print #RptHandle, Tab(5); "Customer Account: "; Tab(23); Using$("#####0", LLRec.CustAcct); Tab(33); "Other Tax Years: "; Tab(55); Using$("$###,###,##0.00", LLRec.PrevBal)
        End If
      ElseIf fpcmbBillType.Text = "PERSONAL ONLY" Then
        If LLRec.TaxYear = PGTaxYear Then
          Print #RptHandle, Tab(5); "Customer Account: "; Tab(23); Using$("#####0", LLRec.CustAcct); Tab(33); "Prev Tax Years: "; Tab(55); Using$("$###,###,##0.00", LLRec.PrevBal)
        Else
          Print #RptHandle, Tab(5); "Customer Account: "; Tab(23); Using$("#####0", LLRec.CustAcct); Tab(33); "Other Tax Years: "; Tab(55); Using$("$###,###,##0.00", LLRec.PrevBal)
        End If
      End If
      Print #RptHandle, Tab(5); "Tax Year: "; Tab(25); Using("###0", LLRec.TaxYear); Tab(33); "Curr Taxes Due: "; Tab(55); Using$("$###,###,##0.00", LLRec.CurrBal)
      Print #RptHandle, Tab(33); "Total Taxes Due: "; Tab(55); Using$("$###,###,##0.00", LLRec.TotBal)
      Print #RptHandle,
      For y = 11 To 20
        Print #RptHandle, LLtrRec.Body(y)
      Next y
      Print #RptHandle,
      Print #RptHandle,
      Print #RptHandle,
      Print #RptHandle,
      Print #RptHandle,
      Print #RptHandle, FF$
    ElseIf NegOpt = "Y" Then
      If LLRec.PrevBal >= 0 Then
        If fpcmbBillType.Text = "REAL ONLY" Then
          If LLRec.TaxYear = RGTaxYear Then
            Print #RptHandle, Tab(5); "Customer Account: "; Tab(23); Using$("#####0", LLRec.CustAcct); Tab(33); "Prev Tax Years: "; Tab(55); Using$("$###,###,##0.00", LLRec.PrevBal)
          Else
            Print #RptHandle, Tab(5); "Customer Account: "; Tab(23); Using$("#####0", LLRec.CustAcct); Tab(33); "Other Tax Years: "; Tab(55); Using$("$###,###,##0.00", LLRec.PrevBal)
          End If
        ElseIf fpcmbBillType.Text = "PERSONAL ONLY" Then
          If LLRec.TaxYear = PGTaxYear Then
            Print #RptHandle, Tab(5); "Customer Account: "; Tab(23); Using$("#####0", LLRec.CustAcct); Tab(33); "Prev Tax Years: "; Tab(55); Using$("$###,###,##0.00", LLRec.PrevBal)
          Else
            Print #RptHandle, Tab(5); "Customer Account: "; Tab(23); Using$("#####0", LLRec.CustAcct); Tab(33); "Other Tax Years: "; Tab(55); Using$("$###,###,##0.00", LLRec.PrevBal)
          End If
        End If
        Print #RptHandle, Tab(5); "Tax Year: "; Tab(25); Using("###0", LLRec.TaxYear); Tab(33); "Curr Taxes Due: "; Tab(55); Using$("$###,###,##0.00", LLRec.CurrBal)
        Print #RptHandle, Tab(33); "Total Taxes Due: "; Tab(55); Using$("$###,###,##0.00", LLRec.TotBal)
        Print #RptHandle,
        For y = 11 To 20
          Print #RptHandle, LLtrRec.Body(y)
        Next y
        Print #RptHandle,
        Print #RptHandle,
        Print #RptHandle,
        Print #RptHandle,
        Print #RptHandle,
        Print #RptHandle, FF$
      ElseIf LLRec.PrevBal < 0 Then
'        If fpcmbBillType.Text = "REAL ONLY" Then
'          If LLRec.TaxYear = RGTaxYear Then
'            Print #RptHandle, Tab(5); "Customer Account: "; Tab(23); Using$("#####0", LLRec.CustAcct)
'        ElseIf fpcmbBillType.Text = "PERSONAL ONLY" Then
'          If LLRec.TaxYear = PGTaxYear Then
'            Print #RptHandle, Tab(5); "Customer Account: "; Tab(23); Using$("#####0", LLRec.CustAcct); Tab(33); "Prev Tax Years: "; Tab(55); Using$("$###,###,##0.00", LLRec.PrevBal)
'          Else
'            Print #RptHandle, Tab(5); "Customer Account: "; Tab(23); Using$("#####0", LLRec.CustAcct); Tab(33); "Other Tax Years: "; Tab(55); Using$("$###,###,##0.00", LLRec.PrevBal)
'          End If
'        End If
        
        Print #RptHandle, Tab(5); "Tax Year: "; Tab(25); Using("###0", LLRec.TaxYear) '; Tab(33); "Curr Taxes Due: "; Tab(55); Using$("$###,###,##0.00", LLRec.CurrBal)
        Print #RptHandle, Tab(5); "Customer Account: "; Tab(23); Using$("#####0", LLRec.CustAcct); Tab(33); "Total Taxes Due: "; Tab(55); Using$("$###,###,##0.00", LLRec.TotBal)
        Print #RptHandle,
        For y = 11 To 20
          Print #RptHandle, LLtrRec.Body(y)
        Next y
        Print #RptHandle,
        Print #RptHandle,
        Print #RptHandle,
        Print #RptHandle,
        Print #RptHandle,
        Print #RptHandle, FF$
      End If
    End If
NextYear:
  Next x
  
  Close
  
  ViewPrint RptFile, "Printing Late Notice Letters", True
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPrintLateNotice", "PrintTextSelfEdit1", Erl)
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
