VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmFAValueRange 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assets By Value Range"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmFAValueRange.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   6930
      Left            =   1965
      TabIndex        =   7
      Top             =   1080
      Width           =   7785
      _Version        =   196609
      _ExtentX        =   13737
      _ExtentY        =   12213
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmFAValueRange.frx":08CA
      Begin LpLib.fpCombo fpcmbValue 
         Height          =   405
         Left            =   3210
         TabIndex        =   1
         ToolTipText     =   "Select the value type this report will display."
         Top             =   2115
         Width           =   3240
         _Version        =   196608
         _ExtentX        =   5715
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
         ColDesigner     =   "frmFAValueRange.frx":08E6
      End
      Begin LpLib.fpCombo fpcomboPrintOpt 
         Height          =   405
         Left            =   3600
         TabIndex        =   6
         ToolTipText     =   "Select  Graphic for a robust report that takes more time to process. Select Text for a faster report."
         Top             =   4995
         Width           =   2370
         _Version        =   196608
         _ExtentX        =   4180
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
         ColDesigner     =   "frmFAValueRange.frx":0BDD
      End
      Begin LpLib.fpCombo fpcmbOrder 
         Height          =   405
         Left            =   3210
         TabIndex        =   0
         ToolTipText     =   "Select the order this report will display data."
         Top             =   1530
         Width           =   3240
         _Version        =   196608
         _ExtentX        =   5715
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
         ColDesigner     =   "frmFAValueRange.frx":0ED4
      End
      Begin LpLib.fpCombo fpcmbYN 
         Height          =   405
         Left            =   5520
         TabIndex        =   3
         ToolTipText     =   "Enter Y to include disposed of fixed assets or N to exclude disposed of fixed assets."
         Top             =   3270
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
         ColDesigner     =   "frmFAValueRange.frx":11CB
      End
      Begin EditLib.fpCurrency fpcurrLeast 
         Height          =   396
         Left            =   4116
         TabIndex        =   4
         ToolTipText     =   "Enter the smallest item value this report will include."
         Top             =   3840
         Width           =   1596
         _Version        =   196608
         _ExtentX        =   2815
         _ExtentY        =   698
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
      Begin EditLib.fpText fptxtDeptNum 
         Height          =   396
         Left            =   3072
         TabIndex        =   2
         ToolTipText     =   "If Report Order is DEPARTMENT NUMBER then enter the desired department number which will appear in this report."
         Top             =   2688
         Width           =   1500
         _Version        =   196608
         _ExtentX        =   2646
         _ExtentY        =   698
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
         ThreeDInsideHighlightColor=   -2147483637
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
         ThreeDTextHighlightColor=   -2147483637
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
         CharValidationText=   "1 2 3 4 5 6 7 8 9 0 - A L a l"
         MaxLength       =   14
         MultiLine       =   0   'False
         PasswordChar    =   ""
         IncHoriz        =   0.25
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ScrollV         =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpCurrency fpcurrMost 
         Height          =   396
         Left            =   4128
         TabIndex        =   5
         ToolTipText     =   "Enter the highest item value this report will include."
         Top             =   4416
         Width           =   1596
         _Version        =   196608
         _ExtentX        =   2815
         _ExtentY        =   698
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
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   690
         Left            =   1590
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   5760
         Width           =   1875
         _Version        =   131072
         _ExtentX        =   3307
         _ExtentY        =   1217
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
         GrayAreaColor   =   13684944
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
         ButtonDesigner  =   "frmFAValueRange.frx":14C2
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   690
         Left            =   4560
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Click this button to create the report based on the parameters entered above."
         Top             =   5760
         Width           =   1890
         _Version        =   131072
         _ExtentX        =   3334
         _ExtentY        =   1217
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
         GrayAreaColor   =   13684944
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
         ButtonDesigner  =   "frmFAValueRange.frx":16A0
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdDept 
         Height          =   405
         Left            =   4710
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Click this button to bring up a list of all current departments."
         Top             =   2685
         Width           =   1350
         _Version        =   131072
         _ExtentX        =   2381
         _ExtentY        =   714
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
         GrayAreaColor   =   13684944
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
         ButtonDesigner  =   "frmFAValueRange.frx":187F
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Include Disposed Of Items (Y/N):"
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
         Left            =   1680
         TabIndex        =   15
         Top             =   3360
         Width           =   3660
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Value Type:"
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
         Left            =   1584
         TabIndex        =   14
         Top             =   2160
         Width           =   1452
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dept #"
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
         Left            =   1920
         TabIndex        =   13
         Top             =   2784
         Width           =   924
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Report Order:"
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
         Left            =   1200
         TabIndex        =   12
         Top             =   1584
         Width           =   1836
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Max Amount:"
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
         Left            =   2352
         TabIndex        =   11
         Top             =   4560
         Width           =   1548
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Least Amount:"
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
         Left            =   2196
         TabIndex        =   10
         Top             =   3936
         Width           =   1692
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   684
         Left            =   1536
         Top             =   432
         Width           =   4908
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Assets By Value Report"
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
         Height          =   492
         Left            =   1584
         TabIndex        =   9
         Top             =   576
         Width           =   4812
      End
      Begin VB.Label Label6 
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
         Left            =   1920
         TabIndex        =   8
         Top             =   5088
         Width           =   1500
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   7164
      Left            =   1836
      Top             =   960
      Width           =   7980
   End
End
Attribute VB_Name = "frmFAValueRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Dim DsplYNFlag As Boolean
  Private Temp_Class As Resize_Class

Private Sub cmdDept_Click()
  frmFADeptList.Show vbModal

End Sub

Private Sub cmdExit_Click()
  frmFAReportMenu.Show
  Close
  DoEvents
  KillFile "valrpt.dat"
  Unload frmFAValueRange

End Sub

Private Sub cmdProcess_Click()
  If fpcomboPrintOpt.Text = "Graphical" Then
    Call PrintGraphics
  ElseIf fpcomboPrintOpt.Text = "Text" Then
    MsgBox "Pitch 17 is recommended for this report."
    Call PrintText
  Else
    Exit Sub
  End If
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  DsplYNFlag = False
  Call LoadMe
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
'    'Me.Visible = False
    Temp_Class.ResizeControls Me
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
    Case vbKeyF8:
      SendKeys "%D"
      Call cmdDept_Click
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
      KillFile "valrpt.dat"
      ClearInUse PWcnt
      MainLog ("FixedAssets.exe terminated via menu bar on frmFAValueRange.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub fpcmbOrder_Change()
  'default is ALL for this combo box
  If QPTrim$(fpcmbOrder.Text) = "TAG NUMBER" Then
    fptxtDeptNum.Enabled = False
    cmdDept.Enabled = False
    fptxtDeptNum.Text = "ALL"
  ElseIf QPTrim$(fpcmbOrder.Text) = "" Then
    fpcmbOrder.Text = "TAG NUMBER"
    fptxtDeptNum.Enabled = False
    cmdDept.Enabled = False
    fptxtDeptNum.Text = "ALL"
  Else
    fptxtDeptNum.Enabled = True
    cmdDept.Enabled = True
  End If
  
End Sub

Private Sub fpcmbOrder_KeyDown(KeyCode As Integer, Shift As Integer)
  'this prevents the user from inadvertently changing data in the combo box when
  'tabbing through the fields
  If KeyCode = vbKeySpace Then
    fpcmbOrder.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbOrder.ListIndex = -1
  End If
  If fpcmbOrder.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
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
  Dim One As Integer
  Dim FileHandle As Integer
  One = 1
  FileHandle = FreeFile
  Open "valrpt.dat" For Output As FileHandle Len = 2
  Print #FileHandle, One
  Close FileHandle
  fpcmbOrder.Text = "TAG NUMBER"
  fpcmbOrder.AddItem "TAG NUMBER"
  fpcmbOrder.AddItem "DEPARTMENT NUMBER"
  fpcmbValue.Text = "PURCHASE PRICE"
  fpcmbValue.AddItem "PURCHASE PRICE"
  fpcmbValue.AddItem "CURRENT VALUE"
  fptxtDeptNum.Text = "ALL"
  fpcomboPrintOpt.AddItem "Graphical"
  fpcomboPrintOpt.AddItem "Text"
  fpcomboPrintOpt.Text = "Graphical"
  fpcmbYN.Text = "N"
  fpcmbYN.AddItem "Y"
  fpcmbYN.AddItem "N"
  
End Sub


Private Function Check4ValidDept() As Boolean
  Dim x As Integer
  Dim DeptIdx As DeptNumbSortIdxType
  Dim DIdxHandle As Integer
  Dim DIdxRecNums As Integer
  Dim ThisDept$
  
  On Error GoTo ERRORSTUFF
  
  Check4ValidDept = True
  OpenDeptIdxFile DIdxHandle
  DIdxRecNums = LOF(DIdxHandle) \ Len(DeptIdx)
  If DIdxRecNums = 0 Then
    MsgBox "No departments saved in index."
    Close
    Check4ValidDept = False
    Exit Function
  End If
  
  If QPTrim$(fptxtDeptNum.Text) = "ALL" Then
    Close
    Exit Function
  End If
  
  ThisDept$ = QPTrim$(fptxtDeptNum.Text)
  
  For x = 1 To DIdxRecNums
    Get DIdxHandle, x, DeptIdx
    If ThisDept$ = QPTrim$(DeptIdx.DeptNumb) Then
      Close
      Exit Function
    End If
  Next x
  
  MsgBox "No department number matches this entry. Please try again."
  Check4ValidDept = False
  fptxtDeptNum.SetFocus
  Close
  
  Exit Function
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAValueRange", "Check4ValidDept", Erl)
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
    ClearInUse (PWcnt)
    Terminate
    Unload Me
End Function

Private Sub fpcmbValue_Change()
  'PURCHASE PRICE is the default
  If QPTrim$(fpcmbValue.Text) = "" Then
    fpcmbValue.Text = "PURCHASE PRICE"
  End If
End Sub

Private Sub fpcmbValue_KeyDown(KeyCode As Integer, Shift As Integer)
  'this prevents the user from inadvertently changing data in the combo box when
  'tabbing through the fields
  If KeyCode = vbKeySpace Then
    fpcmbValue.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbValue.ListIndex = -1
  End If
  If fpcmbValue.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcomboPrintOpt_Change()
  'Graphical is the default
  If QPTrim$(fpcomboPrintOpt.Text) = "" Then
    fpcomboPrintOpt.Text = "Graphical"
  End If
End Sub

Private Sub fpcomboPrintOpt_KeyDown(KeyCode As Integer, Shift As Integer)
  'this prevents the user from inadvertently changing data in the combo box when
  'tabbing through the fields
  If KeyCode = vbKeySpace Then
    fpcomboPrintOpt.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboPrintOpt.ListIndex = -1
  End If
  If fpcomboPrintOpt.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      cmdExit.SetFocus
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
  Dim DOrigCost#(2), DBookTotal#(2), DCDep#(2), DYDep#(2), OrigCost#(2), BookTotal#(2), CDep#(2), YDep#(2)
  Dim ReportFile$
  Dim Dash80$
  Dim FF$
  Dim MaxLines As Integer
  Dim ItemCnt&
  Dim LineCnt&
  Dim Dept$
  Dim BAmt As Double
  Dim EAmt As Double
  Dim RptHandle As Integer
  Dim FAHandle As Integer
  Dim NumOfFARecs As Integer
  Dim FAItemRec As FAItemRecType
  Dim cnt&
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim ThisAmt As Double
  Dim DFlag As Boolean
  Dim AFlag As Boolean
  Dim DeptNumber As Integer
  Dim YTDDep#
  Dim x As Integer
  Dim Nextx As Integer
  Dim DIdxRec As DeptNumbSortIdxType
  Dim DIdxHandle As Integer
  Dim DIdxCnt As Integer
  Dim Index$
  Dim Page As Integer
  Dim DataFlag As Boolean
  Dim TagFlag As Boolean
  Dim AssetRange$
  Dim DeptCnt As Integer
  Dim DeptDescription$
  Dim DeptDescHeader$
  Dim TotalItems As Long
  Dim TagPrint As Boolean
  Dim LifeLeft As String * 3
  Dim WholeLife As String * 3
  Dim LifeData As String * 7
  Dim ValType As String * 1
  Dim FirstFlag As Boolean
  Dim ItemTotal As Long
  
  On Error GoTo ERRORSTUFF
  
  FirstFlag = True
  If QPTrim$(fpcmbValue.Text) = "PURCHASE PRICE" Then
    ValType = "P"
  Else
    ValType = "C"
  End If
  
  TagPrint = False
  If fpcurrLeast.Text < 0 Then
    MsgBox "Please enter a valid value for Least Amount."
    fpcurrLeast.SetFocus
    Exit Sub
  End If
  
  If fpcurrMost.Text = 0 Then
    MsgBox "Please enter a value for Most Amount."
    fpcurrMost.SetFocus
    Exit Sub
  End If
  
  BAmt = fpcurrLeast
  EAmt = fpcurrMost
  If EAmt < BAmt Then
    MsgBox "The Least Amount is greater than the Most Amount. Please reenter these values"
    fpcurrLeast.SetFocus
    Exit Sub
  End If
  
  If Check4ValidDept = False Then Exit Sub
  
  AssetRange = fpcurrLeast + " to " + fpcurrMost
  
  ReportFile$ = "FAVALUERPT.PRN"  'Report File Name
  Dash80$ = String$(80, "=")
  FF$ = Chr$(12)
  
  MaxLines = 56
  LineCnt& = 0
  ItemCnt& = 0
  Dept$ = QPTrim$(fptxtDeptNum.Text)
  
  RptHandle = FreeFile
  Index$ = QPTrim$(fpcmbOrder.Text)
  Open ReportFile$ For Output As #RptHandle
  
  OpenTagIdxFile TagIdxHandle
  NumOfFARecs = LOF(TagIdxHandle) \ Len(TagIdx)
  If NumOfFARecs = 0 Then
    MsgBox "No item records on file."
    Close TagIdxHandle
    Exit Sub
  End If
  
  ReDim TagIdxRecs(1 To NumOfFARecs)
  For x = 1 To NumOfFARecs
    Get TagIdxHandle, x, TagIdx
    TagIdxRecs(x) = TagIdx.DataRecNum
  Next x
  Close TagIdxHandle
  
  OpenDeptIdxFile DIdxHandle
  DIdxCnt = LOF(DIdxHandle) \ Len(DIdxRec)
  ReDim DeptNum(1 To DIdxCnt) As String
  ReDim DeptDesc(1 To DIdxCnt) As String
  
  For x = 1 To DIdxCnt
    Get DIdxHandle, x, DIdxRec
    DeptNum(x) = QPTrim$(DIdxRec.DeptNumb)
    DeptDesc(x) = QPTrim$(DIdxRec.DeptIdxDesc)
  Next x
  Close DIdxHandle
  
  If Dept$ <> "ALL" Then
    DeptNumber = Val(QPTrim$(fptxtDeptNum.Text))
    For x = 1 To DIdxCnt ' + 1
      If DeptNumber = DeptNum(x) Then
        DeptDescription = QPTrim$(DeptDesc(x))
        DeptDescHeader$ = DeptDescription
        Exit For
      End If
    Next x
  Else
    DeptNumber = Val(QPTrim(DeptNum(1)))
    DeptDescription = QPTrim(DeptDesc(1))
    DeptDescHeader$ = ""
  End If
  
  GoSub PrintMasterHeader1
  
  ReDim ATagDOrigCost(1 To DIdxCnt) As Double
  ReDim ATagDBookTotal(1 To DIdxCnt) As Double
  ReDim ATagDYDep(1 To DIdxCnt) As Double
  ReDim ATagDCnt(1 To DIdxCnt) As Long
  OpenFAItemFile FAHandle
  
  TagFlag = False
  
  frmFAShowPctComp.Label1 = "Gathering Item Data"
  frmFAShowPctComp.Show
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdProcess.Enabled = False
  
GetTagTotals: 'to print the report for tag numbers only the report
  'runs through all items one time in tag order and prints pertinent
  'data...then it returns to this spot and runs back through the
  'following loop gathering department totals (just like the
  'DEPARTMENT NUMBERS option does) but does not print items by
  'department...it just prints department totals at the end of the
  'report
  Nextx = 1
  If TagFlag = True Then
    Index = "DEPARTMENT NUMBERS"
    LineCnt = 0
  End If
  
  Do 'this loop iterates once if there is 1 dept requested, DIdxCnt + 1 if
  'department numbers "ALL" is requested and DIdxCnt + 2 for Tag Numbers
  '...Tag Numbers requires one iteration ignoring departments to get a list of
  'all valid tags in numeric order then DIdxCnt + 1 skipping the itemized tag
  'data print out just to allow the dept totals to assimilate...
    DataFlag = False
    For cnt& = 1 To NumOfFARecs
      Get FAHandle, TagIdxRecs(cnt), FAItemRec
      If LineCnt& >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintMasterHeader1
      End If
      
      If ValType = "P" Then 'filter by purchase price
        ThisAmt = FAItemRec.ORGCOST
      Else 'filter by current value
        ThisAmt = FAItemRec.CURRVAL
      End If
      
      If ThisAmt < BAmt Or ThisAmt > EAmt Then
      'filter out items that don't fall inside the date parameters
        GoTo SkipEm1
      End If
      
      If fpcmbYN.Text = "N" Then
        If FAItemRec.DsplFlag = 2 Then GoTo SkipEm1
      End If

      YTDDep# = FAItemRec.DEP2DATE
      
      'TAG NUMBER is actually true (if selected by the user) for the first
      'complete iteration only so that all tag numbers can be printed in numeric
      'order. After that we do not want the tags itemized we only want department
      'totals figured so that at the end of the report a "totals by dept" section
      'can be printed.
      If QPTrim$(Index) = "TAG NUMBER" Then
        GoTo TagOnly1 'prints itemized tag data in numeric order
      ElseIf DeptNumber <> FAItemRec.IDEPT Then 'any time department data
      'is needed then the item falls into this part of the if statement
        GoTo SkipEm1 'if the prevailing department (in the numeric index)
        'doesn't match this item's dept number then we don'y want it now
      End If
      'at this point the item's dept matches the prevailing dept number
      'but since TagFlag (don't want itemized tag data anymore) is true then
      'go ahead and collect dept data
      If TagFlag = True Then GoTo TagOnly2
      
TagOnly1: 'printing valid tag data...skipped if TAG NUMBER is chosen and this is
          'not the first iteration
      If QPTrim$(fpcmbOrder.Text) <> "TAG NUMBER" And DeptCnt = 0 Then
        Print #RptHandle, String$(111, "=")
        Print #RptHandle, "Assets for Dept Number: "; DeptNumber; " "; DeptDescription
        Print #RptHandle, String$(111, "_")
        LineCnt& = LineCnt& + 3
      End If
      
      DataFlag = True
      LifeLeft = CStr(FAItemRec.LifeLeft)
      'format the asset's life data
      If Len(LifeLeft) = 2 Then
        LifeLeft = QPTrim$(LifeLeft)
      ElseIf Len(LifeLeft) = 1 Then
        LifeLeft = " " + QPTrim$(LifeLeft)
      End If
      If FAItemRec.ILIFE = 0 Then
        WholeLife = " 0"
      Else
        WholeLife = CStr(FAItemRec.ILIFE)
      End If
      
      RSet LifeData = QPTrim$(WholeLife) + "/" + LifeLeft
      Print #RptHandle, Tab(1); QPTrim$(FAItemRec.ItemTag); Tab(22); Left$(FAItemRec.IDESC1, 28);
      Print #RptHandle, Tab(51); Using("###0", FAItemRec.IDEPT);
      Print #RptHandle, Tab(59); LifeData;
      Print #RptHandle, Tab(68); Using("###,###,##0.00", CStr(FAItemRec.ORGCOST));
      If QPTrim$(FAItemRec.DEPYN) = "N" And FAItemRec.DsplFlag <> 2 Then
        Print #RptHandle, Tab(82); Using("###,###,##0.00", CStr(YTDDep#)); "*";
      Else
        Print #RptHandle, Tab(82); Using("###,###,##0.00", CStr(YTDDep#));
      End If
      Print #RptHandle, Tab(98); Using("###,###,##0.00", CStr(FAItemRec.CURRVAL))
      If Mid(fpcmbYN.Text, 1, 1) = "Y" Then
        If FAItemRec.DsplFlag = 2 Then
          Print #RptHandle, Tab(19); "Disposed Of On: "; Tab(35); MakeRegDate(FAItemRec.DispDate)
          LineCnt& = LineCnt& + 1
        ElseIf FAItemRec.DsplFlag = 1 Then
          Print #RptHandle, Tab(8); "Scheduled For Disposal On: "; Tab(35); MakeRegDate(FAItemRec.DispDate)
          LineCnt& = LineCnt& + 1
        End If
      End If
      LineCnt& = LineCnt& + 1
      ItemCnt& = ItemCnt& + 1
      ItemTotal = ItemTotal + 1
TagOnly2: 'collects data for each department for reporting totals

      'This if statement filters out the first iteration of TAG NUMBER
      'selection because we do not want to start accumulating dept data
      'until the second iteration
      If TagFlag = False And QPTrim$(Index) = "TAG NUMBER" Then GoTo SkipEm1
      
      'collects grand totals
      OrigCost#(2) = OrigCost#(2) + FAItemRec.ORGCOST
      BookTotal#(2) = BookTotal#(2) + (FAItemRec.CURRVAL)
      YDep#(2) = YDep#(2) + YTDDep#
      
      'collects dept totals
      DeptCnt = DeptCnt + 1
      ATagDCnt(Nextx) = DeptCnt
      TotalItems = TotalItems + 1
      DOrigCost#(2) = DOrigCost#(2) + FAItemRec.ORGCOST
      ATagDOrigCost(Nextx) = DOrigCost#(2)
      DBookTotal#(2) = DBookTotal#(2) + (FAItemRec.CURRVAL)
      ATagDBookTotal(Nextx) = DBookTotal#(2)
      DYDep#(2) = DYDep#(2) + YTDDep#
      ATagDYDep(Nextx) = DYDep#(2)
SkipEm1:

    Next cnt&
    'here we begin the iteration over again but this time TagFlag
    'becomes true so we know that this was originally TAG NUMBERS
    'and the first iteration is done
    If QPTrim$(Index) = "TAG NUMBER" And TagFlag = False Then
      TagFlag = True
      GoTo GetTagTotals
      Exit Do
    End If
    
    If TagFlag = True Then GoTo NoData 'don't want the next dept
    'data to print
    
    If DataFlag = False Then
      GoTo NoData
    End If
    
  'First Print Subtotals
    Print #RptHandle, String$(111, "_")
    Print #RptHandle, "Assets for Dept Number: "; DeptNumber; " "; DeptDescription;
    Print #RptHandle, Tab(68); Using("###,###,##0.00", CStr(DOrigCost#(2)));
    Print #RptHandle, Tab(82); Using("###,###,##0.00", CStr(DYDep#(2)));
    Print #RptHandle, Tab(98); Using("###,###,##0.00", CStr(DBookTotal#(2)))
    Print #RptHandle, "Total Items: "; DeptCnt
    Print #RptHandle, String$(111, "=")
    Print #RptHandle,
    LineCnt& = LineCnt& + 5
NoData:
    frmFAShowPctComp.ShowPctComp Nextx, DIdxCnt
    If frmFAShowPctComp.Out = True Then
      Close
      frmFAShowPctComp.Out = False
      EnableCloseButton Me.hwnd, True
      Me.cmdExit.Enabled = True
      Me.cmdProcess.Enabled = True
      Unload frmFAShowPctComp
      Exit Sub
    End If
    'if "ALL" is not selected then the user has selected a single
    'department so we have all the data we need at this point...exit
    If QPTrim$(Dept$) <> "ALL" Then Exit Do
    'if all the depts have been examined then time to go
    If Nextx = DIdxCnt Then Exit Do
    'move to the next dept
    Nextx = Nextx + 1
    'assign new dept to DeptNumber
    DeptNumber = Val(QPTrim$(DeptNum(Nextx)))
    DeptDescription = QPTrim$(DeptDesc(Nextx))
    'clear all dept totals
'    DOrigCost#(1) = 0
'    DBookTotal#(1) = 0
'    DYDep#(1) = 0
    DOrigCost#(2) = 0
    DBookTotal#(2) = 0
    DYDep#(2) = 0
    DeptCnt = 0
  Loop
  
  Unload frmFAShowPctComp
  frmFAShowPctComp.Out = False
  EnableCloseButton Me.hwnd, True
  Me.cmdExit.Enabled = True
  Me.cmdProcess.Enabled = True
  
  If ItemTotal = 0 Then 'no sense in displaying an empty report
    MsgBox "There are no fixed assets saved for this criteria"
    Close
    Exit Sub
  End If
  
  'only prints if TAG NUMBERS was selected
  If TagFlag = True Then
    GoSub PrintTagDeptTotals
  End If
  
  If TagPrint = False Then GoSub PrintMasterValueEnding1
  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
  
  Close         'Close all open files now
  
  ViewPrint ReportFile$, "Value By Purchase Price", True
  
  KillFile (ReportFile$)
  
  Exit Sub
  
PrintMasterHeader1:
  Page = Page + 1
  If ValType = "P" Then
    Print #RptHandle, Tab(30); "Master Asset Listing : Asset Value by Purchase Price Range"
  Else
    Print #RptHandle, Tab(30); "Master Asset Listing : Asset Value by Current Value Range"
  End If
  If FirstFlag = False Then
    Print #RptHandle, "Dept # "; DeptNumber; " "; DeptDescription
  End If
  Print #RptHandle, "Purchase Price from "; AssetRange
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
  Print #RptHandle, "* = DO NOT DEPRECIATE THIS ASSET"
  Print #RptHandle, Tab(1); "Asset Number"; Tab(22); "Description"; Tab(51); "Dept"; Tab(58); "Life/Left"; Tab(68); "Purchase Price"; Tab(84); "Total Deprec"; Tab(102); "Book Value"
  If DeptCnt > 0 Or fpcmbOrder.Text = "TAG NUMBER" Then
    Print #RptHandle, String$(111, "=")
    LineCnt = LineCnt + 1
  End If
  LineCnt& = 6
  If FirstFlag = True Then
    FirstFlag = False
    LineCnt = 6
  End If
  Return
  
PrintMasterValueEnding1:
  
  Page = Page + 1
  Print #RptHandle, FF$
  Print #RptHandle, Tab(30); "Master Asset Listing : Grand Totals"
  Print #RptHandle, "Purchase Price from "; AssetRange
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
  Print #RptHandle, Tab(18); "Total Items"; Tab(47); "Purchase Price"; Tab(63); "Total Deprec"; Tab(79); "Book Value"
  Print #RptHandle, String$(88, "=")
  Print #RptHandle, "Total Assets ";
  Print #RptHandle, Tab(21); TotalItems;
  Print #RptHandle, Tab(47); Using("###,###,##0.00", CStr(OrigCost#(2)));
  Print #RptHandle, Tab(61); Using("###,###,##0.00", CStr(YDep#(2)));
  Print #RptHandle, Tab(75); Using("###,###,##0.00", CStr(BookTotal#(2)))
  
  Print #RptHandle, FF$
  
  Return
  
PrintTagDeptTotals: 'print only if TAG NUMBERS was selected
  
  Page = Page + 1
  Print #RptHandle, FF$
  Print #RptHandle, Tab(30); "Master Asset Listing : Department Totals"
  Print #RptHandle, "Purchase Price from "; AssetRange
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
  Print #RptHandle, Tab(1); "Department"; Tab(15); "Description"; Tab(40); "Item Count"; Tab(68); "Purchase Price"; Tab(85); "Total Deprec"; Tab(102); "Book Value"
  Print #RptHandle, String$(111, "=")
  LineCnt = 6
  
  For x = 1 To DIdxCnt
    If QPTrim$(DeptNum(x)) = "" Then DeptNum(x) = "0"
    Print #RptHandle, Tab(3); Using("####0", DeptNum(x)); Tab(15); DeptDesc(x); Tab(40); Using("#####0", ATagDCnt(x)); Tab(68); Using("###,###,##0.00", CStr(ATagDOrigCost(x))); Tab(83); Using("###,###,##0.00", CStr(ATagDYDep(x))); Tab(98); Using("###,###,##0.00", CStr(ATagDBookTotal(x)))
    LineCnt = LineCnt + 1
    
    If LineCnt& >= MaxLines And x <> DIdxCnt Then
      LineCnt& = 0
      Page = Page + 1
      Print #RptHandle, FF$
      Print #RptHandle, Tab(20); "Master Asset Listing : Department Totals"
      Print #RptHandle, "Purchase Price from "; AssetRange
      Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
      Print #RptHandle, Tab(1); "Department"; Tab(15); "Description"; Tab(40); "Item Count"; Tab(69); "Purchase Price"; Tab(85); "Total Deprec"; Tab(101); "Book Value"
      Print #RptHandle, String$(111, "=")
      LineCnt = LineCnt + 5
    End If
  Next x
  
  If LineCnt <= 53 Then
    Print #RptHandle, String$(111, "=")
    Print #RptHandle, "Total Assets ";
    Print #RptHandle, Tab(40); Using("#####0", TotalItems);
    Print #RptHandle, Tab(68); Using("###,###,##0.00", CStr(OrigCost#(2)));
    Print #RptHandle, Tab(83); Using("###,###,##0.00", CStr(YDep#(2)));
    Print #RptHandle, Tab(98); Using("###,###,##0.00", CStr(BookTotal#(2)))
  Else
    Print #RptHandle, FF$
    Print #RptHandle, Tab(30); "Master Asset Listing : Department Totals"
    Print #RptHandle, "Purchase Price from "; AssetRange
    Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
    Print #RptHandle, Tab(1); "Department"; Tab(15); "Description"; Tab(40); "Item Count"; Tab(63); "Purchase Price"; Tab(80); "Total Deprec"; Tab(97); "Book Value"
    Print #RptHandle, String$(111, "=")
    Print #RptHandle, String$(111, "=")
    Print #RptHandle, "Total Assets ";
    Print #RptHandle, Tab(40); Using("#####0", TotalItems);
    Print #RptHandle, Tab(68); Using("###,###,##0.00", CStr(OrigCost#(2)));
    Print #RptHandle, Tab(83); Using("###,###,##0.00", CStr(YDep#(2)));
    Print #RptHandle, Tab(98); Using("###,###,##0.00", CStr(BookTotal#(2)))
  End If
  Print #RptHandle, FF$
  TagPrint = True
  
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAValueRange", "PrintText", Erl)
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
    ClearInUse (PWcnt)
    Terminate
    Unload Me
End Sub

Private Sub fptxtDeptNum_Change()
  'ALL is the default value
  If fptxtDeptNum.Text = "" Then
    fptxtDeptNum = "ALL"
  End If
End Sub

Private Sub PrintGraphics()
  Dim DOrigCost#(2), DBookTotal#(2), DCDep#(2), DYDep#(2), OrigCost#(2), BookTotal#(2), CDep#(2), YDep#(2)
  Dim ReportFile$
  Dim TagReportFile$
  Dim TagRptHandle As Integer
  Dim TDReportFile$
  Dim TDRptHandle As Integer
  Dim GTReportFile$
  Dim GTRptHandle As Integer
  Dim ItemCnt&
  Dim Dept$
  Dim BAmt As Double
  Dim EAmt As Double
  Dim RptHandle As Integer
  Dim FAHandle As Integer
  Dim NumOfFARecs As Integer
  Dim FAItemRec As FAItemRecType
  Dim cnt&
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim ThisAmt As Double
  Dim DeptNumber As Integer
  Dim YTDDep#
  Dim x As Integer
  Dim Nextx As Integer
  Dim DIdxRec As DeptNumbSortIdxType
  Dim DIdxHandle As Integer
  Dim DIdxCnt As Integer
  Dim Index$ ', Tab1 As Integer, Tab2 As Integer, Tab3 As Integer
  Dim DataFlag As Boolean
  Dim TagFlag As Boolean
  Dim AssetRange$
  Dim DeptCnt As Integer
  Dim DeptDescription$
  Dim DeptDescHeader$
  Dim TotalItems As Long
  Dim TagPrint As Boolean
  Dim FASetUpRec As FASetupRecType
  Dim FASHandle As Integer
  Dim Employer$
  Dim dlm$
  Dim EndRpt As Integer
  Dim ValType As String * 1
  Dim ItemTotal As Long
  
  On Error GoTo ERRORSTUFF
  
  If QPTrim$(fpcmbValue.Text) = "PURCHASE PRICE" Then
    ValType = "P"
  Else
    ValType = "C"
  End If
  
  If fpcurrLeast.Text < 0 Then
    MsgBox "Please enter a valid value for Least Amount."
    fpcurrLeast.SetFocus
    Exit Sub
  End If
  
  If fpcurrMost.Text = 0 Then
    MsgBox "Please enter a value for Most Amount."
    fpcurrMost.SetFocus
    Exit Sub
  End If
  
  dlm$ = "~"
  OpenFASetUpFile FASHandle
  Get FASHandle, 1, FASetUpRec
  Close FASHandle
  Employer = FASetUpRec.TownName
  
  TagPrint = False
  
  BAmt = fpcurrLeast
  EAmt = fpcurrMost
  If EAmt < BAmt Then
    MsgBox "The Least Amount is greater than the Most Amount. Please reenter these values"
    fpcurrLeast.SetFocus
    Exit Sub
  End If
  
  If Check4ValidDept = False Then Exit Sub
  
  AssetRange = Using$("$###,###,##0.00", CDbl(fpcurrLeast)) + " to " + Using$("$###,###,##0.00", CDbl(fpcurrMost))
  
  ReportFile$ = "FARPTS\FAVALUERPT.RPT"  'Report File Name
  TagReportFile$ = "FARPTS\FATAGVALUERPT.RPT"
  TDReportFile$ = "FARPTS\FATAGDEPTVALUE.RPT"
  GTReportFile$ = "FARPTS\FAGTVALUERPT.RPT"
  ItemCnt& = 0
  Dept$ = QPTrim$(fptxtDeptNum.Text)
  
  Index$ = QPTrim$(fpcmbOrder.Text)
  
  If QPTrim$(Index$) = "TAG NUMBER" Then
    TagRptHandle = FreeFile
    Open TagReportFile For Output As #TagRptHandle
  Else
    RptHandle = FreeFile
    Open ReportFile$ For Output As #RptHandle
  End If
  
  OpenTagIdxFile TagIdxHandle
  NumOfFARecs = LOF(TagIdxHandle) \ Len(TagIdx)
  If NumOfFARecs = 0 Then
    MsgBox "No item records on file."
    Close TagIdxHandle
    Exit Sub
  End If
  ReDim TagIdxRecs(1 To NumOfFARecs)
  For x = 1 To NumOfFARecs
    Get TagIdxHandle, x, TagIdx
    TagIdxRecs(x) = TagIdx.DataRecNum
  Next x
  Close TagIdxHandle
  
  OpenDeptIdxFile DIdxHandle
  DIdxCnt = LOF(DIdxHandle) \ Len(DIdxRec)
  ReDim DeptNum(1 To DIdxCnt) As String
  ReDim DeptDesc(1 To DIdxCnt) As String
  
  For x = 1 To DIdxCnt
    Get DIdxHandle, x, DIdxRec
    DeptNum(x) = QPTrim$(DIdxRec.DeptNumb)
    DeptDesc(x) = QPTrim$(DIdxRec.DeptIdxDesc)
  Next x
  Close DIdxHandle
  
  If Dept$ <> "ALL" Then
    DeptNumber = Val(QPTrim$(fptxtDeptNum.Text))
    For x = 1 To DIdxCnt
      If DeptNumber = DeptNum(x) Then
        DeptDescription = QPTrim$(DeptDesc(x))
        DeptDescHeader$ = DeptDescription
        Exit For
      End If
    Next x
  Else
    DeptNumber = Val(QPTrim(DeptNum(1)))
    DeptDescription = QPTrim(DeptDesc(1))
    DeptDescHeader$ = ""
  End If
  
  ReDim ATagDOrigCost(1 To DIdxCnt) As Double
  ReDim ATagDBookTotal(1 To DIdxCnt) As Double
  ReDim ATagDYDep(1 To DIdxCnt) As Double
  ReDim ATagDCnt(1 To DIdxCnt) As Long
  
  OpenFAItemFile FAHandle
  
  TagFlag = False
  
  frmFAShowPctComp.Label1 = "Gathering Item Data"
  frmFAShowPctComp.Show
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdExit.Enabled = False
  Me.cmdProcess.Enabled = False
  
GetTagTotals: 'to print the report for tag numbers only the report
  'runs through all items one time in tag order and prints pertinent
  'data...then it returns to this spot and runs back through the
  'following loop gathering department totals (just like the
  'DEPARTMENT NUMBERS option does) but does not print items by
  'department...it just prints department totals at the end of the
  'report
  Nextx = 1
  If TagFlag = True Then
    Index = "DEPARTMENT NUMBERS"
  End If
  
  Do 'this loop iterates once if there is 1 dept requested, DIdxCnt + 1 if
  'department numbers "ALL" is requested and DIdxCnt + 2 for Tag Numbers
  '...Tag Numbers requires one iteration ignoring departments to get a list of
  'all valid tags in numeric order then DIdxCnt + 1 skipping the itemized tag
  'data print out just to allow the dept totals to assimilate...
    DataFlag = False
    For cnt& = 1 To NumOfFARecs
      Get FAHandle, TagIdxRecs(cnt), FAItemRec
      If ValType = "P" Then
        ThisAmt = FAItemRec.ORGCOST
      Else
        ThisAmt = FAItemRec.CURRVAL
      End If
      
      If ThisAmt < BAmt Or ThisAmt > EAmt Then
      'filter out items that don't fall inside the date parameters
        GoTo SkipEm1
      End If
      'Can't depreciate something that has no asset life...this should
      'not be necessary unless the asset life was changed to zero after
      'depreciation had already started
      If fpcmbYN.Text = "N" Then
        If FAItemRec.DsplFlag = 2 Then GoTo SkipEm1
      End If

      YTDDep# = FAItemRec.DEP2DATE
      
      'TAG NUMBER is actually true (if selected by the user) for the first
      'complete iteration only so that all tag numbers can be printed in numeric
      'order. After that we do not want the tags itemized we only want department
      'totals figured so that at the end of the report a "totals by dept" section
      'can be printed.
      If QPTrim$(Index) = "TAG NUMBER" Then
        GoTo TagOnly1 'prints itemized tag data in numeric order
      ElseIf DeptNumber <> FAItemRec.IDEPT Then 'any time department data
      'is needed then the item falls into this part of the if statement
        GoTo SkipEm1 'if the prevailing department (in the numeric index)
        'doesn't match this item's dept number then we don'y want it now
      End If
      'at this point the item's dept matches the prevailing dept number
      'but since TagFlag (don't want itemized tag data anymore) is true then
      'go ahead and collect dept data
      If TagFlag = True Then GoTo TagOnly2
      
TagOnly1: 'printing valid tag data...skipped if TAG NUMBER is chosen and this is
          'not the first iteration
      DataFlag = True
      If TagRptHandle > 0 Then
        '                        0              1
        Print #TagRptHandle, Employer; dlm; AssetRange; dlm;
        '                             2                       3
        Print #TagRptHandle, FAItemRec.ItemTag; dlm; FAItemRec.IDESC1; dlm;
        '                            4
        Print #TagRptHandle, FAItemRec.IDEPT; dlm;
        '                            5
        Print #TagRptHandle, FAItemRec.ILIFE; dlm;
        '                            6
        Print #TagRptHandle, FAItemRec.ORGCOST; dlm;
        If QPTrim$(FAItemRec.DEPYN) = "N" And FAItemRec.DsplFlag <> 2 Then
          '                       7           8
          Print #TagRptHandle, YTDDep#; dlm; "*"; dlm;
        Else
          '                       7           8
          Print #TagRptHandle, YTDDep#; dlm; " "; dlm;
        End If
        '                                       9
        Print #TagRptHandle, FAItemRec.CURRVAL; dlm;
        '                     10              11                   12                  13
        Print #TagRptHandle, Dept$; dlm; DeptDescHeader$; dlm; DeptNumber; dlm; DeptDescription; dlm;
        '                         14               15               16                 17
        Print #TagRptHandle, OrigCost#(2); dlm; YDep#(2); dlm; BookTotal#(2); dlm; TotalItems; dlm;
        If DsplYNFlag = False Then
          '                         18                      19                   20         21                22
          Print #TagRptHandle, FAItemRec.DEPYN; dlm; FAItemRec.LifeLeft; dlm; ValType; dlm; ""; dlm; Mid(fpcmbYN.Text, 1, 1)
        Else
          If FAItemRec.DsplFlag = 2 Then
            '                         18                     19                   20                21                                       22
            Print #TagRptHandle, FAItemRec.DEPYN; dlm; FAItemRec.LifeLeft; dlm; ValType; dlm; MakeRegDate(FAItemRec.DispDate); dlm; Mid(fpcmbYN.Text, 1, 1)
          ElseIf FAItemRec.DsplFlag = 1 Then
            '                         18                      19               20             21               22
            Print #TagRptHandle, FAItemRec.DEPYN; dlm; FAItemRec.LifeLeft; dlm; ValType; dlm; "P" + MakeRegDate(FAItemRec.DispDate); dlm; Mid(fpcmbYN.Text, 1, 1)
          Else
            '                         18                      19                   20         21                22
            Print #TagRptHandle, FAItemRec.DEPYN; dlm; FAItemRec.LifeLeft; dlm; ValType; dlm; ""; dlm; Mid(fpcmbYN.Text, 1, 1)
          End If
        End If
      Else
        '                     0              1
        Print #RptHandle, Employer; dlm; AssetRange; dlm;
        '                          2                        3
        Print #RptHandle, FAItemRec.ItemTag; dlm; FAItemRec.IDESC1; dlm;
        '                        4
        Print #RptHandle, FAItemRec.IDEPT; dlm;
        '                        5
        Print #RptHandle, FAItemRec.ILIFE; dlm;
        '                        6
        Print #RptHandle, FAItemRec.ORGCOST; dlm;
        If QPTrim$(FAItemRec.DEPYN) = "N" And FAItemRec.DsplFlag <> 2 Then
          '                    7           8
          Print #RptHandle, YTDDep#; dlm; "*"; dlm;
        Else
          '                    7           8
          Print #RptHandle, YTDDep#; dlm; " "; dlm;
        End If
        '                                     9
        Print #RptHandle, FAItemRec.CURRVAL; dlm;
        '                  10               11                  12                 13
        Print #RptHandle, Dept$; dlm; DeptDescHeader$; dlm; DeptNumber; dlm; DeptDescription; dlm;
        '                      14                15                 16                17
        Print #RptHandle, DOrigCost#(2); dlm; DYDep#(2); dlm; DBookTotal#(2); dlm; DeptCnt; dlm;
        '                      18               19               20                 21
        Print #RptHandle, OrigCost#(2); dlm; YDep#(2); dlm; BookTotal#(2); dlm; TotalItems; dlm;
        If DsplYNFlag = False Then
          '                         22                      23               24          25
          Print #RptHandle, FAItemRec.DEPYN; dlm; FAItemRec.LifeLeft; dlm; ValType; dlm; ""; dlm; Mid(fpcmbYN.Text, 1, 1)
        Else
          If FAItemRec.DsplFlag = 2 Then
            '                         22                      23               24                25
            Print #RptHandle, FAItemRec.DEPYN; dlm; FAItemRec.LifeLeft; dlm; ValType; dlm; MakeRegDate(FAItemRec.DispDate); dlm; Mid(fpcmbYN.Text, 1, 1)
          ElseIf FAItemRec.DsplFlag = 1 Then
            '                         22                      23               24                25
            Print #RptHandle, FAItemRec.DEPYN; dlm; FAItemRec.LifeLeft; dlm; ValType; dlm; "P" + MakeRegDate(FAItemRec.DispDate); dlm; Mid(fpcmbYN.Text, 1, 1)
          Else
            '                         22                      23               24          25
            Print #RptHandle, FAItemRec.DEPYN; dlm; FAItemRec.LifeLeft; dlm; ValType; dlm; ""; dlm; Mid(fpcmbYN.Text, 1, 1)
          End If
        End If
      End If
      
      ItemCnt& = ItemCnt& + 1
      ItemTotal = ItemTotal + 1
TagOnly2: 'collects data for each department for reporting totals

      'This if statement filters out the first iteration of TAG NUMBER
      'selection because we do not want to start accumulating dept data
      'until the second iteration
      If TagFlag = False And QPTrim$(Index) = "TAG NUMBER" Then GoTo SkipEm1
      
      'collects grand totals
      OrigCost#(2) = OrigCost#(2) + FAItemRec.ORGCOST
      BookTotal#(2) = BookTotal#(2) + (FAItemRec.CURRVAL)
      YDep#(2) = YDep#(2) + YTDDep#
      
      'collects dept totals
      DeptCnt = DeptCnt + 1
      ATagDCnt(Nextx) = DeptCnt
      TotalItems = TotalItems + 1
      DOrigCost#(2) = DOrigCost#(2) + FAItemRec.ORGCOST
      ATagDOrigCost(Nextx) = DOrigCost#(2)
      DBookTotal#(2) = DBookTotal#(2) + (FAItemRec.CURRVAL)
      ATagDBookTotal(Nextx) = DBookTotal#(2)
      DYDep#(2) = DYDep#(2) + YTDDep#
      ATagDYDep(Nextx) = DYDep#(2)
SkipEm1:

    Next cnt&
    'here we begin the iteration over again but this time TagFlag
    'becomes true so we know that this was originally TAG NUMBERS
    'and the first iteration is done
    If QPTrim$(Index) = "TAG NUMBER" And TagFlag = False Then
      TagFlag = True
      GoTo GetTagTotals
      Exit Do
    End If
    
    If TagFlag = True Then GoTo NoData 'don't want the next dept
    'data to print
    
NoData:
    frmFAShowPctComp.ShowPctComp Nextx, DIdxCnt ' + 1
    If frmFAShowPctComp.Out = True Then
      Close
      frmFAShowPctComp.Out = False
      EnableCloseButton Me.hwnd, True
      Me.cmdExit.Enabled = True
      Me.cmdProcess.Enabled = True
      Unload frmFAShowPctComp
      Exit Sub
    End If
    'if "ALL" is not selected then the user has selected a single
    'department so we have all the data we need at this point...exit
    If QPTrim$(Dept$) <> "ALL" Then Exit Do
    'if all the depts have been examined then time to go
    If Nextx = DIdxCnt Then Exit Do
    'move to the next dept
    Nextx = Nextx + 1
    'assign new dept to DeptNumber
    DeptNumber = Val(QPTrim$(DeptNum(Nextx)))
    DeptDescription = QPTrim$(DeptDesc(Nextx))
    'clear all dept totals
    DOrigCost#(1) = 0
    DBookTotal#(1) = 0
    DYDep#(1) = 0
    DOrigCost#(2) = 0
    DBookTotal#(2) = 0
    DYDep#(2) = 0
    DeptCnt = 0
   Loop
  
  Unload frmFAShowPctComp
  frmFAShowPctComp.Out = False
  EnableCloseButton Me.hwnd, True
  Me.cmdExit.Enabled = True
  Me.cmdProcess.Enabled = True
  
  'only prints if TAG NUMBERS was selected
  Close         'Close all open files now
  
  If ItemTotal = 0 Then
    MsgBox "There are no fixed assets saved for this criteria"
    Exit Sub
  End If
  If TagFlag = True Then
    GoSub PrintTagDeptTotals
  End If
  
  If TagFlag = True Then
    arFATagValueRpt.Show
  Else
    arFAValueRpt.Show
  End If
  
  frmFALoadReport.Show
  
  Exit Sub
  
  
PrintTagDeptTotals: 'print only if TAG NUMBERS was selected
  
  GTRptHandle = FreeFile
  Open GTReportFile$ For Output As GTRptHandle
  EndRpt = 1
  For x = 1 To DIdxCnt ' + 1
    If QPTrim$(DeptNum(x)) = "" Then DeptNum(x) = "0"
    '                        0                1                2                   3                      4                    5
    Print #GTRptHandle, DeptNum(x); dlm; DeptDesc(x); dlm; ATagDCnt(x); dlm; ATagDOrigCost(x); dlm; ATagDYDep(x); dlm; ATagDBookTotal(x); dlm;
    '                        6                7                8                9
    Print #GTRptHandle, TotalItems; dlm; OrigCost#(2); dlm; YDep#(2); dlm; BookTotal#(2); dlm; EndRpt
  Next x
    EndRpt = 2
      '                        0                1                2                   3                      4                    5
    Print #GTRptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    '                        6                7                8                9
    Print #GTRptHandle, "TotalItems"; dlm; OrigCost#(2); dlm; YDep#(2); dlm; BookTotal#(2); dlm; EndRpt

  Close GTRptHandle
  
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAValueRange", "PrintGraphics", Erl)
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
    ClearInUse (PWcnt)
    Terminate
    Unload Me
End Sub

Private Sub fpcmbYN_Change()
  'default this field to N
  If QPTrim$(fpcmbYN.Text) <> "Y" And QPTrim$(fpcmbYN.Text) <> "N" Then
    fpcmbYN.Text = "N"
  End If
  If QPTrim$(fpcmbYN.Text) = "Y" Then
    DsplYNFlag = True
  ElseIf QPTrim$(fpcmbYN.Text) = "N" Then
    DsplYNFlag = False
  End If
End Sub

Private Sub fpcmbYN_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  If KeyCode = vbKeySpace Then
    fpcmbYN.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbYN.ListIndex = -1
  End If
  If fpcmbYN.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcurrLeast.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

