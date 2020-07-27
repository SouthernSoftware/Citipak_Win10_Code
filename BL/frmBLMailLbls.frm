VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLMailLbls 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Print Customer Mailing Labels"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmBLMailLbls.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   7140
      Left            =   1560
      TabIndex        =   5
      Top             =   862
      Width           =   8505
      _Version        =   196609
      _ExtentX        =   15002
      _ExtentY        =   12594
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmBLMailLbls.frx":08CA
      Begin LpLib.fpCombo fpcmbLabel 
         Height          =   405
         Left            =   2760
         TabIndex        =   1
         Tag             =   $"frmBLMailLbls.frx":08E6
         Top             =   1845
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
         ColDesigner     =   "frmBLMailLbls.frx":09A4
      End
      Begin LpLib.fpCombo fpcmbPrintOrder 
         Height          =   405
         Left            =   3240
         TabIndex        =   0
         Tag             =   $"frmBLMailLbls.frx":0D0F
         Top             =   1290
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
         ColDesigner     =   "frmBLMailLbls.frx":0DBB
      End
      Begin LpLib.fpCombo fpcmbParameters 
         Height          =   405
         Left            =   3150
         TabIndex        =   2
         Tag             =   $"frmBLMailLbls.frx":1126
         Top             =   2400
         Width           =   3705
         _Version        =   196608
         _ExtentX        =   6535
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
         ColDesigner     =   "frmBLMailLbls.frx":1211
      End
      Begin LpLib.fpCombo fpcmbXPar 
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Tag             =   $"frmBLMailLbls.frx":157C
         Top             =   2955
         Width           =   3705
         _Version        =   196608
         _ExtentX        =   6535
         _ExtentY        =   661
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   10.5
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
         ColDesigner     =   "frmBLMailLbls.frx":17A1
      End
      Begin LpLib.fpList fpList 
         Height          =   1740
         Left            =   840
         TabIndex        =   19
         Tag             =   $"frmBLMailLbls.frx":1B0C
         Top             =   4080
         Width           =   6855
         _Version        =   196608
         _ExtentX        =   12091
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
         ColDesigner     =   "frmBLMailLbls.frx":1C85
      End
      Begin EditLib.fpDateTime fptxtXDate 
         Height          =   390
         Left            =   3390
         TabIndex        =   4
         Tag             =   $"frmBLMailLbls.frx":2088
         Top             =   3450
         Width           =   1785
         _Version        =   196608
         _ExtentX        =   3154
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
         UserEntry       =   1
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
         Text            =   "8/12/2003"
         DateCalcMethod  =   0
         DateTimeFormat  =   0
         UserDefinedFormat=   ""
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
         PopUpType       =   0
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
      Begin fpBtnAtlLibCtl.fpBtn cmdAlign 
         Height          =   495
         Left            =   2760
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   $"frmBLMailLbls.frx":21FB
         Top             =   6210
         Width           =   1545
         _Version        =   131072
         _ExtentX        =   2725
         _ExtentY        =   873
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
         ButtonDesigner  =   "frmBLMailLbls.frx":22DA
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   495
         Left            =   4395
         TabIndex        =   10
         TabStop         =   0   'False
         Tag             =   "Press 'Exit' to return to the main Customer Maintenance menu."
         Top             =   6210
         Width           =   1695
         _Version        =   131072
         _ExtentX        =   2990
         _ExtentY        =   873
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
         ButtonDesigner  =   "frmBLMailLbls.frx":24B6
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   495
         Left            =   6195
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   $"frmBLMailLbls.frx":2694
         Top             =   6210
         Width           =   1695
         _Version        =   131072
         _ExtentX        =   2990
         _ExtentY        =   873
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
         ButtonDesigner  =   "frmBLMailLbls.frx":272F
      End
      Begin fpBtnAtlLibCtl.fpBtn fpcmdXList 
         Height          =   405
         Left            =   5250
         TabIndex        =   14
         TabStop         =   0   'False
         Tag             =   $"frmBLMailLbls.frx":290E
         Top             =   3450
         Width           =   1935
         _Version        =   131072
         _ExtentX        =   3413
         _ExtentY        =   714
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
         ButtonDesigner  =   "frmBLMailLbls.frx":2A4C
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
         Height          =   495
         Left            =   645
         TabIndex        =   16
         TabStop         =   0   'False
         Tag             =   $"frmBLMailLbls.frx":2C32
         Top             =   6210
         Width           =   2025
         _Version        =   131072
         _ExtentX        =   3572
         _ExtentY        =   873
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
         ButtonDesigner  =   "frmBLMailLbls.frx":2CCF
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Expiration Parameters:"
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
         Left            =   1080
         TabIndex        =   18
         Top             =   3000
         Width           =   2505
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
         Height          =   255
         Left            =   600
         TabIndex        =   17
         Top             =   6750
         Width           =   2100
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Expiration Date:"
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
         Left            =   1410
         TabIndex        =   13
         Top             =   3555
         Width           =   1830
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Print Range:"
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
         Left            =   1605
         TabIndex        =   12
         Top             =   2490
         Width           =   1410
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   4905
         Left            =   480
         Top             =   1155
         Width           =   7575
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
         Left            =   1275
         TabIndex        =   8
         Top             =   1950
         Width           =   1350
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
         Left            =   1800
         TabIndex        =   7
         Top             =   1395
         Width           =   1305
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Mailing Labels"
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
         Left            =   2250
         TabIndex        =   6
         Top             =   450
         Width           =   3945
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   1830
         Top             =   315
         Width           =   4905
      End
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   450
      Left            =   600
      TabIndex        =   15
      Top             =   6877
      Width           =   780
      _Version        =   131072
      _ExtentX        =   1376
      _ExtentY        =   794
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
      MaxWidth        =   3000
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
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   7410
      Left            =   1440
      Top             =   727
      Width           =   8775
   End
End
Attribute VB_Name = "frmBLMailLbls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class

Private Sub cmdAlign_Click()
  Dim UBRpt As Integer
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
  
  If fpcmbLabel.Text = "2) 1 X 3 1/2 1 Wide Text" Then
    LType = 1
  ElseIf fpcmbLabel.Text = "3) 1 X 3 1/2 3 Wide Text" Then
    LType = 2
  ElseIf fpcmbLabel.Text = "4) 1 X 3 1/2 4 Wide Text" Then
    LType = 3
  Else
    fpcmbLabel.BackColor = &HFFFF&
    frmBLMessageBoxJr.Label1.Caption = "Please make a Label Type selection."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fpcmbLabel.BackColor = &H80000005
    fpcmbLabel.SetFocus
    Exit Sub
  End If
  
  UBRpt = FreeFile
  Open "ARLABEL.RPT" For Output As UBRpt
  Select Case LType
  Case 1
    For cnt = 1 To 5
      Print #UBRpt, Align$
    Next
    Print #UBRpt,
  Case 2
    For cnt = 1 To 5
      Print #UBRpt, Align$; Tab(OSet(2)); Align$; Tab(OSet(3)); Align$
    Next
    Print #UBRpt,
  Case 3
    For cnt = 1 To 5
      Print #UBRpt, Align$; Tab(OSet(2)); Align$; Tab(OSet(3)); Align$; Tab(OSet(4)); Align$
    Next
    Print #UBRpt,
  End Select

  Close UBRpt

  ViewPrint "ARLABEL.RPT", "Mailing Labels Alignment", True
  
  MainLog ("Mailing labels 'Align' feature used.")
  
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLMailLbls", "cmdAlign_Click", Erl)
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

Private Sub cmdExit_Click()
  'mllbls.dat.dat is created so if the user uses the
  'expiration list then when it closes it knows where the command came from
  'and returns data and the screen here
  'this same form is called from the Issue Applications menu
  'so it will return to which ever form has an outstanding
  '.dat file
  KillFile "mllbls.dat"
  If Exist("issueappslics.dat") Then
    frmBLIssueAppsLics.Show
  ElseIf Exist("custrptsmenu.dat") Then
    frmBLCustReportsMenu.Show
  Else
    frmBLCustMaintMenu.Show
  End If
  DoEvents
  Unload frmBLMailLbls
  
End Sub

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "F1 &Turn Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    lblBalloon.Visible = True
    fpcmbPrintOrder.ToolTipText = ""
    fpcmbLabel.ToolTipText = ""
    fpcmbParameters.ToolTipText = ""
    fptxtXDate.ToolTipText = ""
    fpcmdXList.ToolTipText = ""
    cmdAlign.ToolTipText = ""
    cmdProcess.ToolTipText = ""
    cmdExit.ToolTipText = ""
    
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
'    fpcmbPrintOrder.ToolTipText = "Mailing labels can be printed in alphabetical order or in numerical order."
'    fpcmbLabel.ToolTipText = "Select the mailing label option from the drop down box."
'    fpcmbParameters.ToolTipText = "The mailing labels printed can be resticted to just those businesses who have the specified expiration date or they can be printed for all businesses."
'    fptxtXDate.ToolTipText = "If you select expiration date as the printing parameter then enter the license expiration date for which the mailing labels will be printed."
'    fpcmdXList.ToolTipText = "Press for a complete list of customers and their expiration dates."
'    cmdAlign.ToolTipText = "Prints a mailing label template used to help in positioning the mailing labels so they will print accurately."
'    cmdProcess.ToolTipText = "Press to begin the printing process for mailing labels."
'    cmdExit.ToolTipText = "Press to exit this screen."
  End If
End Sub

Private Sub cmdProcess_Click()
  If InStr(fpcmbLabel.Text, "Graphical") Then
    Call PrintGraphics
  ElseIf InStr(fpcmbLabel.Text, "Text") Then
    frmBLMessageBoxJr.Label1.Caption = "Pitch 10 is recommended for this report."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Call PrintText
  Else
    Exit Sub
  End If
End Sub

Private Sub PrintText()
  Dim LType As Integer
  Dim ARFile As Integer
  Dim RptHandle As Integer
  Dim CustIdxHandle As Integer
  Dim CustNameIdxRec As CustNameIdxType
  Dim CustNumIdxRec As CustNumIdxType
  Dim CustSrchIdxRec As CustNameIdxType
  Dim NumOfCustIdx As Integer
  Dim x As Integer
  Dim CustRec As ARCustRecType
  Dim CustRCnt As Integer
  Dim Zip$
  Dim DidCnt As Integer
  Dim LabelCnt As Integer
  Dim PCnt As Integer
  Dim CustPCnt As Integer
  Dim AcctNumber&
  Dim cnt As Integer
  Dim ReportFile$
  Dim BusName As String * 23
  Dim CityName As String * 18
  Dim Address As String * 23
  Dim NameFlag As Boolean
  Dim XFlag As Boolean
  Dim XDate As Integer
  Dim ValidCnt As Integer
  Dim XParFlag As Integer
  Dim SpreadCnt As Integer
  
  On Error GoTo ERRORSTUFF
  
  cmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  
  If QPTrim$(fpcmbParameters.Text) = "All Businesses" Or QPTrim$(fpcmbParameters.Text) = "Spreadsheet Select" Then
    XFlag = False
  Else
    XFlag = True
    XDate = Date2Num(fptxtXDate.Text)
  End If
  
  XParFlag = 0
  
  If XFlag = True Then
    If InStr(fpcmbXPar.Text, "Only") Then
      XParFlag = 1
    Else
      XParFlag = 2
    End If
  End If
  
  NameFlag = False
  If fpcmbLabel.Text = "2) 1 X 3 1/2 1 Wide Text" Then
    LType = 1
  ElseIf fpcmbLabel.Text = "3) 1 X 3 1/2 3 Wide Text" Then
    LType = 2
  ElseIf fpcmbLabel.Text = "4) 1 X 3 1/2 4 Wide Text" Then
    LType = 3
  Else
    fpcmbLabel.BackColor = &HFFFF&
    frmBLMessageBoxJr.Label1.Caption = "Please make a Label Type selection."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fpcmbLabel.BackColor = &HFFFFFF
    fpcmbLabel.SetFocus
    Exit Sub
  End If
  
  OpenCustFile ARFile
  If fpcmbPrintOrder.Text = "Billing Name Order" Then
    NameFlag = True
    OpenCustNameIdxFile CustIdxHandle
    NumOfCustIdx = LOF(CustIdxHandle) / Len(CustSrchIdxRec)
    If NumOfCustIdx = 0 Then
      frmBLMessageBoxJr.Label1.Caption = "There are no business customers indexed."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      Close
      Exit Sub
    End If
    ReDim IdxRec(1 To NumOfCustIdx) As Integer
    For x = 1 To NumOfCustIdx
      Get CustIdxHandle, x, CustSrchIdxRec
      IdxRec(x) = CustSrchIdxRec.CustRec
    Next x
    Close CustIdxHandle
  ElseIf fpcmbPrintOrder.Text = "Account Number Order" Then
    OpenCustNumIdxFile CustIdxHandle
    NumOfCustIdx = LOF(CustIdxHandle) / Len(CustNumIdxRec)
    If NumOfCustIdx = 0 Then
      frmBLMessageBoxJr.Label1.Caption = "There are no business customers indexed."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      Close
      Exit Sub
    End If
    ReDim IdxRec(1 To NumOfCustIdx) As Integer
    For x = 1 To NumOfCustIdx
      Get CustIdxHandle, x, CustNumIdxRec
      IdxRec(x) = CustNumIdxRec.CustRec
    Next x
    Close CustIdxHandle
  Else
    fpcmbPrintOrder.BackColor = &HFFFF&
    frmBLMessageBoxJr.Label1.Caption = "Please select the Printing Order."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fpcmbPrintOrder.BackColor = &HFFFFFF
    fpcmbPrintOrder.SetFocus
    Close
    Exit Sub
  End If
  
  ReDim SpreadIdx(1 To 1) As Integer
  SpreadCnt = 0
  If fpcmbParameters.Text = "Spreadsheet Select" Then
    For x = 0 To NumOfCustIdx - 1
      fpList.Row = x
      If fpList.Selected = True Then
        fpList.ListIndex = x
        fpList.Col = 4
        If QPTrim$(fpList.ColText) = "Active" Then
          fpList.Col = 3
          SpreadCnt = SpreadCnt + 1
          ReDim Preserve SpreadIdx(1 To SpreadCnt) As Integer
          SpreadIdx(SpreadCnt) = CInt(fpList.ColText)
        End If
      End If
    Next x
  End If
  
  If fpcmbParameters.Text = "Spreadsheet Select" And SpreadCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No selections have been made from the spreadsheet. No mailing labels can be printed."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    Close
    fpList.SetFocus
    fpList.ListIndex = 0
    Exit Sub
  End If
  
  ReportFile$ = "ARLABEL.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  ReDim ToPrint(1 To 5, 1 To 5) As String
  
  If SpreadCnt > 0 Then
    NumOfCustIdx = SpreadCnt
    If SpreadCnt < 11 Then
      frmBLMessageBoxJr.Label1.Caption = "Selecting individual businesses for which to print mailing labels can waste mailing label paper if you have selected only a few businesses."
      frmBLMessageBoxJr.Label1.Top = 700
      frmBLMessageBoxJr.Show vbModal
    End If
  End If
  
    
  For cnt = 1 To NumOfCustIdx
    If SpreadCnt > 0 Then
      Get ARFile, SpreadIdx(cnt), CustRec
    Else
      Get ARFile, IdxRec(cnt), CustRec
    End If
    
    If UCase$(CustRec.Deleted) = "Y" Or QPTrim$(CustRec.SortName) = "DELETED" Then
      GoTo NextLabel
    End If
    If QPTrim$(CustRec.Inactive) = "Y" Then GoTo NextLabel
    If XFlag = True Then
      If XParFlag = 1 Then
        If CustRec.VALID <> XDate Then
          GoTo NextLabel
        End If
      ElseIf XParFlag = 2 Then
        If CustRec.VALID > XDate Then
          GoTo NextLabel
        End If
      End If
    End If
  
GoodCust:
    CustPCnt = CustPCnt + 1

    If Mid(CustRec.ZipCode, 7, 1) <> " " Then
      Zip$ = CustRec.ZipCode
      Zip$ = QPTrim$(Zip$)
    Else
      Zip$ = Left$(CustRec.ZipCode, 5)
      Zip$ = QPTrim$(Zip$)
    End If
    
    Select Case LType
    Case 1
      Print #RptHandle, "Cust #" + QPTrim$(CustRec.CustNumb)
      Print #RptHandle, Left$(QPTrim$(CustRec.BillName), 23)
      Print #RptHandle, Left$(QPTrim$(CustRec.ADDRESS1), 23)
      If Len(QPTrim$(CustRec.ADDRESS2)) > 0 Then
        Print #RptHandle, Left$(QPTrim$(CustRec.ADDRESS2), 23)
        Print #RptHandle, Left$(QPTrim$(CustRec.City), 18) + ", " + QPTrim$(CustRec.State) + " " + QPTrim(Zip$)
      Else
        Print #RptHandle, Left$(QPTrim$(CustRec.City), 18) + ", " + QPTrim$(CustRec.State) + " " + QPTrim$(Zip$)
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
      ToPrint(1, LabelCnt) = "Cust #" + QPTrim$(CustRec.CustNumb) 'Str$(AcctNumber&)
      ToPrint(2, LabelCnt) = Left(QPTrim$(CustRec.BillName), 23)
      ToPrint(3, LabelCnt) = Left(QPTrim$(CustRec.ADDRESS1), 23)
      
      If Len(QPTrim$(CustRec.ADDRESS2)) > 0 Then
          ToPrint(4, LabelCnt) = Left(QPTrim$(CustRec.ADDRESS2), 23)
          ToPrint(5, LabelCnt) = Left(QPTrim$(CustRec.City), 18) + ", " + QPTrim$(CustRec.State) + " " + QPTrim$(Zip$)
      Else
          ToPrint(4, LabelCnt) = Left(QPTrim$(CustRec.City), 18) + ", " + QPTrim$(CustRec.State) + " " + QPTrim$(Zip$)
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
      ToPrint(1, LabelCnt) = "Cust #" + QPTrim$(CustRec.CustNumb) 'Str$(AcctNumber&)
      ToPrint(2, LabelCnt) = Left(QPTrim$(CustRec.BillName), 23)
      ToPrint(3, LabelCnt) = Left(QPTrim$(CustRec.ADDRESS1), 23)
      
      If Len(QPTrim$(CustRec.ADDRESS2)) > 0 Then
        ToPrint(4, LabelCnt) = Left(QPTrim$(CustRec.ADDRESS2), 23)
        ToPrint(5, LabelCnt) = Left(QPTrim$(CustRec.City), 18) + ", " + QPTrim$(CustRec.State) + " " + QPTrim(Zip$)
      Else
        ToPrint(4, LabelCnt) = Left(QPTrim$(CustRec.City), 18) + ", " + QPTrim$(CustRec.State) + " " + QPTrim$(Zip$)
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
    frmBLMessageBoxJr.Label1.Caption = "There are no customers within the parameters entered on this screen. Mailing labels not printed."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  If NameFlag = True Then
    ViewPrint ReportFile$, "Mailing Labels Sorted by Search Name", True
  Else
    ViewPrint ReportFile$, "Mailing Labels Sorted by Customer Number", True
  End If
  
  KillFile ReportFile$

  MainLog ("Mailing labels processed in text format.")
  
  Exit Sub

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLMailLbls", "PrintText", Erl)
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
    Case vbKeyF1:
      SendKeys "%T"
      Call cmdHelp_Click
      KeyCode = 0
    Case vbKeyF4:
      SendKeys "%E"
      Call fpcmdXList_Click
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
      KillFile "mllbls.dat"
      ClearInUse PWcnt
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLMailLbls.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  Dim One As Integer
  Dim DHandle As Integer
  Dim x As Integer
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim CustCnt As Integer
  Dim CustIdxRec As CustNameIdxType
  Dim CustIdxHandle As Integer
  Dim CustIdxRecNum As Integer
  Dim Nextx As Integer
  
  On Error Resume Next
  lblBalloon.Visible = False
  
  OpenCustNameIdxFile CustIdxHandle
  CustIdxRecNum = LOF(CustIdxHandle) \ Len(CustIdxRec)
   
  ReDim CustIdx(1 To CustIdxRecNum) As Integer
  For x = 1 To CustIdxRecNum
    Get CustIdxHandle, x, CustIdxRec
    CustIdx(x) = CustIdxRec.CustRec 'load array with record pointers
  Next x
  Close CustIdxHandle
  
  OpenCustFile CHandle
  CustCnt = LOF(CHandle) / Len(CustRec)
  
  For x = 1 To CustIdxRecNum
    Get CHandle, CustIdx(x), CustRec
    If CustRec.Deleted <> "Y" And QPTrim$(CustRec.SortName) <> "DELETED" And QPTrim$(CustRec.Inactive) <> "Y" Then
      fpList.InsertRow = "  " & QPTrim$(CustRec.CustNumb) & Chr$(9) & "  " & QPTrim$(CustRec.BillName) & Chr$(9) & "  " & QPTrim(CustRec.City) & Chr$(9) & CStr(CustIdx(x)) & Chr$(9) & "Active"
    Else
      fpList.InsertRow = "  " & QPTrim$(CustRec.CustNumb) & Chr$(9) & "  " & QPTrim$(CustRec.BillName) & Chr$(9) & "  " & QPTrim(CustRec.City) & Chr$(9) & CStr(CustIdx(x)) & Chr$(9) & "Inactive"
    End If
  Next x
  
  
  fpList.ListIndex = 0
  
  Close
  fpList.Enabled = False
  cmdAlign.Enabled = False
  If Exist("issueappslics.dat") Then
    cmdExit.Tag = "Press 'Exit' to return to the main Applications menu."
  ElseIf Exist("custrptsmenu.dat") Then
    cmdExit.Tag = "Press 'Exit' to return to the main Customer Reports menu."
  Else
    cmdExit.Tag = "Press 'Exit' to return to the main Customer Maintenance menu."
  End If
'  fpcmbPrintOrder.ToolTipText = "Mailing labels can be printed in alphabetical order or in numerical order."
'  fpcmbLabel.ToolTipText = "Select the mailing label option from the drop down box."
'  fpcmbParameters.ToolTipText = "The mailing labels printed can be resticted to just those businesses who have the specified expiration date or they can be printed for all businesses."
'  fptxtXDate.ToolTipText = "If you select expiration date as the printing parameter then enter the license expiration date for which the mailing labels will be printed."
'  fpcmdXList.ToolTipText = "Press for a complete list of customers and their expiration dates."
'  cmdAlign.ToolTipText = "Prints a mailing label template used to help in positioning the mailing labels so they will print accurately."
'  cmdProcess.ToolTipText = "Press to begin the printing process for mailing labels."
'  cmdExit.ToolTipText = "Press to exit this screen."
  
  One = 1
  DHandle = FreeFile
  Open "mllbls.dat" For Output As DHandle Len = 2
  Print #DHandle, One
  Close DHandle
  fpcmbPrintOrder.Text = "Billing Name Order"
  fpcmbPrintOrder.AddItem "Billing Name Order"
  fpcmbPrintOrder.AddItem "Account Number Order"
  fpcmbLabel.Text = "1) 1 X 3 1/2 3 Wide Graphical"
  fpcmbLabel.AddItem "1) 1 X 3 1/2 3 Wide Graphical"
  fpcmbLabel.AddItem "2) 1 X 3 1/2 1 Wide Text"
  fpcmbLabel.AddItem "3) 1 X 3 1/2 3 Wide Text"
  fpcmbLabel.AddItem "4) 1 X 3 1/2 4 Wide Text"
  fpcmbParameters.Text = "All Businesses"
  fpcmbParameters.AddItem "All Businesses"
  fpcmbParameters.AddItem "Expiration Date"
  fpcmbParameters.AddItem "Spreadsheet Select"
  fpcmbXPar.Text = "Up To And Include This Expiration"
  fpcmbXPar.AddItem "Up To And Include This Expiration"
  fpcmbXPar.AddItem "This Expiration Only"
  fptxtXDate = Date
  fptxtXDate.Enabled = False
  fpcmdXList.Enabled = False
  fpcmbXPar.Enabled = False
End Sub

Private Sub fpcmbLabel_Change()
  If QPTrim$(fpcmbLabel.Text) = "" Then
    fpcmbLabel.Text = "1) 1 X 3 1/2 3 Wide Graphical"
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
      fpcmbParameters.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbParameters_Change()
  If QPTrim$(fpcmbParameters.Text) = "" Then
    fpcmbParameters.Text = "All Businesses"
  End If
  
  If QPTrim$(fpcmbParameters.Text) = "All Businesses" Then
    fptxtXDate.Enabled = False
    fpcmdXList.Enabled = False
    fpcmbXPar.Enabled = False
    fpList.Enabled = False
  ElseIf QPTrim$(fpcmbParameters.Text) = "Expiration Date" Then
    fptxtXDate.Enabled = True
    fpcmdXList.Enabled = True
    fpcmbXPar.Enabled = True
    fpList.Enabled = False
  ElseIf QPTrim$(fpcmbParameters.Text) = "Spreadsheet Select" Then
    fpList.Enabled = True
  End If
  
End Sub

Private Sub fpcmbParameters_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbParameters.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbParameters.ListIndex = -1
  End If
  If fpcmbParameters.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      If Mid(fpcmbParameters.Text, 1, 1) = "A" Then
        fpcmbPrintOrder.SetFocus
      Else
        fpcmbXPar.SetFocus
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

Private Sub fpcmbXPar_Change()
  If QPTrim$(fpcmbXPar.Text) = "" Then
    fpcmbXPar.Text = "Up To And Include This Expiration"
  End If
End Sub

Private Sub fpcmbXPar_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbXPar.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbXPar.ListIndex = -1
  End If
  If fpcmbXPar.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      If Mid(fpcmbXPar.Text, 1, 1) = "A" Then
        fpcmbPrintOrder.SetFocus
      Else
        fptxtXDate.SetFocus
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

Private Sub fpcmdXList_Click()
   frmBLXDateList.Show vbModal
End Sub

Private Sub PrintGraphics()
  Dim LType As Integer
  Dim ARFile As Integer
  Dim RptHandle As Integer
  Dim CustIdxHandle As Integer
  Dim CustNameIdxRec As CustNameIdxType
  Dim CustNumIdxRec As CustNumIdxType
  Dim CustSrchIdxRec As CustNameIdxType
  Dim NumOfCustIdx As Integer
  Dim x As Integer
  Dim CustRec As ARCustRecType
  Dim CustRCnt As Integer
  Dim Zip$
  Dim DidCnt As Integer
  Dim LabelCnt As Integer
  Dim PCnt As Integer
  Dim CustPCnt As Integer
  Dim AcctNumber&
  Dim cnt As Integer
  Dim ReportFile$
  Dim BusName As String * 23
  Dim CityName As String * 18
  Dim Address As String * 23
  Dim NameFlag As Boolean
  Dim dlm$
  Dim XFlag As Boolean
  Dim XDate As Integer
  Dim ValidCnt As Integer
  Dim XParFlag As Integer
  Dim SpreadCnt As Integer
  Dim NoDeleteCnt As Integer
  
  On Error GoTo ERRORSTUFF
  
  cmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  
  If QPTrim$(fpcmbParameters.Text) = "All Businesses" Or QPTrim$(fpcmbParameters.Text) = "Spreadsheet Select" Then
    XFlag = False
  Else
    XFlag = True
    XDate = Date2Num(fptxtXDate.Text)
  End If
  
  XParFlag = 0
  
  If XFlag = True Then
    If InStr(fpcmbXPar.Text, "Only") Then
      XParFlag = 1
    Else
      XParFlag = 2
    End If
  End If
  
  dlm = "~"
  NameFlag = False
  
  OpenCustFile ARFile
  If fpcmbPrintOrder.Text = "Billing Name Order" Then
    NameFlag = True
    OpenCustNameIdxFile CustIdxHandle
    NumOfCustIdx = LOF(CustIdxHandle) / Len(CustSrchIdxRec)
    If NumOfCustIdx = 0 Then
      frmBLMessageBoxJr.Label1.Caption = "There are no business customers indexed."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      Close
      Exit Sub
    End If
    ReDim IdxRec(1 To NumOfCustIdx) As Integer
    For x = 1 To NumOfCustIdx
      Get CustIdxHandle, x, CustSrchIdxRec
      IdxRec(x) = CustSrchIdxRec.CustRec
    Next x
    Close CustIdxHandle
  ElseIf fpcmbPrintOrder.Text = "Account Number Order" Then
    OpenCustNumIdxFile CustIdxHandle
    NumOfCustIdx = LOF(CustIdxHandle) / Len(CustNumIdxRec)
    If NumOfCustIdx = 0 Then
      frmBLMessageBoxJr.Label1.Caption = "There are no business customers indexed."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      Close
      Exit Sub
    End If
    ReDim IdxRec(1 To NumOfCustIdx) As Integer
    For x = 1 To NumOfCustIdx
      Get CustIdxHandle, x, CustNumIdxRec
      IdxRec(x) = CustNumIdxRec.CustRec
    Next x
    Close CustIdxHandle
  Else
    fpcmbPrintOrder.BackColor = &HFFFF&
    frmBLMessageBoxJr.Label1.Caption = "Please select the Printing Order."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fpcmbPrintOrder.BackColor = &HFFFFFF
    fpcmbPrintOrder.SetFocus
    Close
    Exit Sub
  End If
  
  ReDim SpreadIdx(1 To 1) As Integer
  SpreadCnt = 0
  If fpcmbParameters.Text = "Spreadsheet Select" Then
    For x = 0 To NumOfCustIdx - 1
      fpList.Row = x
      If fpList.Selected = True Then
        fpList.ListIndex = x
        fpList.Col = 4
        If QPTrim$(fpList.ColText) = "Active" Then
          fpList.Col = 3
          SpreadCnt = SpreadCnt + 1
          ReDim Preserve SpreadIdx(1 To SpreadCnt) As Integer
          SpreadIdx(SpreadCnt) = CInt(fpList.ColText)
        End If
      End If
    Next x
  End If
  
  If fpcmbParameters.Text = "Spreadsheet Select" And SpreadCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No selections have been made from the spreadsheet. No mailing labels can be printed."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    Close
    fpList.SetFocus
    fpList.ListIndex = 0
    Exit Sub
  End If
  
  ReportFile$ = "BLRPTS\ARLABEL.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  ReDim ToPrint(1 To 5, 1 To 5) As String
  
  If SpreadCnt > 0 Then
    NumOfCustIdx = SpreadCnt
    If SpreadCnt < 11 Then
      frmBLMessageBoxJr.Label1.Caption = "Selecting individual businesses for which to print mailing labels can waste mailing label paper if you have selected only a few businesses."
      frmBLMessageBoxJr.Label1.Top = 700
      frmBLMessageBoxJr.Show vbModal
    End If
  End If
  
  For cnt = 1 To NumOfCustIdx
    If SpreadCnt > 0 Then
      Get ARFile, SpreadIdx(cnt), CustRec
    Else
      Get ARFile, IdxRec(cnt), CustRec
    End If
    If QPTrim$(CustRec.Deleted) = "Y" Or QPTrim$(CustRec.SortName) = "DELETED" Then GoTo NextLabel
    If QPTrim$(CustRec.Inactive) = "Y" Then GoTo NextLabel
    If XFlag = True Then
      If XParFlag = 1 Then
        If CustRec.VALID <> XDate Then
          GoTo NextLabel
        End If
      ElseIf XParFlag = 2 Then
        If CustRec.VALID > XDate Then
          GoTo NextLabel
        End If
      End If
    End If
  
GoodCust:
    CustPCnt = CustPCnt + 1
    ValidCnt = ValidCnt + 1

    If Mid(CustRec.ZipCode, 7, 1) <> " " Then
      Zip$ = CustRec.ZipCode
      Zip$ = QPTrim$(Zip$)
    Else
      Zip$ = Left$(CustRec.ZipCode, 5)
      Zip$ = QPTrim$(Zip$)
    End If
    LabelCnt = LabelCnt + 1 'this requires a line to be printed
    'in columns of 3 with each column containing data gathered
    'from different customers...it also must limit the size of some
    'variables to accommodate the limitations of a mailing label's
    'size
    ToPrint(1, LabelCnt) = QPTrim$(CustRec.BillName) 'Str$(AcctNumber&)
    ToPrint(2, LabelCnt) = Left(QPTrim$(CustRec.ADDRESS1), 23)
    ToPrint(3, LabelCnt) = Left(QPTrim$(CustRec.ADDRESS2), 23)
    ToPrint(4, LabelCnt) = Left(QPTrim$(CustRec.City), 18) + ", " + QPTrim$(CustRec.State) + " " + QPTrim$(Zip$)
    
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
    frmBLMessageBoxJr.Label1.Caption = "There are no customers within the parameters entered on this screen. Mailing labels not printed."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  arBLMailLabels.Show
  frmBLLoadReport.Show
  
  MainLog ("Mailing labels processed in graphics format.")
  
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLMailLbls", "PrintGraphics", Erl)
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

Private Sub fptxtXDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fpcmbPrintOrder.SetFocus
  End If
End Sub

