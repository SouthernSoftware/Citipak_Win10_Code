VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmFAItemsAddDelOptRpt 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fixed Assets New or Deleted Item Status Report"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmFAItemsAddDelOptRpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   7068
      Left            =   1968
      TabIndex        =   6
      Top             =   912
      Width           =   7740
      _Version        =   196609
      _ExtentX        =   13652
      _ExtentY        =   12467
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmFAItemsAddDelOptRpt.frx":08CA
      Begin LpLib.fpCombo fpcmbOrder 
         Height          =   405
         Left            =   3210
         TabIndex        =   0
         ToolTipText     =   "Select the order this report will display data."
         Top             =   1485
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
         ColDesigner     =   "frmFAItemsAddDelOptRpt.frx":08E6
      End
      Begin LpLib.fpCombo fpcomboPrintOpt 
         Height          =   405
         Left            =   3510
         TabIndex        =   5
         ToolTipText     =   "Select Graphical for a robust but slower processing report. Select Text for a quicker report."
         Top             =   5280
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
         ColDesigner     =   "frmFAItemsAddDelOptRpt.frx":0BDD
      End
      Begin LpLib.fpCombo fpcmbRptType 
         Height          =   405
         Left            =   3210
         TabIndex        =   1
         ToolTipText     =   "Select the items this report will display."
         Top             =   2070
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
         ColDesigner     =   "frmFAItemsAddDelOptRpt.frx":0ED4
      End
      Begin EditLib.fpText fptxtDeptNum 
         Height          =   396
         Left            =   3072
         TabIndex        =   2
         ToolTipText     =   "If DEPARTMENT is selected in the Report Order field then enter the desired department number on which to report."
         Top             =   2640
         Width           =   1068
         _Version        =   196608
         _ExtentX        =   1884
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
      Begin EditLib.fpDateTime fpDateStart 
         Height          =   444
         Left            =   3744
         TabIndex        =   3
         ToolTipText     =   "Enter the beginning date of this report."
         Top             =   3792
         Width           =   1788
         _Version        =   196608
         _ExtentX        =   3154
         _ExtentY        =   783
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
         Text            =   "1/24/2003"
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
      Begin EditLib.fpDateTime fpDateEnd 
         Height          =   444
         Left            =   3744
         TabIndex        =   4
         ToolTipText     =   "Enter the ending date of this report."
         Top             =   4464
         Width           =   1788
         _Version        =   196608
         _ExtentX        =   3154
         _ExtentY        =   783
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
         Text            =   "1/24/2003"
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
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   690
         Left            =   1590
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Click this button to create the desired report."
         Top             =   6000
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
         ButtonDesigner  =   "frmFAItemsAddDelOptRpt.frx":11CB
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   690
         Left            =   4464
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Click this button to create the desired report."
         Top             =   6000
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
         ButtonDesigner  =   "frmFAItemsAddDelOptRpt.frx":13A7
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdDept 
         Height          =   390
         Left            =   4410
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Click this button to bring up a list of all current department numbers."
         Top             =   2640
         Width           =   1365
         _Version        =   131072
         _ExtentX        =   2408
         _ExtentY        =   688
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
         ButtonDesigner  =   "frmFAItemsAddDelOptRpt.frx":1586
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date:"
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
         Left            =   2256
         TabIndex        =   14
         Top             =   3912
         Width           =   1260
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "End Date:"
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
         Left            =   2304
         TabIndex        =   13
         Top             =   4584
         Width           =   1212
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   300
         Left            =   1440
         TabIndex        =   12
         Top             =   2160
         Width           =   1548
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
         Left            =   1824
         TabIndex        =   11
         Top             =   5364
         Width           =   1500
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Items Added Or Disposed Of Report"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   1350
         TabIndex        =   10
         Top             =   630
         Width           =   5190
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   684
         Left            =   1248
         Top             =   480
         Width           =   5388
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Reporting Period:"
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
         Left            =   3120
         TabIndex        =   9
         Top             =   3264
         Width           =   1884
      End
      Begin VB.Label Label8 
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
         Left            =   1968
         TabIndex        =   8
         Top             =   2736
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
         Left            =   1440
         TabIndex        =   7
         Top             =   1584
         Width           =   1548
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   7260
      Left            =   1860
      Top             =   804
      Width           =   7932
   End
End
Attribute VB_Name = "frmFAItemsAddDelOptRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsFATextBoxOverRider
  Private Temp_Class As Resize_Class

Private Sub cmdDept_Click()
  frmFADeptList.Show vbModal

End Sub

Private Sub cmdExit_Click()
  frmFAReportMenu.Show
  Close
  DoEvents
  On Error Resume Next
  KillFile "disposedofrpt.dat"
  KillFile "newadddelrptopen.dat"
  Unload frmFAItemsAddDelOptRpt

End Sub

Private Sub cmdProcess_Click()
  'this sub directs the printing choice the user makes
  'to the several possible combinations
  If fpcomboPrintOpt.Text = "Graphical" Then
    If QPTrim$(fpcmbRptType.Text) = "BOTH" Then
      Call PrintBothGraphics
    Else
      Call PrintGraphics
    End If
    Exit Sub
  ElseIf fpcomboPrintOpt.Text = "Text" Then
    MsgBox "Pitch 17 is recommended for this report."
    If QPTrim$(fpcmbRptType.Text) = "BOTH" Then
      Call PrintBothText
    Else
      Call PrintText
    End If
  Else
    Exit Sub
  End If
  
End Sub
Private Sub PrintGraphics()
  Dim DOrigCost#(2), DBookTotal#(2), DCDep#(2), DYDep#(2), OrigCost#(2), BookTotal#(2), CDep#(2), YDep#(2)
  Dim ReportFile$
  Dim Dept$
  Dim BDate As Integer
  Dim EDate As Integer
  Dim RptHandle As Integer
  Dim FAHandle As Integer
  Dim NumOfFARecs As Integer
  Dim FAItemRec As FAItemRecType
  Dim cnt&
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim DisposeDate As Integer
  Dim AcquireDate As Integer
  Dim DeptNumber As Integer
  Dim YTDDep#
  Dim x As Integer
  Dim Nextx As Integer
  Dim DIdxRec As DeptNumbSortIdxType
  Dim DIdxHandle As Integer
  Dim DIdxCnt As Integer
  Dim Index$
  Dim DataFlag As Boolean
  Dim TagFlag As Boolean
  Dim TagReportFile$
  Dim TagHandle As Integer
  Dim dlm$
  Dim TagSubHandle As Integer
  Dim TagSubReportFile$
  Dim FASetUpRec As FASetupRecType
  Dim Employer$
  Dim YDispPrice#(2), DYDispPrice#(2)
  Dim DispPrice#
  Dim DCnt As Integer
  Dim Method$
  Dim ActiveX As Long
  Dim HoldRec As Long
  Dim HoldDate As Integer
  Dim ThisDRec As Long
  Dim BigNum As Long
  Dim SmallNum As Long
  Dim HoldTag$
  Dim RptType As String * 1
  Dim DeptDescr As String
  Dim ItemTotal As Long
  
  On Error GoTo ERRORSTUFF
  
  If Check4ValidDept = False Then Exit Sub
  
  If fpcmbRptType.Text = "DISPOSED OF ONLY" Then
    RptType = "D" 'disposed
  Else
    RptType = "A" 'acquired
  End If
  
  OpenFASetUpFile FAHandle
  Get FAHandle, 1, FASetUpRec
  Close FAHandle
  Employer = FASetUpRec.TownName
  
  dlm$ = "~"
  ReportFile$ = "FARPTS\FAADDDELOPT.RPT"  'Report File Name
  TagReportFile$ = "FARPTS\FAADDDELOPTTAG.RPT"
  TagSubReportFile$ = "FARPTS\FAADDDELOPTSUB.RPT"
  Dept$ = QPTrim$(fptxtDeptNum.Text)
  BDate = Date2Num(fpDateStart.Text) 'beginning date
  EDate = Date2Num(fpDateEnd.Text) 'ending date
  
  Index$ = QPTrim$(fpcmbOrder.Text)
  If QPTrim$(Index$) = "DEPARTMENT NUMBER" Then
    RptHandle = FreeFile 'RptHandle is used
    'to print out department data
    Open ReportFile$ For Output As #RptHandle
  Else
    TagHandle = FreeFile
    Open TagReportFile$ For Output As #TagHandle
    'TagHandle is used to print out tags only
    TagSubHandle = FreeFile
    Open TagSubReportFile$ For Output As #TagSubHandle
    'TagSubHandle is used to print out tag
    'totals for an AR sub report
  End If
  
  OpenTagIdxFile TagIdxHandle
  NumOfFARecs = LOF(TagIdxHandle) \ Len(TagIdx)
  If NumOfFARecs = 0 Then
    MsgBox "No item records on file."
    Close TagIdxHandle
    Close
    Exit Sub
  End If
  
  ReDim TagIdxRecs(1 To NumOfFARecs)
  For x = 1 To NumOfFARecs
    Get TagIdxHandle, x, TagIdx
    TagIdxRecs(x) = TagIdx.DataRecNum 'load array with
    'asset records in asset number numerical order
  Next x
  Close TagIdxHandle
  
  OpenDeptIdxFile DIdxHandle
  DIdxCnt = LOF(DIdxHandle) \ Len(DIdxRec)
  ReDim DeptArr(1 To DIdxCnt) As String
  ReDim DeptDesc(1 To DIdxCnt) As String
  
  For x = 1 To DIdxCnt
    Get DIdxHandle, x, DIdxRec 'load arrays with
    'department data
    DeptArr(x) = QPTrim$(DIdxRec.DeptNumb)
    DeptDesc(x) = QPTrim$(DIdxRec.DeptIdxDesc)
  Next x
  Close DIdxHandle
  
  'create arrays for collecting totals
  ReDim TagDOrigCost(1 To DIdxCnt) As Double
  ReDim TagDBookTotal(1 To DIdxCnt) As Double
  ReDim TagDYDep(1 To DIdxCnt) As Double
  ReDim TagDYDispPrice(1 To DIdxCnt) As Double
  ReDim TagDCnt(1 To DIdxCnt) As Integer
  ReDim DsplAmt(1 To DIdxCnt) As Double
  
  If Dept$ <> "ALL" Then 'user wants one dept only
    DeptNumber = Val(QPTrim$(fptxtDeptNum.Text))
    For x = 1 To DIdxCnt
      If DeptNumber = DeptArr(x) Then 'get dept description
        DeptDescr = DeptDesc(x)
        Exit For
      End If
    Next x
  Else
    DeptNumber = Val(QPTrim(DeptArr(1))) 'otherwise just start
    'with the first department when the user wants all departments
    'data displayed
    DeptDescr = DeptDesc(1)
  End If
  
  OpenFAItemFile FAHandle
  
  TagFlag = False
  
  'start date sort
  ReDim ItemsForReport(1 To 1) As Long
  ReDim DatesForReport(1 To 1) As Integer
  For x = 1 To NumOfFARecs
    Get FAHandle, TagIdxRecs(x), FAItemRec
    If RptType = "D" Then 'disposed of only
      If FAItemRec.DispDate >= BDate And FAItemRec.DispDate <= EDate And FAItemRec.DsplFlag > 0 Then
      'this item falls within the beginning and ending dates and is not disposed of
        ActiveX = ActiveX + 1 'count the valid items
        ReDim Preserve ItemsForReport(1 To ActiveX) As Long
        ReDim Preserve DatesForReport(1 To ActiveX) As Integer
        ItemsForReport(ActiveX) = TagIdxRecs(x) 'load array with record numbers
        DatesForReport(ActiveX) = FAItemRec.DispDate 'load array with disposal dates
      End If
    Else 'user wants acquired only
      If FAItemRec.AQURDATE >= BDate And FAItemRec.AQURDATE <= EDate Then
      'this item falls within the beginning and ending dates and it makes no difference
      'if the item has been disposed of or not
        ActiveX = ActiveX + 1 'count all valid items
        ReDim Preserve ItemsForReport(1 To ActiveX) As Long
        ReDim Preserve DatesForReport(1 To ActiveX) As Integer
        ItemsForReport(ActiveX) = TagIdxRecs(x) 'load array with item record numbers
        DatesForReport(ActiveX) = FAItemRec.AQURDATE 'load array with acquired dates
      End If
    End If
  Next x
  
  If ActiveX = 0 Then
    If RptType = "D" Then
      MsgBox "No fixed assets could be found that were disposed of within the time period entered."
    Else
      MsgBox "No fixed assets could be found that were acquired within the time period entered."
    End If
    fpDateStart.SetFocus
    Close
    Exit Sub
  End If
  
  'sort items by dates saved in array (disposal or acquire)
  BigNum = 0
  For x = 1 To ActiveX
    If DatesForReport(x) > BigNum Then
      BigNum = DatesForReport(x)
    End If
  Next x
  
  Nextx = 1
  SmallNum = BigNum + 1
  Do
    For x = Nextx To ActiveX
      If DatesForReport(x) < SmallNum Then
        SmallNum = DatesForReport(x)
        ThisDRec = x
      End If
    Next x
    HoldRec = ItemsForReport(ThisDRec)
    HoldDate = DatesForReport(ThisDRec)
    ItemsForReport(ThisDRec) = ItemsForReport(Nextx)
    DatesForReport(ThisDRec) = DatesForReport(Nextx)
    ItemsForReport(Nextx) = HoldRec
    DatesForReport(Nextx) = HoldDate
    If Nextx = ActiveX Then Exit Do
    Nextx = Nextx + 1
    SmallNum = BigNum + 1
  Loop
  'end date sort
  
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
  If TagFlag = True Then
    Index = "DEPARTMENT NUMBERS"
  End If
  Nextx = 1
  Do 'this loop iterates once if there is 1 dept requested, DIdxCnt + 1 if
  'department numbers "ALL" is requested and DIdxCnt + 2 for Tag Numbers
  '...Tag Numbers requires one iteration ignoring departments to get a list of
  'all valid tags in numeric order then DIdxCnt + 1 skipping the itemized tag
  'data print out just to allow the dept totals to assimilate...
    DataFlag = False
    For cnt& = 1 To ActiveX
      Get FAHandle, ItemsForReport(cnt), FAItemRec 'items are retrieved by
      'last date first
      
      If RptType = "D" Then
        If FAItemRec.DsplFlag = 2 Then
          If QPTrim$(FAItemRec.DsplMethod) = "Salvage" Then
            Method$ = "SLV"
          ElseIf QPTrim$(FAItemRec.DsplMethod) = "Sold" Then
            Method$ = "SLD"
          ElseIf QPTrim$(FAItemRec.DsplMethod) = "Auction" Then
            Method$ = "AUC"
          Else
            Method$ = "OTH"
          End If
        ElseIf FAItemRec.DsplFlag = 1 Then
          Method$ = "PND" 'pending
        End If
        DispPrice# = FAItemRec.DisposAmt
      End If
      
      If FAItemRec.ILIFE > 0 Then 'code carried over from dos
        YTDDep# = FAItemRec.DEP2DATE
      Else
        YTDDep# = 0
      End If
      
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
      
      If RptType = "D" Then 'Disposed of
        If QPTrim$(Index) = "TAG NUMBER" Then 'tag only
          '                   0               1                 2
          Print #TagHandle, DeptDescr; dlm; RptType; dlm; Employer$; dlm;
          '                          3                          4
          Print #TagHandle, FAItemRec.ItemTag; dlm; Left$(FAItemRec.IDESC1, 28); dlm;
          '                        5                     6                       7
          Print #TagHandle, FAItemRec.IDEPT; dlm; Method$; dlm; FAItemRec.ORGCOST; dlm;
          '                         8
          Print #TagHandle, FAItemRec.DEP2DATE; dlm;
          '                                    9                      10                              11
          Print #TagHandle, FAItemRec.CURRVAL; dlm; MakeRegDate(FAItemRec.DispDate); dlm; FAItemRec.DisposAmt; dlm;
          '                         12                      13                        14                    15
          Print #TagHandle, MakeRegDate(BDate); dlm; MakeRegDate(EDate); dlm; FAItemRec.LifeLeft; dlm; FAItemRec.ILIFE
        Else 'dept only
          '                   0               1                2
          Print #RptHandle, DeptDescr; dlm; RptType; dlm; Employer$; dlm;
          '                         3                             4
          Print #RptHandle, FAItemRec.ItemTag; dlm; Left$(FAItemRec.IDESC1, 28); dlm;
          '                        5                    6                       7
          Print #RptHandle, FAItemRec.IDEPT; dlm; Method; dlm; FAItemRec.ORGCOST; dlm;
            '                    8
          Print #RptHandle, FAItemRec.DEP2DATE; dlm;
          '                                   9                                       10
          Print #RptHandle, FAItemRec.CURRVAL; dlm; MakeRegDate(FAItemRec.DispDate); dlm;
        End If
      Else ' Acquired
        If QPTrim$(Index) = "TAG NUMBER" Then 'tag only
          '                   0            1              2
          Print #TagHandle, DeptDescr; dlm; RptType; dlm; Employer$; dlm;
          '                          3                          4
          Print #TagHandle, FAItemRec.ItemTag; dlm; Left$(FAItemRec.IDESC1, 28); dlm;
          '                        5                 6                       7
          Print #TagHandle, FAItemRec.IDEPT; dlm; Method$; dlm; FAItemRec.ORGCOST; dlm;
          '                         8
          Print #TagHandle, FAItemRec.DEP2DATE; dlm;
          '                              9                              10                11
          Print #TagHandle, FAItemRec.CURRVAL; dlm; MakeRegDate(FAItemRec.AQURDATE); dlm; ""; dlm;
          '                         12                      13                         14                     15
          Print #TagHandle, MakeRegDate(BDate); dlm; MakeRegDate(EDate); dlm; FAItemRec.LifeLeft; dlm; FAItemRec.ILIFE
        Else 'dept only
          '                   0               1                2
          Print #RptHandle, DeptDescr; dlm; RptType; dlm; Employer$; dlm;
          '                         3                             4
          Print #RptHandle, FAItemRec.ItemTag; dlm; Left$(FAItemRec.IDESC1, 28); dlm;
          '                        5                    6                       7
          Print #RptHandle, FAItemRec.IDEPT; dlm; Method; dlm; FAItemRec.ORGCOST; dlm;
            '                    8
          Print #RptHandle, FAItemRec.DEP2DATE; dlm;
          '                         9                                   10                11
          Print #RptHandle, FAItemRec.CURRVAL; dlm; MakeRegDate(FAItemRec.AQURDATE); dlm; ""; dlm;
        End If
      End If
      ItemTotal = ItemTotal + 1
TagOnly2: 'collects data for each department for reporting totals

      'This if statement filters out the first iteration of TAG NUMBER
      'selection because we do not want to start accumulating dept data
      'until the second iteration
      If TagFlag = False And QPTrim$(Index) = "TAG NUMBER" Then GoTo SkipEm1
      DCnt = DCnt + 1
      'collects grand totals
      OrigCost#(1) = OrigCost#(1) + FAItemRec.ORGCOST
      BookTotal#(1) = BookTotal#(1) + (FAItemRec.CURRVAL)
      YDep#(1) = YDep#(1) + YTDDep#
      YDispPrice#(1) = YDispPrice#(1) + DYDispPrice#(1)
      DOrigCost#(1) = DOrigCost#(1) + FAItemRec.ORGCOST
      TagDOrigCost(Nextx) = DOrigCost#(1)
      DBookTotal#(1) = DBookTotal#(1) + (FAItemRec.CURRVAL)
      TagDBookTotal(Nextx) = DBookTotal#(1)
      DYDep#(1) = DYDep#(1) + YTDDep#
      TagDYDep(Nextx) = DYDep#(1)
      DYDispPrice#(1) = DYDispPrice#(1) + DispPrice#
      TagDYDispPrice(Nextx) = TagDYDispPrice(Nextx) + FAItemRec.DisposAmt
      TagDCnt(Nextx) = DCnt
      If QPTrim$(Index$) = "DEPARTMENT NUMBER" Then
        If RptType = "D" Then 'only dept only data is printed here
          '                     11               12                 13
          Print #RptHandle, DeptNumber; dlm; DOrigCost#(1); dlm; DYDep#(1); dlm;
          '                     14                   15                16              17
          Print #RptHandle, DBookTotal#(1); dlm; OrigCost#(1); dlm; YDep#(1); dlm; BookTotal#(1); dlm;
          '                     18                    19
          Print #RptHandle, DYDispPrice#(1); dlm; DispPrice#; dlm;
          '                         20                     21                        22                      23
          Print #RptHandle, MakeRegDate(BDate); dlm; MakeRegDate(EDate); dlm; FAItemRec.LifeLeft; dlm; FAItemRec.ILIFE
        Else
          '                     11               12                 13
          Print #RptHandle, DeptNumber; dlm; DOrigCost#(1); dlm; DYDep#(1); dlm;
          '                     14                   15                16              17
          Print #RptHandle, DBookTotal#(1); dlm; OrigCost#(1); dlm; YDep#(1); dlm; BookTotal#(1); dlm;
          '                 18       19
          Print #RptHandle, ""; dlm; "";
          '                         20                      21                     22                         23
          Print #RptHandle, MakeRegDate(BDate); dlm; MakeRegDate(EDate); dlm; FAItemRec.LifeLeft; dlm; FAItemRec.ILIFE
        End If
      End If
      
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
    DeptNumber = Val(QPTrim$(DeptArr(Nextx)))
    DeptDescr = DeptDesc(Nextx)
    'clear all dept totals
    DOrigCost#(1) = 0 'reset for next dept
    DBookTotal#(1) = 0 'reset for next dept
    DYDep#(1) = 0 'reset for next dept
    DOrigCost#(2) = 0 'reset for next dept
    DBookTotal#(2) = 0 'reset for next dept
    DYDep#(2) = 0 'reset for next dept
    DCnt = 0 'reset for next dept
  Loop
  
  Unload frmFAShowPctComp
  frmFAShowPctComp.Out = False
  EnableCloseButton Me.hwnd, True
  Me.cmdExit.Enabled = True
  Me.cmdProcess.Enabled = True
  
  'only prints if TAG NUMBERS was selected
  If ItemTotal = 0 Then
    MsgBox "There are no fixed assets saved for this criteria."
    Close
    Exit Sub
  End If
  
  If TagFlag = True Then
    GoSub PrintTagDeptTotals
  End If
  
  Close         'Close all open files now
  Close RptHandle
  If TagFlag = False Then
    arFADisposedOfRpt.Show 'also includes acquired assets
  Else
    arFADisposedOFTagOnly.Show 'also includes acquired assets
  End If
  
  frmFALoadReport.Show
  
  Exit Sub
  
  
PrintTagDeptTotals: 'print only if TAG NUMBERS was selected
  If RptType = "D" Then
    For x = 1 To DIdxCnt
      If QPTrim$(DeptArr(x)) = "" Then DeptArr(x) = "0"
      '                        0                    1                   2                    3                          4
      Print #TagSubHandle, DeptArr(x); dlm; TagDOrigCost(x); dlm; TagDYDep(x); dlm; TagDBookTotal(x); dlm; TagDYDispPrice(x); dlm; TagDCnt(x); dlm; "D"
    Next x
  Else
    For x = 1 To DIdxCnt
      If QPTrim$(DeptArr(x)) = "" Then DeptArr(x) = "0"
      '                        0                    1                   2                    3                          4
      Print #TagSubHandle, DeptArr(x); dlm; TagDOrigCost(x); dlm; TagDYDep(x); dlm; TagDBookTotal(x); dlm; ""; dlm; TagDCnt(x); dlm; "A"
    Next x
  End If
  
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAItemsAddDelOptRpt", "PrintGraphics", Erl)
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
Private Sub PrintText()
  Dim DOrigCost#(2), DBookTotal#(2), DCDep#(2), DYDep#(2), OrigCost#(2), BookTotal#(2), CDep#(2), YDep#(2)
  Dim TDsplAmt As Double
  Dim ReportFile$
  Dim Dash80$
  Dim FF$
  Dim MaxLines As Integer
  Dim ItemCnt&
  Dim LineCnt&
  Dim Dept$
  Dim BDate As Integer
  Dim EDate As Integer
  Dim RptHandle As Integer
  Dim FAHandle As Integer
  Dim NumOfFARecs As Integer
  Dim FAItemRec As FAItemRecType
  Dim cnt&
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim DisposeDate As Integer
  Dim AcquireDate As Integer
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
  Dim DsplAmt As Double
  Dim Method$
  Dim ActiveX As Long
  Dim HoldRec As Long
  Dim HoldDate As Integer
  Dim ThisDRec As Long
  Dim BigNum As Long
  Dim SmallNum As Long
  Dim HoldTag$
  Dim RptType As String * 1
  Dim DeptDescr$
  Dim HeaderFlag As Boolean
  Dim DItemCnt&
  Dim FirstFlag As Boolean
  Dim ItemTotal As Long
  
  On Error GoTo ERRORSTUFF
  'please refer to PrintGraphics for comments as
  'it is almost identical to this sub
  
  FirstFlag = True
  If QPTrim$(fpcmbOrder.Text) = "TAG NUMBER" Then
    HeaderFlag = False
  Else
    HeaderFlag = True
  End If
  
  If Check4ValidDept = False Then Exit Sub
  
  If fpcmbRptType.Text = "DISPOSED OF ONLY" Then
    RptType = "D"
  Else
    RptType = "A"
  End If
  
  ReportFile$ = "FAMaster.PRN"  'Report File Name
  Dash80$ = String$(80, "=")
  FF$ = Chr$(12)
  
  MaxLines = 58
  LineCnt& = 0
  ItemCnt& = 0
  DItemCnt& = 0
  Dept$ = QPTrim$(fptxtDeptNum.Text)
  BDate = Date2Num(fpDateStart.Text)
  EDate = Date2Num(fpDateEnd.Text)
  
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
  ReDim DeptArr(1 To DIdxCnt) As String
  ReDim DeptDesc(1 To DIdxCnt) As String
  
  For x = 1 To DIdxCnt
    Get DIdxHandle, x, DIdxRec
    DeptArr(x) = QPTrim$(DIdxRec.DeptNumb)
    DeptDesc(x) = QPTrim$(DIdxRec.DeptIdxDesc)
  Next x
  Close DIdxHandle
  
  ReDim TagDOrigCost(1 To DIdxCnt) As Double
  ReDim TagDBookTotal(1 To DIdxCnt) As Double
  ReDim TagDYDep(1 To DIdxCnt) As Double
  ReDim DDsplAmt(1 To DIdxCnt) As Double
  ReDim DptCnt(1 To DIdxCnt) As Long
  
  If Dept$ <> "ALL" Then
    DeptNumber = Val(QPTrim$(fptxtDeptNum.Text))
    For x = 1 To DIdxCnt
      If DeptNumber = Val(QPTrim$(DeptArr(x))) Then
        DeptDescr = QPTrim(DeptDesc(x))
      End If
    Next x
  Else
    DeptNumber = Val(QPTrim(DeptArr(1)))
    DeptDescr = QPTrim$(DeptDesc(1))
  End If
  
  GoSub PrintMasterHeader1
  
  OpenFAItemFile FAHandle
  
  TagFlag = False
  
  'start date sort
  ReDim ItemsForReport(1 To 1) As Long
  ReDim DatesForReport(1 To 1) As Integer
  For x = 1 To NumOfFARecs
    Get FAHandle, TagIdxRecs(x), FAItemRec
    If RptType = "D" Then
      If FAItemRec.DispDate >= BDate And FAItemRec.DispDate <= EDate And FAItemRec.DsplFlag > 0 Then
        ActiveX = ActiveX + 1
        ReDim Preserve ItemsForReport(1 To ActiveX) As Long
        ReDim Preserve DatesForReport(1 To ActiveX) As Integer
        ItemsForReport(ActiveX) = TagIdxRecs(x)
        DatesForReport(ActiveX) = FAItemRec.DispDate
      End If
    ElseIf RptType = "A" Then
      If FAItemRec.AQURDATE >= BDate And FAItemRec.AQURDATE <= EDate Then
        ActiveX = ActiveX + 1
        ReDim Preserve ItemsForReport(1 To ActiveX) As Long
        ReDim Preserve DatesForReport(1 To ActiveX) As Integer
        ItemsForReport(ActiveX) = TagIdxRecs(x)
        DatesForReport(ActiveX) = FAItemRec.AQURDATE
      End If
    End If
  Next x
  
  If ActiveX = 0 Then
    If RptType = "D" Then
      MsgBox "No fixed assets could be found that were disposed of within the reporting period entered."
    Else
      MsgBox "No fixed assets could be found that were acquired within the reporting period entered."
    End If
    fpDateStart.SetFocus
    Close
    Exit Sub
  End If
  
  BigNum = 0 'begin sorting dates
  For x = 1 To ActiveX
    If DatesForReport(x) > BigNum Then
      BigNum = DatesForReport(x)
    End If
  Next x
  
  Nextx = 1
  SmallNum = BigNum + 1
  
  Do
    For x = Nextx To ActiveX ' - 1
      If DatesForReport(x) < SmallNum Then
        SmallNum = DatesForReport(x)
        ThisDRec = x
      End If
    Next x
    HoldRec = ItemsForReport(ThisDRec)
    HoldDate = DatesForReport(ThisDRec)
    ItemsForReport(ThisDRec) = ItemsForReport(Nextx)
    DatesForReport(ThisDRec) = DatesForReport(Nextx)
    ItemsForReport(Nextx) = HoldRec
    DatesForReport(Nextx) = HoldDate
    If Nextx = ActiveX Then Exit Do
    Nextx = Nextx + 1
    SmallNum = BigNum + 1
  Loop
  'end date sort
  
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
    For cnt& = 1 To ActiveX ' - 1   'NumOfFARecs
      Get FAHandle, ItemsForReport(cnt), FAItemRec
      If LineCnt& >= MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintMasterHeader1
      End If
      'Check For Disposed Date
      If RptType = "D" Then
        DisposeDate = FAItemRec.DispDate
        If DisposeDate < BDate Or DisposeDate > EDate Then
        'filter out items that don't fall inside the date parameters
          GoTo SkipEm1
        End If
        If FAItemRec.DsplFlag = 0 Then GoTo SkipEm1
      'Check for Acquired Date
      ElseIf RptType = "A" Then
        AcquireDate = FAItemRec.AQURDATE
        If AcquireDate < BDate Or AcquireDate > EDate Then
          GoTo SkipEm1
        End If
      End If
      'Can't depreciate something that has no asset life...this should
      'not be necessary unless the asset life was changed to zero after
      'depreciation had already started
      If FAItemRec.ILIFE > 0 Then
        YTDDep# = FAItemRec.DEP2DATE
      Else
        YTDDep# = 0
      End If
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
      If DItemCnt = 0 And QPTrim$(fpcmbOrder.Text) = "DEPARTMENT NUMBER" Then
'        Print #RptHandle, String$(113, "-")
        Print #RptHandle, "Department: "; DeptNumber; " "; DeptDescr
        If RptType = "D" Then
          Print #RptHandle, String$(128, "-")
        Else
          Print #RptHandle, String$(113, "-")
        End If
        LineCnt = LineCnt + 2
      End If
      DataFlag = True
      Print #RptHandle, FAItemRec.ItemTag; Tab(21); Left$(FAItemRec.IDESC1, 28);
      Print #RptHandle, Tab(50); FAItemRec.IDEPT;
      If RptType = "D" Then
        If FAItemRec.DsplFlag = 2 Then
          If QPTrim$(FAItemRec.DsplMethod) = "Salvage" Then
            Method$ = "SLV"
          ElseIf QPTrim$(FAItemRec.DsplMethod) = "Sold" Then
            Method$ = "SLD"
          ElseIf QPTrim$(FAItemRec.DsplMethod) = "Auction" Then
            Method$ = "AUC"
          Else
            Method$ = "OTH"
          End If
        ElseIf FAItemRec.DsplFlag = 1 Then
          Method$ = "PND"
        End If
        Print #RptHandle, Tab(57); Method;
      End If
      If RptType = "D" Then
        Print #RptHandle, Tab(65); Using("###,###,##0.00", CStr(FAItemRec.ORGCOST));
        Print #RptHandle, Tab(80); Using("###,###,##0.00", CStr(FAItemRec.CURRVAL));
        Print #RptHandle, Tab(97); Using("###,###,##0.00", CStr(FAItemRec.DisposAmt));
        Print #RptHandle, Tab(119); MakeRegDate(FAItemRec.DispDate)
      Else
        Print #RptHandle, Tab(57); Using$("#0", FAItemRec.ILIFE); "/"; Using$("#0", FAItemRec.LifeLeft);
        Print #RptHandle, Tab(65); Using("###,###,##0.00", CStr(FAItemRec.ORGCOST));
        Print #RptHandle, Tab(80); Using("###,###,##0.00", CStr(FAItemRec.CURRVAL));
        Print #RptHandle, Tab(104); MakeRegDate(FAItemRec.AQURDATE)
      End If
      LineCnt& = LineCnt& + 1
      ItemCnt& = ItemCnt& + 1
      ItemTotal = ItemTotal + 1
      DItemCnt& = DItemCnt& + 1
TagOnly2: 'collects data for each department for reporting totals

      'This if statement filters out the first iteration of TAG NUMBER
      'selection because we do not want to start accumulating dept data
      'until the second iteration
      If TagFlag = False And QPTrim$(Index) = "TAG NUMBER" Then GoTo SkipEm1
      
      'collects grand totals
      OrigCost#(1) = OrigCost#(1) + FAItemRec.ORGCOST
      If RptType = "D" Then
        BookTotal#(1) = BookTotal#(1) + (FAItemRec.CURRVAL)
        DsplAmt = DsplAmt + FAItemRec.DisposAmt
      Else
        BookTotal#(1) = BookTotal#(1) + (FAItemRec.CURRVAL)
      End If
      YDep#(1) = YDep#(1) + YTDDep#
      DOrigCost#(1) = DOrigCost#(1) + FAItemRec.ORGCOST
      TagDOrigCost(Nextx) = DOrigCost#(1)
      DBookTotal#(1) = DBookTotal#(1) + (FAItemRec.CURRVAL)
      TagDBookTotal(Nextx) = DBookTotal#(1)
      DYDep#(1) = DYDep#(1) + YTDDep#
      TagDYDep(Nextx) = DYDep#(1)
      DptCnt(Nextx) = DptCnt(Nextx) + 1
      If RptType = "D" Then
        DDsplAmt(Nextx) = DDsplAmt(Nextx) + FAItemRec.DisposAmt
        TDsplAmt = TDsplAmt + FAItemRec.DisposAmt
      End If
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
    If RptType = "D" Then
      Print #RptHandle, String$(128, "-")
    Else
      Print #RptHandle, String$(113, "-")
    End If
    Print #RptHandle, "Totals for Dept Number: "; DeptNumber; Tab(30); DeptDescr; Tab(48); "#Items: "; CStr(DItemCnt);
    Print #RptHandle, Tab(65); Using("###,###,##0.00", CStr(DOrigCost#(1)));
    If RptType = "D" Then
      Print #RptHandle, Tab(80); Using("###,###,##0.00", CStr(DBookTotal#(1)));
      Print #RptHandle, Tab(97); Using("###,###,##0.00", CStr(DsplAmt))
      Print #RptHandle, String$(128, "=")
    ElseIf RptType = "A" Then
      Print #RptHandle, Tab(77); Using("###,###,##0.00", CStr(DBookTotal#(1)))
      Print #RptHandle, String$(113, "=")
    End If
    Print #RptHandle,
    LineCnt& = LineCnt& + 4
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
    DeptNumber = Val(QPTrim$(DeptArr(Nextx)))
    DeptDescr = DeptDesc(Nextx)
    'clear all dept totals
    DOrigCost#(1) = 0
    DBookTotal#(1) = 0
    DYDep#(1) = 0
    DItemCnt = 0
    If RptType = "D" Then
      DsplAmt = 0
    End If
  Loop
  
  Unload frmFAShowPctComp
  frmFAShowPctComp.Out = False
  EnableCloseButton Me.hwnd, True
  Me.cmdExit.Enabled = True
  Me.cmdProcess.Enabled = True
  
  If ItemTotal = 0 Then
    MsgBox "There are no fixed assets saved for this criteria."
    Close
    Exit Sub
  End If
  
  'only prints if TAG NUMBERS was selected
  If TagFlag = True Then
    GoSub PrintTagDeptTotals
  End If
  
  GoSub PrintMasterValueEnding1
  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
  
  Close         'Close all open files now
  
  ViewPrint ReportFile$, "Fixed Asset Report", True
  
  KillFile (ReportFile$)
  
  Exit Sub
  
PrintMasterHeader1:
  Page = Page + 1
  If RptType = "D" Then
    Print #RptHandle, Tab(30); "Master Asset Listing : Disposed Items"
    If FirstFlag = False Then
      If HeaderFlag = True Then
        Print #RptHandle, "Dept #: "; DeptNumber; " "; DeptDescr ' Dept$
      Else
        Print #RptHandle, "Dept #: ALL"
      End If
    End If
    Print #RptHandle, "Items Disposed From "; MakeRegDate(BDate); " to "; MakeRegDate(EDate)
    Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
    Print #RptHandle, "Asset Number"; Tab(22); "Description"; Tab(50); "Dept"; Tab(55); "Method"; Tab(66); "Original Cost"; Tab(84); "Book Value"; Tab(97); "Disposal Price"; Tab(116); "Disposal Date"
    Print #RptHandle, String$(128, "=")
  ElseIf RptType = "A" Then
    Print #RptHandle, Tab(30); "Master Asset Listing : Acquired Items"
    If FirstFlag = False Then
      If HeaderFlag = True Then
        Print #RptHandle, "Dept #: "; DeptNumber; " "; DeptDescr ' Dept$
      Else
        Print #RptHandle, "Dept #: ALL"
      End If
    End If
    Print #RptHandle, "Items Acquired From "; MakeRegDate(BDate); " to "; MakeRegDate(EDate)
    Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
    Print #RptHandle, "Asset Number"; Tab(22); "Description"; Tab(50); "Dept"; Tab(55); "Life/Left"; Tab(66); "Original Cost"; Tab(84); "Book Value"; Tab(98); "Acquisition Date"
    Print #RptHandle, String$(113, "=")
  End If
  LineCnt& = 7
  If FirstFlag = True Then
    FirstFlag = False
    LineCnt& = 6
  End If
  Return
  
PrintMasterValueEnding1:
  Page = Page + 1
  Print #RptHandle, FF$
  If RptType = "D" Then
    Print #RptHandle, Tab(30); "Master Asset Listing : Disposed Of Totals"
    Print #RptHandle, "Assets Deleted From "; MakeRegDate(BDate); " to "; MakeRegDate(EDate)
    Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
    Print #RptHandle, Tab(14); "#Items"; Tab(24); "Total Purchase Price"; Tab(48); "Total Book Value"; Tab(67); "Total Disp Price "
    Print #RptHandle, String$(83, "=")

    Print #RptHandle, "Total Deleted:"
    Print #RptHandle, Tab(17); CStr(ItemCnt);
    Print #RptHandle, Tab(30); Using("###,###,##0.00", CStr(OrigCost#(1)));
    Print #RptHandle, Tab(50); Using("###,###,##0.00", CStr(BookTotal#(1)));
    Print #RptHandle, Tab(69); Using("###,###,##0.00", CStr(TDsplAmt))

  ElseIf RptType = "A" Then
    Print #RptHandle, Tab(30); "Master Asset Listing : Acquired Totals"
    Print #RptHandle, "Assets Acquired From "; MakeRegDate(BDate); " to "; MakeRegDate(EDate)
    Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
    Print #RptHandle, Tab(17); "#Items"; Tab(26); "Total Purchase Price"; Tab(50); "Total Book Value"
    Print #RptHandle, String$(68, "=")

    Print #RptHandle, "Total Acquired:"; Tab(19); CStr(ItemCnt);
    Print #RptHandle, Tab(32); Using("###,###,##0.00", CStr(OrigCost#(1)));
    Print #RptHandle, Tab(52); Using("###,###,##0.00", CStr(BookTotal#(1)))
  End If
  Print #RptHandle, FF$
  
  Return
  
PrintTagDeptTotals: 'print only if TAG NUMBERS was selected
  Page = Page + 1
    Print #RptHandle, FF$
    Print #RptHandle, Tab(30); "Master Asset Listing : Department Totals"
  If RptType = "D" Then
    Print #RptHandle, "Assets Deleted From "; MakeRegDate(BDate); " to "; MakeRegDate(EDate)
    Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
    Print #RptHandle, "Dept Number"; Tab(14); "#Items"; Tab(24); "Total Purchase Price"; Tab(48); "Total Book Value"; Tab(67); "Total Disp Price"
    Print #RptHandle, String$(83, "=")
  ElseIf RptType = "A" Then
    Print #RptHandle, "Assets Acquired From "; MakeRegDate(BDate); " to "; MakeRegDate(EDate)
    Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
    Print #RptHandle, "Dept Number"; Tab(17); "#Items"; Tab(26); "Total Purchase Price"; Tab(50); "Total Book Value"
    Print #RptHandle, String$(68, "=")
  End If
  LineCnt = 5
  
  
  For x = 1 To DIdxCnt
    If RptType = "D" Then
      Print #RptHandle, Tab(4); Using$("###0", DeptArr(x)); Tab(17); CStr(DptCnt(x)); Tab(30); Using("###,###,##0.00", CStr(TagDOrigCost(x))); Tab(50); Using("###,###,##0.00", CStr(TagDBookTotal(x))); Tab(69); Using("###,###,##0.00", CStr(DDsplAmt(x)))
    ElseIf RptType = "A" Then
      Print #RptHandle, Tab(4); Using$("###0", DeptArr(x)); Tab(19); CStr(DptCnt(x)); Tab(32); Using("###,###,##0.00", CStr(TagDOrigCost(x))); Tab(52); Using("###,###,##0.00", CStr(TagDBookTotal(x)))
    End If
    LineCnt = LineCnt + 1

    If LineCnt& >= MaxLines And x <> DIdxCnt Then
      LineCnt& = 0
      Page = Page + 1
      Print #RptHandle, FF$
      Print #RptHandle, Tab(20); "Master Asset Listing : Department Totals"
      If RptType = "D" Then
        Print #RptHandle, "Assets Deleted From "; MakeRegDate(BDate); " to "; MakeRegDate(EDate)
        Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
        Print #RptHandle, "Dept Number"; Tab(14); "#Items"; Tab(24); "Total Purchase Price"; Tab(48); "Total Book Value"; Tab(67); "Total Disp Price"
      ElseIf RptType = "A" Then
        Print #RptHandle, "Assets Acquired From "; MakeRegDate(BDate); " to "; MakeRegDate(EDate)
        Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
        Print #RptHandle, "Dept Number"; Tab(17); "#Items"; Tab(26); "Total Purchase Price"; Tab(50); "Total Book Value"
      End If
      Print #RptHandle, String$(68, "=")
      LineCnt = LineCnt + 5
    End If
  Next x
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAItemsAddDelOptRpt", "PrintText", Erl)
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

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
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
      KillFile "disposedofrpt.dat"
      KillFile "newadddelrptopen.dat"
      ClearInUse PWcnt
      MainLog ("FixedAssets.exe terminated via menu bar on frmFAItemsAddDelOptRpt.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub fpcmbOrder_Change()
  'disable the department number field if tag number
  'is selected for printing order
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
  'prevents user from inadvertently changing data in this combo box
  'if he is scrolling through the form
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
  'newadddelrptopen.dat is used by the dept list form to know
  'that it was called from this form (so if a dept is double clicked
  'the selected dept will appear on this screen)
  Open "newadddelrptopen.dat" For Output As FileHandle Len = 2
  Print #FileHandle, One
  Close FileHandle
  fpcmbOrder.Text = "TAG NUMBER"
  fpcmbOrder.AddItem "TAG NUMBER"
  fpcmbOrder.AddItem "DEPARTMENT NUMBER"
  fpcmbRptType.Text = "BOTH"
  fpcmbRptType.AddItem "BOTH"
  fpcmbRptType.AddItem "DISPOSED OF ONLY"
  fpcmbRptType.AddItem "ADDITIONS ONLY"
  fptxtDeptNum.Text = "ALL"
  fpcomboPrintOpt.AddItem "Graphical"
  fpcomboPrintOpt.AddItem "Text"
  fpcomboPrintOpt.Text = "Graphical"
  fpDateStart = Date
  fpDateEnd = Date
End Sub

Private Sub fpcmbRptType_Change()
  'if user deletes the value in this field it will
  'default to BOTH if left blank
  If QPTrim$(fpcmbRptType.Text) = "" Then
    fpcmbRptType.Text = "BOTH"
  End If
End Sub

Private Sub fpcmbRptType_KeyDown(KeyCode As Integer, Shift As Integer)
  'prevents user from inadvertently changing data in this combo box
  'if he is scrolling through the form
  If KeyCode = vbKeySpace Then
    fpcmbRptType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbRptType.ListIndex = -1
  End If
  If fpcmbRptType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      If fptxtDeptNum.Enabled = True Then
        fptxtDeptNum.SetFocus
      Else
        cmdDept.SetFocus
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

Private Sub fpcomboPrintOpt_Change()
  'if user deletes this field's value and leaves it
  'blank then it will default to Graphical
  If QPTrim$(fpcomboPrintOpt.Text) = "" Then
    fpcomboPrintOpt.Text = "Graphical"
  End If
End Sub

Private Sub fptxtDeptNum_Change()
  'if user deletes this field's value and leaves it
  'blank then it will default to ALL
  If fptxtDeptNum.Text = "" Then
    fptxtDeptNum = "ALL"
  End If
End Sub


Private Sub fpcomboPrintOpt_KeyDown(KeyCode As Integer, Shift As Integer)
  'prevents user from inadvertently changing data in this combo box
  'if he is scrolling through the form
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

Private Function Check4ValidDept() As Boolean
  Dim x As Integer
  Dim DeptIdx As DeptNumbSortIdxType
  Dim DIdxHandle As Integer
  Dim DIdxRecNums As Integer
  Dim ThisDept$
  
  On Error GoTo ERRORSTUFF
  
  'this function is commented in frmFAEditItemWTabs
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
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAItemsAddDelOptRpt", "Check4ValidDept", Erl)
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

Private Sub Form_KeyPress(KeyAscii As Integer)
'  this causes all characters to be capitalized
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub PrintBothText()
  Dim DOrigCost#(2), DBookTotal#(2), DCDep#(2), DYDep#(2), OrigCost#(2), BookTotal#(2), CDep#(2), YDep#(2)
  Dim ReportFile$
  Dim Dash80$
  Dim FF$
  Dim MaxLines As Integer
  Dim ItemCnt&
  Dim LineCnt&
  Dim Dept$
  Dim BDate As Integer
  Dim EDate As Integer
  Dim RptHandle As Integer
  Dim FAHandle As Integer
  Dim NumOfFARecs As Integer
  Dim FAItemRec As FAItemRecType
  Dim cnt&
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim DisposeDate As Integer
  Dim AcquireDate As Integer
  Dim DFlag As Boolean
  Dim AFlag As Boolean
  Dim DeptNumber As Integer
  Dim YTDDep#
  Dim x As Integer
  Dim Nextx As Integer
  Dim DIdxRec As DeptNumbSortIdxType
  Dim DIdxHandle As Integer
  Dim DIdxCnt As Integer
  Dim Index$ ', Tab1 As Integer, Tab2 As Integer, Tab3 As Integer
  Dim Page As Integer
  Dim DataFlag As Boolean
  Dim TagFlag As Boolean
  Dim PFlag As Boolean
  Dim DeptDescr$
  Dim DItemCnt As Long
  Dim AItemCnt As Long
  Dim HeaderFlag As Boolean
  Dim FirstFlag As Boolean
  Dim DeptCnt As Integer
  Dim ItemTotal As Long
  Dim TotDsplPrice As Double
  
  On Error GoTo ERRORSTUFF
  FirstFlag = True
  If fpcmbOrder.Text = "TAG NUMBER" Then
    HeaderFlag = False
  Else
    HeaderFlag = True
  End If
  
  ReportFile$ = "FANEWADDDEL.PRN"  'Report File Name
  Dash80$ = String$(80, "=")
  FF$ = Chr$(12)
  
  MaxLines = 56
  LineCnt& = 0
  ItemCnt& = 0
  Dept$ = QPTrim$(fptxtDeptNum.Text)
  BDate = Date2Num(fpDateStart.Text) 'beginning date
  EDate = Date2Num(fpDateEnd.Text) 'ending date
 
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
  ReDim DeptArr(1 To DIdxCnt) As String 'create array for
  'department record numbers
  ReDim DeptDesc(1 To DIdxCnt) As String 'create array for
  'department descriptions
  
  For x = 1 To DIdxCnt
    Get DIdxHandle, x, DIdxRec
    DeptArr(x) = QPTrim$(DIdxRec.DeptNumb) 'load array
    DeptDesc(x) = QPTrim$(DIdxRec.DeptIdxDesc) 'load array
  Next x
  Close DIdxHandle
  
  'create arrays to collect totals by dept
  ReDim DTagDOrigCost(1 To DIdxCnt) As Double
  ReDim DTagDBookTotal(1 To DIdxCnt) As Double
  ReDim DTagDYDep(1 To DIdxCnt) As Double
  ReDim DDeptCnt(1 To DIdxCnt) As Long
  ReDim ATagDOrigCost(1 To DIdxCnt) As Double
  ReDim ATagDBookTotal(1 To DIdxCnt) As Double
  ReDim ATagDYDep(1 To DIdxCnt) As Double
  ReDim ADeptCnt(1 To DIdxCnt) As Long
  
  If Dept$ <> "ALL" Then 'user wants one dept only
    DeptNumber = Val(QPTrim$(fptxtDeptNum.Text))
    For x = 1 To DIdxCnt
      If DeptNumber = Val(QPTrim$(DeptArr(x))) Then
        DeptDescr = QPTrim$(DeptDesc(x)) 'this is the
        'description for the selected department
        Exit For 'got it so jump out of loop
      End If
    Next x
  Else 'user wants a report for all departments
    DeptNumber = Val(QPTrim(DeptArr(1))) 'get first record
    DeptDescr = QPTrim$(DeptDesc(1)) 'get first description
  End If
  GoSub PrintMasterHeader1
  
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
      'Check For Disposed Date
      DisposeDate = FAItemRec.DispDate
      'Check for Acquired Date
      AcquireDate = FAItemRec.AQURDATE
      AFlag = False 'acquire flag
      DFlag = False 'disposal flag
      PFlag = False 'pending disposal flag
      
      If DisposeDate >= BDate And DisposeDate <= EDate Or AcquireDate >= BDate And AcquireDate <= EDate Then
      'filter out items that don't fall inside the date parameters
        If DisposeDate >= BDate And DisposeDate <= EDate Then
          If FAItemRec.DsplFlag = 2 Then
            DFlag = True
          ElseIf FAItemRec.DsplFlag = 1 Then
            PFlag = True
          End If
        End If
        
        If AcquireDate >= BDate And AcquireDate <= EDate Then
          AFlag = True
        End If
      Else
        GoTo SkipEm1
      End If
      'Can't depreciate something that has no asset life...this should
      'not be necessary unless the asset life was changed to zero after
      'depreciation had already started
      If FAItemRec.ILIFE > 0 Then
        YTDDep# = FAItemRec.DEP2DATE
      Else
        YTDDep# = 0
      End If
      
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
      If DFlag = True Then
        TotDsplPrice = TotDsplPrice + FAItemRec.DisposAmt
      End If
      If TagFlag = True Then GoTo TagOnly2
      
TagOnly1: 'printing valid tag data...skipped if TAG NUMBER is chosen and this is
          'not the first iteration
      If QPTrim$(fpcmbOrder.Text) = "DEPARTMENT NUMBER" And DeptCnt = 0 Then
'        Print #RptHandle, String$(109, "=")
        Print #RptHandle, "Department: "; DeptNumber; DeptDescr
        Print #RptHandle, String$(109, "-")
        LineCnt& = LineCnt& + 2
      End If
      
      If DFlag = True And AFlag = True Then
        Print #RptHandle, "AD";
      ElseIf PFlag = True And AFlag = True Then
        Print #RptHandle, "AP";
      ElseIf DFlag = True Then
        Print #RptHandle, "D";
      ElseIf AFlag = True Then
        Print #RptHandle, "A";
      ElseIf PFlag = True Then
        Print #RptHandle, "P";
      End If
      DataFlag = True
      DeptCnt = DeptCnt + 1
      Print #RptHandle, FAItemRec.ItemTag; Tab(22); Left$(FAItemRec.IDESC1, 28);
      Print #RptHandle, Tab(51); FAItemRec.IDEPT;
      Print #RptHandle, Tab(58); Using("#0", FAItemRec.ILIFE); "/"; Using("#0", FAItemRec.LifeLeft);
      Print #RptHandle, Tab(66); Using("###,###,##0.00", CStr(FAItemRec.ORGCOST));
      If QPTrim$(FAItemRec.DEPYN) = "N" And FAItemRec.DsplFlag <> 2 Then
        Print #RptHandle, Tab(80); Using("###,###,##0.00", CStr(YTDDep#)); "*";
      Else
        Print #RptHandle, Tab(80); Using("###,###,##0.00", CStr(YTDDep#));
      End If
      Print #RptHandle, Tab(96); Using("###,###,##0.00", CStr(FAItemRec.CURRVAL))
      
      LineCnt& = LineCnt& + 1
      ItemCnt& = ItemCnt& + 1
      ItemTotal = ItemTotal + 1
TagOnly2: 'collects data for each department for reporting totals

      'This if statement filters out the first iteration of TAG NUMBER
      'selection because we do not want to start accumulating dept data
      'until the second iteration
      If TagFlag = False And QPTrim$(Index) = "TAG NUMBER" Then GoTo SkipEm1
      
      'collects grand totals
      If DFlag = True Or PFlag = True Then
        OrigCost#(1) = OrigCost#(1) + FAItemRec.ORGCOST
        BookTotal#(1) = BookTotal#(1) + (FAItemRec.CURRVAL)
        YDep#(1) = YDep#(1) + YTDDep#
      End If
      If AFlag = True Then
        OrigCost#(2) = OrigCost#(2) + FAItemRec.ORGCOST
        BookTotal#(2) = BookTotal#(2) + (FAItemRec.CURRVAL)
        YDep#(2) = YDep#(2) + YTDDep#
      End If
      
      'collects dept totals
      If DFlag = True Or PFlag = True Then
        DOrigCost#(1) = DOrigCost#(1) + FAItemRec.ORGCOST
        DTagDOrigCost(Nextx) = DOrigCost#(1)
        DBookTotal#(1) = DBookTotal#(1) + (FAItemRec.CURRVAL)
        DTagDBookTotal(Nextx) = DBookTotal#(1)
        DYDep#(1) = DYDep#(1) + YTDDep#
        DTagDYDep(Nextx) = DYDep#(1)
        DDeptCnt(Nextx) = DDeptCnt(Nextx) + 1
        DItemCnt = DItemCnt + 1
      End If
      
      If AFlag = True Then
        DOrigCost#(2) = DOrigCost#(2) + FAItemRec.ORGCOST
        ATagDOrigCost(Nextx) = DOrigCost#(2)
        DBookTotal#(2) = DBookTotal#(2) + (FAItemRec.CURRVAL)
        ATagDBookTotal(Nextx) = DBookTotal#(2)
        DYDep#(2) = DYDep#(2) + YTDDep#
        ATagDYDep(Nextx) = DYDep#(2)
        ADeptCnt(Nextx) = ADeptCnt(Nextx) + 1
        AItemCnt = AItemCnt + 1
      End If
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
    Print #RptHandle, String$(109, "-")
    Print #RptHandle, "Totals for: "; DeptNumber; DeptDescr
    Print #RptHandle, Tab(20); "Additions: "; Tab(35); Using("###0", ADeptCnt(Nextx));
    Print #RptHandle, Tab(66); Using("###,###,##0.00", CStr(DOrigCost#(2)));
    Print #RptHandle, Tab(80); Using("###,###,##0.00", CStr(DYDep#(2)));
    Print #RptHandle, Tab(96); Using("###,###,##0.00", CStr(DBookTotal#(2)))
    
    Print #RptHandle, Tab(20); "Deletions: "; Tab(35); Using("###0", DDeptCnt(Nextx));
    Print #RptHandle, Tab(66); Using("###,###,##0.00", CStr(DOrigCost#(1)));
    Print #RptHandle, Tab(80); Using("###,###,##0.00", CStr(DYDep#(1)));
    Print #RptHandle, Tab(96); Using("###,###,##0.00", CStr(DBookTotal#(1)))
    
    Print #RptHandle, String$(109, "=")
    Print #RptHandle,
    LineCnt& = LineCnt& + 6
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
    DeptNumber = Val(QPTrim$(DeptArr(Nextx)))
    DeptDescr = QPTrim$(DeptDesc(Nextx))
    'clear all dept totals
    DOrigCost#(1) = 0 'reset to get ready for the next dept
    DBookTotal#(1) = 0 'reset to get ready for the next dept
    DYDep#(1) = 0 'reset to get ready for the next dept
    DOrigCost#(2) = 0 'reset to get ready for the next dept
    DBookTotal#(2) = 0 'reset to get ready for the next dept
    DYDep#(2) = 0 'reset to get ready for the next dept
    DeptCnt = 0 'reset to get ready for the next dept
  Loop
  
  Unload frmFAShowPctComp
  frmFAShowPctComp.Out = False
  EnableCloseButton Me.hwnd, True
  Me.cmdExit.Enabled = True
  Me.cmdProcess.Enabled = True
  
  If ItemTotal = 0 Then
    MsgBox "There are no fixed assets saved for this criteria."
    Close
    Exit Sub
  End If
  
  'only prints if TAG NUMBERS was selected
  If TagFlag = True Then
    GoSub PrintTagDeptTotals
  End If
  
  GoSub PrintMasterValueEnding1
  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
  
  Close         'Close all open files now
  
  ViewPrint ReportFile$, "Fixed Asset Report", True
  
  KillFile (ReportFile$)
  
  Exit Sub
  
PrintMasterHeader1:
  Page = Page + 1
  Print #RptHandle, Tab(30); "Master Asset Listing : Additions and Deletions"
  If FirstFlag = False Then
    If HeaderFlag = True Then
      Print #RptHandle, "Dept #: "; DeptNumber; " "; DeptDescr ' Dept$
    Else
      Print #RptHandle, "Dept #: ALL"
    End If
  End If
  Print #RptHandle, "Assets Acquired From "; MakeRegDate(BDate); " to "; MakeRegDate(EDate)
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
  Print #RptHandle, "* = DO NOT DEPRECIATE THIS ASSET"
  Print #RptHandle, "Asset Number"; Tab(22); "Description"; Tab(51); "Dept"; Tab(56); "Life/Left"; Tab(67); "Original Cost"; Tab(82); "Total Deprec"; Tab(100); "Book Value"
  Print #RptHandle, String$(109, "=")
  LineCnt& = 7
  If FirstFlag = True Then
    FirstFlag = False
    LineCnt& = 6
  End If
  
  Return
  
PrintMasterValueEnding1:
  Page = Page + 1
  Print #RptHandle, FF$
  Print #RptHandle, Tab(30); "Master Asset Listing : Grand Totals"
  Print #RptHandle, "Assets Acquired From "; MakeRegDate(BDate); " to "; MakeRegDate(EDate)
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
  Print #RptHandle, Tab(20); "# Items"; Tab(35); "Total Purchase Price"; Tab(58); "Total Depreciation"; Tab(78); "Total Book Value"
  Print #RptHandle, String$(93, "=")
  Print #RptHandle, "Total Additions "; Tab(21); Using("####0", AItemCnt);
  Print #RptHandle, Tab(41); Using("###,###,##0.00", CStr(OrigCost#(2)));
  Print #RptHandle, Tab(62); Using("###,###,##0.00", CStr(YDep#(2)));
  Print #RptHandle, Tab(80); Using("###,###,##0.00", CStr(BookTotal#(2)))
  
  Print #RptHandle, "Total Deletions "; Tab(21); Using("####0", DItemCnt);
  Print #RptHandle, Tab(41); Using("###,###,##0.00", CStr(OrigCost#(1)));
  Print #RptHandle, Tab(62); Using("###,###,##0.00", CStr(YDep#(1)));
  Print #RptHandle, Tab(80); Using("###,###,##0.00", CStr(BookTotal#(1)))
  If TotDsplPrice > 0 Then
    Print #RptHandle, "Total Disposal Price: "; Tab(41); Using("###,###,##0.00", TotDsplPrice)
  End If
  Print #RptHandle, FF$
  
  Return
  
PrintTagDeptTotals: 'print only if TAG NUMBERS was selected
  
  Page = Page + 1
  Print #RptHandle, FF$
  Print #RptHandle, Tab(30); "Master Asset Listing : Department Totals"
  Print #RptHandle, "Assets Acquired From "; MakeRegDate(BDate); " to "; MakeRegDate(EDate)
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
  Print #RptHandle, "  Department"; Tab(35); "Total Purchase Price"; Tab(58); "Total Depreciation"; Tab(78); "Total Book Value"
  Print #RptHandle, String$(93, "=")
  LineCnt = 5
  
  
  For x = 1 To DIdxCnt
    Print #RptHandle, DeptArr(x); " "; DeptDesc(x)
    Print #RptHandle, Tab(10); "Additions"; Tab(20); CStr(ADeptCnt(x)); Tab(41); Using("###,###,##0.00", CStr(ATagDOrigCost(x))); Tab(62); Using("###,###,##0.00", CStr(ATagDYDep(x))); Tab(80); Using("###,###,##0.00", CStr(ATagDBookTotal(x)))
    Print #RptHandle, Tab(10); "Deletions"; Tab(20); CStr(DDeptCnt(x)); Tab(41); Using("###,###,##0.00", CStr(DTagDOrigCost(x))); Tab(62); Using("###,###,##0.00", CStr(DTagDYDep(x))); Tab(80); Using("###,###,##0.00", CStr(DTagDBookTotal(x)))
    LineCnt = LineCnt + 3
    
    'If dept total data extends past one page then this will kick in
    If LineCnt& >= MaxLines And x <> DIdxCnt Then
      LineCnt& = 0
      Page = Page + 1
      Print #RptHandle, FF$
      Print #RptHandle, Tab(20); "Master Asset Listing : Department Totals"
      Print #RptHandle, "Assets Acquired From "; MakeRegDate(BDate); " to "; MakeRegDate(EDate)
      Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
      Print #RptHandle, "  Department"; Tab(35); "Total Purchase Price"; Tab(58); "Total Depreciation"; Tab(78); "Total Book Value"
      Print #RptHandle, String$(93, "=")
      LineCnt = LineCnt + 5
    End If
  Next x
  If TotDsplPrice > 0 Then
    Print #RptHandle, "Total Disposal Price: "; Tab(41); Using("###,###,##0.00", TotDsplPrice)
  End If
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAItemsAddDelOptRpt", "PrintBothText", Erl)
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

Private Sub PrintBothGraphics()
  Dim DOrigCost#(2), DBookTotal#(2), DCDep#(2), DYDep#(2), OrigCost#(2), BookTotal#(2), CDep#(2), YDep#(2)
  Dim ReportFile$
  Dim Dept$
  Dim BDate As Integer
  Dim EDate As Integer
  Dim RptHandle As Integer
  Dim FAHandle As Integer
  Dim NumOfFARecs As Integer
  Dim FAItemRec As FAItemRecType
  Dim cnt&
  Dim TagIdx As TagNumbSortIdxType
  Dim TagIdxHandle As Integer
  Dim DisposeDate As Integer
  Dim AcquireDate As Integer
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
  Dim DataFlag As Boolean
  Dim TagFlag As Boolean
  Dim dlm$
  Dim Employer$
  Dim FASetUpRec As FASetupRecType
  Dim DCnt As Integer
  Dim ACnt As Integer
  Dim NoDep$
  Dim TagHandle As Integer
  Dim TagReportFile$
  Dim TagSubHandle As Integer
  Dim TagSubReportFile$
  Dim TDCnt As Integer
  Dim TACnt As Integer
  Dim TagSign As Integer
  Dim TagGrandHandle As Integer
  Dim TagGrandReportFile$
  Dim PFlag As Boolean
  Dim DeptDescr$
  Dim TotDsplPrice As Double
  Dim ItemTotal As Long
  
  'see PrintBothText for comments describing it's operations
  'since it is almost identical to this sub
  
  On Error GoTo ERRORSTUFF
  TagSign = 0
  OpenFASetUpFile FAHandle
  Get FAHandle, 1, FASetUpRec
  Employer = FASetUpRec.TownName
  Close FAHandle
  dlm$ = "~"
  ReportFile$ = "FARPTS\FAADDNEW.RPT"  'Report File Name
  TagReportFile$ = "FARPTS\FAADDDELTAG.RPT"
  TagSubReportFile$ = "FARPTS\SUBTAG.RPT"
  TagGrandReportFile$ = "FARPTS\SUBGRANDTAG.RPT"
  Dept$ = QPTrim$(fptxtDeptNum.Text)
  BDate = Date2Num(fpDateStart.Text)
  EDate = Date2Num(fpDateEnd.Text)
  Index$ = QPTrim$(fpcmbOrder.Text)
  
  If QPTrim$(Index) = "TAG NUMBER" Then
    TagHandle = FreeFile
    Open TagReportFile$ For Output As TagHandle
    TagSubHandle = FreeFile
    Open TagSubReportFile$ For Output As TagSubHandle
    TagGrandHandle = FreeFile
    Open TagGrandReportFile$ For Output As TagGrandHandle
  Else
    RptHandle = FreeFile
    Open ReportFile$ For Output As #RptHandle
  End If
  
  OpenTagIdxFile TagIdxHandle
  NumOfFARecs = LOF(TagIdxHandle) \ Len(TagIdx)
  If NumOfFARecs = 0 Then
    MsgBox "No item records on file."
    Close TagIdxHandle
    Close
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
  ReDim DeptArr(1 To DIdxCnt) As String
  ReDim DeptDesc(1 To DIdxCnt) As String
  
  For x = 1 To DIdxCnt
    Get DIdxHandle, x, DIdxRec
    DeptArr(x) = QPTrim$(DIdxRec.DeptNumb)
    DeptDesc(x) = QPTrim$(DIdxRec.DeptIdxDesc)
  Next x
  Close DIdxHandle
  
  ReDim DTagDOrigCost(1 To DIdxCnt) As Double
  ReDim DTagDBookTotal(1 To DIdxCnt) As Double
  ReDim DTagDYDep(1 To DIdxCnt) As Double
  ReDim ATagDOrigCost(1 To DIdxCnt) As Double
  ReDim ATagDBookTotal(1 To DIdxCnt) As Double
  ReDim ATagDYDep(1 To DIdxCnt) As Double
  ReDim TotalA(1 To DIdxCnt) As Integer
  ReDim TotalD(1 To DIdxCnt) As Integer
  
  If Dept$ <> "ALL" Then
    DeptNumber = Val(QPTrim$(fptxtDeptNum.Text))
    For x = 1 To DIdxCnt
      If DeptNumber = Val(QPTrim$(DeptArr(x))) Then
        DeptDescr = QPTrim$(DeptDesc(x))
        Exit For
      End If
    Next x
  Else
    DeptNumber = Val(QPTrim(DeptArr(1)))
    DeptDescr = QPTrim$(DeptDesc(1))
  End If
  
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
      DFlag = False
      AFlag = False
      PFlag = False
      'Check For Disposed Date
      DisposeDate = FAItemRec.DispDate
      'Check for Acquired Date
      AcquireDate = FAItemRec.AQURDATE
      
      If DisposeDate >= BDate And DisposeDate <= EDate Or AcquireDate >= BDate And AcquireDate <= EDate Then
      'filter out items that don't fall inside the date parameters
        If DisposeDate >= BDate And DisposeDate <= EDate Then
          If FAItemRec.DsplFlag = 2 Then
            DFlag = True
          ElseIf FAItemRec.DsplFlag = 1 Then
            PFlag = True
          End If
        End If
        
        If AcquireDate >= BDate And AcquireDate <= EDate Then
          AFlag = True
        End If
      Else
        GoTo SkipEm1
      End If
      'Can't depreciate something that has no asset life...this should
      'not be necessary unless the asset life was changed to zero after
      'depreciation had already started
      If FAItemRec.ILIFE > 0 Then
        YTDDep# = FAItemRec.DEP2DATE
      Else
        YTDDep# = 0
      End If
      
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
      If DFlag = True Then
        TotDsplPrice = FAItemRec.DisposAmt + TotDsplPrice
      End If
      If TagFlag = True Then GoTo TagOnly2
TagOnly1: 'printing valid tag data...skipped if TAG NUMBER is chosen and this is
          'not the first iteration
      If QPTrim$(Index) = "TAG NUMBER" Then
        '                     0                  1                        2
        Print #TagHandle, Employer$; dlm; MakeRegDate(BDate); dlm; MakeRegDate(EDate); dlm;
        '                                       3
        If DFlag = True And AFlag = True Then
          Print #TagHandle, "DA"; dlm;
        ElseIf PFlag = True And AFlag = True Then
          Print #TagHandle, "AP"; dlm;
        ElseIf DFlag = True Then
          Print #TagHandle, "D"; dlm;
        ElseIf AFlag = True Then
          Print #TagHandle, "A"; dlm;
        ElseIf PFlag = True Then
          Print #TagHandle, "P"; dlm;
        Else
          Print #TagHandle, " "; dlm;
        End If
        
        DataFlag = True
        '                         4                            5
        Print #TagHandle, FAItemRec.ItemTag; dlm; QPTrim(FAItemRec.IDESC1); dlm;
        '                         6
        Print #TagHandle, FAItemRec.IDEPT; dlm;
        '                         7
        Print #TagHandle, FAItemRec.ILIFE; dlm;
        '                         8
        Print #TagHandle, FAItemRec.ORGCOST; dlm;
        If QPTrim$(FAItemRec.DEPYN) = "N" And FAItemRec.DsplFlag <> 2 Then
          NoDep = "*"
        Else
          NoDep = ""
        End If
        '                         9
        Print #TagHandle, YTDDep#; dlm;
        '                                   10     11                12
        Print #TagHandle, FAItemRec.CURRVAL; dlm; NoDep; dlm; FAItemRec.LifeLeft
      Else
        '                     0                  1                        2
        Print #RptHandle, Employer$; dlm; MakeRegDate(BDate); dlm; MakeRegDate(EDate); dlm;
        '                                       3
        If DFlag = True And AFlag = True Then
          Print #RptHandle, "DA"; dlm;
        ElseIf PFlag = True And AFlag = True Then
          Print #RptHandle, "AP"; dlm;
        ElseIf DFlag = True Then
          Print #RptHandle, "D"; dlm;
        ElseIf AFlag = True Then
          Print #RptHandle, "A"; dlm;
        ElseIf PFlag = True Then
          Print #RptHandle, "P"; dlm;
        Else
          Print #RptHandle, " "; dlm;
        End If
        DataFlag = True
        '                         4                            5
        Print #RptHandle, FAItemRec.ItemTag; dlm; QPTrim(FAItemRec.IDESC1); dlm;
        '                         6
        Print #RptHandle, FAItemRec.IDEPT; dlm;
        '                         7
        Print #RptHandle, FAItemRec.ILIFE; dlm;
        '                         8
        Print #RptHandle, FAItemRec.ORGCOST; dlm;
        If QPTrim$(FAItemRec.DEPYN) = "N" And FAItemRec.DsplFlag Then
          NoDep = "*"
        Else
          NoDep = ""
        End If
        '                         9
        Print #RptHandle, YTDDep#; dlm;
        '                         10
        Print #RptHandle, FAItemRec.CURRVAL; dlm;
    End If
    ItemTotal = ItemTotal + 1
TagOnly2: 'collects data for each department for reporting totals

      'This if statement filters out the first iteration of TAG NUMBER
      'selection because we do not want to start accumulating dept data
      'until the second iteration
      If TagFlag = False And QPTrim$(Index) = "TAG NUMBER" Then GoTo SkipEm1
      
      'collects grand totals
      If DFlag = True Or PFlag = True Then
        TDCnt = TDCnt + 1
        OrigCost#(1) = OrigCost#(1) + FAItemRec.ORGCOST
        BookTotal#(1) = BookTotal#(1) + (FAItemRec.CURRVAL)
        YDep#(1) = YDep#(1) + YTDDep#
      End If
      If AFlag = True Then
        TACnt = TACnt + 1
        OrigCost#(2) = OrigCost#(2) + FAItemRec.ORGCOST
        BookTotal#(2) = BookTotal#(2) + (FAItemRec.CURRVAL)
        YDep#(2) = YDep#(2) + YTDDep#
      End If
      
      'collects dept totals
      If DFlag = True Or PFlag = True Then
        DOrigCost#(1) = DOrigCost#(1) + FAItemRec.ORGCOST
        DTagDOrigCost(Nextx) = DOrigCost#(1)
        DBookTotal#(1) = DBookTotal#(1) + (FAItemRec.CURRVAL)
        DTagDBookTotal(Nextx) = DBookTotal#(1)
        DYDep#(1) = DYDep#(1) + YTDDep#
        DTagDYDep(Nextx) = DYDep#(1)
        DCnt = DCnt + 1
        TotalD(Nextx) = TotalD(Nextx) + 1
      End If
      
      If AFlag = True Then
        DOrigCost#(2) = DOrigCost#(2) + FAItemRec.ORGCOST
        ATagDOrigCost(Nextx) = DOrigCost#(2)
        DBookTotal#(2) = DBookTotal#(2) + (FAItemRec.CURRVAL)
        ATagDBookTotal(Nextx) = DBookTotal#(2)
        DYDep#(2) = DYDep#(2) + YTDDep#
        ATagDYDep(Nextx) = DYDep#(2)
        ACnt = ACnt + 1
        TotalA(Nextx) = TotalA(Nextx) + 1
      End If
      
      If TagHandle = 0 Then
        '                     11                  12                 13
        Print #RptHandle, DOrigCost#(1); dlm; DYDep#(1); dlm; DBookTotal#(1); dlm;
        '                     14                  15                 16
        Print #RptHandle, DOrigCost#(2); dlm; DYDep#(2); dlm; DBookTotal#(2); dlm;
        '                     17                  18                 19
        Print #RptHandle, OrigCost#(1); dlm; YDep#(1); dlm; BookTotal#(1); dlm;
        '                     20                  21                 22          23         24         25
        Print #RptHandle, OrigCost#(2); dlm; YDep#(2); dlm; BookTotal#(2); dlm; ACnt; dlm; DCnt; dlm; NoDep; dlm;
        '                   26          27            28                 29                  30
        Print #RptHandle, TACnt; dlm; TDCnt; dlm; DeptDescr; dlm; FAItemRec.LifeLeft; dlm; CStr(TotDsplPrice)
      End If
      
SkipEm1:

    Next cnt&
    'here we begin the iteration over again but this time TagFlag
    'becomes true so we know that this was originally TAG NUMBERS
    'and the first iteration is done
    If QPTrim$(Index) = "TAG NUMBER" And TagFlag = False Then
      TagFlag = True
      TagSign = 1
      GoTo GetTagTotals
      Exit Do
    End If
    
    If TagFlag = True Then GoTo NoData 'don't want the next dept
    'data to print
    
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
    DeptNumber = Val(QPTrim$(DeptArr(Nextx)))
    DeptDescr = QPTrim(DeptDesc(Nextx))
    'clear all dept totals
    DOrigCost#(1) = 0
    DBookTotal#(1) = 0
    DYDep#(1) = 0
    DOrigCost#(2) = 0
    DBookTotal#(2) = 0
    DYDep#(2) = 0
    DCnt = 0
    ACnt = 0
  Loop
  
  Unload frmFAShowPctComp
  frmFAShowPctComp.Out = False
  EnableCloseButton Me.hwnd, True
  Me.cmdExit.Enabled = True
  Me.cmdProcess.Enabled = True
  
  If ItemTotal = 0 Then
    MsgBox "There are no fixed assets saved for this criteria."
    Close
    Exit Sub
  End If
    
  'only prints if TAG NUMBERS was selected
  If TagFlag = True Then
    GoSub PrintTagDeptTotals
  End If
  
  Close         'Close all open files now
  
  If TagFlag = False Then
    arFANewAddDelRpt.Show
  Else
    arFAAddDelTagOnly.Show
  End If
  
  Exit Sub
  
PrintTagDeptTotals: 'print only if TAG NUMBERS was selected
  
  For x = 1 To DIdxCnt
    If QPTrim$(DeptArr(x)) = "" Then DeptArr(x) = "0"
    '                         0                   1                   2
    Print #TagSubHandle, DTagDOrigCost(x); dlm; DTagDYDep(x); dlm; DTagDBookTotal(x); dlm;
    '                         3                       4                   5                   6
    Print #TagSubHandle, ATagDOrigCost(x); dlm; ATagDYDep(x); dlm; ATagDBookTotal(x); dlm; DeptArr(x); dlm;
    '                        7              8
    Print #TagSubHandle, TotalD(x); dlm; TotalA(x)
  Next x
    '                          0                 1                  2
    Print #TagGrandHandle, OrigCost#(1); dlm; YDep#(1); dlm; BookTotal#(1); dlm;
    '                          3                 4                  5
    Print #TagGrandHandle, OrigCost#(2); dlm; YDep#(2); dlm; BookTotal#(2); dlm;
    '                        6           7
    Print #TagGrandHandle, TACnt; dlm; TDCnt
  
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmFAItemsAddDelOptRpt", "PrintBothGraphics", Erl)
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

Private Sub fptxtDeptNum_LostFocus()
  'if user deletes the value in this field and leaves it
  'blank then it will default to ALL
  If QPTrim$(fptxtDeptNum.Text) = "" Then fptxtDeptNum = "ALL"
End Sub

