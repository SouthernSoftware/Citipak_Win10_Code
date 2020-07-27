VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmEmpHistRptSplash 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Earnings History Report"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmEmpHistRptSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8840
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   6156
      Left            =   2088
      TabIndex        =   6
      Top             =   1344
      Width           =   7452
      _Version        =   196609
      _ExtentX        =   13144
      _ExtentY        =   10858
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmEmpHistRptSplash.frx":08CA
      Begin LpLib.fpCombo fpcomboPrintOpt 
         Height          =   405
         Left            =   3480
         TabIndex        =   5
         Top             =   4515
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
         Style           =   0
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
         EditAlignH      =   1
         EditAlignV      =   0
         ColDesigner     =   "frmEmpHistRptSplash.frx":08E6
      End
      Begin LpLib.fpCombo fptxtSummary 
         Height          =   405
         Left            =   4650
         TabIndex        =   4
         ToolTipText     =   "This is an abbreviated version."
         Top             =   3885
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
         EditAlignH      =   1
         EditAlignV      =   0
         ColDesigner     =   "frmEmpHistRptSplash.frx":0BDD
      End
      Begin EditLib.fpDateTime fptxtStartDate 
         Height          =   396
         Left            =   3648
         TabIndex        =   2
         Top             =   2640
         Width           =   1788
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
         Text            =   "11/20/2002"
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
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpText fptxtFirstEmpNo 
         Height          =   396
         Left            =   4176
         TabIndex        =   0
         Top             =   1392
         Width           =   1260
         _Version        =   196608
         _ExtentX        =   2222
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
         CharValidationText=   "1 2 3 4 5 6 7 8 9 0 "
         MaxLength       =   255
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
      Begin EditLib.fpText fptxtLastEmpNo 
         Height          =   396
         Left            =   4176
         TabIndex        =   1
         Top             =   2016
         Width           =   1260
         _Version        =   196608
         _ExtentX        =   2222
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
         CharValidationText=   "1 2 3 4 5 6 7 8 9 0 "
         MaxLength       =   255
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
      Begin EditLib.fpDateTime fptxtEndDate 
         Height          =   396
         Left            =   3648
         TabIndex        =   3
         Top             =   3264
         Width           =   1788
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
         Text            =   "11/20/2002"
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
         ThreeDFrameColor=   13684944
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
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   690
         Left            =   4320
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Press F10 to generate a report detailing employee payroll amounts."
         Top             =   5136
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
         ButtonDesigner  =   "frmEmpHistRptSplash.frx":0ED4
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
         Height          =   690
         Left            =   1392
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Press ESC to exit this screen."
         Top             =   5136
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
         ButtonDesigner  =   "frmEmpHistRptSplash.frx":10B3
      End
      Begin VB.Label Label7 
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
         Height          =   345
         Left            =   1725
         TabIndex        =   13
         Top             =   4605
         Width           =   1500
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
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
         Left            =   1632
         TabIndex        =   12
         Top             =   3360
         Width           =   1788
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Last Employee No:"
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
         TabIndex        =   11
         Top             =   2112
         Width           =   2556
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   732
         Left            =   1296
         Top             =   384
         Width           =   5100
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
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
         Left            =   1632
         TabIndex        =   10
         Top             =   2736
         Width           =   1788
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "First Employee No:"
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
         TabIndex        =   9
         Top             =   1536
         Width           =   2556
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Earnings History Report"
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
         Height          =   444
         Left            =   1296
         TabIndex        =   8
         Top             =   576
         Width           =   5100
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Summaries Only:"
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
         TabIndex        =   7
         Top             =   4032
         Width           =   2028
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   6420
      Left            =   1956
      Top             =   1224
      Width           =   7740
   End
End
Attribute VB_Name = "frmEmpHistRptSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub cmdEscape_Click()
   frmReportsProcessing.Show
   DoEvents
   Unload frmEmpHistRptSplash
   MainLog ("Employee Earnings History Report screen exited.")
End Sub
Private Sub PrintGraphics()

  Dim Emp2Rec As EmpData2Type
  ReDim Emp1Rec(1) As EmpData1Type
  ReDim TransRec(1) As TransRecType
  ReDim Unit(1) As UnitFileRecType
  ReDim EMPHIST(1 To 3) As EmpHistoryRptType
  ReDim DedCodes(1 To 50) As DedCodeRecType
  Dim DedRec As DedCodeRecType
  ReDim ErnCodes(1 To 3) As ErnCodeRecType
  ReDim EmpHistRpt(1) As EmpHistFormType

  ReDim DashLine(1) As String * 132
  Dim TotDeds(1 To 50) As Double
  ReDim TotErns(1 To 3) As Double
  Dim ESubDeds(1 To 50) As Double
  ReDim ESubErns(1 To 3) As Double
  ReDim EmpNo(1) As String * 14
  ReDim RErnP(1) As String * 11
  ReDim EICP(1) As String * 11
  ReDim GPayP(1) As String * 11
  ReDim SSTaxP(1) As String * 11
  ReDim MTaxP(1) As String * 11
  ReDim FTaxP(1) As String * 11
  ReDim STaxP(1) As String * 11
  ReDim RetirP(1) As String * 11
  ReDim NetPayP(1) As String * 11
  ReDim OErnP(1) As String * 11
  ReDim Ded(1) As String * 11
  ReDim Ern(1) As String * 11
  ReDim Pg(1) As String * 5
  ReDim Fill11(1) As String * 11
  ReDim RHrs(1) As String * 11
  ReDim VHrs(1) As String * 11
  ReDim SHrs(1) As String * 11
  ReDim HHrs(1) As String * 11
  ReDim CHrs(1) As String * 11
  ReDim THrs(1) As String * 11

  ReDim PHrs(1) As String * 11
  ReDim OTPaid(1) As String * 11
  ReDim EICP(1) As String * 11
  ReDim RErnP(1) As String * 11
  ReDim EChkDate(1) As String * 11
  ReDim EChkNo(1) As String * 11
  
  Dim Emp2RecLen As Integer, UnitFileRec As UnitFileRecType
  ReDim ErnCodes(1 To 3) As ErnCodeRecType

  ReDim TFedGrs(1) As String * 11
  ReDim TStaGrs(1) As String * 11
  ReDim TSocGrs(1) As String * 11
  ReDim TMedGrs(1) As String * 11
  ReDim TRetGrs(1) As String * 11
  Dim Image2 As String, Image3 As String
  Dim UnitHandle As Integer
  Dim City As String
  Dim ErnCodeFileHandle%, DedCodeFileHandle%, x%, cnt%, LastDed%, LastErn%
  Dim DTitle$(1 To 5), TDed$, ETitle$, TErn$, SumHeader2$
  Dim FirstEmp&, LastEmp&, LowDate%, HiDate%
  Dim EmpRecSize%, TransRecLen%, LineCnt%, MaxLines%, Page%, IdxRecLen%
  Dim NumOfRecs%, EmpIdxLNameHandle%, Emp1RecLen%, EHandle1%
  Dim IdxFileSize&, Today$, FromToDate$, SumFlag%
  Dim RptTitle$, RHandle%
  Dim EmpHistoryRpt$, UsingThisOne As Boolean
  Dim THandle%, DHandle%, RecNo%
  Dim EmpHistHeader As Boolean
  Dim TTaxFring#, TFedGross#, TStaGross#, TSocGross#, TMedGross#, TRetGross#
  Dim TaxFring#, FedGross#, STAGROSS#, SocGross#, MedGross#, RETGROSS#, DAmt#
  Dim TransRecNum&, SalCnt%, HrlCnt%
  Dim FF$, SumDed$(1 To 5), Cnt2%, SumErn$, Nextx As Integer
  Dim tripCnt As Integer
  Dim NumOfDeds As Integer
  Dim DLines As Integer
  Dim dlm$, ChangeFld$
  Dim ErnDsc$(1 To 3), DedDsc$(1 To 50)
  Dim ErnDet#(1 To 3), DedDet#(1 To 50)
  Dim ETTaxFringe#, EmpSumFlag As Boolean
  '-------------------01/04------------------
  Dim FEDTAX As FederalTaxRecType
  Dim FedTaxHandle As Integer
  Dim FedSSMax As Double
  Dim ThisDate As Integer
  Dim TotSocGross As Double
  Dim BegDate$
  Dim SSMaxReachedFlag As Boolean
  Dim ThisDif As Double
  Dim NextDate As Integer
  Dim StopDate As Integer
  Dim SSMaxCode As Integer
  Dim StopDateFlag As Boolean
  Dim z As Integer
  Dim SSTotal As Double
  Dim ThisCnt As Integer
'  Dim One As Integer
'  Dim AHandle As Integer
'
'  One = 1
'  AHandle = FreeFile
'  Open "EHist.dat" For Output As AHandle
  
  On Error GoTo ErrorHandler
  
  OpenFedTaxFile FedTaxHandle
  Get FedTaxHandle, 1, FEDTAX
  Close FedTaxHandle
  FedSSMax = FEDTAX.FTMSSMW
  dlm$ = "~"
  FirstEmp& = Val(fptxtFirstEmpNo.Text)
  LastEmp& = Val(fptxtLastEmpNo.Text)
  
  If CheckValDate(fptxtStartDate.Text) = False Then
    MsgBox "The Start Date is not valid"
    fptxtStartDate.SetFocus
    Exit Sub
  End If
  
  If CheckValDate(fptxtEndDate.Text) = False Then
     MsgBox "The End Date is not valid"
     fptxtEndDate.SetFocus
     Exit Sub
  End If
  
  LowDate = Date2Num(fptxtStartDate.Text)
  HiDate = Date2Num(fptxtEndDate.Text)
  
  If LowDate > HiDate Then
     MsgBox "The Start Date is later than the End Date"
     fptxtStartDate.SetFocus
     Exit Sub
  End If
  
  If InStr("Yy", QPTrim$(fptxtSummary.Text)) Then
    SumFlag = True
  Else
    SumFlag = False
  End If
  
  If HiDate < LowDate Then
    MsgBox "The Ending Date is earlier than the Starting Date"
    fptxtStartDate.SetFocus
    GoTo EndTran
  End If
  
  If LastEmp& < FirstEmp& Then
    MsgBox "The Last Employee Number is less than the First Employee Number."
    fptxtFirstEmpNo.SetFocus
    GoTo EndTran
  End If
  If fptxtStartDate.Text = "" Then
     MsgBox "Please enter a Starting Date"
     fptxtStartDate.SetFocus
     GoTo EndTran
  End If

  If fptxtEndDate.Text = "" Then
     MsgBox "Please enter an Ending Date"
     fptxtEndDate.SetFocus
     GoTo EndTran
  End If
  
  Image2$ = "###0.00"
  Image3$ = "#######0.00"

  OpenUnitFile UnitHandle
  Get UnitHandle, 1, UnitFileRec
  City = QPTrim$(UnitFileRec.UFEMPR)
  Close UnitHandle
    
  OpenDedCodeFile DedCodeFileHandle
  For x = 1 To 50
    Get DedCodeFileHandle, x, DedRec
    DedDsc(x) = QPTrim$(DedRec.DCDESC1)
    If Len(QPTrim$(DedRec.DCDESC1)) > 0 Then
      DedCodes(x) = DedRec
      NumOfDeds = NumOfDeds + 1
    End If
  Next x
  Close DedCodeFileHandle
  
  OpenErnCodeFile ErnCodeFileHandle
  For x = 1 To 3
     Get ErnCodeFileHandle, x, ErnCodes(x)
  Next
  Close ErnCodeFileHandle
  
  For cnt = 1 To 3
    ErnDsc(cnt) = QPTrim$(ErnCodes(cnt).ERNCODE1)
  Next
  
  EmpRecSize = Len(Emp2Rec)
  TransRecLen = Len(TransRec(1))

  OpenEmpIdxNNameFile EmpIdxLNameHandle
  NumOfRecs = LOF(EmpIdxLNameHandle) / 2
  If NumOfRecs = 0 Then
    MsgBox "No records on file"
    Close
    Exit Sub
    MainLog ("Employee Earnings History Report screen exited with no records on file.")
  End If
  
  FrmShowPctComp.Label1 = "Employee Earnings History Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
  
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  
  'load ThisSort with employee list in alphabetical order
  For x = 1 To NumOfRecs
    Get #EmpIdxLNameHandle, x, IdxBuff(x)
  Next x
  Close EmpIdxLNameHandle

  Emp1RecLen = Len(Emp1Rec(1))
  
  OpenEmpData1File EHandle1
  Get EHandle1, IdxBuff(1), Emp1Rec(1)
  EmpHistRpt(1).FirstEmp& = Val(Emp1Rec(1).EmpNo)
  Get EHandle1, IdxBuff(NumOfRecs), Emp1Rec(1)
  EmpHistRpt(1).LastEmp& = Val(Emp1Rec(1).EmpNo)
  Close EHandle1
  
  Today$ = Date$
  
  FromToDate$ = "Report Date: " + QPTrim$(fptxtStartDate.Text) + " to " + QPTrim$(fptxtEndDate.Text)
  RptTitle$ = "Employee Earnings History Report"
  If SumFlag = False Then
    EmpHistoryRpt = "PRRPTS\EMPHISTG.RPT"
  Else
    EmpHistoryRpt = "PRRPTS\EMPHISTSUMG.RPT"
  End If
  
  RHandle = FreeFile
  Open EmpHistoryRpt For Output As RHandle
  
  OpenTransHistFile THandle
  OpenEmpData2File DHandle
  EmpHistHeader = False

  For RecNo = 1 To NumOfRecs
    UsingThisOne = False
    DAmt# = 0
    Get DHandle, IdxBuff(RecNo), Emp2Rec
    If Val(Emp2Rec.EmpNo) >= FirstEmp& And Val(Emp2Rec.EmpNo) <= LastEmp& Then
    'if employee number is in range
      If Emp2Rec.LastTransRec > 0 Then         'if there are any
        TransRecNum& = Emp2Rec.LastTransRec
      Else
        GoTo Skip2NextEmp
      End If
'---------------New as of 01/12/04----------------
      SSMaxReachedFlag = False
      ThisDif = 0
      SSMaxCode = 1

      If Mid(fptxtStartDate.Text, 7, 4) = Mid(fptxtEndDate.Text, 7, 4) Then
        BegDate = ("01/01/" + Mid(fptxtEndDate.Text, 7, 4))
        ThisDate = Date2Num(BegDate)
        TotSocGross = 0
        Do
          Get THandle, TransRecNum&, TransRec(1)
          If (TransRec(1).CheckDate >= ThisDate) And (TransRec(1).CheckDate <= HiDate) Then
            TotSocGross = OldRound(TotSocGross + TransRec(1).SocGrossPay + TransRec(1).TaxFring)
          End If
          If TransRec(1).PrevTransRec > 0 Then
            TransRecNum& = TransRec(1).PrevTransRec
          Else
            Exit Do
          End If
        Loop
        If TotSocGross > FedSSMax Then
          SSMaxReachedFlag = True
          GoSub MaxSSWage
        End If
      End If
      TransRecNum& = Emp2Rec.LastTransRec
'---------------New as of 01/12/04----------------
     
      Do
        Get THandle, TransRecNum&, TransRec(1)
'        If TransRecNum = 18112 Then Stop
        
        If (TransRec(1).CheckDate >= LowDate) And (TransRec(1).CheckDate <= HiDate) Then
        'if this is in the date range
          UsingThisOne = True
          GoSub PrintAndSumEmp
        End If
          If TransRec(1).PrevTransRec > 0 Then
            TransRecNum& = TransRec(1).PrevTransRec
          Else
            If UsingThisOne And SumFlag Then
              GoSub PrintSubOnly
              Exit Do
            Else
              GoTo Skip2NextEmp
            End If
          End If
      Loop
      
    End If
Skip2NextEmp:
'    Print #AHandle, CStr(Emp2Rec.EmpNo) + "~" + Using$("$###,##0.00", FedGross)
    FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      
      GoTo AbortExitHistRpt
    End If
    
    EMPHIST(1) = EMPHIST(2) 'clear EMPHIST(1)
    ReDim ESubErns(1 To 3) As Double

    TaxFring# = 0
    FedGross# = 0
    STAGROSS# = 0
    SocGross# = 0
    MedGross# = 0
    RETGROSS# = 0
    
    For Cnt2 = 1 To NumOfDeds 'LastDed changed 4/30
       ESubDeds(Cnt2) = 0
    Next Cnt2
    
  Next
  
  Close
  
  If ThisCnt = 0 Then
    MsgBox "There are no records that fit the parameters entered."
    EnableCloseButton Me.hwnd, True
    Me.cmdEscape.Enabled = True
    Me.cmdProcess.Enabled = True
    Exit Sub
  End If
  
  If SumFlag = False Then
    arEarningsHistory.Show
    frmLoadingRpt.Show
  Else
    arEarnHistSumOnly.Show
    frmLoadingRpt.Show
  End If
  
AbortExitHistRpt:
  EnableCloseButton Me.hwnd, True
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  MainLog ("Employee Earnings History Report was processed.")
Exit Sub
  
  
PrintAndSumEmp:
  ThisCnt = ThisCnt + 1
  ChangeFld = Emp2Rec.EmpNo
  'employee summary data
  EMPHIST(1).RegHrs = OldRound#(EMPHIST(1).RegHrs + TransRec(1).RegHrsWork)
  EMPHIST(3).RegHrs = OldRound(EMPHIST(3).RegHrs + TransRec(1).RegHrsWork)
  EMPHIST(1).VACHRS = OldRound#(EMPHIST(1).VACHRS + TransRec(1).VacUsed)
  EMPHIST(3).VACHRS = OldRound(EMPHIST(3).VACHRS + TransRec(1).VacUsed)
  EMPHIST(1).SICKHRS = OldRound#(EMPHIST(1).SICKHRS + TransRec(1).SickUsed)
  EMPHIST(3).SICKHRS = OldRound(EMPHIST(3).SICKHRS + TransRec(1).SickUsed)
  EMPHIST(1).HOLHRS = OldRound#(EMPHIST(1).HOLHRS + TransRec(1).HOLHOURS)
  EMPHIST(3).HOLHRS = OldRound(EMPHIST(3).HOLHRS + TransRec(1).HOLHOURS)
  EMPHIST(1).COMPHRS = OldRound#(EMPHIST(1).COMPHRS + TransRec(1).CompUsed)
  EMPHIST(3).COMPHRS = OldRound(EMPHIST(3).COMPHRS + TransRec(1).CompUsed)
  EMPHIST(1).TotalHrs = OldRound(EMPHIST(1).TotalHrs + TransRec(1).RegHrsWork + TransRec(1).VacUsed + TransRec(1).SickUsed + TransRec(1).HOLHOURS + TransRec(1).CompUsed)
  EMPHIST(1).TotalHrs = OldRound(EMPHIST(1).TotalHrs + TransRec(1).PerHours)
  EMPHIST(3).TotalHrs = OldRound(EMPHIST(3).TotalHrs + TransRec(1).RegHrsWork + TransRec(1).VacUsed + TransRec(1).SickUsed + TransRec(1).HOLHOURS + TransRec(1).CompUsed)
  EMPHIST(3).TotalHrs = OldRound(EMPHIST(3).TotalHrs + TransRec(1).PerHours)
  
  EMPHIST(1).PHrs = OldRound#(EMPHIST(1).PHrs + TransRec(1).PerHours)
  EMPHIST(3).PHrs = OldRound(EMPHIST(3).PHrs + TransRec(1).PerHours)

  EMPHIST(1).TOTPaid = OldRound#(EMPHIST(1).TOTPaid + TransRec(1).OTHrsPaid)
  EMPHIST(3).TOTPaid = OldRound(EMPHIST(3).TOTPaid + TransRec(1).OTHrsPaid)
  EMPHIST(1).TotEIC = OldRound#(EMPHIST(1).TotEIC + TransRec(1).EICAmt)
  EMPHIST(3).TotEIC = OldRound(EMPHIST(3).TotEIC + TransRec(1).EICAmt)
  
  EMPHIST(1).TRegWage = OldRound#(EMPHIST(1).TRegWage + TransRec(1).TotRegWage)
  EMPHIST(3).TRegWage = OldRound(EMPHIST(3).TRegWage + TransRec(1).TotRegWage)
  EMPHIST(1).TOTWage = OldRound#(EMPHIST(1).TOTWage + TransRec(1).TotOTWage)
  EMPHIST(3).TOTWage = OldRound(EMPHIST(3).TOTWage + TransRec(1).TotOTWage)
  
  EMPHIST(1).GPay = OldRound#(EMPHIST(1).GPay + TransRec(1).GrossPay)
  EMPHIST(3).GPay = OldRound(EMPHIST(3).GPay + TransRec(1).GrossPay)
  EMPHIST(1).SSTax = OldRound#(EMPHIST(1).SSTax + TransRec(1).SocTaxAmt)
  EMPHIST(3).SSTax = OldRound(EMPHIST(3).SSTax + TransRec(1).SocTaxAmt)
  EMPHIST(1).MTax = OldRound#(EMPHIST(1).MTax + TransRec(1).MedTaxAmt)
  EMPHIST(3).MTax = OldRound(EMPHIST(3).MTax + TransRec(1).MedTaxAmt)
  EMPHIST(1).FTax = OldRound#(EMPHIST(1).FTax + TransRec(1).FedTaxAmt)
  EMPHIST(3).FTax = OldRound(EMPHIST(3).FTax + TransRec(1).FedTaxAmt)
  EMPHIST(1).STax = OldRound#(EMPHIST(1).STax + TransRec(1).StaTaxAmt)
  EMPHIST(3).STax = OldRound(EMPHIST(3).STax + TransRec(1).StaTaxAmt)

  TaxFring# = OldRound#(TaxFring# + TransRec(1).TaxFring)
  
  FedGross# = OldRound#(FedGross# + TransRec(1).FedGrossPay)
  STAGROSS# = OldRound#(STAGROSS# + TransRec(1).StaGrossPay)
'  SocGross# = OldRound#(SocGross# + TransRec(1).SocGrossPay)
  MedGross# = OldRound#(MedGross# + TransRec(1).MedGrossPay)
  RETGROSS# = OldRound#(RETGROSS# + TransRec(1).RetGrossPay)
  If TransRec(1).TaxFring <> 0 Then '12/30/04 changed from > to <>
    EMPHIST(1).GPay = OldRound(EMPHIST(1).GPay + TransRec(1).TaxFring) '12/30/04
    EMPHIST(3).GPay = OldRound(EMPHIST(3).GPay + TransRec(1).TaxFring) '12/30/04
    FedGross# = OldRound#(FedGross# + TransRec(1).TaxFring)
    STAGROSS# = OldRound#(STAGROSS# + TransRec(1).TaxFring)
'    SocGross# = OldRound#(SocGross# + TransRec(1).TaxFring)
    MedGross# = OldRound#(MedGross# + TransRec(1).TaxFring)
'    RETGROSS# = OldRound#(RETGROSS# + TransRec(1).TaxFring)'12/30/04
  End If
  TTaxFring# = OldRound#(TTaxFring# + TransRec(1).TaxFring)
  TFedGross# = OldRound#(TFedGross# + TransRec(1).FedGrossPay + TransRec(1).TaxFring)
  TStaGross# = OldRound#(TStaGross# + TransRec(1).StaGrossPay + TransRec(1).TaxFring)
  
  '------------New as of 01/12/2004-----------------
  If SSMaxReachedFlag = False Then
    SocGross# = OldRound#(SocGross# + TransRec(1).SocGrossPay + TransRec(1).TaxFring)
    TSocGross# = OldRound#(TSocGross# + TransRec(1).SocGrossPay + TransRec(1).TaxFring)
  End If
    
  '---------^^^New as of 01/12/2004^^^--------------
  
  TMedGross# = OldRound#(TMedGross# + TransRec(1).MedGrossPay + TransRec(1).TaxFring)
  TRetGross# = OldRound#(TRetGross# + TransRec(1).RetGrossPay) ' + TransRec(1).TaxFring) 'commented out on 12/30/04

  EMPHIST(1).RETTOT = OldRound(EMPHIST(1).RETTOT + TransRec(1).RetireAmt)
  EMPHIST(3).RETTOT = OldRound(EMPHIST(3).RETTOT + TransRec(1).RetireAmt)
  
  EMPHIST(1).TNetPay = OldRound#(EMPHIST(1).TNetPay + TransRec(1).NetPay)
  EMPHIST(3).TNetPay = OldRound(EMPHIST(3).TNetPay + TransRec(1).NetPay) 'EMPHIST(1).TNetPay)

  Select Case TransRec(1).PayType
  Case "S"
    RSet RHrs(1) = "Salaried"
    SalCnt = SalCnt + 1
  Case Else
    HrlCnt = HrlCnt + 1
  End Select
  
  For Cnt2 = 1 To NumOfDeds 'LastDed changed 4/30
    ESubDeds(Cnt2) = OldRound#(ESubDeds(Cnt2) + TransRec(1).DAmt(Cnt2))
    TotDeds(Cnt2) = OldRound#(TotDeds(Cnt2) + TransRec(1).DAmt(Cnt2))
  Next
  
  For Cnt2 = 1 To 3 'LastErn 9/6
    ESubErns(Cnt2) = OldRound#(ESubErns(Cnt2) + TransRec(1).EAmt(Cnt2))
    TotErns(Cnt2) = OldRound#(TotErns(Cnt2) + TransRec(1).EAmt(Cnt2))
    RSet Ern(1) = Using(Image3$, TransRec(1).EAmt(Cnt2))
    ErnDet(Cnt2) = Ern(1)
  Next
  
  If SumFlag = True Then Return
  '                   0             1            2                          3                         4                                  5
  Print #RHandle, ChangeFld; dlm; City; dlm; fptxtStartDate.Text; dlm; fptxtEndDate.Text; dlm; QPTrim$(Emp2Rec.EmpNo); dlm; QPTrim$(Emp2Rec.EmpLName) + ", " + QPTrim$(Emp2Rec.EmpFName);
  '                        6              7               8                                                           16
  Print #RHandle, dlm; ErnDsc(3); dlm; ErnDsc(2); dlm; ErnDsc(1);
  
  For x = 1 To 50
    Print #RHandle, dlm; DedDsc(x);
  Next x
  
  '                               59                                       60                                   61
  Print #RHandle, dlm; MakeRegDate(TransRec(1).CheckDate); dlm; Str$(TransRec(1).CheckNum); dlm; Using(Image3$, TransRec(1).TaxFring); dlm;
  '                             62                                           63                                           64                                   65                                           66
  Print #RHandle, Using(Image2$, TransRec(1).RegHrsWork); dlm; Using(Image2$, TransRec(1).VacUsed); dlm; Using(Image2$, TransRec(1).SickUsed); dlm; Using(Image2$, TransRec(1).HOLHOURS); dlm; Using(Image2$, TransRec(1).CompUsed);
  '                                        67                                68                                           69                                        70
  Print #RHandle, dlm; Using(Image2$, TransRec(1).PerHours); dlm; Using(Image2$, TransRec(1).RegHrsPaid); dlm; Using(Image2$, TransRec(1).OTHrsPaid); dlm; Using(Image2$, TransRec(1).EICAmt); dlm;
  '                            71                                           72                                         73                              74                            75                                 76                                    77
  Print #RHandle, Using(Image3$, TransRec(1).TotRegWage); dlm; Using(Image3$, TransRec(1).TotOTWage); dlm; Using(Image3, ErnDet(3)); dlm; Using(Image3, ErnDet(2)); dlm; Using(Image3, ErnDet(1)); dlm; Using(Image3$, TransRec(1).GrossPay + TransRec(1).TaxFring); dlm; Using(Image3$, TransRec(1).SocTaxAmt); '12/30/04 added tax fring to #76
  '                            78                                                79                                         80                                               81
  Print #RHandle, dlm; Using(Image3$, TransRec(1).MedTaxAmt); dlm; Using(Image3$, TransRec(1).FedTaxAmt); dlm; Using(Image3$, TransRec(1).StaTaxAmt); dlm; Using(Image3$, TransRec(1).RetireAmt); dlm;
  '                           82
  Print #RHandle, Using(Image3$, TransRec(1).NetPay); dlm;
  
  For x = 1 To 50
    If Len(DedDsc(x)) = 0 Then
      Image3$ = ""
    End If
    '                           83 - 132
    Print #RHandle, Using(Image3$, TransRec(1).DAmt(x)); dlm;
    Image3$ = "#######0.00" 'resets Image3 to accommodate
    'if Image3 was set to ""
  Next x
  '                               133                               134                                                  135                                            136
  Print #RHandle, Using("#######0.00", TaxFring#); dlm; Using("#######0.00", EMPHIST(1).RegHrs); dlm; Using("#######0.00", EMPHIST(1).VACHRS); dlm; Using("#######0.00", EMPHIST(1).SICKHRS); dlm;
  '                               137                                              138                                             139                                       140
  Print #RHandle, Using("#######0.00", EMPHIST(1).HOLHRS); dlm; Using("#######0.00", EMPHIST(1).COMPHRS); dlm; Using("#######0.00", EMPHIST(1).PHrs); dlm; Using("#######0.00", EMPHIST(1).TotalHrs); dlm;
  '                 141              142             143                            144                                               145
  Print #RHandle, ErnDsc(3); dlm; ErnDsc(2); dlm; ErnDsc(1); dlm; Using("#######0.00", EMPHIST(1).TOTPaid); dlm; Using("#######0.00", EMPHIST(1).TotEIC); dlm;
  '                                     146                                      147                                           148                                     149                                        150
  Print #RHandle, Using("#######0.00", EMPHIST(1).TRegWage); dlm; Using("#######0.00", EMPHIST(1).TOTWage); dlm; Using("#######0.00", ESubErns(3)); dlm; Using("#######0.00", ESubErns(2)); dlm; Using("#######0.00", ESubErns(1)); dlm;
  '                                     151                                       152                                        153                          154             155             156
  Print #RHandle, Using("#######0.00", EMPHIST(1).GPay); dlm; Using("#######0.00", EMPHIST(1).SSTax); dlm; Using("#######0.00", EMPHIST(1).MTax); dlm; DedDsc(1); dlm; DedDsc(2); dlm; DedDsc(3); dlm;
  '                 157              158                         159                                             160                                          161
  Print #RHandle, DedDsc(4); dlm; DedDsc(5); dlm; Using("#######0.00", EMPHIST(1).FTax); dlm; Using("#######0.00", EMPHIST(1).STax); dlm; Using("#######0.00", EMPHIST(1).RETTOT); dlm;
  '                               162
  Print #RHandle, Using("#######0.00", EMPHIST(1).TNetPay); dlm;
  '
  For x = 1 To 5
    If x <= NumOfDeds Then
      '                        163 - 167
      Print #RHandle, Using("#######0.00", ESubDeds(x)); dlm;
    Else
      '            163 - 167
      Print #RHandle, ""; dlm;
    End If
  Next x
  '
  For x = 6 To 15
    '          168 - 177
    Print #RHandle, DedDsc(x); dlm;
  Next x
  '
  For x = 6 To 15
    If x <= NumOfDeds Then
      '                        178 - 187
      Print #RHandle, Using("#######0.00", ESubDeds(x)); dlm;
    Else
      '            178 - 187
      Print #RHandle, ""; dlm;
    End If
  Next x
  
  For x = 16 To 25
    '          188 - 197
    Print #RHandle, DedDsc(x); dlm;
  Next x
  '
  For x = 16 To 25
    If x <= NumOfDeds Then
      '                        198 - 207
      Print #RHandle, Using("#######0.00", ESubDeds(x)); dlm;
    Else
      '            198 - 207
      Print #RHandle, ""; dlm;
    End If
  Next x
  
  For x = 26 To 35
    '          208 - 217
    Print #RHandle, DedDsc(x); dlm;
  Next x
  '
  For x = 26 To 35
    If x <= NumOfDeds Then
      '                        218 - 227
      Print #RHandle, Using("#######0.00", ESubDeds(x)); dlm;
    Else
      '            218 - 227
      Print #RHandle, ""; dlm;
    End If
  Next x

  For x = 36 To 45
    '          228 - 237
    Print #RHandle, DedDsc(x); dlm;
  Next x
  '
  For x = 36 To 45
    If x <= NumOfDeds Then
      '                        238 - 247
      Print #RHandle, Using("#######0.00", ESubDeds(x)); dlm;
    Else
      '            238 - 247
      Print #RHandle, ""; dlm;
    End If
  Next x
  '
  For x = 46 To 50
    '          248 - 252
    Print #RHandle, DedDsc(x); dlm;
  Next x
  '
  For x = 46 To 50
    If x <= NumOfDeds Then
      '                        253 - 257
      Print #RHandle, Using("#######0.00", ESubDeds(x)); dlm;
    Else
      '            255 - 257
      Print #RHandle, ""; dlm;
    End If
  Next x
  '                               258
  Print #RHandle, Using("#######0.00", FedGross#); dlm;
  '                               259
  Print #RHandle, Using("#######0.00", STAGROSS#); dlm;
  '                               260
  Print #RHandle, Using("#######0.00", SocGross#); dlm;
  '                               261
  Print #RHandle, Using("#######0.00", MedGross#); dlm;
  '                               262
  Print #RHandle, Using("#######0.00", RETGROSS#); dlm;
  '                      263                               264
  Print #RHandle, QPTrim$(Emp2Rec.EmpFName); dlm; QPTrim$(Emp2Rec.EmpLName); dlm;
  '                         265                            266                                              267
  Print #RHandle, Using(Image3$, TTaxFring#); dlm; Using(Image3$, EMPHIST(3).RegHrs); dlm; Using(Image3$, EMPHIST(3).VACHRS); dlm;
  '                           268                                     269                                        270
  Print #RHandle, Using(Image3$, EMPHIST(3).SICKHRS); dlm; Using(Image3$, EMPHIST(3).HOLHRS); dlm; Using(Image3$, EMPHIST(3).COMPHRS); dlm;
  '                           271                                  272                              273              274            275
  Print #RHandle, Using(Image3$, EMPHIST(3).PHrs); dlm; Using(Image3$, EMPHIST(3).TotalHrs); dlm; ErnDsc(3); dlm; ErnDsc(2); dlm; ErnDsc(1); dlm;
  '                           276                                        277                                     278
  Print #RHandle, Using(Image3$, EMPHIST(3).TOTPaid); dlm; Using(Image3$, EMPHIST(3).TotEIC); dlm; Using(Image3$, EMPHIST(3).TRegWage); dlm;
  '                           279                                        280                              281                             282
  Print #RHandle, Using(Image3$, EMPHIST(3).TOTWage); dlm; Using(Image3$, TotErns(3)); dlm; Using(Image3$, TotErns(2)); dlm; Using(Image3$, TotErns(1)); dlm;
  '                            283                                   284                                    285                        286              287
  Print #RHandle, Using(Image3$, EMPHIST(3).GPay); dlm; Using(Image3$, EMPHIST(3).SSTax); dlm; Using(Image3$, EMPHIST(3).MTax); dlm; DedDsc(1); dlm; DedDsc(2); dlm;
  '                 288             289             290                     291                                      292                                   293
  Print #RHandle, DedDsc(3); dlm; DedDsc(4); dlm; DedDsc(5); dlm; Using(Image3$, EMPHIST(3).FTax); dlm; Using(Image3$, EMPHIST(3).STax); dlm; Using(Image3$, EMPHIST(3).RETTOT); dlm;
  '                           294
  Print #RHandle, Using(Image3$, EMPHIST(3).TNetPay); dlm;
  '
  For x = 1 To 5
    If x <= NumOfDeds Then
      '                        295 - 299
      Print #RHandle, Using(Image3$, TotDeds(x)); dlm;
    Else
      '            295 - 299
      Print #RHandle, ""; dlm;
    End If
  Next x
  '
  For x = 6 To 15
    '          300 - 309
    Print #RHandle, DedDsc(x); dlm;
  Next x
  '
  For x = 6 To 15
    If x <= NumOfDeds Then
      '                        310 - 319
      Print #RHandle, Using(Image3$, TotDeds(x)); dlm;
    Else
      '            300 - 319
      Print #RHandle, ""; dlm;
    End If
  Next x
  
  For x = 16 To 25
    '          320 - 329
    Print #RHandle, DedDsc(x); dlm;
  Next x
  '
  For x = 16 To 25
    If x <= NumOfDeds Then
      '                        330 - 339
      Print #RHandle, Using(Image3$, TotDeds(x)); dlm;
    Else
      '            330 - 339
      Print #RHandle, ""; dlm;
    End If
  Next x
  
  For x = 26 To 35
    '          340 - 349
    Print #RHandle, DedDsc(x); dlm;
  Next x
  '
  For x = 26 To 35
    If x <= NumOfDeds Then
      '                        350 - 359
      Print #RHandle, Using(Image3$, TotDeds(x)); dlm;
    Else
      '            350 - 359
      Print #RHandle, ""; dlm;
    End If
  Next x

  For x = 36 To 45
    '          360 - 369
    Print #RHandle, DedDsc(x); dlm;
  Next x
  '
  For x = 36 To 45
    If x <= NumOfDeds Then
      '                        370 - 379
      Print #RHandle, Using(Image3$, TotDeds(x)); dlm;
    Else
      '            370 - 379
      Print #RHandle, ""; dlm;
    End If
  Next x
  '
  For x = 46 To 50
    '          380 - 384
    Print #RHandle, DedDsc(x); dlm;
  Next x
  '
  For x = 46 To 50
    If x <= NumOfDeds Then
      '                        385 - 389
      Print #RHandle, Using(Image3$, TotDeds(x)); dlm;
    Else
      '            385 - 389
      Print #RHandle, ""; dlm;
    End If
  Next x
  '                               390
  Print #RHandle, Using(Image3$, TFedGross#); dlm;
  '                               391
  Print #RHandle, Using(Image3$, TStaGross#); dlm;
  '                               392
  Print #RHandle, Using(Image3$, TSocGross#); dlm;
  '                               393
  Print #RHandle, Using(Image3$, TMedGross#); dlm;
  '                               394
  Print #RHandle, Using(Image3$, TRetGross#)
    
Return
  
PrintSubOnly:
  '                   0             1            2                          3                         4                                  5
  Print #RHandle, ChangeFld; dlm; City; dlm; fptxtStartDate.Text; dlm; fptxtEndDate.Text; dlm; QPTrim$(Emp2Rec.EmpNo); dlm; QPTrim$(Emp2Rec.EmpFName) + "  " + QPTrim$(Emp2Rec.EmpLName);
  '                        6              7               8
  Print #RHandle, dlm; ErnDsc(3); dlm; ErnDsc(2); dlm; ErnDsc(1);

  For x = 1 To 50
    '                     9 - 58
    Print #RHandle, dlm; DedDsc(x);
  Next x
  '                             59                                           60                                       61                                        62
  Print #RHandle, dlm; Using("#######0.00", TaxFring#); dlm; Using("#######0.00", EMPHIST(1).RegHrs); dlm; Using("#######0.00", EMPHIST(1).VACHRS); dlm; Using("#######0.00", EMPHIST(1).SICKHRS); dlm;
  '                                     63                                 64                                               65                                       66
  Print #RHandle, Using("#######0.00", EMPHIST(1).HOLHRS); dlm; Using("#######0.00", EMPHIST(1).COMPHRS); dlm; Using("#######0.00", EMPHIST(1).PHrs); dlm; Using("#######0.00", EMPHIST(1).TotalHrs); dlm;
  '                            67                                                         68
  Print #RHandle, Using("#######0.00", EMPHIST(1).TOTPaid); dlm; Using("#######0.00", EMPHIST(1).TotEIC); dlm;
  '                         69                                                   70                                                71                                    72                                       73
  Print #RHandle, Using("#######0.00", EMPHIST(1).TRegWage); dlm; Using("#######0.00", EMPHIST(1).TOTWage); dlm; Using("#######0.00", ESubErns(3)); dlm; Using("#######0.00", ESubErns(2)); dlm; Using("#######0.00", ESubErns(1)); dlm;
  '                                   74                                        75                                          76
  Print #RHandle, Using("#######0.00", EMPHIST(1).GPay); dlm; Using("#######0.00", EMPHIST(1).SSTax); dlm; Using("#######0.00", EMPHIST(1).MTax); dlm; '
  '                                           77                                           78                                               79
  Print #RHandle, Using("#######0.00", EMPHIST(1).FTax); dlm; Using("#######0.00", EMPHIST(1).STax); dlm; Using("#######0.00", EMPHIST(1).RETTOT); dlm;
  '                    80
  Print #RHandle, Using("#######0.00", EMPHIST(1).TNetPay); dlm;
  '
  For x = 1 To 5
    If x <= NumOfDeds Then
      '                        81 - 85
      Print #RHandle, Using("#######0.00", ESubDeds(x)); dlm;
    Else
      '            81 - 85
      Print #RHandle, ""; dlm;
    End If
  Next x
  '
  For x = 6 To 15
    If x <= NumOfDeds Then
      '                        86 - 95
      Print #RHandle, Using("#######0.00", ESubDeds(x)); dlm;
    Else
      '            86 - 95
      Print #RHandle, ""; dlm;
    End If
  Next x
  
  For x = 16 To 25
    If x <= NumOfDeds Then
      '                        96 - 105
      Print #RHandle, Using("#######0.00", ESubDeds(x)); dlm;
    Else
      '            96 - 105
      Print #RHandle, ""; dlm;
    End If
  Next x
  
  For x = 26 To 35
    If x <= NumOfDeds Then
      '                        106 - 115
      Print #RHandle, Using("#######0.00", ESubDeds(x)); dlm;
    Else
      '            106 - 115
      Print #RHandle, ""; dlm;
    End If
  Next x

  For x = 36 To 45
    If x <= NumOfDeds Then
      '                        116 - 125
      Print #RHandle, Using("#######0.00", ESubDeds(x)); dlm;
    Else
      '            116 - 125
      Print #RHandle, ""; dlm;
    End If
  Next x
  '
  For x = 46 To 50
    If x <= NumOfDeds Then
      '                        126 - 130
      Print #RHandle, Using("#######0.00", ESubDeds(x)); dlm;
    Else
      '            126 - 130
      Print #RHandle, ""; dlm;
    End If
  Next x
  '                               131
  Print #RHandle, Using("#######0.00", FedGross#); dlm;
  '                               132
  Print #RHandle, Using("#######0.00", STAGROSS#); dlm;
  '                               133
  Print #RHandle, Using("#######0.00", SocGross#); dlm;
  '                               134
  Print #RHandle, Using("#######0.00", MedGross#); dlm;
  '                               135
  Print #RHandle, Using("#######0.00", RETGROSS#); dlm;
  '                          136                                     137                                 138
  Print #RHandle, Using(Image3$, TTaxFring#); dlm; Using(Image3$, EMPHIST(3).RegHrs); dlm; Using(Image3$, EMPHIST(3).VACHRS); dlm;
  '                           139                                     140                                       141
  Print #RHandle, Using(Image3$, EMPHIST(3).SICKHRS); dlm; Using(Image3$, EMPHIST(3).HOLHRS); dlm; Using(Image3$, EMPHIST(3).COMPHRS); dlm;
  '                           142                                 143
  Print #RHandle, Using(Image3$, EMPHIST(3).PHrs); dlm; Using(Image3$, EMPHIST(3).TotalHrs); dlm;
  '                           144                                        145                                     146
  Print #RHandle, Using(Image3$, EMPHIST(3).TOTPaid); dlm; Using(Image3$, EMPHIST(3).TotEIC); dlm; Using(Image3$, EMPHIST(3).TRegWage); dlm;
  '                           147                                       148                              149                             150
  Print #RHandle, Using(Image3$, EMPHIST(3).TOTWage); dlm; Using(Image3$, TotErns(3)); dlm; Using(Image3$, TotErns(2)); dlm; Using(Image3$, TotErns(1)); dlm;
  '                            151                                 152                                    153
  Print #RHandle, Using(Image3$, EMPHIST(3).GPay); dlm; Using(Image3$, EMPHIST(3).SSTax); dlm; Using(Image3$, EMPHIST(3).MTax); dlm;
  '                                      154                             155                               156
  Print #RHandle, Using(Image3$, EMPHIST(3).FTax); dlm; Using(Image3$, EMPHIST(3).STax); dlm; Using(Image3$, EMPHIST(3).RETTOT); dlm;
  '                           157
  Print #RHandle, Using(Image3$, EMPHIST(3).TNetPay); dlm;
  '
  For x = 1 To 5
    If x <= NumOfDeds Then
      '                        158 - 162
      Print #RHandle, Using(Image3$, TotDeds(x)); dlm;
    Else
      '            158 - 162
      Print #RHandle, ""; dlm;
    End If
  Next x
  '
  For x = 6 To 15
    If x <= NumOfDeds Then
      '                        163 - 172
      Print #RHandle, Using(Image3$, TotDeds(x)); dlm;
    Else
      '            163 - 172
      Print #RHandle, ""; dlm;
    End If
  Next x
  
  For x = 16 To 25
    If x <= NumOfDeds Then
      '                        173 - 182
      Print #RHandle, Using(Image3$, TotDeds(x)); dlm;
    Else
      '            173 - 182
      Print #RHandle, ""; dlm;
    End If
  Next x
  
  For x = 26 To 35
    If x <= NumOfDeds Then
      '                        183 - 192
      Print #RHandle, Using(Image3$, TotDeds(x)); dlm;
    Else
      '           183 - 192
      Print #RHandle, ""; dlm;
    End If
  Next x

  For x = 36 To 45
    If x <= NumOfDeds Then
      '                        193 - 202
      Print #RHandle, Using(Image3$, TotDeds(x)); dlm;
    Else
      '            193 - 202
      Print #RHandle, ""; dlm;
    End If
  Next x
  '
  For x = 46 To 50
    If x <= NumOfDeds Then
      '                        203 - 207
      Print #RHandle, Using(Image3$, TotDeds(x)); dlm;
    Else
      '            203 - 207
      Print #RHandle, ""; dlm;
    End If
  Next x
  '                               208
  Print #RHandle, Using(Image3$, TFedGross#); dlm;
  '                               209
  Print #RHandle, Using(Image3$, TStaGross#); dlm;
  '                               210
  Print #RHandle, Using(Image3$, TSocGross#); dlm;
  '                               211
  Print #RHandle, Using(Image3$, TMedGross#); dlm;
  '                               212
  Print #RHandle, Using(Image3$, TRetGross#); dlm;
  '
  Print #RHandle, "#" & Emp2Rec.EmpNo

Return
  
MaxSSWage:
  
  StopDateFlag = False
  TransRecNum& = Emp2Rec.LastTransRec
  ReDim SSTempDates(1 To 1) As Integer
  ReDim SSTempAmts(1 To 1) As Double
  Do
    Get THandle, TransRecNum&, TransRec(1)
    If (TransRec(1).CheckDate >= ThisDate) And (TransRec(1).CheckDate <= HiDate) Then
      NextDate = NextDate + 1
      ReDim Preserve SSTempDates(1 To NextDate) As Integer
      ReDim Preserve SSTempAmts(1 To NextDate) As Double
      SSTempDates(NextDate) = TransRec(1).CheckDate
      SSTempAmts(NextDate) = OldRound(TransRec(1).SocGrossPay + TransRec(1).TaxFring)
    End If
    If TransRec(1).PrevTransRec > 0 Then
      TransRecNum& = TransRec(1).PrevTransRec
    Else
      Exit Do
    End If
  Loop
    
  ReDim SSAmts(1 To NextDate) As Double
  ReDim SSDates(1 To NextDate) As Integer
  'assign dates/amts in forward order
  z = NextDate
  For x = 1 To NextDate
    SSDates(x) = SSTempDates(z)
    SSAmts(x) = SSTempAmts(z)
    z = z - 1
  Next x
  
  TotSocGross = 0
    
  For x = 1 To NextDate
    TotSocGross = OldRound(TotSocGross + SSAmts(x))
    If SSDates(x) >= LowDate Then SSTotal = SSTotal + SSAmts(x)
    If TotSocGross > FedSSMax Then
      If StopDateFlag = False Then
        StopDateFlag = True
        StopDate = SSDates(x)
        If NextDate > 1 Then
          ThisDif = OldRound(TotSocGross - FedSSMax)
          Exit For
        Else
          SSMaxCode = 4
          Exit For
        End If
      End If
    End If
  Next x
  
  If SSMaxCode = 4 Then
    SocGross# = FedSSMax
    TSocGross# = TSocGross# + FedSSMax
    Return
  End If
  
  If StopDate < LowDate Then
    SocGross# = 0
    TSocGross# = TSocGross#
  Else
    SocGross# = SSTotal - ThisDif
    TSocGross# = TSocGross# + SocGross
  End If
  
  Return
  
  
ErrorHandler:
  Close
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  Unload FrmShowPctComp
  MsgBox "ERROR: If this problem persists please consult Southern Software."

EndTran:

End Sub
Private Sub SetupHistReportForm()

   Dim EmpData1Handle As Integer, EmpIdxLNameHandle As Integer
   Dim EmpData1Rec As EmpData1Type
   Dim IdxRecPointer As Integer, NumOfRecs As Integer

   OpenEmpData1File EmpData1Handle
   OpenEmpIdxNNameFile EmpIdxLNameHandle
   NumOfRecs = LOF(EmpIdxLNameHandle) / 2
   If NumOfRecs = 0 Then
     MsgBox "No records on file."
     Close
     Exit Sub
   End If
   Get #EmpIdxLNameHandle, 1, IdxRecPointer
   Get #EmpData1Handle, IdxRecPointer, EmpData1Rec
   fptxtFirstEmpNo.Text = Val(EmpData1Rec.EmpNo)
   
   Get #EmpIdxLNameHandle, NumOfRecs, IdxRecPointer
   Get #EmpData1Handle, IdxRecPointer, EmpData1Rec
'   Stop
   fptxtLastEmpNo.Text = Val(EmpData1Rec.EmpNo)
  
   Close EmpIdxLNameHandle, EmpData1Handle
   fptxtSummary.AddItem "Y"
   fptxtSummary.AddItem "N"
   fptxtSummary.Text = "N"
   fptxtEndDate.Text = Date$
   fptxtStartDate.Text = "01-01-" + Right$(Date$, 4)
   fpcomboPrintOpt.AddItem "Graphical"
   fpcomboPrintOpt.AddItem "Text"
   fpcomboPrintOpt.Text = "Graphical"

End Sub

Private Sub cmdProcess_Click()
  If fpcomboPrintOpt.Text = "Graphical" Then
    Call PrintGraphics
    Exit Sub
  ElseIf fpcomboPrintOpt.Text = "Text" Then
    Call PrintText
  Else
    Exit Sub
  End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%C"
      Call cmdEscape_Click
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      Call cmdProcess_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  MainLog ("Employee Earnings History Report screen accessed.")
  Call SetupHistReportForm
  Me.HelpContextID = hlpEarnings
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub fptxtSummary_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fptxtSummary.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fptxtSummary.ListIndex = -1
  End If
  If fptxtSummary.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        fptxtEndDate.SetFocus
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub fptxtSummary_LostFocus()
  fptxtSummary.Action = ActionClearSearchBuffer
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdEscape.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmEmpHistRptSplash.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Function NumOfDedLines(NumOfDeds As Integer) As Integer
  Select Case NumOfDeds
  Case 0 To 12:
    NumOfDedLines = 1
  Case 13 To 24:
    NumOfDedLines = 2
  Case 25 To 36:
    NumOfDedLines = 3
  Case 37 To 48:
    NumOfDedLines = 4
  Case 49 To 50:
    NumOfDedLines = 5
  Case Else:
    NumOfDedLines = 0
  End Select

End Function

Private Sub PrintText()

  Dim Emp2Rec As EmpData2Type
  ReDim Emp1Rec(1) As EmpData1Type
  ReDim TransRec(1) As TransRecType
  ReDim Unit(1) As UnitFileRecType
  ReDim EMPHIST(1 To 3) As EmpHistoryRptType
  ReDim DedCodes(1 To 50) As DedCodeRecType
  Dim DedRec As DedCodeRecType
  ReDim ErnCodes(1 To 3) As ErnCodeRecType
  ReDim EmpHistRpt(1) As EmpHistFormType

  ReDim DashLine(1) As String * 132
  Dim TotDeds(1 To 50) As Double
  ReDim TotErns(1 To 3) As Double
  Dim ESubDeds(1 To 50) As Double
  ReDim ESubErns(1 To 3) As Double
  ReDim EmpNo(1) As String * 14
  ReDim RErnP(1) As String * 11
  ReDim EICP(1) As String * 11
  ReDim GPayP(1) As String * 11
  ReDim SSTaxP(1) As String * 11
  ReDim MTaxP(1) As String * 11
  ReDim FTaxP(1) As String * 11
  ReDim STaxP(1) As String * 11
  ReDim RetirP(1) As String * 11
  ReDim NetPayP(1) As String * 11
  ReDim OErnP(1) As String * 11
  ReDim Ded(1) As String * 11
  ReDim Ern(1) As String * 11
  ReDim Pg(1) As String * 5
  ReDim Fill11(1) As String * 11
  ReDim RHrs(1) As String * 11
  ReDim VHrs(1) As String * 11
  ReDim SHrs(1) As String * 11
  ReDim HHrs(1) As String * 11
  ReDim CHrs(1) As String * 11
  ReDim THrs(1) As String * 11

  ReDim PHrs(1) As String * 11
  ReDim OTPaid(1) As String * 11
  ReDim EICP(1) As String * 11
  ReDim RErnP(1) As String * 11
  ReDim EChkDate(1) As String * 11
  ReDim EChkNo(1) As String * 11
  
  Dim Emp2RecLen As Integer, UnitFileRec As UnitFileRecType
  ReDim ErnCodes(1 To 3) As ErnCodeRecType

  ReDim TFedGrs(1) As String * 11
  ReDim TStaGrs(1) As String * 11
  ReDim TSocGrs(1) As String * 11
  ReDim TMedGrs(1) As String * 11
  ReDim TRetGrs(1) As String * 11
  Dim Image2 As String, Image3 As String
  Dim UnitHandle As Integer
  Dim City As String
  Dim ErnCodeFileHandle%, DedCodeFileHandle%, x%, cnt%, LastDed%, LastErn%
  Dim DTitle$(1 To 5), TDed$, ETitle$, TErn$, SumHeader2$
  Dim FirstEmp&, LastEmp&, LowDate%, HiDate%
  Dim EmpRecSize%, TransRecLen%, LineCnt%, MaxLines%, Page%, IdxRecLen%
  Dim NumOfRecs%, EmpIdxLNameHandle%, Emp1RecLen%, EHandle1%
  Dim IdxFileSize&, Today$, FromToDate$, SumFlag%
  Dim RptTitle$, RHandle%
  Dim EmpHistoryRpt$, UsingThisOne As Boolean
  Dim THandle%, DHandle%, RecNo%
  Dim EmpHistHeader As Boolean
  Dim TTaxFring#, TFedGross#, TStaGross#, TSocGross#, TMedGross#, TRetGross#
  Dim TaxFring#, FedGross#, STAGROSS#, SocGross#, MedGross#, RETGROSS#, DAmt#
  Dim TransRecNum&, SalCnt%, HrlCnt%
  Dim FF$, SumDed$(1 To 5), Cnt2%, SumErn$, Nextx As Integer
  Dim tripCnt As Integer
  Dim NumOfDeds As Integer
  Dim DLines As Integer
  '-------------------01/04------------------
  Dim FEDTAX As FederalTaxRecType
  Dim FedTaxHandle As Integer
  Dim FedSSMax As Double
  Dim ThisDate As Integer
  Dim TotSocGross As Double
  Dim BegDate$
  Dim SSMaxReachedFlag As Boolean
  Dim ThisDif As Double
  Dim NextDate As Integer
  Dim StopDate As Integer
  Dim SSMaxCode As Integer
  Dim StopDateFlag As Boolean
  Dim z As Integer
  Dim SSTotal As Double
  Dim ThisCnt As Integer
  
  OpenFedTaxFile FedTaxHandle
  Get FedTaxHandle, 1, FEDTAX
  Close FedTaxHandle
  FedSSMax = FEDTAX.FTMSSMW
  
  FirstEmp& = Val(fptxtFirstEmpNo.Text)
  LastEmp& = Val(fptxtLastEmpNo.Text)
  If CheckValDate(fptxtStartDate.Text) = False Then
    MsgBox "The Start Date is not valid"
    fptxtStartDate.SetFocus
    Exit Sub
  End If
  
  If CheckValDate(fptxtEndDate.Text) = False Then
     MsgBox "The End Date is not valid"
     fptxtEndDate.SetFocus
     Exit Sub
  End If
  
  LowDate = Date2Num(fptxtStartDate.Text)
  HiDate = Date2Num(fptxtEndDate.Text)
  
  If LowDate > HiDate Then
     MsgBox "The Start Date is later than the End Date"
     fptxtStartDate.SetFocus
     Exit Sub
  End If
  
  If InStr("Yy", QPTrim$(fptxtSummary.Text)) Then
    SumFlag = True
  Else
    SumFlag = False
  End If
  
  If HiDate < LowDate Then
    MsgBox "The Ending Date is earlier than the Starting Date"
    fptxtStartDate.SetFocus
    GoTo EndTran
  End If
  
  If LastEmp& < FirstEmp& Then
    MsgBox "The Last Employee Number is less than the First Employee Number."
    fptxtFirstEmpNo.SetFocus
    GoTo EndTran
  End If
  If fptxtStartDate.Text = "" Then
     MsgBox "Please enter a Starting Date"
     fptxtStartDate.SetFocus
     GoTo EndTran
  End If

  If fptxtEndDate.Text = "" Then
     MsgBox "Please enter an Ending Date"
     fptxtEndDate.SetFocus
     GoTo EndTran
  End If
  
  FF$ = Chr$(12)
  Image2$ = "###0.00"
  Image3$ = "#######0.00"

  LSet Fill11(1) = ""
  LSet DashLine(1) = String$(132, "-")
  
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, UnitFileRec
  City = QPTrim$(UnitFileRec.UFEMPR)
  Close UnitHandle
    
  OpenDedCodeFile DedCodeFileHandle
  For x = 1 To 50
    Get DedCodeFileHandle, x, DedRec
    If Len(QPTrim$(DedRec.DCDESC1)) > 0 Then
      DedCodes(x) = DedRec
      NumOfDeds = NumOfDeds + 1
    End If
  Next x
  Close DedCodeFileHandle
  
  DLines = NumOfDedLines(NumOfDeds)
  
  OpenErnCodeFile ErnCodeFileHandle
  For x = 1 To 3
     Get ErnCodeFileHandle, x, ErnCodes(x)
  Next
  Close ErnCodeFileHandle
  
'*** Create the voluntary deduction description line
  For x = 1 To 5 'added 4/30
    DTitle$(x) = ""
  Next x
  tripCnt = 1
  Nextx = 1
  For cnt = 1 To NumOfDeds
    If tripCnt = 13 Then 'added 4/30
      tripCnt = 1
      Nextx = Nextx + 1
    End If
    TDed$ = QPTrim$(DedCodes(cnt).DCDESC1)
    If Len(TDed$) > 0 Then
      RSet Ded(1) = TDed$
      DTitle$(Nextx) = DTitle$(Nextx) + Ded(1)
    Else
      Exit For
    End If
    tripCnt = tripCnt + 1
  Next
'*** Create the alternate earnings description line
  ETitle$ = ""
  For cnt = 1 To 3
    TErn$ = QPTrim$(ErnCodes(cnt).ERNCODE1)
    If Len(TErn$) > 0 Then
      LastErn = LastErn + 1
      RSet Ern(1) = TErn$
      ETitle$ = ETitle$ + Ern(1)
    Else
      Exit For
    End If
  Next
  If LastErn < 3 Then
    ETitle$ = Space$(11 * (3 - LastErn)) + ETitle$
  End If

  SumHeader2$ = "  Reg Wages  O/T Wages" + ETitle$
  ETitle$ = "   Reg Earn   O/T Earn" + ETitle$ + "  Gross Pay    Soc Sec   Medicare        FWT        SWT     Retire    Net Pay"
  '------------------------------------------------------------------
  EmpRecSize = Len(Emp2Rec)
  TransRecLen = Len(TransRec(1))
  LineCnt = 0
  MaxLines = 50
  Page = 1

  OpenEmpIdxNNameFile EmpIdxLNameHandle
  NumOfRecs = LOF(EmpIdxLNameHandle) / 2
  If NumOfRecs = 0 Then
    MsgBox "No records on file"
    Close
    Exit Sub
    MainLog ("Employee Earnings History Report screen exited with no records on file.")
  End If
  
  FrmShowPctComp.Label1 = "Employee Earnings History Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
  
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  
  
  'load ThisSort with employee list in alphabetical order
  For x = 1 To NumOfRecs
    Get #EmpIdxLNameHandle, x, IdxBuff(x)
  Next x
  Close EmpIdxLNameHandle

  Emp1RecLen = Len(Emp1Rec(1))
  
  OpenEmpData1File EHandle1
  Get EHandle1, IdxBuff(1), Emp1Rec(1)
  EmpHistRpt(1).FirstEmp& = Val(Emp1Rec(1).EmpNo)
  Get EHandle1, IdxBuff(NumOfRecs), Emp1Rec(1)
  EmpHistRpt(1).LastEmp& = Val(Emp1Rec(1).EmpNo)
  Close EHandle1

  Today$ = Date$
  EmpHistRpt(1).SumOnly = "N"
  
  FromToDate$ = "Report Date: " + QPTrim$(fptxtStartDate.Text) + " to " + QPTrim$(fptxtEndDate.Text)
  RptTitle$ = "Employee Earnings History Report"
  EmpHistoryRpt = "EMPHIST.RPT"
  
  RHandle = FreeFile
  Open EmpHistoryRpt For Output As RHandle
  RPTSetupPRN 4, RHandle
  
  OpenTransHistFile THandle
  OpenEmpData2File DHandle
  EmpHistHeader = False

  For RecNo = 1 To NumOfRecs
    UsingThisOne = False
    If Not SumFlag Then
      EmpHistHeader = False
    End If
    DAmt# = 0
    Get DHandle, IdxBuff(RecNo), Emp2Rec
    If Val(Emp2Rec.EmpNo) >= FirstEmp& And Val(Emp2Rec.EmpNo) <= LastEmp& Then
    'if employee number is in range
      If Emp2Rec.LastTransRec > 0 Then         'if there are any
        TransRecNum& = Emp2Rec.LastTransRec
      Else
        GoTo Skip2NextEmp
      End If
      
'---------------New as of 01/12/04----------------
      SSMaxReachedFlag = False
      ThisDif = 0
      SSMaxCode = 1

      If Mid(fptxtStartDate.Text, 7, 4) = Mid(fptxtEndDate.Text, 7, 4) Then
        BegDate = ("01/01/" + Mid(fptxtEndDate.Text, 7, 4))
        ThisDate = Date2Num(BegDate)
        TotSocGross = 0
        Do
          Get THandle, TransRecNum&, TransRec(1)
          If (TransRec(1).CheckDate >= ThisDate) And (TransRec(1).CheckDate <= HiDate) Then
            TotSocGross = OldRound(TotSocGross + TransRec(1).SocGrossPay + TransRec(1).TaxFring)
          End If
          If TransRec(1).PrevTransRec > 0 Then
            TransRecNum& = TransRec(1).PrevTransRec
          Else
            Exit Do
          End If
        Loop
        If TotSocGross > FedSSMax Then
          SSMaxReachedFlag = True
          GoSub MaxSSWage
        End If
      End If
      TransRecNum& = Emp2Rec.LastTransRec
'---------------New as of 01/12/04----------------
      
      Do
        Get THandle, TransRecNum&, TransRec(1)
        If (TransRec(1).CheckDate >= LowDate) And (TransRec(1).CheckDate <= HiDate) Then
        'if this is in the date range
          UsingThisOne = True
          GoSub PrintAndSumEmp
          If LineCnt >= MaxLines Then
            Print #RHandle, FF$
            LineCnt = 0
            GoSub PrintEmpHistoryHeader
          End If
        End If 'ELSE
          If TransRec(1).PrevTransRec > 0 Then
            TransRecNum& = TransRec(1).PrevTransRec
          Else
            If UsingThisOne Then
              GoSub PrintSubTotal
              Exit Do
            Else
              GoTo Skip2NextEmp
            End If
          End If
      Loop
      EMPHIST(1) = EMPHIST(2)
      ReDim ESubErns(1 To 3) As Double

      TTaxFring# = OldRound#(TTaxFring# + TaxFring#)
      TFedGross# = OldRound#(TFedGross# + FedGross#)
      TStaGross# = OldRound#(TStaGross# + STAGROSS#)
      TSocGross# = OldRound#(TSocGross# + SocGross#)
      TMedGross# = OldRound#(TMedGross# + MedGross#)
      TRetGross# = OldRound#(TRetGross# + RETGROSS#)

      TaxFring# = 0
      FedGross# = 0
      STAGROSS# = 0
      SocGross# = 0
      MedGross# = 0
      RETGROSS# = 0

    End If

Skip2NextEmp:
    FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      
      GoTo AbortExitHistRpt
    End If
  Next
  GoSub PrintGrandTotals
  RPTSetupPRN 123, RHandle '7/24
  
  Close
  
  If ThisCnt = 0 Then
    MsgBox "There are no records that fit the parameters entered."
    EnableCloseButton Me.hwnd, True
    Me.cmdEscape.Enabled = True
    Me.cmdProcess.Enabled = True
    Exit Sub
  End If
  
  ViewPrint EmpHistoryRpt, RptTitle$, True

AbortExitHistRpt:
  EnableCloseButton Me.hwnd, True
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  MainLog ("Employee Earnings History Report was processed.")
Exit Sub
  
PrintEmpHistoryHeader:
  RSet Pg(1) = Str$(Page)
  
  Print #RHandle, City + Space$(86) + "Page:" + Pg(1)
  Print #RHandle, "Employee Earnings History Report" + Space$(63) + FromToDate$
  If SumFlag Then
    Print #RHandle, DashLine(1)
    LineCnt = 4
    Page = Page + 1 '7/25
  Else
    Print #RHandle,
    LSet EmpNo(1) = QPTrim$(Emp2Rec.EmpNo)
    Print #RHandle, EmpNo(1) + QPTrim$(Emp2Rec.EmpLName) + ", " + QPTrim$(Emp2Rec.EmpFName)
    Print #RHandle, " Trans Date   Check No  Tax Fring   Reg Hrs      Vacat       Sick        Hol       Comp    Personal      Total   O/T Paid        EIC"
    Print #RHandle, ETitle$
    For x = 1 To DLines
      Print #RHandle, DTitle$(x)
    Next x
    Print #RHandle, DashLine(1)
    LineCnt = 8 + DLines
    Page = Page + 1
  End If
  
  Return
  
PrintAndSumEmp:
  If Not EmpHistHeader Then
    EmpHistHeader = True
    GoSub PrintEmpHistoryHeader
  End If
  ThisCnt = ThisCnt + 1
  EMPHIST(1).RegHrs = OldRound#(EMPHIST(1).RegHrs + TransRec(1).RegHrsWork)
  EMPHIST(1).VACHRS = OldRound#(EMPHIST(1).VACHRS + TransRec(1).VacUsed)
  EMPHIST(1).SICKHRS = OldRound#(EMPHIST(1).SICKHRS + TransRec(1).SickUsed)
  EMPHIST(1).HOLHRS = OldRound#(EMPHIST(1).HOLHRS + TransRec(1).HOLHOURS)
  EMPHIST(1).COMPHRS = OldRound#(EMPHIST(1).COMPHRS + TransRec(1).CompUsed)
  EMPHIST(1).TotalHrs = OldRound(EMPHIST(1).TotalHrs + TransRec(1).RegHrsWork + TransRec(1).VacUsed + TransRec(1).SickUsed + TransRec(1).HOLHOURS + TransRec(1).CompUsed)
  EMPHIST(1).TotalHrs = OldRound(EMPHIST(1).TotalHrs + TransRec(1).PerHours)
  
  EMPHIST(1).PHrs = OldRound#(EMPHIST(1).PHrs + TransRec(1).PerHours)

  EMPHIST(1).TOTPaid = OldRound#(EMPHIST(1).TOTPaid + TransRec(1).OTHrsPaid)
  EMPHIST(1).TotEIC = OldRound#(EMPHIST(1).TotEIC + TransRec(1).EICAmt)
  
  EMPHIST(1).TRegWage = OldRound#(EMPHIST(1).TRegWage + TransRec(1).TotRegWage)
  EMPHIST(1).TOTWage = OldRound#(EMPHIST(1).TOTWage + TransRec(1).TotOTWage)
  
  EMPHIST(1).GPay = OldRound#(EMPHIST(1).GPay + TransRec(1).GrossPay)
  EMPHIST(1).SSTax = OldRound#(EMPHIST(1).SSTax + TransRec(1).SocTaxAmt)
  EMPHIST(1).MTax = OldRound#(EMPHIST(1).MTax + TransRec(1).MedTaxAmt)
  EMPHIST(1).FTax = OldRound#(EMPHIST(1).FTax + TransRec(1).FedTaxAmt)
  EMPHIST(1).STax = OldRound#(EMPHIST(1).STax + TransRec(1).StaTaxAmt)

  TaxFring# = OldRound#(TaxFring# + TransRec(1).TaxFring)

  FedGross# = OldRound#(FedGross# + TransRec(1).FedGrossPay)
  STAGROSS# = OldRound#(STAGROSS# + TransRec(1).StaGrossPay)
  'SocGross is tallied inside the SSMax code below
'  SocGross# = OldRound#(SocGross# + TransRec(1).SocGrossPay)
  MedGross# = OldRound#(MedGross# + TransRec(1).MedGrossPay)
  RETGROSS# = OldRound#(RETGROSS# + TransRec(1).RetGrossPay)

  If TransRec(1).TaxFring <> 0 Then '12/30/04 changed > to <>
    EMPHIST(1).GPay = OldRound(EMPHIST(1).GPay + TransRec(1).TaxFring) '12/30/04
    FedGross# = OldRound#(FedGross# + TransRec(1).TaxFring)
    STAGROSS# = OldRound#(STAGROSS# + TransRec(1).TaxFring)
'    SocGross# = OldRound#(SocGross# + TransRec(1).TaxFring)
    MedGross# = OldRound#(MedGross# + TransRec(1).TaxFring)
'    RETGROSS# = OldRound#(RETGROSS# + TransRec(1).TaxFring)'12/30/04
  End If

  '------------New as of 01/12/2004-----------------
  If SSMaxReachedFlag = False Then 'SSMaxReachedFlag is set
  'as a new employee is introduced to the process
    SocGross# = OldRound#(SocGross# + TransRec(1).SocGrossPay + TransRec(1).TaxFring)
  End If
  
  '---------^^^New as of 01/12/2004^^^--------------
  EMPHIST(1).RETTOT = OldRound(EMPHIST(1).RETTOT + TransRec(1).RetireAmt)
  
  EMPHIST(1).TNetPay = OldRound#(EMPHIST(1).TNetPay + TransRec(1).NetPay)
  
  If Not SumFlag Then
    LSet EChkDate(1) = MakeRegDate(TransRec(1).CheckDate)           'LTRIM$(EmpRec1(1).EmpNo)
    RSet EChkNo(1) = Str$(TransRec(1).CheckNum)
    RSet RHrs(1) = Using(Image2$, TransRec(1).RegHrsWork)
    RSet VHrs(1) = Using(Image2$, TransRec(1).VacUsed)
    RSet SHrs(1) = Using(Image2$, TransRec(1).SickUsed)
    RSet HHrs(1) = Using(Image2$, TransRec(1).HOLHOURS)
    RSet CHrs(1) = Using(Image2$, TransRec(1).CompUsed)
    RSet THrs(1) = Using(Image2$, TransRec(1).RegHrsPaid)
    RSet PHrs(1) = Using(Image2$, TransRec(1).PerHours)
    RSet OTPaid(1) = Using(Image2$, TransRec(1).OTHrsPaid)
    RSet EICP(1) = Using(Image2$, TransRec(1).EICAmt)
    RSet Fill11(1) = Using(Image3$, TransRec(1).TaxFring)
    RSet RErnP(1) = Using(Image3$, TransRec(1).TotRegWage)
    RSet OErnP(1) = Using(Image3$, TransRec(1).TotOTWage)
    RSet GPayP(1) = Using(Image3$, TransRec(1).GrossPay + TransRec(1).TaxFring) '12/30/04)
    RSet SSTaxP(1) = Using(Image3$, TransRec(1).SocTaxAmt)
    RSet MTaxP(1) = Using(Image3$, TransRec(1).MedTaxAmt)
    RSet FTaxP(1) = Using(Image3$, TransRec(1).FedTaxAmt)
    RSet STaxP(1) = Using(Image3$, TransRec(1).StaTaxAmt)
    RSet RetirP(1) = Using(Image3$, TransRec(1).RetireAmt)
    RSet NetPayP(1) = Using(Image3$, TransRec(1).NetPay)
  End If

  Select Case TransRec(1).PayType
  Case "S"
    RSet RHrs(1) = "Salaried"
    SalCnt = SalCnt + 1
  Case Else
    HrlCnt = HrlCnt + 1
  End Select
  
  For x = 1 To 5
    SumDed$(x) = ""
  Next x
  Nextx = 1
  tripCnt = 1
  
  For Cnt2 = 1 To NumOfDeds 'LastDed changed 4/30
    If tripCnt = 13 Then
      tripCnt = 1
      Nextx = Nextx + 1
    End If
    ESubDeds(Cnt2) = OldRound#(ESubDeds(Cnt2) + TransRec(1).DAmt(Cnt2))
    TotDeds(Cnt2) = OldRound#(TotDeds(Cnt2) + TransRec(1).DAmt(Cnt2))
    RSet Ded(1) = Using(Image3$, TransRec(1).DAmt(Cnt2))
    SumDed$(Nextx) = SumDed$(Nextx) + Ded(1)
    tripCnt = tripCnt + 1
  Next
  
  SumErn$ = ""
  For Cnt2 = 1 To LastErn
    ESubErns(Cnt2) = OldRound#(ESubErns(Cnt2) + TransRec(1).EAmt(Cnt2))
    TotErns(Cnt2) = OldRound#(TotErns(Cnt2) + TransRec(1).EAmt(Cnt2))
    RSet Ern(1) = Using(Image3$, TransRec(1).EAmt(Cnt2))
    SumErn$ = SumErn$ + Ern(1)
  Next
  If LastErn < 3 Then
    SumErn$ = Space$(11 * (3 - LastErn)) + SumErn$
  End If
  If Not SumFlag Then
    Print #RHandle, EChkDate(1); EChkNo(1); Fill11(1); RHrs(1); VHrs(1); SHrs(1); HHrs(1);
    Print #RHandle, CHrs(1) + PHrs(1) + THrs(1) + OTPaid(1) + EICP(1)
    Print #RHandle, RErnP(1); OErnP(1); SumErn$; GPayP(1); SSTaxP(1); MTaxP(1) + FTaxP(1); STaxP(1);
    Print #RHandle, RetirP(1); NetPayP(1)
    For x = 1 To DLines '(NumOfDeds / 12) + 1
      Print #RHandle, SumDed$(x)
    Next x
    Print #RHandle,
    LineCnt = LineCnt + 3 + DLines '(NumOfDeds / 12) '8 + (Nextx - 1) ' 4  changed 4/30
  End If

Return
  
PrintSubTotal:
  
  RSet THrs(1) = Using(Image3$, EMPHIST(1).TotalHrs)
  RSet RHrs(1) = Using(Image3$, EMPHIST(1).RegHrs)
  RSet VHrs(1) = Using(Image3$, EMPHIST(1).VACHRS)
  RSet SHrs(1) = Using(Image3$, EMPHIST(1).SICKHRS)
  RSet HHrs(1) = Using(Image3$, EMPHIST(1).HOLHRS)
  RSet CHrs(1) = Using(Image3$, EMPHIST(1).COMPHRS)


  RSet PHrs(1) = Using(Image3$, EMPHIST(1).PHrs)
  RSet OTPaid(1) = Using(Image3$, EMPHIST(1).TOTPaid)
  RSet EICP(1) = Using(Image3$, EMPHIST(1).TotEIC)
  
  RSet RErnP(1) = Using(Image3$, EMPHIST(1).TRegWage)
  RSet OErnP(1) = Using(Image3$, EMPHIST(1).TOTWage)
  
  RSet GPayP(1) = Using(Image3$, EMPHIST(1).GPay) ' + TaxFring#) '12/30/04)
  RSet SSTaxP(1) = Using(Image3$, EMPHIST(1).SSTax)
  RSet MTaxP(1) = Using(Image3$, EMPHIST(1).MTax)
  RSet FTaxP(1) = Using(Image3$, EMPHIST(1).FTax)
  RSet STaxP(1) = Using(Image3$, EMPHIST(1).STax)
  RSet RetirP(1) = Using(Image3$, EMPHIST(1).RETTOT)
  RSet NetPayP(1) = Using(Image3$, EMPHIST(1).TNetPay)

  RSet Fill11(1) = Using(Image3$, TaxFring#)
  RSet TFedGrs(1) = Using(Image3$, FedGross#)
  RSet TStaGrs(1) = Using(Image3$, STAGROSS#)
  RSet TSocGrs(1) = Using(Image3$, SocGross#)
  RSet TMedGrs(1) = Using(Image3$, MedGross#)
  RSet TRetGrs(1) = Using(Image3$, RETGROSS#)
  '---------------------------------------------------------------
  
  EMPHIST(3).TotalHrs = OldRound(EMPHIST(3).TotalHrs + EMPHIST(1).TotalHrs)
  EMPHIST(3).RegHrs = OldRound(EMPHIST(3).RegHrs + EMPHIST(1).RegHrs)
  EMPHIST(3).VACHRS = OldRound(EMPHIST(3).VACHRS + EMPHIST(1).VACHRS)
  EMPHIST(3).SICKHRS = OldRound(EMPHIST(3).SICKHRS + EMPHIST(1).SICKHRS)
  EMPHIST(3).HOLHRS = OldRound(EMPHIST(3).HOLHRS + EMPHIST(1).HOLHRS)
  EMPHIST(3).COMPHRS = OldRound(EMPHIST(3).COMPHRS + EMPHIST(1).COMPHRS)
  EMPHIST(3).PHrs = OldRound(EMPHIST(3).PHrs + EMPHIST(1).PHrs)
  EMPHIST(3).TOTPaid = OldRound(EMPHIST(3).TOTPaid + EMPHIST(1).TOTPaid)
  EMPHIST(3).TotEIC = OldRound(EMPHIST(3).TotEIC + EMPHIST(1).TotEIC)
  EMPHIST(3).TRegWage = OldRound(EMPHIST(3).TRegWage + EMPHIST(1).TRegWage)
  EMPHIST(3).TOTWage = OldRound(EMPHIST(3).TOTWage + EMPHIST(1).TOTWage)
  EMPHIST(3).GPay = OldRound(EMPHIST(3).GPay + EMPHIST(1).GPay)
  EMPHIST(3).SSTax = OldRound(EMPHIST(3).SSTax + EMPHIST(1).SSTax)
  EMPHIST(3).MTax = OldRound(EMPHIST(3).MTax + EMPHIST(1).MTax)
  EMPHIST(3).FTax = OldRound(EMPHIST(3).FTax + EMPHIST(1).FTax)
  EMPHIST(3).STax = OldRound(EMPHIST(3).STax + EMPHIST(1).STax)
  EMPHIST(3).RETTOT = OldRound(EMPHIST(3).RETTOT + EMPHIST(1).RETTOT)
  EMPHIST(3).TNetPay = OldRound(EMPHIST(3).TNetPay + EMPHIST(1).TNetPay)
  '---------------------------------------------------------------
  For x = 1 To 5 'added 4/30
    SumDed$(x) = ""
  Next x
  Nextx = 1
  tripCnt = 1
  For Cnt2 = 1 To NumOfDeds ' LastDed
    If tripCnt = 13 Then 'added 4/30
      tripCnt = 1
      Nextx = Nextx + 1
    End If
    RSet Ded(1) = Using(Image3$, ESubDeds(Cnt2))
    SumDed$(Nextx) = SumDed$(Nextx) + Ded(1)
    tripCnt = tripCnt + 1
  Next
  
  For x = 1 To NumOfDeds
    ESubDeds(x) = 0
  Next x
  '---------------------------------------------------------
  SumErn$ = ""
  For Cnt2 = 1 To LastErn
    RSet Ern(1) = Using(Image3$, ESubErns(Cnt2))
    SumErn$ = SumErn$ + Ern(1)
  Next
  If LastErn < 3 Then
    SumErn$ = Space$(11 * (3 - LastErn)) + SumErn$
  End If
  
  '--------------NEW----------------------------
  RSet Pg(1) = Str$(Page)
  If Not SumFlag Then
    Print #RHandle, DashLine(1)
  End If
  If SumFlag Then
    LSet EmpNo(1) = QPTrim$(Emp2Rec.EmpNo)
    Print #RHandle, EmpNo(1) + QPTrim$(Emp2Rec.EmpLName) + ", " + QPTrim$(Emp2Rec.EmpFName)
    Print #RHandle, "                        Tax Fring    Reg Hrs      Vacat       Sick        Hol       Comp   Personal      Total   O/T Paid        EIC"
  Else
    Print #RHandle, "Employee Totals:        Tax Fring    Reg Hrs      Vacat       Sick        Hol       Comp   Personal      Total   O/T Paid        EIC"
  End If

  Print #RHandle, Space$(22) + Fill11(1) + RHrs(1) + VHrs(1) + SHrs(1) + HHrs(1);
  
  Print #RHandle, CHrs(1) + PHrs(1) + THrs(1) + OTPaid(1) + EICP(1)
'  Print #RHandle,
  Print #RHandle, SumHeader2$ + "  Gross Pay    Soc Sec   Medicare        FWT        SWT  Ret Total    Net Pay"
  Print #RHandle, RErnP(1) + OErnP(1) + SumErn$ + GPayP(1) + SSTaxP(1) + MTaxP(1) + FTaxP(1);
  Print #RHandle, STaxP(1) + RetirP(1) + NetPayP(1)
'  Print #RHandle,
  For x = 1 To DLines '((NumOfDeds / 12) + 1)
    Print #RHandle, DTitle$(x)
    Print #RHandle, SumDed$(x)
  Next x
'  Print #RHandle,
  Print #RHandle, "  Fed Gross  Sta Gross  Soc Gross  Med Gross  Ret Gross"
  Print #RHandle, TFedGrs(1) + TStaGrs(1) + TSocGrs(1) + TMedGrs(1) + TRetGrs(1)
  Print #RHandle,
  Print #RHandle, DashLine(1)
  If Not SumFlag Then
    Print #RHandle, FF$
    LineCnt = 0
  Else
    LineCnt = LineCnt + 14 + DLines '(NumOfDeds / 12)
  End If
  Return
  '-----------------------------------------------------------------------
PrintGrandTotals:
  RSet Fill11(1) = Using(Image3$, TTaxFring#)
  RSet THrs(1) = Using(Image3$, EMPHIST(3).TotalHrs)
  RSet RHrs(1) = Using(Image3$, EMPHIST(3).RegHrs)
  RSet VHrs(1) = Using(Image3$, EMPHIST(3).VACHRS)
  RSet SHrs(1) = Using(Image3$, EMPHIST(3).SICKHRS)
  RSet HHrs(1) = Using(Image3$, EMPHIST(3).HOLHRS)
  RSet CHrs(1) = Using(Image3$, EMPHIST(3).COMPHRS)

  RSet PHrs(1) = Using(Image3$, EMPHIST(3).PHrs)
  RSet OTPaid(1) = Using(Image3$, EMPHIST(3).TOTPaid)
  RSet EICP(1) = Using(Image3$, EMPHIST(3).TotEIC)
  RSet RErnP(1) = Using(Image3$, EMPHIST(3).TRegWage)
  RSet OErnP(1) = Using(Image3$, EMPHIST(3).TOTWage)
  
  RSet GPayP(1) = Using(Image3$, EMPHIST(3).GPay + TaxFring#) '12/30/04)
  RSet SSTaxP(1) = Using(Image3$, EMPHIST(3).SSTax)
  RSet MTaxP(1) = Using(Image3$, EMPHIST(3).MTax)
  RSet FTaxP(1) = Using(Image3$, EMPHIST(3).FTax)
  RSet STaxP(1) = Using(Image3$, EMPHIST(3).STax)
  RSet RetirP(1) = Using(Image3$, EMPHIST(3).RETTOT)
  RSet NetPayP(1) = Using(Image3$, EMPHIST(3).TNetPay)

  RSet TFedGrs(1) = Using(Image3$, TFedGross#)
  RSet TStaGrs(1) = Using(Image3$, TStaGross#)
  RSet TSocGrs(1) = Using(Image3$, TSocGross#)
  RSet TMedGrs(1) = Using(Image3$, TMedGross#)
  RSet TRetGrs(1) = Using(Image3$, TRetGross#)

  For x = 1 To 5
    SumDed$(x) = ""
  Next x
  
  Nextx = 1
  tripCnt = 1
  For Cnt2 = 1 To NumOfDeds 'LastDed changed 4/30
    If tripCnt = 13 Then
      tripCnt = 1
      Nextx = Nextx + 1
    End If
    RSet Ded(1) = Using(Image3$, TotDeds(Cnt2))
    SumDed$(Nextx) = SumDed$(Nextx) + Ded(1)
    tripCnt = tripCnt + 1
  Next
  
  SumErn$ = ""
  For Cnt2 = 1 To LastErn
    RSet Ern(1) = Using(Image3$, TotErns(Cnt2))
    SumErn$ = SumErn$ + Ern(1)
  Next
  If LastErn < 3 Then
    SumErn$ = Space$(11 * (3 - LastErn)) + SumErn$
  End If
  If SumFlag Then
    Print #RHandle, FF$
  End If
  RSet Pg(1) = Str$(Page)
  Print #RHandle, City + Space$(86) + "Page:" + Pg(1)
  Print #RHandle, "Employee Earnings History Report" + Space$(63) + FromToDate$
  Print #RHandle, DashLine(1)
  Print #RHandle,
  Print #RHandle, "Report Totals:          Tax Fring    Reg Hrs      Vacat       Sick        Hol       Comp   Personal      Total   O/T Paid        EIC"
  Print #RHandle, Space$(22) + Fill11(1) + RHrs(1) + VHrs(1) + SHrs(1) + HHrs(1);
  Print #RHandle, CHrs(1) + PHrs(1) + THrs(1) + OTPaid(1) + EICP(1)
  Print #RHandle,
  Print #RHandle, SumHeader2$ + "  Gross Pay    Soc Sec   Medicare        FWT        SWT  Ret Total    Net Pay"
'add grand totals here
  Print #RHandle, RErnP(1) + OErnP(1) + SumErn$ + GPayP(1) + SSTaxP(1) + MTaxP(1) + FTaxP(1);
  Print #RHandle, STaxP(1) + RetirP(1) + NetPayP(1)
  Print #RHandle,
  For x = 1 To ((NumOfDeds / 12) + 1) 'added 4/30
    Print #RHandle, DTitle$(x)
    Print #RHandle, SumDed$(x)
  Next x
  Print #RHandle,
  Print #RHandle, "  Fed Gross  Sta Gross  Soc Gross  Med Gross  Ret Gross"
  Print #RHandle, TFedGrs(1) + TStaGrs(1) + TSocGrs(1) + TMedGrs(1) + TRetGrs(1)
  Print #RHandle, FF$
  Return
  
MaxSSWage:
  
  StopDateFlag = False
  TransRecNum& = Emp2Rec.LastTransRec
  ReDim SSTempDates(1 To 1) As Integer
  ReDim SSTempAmts(1 To 1) As Double
  Do
    Get THandle, TransRecNum&, TransRec(1)
    If (TransRec(1).CheckDate >= ThisDate) And (TransRec(1).CheckDate <= HiDate) Then
      NextDate = NextDate + 1
      ReDim Preserve SSTempDates(1 To NextDate) As Integer
      ReDim Preserve SSTempAmts(1 To NextDate) As Double
      SSTempDates(NextDate) = TransRec(1).CheckDate
      SSTempAmts(NextDate) = OldRound(TransRec(1).SocGrossPay + TransRec(1).TaxFring)
    End If
    If TransRec(1).PrevTransRec > 0 Then
      TransRecNum& = TransRec(1).PrevTransRec
    Else
      Exit Do
    End If
  Loop
    
  ReDim SSAmts(1 To NextDate) As Double
  ReDim SSDates(1 To NextDate) As Integer
  'assign dates/amts in forward order
  z = NextDate
  For x = 1 To NextDate
    SSDates(x) = SSTempDates(z)
    SSAmts(x) = SSTempAmts(z)
    z = z - 1
  Next x
  
  TotSocGross = 0
    
  For x = 1 To NextDate
    TotSocGross = OldRound(TotSocGross + SSAmts(x))
    If SSDates(x) >= LowDate Then SSTotal = SSTotal + SSAmts(x)
    If TotSocGross > FedSSMax Then
      If StopDateFlag = False Then
        StopDateFlag = True
        StopDate = SSDates(x)
        If NextDate > 1 Then
          ThisDif = OldRound(TotSocGross - FedSSMax)
          Exit For
        Else
          SSMaxCode = 4
          Exit For
        End If
      End If
    End If
  Next x
  
  If SSMaxCode = 4 Then
    SocGross# = FedSSMax
'    TSocGross# = TSocGross# + FedSSMax
    Return
  End If
  
  If StopDate < LowDate Then
    SocGross# = 0
'    TSocGross# = TSocGross#
  Else
    SocGross# = SSTotal - ThisDif
'    TSocGross# = TSocGross# + SocGross
  End If
  
  Return
  
  
EndTran:
End Sub

Private Sub fpcomboPrintOpt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcomboPrintOpt.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcomboPrintOpt.ListIndex = -1
  End If
  If fpcomboPrintOpt.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      cmdEscape.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
End Sub

