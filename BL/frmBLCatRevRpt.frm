VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLCatRevRpt 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Category Revenue Report"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11640
   Icon            =   "frmBLCatRevRpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   5910
      Left            =   1943
      TabIndex        =   4
      Top             =   1393
      Width           =   7785
      _Version        =   196609
      _ExtentX        =   13732
      _ExtentY        =   10425
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmBLCatRevRpt.frx":08CA
      Begin LpLib.fpCombo fpcmbPrintOpt 
         Height          =   405
         Left            =   2970
         TabIndex        =   3
         Tag             =   $"frmBLCatRevRpt.frx":08E6
         Top             =   3645
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
         ColDesigner     =   "frmBLCatRevRpt.frx":099F
      End
      Begin LpLib.fpCombo fpcmbBalance 
         Height          =   405
         Left            =   5115
         TabIndex        =   2
         Tag             =   $"frmBLCatRevRpt.frx":0C96
         Top             =   3000
         Width           =   1020
         _Version        =   196608
         _ExtentX        =   1799
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
         ColDesigner     =   "frmBLCatRevRpt.frx":0D25
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   645
         Left            =   3210
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "Press the 'Cancel' button to exit this screen and return to the main 'Business License Reports' menu."
         Top             =   4725
         Width           =   1890
         _Version        =   131072
         _ExtentX        =   3334
         _ExtentY        =   1138
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
         ButtonDesigner  =   "frmBLCatRevRpt.frx":101C
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   645
         Left            =   5235
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   $"frmBLCatRevRpt.frx":11FA
         Top             =   4725
         Width           =   1830
         _Version        =   131072
         _ExtentX        =   3228
         _ExtentY        =   1138
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
         ButtonDesigner  =   "frmBLCatRevRpt.frx":12A5
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
         Height          =   645
         Left            =   915
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   $"frmBLCatRevRpt.frx":1484
         Top             =   4725
         Width           =   2175
         _Version        =   131072
         _ExtentX        =   3836
         _ExtentY        =   1138
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
         ButtonDesigner  =   "frmBLCatRevRpt.frx":1554
      End
      Begin EditLib.fpDateTime fptxtBDate 
         Height          =   375
         Left            =   4005
         TabIndex        =   0
         Tag             =   "Enter the date the report will use as the beginning date for the transaction search."
         Top             =   1680
         Width           =   1740
         _Version        =   196608
         _ExtentX        =   3069
         _ExtentY        =   653
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
         BackColor       =   16777215
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
         ButtonColor     =   13684944
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDateTime fptxtEDate 
         Height          =   370
         Left            =   3885
         TabIndex        =   1
         Tag             =   "Enter the last date the report will look for as it searches through the transaction records for data for this report."
         Top             =   2355
         Width           =   1740
         _Version        =   196608
         _ExtentX        =   3069
         _ExtentY        =   653
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
         ButtonColor     =   13684944
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Include Zero Balances (Y/N)?:"
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
         Left            =   1560
         TabIndex        =   14
         Top             =   3075
         Width           =   3375
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2040
         TabIndex        =   13
         Top             =   1740
         Width           =   1740
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2115
         TabIndex        =   12
         Top             =   2400
         Width           =   1545
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
         Left            =   1320
         TabIndex        =   10
         Top             =   3750
         Width           =   1500
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   2985
         Left            =   990
         Top             =   1395
         Width           =   5895
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Category Revenue Report"
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
         Height          =   396
         Left            =   2016
         TabIndex        =   9
         Top             =   576
         Width           =   3948
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
         Left            =   960
         TabIndex        =   8
         Top             =   5400
         Width           =   2100
      End
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   450
      Left            =   840
      TabIndex        =   11
      Top             =   5923
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
      Height          =   6214
      Left            =   1733
      Top             =   1258
      Width           =   8175
   End
End
Attribute VB_Name = "frmBLCatRevRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class

Private Sub cmdExit_Click()
  frmBLCustReportsMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "F1 &Turn Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    lblBalloon.Visible = True
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
  End If
End Sub

Private Sub cmdProcess_Click()
  If fpcmbPrintOpt.Text = "Graphical" Then
    Call PrintGraphics
  Else
    Call PrintText
  End If

End Sub
Private Sub PrintGraphics()
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeHandle As Integer
  Dim CodeIdxRec As CatCodeIdxType
  Dim IdxHandle As Integer
  Dim NumOfCodes As Integer
  Dim TransRec As ARTransRecType
  Dim THandle As Integer
  Dim NumOfTransRecs As Double
  Dim x As Integer, y As Integer
  Dim BDate As Integer
  Dim EDate As Integer
  Dim TotAmt As Double
  Dim ThisCode As Integer
  Dim NegPos As Integer
  Dim dlm$
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim ThisTown$
  Dim ThisPct As Double
  Dim ThisTotal As Double
  Dim SortNum As Double
  Dim SortCnt As Integer
  Dim Nextx As Integer
  Dim Thisx As Integer
  Dim HoldRev As Double
  Dim HoldNum$
  Dim HoldDesc$
  Dim ReportFile$
  Dim RptHandle As Integer
  Dim HoldSortNum As Double
  Dim PosCnt As Integer
  Dim RunningPct As Double
  Dim AllZeros As Boolean
  Dim RevTally As Double
  
  On Error GoTo ERRORSTUFF
  
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  
  ThisTown = QPTrim$(TownRec.TownName)
  
  dlm$ = "~"
  BDate = Date2Num(fptxtBDate.Text)
  EDate = Date2Num(fptxtEDate.Text)
  OpenCatCodeIdxFile IdxHandle
  NumOfCodes = LOF(IdxHandle) / Len(CodeIdxRec)
  
  ReDim CodeIdx(1 To NumOfCodes) As Integer
  For x = 1 To NumOfCodes
    Get IdxHandle, x, CodeIdxRec
    CodeIdx(x) = CodeIdxRec.CatCodeRec 'load array with record pointers
  Next x
  Close IdxHandle
  
  ReDim ThisRev(1 To NumOfCodes) As Double
  ReDim ThisDesc(1 To NumOfCodes) As String
  ReDim ThisNum(1 To NumOfCodes) As String
  
  OpenCatCodeFile CodeHandle
  OpenTransFile THandle
  NumOfTransRecs = LOF(THandle) / Len(TransRec)
  frmBLShowPctComp.Label1 = "Gathering Category Revenue Data"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  cmdHelp.Enabled = False
  ThisTotal = 0
  For x = 1 To NumOfCodes
    Get CodeHandle, CodeIdx(x), CodeRec
    ThisCode = CodeIdx(x)
'    ThisDesc(x) = QPTrim$(CodeRec.CODEDESC)'changed on 02/02/05
'    ThisNum(x) = QPTrim$(CodeRec.CatCode)
    ThisDesc(ThisCode) = QPTrim$(CodeRec.CODEDESC) 'changed on 02/02/05
    ThisNum(ThisCode) = QPTrim$(CodeRec.CatCode)
    For y = 1 To NumOfTransRecs
      Get THandle, y, TransRec
        If TransRec.TransDate < BDate Or TransRec.TransDate > EDate Then GoTo SkipTrans
        If TransRec.TransAmount <= 0 Then GoTo SkipTrans
        If TransRec.CatCodeRec1 = ThisCode Then
          GoSub RevType
          If NegPos = 1 Then
            ThisRev(ThisCode) = ThisRev(ThisCode) + TransRec.CatLicAmt1
            ThisTotal = ThisTotal + TransRec.CatLicAmt1
          ElseIf NegPos = 0 Then
            ThisRev(ThisCode) = ThisRev(ThisCode) - TransRec.CatLicAmt1
            ThisTotal = ThisTotal - TransRec.CatLicAmt1
          End If
          
          GoTo SkipTrans
        ElseIf TransRec.CatCodeRec2 = ThisCode Then
          GoSub RevType
          If NegPos = 1 Then
            ThisRev(ThisCode) = ThisRev(ThisCode) + TransRec.CatLicAmt2
            ThisTotal = ThisTotal + TransRec.CatLicAmt2
          ElseIf NegPos = 0 Then
            ThisRev(ThisCode) = ThisRev(ThisCode) - TransRec.CatLicAmt2
            ThisTotal = ThisTotal - TransRec.CatLicAmt2
          End If
          GoTo SkipTrans
        ElseIf TransRec.CatCodeRec3 = ThisCode Then
          GoSub RevType
          If NegPos = 1 Then
            ThisRev(ThisCode) = ThisRev(ThisCode) + TransRec.CatLicAmt3
            ThisTotal = ThisTotal + TransRec.CatLicAmt3
          ElseIf NegPos = 0 Then
            ThisRev(ThisCode) = ThisRev(ThisCode) - TransRec.CatLicAmt3
            ThisTotal = ThisTotal - TransRec.CatLicAmt3
          End If
          GoTo SkipTrans
        ElseIf TransRec.CatCodeRec4 = ThisCode Then
          GoSub RevType
          If NegPos = 1 Then
            ThisRev(ThisCode) = ThisRev(ThisCode) + TransRec.CatLicAmt4
            ThisTotal = ThisTotal + TransRec.CatLicAmt4
          ElseIf NegPos = 0 Then
            ThisRev(ThisCode) = ThisRev(ThisCode) - TransRec.CatLicAmt4
            ThisTotal = ThisTotal - TransRec.CatLicAmt4
          End If
          GoTo SkipTrans
        ElseIf TransRec.CatCodeRec5 = ThisCode Then
          GoSub RevType
          If NegPos = 1 Then
            ThisRev(ThisCode) = ThisRev(ThisCode) + TransRec.CatLicAmt5
            ThisTotal = ThisTotal + TransRec.CatLicAmt5
          ElseIf NegPos = 0 Then
            ThisRev(ThisCode) = ThisRev(ThisCode) - TransRec.CatLicAmt5
            ThisTotal = ThisTotal - TransRec.CatLicAmt5
          End If
          GoTo SkipTrans
        End If
SkipTrans:
    Next y
    frmBLShowPctComp.ShowPctComp x, NumOfCodes
    If frmBLShowPctComp.Out = True Then
      Close
      frmBLShowPctComp.Out = False
      Unload frmBLShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      cmdHelp.Enabled = True
      Exit Sub
    End If
  Next x
  
  Close
  
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  cmdHelp.Enabled = True
  
  If ThisTotal = 0 And fpcmbBalance = "No" Then
    frmBLMessageBoxJr.Label1.Caption = "No category has a positive balance."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  ReportFile$ = "BLRPTS\ARCatRpt.Rpt"
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  AllZeros = False 'added 1/20/05 to prevent a crash if all revenues are zero
  GoSub SortIt
  If AllZeros = True Then
     For x = 1 To NumOfCodes
        Print #RptHandle, ThisNum(x); dlm; ThisDesc(x); dlm; CStr(ThisRev(x)); dlm; ThisTown; dlm; fptxtBDate.Text; dlm; fptxtEDate.Text; dlm; "0"; dlm; RunningPct
     Next x
     GoTo Zeros
  End If
  
  For x = 1 To NumOfCodes
    ThisPct = (ThisRev(x) / ThisTotal)
    RunningPct = RunningPct + ThisPct
    If ThisRev(x) = 0 Then
      If fpcmbBalance.Text = "Yes" Then
        Print #RptHandle, ThisNum(x); dlm; ThisDesc(x); dlm; CStr(ThisRev(x)); dlm; ThisTown; dlm; fptxtBDate.Text; dlm; fptxtEDate.Text; dlm; "0"; dlm; RunningPct
      End If
    Else
      Print #RptHandle, ThisNum(x); dlm; ThisDesc(x); dlm; CStr(ThisRev(x)); dlm; ThisTown; dlm; fptxtBDate.Text; dlm; fptxtEDate.Text; dlm; ThisPct; dlm; RunningPct
    End If
  Next x
Zeros:
  
  Close
  
  arBLCatRevRpt.Show
  frmBLLoadReport.Show

  Exit Sub
  
RevType:
  NegPos = 1 '1 = leave as is, -1 = omit rev, 0 = subtract rev
  
  If TransRec.DetailTransType <= 0 Then
    Select Case TransRec.TransType
      Case 2, 6, 13 'no impact on category charges (these are payment or penalty transactions)
        NegPos = -1
      Case 23 'adjust bill down
        NegPos = 0
    End Select
  Else
    Select Case TransRec.DetailTransType
      Case 101, 211, 210, 201, 301, 401 'no impact on category charges (these are payment or penalty transactions)
        NegPos = -1
      Case 311, 310 'adjust lic down
        NegPos = 0
    End Select
  End If
  
Return

SortIt:
  SortCnt = 0
  RevTally = 0 'added 01/20/05 to prevent a crash if all revenues are zero
  For x = 1 To NumOfCodes 'added 1/20/05
     RevTally = RevTally + ThisRev(x)
  Next x
  If RevTally = 0 Then 'added 1/20/05
     AllZeros = True
     Return
  End If
  
  ReDim SortIdx(1 To NumOfCodes) As Integer
  Nextx = 1
  SortNum = 0
  Do
    For x = Nextx To NumOfCodes
      If ThisRev(x) > SortNum Then
        SortNum = ThisRev(x)
        Thisx = x
      End If
    Next x
    SortIdx(Nextx) = Thisx
    HoldRev = ThisRev(Nextx)
    HoldNum = ThisNum(Nextx)
    HoldDesc = ThisDesc(Nextx)
    ThisRev(Nextx) = ThisRev(Thisx)
    ThisNum(Nextx) = ThisNum(Thisx)
    ThisDesc(Nextx) = ThisDesc(Thisx)
    ThisRev(Thisx) = HoldRev
    ThisNum(Thisx) = HoldNum
    ThisDesc(Thisx) = HoldDesc
    Nextx = Nextx + 1
    If Nextx = NumOfCodes + 1 Then Exit Do
    SortNum = 0
  Loop
  
  Nextx = 1
  SortNum = 0
  For x = 1 To NumOfCodes
    If ThisRev(x) <> 0 Then
      PosCnt = PosCnt + 1
      GoTo NotZero
    End If
    If Val(ThisNum(x)) > SortNum Then
      SortNum = Val(ThisNum(x))
    End If
NotZero:
  Next x
  
  HoldSortNum = SortNum + 1
  
  If PosCnt = NumOfCodes Then GoTo GetOut
  
  PosCnt = PosCnt + 1
  
  SortNum = HoldSortNum
  Nextx = 1
  Do
    For x = Nextx To NumOfCodes
      If ThisRev(x) <> 0 Then GoTo NotZero2
      If x >= PosCnt Then
        If Val(ThisNum(x)) < SortNum Then
          SortNum = Val(ThisNum(x))
          Thisx = x
        End If
      End If
NotZero2:
    Next x
    HoldRev = ThisRev(PosCnt)
    HoldNum = ThisNum(PosCnt)
    HoldDesc = ThisDesc(PosCnt)
    ThisRev(PosCnt) = ThisRev(Thisx)
    ThisNum(PosCnt) = ThisNum(Thisx)
    ThisDesc(PosCnt) = ThisDesc(Thisx)
    ThisRev(Thisx) = HoldRev
    ThisNum(Thisx) = HoldNum
    ThisDesc(Thisx) = HoldDesc
    SortNum = HoldSortNum
    PosCnt = PosCnt + 1
    Nextx = Nextx + 1
    If PosCnt = NumOfCodes + 1 Then Exit Do
  Loop
GetOut:
  Return
  
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustBalListing", "PrintGraphics", Erl)
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
      SendKeys "%P"
      Call cmdProcess_Click
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
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLCatRevRpt.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  lblBalloon.Visible = False
  fptxtBDate = "01/01/" + Mid(Date, 7, 4)
  fptxtEDate = Date
  fpcmbPrintOpt.AddItem "Graphical"
  fpcmbPrintOpt.AddItem "Text"
  fpcmbPrintOpt.Text = "Graphical"
  fpcmbBalance.AddItem "Yes"
  fpcmbBalance.AddItem "No"
  fpcmbBalance.Text = "No"
End Sub

Private Sub PrintText()
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeHandle As Integer
  Dim CodeIdxRec As CatCodeIdxType
  Dim IdxHandle As Integer
  Dim NumOfCodes As Integer
  Dim TransRec As ARTransRecType
  Dim THandle As Integer
  Dim NumOfTransRecs As Double
  Dim x As Integer, y As Integer
  Dim BDate As Integer
  Dim EDate As Integer
  Dim TotAmt As Double
  Dim ThisCode As Integer
  Dim NegPos As Integer
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim ThisTown$
  Dim ThisPct As Double
  Dim ThisTotal As Double
  Dim SortNum As Double
  Dim SortCnt As Integer
  Dim Nextx As Integer
  Dim Thisx As Integer
  Dim HoldRev As Double
  Dim HoldNum$
  Dim HoldDesc$
  Dim ReportFile$
  Dim RptHandle As Integer
  Dim HoldSortNum As Double
  Dim PosCnt As Integer
  Dim FF$, Page As Integer
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim TotPct As Double
  Dim RunPct As Double
  Dim AllZeros As Boolean
  Dim RevTally As Double
  
  On Error GoTo ERRORSTUFF
  
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  
  ThisTown = QPTrim$(TownRec.TownName)
  
  BDate = Date2Num(fptxtBDate.Text)
  EDate = Date2Num(fptxtEDate.Text)
  
  FF$ = Chr$(12)
  MaxLines = 58
  LineCnt = 0
  
  OpenCatCodeIdxFile IdxHandle
  NumOfCodes = LOF(IdxHandle) / Len(CodeIdxRec)
  
  ReDim CodeIdx(1 To NumOfCodes) As Integer
  For x = 1 To NumOfCodes
    Get IdxHandle, x, CodeIdxRec
    CodeIdx(x) = CodeIdxRec.CatCodeRec 'load array with record pointers
  Next x
  Close IdxHandle
  
  ReDim ThisRev(1 To NumOfCodes) As Double
  ReDim ThisDesc(1 To NumOfCodes) As String
  ReDim ThisNum(1 To NumOfCodes) As String
  
  OpenCatCodeFile CodeHandle
  OpenTransFile THandle
  NumOfTransRecs = LOF(THandle) / Len(TransRec)
  frmBLShowPctComp.Label1 = "Gathering Category Revenue Data"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  cmdHelp.Enabled = False
  ThisTotal = 0
  For x = 1 To NumOfCodes
    Get CodeHandle, CodeIdx(x), CodeRec
    ThisCode = CodeIdx(x)
'    ThisDesc(x) = QPTrim$(CodeRec.CODEDESC)'changed 02/02/05
'    ThisNum(x) = QPTrim$(CodeRec.CatCode)
    ThisDesc(ThisCode) = QPTrim$(CodeRec.CODEDESC) 'changed 02/02/05
    ThisNum(ThisCode) = QPTrim$(CodeRec.CatCode)
    For y = 1 To NumOfTransRecs
      Get THandle, y, TransRec
        If TransRec.TransDate < BDate Or TransRec.TransDate > EDate Then GoTo SkipTrans
        If TransRec.TransAmount <= 0 Then GoTo SkipTrans
        If TransRec.CatCodeRec1 = ThisCode Then
          GoSub RevType
          If NegPos = 1 Then
            ThisRev(ThisCode) = ThisRev(ThisCode) + TransRec.CatLicAmt1
            ThisTotal = ThisTotal + TransRec.CatLicAmt1
          ElseIf NegPos = 0 Then
            ThisRev(ThisCode) = ThisRev(ThisCode) - TransRec.CatLicAmt1
            ThisTotal = ThisTotal - TransRec.CatLicAmt1
          End If
          
          GoTo SkipTrans
        ElseIf TransRec.CatCodeRec2 = ThisCode Then
          GoSub RevType
          If NegPos = 1 Then
            ThisRev(ThisCode) = ThisRev(ThisCode) + TransRec.CatLicAmt2
            ThisTotal = ThisTotal + TransRec.CatLicAmt2
          ElseIf NegPos = 0 Then
            ThisRev(ThisCode) = ThisRev(ThisCode) - TransRec.CatLicAmt2
            ThisTotal = ThisTotal - TransRec.CatLicAmt2
          End If
          GoTo SkipTrans
        ElseIf TransRec.CatCodeRec3 = ThisCode Then
          GoSub RevType
          If NegPos = 1 Then
            ThisRev(ThisCode) = ThisRev(ThisCode) + TransRec.CatLicAmt3
            ThisTotal = ThisTotal + TransRec.CatLicAmt3
          ElseIf NegPos = 0 Then
            ThisRev(ThisCode) = ThisRev(ThisCode) - TransRec.CatLicAmt3
            ThisTotal = ThisTotal - TransRec.CatLicAmt3
          End If
          GoTo SkipTrans
        ElseIf TransRec.CatCodeRec4 = ThisCode Then
          GoSub RevType
          If NegPos = 1 Then
            ThisRev(ThisCode) = ThisRev(ThisCode) + TransRec.CatLicAmt4
            ThisTotal = ThisTotal + TransRec.CatLicAmt4
          ElseIf NegPos = 0 Then
            ThisRev(ThisCode) = ThisRev(ThisCode) - TransRec.CatLicAmt4
            ThisTotal = ThisTotal - TransRec.CatLicAmt4
          End If
          GoTo SkipTrans
        ElseIf TransRec.CatCodeRec5 = ThisCode Then
          GoSub RevType
          If NegPos = 1 Then
            ThisRev(ThisCode) = ThisRev(ThisCode) + TransRec.CatLicAmt5
            ThisTotal = ThisTotal + TransRec.CatLicAmt5
          ElseIf NegPos = 0 Then
            ThisRev(ThisCode) = ThisRev(ThisCode) - TransRec.CatLicAmt5
            ThisTotal = ThisTotal - TransRec.CatLicAmt5
          End If
          GoTo SkipTrans
        End If
SkipTrans:
    Next y
    frmBLShowPctComp.ShowPctComp x, NumOfCodes
    If frmBLShowPctComp.Out = True Then
      Close
      frmBLShowPctComp.Out = False
      Unload frmBLShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      cmdHelp.Enabled = True
      Exit Sub
    End If
  Next x
  
  Close
  
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  cmdHelp.Enabled = True
  
  If ThisTotal = 0 And fpcmbBalance = "No" Then
    frmBLMessageBoxJr.Label1.Caption = "No category has a positive balance."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  ReportFile$ = "ARCatRev.PRN"  'Report File Name
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  GoSub PrintHeader
  AllZeros = False 'added 1/20/05 to prevent a crash if all revenues are zero
  GoSub SortIt
  If AllZeros = True Then
     For x = 1 To NumOfCodes
        Print #RptHandle, Tab(2); ThisNum(x); Tab(10); ThisDesc(x); Tab(45); Using("$#,###,##0.00", ThisRev(x)); Tab(63); Using("##0.00%", ThisPct); Tab(74); Using("##0.00%", RunPct)
        LineCnt = LineCnt + 1
        If LineCnt > MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintHeader
        End If
     Next x
     GoTo Zeros
  End If
  
  For x = 1 To NumOfCodes
    ThisPct = (ThisRev(x) / ThisTotal)
    RunPct = RunPct + ThisPct
    If ThisRev(x) = 0 Then
      If fpcmbBalance.Text = "Yes" Then
        Print #RptHandle, Tab(2); ThisNum(x); Tab(10); ThisDesc(x); Tab(45); Using("$#,###,##0.00", ThisRev(x)); Tab(63); Using("##0.00%", ThisPct); Tab(74); Using("##0.00%", RunPct)
        LineCnt = LineCnt + 1
        If LineCnt > MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintHeader
        End If
      End If
    Else
      Print #RptHandle, Tab(2); ThisNum(x); Tab(10); ThisDesc(x); Tab(45); Using("$#,###,##0.00", ThisRev(x)); Tab(63); Using("##0.00%", ThisPct); Tab(74); Using("##0.00%", RunPct)
      LineCnt = LineCnt + 1
      TotPct = TotPct + ThisPct
      TotAmt = TotAmt + ThisRev(x)
      If LineCnt > MaxLines Then
        Print #RptHandle, FF$
        GoSub PrintHeader
      End If
    End If
  Next x
Zeros:
  
  If LineCnt > MaxLines - 5 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  
  GoSub PrintFooter
  Close
  
  ViewPrint ReportFile$, "Customer Balance Listing", True
  KillFile ReportFile$

  Exit Sub
  
PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(20); "Business License: Category Revenue Report"
  Print #RptHandle, "Report Date: "; Date$; Tab(65); "Page #"; Page
  Print #RptHandle, "From: " + fptxtBDate.Text + " to " + fptxtEDate.Text
  Print #RptHandle, "Note: Numbers reflect license charges only"
  Print #RptHandle,
  Print #RptHandle, Tab(2); "Cat #"; Tab(10); "Description"; Tab(51); "Revenue"; Tab(60); "% to Total"; Tab(74); "Accum %"
  Print #RptHandle, String$(80, "=")
  LineCnt = 7
  
  Return
  
PrintFooter:
  Print #RptHandle, String$(80, "-")
  Print #RptHandle, Tab(2); "Total Categories: "; Using("#,###0", NumOfCodes);
  Print #RptHandle, Tab(27); "Total Revenue:"; Tab(45); Using("$#,###,##0.00", TotAmt#)
  Print #RptHandle, FF$
  
  Return
  

RevType:
  NegPos = 1 '1 = leave as is, -1 = omit rev, 0 = subtract rev
  
  If TransRec.DetailTransType <= 0 Then
    Select Case TransRec.TransType
      Case 2, 6, 13 'no impact on category charges (these are payment or penalty transactions)
        NegPos = -1
      Case 23 'adjust bill down
        NegPos = 0
    End Select
  Else
    Select Case TransRec.DetailTransType
      Case 101, 211, 210, 201, 301, 401 'no impact on category charges (these are payment or penalty transactions)
        NegPos = -1
      Case 311, 310 'adjust lic down
        NegPos = 0
    End Select
  End If
  
Return

SortIt:
  SortCnt = 0
  RevTally = 0 'added 01/20/05 to prevent a crash if all revenues are zero
  For x = 1 To NumOfCodes 'added 1/20/05
     RevTally = RevTally + ThisRev(x)
  Next x
  If RevTally = 0 Then 'added 1/20/05
     AllZeros = True
     Return
  End If
  
  ReDim SortIdx(1 To NumOfCodes) As Integer
  Nextx = 1
  SortNum = 0
  Do
    For x = Nextx To NumOfCodes
      If ThisRev(x) > SortNum Then
        SortNum = ThisRev(x)
        Thisx = x
      End If
    Next x
    SortIdx(Nextx) = Thisx
    HoldRev = ThisRev(Nextx)
    HoldNum = ThisNum(Nextx)
    HoldDesc = ThisDesc(Nextx)
    ThisRev(Nextx) = ThisRev(Thisx)
    ThisNum(Nextx) = ThisNum(Thisx)
    ThisDesc(Nextx) = ThisDesc(Thisx)
    ThisRev(Thisx) = HoldRev
    ThisNum(Thisx) = HoldNum
    ThisDesc(Thisx) = HoldDesc
    Nextx = Nextx + 1
    If Nextx = NumOfCodes + 1 Then Exit Do
    SortNum = 0
  Loop
  
  Nextx = 1
  SortNum = 0
  For x = 1 To NumOfCodes
    If ThisRev(x) <> 0 Then
      PosCnt = PosCnt + 1
      GoTo NotZero
    End If
    If Val(ThisNum(x)) > SortNum Then
      SortNum = Val(ThisNum(x))
    End If
NotZero:
  Next x
  
  HoldSortNum = SortNum + 1
  
  If PosCnt = NumOfCodes Then GoTo GetOut
  PosCnt = PosCnt + 1
  
  SortNum = HoldSortNum
  Nextx = 1
  Do
    For x = Nextx To NumOfCodes
      If ThisRev(x) <> 0 Then GoTo NotZero2
      If x >= PosCnt Then
        If Val(ThisNum(x)) < SortNum Then
          SortNum = Val(ThisNum(x))
          Thisx = x
        End If
      End If
NotZero2:
    Next x
    HoldRev = ThisRev(PosCnt)
    HoldNum = ThisNum(PosCnt)
    HoldDesc = ThisDesc(PosCnt)
    ThisRev(PosCnt) = ThisRev(Thisx)
    ThisNum(PosCnt) = ThisNum(Thisx)
    ThisDesc(PosCnt) = ThisDesc(Thisx)
    ThisRev(Thisx) = HoldRev
    ThisNum(Thisx) = HoldNum
    ThisDesc(Thisx) = HoldDesc
    SortNum = HoldSortNum
    PosCnt = PosCnt + 1
    Nextx = Nextx + 1
    If PosCnt = NumOfCodes + 1 Then Exit Do
  Loop
GetOut:
  
  Return
  
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLCustBalListing", "PrintGraphics", Erl)
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

Private Sub fpcmbBalance_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbBalance.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbBalance.ListIndex = -1
  End If
  If fpcmbBalance.ListDown <> True Then
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
      fptxtBDate.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub
