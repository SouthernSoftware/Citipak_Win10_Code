VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "IMP32X30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "EDT32X30.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmBLTransJournal 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Transactions by Payment Type"
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   11652
   Icon            =   "frmBLTransJournal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   6084
      Left            =   1920
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1392
      Width           =   7788
      _Version        =   196609
      _ExtentX        =   13737
      _ExtentY        =   10731
      _StockProps     =   70
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmBLTransJournal.frx":08CA
      Begin LpLib.fpCombo fpcmbPrintOpt 
         Height          =   384
         Left            =   2976
         TabIndex        =   3
         Tag             =   $"frmBLTransJournal.frx":08E6
         Top             =   3888
         Width           =   3564
         _Version        =   196608
         _ExtentX        =   6286
         _ExtentY        =   677
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
         ColDesigner     =   "frmBLTransJournal.frx":099F
      End
      Begin LpLib.fpCombo fpcmbCategory 
         Height          =   384
         Left            =   2928
         TabIndex        =   0
         Tag             =   $"frmBLTransJournal.frx":0C5E
         Top             =   1824
         Width           =   3564
         _Version        =   196608
         _ExtentX        =   6286
         _ExtentY        =   677
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
         ColDesigner     =   "frmBLTransJournal.frx":0F1F
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   684
         Left            =   3120
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "Press the 'Cancel' button to exit this screen and return to the main 'Business License Reports' menu."
         Top             =   4896
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
         ButtonDesigner  =   "frmBLTransJournal.frx":11DE
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   684
         Left            =   5328
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   $"frmBLTransJournal.frx":13BC
         Top             =   4896
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
         ButtonDesigner  =   "frmBLTransJournal.frx":1467
      End
      Begin EditLib.fpDateTime fptxtBDate 
         Height          =   348
         Left            =   3840
         TabIndex        =   1
         Tag             =   "Enter the date the report will use as the beginning date for it's transaction search."
         Top             =   2544
         Width           =   1740
         _Version        =   196608
         _ExtentX        =   3069
         _ExtentY        =   614
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
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDateTime fptxtEDate 
         Height          =   348
         Left            =   3840
         TabIndex        =   2
         Tag             =   "Enter the last date the report will look for as it searches through the transaction records for data for this report."
         Top             =   3216
         Width           =   1740
         _Version        =   196608
         _ExtentX        =   3069
         _ExtentY        =   614
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
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
         Height          =   684
         Left            =   816
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   $"frmBLTransJournal.frx":1646
         Top             =   4896
         Width           =   2028
         _Version        =   131072
         _ExtentX        =   3577
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
         ButtonDesigner  =   "frmBLTransJournal.frx":1716
      End
      Begin VB.Label lblBalloon 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "HELP BALLOONS ON"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   252
         Left            =   768
         TabIndex        =   13
         Top             =   5616
         Width           =   2100
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
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
         Height          =   348
         Left            =   1296
         TabIndex        =   11
         Top             =   3984
         Width           =   1500
      End
      Begin VB.Label Label3 
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
         Left            =   2064
         TabIndex        =   10
         Top             =   3312
         Width           =   1548
      End
      Begin VB.Label Label2 
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
         Left            =   1872
         TabIndex        =   9
         Top             =   2592
         Width           =   1740
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   3132
         Left            =   1008
         Top             =   1488
         Width           =   5964
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Type:"
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
         Left            =   1344
         TabIndex        =   8
         Top             =   1920
         Width           =   1356
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Transactions by Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   396
         Left            =   1776
         TabIndex        =   7
         Top             =   576
         Width           =   4332
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
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   444
      Left            =   1776
      TabIndex        =   14
      Top             =   7680
      Width           =   780
      _Version        =   131072
      _ExtentX        =   1376
      _ExtentY        =   783
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.2
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
      ShapeRoundWidth =   192
      ShapeRoundHeight=   192
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
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   6348
      Left            =   1800
      Top             =   1260
      Width           =   8052
   End
End
Attribute VB_Name = "frmBLTransJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class

Private Sub cmdCodeList_Click()
  frmBLCategoryList.Show vbModal
  DoEvents
End Sub

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "F1 &Turn Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    lblBalloon.Visible = True
    fpcmbCategory.ToolTipText = ""
    fptxtBDate.ToolTipText = ""
    fptxtEDate.ToolTipText = ""
    fpcmbPrintOpt.ToolTipText = ""
    cmdHelp.ToolTipText = ""
    cmdExit.ToolTipText = ""
    cmdProcess.ToolTipText = ""
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
'    fpcmbCategory.ToolTipText = "Select one of the transaction types from the drop down box."
'    fptxtBDate.ToolTipText = "Enter the date the transaction search should begin."
'    fptxtEDate.ToolTipText = "Enter the date the report will use as the last date to search for transactions."
'    fpcmbPrintOpt.ToolTipText = "Select graphical to print on a laser printer or choose text to print on a dot matrix printer."
'    cmdHelp.ToolTipText = "Press to activate or deactivate instructional balloons."
'    cmdExit.ToolTipText = "Press to exit screen."
'    cmdProcess.ToolTipText = "Press to activate the reporting process."
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
    Me.Visible = False
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
      Call cmdExit_Click
      SendKeys "%C"
      KeyCode = 0
    Case vbKeyF10:
      Call cmdProcess_Click
      SendKeys "%P"
      KeyCode = 0
    Case vbKeyF1:
      Call cmdHelp_Click
      SendKeys "%T"
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
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLTransJournal.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  lblBalloon.Visible = False
'  fpcmbCategory.ToolTipText = "Select one of the transaction types from the drop down box."
'  fptxtBDate.ToolTipText = "Enter the date the transaction search should begin."
'  fptxtEDate.ToolTipText = "Enter the date the report will use as the last date to search for transactions."
'  fpcmbPrintOpt.ToolTipText = "Select graphical to print on a laser printer or choose text to print on a dot matrix printer."
'  cmdHelp.ToolTipText = "Press to activate or deactivate instructional balloons."
'  cmdExit.ToolTipText = "Press to exit screen."
'  cmdProcess.ToolTipText = "Press to activate the reporting process."
  fptxtBDate = "01/01/" + Mid(Date, 7, 4)
  fptxtEDate = Date
  fpcmbCategory.Text = "All Charges"
  fpcmbCategory.AddItem "All Charges"
  fpcmbCategory.AddItem "All Penalties"
  fpcmbCategory.AddItem "All Payments"
  fpcmbCategory.AddItem "All Adjustments Down"
  fpcmbCategory.AddItem "All Adjustments Up"
  fpcmbCategory.AddItem "Charge Licenses"
  fpcmbCategory.AddItem "Pay Licenses"
  fpcmbCategory.AddItem "License Adjustments Down"
  fpcmbCategory.AddItem "License Adjustments Up"
  fpcmbCategory.AddItem "Charge Penalties"
  fpcmbCategory.AddItem "Pay Penalties"
  fpcmbCategory.AddItem "Penalty Adjustments Down"
  fpcmbCategory.AddItem "Penalty Adjustments Up"
  fpcmbCategory.AddItem "Charge Issuance Fees"
  fpcmbCategory.AddItem "Pay Issuance Fees"
  fpcmbCategory.AddItem "Issuance Adjustments Down"
  fpcmbCategory.AddItem "Issuance Adjustments Up"
  fpcmbPrintOpt.AddItem "Graphical"
  fpcmbPrintOpt.AddItem "Text"
  fpcmbPrintOpt.Text = "Graphical"

End Sub

Private Sub fpcmbCategory_Change()
  If QPTrim$(fpcmbCategory.Text) = "" Then
    fpcmbCategory.Text = "All Charges"
  End If
End Sub

Private Sub fpcmbCategory_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbCategory.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbCategory.ListIndex = -1
  End If
  If fpcmbCategory.ListDown <> True Then
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

Private Sub fpcmbPrintOpt_Change()
  If QPTrim$(fpcmbPrintOpt.Text) = "" Then
    fpcmbPrintOpt.Text = "Graphical"
  End If
End Sub

Private Sub fpcmbPrintOpt_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  If KeyCode = vbKeySpace Then
    fpcmbPrintOpt.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPrintOpt.ListIndex = -1
  End If
  If fpcmbPrintOpt.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbCategory.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
End Sub

Private Sub cmdExit_Click()
'  frmBLTransJrnlMenu.Show
  frmBLCustReportsMenu.Show
  DoEvents
  Unload frmBLTransJournal
End Sub

Private Sub cmdProcess_Click()
  If Not Exist("artrans.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "No transactions records saved."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  If fpcmbPrintOpt.Text = "Graphical" Then
    Call PrintGraphics
  ElseIf fpcmbPrintOpt.Text = "Text" Then
    If InStr(fpcmbCategory.Text, "All") Then
      frmBLMessageBoxJr.Label1.Caption = "Pitch 17 is recommended for this report."
    Else
      frmBLMessageBoxJr.Label1.Caption = "Pitch 10 is recommended for this report."
    End If
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Call PrintText
  Else
    Exit Sub
  End If
End Sub

Private Sub PrintText()
  Dim BegDate$
  Dim BegDateNum As Integer
  Dim EndDate$
  Dim EndDateNum As Integer
  Dim ReportFile$
  Dim RptHandle As Integer
  Dim TransCnt As Double
  Dim TransRec As ARTransRecType
  Dim THandle As Integer
  Dim LineCnt As Integer
  Dim MaxLines As Integer
  Dim FF$, cnt As Double
  Dim TotalTrans As Double
  Dim TotalAmt As Double
  Dim TotalPaid As Double
  Dim FeePd As Double
  Dim Category$
  Dim LeftOver As Double
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim BILLCAT1$
  Dim BILLCAT2$
  Dim BILLCAT3$
  Dim BILLCAT4$
  Dim BILLCAT5$
  Dim Fee1#
  Dim Fee2#
  Dim Fee3#
  Dim Fee4#
  Dim Fee5#
  Dim CatCnt!, CatFnd!
  Dim ll As Double
  Dim CategoryDesc$
  Dim CodeRec As ARNewCatCodeRecType
  Dim NumOfCatRecs As Integer
  Dim COHandle As Integer
  Dim LCnt As Integer
  Dim Page As Integer
  Dim TRNumRecs As Double
  Dim CountNum As Double
  Dim BigNum$
  Dim HoldThis As TransIdxType
  Dim ThisRec As Double
  Dim SmallNum$
  Dim Nextx As Double
  Dim x As Double, y As Integer, z As Integer
  Dim CatTotalAmt As Double
  Dim ThisTransType$
  Dim HeadLen1 As Integer
  Dim HeadLen2 As Integer
  Dim PrintCnt As Integer
  Dim LicAndPenFlag As Boolean
  Dim PTotal As Double
  Dim LTotal As Double
  Dim LPTotal As Double
  Dim LPCnt As Integer
  Dim ITotal As Double
  Dim NCnt As Integer
  Dim ThisTotal As Double
  
  On Error GoTo ERRORSTUFF
  
  cmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  
  ReportFile$ = "ARTRANS.PRN"
  FF$ = Chr$(12)
  MaxLines = 53
  LineCnt = 0
  ReDim Cat$(300), CatAmt#(300), GTotalAmt#(103), TypeCnt%(103)
  
  FF$ = Chr(12)
  
  BegDate = fptxtBDate.Text
  BegDateNum = Date2Num(fptxtBDate.Text)
  EndDate = fptxtEDate.Text
  EndDateNum = Date2Num(fptxtEDate.Text)
  If EndDateNum < BegDateNum Then
    fptxtEDate.BackColor = &HFFFF&
    fptxtBDate.BackColor = &HFFFF&
    frmBLMessageBoxJr.Label1.Caption = "The ending date comes before the beginning date. Please re-enter these values."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    Close
    fptxtEDate.BackColor = &HFFFFFF
    fptxtBDate.BackColor = &HFFFFFF
    fptxtBDate.SetFocus
    Exit Sub
  End If
  
  OpenTransFile THandle 'used also in GetReportInformation2
  ThisTransType = QPTrim$(fpcmbCategory.Text)
  
  HeadLen1 = Len("Transaction Type: " + ThisTransType)
  HeadLen1 = HeadLen1 / 2
  HeadLen1 = Abs(40 - HeadLen1)
  HeadLen2 = Len("Total " + ThisTransType + " Transactions")
  HeadLen2 = HeadLen2 / 2
  HeadLen2 = Abs(40 - HeadLen2)
  
  ReDim TransIdx(1 To 1) As TransIdxType
  'this report has two versions determined by what
  'is returned by GetReportInformation2
  GoSub GetReportInformation2
  
  If LicAndPenFlag = True Then
    ReDim LTotalAmt#(103), PTotalAmt#(103), ITotalAmt#(103)
  End If
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  GoSub PrintRptHeader2
  OpenCustFile CHandle
  TransCnt = LOF(THandle) / Len(TransRec)
  frmBLShowPctComp.Label1 = "Loading Transaction List"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  cmdHelp.Enabled = False
  
  For cnt = CountNum To 1 Step -1
    Get THandle, TransIdx(cnt).TransRecNum, TransRec
    If Val(TransRec.CustomerNumber) = 0 Then
      GoTo BadCustSkip
    End If
    
    If LineCnt >= MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintRptHeader2
    End If
    
    'Get Customer
    Get CHandle, Val(TransRec.CustomerNumber), CustRec
    If QPTrim$(CustRec.Deleted) = "Y" Then GoTo BadCustSkip
    
    If TransIdx(cnt).TransAmt > 0 Then
      ThisTotal = TransIdx(cnt).TransAmt
    Else
      ThisTotal = TransRec.TransAmount
    End If
    
    Print #RptHandle, MakeRegDate(TransRec.TransDate);
    Print #RptHandle, Tab(13); Left$(CustRec.CustName, 25);
    
    If LicAndPenFlag = False Then
      Select Case TransRec.TransType
      Case 1
        Print #RptHandle, Tab(53); "License Charges";
        PrintCnt = PrintCnt + 1
      Case 2
        Print #RptHandle, Tab(52); "Payment";
        PrintCnt = PrintCnt + 1
      Case 6
        Print #RptHandle, Tab(52); "Penalty Charges";
        PrintCnt = PrintCnt + 1
      Case 9
        Print #RptHandle, Tab(50); "Beg Bal";
        PrintCnt = PrintCnt + 1
      Case 13
        Print #RptHandle, Tab(45); "DOWN Pay Adjustment";
        PrintCnt = PrintCnt + 1
      Case 23
        Print #RptHandle, Tab(45); "DOWN Bill Adjustment";
        PrintCnt = PrintCnt + 1
      Case 24
        Print #RptHandle, Tab(45); "UP Bill Adjustment";
        PrintCnt = PrintCnt + 1
      Case 100
        Print #RptHandle, Tab(45); "DOWN Adjustment";
        PrintCnt = PrintCnt + 1
      Case 101
        Print #RptHandle, Tab(45); "UP Adjustment";
        PrintCnt = PrintCnt + 1
      Case Else
      End Select
    Else
      Select Case TransRec.TransType
      Case 1
        Print #RptHandle, Tab(40); "Charge Lic";
        PrintCnt = PrintCnt + 1
      Case 2
        Print #RptHandle, Tab(40); "Payment";
        PrintCnt = PrintCnt + 1
      Case 6
        Print #RptHandle, Tab(40); "Charge Pen ";
        PrintCnt = PrintCnt + 1
      Case 9
        Print #RptHandle, Tab(40); "Beg Bal";
        PrintCnt = PrintCnt + 1
      Case 13
        Print #RptHandle, Tab(40); "DOWN Pay Adj";
        PrintCnt = PrintCnt + 1
      Case 23
        Print #RptHandle, Tab(40); "DOWN Bill Adj";
        PrintCnt = PrintCnt + 1
      Case 24
        Print #RptHandle, Tab(40); "UP Bill Adj";
        PrintCnt = PrintCnt + 1
      Case 100
        Print #RptHandle, Tab(40); "DOWN Adj";
        PrintCnt = PrintCnt + 1
      Case 101
        Print #RptHandle, Tab(40); "UP Adj";
        PrintCnt = PrintCnt + 1
      Case Else
      End Select
    End If
    If LicAndPenFlag = False Then 'LicAndPenFlag is true if the report
    'prints multiple totals
      Print #RptHandle, Tab(70); Using("$###,##0.00", ThisTotal)
      TotalAmt# = TotalAmt# + ThisTotal
    Else
      Print #RptHandle, Tab(53); Using("$##,##0.00", TransRec.PenAmt); Tab(66); Using("$#,##0.00", TransRec.IssAmt); Tab(77); Using("$###,##0.00", TransRec.LicAmt); Tab(89); Using("$###,##0.00", TransRec.TransAmount)
      TotalAmt# = TotalAmt# + TransRec.TransAmount
    End If
    TotalTrans = TotalTrans + 1
    TypeCnt(TransRec.TransType) = TypeCnt(TransRec.TransType) + 1
    LineCnt = LineCnt + 1
    Rem total by category
TotalUp:
    If ThisTotal > 0 Then 'ThisTotal > 0 if LicAndPenFlag = False
      GTotalAmt#(TransRec.TransType) = GTotalAmt#(TransRec.TransType) + ThisTotal
    Else
      GTotalAmt#(TransRec.TransType) = GTotalAmt#(TransRec.TransType) + TransRec.TransAmount
    End If
    
    If LicAndPenFlag = True Then
      LTotalAmt#(TransRec.TransType) = LTotalAmt#(TransRec.TransType) + TransRec.LicAmt
      PTotalAmt#(TransRec.TransType) = PTotalAmt#(TransRec.TransType) + TransRec.PenAmt
      ITotalAmt#(TransRec.TransType) = ITotalAmt#(TransRec.TransType) + TransRec.IssAmt
    End If
    
BadCustSkip:
    frmBLShowPctComp.ShowPctComp cnt, CountNum
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
    
  Next cnt
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  cmdHelp.Enabled = True
  
  GoSub PrintRptEnding2
  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
  Close         'Close all open files now
  
  If PrintCnt = 0 Then
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Label1.Caption = "There are no transactions saved between " + fptxtBDate.Text + " and " + fptxtEDate.Text + " for " + QPTrim$(fpcmbCategory.Text) + "."
    frmBLMessageBoxJr.Show vbModal
  Else
    ViewPrint ReportFile, "Transaction Journal", True
  End If
  
  KillFile ReportFile$
  
  MainLog ("'Transaction By Type' report processed for " + QPTrim$(fpcmbCategory.Text) + " beginning on " + fptxtBDate + " thru " + fptxtEDate + " in text format.")
  
  Exit Sub
  
  
PrintRptHeader2:
  Page = Page + 1
  Print #RptHandle, Tab(21); "Business License : Transactions Journal"
  Print #RptHandle, Tab(HeadLen1); "Transaction Type: " + ThisTransType
  Print #RptHandle, Tab(29); "Report Date: "; Date$; Tab(75); "Page #"; Page
  Print #RptHandle, ""
  Print #RptHandle, "Beginning Date: "; BegDate$
  Print #RptHandle, "   Ending Date: "; EndDate$
  If fpcmbCategory.Text = "Charge Licenses" Then
    Print #RptHandle, "*Amounts Include Issuance Fees"
  Else
    Print #RptHandle, ""
  End If
  If LicAndPenFlag = False Then
    Print #RptHandle, "  Date"; Tab(13); "Customer Name"; Tab(50); "Description"; Tab(75); "Amount"
    Print #RptHandle, String$(80, "=")
  Else
    Print #RptHandle, "  Date"; Tab(13); "Customer Name"; Tab(40); "Desc"; Tab(54); "Pen Total"; Tab(66); "Iss Total"; Tab(79); "Lic Total"; Tab(95); "Total"
    Print #RptHandle, String$(99, "=")
  End If
  LineCnt = 9
  Return
  
PrintRptEnding2:
  If LicAndPenFlag = False Then
    Print #RptHandle, String$(80, "-")
    Print #RptHandle, "Totals:"
    Print #RptHandle,
  Else
    Print #RptHandle, FF$
    Page = Page + 1
    Print #RptHandle, Tab(21); "Business License : Transactions Journal"
    Print #RptHandle, Tab(HeadLen1); "Transaction Type: " + ThisTransType
    Print #RptHandle, Tab(29); "Report Date: "; Date$; Tab(65); "Page #"; Page
    Print #RptHandle, Tab(34); "Report Totals"
    Print #RptHandle, ""
    Print #RptHandle, "Beginning Date: "; BegDate$
    Print #RptHandle, "   Ending Date: "; EndDate$
    Print #RptHandle, ""
    Print #RptHandle, Tab(3); "Transaction Type"; Tab(25); "# Transactions"; Tab(47); "Pen Total"; Tab(60); "Iss Total"; Tab(73); "Lic Total"; Tab(93); "Total"
    Print #RptHandle, String$(97, "-")
  End If
  
  For cnt = 1 To 101
    If GTotalAmt#(cnt) <> 0 Then
      Print #RptHandle, Tab(3);
      Select Case cnt
      Case 1
        Print #RptHandle, "License Charges";
      Case 2
        Print #RptHandle, "Payments";
      Case 6
        Print #RptHandle, "Penalty Charges";
      Case 9
        Print #RptHandle, "Beg Bal  ";
      Case 13
        Print #RptHandle, "DOWN Payment Adjustment";
      Case 23
        Print #RptHandle, "DOWN Bill Adjustment";
      Case 24
        Print #RptHandle, "UP Bill Adjustment";
    
      End Select
      If LicAndPenFlag = False Then
        Print #RptHandle, Tab(45); "# Transactions: " + CStr(TypeCnt(cnt)); Tab(68); Using("$#,###,##0.00", GTotalAmt#(cnt))
      Else
        PTotal = PTotal + PTotalAmt#(cnt)
        ITotal = ITotal + ITotalAmt#(cnt)
        LTotal = LTotal + LTotalAmt#(cnt)
        LPTotal = LPTotal + GTotalAmt#(cnt)
        LPCnt = LPCnt + TypeCnt(cnt)
        Print #RptHandle, Tab(28); Using("####0", CStr(TypeCnt(cnt))); Tab(45); Using("$###,##0.00", PTotalAmt#(cnt)); Tab(58); Using("$###,##0.00", ITotalAmt#(cnt)); Tab(69); Using("$#,###,##0.00", LTotalAmt#(cnt)); Tab(85); Using("$#,###,##0.00", GTotalAmt#(cnt))
      End If
    End If
  Next cnt
  If LicAndPenFlag = True Then
    Print #RptHandle, String$(97, "-")
    Print #RptHandle, Tab(3); "Grand Totals: "; Tab(28); Using("####0", CStr(LPCnt)); Tab(43); Using("$#,###,##0.00", PTotal); Tab(59); Using("$##,##0.00", ITotal); Tab(69); Using("$#,###,##0.00", LTotal); Tab(85); Using("$#,###,##0.00", LPTotal)
    Print #RptHandle, FF$
  End If
  
  Return

GetReportInformation2:
  TRNumRecs = LOF(THandle) / Len(TransRec)
  frmBLShowPctComp.Label1 = "Building Index"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  cmdHelp.Enabled = False
  'we need to find the transactions that pertain to the category
  'selected by the user
  For cnt = 1 To TRNumRecs
    Get THandle, cnt, TransRec
    'start by sorting transactions between those done before
    'using this version of business license and those conducted
    'using this version...if there is a value in the DetailTransType
    'then the transaction took place using this version
    If TransRec.DetailTransType = 0 And TransRec.TransType > 0 Then
      If QPTrim$(fpcmbCategory.Text) = "All Adjustments Down" Then
        If TransRec.TransType = 100 Then 'not available in this version
          'sort out non qualifying dates
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            'TransWho is used in sorting transactions
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            'TransRecNum is used in getting customer data for this type of transaction only
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = 0
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "All Adjustments Up" Then
        If TransRec.TransType = 101 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = 0
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "All Charges" Then
        If TransRec.TransType = 1 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = 0
            NCnt = NCnt + 1
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "All Payments" Then
        If TransRec.TransType = 2 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = 0
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "All Penalties" Then
        If TransRec.TransType = 6 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = 0
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "Charge Licenses" Then
        If TransRec.TransType = 1 Then
          If TransRec.TransAmount > 0 Then
            If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
              CountNum = CountNum + 1
              ReDim Preserve TransIdx(1 To CountNum)
              TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
              TransIdx(CountNum).TransRecNum = cnt
              TransIdx(CountNum).TransAmt = TransRec.TransAmount
            End If
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "Pay Licenses" Then
        If TransRec.TransType = 2 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = TransRec.TransAmount
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "License Adjustments Down" Then
        If TransRec.TransType = 100 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = TransRec.TransAmount
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "License Adjustments Up" Then
        If TransRec.TransType = 101 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = TransRec.TransAmount
          End If
        End If
      End If
    ElseIf TransRec.DetailTransType > 0 Then 'DetailTransType was not used
    'before this version of business license
      If QPTrim$(fpcmbCategory.Text) = "All Adjustments Down" Then
        LicAndPenFlag = True 'this flag indicates a type of report that requires more detail and
        'is true for all categories with 'All' in it's caption
        If TransRec.DetailTransType = 310 Or TransRec.DetailTransType = 301 Or TransRec.DetailTransType = 311 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = 0
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "Penalty Adjustments Down" Then
        If TransRec.DetailTransType = 301 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = TransRec.PenAmt
          End If
        ElseIf TransRec.DetailTransType = 311 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = TransRec.PenAmt
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "Penalty Adjustments Up" Then
        If TransRec.DetailTransType = 401 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = TransRec.PenAmt
          End If
        ElseIf TransRec.DetailTransType = 411 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = TransRec.PenAmt
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "All Adjustments Up" Then
        LicAndPenFlag = True
        If TransRec.DetailTransType = 410 Or TransRec.DetailTransType = 401 Or TransRec.DetailTransType = 411 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = 0
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "License Adjustments Down" Then
        If TransRec.DetailTransType = 310 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = TransRec.LicAmt
          End If
        ElseIf TransRec.DetailTransType = 311 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = TransRec.LicAmt
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "Issuance Adjustments Down" Then
        If TransRec.DetailTransType = 310 Then
          If TransRec.IssAmt > 0 Then
            If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
              CountNum = CountNum + 1
              ReDim Preserve TransIdx(1 To CountNum)
              TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
              TransIdx(CountNum).TransRecNum = cnt
              TransIdx(CountNum).TransAmt = TransRec.IssAmt
            End If
          End If
        ElseIf TransRec.DetailTransType = 311 Then
          If TransRec.IssAmt > 0 Then
            If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
              CountNum = CountNum + 1
              ReDim Preserve TransIdx(1 To CountNum)
              TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
              TransIdx(CountNum).TransRecNum = cnt
              TransIdx(CountNum).TransAmt = TransRec.IssAmt
            End If
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "License Adjustments Up" Then
        If TransRec.DetailTransType = 410 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = TransRec.LicAmt
          End If
        ElseIf TransRec.DetailTransType = 411 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = TransRec.LicAmt
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "Issuance Adjustments Up" Then
        If TransRec.DetailTransType = 410 Then
          If TransRec.IssAmt > 0 Then
            If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
              CountNum = CountNum + 1
              ReDim Preserve TransIdx(1 To CountNum)
              TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
              TransIdx(CountNum).TransRecNum = cnt
              TransIdx(CountNum).TransAmt = TransRec.IssAmt
            End If
          End If
        ElseIf TransRec.DetailTransType = 411 Then
          If TransRec.IssAmt > 0 Then
            If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
              CountNum = CountNum + 1
              ReDim Preserve TransIdx(1 To CountNum)
              TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
              TransIdx(CountNum).TransRecNum = cnt
              TransIdx(CountNum).TransAmt = TransRec.IssAmt
            End If
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "All Charges" Then
        LicAndPenFlag = True
        If TransRec.DetailTransType = 101 Or TransRec.DetailTransType = 110 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = 0
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "Charge Penalties" Then
        If TransRec.DetailTransType = 101 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = TransRec.PenAmt
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "Charge Licenses" Then
        If TransRec.DetailTransType = 110 Then
          If TransRec.LicAmt > 0 Then
            If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
              CountNum = CountNum + 1
              ReDim Preserve TransIdx(1 To CountNum)
              TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
              TransIdx(CountNum).TransRecNum = cnt
              TransIdx(CountNum).TransAmt = TransRec.LicAmt
            End If
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "Charge Issuance Fees" Then
        If TransRec.DetailTransType = 110 Then
          If TransRec.IssAmt > 0 Then
            If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
              CountNum = CountNum + 1
              ReDim Preserve TransIdx(1 To CountNum)
              TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
              TransIdx(CountNum).TransRecNum = cnt
              TransIdx(CountNum).TransAmt = TransRec.IssAmt
            End If
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "All Payments" Then
        LicAndPenFlag = True
        If TransRec.DetailTransType = 201 Or TransRec.DetailTransType = 210 Or TransRec.DetailTransType = 211 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = 0
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "Pay Penalties" Then
        If TransRec.DetailTransType = 201 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = TransRec.PenAmt
          End If
        ElseIf TransRec.DetailTransType = 211 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = TransRec.PenAmt
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "Pay Licenses" Then
        If TransRec.DetailTransType = 210 Then
          If TransRec.LicAmt > 0 Then
            If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
              CountNum = CountNum + 1
              ReDim Preserve TransIdx(1 To CountNum)
              TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
              TransIdx(CountNum).TransRecNum = cnt
              TransIdx(CountNum).TransAmt = TransRec.LicAmt
            End If
          End If
        ElseIf TransRec.DetailTransType = 211 Then
          If TransRec.LicAmt > 0 Then
            If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
              CountNum = CountNum + 1
              ReDim Preserve TransIdx(1 To CountNum)
              TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
              TransIdx(CountNum).TransRecNum = cnt
              TransIdx(CountNum).TransAmt = TransRec.LicAmt
            End If
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "Pay Issuance Fees" Then
        If TransRec.DetailTransType = 210 Then
          If TransRec.IssAmt > 0 Then
            If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
              CountNum = CountNum + 1
              ReDim Preserve TransIdx(1 To CountNum)
              TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
              TransIdx(CountNum).TransRecNum = cnt
              TransIdx(CountNum).TransAmt = TransRec.IssAmt
            End If
          End If
        ElseIf TransRec.DetailTransType = 211 Then
          If TransRec.IssAmt > 0 Then
            If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
              CountNum = CountNum + 1
              ReDim Preserve TransIdx(1 To CountNum)
              TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
              TransIdx(CountNum).TransRecNum = cnt
              TransIdx(CountNum).TransAmt = TransRec.IssAmt
            End If
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "All Penalties" Then
        LicAndPenFlag = True
        'penalty activity shows up in each of the following
        'detailtranstypes
        If TransRec.DetailTransType = 101 Or TransRec.DetailTransType = 200 _
        Or TransRec.DetailTransType = 301 Or TransRec.DetailTransType = 401 _
        Or TransRec.DetailTransType = 211 Or TransRec.DetailTransType = 311 _
        Or TransRec.DetailTransType = 411 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = 0
          End If
        End If
      End If
    End If
    frmBLShowPctComp.ShowPctComp cnt, TRNumRecs
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
  Next cnt
  
  If CountNum = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "There are no transactions saved between " + fptxtBDate.Text + " and " + fptxtEDate.Text + " for " + QPTrim$(fpcmbCategory.Text) + "."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    Close
    EnableCloseButton Me.hwnd, True
    cmdExit.Enabled = True
    cmdProcess.Enabled = True
    cmdHelp.Enabled = True
    Exit Sub
  End If
  
  'sort transactions based on the transaction date and customer number
  BigNum = "A"
  For x = 1 To CountNum
    If QPTrim$(TransIdx(x).TransWho) > BigNum Then
      BigNum = QPTrim$(TransIdx(x).TransWho)
    End If
  Next x
  
  SmallNum = BigNum + "A"
  Nextx = 1
  frmBLShowPctComp.Label1 = "Sorting Transactions"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  cmdHelp.Enabled = False
  
  Do
    For x = Nextx To CountNum
      If QPTrim$(TransIdx(x).TransWho) < SmallNum Then
        SmallNum = QPTrim$(TransIdx(x).TransWho)
        ThisRec = x
      End If
    Next x
    HoldThis = TransIdx(Nextx)
    TransIdx(Nextx) = TransIdx(ThisRec)
    TransIdx(ThisRec) = HoldThis
    If Nextx = CountNum Then Exit Do
    SmallNum = BigNum + "A"
    Nextx = Nextx + 1
    frmBLShowPctComp.ShowPctComp Nextx, CountNum
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
  Loop
  'at this point you have been garnered all pertinent transactions
  'and sorted this in the array TransIdx
  Return
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLTransJournal", "PrintText", Erl)
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

Private Sub PrintGraphics()
  Dim BegDate$
  Dim BegDateNum As Integer
  Dim EndDate$
  Dim EndDateNum As Integer
  Dim ReportFile$
  Dim SubReportFile$
  Dim RptHandle As Integer
  Dim SubRptHandle As Integer
  Dim TransCnt As Double
  Dim TransRec As ARTransRecType
  Dim THandle As Integer
  Dim cnt As Double
  Dim TotalTrans As Double
  Dim TotalAmt As Double
  Dim TotalPaid As Double
  Dim FeePd As Double
  Dim Category$
  Dim LeftOver As Double
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim BILLCAT1$
  Dim BILLCAT2$
  Dim BILLCAT3$
  Dim BILLCAT4$
  Dim BILLCAT5$
  Dim Fee1#
  Dim Fee2#
  Dim Fee3#
  Dim Fee4#
  Dim Fee5#
  Dim CatCnt!, CatFnd!
  Dim ll As Double
  Dim CategoryDesc$
  Dim CodeRec As ARNewCatCodeRecType
  Dim NumOfARCatRecs As Integer
  Dim COHandle As Integer
  Dim LCnt As Integer
  Dim TRNumRecs As Double
  Dim CountNum As Double
  Dim BigNum$
  Dim HoldThis As TransIdxType
  Dim ThisRec As Double
  Dim SmallNum$
  Dim Nextx As Double
  Dim x As Double
  Dim dlm$
  Dim TownName$
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim PrintCnt As Integer
  Dim ZCnt As Integer
  Dim LicAndPenFlag As Boolean
  Dim ThisTotal As Double
  
  On Error GoTo ERRORSTUFF
  
  cmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  
  LicAndPenFlag = False
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  TownName$ = QPTrim$(TownRec.TownName)
  dlm$ = "~"
  
  ReDim Cat$(300), CatAmt#(300), GTotalAmt#(103), TypeCnt%(103), TypeDesc$(103)
  
  BegDate = fptxtBDate.Text
  BegDateNum = Date2Num(fptxtBDate.Text)
  EndDate = fptxtEDate.Text
  EndDateNum = Date2Num(fptxtEDate.Text)
  If EndDateNum < BegDateNum Then
    fptxtEDate.BackColor = &HFFFF&
    fptxtBDate.BackColor = &HFFFF&
    frmBLMessageBoxJr.Label1.Caption = "The ending date comes before the beginning date. Please re-enter these values."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    fptxtEDate.BackColor = &HFFFFFF
    fptxtBDate.BackColor = &HFFFFFF
    fptxtBDate.SetFocus
    Exit Sub
  End If
  
  OpenTransFile THandle 'opened here and used also in GetReportInformation2

  ReDim TransIdx(1 To 1) As TransIdxType
  'GetReportInformation2 is used to determine which type
  'of report will be needed (one is more detailed than the other
  'depending on if the report wants 'All' data or more specific
  'data) plus transaction data is scoured for valid entries and
  'the transaction data is sorted by date and customer number
  GoSub GetReportInformation2
  
  If LicAndPenFlag = False Then
    ReportFile$ = "BLRPTS\ARTRANS.RPT"
    SubReportFile$ = "BLRPTS\ARTYPSUB.RPT"
  Else
    ReportFile$ = "BLRPTS\ARTRNLP.RPT"
    SubReportFile$ = "BLRPTS\ARTYSBLP.RPT"
  End If
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  OpenCustFile CHandle
  TransCnt = LOF(THandle) / Len(TransRec)
  frmBLShowPctComp.Label1 = "Loading Transaction List"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  cmdHelp.Enabled = False
  
  If LicAndPenFlag = True Then GoTo LicAndPenCalc
  'LicAndPenFlag is set in GetReportInformation2 and TransIdx()
  'is set in GetReportInformation2
  For cnt = CountNum To 1 Step -1
    Get THandle, TransIdx(cnt).TransRecNum, TransRec
    If Val(TransRec.CustomerNumber) = 0 Then
      GoTo BadCustSkip
    End If
    'Get Customer
    Get CHandle, Val(TransRec.CustomerNumber), CustRec
    If QPTrim$(CustRec.Deleted) = "Y" Then GoTo BadCustSkip
    If TransIdx(cnt).TransAmt > 0 Then
      ThisTotal = TransIdx(cnt).TransAmt
    Else
      ThisTotal = TransRec.TransAmount
    End If
    '                     0
    Print #RptHandle, TownName$; dlm;
    '                           1
    Print #RptHandle, MakeRegDate(TransRec.TransDate); dlm;
    '                           2
    Print #RptHandle, QPTrim$(CustRec.CustName); dlm;
    'now print the type of transaction
    Select Case TransRec.TransType
      Case 1
        '                    3
        Print #RptHandle, "License Charges"; dlm;
        PrintCnt = PrintCnt + 1
      Case 2
        '                    3
        Print #RptHandle, "Payment"; dlm;
        PrintCnt = PrintCnt + 1
      Case 6
        '                    3
        Print #RptHandle, "Penalty Charges"; dlm;
        PrintCnt = PrintCnt + 1
      Case 9
        '                    3
        Print #RptHandle, "Beg Bal"; dlm;
        PrintCnt = PrintCnt + 1
      Case 13
        '                    3
        Print #RptHandle, "DOWN Pay Adjustment"; dlm;
        PrintCnt = PrintCnt + 1
      Case 23
        '                    3
        Print #RptHandle, "DOWN Bill Adjustment"; dlm;
        PrintCnt = PrintCnt + 1
      Case 24
        '                    3
        Print #RptHandle, "UP Bill Adjustment"; dlm;
        PrintCnt = PrintCnt + 1
      Case 100
        '                    3
        Print #RptHandle, "DOWN Adjustment"; dlm;
        PrintCnt = PrintCnt + 1
      Case 101
        '                    3
        Print #RptHandle, "UP Adjustment"; dlm;
        PrintCnt = PrintCnt + 1
      Case Else
    End Select
    
    'print
    '                 4
    Print #RptHandle, ""; dlm;
    '                          5                     6             7                         8                              9
    Print #RptHandle, ThisTotal; dlm; BegDate$; dlm; EndDate$; dlm; QPTrim$(fpcmbCategory.Text); dlm; CStr(TransRec.DetailTransType)
    'now gather summary data
    TotalTrans = TotalTrans + 1
    TotalAmt# = TotalAmt# + ThisTotal
    TypeDesc$(TransRec.TransType) = QPTrim$(TransRec.TransDesc)
    TypeCnt%(TransRec.TransType) = TypeCnt%(TransRec.TransType) + 1
    Rem total by category
TotalUp:
    GTotalAmt#(TransRec.TransType) = GTotalAmt#(TransRec.TransType) + ThisTotal
    
    
BadCustSkip:
    frmBLShowPctComp.ShowPctComp cnt, CountNum
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
    
  Next cnt
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  cmdHelp.Enabled = True
  
  
  Close         'Close all open files now
  
  If PrintCnt = 0 Then
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Label1.Caption = "There are no " + QPTrim$(fpcmbCategory.Text) + " transactions on file."
    frmBLMessageBoxJr.Show vbModal
  Else
    GoSub PrintEnd
    arBLTransJournal.Show
    frmBLLoadReport.Show
  End If
  
  Exit Sub
  
LicAndPenCalc:
  'these reports are more detailed
  ReDim LTotalAmt#(103), PTotalAmt#(103), ITotalAmt#(103)
  For cnt = CountNum To 1 Step -1
    Get THandle, TransIdx(cnt).TransRecNum, TransRec
    If Val(TransRec.CustomerNumber) = 0 Then
      GoTo BadCustSkip2
    End If
    
    'Get Customer
    Get CHandle, Val(TransRec.CustomerNumber), CustRec
    If QPTrim$(CustRec.Deleted) = "Y" Then GoTo BadCustSkip2
    '                     0
    Print #RptHandle, TownName$; dlm;
    '                           1
    Print #RptHandle, MakeRegDate(TransRec.TransDate); dlm;
    '                           2
    Print #RptHandle, QPTrim$(CustRec.CustName); dlm;
    Select Case TransRec.TransType
      Case 1
        '                    3
        Print #RptHandle, "License Charges"; dlm;
        PrintCnt = PrintCnt + 1
      Case 2
        '                    3
        Print #RptHandle, "Payment"; dlm;
        PrintCnt = PrintCnt + 1
      Case 6
        '                    3
        Print #RptHandle, "Penalty Charges"; dlm;
        PrintCnt = PrintCnt + 1
      Case 9
        '                    3
        Print #RptHandle, "Beg Bal"; dlm;
        PrintCnt = PrintCnt + 1
      Case 13
        '                    3
        Print #RptHandle, "DOWN Pay Adjustment"; dlm;
        PrintCnt = PrintCnt + 1
      Case 23
        '                    3
        Print #RptHandle, "DOWN Bill Adjustment"; dlm;
        PrintCnt = PrintCnt + 1
      Case 24
        '                    3
        Print #RptHandle, "UP Bill Adjustment"; dlm;
        PrintCnt = PrintCnt + 1
      Case 100
        '                    3
        Print #RptHandle, "DOWN Adjustment"; dlm;
        PrintCnt = PrintCnt + 1
      Case 101
        '                    3
        Print #RptHandle, "UP Adjustment"; dlm;
        PrintCnt = PrintCnt + 1
      Case Else
        '                    3
        Print #RptHandle, "Unknown"; dlm;
        PrintCnt = PrintCnt + 1
    End Select
    
    'print
    '                          4
    Print #RptHandle, ""; dlm;
    '                          5                     6             7                         8
    Print #RptHandle, TransRec.TransAmount; dlm; BegDate$; dlm; EndDate$; dlm; QPTrim$(fpcmbCategory.Text); dlm;
    '                          9                  10                   11
    Print #RptHandle, TransRec.LicAmt; dlm; TransRec.PenAmt; dlm; TransRec.IssAmt
    
    TotalTrans = TotalTrans + 1
    TotalAmt# = TotalAmt# + TransRec.TransAmount
    TypeDesc$(TransRec.TransType) = QPTrim$(TransRec.TransDesc)
    TypeCnt%(TransRec.TransType) = TypeCnt%(TransRec.TransType) + 1
    Rem total by category
TotalUp2: 'G = Grand Total, L = License Totals, P = Penalty Totals and I = Issuance Fee Totals
    GTotalAmt#(TransRec.TransType) = GTotalAmt#(TransRec.TransType) + TransRec.TransAmount
    LTotalAmt#(TransRec.TransType) = LTotalAmt#(TransRec.TransType) + TransRec.LicAmt
    PTotalAmt#(TransRec.TransType) = PTotalAmt#(TransRec.TransType) + TransRec.PenAmt
    ITotalAmt#(TransRec.TransType) = ITotalAmt#(TransRec.TransType) + TransRec.IssAmt
    
BadCustSkip2:
    frmBLShowPctComp.ShowPctComp cnt, CountNum
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
    
  Next cnt
  
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  cmdHelp.Enabled = True
  
  Close         'Close all open files now
  
  If PrintCnt = 0 Then
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Label1.Caption = "There are no " + QPTrim$(fpcmbCategory.Text) + " transactions on file."
    frmBLMessageBoxJr.Show vbModal
  Else
    GoSub PrintEndLP
    arBLTransJrnalLP.Show
    frmBLLoadReport.Show
  End If
  
  MainLog ("'Transaction By Type' report processed for " + QPTrim$(fpcmbCategory.Text) + " beginning on " + fptxtBDate + " thru " + fptxtEDate + " in graphics format.")
  
  Exit Sub

GetReportInformation2: 'see PrintText for more explanation
  'for GetReportInformation2
  TRNumRecs = LOF(THandle) / Len(TransRec)
  frmBLShowPctComp.Label1 = "Building Index"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  cmdHelp.Enabled = False
  
  For cnt = 1 To TRNumRecs
    Get THandle, cnt, TransRec
    If TransRec.DetailTransType = 0 And TransRec.TransType > 0 Then
      If QPTrim$(fpcmbCategory.Text) = "All Adjustments Down" Then
        If TransRec.TransType = 100 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = 0
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "All Adjustments Up" Then
        If TransRec.TransType = 101 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = 0
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "All Charges" Then
        If TransRec.TransType = 1 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = 0
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "All Payments" Then
        If TransRec.TransType = 2 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = 0
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "All Penalties" Then
        If TransRec.TransType = 6 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = 0
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "Charge Licenses" Then
        If TransRec.TransType = 1 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = TransRec.TransAmount
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "Pay Licenses" Then
        If TransRec.TransType = 2 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = TransRec.TransAmount
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "License Adjustments Down" Then
        If TransRec.TransType = 100 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = TransRec.TransAmount
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "License Adjustments Up" Then
        If TransRec.TransType = 101 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = TransRec.TransAmount
          End If
        End If
      End If
    '----------------------------------------------------------
    ElseIf TransRec.DetailTransType > 0 Then 'occurs only after this version
    'of business license is installed
      If QPTrim$(fpcmbCategory.Text) = "All Adjustments Down" Then
        LicAndPenFlag = True
        If TransRec.DetailTransType = 310 Or TransRec.DetailTransType = 301 Or TransRec.DetailTransType = 311 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = 0
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "Penalty Adjustments Down" Then
        If TransRec.DetailTransType = 301 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = TransRec.PenAmt
          End If
        ElseIf TransRec.DetailTransType = 311 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = TransRec.PenAmt
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "Penalty Adjustments Up" Then
        If TransRec.DetailTransType = 401 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = TransRec.PenAmt
          End If
        ElseIf TransRec.DetailTransType = 411 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = TransRec.PenAmt
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "All Adjustments Up" Then
        LicAndPenFlag = True
        If TransRec.DetailTransType = 410 Or TransRec.DetailTransType = 401 Or TransRec.DetailTransType = 411 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = 0
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "License Adjustments Down" Then
        If TransRec.DetailTransType = 310 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = TransRec.LicAmt
          End If
        ElseIf TransRec.DetailTransType = 311 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = TransRec.LicAmt
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "Issuance Adjustments Down" Then
        If TransRec.DetailTransType = 310 Then
          If TransRec.IssAmt > 0 Then
            If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
              CountNum = CountNum + 1
              ReDim Preserve TransIdx(1 To CountNum)
              TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
              TransIdx(CountNum).TransRecNum = cnt
              TransIdx(CountNum).TransAmt = TransRec.IssAmt
            End If
          End If
        ElseIf TransRec.DetailTransType = 311 Then
          If TransRec.IssAmt > 0 Then
            If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
              CountNum = CountNum + 1
              ReDim Preserve TransIdx(1 To CountNum)
              TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
              TransIdx(CountNum).TransRecNum = cnt
              TransIdx(CountNum).TransAmt = TransRec.IssAmt
            End If
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "License Adjustments Up" Then
        If TransRec.DetailTransType = 410 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = TransRec.LicAmt
          End If
        ElseIf TransRec.DetailTransType = 411 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = TransRec.LicAmt
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "Issuance Adjustments Up" Then
        If TransRec.DetailTransType = 410 Then
          If TransRec.IssAmt > 0 Then
            If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
              CountNum = CountNum + 1
              ReDim Preserve TransIdx(1 To CountNum)
              TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
              TransIdx(CountNum).TransRecNum = cnt
              TransIdx(CountNum).TransAmt = TransRec.IssAmt
            End If
          End If
        ElseIf TransRec.DetailTransType = 411 Then
          If TransRec.IssAmt > 0 Then
            If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
              CountNum = CountNum + 1
              ReDim Preserve TransIdx(1 To CountNum)
              TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
              TransIdx(CountNum).TransRecNum = cnt
              TransIdx(CountNum).TransAmt = TransRec.IssAmt
            End If
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "All Charges" Then
        LicAndPenFlag = True
        If TransRec.DetailTransType = 101 Or TransRec.DetailTransType = 110 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = 0
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "Charge Penalties" Then
        If TransRec.DetailTransType = 101 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = TransRec.PenAmt
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "Charge Licenses" Then
        If TransRec.DetailTransType = 110 Then
          If TransRec.LicAmt > 0 Then 'in case issuance fee was the only charge
            If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
              CountNum = CountNum + 1
              ReDim Preserve TransIdx(1 To CountNum)
              TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
              TransIdx(CountNum).TransRecNum = cnt
              TransIdx(CountNum).TransAmt = TransRec.LicAmt
            End If
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "Charge Issuance Fees" Then
        If TransRec.DetailTransType = 110 Then
          If TransRec.IssAmt > 0 Then
            If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
              CountNum = CountNum + 1
              ReDim Preserve TransIdx(1 To CountNum)
              TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
              TransIdx(CountNum).TransRecNum = cnt
              TransIdx(CountNum).TransAmt = TransRec.IssAmt
            End If
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "All Payments" Then
        LicAndPenFlag = True
        If TransRec.DetailTransType = 201 Or TransRec.DetailTransType = 210 Or TransRec.DetailTransType = 211 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = 0
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "Pay Penalties" Then
        If TransRec.DetailTransType = 201 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = TransRec.PenAmt
          End If
        ElseIf TransRec.DetailTransType = 211 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = TransRec.PenAmt
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "Pay Licenses" Then
        If TransRec.DetailTransType = 210 Then
          If TransRec.LicAmt > 0 Then
            If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
              CountNum = CountNum + 1
              ReDim Preserve TransIdx(1 To CountNum)
              TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
              TransIdx(CountNum).TransRecNum = cnt
              TransIdx(CountNum).TransAmt = TransRec.LicAmt
            End If
          End If
        ElseIf TransRec.DetailTransType = 211 Then
          If TransRec.LicAmt > 0 Then
            If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
              CountNum = CountNum + 1
              ReDim Preserve TransIdx(1 To CountNum)
              TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
              TransIdx(CountNum).TransRecNum = cnt
              TransIdx(CountNum).TransAmt = TransRec.LicAmt
            End If
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "Pay Issuance Fees" Then
        If TransRec.DetailTransType = 210 Then
          If TransRec.IssAmt > 0 Then
            If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
              CountNum = CountNum + 1
              ReDim Preserve TransIdx(1 To CountNum)
              TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
              TransIdx(CountNum).TransRecNum = cnt
              TransIdx(CountNum).TransAmt = TransRec.IssAmt
            End If
          End If
        ElseIf TransRec.DetailTransType = 211 Then
          If TransRec.IssAmt > 0 Then
            If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
              CountNum = CountNum + 1
              ReDim Preserve TransIdx(1 To CountNum)
              TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
              TransIdx(CountNum).TransRecNum = cnt
              TransIdx(CountNum).TransAmt = TransRec.IssAmt
            End If
          End If
        End If
      End If
      If QPTrim$(fpcmbCategory.Text) = "All Penalties" Then
        LicAndPenFlag = True
        If TransRec.DetailTransType = 101 Or TransRec.DetailTransType = 200 _
        Or TransRec.DetailTransType = 301 Or TransRec.DetailTransType = 401 _
        Or TransRec.DetailTransType = 211 Or TransRec.DetailTransType = 311 _
        Or TransRec.DetailTransType = 411 Then
          If TransRec.TransDate >= BegDateNum And TransRec.TransDate <= EndDateNum Then
            CountNum = CountNum + 1
            ReDim Preserve TransIdx(1 To CountNum)
            TransIdx(CountNum).TransWho = LTrim$(Str$(TransRec.TransDate)) + TransRec.CustomerNumber + String$(4, " ")
            TransIdx(CountNum).TransRecNum = cnt
            TransIdx(CountNum).TransAmt = 0
          End If
        End If
      End If
    frmBLShowPctComp.ShowPctComp cnt, TRNumRecs
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
  End If
  Next cnt
  
  If CountNum = 0 Then
    Unload frmBLShowPctComp
    frmBLMessageBoxJr.Label1.Caption = "There are no transactions saved between " + fptxtBDate.Text + " and " + fptxtEDate.Text + " for " + QPTrim$(fpcmbCategory.Text) + "."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    Close
    EnableCloseButton Me.hwnd, True
    cmdExit.Enabled = True
    cmdProcess.Enabled = True
    cmdHelp.Enabled = True
    Exit Sub
  End If
  
  BigNum = "A"
  For x = 1 To CountNum
    If QPTrim$(TransIdx(x).TransWho) > BigNum Then
      BigNum = QPTrim$(TransIdx(x).TransWho)
    End If
  Next x
  
  SmallNum = BigNum + "A"
  Nextx = 1
  frmBLShowPctComp.Label1 = "Sorting Transactions"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  cmdHelp.Enabled = False
  
  Do
    For x = Nextx To CountNum
      If QPTrim$(TransIdx(x).TransWho) < SmallNum Then
        SmallNum = QPTrim$(TransIdx(x).TransWho)
        ThisRec = x
      End If
    Next x
    HoldThis = TransIdx(Nextx)
    TransIdx(Nextx) = TransIdx(ThisRec)
    TransIdx(ThisRec) = HoldThis
    If Nextx = CountNum Then Exit Do
    SmallNum = BigNum + "A"
    Nextx = Nextx + 1
    frmBLShowPctComp.ShowPctComp Nextx, CountNum
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
  Loop
  
  Return
  
PrintEnd:
  SubRptHandle = FreeFile
  Open SubReportFile$ For Output As #SubRptHandle
  For x = 1 To 103
    If GTotalAmt#(x) > 0 Then
    Select Case x
      Case 1
        '                    3
        Print #SubRptHandle, "License Charges"; dlm;
        PrintCnt = PrintCnt + 1
      Case 2
        '                    3
        Print #SubRptHandle, "Payment"; dlm;
        PrintCnt = PrintCnt + 1
      Case 6
        '                    3
        Print #SubRptHandle, "Penalty Charges"; dlm;
        PrintCnt = PrintCnt + 1
      Case 9
        '                    3
        Print #SubRptHandle, "Beg Bal"; dlm;
        PrintCnt = PrintCnt + 1
      Case 13
        '                    3
        Print #SubRptHandle, "DOWN Pay Adjustment"; dlm;
        PrintCnt = PrintCnt + 1
      Case 23
        '                    3
        Print #SubRptHandle, "DOWN Bill Adjustment"; dlm;
        PrintCnt = PrintCnt + 1
      Case 24
        '                    3
        Print #SubRptHandle, "UP Bill Adjustment"; dlm;
        PrintCnt = PrintCnt + 1
      Case 100
        '                    3
        Print #SubRptHandle, "DOWN Adjustment"; dlm;
        PrintCnt = PrintCnt + 1
      Case 101
        '                    3
        Print #SubRptHandle, "UP Adjustment"; dlm;
        PrintCnt = PrintCnt + 1
      Case Else
        '                    3
        Print #SubRptHandle, "Unknown"; dlm;
        PrintCnt = PrintCnt + 1
    End Select
      Print #SubRptHandle, CStr(GTotalAmt#(x)); dlm; CStr(TypeCnt%(x))
    End If
  Next x
  
  Close SubRptHandle
  Return
  
PrintEndLP:
  SubRptHandle = FreeFile
  Open SubReportFile$ For Output As #SubRptHandle
  For x = 1 To 103
    If GTotalAmt#(x) > 0 Then
    Select Case x
      Case 1
        '                    3
        Print #SubRptHandle, "Lic Charge"; dlm;
        PrintCnt = PrintCnt + 1
      Case 2
        '                    3
        Print #SubRptHandle, "Payment"; dlm;
        PrintCnt = PrintCnt + 1
      Case 6
        '                    3
        Print #SubRptHandle, "Pen Charge"; dlm;
        PrintCnt = PrintCnt + 1
      Case 9
        '                    3
        Print #SubRptHandle, "Beg Bal"; dlm;
        PrintCnt = PrintCnt + 1
      Case 13
        '                    3
        Print #SubRptHandle, "DOWN Pay Adjustment"; dlm;
        PrintCnt = PrintCnt + 1
      Case 23
        '                    3
        Print #SubRptHandle, "DOWN Bill Adjustment"; dlm;
        PrintCnt = PrintCnt + 1
      Case 24
        '                    3
        Print #SubRptHandle, "UP Bill Adjustment"; dlm;
        PrintCnt = PrintCnt + 1
      Case 100
        '                    3
        Print #SubRptHandle, "DOWN Adjustment"; dlm;
        PrintCnt = PrintCnt + 1
      Case 101
        '                    3
        Print #SubRptHandle, "UP Adjustment"; dlm;
        PrintCnt = PrintCnt + 1
      Case Else
    End Select
      Print #SubRptHandle, CStr(GTotalAmt#(x)); dlm; CStr(TypeCnt%(x)); dlm; CStr(LTotalAmt#(x)); dlm; CStr(PTotalAmt#(x)); dlm; CStr(ITotalAmt(x))
    End If
  Next x
  
  Close SubRptHandle
  Return

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLTransJournal", "PrintGraphics", Erl)
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

