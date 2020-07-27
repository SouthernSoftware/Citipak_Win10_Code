VERSION 5.00
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "IMP32X30.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "EDT32X30.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Begin VB.Form frmFANewAddDelRpt 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Additions/Deletions"
   ClientHeight    =   8868
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   11652
   Icon            =   "frmFANewAddDelRpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8868
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   6156
      Left            =   1932
      TabIndex        =   8
      Top             =   1356
      Width           =   7788
      _Version        =   196609
      _ExtentX        =   13737
      _ExtentY        =   10858
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmFANewAddDelRpt.frx":08CA
      Begin LpLib.fpCombo fpcmbOrder 
         Height          =   384
         Left            =   3216
         TabIndex        =   0
         Top             =   1536
         Width           =   3228
         _Version        =   196608
         _ExtentX        =   5694
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
         ColDesigner     =   "frmFANewAddDelRpt.frx":08E6
      End
      Begin LpLib.fpCombo fpcomboPrintOpt 
         Height          =   384
         Left            =   3552
         TabIndex        =   5
         Top             =   4236
         Width           =   2364
         _Version        =   196608
         _ExtentX        =   4170
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
         ColDesigner     =   "frmFANewAddDelRpt.frx":0BA5
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "ESC &Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   684
         Left            =   1584
         TabIndex        =   6
         Top             =   5040
         Width           =   1884
      End
      Begin VB.CommandButton cmdProcess 
         Caption         =   "F10 &Process"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   684
         Left            =   4560
         TabIndex        =   7
         Top             =   5040
         Width           =   1884
      End
      Begin VB.CommandButton cmdDept 
         Caption         =   "F8 &Dept List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   396
         Left            =   4704
         TabIndex        =   2
         Top             =   2208
         Width           =   1356
      End
      Begin EditLib.fpText fptxtDeptNum 
         Height          =   396
         Left            =   3072
         TabIndex        =   1
         ToolTipText     =   $"frmFANewAddDelRpt.frx":0E64
         Top             =   2208
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
      Begin EditLib.fpDateTime fpDateStart 
         Height          =   444
         Left            =   3840
         TabIndex        =   3
         Top             =   2856
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
         Text            =   "01/24/2003"
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
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDateTime fpDateEnd 
         Height          =   444
         Left            =   3840
         TabIndex        =   4
         Top             =   3528
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
         Text            =   "01/24/2003"
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
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
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
         TabIndex        =   14
         Top             =   2304
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
         TabIndex        =   13
         Top             =   1584
         Width           =   1836
      End
      Begin VB.Label Label9 
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
         Left            =   2400
         TabIndex        =   12
         Top             =   3648
         Width           =   1212
      End
      Begin VB.Label Label8 
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
         Left            =   2304
         TabIndex        =   11
         Top             =   2976
         Width           =   1308
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
         Caption         =   "New Additions/Deletions"
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
         Height          =   492
         Left            =   1584
         TabIndex        =   10
         Top             =   576
         Width           =   4812
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
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
         Left            =   1872
         TabIndex        =   9
         Top             =   4320
         Width           =   1500
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   6348
      Left            =   1836
      Top             =   1260
      Width           =   7980
   End
End
Attribute VB_Name = "frmFANewAddDelRpt"
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
  DoEvents
  KillFile "newadddelrptopen.dat"
  Unload frmFANewAddDelRpt

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
  Dim CurDep#
  Dim MaxDep#
  Dim YTDDep#
  Dim x As Integer
  Dim Nextx As Integer
  Dim DIdxRec As DeptNumbSortIdxType
  Dim DIdxHandle As Integer
  Dim DIdxCnt As Integer
  Dim Index$, Tab1 As Integer, Tab2 As Integer, Tab3 As Integer
  Dim Page As Integer
  Dim DataFlag As Boolean
  Dim TagFlag As Boolean
  
  ReportFile$ = "FANEWADDDEL.PRN"  'Report File Name
  Dash80$ = String$(80, "=")
  FF$ = Chr$(12)
  
  MaxLines = 50
  LineCnt& = 0
  ItemCnt& = 0
  Dept$ = QPTrim$(fptxtDeptNum.Text)
  BDate = Date2Num(fpDateStart.Text)
  EDate = Date2Num(fpDateEnd.Text)
  
  RptHandle = FreeFile
  Index$ = QPTrim$(fpcmbOrder.Text)
  Open ReportFile$ For Output As #RptHandle
  
  GoSub PrintMasterHeader1
  
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
  ReDim DeptArr(1 To DIdxCnt + 1) As String
  
  For x = 1 To DIdxCnt
    Get DIdxHandle, x, DIdxRec
    DeptArr(x) = QPTrim$(DIdxRec.DeptNumb)
  Next x
  DeptArr(x) = ""
  Close DIdxHandle
  
  ReDim DTagDOrigCost(1 To DIdxCnt + 1) As Double
  ReDim DTagDBookTotal(1 To DIdxCnt + 1) As Double
  ReDim DTagDYDep(1 To DIdxCnt + 1) As Double
  ReDim ATagDOrigCost(1 To DIdxCnt + 1) As Double
  ReDim ATagDBookTotal(1 To DIdxCnt + 1) As Double
  ReDim ATagDYDep(1 To DIdxCnt + 1) As Double
  
  
  If Dept$ <> "ALL" Then
    DeptNumber = Val(QPTrim$(fptxtDeptNum.Text))
  Else
    DeptNumber = Val(QPTrim(DeptArr(1)))
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
      
      If DisposeDate >= BDate And DisposeDate <= EDate Or AcquireDate >= BDate And AcquireDate <= EDate Then
      'filter out items that don't fall inside the date parameters
        If DisposeDate >= BDate And DisposeDate <= EDate Then
          DFlag = True
        Else
          DFlag = False
        End If
        If AcquireDate >= BDate And AcquireDate <= EDate Then
          AFlag = True
        Else
          AFlag = False
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
      If TagFlag = True Then GoTo TagOnly2
      
TagOnly1: 'printing valid tag data...skipped if TAG NUMBER is chosen and this is
          'not the first iteration
      If DFlag = True Then Print #RptHandle, "D";
      If AFlag = True Then Print #RptHandle, "A";
      
      DataFlag = True
      
      Print #RptHandle, FAItemRec.ItemTag; Tab(22); Left$(FAItemRec.IDESC1, 28);
      Print #RptHandle, Tab(51); FAItemRec.IDEPT;
      Print #RptHandle, Tab(58); Using("###", FAItemRec.ILIFE);
      Print #RptHandle, Tab(63); Using("###,###,##0.00", CStr(FAItemRec.ORGCOST));
      If QPTrim$(FAItemRec.DEPYN) = "N" And FAItemRec.DsplFlag <> 2 Then
        Print #RptHandle, Tab(77); Using("###,###,##0.00", CStr(YTDDep#)); "*";
      Else
        Print #RptHandle, Tab(77); Using("###,###,##0.00", CStr(YTDDep#));
      End If
      Print #RptHandle, Tab(93); Using("###,###,##0.00", CStr(FAItemRec.CURRVAL))
      LineCnt& = LineCnt& + 1
      ItemCnt& = ItemCnt& + 1
      
TagOnly2: 'collects data for each department for reporting totals

      'This if statement filters out the first iteration of TAG NUMBER
      'selection because we do not want to start accumulating dept data
      'until the second iteration
      If TagFlag = False And QPTrim$(Index) = "TAG NUMBER" Then GoTo SkipEm1
      
      'collects grand totals
      If DFlag = True Then
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
      If DFlag = True Then
        DOrigCost#(1) = DOrigCost#(1) + FAItemRec.ORGCOST
        DTagDOrigCost(Nextx) = DOrigCost#(1)
        DBookTotal#(1) = DBookTotal#(1) + (FAItemRec.CURRVAL)
        DTagDBookTotal(Nextx) = DBookTotal#(1)
        DYDep#(1) = DYDep#(1) + YTDDep#
        DTagDYDep(Nextx) = DYDep#(1)
      End If
      
      If AFlag = True Then
        DOrigCost#(2) = DOrigCost#(2) + FAItemRec.ORGCOST
        ATagDOrigCost(Nextx) = DOrigCost#(2)
        DBookTotal#(2) = DBookTotal#(2) + (FAItemRec.CURRVAL)
        ATagDBookTotal(Nextx) = DBookTotal#(2)
        DYDep#(2) = DYDep#(2) + YTDDep#
        ATagDYDep(Nextx) = DYDep#(2)
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
      Print #RptHandle, "NO ADDITIONS OR DELETIONS FOR DEPARTMENT "; DeptNumber
      Print #RptHandle, String$(106, "=")
      LineCnt = LineCnt + 1
      GoTo NoData
    End If
    
  'First Print Subtotals
    Print #RptHandle,
    Print #RptHandle, "Additions for Dept Number: "; DeptNumber;
    Print #RptHandle, Tab(63); Using("###,###,##0.00", CStr(DOrigCost#(2)));
    Print #RptHandle, Tab(77); Using("###,###,##0.00", CStr(DYDep#(2)));
    Print #RptHandle, Tab(93); Using("###,###,##0.00", CStr(DBookTotal#(2)))
    
    Print #RptHandle, "Deletions for Dept Number: "; DeptNumber;
    Print #RptHandle, Tab(63); Using("###,###,##0.00", CStr(DOrigCost#(1)));
    Print #RptHandle, Tab(77); Using("###,###,##0.00", CStr(DYDep#(1)));
    Print #RptHandle, Tab(93); Using("###,###,##0.00", CStr(DBookTotal#(1)))
    
    Print #RptHandle, String$(106, "=")
    LineCnt& = LineCnt& + 4
NoData:
    frmFAShowPctComp.ShowPctComp Nextx, DIdxCnt + 1
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
    If Nextx = DIdxCnt + 1 Then Exit Do
    'move to the next dept
    Nextx = Nextx + 1
    'assign new dept to DeptNumber
    DeptNumber = Val(QPTrim$(DeptArr(Nextx)))
    'clear all dept totals
    DOrigCost#(1) = 0
    DBookTotal#(1) = 0
    DYDep#(1) = 0
    DOrigCost#(2) = 0
    DBookTotal#(2) = 0
    DYDep#(2) = 0
  
  Loop
  
  Unload frmFAShowPctComp
  frmFAShowPctComp.Out = False
  EnableCloseButton Me.hwnd, True
  Me.cmdExit.Enabled = True
  Me.cmdProcess.Enabled = True
  
  'only prints if TAG NUMBERS was selected
  If TagFlag = True Then
    GoSub PrintTagDeptTotals
  End If
  
  GoSub PrintMasterValueEnding1
  Print #RptHandle, Chr$(18);   ' oki 320 10 cpi
  
  Close         'Close all open files now
  
  Close
  ViewPrint ReportFile$, "Master Asset Listing", True
  
  KillFile (ReportFile$)
  
  Exit Sub
  
PrintMasterHeader1:
  Page = Page + 1
  Print #RptHandle, Tab(30); "Master Asset Listing : Additions and Deletions"
  Print #RptHandle, "Dept # "; Dept$
  Print #RptHandle, "Assets Acquired From "; MakeRegDate(BDate); " to "; MakeRegDate(EDate)
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
  Print #RptHandle, "* = DO NOT DEPRECIATE THIS ASSET"
  Print #RptHandle, "Asset Number"; Tab(22); "Description"; Tab(51); "Dept"; Tab(58); "Life"; Tab(64); "Original Cost"; Tab(79); "Total Deprec"; Tab(96); "Book Value"
  Print #RptHandle, String$(106, "=")
  LineCnt& = 7
  Return
  
PrintMasterValueEnding1:
  Page = Page + 1
  Print #RptHandle, FF$
  Print #RptHandle, Tab(30); "Master Asset Listing : Grand Totals"
  Print #RptHandle, "Assets Acquired From "; MakeRegDate(BDate); " to "; MakeRegDate(EDate)
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
  Print #RptHandle, Tab(35); "Total Purchase Price"; Tab(58); "Total Depreciation"; Tab(78); "Total Book Value"
  Print #RptHandle, String$(106, "=")
  Print #RptHandle, "Total Additions ";
  Print #RptHandle, Tab(41); Using("###,###,##0.00", CStr(OrigCost#(2)));
  Print #RptHandle, Tab(62); Using("###,###,##0.00", CStr(YDep#(2)));
  Print #RptHandle, Tab(80); Using("###,###,##0.00", CStr(BookTotal#(2)))
  
  Print #RptHandle, "Total Deletions ";
  Print #RptHandle, Tab(41); Using("###,###,##0.00", CStr(OrigCost#(1)));
  Print #RptHandle, Tab(62); Using("###,###,##0.00", CStr(YDep#(1)));
  Print #RptHandle, Tab(80); Using("###,###,##0.00", CStr(BookTotal#(1)))
  
  Print #RptHandle, FF$
  
  Return
  
PrintTagDeptTotals: 'print only if TAG NUMBERS was selected
  
  Page = Page + 1
  Print #RptHandle, FF$
  Print #RptHandle, Tab(30); "Master Asset Listing : Department Totals"
  Print #RptHandle, "Assets Acquired From "; MakeRegDate(BDate); " to "; MakeRegDate(EDate)
  Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
  Print #RptHandle, "Dept Number"; Tab(18); "Added/Deleted"; Tab(35); "Total Purchase Price"; Tab(58); "Total Depreciation"; Tab(78); "Total Book Value"
  Print #RptHandle, String$(106, "=")
  LineCnt = 5
  
  
  For x = 1 To DIdxCnt + 1
    If QPTrim$(DeptArr(x)) = "" Then DeptArr(x) = "0"
    Print #RptHandle, DeptArr(x); Tab(20); "Additions"; Tab(41); Using("###,###,##0.00", CStr(ATagDOrigCost(x))); Tab(62); Using("###,###,##0.00", CStr(ATagDYDep(x))); Tab(80); Using("###,###,##0.00", CStr(ATagDBookTotal(x)))
    Print #RptHandle, Tab(20); "Deletions"; Tab(41); Using("###,###,##0.00", CStr(DTagDOrigCost(x))); Tab(62); Using("###,###,##0.00", CStr(DTagDYDep(x))); Tab(80); Using("###,###,##0.00", CStr(DTagDBookTotal(x)))
    LineCnt = LineCnt + 2
    
    If LineCnt& >= MaxLines And x <> DIdxCnt + 1 Then
      LineCnt& = 0
      Page = Page + 1
      Print #RptHandle, FF$
      Print #RptHandle, Tab(20); "Master Asset Listing : Department Totals"
      Print #RptHandle, "Assets Acquired From "; MakeRegDate(BDate); " to "; MakeRegDate(EDate)
      Print #RptHandle, "Report Date: "; Date$; Tab(75); "Page #"; Page
      Print #RptHandle, "Dept Number"; Tab(18); "Added/Deleted"; Tab(35); "Total Purchase Price"; Tab(58); "Total Depreciation"; Tab(78); "Total Book Value"
      Print #RptHandle, String$(106, "=")
      LineCnt = LineCnt + 5
    End If
  Next x
  Return

End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsFATextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call Loadme
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
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
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%P"
      KeyCode = 0
    Case vbKeyF8:
      SendKeys "%D"
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
      KillFile "newadddelrptopen.dat"
      ClearInUse PWcnt
      MainLog ("FixedAssets.exe terminated via menu bar on frmFANewAddDelRpt.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub fpcmbOrder_Change()
  If QPTrim$(fpcmbOrder.Text) = "TAG NUMBER" Then
    fptxtDeptNum.Enabled = False
    fptxtDeptNum.Text = "ALL"
  ElseIf QPTrim$(fpcmbOrder.Text) = "" Then
    fpcmbOrder.Text = "TAG NUMBER"
    fptxtDeptNum.Enabled = False
    fptxtDeptNum.Text = "ALL"
  Else
    fptxtDeptNum.Enabled = True
  End If

End Sub

Private Sub fpcmbOrder_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Loadme()
  Dim One As Integer
  Dim FileHandle As Integer
  One = 1
  FileHandle = FreeFile
  Open "newadddelrptopen.dat" For Output As FileHandle Len = 2
  Print #FileHandle, One
  Close FileHandle
  fpcmbOrder.Text = "TAG NUMBER"
  fpcmbOrder.AddItem "TAG NUMBER"
  fpcmbOrder.AddItem "DEPARTMENT NUMBER"
  fptxtDeptNum.Text = "ALL"
  fpcomboPrintOpt.AddItem "Graphical"
  fpcomboPrintOpt.AddItem "Text"
  fpcomboPrintOpt.Text = "Graphical"
  fpDateStart = Date
  fpDateEnd = Date
End Sub

Private Sub fpcomboPrintOpt_Change()
  If QPTrim$(fpcomboPrintOpt.Text) = "" Then
    fpcomboPrintOpt.Text = "Graphical"
  End If
End Sub

Private Sub fptxtDeptNum_DblClick(Button As Integer)
  Dim This$
  This$ = Clipboard.GetText
  If This$ = "" Then Exit Sub
  fptxtDeptNum = Clipboard.GetText
  Clipboard.Clear

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
  Dim DFlag As Boolean
  Dim AFlag As Boolean
  Dim DeptNumber As Integer
  Dim CurDep#
  Dim MaxDep#
  Dim YTDDep#
  Dim x As Integer
  Dim Nextx As Integer
  Dim DIdxRec As DeptNumbSortIdxType
  Dim DIdxHandle As Integer
  Dim DIdxCnt As Integer
  Dim Index$, Tab1 As Integer, Tab2 As Integer, Tab3 As Integer
  Dim Page As Integer
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
  ReDim DeptArr(1 To DIdxCnt + 1) As String
  
  For x = 1 To DIdxCnt
    Get DIdxHandle, x, DIdxRec
    DeptArr(x) = QPTrim$(DIdxRec.DeptNumb)
  Next x
  DeptArr(x) = ""
  Close DIdxHandle
  
  ReDim DTagDOrigCost(1 To DIdxCnt + 1) As Double
  ReDim DTagDBookTotal(1 To DIdxCnt + 1) As Double
  ReDim DTagDYDep(1 To DIdxCnt + 1) As Double
  ReDim ATagDOrigCost(1 To DIdxCnt + 1) As Double
  ReDim ATagDBookTotal(1 To DIdxCnt + 1) As Double
  ReDim ATagDYDep(1 To DIdxCnt + 1) As Double
  ReDim TotalA(1 To DIdxCnt + 1) As Integer
  ReDim TotalD(1 To DIdxCnt + 1) As Integer
  
  If Dept$ <> "ALL" Then
    DeptNumber = Val(QPTrim$(fptxtDeptNum.Text))
  Else
    DeptNumber = Val(QPTrim(DeptArr(1)))
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
'    LineCnt = 0
  End If
  
  Do 'this loop iterates once if there is 1 dept requested, DIdxCnt + 1 if
  'department numbers "ALL" is requested and DIdxCnt + 2 for Tag Numbers
  '...Tag Numbers requires one iteration ignoring departments to get a list of
  'all valid tags in numeric order then DIdxCnt + 1 skipping the itemized tag
  'data print out just to allow the dept totals to assimilate...
    DataFlag = False
    For cnt& = 1 To NumOfFARecs
      Get FAHandle, TagIdxRecs(cnt), FAItemRec
      'Check For Disposed Date
      DisposeDate = FAItemRec.DispDate
      'Check for Acquired Date
      AcquireDate = FAItemRec.AQURDATE
      
      If DisposeDate >= BDate And DisposeDate <= EDate Or AcquireDate >= BDate And AcquireDate <= EDate Then
      'filter out items that don't fall inside the date parameters
        If DisposeDate >= BDate And DisposeDate <= EDate Then
          DFlag = True
'          GoTo DFlagIsTrue
        Else
          DFlag = False
        End If
        If AcquireDate >= BDate And AcquireDate <= EDate Then
          AFlag = True
        Else
          AFlag = False
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
      If TagFlag = True Then GoTo TagOnly2
TagOnly1: 'printing valid tag data...skipped if TAG NUMBER is chosen and this is
          'not the first iteration
      If QPTrim$(Index) = "TAG NUMBER" Then
        '                     0                  1                        2
        Print #TagHandle, Employer$; dlm; MakeRegDate(BDate); dlm; MakeRegDate(EDate); dlm;
        '                                       3
        If DFlag = True And AFlag = True Then
          Print #TagHandle, "DA"; dlm;
        ElseIf DFlag = True Then
          Print #TagHandle, "D"; dlm;
        ElseIf AFlag = True Then
          Print #TagHandle, "A"; dlm;
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
        '                                   10                           11
        Print #TagHandle, FAItemRec.CURRVAL; dlm; NoDep
      Else
        '                     0                  1                        2
        Print #RptHandle, Employer$; dlm; MakeRegDate(BDate); dlm; MakeRegDate(EDate); dlm;
        '                                       3
        If DFlag = True And AFlag = True Then
          Print #RptHandle, "DA"; dlm;
        ElseIf DFlag = True Then
          Print #RptHandle, "D"; dlm;
        ElseIf AFlag = True Then
          Print #RptHandle, "A"; dlm;
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
        If QPTrim$(FAItemRec.DEPYN) = "N" And FAItemRec.DsplFlag <> 2 Then
          NoDep = "*"
        Else
          NoDep = ""
        End If
        '                         9
        Print #RptHandle, YTDDep#; dlm;
        '                                   10
        Print #RptHandle, FAItemRec.CURRVAL; dlm;
    End If
TagOnly2: 'collects data for each department for reporting totals

      'This if statement filters out the first iteration of TAG NUMBER
      'selection because we do not want to start accumulating dept data
      'until the second iteration
      If TagFlag = False And QPTrim$(Index) = "TAG NUMBER" Then GoTo SkipEm1
      
      'collects grand totals
      If DFlag = True Then
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
      If DFlag = True Then
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
        '                   26          27
        Print #RptHandle, TACnt; dlm; TDCnt
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
    frmFAShowPctComp.ShowPctComp Nextx, DIdxCnt + 1
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
    If Nextx = DIdxCnt + 1 Then Exit Do
    'move to the next dept
    Nextx = Nextx + 1
    'assign new dept to DeptNumber
    DeptNumber = Val(QPTrim$(DeptArr(Nextx)))
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
  
  For x = 1 To DIdxCnt + 1
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

