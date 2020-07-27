VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmRetRpt 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Retirement Report"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmRetRpt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8840
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   355
      Left            =   1032
      Top             =   672
   End
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   5292
      Left            =   2220
      TabIndex        =   5
      Top             =   1794
      Width           =   7356
      _Version        =   196609
      _ExtentX        =   12975
      _ExtentY        =   9334
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.27
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
      Picture         =   "frmRetRpt.frx":08CA
      Begin LpLib.fpCombo fpcomboPrintOpt 
         Height          =   405
         Left            =   3450
         TabIndex        =   4
         Top             =   3600
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
         ColDesigner     =   "frmRetRpt.frx":08E6
      End
      Begin LpLib.fpCombo fpcmbMagDiskYN 
         Height          =   405
         Left            =   5715
         TabIndex        =   2
         Top             =   2445
         Width           =   975
         _Version        =   196608
         _ExtentX        =   1720
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
         AutoSearchFill  =   -1  'True
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
         ColDesigner     =   "frmRetRpt.frx":0D1E
      End
      Begin LpLib.fpCombo fpcmbDestination 
         Height          =   405
         Left            =   3840
         TabIndex        =   3
         Top             =   3030
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
         ColDesigner     =   "frmRetRpt.frx":1156
      End
      Begin EditLib.fpDateTime fptxtStart 
         Height          =   372
         Left            =   3600
         TabIndex        =   0
         Top             =   1344
         Width           =   1836
         _Version        =   196608
         _ExtentX        =   3238
         _ExtentY        =   649
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
      Begin EditLib.fpDateTime fptxtEnd 
         Height          =   348
         Left            =   3600
         TabIndex        =   1
         Top             =   1920
         Width           =   1836
         _Version        =   196608
         _ExtentX        =   3238
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
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   690
         Left            =   4176
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Press F10 to generate the retirement report."
         Top             =   4224
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
         ButtonDesigner  =   "frmRetRpt.frx":158E
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdEscape 
         Height          =   690
         Left            =   1320
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Press ESC to exit this screen."
         Top             =   4224
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
         ButtonDesigner  =   "frmRetRpt.frx":17A5
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Report Destination:"
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
         Left            =   1152
         TabIndex        =   11
         Top             =   3120
         Width           =   2316
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         BorderWidth     =   3
         X1              =   48
         X2              =   7296
         Y1              =   2688
         Y2              =   2688
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Create Retirement Submission Report?"
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
         Left            =   480
         TabIndex        =   10
         Top             =   2544
         Width           =   4716
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
         Left            =   1536
         TabIndex        =   9
         Top             =   3696
         Width           =   1500
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   684
         Left            =   1488
         Top             =   336
         Width           =   4476
      End
      Begin VB.Label Label3 
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
         Left            =   1920
         TabIndex        =   8
         Top             =   1968
         Width           =   1212
      End
      Begin VB.Label Label2 
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
         Height          =   252
         Left            =   1824
         TabIndex        =   7
         Top             =   1488
         Width           =   1260
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Retirement Report"
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
         TabIndex        =   6
         Top             =   480
         Width           =   4284
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Retirement Report"
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
      Height          =   615
      Left            =   975
      TabIndex        =   12
      Top             =   600
      Width           =   9705
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      Height          =   1050
      Left            =   870
      Top             =   360
      Width           =   9900
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   4
      Height          =   5616
      Left            =   2076
      Top             =   1626
      Width           =   7692
   End
End
Attribute VB_Name = "frmRetRpt"
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
   Unload frmRetRpt
End Sub

Private Sub cmdProcess_Click()
   Dim UnitHandle As Integer
   Dim UnitFileRec As UnitFileRecType
   Dim State As String
   
   OpenUnitFile UnitHandle
   Get UnitHandle, 1, UnitFileRec
   Close UnitHandle
   State = QPTrim$(UnitFileRec.UFSTATE)
   
   Timer1.Enabled = False '5/27/04
   Label5.Visible = False
   Shape3.Visible = False
   
   If fpcomboPrintOpt.Text = "Graphical" Then
     RptOpt = 1
   ElseIf fpcomboPrintOpt.Text = "Text" Then
     RptOpt = 2
   Else
     Exit Sub
   End If
   Select Case State
   Case "NC"
     If RptOpt = 2 Then
       Call NCRetirementRptT
     ElseIf RptOpt = 1 Then
       Call NCRetirementRptG
     Else
       Exit Sub
     End If
   Case "SC"
     If RptOpt = 2 Then
       Call SCRetirementRptT
     ElseIf RptOpt = 1 Then
       Call SCRetirementRptG
     Else
       Exit Sub
     End If
   Case Else
     If RptOpt = 2 Then
       Call VARetirementRptT
     ElseIf RptOpt = 1 Then
       Call VARetirementRptG
     Else
       Exit Sub
     End If
   End Select
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
  Call LoadRetRptScreen
  Me.HelpContextID = hlpRetirement
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If

End Sub
Private Sub LoadRetRptScreen()
   Dim UnitHandle As Integer
   Dim UnitFileRec As UnitFileRecType
   Dim State As String
   
   OpenUnitFile UnitHandle
   Get UnitHandle, 1, UnitFileRec
   Close UnitHandle
   State = QPTrim$(UnitFileRec.UFSTATE)
   Dim Today As String * 10
   Today = Date
   Line1.Visible = True
   Label5.Visible = False
   Shape3.Visible = False
   fpcmbMagDiskYN.Visible = True
   Label9.Visible = True
   Label4.Visible = True
   fpcmbDestination.Enabled = False
   fptxtStart.Text = Mid(Today, 1, 2) + "-" + "01" + "-" + Mid(Today, 7, 4)
   fptxtEnd.Text = Today
   If State = "NC" Then
     fpcmbMagDiskYN.AddItem "Yes"
     fpcmbMagDiskYN.AddItem "No"
     fpcmbMagDiskYN.Text = "No"
     fpcmbDestination.AddItem "Citipak Directory"
     fpcmbDestination.AddItem "Magnetic Disk"
     fpcmbDestination.Text = "Citipak Directory"
     Line1.Visible = False
   Else
     fpcmbMagDiskYN.Visible = False
     Label9.Visible = False
     Label4.Visible = False
     fpcmbDestination.Visible = False
     cmdEscape.Top = 3000
     cmdProcess.Top = 3000
     fpcomboPrintOpt.Top = 2600
     Label6.Top = 2600
   End If
   fpcomboPrintOpt.AddItem "Graphical"
   fpcomboPrintOpt.AddItem "Text"
   fpcomboPrintOpt.Text = "Graphical"
   MainLog ("Retirement Report opening screen loaded.")
End Sub
Sub SCRetirementRptG()
  
  Dim LowDate As Long, HighDate As Long
  Dim MonthNum As Integer
  Dim RptName$, Page As Integer, cnt As Integer
  Dim Dash As String * 78, UnitHandle As Integer
  Dim EmpRecSize As Long, TRecSize As Long
  Dim IdxRecLen As Integer, IdxFileSize&
  Dim NumOfRecs As Long, EmpIdxLNameHandle As Integer
  Dim UTemp$, MaxLines As Integer, PrnDef$
  Dim RptTitle$, TCol As Integer, PctRow As Integer
  Dim RHandle As Integer, THandle As Integer, DHandle As Integer
  Dim Pass As Integer, UsingThisOne As Boolean
  Dim RecNo As Long, RETAMT#, RetWage#, MatAmt#
  Dim ThisEmp&, TransRecNum&, FF$, LineCnt As Integer
  Dim GRTotal#, RTotal#, NWTotal#, x As Integer
  Dim GWTotal#, WTotal#, GETotal#, ETotal#
  Dim Emp2Rec As EmpData2Type
  Dim WLTotal#, RLTotal#, MLTotal#, WGTotal#, RGTotal#, MGTotal#
  Dim dlm$, ChangeFld$
  Dim WNTotal#, RNTotal#, MNTotal#
  Dim ThisCnt As Integer
  
  dlm$ = "~"
  FF$ = Chr$(12)
  
  LowDate = Date2Num(fptxtStart.Text) '8/26 reworked the following
  'error check
  HighDate = Date2Num(fptxtEnd.Text)
  MonthNum = Mid(fptxtEnd.Text, 1, 2)
  If LowDate > HighDate Then
    MsgBox "ERROR: The ending date is before the starting date."
    fptxtStart.SetFocus
    Exit Sub
  End If

  RptName$ = "PRRPTS\SCRETIREG.RPT"

  ReDim TransHRec(1) As TransRecType
  ReDim Unit(1) As UnitFileRecType
  ReDim Pg(1) As String * 3

  Dash = String$(78, "-")
  EmpRecSize = Len(Emp2Rec)
  TRecSize = Len(TransHRec(1))
  IdxRecLen = 2
  IdxFileSize& = FileSize(PRData + EmpIdxLName)
  NumOfRecs = IdxFileSize& \ IdxRecLen
  If NumOfRecs = 0 Then
    MsgBox "No records on file"
    Close
    Exit Sub
  End If
  
  FrmShowPctComp.Label1 = "Employee Retirement Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
  
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  OpenEmpIdxLNameFile EmpIdxLNameHandle
  For x = 1 To NumOfRecs
     Get EmpIdxLNameHandle, x, IdxBuff(x)
  Next x
  Close EmpIdxLNameHandle
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, Unit(1)
  Close UnitHandle
  
  UTemp$ = "Reporting Unit: " + QPTrim(UCase$(Unit(1).UFEMPR))

  RptTitle$ = "Employee Retirement Report."

  RHandle = FreeFile
  On Error GoTo ErrorHandler
  
  Open RptName$ For Output As RHandle
  
  THandle = FreeFile
  Open PRData + TransHistFileName For Random As THandle Len = TRecSize
  OpenEmpData2File DHandle

  Pass = 1
PassLoop:
  For RecNo = 1 To NumOfRecs
    UsingThisOne = False
    RETAMT# = 0
    RetWage# = 0
    MatAmt# = 0

    ThisEmp& = IdxBuff(RecNo)
    Get DHandle, IdxBuff(RecNo), Emp2Rec

    If Pass = 1 Then
      ChangeFld = "G"
      If Left$(Emp2Rec.EMPRETTP, 1) <> "G" Then
        GoTo SCSkipEm
      End If
    Else
      ChangeFld = "L"
      If Left$(Emp2Rec.EMPRETTP, 1) <> "L" Then
        GoTo SCSkipEm
      End If
    End If
        
    If Emp2Rec.LastTransRec <= 0 Then
      GoTo SCSkipEm
    End If
    TransRecNum& = CLng(Emp2Rec.LastTransRec)

    Do
      Get THandle, TransRecNum&, TransHRec(1) '
      Select Case TransHRec(1).CheckDate
      Case LowDate To HighDate
        If TransHRec(1).RetireAmt <> 0 Then
          RETAMT# = OldRound(RETAMT# + TransHRec(1).RetireAmt)
          RetWage# = OldRound(RetWage# + TransHRec(1).RetGrossPay)
          MatAmt# = OldRound(MatAmt# + TransHRec(1).MatchRetAmt)
          UsingThisOne = True
        End If
      Case Else
      End Select

      If TransHRec(1).PrevTransRec <= 0 Then
        If UsingThisOne Then
          If ChangeFld = "G" Then
            WLTotal = WLTotal + RetWage
            RLTotal = RLTotal + RETAMT
            MLTotal = MLTotal + MatAmt
            GoSub SCPrintEmpRetLine
          ElseIf ChangeFld = "L" Then
            WGTotal = WGTotal + RetWage
            RGTotal = RGTotal + RETAMT
            MGTotal = MGTotal + MatAmt
            GoSub SCPrintEmpRetLine
          End If
        End If
        Exit Do
      Else
        TransRecNum& = CLng(TransHRec(1).PrevTransRec)
      End If

    Loop
SCSkipEm:
    FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      Exit Sub
    End If

  Next

  If Pass = 1 Then
    GRTotal# = WLTotal# + WGTotal# + WNTotal#
    GWTotal# = RLTotal# + RGTotal# + RNTotal#
    GETotal# = MLTotal# + MGTotal# + MNTotal#

    Pass = Pass + 1
    GoTo PassLoop
  End If
  
  ThisEmp& = 0
  For RecNo = 1 To NumOfRecs
    UsingThisOne = False
    RETAMT# = 0
    RetWage# = 0
    MatAmt# = 0

    ThisEmp& = CLng(IdxBuff(RecNo))
    Get DHandle, ThisEmp&, Emp2Rec
    If Emp2Rec.LastTransRec <= 0 Then
      GoTo SCSkipNOEm
    End If
    TransRecNum& = CLng(Emp2Rec.LastTransRec)

    Do
      Get THandle, TransRecNum&, TransHRec(1)
      Select Case TransHRec(1).CheckDate
      Case LowDate To HighDate
        RETAMT# = OldRound(RETAMT# + TransHRec(1).RetireAmt)
        RetWage# = OldRound(RetWage# + TransHRec(1).RetGrossPay)
        MatAmt# = OldRound(MatAmt# + TransHRec(1).MatchRetAmt)
      Case Else
      End Select

      If TransHRec(1).PrevTransRec <= 0 Then
        If (RETAMT# = 0 And MatAmt# = 0) And RetWage# > 0 Then
          ChangeFld = "N"
          WNTotal = WNTotal + RetWage
          RNTotal = RNTotal + RETAMT
          MNTotal = MNTotal + MatAmt
          GoSub SCPrintEmpRetLine
        End If
        Exit Do
      Else
        TransRecNum& = CLng(TransHRec(1).PrevTransRec)
      End If
    Loop
SCSkipNOEm:

    FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      Exit Sub
    End If

  Next

  Close THandle
  Close DHandle   'open employee data file
  Close RHandle
  Close
  
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  
  If ThisCnt = 0 Then
    MsgBox "There are no records for the parameters entered."
    Exit Sub
  End If
  
  arRetRptSC.Show
  frmLoadingRpt.Show
  MainLog ("South Carolina retirement report processed.")

Exit Sub

SCPrintEmpRetLine:

  If ChangeFld = "G" Then
    ThisCnt = ThisCnt + 1
    '                         0
    Print #RHandle, Left$(Emp2Rec.EmpSSN, 3); "-"; Mid$(Emp2Rec.EmpSSN, 4, 2); "-"; Mid$(Emp2Rec.EmpSSN, 6, 4);
    '                                 1
    Print #RHandle, dlm; QPTrim$(Emp2Rec.EmpLName); " , "; QPTrim$(Emp2Rec.EmpFName);
    '                                 2                                   3
    Print #RHandle, dlm; Using("###,##0.00", RetWage#); dlm; Using("###,##0.00", RETAMT#); dlm;
    '                            4                            5                                6
    Print #RHandle, Using("###,##0.00", MatAmt#); dlm; "Total Government"; dlm; Using("##,###,##0.00", WLTotal); dlm;
    '                              7                                      8                       9              10
    Print #RHandle, Using("##,###,##0.00", RLTotal); dlm; Using("##,###,##0.00", MLTotal); dlm; "GENERAL "; dlm; "G";
    '                         11                 12                         13                        14
    Print #RHandle, dlm; "Soc Sec #"; dlm; "Employee Name"; dlm; "Retirement Deductions"; dlm; "Employee Match"; dlm;
    '
    Print #RHandle, fptxtStart.Text & "  to  " & fptxtEnd.Text; dlm; QPTrim$(Unit(1).UFEMPR)
  ElseIf ChangeFld = "L" Then
    ThisCnt = ThisCnt + 1
    '                         0
    Print #RHandle, Left$(Emp2Rec.EmpSSN, 3); "-"; Mid$(Emp2Rec.EmpSSN, 4, 2); "-"; Mid$(Emp2Rec.EmpSSN, 6, 4);
    '                               1
    Print #RHandle, dlm; QPTrim$(Emp2Rec.EmpLName); " , "; QPTrim$(Emp2Rec.EmpFName);
    '                                  2                                  3
    Print #RHandle, dlm; Using("###,##0.00", RetWage#); dlm; Using("###,##0.00", RETAMT#); dlm;
    '                            4                               5                                6
    Print #RHandle, Using("###,##0.00", MatAmt#); dlm; "Total Law Enforcement"; dlm; Using("##,###,##0.00", WGTotal); dlm;
    '                            7                                        8                       9                      10
    Print #RHandle, Using("##,###,##0.00", RGTotal); dlm; Using("#,###,##0.00", MGTotal); dlm; "LAW ENFORCEMENT "; dlm; "L";
    '                         11                12                        13                         14
    Print #RHandle, dlm; "Soc Sec #"; dlm; "Employee Name"; dlm; "Retirement Deductions"; dlm; "Employee Match"; dlm;
    '
    Print #RHandle, fptxtStart.Text & "  to  " & fptxtEnd.Text; dlm; QPTrim$(Unit(1).UFEMPR)
  ElseIf ChangeFld = "N" Then
    ThisCnt = ThisCnt + 1
    '                         0
    Print #RHandle, Left$(Emp2Rec.EmpSSN, 3); "-"; Mid$(Emp2Rec.EmpSSN, 4, 2); "-"; Mid$(Emp2Rec.EmpSSN, 6, 4);
    '                                 1
    Print #RHandle, dlm; QPTrim$(Emp2Rec.EmpLName); "  "; QPTrim$(Emp2Rec.EmpFName);
    '                               2                                     3
    Print #RHandle, dlm; Using("###,##0.00", RetWage#); dlm; Using("###,##0.00", RETAMT#); dlm;
    '                             4                               5                              6
    Print #RHandle, Using("###,##0.00", MatAmt#); dlm; "Total Non-Retirement"; dlm; Using("##,###,##0.00", WNTotal); dlm;
    '                7       8                   9                  10
    Print #RHandle, ""; dlm; ""; dlm; "NON-RETIREMENT WAGES"; dlm; "N";
    '                        11                   12             13       14
    Print #RHandle, dlm; "Soc Sec #"; dlm; "Employee Name"; dlm; ""; dlm; ""; dlm;
    '
    Print #RHandle, fptxtStart.Text & "  to  " & fptxtEnd.Text; dlm; QPTrim$(Unit(1).UFEMPR)
  End If

Return

ErrorHandler:
  Close
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  Unload FrmShowPctComp
  MsgBox "ERROR: If this problem persists please consult Southern Software."

End Sub

Sub VARetirementRptG()
  
  Dim LowDate As Long, HighDate As Long
  Dim MonthNum As Integer, x As Integer
  Dim RptName$, EmpRecSize As Long
  Dim Dash As String * 78, TRecSize As Long
  Dim IdxRecLen As Integer, IdxFileSize&
  Dim NumOfRecs As Long, EmpIdxLNameHandle As Integer
  Dim UnitHandle As Integer, UTemp$
  Dim RptTitle$, WTotal#, RTotal#, MTotal#
  Dim RHandle As Integer, DHandle As Integer, THandle As Integer
  Dim RecNo As Long, UsingThisOne As Boolean
  Dim RETAMT#, RetWage#, MatAmt#, TransRecNum&
  Dim Emp2Rec As EmpData2Type, cnt As Integer
  Dim dlm$
  Dim ThisCnt As Integer
  
  dlm$ = "~"
  LowDate = Date2Num(fptxtStart.Text) '8/26 reworked the
  'following error check
  HighDate = Date2Num(fptxtEnd.Text)
  MonthNum = Mid(fptxtEnd.Text, 1, 2)
  If LowDate > HighDate Then
    MsgBox "ERROR: The ending date is before the starting date."
    fptxtStart.SetFocus
    Exit Sub
  End If

  RptName$ = "PRRPTS\VARETIREG.RPT"

  ReDim TransHRec(1) As TransRecType
  ReDim Unit(1) As UnitFileRecType
  ReDim Pg(1) As String * 3

  Dash = String$(78, "-")
  EmpRecSize = Len(Emp2Rec)
  TRecSize = Len(TransHRec(1))
  IdxRecLen = 2
  IdxFileSize& = FileSize(PRData + EmpIdxLName)
  NumOfRecs = IdxFileSize& \ IdxRecLen
  
  If NumOfRecs = 0 Then
    MsgBox "No records on file"
    Close
    Exit Sub
  End If
  
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  OpenEmpIdxLNameFile EmpIdxLNameHandle
  For x = 1 To NumOfRecs
     Get EmpIdxLNameHandle, x, IdxBuff(x)
  Next x
  Close EmpIdxLNameHandle
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, Unit(1)
  Close UnitHandle
  
  RptTitle$ = "Employee Retirement Report."
  FrmShowPctComp.Label1 = "Employee Retirement Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False

  RHandle = FreeFile
  On Error GoTo ErrorHandler
  Open RptName$ For Output As RHandle
  THandle = FreeFile
  Open PRData + TransHistFileName For Random As THandle Len = TRecSize
  OpenEmpData2File DHandle

  For RecNo = 1 To NumOfRecs
    UsingThisOne = False
    RETAMT# = 0
    RetWage# = 0
    MatAmt# = 0
    Get DHandle, CLng(IdxBuff(RecNo)), Emp2Rec

    If Emp2Rec.LastTransRec <= 0 Then
      GoTo oSkipEm6
    End If

    TransRecNum& = CLng(Emp2Rec.LastTransRec)

    Do
      Get THandle, TransRecNum&, TransHRec(1)
      Select Case TransHRec(1).CheckDate
      Case LowDate To HighDate
      'added MatchRetAmt for places that contribute 100% of retirement
      'to employees...employees contribute nothing ...6/27/02
        If TransHRec(1).RetireAmt <> 0 Then
          RETAMT# = OldRound(RETAMT# + TransHRec(1).RetireAmt)
          RetWage# = OldRound(RetWage# + TransHRec(1).RetGrossPay)
          MatAmt# = OldRound(MatAmt# + TransHRec(1).MatchRetAmt)
          UsingThisOne = True
        End If
      Case Else
      End Select

      If TransHRec(1).PrevTransRec <= 0 Then
        If UsingThisOne Then
          WTotal = WTotal + RetWage
          RTotal = RTotal + RETAMT
          MTotal = MTotal + MatAmt
          GoSub oPrintEmpRetLine
        End If
        Exit Do
      Else
        TransRecNum& = CLng(TransHRec(1).PrevTransRec)
      End If
    Loop
oSkipEm6:
    FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next

  Close THandle
  Close DHandle   'open employee data file
  Close RHandle
  Close
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  
  If ThisCnt = 0 Then
    MsgBox "There are no records for the parameters entered."
    Exit Sub
  End If
  
  arRetRptVA.Show
  frmLoadingRpt.Show
  
  MainLog ("Virginia retirement report processed.")
  
  Exit Sub


oPrintEmpRetLine:
  ThisCnt = ThisCnt + 1
  '                              0
  Print #RHandle, Left$(Emp2Rec.EmpSSN, 3); "-"; Mid$(Emp2Rec.EmpSSN, 4, 2); "-"; Mid$(Emp2Rec.EmpSSN, 6, 4);
  '                              1
  Print #RHandle, dlm; QPTrim$(Emp2Rec.EmpLName); "  "; QPTrim$(Emp2Rec.EmpFName);
  '                              2                                      3
  Print #RHandle, dlm; Using("###,##0.00", RetWage#); dlm; Using("###,##0.00", RETAMT#); dlm;
  '                              4                               5
  Print #RHandle, Using("###,##0.00", MatAmt#); dlm; Using("##,###,##0.00", WTotal); dlm;
  '                              6                                      7                    8
  Print #RHandle, Using("##,###,##0.00", RTotal); dlm; Using("##,###,##0.00", MTotal); dlm; "VA"; dlm;
  '
  Print #RHandle, fptxtStart.Text & "  to  " & fptxtEnd.Text; dlm; QPTrim$(Unit(1).UFEMPR)

Return

ErrorHandler:
  Close
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  Unload FrmShowPctComp
  MsgBox "ERROR: If this problem persists please consult Southern Software."

End Sub

Sub NCRetirementRptG()
  
  Dim LowDate As Long, HighDate As Long
  Dim MonthNum As Integer, YearNum As Integer
  Dim cnt As Integer
  Dim RptName$, EmpRecSize As Long
  Dim TRecSize As Long, IdxRecLen As Integer
  Dim IdxFileSize&, NumOfRecs As Long
  Dim EmpIdxLNameHandle As Integer
  Dim UnitHandle As Integer, MaxLines As Integer
  Dim RptTitle$, TCol As Integer, PctRow As Integer
  Dim RHandle As Integer, DHandle As Integer
  Dim THandle As Integer, RecNo As Long
  Dim UsingThisOne As Boolean, RetWage#, RETAMT#
  Dim MatAmt#, MatchCnt As Integer, TransRecNum&
  Dim PGTotal#, PRTotal#, PMTotal#, FF$
  Dim GTotal#, RTotal#, MTotal#, RGTotal#, MGTotal#
  Dim WTotal#, WGTotal#
  Dim LineCnt As Integer, x As Integer
  Dim LPage As Integer, GPage As Integer
  Dim Emp2Rec As EmpData2Type
  Dim dlm As String
  Dim Gov$
  Dim Law$
  Dim LTotal As Double
  Dim Total As Double
  Dim GHHeader$
  Dim ghTitle$
  Dim SubTitle$
  Dim UnitCode$
  Dim ThisFile$
  Dim ThisCnt As Integer
  
  ThisFile = "NONE"
  dlm = "~"
  FF$ = Chr$(12)
  
  LowDate = Date2Num(fptxtStart.Text) '8/26 reworked the
  'following error check
  HighDate = Date2Num(fptxtEnd.Text)
  MonthNum = Mid(fptxtEnd.Text, 1, 2)
  YearNum = Mid(fptxtEnd.Text, 7, 4)
  If LowDate > HighDate Then
    MsgBox "ERROR: The ending date is before the starting date."
    fptxtStart.SetFocus
    Exit Sub
  End If
  If fpcmbMagDiskYN.Text = "Yes" Then Call Ret2Disk(ThisFile)
  DoEvents
  RptName$ = "PRRPTS\RETIREG.RPT"
  
  ReDim TransHRec(1) As TransRecType
  ReDim Unit(1) As UnitFileRecType
  ReDim Pg(1) As String * 3
  Dim Dash As String * 91
  
  Dash = String$(91, "-")
  
  EmpRecSize = Len(Emp2Rec)
  TRecSize = Len(TransHRec(1))
  IdxRecLen = 2
  
  IdxFileSize& = FileSize(PRData + EmpIdxLName)
  NumOfRecs = IdxFileSize& \ IdxRecLen
  If NumOfRecs = 0 Then
    MsgBox "No records on file"
    Close
    Exit Sub
  End If
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  OpenEmpIdxLNameFile EmpIdxLNameHandle
  For x = 1 To NumOfRecs
     Get EmpIdxLNameHandle, x, IdxBuff(x)
  Next x
  Close EmpIdxLNameHandle
  
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, Unit(1)
  Close UnitHandle
  
  MaxLines = 50
  RptTitle$ = "Employee Retirement Report."
  FrmShowPctComp.Label1 = "Employee Retirement Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
  
  TCol = 40 - (Len(RptTitle$) \ 2) + 1
  PctRow = 11
  
  RHandle = FreeFile
  On Error GoTo ErrorHandler
  Open RptName$ For Output As RHandle
  THandle = FreeFile
  Open PRData + TransHistFileName For Random As THandle Len = TRecSize
  OpenEmpData2File DHandle
  
  WTotal# = 0
  RTotal# = 0
  MTotal# = 0
  For RecNo = 1 To NumOfRecs
    UsingThisOne = False
    RetWage# = 0
    RETAMT# = 0
    MatAmt# = 0
    MatchCnt = 0
    Get DHandle, IdxBuff(RecNo), Emp2Rec
    If Emp2Rec.LastTransRec <= 0 Then
      GoTo SkipEm6
    End If
    TransRecNum& = CLng(Emp2Rec.LastTransRec)
    Do
      Get THandle, TransRecNum&, TransHRec(1)
      Select Case TransHRec(1).CheckDate
      Case LowDate To HighDate
'        If UCase$(Left$(LTrim$(Emp2Rec.EMPRETTP), 1)) = "L" Then
        If Mid(Emp2Rec.EMPRETNO, 1, 1) = "R" And TransHRec(1).MatchRetAmt = 0 Then GoTo Retired
'        If UCase$(Left$(LTrim$(Emp2Rec.EMPRETTP), 1)) = "L" Then
        If Mid(Emp2Rec.EMPRETNO, 1, 1) = "R" Then
          Emp2Rec.EMPRETNO = Mid(Emp2Rec.EMPRETNO, 2, Len(Emp2Rec.EMPRETNO))
        End If
        If UCase$(Left$(LTrim$(Emp2Rec.EMPRETTP), 1)) = "L" And Mid(Emp2Rec.EMPRETNO, 1, 1) <> "T" Then '7/22/2004
          UnitCode = QPTrim$(Unit(1).UFRETIDL)
          GHHeader = "Law"
          ghTitle$ = "North Carolina Law Enforcement Officers Benefit and Retirement Fund"
          SubTitle = "Total Law Enforcement"
          RetWage# = OldRound(RetWage# + TransHRec(1).RetGrossPay)
          RETAMT# = OldRound(RETAMT# + TransHRec(1).RetireAmt)
          MatAmt# = OldRound(MatAmt# + TransHRec(1).MatchRetAmt)
          MatchCnt = MatchCnt + 1
          UsingThisOne = True
        End If
      Case Else
      End Select
Retired:
      If TransHRec(1).PrevTransRec <= 0 Then
        If UsingThisOne Then
          WTotal = WTotal + RetWage
          RTotal = RTotal + RETAMT
          MTotal = MTotal + MatAmt
          GoSub PrintEmpRetLine
        End If
        Exit Do
      Else
        TransRecNum& = CLng(TransHRec(1).PrevTransRec)
      End If
    Loop
'the progress bar is split into two parts...the first 50%
'is here and the second is under SkipEm7

SkipEm6:
  If RecNo / NumOfRecs < 0.5 Then
    FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      EnableCloseButton Me.hwnd, True
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      Unload FrmShowPctComp
      Exit Sub
    End If
  End If

  Next
  
  RptTitle$ = "Employee Retirement Report."
  FrmShowPctComp.Label1 = "Employee Retirement Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
  
  WTotal = 0
  RTotal = 0
  MTotal = 0
  
  For RecNo = 1 To NumOfRecs
    UsingThisOne = False
    RetWage# = 0
    RETAMT# = 0
    MatAmt# = 0
    Get DHandle, CLng(IdxBuff(RecNo)), Emp2Rec
    If Emp2Rec.LastTransRec <= 0 Then
      GoTo SkipEm7
    End If
    
    TransRecNum& = CLng(Emp2Rec.LastTransRec)
    Do
      Get THandle, TransRecNum&, TransHRec(1)
      Select Case TransHRec(1).CheckDate
      Case LowDate To HighDate
        If Mid(Emp2Rec.EMPRETNO, 1, 1) = "R" And TransHRec(1).MatchRetAmt = 0 Then GoTo Retired2
        If Mid(Emp2Rec.EMPRETNO, 1, 1) = "R" Then
          Emp2Rec.EMPRETNO = Mid(Emp2Rec.EMPRETNO, 2, Len(Emp2Rec.EMPRETNO))
        End If
'      If UCase$(Left$(LTrim$(Emp2Rec.EMPRETTP), 1)) = "G" Then
        If UCase$(Left$(LTrim$(Emp2Rec.EMPRETTP), 1)) = "G" And Mid(Emp2Rec.EMPRETNO, 1, 1) <> "T" Then '7/22/2004
          UnitCode = QPTrim$(Unit(1).UFRETID)
          SubTitle = "Total Government"
          GHHeader = "Gov"
          ghTitle$ = "North Carolina Local Government Retirement System"
          RetWage# = OldRound(RetWage# + TransHRec(1).RetGrossPay) 'was GrossPay
          RETAMT# = OldRound(RETAMT# + TransHRec(1).RetireAmt)
          MatAmt# = OldRound(MatAmt# + TransHRec(1).MatchRetAmt)
          UsingThisOne = True
        End If
      Case Else
      End Select
Retired2:
      If TransHRec(1).PrevTransRec <= 0 Then
        If UsingThisOne Then
          WTotal = WTotal + RetWage
          RTotal = RTotal + RETAMT
          MTotal = MTotal + MatAmt
          GoSub PrintEmpRetLine
        End If
        Exit Do
      Else
        TransRecNum& = CLng(TransHRec(1).PrevTransRec)
      End If
    Loop
    
SkipEm7:
    If RecNo / NumOfRecs > 0.5 Then
      FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        Me.cmdEscape.Enabled = True
        Me.cmdProcess.Enabled = True
        EnableCloseButton Me.hwnd, True
        Unload FrmShowPctComp
        Exit Sub
      End If
    End If

  Next
  Close RHandle
  Close THandle
  Close DHandle                'open employee data file
  Close
  
  If ThisCnt = 0 Then
    MsgBox "There are no records for the parameters entered."
    Me.cmdEscape.Enabled = True
    Me.cmdProcess.Enabled = True
    EnableCloseButton Me.hwnd, True
    Exit Sub
  End If
    
  arRetRpt.Show
  frmLoadingRpt.Show
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  MainLog ("North Carolina retirement report processed.")
  
  If ThisFile <> "NONE" Then
   Label5.Visible = True
   Shape3.Visible = True
   Label5.Caption = ThisFile + " saved successfully."
   Call Timer1_Timer 'added 5/27/04
  End If
  
  Exit Sub
  
PrintEmpRetLine:
    ThisCnt = ThisCnt + 1
    '                  0
    Print #RHandle, GHHeader;
    '                                 1                                        2
    Print #RHandle, dlm; QPTrim$(Left$(Emp2Rec.EMPRETNO, 12)); dlm; QPTrim$(Emp2Rec.EmpLName); "  "; QPTrim$(Emp2Rec.EmpFName);
    '                                 3                                    4
    Print #RHandle, dlm; Using("###,##0.00", RetWage#); dlm; Using("###,##0.00", RETAMT#); dlm;
    '                            5                        6                          7
    Print #RHandle, Using("###,##0.00", MatAmt#); dlm; SubTitle; dlm; Using("##,###,##0.00", WTotal); dlm;
    '                             8                                  9                                      10                  11                                         12
    Print #RHandle, Using("##,###,##0.00", RTotal); dlm; Using("##,###,##0.00", MTotal); dlm; "Unit Code: " & UnitCode; dlm; ghTitle; dlm; Left$(Emp2Rec.EmpSSN, 3); "-"; Mid$(Emp2Rec.EmpSSN, 4, 2); "-"; Mid$(Emp2Rec.EmpSSN, 6, 4);
    '                         13             14                 15                       16                       17
    Print #RHandle, dlm; "Soc Sec #"; dlm; "Ret #"; dlm; "Employee Name"; dlm; "Retirement Deduction"; dlm; "Employer Match"; dlm;
    '
    Print #RHandle, MonthName$(MonthNum) & " " & YearNum; dlm; QPTrim$(Unit(1).UFEMPR)
Return
  
ErrorHandler:
  Close
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  Unload FrmShowPctComp
  MsgBox "ERROR: If this problem persists please consult Southern Software."
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdEscape.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmRetRpt.")
      Call Terminate
      End
    End If
  End If
End Sub

Sub SCRetirementRptT()
  
  Dim LowDate As Long, HighDate As Long
  Dim MonthNum As Integer
  Dim RptName$, Page As Integer, cnt As Integer
  Dim Dash As String * 78, UnitHandle As Integer
  Dim EmpRecSize As Long, TRecSize As Long
  Dim IdxRecLen As Integer, IdxFileSize&
  Dim NumOfRecs As Long, EmpIdxLNameHandle As Integer
  Dim UTemp$, MaxLines As Integer, PrnDef$
  Dim RptTitle$, TCol As Integer, PctRow As Integer
  Dim RHandle As Integer, THandle As Integer, DHandle As Integer
  Dim Pass As Integer, UsingThisOne As Boolean
  Dim RecNo As Long, RETAMT#, RetWage#, MatAmt#
  Dim ThisEmp&, TransRecNum&, FF$, LineCnt As Integer
  Dim GRTotal#, RTotal#, NWTotal#, x As Integer
  Dim GWTotal#, WTotal#, GETotal#, ETotal#
  Dim Emp2Rec As EmpData2Type
  Dim ThisCnt As Integer
  
  FF$ = Chr$(12)
  
  LowDate = Date2Num(fptxtStart.Text) '8/26 reworked the following
  'error check
  HighDate = Date2Num(fptxtEnd.Text)
  MonthNum = Mid(fptxtEnd.Text, 1, 2)
  If LowDate > HighDate Then
    MsgBox "ERROR: The ending date is before the starting date."
    fptxtStart.SetFocus
    Exit Sub
  End If

  RptName$ = "PRRPTS\RETIRE.RPT"

  ReDim TransHRec(1) As TransRecType
  ReDim Unit(1) As UnitFileRecType
  ReDim Pg(1) As String * 3

  Dash = String$(78, "-")
  EmpRecSize = Len(Emp2Rec)
  TRecSize = Len(TransHRec(1))
  IdxRecLen = 2
  IdxFileSize& = FileSize(PRData + EmpIdxLName)
  NumOfRecs = IdxFileSize& \ IdxRecLen
  If NumOfRecs = 0 Then
    MsgBox "No records on file"
    Close
    Exit Sub
  End If
  
  FrmShowPctComp.Label1 = "Employee Retirement Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
  
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  OpenEmpIdxLNameFile EmpIdxLNameHandle
  For x = 1 To NumOfRecs
     Get EmpIdxLNameHandle, x, IdxBuff(x)
  Next x
  Close EmpIdxLNameHandle
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, Unit(1)
  Close UnitHandle
  
  UTemp$ = "Reporting Unit: " + QPTrim(UCase$(Unit(1).UFEMPR))

  MaxLines = 55
  RptTitle$ = "Employee Retirement Report."

  TCol = 40 - (Len(RptTitle$) \ 2) + 1
  PctRow = 11

  RHandle = FreeFile
  Open RptName$ For Output As RHandle
  RPTSetupPRN 11, RHandle
  
  THandle = FreeFile
  Open PRData + TransHistFileName For Random As THandle Len = TRecSize
  OpenEmpData2File DHandle

  Pass = 1
PassLoop:
  GoSub SCRetRptHeader
  For RecNo = 1 To NumOfRecs
    UsingThisOne = False
    RETAMT# = 0
    RetWage# = 0
    MatAmt# = 0

    ThisEmp& = IdxBuff(RecNo)
    Get DHandle, IdxBuff(RecNo), Emp2Rec

    If Pass = 1 Then
      If Left$(Emp2Rec.EMPRETTP, 1) <> "G" Then
        GoTo SCSkipEm
      End If
    Else
      If Left$(Emp2Rec.EMPRETTP, 1) <> "L" Then
        GoTo SCSkipEm
      End If
    End If
        
    If Emp2Rec.LastTransRec <= 0 Then
      GoTo SCSkipEm
    End If
    TransRecNum& = CLng(Emp2Rec.LastTransRec)

    Do
      Get THandle, TransRecNum&, TransHRec(1)
      Select Case TransHRec(1).CheckDate
      Case LowDate To HighDate
        If TransHRec(1).RetireAmt <> 0 Then
          RETAMT# = OldRound(RETAMT# + TransHRec(1).RetireAmt)
          RetWage# = OldRound(RetWage# + TransHRec(1).RetGrossPay)
          MatAmt# = OldRound(MatAmt# + TransHRec(1).MatchRetAmt)
          UsingThisOne = True
        End If
      Case Else
      End Select

      If TransHRec(1).PrevTransRec <= 0 Then
        If UsingThisOne Then
          GoSub SCPrintEmpRetLine
          If LineCnt >= MaxLines Then
            Print #RHandle, FF$
            GoSub SCRetRptHeader
          End If
        End If
        Exit Do
      Else
        TransRecNum& = CLng(TransHRec(1).PrevTransRec)
      End If

    Loop
SCSkipEm:
    FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      Exit Sub
    End If

  Next

  If Pass = 1 Then
    GoSub SCRetTotals
    GRTotal# = RTotal#
    GWTotal# = WTotal#
    GETotal# = ETotal#

    RTotal# = 0
    WTotal# = 0
    ETotal# = 0
    
    Pass = Pass + 1
    GoTo PassLoop
  End If

  GoSub SCRetGTotals
'***********************************************************
  GoSub SCNORetRptHeader

  ThisEmp& = 0
  For RecNo = 1 To NumOfRecs
    UsingThisOne = False
    RETAMT# = 0
    RetWage# = 0
    MatAmt# = 0

    ThisEmp& = CLng(IdxBuff(RecNo))
    Get DHandle, ThisEmp&, Emp2Rec
    If Emp2Rec.LastTransRec <= 0 Then
      GoTo SCSkipNOEm
    End If
    TransRecNum& = CLng(Emp2Rec.LastTransRec)

    Do
      Get THandle, TransRecNum&, TransHRec(1)
      Select Case TransHRec(1).CheckDate
      Case LowDate To HighDate
        RETAMT# = OldRound(RETAMT# + TransHRec(1).RetireAmt)
        RetWage# = OldRound(RetWage# + TransHRec(1).RetGrossPay)
        MatAmt# = OldRound(MatAmt# + TransHRec(1).MatchRetAmt)
      Case Else
      End Select

      If TransHRec(1).PrevTransRec <= 0 Then
        If (RETAMT# = 0 And MatAmt# = 0) And RetWage# > 0 Then
          GoSub SCPrintEmpNORetLine
          If LineCnt >= MaxLines Then
            Print #RHandle, FF$
            GoSub SCNORetRptHeader
          End If
        End If
        Exit Do
      Else
        TransRecNum& = CLng(TransHRec(1).PrevTransRec)
      End If
    Loop
SCSkipNOEm:

    FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      Exit Sub
    End If

  Next

  GoSub SCNORetTotals

'****************************************************************

  Close THandle
  Close DHandle   'open employee data file
  RPTSetupPRN 123, RHandle '8/26...123 is the default end code
  Close RHandle
  Close
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  
  If ThisCnt = 0 Then
    MsgBox "There are no records for the parameters entered."
    Exit Sub
  End If
  
  ViewPrint RptName$, RptTitle$, True
  EnableCloseButton Me.hwnd, True
  MainLog ("South Carolina retirement report processed.")

Exit Sub

SCPrintEmpRetLine:
  ThisCnt = ThisCnt + 1
  Print #RHandle, Left$(Emp2Rec.EmpSSN, 3); "-"; Mid$(Emp2Rec.EmpSSN, 4, 2); "-"; Mid$(Emp2Rec.EmpSSN, 6, 4);
  Print #RHandle, Tab(16); QPTrim$(Emp2Rec.EmpLName); ", "; QPTrim$(Emp2Rec.EmpFName);
  Print #RHandle, Tab(41); Using("###,##0.00", RetWage#);
  Print #RHandle, Tab(55); Using("###,##0.00", RETAMT#);
  Print #RHandle, Tab(69); Using("###,##0.00", MatAmt#)
  LineCnt = LineCnt + 1     'employeesprinted = employeesprinted + 1
  RTotal# = OldRound(RTotal# + RETAMT#)
  WTotal# = OldRound(WTotal# + RetWage#)
  ETotal# = OldRound(ETotal# + MatAmt#)

Return

SCPrintEmpNORetLine:
  ThisCnt = ThisCnt + 1
  Print #RHandle, Left$(Emp2Rec.EmpSSN, 3); "-"; Mid$(Emp2Rec.EmpSSN, 4, 2); "-"; Mid$(Emp2Rec.EmpSSN, 6, 4);
  Print #RHandle, Tab(16); QPTrim$(Emp2Rec.EmpLName); ", "; QPTrim$(Emp2Rec.EmpFName);
  Print #RHandle, Tab(47); Using("###,##0.00", RetWage#)
  LineCnt = LineCnt + 1     'employeesprinted = employeesprinted + 1
  NWTotal# = OldRound(NWTotal# + RetWage#)

Return

SCRetTotals:
  Print #RHandle, Dash
  Print #RHandle, Tab(28); "Totals:";
  Print #RHandle, Tab(38); Using("#,###,##0.00", WTotal#);
  Print #RHandle, Tab(52); Using("#,###,##0.00", RTotal#);
  Print #RHandle, Tab(67); Using("###,##0.00", ETotal#)
  Print #RHandle, FF$
Return

SCRetGTotals:
  Print #RHandle, Dash
  Print #RHandle, Tab(28); "Totals:";
  Print #RHandle, Tab(38); Using("#,###,##0.00", WTotal#);
  Print #RHandle, Tab(52); Using("#,###,##0.00", RTotal#);
  Print #RHandle, Tab(67); Using("###,##0.00", ETotal#)
  Print #RHandle, Tab(22); "Grand Totals:";
  Print #RHandle, Tab(38); Using("#,###,##0.00", OldRound#(GWTotal# + WTotal#));
  Print #RHandle, Tab(52); Using("#,###,##0.00", OldRound#(GRTotal# + RTotal#));
  Print #RHandle, Tab(67); Using("###,##0.00", OldRound#(GETotal# + ETotal#))
  Print #RHandle, FF$
Return

SCNORetTotals:
  Print #RHandle, Dash
  Print #RHandle, Tab(28); "Totals:";
  Print #RHandle, Tab(45); Using("#,###,##0.00", NWTotal#)
  Print #RHandle, FF$
Return

SCRetRptHeader:
  Page = Page + 1
  Print #RHandle, "S.C. Retirement System Report"; Tab(68); "Page:"; Page
  Print #RHandle, "Monthly Report of Subject Wages and Retirement Contributions."
  Print #RHandle, "Report Date:  "; MakeRegDate(LowDate); " to "; MakeRegDate(HighDate)
  Print #RHandle, UTemp$
  If Pass = 1 Then
    Print #RHandle, "General";
  Else
    Print #RHandle, "Law Enforcement";
  End If
  Print #RHandle, Tab(38); "Wages Subject    Retirement    Employer"
  Print #RHandle, "Soc Sec #      Employee Name         to Retirement    Deductions    Matching"
  Print #RHandle, Dash
  LineCnt = 7
Return

SCNORetRptHeader:
  Page = Page + 1
  Print #RHandle, "S.C. Retirement System Report"; Tab(68); "Page:"; Page
  Print #RHandle, "Monthly Report of Subject Wages and Retirement Contributions."
  Print #RHandle, "Report Date:  "; MakeRegDate(LowDate); " to "; MakeRegDate(HighDate)
  Print #RHandle, UTemp$
  Print #RHandle,
  Print #RHandle, "Soc Sec #      Employee Name        NON-Retirement Wages"
  Print #RHandle, Dash
  LineCnt = 6
Return

End Sub

Sub VARetirementRptT()
  
  Dim LowDate As Long, HighDate As Long, FF$
  Dim MonthNum As Integer, x As Integer
  Dim RptName$, EmpRecSize As Long
  Dim Dash As String * 78, TRecSize As Long
  Dim IdxRecLen As Integer, IdxFileSize&
  Dim NumOfRecs As Long, EmpIdxLNameHandle As Integer
  Dim UnitHandle As Integer, UTemp$, MaxLines As Integer
  Dim RptTitle$, TCol As Integer, PctRow As Integer
  Dim RHandle As Integer, DHandle As Integer, THandle As Integer
  Dim RecNo As Long, UsingThisOne As Boolean
  Dim RETAMT#, RetWage#, MatAmt#, TransRecNum&
  Dim LineCnt As Integer, PrnDef$, RTotal#, WTotal#, ETotal#
  Dim Emp2Rec As EmpData2Type, cnt As Integer
  Dim ThisCnt As Integer
  
  FF$ = Chr$(12)

  LowDate = Date2Num(fptxtStart.Text) '8/26 reworked the
  'following error check
  HighDate = Date2Num(fptxtEnd.Text)
  MonthNum = Mid(fptxtEnd.Text, 1, 2)
  If LowDate > HighDate Then
    MsgBox "ERROR: The ending date is before the starting date."
    fptxtStart.SetFocus
    Exit Sub
  End If

  RptName$ = "PRRPTS\RETIRE.RPT"

  ReDim TransHRec(1) As TransRecType
  ReDim Unit(1) As UnitFileRecType
  ReDim Pg(1) As String * 3

  Dash = String$(78, "-")
  EmpRecSize = Len(Emp2Rec)
  TRecSize = Len(TransHRec(1))
  IdxRecLen = 2
  IdxFileSize& = FileSize(PRData + EmpIdxLName)
  NumOfRecs = IdxFileSize& \ IdxRecLen
  
  If NumOfRecs = 0 Then
    MsgBox "No records on file"
    Close
    Exit Sub
  End If
  
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  OpenEmpIdxLNameFile EmpIdxLNameHandle
  For x = 1 To NumOfRecs
     Get EmpIdxLNameHandle, x, IdxBuff(x)
  Next x
  Close EmpIdxLNameHandle
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, Unit(1)
  Close UnitHandle
  
  UTemp$ = "Reporting Unit: " + QPTrim(UCase$(Unit(1).UFEMPR))

  MaxLines = 55

  RptTitle$ = "Employee Retirement Report."
  FrmShowPctComp.Label1 = "Employee Retirement Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False

  TCol = 40 - (Len(RptTitle$) \ 2) + 1
  PctRow = 11

  RHandle = FreeFile
  Open RptName$ For Output As RHandle
  RPTSetupPRN 11, RHandle
  THandle = FreeFile
  Open PRData + TransHistFileName For Random As THandle Len = TRecSize
  OpenEmpData2File DHandle

  GoSub oLRetRptHeader
  For RecNo = 1 To NumOfRecs
    UsingThisOne = False
    RETAMT# = 0
    RetWage# = 0
    MatAmt# = 0
    Get DHandle, CLng(IdxBuff(RecNo)), Emp2Rec

    If Emp2Rec.LastTransRec <= 0 Then
      GoTo oSkipEm6
    End If

    TransRecNum& = CLng(Emp2Rec.LastTransRec)

    Do
      Get THandle, TransRecNum&, TransHRec(1)
      Select Case TransHRec(1).CheckDate
      Case LowDate To HighDate
      'added MatchRetAmt for places that contribute 100% of retirement
      'to employees...employees contribute nothing ...6/27/02
        If TransHRec(1).RetireAmt <> 0 Then
          RETAMT# = OldRound(RETAMT# + TransHRec(1).RetireAmt)
          RetWage# = OldRound(RetWage# + TransHRec(1).RetGrossPay)
          MatAmt# = OldRound(MatAmt# + TransHRec(1).MatchRetAmt)
          UsingThisOne = True
        End If
      Case Else
      End Select

      If TransHRec(1).PrevTransRec <= 0 Then
        If UsingThisOne Then
          GoSub oPrintEmpRetLine
          If LineCnt >= MaxLines Then
            Print #RHandle, FF$
            GoSub oLRetRptHeader
          End If
        End If
        Exit Do
      Else
        TransRecNum& = CLng(TransHRec(1).PrevTransRec)
      End If
    Loop
oSkipEm6:
    FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Me.cmdEscape.Enabled = True
      Me.cmdProcess.Enabled = True
      EnableCloseButton Me.hwnd, True
      Unload FrmShowPctComp
      Exit Sub
    End If
  Next
  GoSub oRetLTotals

  Close THandle
  Close DHandle   'open employee data file
  RPTSetupPRN 123, RHandle '8/26...123 is the default end code
  Close RHandle
  Close
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  
  If ThisCnt = 0 Then
    MsgBox "There are no records for the parameters entered."
    Exit Sub
  End If
  
  ViewPrint RptName$, RptTitle$, True
  MainLog ("Virginia retirment report processed.")
  
Exit Sub

oPrintEmpRetLine:
  ThisCnt = ThisCnt + 1
  Print #RHandle, Left$(Emp2Rec.EmpSSN, 3); "-"; Mid$(Emp2Rec.EmpSSN, 4, 2); "-"; Mid$(Emp2Rec.EmpSSN, 6, 4);
  Print #RHandle, Tab(16); QPTrim$(Emp2Rec.EmpLName); ", "; QPTrim$(Emp2Rec.EmpFName);
  Print #RHandle, Tab(41); Using("###,##0.00", RetWage#);
  Print #RHandle, Tab(55); Using("###,##0.00", RETAMT#);
  Print #RHandle, Tab(67); Using("###,##0.00", MatAmt#)
  LineCnt = LineCnt + 1     'employeesprinted = employeesprinted + 1
  RTotal# = OldRound(RTotal# + RETAMT#)
  WTotal# = OldRound(WTotal# + RetWage#)
  ETotal# = OldRound(ETotal# + MatAmt#)

Return

oRetLTotals:
  Print #RHandle, Dash
  Print #RHandle, Tab(28); "Totals:";
  Print #RHandle, Tab(39); Using("#,###,##0.00", WTotal#);
  Print #RHandle, Tab(53); Using("#,###,##0.00", RTotal#);
  Print #RHandle, Tab(67); Using("###,##0.00", ETotal#)
  Print #RHandle, FF$
Return

oLRetRptHeader:
  Print #RHandle, "Retirement Deduction Report"
  Print #RHandle, ""
  Print #RHandle, "Report Date:  "; MakeRegDate(LowDate); " to "; MakeRegDate(HighDate)
  Print #RHandle,
  Print #RHandle, "                                     Wages Subject    Retirement    Employer"
  Print #RHandle, "Soc Sec #      Employee Name         to Retirement    Deductions    Matching"
  Print #RHandle, Dash
  LineCnt = 7
Return

End Sub

Sub NCRetirementRptT()
  
  Dim LowDate As Long, HighDate As Long
  Dim MonthNum As Integer, YearNum As Integer
  Dim cnt As Integer
  Dim RptName$, EmpRecSize As Long
  Dim TRecSize As Long, IdxRecLen As Integer
  Dim IdxFileSize&, NumOfRecs As Long
  Dim EmpIdxLNameHandle As Integer
  Dim UnitHandle As Integer, MaxLines As Integer
  Dim RptTitle$, TCol As Integer, PctRow As Integer
  Dim RHandle As Integer, DHandle As Integer
  Dim THandle As Integer, RecNo As Long
  Dim UsingThisOne As Boolean, RetWage#, RETAMT#
  Dim MatAmt#, MatchCnt As Integer, TransRecNum&
  Dim PGTotal#, PRTotal#, PMTotal#, FF$
  Dim GTotal#, RTotal#, MTotal#
  Dim LineCnt As Integer, x As Integer
  Dim LPage As Integer, GPage As Integer
  Dim Emp2Rec As EmpData2Type
  Dim PageGTot#, PageRTot#, PageMTot#
  Dim PageFF As Boolean
  Dim ThisFile$
  Dim ThisCnt As Integer
  
  ThisFile = "NONE"
  FF$ = Chr$(12)
  PageFF = False
  LowDate = Date2Num(fptxtStart.Text) '8/26 reworked the
  'following error check
  HighDate = Date2Num(fptxtEnd.Text)
  MonthNum = Mid(fptxtEnd.Text, 1, 2)
  YearNum = Mid(fptxtEnd.Text, 7, 4)
  If LowDate > HighDate Then
    MsgBox "ERROR: The ending date is before the starting date."
    fptxtStart.SetFocus
    Exit Sub
  End If
  
  If fpcmbMagDiskYN.Text = "Yes" Then Call Ret2Disk(ThisFile)
  DoEvents
    
  RptName$ = "PRRPTS\RETIRE.RPT"
  
  ReDim TransHRec(1) As TransRecType
  ReDim Unit(1) As UnitFileRecType
  ReDim Pg(1) As String * 3
  Dim Dash As String * 91
  
  Dash = String$(96, "-")
  
  EmpRecSize = Len(Emp2Rec)
  TRecSize = Len(TransHRec(1))
  IdxRecLen = 2
  
  IdxFileSize& = FileSize(PRData + EmpIdxLName)
  NumOfRecs = IdxFileSize& \ IdxRecLen
  If NumOfRecs = 0 Then
    MsgBox "No records on file"
    Close
    Exit Sub
  End If
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  OpenEmpIdxLNameFile EmpIdxLNameHandle
  For x = 1 To NumOfRecs
     Get EmpIdxLNameHandle, x, IdxBuff(x)
  Next x
  Close EmpIdxLNameHandle
  
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, Unit(1)
  Close UnitHandle
  
  MaxLines = 50
  RptTitle$ = "Employee Retirement Report."
  FrmShowPctComp.Label1 = "Employee Retirement Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
  
  TCol = 40 - (Len(RptTitle$) \ 2) + 1
  PctRow = 11
  
  RHandle = FreeFile
  Open RptName$ For Output As RHandle
  RPTSetupPRN 11, RHandle
  THandle = FreeFile
  Open PRData + TransHistFileName For Random As THandle Len = TRecSize
  OpenEmpData2File DHandle
  PageGTot# = 0
  PageRTot# = 0
  PageMTot# = 0
  GoSub LRetRptHeader
  For RecNo = 1 To NumOfRecs
    UsingThisOne = False
    RetWage# = 0
    RETAMT# = 0
    MatAmt# = 0
    MatchCnt = 0
    Get DHandle, IdxBuff(RecNo), Emp2Rec
    If Emp2Rec.LastTransRec <= 0 Then
      GoTo SkipEm6
    End If
    TransRecNum& = CLng(Emp2Rec.LastTransRec)
    Do
      Get THandle, TransRecNum&, TransHRec(1)
      Select Case TransHRec(1).CheckDate
      Case LowDate To HighDate
        If Mid(Emp2Rec.EMPRETNO, 1, 1) = "R" And TransHRec(1).MatchRetAmt = 0 Then GoTo Retired
'        If UCase$(Left$(LTrim$(Emp2Rec.EMPRETTP), 1)) = "L" Then
        If Mid(Emp2Rec.EMPRETNO, 1, 1) = "R" Then
          Emp2Rec.EMPRETNO = Mid(Emp2Rec.EMPRETNO, 2, Len(Emp2Rec.EMPRETNO))
        End If
        If UCase$(Left$(LTrim$(Emp2Rec.EMPRETTP), 1)) = "L" And Mid(Emp2Rec.EMPRETNO, 1, 1) <> "T" Then '7/22/2004
          RetWage# = OldRound(RetWage# + TransHRec(1).RetGrossPay)
          RETAMT# = OldRound(RETAMT# + TransHRec(1).RetireAmt)
          MatAmt# = OldRound(MatAmt# + TransHRec(1).MatchRetAmt)
          MatchCnt = MatchCnt + 1
          UsingThisOne = True
        End If
      Case Else
      End Select
Retired:
      If TransHRec(1).PrevTransRec <= 0 Then
        If UsingThisOne Then
          GoSub PrintEmpRetLine
          If LineCnt >= MaxLines Then
            PageFF = True
            GoSub PrintPageTotals
            PageFF = False
            PageGTot# = 0
            PageRTot# = 0
            PageMTot# = 0
            Print #RHandle, FF$
            GoSub LRetRptHeader
          End If
        End If
        Exit Do
      Else
        TransRecNum& = CLng(TransHRec(1).PrevTransRec)
      End If
    Loop
'the progress bar is split into two parts...the first 50%
'is here and the second is under SkipEm7

SkipEm6:
    If RecNo / NumOfRecs < 0.5 Then
      FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        EnableCloseButton Me.hwnd, True
        Me.cmdEscape.Enabled = True
        Me.cmdProcess.Enabled = True
        Unload FrmShowPctComp
        Exit Sub
      End If
    End If

  Next
  GoSub PrintPageTotals
  PageGTot# = 0
  PageRTot# = 0
  PageMTot# = 0
  GoSub RetLTotals
  
  GTotal# = 0
  RTotal# = 0
  MTotal# = 0
  GoSub GRetRptHeader
  
  RptTitle$ = "Employee Retirement Report."
  FrmShowPctComp.Label1 = "Employee Retirement Report"
  FrmShowPctComp.Show , Me
  DoEvents
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
  
  PageGTot# = 0
  PageRTot# = 0
  PageMTot# = 0
  
  For RecNo = 1 To NumOfRecs
    UsingThisOne = False
    RetWage# = 0
    RETAMT# = 0
    MatAmt# = 0
    Get DHandle, CLng(IdxBuff(RecNo)), Emp2Rec
    
    If Emp2Rec.LastTransRec <= 0 Then
      GoTo SkipEm7
    End If
    
    TransRecNum& = CLng(Emp2Rec.LastTransRec)
    
    Do
      Get THandle, TransRecNum&, TransHRec(1)
      Select Case TransHRec(1).CheckDate
      Case LowDate To HighDate
        If Mid(Emp2Rec.EMPRETNO, 1, 1) = "R" And TransHRec(1).MatchRetAmt = 0 Then GoTo Retired2
        If Mid(Emp2Rec.EMPRETNO, 1, 1) = "R" Then
          Emp2Rec.EMPRETNO = Mid(Emp2Rec.EMPRETNO, 2, Len(Emp2Rec.EMPRETNO))
        End If
'        If UCase$(Left$(LTrim$(Emp2Rec.EMPRETTP), 1)) = "G" Then
        If UCase$(Left$(LTrim$(Emp2Rec.EMPRETTP), 1)) = "G" And Mid(Emp2Rec.EMPRETNO, 1, 1) <> "T" Then '7/22/2004
          RetWage# = OldRound(RetWage# + TransHRec(1).RetGrossPay) 'was GrossPay
          RETAMT# = OldRound(RETAMT# + TransHRec(1).RetireAmt)
          MatAmt# = OldRound(MatAmt# + TransHRec(1).MatchRetAmt)
          UsingThisOne = True
        End If
      Case Else
      End Select
Retired2:
      If TransHRec(1).PrevTransRec <= 0 Then
        If UsingThisOne Then
          GoSub PrintEmpRetLine
          If LineCnt >= MaxLines Then
            PageFF = True
            GoSub PrintPageTotals
            PageFF = False
            PageGTot# = 0
            PageRTot# = 0
            PageMTot# = 0
            Print #RHandle, FF$
            GoSub GRetRptHeader
          End If
        End If
        Exit Do
      Else
        TransRecNum& = CLng(TransHRec(1).PrevTransRec)
      End If
    Loop
    
SkipEm7:
    If RecNo / NumOfRecs > 0.5 Then
      FrmShowPctComp.ShowPctComp RecNo, NumOfRecs
      If FrmShowPctComp.Out = True Then
        Close
        FrmShowPctComp.Out = False
        Me.cmdEscape.Enabled = True
        Me.cmdProcess.Enabled = True
        EnableCloseButton Me.hwnd, True
        Unload FrmShowPctComp
        Exit Sub
      End If
    End If

  Next
  GoSub PrintPageTotals
  PageGTot# = 0
  PageRTot# = 0
  PageMTot# = 0
  GoSub RetGTotals
  RPTSetupPRN 123, RHandle '8/26...123 is the default end code
  
  Close RHandle
  Close THandle
  Close DHandle                'open employee data file
  Close
  
  If ThisCnt = 0 Then
    MsgBox "There are no records for the parameters entered."
    Me.cmdEscape.Enabled = True
    Me.cmdProcess.Enabled = True
    EnableCloseButton Me.hwnd, True
    Exit Sub
  End If
  
  ViewPrint RptName$, RptTitle$, True
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True
  EnableCloseButton Me.hwnd, True
  
  If ThisFile <> "NONE" Then
   Label5.Visible = True
   Shape3.Visible = True
   Label5.Caption = ThisFile + " saved successfully."
   Call Timer1_Timer 'added 5/27/04
  End If
  
  MainLog ("North Carolina retirement report processed.")
  
  Exit Sub
  
PrintEmpRetLine:
  ThisCnt = ThisCnt + 1
  Print #RHandle, Left$(Emp2Rec.EmpSSN, 3); "-"; Mid$(Emp2Rec.EmpSSN, 4, 2); "-"; Mid$(Emp2Rec.EmpSSN, 6, 4);
  Print #RHandle, Tab(14); QPTrim$(Left$(Emp2Rec.EMPRETNO, 12));
  Print #RHandle, Tab(28); QPTrim$(Emp2Rec.EmpLName); ", "; QPTrim$(Emp2Rec.EmpFName);
  Print #RHandle, Tab(56); Using("###,##0.00", RetWage#); Tab(70); Using("###,##0.00", RETAMT#);
  Print #RHandle, Tab(82); Using("###,##0.00", MatAmt#)
  LineCnt = LineCnt + 1         'employeesprinted = employeesprinted + 1

  PGTotal# = OldRound(PGTotal# + RetWage#)
  GTotal# = OldRound(GTotal# + RetWage#)
  PageGTot# = OldRound(PageGTot# + RetWage#)
  PRTotal# = OldRound(PRTotal# + RETAMT#)
  RTotal# = OldRound(RTotal# + RETAMT#)
  PageRTot# = OldRound(PageRTot# + RETAMT#)
  PMTotal# = OldRound(PMTotal# + MatAmt#)
  MTotal# = OldRound(MTotal# + MatAmt#)
  PageMTot# = OldRound(PageMTot# + MatAmt#)
  
Return
  
PrintPageTotals:
  Print #RHandle, Dash
  Print #RHandle, Tab(28); "Sub Total:";
  Print #RHandle, Tab(54); Using("#,###,##0.00", PageGTot#); Tab(68); Using("#,###,##0.00", PageRTot#);
  Print #RHandle, Tab(80); Using("#,###,##0.00", PageMTot#)
  If PageFF = False Then
    Print #RHandle, FF$
  End If
Return
  
  
RetLTotals:
  Print #RHandle, "Department of State Treasurer" '; Tab(70); '"     Page:"; LPage
  Print #RHandle, "North Carolina Law Enforcement Officers Benefit and Retirement Fund"
  Print #RHandle, "Monthly Report of Subject Wages And Retirement Contributions"
  Print #RHandle,
  Print #RHandle, "Reporting Unit: "; QPTrim$(Unit(1).UFEMPR); Tab(70); "Unit Code: "; QPTrim$(Unit(1).UFRETIDL)
  Print #RHandle, "Month: "; MonthName$(MonthNum); " "; YearNum
  Print #RHandle,
  Print #RHandle, "                                                   Wages Subject    Retirement     Employer"
  Print #RHandle, "                                                   To Retirement    Deductions        Match"
  'inserted above
  Print #RHandle, Dash
  Print #RHandle, Tab(28); "Law Enforcement Totals:";
  Print #RHandle, Tab(53); Using("#,###,##0.00", GTotal#); Tab(67); Using("#,###,##0.00", RTotal#);
  Print #RHandle, Tab(80); Using("#,###,##0.00", MTotal#)
  Print #RHandle, FF$
Return

LRetRptHeader:
  LPage = LPage + 1
  Print #RHandle, "Department of State Treasurer"; Tab(70); "     Page:"; LPage
  Print #RHandle, "North Carolina Law Enforcement Officers Benefit and Retirement Fund"
  Print #RHandle, "Monthly Report of Subject Wages And Retirement Contributions"
  Print #RHandle,
  Print #RHandle, "Reporting Unit: "; QPTrim$(Unit(1).UFEMPR); Tab(70); "Unit Code: "; QPTrim$(Unit(1).UFRETIDL)
  Print #RHandle, "Month: "; MonthName$(MonthNum); " "; YearNum
  Print #RHandle,
  Print #RHandle, "                                                    Wages Subject    Retirement    Employer"
  Print #RHandle, "Soc Sec #    Ret #         Employee Name            To Retirement    Deductions       Match"
  Print #RHandle, Dash
  LineCnt = 10
Return
  
GRetRptHeader:
  GPage = GPage + 1
  Print #RHandle, "Department of State Treasurer"; Tab(70); "     Page:"; GPage
  Print #RHandle, "North Carolina Local Government Retirement System."
  Print #RHandle, "Monthly Report of Subject Wages And Retirement Contributions"
  Print #RHandle,
  Print #RHandle, "Reporting Unit: "; QPTrim$(Unit(1).UFEMPR); Tab(70); "Unit Code: "; QPTrim$(Unit(1).UFRETID)
  Print #RHandle, "Month: "; MonthName$(MonthNum); " "; YearNum
  Print #RHandle,
  Print #RHandle, "                                                    Wages Subject    Retirement    Employer"
  Print #RHandle, "Soc Sec #    Ret #         Employee Name            To Retirement    Deductions       Match"
  Print #RHandle, Dash
  LineCnt = 10
Return

RetGTotals:
  Print #RHandle, "Department of State Treasurer" '; Tab(70); '"     Page:"; GPage
  Print #RHandle, "North Carolina Local Government Retirement System."
  Print #RHandle, "Monthly Report of Subject Wages And Retirement Contributions"
  Print #RHandle,
  Print #RHandle, "Reporting Unit: "; QPTrim$(Unit(1).UFEMPR); Tab(70); "Unit Code: "; QPTrim$(Unit(1).UFRETID)
  Print #RHandle, "Month: "; MonthName$(MonthNum); " "; YearNum
  Print #RHandle,
  Print #RHandle, "                                                   Wages Subject    Retirement     Employer"
  Print #RHandle, "                                                   To Retirement    Deductions        Match"
  Print #RHandle, Dash
  Print #RHandle, Tab(28); "General Totals:";
  Print #RHandle, Tab(53); Using("#,###,##0.00", GTotal#); Tab(67); Using("#,###,##0.00", RTotal#);
  Print #RHandle, Tab(80); Using("#,###,##0.00", MTotal#)
  Print #RHandle, FF$
Return

End Sub

Private Sub fpcmbDestination_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbDestination.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbDestination.ListIndex = -1
  End If
  If fpcmbDestination.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcomboPrintOpt.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbMagDiskYN_Change()
  If QPTrim$(fpcmbMagDiskYN.Text) = "" Then
    fpcmbMagDiskYN.Text = "No"
    fpcmbDestination.Enabled = False
  ElseIf QPTrim$(fpcmbMagDiskYN.Text) = "No" Then
    fpcmbDestination.Enabled = False
  ElseIf QPTrim$(fpcmbMagDiskYN.Text) = "Yes" Then
    fpcmbDestination.Enabled = True
  End If

End Sub

Private Sub fpcmbMagDiskYN_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbMagDiskYN.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbMagDiskYN.ListIndex = -1
  End If
  If fpcmbMagDiskYN.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbDestination.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If
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
Private Sub Ret2DiskOld()
  Dim Year$
  Dim LowDate As Integer
  Dim HiDate As Integer
  Dim NumOfRecs As Long
  Dim x As Long
  Dim hFile As Integer
  Dim EFile As Integer
  Dim RecNo As Long
  Dim TRec(1) As TransRecType
  Dim EGro#
  Dim EHrs#
  Dim OTHrs#
  Dim TransRecNum&
  Dim Term$
  Dim Birth$
  Dim HDay$
  Dim City$
  Dim UsingThisOne As Boolean
  Dim ERet#
  Dim TRecSize As Integer
  Dim EmpRecSize As Integer
  ReDim E2Rec(1) As EmpData2Type
  Dim ThisMonth$
  Dim MonthInt$
  Dim IdxHandle As Integer
  Dim RptName$
  
  If Date2Num(fptxtStart.Text) > Date2Num(fptxtEnd.Text) Then
    MsgBox "The beginning date comes after the ending date. Please re-enter the dates."
    fptxtStart.SetFocus
    Close
    Exit Sub
  End If
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
    
  frmProcessing.Show , Me
  DoEvents
  TRecSize = Len(TRec(1))
  EmpRecSize = Len(E2Rec(1))
  
  Year$ = Mid(fptxtStart.Text, 7, 4)
  
  LowDate = Date2Num(fptxtStart.Text)
  HiDate = Date2Num(fptxtEnd.Text)

  OpenEmpIdxNNameFile IdxHandle
  NumOfRecs = LOF(IdxHandle) / 2
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  For x = 1 To NumOfRecs
    Get IdxHandle, x, IdxBuff(x)
  Next x
  Close IdxHandle
  
  On Local Error GoTo ErrorHandler
  
  RptName$ = "A:\MAGRET" + ".RPT"
  Open RptName$ For Output As #1 Len = 16384
  OpenTransHistFile hFile
  OpenEmpData2File EFile
  GoSub PrintHeader
  
  For RecNo = 1 To NumOfRecs
    Get #EFile, IdxBuff(RecNo), E2Rec(1)
    E2Rec(1).EmpLName = E2Rec(1).EmpLName
    If E2Rec(1).LastTransRec <= 0 Then
      GoTo SkipEm2
    End If
    TransRecNum& = CLng(E2Rec(1).LastTransRec)
    Do
      Get #hFile, TransRecNum&, TRec(1)
      Select Case TRec(1).CheckDate
      Case LowDate To HiDate
        EGro# = RoundDbl(EGro# + TRec(1).GrossPay)
        EHrs# = RoundDbl(EHrs# + TRec(1).RegHrsWork + TRec(1).SickUsed + TRec(1).CompUsed + TRec(1).VacUsed)
        OTHrs# = RoundDbl(OTHrs# + TRec(1).OTHrsPaid)
        UsingThisOne = True
      Case Else
      End Select
      If TRec(1).PrevTransRec <= 0 Then
        If UsingThisOne Then
          GoSub PrintThisOne
        End If
        Exit Do
      Else
        TransRecNum& = CLng(TRec(1).PrevTransRec)
      End If
    Loop
SkipEm2:
  Next
  
  Unload frmProcessing
  
RetExitRpt:
  Close
  EnableCloseButton Me.hwnd, True
  Me.cmdEscape.Enabled = True
  Me.cmdProcess.Enabled = True

  Exit Sub
  
ErrorHandler:
  MsgBox "An error has occurred in accessing Drive A: and data has not been saved to disk."
  Close
  Unload frmProcessing
  Exit Sub
  
Return

PrintThisOne:
  If E2Rec(1).EMPTDATE > 0 Then
    Term$ = MakeRegDateDash(E2Rec(1).EMPTDATE)
  Else
    Term$ = "  -  -    "
  End If
  
  If E2Rec(1).EMPBDAY > -29219 Then
    Birth$ = MakeRegDateDash(E2Rec(1).EMPBDAY)
  Else
    Birth$ = "  -  -    "
  End If
  If E2Rec(1).EMPHDATE > -29219 Then
    HDay$ = MakeRegDateDash(E2Rec(1).EMPHDATE)
  Else
    HDay$ = "  -  -    "
  End If
  City$ = Space$(30)
  LSet City$ = E2Rec(1).EmpCity
  Mid$(City$, 20) = (E2Rec(1).EmpState + E2Rec(1).EmpZip)
  Write #1, QPTrim$(E2Rec(1).EmpSSN), (QPTrim$(E2Rec(1).EmpLName) + " " + QPTrim$(E2Rec(1).EmpFName)), QPTrim$(E2Rec(1).EmpAddr1), City$, Birth$, HDay$, Term$, QPTrim$(Using$("####0.00", EHrs#)), QPTrim$(Using$("####0.00", OTHrs#)), _
QPTrim$(Using$("#####0.00", EGro#))
  UsingThisOne = False

  EHrs# = 0
  OTHrs# = 0
  EGro# = 0
  ERet# = 0
  Return
  
PrintHeader:
  Write #1, "SSN", "LastName  FirstName", "Addrs", "City    ST     Zip", "BirthDate", "HireDate", "Term Date", "REGHours", "OT Hours", "Gross"
  Return
  
End Sub

Function RoundDbl#(DblNum#)
  RoundDbl# = (Int((DblNum# * 100) + 0.5) / 100)
End Function

Private Sub Ret2Disk(ByRef ThisFile$)
  Dim EYear$
  Dim RetType$
  Dim BMonth$
  Dim EMonth$
  Dim Year$, Year2$
  Dim hFile As Integer, EFile As Integer
  Dim LowDate As Integer
  Dim HiDate As Integer
  Dim EmpFirst$
  Dim x As Integer
  Dim IdxRec As NumbSortIdxType
  Dim IdxHandle As Integer
  Dim EGro#
  Dim ERet#
  Dim UsingThisOne As Boolean
  ReDim TRec(1) As TransRecType
  ReDim E2Rec(1) As EmpData2Type
  ReDim RetRec(1 To 2) As RetRecType
  Dim RHandle As Integer
  Dim Wage$
  Dim DotPos As Integer
  Dim RetA$
  Dim OTHrs#
  Dim EHrs#
  Dim NumOfRecs As Long
  Dim RptName$, TransRecNum&
  Dim RecNo As Long
  Dim UnitRec As UnitFileRecType
  Dim UnitHandle As Integer
  
  If Date2Num(fptxtStart.Text) > Date2Num(fptxtEnd.Text) Then
    MsgBox "The beginning date comes after the ending date. Please re-enter the dates."
    fptxtStart.SetFocus
    Close
    Exit Sub
  End If
  EnableCloseButton Me.hwnd, False
  Me.cmdEscape.Enabled = False
  Me.cmdProcess.Enabled = False
    
  BMonth = Mid(fptxtStart.Text, 1, 2)
  EMonth = Mid(fptxtEnd.Text, 1, 2)
  Year$ = Mid(fptxtStart.Text, 7, 4)
  Year2$ = Mid(fptxtEnd.Text, 7, 4)
  If BMonth <> EMonth Then
    If MsgBox("the beginning month is not the same as the ending month. The ending month will be reported. OK to continue?", vbYesNo) = vbNo Then
      Close
      fptxtStart.SetFocus
      Exit Sub
    End If
  End If
  
  If Year <> Year2 Then
    If MsgBox("The beginning year and the ending year are not the same. Using ending year. OK to continue?", vbYesNo) = vbNo Then
      Close
      fptxtStart.SetFocus
      Exit Sub
    End If
  End If
  
  EYear = Mid(fptxtEnd.Text, 9, 2)
  LowDate = Date2Num(fptxtStart.Text)
  HiDate = Date2Num(fptxtEnd.Text)
  
  OpenEmpIdxNNameFile IdxHandle
  NumOfRecs = LOF(IdxHandle) \ 2
  ReDim IdxBuff(1 To NumOfRecs) As Integer
  For x = 1 To NumOfRecs
    Get IdxHandle, x, IdxBuff(x)
  Next x
  Close IdxHandle
  
  frmRet2DiskModal.Show vbModal
  '------------------altered below 5/27/04
  If frmRet2DiskModal.fptxtChoice.Text = "general" Then
    Unload frmRet2DiskModal
    RetType$ = "G"
  ElseIf frmRet2DiskModal.fptxtChoice.Text = "law" Then
    Unload frmRet2DiskModal
    RetType$ = "L"
  Else
    Unload frmRet2DiskModal
    Close
    Exit Sub
  End If
  '-------^^^^-------altered above 5/27/04
  
  OpenUnitFile UnitHandle
  Get UnitHandle, 1, UnitRec
  Close UnitHandle
  
  Select Case RetType$
  Case "G"
    RptName$ = "GEN"
    RetRec(1).Unit = QPTrim$(UnitRec.UFRETID) '03/31/04 "99202" '"96511"
'    RetRec(1).Unit = "96511"
    RetRec(2).Unit = RetRec(1).Unit
  Case "L"
    RptName$ = "LEO"
    RetRec(1).Unit = QPTrim$(UnitRec.UFRETIDL) '03/31/04"71385" '"73165"
'    RetRec(1).Unit = "73165"
    RetRec(2).Unit = RetRec(1).Unit
  End Select
  ThisFile = RptName$ + EMonth$ + EYear$ + ".txt"
  
  RetRec(2).CrLf = Chr$(13) + Chr$(10)
  
  If QPTrim$(fpcmbDestination.Text) = "Citipak Directory" Then
    On Error GoTo DirectoryError
    frmProcessing.Show
    DoEvents
    frmProcessing.Label1.Caption = "Saving " + ThisFile
    DoEvents
    RptName$ = StartPath + "\" + RptName$ + EMonth$ + EYear$ + ".txt"
    ThisFile = RptName$
  Else
    On Error GoTo MagneticError
    frmProcessing.Show
    DoEvents
    frmProcessing.Label1.Caption = "Saving A:\" + RptName$ + EMonth$ + EYear$
    DoEvents
    KillFile "A:\" + RptName$ + EMonth$ + EYear$
    RptName$ = "A:\" + RptName$ + EMonth$ + EYear$
    ThisFile = RptName$
  End If
  
  Open RptName$ For Random As #1 Len = Len(RetRec(1))
  On Local Error GoTo ErrorHandler

  OpenTransHistFile hFile
  OpenEmpData2File EFile

  For RecNo = 1 To NumOfRecs
    Get #EFile, IdxBuff(RecNo), E2Rec(1)
'    If Left$(E2Rec(1).EMPRETTP, 1) = RetType$ Then
    If Left$(E2Rec(1).EMPRETTP, 1) = RetType$ And Mid(E2Rec(1).EMPRETNO, 1, 1) <> "T" Then '7/22/2004
      If E2Rec(1).LastTransRec <= 0 Then
        GoTo SkipEm2
      End If
      TransRecNum& = CLng(E2Rec(1).LastTransRec)
      Do
        Get #hFile, TransRecNum&, TRec(1)

        Select Case TRec(1).CheckDate
        Case LowDate To HiDate
          If Mid(E2Rec(1).EMPRETNO, 1, 1) = "R" And TRec(1).MatchRetAmt = 0 Then GoTo Retired
          If Mid(E2Rec(1).EMPRETNO, 1, 1) = "R" Then
            E2Rec(1).EMPRETNO = Mid(E2Rec(1).EMPRETNO, 2, Len(E2Rec(1).EMPRETNO))
          End If
          EGro# = RoundDbl(EGro# + TRec(1).RetGrossPay)
          ERet# = RoundDbl(ERet# + TRec(1).RetireAmt)
          UsingThisOne = True
        Case Else
        End Select
Retired:
        If TRec(1).PrevTransRec <= 0 Then
          If UsingThisOne Then
            GoSub PrintThisOne
          End If
          Exit Do
        Else
          TransRecNum& = CLng(TRec(1).PrevTransRec)
        End If
      Loop
SkipEm2:
    End If
  Next
  Unload frmProcessing
  DoEvents
  
RetExitRpt:
  Close
  DoEvents
  

  Exit Sub
  
ErrorHandler:
  MsgBox "ERROR: If this problem persists please consult Southern Software."
  Close
  Unload frmProcessing
  Exit Sub
  Return
  
DirectoryError:
  MsgBox "ERROR: There was a problem writing the file to a magnetic disk. File not saved."
  Close
  ThisFile = "NONE"
  Unload frmProcessing
  Exit Sub
  Return
  
MagneticError:
  MsgBox "ERROR: There was a problem writing the file to a magnetic disk. File not saved."
  Close
  ThisFile = "NONE"
  Unload frmProcessing
  Exit Sub
  Return
  
PrintThisOne:
  LSet RetRec(1) = RetRec(2)

  RetRec(1).RetNum = QPTrim$(E2Rec(1).EMPRETNO)  '6   '6-11
  EmpFirst$ = QPTrim$(E2Rec(1).EmpFName)
  RetRec(1).FirstN = Left$(EmpFirst$, 1)         '1
  RetRec(1).MidN = Right$(EmpFirst$, 1)          '1   '13
  RetRec(1).LastN = QPTrim$(E2Rec(1).EmpLName)   '11  '14-24
  RetRec(1).SSN = QPTrim$(E2Rec(1).EmpSSN)       ' 9   '25-33
  RetRec(1).Fill1 = ""

  RetRec(1).EMonth = EMonth$
  RetRec(1).EYear = EYear$

  RetRec(1).Fill2 = ""

  Wage$ = QPTrim$(Using$("######.##", (EGro#)))
  DotPos = InStr(Wage$, ".")
  Wage$ = Mid$(Wage$, 1, DotPos - 1) + Right$(Wage$, 2)
  Wage$ = "0000000" + Wage$
  RetRec(1).WageAmt = Right$(Wage$, 7)    '7   '51-57
  RetRec(1).NegWage = ""             '1   '58

  RetA$ = QPTrim$(Using$("######.##", (ERet#)))
  DotPos = InStr(RetA$, ".")
  RetA$ = Mid$(RetA$, 1, DotPos - 1) + Right$(RetA$, 2)
  RetA$ = "0000000" + RetA$
  RetRec(1).RETAMT = Right$(RetA$, 7)      '7   '59-65
  RetRec(1).NegRet = ""              '1   '66
  RetRec(1).Fill3 = ""
  Put #1, , RetRec(1)


  UsingThisOne = False

  EHrs# = 0
  OTHrs# = 0
  EGro# = 0
  ERet# = 0
  Return
End Sub

Private Sub Timer1_Timer() 'added 5/27/04
  Static tog As Boolean
  
  Timer1.Enabled = True
  tog = Not tog
  If tog Then
    Label5.Visible = False
  Else
    Label5.Visible = True
  End If
 
End Sub
