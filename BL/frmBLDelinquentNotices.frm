VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLDelinquentNotices 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Delinquent Notice Printning"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmBLDelinquentNotices.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   7284
      Left            =   1416
      TabIndex        =   9
      Top             =   768
      Width           =   8796
      _Version        =   196609
      _ExtentX        =   15515
      _ExtentY        =   12848
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
      Picture         =   "frmBLDelinquentNotices.frx":08CA
      Begin LpLib.fpCombo fpcmbPrintOrder 
         Height          =   405
         Left            =   3360
         TabIndex        =   0
         Tag             =   $"frmBLDelinquentNotices.frx":08E6
         Top             =   1920
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
         ColDesigner     =   "frmBLDelinquentNotices.frx":0992
      End
      Begin LpLib.fpCombo fpcmbPrintOpt 
         Height          =   405
         Left            =   3165
         TabIndex        =   6
         Tag             =   $"frmBLDelinquentNotices.frx":0C89
         Top             =   4995
         Width           =   3810
         _Version        =   196608
         _ExtentX        =   6720
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
         ColDesigner     =   "frmBLDelinquentNotices.frx":0D58
      End
      Begin LpLib.fpCombo fpcmbFees 
         Height          =   405
         Left            =   4845
         TabIndex        =   5
         Tag             =   $"frmBLDelinquentNotices.frx":104F
         Top             =   4365
         Width           =   885
         _Version        =   196608
         _ExtentX        =   1561
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
         ColDesigner     =   "frmBLDelinquentNotices.frx":119C
      End
      Begin LpLib.fpCombo fpcmbRange 
         Height          =   405
         Left            =   3120
         TabIndex        =   4
         Tag             =   $"frmBLDelinquentNotices.frx":1493
         Top             =   3750
         Width           =   4050
         _Version        =   196608
         _ExtentX        =   7144
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
         ColDesigner     =   "frmBLDelinquentNotices.frx":16C3
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   630
         Left            =   3675
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "Press 'Cancel' to exit this screen and return to the 'Penalty Processing' menu."
         Top             =   6000
         Width           =   1875
         _Version        =   131072
         _ExtentX        =   3307
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
         ButtonDesigner  =   "frmBLDelinquentNotices.frx":19BA
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   636
         Left            =   5808
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   $"frmBLDelinquentNotices.frx":1B99
         Top             =   6000
         Width           =   1884
         _Version        =   131072
         _ExtentX        =   3323
         _ExtentY        =   1122
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
         ButtonDesigner  =   "frmBLDelinquentNotices.frx":1C58
      End
      Begin EditLib.fpDateTime fptxtNoticeDate 
         Height          =   348
         Left            =   2448
         TabIndex        =   1
         Tag             =   "Enter the date in the 'Notice Date' field you wish to appear on the delinquent notice as the day the notice was issued."
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
         Text            =   "04/28/2003"
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
      Begin EditLib.fpDateTime fptxtPayByDate 
         Height          =   348
         Left            =   6144
         TabIndex        =   2
         Tag             =   $"frmBLDelinquentNotices.frx":1E37
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
         Text            =   "04/28/2003"
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
      Begin EditLib.fpDateTime fptxtXDate 
         Height          =   348
         Left            =   3504
         TabIndex        =   3
         Tag             =   $"frmBLDelinquentNotices.frx":1ECF
         Top             =   3120
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
         Text            =   "04/28/2003"
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
      Begin EditLib.fpText fptxtDlqNum 
         Height          =   396
         Left            =   5136
         TabIndex        =   16
         TabStop         =   0   'False
         Tag             =   $"frmBLDelinquentNotices.frx":1FAF
         Top             =   1152
         Width           =   732
         _Version        =   196608
         _ExtentX        =   1291
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
         ControlType     =   1
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   150
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
      Begin fpBtnAtlLibCtl.fpBtn fpcmdXList 
         Height          =   348
         Left            =   5280
         TabIndex        =   19
         TabStop         =   0   'False
         Tag             =   $"frmBLDelinquentNotices.frx":204C
         Top             =   3120
         Width           =   1932
         _Version        =   131072
         _ExtentX        =   3408
         _ExtentY        =   614
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
         ButtonDesigner  =   "frmBLDelinquentNotices.frx":2143
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
         Height          =   636
         Left            =   1248
         TabIndex        =   21
         TabStop         =   0   'False
         Tag             =   $"frmBLDelinquentNotices.frx":2329
         Top             =   6000
         Width           =   2172
         _Version        =   131072
         _ExtentX        =   3831
         _ExtentY        =   1122
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
         ButtonDesigner  =   "frmBLDelinquentNotices.frx":23F9
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
         Height          =   252
         Left            =   1296
         TabIndex        =   22
         Top             =   6672
         Width           =   2100
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1632
         TabIndex        =   20
         Top             =   3840
         Width           =   1356
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   3996
         Left            =   528
         Top             =   1680
         Width           =   7788
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Include Fees?:"
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
         Left            =   3024
         TabIndex        =   18
         Top             =   4464
         Width           =   1692
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Delinquent Notice # #:"
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
         Left            =   2784
         TabIndex        =   17
         Top             =   1248
         Width           =   2172
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1536
         TabIndex        =   15
         Top             =   3168
         Width           =   1788
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
         Left            =   1872
         TabIndex        =   14
         Top             =   2016
         Width           =   1308
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Print Delinquent Notices"
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
         Left            =   2256
         TabIndex        =   13
         Top             =   480
         Width           =   4332
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   684
         Left            =   2016
         Top             =   336
         Width           =   4908
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
         Height          =   348
         Left            =   1536
         TabIndex        =   12
         Top             =   5088
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Notice Date:"
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
         Left            =   768
         TabIndex        =   11
         Top             =   2592
         Width           =   1500
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pay By Date:"
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
         Left            =   4560
         TabIndex        =   10
         Top             =   2592
         Width           =   1404
      End
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   444
      Left            =   1704
      TabIndex        =   23
      Top             =   8256
      Width           =   780
      _Version        =   131072
      _ExtentX        =   1376
      _ExtentY        =   783
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
      Height          =   7548
      Left            =   1296
      Top             =   660
      Width           =   9060
   End
End
Attribute VB_Name = "frmBLDelinquentNotices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class
  Dim ExpDate$
  Dim PayDate$
  Dim XYear As String
  Dim PenAmt As Boolean
  Dim FormNum As Integer
  
Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "F1 &Turn Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    lblBalloon.Visible = True
    fptxtDlqNum.ToolTipText = ""
    fpcmbPrintOrder.ToolTipText = ""
    fptxtNoticeDate.ToolTipText = ""
    fptxtPayByDate.ToolTipText = ""
    fptxtXDate.ToolTipText = ""
    fpcmdXList.ToolTipText = ""
    fpcmbRange.ToolTipText = ""
    fpcmbFees.ToolTipText = ""
    fpcmbPrintOpt.ToolTipText = ""
    cmdHelp.ToolTipText = ""
    cmdExit.ToolTipText = ""
    cmdProcess.ToolTipText = ""
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
'    fptxtDlqNum.ToolTipText = "This number refers to the Delinquent Notice number selected on the Town Setup screen."
'    fpcmbPrintOrder.ToolTipText = "Notices can be printed in alphabetical order or in numerical order."
'    fptxtNoticeDate.ToolTipText = "Enter the date the delinquent notice will indicate is the date it was issued."
'    fptxtPayByDate.ToolTipText = "Enter the date the recipient of this notice must pay the delinquent business license fee."
'    fptxtXDate.ToolTipText = "Enter the date the recipient's business license expires."
'    fpcmdXList.ToolTipText = "Press for a concise explanation of the details of this screen."
'    fpcmbRange.ToolTipText = "You can elect to print all those delinquent up to and including the expiration date or just those delinquent on the date entered."
'    fpcmbFees.ToolTipText = "Select 'Yes' if you want the delinquent forms to show all business license fees currently past due."
'    fpcmbPrintOpt.ToolTipText = "Select graphical to print on a laser printer or choose text to print on a dot matrix printer."
'    cmdHelp.ToolTipText = "Press 'Turn Help On' to activate informational balloons. Press 'Turn Help Off' to deactivate informational balloons."
'    cmdExit.ToolTipText = "Press to exit this screen."
'    cmdProcess.ToolTipText = "Press to activate the print delinquent notices to the screen."
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
    Case vbKeyF4:
      SendKeys "%E"
      Call fpcmdXList_Click
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
      KillFile "dlnqnotice.dat"
      ClearInUse PWcnt
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLDelinquentNotices.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub fpcmbFees_Change()
  If QPTrim$(fpcmbFees.Text) = "" Then
    fpcmbFees.Text = "Yes"
  End If
End Sub

Private Sub fpcmbFees_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbFees.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbFees.ListIndex = -1
  End If
  If fpcmbFees.ListDown <> True Then
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

Private Sub fpcmbPrintOpt_Change()
  If QPTrim$(fpcmbPrintOpt.Text) = "" Then
    fpcmbPrintOpt.Text = "Graphical"
  End If
End Sub

Private Sub fpcmbPrintOrder_Change()
  If QPTrim$(fpcmbPrintOrder.Text) = "" Then
    fpcmbPrintOrder.Text = "Billing Name Order"
  End If
End Sub

Private Sub fpcmbPrintOrder_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  If KeyCode = vbKeySpace Then
    fpcmbPrintOrder.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPrintOrder.ListIndex = -1
  End If
  If fpcmbPrintOrder.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtNoticeDate.SetFocus
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

Private Sub cmdExit_Click()
  KillFile "dlnqnotice.dat"
  frmBLPenProcMenu.Show
  DoEvents
  Unload frmBLDelinquentNotices
End Sub

Private Sub cmdProcess_Click()
  If fpcmbPrintOpt.Text = "Graphical" Then
    Call PrintGraphics
  ElseIf fpcmbPrintOpt.Text = "Text" Then
    If QPTrim$(fptxtDlqNum.Text) = "4" Then
      frmBLMessageBoxJr.Label1.Caption = "Delinquent form #4 is only available in graphics format. Business license is switching to that format for printing."
      frmBLMessageBoxJr.Label1.Top = 700
      frmBLMessageBoxJr.Show vbModal
      Call PrintGraphics
      Close
      Exit Sub
    End If
    frmBLMessageBoxJr.Label1.Caption = "Pitch 10 is recommended for this report."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Call PrintText
  Else
    Exit Sub
  End If
End Sub

Private Sub LoadMe()
  Dim Year As Integer
  Dim PayOn As Integer
  Dim TownRec As TownSetUpType
  Dim THandle As Integer
  Dim One As Integer
  Dim DHandle As Integer
  
  On Error GoTo ERRORSTUFF
  
  lblBalloon.Visible = False
'  fptxtDlqNum.ToolTipText = "This number refers to the Delinquent Notice number selected on the Town Setup screen."
'  fpcmbPrintOrder.ToolTipText = "Notices can be printed in alphabetical order or in numerical order."
'  fptxtNoticeDate.ToolTipText = "Enter the date the delinquent notice will indicate is the date it was issued."
'  fptxtPayByDate.ToolTipText = "Enter the date the recipient of this notice must pay the delinquent business license fee."
'  fptxtXDate.ToolTipText = "Enter the date the recipient's business license expires."
'  fpcmdXList.ToolTipText = "Press for a concise explanation of the details of this screen."
'  fpcmbRange.ToolTipText = "You can elect to print all those delinquent up to and including the expiration date or just those delinquent on the date entered."
'  fpcmbFees.ToolTipText = "Select 'Yes' if you want the delinquent forms to show all business license fees currently past due."
'  fpcmbPrintOpt.ToolTipText = "Select graphical to print on a laser printer or choose text to print on a dot matrix printer."
'  cmdHelp.ToolTipText = "Press 'Turn Help On' to activate informational balloons. Press 'Turn Help Off' to deactivate informational balloons."
'  cmdExit.ToolTipText = "Press to exit this screen."
'  cmdProcess.ToolTipText = "Press to activate the print delinquent notices to the screen."
  One = 1
  DHandle = FreeFile
  Open "dlnqnotice.dat" For Output As DHandle Len = 2
  Print #DHandle, One
  Close DHandle
  
  OpenTownFile THandle
  Get THandle, 1, TownRec
  Close THandle
  PenAmt = False
  
Form3:
  fptxtDlqNum.Text = TownRec.DLQNotice
  
  'no fees are displayed on delinquent form 3 anyway
  If TownRec.DLQNotice = 3 Then
    fpcmbFees.Enabled = False
    fpcmbFees.Text = "NA"
  Else
    fpcmbFees.Text = "Yes"
    fpcmbFees.AddItem "Yes"
    fpcmbFees.AddItem "No"
  End If
  
  fptxtNoticeDate.Text = Date$
  PayOn = Date2Num(Date$)
  PayOn = PayOn + 30
  fptxtPayByDate.Text = MakeRegDate(PayOn)
  Year = Val(Mid(Date$, 7, 4))
  ExpDate$ = "12-31-" + QPTrim$(Str$(Year))
  fptxtXDate.Text = ExpDate$
  XYear = Mid(fptxtXDate.Text, 7, 4)
  If TownRec.DLQNotice = 3 Then
    Label8.Tag = "Delinquent form #3 does not show any license fees so the option to show fees is disabled."
  Else
    fpcmbFees.Tag = "You have the option of including or excluding business license related fees on the delinquent notice."
  End If
  PayDate$ = "03-31-" + Mid(Date$, 7, 4)
  fpcmbPrintOrder.Text = "Billing Name Order"
  fpcmbPrintOrder.AddItem "Billing Name Order"
  fpcmbPrintOrder.AddItem "Account Number Order"
  fpcmbPrintOpt.AddItem "Graphical"
  fpcmbPrintOpt.AddItem "Text"
  fpcmbPrintOpt.Text = "Graphical"
  fpcmbRange.Text = "Up To And Include This Expiration"
  fpcmbRange.AddItem "Up To And Include This Expiration"
  fpcmbRange.AddItem "This Expiration Only"
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLDelinquentNotices", "LoadMe", Erl)
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

Private Sub PrintText()
  Dim ReportFile$
  Dim FF$, x As Double
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeHandle As Integer
  Dim NumOfARCatRecs As Integer
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Double
  Dim CustNameIdxRec As CustNameIdxType ' CustSearchNameIdxType
  Dim CustNumIdxRec As CustNumIdxType
  Dim IdxHandle As Integer
  Dim NumOfCustIdxRecs As Double
  Dim NameFlag As Boolean
  Dim NumFlag As Boolean
  Dim RptHandle As Integer
  Dim Page As Integer
  Dim TCat$, ZCnt&, cnt&
  Dim LNotDate$, LPayDate$
  Dim LNDLen As Integer
  Dim LNDTab As Integer
  Dim ExpDate As Integer
  Dim PenPct$
  Dim CatLine$, C$, Year As Integer
  Dim Lp As Integer
  Dim CatCode$, LicTotal#, ZZCnt As Integer
  Dim TotalCust As Integer
  Dim CODEDESC$, CodeType$
  Dim CustomerNumber As Integer
  Dim ZZ As Integer, Snt As Long
  Dim Amt#
  Dim DESC$, DESC1$
  Dim BaseAmt1#, Revenue1#, Percent1#, Maximum1#
  Dim BaseAmt2#, Revenue2#, Percent2#, Maximum2#
  Dim BaseAmt3#, Revenue3#, Percent3#, Maximum3#
  Dim BaseAmt4#, Revenue4#, Percent4#, Maximum4#
  Dim BaseAmt5#, Revenue5#, Percent5#, Maximum5#
  Dim BaseAmt6#, Revenue6#, Percent6#, Maximum6#
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim DLQType As Integer
  Dim XDatePLusOne As Integer
  Dim TownLen As Integer
  Dim DateLen As Integer
  Dim TabLen1 As Integer
  Dim TabLen2 As Integer
  Dim TabLen3 As Integer
  Dim TabLen4 As Integer
  Dim ThisTab As Integer
  Dim Nextcnt As Integer
  Dim GoodThru$
  Dim SubTotal As Double
  Dim GTotal As Double
  Dim PctTotal As Double
  Dim UseThisPct As Double
  Dim PctCnt As Integer
  Dim PrintCnt As Integer
  Dim One As Integer
  Dim DHandle As Integer
  
  On Error GoTo ERRORSTUFF
  
  cmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  
  If QPTrim$(fpcmbFees.Text) = "Yes" Then
    If Exist("artmppen.dat") Then
      frmBLMessageBoxJrWOpts.Label1.Caption = "There are currently unposted penalty calculations on file. These unposted penalty calculations will not appear on any delinquent notices printed now. Do you want to post the latest penalty calculations now?"
      frmBLMessageBoxJrWOpts.Label1.Top = 600
      frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Post"
      frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
      frmBLMessageBoxJrWOpts.Show vbModal
      If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
        Unload frmBLMessageBoxJrWOpts
        Call frmBLPenProcMenu.cmdPost_Click
      Else
        DoEvents
        Unload frmBLMessageBoxJrWOpts
        frmBLMessageBoxJrWOpts.Label1.Caption = "Do you wish to continue printing the delinquent notice even though the penalty amounts will exclude the latest calculations?"
        frmBLMessageBoxJrWOpts.Label1.Top = 800
        frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
        frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
        frmBLMessageBoxJrWOpts.Show vbModal
        If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
          Unload frmBLMessageBoxJrWOpts
          MainLog ("The user was warned that they would be printing delinquent notices but that the amounts showing would exclude the current penalty posting calculations.")
        Else
          Unload frmBLMessageBoxJrWOpts
          Close
          Exit Sub
        End If
      End If
    End If
  End If
  'dlnqnotice.dat is created when this screen loads and is
  'deleted when this screen exits...however, if there is an
  'unposted penalty file then when the user attempts to print
  'delinquent notices he will be asked to decide if he wishes
  'to post first...if he chooses yes then the post warning pop-up
  'activates from the penalty processing menu...if at that
  'time the user elects to abort the post then the code on
  'the penalty processing menu (post) deletes the dlnqnotice.dat
  'file as a way to signal this screen that the user aborted the
  'post...the program returns to this point from that code
  If Not Exist("dlnqnotice.dat") Then
    One = 1
    DHandle = FreeFile
    Open "dlnqnotice.dat" For Output As DHandle Len = 2
    Print #DHandle, One
    Close DHandle
    frmBLMessageBoxJrWOpts.Label1.Caption = "Do you wish to continue printing the delinquent notice even though the penalty amounts will exclude the latest calculations?"
    frmBLMessageBoxJrWOpts.Label1.Top = 800
    frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
    frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
    frmBLMessageBoxJrWOpts.Show vbModal
    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
      Unload frmBLMessageBoxJrWOpts
      MainLog ("The user was warned that they would be printing delinquent notices but that the amounts showing would exclude the current penalty posting calculations.")
    Else
      Unload frmBLMessageBoxJrWOpts
      Close
      Exit Sub
    End If
  End If
  
  Nextcnt = 1
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  
  DLQType = TownRec.DLQNotice
  
  ReportFile$ = "AREXPLIC.PRN"  'Report File Name
  FF$ = Chr$(12)
  MaxLines = 53
  LineCnt = 0
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  If QPTrim$(fpcmbPrintOrder.Text) = "Billing Name Order" Then
    NameFlag = True
    NumFlag = False
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "Account Number Order" Then
    NumFlag = True
    NameFlag = False
  Else
    fpcmbPrintOrder.BackColor = &HFFFF&
    frmBLMessageBoxJr.Label1.Caption = "Please make a selection for Print Order."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fpcmbPrintOrder.BackColor = &HFFFFFF
    fpcmbPrintOrder.SetFocus
    Close
    Exit Sub
  End If
  
  If NameFlag = True Then
'    OpenSrchNameIdxFile IdxHandle
    OpenCustNameIdxFile IdxHandle
    NumOfCustIdxRecs = LOF(IdxHandle) / Len(CustNameIdxRec)
  Else
    OpenCustNumIdxFile IdxHandle
    NumOfCustIdxRecs = LOF(IdxHandle) / Len(CustNumIdxRec)
  End If
  
  If NumOfCustIdxRecs = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "There are no business customers indexed."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  OpenCustFile CustHandle
  
  ReDim IdxRecs(1 To NumOfCustIdxRecs) As Double
  If NameFlag = True Then
    For x = 1 To NumOfCustIdxRecs
      Get IdxHandle, x, CustNameIdxRec
      IdxRecs(x) = CustNameIdxRec.CustRec
    Next x
  Else
      For x = 1 To NumOfCustIdxRecs
      Get IdxHandle, x, CustNumIdxRec
      IdxRecs(x) = CustNumIdxRec.CustRec
    Next x
  End If
  Close IdxHandle
  
  OpenCatCodeFile CodeHandle
  NumOfARCatRecs = LOF(CodeHandle) \ Len(CodeRec)
  
  If NumOfARCatRecs = 0 Then
    frmBLMessageBoxJrWOpts.Label1.Caption = "There are no category codes saved. Delinquent notices require category codes. Would you like to jump to the category edit screen now?"
    frmBLMessageBoxJrWOpts.Label1.Top = 700
    frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Jump"
    frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
    frmBLMessageBoxJrWOpts.Show vbModal
    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
      Unload frmBLMessageBoxJrWOpts
      Close
      frmBLCatEdit.Show
      DoEvents
      Unload frmBLDelinquentNotices
      Exit Sub
    Else
      Unload frmBLMessageBoxJrWOpts
      Close
      Exit Sub
    End If
  End If
  
  LNotDate$ = MakeLongDate(fptxtNoticeDate.Text)
  LPayDate$ = MakeLongDate(fptxtPayByDate.Text)
  LNDLen = Len(LNotDate$)
  LNDTab = (40 - (LNDLen / 2))
  LNDTab = LNDTab - 1
  ExpDate = Date2Num%(fptxtXDate.Text)
  
  Year = Val(Right$(Date$, 4))
  GoodThru = Mid(fptxtXDate.Text, 1, 6) + CStr(Year + 1)
  Year = Year - 1
  
  frmBLShowPctComp.Label1 = "Loading Detailed Customer List"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  cmdHelp.Enabled = False
  
  ' Print Main Body
  For cnt = 1 To NumOfCustIdxRecs 'NumOfCustRecs
    Get CustHandle, IdxRecs(cnt), CustRec
    If QPTrim$(CustRec.Inactive) = "Y" Then GoTo Inactive
    CustomerNumber = IdxRecs(cnt)
    If InStr(fpcmbRange.Text, "Only") Then
      If CustRec.VALID <> ExpDate Then GoTo Inactive
    Else
      If CustRec.VALID > ExpDate Then GoTo Inactive
    End If
    If UCase$(CustRec.Deleted) <> "Y" Or QPTrim$(CustRec.SortName) <> "DELETED" Then
      If CustRec.AcctBal > 0 Then
        If DLQType = 1 Then GoSub DLQType1
        If DLQType = 2 Then GoSub DLQType2
        If DLQType = 3 Then GoSub DLQType3
        PrintCnt = PrintCnt + 1
      End If
    End If
    frmBLShowPctComp.ShowPctComp cnt, NumOfCustIdxRecs 'NumOfCustRecs
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
Inactive:
  Next cnt
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  cmdHelp.Enabled = True
  
  Close         'Close all open files now
  If PrintCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No business license fees are delinquent for expiration date " + fptxtXDate + "."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  ViewPrint ReportFile$, "Delinquent Notice Printing", True
  KillFile ReportFile$
  
  MainLog ("Delinquent notices (#" + QPTrim$(fptxtDlqNum.Text) + ") printed in text format.")
  
  Exit Sub
  
DLQType1:
  TownLen = Len(QPTrim$(TownRec.DlqTownName))
  TabLen1 = TownLen / 2
  TabLen1 = 39 - TabLen1
  DateLen = Len(QPTrim$(LNotDate$))
  TabLen2 = DateLen / 2
  TabLen2 = Abs(39 - TabLen2)
  TownLen = Len(QPTrim$(TownRec.DlqAdd1))
  TabLen3 = TownLen / 2
  TabLen3 = Abs(39 - TabLen3)
  TownLen = Len(QPTrim$(TownRec.DlqCity)) + Len(QPTrim$(TownRec.DlqState)) + Len(QPTrim$(TownRec.DlqZip))
  TownLen = TownLen + 3
  TabLen4 = TownLen / 2
  TabLen4 = Abs(39 - TabLen4)
  Print #RptHandle, Tab(27); "-----------------------"
  Print #RptHandle, Tab(27); "***DELINQUENT NOTICE***"
  Print #RptHandle, Tab(27); "-----------------------"
  Print #RptHandle, ""
  Print #RptHandle, Tab(TabLen2); LNotDate$
  Print #RptHandle, ""
  Print #RptHandle, Tab(TabLen1); QPTrim$(TownRec.DlqTownName) '"TOWN OF CAROLINA BEACH"
  Print #RptHandle, Tab(TabLen3); QPTrim$(TownRec.DlqAdd1) ' "1121 NORTH LAKE PARK BLVD."
  Print #RptHandle, Tab(TabLen4); QPTrim$(TownRec.DlqCity) + ", "; QPTrim$(TownRec.DlqState) + " " + QPTrim$(TownRec.DlqZip) '"CAROLINA BEACH, N.C. 28428"
  Print #RptHandle, Tab(30); "TEL " + QPTrim$(TownRec.DlqPhone) '"TEL 910-458-2999"
  Print #RptHandle, ""
  Print #RptHandle, ""
  Print #RptHandle, ""
  Print #RptHandle, Tab(52); "Account ID: " + CStr(CustomerNumber)
  Print #RptHandle, Tab(7); QPTrim$(CustRec.BillName)
  Print #RptHandle, Tab(7); QPTrim$(CustRec.ADDRESS1)
  Print #RptHandle, Tab(7); QPTrim$(CustRec.ADDRESS2)
  Print #RptHandle, Tab(7); QPTrim$(CustRec.City); ", "; QPTrim$(CustRec.State); " "; QPTrim$(CustRec.ZipCode)
  Print #RptHandle, ""
  Print #RptHandle, ""
  Print #RptHandle, Tab(7); "According to our records, your " + XYear + " Business License has not been"
  If PenAmt = True Then
    Print #RptHandle, Tab(7); "purchased as of today. All licenses are now subject to a penalty"
  Else
    Print #RptHandle, Tab(7); "purchased as of today. All licenses are now subject to a penalty"
  End If
  Print #RptHandle, Tab(7); "and will NOT be issued unless the penalty amount is included with"
  Print #RptHandle, Tab(7); "your payment. We realize that you are very busy, but we would like"
  Print #RptHandle, Tab(7); "for you to take the time to purchase this license."
  Print #RptHandle, ""
  Print #RptHandle, ""
  Print #RptHandle, Tab(16); "We show your account delinquent for the following"
  Print #RptHandle, "                license code(s):"
  Print #RptHandle, ""
  If fpcmbFees.Text = "Yes" Then
    Print #RptHandle, Tab(16); CustRec.DESC1; Tab(50); Using("$##,###,##0.00", CustRec.Fee1)
    If QPTrim$(CustRec.DESC2) <> "" Then
      Print #RptHandle, Tab(16); CustRec.DESC2; Tab(50); Using("$##,###,##0.00", CustRec.Fee2)
    Else
      Print #RptHandle, ""
    End If
    If QPTrim$(CustRec.DESC3) <> "" Then
      Print #RptHandle, Tab(16); CustRec.DESC3; Tab(50); Using("$##,###,##0.00", CustRec.Fee3)
    Else
      Print #RptHandle, ""
    End If
    If QPTrim$(CustRec.DESC4) <> "" Then
      Print #RptHandle, Tab(16); CustRec.DESC4; Tab(50); Using("$##,###,##0.00", CustRec.Fee4)
    Else
      Print #RptHandle, ""
    End If
    If QPTrim$(CustRec.DESC5) <> "" Then
      Print #RptHandle, Tab(16); CustRec.DESC5; Tab(50); Using("$##,###,##0.00", CustRec.Fee5)
    Else
      Print #RptHandle, ""
    End If
    Print #RptHandle, ""
    
    SubTotal = CustRec.Fee1 + CustRec.Fee2 + CustRec.Fee3 + CustRec.Fee4 + CustRec.Fee5
    PctTotal = CustRec.PenBal
    If PctTotal >= CustRec.PenBal Then 'use the penalty that applies only to this license
      UseThisPct = PctTotal
      If PctTotal > CustRec.PenBal Then
        PctCnt = PctCnt + 1
      End If
    Else
      UseThisPct = CustRec.PenBal - PctTotal '.PenBal may already have another penalty charge from another transaction
    End If
    
    GTotal = SubTotal + PctTotal
    Print #RptHandle, Tab(16); "Penalty Charge"; Tab(50); Using("$##,###,##0.00", UseThisPct)
    Print #RptHandle, ""
    Print #RptHandle, Tab(16); "Total Business License Charges:"; Tab(50); Using("$##,###,##0.00", GTotal)
    Print #RptHandle, ""
    Print #RptHandle, Tab(16); "Total Outstanding Balance:"; Tab(50); Using("$##,###,##0.00", CustRec.AcctBal)
    Print #RptHandle, ""
  Else
    Print #RptHandle, Tab(16); CustRec.DESC1
    If QPTrim$(CustRec.DESC2) <> "" Then
      Print #RptHandle, Tab(16); CustRec.DESC2
    Else
      Print #RptHandle, ""
    End If
    If QPTrim$(CustRec.DESC3) <> "" Then
      Print #RptHandle, Tab(16); CustRec.DESC3
    Else
      Print #RptHandle, ""
    End If
    If QPTrim$(CustRec.DESC4) <> "" Then
      Print #RptHandle, Tab(16); CustRec.DESC4
    Else
      Print #RptHandle, ""
    End If
    If QPTrim$(CustRec.DESC5) <> "" Then
      Print #RptHandle, Tab(16); CustRec.DESC5
    Else
      Print #RptHandle, ""
    End If
    Print #RptHandle, ""
  End If
  
  Print #RptHandle, Tab(7); "Please remit your payment (including penalty) to this office"
  Print #RptHandle, Tab(7); "NO later than: "; LPayDate$; "."; " If you have questions"
  Print #RptHandle, Tab(7); "regarding your license, please feel free to contact our office."
  Print #RptHandle, ""
  Print #RptHandle, Tab(7); "If payment has been made prior to receiving this notice, please"
  Print #RptHandle, Tab(7); "disregard this notice."
  Print #RptHandle, ""
  Print #RptHandle, ""
  Print #RptHandle, ""
  Print #RptHandle, Tab(52); "Sincerely,"
  Print #RptHandle, ""
  Print #RptHandle, Tab(52); QPTrim$(TownRec.DlqAdminName) '"DAWN S. JOHNSON"
  Print #RptHandle, Tab(52); QPTrim$(TownRec.DlqAdminTitle) ' "FINANCE DIRECTOR"
  Print #RptHandle, Chr$(12);

  Return
  
GetCode:
  For Snt& = 1 To NumOfARCatRecs
    Get CodeHandle, Snt&, CodeRec
    If QPTrim$(CodeRec.CatCode) = CatCode$ Then
      CODEDESC$ = QPTrim$(CodeRec.CODEDESC)
      Select Case CodeRec.CodeType
      Case "F"
        Amt# = CodeRec.Fee
        CodeType$ = CodeRec.CodeType
      Case "M"
        DESC1$ = "Per Each"
        Amt# = CodeRec.Fee
        CodeType$ = CodeRec.CodeType
      Case Is = "S"
        BaseAmt1# = CodeRec.BaseAmt1
        Revenue1# = CodeRec.Recpt1
        Percent1# = CodeRec.Percent1
        Maximum1# = CodeRec.Maximum1
        BaseAmt2# = CodeRec.BaseAmt2
        Revenue2# = CodeRec.Recpt2
        Percent2# = CodeRec.Percent2
        Maximum2# = CodeRec.Maximum2
        BaseAmt3# = CodeRec.BaseAmt3
        Revenue3# = CodeRec.Recpt3
        Percent3# = CodeRec.Percent3
        Maximum3# = CodeRec.Maximum3
        BaseAmt4# = CodeRec.BaseAmt4
        Revenue4# = CodeRec.Recpt4
        Percent4# = CodeRec.Percent4
        Maximum4# = CodeRec.Maximum4
        BaseAmt5# = CodeRec.BaseAmt5
        Revenue5# = CodeRec.Recpt5
        Percent5# = CodeRec.Percent5
        Maximum5# = CodeRec.Maximum5
        BaseAmt6# = CodeRec.BaseAmt6
        Revenue6# = CodeRec.Recpt6
        Percent6# = CodeRec.Percent6
        Maximum6# = CodeRec.Maximum6
        CodeType$ = CodeRec.CodeType
      Case Else
        CodeType$ = "N"
      End Select
      Exit For
    End If
  Next Snt&
  

GotCode:
  Return
  
DLQType2:
  TownLen = Len(QPTrim$(TownRec.DlqTownName))

  ThisTab = TownLen / 2
  ThisTab = Abs(40 - ThisTab)

  Print #RptHandle, ""
  Print #RptHandle, Tab(2); CStr(Nextcnt); Tab(28); "***DELINQUENT NOTICE***"
  Print #RptHandle, Tab(28); "ANNUAL BUSINESS LICENSE"
  Print #RptHandle, Tab(ThisTab); QPTrim$(TownRec.DlqTownName)
  Print #RptHandle, Tab(LNDTab); UCase(LNotDate$)
  Print #RptHandle, ""
  Print #RptHandle, Tab(5); CustRec.BillName; Tab(45); "BUSINESS ACCOUNT # "; IdxRecs(cnt)
  Print #RptHandle, Tab(5); CustRec.ADDRESS1
  Print #RptHandle, Tab(5); CustRec.ADDRESS2
  Print #RptHandle, Tab(5); RTrim$(CustRec.City); " " + RTrim$(CustRec.State) + " " + RTrim$(CustRec.ZipCode)
  Print #RptHandle, ""
  Print #RptHandle, Tab(2); "ACCORDING TO OUR RECORDS YOUR " + XYear + " BUSINESS LICENSE HAS NOT BEEN PURCHASED "
  Print #RptHandle, Tab(2); "AS OF TODAY. ALL BUSINESS LICENSE FEES ARE NOW SUBJECT TO A PENALTY CHARGE."
  Print #RptHandle, Tab(2); "PLEASE REMIT YOUR PAYMENT (INCLUDING THE PENALTY) NO LATER THAN "
  Print #RptHandle, Tab(2); UCase(LPayDate$) + ". IF PAYMENT HAS ALREADY BEEN MADE PRIOR TO THE "
  Print #RptHandle, Tab(2); "DATE ABOVE THEN PLEASE DISREGARD THIS NOTICE AND THANK YOU FOR YOUR PAYMENT. "
  Print #RptHandle, ""
  Print #RptHandle, String$(79, "-")
  Print #RptHandle, Tab(2); "Code"; Tab(9); "Type of License"
  Print #RptHandle, String$(79, "-")
  Lp = 21
'-----------------------------------------------------------
  CatCode$ = QPTrim$(CustRec.BILLCAT1)
  GoSub GetCode
  Print #RptHandle, Tab(2); CustRec.BILLCAT1;
  Print #RptHandle, Tab(9); CustRec.DESC1; Tab(57); "BASIS AMT"; Tab(69); "LICENSE AMT"
  Lp = Lp + 1
  If CodeType$ = "S" Then
    Print #RptHandle, Tab(2); "Min Due"; Tab(13); "For Recpts Up To"; Tab(31); "Plus"; Tab(39); "Of Recpts Over"
    Lp = Lp + 1
    If BaseAmt1# > 0 Then
      Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt1#); Tab(14); Using("$###,###,##0.00", Revenue1#); Tab(30); Using("#0.00%  ", (Percent1# / 100)); Tab(38); Using("$###,###,##0.00", Maximum1#)
      Lp = Lp + 1
    End If
    If BaseAmt2# > 0 Then
      Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt2#); Tab(14); Using("$###,###,##0.00", Revenue2#); Tab(30); Using("#0.00%  ", (Percent2# / 100)); Tab(38); Using("$###,###,##0.00", Maximum2#)
      Lp = Lp + 1
    End If
    If BaseAmt3# > 0 Then
      Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt3#); Tab(14); Using("$###,###,##0.00", Revenue3#); Tab(30); Using("#0.00%  ", (Percent3# / 100)); Tab(38); Using("$###,###,##0.00", Maximum3#)
      Lp = Lp + 1
    End If
    If BaseAmt4# > 0 Then
      Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt4#); Tab(14); Using("$###,###,##0.00", Revenue4#); Tab(30); Using("#0.00%  ", (Percent4# / 100)); Tab(38); Using("$###,###,##0.00", Maximum4#)
      Lp = Lp + 1
    End If
    If BaseAmt5# > 0 Then
      Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt5#); Tab(14); Using("$###,###,##0.00", Revenue5#); Tab(30); Using("#0.00%  ", (Percent5# / 100)); Tab(38); Using("$###,###,##0.00", Maximum5#)
      Lp = Lp + 1
    End If
    If BaseAmt6# > 0 Then
      Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt6#); Tab(14); Using("$###,###,##0.00", Revenue6#); Tab(30); Using("#0.00%  ", (Percent6# / 100)); Tab(38); Using("$###,###,##0.00", Maximum6#)
      Lp = Lp + 1
    End If
    If fpcmbFees.Text = "Yes" Then
      Print #RptHandle, ; Tab(52); Using("$##,###,##0.00", CustRec.REV1); Tab(69); Using("$###,##0.00", CustRec.Fee1)
    Else
      Print #RptHandle, ; Tab(55); "___________ "; Tab(68); "____________ "
    End If
    Lp = Lp + 1
  End If
  If CodeType$ = "F" Then
    Print #RptHandle, Tab(57); "Flat Fee: "; Tab(69); Using("$###,##0.00", Amt#)
    Lp = Lp + 1
  End If
  If CodeType$ = "M" Then
    If fpcmbFees.Text = "Yes" Then
      Print #RptHandle, Tab(9); "Rate Per Unit: "; Tab(29); Using("$#,###,##0.00", Amt#)
      Print #RptHandle, Tab(9); "Times Number Of Units: "; Tab(37); Using("####0", CustRec.REV1); Tab(57); "*********"; Tab(67); Using("$#,###,##0.00", CustRec.Fee1)
    Else
      Print #RptHandle, Tab(9); "Rate Per Unit: "; Tab(29); Using("$#,###,##0.00", Amt#)
      Print #RptHandle, Tab(9); "Times Number Of Units: "; Tab(36); "______"; Tab(57); "*********"; Tab(67); "_____________"
    End If
    Lp = Lp + 2
  End If
  Print #RptHandle, String$(79, "-")
  Lp = Lp + 1

'-----------------------------------------------------------
  If Len(QPTrim$(CustRec.BILLCAT2)) = 0 Then GoTo EndAtmore1
  CatCode$ = QPTrim$(CustRec.BILLCAT2)
  GoSub GetCode
  Print #RptHandle, Tab(2); CustRec.BILLCAT2;
  Print #RptHandle, Tab(9); CustRec.DESC2; Tab(57); "BASIS AMT"; Tab(69); "LICENSE AMT"
  Lp = Lp + 1
  If CodeType$ = "S" Then
    Print #RptHandle, Tab(2); "Min Due"; Tab(13); "For Recpts Up To"; Tab(31); "Plus"; Tab(39); "Of Recpts Over"
    Lp = Lp + 1
    If BaseAmt1# > 0 Then
      Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt1#); Tab(14); Using("$###,###,##0.00", Revenue1#); Tab(30); Using("#0.00%  ", (Percent1# / 100)); Tab(38); Using("$###,###,##0.00", Maximum1#)
      Lp = Lp + 1
    End If
    If BaseAmt2# > 0 Then
      Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt2#); Tab(14); Using("$###,###,##0.00", Revenue2#); Tab(30); Using("#0.00%  ", (Percent2# / 100)); Tab(38); Using("$###,###,##0.00", Maximum2#)
      Lp = Lp + 1
    End If
    If BaseAmt3# > 0 Then
      Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt3#); Tab(14); Using("$###,###,##0.00", Revenue3#); Tab(30); Using("#0.00%  ", (Percent3# / 100)); Tab(38); Using("$###,###,##0.00", Maximum3#)
      Lp = Lp + 1
    End If
    If BaseAmt4# > 0 Then
      Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt4#); Tab(14); Using("$###,###,##0.00", Revenue4#); Tab(30); Using("#0.00%  ", (Percent4# / 100)); Tab(38); Using("$###,###,##0.00", Maximum4#)
      Lp = Lp + 1
    End If
    If BaseAmt5# > 0 Then
      Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt5#); Tab(14); Using("$###,###,##0.00", Revenue5#); Tab(30); Using("#0.00%  ", (Percent5# / 100)); Tab(38); Using("$###,###,##0.00", Maximum5#)
      Lp = Lp + 1
    End If
    If BaseAmt6# > 0 Then
      Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt6#); Tab(14); Using("$###,###,##0.00", Revenue6#); Tab(30); Using("#0.00%  ", (Percent6# / 100)); Tab(38); Using("$###,###,##0.00", Maximum6#)
      Lp = Lp + 1
    End If
    If fpcmbFees.Text = "Yes" Then
      Print #RptHandle, ; Tab(52); Using("$##,###,##0.00", CustRec.REV2); Tab(69); Using("$###,##0.00", CustRec.Fee2)
    Else
      Print #RptHandle, ; Tab(55); "___________ "; Tab(68); "____________ "
    End If
    Lp = Lp + 1
  End If
  If CodeType$ = "F" Then
    Print #RptHandle, Tab(57); "Flat Fee: "; Tab(69); Using("$###,##0.00", Amt#)
    Lp = Lp + 1
  End If
  If CodeType$ = "M" Then
    If fpcmbFees.Text = "Yes" Then
      Print #RptHandle, Tab(9); "Rate Per Unit: "; Tab(29); Using("$#,###,##0.00", Amt#)
      Print #RptHandle, Tab(9); "Times Number Of Units: "; Tab(37); Using("####0", CustRec.REV2); Tab(57); "*********"; Tab(67); Using("$#,###,##0.00", CustRec.Fee2)
    Else
      Print #RptHandle, Tab(9); "Rate Per Unit: "; Tab(29); Using("$#,###,##0.00", Amt#)
      Print #RptHandle, Tab(9); "Times Number Of Units: "; Tab(36); "______"; Tab(57); "*********"; Tab(67); "_____________"
    End If
    Lp = Lp + 2
  End If
  Print #RptHandle, String$(79, "-")
  Lp = Lp + 1
'-----------------------------------------------------------
  If Len(QPTrim$(CustRec.BILLCAT3)) = 0 Then GoTo EndAtmore1
  CatCode$ = QPTrim$(CustRec.BILLCAT3)
  GoSub GetCode
  Print #RptHandle, Tab(2); CustRec.BILLCAT3;
  Print #RptHandle, Tab(9); CustRec.DESC3; Tab(57); "BASIS AMT"; Tab(69); "LICENSE AMT"
  Lp = Lp + 1
  If CodeType$ = "S" Then
    Print #RptHandle, Tab(2); "Min Due"; Tab(13); "For Recpts Up To"; Tab(31); "Plus"; Tab(39); "Of Recpts Over"
    Lp = Lp + 1
    If BaseAmt1# > 0 Then
      Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt1#); Tab(14); Using("$###,###,##0.00", Revenue1#); Tab(30); Using("#0.00%  ", (Percent1# / 100)); Tab(38); Using("$###,###,##0.00", Maximum1#)
      Lp = Lp + 1
    End If
    If BaseAmt2# > 0 Then
      Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt2#); Tab(14); Using("$###,###,##0.00", Revenue2#); Tab(30); Using("#0.00%  ", (Percent2# / 100)); Tab(38); Using("$###,###,##0.00", Maximum2#)
      Lp = Lp + 1
    End If
    If BaseAmt3# > 0 Then
      Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt3#); Tab(14); Using("$###,###,##0.00", Revenue3#); Tab(30); Using("#0.00%  ", (Percent3# / 100)); Tab(38); Using("$###,###,##0.00", Maximum3#)
      Lp = Lp + 1
    End If
    If BaseAmt4# > 0 Then
      Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt4#); Tab(14); Using("$###,###,##0.00", Revenue4#); Tab(30); Using("#0.00%  ", (Percent4# / 100)); Tab(38); Using("$###,###,##0.00", Maximum4#)
      Lp = Lp + 1
    End If
    If BaseAmt5# > 0 Then
      Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt5#); Tab(14); Using("$###,###,##0.00", Revenue5#); Tab(30); Using("#0.00%  ", (Percent5# / 100)); Tab(38); Using("$###,###,##0.00", Maximum5#)
      Lp = Lp + 1
    End If
    If BaseAmt6# > 0 Then
      Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt6#); Tab(14); Using("$###,###,##0.00", Revenue6#); Tab(30); Using("#0.00%  ", (Percent6# / 100)); Tab(38); Using("$###,###,##0.00", Maximum6#)
      Lp = Lp + 1
    End If
    If fpcmbFees.Text = "Yes" Then
      Print #RptHandle, ; Tab(52); Using("$##,###,##0.00", CustRec.REV3); Tab(69); Using("$###,##0.00", CustRec.Fee3)
    Else
      Print #RptHandle, ; Tab(55); "___________ "; Tab(68); "____________ "
    End If
    Lp = Lp + 1
  End If
  If CodeType$ = "F" Then
    Print #RptHandle, Tab(57); "Flat Fee: "; Tab(69); Using("$###,##0.00", Amt#)
    Lp = Lp + 1
  End If
  If CodeType$ = "M" Then
    If fpcmbFees.Text = "Yes" Then
      Print #RptHandle, Tab(9); "Rate Per Unit: "; Tab(29); Using("$#,###,##0.00", Amt#)
      Print #RptHandle, Tab(9); "Times Number Of Units: "; Tab(37); Using("####0", CustRec.REV3); Tab(57); "*********"; Tab(67); Using("$#,###,##0.00", CustRec.Fee3)
    Else
      Print #RptHandle, Tab(9); "Rate Per Unit: "; Tab(29); Using("$#,###,##0.00", Amt#)
      Print #RptHandle, Tab(9); "Times Number Of Units: "; Tab(36); "______"; Tab(57); "*********"; Tab(67); "_____________"
    End If
    Lp = Lp + 2
  End If
  Print #RptHandle, String$(79, "-")
  Lp = Lp + 1
  If Lp >= 52 Then 'if this customer has 4 full categories (10 lines each) then
  'if the page does not break here it will run over with the fifth code
    GoSub PrintHeader2
  End If

'-----------------------------------------------------------
  If Len(QPTrim$(CustRec.BILLCAT4)) = 0 Then GoTo EndAtmore1
  CatCode$ = QPTrim$(CustRec.BILLCAT4)
  GoSub GetCode
  Print #RptHandle, Tab(2); CustRec.BILLCAT4;
  Print #RptHandle, Tab(9); CustRec.DESC4; Tab(57); "BASIS AMT"; Tab(69); "LICENSE AMT"
  Lp = Lp + 1
  If CodeType$ = "S" Then
    Print #RptHandle, Tab(2); "Min Due"; Tab(13); "For Recpts Up To"; Tab(31); "Plus"; Tab(39); "Of Recpts Over"
    Lp = Lp + 1
    If BaseAmt1# > 0 Then
      Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt1#); Tab(14); Using("$###,###,##0.00", Revenue1#); Tab(30); Using("#0.00%  ", (Percent1# / 100)); Tab(38); Using("$###,###,##0.00", Maximum1#)
      Lp = Lp + 1
    End If
    If BaseAmt2# > 0 Then
      Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt2#); Tab(14); Using("$###,###,##0.00", Revenue2#); Tab(30); Using("#0.00%  ", (Percent2# / 100)); Tab(38); Using("$###,###,##0.00", Maximum2#)
      Lp = Lp + 1
    End If
    If BaseAmt3# > 0 Then
      Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt3#); Tab(14); Using("$###,###,##0.00", Revenue3#); Tab(30); Using("#0.00%  ", (Percent3# / 100)); Tab(38); Using("$###,###,##0.00", Maximum3#)
      Lp = Lp + 1
    End If
    If BaseAmt4# > 0 Then
      Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt4#); Tab(14); Using("$###,###,##0.00", Revenue4#); Tab(30); Using("#0.00%  ", (Percent4# / 100)); Tab(38); Using("$###,###,##0.00", Maximum4#)
      Lp = Lp + 1
    End If
    If BaseAmt5# > 0 Then
      Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt5#); Tab(14); Using("$###,###,##0.00", Revenue5#); Tab(30); Using("#0.00%  ", (Percent5# / 100)); Tab(38); Using("$###,###,##0.00", Maximum5#)
      Lp = Lp + 1
    End If
    If BaseAmt6# > 0 Then
      Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt6#); Tab(14); Using("$###,###,##0.00", Revenue6#); Tab(30); Using("#0.00%  ", (Percent6# / 100)); Tab(38); Using("$###,###,##0.00", Maximum6#)
      Lp = Lp + 1
    End If
    If fpcmbFees.Text = "Yes" Then
      Print #RptHandle, ; Tab(52); Using("$##,###,##0.00", CustRec.REV4); Tab(69); Using("$###,##0.00", CustRec.Fee4)
    Else
      Print #RptHandle, ; Tab(55); "___________ "; Tab(68); "____________ "
    End If
    Lp = Lp + 1
  End If
  If CodeType$ = "F" Then
    Print #RptHandle, Tab(57); "Flat Fee: "; Tab(69); Using("$###,##0.00", Amt#)
    Lp = Lp + 1
  End If
  If CodeType$ = "M" Then
    If fpcmbFees.Text = "Yes" Then
      Print #RptHandle, Tab(9); "Rate Per Unit: "; Tab(29); Using("$#,###,##0.00", Amt#)
      Print #RptHandle, Tab(9); "Times Number Of Units: "; Tab(37); Using("####0", CustRec.REV4); Tab(57); "*********"; Tab(67); Using("$#,###,##0.00", CustRec.Fee4)
    Else
      Print #RptHandle, Tab(9); "Rate Per Unit: "; Tab(29); Using("$#,###,##0.00", Amt#)
      Print #RptHandle, Tab(9); "Times Number Of Units: "; Tab(36); "______"; Tab(57); "*********"; Tab(67); "_____________"
    End If
    Lp = Lp + 2
  End If
  Print #RptHandle, String$(79, "-")
  Lp = Lp + 1
  If Lp >= 52 Then 'if this customer has 4 full categories (10 lines each) then
  'if the page does not break here it will run over with the fifth code
    GoSub PrintHeader2
  End If

'-----------------------------------------------------------
  If Len(QPTrim$(CustRec.BILLCAT5)) = 0 Then GoTo EndAtmore1
  CatCode$ = QPTrim$(CustRec.BILLCAT5)
  GoSub GetCode
  Print #RptHandle, Tab(2); CustRec.BILLCAT5;
  Print #RptHandle, Tab(9); CustRec.DESC5; Tab(57); "BASIS AMT"; Tab(69); "LICENSE AMT"
  Lp = Lp + 1
  If CodeType$ = "S" Then
    Print #RptHandle, Tab(2); "Min Due"; Tab(13); "For Recpts Up To"; Tab(31); "Plus"; Tab(39); "Of Recpts Over"
    If BaseAmt1# > 0 Then
      Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt1#); Tab(14); Using("$###,###,##0.00", Revenue1#); Tab(30); Using("#0.00%  ", (Percent1# / 100)); Tab(38); Using("$###,###,##0.00", Maximum1#)
      Lp = Lp + 1
    End If
    If BaseAmt2# > 0 Then
      Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt2#); Tab(14); Using("$###,###,##0.00", Revenue2#); Tab(30); Using("#0.00%  ", (Percent2# / 100)); Tab(38); Using("$###,###,##0.00", Maximum2#)
      Lp = Lp + 1
    End If
    If BaseAmt3# > 0 Then
      Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt3#); Tab(14); Using("$###,###,##0.00", Revenue3#); Tab(30); Using("#0.00%  ", (Percent3# / 100)); Tab(38); Using("$###,###,##0.00", Maximum3#)
      Lp = Lp + 1
    End If
    If BaseAmt4# > 0 Then
      Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt4#); Tab(14); Using("$###,###,##0.00", Revenue4#); Tab(30); Using("#0.00%  ", (Percent4# / 100)); Tab(38); Using("$###,###,##0.00", Maximum4#)
      Lp = Lp + 1
    End If
    If BaseAmt5# > 0 Then
      Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt5#); Tab(14); Using("$###,###,##0.00", Revenue5#); Tab(30); Using("#0.00%  ", (Percent5# / 100)); Tab(38); Using("$###,###,##0.00", Maximum5#)
      Lp = Lp + 1
    End If
    If BaseAmt6# > 0 Then
      Print #RptHandle, Tab(2); Using("$##0.00", BaseAmt6#); Tab(14); Using("$###,###,##0.00", Revenue6#); Tab(30); Using("#0.00%  ", (Percent6# / 100)); Tab(38); Using("$###,###,##0.00", Maximum6#)
      Lp = Lp + 1
    End If
    If fpcmbFees.Text = "Yes" Then
      Print #RptHandle, ; Tab(52); Using("$##,###,##0.00", CustRec.REV5); Tab(69); Using("$###,##0.00", CustRec.Fee5)
    Else
      Print #RptHandle, ; Tab(55); "___________ "; Tab(68); "____________ "
    End If
    Lp = Lp + 1
  End If
  If CodeType$ = "F" Then
    Print #RptHandle, Tab(57); "Flat Fee: "; Tab(69); Using("$###,##0.00", Amt#)
    Lp = Lp + 1
  End If
  If CodeType$ = "M" Then
    If fpcmbFees.Text = "Yes" Then
      Print #RptHandle, Tab(9); "Rate Per Unit: "; Tab(29); Using("$#,###,##0.00", Amt#)
      Print #RptHandle, Tab(9); "Times Number Of Units: "; Tab(37); Using("####0", CustRec.REV5); Tab(57); "*********"; Tab(67); Using("$#,###,##0.00", CustRec.Fee5)
    Else
      Print #RptHandle, Tab(9); "Rate Per Unit: "; Tab(29); Using("$#,###,##0.00", Amt#)
      Print #RptHandle, Tab(9); "Times Number Of Units: "; Tab(36); "______"; Tab(57); "*********"; Tab(67); "_____________"
    End If
    Lp = Lp + 2
  End If
      Print #RptHandle, String$(79, "-")
      Lp = Lp + 1
EndAtmore1:
    
    If Lp >= 34 Then
      GoSub PrintHeader2
    End If
    
    Print #RptHandle,
    If fpcmbFees.Text = "Yes" Then
      SubTotal = CustRec.Fee1 + CustRec.Fee2 + CustRec.Fee3 + CustRec.Fee4 + CustRec.Fee5
      PctTotal = CustRec.PenBal
      GTotal = SubTotal + PctTotal
      Print #RptHandle, Tab(5); "MAKE CHECKS PAYABLE TO:"; Tab(45); "LICENSE TOTAL: "; Tab(66); Using$("$##,###,##0.00", SubTotal)
      Print #RptHandle, Tab(5); QPTrim$(TownRec.DlqTownName)
      Print #RptHandle, Tab(5); QPTrim$(TownRec.DlqAdd1); Tab(45); "PENALTY:"; Tab(70); Using("$##,##0.00", PctTotal)
      Print #RptHandle, Tab(5); QPTrim$(TownRec.DlqCity) + ", " + QPTrim$(TownRec.DlqState) + " " + QPTrim$(TownRec.DlqZip)
      Print #RptHandle, Tab(45); "TOTAL LICENSE FEES:  "; Tab(66); Using("$##,###,##0.00", GTotal)
      Print #RptHandle, ""
      Print #RptHandle, Tab(45); "TOTAL BALANCE:  "; Tab(66); Using("$##,###,##0.00", CustRec.AcctBal)
    Else
      Print #RptHandle, Tab(5); "MAKE CHECKS PAYABLE TO:"; Tab(45); "LICENSE TOTAL: ____________________"
      Print #RptHandle, Tab(5); QPTrim$(TownRec.DlqTownName)
      Print #RptHandle, Tab(5); QPTrim$(TownRec.DlqAdd1); Tab(45); "PENALTY:"
      Print #RptHandle, Tab(5); QPTrim$(TownRec.DlqCity) + ", " + QPTrim$(TownRec.DlqState) + " " + QPTrim$(TownRec.DlqZip)
      Print #RptHandle, Tab(45); "TOTAL LICENSE FEES:     ____________________"
      Print #RptHandle, ""
      Print #RptHandle, Tab(45); "TOTAL BALANCE:          ____________________  ";
    End If
    
    Print #RptHandle, ""
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "WHERE APPLICABLE, ESTABLISHMENTS NOT PURCHASING A LICENSE BY " + fptxtPayByDate.Text
    Print #RptHandle, Tab(5); "WILL BE REPORTED TO THE ABC COMMISSION."
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "RENEWALS LICENSE VALID UNTIL " + GoodThru + "."
    Print #RptHandle, ""
    Print #RptHandle, Tab(5); "PLEASE CONTACT THE TOWN OFFICE WITH ANY QUESTIONS."
    Print #RptHandle, ""
    If QPTrim$(TownRec.DlqPhone2) = "(" Then
      TownRec.DlqPhone2 = ""
    End If
    If QPTrim$(TownRec.DlqFax) = "(" Then
      TownRec.DlqFax = ""
    End If
    If QPTrim$(TownRec.DlqPhone2) = "" And QPTrim$(TownRec.DlqFax) = "" Then
      Print #RptHandle, Tab(5); "TELEPHONE:" + QPTrim$(TownRec.DlqPhone)
    Else
      Print #RptHandle, Tab(5); "TELEPHONE:" + QPTrim$(TownRec.DlqPhone) + " OR " + QPTrim$(TownRec.DlqPhone2) + "  FAX:" + QPTrim$(TownRec.DlqFax)
    End If
    Print #RptHandle, ""
    Print #RptHandle, Chr$(12);
    Nextcnt = Nextcnt + 1

  Return
  
PrintHeader2:
  Print #RptHandle, Chr$(12)
  Print #RptHandle, Tab(28); "***DELINQUENT NOTICE***"
  Print #RptHandle, Tab(28); "ANNUAL BUSINESS LICENSE"
  Print #RptHandle, Tab(ThisTab); QPTrim$(TownRec.DlqTownName)
  Print #RptHandle, Tab(LNDTab); UCase(LNotDate$)
  Print #RptHandle, ""
  Lp = 5

  Return
'-----------------------------------------------------

DLQType3:
  'GOSUB MakeCatLine
  
  If Len(QPTrim$(TownRec.DlqTownName)) = 0 Then
    TownLen = Len(QPTrim$(TownRec.TownPhone))
  Else
    TownLen = Len(QPTrim$(TownRec.DlqTownName))
  End If
  TabLen1 = TownLen / 2
  TabLen1 = 39 - TabLen1
  DateLen = Len(QPTrim$(LNotDate$))
  TabLen2 = DateLen / 2
  TabLen2 = 39 - TabLen2
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  If Len(QPTrim$(TownRec.DlqTownName)) = 0 Then
    Print #RptHandle, Tab(TabLen1); QPTrim$(TownRec.TownName) '"TOWN OF EXMORE"
  Else
    Print #RptHandle, Tab(TabLen1); QPTrim$(TownRec.DlqTownName) '"TOWN OF EXMORE"
  End If
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(TabLen2); LNotDate$ 'LNDTab
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(5); QPTrim$(CustRec.BillName)
  Print #RptHandle, Tab(5); QPTrim$(CustRec.ADDRESS1)
  Print #RptHandle, Tab(5); QPTrim$(CustRec.ADDRESS2)
  Print #RptHandle, Tab(5); QPTrim$(CustRec.City); ", "; QPTrim$(CustRec.State); " "; QPTrim$(CustRec.ZipCode)
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(5); "Dear Business Owner:"
  Print #RptHandle,
  Print #RptHandle, Tab(5); "  According to our records, your APPLICATION FOR TOWN LICENSE(S) has not"
  Print #RptHandle, Tab(5); "yet been submitted for processing your " + XYear + " BUSINESS, PROFESSIONAL and"
  Print #RptHandle, Tab(5); "OCCUPATIONAL LICENSE(BPOL) TAX.  This application is required prior to"
  If Len(QPTrim$(TownRec.DlqCity)) = 0 Then
    Print #RptHandle, Tab(5); "issuance of any " + QPTrim$(TownRec.City) + " Business License.  The application form specifies"
  Else
    Print #RptHandle, Tab(5); "issuance of any " + QPTrim$(TownRec.DlqCity) + " Business License.  The application form"
  End If
  Print #RptHandle, Tab(5); "specifies a " + fptxtXDate.Text + " deadline for filing, and states that a penalty "
  Print #RptHandle, Tab(5); "may be assessed on delinquent applications. The application also states"
  Print #RptHandle, Tab(5); "a deadline of " + LPayDate$ + " for payment of applicable BPOL Tax,"
  Print #RptHandle, Tab(5); "as stated in Code of Virginia 58.1-3703.1-Uniform Ordinance Provisions."
  Print #RptHandle, Tab(5); "To avoid further action, please complete and return the APPLICATION FOR"
  Print #RptHandle, Tab(5); "TOWN LICENSE(S) immediately.  Failure to comply may result in legal"
  Print #RptHandle, Tab(5); "action including, but not limited to, an audit of business records, as"
  Print #RptHandle, Tab(5); "permitted in Code of Virginia 58.1-3110 and 58.1-3939.1"
  Print #RptHandle, Tab(5); ""
  Print #RptHandle, Tab(5); ""
  Print #RptHandle, Tab(5); ""
  Print #RptHandle, Tab(5); "  If there are any questions or if assistance is needed in completing the"
  Print #RptHandle, Tab(5); "form, please call " + QPTrim$(TownRec.DlqClerkName) + " at " + QPTrim$(TownRec.DlqPhone) + ", " + QPTrim(TownRec.DlqFirstDay) + " to " + QPTrim$(TownRec.DlqLastDay)   '  Monday - Friday 8:00 A.M."
  Print #RptHandle, Tab(5); "from " + QPTrim$(TownRec.DlqFirstHour) + "to " + QPTrim$(TownRec.DlqLastHour) + "."
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(5); "Cordially,"
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(5); QPTrim$(TownRec.DlqMayorCouncil) '"Mayor and Council"
  Print #RptHandle, Tab(5); QPTrim$(TownRec.DlqTownName) '"Town Of Exmore"
  Print #RptHandle,
  Print #RptHandle,
  Print #RptHandle, Tab(5); QPTrim$(TownRec.DlqAdminName) ' "Donald P. Kellam, Sr."
  Print #RptHandle, Tab(5); QPTrim$(TownRec.DlqAdminTitle) ' "BPOL Commissioner"
  Print #RptHandle, Chr$(12);

  Return

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLDelinquentNotices", "PrintText", Erl)
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
  Dim ReportFile$
  Dim FF$, x As Double, y As Integer
  Dim MaxLines As Integer
  Dim LineCnt As Integer
  Dim CodeRec As ARNewCatCodeRecType
  Dim CodeHandle As Integer
  Dim NumOfARCatRecs As Integer
  Dim CustRec As ARCustRecType
  Dim CustHandle As Integer
  Dim NumOfCustRecs As Double
  Dim CustNameIdxRec As CustNameIdxType ' CustSearchNameIdxType
  Dim CustNumIdxRec As CustNumIdxType
  Dim IdxHandle As Integer
  Dim NumOfCustIdxRecs As Double
  Dim NameFlag As Boolean
  Dim NumFlag As Boolean
  Dim RptHandle As Integer
  Dim Page As Integer
  Dim TCat$, ZCnt&, cnt&
  Dim LNotDate$, LPayDate$
  Dim LNDLen As Integer
  Dim LNDTab As Integer
  Dim ExpDate As Integer
  Dim PenPct$
  Dim CatLine$, C$, Year As Integer
  Dim Lp As Integer
  Dim CatCode$, LicTotal#, ZZCnt As Integer
  Dim TotalCust As Integer
  Dim CODEDESC$, CodeType$
  Dim CustomerNumber As Integer
  Dim ZZ As Integer, Snt As Long
  Dim Amt#
  Dim DESC$, DESC1$
  Dim BaseAmt1#, Revenue1#, Percent1#, Maximum1#
  Dim BaseAmt2#, Revenue2#, Percent2#, Maximum2#
  Dim BaseAmt3#, Revenue3#, Percent3#, Maximum3#
  Dim BaseAmt4#, Revenue4#, Percent4#, Maximum4#
  Dim BaseAmt5#, Revenue5#, Percent5#, Maximum5#
  Dim BaseAmt6#, Revenue6#, Percent6#, Maximum6#
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim DLQType As Integer
  Dim XDatePLusOne As Integer
  Dim TownLen As Integer
  Dim DateLen As Integer
  Dim TabLen1 As Integer
  Dim TabLen2 As Integer
  Dim dlm$, ExpYear$
  Dim GoodThru$
  Dim DlqCnt As Integer
  Dim SubTotal As Double
  Dim GTotal As Double
  Dim PctTotal As Double
  Dim UseThisPct As Double
  Dim PctCnt As Integer
  Dim One As Integer
  Dim DHandle As Integer
  Dim LaserRec4 As LaserLetterType4
  Dim LHandle As Integer
  Dim AddEmptyFields As Integer
  
  On Error GoTo ERRORSTUFF
  
  cmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  
  If QPTrim$(fpcmbFees.Text) = "Yes" Then
    If Exist("artmppen.dat") Then
      frmBLMessageBoxJrWOpts.Label1.Caption = "There are currently unposted penalty calculations on file. These unposted penalty calculations will not appear on any delinquent notices printed now. Do you want to post the latest penalty calculations now?"
      frmBLMessageBoxJrWOpts.Label1.Top = 600
      frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Post"
      frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
      frmBLMessageBoxJrWOpts.Show vbModal
      If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
        Unload frmBLMessageBoxJrWOpts
        Call frmBLPenProcMenu.cmdPost_Click
      Else
        DoEvents
        Unload frmBLMessageBoxJrWOpts
        frmBLMessageBoxJrWOpts.Label1.Caption = "Do you wish to continue printing the delinquent notice even though the penalty amounts will exclude the latest calculations?"
        frmBLMessageBoxJrWOpts.Label1.Top = 800
        frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
        frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
        frmBLMessageBoxJrWOpts.Show vbModal
        If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
          Unload frmBLMessageBoxJrWOpts
          MainLog ("The user was warned that they would be printing delinquent notices but that the amounts showing would exclude the current penalty posting calculations.")
        Else
          Unload frmBLMessageBoxJrWOpts
          Close
          Exit Sub
        End If
      End If
    End If
  End If
  'dlnqnotice.dat is created when this screen loads and is
  'deleted when this screen exits...however, if there is an
  'unposted penalty file then when the user attempts to print
  'delinquent notices he will be asked to decide if he wishes
  'to post first...if he chooses yes then the post warning pop-up
  'activates from the penalty processing menu...if at that
  'time the user elects to abort the post then the code on
  'the penalty processing menu (post) deletes the dlnqnotice.dat
  'file as a way to signal this screen that the user aborted the
  'post...the program returns to this point from that code
  If Not Exist("dlnqnotice.dat") Then
    One = 1
    DHandle = FreeFile
    Open "dlnqnotice.dat" For Output As DHandle Len = 2
    Print #DHandle, One
    Close DHandle
    frmBLMessageBoxJrWOpts.Label1.Caption = "Do you wish to continue printing the delinquent notice even though the penalty amounts will exclude the latest calculations?"
    frmBLMessageBoxJrWOpts.Label1.Top = 800
    frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Continue"
    frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
    frmBLMessageBoxJrWOpts.Show vbModal
    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
      Unload frmBLMessageBoxJrWOpts
      MainLog ("The user was warned that they would be printing delinquent notices but that the amounts showing would exclude the current penalty posting calculations.")
    Else
      Unload frmBLMessageBoxJrWOpts
      Close
      Exit Sub
    End If
  End If
  
  DlqCnt = 0
  dlm$ = "~"
  OpenTownFile TownHandle
  Get TownHandle, 1, TownRec
  Close TownHandle
  
  LNotDate$ = MakeLongDate(fptxtNoticeDate.Text)
  LPayDate$ = MakeLongDate(fptxtPayByDate.Text)
  ExpDate = Date2Num%(fptxtXDate.Text)
  Year = Val(Right$(Date$, 4))
  GoodThru = Mid(fptxtXDate.Text, 1, 6) + CStr(Year + 1)
  Year = Year - 1
  
  DLQType = TownRec.DLQNotice
  
  Select Case DLQType
    Case 1
      ReportFile$ = "BLRPTS\DLQ1.RPT"
    Case 2
      ReportFile$ = "BLRPTS\DLQ2.RPT"
    Case 3
      ReportFile$ = "BLRPTS\DLQ3.RPT"
    Case 4
      OpenLaserFile4 LHandle
      Get LHandle, 1, LaserRec4
      Close LHandle
      ReportFile$ = "BLRPTS\DLQ4.RPT"
    Case Else
      Close
      frmBLMessageBoxJr.Label1.Caption = "No delinquent notice has been selected."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      Exit Sub
  End Select
  
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  If QPTrim$(fpcmbPrintOrder.Text) = "Billing Name Order" Then
    NameFlag = True
    NumFlag = False
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "Account Number Order" Then
    NumFlag = True
    NameFlag = False
  Else
    fpcmbPrintOrder.BackColor = &HFFFF&
    frmBLMessageBoxJr.Label1.Caption = "Please make a selection for Print Order."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    fpcmbPrintOrder.BackColor = &HFFFFFF
    fpcmbPrintOrder.SetFocus
    Close
    Exit Sub
  End If
  
  If NameFlag = True Then
'    OpenSrchNameIdxFile IdxHandle
    OpenCustNameIdxFile IdxHandle
    NumOfCustIdxRecs = LOF(IdxHandle) / Len(CustNameIdxRec)
  Else
    OpenCustNumIdxFile IdxHandle
    NumOfCustIdxRecs = LOF(IdxHandle) / Len(CustNumIdxRec)
  End If
  
  If NumOfCustIdxRecs = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "There are no business customers indexed."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  OpenCustFile CustHandle
  
  ReDim IdxRecs(1 To NumOfCustIdxRecs) As Double
  If NameFlag = True Then
    For x = 1 To NumOfCustIdxRecs
      Get IdxHandle, x, CustNameIdxRec
      IdxRecs(x) = CustNameIdxRec.CustRec
    Next x
  Else
      For x = 1 To NumOfCustIdxRecs
      Get IdxHandle, x, CustNumIdxRec
      IdxRecs(x) = CustNumIdxRec.CustRec
    Next x
  End If
  Close IdxHandle
  
  OpenCatCodeFile CodeHandle
  NumOfARCatRecs = LOF(CodeHandle) \ Len(CodeRec)
  
  If NumOfARCatRecs = 0 Then
    frmBLMessageBoxJrWOpts.Label1.Caption = "There are no category codes saved. Delinquent notices require category codes. Would you like to jump to the category edit screen now?"
    frmBLMessageBoxJrWOpts.Label1.Top = 700
    frmBLMessageBoxJrWOpts.cmdCont.Text = "F10 Jump"
    frmBLMessageBoxJrWOpts.cmdExit.Text = "ESC Cancel"
    frmBLMessageBoxJrWOpts.Show vbModal
    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "continue" Then
      Unload frmBLMessageBoxJrWOpts
      Close
      frmBLCatEdit.Show
      DoEvents
      Unload frmBLDelinquentNotices
      Exit Sub
    Else
      Unload frmBLMessageBoxJrWOpts
      Close
      Exit Sub
    End If
  End If
  
  LNotDate$ = MakeLongDate(fptxtNoticeDate.Text)
  LPayDate$ = MakeLongDate(fptxtPayByDate.Text)
  ExpDate = Date2Num%(fptxtXDate.Text)
  ExpYear$ = Mid(fptxtXDate.Text, 7, 4)
  Year = Val(Right$(Date$, 4))
  Year = Year - 1
  
  frmBLShowPctComp.Label1 = "Loading Detailed Customer List"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  cmdHelp.Enabled = False
  
  ' Print Main Body
  For cnt = 1 To NumOfCustIdxRecs
    Get CustHandle, IdxRecs(cnt), CustRec
    If QPTrim$(CustRec.Inactive) = "Y" Then GoTo Inactive
    CustomerNumber = IdxRecs(cnt)
    If InStr(fpcmbRange.Text, "Only") Then
      If CustRec.VALID <> ExpDate Then GoTo Inactive
    Else
      If CustRec.VALID > ExpDate Then GoTo Inactive
    End If
    If UCase$(CustRec.Deleted) <> "Y" Or QPTrim$(CustRec.SortName) <> "DELETED" Then
      If CustRec.AcctBal > 0 Then
        If DLQType = 1 Then GoSub DLQType1
        If DLQType = 2 Then GoSub DLQType2
        If DLQType = 3 Then GoSub DLQType3
        If DLQType = 4 Then GoSub DLQType4
      End If
    End If
    frmBLShowPctComp.ShowPctComp cnt, NumOfCustIdxRecs 'NumOfCustRecs
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
Inactive:
  Next cnt
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  cmdHelp.Enabled = True
  
  Close         'Close all open files now
  
  If DlqCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "No business license fees are delinquent for expiration date " + fptxtXDate + "."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If
  
  Select Case DLQType
    Case 1
      arBLDLQ1.Show
    Case 2
      arBLDLQ2.Show
    Case 3
      arBLDLQ3.Show
    Case 4
      arBLFreeFormatDlnq.Show
    Case Else
      Close
      frmBLMessageBoxJr.Label1.Caption = "No delinquent notice has been selected."
      frmBLMessageBoxJr.Label1.Top = 900
      frmBLMessageBoxJr.Show vbModal
      Exit Sub
  End Select
  
  frmBLLoadReport.Show
  MainLog ("Delinquent notices (#" + QPTrim$(fptxtDlqNum.Text) + ") printed in graphics format.")
  
  Exit Sub
  
MakeCatLine:
  CatLine$ = ""
  C$ = QPTrim$(CustRec.BILLCAT1)
  If Len(C$) > 0 Then
    CatLine$ = CatLine$ + C$
  End If
  
  C$ = QPTrim$(CustRec.BILLCAT2)
  If Len(C$) > 0 Then
    If Len(CatLine$) > 0 Then
      CatLine$ = CatLine$ + ", " + C$
    Else
      CatLine$ = CatLine$ + C$
    End If
  End If
  C$ = QPTrim$(CustRec.BILLCAT3)
  If Len(C$) > 0 Then
    If Len(CatLine$) > 0 Then
      CatLine$ = CatLine$ + ", " + C$
    Else
      CatLine$ = CatLine$ + C$
    End If
  End If
  C$ = QPTrim$(CustRec.BILLCAT4)
  If Len(C$) > 0 Then
    If Len(CatLine$) > 0 Then
      CatLine$ = CatLine$ + ", " + C$
    Else
      CatLine$ = CatLine$ + C$
    End If
  End If
  C$ = QPTrim$(CustRec.BILLCAT5)
  If Len(C$) > 0 Then
    If Len(CatLine$) > 0 Then
      CatLine$ = CatLine$ + ", " + C$
    Else
      CatLine$ = CatLine$ + C$
    End If
  End If
  
  Return
  
'CarolinaBeach:
DLQType1:
  GoSub MakeCatLine
  
  '                           0
  Print #RptHandle, "***DELINQUENT NOTICE***"; dlm;
  '                     1
  Print #RptHandle, LNotDate$; dlm;
  '                               2
  Print #RptHandle, QPTrim$(TownRec.DlqTownName); dlm; '"TOWN OF CAROLINA BEACH"
  '                            3
  Print #RptHandle, QPTrim$(TownRec.DlqAdd1); dlm; ' "1121 NORTH LAKE PARK BLVD."
  '                            4
  Print #RptHandle, QPTrim$(TownRec.DlqCity) + ", " + QPTrim$(TownRec.DlqState) + " " + QPTrim$(TownRec.DlqZip); dlm; '"CAROLINA BEACH, N.C. 28428"
  '                            5
  Print #RptHandle, "TEL " + QPTrim$(TownRec.DlqPhone); dlm; '"TEL 910-458-2999"
  '                            6
  Print #RptHandle, "Account ID: " + CStr(CustomerNumber); dlm;
  '                            7
  Print #RptHandle, QPTrim$(CustRec.BillName); dlm;
  '                            8
  Print #RptHandle, QPTrim$(CustRec.ADDRESS1); dlm;
  '                            9
  Print #RptHandle, QPTrim$(CustRec.ADDRESS2); dlm;
  '                           10
  Print #RptHandle, QPTrim$(CustRec.City) + ", " + QPTrim$(CustRec.State) + " " + QPTrim$(CustRec.ZipCode); dlm;
  '                           11
  Print #RptHandle, "According   to   our   records,   your  " + XYear$ + "   Business   License   has   not   been"; dlm;
    '                           12
  Print #RptHandle, "purchased  as  of  today.  All  licenses  are  now  subject  to  a  penalty"; dlm;
  '                           13
  Print #RptHandle, "and   will   NOT   be   issued   unless   the   penalty   amount   is   included   with"; dlm;
  '                           14
  Print #RptHandle, "your   payment.   We   realize   that   you   are   very   busy,   but   we   would   like"; dlm;
  '                           15
  Print #RptHandle, "for   you   to   take   the   time   to   purchase   this   license."; dlm;
  '                           16
  Print #RptHandle, "We   show   your   account   delinquent   for   the   following"; dlm;
  '                           17
  Print #RptHandle, "license   code(s):"; dlm;
  If fpcmbFees.Text = "Yes" Then
    '                           18                             19
    Print #RptHandle, CustRec.DESC1; dlm; Using("$##,###,##0.00", CustRec.Fee1); dlm;
    '                           20                             21
    Print #RptHandle, CustRec.DESC2; dlm; Using("$##,###,##0.00", CustRec.Fee2); dlm;
    '                           22                             23
    Print #RptHandle, CustRec.DESC3; dlm; Using("$##,###,##0.00", CustRec.Fee3); dlm;
    '                           24                             25
    Print #RptHandle, CustRec.DESC4; dlm; Using("$##,###,##0.00", CustRec.Fee4); dlm;
    '                           26                             27
    Print #RptHandle, CustRec.DESC5; dlm; Using("$##,###,##0.00", CustRec.Fee5); dlm;
    
    SubTotal = CustRec.Fee1 + CustRec.Fee2 + CustRec.Fee3 + CustRec.Fee4 + CustRec.Fee5
    PctTotal = CustRec.PenBal
    GTotal = PctTotal + SubTotal
    UseThisPct = PctTotal
    '                           28                 29
    Print #RptHandle, "Penalty Charge"; dlm; CStr(UseThisPct); dlm;
    '                           30                                   31                     32                            33
    Print #RptHandle, "Total Business License Charges: "; dlm; CStr(GTotal); dlm; "Total Outstanding Balance"; dlm; CustRec.AcctBal; dlm;
  Else
    '                           18        19
    Print #RptHandle, CustRec.DESC1; dlm; ""; dlm;
    '                           20        21
    Print #RptHandle, CustRec.DESC2; dlm; ""; dlm;
    '                           22        23
    Print #RptHandle, CustRec.DESC3; dlm; ""; dlm;
    '                           24        25
    Print #RptHandle, CustRec.DESC4; dlm; ""; dlm;
    '                           26        27
    Print #RptHandle, CustRec.DESC5; dlm; ""; dlm;
    '                 28       29
    Print #RptHandle, ""; dlm; ""; dlm;
    '                 30       31       32       33
    Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
  End If
  '                           34
  Print #RptHandle, "Please   remit   your   payment   (including penalty)   to   this   office"; dlm;
  '                           35
  Print #RptHandle, "NO   later   than:   " + LPayDate$ + ".   If   you   have   questions"; dlm;
  '                           36
  Print #RptHandle, "regarding   your   license,   please   feel   free   to   contact   our   office."; dlm;
  '                           37
  Print #RptHandle, "If   payment   has   been   made   prior   to   receiving   this   notice,   please"; dlm;
  '                           38
  Print #RptHandle, "disregard   this   notice."; dlm;
  '                      39
  Print #RptHandle, "Sincerely,"; dlm;
  '                           40
  Print #RptHandle, QPTrim$(TownRec.DlqAdminName); dlm; '"DAWN S. JOHNSON"
  '                           41
  Print #RptHandle, QPTrim$(TownRec.DlqAdminTitle) ' "FINANCE DIRECTOR"
  
  DlqCnt = DlqCnt + 1
  
  Return
  
GetCode:
  For Snt& = 1 To NumOfARCatRecs
    Get CodeHandle, Snt&, CodeRec
    If QPTrim$(CodeRec.CatCode) = CatCode$ Then
      CODEDESC$ = QPTrim$(CodeRec.CODEDESC)
      Select Case CodeRec.CodeType
      Case "F"
        Amt# = CodeRec.Fee
        CodeType$ = CodeRec.CodeType
      Case "M"
        DESC1$ = "Per Each"
        Amt# = CodeRec.Fee
        CodeType$ = CodeRec.CodeType
      Case Is = "S"
        BaseAmt1# = CodeRec.BaseAmt1
        Revenue1# = CodeRec.Recpt1
        Percent1# = CodeRec.Percent1
        Maximum1# = CodeRec.Maximum1
        BaseAmt2# = CodeRec.BaseAmt2
        Revenue2# = CodeRec.Recpt2
        Percent2# = CodeRec.Percent2
        Maximum2# = CodeRec.Maximum2
        BaseAmt3# = CodeRec.BaseAmt3
        Revenue3# = CodeRec.Recpt3
        Percent3# = CodeRec.Percent3
        Maximum3# = CodeRec.Maximum3
        BaseAmt4# = CodeRec.BaseAmt4
        Revenue4# = CodeRec.Recpt4
        Percent4# = CodeRec.Percent4
        Maximum4# = CodeRec.Maximum4
        BaseAmt5# = CodeRec.BaseAmt5
        Revenue5# = CodeRec.Recpt5
        Percent5# = CodeRec.Percent5
        Maximum5# = CodeRec.Maximum5
        BaseAmt6# = CodeRec.BaseAmt6
        Revenue6# = CodeRec.Recpt6
        Percent6# = CodeRec.Percent6
        Maximum6# = CodeRec.Maximum6
        CodeType$ = CodeRec.CodeType
      Case Else
        CodeType$ = "N"
      End Select
      Exit For
    End If
  Next Snt&
  
GotCode:
  Return
  
'SunSetBeach:
DLQType2:
  '                            0
  Print #RptHandle, "***DELINQUENT NOTICE***"; dlm;
  '                            1
  Print #RptHandle, "ANNUAL BUSINESS LICENSE"; dlm;
  '                            2
  Print #RptHandle, QPTrim$(TownRec.AppTownOf); dlm; '"CITY OF ATMORE"
  '                            3
  Print #RptHandle, UCase(LNotDate$); dlm;
  '                            4
  Print #RptHandle, CustRec.BillName; dlm;
  '                            5
  Print #RptHandle, "BUSINESS ACCOUNT # " + CStr(IdxRecs(cnt)); dlm;
  '                            6
  Print #RptHandle, CustRec.ADDRESS1; dlm;
  '                            7
  Print #RptHandle, CustRec.ADDRESS2; dlm;
  '                            8
  Print #RptHandle, RTrim$(CustRec.City) + " " + RTrim$(CustRec.State) + " " + RTrim$(CustRec.ZipCode); dlm;
  '                            9
  Print #RptHandle, "ACCORDING   TO   OUR   RECORDS   YOUR   " + XYear + "   BUSINESS   LICENSE   HAS   NOT   BEEN   PURCHASED"; dlm;
  '                           10
  Print #RptHandle, "AS   OF   TODAY.   ALL   BUSINESS   LICENSE   FEES   ARE   NOW   SUBJECT   TO   A  PENALTY  CHARGE."; dlm;
  '                             11
  Print #RptHandle, "PLEASE   REMIT   YOUR   PAYMENT   (INCLUDING   THE   PENALTY)   NO   LATER   THAN "; dlm;
  '                             12
  Print #RptHandle, UCase(LPayDate$) + ".   IF   PAYMENT   HAS   ALREADY   BEEN   MADE   PRIOR   TO   THE "; dlm;
  '                             13
  Print #RptHandle, "DATE   ABOVE   THEN   PLEASE   DISREGARD   THIS   NOTICE   AND   THANK   YOU   FOR   YOUR   PAYMENT. "; dlm;
  '                             14
  Print #RptHandle, "Code         Type of License"; dlm;
      
'-----------------------------------------------------------
  CatCode$ = QPTrim$(CustRec.BILLCAT1)
  GoSub GetCode
  '                       15                    16               17                18
  Print #RptHandle, CustRec.BILLCAT1; dlm; CustRec.DESC1; dlm; "BASIS AMT"; dlm; "LICENSE AMT"; dlm;
  
  If CodeType$ = "S" Then
    '                    19                   20                 21                 22
    Print #RptHandle, "Min Due"; dlm; "For Recpts Up To"; dlm; "Plus"; dlm; "Of Recpts Over"; dlm;
    
    If BaseAmt1# > 0 Then
      '                    23             24                25                 26
      Print #RptHandle, BaseAmt1; dlm; Revenue1; dlm; Percent1# / 100; dlm; Maximum1; dlm;
    Else
      '                 23       24       25       26
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    End If
    
    If BaseAmt2# > 0 Then
      '                    27             28              29                  30
      Print #RptHandle, BaseAmt2; dlm; Revenue2; dlm; Percent2# / 100; dlm; Maximum2; dlm;
    Else
      '                 27       28       29       30
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    End If
    
    If BaseAmt3# > 0 Then
      '                    31             32                33                34
      Print #RptHandle, BaseAmt3; dlm; Revenue3; dlm; Percent3# / 100; dlm; Maximum3; dlm;
    Else
      '                 31       32       33       34
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    End If
    
    If BaseAmt4# > 0 Then
      '                    35             36                37                38
      Print #RptHandle, BaseAmt4; dlm; Revenue4; dlm; Percent4# / 100; dlm; Maximum4; dlm;
    Else
      '                 35       36       37       38
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    End If
    
    If BaseAmt5# > 0 Then
      '                    39             40                41                 42
      Print #RptHandle, BaseAmt5; dlm; Revenue5; dlm; Percent5# / 100; dlm; Maximum5; dlm;
    Else
      '                 39       40       41       42
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    End If
    
    If BaseAmt6# > 0 Then
      '                    43            44                 45                46
      Print #RptHandle, BaseAmt6; dlm; Revenue6; dlm; Percent6# / 100; dlm; Maximum6; dlm;
    Else
      '                 43       44       45       46
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    End If
  Else
    For x = 19 To 46
      Print #RptHandle, ""; dlm;
    Next x
  End If
  
  If CodeType$ = "F" Then
    '                     47             48
    Print #RptHandle, "Flat Fee: "; dlm; Amt#; dlm;
  Else
    '                 47       48
    Print #RptHandle, ""; dlm; ""; dlm;
  End If
  
  If CodeType$ = "M" Then
    '                       49                 50                 51
    Print #RptHandle, "Rate Per Unit: "; dlm; Amt#; dlm; "Times Number Of Units: "; dlm;
  Else
    '                 49       50       51
    Print #RptHandle, ""; dlm; ""; dlm; ""; dlm;
  End If

'-----------------------------------------------------------
  If Len(QPTrim$(CustRec.BILLCAT2)) = 0 Then
    For x = 52 To 199
      Print #RptHandle, ""; dlm;
    Next x
    GoTo EndAtmore1
  End If

  CatCode$ = QPTrim$(CustRec.BILLCAT2)
  GoSub GetCode
  '                      52                      53                 54                 55
  Print #RptHandle, CustRec.BILLCAT2; dlm; CustRec.DESC2; dlm; "BASIS AMT"; dlm; "LICENSE AMT"; dlm;
  If CodeType$ = "S" Then
    '                     56                  57                 58               59
    Print #RptHandle, "Min Due"; dlm; "For Recpts Up To"; dlm; "Plus"; dlm; "Of Recpts Over"; dlm;
    If BaseAmt1# > 0 Then
      '                    60              61                62                  63
      Print #RptHandle, BaseAmt1#; dlm; Revenue1#; dlm; Percent1# / 100; dlm; Maximum1#; dlm;
    Else
      '                 60       61       62       63
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    End If
    
    If BaseAmt2# > 0 Then
      '                     64              65                  66               67
      Print #RptHandle, BaseAmt2#; dlm; Revenue2#; dlm; Percent2# / 100; dlm; Maximum2#; dlm;
    Else
      '                 64       65       66       67
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    End If
    
    If BaseAmt3# > 0 Then
      '                    68              69                 70                 71
      Print #RptHandle, BaseAmt3#; dlm; Revenue3#; dlm; Percent3# / 100; dlm; Maximum3#; dlm;
    Else
      '                 68       69       70       71
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    End If
    
    If BaseAmt4# > 0 Then
      '                    72               73                74                 75
      Print #RptHandle, BaseAmt4#; dlm; Revenue4#; dlm; Percent4# / 100; dlm; Maximum4#; dlm;
    Else
      '                 72       73       74       75
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    End If
    
    If BaseAmt5# > 0 Then
      '                    76               77                  78                79
      Print #RptHandle, BaseAmt5#; dlm; Revenue5#; dlm; Percent5# / 100; dlm; Maximum5#; dlm;
    Else
      '                 76       77       78       79
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    End If
    
    If BaseAmt6# > 0 Then
      '                     80             81                   82                 83
      Print #RptHandle, BaseAmt6#; dlm; Revenue6#; dlm; Percent6# / 100; dlm; Maximum6#; dlm;
    Else
      '                 80       81       82       83
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    End If
  Else
    For x = 56 To 83
      Print #RptHandle, ""; dlm;
    Next x
  End If
  
  If CodeType$ = "F" Then
    '                    84               85
    Print #RptHandle, "Flat Fee: "; dlm; Amt#; dlm;
  Else
    '                 84       85
    Print #RptHandle, ""; dlm; ""; dlm;
  End If
  
  If CodeType$ = "M" Then
    '                       86                 87                88
    Print #RptHandle, "Rate Per Unit: "; dlm; Amt#; dlm; "Times Number Of Units: "; dlm;
  Else
    '                 86       87       88
    Print #RptHandle, ""; dlm; ""; dlm; ""; dlm;
  End If
'-----------------------------------------------------------
  If Len(QPTrim$(CustRec.BILLCAT3)) = 0 Then
    For x = 89 To 199
      Print #RptHandle, ""; dlm;
    Next x
    GoTo EndAtmore1
  End If
  
  CatCode$ = QPTrim$(CustRec.BILLCAT3)
  GoSub GetCode
  '                        89                   90                 91                92
  Print #RptHandle, CustRec.BILLCAT3; dlm; CustRec.DESC3; dlm; "BASIS AMT"; dlm; "LICENSE AMT"; dlm;
  If CodeType$ = "S" Then
    '                     93                  94                 95                96
    Print #RptHandle, "Min Due"; dlm; "For Recpts Up To"; dlm; "Plus"; dlm; "Of Recpts Over"; dlm;
    If BaseAmt1# > 0 Then
      '                    97              98                 99                 100
      Print #RptHandle, BaseAmt1#; dlm; Revenue1#; dlm; Percent1# / 100; dlm; Maximum1#; dlm;
    Else
      '                 97       98       99       100
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    End If
    
    If BaseAmt2# > 0 Then
      '                     101             102                 103                 104
      Print #RptHandle, BaseAmt2#; dlm; Revenue2#; dlm; Percent2# / 100; dlm; Maximum2#; dlm;
    Else
      '                 101      102      103      104
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    End If
    
    If BaseAmt3# > 0 Then
      '                    105             106                 107               108
      Print #RptHandle, BaseAmt3#; dlm; Revenue3#; dlm; Percent3# / 100; dlm; Maximum3#; dlm;
    Else
      '                 105      106      107      108
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    End If
    
    If BaseAmt4# > 0 Then
      '                    109             110               111                 112
      Print #RptHandle, BaseAmt4#; dlm; Revenue4#; dlm; Percent4# / 100; dlm; Maximum4#; dlm;
    Else
      '                    109             110               111                 112
      Print #RptHandle, BaseAmt4#; dlm; Revenue4#; dlm; Percent4# / 100; dlm; Maximum4#; dlm;
    End If
    
    If BaseAmt5# > 0 Then
      '                    113             114               115                 116
      Print #RptHandle, BaseAmt5#; dlm; Revenue5#; dlm; Percent5# / 100; dlm; Maximum5#; dlm;
    Else
      '                 113      114      115      116
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    End If
    
    If BaseAmt6# > 0 Then
      '                    117             118               119                 120
      Print #RptHandle, BaseAmt6#; dlm; Revenue6#; dlm; Percent6# / 100; dlm; Maximum6#; dlm;
    Else
      '                 117      118      119      120
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    End If
  Else
    For x = 93 To 120
      Print #RptHandle, ""; dlm;
    Next x
  End If
  
  If CodeType$ = "F" Then
    '                     121            122
    Print #RptHandle, "Flat Fee: "; dlm; Amt#; dlm;
  Else
    '                 121     122
    Print #RptHandle, ""; dlm; ""; dlm;
  End If
  
  If CodeType$ = "M" Then
    '                      123                124                 125
    Print #RptHandle, "Rate Per Unit: "; dlm; Amt#; dlm; "Times Number Of Units: "; dlm;
  Else
    '                 123      124     125
    Print #RptHandle, ""; dlm; ""; dlm; ""; dlm;
  End If

'-----------------------------------------------------------
  If Len(QPTrim$(CustRec.BILLCAT4)) = 0 Then
    For x = 126 To 199
      Print #RptHandle, ""; dlm;
    Next x
    GoTo EndAtmore1
  End If
  
  CatCode$ = QPTrim$(CustRec.BILLCAT4)
  GoSub GetCode
  '                       126                  127                 128                129
  Print #RptHandle, CustRec.BILLCAT4; dlm; CustRec.DESC4; dlm; "BASIS AMT"; dlm; "LICENSE AMT"; dlm;
  If CodeType$ = "S" Then
    '                    130                 131                132               133
    Print #RptHandle, "Min Due"; dlm; "For Recpts Up To"; dlm; "Plus"; dlm; "Of Recpts Over"; dlm;
    If BaseAmt1# > 0 Then
      '                    134             135               136                 137
      Print #RptHandle, BaseAmt1#; dlm; Revenue1#; dlm; Percent1# / 100; dlm; Maximum1#; dlm;
    Else
      '                 134      135      136      137
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    End If
    
    If BaseAmt2# > 0 Then
      '                    138             139               140                 141
      Print #RptHandle, BaseAmt2#; dlm; Revenue2#; dlm; Percent2# / 100; dlm; Maximum2#; dlm;
    Else
      '                 138      139      140      141
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    End If
    
    If BaseAmt3# > 0 Then
      '                    142             143                144                145
      Print #RptHandle, BaseAmt3#; dlm; Revenue3#; dlm; Percent3# / 100; dlm; Maximum3#; dlm;
    Else
      '                 142      143      144      145
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    End If
    
    If BaseAmt4# > 0 Then
      '                    146             147                148                149
      Print #RptHandle, BaseAmt4#; dlm; Revenue4#; dlm; Percent4# / 100; dlm; Maximum4#; dlm;
    Else
      '                 146      147      148      149
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    End If
    
    If BaseAmt5# > 0 Then
      '                    150             151                152                153
      Print #RptHandle, BaseAmt5#; dlm; Revenue5#; dlm; Percent5# / 100; dlm; Maximum5#; dlm;
    Else
      '                 150      151     152      153
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    End If
    
    If BaseAmt6# > 0 Then
      '                    154             155                156               157
      Print #RptHandle, BaseAmt6#; dlm; Revenue6#; dlm; Percent6# / 100; dlm; Maximum6#; dlm;
    Else
      '                 154      155      156      157
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    End If
  Else
    For x = 130 To 157
      Print #RptHandle, ""; dlm;
    Next x
  End If
  
  If CodeType$ = "F" Then
    '                     158            159
    Print #RptHandle, "Flat Fee: "; dlm; Amt#; dlm;
  Else
    '                 158      159
    Print #RptHandle, ""; dlm; ""; dlm;
  End If
  
  If CodeType$ = "M" Then
    '                       160               161                 162
    Print #RptHandle, "Rate Per Unit: "; dlm; Amt#; dlm; "Times Number Of Units: "; dlm;
  Else
    '                 160      161      162
    Print #RptHandle, ""; dlm; ""; dlm; ""; dlm;
  End If

'-----------------------------------------------------------
  If Len(QPTrim$(CustRec.BILLCAT5)) = 0 Then
    For x = 163 To 199
      Print #RptHandle, ""; dlm;
    Next x
    GoTo EndAtmore1
  End If
  
  CatCode$ = QPTrim$(CustRec.BILLCAT5)
  GoSub GetCode
  '                       163                   164               165                 166
  Print #RptHandle, CustRec.BILLCAT5; dlm; CustRec.DESC5; dlm; "BASIS AMT"; dlm; "LICENSE AMT"; dlm;
  
  If CodeType$ = "S" Then
    '                    167                 168                 169             170
    Print #RptHandle, "Min Due"; dlm; "For Recpts Up To"; dlm; "Plus"; dlm; "Of Recpts Over"; dlm;
    If BaseAmt1# > 0 Then
      '                    171             172                173                174
      Print #RptHandle, BaseAmt1#; dlm; Revenue1#; dlm; Percent1# / 100; dlm; Maximum1#; dlm;
    Else
      '                 171      172      173      174
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    End If
    
    If BaseAmt2# > 0 Then
      '                    175             176               177                 178
      Print #RptHandle, BaseAmt2#; dlm; Revenue2#; dlm; Percent2# / 100; dlm; Maximum2#; dlm;
    Else
      '                 175      176      177      178
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    End If
    
    If BaseAmt3# > 0 Then
      '                    179             180               181                 182
      Print #RptHandle, BaseAmt3#; dlm; Revenue3#; dlm; Percent3# / 100; dlm; Maximum3#; dlm;
    Else
      '                 179      180      181      182
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    End If
    
    If BaseAmt4# > 0 Then
      '                    183             184               185                 186
      Print #RptHandle, BaseAmt4#; dlm; Revenue4#; dlm; Percent4# / 100; dlm; Maximum4#; dlm;
    Else
      '                 183      184      185      186
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    End If
    
    If BaseAmt5# > 0 Then
      '                     187            188               189                 190
      Print #RptHandle, BaseAmt5#; dlm; Revenue5#; dlm; Percent5# / 100; dlm; Maximum5#; dlm;
    Else
      '                 187      188      189      190
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    End If
    
    If BaseAmt6# > 0 Then
      '                    191             192                193                194
      Print #RptHandle, BaseAmt6#; dlm; Revenue6#; dlm; Percent6# / 100; dlm; Maximum6#; dlm;
    Else
      '                 191      192      193      194
      Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
    End If
  Else
    For x = 167 To 194
      Print #RptHandle, ""; dlm;
    Next x
  End If
  
  If CodeType$ = "F" Then
    '                    195             196
    Print #RptHandle, "Flat Fee: "; dlm; Amt#; dlm;
  Else
    '                 195      196
    Print #RptHandle, ""; dlm; ""; dlm;
  End If
  
  If CodeType$ = "M" Then
    '                        197               198                  199
    Print #RptHandle, "Rate Per Unit: "; dlm; Amt#; dlm; "Times Number Of Units: "; dlm;
  Else
    '                 197      198      199
    Print #RptHandle, ""; dlm; ""; dlm; ""; dlm;
  End If

EndAtmore1:
  '                        200                            201
  Print #RptHandle, "MAKE CHECKS PAYABLE TO:"; dlm; "LICENSE  TOTAL: "; dlm;
  If PenAmt = True Then
    '                               202                   203
    Print #RptHandle, QPTrim$(TownRec.AppTownOf); dlm; "PENALTY:   +  "; dlm; '+ fpcurrFee.Text + " ="; dlm;
  Else
    '                               202                   203
    Print #RptHandle, QPTrim$(TownRec.AppTownOf); dlm; "PENALTY:"; dlm; '   TOTAL  * " + Using("##0.00%", Val(fptxtPenPct.Text) / 100) + " ="; dlm;
  End If
  '                           204                      205
  Print #RptHandle, QPTrim$(TownRec.AppAdd1); dlm; "TOTAL  LICENSE  FEES: "; dlm;
  '                              206
  Print #RptHandle, QPTrim$(TownRec.AppCity) + ", " + QPTrim$(TownRec.AppState) + " " + QPTrim$(TownRec.AppZip); dlm;
  '                              207
  Print #RptHandle, "WHERE  APPLICABLE,  ESTABLISHMENTS  NOT  PURCHASING  A  LICENSE  BY  " + fptxtPayByDate.Text; dlm;
  '                              208
  Print #RptHandle, "WILL  BE  REPORTED  TO  THE  ABC  COMMISSION."; dlm;
  '                             209
  Print #RptHandle, "RENEWALS LICENSE VALID UNTIL " + GoodThru + "."; dlm;
  '                             210
  Print #RptHandle, "PLEASE CONTACT THE TOWN OFFICE WITH ANY QUESTIONS."; dlm;
  '                             211
  If QPTrim$(TownRec.DlqPhone2) = "(" Then
    Print #RptHandle, "TELEPHONE:  " + QPTrim$(TownRec.DlqPhone); dlm;
  Else
    Print #RptHandle, "TELEPHONE:  " + QPTrim$(TownRec.DlqPhone) + " OR " + QPTrim$(TownRec.DlqPhone2); dlm;
  End If
  '                             212
  If QPTrim$(TownRec.DlqFax) = "(" Then
    Print #RptHandle, ""; dlm;
  Else
    Print #RptHandle, "FAX:  " + QPTrim$(TownRec.DlqFax); dlm;
  End If
  
  DlqCnt = DlqCnt + 1
  If fpcmbFees.Text = "Yes" Then
    SubTotal = CustRec.Fee1 + CustRec.Fee2 + CustRec.Fee3 + CustRec.Fee4 + CustRec.Fee5
    PctTotal = CustRec.PenBal
    GTotal = SubTotal + PctTotal
    '                      213               214                 215                216                217
    Print #RptHandle, CustRec.Fee1; dlm; CustRec.Fee2; dlm; CustRec.Fee3; dlm; CustRec.Fee4; dlm; CustRec.Fee5; dlm;
    '                    218           219            220            221                 222              223                224                 225                 226
    Print #RptHandle, SubTotal; dlm; PctTotal; dlm; GTotal; dlm; CustRec.REV1; dlm; CustRec.REV2; dlm; CustRec.REV3; dlm; CustRec.REV4; dlm; CustRec.REV5; dlm; CustRec.AcctBal; dlm; "TOTAL  BALANCE:"
  Else
    '                 213      214      215      216      217      218      219      220      221      222      223      224      225      226
    Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""; dlm; ""
  End If
  
  UseThisPct = PctTotal
  
  Return
'-----------------------------------------------------

DLQType3:
  '             0
  Print #RptHandle, QPTrim$(TownRec.DlqTownName); dlm; '"TOWN OF EXMORE"
  '             1
  Print #RptHandle, LNotDate$; dlm; 'LNDTab
  '             2
  Print #RptHandle, QPTrim$(CustRec.BillName); dlm;
  '             3
  Print #RptHandle, QPTrim$(CustRec.ADDRESS1); dlm;
  '             4
  Print #RptHandle, QPTrim$(CustRec.ADDRESS2); dlm;
  '             5
  Print #RptHandle, QPTrim$(CustRec.City) + ", " + QPTrim$(CustRec.State) + " " + QPTrim$(CustRec.ZipCode); dlm;
  '             6
  Print #RptHandle, "Dear  Business  Owner:"; dlm;
  '             7
  Print #RptHandle, "       According  to  our  records,  your  APPLICATION  FOR  TOWN  LICENSE(S)  has  not"; dlm;
  '             8
  Print #RptHandle, "yet   been   submitted   for   processing   your  " + XYear + "  BUSINESS,  PROFESSIONAL  and"; dlm;
  '             9
  Print #RptHandle, "OCCUPATIONAL   LICENSE  (BPOL)   TAX.    This   application   is   required  prior  to"; dlm;
  '            10
  Print #RptHandle, "issuance   of   any   " + QPTrim$(TownRec.DlqCity) + "   Business   License.   The   application   form"; dlm;
  '            11
  Print #RptHandle, "specifies    a    " + fptxtXDate.Text + "    deadline    for    filing,    and    states    that    a    penalty "; dlm;
  '            12
  Print #RptHandle, "may   be   assessed   on   delinquent  applications.  The   application  also  states"; dlm;
  '            13
  Print #RptHandle, "a   deadline   of   " + LPayDate$ + "   for   payment   of   applicable   BPOL   Tax,"; dlm;
  '            14
  Print #RptHandle, "as    stated    in    Code    of    Virginia    58.1-3703.1-Uniform    Ordinance    Provisions."; dlm;
  '            15
  Print #RptHandle, "To   avoid   further   action,   please   complete   and   return   the   APPLICATION   FOR"; dlm;
  '            16
  Print #RptHandle, "TOWN    LICENSE(S)    immediately.   Failure    to    comply    may    result    in    legal"; dlm;
  '            17
  Print #RptHandle, "action    including,    but    not    limited    to,    an    audit    of    business    records,    as"; dlm;
  '            18
  Print #RptHandle, "permitted    in    Code    of    Virginia  58.1-3110  and  58.1-3939.1"; dlm;
  '            19
  Print #RptHandle, "       If    there    are    any    questions   or   if   assistance   is   needed   in   completing   the"; dlm;
  '            20
  Print #RptHandle, "form,      please      call      " + QPTrim$(TownRec.DlqClerkName) + "    at    " + QPTrim$(TownRec.DlqPhone) + ",   " + QPTrim(TownRec.DlqFirstDay) + "    to    " + QPTrim$(TownRec.DlqLastDay); dlm;   '  Monday - Friday 8:00 A.M."
  '            21
  Print #RptHandle, "from   " + QPTrim$(TownRec.DlqFirstHour) + "   to   " + QPTrim$(TownRec.DlqLastHour) + "."; dlm;
  '            22
  Print #RptHandle, "Cordially,"; dlm;
  '            23
  Print #RptHandle, QPTrim$(TownRec.DlqMayorCouncil); dlm; '"Mayor and Council"
  '            24
  Print #RptHandle, QPTrim$(TownRec.DlqTownName); dlm; '"Town Of Exmore"
  '            25
  Print #RptHandle, QPTrim$(TownRec.DlqAdminName); dlm; ' "Donald P. Kellam, Sr."
  '            26
  Print #RptHandle, QPTrim$(TownRec.DlqAdminTitle) ' "BPOL Commissioner"

  DlqCnt = DlqCnt + 1

  Return
'-----------------------------------------------------

DLQType4:
  AddEmptyFields = 0
  '                              0                                2
  Print #RptHandle, QPTrim$(LaserRec4.Line1(0)); dlm; QPTrim$(LaserRec4.Line1(1)); dlm;
  '                              2                                3
  Print #RptHandle, QPTrim$(LaserRec4.Line1(2)); dlm; QPTrim$(LaserRec4.Line1(3)); dlm;
  '                              4                                5
  Print #RptHandle, QPTrim$(LaserRec4.Phone); dlm; QPTrim$(fptxtNoticeDate.Text); dlm;
  '                              6                                7
  Print #RptHandle, LaserRec4.Line2(0); dlm; LaserRec4.Line2(1); dlm;
  '                              8                                9
  Print #RptHandle, LaserRec4.Line2(2); dlm; LaserRec4.Line2(3); dlm;
  '                              10                               11
  Print #RptHandle, LaserRec4.Line2(4); dlm; LaserRec4.Line2(5); dlm;
  '                              12                               13
  Print #RptHandle, LaserRec4.Line2(6); dlm; LaserRec4.Line2(7); dlm;
  '                              14                               15
  Print #RptHandle, QPTrim$(LaserRec4.Signer); dlm; QPTrim$(CustRec.BillName); dlm;
  '                              16                               17
  Print #RptHandle, QPTrim$(CustRec.CustNumb); dlm; QPTrim$(CustRec.ADDRESS1); dlm;
  '                              18
  Print #RptHandle, QPTrim$(CustRec.ADDRESS2); dlm;
  
  If QPTrim$(CustRec.BILLCAT1) <> "" Then
    If fpcmbFees.Text = "Yes" Then
      '                              19                               20                     21
      Print #RptHandle, QPTrim$(CustRec.BILLCAT1); dlm; QPTrim$(CustRec.DESC1); dlm; CustRec.FeeLicBal1; dlm;
    Else
      Print #RptHandle, QPTrim$(CustRec.BILLCAT1); dlm; QPTrim$(CustRec.DESC1); dlm; ""; dlm;
    End If
  Else
      AddEmptyFields = AddEmptyFields + 3
  End If
  
  If QPTrim$(CustRec.BILLCAT2) <> "" Then
    If fpcmbFees.Text = "Yes" Then
      '                              22                               23                      24
      Print #RptHandle, QPTrim$(CustRec.BILLCAT2); dlm; QPTrim$(CustRec.DESC2); dlm; CustRec.FeeLicBal2; dlm;
    Else
      Print #RptHandle, QPTrim$(CustRec.BILLCAT2); dlm; QPTrim$(CustRec.DESC2); dlm; ""; dlm;
    End If
  Else
      AddEmptyFields = AddEmptyFields + 3
  End If
    
  If QPTrim$(CustRec.BILLCAT3) <> "" Then
    If fpcmbFees.Text = "Yes" Then
      '                              25                               26                      27
      Print #RptHandle, QPTrim$(CustRec.BILLCAT3); dlm; QPTrim$(CustRec.DESC3); dlm; CustRec.FeeLicBal3; dlm;
    Else
      Print #RptHandle, QPTrim$(CustRec.BILLCAT3); dlm; QPTrim$(CustRec.DESC3); dlm; ""; dlm;
    End If
  Else
      AddEmptyFields = AddEmptyFields + 3
  End If
  
  If QPTrim$(CustRec.BILLCAT4) <> "" Then
    If fpcmbFees.Text = "Yes" Then
      '                              28                               29                      30
      Print #RptHandle, QPTrim$(CustRec.BILLCAT4); dlm; QPTrim$(CustRec.DESC4); dlm; CustRec.FeeLicBal4; dlm;
    Else
      Print #RptHandle, QPTrim$(CustRec.BILLCAT4); dlm; QPTrim$(CustRec.DESC4); dlm; ""; dlm;
    End If
  Else
      AddEmptyFields = AddEmptyFields + 3
  End If
  
  If QPTrim$(CustRec.BILLCAT5) <> "" Then
    If fpcmbFees.Text = "Yes" Then
      '                              31                               32                     33
      Print #RptHandle, QPTrim$(CustRec.BILLCAT5); dlm; QPTrim$(CustRec.DESC5); dlm; CustRec.FeeLicBal5; dlm;
    Else
      Print #RptHandle, QPTrim$(CustRec.BILLCAT5); dlm; QPTrim$(CustRec.DESC5); dlm; ""; dlm;
    End If
  Else
      AddEmptyFields = AddEmptyFields + 3
  End If
  
  For y = 1 To AddEmptyFields
    '
    Print #RptHandle, ""; dlm;
  Next y
  
  If fpcmbFees.Text = "Yes" Then
    '                       34                     35                     36                   37
    Print #RptHandle, CustRec.FeeBal; dlm; CustRec.IssuanceBal; dlm; CustRec.PenBal; dlm; CustRec.AcctBal; dlm;
  Else
    '                 34       35       36       37
    Print #RptHandle, ""; dlm; ""; dlm; ""; dlm; ""; dlm;
  End If
  
  '                           38
  Print #RptHandle, QPTrim$(CustRec.City) + ", " + QPTrim$(CustRec.State) + " " + QPTrim$(CustRec.ZipCode); dlm;
  '                           39
  Print #RptHandle, QPTrim$(fpcmbFees.Text)
  
  DlqCnt = DlqCnt + 1
  
  Return

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLDelinquentNotices", "PrintGraphics", Erl)
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

Private Sub fpcmbRange_Change()
  If QPTrim$(fpcmbRange.Text) = "" Then
    fpcmbRange.Text = "Up To And Include This Expiration"
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
      If fpcmbFees.Enabled = True Then
        fpcmbFees.SetFocus
      Else
        fpcmbPrintOpt.SetFocus
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

Private Sub fptxtXDate_Change()
  XYear = Mid(fptxtXDate.Text, 7, 4)
End Sub

