VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLDlqntMailLbls 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Mailing Labels"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmBLDlqntMailLbls.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   6084
      Left            =   1968
      TabIndex        =   4
      Top             =   1392
      Width           =   7788
      _Version        =   196609
      _ExtentX        =   13737
      _ExtentY        =   10731
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmBLDlqntMailLbls.frx":08CA
      Begin LpLib.fpCombo fpcmbRange 
         Height          =   405
         Left            =   2445
         TabIndex        =   2
         Tag             =   $"frmBLDlqntMailLbls.frx":08E6
         Top             =   3120
         Width           =   4380
         _Version        =   196608
         _ExtentX        =   7726
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
         ColDesigner     =   "frmBLDlqntMailLbls.frx":0A97
      End
      Begin LpLib.fpCombo fpcmbLabel 
         Height          =   405
         Left            =   2445
         TabIndex        =   1
         Tag             =   $"frmBLDlqntMailLbls.frx":0D92
         Top             =   2445
         Width           =   4380
         _Version        =   196608
         _ExtentX        =   7726
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
         ColDesigner     =   "frmBLDlqntMailLbls.frx":0E58
      End
      Begin LpLib.fpCombo fpcmbPrintOrder 
         Height          =   405
         Left            =   2880
         TabIndex        =   0
         Tag             =   $"frmBLDlqntMailLbls.frx":1153
         Top             =   1770
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
         ColDesigner     =   "frmBLDlqntMailLbls.frx":11FF
      End
      Begin EditLib.fpDateTime fptxtXDate 
         Height          =   396
         Left            =   2976
         TabIndex        =   3
         Tag             =   $"frmBLDlqntMailLbls.frx":14FA
         Top             =   3792
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
         Text            =   "05/08/2003"
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
         Height          =   480
         Left            =   2550
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   $"frmBLDlqntMailLbls.frx":1725
         Top             =   4950
         Width           =   1350
         _Version        =   131072
         _ExtentX        =   2381
         _ExtentY        =   847
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
         ButtonDesigner  =   "frmBLDlqntMailLbls.frx":1804
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   480
         Left            =   3990
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "Press 'Cancel' to exit this screen and return to the 'Penalty Processing' menu."
         Top             =   4950
         Width           =   1635
         _Version        =   131072
         _ExtentX        =   2884
         _ExtentY        =   847
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
         ButtonDesigner  =   "frmBLDlqntMailLbls.frx":19E0
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   480
         Left            =   5715
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   $"frmBLDlqntMailLbls.frx":1BBE
         Top             =   4950
         Width           =   1695
         _Version        =   131072
         _ExtentX        =   2990
         _ExtentY        =   847
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
         ButtonDesigner  =   "frmBLDlqntMailLbls.frx":1C4F
      End
      Begin fpBtnAtlLibCtl.fpBtn fpcmdXList 
         Height          =   396
         Left            =   4800
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   $"frmBLDlqntMailLbls.frx":1E2E
         Top             =   3792
         Width           =   1932
         _Version        =   131072
         _ExtentX        =   3408
         _ExtentY        =   698
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
         ButtonDesigner  =   "frmBLDlqntMailLbls.frx":1F19
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdHelp 
         Height          =   480
         Left            =   435
         TabIndex        =   14
         TabStop         =   0   'False
         Tag             =   $"frmBLDlqntMailLbls.frx":20FF
         Top             =   4950
         Width           =   2025
         _Version        =   131072
         _ExtentX        =   3572
         _ExtentY        =   847
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
         ButtonDesigner  =   "frmBLDlqntMailLbls.frx":219C
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "For Customers With Balances"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Left            =   2010
         TabIndex        =   17
         Top             =   840
         Width           =   3945
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
         Left            =   384
         TabIndex        =   15
         Top             =   5472
         Width           =   2100
      End
      Begin VB.Label Label3 
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
         Left            =   912
         TabIndex        =   13
         Top             =   3216
         Width           =   1356
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
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
         Height          =   348
         Left            =   1104
         TabIndex        =   11
         Top             =   3888
         Width           =   1788
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
         Caption         =   "Delinquent Mailing Labels"
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
         Height          =   390
         Left            =   2010
         TabIndex        =   10
         Top             =   480
         Width           =   3945
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
         Left            =   1392
         TabIndex        =   9
         Top             =   1872
         Width           =   1308
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
         Left            =   912
         TabIndex        =   8
         Top             =   2544
         Width           =   1356
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   3180
         Left            =   720
         Top             =   1392
         Width           =   6396
      End
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   444
      Left            =   1920
      TabIndex        =   16
      Top             =   7728
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
      Height          =   6348
      Left            =   1800
      Top             =   1260
      Width           =   8052
   End
End
Attribute VB_Name = "frmBLDlqntMailLbls"
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
  
  On Error Resume Next
  
  cmdHelp.Text = "F1 &Turn Help Off"
  btnHelp.AutoScan = fpAutoScanPopupOnly
  lblBalloon.Visible = True
  
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
    frmBLMessageBoxJr.Label1.Top = 800
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
  MainLog ("Delinquent mailing labels allignment feature used.")
  
End Sub

Private Sub cmdExit_Click()
  KillFile "dlnqmllbls.dat"
  frmBLPenProcMenu.Show
  DoEvents
  Unload frmBLDlqntMailLbls
End Sub

Private Sub cmdHelp_Click()
  If InStr(cmdHelp.Text, "On") Then
    cmdHelp.Text = "F1 &Turn Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    lblBalloon.Visible = True
    cmdHelp.ToolTipText = ""
    fpcmbPrintOrder.ToolTipText = ""
    fpcmbLabel.ToolTipText = ""
    fpcmbRange.ToolTipText = ""
    fptxtXDate.ToolTipText = ""
    fpcmdXList.ToolTipText = ""
    cmdAlign.ToolTipText = ""
    cmdExit.ToolTipText = ""
    cmdProcess.ToolTipText = ""
  ElseIf InStr(cmdHelp.Text, "Off") Then
    cmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
'    fpcmdHelp.ToolTipText = "Press 'Turn Help On' to activate instructional balloons that will appear when you place the cursor over any field on the screen. Press 'Turn Help Off' to deactivate the instructional balloons."
'    fpcmbPrintOrder.ToolTipText = "Mailing labels can be printed in alphabetical order or in numerical order."
'    fpcmbLabel.ToolTipText = "Select the kind of mailing labels to print from the drop down list."
'    fpcmbRange.ToolTipText = "You can elect to print all those delinquent up to and including the expiration date or just those delinquent on the date entered."
'    fptxtXDate.ToolTipText = "Only customers whose business licenses expire on or before this date will have a mailing label printed."
'    fpcmdXList.ToolTipText = "Press for a concise explanation of the details of this screen."
'    cmdAlign.ToolTipText = "Press to print out a template from which you can check your mailing label alignment."
'    cmdExit.ToolTipText = "Press to exit this screen."
'    cmdProcess.ToolTipText = "Press to begin the mailing label printing process."
  
  
  End If
End Sub

Private Sub cmdProcess_Click()
  If InStr(fpcmbLabel.Text, "Graphical") Then
    Call PrintGraphics
  ElseIf InStr(fpcmbLabel.Text, "Text") Then
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
  Dim CustSrchIdxRec As CustNameIdxType ' CustSearchNameIdxType
  Dim NumOfCustIdx As Integer
  Dim x As Integer, y As Integer
  Dim CustRec As ARCustRecType
  Dim CustRCnt As Integer
  Dim CustHandle As Integer
  Dim Zip$
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
  Dim ExpDate As Integer
  Dim Nextx As Integer
  Dim NumOfCustRecs As Integer
  Dim DlnqntCnt As Integer
  Dim ValidCnt As Integer
  
  On Error GoTo ERRORSTUFF
  
  cmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  
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
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    fpcmbLabel.BackColor = &HFFFFFF
    fpcmbLabel.SetFocus
    Exit Sub
  End If
  
  If QPTrim$(fptxtXDate.Text) = "" Then
    fptxtXDate.BackColor = &HFFFF&
    frmBLMessageBoxJr.Label1.Caption = "Please enter a valid expiration date."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    fptxtXDate.BackColor = &HFFFFFF
    fptxtXDate.SetFocus
    Exit Sub
  End If
  
  ExpDate = Date2Num(fptxtXDate.Text)
  
  ReportFile$ = "ARDLQBL.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  If fpcmbPrintOrder.Text = "Billing Name Order" Then
    NameFlag = True
'    OpenSrchNameIdxFile CustIdxHandle
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
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    fpcmbPrintOrder.BackColor = &HFFFFFF
    fpcmbPrintOrder.SetFocus
    Close
    Exit Sub
  End If

  OpenCustFile CustHandle
  NumOfCustRecs = LOF(CustHandle) / Len(CustRec)
  ReDim ToPrint(1 To 5, 1 To 5) As String
  For cnt = 1 To NumOfCustIdx 'NumOfCustRecs
    Get CustHandle, IdxRec(cnt), CustRec
    If CustRec.AcctBal <= 0 Then GoTo NotThisOne
    If QPTrim$(CustRec.Inactive) = "Y" Or QPTrim$(CustRec.Deleted) = "Y" Or QPTrim$(CustRec.SortName) = "DELETED" Then GoTo NotThisOne
    If InStr(fpcmbRange.Text, "Only") Then
      If CustRec.VALID <> ExpDate Then GoTo NotThisOne
    Else
      If CustRec.VALID > ExpDate Then GoTo NotThisOne
    End If
    GoSub PastDueCust
NotThisOne:
  Next cnt
  
  
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
    frmBLMessageBoxJr.Label1.Caption = "There are no customers that fall within the parameters as set on the screen. No labels printed."
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
  MainLog ("Delinquent mailing labels processed in text format.")
  
  Exit Sub
  
PastDueCust:
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
  Return

ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLDlqntMailLbls", "PrintText", Erl)
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
    Case vbKeyF5:
      SendKeys "%A"
      Call cmdAlign_Click
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
      KillFile "dlnqmllbls.dat"
      ClearInUse PWcnt
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLDlqntMailLbls.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  Dim One As Integer
  Dim DHandle As Integer
  
  On Error Resume Next
  
  cmdAlign.Enabled = False
  lblBalloon.Visible = False
'  fpcmbPrintOrder.ToolTipText = "Mailing labels can be printed in alphabetical order or in numerical order."
'  fpcmbLabel.ToolTipText = "Select the kind of mailing labels to print from the drop down list."
'  fpcmbRange.ToolTipText = "You can elect to print all those delinquent up to and including the expiration date or just those delinquent on the date entered."
'  fptxtXDate.ToolTipText = "Only customers whose business licenses expire on or before this date will have a mailing label printed."
'  fpcmdXList.ToolTipText = "Press for a concise explanation of the details of this screen."
'  cmdAlign.ToolTipText = "Press to print out a template from which you can check your mailing label alignment."
'  cmdExit.ToolTipText = "Press to exit this screen."
'  cmdProcess.ToolTipText = "Press to begin the mailing label printing process."
'  fpcmdHelp.ToolTipText = "Press 'Turn Help On' to activate instructional balloons that will appear when you place the cursor over any field on the screen. Press 'Turn Help Off' to deactivate the instructional balloons."
  One = 1
  DHandle = FreeFile
  Open "dlnqmllbls.dat" For Output As DHandle Len = 2
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
  fpcmbRange.Text = "Up To And Include This Expiration"
  fpcmbRange.AddItem "Up To And Include This Expiration"
  fpcmbRange.AddItem "This Expiration Only"
  fptxtXDate.Text = Date
  cmdAlign.Enabled = False
End Sub

Private Sub fpcmbLabel_Change()
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
      fpcmbRange.SetFocus
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
      If fptxtXDate.Enabled = True Then
        fptxtXDate.SetFocus
      Else
        fpcmbPrintOrder.SetFocus
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
  Dim CustSrchIdxRec As CustNameIdxType ' CustSearchNameIdxType
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
  Dim XDate As Integer
  
  On Error GoTo ERRORSTUFF
  
  XDate = Date2Num(fptxtXDate.Text)
  
  dlm = "~"
  NameFlag = False
  
  OpenCustFile ARFile
  If fpcmbPrintOrder.Text = "Billing Name Order" Then
    NameFlag = True
'    OpenSrchNameIdxFile CustIdxHandle
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
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    fpcmbPrintOrder.BackColor = &HFFFFFF
    fpcmbPrintOrder.SetFocus
    Close
    Exit Sub
  End If
  
  ReportFile$ = "BLRPTS\ARLABEL.RPT"
  RptHandle = FreeFile
  Open ReportFile$ For Output As #RptHandle
  
  ReDim ToPrint(1 To 5, 1 To 5) As String
  
  For cnt = 1 To NumOfCustIdx
    Get ARFile, IdxRec(cnt), CustRec
    If CustRec.AcctBal <= 0 Then GoTo NextLabel
    If QPTrim$(CustRec.Inactive) = "Y" Then GoTo NextLabel
    If UCase$(CustRec.Deleted) = "Y" Then
      GoTo NextLabel
    End If
    If InStr(fpcmbRange.Text, "Only") Then
      If CustRec.VALID <> XDate Then GoTo NextLabel
    Else
      If CustRec.VALID > XDate Then GoTo NextLabel
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
  
  If CustPCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "There are no customers that fall within the parameters as set on the screen. No labels printed."
    frmBLMessageBoxJr.Label1.Top = 700
    frmBLMessageBoxJr.Show vbModal
    Close
    Exit Sub
  End If
  
  arBLMailLabels.Show
  frmBLLoadReport.Show
  
  MainLog ("Delinquent mailing labels processed in graphics format.")
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
