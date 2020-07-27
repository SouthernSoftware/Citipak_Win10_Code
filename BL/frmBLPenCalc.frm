VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmBLPenCalc 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business License Penalty Calculation"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   Icon            =   "frmBLPenCalc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   6090
      Left            =   1965
      TabIndex        =   8
      Top             =   1635
      Width           =   7785
      _Version        =   196609
      _ExtentX        =   13732
      _ExtentY        =   10742
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmBLPenCalc.frx":08CA
      Begin LpLib.fpCombo fpcmbRange 
         Height          =   375
         Left            =   3210
         TabIndex        =   5
         Tag             =   $"frmBLPenCalc.frx":08E6
         Top             =   4035
         Width           =   3720
         _Version        =   196608
         _ExtentX        =   6562
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
         ColDesigner     =   "frmBLPenCalc.frx":0AA1
      End
      Begin LpLib.fpCombo fpcmbBalType 
         Height          =   405
         Left            =   3075
         TabIndex        =   0
         Tag             =   $"frmBLPenCalc.frx":0DD0
         Top             =   1515
         Width           =   3555
         _Version        =   196608
         _ExtentX        =   6271
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
         ColDesigner     =   "frmBLPenCalc.frx":0FA1
      End
      Begin EditLib.fpCurrency fpcurrFee 
         Height          =   345
         Left            =   4035
         TabIndex        =   2
         Tag             =   $"frmBLPenCalc.frx":12D0
         Top             =   2190
         Width           =   1650
         _Version        =   196608
         _ExtentX        =   2900
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
         Height          =   495
         Left            =   3075
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "Press 'Cancel' to exit this screen and return to the 'Penalty Processing' menu."
         Top             =   4920
         Width           =   1875
         _Version        =   131072
         _ExtentX        =   3307
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
         ButtonDesigner  =   "frmBLPenCalc.frx":138C
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   495
         Left            =   5370
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   $"frmBLPenCalc.frx":156A
         Top             =   4920
         Width           =   1890
         _Version        =   131072
         _ExtentX        =   3334
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
         ButtonDesigner  =   "frmBLPenCalc.frx":17BC
      End
      Begin EditLib.fpText fptxtPct 
         Height          =   390
         Left            =   4035
         TabIndex        =   1
         Tag             =   $"frmBLPenCalc.frx":199B
         Top             =   2130
         Width           =   1650
         _Version        =   196608
         _ExtentX        =   2900
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
         BackColor       =   16777215
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
         CharValidationText=   "0 1 2 3 4 5 6 7 8 9 ."
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
      Begin EditLib.fpDateTime fptxtXDate 
         Height          =   390
         Left            =   2970
         TabIndex        =   4
         Tag             =   $"frmBLPenCalc.frx":1AD4
         Top             =   3360
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
      Begin fpBtnAtlLibCtl.fpBtn fpcmdXList 
         Height          =   390
         Left            =   4845
         TabIndex        =   13
         TabStop         =   0   'False
         Tag             =   $"frmBLPenCalc.frx":1BB1
         Top             =   3360
         Width           =   1935
         _Version        =   131072
         _ExtentX        =   3413
         _ExtentY        =   688
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
         ButtonDesigner  =   "frmBLPenCalc.frx":1C9A
      End
      Begin fpBtnAtlLibCtl.fpBtn fpcmdHelp 
         Height          =   495
         Left            =   630
         TabIndex        =   15
         TabStop         =   0   'False
         Tag             =   $"frmBLPenCalc.frx":1E80
         Top             =   4920
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
         ButtonDesigner  =   "frmBLPenCalc.frx":1F1D
      End
      Begin EditLib.fpDateTime fptxtTransDate 
         Height          =   390
         Left            =   4080
         TabIndex        =   3
         Tag             =   "In the 'Transaction Date' enter the date you wish for this penalty calculation operation to be conducted."
         Top             =   2760
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
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Date:"
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
         Left            =   1800
         TabIndex        =   18
         Top             =   2880
         Width           =   2190
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
         Left            =   570
         TabIndex        =   16
         Top             =   5445
         Width           =   2100
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Calculation Range:"
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
         Left            =   1050
         TabIndex        =   14
         Top             =   4080
         Width           =   2025
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Past Due Date:"
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
         Left            =   1005
         TabIndex        =   12
         Top             =   3450
         Width           =   1830
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Percent:"
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
         Left            =   2010
         TabIndex        =   11
         Top             =   2235
         Width           =   1830
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   684
         Left            =   1488
         Top             =   288
         Width           =   4908
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Penalty Calculations"
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
         Left            =   1824
         TabIndex        =   10
         Top             =   432
         Width           =   4332
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Balance Type:"
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
         Left            =   1350
         TabIndex        =   9
         Top             =   1605
         Width           =   1545
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   3495
         Left            =   870
         Top             =   1230
         Width           =   6210
      End
   End
   Begin fpBtnAtlLibCtl.fpBln btnHelp 
      Height          =   450
      Left            =   1920
      TabIndex        =   17
      Top             =   8040
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
      ShapeRoundWidth =   180
      ShapeRoundHeight=   180
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
      Height          =   6390
      Left            =   1770
      Top             =   1470
      Width           =   8100
   End
End
Attribute VB_Name = "frmBLPenCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsBLTextBoxOverrider
  Private Temp_Class As Resize_Class
  Dim PctFlag As Boolean

Private Sub cmdProcess_Click()
  Dim TBalFlag As Integer
  Dim Amt#, cnt As Integer
  Dim TransRec As TempPenaltyCharges
  Dim THandle As Integer
  Dim NumOfTransRecs As Double
  Dim CustRec As ARCustRecType
  Dim CHandle As Integer
  Dim NumOfCustRecs As Integer
  Dim CustBal#, CustPen#
  Dim PenCnt As Integer
  Dim PenAsTotal As Boolean
  Dim ThisAmt As Double
  Dim x As Integer, Nextx As Integer
  Dim ThisDate As Integer
  Dim RangeFlag As Integer
  Dim y As Integer
  
  On Error GoTo ERRORSTUFF
  
  fpcmdHelp.Text = "F1 &Turn Help On"
  btnHelp.AutoScan = fpAutoScanOff
  lblBalloon.Visible = False
  
  If Not Exist("artownsu.dat") Then
    frmBLMessageBoxJr.Label1.Caption = "Please complete the Town Setup data before continuing."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    Exit Sub
  End If

  If Exist("artmppen.dat") Then
    frmBLMessageBoxJrWOpts.Label1.Caption = "A penalty processing file already exists. Do you wish to overwrite?"
    frmBLMessageBoxJrWOpts.Label1.Top = 1000
    frmBLMessageBoxJrWOpts.Show vbModal
    If frmBLMessageBoxJrWOpts.fptxtChoice.Text = "abort" Then
      Unload frmBLMessageBoxJrWOpts
      Exit Sub
    Else
      Unload frmBLMessageBoxJrWOpts
    End If
  End If
  
  If QPTrim$(fpcmbBalType.Text) = "Total Balance" Then
    TBalFlag = 1
  ElseIf QPTrim$(fpcmbBalType.Text) = "License Only" Then
    TBalFlag = 2
  ElseIf QPTrim$(fpcmbBalType.Text) = "Using Fixed Amount" Then
    TBalFlag = 3
  End If
  
  If PctFlag = True Then
    If QPTrim$(fptxtPct.Text) = "" Then
     fptxtPct.BackColor = &HFFFF&
     frmBLMessageBoxJr.Label1.Caption = "Please enter a value for percent."
     frmBLMessageBoxJr.Label1.Top = 900
     frmBLMessageBoxJr.Show vbModal
     fptxtPct.BackColor = &HFFFFFF
     fptxtPct.SetFocus
     Close
     Exit Sub
    End If
  Else
    If fpcurrFee.DoubleValue = 0 Then
     fpcurrFee.BackColor = &HFFFF&
     frmBLMessageBoxJr.Label1.Caption = "Please enter a value for amount."
     frmBLMessageBoxJr.Label1.Top = 900
     frmBLMessageBoxJr.Show vbModal
     fpcurrFee.BackColor = &HFFFFFF
     fpcurrFee.SetFocus
     Close
     Exit Sub
    End If
  End If
  
  If PctFlag Then
    Amt# = CDbl(ReplaceString(fptxtPct.Text, "%", ""))
    Amt# = Amt# / 100
  Else
    Amt# = fpcurrFee.DoubleValue
  End If

  Call KillFile("artmppen.dat")
  OpenPenTransFile THandle

  OpenCustFile CHandle
  NumOfCustRecs = LOF(CHandle) / Len(CustRec)
  
  ReDim CatCharge(1 To 5) As Double
  ThisDate = Date2Num(fptxtXDate)
  
  If InStr(fpcmbRange.Text, "Only") Then
    RangeFlag = 1
  Else
    RangeFlag = 2
  End If
  OmitCnt = 0
  
  'NOT OK to calculate penalty fees if customer is involved in an
  'unposted license fee file
  If Exist("artmppst.dat") Then
    For cnt = 1 To NumOfCustRecs
      Get #CHandle, cnt, CustRec
      If QPTrim$(CustRec.IssueLicense) <> "Y" Then GoTo SkipThisOne
      
      If TBalFlag = 1 Then
        If CustRec.AcctBal <= 0 Then GoTo SkipThisOne
      ElseIf TBalFlag > 1 Then
        If CustRec.LicBal <= 0 Then GoTo SkipThisOne
      End If
      
      If RangeFlag = 1 Then
        If CustRec.VALID <> ThisDate Then GoTo SkipThisOne
      Else
        If CustRec.VALID > ThisDate Then GoTo SkipThisOne
      End If
      
      If EmpInLicProcess(QPTrim$(CustRec.CustNumb)) = True Then
        OmitCnt = OmitCnt + 1
        ReDim Preserve OmitList(1 To OmitCnt) As Long
        OmitList(OmitCnt) = cnt
      End If
SkipThisOne:
    Next cnt
  End If
  
  If OmitCnt > 0 Then
    frmBLOmitList.Label1.Alignment = 0
    frmBLOmitList.Label1.Caption = "The following is a list of all customers who qualify for a penalty fee but are currently included in an unposted license fee calculation.                                                                             Press F10: Exclude this list from penalty fees                                Press F5: Include this list in penalty fees and delete license file.      Press ESC: Abort penalty fee processing.                                Press F3: Print list."
    frmBLOmitList.Show vbModal
    If frmBLOmitList.fptxtChoice.Text = "delete" Then
      Unload frmBLOmitList
      KillFile "artmppst.dat"
      KillFile "artmplic.dat"
      KillFile "licprnOK.dat"
      OmitCnt = 0
      MainLog ("User elected to delete the 'artmppst.dat', 'artmplic.dat' and 'licprnOK.dat' files in order to allow all penalty fees to be calculated.")
    ElseIf frmBLOmitList.fptxtChoice.Text = "continue" Then
      Unload frmBLOmitList
      MainLog ("User presented with a list of all customers who would not be processed for penalty fees because they were involved in an unposted license fee process. User elected to continue penalty process excluding customers on list.")
    ElseIf frmBLOmitList.fptxtChoice.Text = "abort" Then
      Unload frmBLOmitList
      MainLog ("User elected to abort penalty calculations after being shown a list of all customers ineligible because they were already involved in a license fee file.")
      fpcmbBalType.SetFocus
      Close
      Exit Sub
    End If
  End If
  
  PayOmitCnt = 0

  frmBLShowPctComp.Label1 = "Searching For Customer Data"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  fpcmdHelp.Enabled = False

  For cnt = 1 To NumOfCustRecs
    Get #CHandle, cnt, CustRec

    If TBalFlag = 1 Then
      If CustRec.AcctBal <= 0 Then GoTo SkipThisOne1
    ElseIf TBalFlag > 1 Then
      If CustRec.LicBal <= 0 Then GoTo SkipThisOne1
    End If

    If RangeFlag = 1 Then
      If CustRec.VALID <> ThisDate Then GoTo SkipThisOne1
    Else
      If CustRec.VALID > ThisDate Then GoTo SkipThisOne1
    End If

    If EmpInPayProcess(QPTrim$(CustRec.CustNumb)) = True Then
      PayOmitCnt = PayOmitCnt + 1
      ReDim Preserve InPayOmit(1 To PayOmitCnt) As Long
      InPayOmit(PayOmitCnt) = cnt
    End If
SkipThisOne1:
    frmBLShowPctComp.ShowPctComp cnt, NumOfCustRecs
    If frmBLShowPctComp.Out = True Then
      Close
      frmBLShowPctComp.Out = False
      Unload frmBLShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdExit.Enabled = True
      cmdProcess.Enabled = True
      fpcmdHelp.Enabled = True
      Exit Sub
    End If
  Next cnt
  
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  fpcmdHelp.Enabled = True

  If PayOmitCnt > 0 Then
    frmBLPayOmitList.Label1.Alignment = 2
    frmBLPayOmitList.Label1.Height = 1000
    frmBLPayOmitList.Label1.Top = 700
    frmBLPayOmitList.cmdExit.Visible = False
    frmBLPayOmitList.cmdPrint.Left = 2680
    frmBLPayOmitList.cmdContinue.Text = "F10 Continue"
    frmBLPayOmitList.Label1.Caption = "The following is a list of all customers who qualify for a penalty fee and are currently included in an unposted payment file."
    frmBLPayOmitList.Show vbModal
'    If frmBLOmitList.fptxtChoice.Text = "continue" Then
'      Unload frmBLPayOmitList
'      MainLog ("User presented with a list of all customers who would not be processed for penalty fees because they were involved in an unposted payment file. User elected to continue penalty process excluding customers on list.")
'    ElseIf frmBLPayOmitList.fptxtChoice.Text = "abort" Then
'      Unload frmBLPayOmitList
'      MainLog ("User elected to abort penalty calculations after being shown a list of all customers ineligible because they were already involved in an unposted payment file.")
'      fpcmbBalType.SetFocus
'      Close
'      Exit Sub
'    End If
  End If
  
  frmBLShowPctComp.Label1 = "Calculating Penalty Amounts"
  frmBLShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  cmdExit.Enabled = False
  cmdProcess.Enabled = False
  fpcmdHelp.Enabled = False

  For cnt = 1 To NumOfCustRecs
    Get #CHandle, cnt, CustRec
      For y = 1 To OmitCnt
        If CInt(CustRec.CustNumb) = OmitList(y) Then
          GoTo NotThisOne
        End If
      Next y
      If QPTrim$(CustRec.Inactive) = "Y" Then GoTo NotThisOne
      If QPTrim$(CustRec.Deleted) = "Y" Or QPTrim$(CustRec.SortName) = "DELETED" Then GoTo NotThisOne
      If TBalFlag = 1 Then
        CustBal# = CDbl(CustRec.AcctBal)
      ElseIf TBalFlag > 1 Then
        CustBal# = CDbl(CustRec.LicBal)
      End If
      
      If RangeFlag = 1 Then
        If CustRec.VALID <> ThisDate Then GoTo NotThisOne
      Else
        If CustRec.VALID > ThisDate Then GoTo NotThisOne
      End If
      
      'balances in dos often won't work in windows
      'because when saved in dos as empty it comes
      'through here as really huge or infinitely small
      'numbers
      If CustBal# > 0.01 And CustBal# < 1000000000 Then
        PenCnt = PenCnt + 1
        
        If PctFlag Then
          CustPen# = OldRound#(CustBal# * Amt#)
        Else
          CustPen# = Amt#
        End If
        
        TransRec.CustomerNumber = Str$(cnt)
        TransRec.TransDate = Date2Num(fptxtTransDate.Text)
        TransRec.TransAmount = CustPen#
        TransRec.PenAmt = CustPen#
        Put #THandle, PenCnt, TransRec
      End If

      frmBLShowPctComp.ShowPctComp cnt, NumOfCustRecs
      If frmBLShowPctComp.Out = True Then
        Close
        frmBLShowPctComp.Out = False
        Unload frmBLShowPctComp
        EnableCloseButton Me.hwnd, True
        cmdExit.Enabled = True
        cmdProcess.Enabled = True
        fpcmdHelp.Enabled = True
        Exit Sub
      End If
NotThisOne:
  Next
  Unload frmBLShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdExit.Enabled = True
  cmdProcess.Enabled = True
  fpcmdHelp.Enabled = True

  Close
  
  If PenCnt = 0 Then
    frmBLMessageBoxJr.Label1.Caption = "There are no delinquent customers for the criteria entered."
    frmBLMessageBoxJr.Label1.Top = 800
    frmBLMessageBoxJr.Show vbModal
    MainLog ("Penalty processing attempt found no delinquent customers. File 'artmppen.dat' deleted.")
    KillFile "artmppen.dat"
    Exit Sub
  End If
  
  If TBalFlag = 1 Then
    frmBLMessageBoxJr.Label1.Caption = "Penalty calculations on total balances for " + Str(PenCnt) + " customers completed successfully."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    MainLog ("Penalty calculations (using " + fptxtPct.Text + "%) on total balances completed successfully for " + Str(PenCnt) + " customers.")
  ElseIf TBalFlag = 2 Then
    frmBLMessageBoxJr.Label1.Caption = "Penalty calculations on licenses only for " + Str(PenCnt) + " customers completed successfully."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    MainLog ("Penalty calculations (using " + fptxtPct.Text + "%) on license balances completed successfully for " + Str(PenCnt) + " customers.")
  ElseIf TBalFlag = 3 Then
    frmBLMessageBoxJr.Label1.Caption = "Penalty calculations for " + Str(PenCnt) + " customers completed successfully."
    frmBLMessageBoxJr.Label1.Top = 900
    frmBLMessageBoxJr.Show vbModal
    MainLog ("Penalty calculations (using " + fpcurrFee.Text + ") on total balances completed successfully for " + Str(PenCnt) + " customers.")
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Unload frmBLShowPctComp
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmBLPenCalc", "cmdProcess_Click", Erl)
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
    Case vbKeyF3:
      SendKeys "%H"
      Call fpcmdHelp_Click
      KeyCode = 0
    Case vbKeyF4:
      SendKeys "%E"
      Call fpcmdXList_Click
      KeyCode = 0
    Case vbKeyF1:
      SendKeys "%T"
      Call fpcmdHelp_Click
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
      KillFile "pencalcscr.dat"
      ClearInUse PWcnt
      MainLog ("BusinessLicense.exe terminated via menu bar on frmBLPenCalc.")
      Call Terminate
      End
    End If
  End If
End Sub

Private Sub LoadMe()
  Dim TownRec As TownSetUpType
  Dim TownHandle As Integer
  Dim Towncnt As Integer
  Dim FileHandle As Integer
  Dim One As Integer
  Dim DHandle As Integer
  
  On Error Resume Next
  lblBalloon.Visible = False
'  fpcmbBalType.ToolTipText = "Penalty calculations can be conducted on the customer's total balance or the calculation can be restricted to solely to their license balance."
'  fpcurrFee.ToolTipText = "The amount or percent is set in the Town Setup screen. Values for each are entered here."
'  fptxtPct.ToolTipText = "The amount or percent is set in the Town Setup screen. Values for each are entered here."
'  fptxtXDate.ToolTipText = "Penalties are assessed to customers whose business license expires on this date."
'  fpcmdXList.ToolTipText = "Press to bring up a complete list of all active customers and their expiration dates."
'  fpcmbRange.ToolTipText = "You can elect to print all those delinquent up to and including the expiration date or just those delinquent on the date entered."
'  cmdProcess.ToolTipText = "Press to begin penalty calculations."
'  cmdExit.ToolTipText = "Press to exit this screen."
'  fpcmdHelp.ToolTipText = "Press 'Turn Help On' to activate instructional balloons that will appear when you place the cursor over any field on the screen. Press 'Turn Help Off' to deactivate the instructional balloons."
  One = 1
  DHandle = FreeFile
  'pencalcscr.dat is used by the expiration list to know
  'where to return data and the focus when it is closed
  Open "pencalcscr.dat" For Output As DHandle Len = 2
  Print #DHandle, One
  Close DHandle
  
  OpenTownFile TownHandle
  Towncnt = LOF(TownHandle) / Len(TownRec)
  Get TownHandle, 1, TownRec
  Close TownHandle
  
  fptxtXDate = Date
  fptxtTransDate = Date
  If QPTrim$(TownRec.UseAmtPctYN) = "Amt" Then
    fpcmbBalType.Enabled = False
    fptxtPct.Visible = False
    PctFlag = False
    Label1.Caption = "Enter Amount"
    fptxtPct.TabIndex = 1
    fpcmbBalType.Text = "Using Fixed Amount"
  Else
    fpcurrFee.Visible = False
    PctFlag = True
    Label1.Caption = "Enter Percent"
    fpcurrFee.TabIndex = 1
    fpcmbBalType.Text = "Total Balance"
    fpcmbBalType.AddItem "Total Balance"
    fpcmbBalType.AddItem "License Only"
  End If
  
  fpcmbRange.Text = "Up To And Include This Expiration"
  fpcmbRange.AddItem "Up To And Include This Expiration"
  fpcmbRange.AddItem "This Expiration Only"
  MainLog ("Penalty calculation screen opened.")
End Sub

Private Sub fpcmbBalType_Change()
  If QPTrim$(fpcmbBalType.Text) = "" Then
    fpcmbBalType.Text = "Total Balance"
  End If
End Sub

Private Sub fpcmbBalType_KeyDown(KeyCode As Integer, Shift As Integer)
  'this keeps the user from inadvertently changing data on this
  'combo box if they are scrolling through the form
  If KeyCode = vbKeySpace Then
    fpcmbBalType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbBalType.ListIndex = -1
  End If
  If fpcmbBalType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      If PctFlag = True Then
        fptxtPct.SetFocus
      Else
        fpcurrFee.SetFocus
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

Private Sub cmdExit_Click()
  KillFile "pencalcscr.dat"
  frmBLPenProcMenu.Show
  DoEvents
  Unload frmBLPenCalc
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
      If fpcmbBalType.Enabled = True Then
        fpcmbBalType.SetFocus
      ElseIf fptxtPct.Visible = True Then
        fptxtPct.SetFocus
      Else
        fpcurrFee.SetFocus
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

Private Sub fpcmdHelp_Click()
  If InStr(fpcmdHelp.Text, "On") Then
    fpcmdHelp.Text = "F1 &Turn Help Off"
    btnHelp.AutoScan = fpAutoScanPopupOnly
    lblBalloon.Visible = True
    fpcmbBalType.ToolTipText = ""
    fpcurrFee.ToolTipText = ""
    fptxtPct.ToolTipText = ""
    fptxtXDate.ToolTipText = ""
    fpcmdXList.ToolTipText = ""
    fpcmbRange.ToolTipText = ""
    cmdProcess.ToolTipText = ""
    cmdExit.ToolTipText = ""
    fpcmdHelp.ToolTipText = ""
  ElseIf InStr(fpcmdHelp.Text, "Off") Then
    fpcmdHelp.Text = "F1 &Turn Help On"
    btnHelp.AutoScan = fpAutoScanOff
    lblBalloon.Visible = False
'    fpcmbBalType.ToolTipText = "Penalty calculations can be conducted on the customer's total balance or the calculation can be restricted to solely to their license balance."
'    fpcurrFee.ToolTipText = "The amount or percent is set in the Town Setup screen. Values for each are entered here."
'    fptxtPct.ToolTipText = "The amount or percent is set in the Town Setup screen. Values for each are entered here."
'    fptxtXDate.ToolTipText = "Penalties are assessed to customers whose business license expires on this date."
'    fpcmdXList.ToolTipText = "Press to bring up a complete list of all active customers and their expiration dates."
'    fpcmbRange.ToolTipText = "You can elect to print all those delinquent up to and including the expiration date or just those delinquent on the date entered."
'    cmdProcess.ToolTipText = "Press to begin penalty calculations."
'    cmdExit.ToolTipText = "Press to exit this screen."
'    fpcmdHelp.ToolTipText = "Press 'Turn Help On' to activate instructional balloons that will appear when you place the cursor over any field on the screen. Press 'Turn Help Off' to deactivate the instructional balloons."
  End If
  
End Sub

Private Sub fpcmdXList_Click()
  frmBLXDateList.Show vbModal
End Sub

Private Sub fptxtXDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fpcmbBalType.SetFocus
  ElseIf KeyCode = vbKeyUp Then
    If PctFlag = False Then
      fpcurrFee.SetFocus
    Else
      fptxtPct.SetFocus
    End If
  End If

End Sub

