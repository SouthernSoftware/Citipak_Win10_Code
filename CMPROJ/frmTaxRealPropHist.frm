VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "BTN32A20.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "FLP32A30.OCX"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmTaxRealPropHist 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Real Property History"
   ClientHeight    =   8736
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   11640
   Icon            =   "frmTaxRealPropHist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8736
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   7110
      Left            =   1920
      TabIndex        =   9
      Top             =   810
      Width           =   7785
      _Version        =   196609
      _ExtentX        =   13732
      _ExtentY        =   12541
      _StockProps     =   70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.81
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
      Picture         =   "frmTaxRealPropHist.frx":08CA
      Begin LpLib.fpCombo fpcmbDetSum 
         Height          =   384
         Left            =   3528
         TabIndex        =   7
         Top             =   4800
         Width           =   2844
         _Version        =   196608
         _ExtentX        =   5016
         _ExtentY        =   677
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
         ColDesigner     =   "frmTaxRealPropHist.frx":08E6
      End
      Begin LpLib.fpCombo fpcmbTransType 
         Height          =   384
         Left            =   3408
         TabIndex        =   4
         Top             =   3072
         Width           =   3060
         _Version        =   196608
         _ExtentX        =   5397
         _ExtentY        =   677
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
         ColDesigner     =   "frmTaxRealPropHist.frx":0D31
      End
      Begin LpLib.fpCombo fpcmbPrintOpt 
         Height          =   384
         Left            =   3048
         TabIndex        =   8
         Top             =   5328
         Width           =   3564
         _Version        =   196608
         _ExtentX        =   6286
         _ExtentY        =   677
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
         ColDesigner     =   "frmTaxRealPropHist.frx":117C
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   636
         Left            =   3120
         TabIndex        =   10
         TabStop         =   0   'False
         Tag             =   "Press the 'Cancel' button to exit this screen and return to the main 'Business License Reports' menu."
         Top             =   6096
         Width           =   1740
         _Version        =   131072
         _ExtentX        =   3069
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
         ButtonDesigner  =   "frmTaxRealPropHist.frx":15C7
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   636
         Left            =   5232
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   $"frmTaxRealPropHist.frx":17DD
         Top             =   6096
         Width           =   1740
         _Version        =   131072
         _ExtentX        =   3069
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
         ButtonDesigner  =   "frmTaxRealPropHist.frx":1888
      End
      Begin EditLib.fpDateTime fptxtBegDate 
         Height          =   375
         Left            =   3960
         TabIndex        =   5
         Top             =   3720
         Width           =   1620
         _Version        =   196608
         _ExtentX        =   2857
         _ExtentY        =   661
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.4
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
         Text            =   "02/24/2005"
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
      Begin EditLib.fpDateTime fptxtEndDate 
         Height          =   375
         Left            =   3960
         TabIndex        =   6
         Top             =   4200
         Width           =   1620
         _Version        =   196608
         _ExtentX        =   2857
         _ExtentY        =   661
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.4
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
         Text            =   "02/24/2005"
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
      Begin EditLib.fpText fptxtPin 
         Height          =   390
         Left            =   3600
         TabIndex        =   0
         Top             =   1320
         Width           =   2295
         _Version        =   196608
         _ExtentX        =   4048
         _ExtentY        =   688
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   10.8
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
         AutoCase        =   1
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
         CharValidationText=   ""
         MaxLength       =   20
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
      Begin fpBtnAtlLibCtl.fpBtn cmdLookup 
         Height          =   375
         Left            =   600
         TabIndex        =   1
         Top             =   1920
         Width           =   3015
         _Version        =   131072
         _ExtentX        =   5318
         _ExtentY        =   661
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
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
         ButtonDesigner  =   "frmTaxRealPropHist.frx":1A9F
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdLookUpByOwner 
         Height          =   375
         Left            =   4080
         TabIndex        =   2
         Top             =   1920
         Width           =   3015
         _Version        =   131072
         _ExtentX        =   5318
         _ExtentY        =   661
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
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
         ButtonDesigner  =   "frmTaxRealPropHist.frx":1CC4
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdDetail 
         Height          =   636
         Left            =   1020
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   6096
         Width           =   1740
         _Version        =   131072
         _ExtentX        =   3069
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
         ButtonDesigner  =   "frmTaxRealPropHist.frx":1EEA
      End
      Begin EditLib.fpText fptxtCurrOwner 
         Height          =   390
         Left            =   2280
         TabIndex        =   3
         Top             =   2400
         Width           =   4935
         _Version        =   196608
         _ExtentX        =   8705
         _ExtentY        =   688
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   10.8
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
         AutoCase        =   1
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
         CharValidationText=   ""
         MaxLength       =   70
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
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Current Owner:"
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
         Left            =   480
         TabIndex        =   20
         Top             =   2475
         Width           =   1665
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Property Pin #:"
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
         Left            =   1680
         TabIndex        =   18
         Top             =   1395
         Width           =   1815
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
         Left            =   1395
         TabIndex        =   17
         Top             =   5430
         Width           =   1500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type:"
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
         TabIndex        =   16
         Top             =   3165
         Width           =   2055
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   1290
         Top             =   435
         Width           =   5265
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Transaction History By Property"
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
         Height          =   390
         Left            =   1440
         TabIndex        =   15
         Top             =   570
         Width           =   4935
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   2940
         Left            =   1005
         Top             =   2925
         Width           =   5970
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
         Height          =   300
         Left            =   1800
         TabIndex        =   14
         Top             =   3795
         Width           =   2055
      End
      Begin VB.Label Label4 
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
         Height          =   300
         Left            =   1800
         TabIndex        =   13
         Top             =   4260
         Width           =   2055
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Detail/Summary:"
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
         Left            =   1440
         TabIndex        =   12
         Top             =   4890
         Width           =   1905
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000004&
         X1              =   1005
         X2              =   6955
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000004&
         X1              =   1005
         X2              =   6955
         Y1              =   4680
         Y2              =   4680
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   7380
      Left            =   1800
      Top             =   675
      Width           =   8055
   End
End
Attribute VB_Name = "frmTaxRealPropHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim Town$
  Public GRealRec As Long
  Dim GAdd As String
  Dim Opt1Desc$
  Dim Opt2Desc$
  Dim Opt3Desc$
  Dim GValidPin As Boolean

Private Sub cmdDetail_Click()
  If QPTrim$(fptxtPin.Text) <> "" Then
    frmTaxRealDetail.Show vbModal
    DoEvents
  Else
    Call TaxMsg(900, "Please enter a pin number.")
    fptxtPin.SetFocus
  End If
End Sub

Private Sub cmdExit_Click()
  KillFile "realhist.dat"
  'frmTaxReportsMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdLookup_Click()
  'frmTaxRealLookup.Show
  DoEvents
  
End Sub

Private Sub cmdLookUpByOwner_Click()
 ' frmTaxRealLookupByOwner.Show
  DoEvents
End Sub

Private Sub cmdProcess_Click()
  If GValidPin = False Then
    Call TaxMsg(900, "The pin number entered is not valid.")
    Exit Sub
  End If
  If fpcmbDetSum.Text = "Summary" Then
    If fpcmbPrintOpt.Text = "Graphical" Then
      Call PrintGraphics
    Else
      Call TaxMsg(900, "Pitch 12 is recommended for this printout.")
      Call PrintText
    End If
  Else
    If fpcmbPrintOpt.Text = "Graphical" Then
      Call PrintGraphicsDet
    Else
      Call TaxMsg(900, "Pitch 12 is recommended for this printout.")
      Call PrintTextDet
    End If
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
    Case vbKeyF7:
      SendKeys "%L"
      Call cmdLookup_Click
      KeyCode = 0
    Case vbKeyF8:
      SendKeys "%O"
      Call cmdLookUpByOwner_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  Call Loadme
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      KillFile "realhist.dat"
      TXLog ("cm.exe terminated via menu bar on frmTaxRealPropHist.")
      Call CMTerminate
      End
    End If
  End If

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

Private Sub fpcmbPrintOpt_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbPrintOpt.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPrintOpt.ListIndex = -1
  End If
  If fpcmbPrintOpt.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtPin.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbDetSum_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbDetSum.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbDetSum.ListIndex = -1
  End If
  If fpcmbDetSum.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbTransType.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbTransType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbTransType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbTransType.ListIndex = -1
  End If
  If fpcmbTransType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fptxtBegDate.SetFocus
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
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim One As Integer
  Dim AHandle As Integer
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  Town$ = QPTrim$(TaxMasterRec.Name)
  Opt1Desc = QPTrim$(TaxMasterRec.OptRev1)
  Opt2Desc = QPTrim$(TaxMasterRec.OptRev2)
  Opt3Desc = QPTrim$(TaxMasterRec.OptRev3)
  fpcmbTransType.Text = " 0) All"
  fpcmbTransType.AddItem " 0) All"
  fpcmbTransType.AddItem " 1) Billing" 'include #9
  fpcmbTransType.AddItem " 2) Payment"
  fpcmbTransType.AddItem " 3) Release"
  fpcmbTransType.AddItem " 4) Interest"
  fpcmbTransType.AddItem " 5) Penalty"
  fpcmbTransType.AddItem " 6) Advertising"
  fpcmbTransType.AddItem " 7) Adjust Pay Down"
  fpcmbTransType.AddItem "13) Adjust Bill Down" 'include #23
  fpcmbTransType.AddItem "14) Adjust Bill Up" 'include #24
'  fpcmbTransType.AddItem "21) Overpayment" 'include #22
  One = 1
  AHandle = FreeFile
  Open "realhist.dat" For Output As AHandle
  Print #AHandle, One
  Close AHandle
  
  GValidPin = False
  fptxtBegDate = Date
  fptxtEndDate = Date
  fpcmbPrintOpt.Text = "Graphical"
  fpcmbPrintOpt.AddItem "Graphical"
  fpcmbPrintOpt.AddItem "Text"
  fpcmbDetSum.Text = "Summary"
  fpcmbDetSum.AddItem "Detail"
  fpcmbDetSum.AddItem "Summary"

End Sub

Public Sub LoadRealRec()
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRealRecs As Long
  Dim CustPin As Long
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  
  On Error GoTo BooBoo
  GValidPin = True
  OpenRealPropFile RHandle, NumOfRealRecs
  Get RHandle, GRealRec, RealRec
  Close RHandle
  If QPTrim$(RealRec.PropAddr) <> "" Then
    GAdd = QPTrim$(RealRec.PropAddr)
  Else
    GAdd = "No Address Saved"
  End If
  fptxtPin.Text = QPTrim$(RealRec.RealPin)
  CustPin = RealRec.CustPin
  OpenTaxCustFile TCHandle, NumOfTCRecs
  Get TCHandle, CustPin, TaxCust
  Close TCHandle
  
  fptxtCurrOwner.Text = "#" + QPTrim$(Using$("####0", RealRec.CustPin)) + "   " + QPTrim$(TaxCust.CustName)
  
  Exit Sub
  
BooBoo:
  fptxtCurrOwner.Text = "Error: Owner Not Found"
  
End Sub

Private Function Check4ValidPin() As Boolean
  Dim RealRec As PropertyRecType
  Dim RHandle As Integer
  Dim NumOfRRecs As Long
  Dim x As Long
  Dim ThisPin$
  
  GValidPin = False
  Check4ValidPin = False
  If QPTrim$(fptxtPin.Text) = "" Then
    Exit Function
  Else
    ThisPin = QPTrim$(fptxtPin.Text)
  End If
  OpenRealPropFile RHandle, NumOfRRecs
  For x = 1 To NumOfRRecs
    Get RHandle, x, RealRec
    If RealRec.Deleted = -1 Then GoTo Deleted
    If ThisPin = QPTrim$(RealRec.RealPin) Then
      Check4ValidPin = True
      GValidPin = True
      GRealRec = x
      Exit For
    End If
Deleted:
  Next x
  If x > NumOfRRecs Then
    Call TaxMsg(900, "The pin number entered could not be found.")
  End If
  
End Function

Private Sub fptxtPin_LostFocus()
  If Check4ValidPin = False Then Exit Sub
  Call LoadRealRec
End Sub

Private Sub PrintGraphicsDet()
  Dim x As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim ThisPin$
  Dim ThisClass As Integer
  Dim ThisType As String
  Dim BegDate As Integer
  Dim EndDate As Integer
  Dim RptFile$
  Dim RptHandle As Integer
  Dim dlm$
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim Owner$
  Dim TCnt As Long
  Dim TotAmt As Double
  Dim PrincDif As Double
  Dim IntDif As Double
  Dim AdvDif As Double
  Dim LateListDif As Double
  Dim Opt1Dif As Double
  Dim Opt2Dif As Double
  Dim Opt3Dif As Double
  Dim ThisTransType As String
  Dim BillToOwner$
  Dim CustRec As Long
  Dim BillCustRec As Long
  Dim ThisBal As Double
  
  dlm$ = "~"
  If Mid(fpcmbTransType.Text, 1, 1) = "" Then
    ThisClass = CInt(Mid(fpcmbTransType.Text, 2, 1))
  Else
    ThisClass = CInt(Mid(fpcmbTransType.Text, 1, 2))
  End If

  Select Case ThisClass
    Case 1
      ThisType = "Billing"
    Case 2
      ThisType = "Payment"
    Case 3
      ThisType = "Release"
    Case 4
      ThisType = "Interest"
    Case 5
      ThisType = "Unknown"
    Case 6
      ThisType = "Advertising Charge"
    Case 7
      ThisType = "Adjust Pay Down"
    Case 13
      ThisType = "Adjust Bill Down"
    Case 14
      ThisType = "Adjust Bill Up"
    Case Else
      ThisType = "All"
  End Select

  BegDate = Date2Num(fptxtBegDate.Text)
  EndDate = Date2Num(fptxtEndDate.Text)
  ThisPin = QPTrim$(fptxtPin.Text)
  RptFile$ = "TAXRPTS\REALHIST.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle

  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  FrmShowPctComp.Label1 = "Gathering Tax Transaction Data"
  FrmShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
    If QPTrim$(TaxTrans.RealPin) = ThisPin Then
      If ThisClass <> 0 And TaxTrans.TranType <> ThisClass Then GoTo SkipIt
      If TaxTrans.TransDate >= BegDate And TaxTrans.TransDate <= EndDate Then
        Select Case TaxTrans.TranType
          Case 1
            ThisTransType = "Billing"
         Case 2
            ThisTransType = "Payment"
          Case 3
            ThisTransType = "Release"
          Case 4
            ThisTransType = "Interest"
          Case 5
            ThisTransType = "Unknown"
          Case 6
            ThisTransType = "Advertising Charge"
          Case 7
            ThisTransType = "Adjust Pay Down"
          Case 9
            ThisTransType = "Credit Applied at Billing"
          Case 13
            ThisTransType = "Adjust Bill Down"
          Case 14
            ThisTransType = "Adjust Bill Up"
          Case 21
            ThisTransType = "Billpay/Overpay"
'        Case 22
'          ThisTransType = "Overpayment"
          Case 24
            ThisTransType = "Adjust Bill Up Affecting Credit Balance"
          Case Else
            ThisTransType = "Unknown"
        End Select
        TCnt = TCnt + 1
        TotAmt = OldRound(TotAmt + TaxTrans.Amount)
        PrincDif = OldRound(TaxTrans.Revenue.Principle1 - TaxTrans.Revenue.Principle1Pd)
        IntDif = OldRound(TaxTrans.Revenue.Interest - TaxTrans.Revenue.InterestPd)
        AdvDif = OldRound(TaxTrans.Revenue.Collection - TaxTrans.Revenue.CollectionPd)
        LateListDif = OldRound(TaxTrans.Revenue.LateList - TaxTrans.Revenue.LateListPd)
        Opt1Dif = OldRound(TaxTrans.Revenue.RevOpt1 - TaxTrans.Revenue.RevOpt1Pd)
        Opt2Dif = OldRound(TaxTrans.Revenue.RevOpt2 - TaxTrans.Revenue.RevOpt2Pd)
        Opt3Dif = OldRound(TaxTrans.Revenue.RevOpt3 - TaxTrans.Revenue.RevOpt3Pd)
        ThisBal = OldRound(PrincDif + IntDif + AdvDif + LateListDif + Opt1Dif + Opt2Dif + Opt3Dif)
        GoSub GetOwner
        '                   0            1                    2                              3
        Print #RptHandle, Town$; dlm; Owner; dlm; MakeRegDate(TaxTrans.TransDate); dlm; ThisPin; dlm;
        '                     4                   5                         6                      7
        Print #RptHandle, ThisType; dlm; MakeRegDate(BegDate); dlm; MakeRegDate(EndDate); dlm; TaxTrans.TaxYear; dlm;
        '                        8                            9                        10
        Print #RptHandle, TaxTrans.Amount; dlm; QPTrim$(TaxTrans.Description); dlm; ThisTransType; dlm;
        '                               11                              12                           13
       Print #RptHandle, TaxTrans.Revenue.Principle1; dlm; TaxTrans.Revenue.Principle1Pd; ; dlm; PrincDif; dlm;
        '                               14                              15                      16
        Print #RptHandle, TaxTrans.Revenue.Interest; dlm; TaxTrans.Revenue.InterestPd; ; dlm; IntDif; dlm;
        '                               17                              18                        19
        Print #RptHandle, TaxTrans.Revenue.Collection; dlm; TaxTrans.Revenue.CollectionPd; dlm; AdvDif; dlm;
        '                               20                              21                       22
        Print #RptHandle, TaxTrans.Revenue.LateList; dlm; TaxTrans.Revenue.LateListPd; dlm; LateListDif; dlm;
        '                               23                              24                   25
        Print #RptHandle, TaxTrans.Revenue.RevOpt1; dlm; TaxTrans.Revenue.RevOpt1Pd; dlm; Opt1Dif; dlm;
        '                               26                              27                   28
        Print #RptHandle, TaxTrans.Revenue.RevOpt2; dlm; TaxTrans.Revenue.RevOpt2Pd; dlm; Opt2Dif; dlm;
        '                               29                              30                   31
        Print #RptHandle, TaxTrans.Revenue.RevOpt3; dlm; TaxTrans.Revenue.RevOpt3Pd; dlm; Opt3Dif; dlm;
        '                    32             33             34           35           36             37
        Print #RptHandle, Opt1Desc; dlm; Opt2Desc; dlm; Opt3Desc; dlm; GAdd; dlm; CustRec; dlm; BillCustRec; dlm;

        If TaxTrans.BelongTo > 0 Then
          Get TTHandle, TaxTrans.BelongTo, TaxTrans
          GoSub GetOwner
          '                             38                               39             40          41
          Print #RptHandle, ParseBillNum(TaxTrans.Description); dlm; BillToOwner; dlm; TCnt; dlm; ThisBal
        Else
          '                 38          39             40           41
          Print #RptHandle, 0; dlm; BillToOwner; dlm; TCnt; dlm; ThisBal
        End If
      End If
    End If
SkipIt:
    FrmShowPctComp.ShowPctComp x, NumOfTTRecs
    If FrmShowPctComp.Out = True Then
      Close
      FrmShowPctComp.Out = False
      Unload FrmShowPctComp
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  Unload FrmShowPctComp
  EnableCloseButton Me.hwnd, True
  
  Close
  If TCnt = 0 Then
    Call TaxMsg(900, "No transactions are recorded for this property.")
    Exit Sub
  End If
  Call arTaxRealHistRpt.Show
  frmTaxLoadReport.Show
  
  Exit Sub

GetOwner:
  Get TCHandle, TaxTrans.CustomerRec, TaxCust
  Owner = QPTrim$(TaxCust.CustName)
  BillToOwner = Owner
  CustRec = TaxCust.PIN
  BillCustRec = CustRec
  Return
End Sub

Private Sub PrintTextDet()
  Dim x As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim ThisPin$
  Dim ThisClass As Integer
  Dim ThisType As String
  Dim BegDate As Integer
  Dim EndDate As Integer
  Dim RptFile$
  Dim RptHandle As Integer
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim Owner$
  Dim TCnt As Long
  Dim TotAmt As Double
  Dim PrincDif As Double
  Dim IntDif As Double
  Dim AdvDif As Double
  Dim LateListDif As Double
  Dim Opt1Dif As Double
  Dim Opt2Dif As Double
  Dim Opt3Dif As Double
  Dim ThisTransType As String
  Dim BillToOwner$
  Dim CustRec As Long
  Dim BillCustRec As Long
  Dim LineCnt As Integer
  Dim MaxLines As Integer
  Dim Page As Integer
  Dim FF$
  Dim ThisBal As Double
  Dim ThisBillNum$
  Dim ThisRec As Long
  
  FF$ = Chr(12)
  MaxLines = 58
  
  If Mid(fpcmbTransType.Text, 1, 1) = "" Then
    ThisClass = CInt(Mid(fpcmbTransType.Text, 2, 1))
  Else
    ThisClass = CInt(Mid(fpcmbTransType.Text, 1, 2))
  End If

  Select Case ThisClass
    Case 1
      ThisType = "Billing"
    Case 2
      ThisType = "Payment"
    Case 3
      ThisType = "Release"
    Case 4
      ThisType = "Interest"
    Case 5
      ThisType = "Unknown"
    Case 6
      ThisType = "Advertising Charge"
    Case 7
      ThisType = "Adjust Pay Down"
    Case 13
      ThisType = "Adjust Bill Down"
    Case 14
      ThisType = "Adjust Bill Up"
    Case Else
      ThisType = "All"
  End Select

  BegDate = Date2Num(fptxtBegDate.Text)
  EndDate = Date2Num(fptxtEndDate.Text)
  ThisPin = QPTrim$(fptxtPin.Text)
  RptFile$ = "TAXRPTS\REALHIST.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle

  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  frmTaxShowPctComp.Label1 = "Gathering Tax Transaction Data"
  frmTaxShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  GoSub PrintHeader
  
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
    If QPTrim$(TaxTrans.RealPin) = ThisPin Then
      If ThisClass <> 0 And TaxTrans.TranType <> ThisClass Then GoTo SkipIt
      If TaxTrans.TransDate >= BegDate And TaxTrans.TransDate <= EndDate Then
        Select Case TaxTrans.TranType
          Case 1
            ThisTransType = "Billing"
         Case 2
            ThisTransType = "Payment"
          Case 3
            ThisTransType = "Release"
          Case 4
            ThisTransType = "Interest"
          Case 5
            ThisTransType = "Unknown"
          Case 6
            ThisTransType = "Advertising Charge"
          Case 7
            ThisTransType = "Adjust Pay Down"
          Case 9
            ThisTransType = "Credit Applied at Billing"
          Case 13
            ThisTransType = "Adjust Bill Down"
          Case 14
            ThisTransType = "Adjust Bill Up"
          Case 21
            ThisTransType = "Billpay/Overpay"
'        Case 22
'          ThisTransType = "Overpayment"
          Case 24
            ThisTransType = "Adjust Bill Up Affecting Credit Balance"
          Case Else
            ThisTransType = "Unknown"
        End Select
        TCnt = TCnt + 1
        TotAmt = OldRound(TotAmt + TaxTrans.Amount)
        PrincDif = OldRound(TaxTrans.Revenue.Principle1 - TaxTrans.Revenue.Principle1Pd)
        IntDif = OldRound(TaxTrans.Revenue.Interest - TaxTrans.Revenue.InterestPd)
        AdvDif = OldRound(TaxTrans.Revenue.Collection - TaxTrans.Revenue.CollectionPd)
        LateListDif = OldRound(TaxTrans.Revenue.LateList - TaxTrans.Revenue.LateListPd)
        Opt1Dif = OldRound(TaxTrans.Revenue.RevOpt1 - TaxTrans.Revenue.RevOpt1Pd)
        Opt2Dif = OldRound(TaxTrans.Revenue.RevOpt2 - TaxTrans.Revenue.RevOpt2Pd)
        Opt3Dif = OldRound(TaxTrans.Revenue.RevOpt3 - TaxTrans.Revenue.RevOpt3Pd)
        GoSub GetOwner
        ThisBal = OldRound(PrincDif + IntDif + AdvDif + LateListDif + Opt1Dif + Opt2Dif + Opt3Dif)
        Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); QPTrim$(TaxTrans.Description);
        Print #RptHandle, Tab(37); Using$("###0", TaxTrans.TaxYear); Tab(45); Using$("$##,##0.00", TaxTrans.Amount); Tab(63);
        LineCnt = LineCnt + 1
        If TaxTrans.BelongTo > 0 Then
          Get TTHandle, TaxTrans.BelongTo, TaxTrans
          ThisBillNum = ParseBillNum(TaxTrans.Description)
          If IsNumeric(ThisBillNum) Then
            Print #RptHandle, Using$("######", CDbl(ThisBillNum));
          Else
            Print #RptHandle, "   " + ThisBillNum;
          End If
        Else
          Print #RptHandle, "    N/A";
        End If
      
        Get TTHandle, x, TaxTrans
        Print #RptHandle, Tab(79); ThisTransType
        If LineCnt >= MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintHeader
        End If
        If ThisTransType <> "Billing" Then
          Print #RptHandle, Tab(15); "Principal         "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle1); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle1Pd)
          Print #RptHandle, Tab(15); "Interest          "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Interest); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.InterestPd)
          Print #RptHandle, Tab(15); "Advertising       "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Collection); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.CollectionPd)
          Print #RptHandle, Tab(15); "Late Listing      "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.LateList); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.LateListPd)
          LineCnt = LineCnt + 4
          If Len(Opt1Desc) > 0 Then
            Print #RptHandle, Tab(15); Opt1Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1Pd)
            LineCnt = LineCnt + 1
          End If
          If Len(Opt2Desc) > 0 Then
            Print #RptHandle, Tab(15); Opt2Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2Pd)
            LineCnt = LineCnt + 1
          End If
            If Len(Opt3Desc) > 0 Then
            Print #RptHandle, Tab(15); Opt3Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3Pd)
            LineCnt = LineCnt + 1
          End If
        Else
          Print #RptHandle, Tab(15); "Principal         "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Principle1); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.Principle1Pd); Tab(80); Using$("$##,##0.00", PrincDif)
          Print #RptHandle, Tab(15); "Interest          "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Interest); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.InterestPd); Tab(80); Using$("$##,##0.00", IntDif)
          Print #RptHandle, Tab(15); "Advertising       "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.Collection); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.CollectionPd); Tab(80); Using$("$##,##0.00", AdvDif)
          Print #RptHandle, Tab(15); "Late Listing      "; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.LateList); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.LateListPd); Tab(80); Using$("$##,##0.00", LateListDif)
          LineCnt = LineCnt + 4
          If Len(Opt1Desc) > 0 Then
            Print #RptHandle, Tab(15); Opt1Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt1Pd); Tab(80); Using$("$##,##0.00", Opt1Dif)
            LineCnt = LineCnt + 1
          End If
          If Len(Opt2Desc) > 0 Then
            Print #RptHandle, Tab(15); Opt2Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt2Pd); Tab(80); Using$("$##,##0.00", Opt2Dif)
            LineCnt = LineCnt + 1
          End If
            If Len(Opt3Desc) > 0 Then
            Print #RptHandle, Tab(15); Opt3Desc; Tab(40); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3); Tab(60); Using$("$##,##0.00", TaxTrans.Revenue.RevOpt3Pd); Tab(80); Using$("$##,##0.00", Opt3Dif)
            LineCnt = LineCnt + 1
          End If
          Print #RptHandle, Tab(15); "Bill Balance:"; Tab(80); Using$("$##,##0.00", ThisBal)
        End If
        Print #RptHandle, Tab(12); "Transaction Owner: " + QPTrim$(Owner)
        Print #RptHandle, String(89, "-")
        LineCnt = LineCnt + 2
      End If
    End If
SkipIt:
    frmTaxShowPctComp.ShowPctComp x, NumOfTTRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  If LineCnt >= MaxLines - 1 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  
  Print #RptHandle, "Total Transaction Count: " + Using("####0", TCnt)
  
  Print #RptHandle, FF$
  Close
  If TCnt = 0 Then
    Call TaxMsg(900, "No transactions are recorded for this property.")
    Exit Sub
  End If
 
  ViewPrint RptFile, "Tax Real Estate History Journal", True
  
  Exit Sub

GetOwner:
  Get TCHandle, TaxTrans.CustomerRec, TaxCust
  Owner = QPTrim$(TaxCust.CustName)
  BillToOwner = Owner
  CustRec = TaxCust.PIN
  BillCustRec = CustRec
  
  Return

PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(30); "Tax Real Estate History Journal"
  Print #RptHandle, Town$; Tab(75); "Page #: " + CStr(Page)
  Print #RptHandle, "For Pin #: " + ThisPin
  Print #RptHandle, "Address: " + GAdd
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle, "Transaction Type: " + ThisType
  Print #RptHandle, "Date Range: " + fptxtBegDate.Text + " to " + fptxtEndDate.Text
  Print #RptHandle, "Trans Date"; Tab(12); "Description"; Tab(35); "Tax Year"; Tab(47); "Trans Amt"; Tab(60); "Belongs To"; Tab(77); "Trans Type"
  Print #RptHandle, Tab(12); "Revenue"; Tab(44); "Amount"; Tab(59); "Amount Paid"; Tab(83); "Balance"
  Print #RptHandle, String(89, "-")
  LineCnt = 10
  
  Return

PrintRealHeader:
  Print #RptHandle, "Pin #: " + ThisPin; Tab(30); "Address: " + GAdd
  Print #RptHandle, String(89, "-")
  LineCnt = LineCnt + 2
  
  Return
End Sub

Private Sub PrintGraphics()
  Dim x As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim ThisPin$
  Dim ThisClass As Integer
  Dim ThisType As String
  Dim BegDate As Integer
  Dim EndDate As Integer
  Dim RptFile$
  Dim RptHandle As Integer
  Dim dlm$
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim Owner$
  Dim TCnt As Long
  Dim TotAmt As Double
  Dim PrincDif As Double
  Dim IntDif As Double
  Dim AdvDif As Double
  Dim LateListDif As Double
  Dim Opt1Dif As Double
  Dim Opt2Dif As Double
  Dim Opt3Dif As Double
  Dim ThisTransType As String
  Dim BillToOwner$
  Dim CustRec As Long
  Dim BillCustRec As Long
  Dim ThisBal As Double
  
  dlm$ = "~"
  If Mid(fpcmbTransType.Text, 1, 1) = "" Then
    ThisClass = CInt(Mid(fpcmbTransType.Text, 2, 1))
  Else
    ThisClass = CInt(Mid(fpcmbTransType.Text, 1, 2))
  End If

  Select Case ThisClass
    Case 1
      ThisType = "Billing"
    Case 2
      ThisType = "Payment"
    Case 3
      ThisType = "Release"
    Case 4
      ThisType = "Interest"
    Case 5
      ThisType = "Unknown"
    Case 6
      ThisType = "Advertising Charge"
    Case 7
      ThisType = "Adjust Pay Down"
    Case 13
      ThisType = "Adjust Bill Down"
    Case 14
      ThisType = "Adjust Bill Up"
    Case Else
      ThisType = "All"
  End Select

  BegDate = Date2Num(fptxtBegDate.Text)
  EndDate = Date2Num(fptxtEndDate.Text)
  ThisPin = QPTrim$(fptxtPin.Text)
  RptFile$ = "TAXRPTS\REALHISTSUM.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle

  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  frmTaxShowPctComp.Label1 = "Gathering Tax Transaction Data"
  frmTaxShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
    If QPTrim$(TaxTrans.RealPin) = ThisPin Then
      If ThisClass <> 0 And TaxTrans.TranType <> ThisClass Then GoTo SkipIt
      If TaxTrans.TransDate >= BegDate And TaxTrans.TransDate <= EndDate Then
        Select Case TaxTrans.TranType
          Case 1
            ThisTransType = "Billing"
         Case 2
            ThisTransType = "Payment"
          Case 3
            ThisTransType = "Release"
          Case 4
            ThisTransType = "Interest"
          Case 5
            ThisTransType = "Unknown"
          Case 6
            ThisTransType = "Advertising Charge"
          Case 7
            ThisTransType = "Adjust Pay Down"
          Case 9
            ThisTransType = "Credit Applied at Billing"
          Case 13
            ThisTransType = "Adjust Bill Down"
          Case 14
            ThisTransType = "Adjust Bill Up"
          Case 21
            ThisTransType = "Billpay/Overpay"
'        Case 22
'          ThisTransType = "Overpayment"
          Case 24
            ThisTransType = "Adjust Bill Up Affecting Credit Balance"
          Case Else
            ThisTransType = "Unknown"
        End Select
        TCnt = TCnt + 1
        TotAmt = OldRound(TotAmt + TaxTrans.Amount)
        PrincDif = OldRound(TaxTrans.Revenue.Principle1 - TaxTrans.Revenue.Principle1Pd)
        IntDif = OldRound(TaxTrans.Revenue.Interest - TaxTrans.Revenue.InterestPd)
        AdvDif = OldRound(TaxTrans.Revenue.Collection - TaxTrans.Revenue.CollectionPd)
        LateListDif = OldRound(TaxTrans.Revenue.LateList - TaxTrans.Revenue.LateListPd)
        Opt1Dif = OldRound(TaxTrans.Revenue.RevOpt1 - TaxTrans.Revenue.RevOpt1Pd)
        Opt2Dif = OldRound(TaxTrans.Revenue.RevOpt2 - TaxTrans.Revenue.RevOpt2Pd)
        Opt3Dif = OldRound(TaxTrans.Revenue.RevOpt3 - TaxTrans.Revenue.RevOpt3Pd)
        ThisBal = OldRound(PrincDif + IntDif + AdvDif + LateListDif + Opt1Dif + Opt2Dif + Opt3Dif)
        GoSub GetOwner
        '                   0            1                    2                              3
        Print #RptHandle, Town$; dlm; Owner; dlm; MakeRegDate(TaxTrans.TransDate); dlm; ThisPin; dlm;
        '                     4                   5                         6                      7
        Print #RptHandle, ThisType; dlm; MakeRegDate(BegDate); dlm; MakeRegDate(EndDate); dlm; TaxTrans.TaxYear; dlm;
        '                        8                            9                        10
        Print #RptHandle, TaxTrans.Amount; dlm; QPTrim$(TaxTrans.Description); dlm; ThisTransType; dlm;
        '                  11          12              13
        Print #RptHandle, GAdd; dlm; CustRec; dlm; BillCustRec; dlm;

        If TaxTrans.BelongTo > 0 Then
          Get TTHandle, TaxTrans.BelongTo, TaxTrans
          GoSub GetOwner
          '                             14                               15             16          17
          Print #RptHandle, ParseBillNum(TaxTrans.Description); dlm; BillToOwner; dlm; TCnt; dlm; ThisBal
        Else
          '                 14          15             16          17
          Print #RptHandle, 0; dlm; BillToOwner; dlm; TCnt; dlm; ThisBal
        End If
      End If
    End If
SkipIt:
    frmTaxShowPctComp.ShowPctComp x, NumOfTTRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  
  Close
  If TCnt = 0 Then
    Call TaxMsg(900, "No transactions are recorded for this property.")
    Exit Sub
  End If
  Call arTaxRealHistSum.Show
  frmTaxLoadReport.Show
  
  Exit Sub

GetOwner:
  Get TCHandle, TaxTrans.CustomerRec, TaxCust
  Owner = QPTrim$(TaxCust.CustName)
  BillToOwner = Owner
  CustRec = TaxCust.PIN
  BillCustRec = CustRec
  Return

End Sub

Private Sub PrintText()
  Dim x As Long
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim ThisPin$
  Dim ThisClass As Integer
  Dim ThisType As String
  Dim BegDate As Integer
  Dim EndDate As Integer
  Dim RptFile$
  Dim RptHandle As Integer
  Dim TaxCust As TaxCustType
  Dim PrincDif As Double
  Dim IntDif As Double
  Dim AdvDif As Double
  Dim LateListDif As Double
  Dim Opt1Dif As Double
  Dim Opt2Dif As Double
  Dim Opt3Dif As Double
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim Owner$
  Dim TCnt As Long
  Dim TotAmt As Double
  Dim ThisTransType As String
  Dim BillToOwner$
  Dim CustRec As Long
  Dim BillCustRec As Long
  Dim LineCnt As Integer
  Dim MaxLines As Integer
  Dim Page As Integer
  Dim FF$
  Dim ThisBal As Double
  Dim ThisBillNum$
  Dim ThisRec As Long
  
  FF$ = Chr(12)
  MaxLines = 58
  
  If Mid(fpcmbTransType.Text, 1, 1) = "" Then
    ThisClass = CInt(Mid(fpcmbTransType.Text, 2, 1))
  Else
    ThisClass = CInt(Mid(fpcmbTransType.Text, 1, 2))
  End If

  Select Case ThisClass
    Case 1
      ThisType = "Billing"
    Case 2
      ThisType = "Payment"
    Case 3
      ThisType = "Release"
    Case 4
      ThisType = "Interest"
    Case 5
      ThisType = "Unknown"
    Case 6
      ThisType = "Advertising Charge"
    Case 7
      ThisType = "Adjust Pay Down"
    Case 13
      ThisType = "Adjust Bill Down"
    Case 14
      ThisType = "Adjust Bill Up"
    Case Else
      ThisType = "All"
  End Select

  BegDate = Date2Num(fptxtBegDate.Text)
  EndDate = Date2Num(fptxtEndDate.Text)
  ThisPin = QPTrim$(fptxtPin.Text)
  RptFile$ = "TAXRPTS\REALHIST.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle

  OpenTaxTransFile TTHandle, NumOfTTRecs
  OpenTaxCustFile TCHandle, NumOfTCRecs
  frmTaxShowPctComp.Label1 = "Gathering Tax Transaction Data"
  frmTaxShowPctComp.Show
  EnableCloseButton Me.hwnd, False
  GoSub PrintHeader
  
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
    If QPTrim$(TaxTrans.RealPin) = ThisPin Then
      If ThisClass <> 0 And TaxTrans.TranType <> ThisClass Then GoTo SkipIt
      If TaxTrans.TransDate >= BegDate And TaxTrans.TransDate <= EndDate Then
        Select Case TaxTrans.TranType
          Case 1
            ThisTransType = "Billing"
         Case 2
            ThisTransType = "Payment"
          Case 3
            ThisTransType = "Release"
          Case 4
            ThisTransType = "Interest"
          Case 5
            ThisTransType = "Unknown"
          Case 6
            ThisTransType = "Advertising Charge"
          Case 7
            ThisTransType = "Adjust Pay Down"
          Case 9
            ThisTransType = "Credit Applied at Billing"
          Case 13
            ThisTransType = "Adjust Bill Down"
          Case 14
            ThisTransType = "Adjust Bill Up"
          Case 21
            ThisTransType = "Billpay/Overpay"
'        Case 22
'          ThisTransType = "Overpayment"
          Case 24
            ThisTransType = "Adjust Bill Up Affecting Credit Balance"
          Case Else
            ThisTransType = "Unknown"
        End Select
        TCnt = TCnt + 1
        GoSub GetOwner
        PrincDif = OldRound(TaxTrans.Revenue.Principle1 - TaxTrans.Revenue.Principle1Pd)
        IntDif = OldRound(TaxTrans.Revenue.Interest - TaxTrans.Revenue.InterestPd)
        AdvDif = OldRound(TaxTrans.Revenue.Collection - TaxTrans.Revenue.CollectionPd)
        LateListDif = OldRound(TaxTrans.Revenue.LateList - TaxTrans.Revenue.LateListPd)
        Opt1Dif = OldRound(TaxTrans.Revenue.RevOpt1 - TaxTrans.Revenue.RevOpt1Pd)
        Opt2Dif = OldRound(TaxTrans.Revenue.RevOpt2 - TaxTrans.Revenue.RevOpt2Pd)
        Opt3Dif = OldRound(TaxTrans.Revenue.RevOpt3 - TaxTrans.Revenue.RevOpt3Pd)
        ThisBal = OldRound(PrincDif + IntDif + AdvDif + LateListDif + Opt1Dif + Opt2Dif + Opt3Dif)
        Print #RptHandle, MakeRegDate(TaxTrans.TransDate); Tab(12); QPTrim$(TaxTrans.Description);
        Print #RptHandle, Tab(37); Using$("###0", TaxTrans.TaxYear); Tab(45); Using$("$##,##0.00", TaxTrans.Amount); Tab(63);
        LineCnt = LineCnt + 1
        If TaxTrans.BelongTo > 0 Then
          Get TTHandle, TaxTrans.BelongTo, TaxTrans
          ThisBillNum = ParseBillNum(TaxTrans.Description)
          If IsNumeric(ThisBillNum) Then
            Print #RptHandle, Using$("######", CDbl(ThisBillNum));
          Else
            Print #RptHandle, "   " + ThisBillNum;
          End If
        Else
          Print #RptHandle, "    N/A";
        End If
      
        Get TTHandle, x, TaxTrans
        Print #RptHandle, Tab(79); ThisTransType
        If LineCnt >= MaxLines Then
          Print #RptHandle, FF$
          GoSub PrintHeader
        End If
        If ThisTransType = "Billing" Then
          Print #RptHandle, Tab(5); "Transaction Owner: " + QPTrim$(Owner); Tab(56); "Bill Balance: " + Using$("##,###,##0.00", ThisBal)
        Else
          Print #RptHandle, Tab(5); "Transaction Owner: " + QPTrim$(Owner)
        End If
        Print #RptHandle, String(89, "-")
        LineCnt = LineCnt + 2
      End If
    End If
SkipIt:
    frmTaxShowPctComp.ShowPctComp x, NumOfTTRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      EnableCloseButton Me.hwnd, True
      Exit Sub
    End If
  Next x
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  If LineCnt >= MaxLines - 1 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
    GoSub PrintRealHeader
  End If
  
  Print #RptHandle, "Total Transaction Count: " + Using("####0", TCnt)
  
  Print #RptHandle, FF$
  Close
  If TCnt = 0 Then
    Call TaxMsg(900, "No transactions are recorded for this property.")
    Exit Sub
  End If
 
  ViewPrint RptFile, "Tax Real Estate History Journal", True
  
  Exit Sub

GetOwner:
  Get TCHandle, TaxTrans.CustomerRec, TaxCust
  Owner = QPTrim$(TaxCust.CustName)
  BillToOwner = Owner
  CustRec = TaxCust.PIN
  BillCustRec = CustRec
  
  Return

PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(30); "Tax Real Estate History Journal"
  Print #RptHandle, Town$; Tab(75); "Page #: " + CStr(Page)
  Print #RptHandle, "For Pin #: " + ThisPin
  Print #RptHandle, "Address: " + GAdd
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle, "Transaction Type: " + ThisType
  Print #RptHandle, "Date Range: " + fptxtBegDate.Text + " to " + fptxtEndDate.Text
  Print #RptHandle, "Trans Date"; Tab(12); "Description"; Tab(35); "Tax Year"; Tab(47); "Trans Amt"; Tab(60); "Belongs To"; Tab(77); "Trans Type"
  Print #RptHandle, String(89, "-")
  LineCnt = 10
  
  Return

PrintRealHeader:
  Print #RptHandle, "Pin #: " + ThisPin; Tab(30); "Address: " + GAdd
  Print #RptHandle, String(89, "-")
  LineCnt = LineCnt + 2
  
  Return

End Sub
