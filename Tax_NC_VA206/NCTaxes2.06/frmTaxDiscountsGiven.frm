VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{990AFBE3-7E6C-101C-A7FD-4A79242FD97B}#3.1#0"; "Imp32x30.ocx"
Begin VB.Form frmTaxDiscountsGiven 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Discounts Applied to Payments"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   Icon            =   "frmTaxDiscountsGiven.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ImpproLib.vaImprint vaImprint1 
      Height          =   6840
      Left            =   1920
      TabIndex        =   8
      Top             =   960
      Width           =   7788
      _Version        =   196609
      _ExtentX        =   13737
      _ExtentY        =   12065
      _StockProps     =   70
      BackColor       =   13684944
      Caption         =   ""
      FrameColor      =   -2147483630
      FrameThreeDStyle=   1
      FrameWidth      =   2
      Picture         =   "frmTaxDiscountsGiven.frx":08CA
      Begin LpLib.fpCombo fpcmbPrintOrder 
         Height          =   384
         Left            =   2808
         TabIndex        =   7
         Top             =   4956
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
         ColDesigner     =   "frmTaxDiscountsGiven.frx":08E6
      End
      Begin LpLib.fpCombo fpcmbPrintOpt 
         Height          =   384
         Left            =   2808
         TabIndex        =   6
         Top             =   4368
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
         ColDesigner     =   "frmTaxDiscountsGiven.frx":0C51
      End
      Begin LpLib.fpCombo fpcmbTownship 
         Height          =   384
         Left            =   2808
         TabIndex        =   5
         Top             =   3792
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
         ColDesigner     =   "frmTaxDiscountsGiven.frx":0FBC
      End
      Begin LpLib.fpCombo fpcmbTaxYear 
         Height          =   384
         Left            =   4608
         TabIndex        =   4
         Top             =   2880
         Width           =   1380
         _Version        =   196608
         _ExtentX        =   2434
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
         ColDesigner     =   "frmTaxDiscountsGiven.frx":1327
      End
      Begin VB.OptionButton OptTaxYear 
         BackColor       =   &H00D0D0D0&
         Caption         =   "Tax Year"
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
         Left            =   1080
         TabIndex        =   3
         Top             =   2880
         Width           =   1332
      End
      Begin VB.OptionButton OptFromTo 
         BackColor       =   &H00D0D0D0&
         Caption         =   "From To"
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
         Left            =   1080
         TabIndex        =   0
         Top             =   1920
         Width           =   1332
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdExit 
         Height          =   645
         Left            =   1800
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   "Press the 'Cancel' button to exit this screen and return to the main 'Business License Reports' menu."
         Top             =   5850
         Width           =   1740
         _Version        =   131072
         _ExtentX        =   3069
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
         ButtonDesigner  =   "frmTaxDiscountsGiven.frx":1692
      End
      Begin fpBtnAtlLibCtl.fpBtn cmdProcess 
         Height          =   645
         Left            =   4155
         TabIndex        =   10
         TabStop         =   0   'False
         Tag             =   $"frmTaxDiscountsGiven.frx":1870
         Top             =   5850
         Width           =   1740
         _Version        =   131072
         _ExtentX        =   3069
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
         ButtonDesigner  =   "frmTaxDiscountsGiven.frx":191B
      End
      Begin EditLib.fpDateTime fptxtBegDate 
         Height          =   372
         Left            =   4800
         TabIndex        =   1
         Top             =   1560
         Width           =   1740
         _Version        =   196608
         _ExtentX        =   3069
         _ExtentY        =   656
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
         Height          =   372
         Left            =   4800
         TabIndex        =   2
         Top             =   2160
         Width           =   1740
         _Version        =   196608
         _ExtentX        =   3069
         _ExtentY        =   656
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
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
         AutoAdvance     =   0   'False
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
      Begin VB.Line Line4 
         BorderColor     =   &H80000004&
         BorderWidth     =   2
         X1              =   2640
         X2              =   2640
         Y1              =   3480
         Y2              =   1440
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Year:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   3360
         TabIndex        =   17
         Top             =   2952
         Width           =   1092
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000004&
         BorderWidth     =   2
         X1              =   876
         X2              =   6836
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   2052
         Left            =   888
         Top             =   1440
         Width           =   5976
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
         Left            =   3120
         TabIndex        =   16
         Top             =   2256
         Width           =   1572
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
         Left            =   2880
         TabIndex        =   15
         Top             =   1656
         Width           =   1812
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
         Left            =   1152
         TabIndex        =   14
         Top             =   4476
         Width           =   1500
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   696
         Left            =   960
         Top             =   432
         Width           =   5868
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Discounts Applied to Payments"
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
         Left            =   1200
         TabIndex        =   13
         Top             =   600
         Width           =   5412
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
         Left            =   1356
         TabIndex        =   12
         Top             =   5040
         Width           =   1308
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         FillColor       =   &H00FFFFFF&
         Height          =   1836
         Left            =   888
         Top             =   3672
         Width           =   5976
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D0D0D0&
         Caption         =   "Township:"
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
         Top             =   3912
         Width           =   1500
      End
   End
   Begin EditLib.fpText fptxtMessage 
      Height          =   372
      Left            =   420
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   8040
      Visible         =   0   'False
      Width           =   10812
      _Version        =   196608
      _ExtentX        =   19071
      _ExtentY        =   656
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8454143
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
      MaxLength       =   150
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
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   7080
      Left            =   1800
      Top             =   840
      Width           =   8052
   End
End
Attribute VB_Name = "frmTaxDiscountsGiven"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim ThisOpt$

Private Sub cmdExit_Click()
  frmTaxReportsMenu.Show
  DoEvents
  Unload Me
End Sub

Private Sub cmdProcess_Click()
  If fpcmbPrintOpt.Text = "Graphical" Then
    Call PrintGraphics
  Else
    frmTaxMsg.Label1.Caption = "Pitch 12 is recommended for this printout."
    frmTaxMsg.Label1.Top = 900
    frmTaxMsg.Show vbModal
    Call PrintText
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%x"
      Call cmdExit_Click
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
  Me.HelpContextID = hlpPayment
  Call LoadMe
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("CitiTaxes.exe terminated via menu bar on frmTaxDiscountsGiven.")
      Call Terminate
      End
    End If
  End If

End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    'Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub LoadMe()
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim TSRec As TownshipType
  Dim TSCnt As Integer
  Dim TSHandle As Integer
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim x As Long, y As Integer
  Dim Town$
  Dim YrCnt As Integer
  Dim BigYr As Integer
  Dim ThisBigYr As Integer
  Dim HoldYr As Integer
  Dim Nextx As Integer
  Dim Thisx As Integer
  Dim BHandle As Integer
  Dim CDateStr$
  
  'on error goto ERRORSTUFF
  
  If Exist("cnvtdate.dat") Then
    BHandle = FreeFile
    Open "cnvtdate.dat" For Input As BHandle
    Input #BHandle, CDateStr$
    Close BHandle
    CDateStr$ = MakeRegDate(CInt(CDateStr$))
    fptxtMessage.Visible = True
    fptxtMessage.Text = "Only transactions affecting tax bills posted on or after " + CDateStr$ + " are reported."
  End If
  
  OptFromTo.Value = 1
  fpcmbTaxYear.Enabled = False
  fptxtBegDate.Enabled = True
  fptxtBegDate = CStr(Date)
  fptxtEndDate.Enabled = True
  fptxtEndDate = CStr(Date)
  
  If Exist(TaxTownships) Then
    fpcmbTownship.Text = "All"
    fpcmbTownship.AddItem "All"
    OpenTownshipFile TSHandle, TSCnt
    For x = 1 To TSCnt
      Get TSHandle, x, TSRec
      fpcmbTownship.AddItem QPTrim$(TSRec.TownShip)
    Next x
    Close TSHandle
  Else
    fpcmbTownship.Text = "No Townships Saved"
  End If
  
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  Town = QPTrim$(TaxMasterRec.Name)
  
  fpcmbPrintOpt.Text = "Graphical"
  fpcmbPrintOpt.AddItem "Graphical"
  fpcmbPrintOpt.AddItem "Text"
  
  fpcmbPrintOrder.Text = "Name Order"
  fpcmbPrintOrder.AddItem "Name Order"
  fpcmbPrintOrder.AddItem "Acct Number Order"
  fpcmbPrintOrder.AddItem "Search Name"
  
  ThisOpt = QPTrim$(TaxMasterRec.OptSrchCust)
  If ThisOpt <> "" Then
    fpcmbPrintOrder.AddItem ThisOpt + " Order"
  End If
  
  frmTaxLoadReport.Label1.Caption = "Loading Years"
  frmTaxLoadReport.Show
  DoEvents
  ReDim Years(1 To 1) As Integer
  YrCnt = 0
  
  OpenTaxTransFile TTHandle, NumOfTTRecs
  For x = 1 To NumOfTTRecs
    Get TTHandle, x, TaxTrans
    If YrCnt = 0 Then
      If TaxTrans.TaxYear > 0 Then
        YrCnt = YrCnt + 1
        ReDim Preserve Years(1 To YrCnt) As Integer
        Years(YrCnt) = TaxTrans.TaxYear
      End If
    Else
      For y = 1 To YrCnt
        If TaxTrans.TaxYear = Years(y) Then
          Exit For
        End If
      Next y
      If y > YrCnt Then
        YrCnt = YrCnt + 1
        ReDim Preserve Years(1 To YrCnt) As Integer
        Years(YrCnt) = TaxTrans.TaxYear
      End If
    End If
  Next x
  Close TTHandle
  BigYr = 0
  For x = 1 To YrCnt
    If Years(x) > BigYr Then
      BigYr = Years(x)
    End If
  Next x
  
  Nextx = 1
  ThisBigYr = BigYr + 1
  Do While Nextx <= YrCnt
    For x = Nextx To YrCnt
      If Years(x) < ThisBigYr Then
        ThisBigYr = Years(x)
        Thisx = x
      End If
    Next x
    HoldYr = Years(Nextx)
    Years(Nextx) = Years(Thisx)
    Years(Thisx) = HoldYr
    Nextx = Nextx + 1
    ThisBigYr = BigYr + 1
  Loop
    
  fpcmbTaxYear.Text = "All"
  fpcmbTaxYear.AddItem "All"
  
  For x = YrCnt To 1 Step -1
    fpcmbTaxYear.AddItem CStr(Years(x))
  Next x
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxSeniorDscRpt", "LoadMe", Erl)
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

Private Sub fpcmbPrintOpt_KeyDown(KeyCode As Integer, Shift As Integer)
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
Private Sub fpcmbPrintOrder_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbPrintOrder.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbPrintOrder.ListIndex = -1
  End If
  If fpcmbPrintOrder.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      If fptxtBegDate.Enabled = True Then
        OptFromTo.SetFocus
      Else
        OptTaxYear.SetFocus
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

Private Sub fpcmbTaxYear_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbTaxYear.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbTaxYear.ListIndex = -1
  End If
  If fpcmbTaxYear.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      fpcmbTownship.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub fpcmbTownship_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbTownship.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbTownship.ListIndex = -1
  End If
  If fpcmbTownship.ListDown <> True Then
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

Private Sub fptxtEndDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Then
    fptxtEndDate.IncYear = 0
    fpcmbTownship.SetFocus
  ElseIf KeyCode = vbKeyUp Then
    fptxtBegDate.SetFocus
  End If
End Sub

Private Sub OptFromTo_Click()
  If OptFromTo.Value = True Then
    fptxtBegDate.Enabled = True
    fptxtEndDate.Enabled = True
    fpcmbTaxYear.Enabled = False
  Else
    fptxtBegDate.Enabled = False
    fptxtEndDate.Enabled = False
    fpcmbTaxYear.Enabled = True
  End If
End Sub

Private Sub OptTaxYear_Click()
  If OptTaxYear.Value = True Then
    fptxtBegDate.Enabled = False
    fptxtEndDate.Enabled = False
    fpcmbTaxYear.Enabled = True
  Else
    fptxtBegDate.Enabled = True
    fptxtEndDate.Enabled = True
    fpcmbTaxYear.Enabled = False
  End If

End Sub

Private Sub PrintGraphics()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim RptFile$
  Dim RptHandle As Integer
  Dim SubRptFile$
  Dim SubRptHandle As Integer
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim dlm$
  Dim InactiveFlag As Boolean
  Dim x As Long, y As Long
  Dim NextRec As Long
  Dim PropAdd$, PropTownShip$
  Dim CustRec As Long
  Dim CustName$
  Dim ThisTownship$
  Dim RealTotVal As Double
  Dim PersTotVal As Double
  Dim TotVal As Double
  Dim TotLLCnt As Long
  Dim TotRealLLCnt As Long
  Dim TotPersLLCnt As Long
  Dim ThisPersVal As Double
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim Balance As Double
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim TownName$
  Dim Charged#
  Dim Paid#
  Dim TaxYear$
  Dim ThisTaxYear As Integer
  Dim PrintCnt As Long
  Dim Revenues$
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim OptFlag As Integer
  Dim TownShip$
  Dim Beg As Integer
  Dim Last As Integer
  Dim YrCnt As Integer
  Dim TotAmt As Double
  Dim TotDiscs As Double
  
  'on error goto ERRORSTUFF
  
  TaxYear = fpcmbTaxYear.Text
  TotAmt = 0
  TotDiscs = 0
  ReDim YrAmt(1 To 1) As Double
  ReDim YrDisc(1 To 1) As Double
  ReDim Year(1 To 1) As Integer
  TaxYear = QPTrim$(fpcmbTaxYear.Text)
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  dlm$ = "~"
  TownName = QPTrim$(TaxMasterRec.Name)
'  TaxYear = CInt(fpcmbTaxYear.Text)
  TownShip = QPTrim$(fpcmbTownship.Text)
  If OptFromTo.Value = True Then
    OptFlag = 1
    Beg = Date2Num(fptxtBegDate.Text)
    Last = Date2Num(fptxtEndDate.Text)
  Else
    OptFlag = 2
    Beg = 0
    Last = 0
  End If
  If QPTrim$(fpcmbPrintOrder.Text) = "Name Order" Then
    OpenNameIdxFile IdxHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmTaxMsg.Label1.Caption = "There are no customers saved."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show
      fpcmbPrintOrder.SetFocus
      Close IdxHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get IdxHandle, x, IdxRec
      IdxArray(x) = IdxRec.CustRec
    Next x
    Close IdxHandle
    IdxFlag = True
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "Search Name" Then
    OpenSrchNameIdxFile SHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmTaxMsg.Label1.Caption = "There are no search names indexed."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show
      fpcmbPrintOrder.SetFocus
      Close SHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get SHandle, x, SearchRec
      IdxArray(x) = SearchRec.CustRec
    Next x
    Close SHandle
    IdxFlag = True
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = ThisOpt + " Order" Then
    OpenCustOptSearchFile OHandle, NumOfIdx
    If NumOfIdx = 0 Then
      Call TaxMsg(900, "There are no " + ThisOpt + "descriptions indexed.")
      Close OHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long
    
    For x = 1 To NumOfIdx
      Get OHandle, x, OptRec
      IdxArray(x) = OptRec.CustRec
    Next x
    Close OHandle
    IdxFlag = True
  End If

  RptFile$ = "TAXRPTS\DISCATPAY.RPT"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  
  SubRptFile$ = "TAXRPTS\SUBDISCATPAY.RPT"
  SubRptHandle = FreeFile
  Open SubRptFile For Output As #SubRptHandle
  
  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  If IdxFlag = True Then NumOfTCRecs = NumOfIdx
  
  OpenTaxTransFile TTHandle, NumOfTTRecs

  frmTaxShowPctComp.Label1 = "Gathering Discount Data"
  frmTaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
      CustRec = IdxArray(x)
    Else
      Get TCHandle, x, TaxCust
      CustRec = x
    End If
    If TaxCust.Deleted <> 0 Then GoTo SkipIt
    If TownShip <> "All" And TownShip <> "No Townships Saved" Then
      If QPTrim$(TaxCust.TownShip) <> TownShip Then GoTo SkipIt
    End If
    CustName = QPTrim$(TaxCust.CustName)
    NextRec = TaxCust.LastTrans
    If NextRec = 0 Then GoTo SkipIt
      Do While NextRec > 0
        Get TTHandle, NextRec, TaxTrans
        If TaxTrans.TranType <> 1 Then GoTo Another 'bill trans reflect any discounts
        'or adjustments to discounts
        If OptFlag = 1 Then
          If TaxTrans.TransDate < Beg Or TaxTrans.TransDate > Last Then GoTo Another
        ElseIf OptFlag = 2 Then
          If TaxYear = "All" Then GoTo All
          If TaxTrans.TaxYear <> CInt(TaxYear) Then GoTo Another
        End If
All:
        If TaxTrans.DiscAmt > 0 Then
          If YrCnt = 0 Then
            YrCnt = YrCnt + 1
            ReDim Preserve Year(1 To YrCnt) As Integer
            ReDim Preserve YrAmt(1 To YrCnt) As Double
            ReDim Preserve YrDisc(1 To YrCnt) As Double
            Year(YrCnt) = TaxTrans.TaxYear
            YrAmt(YrCnt) = TaxTrans.Amount
            YrDisc(YrCnt) = TaxTrans.DiscAmt
          Else
            For y = 1 To YrCnt
              If TaxTrans.TaxYear = Year(y) Then
                YrAmt(y) = OldRound(YrAmt(y) + TaxTrans.Amount)
                YrDisc(y) = OldRound(YrDisc(y) + TaxTrans.DiscAmt)
                Exit For
              End If
            Next y
            If y > YrCnt Then
              YrCnt = YrCnt + 1
              ReDim Preserve Year(1 To YrCnt) As Integer
              ReDim Preserve YrAmt(1 To YrCnt) As Double
              ReDim Preserve YrDisc(1 To YrCnt) As Double
              Year(YrCnt) = TaxTrans.TaxYear
              YrAmt(YrCnt) = TaxTrans.Amount
              YrDisc(YrCnt) = TaxTrans.DiscAmt
            End If
          End If
          GoSub PrintMe
        End If
Another:
        NextRec = TaxTrans.LastTrans
      Loop
SkipIt:
    frmTaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      Exit Sub
    End If
  Next x
  
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  If YrCnt = 0 Then
    Call TaxMsg(900, "There are no discount amounts recorded for the parameters entered.")
    Close
    Exit Sub
  End If
  GoSub PrintSub
  Close
  
  arTaxDiscsAtPayment.Show
  
  Exit Sub
  
PrintMe:
  TotAmt = OldRound(TotAmt + TaxTrans.Amount)
  TotDiscs = OldRound(TotDiscs + TaxTrans.DiscAmt)
  
  '                     0                        1                           2                3
  Print #RptHandle, TownName; dlm; MakeRegDate(TaxTrans.TransDate); dlm; CustName; dlm; TaxTrans.CustPin; dlm;
  '                    4                 5                      6                     7
  Print #RptHandle, OptFlag; dlm; TaxTrans.DiscAmt; dlm; TaxTrans.Amount; dlm; TaxTrans.TaxYear; dlm;
  '                     8             9            10                11                       12                 13
  Print #RptHandle, TownShip; dlm; TotAmt; dlm; TotDiscs; dlm; fptxtBegDate.Text; dlm; fptxtEndDate.Text; dlm; TaxYear
  Return
  
PrintSub:
  For x = 1 To YrCnt
    '                       0             1              2
    Print #SubRptHandle, Year(x); dlm; YrAmt(x); dlm; YrDisc(x)
  Next x
  
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxDiscountGiven", "PrintGraphics", Erl)
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

End Sub

Private Sub PrintText()
  Dim TaxCust As TaxCustType
  Dim TCHandle As Integer
  Dim NumOfTCRecs As Long
  Dim IdxRec As CustNameIdxType
  Dim IdxHandle As Integer
  Dim NumOfIdx As Long
  Dim RptFile$
  Dim RptHandle As Integer
  Dim IdxFlag As Boolean
  Dim SearchRec As SrchNameIdxType
  Dim SHandle As Integer
  Dim InactiveFlag As Boolean
  Dim x As Long, y As Long
  Dim NextRec As Long
  Dim PropAdd$, PropTownShip$
  Dim CustRec As Long
  Dim CustName$
  Dim ThisTownship$
  Dim RealTotVal As Double
  Dim PersTotVal As Double
  Dim TotVal As Double
  Dim TotLLCnt As Long
  Dim TotRealLLCnt As Long
  Dim TotPersLLCnt As Long
  Dim ThisPersVal As Double
  Dim TaxTrans As TaxTransactionType
  Dim TTHandle As Integer
  Dim NumOfTTRecs As Long
  Dim Balance As Double
  Dim TaxMasterRec As TaxMasterType
  Dim TMHandle As Integer
  Dim TownName$
  Dim Charged#
  Dim Paid#
  Dim TaxYear$
  Dim ThisTaxYear As Integer
  Dim PrintCnt As Long
  Dim Revenues$
  Dim OptRec As OptCustIdxType
  Dim OHandle As Integer
  Dim NumOfORecs As Long
  Dim OptFlag As Integer
  Dim TownShip$
  Dim Beg As Integer
  Dim Last As Integer
  Dim YrCnt As Integer
  Dim TotAmt As Double
  Dim TotDiscs As Double
  Dim LineCnt As Integer
  Dim MaxLines As Integer
  Dim FF$
  Dim Page As Integer
  Dim ThisName$
  Dim ThatName$
  
  'on error goto ERRORSTUFF
  
  MaxLines = 58
  FF$ = Chr(12)
  TotAmt = 0
  TotDiscs = 0
  ReDim YrAmt(1 To 1) As Double
  ReDim YrDisc(1 To 1) As Double
  ReDim Year(1 To 1) As Integer
  TaxYear = QPTrim$(fpcmbTaxYear.Text)
  OpenTaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
 
  TownName = QPTrim$(TaxMasterRec.Name)
  TaxYear = fpcmbTaxYear.Text
  TownShip = QPTrim$(fpcmbTownship.Text)
  If OptFromTo.Value = True Then
    OptFlag = 1
    Beg = Date2Num(fptxtBegDate.Text)
    Last = Date2Num(fptxtEndDate.Text)
  Else
    OptFlag = 2
    Beg = 0
    Last = 0
  End If
  If QPTrim$(fpcmbPrintOrder.Text) = "Name Order" Then
    OpenNameIdxFile IdxHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmTaxMsg.Label1.Caption = "There are no customers saved."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show
      fpcmbPrintOrder.SetFocus
      Close IdxHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get IdxHandle, x, IdxRec
      IdxArray(x) = IdxRec.CustRec
    Next x
    Close IdxHandle
    IdxFlag = True
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = "Search Name" Then
    OpenSrchNameIdxFile SHandle, NumOfIdx
    If NumOfIdx = 0 Then
      frmTaxMsg.Label1.Caption = "There are no search names indexed."
      frmTaxMsg.Label1.Top = 900
      frmTaxMsg.Show
      fpcmbPrintOrder.SetFocus
      Close SHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long

    For x = 1 To NumOfIdx
      Get SHandle, x, SearchRec
      IdxArray(x) = SearchRec.CustRec
    Next x
    Close SHandle
    IdxFlag = True
  ElseIf QPTrim$(fpcmbPrintOrder.Text) = ThisOpt + " Order" Then
    OpenCustOptSearchFile OHandle, NumOfIdx
    If NumOfIdx = 0 Then
      Call TaxMsg(900, "There are no " + ThisOpt + "descriptions indexed.")
      Close OHandle
      Exit Sub
    End If
    ReDim IdxArray(1 To NumOfIdx) As Long
    
    For x = 1 To NumOfIdx
      Get OHandle, x, OptRec
      IdxArray(x) = OptRec.CustRec
    Next x
    Close OHandle
    IdxFlag = True
  End If

  RptFile$ = "TAXRPTS\DISCATPAY.PRN"
  RptHandle = FreeFile
  Open RptFile For Output As #RptHandle
  GoSub PrintHeader
  OpenTaxCustFile TCHandle, NumOfTCRecs
  
  If IdxFlag = True Then NumOfTCRecs = NumOfIdx
  
  OpenTaxTransFile TTHandle, NumOfTTRecs

  frmTaxShowPctComp.Label1 = "Gathering Discount Data"
  frmTaxShowPctComp.Show , Me
  EnableCloseButton Me.hwnd, False
  cmdProcess.Enabled = False
  cmdExit.Enabled = False
  For x = 1 To NumOfTCRecs
    If IdxFlag = True Then
      Get TCHandle, IdxArray(x), TaxCust
      CustRec = IdxArray(x)
    Else
      Get TCHandle, x, TaxCust
      CustRec = x
    End If
    If TaxCust.Deleted <> 0 Then GoTo SkipIt
    If TownShip <> "All" And TownShip <> "No Townships Saved" Then
      If QPTrim$(TaxCust.TownShip) <> TownShip Then GoTo SkipIt
    End If
    CustName = QPTrim$(TaxCust.CustName)
    ThisName = CustName
    NextRec = TaxCust.LastTrans
    If NextRec = 0 Then GoTo SkipIt
      Do While NextRec > 0
        Get TTHandle, NextRec, TaxTrans
        If TaxTrans.TranType <> 1 Then GoTo Another 'bill trans reflect any discounts
        'or adjustments to discounts
        If OptFlag = 1 Then
          If TaxTrans.TransDate < Beg Or TaxTrans.TransDate > Last Then GoTo Another
        ElseIf OptFlag = 2 Then
          If TaxYear = "All" Then GoTo All
          If TaxTrans.TaxYear <> CInt(TaxYear) Then GoTo Another
        End If
All:
        If TaxTrans.DiscAmt > 0 Then
          If YrCnt = 0 Then
            YrCnt = YrCnt + 1
            ReDim Preserve Year(1 To YrCnt) As Integer
            ReDim Preserve YrAmt(1 To YrCnt) As Double
            ReDim Preserve YrDisc(1 To YrCnt) As Double
            Year(YrCnt) = TaxTrans.TaxYear
            YrAmt(YrCnt) = TaxTrans.Amount
            YrDisc(YrCnt) = TaxTrans.DiscAmt
          Else
            For y = 1 To YrCnt
              If TaxTrans.TaxYear = Year(y) Then
                YrAmt(y) = OldRound(YrAmt(y) + TaxTrans.Amount)
                YrDisc(y) = OldRound(YrDisc(y) + TaxTrans.DiscAmt)
                Exit For
              End If
            Next y
            If y > YrCnt Then
              YrCnt = YrCnt + 1
              ReDim Preserve Year(1 To YrCnt) As Integer
              ReDim Preserve YrAmt(1 To YrCnt) As Double
              ReDim Preserve YrDisc(1 To YrCnt) As Double
              Year(YrCnt) = TaxTrans.TaxYear
              YrAmt(YrCnt) = TaxTrans.Amount
              YrDisc(YrCnt) = TaxTrans.DiscAmt
            End If
          End If
          If ThisName <> ThatName Then
            ThatName = ThisName
            
            GoSub PrintCustHeader
          End If
          GoSub PrintMe
        End If
Another:
        NextRec = TaxTrans.LastTrans
      Loop
SkipIt:
    frmTaxShowPctComp.ShowPctComp x, NumOfTCRecs
    If frmTaxShowPctComp.Out = True Then
      Close
      frmTaxShowPctComp.Out = False
      Unload frmTaxShowPctComp
      EnableCloseButton Me.hwnd, True
      cmdProcess.Enabled = True
      cmdExit.Enabled = True
      Exit Sub
    End If
  Next x
  
  Unload frmTaxShowPctComp
  EnableCloseButton Me.hwnd, True
  cmdProcess.Enabled = True
  cmdExit.Enabled = True
  If YrCnt = 0 Then
    Call TaxMsg(900, "There are no discount amounts recorded for the parameters entered.")
    Close
    Exit Sub
  End If
  Print #RptHandle,
  Print #RptHandle,
  LineCnt = LineCnt + 2
  GoSub PrintSubHeader
  GoSub PrintSub
  Close
  
  ViewPrint RptFile, "Discounts Applied At Payment", True
  
  Exit Sub
  
PrintHeader:
  Page = Page + 1
  Print #RptHandle, Tab(30); "Discounts Applied At Payment Report"
  Print #RptHandle, TownName; Tab(65); "Page #: " + CStr(Page)
  Print #RptHandle, "Report Date: " + CStr(Date)
  Print #RptHandle, "Township:"; Tab(11); TownShip
  If OptFromTo.Value = True Then
    Print #RptHandle, "From: " + fptxtBegDate.Text + "  To: " + fptxtEndDate.Text
  Else
    Print #RptHandle, "Tax Year: " + TaxYear
  End If
  Print #RptHandle,
  Print #RptHandle, "Customer Name"; Tab(50); "Cust Pin";
  Print #RptHandle, Tab(5); "Trans Date"; Tab(22); "Tax Year"; Tab(36); "Amount Paid"; Tab(54); "Discounts Applied"
  Print #RptHandle, String(74, "-")
  LineCnt = 8
  Return
  
PrintCustHeader:
  If LineCnt > MaxLines - 4 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
  End If
  If LineCnt <> 8 Then
    Print #RptHandle, String(74, "-")
    Print #RptHandle,
    LineCnt = LineCnt + 2
  End If
  Print #RptHandle, CustName; Tab(50); Using$("#######0", TaxTrans.CustPin)
  Print #RptHandle, String(74, "-")
  LineCnt = LineCnt + 2
  Return

PrintMe:
  TotAmt = OldRound(TotAmt + TaxTrans.Amount)
  TotDiscs = OldRound(TotDiscs + TaxTrans.DiscAmt)
  If LineCnt >= MaxLines Then
    Print #RptHandle, FF$
    GoSub PrintHeader
    GoSub PrintCustHeader
  End If
  Print #RptHandle, Tab(5); MakeRegDate(TaxTrans.TransDate); Tab(24); Using$("###0", TaxTrans.TaxYear); Tab(32); Using$("$###,###,##0.00", TaxTrans.Amount); Tab(61); Using$("$##,##0.00", TaxTrans.DiscAmt)
  LineCnt = LineCnt + 1
  Return
  
PrintSubHeader:
  If LineCnt <> 7 Then
    Print #RptHandle, String(74, "-")
    LineCnt = LineCnt + 1
  End If
  Print #RptHandle, "Summary:"
  Print #RptHandle, String(74, "-")
  Print #RptHandle, Tab(22); "Tax Year"; Tab(36); "Total Paid"; Tab(54); "Total Discounts"
  Print #RptHandle, String(74, "-")
  LineCnt = LineCnt + 4
  Return
  
PrintSub:
  For x = 1 To YrCnt
    If LineCnt > MaxLines Then
      Print #RptHandle, FF$
      GoSub PrintHeader
      GoSub PrintSubHeader
    End If
    Print #RptHandle, Tab(24); Using$("###0", Year(x)); Tab(32); Using$("$###,###,##0.00", YrAmt(x)); Tab(61); Using$("$##,##0.00", YrDisc(x))
    LineCnt = LineCnt + 1
  Next x
  If LineCnt > MaxLines - 2 Then
    Print #RptHandle, FF$
    GoSub PrintHeader
    GoSub PrintSubHeader
  End If
  Print #RptHandle,
  Print #RptHandle, "Total:"; Tab(32); Using$("$###,###,##0.00", TotAmt); Tab(61); Using$("$##,##0.00", TotDiscs)
  
  Return

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmTaxDiscountGiven", "PrintGraphics", Erl)
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

End Sub
