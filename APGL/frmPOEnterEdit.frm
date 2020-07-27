VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{48932A52-981F-101B-A7FB-4A79242FD97B}#3.1#0"; "Tab32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPOEnterEdit 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Order Edit/Entry"
   ClientHeight    =   8850
   ClientLeft      =   30
   ClientTop       =   540
   ClientWidth     =   12195
   ClipControls    =   0   'False
   Icon            =   "frmPOEnterEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin TabproLib.vaTabPro vaTabPro1 
      Height          =   6060
      Left            =   1005
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   10050
      _Version        =   196609
      _ExtentX        =   17738
      _ExtentY        =   10689
      _StockProps     =   100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   9405029
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabsPerRow      =   1
      TabCount        =   1
      ThreeD          =   0   'False
      ShowFocusRect   =   0   'False
      ActiveTabBold   =   0   'False
      OffsetFromClientTop=   -1  'True
      PageMax         =   1
      PageEarMarkType =   0
      DataFormat      =   ""
      AutoSizeChildren=   3
      BookCornerGuardWidth=   90
      BookCornerGuardLength=   375
      ThreeDInnerWidthActive=   0
      DrawFocusRect   =   1
      DataField       =   ""
      TabCaption      =   "frmPOEnterEdit.frx":08CA
      PageEarMarkPictureNext=   "frmPOEnterEdit.frx":15D8
      PageEarMarkPicturePrev=   "frmPOEnterEdit.frx":15F4
      EarMarkPictureNext=   "frmPOEnterEdit.frx":1610
      EarMarkPicturePrev=   "frmPOEnterEdit.frx":162C
      Begin LpLib.fpCombo fpcboAcctNumNa 
         Height          =   405
         Left            =   -16215
         TabIndex        =   22
         Top             =   -12705
         Width           =   4200
         _Version        =   196608
         _ExtentX        =   7408
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
         Enabled         =   0   'False
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Text            =   ""
         Columns         =   4
         Sorted          =   0
         SelDrawFocusRect=   -1  'True
         ColumnSeparatorChar=   9
         ColumnSearch    =   3
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
         ScrollBarH      =   3
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
         AutoSearchFillDelay=   100
         EditMarginLeft  =   1
         EditMarginTop   =   1
         EditMarginRight =   0
         EditMarginBottom=   3
         ResizeRowToFont =   0   'False
         TextTipMultiLine=   0
         AutoMenu        =   -1  'True
         EditAlignH      =   0
         EditAlignV      =   0
         ColDesigner     =   "frmPOEnterEdit.frx":1648
      End
      Begin LpLib.fpCombo fpcboDepartment 
         Height          =   405
         Left            =   4710
         TabIndex        =   0
         Top             =   330
         Width           =   2175
         _Version        =   196608
         _ExtentX        =   3836
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
         Columns         =   3
         Sorted          =   0
         SelDrawFocusRect=   -1  'True
         ColumnSeparatorChar=   9
         ColumnSearch    =   1
         ColumnWidthScale=   2
         RowHeight       =   -1
         WrapList        =   0   'False
         WrapWidth       =   0
         AutoSearch      =   2
         SearchMethod    =   2
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
         AutoSearchFillDelay=   100
         EditMarginLeft  =   1
         EditMarginTop   =   1
         EditMarginRight =   0
         EditMarginBottom=   3
         ResizeRowToFont =   0   'False
         TextTipMultiLine=   0
         AutoMenu        =   -1  'True
         EditAlignH      =   0
         EditAlignV      =   0
         ColDesigner     =   "frmPOEnterEdit.frx":1B0B
      End
      Begin LpLib.fpList fplstVendor 
         Height          =   1485
         Left            =   555
         TabIndex        =   36
         Top             =   1560
         Width           =   4350
         _Version        =   196608
         _ExtentX        =   7673
         _ExtentY        =   2619
         TextAlias       =   ""
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
         Columns         =   0
         Sorted          =   0
         LineWidth       =   1
         SelDrawFocusRect=   -1  'True
         ColumnSeparatorChar=   9
         ColumnSearch    =   -1
         ColumnWidthScale=   2
         RowHeight       =   -1
         MultiSelect     =   0
         WrapList        =   0   'False
         WrapWidth       =   0
         SelMax          =   -1
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
         ScrollBarV      =   3
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
         DataField       =   ""
         OLEDragMode     =   0
         OLEDropMode     =   0
         Redraw          =   -1  'True
         ResizeRowToFont =   0   'False
         TextTipMultiLine=   0
         ColDesigner     =   "frmPOEnterEdit.frx":1F7E
      End
      Begin LpLib.fpCombo fpcboVendCode 
         Height          =   405
         Left            =   2430
         TabIndex        =   3
         Top             =   1065
         Width           =   2490
         _Version        =   196608
         _ExtentX        =   4392
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
         Columns         =   2
         Sorted          =   0
         SelDrawFocusRect=   -1  'True
         ColumnSeparatorChar=   9
         ColumnSearch    =   0
         ColumnWidthScale=   2
         RowHeight       =   -1
         WrapList        =   0   'False
         WrapWidth       =   0
         AutoSearch      =   2
         SearchMethod    =   1
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
         ScrollBarH      =   3
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
         AutoSearchFillDelay=   100
         EditMarginLeft  =   5
         EditMarginTop   =   1
         EditMarginRight =   0
         EditMarginBottom=   3
         ResizeRowToFont =   0   'False
         TextTipMultiLine=   0
         AutoMenu        =   -1  'True
         EditAlignH      =   0
         EditAlignV      =   0
         ColDesigner     =   "frmPOEnterEdit.frx":2302
      End
      Begin EditLib.fpCurrency fpBudget 
         Height          =   300
         Left            =   -13392
         TabIndex        =   63
         Top             =   -12624
         Width           =   1380
         _Version        =   196608
         _ExtentX        =   2434
         _ExtentY        =   529
         Enabled         =   0   'False
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483637
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   0
         ThreeDInsideHighlightColor=   -2147483637
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   0
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
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
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
         ControlType     =   1
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
         ThreeDFrameColor=   -2147483637
         Appearance      =   0
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDoubleSingle txtQty 
         Height          =   372
         Left            =   -13152
         TabIndex        =   20
         Top             =   -12696
         Width           =   1140
         _Version        =   196608
         _ExtentX        =   2011
         _ExtentY        =   656
         Enabled         =   0   'False
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
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         CaretInsert     =   2
         CaretOverWrite  =   2
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
         OnFocusPosition =   1
         ControlType     =   0
         Text            =   "0"
         DecimalPlaces   =   3
         DecimalPoint    =   ""
         FixedPoint      =   0   'False
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "-9000000000"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   0
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpText txtDesc 
         Height          =   372
         Left            =   -16068
         TabIndex        =   19
         Top             =   -12696
         Width           =   4056
         _Version        =   196608
         _ExtentX        =   7154
         _ExtentY        =   656
         Enabled         =   0   'False
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
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   2
         CaretOverWrite  =   2
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
         OnFocusNoSelect =   -1  'True
         OnFocusPosition =   1
         ControlType     =   0
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   40
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
      Begin EditLib.fpText txtStock 
         Height          =   372
         Left            =   -13176
         TabIndex        =   18
         Top             =   -12696
         Width           =   1164
         _Version        =   196608
         _ExtentX        =   2053
         _ExtentY        =   656
         Enabled         =   0   'False
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
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   2
         CaretOverWrite  =   2
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
         OnFocusNoSelect =   -1  'True
         OnFocusPosition =   1
         ControlType     =   0
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   8
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
      Begin EditLib.fpText txtAddinst3 
         Height          =   375
         Left            =   5775
         TabIndex        =   16
         Top             =   4890
         Width           =   2940
         _Version        =   196608
         _ExtentX        =   5186
         _ExtentY        =   656
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
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   2
         CaretOverWrite  =   2
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
         OnFocusNoSelect =   -1  'True
         OnFocusPosition =   1
         ControlType     =   0
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   30
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
      Begin EditLib.fpText txtAddinst2 
         Height          =   375
         Left            =   5775
         TabIndex        =   15
         Top             =   4470
         Width           =   2940
         _Version        =   196608
         _ExtentX        =   5186
         _ExtentY        =   656
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
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   2
         CaretOverWrite  =   2
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
         OnFocusNoSelect =   -1  'True
         OnFocusPosition =   1
         ControlType     =   0
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   30
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
      Begin EditLib.fpText txtAddinst1 
         Height          =   375
         Left            =   5775
         TabIndex        =   14
         Top             =   4035
         Width           =   2940
         _Version        =   196608
         _ExtentX        =   5186
         _ExtentY        =   656
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
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   2
         CaretOverWrite  =   2
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
         OnFocusNoSelect =   -1  'True
         OnFocusPosition =   1
         ControlType     =   0
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   30
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
      Begin EditLib.fpText txtShipTo5 
         Height          =   390
         Left            =   5775
         TabIndex        =   13
         Top             =   3045
         Width           =   3780
         _Version        =   196608
         _ExtentX        =   6667
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
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   2
         CaretOverWrite  =   2
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
         OnFocusNoSelect =   -1  'True
         OnFocusPosition =   1
         ControlType     =   0
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   30
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
      Begin EditLib.fpText txtShipTo4 
         Height          =   390
         Left            =   5775
         TabIndex        =   12
         Top             =   2595
         Width           =   3780
         _Version        =   196608
         _ExtentX        =   6667
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
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   2
         CaretOverWrite  =   2
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
         OnFocusNoSelect =   -1  'True
         OnFocusPosition =   1
         ControlType     =   0
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   30
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
      Begin EditLib.fpText txtShipTo3 
         Height          =   390
         Left            =   5775
         TabIndex        =   11
         Top             =   2145
         Width           =   3780
         _Version        =   196608
         _ExtentX        =   6667
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
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   2
         CaretOverWrite  =   2
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
         OnFocusNoSelect =   -1  'True
         OnFocusPosition =   1
         ControlType     =   0
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   30
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
      Begin EditLib.fpText txtShipTo2 
         Height          =   390
         Left            =   5775
         TabIndex        =   10
         Top             =   1680
         Width           =   3780
         _Version        =   196608
         _ExtentX        =   6667
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
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   2
         CaretOverWrite  =   2
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
         OnFocusNoSelect =   -1  'True
         OnFocusPosition =   1
         ControlType     =   0
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   30
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
      Begin EditLib.fpText txtShipTo1 
         Height          =   390
         Left            =   5775
         TabIndex        =   9
         Top             =   1230
         Width           =   3780
         _Version        =   196608
         _ExtentX        =   6667
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
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   2
         CaretOverWrite  =   2
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
         OnFocusNoSelect =   -1  'True
         OnFocusPosition =   1
         ControlType     =   0
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   30
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
      Begin EditLib.fpText txtShipOn 
         Height          =   390
         Left            =   2085
         TabIndex        =   7
         Top             =   4725
         Width           =   2850
         _Version        =   196608
         _ExtentX        =   5016
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
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   2
         CaretOverWrite  =   2
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
         OnFocusNoSelect =   -1  'True
         OnFocusPosition =   1
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
      Begin EditLib.fpText txtTerms 
         Height          =   390
         Left            =   2085
         TabIndex        =   6
         Top             =   4245
         Width           =   2850
         _Version        =   196608
         _ExtentX        =   5016
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
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   2
         CaretOverWrite  =   2
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
         OnFocusNoSelect =   -1  'True
         OnFocusPosition =   1
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
      Begin EditLib.fpText txtShipVia 
         Height          =   390
         Left            =   2085
         TabIndex        =   5
         Top             =   3765
         Width           =   2850
         _Version        =   196608
         _ExtentX        =   5016
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
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   2
         CaretOverWrite  =   2
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
         OnFocusNoSelect =   -1  'True
         OnFocusPosition =   1
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
      Begin EditLib.fpText txtFOB 
         Height          =   390
         Left            =   2085
         TabIndex        =   4
         Top             =   3285
         Width           =   2850
         _Version        =   196608
         _ExtentX        =   5016
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
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   2
         CaretOverWrite  =   2
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
         OnFocusNoSelect =   -1  'True
         OnFocusPosition =   1
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
      Begin VB.CommandButton cmdAddDist 
         Caption         =   "F9 &Add Distribution"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   576
         Left            =   -13344
         TabIndex        =   23
         Top             =   -12900
         Width           =   1332
      End
      Begin VB.CommandButton cmdPage1 
         Caption         =   "<-  &Page 1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   -13680
         TabIndex        =   32
         Top             =   -12636
         Width           =   1668
      End
      Begin VB.CommandButton cmdDist 
         BackColor       =   &H00D0D0D0&
         Caption         =   "Distri&butions ->"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7890
         MaskColor       =   &H00D0D0D0&
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   5595
         Width           =   1860
      End
      Begin VB.TextBox txtPONumber 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   330
         Width           =   1230
      End
      Begin EditLib.fpCurrency txtTotPOAmt 
         Height          =   375
         Left            =   2085
         TabIndex        =   8
         Top             =   5415
         Width           =   1860
         _Version        =   196608
         _ExtentX        =   3281
         _ExtentY        =   656
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
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         CaretInsert     =   2
         CaretOverWrite  =   2
         UserEntry       =   0
         HideSelection   =   0   'False
         InvalidColor    =   -2147483643
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483643
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   1
         ControlType     =   0
         Text            =   "$0.00"
         CurrencyDecimalPlaces=   2
         CurrencyNegFormat=   0
         CurrencyPlacement=   0
         CurrencySymbol  =   "$"
         DecimalPoint    =   ""
         FixedPoint      =   -1  'True
         LeadZero        =   0
         MaxValue        =   "999999999.99"
         MinValue        =   "-999999999.99"
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   0
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDateTime txtDate 
         Height          =   375
         Left            =   7860
         TabIndex        =   2
         Top             =   330
         Width           =   1695
         _Version        =   196608
         _ExtentX        =   2984
         _ExtentY        =   656
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
         ButtonStyle     =   2
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         CaretInsert     =   2
         CaretOverWrite  =   2
         UserEntry       =   0
         HideSelection   =   0   'False
         InvalidColor    =   -2147483643
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483643
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   -1  'True
         OnFocusPosition =   1
         ControlType     =   0
         Text            =   "10/03/2001"
         DateCalcMethod  =   0
         DateTimeFormat  =   5
         UserDefinedFormat=   "mm/dd/yyyy"
         DateMax         =   "20350101"
         DateMin         =   "19800101"
         TimeMax         =   "000000"
         TimeMin         =   "000000"
         TimeString1159  =   ""
         TimeString2359  =   ""
         DateDefault     =   "20010101"
         TimeDefault     =   "000000"
         TimeStyle       =   0
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
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
         ButtonColor     =   -2147483637
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3180
         Left            =   -21924
         TabIndex        =   24
         Top             =   -15504
         Width           =   9912
         _Version        =   196613
         _ExtentX        =   17621
         _ExtentY        =   6138
         _StockProps     =   64
         Enabled         =   0   'False
         BackColorStyle  =   1
         ButtonDrawMode  =   4
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   7
         MaxRows         =   36
         OperationMode   =   1
         ScrollBars      =   2
         SelectBlockOptions=   0
         SpreadDesigner  =   "frmPOEnterEdit.frx":2749
         VisibleCols     =   6
         VisibleRows     =   10
         ScrollBarTrack  =   1
      End
      Begin EditLib.fpCurrency txtTotDistAmt 
         Height          =   348
         Left            =   -13872
         TabIndex        =   37
         Top             =   -12672
         Width           =   1860
         _Version        =   196608
         _ExtentX        =   3281
         _ExtentY        =   614
         Enabled         =   0   'False
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
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         CaretInsert     =   2
         CaretOverWrite  =   2
         UserEntry       =   0
         HideSelection   =   0   'False
         InvalidColor    =   -2147483643
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483643
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   1
         ControlType     =   1
         Text            =   "$0.00"
         CurrencyDecimalPlaces=   2
         CurrencyNegFormat=   0
         CurrencyPlacement=   0
         CurrencySymbol  =   "$"
         DecimalPoint    =   ""
         FixedPoint      =   -1  'True
         LeadZero        =   0
         MaxValue        =   "999999999.99"
         MinValue        =   "-999999999.99"
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   0
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpCurrency txtPrice 
         Height          =   348
         Left            =   -13536
         TabIndex        =   21
         Top             =   -12672
         Width           =   1524
         _Version        =   196608
         _ExtentX        =   2688
         _ExtentY        =   614
         Enabled         =   0   'False
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
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         CaretInsert     =   2
         CaretOverWrite  =   2
         UserEntry       =   0
         HideSelection   =   0   'False
         InvalidColor    =   -2147483643
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483643
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   1
         ControlType     =   0
         Text            =   "$0"
         CurrencyDecimalPlaces=   4
         CurrencyNegFormat=   0
         CurrencyPlacement=   0
         CurrencySymbol  =   "$"
         DecimalPoint    =   ""
         FixedPoint      =   0   'False
         LeadZero        =   0
         MaxValue        =   "999999999"
         MinValue        =   "-999999999"
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   0
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpCurrency txtTot2 
         Height          =   324
         Left            =   -13728
         TabIndex        =   47
         Top             =   -12648
         Width           =   1716
         _Version        =   196608
         _ExtentX        =   3027
         _ExtentY        =   572
         Enabled         =   0   'False
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   10.5
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
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
         CaretInsert     =   2
         CaretOverWrite  =   2
         UserEntry       =   0
         HideSelection   =   0   'False
         InvalidColor    =   -2147483643
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483643
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   1
         ControlType     =   1
         Text            =   "$0.00"
         CurrencyDecimalPlaces=   2
         CurrencyNegFormat=   0
         CurrencyPlacement=   0
         CurrencySymbol  =   "$"
         DecimalPoint    =   ""
         FixedPoint      =   -1  'True
         LeadZero        =   0
         MaxValue        =   "999999999.99"
         MinValue        =   "-999999999.99"
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   0
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpCurrency fptotPost 
         Height          =   300
         Left            =   -13392
         TabIndex        =   60
         Top             =   -12624
         Width           =   1380
         _Version        =   196608
         _ExtentX        =   2434
         _ExtentY        =   529
         Enabled         =   0   'False
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483637
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   0
         ThreeDInsideHighlightColor=   -2147483637
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   0
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
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
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
         ControlType     =   1
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
         ThreeDFrameColor=   -2147483637
         Appearance      =   0
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpCurrency fpEncumb 
         Height          =   300
         Left            =   -13392
         TabIndex        =   61
         Top             =   -12624
         Width           =   1380
         _Version        =   196608
         _ExtentX        =   2434
         _ExtentY        =   529
         Enabled         =   0   'False
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483637
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   0
         ThreeDInsideHighlightColor=   -2147483637
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   0
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
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
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
         ControlType     =   1
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
         ThreeDFrameColor=   -2147483637
         Appearance      =   0
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpCurrency fpYTD 
         Height          =   300
         Left            =   -13392
         TabIndex        =   62
         Top             =   -12624
         Width           =   1380
         _Version        =   196608
         _ExtentX        =   2434
         _ExtentY        =   529
         Enabled         =   0   'False
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483637
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   0
         ThreeDInsideHighlightColor=   -2147483637
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   0
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
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
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
         ControlType     =   1
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
         ThreeDFrameColor=   -2147483637
         Appearance      =   0
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Line Line2 
         X1              =   -16200
         X2              =   -12024
         Y1              =   -12336
         Y2              =   -12336
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Balance Including This PO -"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   264
         Index           =   1
         Left            =   -14280
         TabIndex        =   68
         Top             =   -12588
         Width           =   2268
      End
      Begin VB.Label Label13b 
         Alignment       =   2  'Center
         Caption         =   "Encumbered"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   276
         Index           =   0
         Left            =   -13392
         TabIndex        =   67
         Top             =   -12600
         Width           =   1380
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "YTD"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   228
         Left            =   -13392
         TabIndex        =   66
         Top             =   -12552
         Width           =   1380
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "{Posted Amts}...."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   276
         Left            =   -13464
         TabIndex        =   65
         Top             =   -12600
         Width           =   1452
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "Budget"
         Enabled         =   0   'False
         Height          =   228
         Left            =   -13392
         TabIndex        =   64
         Top             =   -12552
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PO Number"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   330
         Left            =   270
         TabIndex        =   59
         Top             =   360
         Width           =   1425
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   390
         TabIndex        =   58
         Top             =   1035
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000009&
         BorderWidth     =   3
         X1              =   270
         X2              =   11025
         Y1              =   870
         Y2              =   870
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dept No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   330
         Index           =   0
         Left            =   3630
         TabIndex        =   57
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ship To Information:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   330
         Left            =   5265
         TabIndex        =   56
         Top             =   945
         Width           =   2295
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Index           =   1
         Left            =   7065
         TabIndex        =   55
         Top             =   345
         Width           =   660
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "FOB Point"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   345
         Left            =   870
         TabIndex        =   54
         Top             =   3360
         Width           =   1140
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Additional Instructions:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   330
         Left            =   5370
         TabIndex        =   53
         Top             =   3720
         Width           =   2625
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ship On"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   420
         Left            =   1065
         TabIndex        =   52
         Top             =   4740
         Width           =   945
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ship Via"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   345
         Left            =   945
         TabIndex        =   51
         Top             =   3810
         Width           =   1050
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Terms"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   345
         Left            =   1065
         TabIndex        =   50
         Top             =   4275
         Width           =   930
      End
      Begin VB.Label Label2b 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total PO:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   49
         Top             =   5445
         Width           =   1065
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000004&
         Height          =   732
         Left            =   -13992
         Top             =   -13056
         Width           =   1980
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Total PO:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   288
         Index           =   0
         Left            =   -12960
         TabIndex        =   48
         Top             =   -12612
         Width           =   948
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Dbl-Click Or F11 On Row To Edit."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   444
         Left            =   -13920
         TabIndex        =   46
         Top             =   -12768
         Width           =   1908
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "G/L Account"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   348
         Index           =   2
         Left            =   -13440
         TabIndex        =   45
         Top             =   -12672
         Width           =   1428
      End
      Begin VB.Label Label2b 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   348
         Index           =   0
         Left            =   -12600
         TabIndex        =   44
         Top             =   -12672
         Width           =   588
      End
      Begin VB.Label Label3b 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desc"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   348
         Index           =   1
         Left            =   -12744
         TabIndex        =   43
         Top             =   -12672
         Width           =   732
      End
      Begin VB.Label Label3b 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   348
         Index           =   0
         Left            =   -12504
         TabIndex        =   42
         Top             =   -12672
         Width           =   492
      End
      Begin VB.Label Label3b 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stock #"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   348
         Index           =   2
         Left            =   -12936
         TabIndex        =   41
         Top             =   -12672
         Width           =   924
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Distributions:"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   288
         Left            =   -13896
         TabIndex        =   38
         Top             =   -12612
         Width           =   1884
      End
   End
   Begin VB.CommandButton cmdList 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F5 &List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   5430
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   7512
      Width           =   1332
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F2 &New"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2276
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   7512
      Width           =   1332
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F4 &Edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3853
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   7512
      Width           =   1332
   End
   Begin VB.CommandButton cmdDelDist 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F6 Del D&ist"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   7007
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   7512
      Width           =   1332
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F3 &Delete"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   8584
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   7512
      Width           =   1332
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00D0D0D0&
      Caption         =   "Esc E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   10161
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   7512
      Width           =   1332
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F10 &Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   699
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   7512
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   33
      Top             =   8496
      Width           =   12192
      _ExtentX        =   21511
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7117
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "8:48 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7117
            TextSave        =   "6/4/2018"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblNew 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "New Purchase Order"
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
      Left            =   888
      TabIndex        =   40
      Top             =   984
      Visible         =   0   'False
      Width           =   2964
   End
   Begin VB.Label lblEdit 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "Edit Purchase Order"
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
      Left            =   8232
      TabIndex        =   39
      Top             =   984
      Visible         =   0   'False
      Width           =   3132
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   636
      Left            =   2580
      Top             =   288
      Width           =   7020
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter/Edit Purchase Orders"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4092
      TabIndex        =   34
      Top             =   432
      Width           =   4020
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00D0D0D0&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   780
      Left            =   2592
      Top             =   168
      Width           =   7020
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuPrnScn 
         Caption         =   "Prin&t Screen"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmPOEnterEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim LPDate As Integer, HPDate As Integer
Dim POControl As POControlRecType
Dim POEdit As POFORMRecType2
Dim Vendor As VendorRecType
Dim VendorIdx As VendorIdxRecType
Dim GLAcct As GLAcctRecType
Dim FY1BegDate As Integer, FY1EndDate As Integer, FY2BegDate As Integer, FY2EndDate As Integer
Private Temp_Class As Resize_Class
Dim EMode As Boolean, RecNum As Integer, RecLok As Boolean
Dim OldRec As Integer, skip As Boolean
'**************************
'Use Emode to determine if New record or Editing, so true if editing.
'RecNum passed from Listing to load chosen record to form
'******************************
Private Sub cmdDist_Click()
  If Okgopg2 = True Then
    vaTabPro1.ActivePage = 1
    cmdDelDist.Enabled = True
    txtStock.SetFocus
  Else
    MsgBox "You Must Enter The Vendor Code and Purchase Order Amount Before Entering the Distributions.", vbOKOnly, "PO Entry"
  End If
End Sub
Private Sub GetBudInfo()
  Dim ac As Integer, TD As Integer, ActiveYear As Integer
  If CheckValDate(QPTrim(txtDate)) = True Then
    TD = DateDiff("d", "12/31/1979", txtDate)
    fpcboAcctNumNa.col = 0
    ac = fpcboAcctNumNa.ColText
    If TD >= FY2BegDate Then
      ActiveYear = 2
    Else
      ActiveYear = 1
    End If
  If GetAcctType$(ac) = "E" Then
    Call GetAcctAmts(ac, ActiveYear)
  Else
    fpBudget = 0
    fpYTD = 0
    fpEncumb = 0
    fptotPost = 0
  End If
End If
End Sub
Private Sub GetTot4Bud(ac As Integer)
  Dim cntB As Integer, TmpAmt As Double, TmpTot As Double
  TmpAmt = 0
  TmpTot = 0
 ' If vaSpread1.DataRowCnt > 0 Then
  For cntB = 1 To vaSpread1.DataRowCnt
    vaSpread1.Row = cntB
    vaSpread1.col = 1
    If vaSpread1.Text = ac Then
      vaSpread1.col = 6
      TmpAmt = Round(TmpAmt + vaSpread1.Text)
    End If
  Next
  TmpTot = Round(Val(txtQty * txtPrice.DoubleValue))
  TmpTot = Round(TmpTot + TmpAmt)
  fptotPost = Round(fpBudget.DoubleValue - (fpYTD.DoubleValue + fpEncumb.DoubleValue + TmpTot))
  If fptotPost.DoubleValue < 0 Then
    fptotPost.BackColor = &HC0&
  Else
    fptotPost.BackColor = &H8000000F
  End If
      
End Sub

Private Sub GetAcctAmts(ac As Integer, ActiveYear As Integer)
  Dim AcctFileNum As Integer, NumAccts As Integer
  OpenAcctFile AcctFileNum
  NumAccts = LOF(AcctFileNum) / Len(GLAcct)
  Get AcctFileNum, ac, GLAcct
    If ActiveYear = 1 Then
      fpBudget = GLAcct.Bgt
    Else
      fpBudget = GLAcct.NYApp
    End If
    fpYTD = GLAcct.YTD
    fpEncumb = GLAcct.Encumb
  Close AcctFileNum
  GetTot4Bud ac
End Sub
Private Sub ClearBuds()
  fpcboAcctNumNa.ListIndex = -1
  fpBudget = 0
  fpYTD = 0
  fpEncumb = 0
  fptotPost = 0
  fptotPost.BackColor = &H8000000F

End Sub


Private Sub txtPrice_Change()
  If fpcboAcctNumNa.ListIndex <> -1 And skip = False Then
    GetBudInfo
  End If
End Sub

Private Sub fpcboAcctNumNa_Click()
  If fpcboAcctNumNa.ListIndex <> -1 Then
    GetBudInfo
  End If
End Sub
Private Sub txtQty_Change()
  If fpcboAcctNumNa.ListIndex <> -1 And skip = False Then
    GetBudInfo
  End If
End Sub
Private Sub cmdPage1_Click()
  vaTabPro1.ActivePage = 0
  cmdDelDist.Enabled = False
End Sub
Private Function Okgopg2()
  If fpcboVendCode.ListIndex <> -1 Then
    If txtTotPOAmt <> 0 Then
      Okgopg2 = True
    Else
      Okgopg2 = False
      txtTotPOAmt.SetFocus
    End If
  Else
    Okgopg2 = False
    fpcboVendCode.SetFocus
  End If
End Function
Public Sub FirstOpenPOs()
  If RecLok = True Then
    frmPOListing.Show 1, frmPOEnterEdit
  End If
End Sub


Private Sub fpcboVendCode_Click()
    If fpcboVendCode.ListIndex <> -1 Then
      LoadUp
    End If
End Sub
Private Function SetScreen()
  If EMode = False Then  'This is in New Mode
    cmdNew.Enabled = False
    cmdEdit.Enabled = True
    cmdDelete.Enabled = False
    lblNew.Visible = True
    lblEdit.Visible = False
  Else               'This is in Edit Mode
    cmdNew.Enabled = True
    cmdEdit.Enabled = False
    cmdDelete.Enabled = True
    lblNew.Visible = False
    lblEdit.Visible = True
  End If
  
End Function
Private Sub cmdDelete_Click()
  Dim POEditFile As Integer, NumEdTrans As Integer
  Dim POBusy As Boolean
  POBusy = False
  If Exist("APPED.DAT") Then POBusy = GetAttr("apped.dat") And vbReadOnly
  If Not POBusy Then
    If EMode = True Then
      If MsgBox("Are you sure you wish to delete this entry?", vbYesNo, "Delete GJEntry") = vbYes Then
        OpenPOEditFile POEditFile, NumEdTrans
        POEdit.Deleted = -1
        POEdit.LOCKED = False
        Put POEditFile, RecNum, POEdit
        Close POEditFile
        POEdit.Deleted = 0
        ClearScn
      Else
        vaTabPro1.ActivePage = 0
        fpcboDepartment.SetFocus
      End If
    Else
      ClearScn
    End If
  Else
    MsgBox "Posting Is In Progress, Editing Not Allowed At This Time.", vbOKOnly, "Canceled"
    frmPOProcessMenu.Show
    Unload frmPOEnterEdit
  End If
End Sub
Private Sub cmdDelDist_Click()
  If vaSpread1.ActiveRow > 0 Then
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.col = 1
    If vaSpread1.Text <> "" Then
      If MsgBox("You Wish to Delete this Distribution?", vbYesNo, "Delete Distribution") = vbYes Then
        vaSpread1.col = 6
        txtTotDistAmt = (txtTotDistAmt.DoubleValue - vaSpread1.Text)
        
        vaSpread1.DeleteRows vaSpread1.Row, 1
        txtStock.SetFocus
      End If
    End If
  End If

End Sub
Public Sub ClearFields()
  fpcboDepartment.ListIndex = -1
  fpcboVendCode.ListIndex = -1
  fplstVendor.Clear
  txtFOB = ""
  txtShipVia = ""
  txtTerms = ""
  txtShipOn = ""
  txtTotPOAmt = 0
  txtShipTo1 = ""
  txtShipTo2 = ""
  txtShipTo3 = ""
  txtShipTo4 = ""
  txtShipTo5 = ""
  txtAddinst1 = ""
  txtAddinst2 = ""
  txtAddinst3 = ""
  txtStock = ""
  txtDesc = ""
  txtQty = ""
  txtPrice = 0
  fpcboAcctNumNa.ListIndex = -1
  ClearBuds
'*****Clear data in the spreadsheet
  vaSpread1.ClearRange 1, 1, 7, 36, True
  txtTotDistAmt = 0
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If Changed = False Then
      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
        Cancel = True
      Else
        Undolok RecNum
        MainLog "Close AP"
        ClearInUse PWcnt
      End If
    Else
      If MsgBox("Changes Have Been Made to the Current Record." & Chr(13) & Chr(13) & "                          Select OK to Abandon and Close Program," & Chr(13) & Chr(13) & "       or Cancel to Remain on Entry/Edit Screen.", vbOKCancel, "Abandon Changes?") = vbOK Then
        Undolok RecNum
        MainLog "Close AP"
        ClearInUse PWcnt
      Else
        Cancel = True
      End If
    End If
  End If
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  GetPostDates LPDate, HPDate  'In Main Module to get dates from setup
  GetFYDates FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate
  StatusBar1.Panels.Item(1).Text = GLUserName
  Me.HelpContextID = hlpEnterPO
  DeptList fpcboDepartment
  fpcboDepartment.RemoveItem 0
  VendCodeList fpcboVendCode
  Fixspread
  FillAcctNumName fpcboAcctNumNa
  vaTabPro1.ActivePage = 0
  EdorNewEntry
  cmdDelDist.Enabled = False
End Sub
Private Sub LoadUp()
  Dim VendorFile As Integer, NumVRecs As Integer, VRecNum As Integer
  Dim Last As Integer, cnt As Integer, Dcnt As Integer, TmpAcct As Integer
  fpcboVendCode.col = 1
  VRecNum = fpcboVendCode.ColText
  fplstVendor.Clear
  If VRecNum > 0 Then
    OpenVendorFile VendorFile, NumVRecs
    Get VendorFile, VRecNum, Vendor
    fplstVendor.Row = -1
    fplstVendor.InsertRow = Vendor.VNAME
    fplstVendor.InsertRow = Vendor.Addr1
    fplstVendor.InsertRow = Vendor.Addr2
    fplstVendor.InsertRow = QPTrim$(Vendor.City) + ", " + Vendor.State + " " + Vendor.Zip
  End If
 Close
End Sub
Public Sub Rec2Form(RecordNumber)
  Dim POEditFile As Integer, NumEdTrans As Integer
  Dim cnt As Integer, TmpAcct As Integer, tempcode As Integer
  Dim POBusy As Boolean
  POBusy = False
  If Exist("APPED.DAT") Then POBusy = GetAttr("apped.dat") And vbReadOnly
  If Not POBusy Then
    OldRec = RecNum
    RecNum = RecordNumber
    OpenPOEditFile POEditFile, NumEdTrans
    Get POEditFile, RecordNumber, POEdit
    If POEdit.LOCKED = False Then
      POEdit.LOCKED = True
      Put POEditFile, RecNum, POEdit
      txtPONumber = POEdit.PONum
      fpcboDepartment.col = 1
      fpcboDepartment.SearchText = QPTrim(POEdit.REQNUM)
      fpcboDepartment.Action = 0
      If fpcboDepartment.SearchIndex <> -1 Then
        fpcboDepartment.ListIndex = fpcboDepartment.SearchIndex
      End If
      'fpcboDepartment.Text = POEdit.REQNUM
      txtDate.Text = Format(DateAdd("d", (POEdit.PODATE), "12-31-1979"), "mm/dd/yyyy")
  
      fpcboVendCode.col = 1
      fpcboVendCode.SearchText = QPTrim(POEdit.VNDRCODE)
      fpcboVendCode.Action = 0
      If fpcboVendCode.SearchIndex <> -1 Then
        fpcboVendCode.ListIndex = fpcboVendCode.SearchIndex
        fpcboVendCode.col = 0
        fpcboVendCode.ColText = POEdit.VNDRCODE
        fpcboVendCode.col = 1
        fpcboVendCode.ColText = POEdit.VNDRREC
      Else
        MsgBox "Invalid Vendor Code, Please re-enter.", vbOKOnly, "Invalid Vendor"
      End If
      LoadUp
      txtFOB = Trim(POEdit.FOB)
      txtShipVia = Trim(POEdit.Shipvia)
      txtTerms = POEdit.Terms
      txtShipOn = POEdit.SHIPON
      txtTotPOAmt = POEdit.POAmt
      txtShipTo1 = POEdit.SHPLINE1
      txtShipTo2 = POEdit.SHPLINE2
      txtShipTo3 = POEdit.SHPLINE3
      txtShipTo4 = POEdit.SHPLINE4
      txtShipTo5 = POEdit.SHPLINE5
      txtAddinst1 = POEdit.Addinst1
      txtAddinst2 = POEdit.Addinst2
      txtAddinst3 = POEdit.Addinst3
      If UBound(POEdit.ITEMS) > 0 Then
        txtTotDistAmt = 0
        For cnt = 1 To UBound(POEdit.ITEMS)
          TmpAcct = AcctFind(QPTrim(POEdit.ITEMS(cnt).ACCTNO))
            If TmpAcct > 0 Then
              vaSpread1.Row = vaSpread1.DataRowCnt + 1
              vaSpread1.col = 1
              vaSpread1.Text = POEdit.ITEMS(cnt).AcctRec
              vaSpread1.col = 2
              vaSpread1.Text = QPTrim$(POEdit.ITEMS(cnt).STKNO)
              vaSpread1.col = 3
              vaSpread1.Text = QPTrim$(POEdit.ITEMS(cnt).Desc)
              vaSpread1.col = 4
              vaSpread1.Text = Trim(POEdit.ITEMS(cnt).QUAN)
              vaSpread1.col = 5
              vaSpread1.Text = Trim(POEdit.ITEMS(cnt).PRICE)
              vaSpread1.col = 6
              vaSpread1.Text = Trim(POEdit.ITEMS(cnt).EXT)
              vaSpread1.col = 7
              vaSpread1.Text = (QPTrim(POEdit.ITEMS(cnt).ACCTNO))
              txtTotDistAmt = (txtTotDistAmt.DoubleValue + POEdit.ITEMS(cnt).EXT)
            'Else
              'MsgBox "Account could not be found and will not be loaded.", vbOKOnly, "Invalid Account"
            End If
       Next
      End If
      EMode = True
      cmdDelete.Enabled = True
    Else
      MsgBox "Record Is Being Edited By Another User.", vbOKOnly, "Record Unavailable"
      RecNum = OldRec
      Close POEditFile
    End If
  Else
    MsgBox "Posting In Progress, Editing May Not Continue At This Time.", vbOKOnly, "Canceled"
    frmPOProcessMenu.Show
    Unload frmPOEnterEdit
  End If
End Sub

Private Sub ClearScn()
    EMode = False
    SetScreen
    LoadControl
    txtDate.Text = Format(Now, "mm/dd/yyyy")
    
    fpcboVendCode.ListIndex = -1
    fplstVendor.Clear
    txtTotPOAmt = 0
    fpcboDepartment.ListIndex = -1
    vaSpread1.ClearRange 1, 1, 7, 36, True
    txtTotDistAmt = 0
    ClearBuds
End Sub
Private Sub LoadControl()
'loads info from control file
  Dim POFile As Integer, POFileLen As Integer, NumRecs As Integer
  ReDim POCont(1) As POControlRecType
  OpenPOFile POFile, NumRecs
  If LOF(POFile) > 0 Then
    Get POFile, 1, POCont(1)
    txtPONumber = "N/A"
    txtShipTo1 = QPTrim(POCont(1).Shipto1)
    txtShipTo2 = QPTrim(POCont(1).Shipto2)
    txtShipTo3 = QPTrim(POCont(1).Shipto3)
    txtShipTo4 = QPTrim(POCont(1).Shipto4)
    txtShipTo5 = QPTrim(POCont(1).Shipto5)
    txtFOB = QPTrim(POCont(1).FOB)
    txtShipVia = QPTrim(POCont(1).Shipvia)
    txtTerms = QPTrim(POCont(1).Terms)
    txtAddinst1 = QPTrim(POCont(1).Addinst1)
    txtAddinst2 = QPTrim(POCont(1).Addinst2)
    txtAddinst3 = QPTrim(POCont(1).Addinst3)
  End If
  Close POFile
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ' Me.Visible = False
    DoEvents
    Temp_Class.ResizeControls Me
    'DoEvents
    Me.Visible = True
  '  Me.SetFocus
    DoEvents
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
      cmdExit_Click
      KeyCode = 0
    Case vbKeyF10:
      cmdSave_Click
      KeyCode = 0
    Case vbKeyF9:
      cmdAddDist_Click
      KeyCode = 0
    Case vbKeyF2:
      cmdNew_Click
      KeyCode = 0
    Case vbKeyF4:
      cmdEdit_Click
      KeyCode = 0
    Case vbKeyF5:
      cmdList_Click
      KeyCode = 0
    Case vbKeyF3:
      cmdDelete_Click
      KeyCode = 0
    Case vbKeyF6:
      cmdDelDist_Click
      KeyCode = 0
    Case vbKeyPageDown:
      Call cmdDist_Click
      KeyCode = 0
    Case vbKeyPageUp:
      Call cmdPage1_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub
'This is to fix spreadsheet for various resolutions
Public Function Fixspread()
'    Select Case screenW
'      Case 1280
'      If Screen.TwipsPerPixelX <> 12 Then
'        coladj = 12.5
'        vaSpread1.RowHeight(-1) = 22
'        vaSpread1.RowHeight(0) = 22
'      Else
'        coladj = 8.3
'        vaSpread1.RowHeight(-1) = 19
'        vaSpread1.RowHeight(0) = 19
'      End If
'      Case 1152
'      If Screen.TwipsPerPixelX <> 12 Then
'        coladj = 10.5
'        vaSpread1.RowHeight(0) = 18
'        vaSpread1.RowHeight(-1) = 18
'      Else
'        coladj = 6.6
'        vaSpread1.RowHeight(0) = 15.5
'        vaSpread1.RowHeight(-1) = 15.5
'      End If
'      Case 1024
'      If Screen.TwipsPerPixelX <> 12 Then
'        coladj = 8.2
'        vaSpread1.RowHeight(0) = 16
'        vaSpread1.RowHeight(-1) = 16
'      Else
'        coladj = 4.75
'      End If
'      Case 800
'        coladj = 4.55
'        vaSpread1.Font.Size = 10
'        vaSpread1.RowHeight(-1) = 12
'      Case Else
'        'don't worry be happpy
'    End Select
    'vaSpread1.ColWidth(-1) = vaSpread1.ColWidth(-1) + coladj
    vaSpread1.Font.Size = 8
End Function

Private Sub cmdExit_Click()
  If Changed = False Then
    Undolok RecNum
    frmPOProcessMenu.Show
    Unload frmPOEnterEdit
  Else
    If MsgBox("Changes Have Been Made to the Current Record." & Chr(13) & Chr(13) & "                          Select OK to Abandon," & Chr(13) & Chr(13) & "       or Cancel to Remain on Entry/Edit Screen.", vbOKCancel, "Abandon Changes?") = vbOK Then
      Undolok RecNum
      frmPOProcessMenu.Show
      Unload frmPOEnterEdit
    End If
'*****      'figure out how to get focus to proper place.
     ' fpcbo.SetFocus
    
  End If
End Sub
Private Sub EdorNewEntry()
  Dim POEditFile As Integer, NumEdTrans As Integer
  Dim Rec As Integer, FileName As String, EdLen As Integer
  ReDim POCont(1) As POControlRecType
  '--get first active record number and set Editing Mode
  OpenPOEditFile POEditFile, NumEdTrans

  If NumEdTrans > 0 Then
    For Rec = 1 To NumEdTrans
      Get POEditFile, Rec, POEdit
      If POEdit.Deleted <> True Then
        If QPTrim(POEdit.PONum) = "N/A" Then
        'RecNum = Rec
        'EMode = True
        RecLok = True
        Exit For
        End If
      End If
    Next
  End If
  
  Close POEditFile
  If RecLok = True Then
    RecNum = NumEdTrans + 1
  Else
    RecNum = 1
  End If
    EMode = False
    SetScreen
    LoadControl
    txtDate.Text = Format(Now, "mm/dd/yyyy")
    fpcboVendCode.ListIndex = -1
    fplstVendor.Clear
    txtTotPOAmt = 0
    fpcboDepartment.ListIndex = -1
'***** spreadsheet do Not have to set blank fields on load ..
    txtTotDistAmt = 0
    ClearBuds

  If QPTrim$(txtPONumber) <> "N/A" Then
    fpcboVendCode.Enabled = False
  End If
End Sub
Private Function Changed()
  Dim POFile As Integer, POFileLen As Integer, NumRecs As Integer
  Dim POEditFile As Integer, NumEdTrans As Integer
  Dim cnt As Integer
  ReDim POCont(1) As POControlRecType
  
  If Val(txtStock) <> 0 Then GoTo DoChange
  If Val(txtDesc) <> 0 Then GoTo DoChange
  If Val(txtQty) <> 0 Then GoTo DoChange
  If txtPrice <> 0 Then GoTo DoChange
  If fpcboAcctNumNa.ListIndex <> -1 Then GoTo DoChange
  If EMode = False Then
    If fpcboDepartment.ListIndex <> -1 Then GoTo DoChange
    If fpcboVendCode.ListIndex <> -1 Then GoTo DoChange
    If txtTotPOAmt <> 0 Then GoTo DoChange
    If Val(txtShipOn) <> 0 Then GoTo DoChange
    OpenPOFile POFile, NumRecs
    If LOF(POFile) > 0 Then
      Get POFile, 1, POCont(1)
      If QPTrim(txtFOB) <> QPTrim(POCont(1).FOB) Then GoTo DoChange
      If QPTrim(txtShipVia) <> QPTrim(POCont(1).Shipvia) Then GoTo DoChange
      If QPTrim(txtTerms) <> QPTrim(POCont(1).Terms) Then GoTo DoChange
      If QPTrim(txtShipTo1) <> QPTrim(POCont(1).Shipto1) Then GoTo DoChange
      If QPTrim(txtShipTo2) <> QPTrim(POCont(1).Shipto2) Then GoTo DoChange
      If QPTrim(txtShipTo3) <> QPTrim(POCont(1).Shipto3) Then GoTo DoChange
      If QPTrim(txtShipTo4) <> QPTrim(POCont(1).Shipto4) Then GoTo DoChange
      If QPTrim(txtShipTo5) <> QPTrim(POCont(1).Shipto5) Then GoTo DoChange
      If QPTrim(txtAddinst1) <> QPTrim(POCont(1).Addinst1) Then GoTo DoChange
      If QPTrim(txtAddinst2) <> QPTrim(POCont(1).Addinst2) Then GoTo DoChange
      If QPTrim(txtAddinst3) <> QPTrim(POCont(1).Addinst3) Then GoTo DoChange
    End If
    Close POFile
    vaSpread1.Row = 1
    vaSpread1.col = 1
    If Val(vaSpread1.Text) <> 0 Then GoTo DoChange
    Changed = False
  Else
    OpenPOEditFile POEditFile, NumEdTrans
    Get POEditFile, RecNum, POEdit
    If txtDate <> Format(DateAdd("d", (POEdit.PODATE), "12-31-1979"), "mm/dd/yyyy") Then GoTo DoChangeClose
    fpcboDepartment.col = 1
    If fpcboDepartment.ColText <> POEdit.REQNUM Then GoTo DoChangeClose
    fpcboVendCode.col = 1
    If fpcboVendCode.ColText <> POEdit.VNDRREC Then GoTo DoChangeClose
    If QPTrim(txtFOB) <> QPTrim(POEdit.FOB) Then GoTo DoChangeClose
    If QPTrim(txtShipVia) <> QPTrim(POEdit.Shipvia) Then GoTo DoChangeClose
    If QPTrim(txtTerms) <> QPTrim(POEdit.Terms) Then GoTo DoChangeClose
    If QPTrim(txtShipOn) <> QPTrim(POEdit.SHIPON) Then GoTo DoChangeClose
    If txtTotPOAmt.DoubleValue <> POEdit.POAmt Then GoTo DoChangeClose
    If QPTrim(txtShipTo1) <> QPTrim(POEdit.SHPLINE1) Then GoTo DoChangeClose
    If QPTrim(txtShipTo2) <> QPTrim(POEdit.SHPLINE2) Then GoTo DoChangeClose
    If QPTrim(txtShipTo3) <> QPTrim(POEdit.SHPLINE3) Then GoTo DoChangeClose
    If QPTrim(txtShipTo4) <> QPTrim(POEdit.SHPLINE4) Then GoTo DoChangeClose
    If QPTrim(txtShipTo5) <> QPTrim(POEdit.SHPLINE5) Then GoTo DoChangeClose
    If QPTrim(txtAddinst1) <> QPTrim(POEdit.Addinst1) Then GoTo DoChangeClose
    If QPTrim(txtAddinst2) <> QPTrim(POEdit.Addinst2) Then GoTo DoChangeClose
    If QPTrim(txtAddinst3) <> QPTrim(POEdit.Addinst3) Then GoTo DoChangeClose

      For cnt = 1 To 36
        vaSpread1.Row = cnt
        vaSpread1.col = 1
        If Val(vaSpread1.Text) <> POEdit.ITEMS(cnt).AcctRec Then GoTo DoChangeClose
        If Val(vaSpread1.Text) = 0 Then
          Changed = False
          Exit For
        End If
        vaSpread1.col = 2
        If vaSpread1.Text <> QPTrim(POEdit.ITEMS(cnt).STKNO) Then GoTo DoChangeClose
        vaSpread1.col = 3
        If vaSpread1.Text <> QPTrim(POEdit.ITEMS(cnt).Desc) Then GoTo DoChangeClose
        vaSpread1.col = 4
        If vaSpread1.Text <> POEdit.ITEMS(cnt).QUAN Then GoTo DoChangeClose
        vaSpread1.col = 5
        If vaSpread1.Text <> POEdit.ITEMS(cnt).PRICE Then GoTo DoChangeClose
        vaSpread1.col = 7
        If vaSpread1.Text <> QPTrim(POEdit.ITEMS(cnt).ACCTNO) Then GoTo DoChangeClose
        Changed = False
      Next
    Close POEditFile
    End If
    Exit Function
DoChange:
   Changed = True
   Exit Function
DoChangeClose:
   Changed = True
   Close POEditFile
   Exit Function

End Function
Private Function Check4Trans()
  Dim POEditFile As Integer, NumEdTrans As Integer
  Dim cnt As Integer, Good As Integer
  Good = 0
  If Exist("APPED.dat") Then
    OpenPOEditFile POEditFile, NumEdTrans
    If NumEdTrans > 0 Then
      For cnt = 1 To NumEdTrans
        Get POEditFile, cnt, POEdit
        If POEdit.Deleted <> True Then
          If QPTrim(POEdit.PONum) = "N/A" Then
            Good = Good + 1
          End If
        End If
      Next
    Else
      Check4Trans = False
    End If
  Else
    Check4Trans = False
  End If
  If Good > 0 Then
    Check4Trans = True
  Else
    Check4Trans = False
  End If
 Close POEditFile
 End Function

Private Sub cmdEdit_Click()
 If Changed = True Then
    If MsgBox("Changes Were Made to the Current Information on Screen and Not Saved." & Chr(13) & "Select OK to View Edit List," & Chr(13) & "or Cancel to Remain on Current Record.", vbOKCancel, "View Edit List?") = vbCancel Then
      vaTabPro1.ActivePage = 0
      txtDate.SetFocus
      Exit Sub
    End If
  End If
  Undolok RecNum
  NextNew
  If Check4Trans = True Then
    frmPOListing.Show 1, frmPOEnterEdit
    SetScreen
    If EMode = True Then
      
      'DisplayTotals
      txtDate.SetFocus
      
    End If
  Else
    MsgBox "No Entries To Display.", vbOKOnly, "No Entries"
    txtDate.SetFocus
  End If
End Sub
Private Sub Undolok(OldRec)
  Dim POEditFile As Integer, NumEdTrans As Integer
  Dim POBusy As Boolean
  If Exist("APPEd.DAT") Then POBusy = GetAttr("APPEd.DAT") And vbReadOnly
  If Not POBusy Then
    OpenPOEditFile POEditFile, NumEdTrans
      If OldRec <= NumEdTrans Then
        Get POEditFile, RecNum, POEdit
        POEdit.LOCKED = False
        Put POEditFile, RecNum, POEdit
      End If
      Close POEditFile
  Else
    MsgBox "Posting In Progress, Editing May Not Continue At This Time.", vbOKOnly, "Canceled"
    frmPOProcessMenu.Show
    Unload frmPOEnterEdit
  End If
End Sub

Private Sub cmdList_Click()
  If Changed = True Then
    If MsgBox("Changes Were Made to the Current Information on Screen and Not Saved." & Chr(13) & "Select OK to View List," & Chr(13) & "or Cancel to Remain on Current Record.", vbOKCancel, "View List?") = vbCancel Then
      vaTabPro1.ActivePage = 0
      txtDate.SetFocus
      Exit Sub
    End If
  End If
  Undolok RecNum
  NextNew
  If Check4Trans = True Then
    frmPOListing.Show 1, frmPOEnterEdit
    SetScreen
    If EMode = True Then
      
      vaTabPro1.ActivePage = 0
      txtDate.SetFocus
    End If
  Else
    MsgBox "No Entries To Display.", vbOKOnly, "No Entries"
    vaTabPro1.ActivePage = 0
    cmdDelDist.Enabled = False
    txtDate.SetFocus
  End If
End Sub

Private Sub cmdNew_Click()
  Dim POBusy As Boolean
  POBusy = False
  If Exist("APPED.DAT") Then POBusy = GetAttr("APPED.DAT") And vbReadOnly
  If Not POBusy Then
    If Changed = True Then
      If MsgBox("Changes Have Been Made to the Current Record." & Chr(13) & "Select OK to Abandon," & Chr(13) & "or Cancel to Remain on Current Record.", vbOKCancel, "Abandon Changes?") = vbCancel Then
        txtDate.SetFocus
        Exit Sub
      End If
    End If
    Undolok RecNum
    NextNew
    vaTabPro1.ActivePage = 0
    cmdDelDist.Enabled = False
    txtDate.SetFocus
  Else
    MsgBox "Posting Is In Progress, Editing Not Allowed At This Time.", vbOKOnly, "Canceled"
    frmPOProcessMenu.Show
    Unload frmPOEnterEdit
  End If
End Sub

Private Sub cmdAddDist_Click()
  Dim TempAmt As Double
  If VerifyEntered = False Then
    MsgBox "The Information In The Top Section Must Be Completed Before Adding Distributions.", vbOKOnly, "Cash Disbursement"
  Else
    If fpcboAcctNumNa.Text <> "" And txtPrice.DoubleValue > 0 And txtQty > 0 Then
      If vaSpread1.DataRowCnt < 36 Then
      vaSpread1.Row = vaSpread1.DataRowCnt + 1
      vaSpread1.col = 1
      fpcboAcctNumNa.col = 0
      vaSpread1.Text = fpcboAcctNumNa.ColText
      vaSpread1.col = 2
      fpcboAcctNumNa.col = 1
      vaSpread1.Text = txtStock
      vaSpread1.col = 3
      vaSpread1.Text = txtDesc
      vaSpread1.col = 4
      vaSpread1.Text = txtQty
      vaSpread1.col = 5
      vaSpread1.Text = txtPrice
      vaSpread1.col = 6
      TempAmt = Round(Val(txtQty * txtPrice.DoubleValue))
      vaSpread1.Text = TempAmt
      vaSpread1.col = 7
      fpcboAcctNumNa.col = 1
      vaSpread1.Text = fpcboAcctNumNa.ColText
      txtTotDistAmt = Round(TempAmt + txtTotDistAmt.DoubleValue)
      fpcboAcctNumNa.ListIndex = -1
      txtPrice = 0
      txtStock = ""
      txtDesc = ""
      txtQty = 0
      ClearBuds
      txtStock.SetFocus
    Else
      MsgBox "Only 36 Distributions Allowed Per PO.", vbOKOnly, "Limit Reached."
    End If
    Else
      MsgBox "The Quantity, Price and Account Must Be Entered Before Adding To The Distribution List.", vbOKOnly, "Add Distribution Denied"
    End If
  End If
  
End Sub
Private Function VerifyEntered()
  If txtQty > 0 Then
    If txtPrice > 0 Then
      If fpcboAcctNumNa.ListIndex <> -1 Then
        VerifyEntered = True
      Else
        VerifyEntered = False
        fpcboAcctNumNa.SetFocus
        Exit Function
      End If
    Else
      VerifyEntered = False
      txtPrice.SetFocus
      Exit Function
    End If
  Else
    VerifyEntered = False
    txtQty.SetFocus
    Exit Function
  End If
End Function

 
Private Function Ready2Save()
  Dim TempDate As Integer, cnt As Integer
  Dim TempDist As Double
  TempDist = 0
  'Take care of Invalid Data and Messages in this Section
  'CheckValDate is in main module to verify date entered w/correct format
  If CheckValDate(txtDate) = True Then
    TempDate = DateDiff("d", "12/31/1979", txtDate)
  'Also compare date with Hi/Lo range
    If (TempDate < LPDate) Or (TempDate > HPDate) Then
      MsgBox "This Date Is Not Within Allowable Posting Range. Please Correct or Change Setup.", vbOKOnly, "Invalid Date"
      Ready2Save = False
      Exit Function
    Else
      Ready2Save = True
    End If
  Else
    MsgBox "This Date Is Not Valid. Please Correct.", vbOKOnly, "Invalid Date"
    Ready2Save = False
    Exit Function
  End If
  'Not allow Zero Total or Unequal Distritbutions
  If txtTotPOAmt <> 0 Then
    If txtTotDistAmt = 0 Or txtTotPOAmt <> txtTotDistAmt Then
      MsgBox "The Total Purchase Order Does Not Equal The Amount of The Distributions." & Chr$(13) & "Please Correct Before Saving.", vbOKOnly, "PO Entry"
      Ready2Save = False
      Exit Function
    Else
      
      For cnt = 1 To 36
        vaSpread1.col = 6
        vaSpread1.Row = cnt
        If vaSpread1.Text <> "" Then
          TempDist = Round(vaSpread1.Text + TempDist)
        Else
          Exit For
        End If
      Next
      If TempDist <> txtTotDistAmt Or TempDist <> txtTotPOAmt Then
        MsgBox "Totals Are Not In Balance. Please Correct.", vbOKOnly, "PO Entry"
        Ready2Save = False
        Exit Function
      Else
        Ready2Save = True
      End If
     End If
  Else
    MsgBox "You May Not Save A Purchase Order With A $0.00 Total.", vbOKOnly, "PO Entry"
    Ready2Save = False
  End If
End Function
Private Sub fpcboAcctNumNa_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboAcctNumNa.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    ClearBuds
    fpcboAcctNumNa.ListIndex = -1
    fpcboAcctNumNa.Action = ActionClearSearchBuffer
  End If
  If fpcboAcctNumNa.ListDown <> True Then
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

Private Sub fpcboDepartment_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboDepartment.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboDepartment.ListIndex = -1
    fpcboDepartment.Action = ActionClearSearchBuffer
  End If
  If fpcboDepartment.ListDown <> True Then
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
Private Sub fpcboAcctNumNa_LostFocus()
  fpcboAcctNumNa.Action = ActionClearSearchBuffer
End Sub
Private Sub fpcboDepartment_LostFocus()
  fpcboDepartment.Action = ActionClearSearchBuffer
End Sub

Private Sub cmdSave_Click()
  If Ready2Save = True Then
    SavePO
    Call NextNew
  Else
    MsgBox "             Save Canceled.", vbOKOnly, "PO Entry"
  End If
End Sub

Private Sub SavePO()
  Dim POEditFile As Integer, NumEdTrans As Integer, cnt As Integer
  Dim POBusy As Boolean
  POBusy = False
  If Exist("APPED.DAT") Then POBusy = GetAttr("APPED.DAT") And vbReadOnly
  If Not POBusy Then
    OpenPOEditFile POEditFile, NumEdTrans
    POEdit.Deleted = 0
    POEdit.LOCKED = False
    POEdit.PONum = Trim(txtPONumber)
    fpcboDepartment.col = 1
    POEdit.REQNUM = QPTrim(fpcboDepartment.ColText)
    'POEdit.REQNUM = QPTrim(fpcboDepartment.Text)
    POEdit.PODATE = DateDiff("d", "12/31/1979", txtDate)
    fpcboVendCode.col = 0
    POEdit.VNDRCODE = fpcboVendCode.ColText
    fpcboVendCode.col = 1
    POEdit.VNDRREC = fpcboVendCode.ColText
    fplstVendor.col = -1
    fplstVendor.Selected(0) = True
    POEdit.VNDRINF1 = QPTrim(fplstVendor.Text)
    fplstVendor.Selected(1) = True
    POEdit.VNDRINF2 = QPTrim(fplstVendor.Text)
    fplstVendor.Selected(2) = True
    POEdit.VNDRINF3 = QPTrim(fplstVendor.Text)
    fplstVendor.Selected(3) = True
    POEdit.VNDRINF4 = QPTrim(fplstVendor.Text)
    fplstVendor.Selected(4) = True
    POEdit.VNDRINF5 = QPTrim(fplstVendor.Text)
    POEdit.FOB = QPTrim(txtFOB)
    POEdit.Shipvia = QPTrim(txtShipVia)
    POEdit.Terms = QPTrim(txtTerms)
    POEdit.SHIPON = QPTrim(txtShipOn)
    POEdit.POAmt = txtTotPOAmt
    POEdit.SHPLINE1 = QPTrim(txtShipTo1)
    POEdit.SHPLINE2 = QPTrim(txtShipTo2)
    POEdit.SHPLINE3 = QPTrim(txtShipTo3)
    POEdit.SHPLINE4 = QPTrim(txtShipTo4)
    POEdit.SHPLINE5 = QPTrim(txtShipTo5)
    POEdit.Addinst1 = QPTrim(txtAddinst1)
    POEdit.Addinst2 = QPTrim(txtAddinst2)
    POEdit.Addinst3 = QPTrim(txtAddinst3)
    
    For cnt = 1 To 36
      vaSpread1.Row = cnt
      vaSpread1.col = 1
      If vaSpread1.Text = "" Then
        POEdit.ITEMS(cnt).AcctRec = 0
        POEdit.ITEMS(cnt).STKNO = ""
        POEdit.ITEMS(cnt).ACCTNO = ""
        POEdit.ITEMS(cnt).Desc = ""
        POEdit.ITEMS(cnt).EXT = 0
        POEdit.ITEMS(cnt).PRICE = 0
        POEdit.ITEMS(cnt).QUAN = 0
      Else
        POEdit.ITEMS(cnt).AcctRec = vaSpread1.Text
        vaSpread1.col = 2
        POEdit.ITEMS(cnt).STKNO = QPTrim(vaSpread1.Text)
        vaSpread1.col = 3
        POEdit.ITEMS(cnt).Desc = QPTrim(vaSpread1.Text)
        vaSpread1.col = 4
        POEdit.ITEMS(cnt).QUAN = vaSpread1.Text
        vaSpread1.col = 5
        POEdit.ITEMS(cnt).PRICE = vaSpread1.Text
        vaSpread1.col = 6
        POEdit.ITEMS(cnt).EXT = vaSpread1.Text
        vaSpread1.col = 7
        POEdit.ITEMS(cnt).ACCTNO = QPTrim(vaSpread1.Text)
      End If
    Next
    If EMode = False Then
      If NumEdTrans > 0 Then
        RecNum = NumEdTrans + 1
      Else
        RecNum = 1
      End If
    End If
    Put POEditFile, RecNum, POEdit
    Close POEditFile
  Else
    MsgBox "Posting Is In Progress, Editing Not Allowed At This Time.", vbOKOnly, "Canceled"
    frmPOProcessMenu.Show
    Unload frmPOEnterEdit
  End If
End Sub


Private Sub mnuExit_Click()
  Call cmdExit_Click
End Sub

Private Sub NextNew()
  Dim POEditFile As Integer, NumEdTrans As Integer
  OpenPOEditFile POEditFile, NumEdTrans
  Close POEditFile
   If NumEdTrans > 0 Then
     RecNum = NumEdTrans + 1
   Else
     RecNum = 1
   End If

   EMode = False
   ClearFields
   SetScreen
   LoadControl
   vaTabPro1.ActivePage = 0
   cmdDelDist.Enabled = False
   fpcboDepartment.SetFocus
End Sub

Private Sub fpcboVendCode_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboVendCode.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboVendCode.ListIndex = -1
    fpcboVendCode.Action = ActionClearSearchBuffer
    fplstVendor.Clear
  End If
  If fpcboVendCode.ListDown <> True Then
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

Private Sub fpcboVendCode_LostFocus()
  fpcboVendCode.Action = ActionClearSearchBuffer
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub


Private Sub txtDate_LostFocus()
  If CheckValDate(txtDate) = False Then
    MsgBox "Invalid Date, Please Correct.", vbOKOnly, "Invalid Date"
    txtDate.SetFocus
  End If
End Sub

Private Sub txtFOB_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub


Private Sub txtShipVia_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub txtTerms_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub txtStock_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub txtDesc_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub txtAddinst1_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub txtAddinst2_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub txtAddinst3_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub txtShipOn_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub txtShipTo1_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub txtShipTo2_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub txtShipTo3_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub txtShipTo4_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub
Private Sub txtShipTo5_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub txtTotPOAmt_Change()
  txtTot2 = txtTotPOAmt
End Sub

Private Sub vaSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF11 Then
    Call vaSpread1_DblClick(vaSpread1.ActiveCol, vaSpread1.ActiveRow)
  End If
End Sub

Private Sub vaSpread1_DblClick(ByVal col As Long, ByVal Row As Long)
  Dim TempAcct As String, TempAmt As Double
  Dim TempCol As Long, TempRow As Long
  skip = True
  TempRow = Row
  TempCol = col
  If TempRow > 0 Then
    vaSpread1.Row = TempRow
    vaSpread1.col = 7
    TempAcct = QPTrim(vaSpread1.Text)
    If vaSpread1.Text <> "" Then
      If fpcboAcctNumNa.ListIndex <> -1 Or txtPrice <> 0 Then
        If MsgBox("Do You Wish To Abandon Current Distribution?, 'Yes' or 'No' Complete PO Entry.", vbYesNo, "Clear??") = vbNo Then
          cmdAddDist.SetFocus
          Exit Sub
        Else
          fpcboAcctNumNa.ListIndex = -1
          txtPrice = 0
        End If
      End If
        vaSpread1.Row = TempRow
        fpcboAcctNumNa.SearchText = QPStrip(TempAcct)
        fpcboAcctNumNa.Action = 0
        If fpcboAcctNumNa.SearchIndex <> -1 Then
          fpcboAcctNumNa.ListIndex = fpcboAcctNumNa.SearchIndex
        End If
          vaSpread1.Row = TempRow
          vaSpread1.col = 1
          fpcboAcctNumNa.col = 0
          fpcboAcctNumNa.ColText = vaSpread1.Text
          vaSpread1.Row = TempRow
          vaSpread1.col = 2
          txtStock = vaSpread1.Text
          vaSpread1.Row = TempRow
          vaSpread1.col = 3
          txtDesc = vaSpread1.Text
          vaSpread1.Row = TempRow
          vaSpread1.col = 4
          txtQty = vaSpread1.Text
          'vaSpread1.col = 5
          'txtPrice = vaSpread1.Text
          vaSpread1.Row = TempRow
          vaSpread1.col = 6
          TempAmt = vaSpread1.Text
          'vaSpread1.col = 5
          txtPrice = TempAmt / txtQty
          'vaSpread1.col = 7
          'fpcboAcctNumNa.col = 1
          'fpcboAcctNumNa.ColText = vaSpread1.Text
          txtTotDistAmt = Round(txtTotDistAmt.DoubleValue - TempAmt)
          vaSpread1.DeleteRows TempRow, 1
          txtStock.SetFocus
    End If
  End If
  skip = False
End Sub


