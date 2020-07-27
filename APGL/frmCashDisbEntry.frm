VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Begin VB.Form frmCashDisbEntry 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash Disbursement Entry"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11655
   ForeColor       =   &H80000007&
   Icon            =   "frmCashDisbEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo txtBanks 
      Height          =   375
      Left            =   9075
      TabIndex        =   6
      Top             =   1950
      Width           =   1095
      _Version        =   196608
      _ExtentX        =   1931
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
      Object.TabStop         =   0   'False
      BackColor       =   16777215
      ForeColor       =   0
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
      AutoSearch      =   1
      SearchMethod    =   0
      VirtualMode     =   0   'False
      VRowCount       =   0
      DataSync        =   3
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   14737632
      ThreeDInsideShadowColor=   0
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   14737632
      ThreeDOutsideShadowColor=   8421504
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   12632256
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   4210752
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
      ColumnEdit      =   0
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   3504
      EditHeight      =   -1
      GrayAreaColor   =   12632256
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   150
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   12632256
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   0   'False
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   12632256
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
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmCashDisbEntry.frx":08CA
   End
   Begin LpLib.fpCombo fpcboAcctNumNa 
      Height          =   405
      Left            =   1680
      TabIndex        =   7
      Top             =   2835
      Width           =   4935
      _Version        =   196608
      _ExtentX        =   8705
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
      BackColor       =   16777215
      ForeColor       =   0
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
      ThreeDInsideHighlightColor=   14737632
      ThreeDInsideShadowColor=   0
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   14737632
      ThreeDOutsideShadowColor=   8421504
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
      BorderWidth     =   1
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   12632256
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   4210752
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
      GrayAreaColor   =   12632256
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   150
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   12632256
      GroupHeaderHeight=   -1
      GroupHeaderShow =   0   'False
      AllowGrpResize  =   0
      AllowGrpDragDrop=   0
      MergeAdjustView =   0   'False
      ColumnHeaderShow=   0   'False
      ColumnHeaderHeight=   -1
      GrpsFrozen      =   0
      BorderGrayAreaColor=   12632256
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
      ColDesigner     =   "frmCashDisbEntry.frx":0CC4
   End
   Begin VB.CommandButton cmdDelDist 
      Appearance      =   0  'Flat
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
      Left            =   5160
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7728
      Width           =   1332
   End
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
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
      Left            =   312
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7728
      Width           =   1332
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   3140
      Left            =   1740
      TabIndex        =   10
      Top             =   3690
      Width           =   8265
      _Version        =   196613
      _ExtentX        =   14579
      _ExtentY        =   5539
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
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
      GrayAreaBackColor=   13684944
      MaxCols         =   4
      MaxRows         =   36
      Protect         =   0   'False
      ShadowColor     =   13684944
      ShadowDark      =   8421504
      ShadowText      =   0
      SpreadDesigner  =   "frmCashDisbEntry.frx":1126
      VisibleCols     =   3
      VisibleRows     =   10
      ScrollBarTrack  =   1
   End
   Begin EditLib.fpCurrency txtDebAmt 
      Height          =   372
      Left            =   6960
      TabIndex        =   8
      Top             =   2832
      Width           =   1452
      _Version        =   196608
      _ExtentX        =   2561
      _ExtentY        =   656
      Enabled         =   -1  'True
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
      BackColor       =   16777215
      ForeColor       =   0
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   14737632
      ThreeDInsideShadowColor=   0
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   14737632
      ThreeDOutsideShadowColor=   8421504
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
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
      ThreeDTextHighlightColor=   14737632
      ThreeDTextShadowColor=   8421504
      ThreeDTextOffset=   1
      AlignTextH      =   2
      AlignTextV      =   1
      AllowNull       =   0   'False
      NoSpecialKeys   =   0
      AutoAdvance     =   -1  'True
      AutoBeep        =   0   'False
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   16777215
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   16777215
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   0
      ControlType     =   0
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   -1
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   "$"
      DecimalPoint    =   ""
      FixedPoint      =   -1  'True
      LeadZero        =   0
      MaxValue        =   "999999999"
      MinValue        =   "-999999999"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   12632256
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   12632256
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   4210752
      BorderDropShadowWidth=   3
      ButtonColor     =   12632256
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.CommandButton cmdEdit 
      Appearance      =   0  'Flat
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
      Left            =   6776
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7728
      Width           =   1332
   End
   Begin VB.CommandButton cmdNew 
      Appearance      =   0  'Flat
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
      Left            =   1928
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7728
      Width           =   1332
   End
   Begin VB.CommandButton cmdList 
      Appearance      =   0  'Flat
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
      Left            =   8392
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7728
      Width           =   1332
   End
   Begin EditLib.fpText txtCheckNum 
      Height          =   324
      Left            =   9072
      TabIndex        =   2
      Top             =   1296
      Width           =   1092
      _Version        =   196608
      _ExtentX        =   1926
      _ExtentY        =   572
      Enabled         =   -1  'True
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
      BackColor       =   16777215
      ForeColor       =   0
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   14737632
      ThreeDInsideShadowColor=   0
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   14737632
      ThreeDOutsideShadowColor=   4210752
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
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
      ThreeDTextHighlightColor=   14737632
      ThreeDTextShadowColor=   8421504
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
      UserEntry       =   0
      HideSelection   =   0   'False
      InvalidColor    =   16777215
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   16777215
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   8
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   12632256
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   12632256
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   4210752
      BorderDropShadowWidth=   3
      ButtonColor     =   12632256
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpCurrency txtDist 
      Height          =   372
      Left            =   7680
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   7104
      Width           =   1812
      _Version        =   196608
      _ExtentX        =   3196
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
      BackColor       =   16777215
      ForeColor       =   0
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   16777215
      ThreeDInsideShadowColor=   0
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   14737632
      ThreeDOutsideShadowColor=   8421504
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
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
      ThreeDTextHighlightColor=   14737632
      ThreeDTextShadowColor=   8421504
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
      InvalidColor    =   16777215
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   16777215
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
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   12632256
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   0
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   8421504
      BorderDropShadowWidth=   3
      ButtonColor     =   12632256
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpCurrency txtUndist 
      Height          =   372
      Left            =   3816
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   7104
      Width           =   1812
      _Version        =   196608
      _ExtentX        =   3196
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
      BackColor       =   16777215
      ForeColor       =   0
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   14737632
      ThreeDInsideShadowColor=   0
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   14737632
      ThreeDOutsideShadowColor=   8421504
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
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
      ThreeDTextHighlightColor=   14737632
      ThreeDTextShadowColor=   4210752
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
      InvalidColor    =   16777215
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   16777215
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
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
      MaxValue        =   "9000000000"
      MinValue        =   "-9000000000"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   12632256
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   12632256
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   8421504
      BorderDropShadowWidth=   3
      ButtonColor     =   12632256
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText txtDesc 
      Height          =   324
      Left            =   4944
      TabIndex        =   1
      Top             =   1320
      Width           =   2700
      _Version        =   196608
      _ExtentX        =   4762
      _ExtentY        =   572
      Enabled         =   -1  'True
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
      BackColor       =   16777215
      ForeColor       =   0
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   14737632
      ThreeDInsideShadowColor=   0
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   14737632
      ThreeDOutsideShadowColor=   4210752
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
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
      ThreeDTextHighlightColor=   14737632
      ThreeDTextShadowColor=   4210752
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
      UserEntry       =   0
      HideSelection   =   0   'False
      InvalidColor    =   16777215
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   16777215
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   20
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   12632256
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   12632256
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   8421504
      BorderDropShadowWidth=   3
      ButtonColor     =   12632256
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpCurrency txtAmount 
      Height          =   372
      Left            =   2208
      TabIndex        =   4
      Top             =   1896
      Width           =   1812
      _Version        =   196608
      _ExtentX        =   3196
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
      BackColor       =   16777215
      ForeColor       =   0
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   12632256
      ThreeDInsideShadowColor=   0
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   14737632
      ThreeDOutsideShadowColor=   8421504
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
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
      ThreeDTextHighlightColor=   14737632
      ThreeDTextShadowColor=   8421504
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
      InvalidColor    =   16777215
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   16777215
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
      MaxValue        =   "999999999"
      MinValue        =   "-999999999"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   12632256
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   12632256
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   8421504
      BorderDropShadowWidth=   3
      ButtonColor     =   12632256
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpDateTime txtDate 
      Height          =   372
      Left            =   2232
      TabIndex        =   0
      Top             =   1344
      Width           =   1692
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
      BackColor       =   16777215
      ForeColor       =   0
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   14737632
      ThreeDInsideShadowColor=   0
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   14737632
      ThreeDOutsideShadowColor=   4210752
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
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
      ThreeDTextHighlightColor=   14737632
      ThreeDTextShadowColor=   4210752
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
      InvalidColor    =   16777215
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   16777215
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   "10/01/2001"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
      DateMax         =   "20350101"
      DateMin         =   "19800101"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "19800101"
      TimeDefault     =   "000000"
      TimeStyle       =   0
      BorderGrayAreaColor=   12632256
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   12632256
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   4210752
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
      ButtonColor     =   12632256
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.CommandButton cmdDelete 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "F3 &Delete"
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
      Left            =   3544
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7728
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   264
      Left            =   0
      TabIndex        =   24
      Top             =   8376
      Width           =   11652
      _ExtentX        =   20558
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6800
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   6800
            TextSave        =   "8:14 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   6800
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
   Begin VB.CommandButton cmdAddDist 
      Appearance      =   0  'Flat
      BackColor       =   &H00D0D0D0&
      Caption         =   "F9 &Add Distribution"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   8568
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2592
      Width           =   1644
   End
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
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
      Left            =   10008
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7728
      Width           =   1332
   End
   Begin EditLib.fpText txtBatch 
      Height          =   324
      Left            =   9072
      TabIndex        =   3
      Top             =   1608
      Width           =   1092
      _Version        =   196608
      _ExtentX        =   1926
      _ExtentY        =   572
      Enabled         =   -1  'True
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
      BackColor       =   16777215
      ForeColor       =   0
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   14737632
      ThreeDInsideShadowColor=   0
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   14737632
      ThreeDOutsideShadowColor=   4210752
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
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
      ThreeDTextHighlightColor=   14737632
      ThreeDTextShadowColor=   4210752
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
      UserEntry       =   0
      HideSelection   =   0   'False
      InvalidColor    =   16777215
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
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   8
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   12632256
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   12632256
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   4210752
      BorderDropShadowWidth=   3
      ButtonColor     =   12632256
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText txtAdditDesc 
      Height          =   324
      Left            =   4920
      TabIndex        =   5
      Top             =   1944
      Width           =   3348
      _Version        =   196608
      _ExtentX        =   5905
      _ExtentY        =   572
      Enabled         =   -1  'True
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
      BackColor       =   16777215
      ForeColor       =   0
      ThreeDInsideStyle=   1
      ThreeDInsideHighlightColor=   14737632
      ThreeDInsideShadowColor=   0
      ThreeDInsideWidth=   1
      ThreeDOutsideStyle=   1
      ThreeDOutsideHighlightColor=   14737632
      ThreeDOutsideShadowColor=   8421504
      ThreeDOutsideWidth=   1
      ThreeDFrameWidth=   0
      BorderStyle     =   0
      BorderColor     =   0
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
      ThreeDTextHighlightColor=   14737632
      ThreeDTextShadowColor=   8421504
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
      UserEntry       =   0
      HideSelection   =   0   'False
      InvalidColor    =   16777215
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   16777215
      OnFocusAlignH   =   0
      OnFocusAlignV   =   0
      OnFocusNoSelect =   0   'False
      OnFocusPosition =   1
      ControlType     =   0
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   32
      MultiLine       =   0   'False
      PasswordChar    =   ""
      IncHoriz        =   0.25
      BorderGrayAreaColor=   12632256
      NoPrefix        =   0   'False
      ScrollV         =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   12632256
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   8421504
      BorderDropShadowWidth=   3
      ButtonColor     =   12632256
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Label Label3b 
      BackStyle       =   0  'Transparent
      Caption         =   "Additional Desc"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   3
      Left            =   4344
      TabIndex        =   35
      Top             =   1680
      Width           =   2052
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Distributions :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1584
      TabIndex        =   34
      Top             =   3312
      Width           =   1500
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   3255
      Left            =   1680
      Top             =   3600
      Width           =   8415
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   4548
      Left            =   1440
      Top             =   2448
      Width           =   8892
   End
   Begin VB.Label Label3b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Debit Amount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   5
      Left            =   6720
      TabIndex        =   33
      Top             =   2520
      Width           =   1692
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "GL Account Number/Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   0
      Left            =   2400
      TabIndex        =   32
      Top             =   2520
      Width           =   3132
   End
   Begin VB.Label Label3b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Batch"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   2
      Left            =   8232
      TabIndex        =   31
      Top             =   1656
      Width           =   732
   End
   Begin VB.Label lblEdit 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "Edit Cash Disbursement"
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
      Left            =   6960
      TabIndex        =   30
      Top             =   864
      Visible         =   0   'False
      Width           =   3132
   End
   Begin VB.Label lblNew 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "New Cash Disbursement"
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
      Left            =   1824
      TabIndex        =   29
      Top             =   864
      Visible         =   0   'False
      Width           =   2964
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      Height          =   492
      Left            =   1440
      Top             =   6984
      Width           =   8892
   End
   Begin VB.Label lblCredits 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Distributed"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   3
      Left            =   6264
      TabIndex        =   28
      Top             =   7104
      Width           =   1332
   End
   Begin VB.Label lblDebits 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Undistributed"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   2
      Left            =   2256
      TabIndex        =   27
      Top             =   7104
      Width           =   1452
   End
   Begin VB.Label Label3b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Check No"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   0
      Left            =   7848
      TabIndex        =   26
      Top             =   1320
      Width           =   1116
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
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   1
      Left            =   1536
      TabIndex        =   25
      Top             =   1344
      Width           =   612
   End
   Begin VB.Label Label3b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Desc"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   1
      Left            =   4080
      TabIndex        =   23
      Top             =   1344
      Width           =   732
   End
   Begin VB.Label Label4b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bank"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   1
      Left            =   8208
      TabIndex        =   22
      Top             =   1992
      Width           =   732
   End
   Begin VB.Label Label2b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Amt"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   1
      Left            =   1464
      TabIndex        =   21
      Top             =   1896
      Width           =   636
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000016&
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   1236
      Index           =   0
      Left            =   1440
      Top             =   1224
      Width           =   8892
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Disbursement Entry/Edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3600
      TabIndex        =   20
      Top             =   312
      Width           =   4452
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   612
      Left            =   2760
      Top             =   192
      Width           =   6132
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      Height          =   732
      Left            =   2760
      Top             =   72
      Width           =   6132
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
Attribute VB_Name = "frmCashDisbEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GLSetup As GLSetupRecType
Dim GLAcct As GLAcctRecType
Dim GJEdit As TrEditRecType
Dim GLTrans As GLTransRecType
Dim GLCDEd(1) As CJEditRecType
Dim Over As clsTextBoxOverRider
Dim LPDate As Integer, HPDate As Integer
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim Emode As Boolean, RecNum As Integer, CJType As Integer, RecLok As Boolean
Dim ActLen As Integer, OldRec As Integer, DefBnk As Integer
Dim CDActive As String, CashAcct As String, CDCash As String, CDDue As String
Private Temp_Class As Resize_Class
'This is to fix spreadsheet for various resolutions
Public Function Fixspread()
  
'    Select Case screenW
'      Case 1280
'      If Screen.TwipsPerPixelX <> 12 Then
'        coladj = 27.3
'        vaSpread1.RowHeight(-1) = 22.5
'        vaSpread1.RowHeight(0) = 22.5
'      Else
'        coladj = 20
'        vaSpread1.RowHeight(-1) = 19
'        vaSpread1.RowHeight(0) = 19
'      End If
'      Case 1152
'      If Screen.TwipsPerPixelX <> 12 Then
'        coladj = 23.5
'        vaSpread1.RowHeight(-1) = 19.5
'        vaSpread1.RowHeight(0) = 19.5
'      Else
'        coladj = 17
'        vaSpread1.RowHeight(-1) = 16
'        vaSpread1.RowHeight(0) = 16
'      End If
'      Case 1024
'      If Screen.TwipsPerPixelX <> 12 Then
'        coladj = 20
'        vaSpread1.RowHeight(-1) = 18
'        vaSpread1.RowHeight(0) = 18
'      Else
'        coladj = 14
'      End If
'      Case 800
'        coladj = 13.5
'        'vaSpread1.Font.Size = 8
'        vaSpread1.RowHeight(-1) = 13
'      Case Else
'        'don't worry be happpy
'    End Select
    vaSpread1.Font.Size = 6
    vaSpread1.RowHeight(0) = 18
    vaSpread1.RowHeight(-1) = 18
'    vaSpread1.ColHidden(0) = True
    'vaSpread1.ColWidth(-1) = vaSpread1.ColWidth(-1) + coladj
End Function

'**************************
'Use Emode to determine if New record or Editing, so true if editing.
'RecNum passed from Listing to load chosen record to form
'******************************
Private Sub cmdDelete_Click()
  Dim CJEditFile As Integer, CJEditFileNum As Integer, NumEdTrans As Integer
  Dim CDBusy As Boolean
  If Exist("GLCDEd.DAT") Then CDBusy = GetAttr("GLCDEd.DAT") And vbReadOnly
  If Not CDBusy Then
    If MsgBox("Are you sure you wish to delete this entry?", vbYesNo, "Delete GJEntry") = vbYes Then
      If Emode = True Then
        OpenCJEditFile CJEditFileNum, NumEdTrans, CJType
        Get CJEditFileNum, RecNum, GLCDEd(1)
        GLCDEd(1).LOCKED = False
        GLCDEd(1).DelFlag = -1
        Put CJEditFileNum, RecNum, GLCDEd(1)
        Close CJEditFileNum
        GLCDEd(1).DelFlag = 0
        Call MainLog("CD Deleted.")
        Call NextNew
      Else
       Call NextNew
      End If
    Else
      txtDate.SetFocus
    End If
  Else
    MsgBox "Posting In Progress, Editing May Not Continue At This Time.", vbOKOnly, "Canceled"
    frmCashDisbMenu.Show
    Unload frmCashDisbEntry
  End If
End Sub
Private Sub cmdDelDist_Click()
  If vaSpread1.ActiveRow > 0 Then
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 4
    If vaSpread1.Text <> "" Then
      If MsgBox("You Wish to Delete this Distribution?", vbYesNo, "Delete Distribution") = vbYes Then
        vaSpread1.Col = 4
        txtDist = Round#(txtDist.DoubleValue - vaSpread1.Text)
        txtUndist = Round#(txtAmount.DoubleValue - txtDist.DoubleValue)
        
        vaSpread1.DeleteRows vaSpread1.Row, 1
        fpcboAcctNumNa.SetFocus
      End If
    End If
  End If

End Sub

Private Sub cmdEdit_Click()
  Dim RetValue As Integer
 
  If Changed = True Then
    If MsgBox("Changes Were Made to the Current Information on Screen and Not Saved." & Chr(13) & "Select OK to View Edit List," & Chr(13) & "or Cancel to Review Current Record.", vbOKCancel, "View Edit List?") = vbCancel Then
      txtDate.SetFocus
      Exit Sub
    End If
  End If
  Undolok RecNum
  NextNew
  If Check4Trans = True Then
    frmCDListing.Show 1, frmCashDisbEntry
    If Emode = True Then
      SetScreen
      'DisplayTotals
      txtDate.SetFocus
      cmdDelete.Enabled = True
    End If
  Else
    MsgBox "No Entries To Display.", vbOKOnly, "No Entries"
    'RetValue = sndPlaySound("noentries.wav", SND_ASYNC Or SND_NODEFAULT)
    txtDate.SetFocus
  End If
End Sub

Private Sub cmdList_Click()
  If Changed = True Then
    If MsgBox("Changes Were Made to the Current Information on Screen and Not Saved." & Chr(13) & "Select OK to View List," & Chr(13) & "or Cancel to Remain on Current Record.", vbOKCancel, "View List?") = vbCancel Then
      txtDate.SetFocus
      Exit Sub
    End If
  End If
  Undolok RecNum
  NextNew
  If Check4Trans = True Then
    frmCDListing.Show 1, frmCashDisbEntry
    If Emode = True Then
      
      SetScreen
      txtDate.SetFocus
    End If
  Else
    MsgBox "No Entries To Display.", vbOKOnly, "No Entries"
    txtDate.SetFocus
  End If
End Sub

Private Sub cmdNew_Click()
  Dim CJEditFile As Integer, CJEditFileNum As Integer, NumEdTrans As Integer
  Dim CDBusy As Boolean
  If Exist("GLCDEd.DAT") Then CDBusy = GetAttr("GLCDEd.DAT") And vbReadOnly
  If Not CDBusy Then
    OpenCJEditFile CJEditFileNum, NumEdTrans, CJType
    If Changed = True Then
      If MsgBox("Changes Have Been Made to the Current Record." & Chr(13) & "Select OK to Abandon," & Chr(13) & "or Cancel to Remain on Current Record.", vbOKCancel, "Abandon Changes?") = vbCancel Then
        txtDate.SetFocus
        Exit Sub
      End If
    End If
    Get CJEditFileNum, RecNum, GLCDEd(1)
    GLCDEd(1).LOCKED = False
    Put CJEditFileNum, RecNum, GLCDEd(1)
    If NumEdTrans > 0 Then
      RecNum = NumEdTrans + 1
    Else
      RecNum = 1
    End If
    Close CJEditFileNum
    Emode = False
    ClearFields
    txtDate.SetFocus
    SetScreen
    txtDate.SetFocus
  Else
    MsgBox "Posting In Progress, Editing May Not Continue At This Time.", vbOKOnly, "Canceled"
    frmCashDisbMenu.Show
    Unload frmCashDisbEntry
  End If
End Sub

Private Sub cmdAddDist_Click()
  If VerifyEntered = False Then
    MsgBox "The Information In The Top Section Must Be Completed Before Adding Distributions.", vbOKOnly, "Cash Disbursement"
  Else
    If fpcboAcctNumNa.Text <> "" And txtDebAmt.DoubleValue <> 0 Then
    If vaSpread1.DataRowCnt < 36 Then
      vaSpread1.Row = vaSpread1.DataRowCnt + 1
      vaSpread1.Col = 1
      fpcboAcctNumNa.Col = 0
      vaSpread1.Text = fpcboAcctNumNa.ColText
      vaSpread1.Col = 2
      fpcboAcctNumNa.Col = 1
      vaSpread1.Text = fpcboAcctNumNa.ColText
      vaSpread1.Col = 3
      fpcboAcctNumNa.Col = 2
      vaSpread1.Text = fpcboAcctNumNa.ColText
      vaSpread1.Col = 4
      txtDist = Round#(txtDebAmt.DoubleValue + txtDist.DoubleValue)
      vaSpread1.Text = txtDebAmt
      fpcboAcctNumNa.ListIndex = -1
      txtDebAmt = 0
      fpcboAcctNumNa.SetFocus
      txtUndist = Round#(txtAmount.DoubleValue - txtDist.DoubleValue)
    Else
      MsgBox "Only 36 Distributions Allowed.", vbOKOnly, "Limit Reached."
    End If
    Else
      MsgBox "The Account and Amount Must Be Entered Before Adding To The Distribution List.", vbOKOnly, "Add Distribution Denied"
    End If
  End If
End Sub

  
'  If Emode = True Then
'    If Changed = False Then
'      If MsgBox("This Entry Has Not Been Changed, Would you like to Make a New Entry?", vbYesNo, "Go to New") = vbNo Then
'        txtDate.SetFocus
'        Exit Sub
'      End If
'    End If
'  End If
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
  If txtAmount <> 0 Then
    If txtUndist <> 0 Or txtAmount <> txtDist Then
      MsgBox "The Total Distributed Does Not Equal The Amount of The Disbursement." & Chr$(13) & "Please Correct Before Saving.", vbOKOnly, "Cash Disbursement"
      Ready2Save = False
      Exit Function
    Else
      
      For cnt = 1 To 36
        vaSpread1.Col = 4
        vaSpread1.Row = cnt
        If vaSpread1.Text <> "" Then
          TempDist = Round#(vaSpread1.Text + TempDist)
        Else
          Exit For
        End If
      Next
      If TempDist <> txtDist Or TempDist <> txtAmount Then
        MsgBox "Totals Are Not In Balance. Please Correct.", vbOKOnly, "Cash Disbursements"
        Ready2Save = False
        Exit Function
      Else
        Ready2Save = True
      End If
     End If
  Else
    MsgBox "You May Not Save A Cash Disbursement With A $0.00 Total.", vbOKOnly, "Cash Disbursements"
    Ready2Save = False
  End If
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbYes Then
      If Changed = False Then
        Undolok RecNum
        ClearInUse PWcnt
      Else
        If MsgBox("Abandon Changes and Close?", vbYesNo, "Close?") = vbYes Then
          Undolok RecNum
          ClearInUse PWcnt
        Else
          Cancel = True
        End If
      End If
    Else
      Cancel = True
    End If
  End If
End Sub

Private Sub cmdSave_Click()
  If Ready2Save = True Then
    SaveCashDisb
    Call NextNew
  Else
    MsgBox "             Save Canceled.", vbOKOnly, "Cash Disbursement"
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
      SendKeys "%X"
      KeyCode = 0
    Case vbKeyF9:
      SendKeys "%A"
      KeyCode = 0
    Case vbKeyF3:
      SendKeys "%D"
      KeyCode = 0
    Case vbKeyF6:
      SendKeys "%i"
      KeyCode = 0
    Case vbKeyF2:
      SendKeys "%N"
      KeyCode = 0
    Case vbKeyF4:
      SendKeys "%E"
      KeyCode = 0
    Case vbKeyF5:
      SendKeys "%L"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
      KeyCode = 0
    Case Else:
  End Select
End Sub


Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub
Private Sub Undolok(OldRec)
  Dim CJEditFile As Integer, CJEditFileNum As Integer, NumEdTrans As Integer, cnt As Integer
  Dim CDBusy As Boolean
  If Exist("GLCDEd.DAT") Then CDBusy = GetAttr("GLCDEd.DAT") And vbReadOnly
  If Not CDBusy Then
    OpenCJEditFile CJEditFileNum, NumEdTrans, CJType
      If OldRec <= NumEdTrans Then
        Get CJEditFileNum, OldRec, GLCDEd(1)
        GLCDEd(1).LOCKED = False
        Put CJEditFileNum, OldRec, GLCDEd(1)
      End If
      Close CJEditFileNum
  Else
    MsgBox "Posting In Progress, Editing May Not Continue At This Time.", vbOKOnly, "Canceled"
    frmCashDisbMenu.Show
    Unload frmCashDisbEntry
  End If
End Sub


Private Sub cmdExit_Click()
  If Changed = False Then
    Undolok RecNum
    frmCashDisbMenu.Show
    Unload frmCashDisbEntry
  Else
    If MsgBox("Changes Have Been Made to the Current Record." & Chr(13) & Chr(13) & "                          Select OK to Abandon," & Chr(13) & Chr(13) & "       or Cancel to Remain on Entry/Edit Screen.", vbOKCancel, "Abandon Changes?") = vbOK Then
      Undolok RecNum
      frmCashDisbMenu.Show
      Unload frmCashDisbEntry
    Else
      txtDate.SetFocus
    End If
  End If
End Sub

'***********Use CJType 2 (for Disbursement) Pass to open file,
'********* so will open correct file - GLCDEd.dat
Private Sub Form_Load()
  Dim SetupFile As Integer, Rec As Integer
  Dim GLAcctidx As GLAcctIndexType
  Dim CJEditFile As Integer, CJEditFileNum As Integer, NumEdTrans As Integer
  Dim AcctIdxFileNum As Integer, actlist As String
  Dim NumAIdxRecs As Integer, x As Integer
  Dim cnt As Integer, LCnt As Integer
  RecLok = False
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  CJType = 2
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  ActLen = (GLFundLen + GLAcctLen + GLDetLen + 2)
  GetPostDates LPDate, HPDate  'In Main Module to get dates from setup
  StatusBar1.Panels.Item(1).Text = GLUserName
  Me.HelpContextID = hlpEntryEditCash
  GetCentDep CDActive, CashAcct, CDCash, CDDue
  Fixspread
  FillAcctNumName fpcboAcctNumNa
'  If CDActive = "Y" Then
'    txtBanks.AddItem ("99 System")
'  Else
  GetBankList txtBanks
  SetDefBank "D", DefBnk
  If DefBnk > 0 Then
    'txtBanks.Col = 0
    txtBanks.SearchText = Trim(DefBnk)
    txtBanks.Action = 0
    If txtBanks.SearchIndex <> -1 Then
      txtBanks.ListIndex = txtBanks.SearchIndex
    End If
  Else
    txtBanks.ListIndex = 0
  End If

'  End If
  OpenCJEditFile CJEditFileNum, NumEdTrans, CJType
  If NumEdTrans > 0 Then
    For Rec = 1 To NumEdTrans
      Get CJEditFileNum, Rec, GLCDEd(1)
      If GLCDEd(1).DelFlag = 0 Then
        'RecNum = Rec
        'Emode = True
        RecLok = True
        Exit For
      End If
    Next
  End If
  Close CJEditFileNum
  If RecLok = True Then
    RecNum = NumEdTrans + 1
  Else
    RecNum = 1
  End If
    Emode = False
    SetScreen
    txtDate.Text = Format(Now, "mm/dd/yyyy")
    txtDesc = ""
    txtAdditDesc = ""
    txtCheckNum = ""
    txtAmount = 0
    txtBatch = ""
    'txtBanks.ListIndex = -1
    fpcboAcctNumNa.ListIndex = -1
    txtDebAmt = 0
'***** spreadsheet do Not have to set blank fields on load ..
    txtUndist = 0
    txtDist = 0
  
End Sub
Public Sub FirstOpenCD()
  If RecLok = True Then
    frmCDListing.Show 1, frmCashDisbEntry
  End If
End Sub

Private Sub fpcboAcctNumNa_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboAcctNumNa.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboAcctNumNa.ListIndex = -1
    fpcboAcctNumNa.Action = ActionClearSearchBuffer
  End If
  If fpcboAcctNumNa.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
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

Private Sub mnuExit_Click()
  Call cmdExit_Click
End Sub

Private Sub mnuPrint_Click()
'Printer.Print
End Sub

Private Sub SaveCashDisb()
  Dim CJEditFile As Integer, CJEditFileNum As Integer, NumEdTrans As Integer, cnt As Integer
  Dim CDBusy As Boolean
  If Exist("GLCDEd.DAT") Then CDBusy = GetAttr("GLCDEd.DAT") And vbReadOnly
  If Not CDBusy Then
    OpenCJEditFile CJEditFileNum, NumEdTrans, CJType
    If Emode = False Then
      If NumEdTrans > 0 Then
        RecNum = NumEdTrans + 1
      Else
        RecNum = 1
      End If
    Else
      Get CJEditFileNum, RecNum, GLCDEd(1)
    End If
  
    GLCDEd(1).DelFlag = 0
    GLCDEd(1).TRDATE = DateDiff("d", "12/31/1979", txtDate)
    GLCDEd(1).Desc = Trim(txtDesc)
    GLCDEd(1).LDesc = Trim(txtAdditDesc)
    GLCDEd(1).LOCKED = False
    GLCDEd(1).DOCREF = Trim(txtCheckNum)
    GLCDEd(1).Amt = txtAmount.DoubleValue()
    GLCDEd(1).BATCHNUM = Trim(txtBatch)
    GLCDEd(1).RECCODE = txtBanks.Text
    For cnt = 1 To 36
      vaSpread1.Row = cnt
      vaSpread1.Col = 1
      If vaSpread1.Text = "" Then
        GLCDEd(1).Dist(cnt).DACREC = 0
        GLCDEd(1).Dist(cnt).DACN = 0
        GLCDEd(1).Dist(cnt).DACNM = 0
        GLCDEd(1).Dist(cnt).DAMT = 0
      Else
        GLCDEd(1).Dist(cnt).DACREC = vaSpread1.Text
        vaSpread1.Col = 2
        GLCDEd(1).Dist(cnt).DACN = QPTrim(vaSpread1.Text)
        vaSpread1.Col = 3
        GLCDEd(1).Dist(cnt).DACNM = QPTrim(vaSpread1.Text)
        vaSpread1.Col = 4
        GLCDEd(1).Dist(cnt).DAMT = vaSpread1.Text
        'txtDist = (txtDist.DoubleValue + GLCDEd(1).Dist(Cnt).DAMT)
      
        'Exit For
      End If
    Next
    Put CJEditFileNum, RecNum, GLCDEd(1)
    Close CJEditFileNum
    Call MainLog("CD Saved.")
  Else
    MsgBox "Posting In Progress, Editing May Not Continue At This Time.", vbOKOnly, "Canceled"
    frmCashDisbMenu.Show
    Unload frmCashDisbEntry
  End If
End Sub
  
Private Function SetScreen()
  If Emode = False Then  'This is in New Mode
    cmdNew.Enabled = False
    cmdEdit.Enabled = True
    lblNew.Visible = True
    lblEdit.Visible = False
  Else               'This is in Edit Mode
    cmdNew.Enabled = True
    cmdEdit.Enabled = False
    lblNew.Visible = False
    lblEdit.Visible = True
  End If
End Function
Public Function Rec2Form(TempRec)
  Dim CJEditFile As Integer, CJEditFileNum As Integer, NumEdTrans As Integer
  Dim CurrRec As Integer, NextRec As Integer, cnt As Integer, Last As Integer
  Dim CDBusy As Boolean
  If Exist("GLCDEd.DAT") Then CDBusy = GetAttr("GLCDEd.DAT") And vbReadOnly
  If Not CDBusy Then
    OpenCJEditFile CJEditFileNum, NumEdTrans, CJType
    'If used list or edit the temprec was selected there, need to transfer to recnum
    OldRec = RecNum
    RecNum = TempRec
    Get CJEditFileNum, RecNum, GLCDEd(1)
    If GLCDEd(1).LOCKED = False Then
      txtDate = Format(DateAdd("d", (GLCDEd(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
      txtDesc = GLCDEd(1).Desc
      txtAdditDesc = GLCDEd(1).LDesc
      txtCheckNum = GLCDEd(1).DOCREF
      txtAmount = GLCDEd(1).Amt
      txtBatch = GLCDEd(1).BATCHNUM
      txtBanks.Text = GLCDEd(1).RECCODE
      Last = UBound(GLCDEd(1).Dist)
    
      For cnt = 1 To Last
        'If Trim(GLCDEd(1).Dist(Cnt).DACN) <> "" Then
        If GLCDEd(1).Dist(cnt).DACREC <> 0 Then
          vaSpread1.Row = vaSpread1.DataRowCnt + 1
    '***** Until complete testing display the record num here to make sure correct.
    '****** Fixed the Spread to Hide this column so can leave
          vaSpread1.Col = 1
          vaSpread1.Text = GLCDEd(1).Dist(cnt).DACREC
          vaSpread1.Col = 2
          vaSpread1.Text = GLCDEd(1).Dist(cnt).DACN
          vaSpread1.Col = 3
          vaSpread1.Text = GLCDEd(1).Dist(cnt).DACNM
          vaSpread1.Col = 4
          vaSpread1.Text = GLCDEd(1).Dist(cnt).DAMT
          txtDist = Round#(txtDist.DoubleValue + GLCDEd(1).Dist(cnt).DAMT)
        Else
          Exit For
        End If
      Next
      txtUndist = Round#(txtAmount.DoubleValue - txtDist.DoubleValue)
      GLCDEd(1).LOCKED = True
      Put CJEditFileNum, RecNum, GLCDEd(1)
      Emode = True
      Close CJEditFileNum
      Undolok OldRec
    Else
      MsgBox "Record Is Being Edited By Another User.", vbOKOnly, "Record Unavailable"
      RecNum = OldRec
      Close CJEditFileNum
    End If
    SetScreen
  Else
    MsgBox "Posting In Progress, Editing May Not Continue At This Time.", vbOKOnly, "Canceled"
    frmCashDisbMenu.Show
    Unload frmCashDisbEntry
  End If
End Function
Private Function Changed()
  Dim CJEditFile As Integer, CJEditFileNum As Integer, NumEdTrans As Integer
  Dim cnt As Integer
  If Emode = False Then
    If Val(txtDesc) <> 0 Then
      Changed = True
      Exit Function
    Else
    If Len(QPTrim$(txtAdditDesc)) > 0 Then
      Changed = True
      Exit Function
    Else
      If Val(txtCheckNum) <> 0 Then
        Changed = True
        Exit Function
      Else
        If txtAmount <> 0 Then
          Changed = True
          Exit Function
        Else
          If Val(txtBatch) <> 0 Then
            Changed = True
            Exit Function
          Else
            'If Val(txtBanks.Text) <> 0 Then
              'Changed = True
              'Exit Function
            'Else
             ' Changed = False
            'End If
          End If
        End If
      End If
      End If
    End If
  Else
    OpenCJEditFile CJEditFileNum, NumEdTrans, CJType
    Get CJEditFileNum, RecNum, GLCDEd(1)
    If txtDate <> Format(DateAdd("d", (GLCDEd(1).TRDATE), "12-31-1979"), "mm/dd/yyyy") Then
      Changed = True
      Close CJEditFileNum
      Exit Function
    Else
      If txtDesc <> GLCDEd(1).Desc Then
        Changed = True
        Close CJEditFileNum
        Exit Function
      Else
      If QPTrim$(txtAdditDesc) <> GLCDEd(1).LDesc Then
        Changed = True
        Close CJEditFileNum
        Exit Function
      Else
        If txtCheckNum <> GLCDEd(1).DOCREF Then
          Changed = True
          Close CJEditFileNum
          Exit Function
        Else
          If txtAmount.DoubleValue() <> GLCDEd(1).Amt Then
            Changed = True
            Close CJEditFileNum
            Exit Function
          Else
            If txtBatch <> GLCDEd(1).BATCHNUM Then
              Changed = True
              Close CJEditFileNum
              Exit Function
            Else
              If txtBanks.Text <> GLCDEd(1).RECCODE Then
                Changed = True
                Close CJEditFileNum
                Exit Function
              Else
                Changed = False
                
              End If
            End If
          End If
        End If
        End If
      End If
    End If
    If txtAmount.DoubleValue <> txtDist.DoubleValue Then
      Changed = True
      Close CJEditFileNum
      Exit Function
    End If
    If fpcboAcctNumNa.Text <> "" Then
      Changed = True
      Close CJEditFileNum
      Exit Function
    Else
      For cnt = 1 To 36
        vaSpread1.Row = cnt
        vaSpread1.Col = 1
        If Val(vaSpread1.Text) = GLCDEd(1).Dist(cnt).DACREC Then
          If Val(vaSpread1.Text) = 0 Then
            Changed = False
            Exit For
          Else
            vaSpread1.Col = 2
            If vaSpread1.Text = GLCDEd(1).Dist(cnt).DACN Then
              vaSpread1.Col = 3
              If vaSpread1.Text = GLCDEd(1).Dist(cnt).DACNM Then
                vaSpread1.Col = 4
                If vaSpread1.Text = GLCDEd(1).Dist(cnt).DAMT Then
                  Changed = False
                Else
                  Changed = True
                  Exit For
                End If
              Else
                Changed = True
                Exit For
              End If
            Else
              Changed = True
              Exit For
            End If
          End If
        Else
          Changed = True
          Exit For
        End If
      Next
    Close CJEditFileNum
    End If
  End If
End Function

Private Sub NextNew()
  Dim CJEditFile As Integer, CJEditFileNum As Integer, NumEdTrans As Integer
  OpenCJEditFile CJEditFileNum, NumEdTrans, CJType
  Close CJEditFileNum
   If NumEdTrans > 0 Then
     RecNum = NumEdTrans + 1
   Else
     RecNum = 1
   End If

   Emode = False
   ClearFields
   SetScreen
   txtDate.SetFocus
  
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

'Private Sub fpcboAcctNumNa_GotFocus()
'  If VerifyEntered = False Then
'    MsgBox "The Information In The Top Section Must Be Completed Before The Distributions.", vbOKOnly, "Cash Disbursement"
'
'  End If
'End Sub

Private Sub txtAmount_Change()
  txtUndist = txtAmount - txtDist
End Sub


Private Sub txtBanks_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    txtBanks.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    txtBanks.ListIndex = -1
    txtBanks.Action = ActionClearSearchBuffer
  End If
  If txtBanks.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      SendKeys "{Tab}"
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        SendKeys "+{Tab}"
        KeyCode = 0
      End If
    End If
  txtBanks.Action = ActionClearSearchBuffer
  End If
End Sub

Private Sub txtDate_LostFocus()
  If CheckValDate(txtDate) = False Then
    MsgBox "Invalid Date, Please Correct.", vbOKOnly, "Invalid Date"
    txtDate.SetFocus
  End If
End Sub
Private Function Check4Trans()
  Dim CJEditFile As Integer, CJEditFileNum As Integer, NumEdTrans As Integer
  Dim cnt As Integer, Good As Integer
  Good = 0
  If Exist("GLCDEd.dat") Then
    OpenCJEditFile CJEditFileNum, NumEdTrans, CJType
    If NumEdTrans > 0 Then
      For cnt = 1 To NumEdTrans
        Get CJEditFileNum, cnt, GLCDEd(1)
        If GLCDEd(1).DelFlag = 0 Then
          Good = Good + 1
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
 Close CJEditFileNum
 End Function
'**** To PS 9-21
'******What was I Doing, Check this
Private Function VerifyEntered()
  If txtDesc <> "" Or txtDesc <> " " Then
    If txtCheckNum <> "" Or txtCheckNum <> " " Then
      If txtAmount <> 0 Then
        If txtBatch <> "" Or txtBatch <> " " Then
          If txtBanks.Text <> "" Then
            VerifyEntered = True
          Else
            VerifyEntered = False
            txtBanks.SetFocus
            Exit Function
          End If
        Else
          VerifyEntered = False
          txtBatch.SetFocus
          Exit Function
        End If
      Else
        VerifyEntered = False
        txtAmount.SetFocus
        Exit Function
      End If
    Else
      VerifyEntered = False
      txtCheckNum.SetFocus
      Exit Function
    End If
  Else
    VerifyEntered = False
    txtDesc.SetFocus
    Exit Function
  End If
End Function

'Private Sub txtDebAmt_GotFocus()
'  If VerifyEntered = False Then
'    MsgBox "The Information In The Top Section Must Be Filled Out Before The Distributions.", vbOKOnly, "Cash Disbursement"
'  End If
'End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
  Dim TempAcct As String
  Dim TempCol As Long, TempRow As Long
  TempRow = Row
  TempCol = Col
  If TempRow > 0 Then
    vaSpread1.Row = TempRow
    vaSpread1.Col = 2
    TempAcct = QPTrim(vaSpread1.Text)
    If vaSpread1.Text <> "" Then
      If fpcboAcctNumNa.ListIndex <> -1 Or txtDebAmt <> 0 Then
        If MsgBox("Do You Wish To Abandon Current Distribution?, 'Yes' or 'No' Complete Distribution Entry.", vbYesNo, "Clear??") = vbNo Then
          cmdAddDist.SetFocus
          Exit Sub
        Else
          fpcboAcctNumNa.ListIndex = -1
          txtDebAmt = 0
        End If
      End If
    
        fpcboAcctNumNa.SearchText = QPStrip(TempAcct)
        fpcboAcctNumNa.Action = 0
        If fpcboAcctNumNa.SearchIndex <> -1 Then
          fpcboAcctNumNa.ListIndex = fpcboAcctNumNa.SearchIndex
        End If
          vaSpread1.Col = 1
          fpcboAcctNumNa.Col = 0
          fpcboAcctNumNa.ColText = vaSpread1.Text
          vaSpread1.Col = 2
          fpcboAcctNumNa.Col = 1
          fpcboAcctNumNa.ColText = vaSpread1.Text
          vaSpread1.Col = 3
          fpcboAcctNumNa.Col = 2
          fpcboAcctNumNa.ColText = vaSpread1.Text
          vaSpread1.Col = 4
          txtDebAmt = vaSpread1.Text
          txtDist = Round#(txtDist.DoubleValue - txtDebAmt.DoubleValue)
          txtUndist = Round#(txtAmount.DoubleValue - txtDist.DoubleValue)
          'vaSpread1.ClearRange TempCol, TempRow, 4, TempRow, True
          vaSpread1.DeleteRows TempRow, 1
          fpcboAcctNumNa.SetFocus
    End If
  End If
End Sub
Public Sub ClearFields()
  
  txtDesc = ""
  txtAdditDesc = ""
  txtCheckNum = ""
  txtAmount = 0
  txtBatch = ""
  txtBanks.Text = DefBnk
  
  fpcboAcctNumNa.ListIndex = -1
  txtDebAmt = 0
'*****Clear data in the spreadsheet
  vaSpread1.ClearRange 1, 1, 4, 36, True
  txtUndist = 0
  txtDist = 0
'  txtDate.SetFocus
End Sub


