VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Begin VB.Form frmCashReceiptEntry 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash Receipt Entry"
   ClientHeight    =   8640
   ClientLeft      =   30
   ClientTop       =   540
   ClientWidth     =   12195
   Icon            =   "frmCashReceiptsEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboAcctNumNa 
      Height          =   405
      Left            =   1875
      TabIndex        =   7
      Top             =   2910
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
      ColDesigner     =   "frmCashReceiptsEntry.frx":08CA
   End
   Begin LpLib.fpCombo txtBanks 
      Height          =   375
      Left            =   9480
      TabIndex        =   6
      Top             =   2010
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
      ScrollBarH      =   3
      DataFieldList   =   ""
      ColumnEdit      =   0
      ColumnBound     =   -1
      Style           =   2
      MaxDrop         =   8
      ListWidth       =   3504
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
      EditAlignH      =   0
      EditAlignV      =   0
      ColDesigner     =   "frmCashReceiptsEntry.frx":0D3B
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
      Left            =   5418
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7728
      Width           =   1356
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
      Left            =   10362
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7728
      Width           =   1356
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
      Left            =   8856
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2664
      Width           =   1668
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
      Left            =   3770
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7728
      Width           =   1356
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
      Left            =   8714
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7728
      Width           =   1356
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
      Left            =   2122
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7728
      Width           =   1356
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
      Left            =   7066
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7728
      Width           =   1356
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
      Left            =   474
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7728
      Width           =   1356
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   3060
      Left            =   1770
      TabIndex        =   10
      Top             =   3720
      Width           =   8625
      _Version        =   196613
      _ExtentX        =   15214
      _ExtentY        =   5397
      _StockProps     =   64
      AutoClipboard   =   0   'False
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
      GridColor       =   8421504
      MaxCols         =   4
      MaxRows         =   36
      OperationMode   =   3
      ScrollBars      =   2
      SelectBlockOptions=   0
      ShadowColor     =   13684944
      ShadowDark      =   12632256
      ShadowText      =   0
      SpreadDesigner  =   "frmCashReceiptsEntry.frx":1144
      VisibleCols     =   3
      VisibleRows     =   10
      ScrollBarTrack  =   1
   End
   Begin EditLib.fpCurrency txtCrAmt 
      Height          =   372
      Left            =   7152
      TabIndex        =   8
      Top             =   2928
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
      AlignTextV      =   1
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
   Begin EditLib.fpText txtDocRef 
      Height          =   324
      Left            =   9480
      TabIndex        =   2
      Top             =   1344
      Width           =   1092
      _Version        =   196608
      _ExtentX        =   1926
      _ExtentY        =   572
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
   Begin EditLib.fpCurrency txtDist 
      Height          =   372
      Left            =   7872
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   7128
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
   Begin EditLib.fpCurrency txtUndist 
      Height          =   372
      Left            =   4008
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   7128
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
   Begin EditLib.fpCurrency txtAmount 
      Height          =   372
      Left            =   2664
      TabIndex        =   4
      Top             =   1920
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
   Begin EditLib.fpDateTime txtDate 
      Height          =   372
      Left            =   2448
      TabIndex        =   0
      Top             =   1392
      Width           =   1740
      _Version        =   196608
      _ExtentX        =   3069
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
      Text            =   "10/16/2001"
      DateCalcMethod  =   0
      DateTimeFormat  =   5
      UserDefinedFormat=   "mm/dd/yyyy"
      DateMax         =   "20350101"
      DateMin         =   "19800101"
      TimeMax         =   "000000"
      TimeMin         =   "000000"
      TimeString1159  =   ""
      TimeString2359  =   ""
      DateDefault     =   "00000000"
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
      ButtonColor     =   14737632
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   20
      Top             =   8400
      Width           =   12192
      _ExtentX        =   21511
      _ExtentY        =   423
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
            TextSave        =   "8:14 AM"
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
   Begin EditLib.fpText txtDesc 
      Height          =   324
      Left            =   5424
      TabIndex        =   1
      Top             =   1392
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
   Begin EditLib.fpText txtAdditDesc 
      Height          =   324
      Left            =   5400
      TabIndex        =   5
      Top             =   2016
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
      Text            =   ""
      CharValidationText=   ""
      MaxLength       =   32
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
   Begin EditLib.fpText txtBatch 
      Height          =   324
      Left            =   9480
      TabIndex        =   3
      Top             =   1680
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
      ForeColor       =   &H8000000E&
      Height          =   372
      Index           =   1
      Left            =   8616
      TabIndex        =   35
      Top             =   2064
      Width           =   732
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
      ForeColor       =   &H8000000E&
      Height          =   372
      Index           =   2
      Left            =   8640
      TabIndex        =   34
      Top             =   1728
      Width           =   732
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
      ForeColor       =   &H8000000E&
      Height          =   372
      Index           =   1
      Left            =   4560
      TabIndex        =   33
      Top             =   1416
      Width           =   732
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
      ForeColor       =   &H8000000E&
      Height          =   372
      Index           =   3
      Left            =   4824
      TabIndex        =   32
      Top             =   1752
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
      ForeColor       =   &H8000000E&
      Height          =   348
      Left            =   1440
      TabIndex        =   31
      Top             =   3360
      Width           =   1884
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   612
      Left            =   2952
      Top             =   264
      Width           =   6132
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Receipt Entry/Edit"
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
      Left            =   3792
      TabIndex        =   30
      Top             =   384
      Width           =   4452
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000016&
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   1188
      Index           =   0
      Left            =   1488
      Top             =   1272
      Width           =   9180
   End
   Begin VB.Label Label2b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
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
      Height          =   372
      Index           =   1
      Left            =   1608
      TabIndex        =   29
      Top             =   1896
      Width           =   972
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
      Height          =   372
      Index           =   1
      Left            =   1728
      TabIndex        =   28
      Top             =   1416
      Width           =   612
   End
   Begin VB.Label Label3b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Doc Ref"
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
      Height          =   372
      Index           =   0
      Left            =   8184
      TabIndex        =   27
      Top             =   1416
      Width           =   1188
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
      ForeColor       =   &H8000000E&
      Height          =   372
      Index           =   2
      Left            =   2448
      TabIndex        =   26
      Top             =   7128
      Width           =   1452
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
      ForeColor       =   &H8000000E&
      Height          =   372
      Index           =   3
      Left            =   6456
      TabIndex        =   25
      Top             =   7128
      Width           =   1332
   End
   Begin VB.Label lblNew 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "New Cash Receipt"
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
      Left            =   2952
      TabIndex        =   24
      Top             =   984
      Visible         =   0   'False
      Width           =   2772
   End
   Begin VB.Label lblEdit 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "Edit Cash Receipt"
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
      Left            =   6432
      TabIndex        =   23
      Top             =   984
      Visible         =   0   'False
      Width           =   2652
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
      ForeColor       =   &H8000000E&
      Height          =   372
      Index           =   0
      Left            =   2592
      TabIndex        =   22
      Top             =   2544
      Width           =   3132
   End
   Begin VB.Label Label3b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Amount"
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
      Height          =   372
      Index           =   5
      Left            =   6912
      TabIndex        =   21
      Top             =   2544
      Width           =   1692
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   4572
      Left            =   1488
      Top             =   2448
      Width           =   9180
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      Height          =   732
      Left            =   2952
      Top             =   144
      Width           =   6132
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   3225
      Left            =   1680
      Top             =   3675
      Width           =   8790
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      Height          =   492
      Left            =   1488
      Top             =   7008
      Width           =   9180
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
Attribute VB_Name = "frmCashReceiptEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GLSetup As GLSetupRecType
Dim GLAcct As GLAcctRecType
Dim GJEdit As TrEditRecType
Dim GLTrans As GLTransRecType
Dim GLCREd(1) As CJEditRecType
Dim Over As clsTextBoxOverRider
Dim LPDate As Integer, HPDate As Integer, CJType As Integer
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim Emode As Boolean, RecNum As Integer, RecLok As Boolean, OldRec As Integer
Dim ActLen As Integer, DefBnk As Integer
Dim CDActive As String, CashAcct As String, CDCash As String, CDDue As String
Dim RetValue As Integer
Private Temp_Class As Resize_Class
'This is to fix spreadsheet for various resolutions
Public Function Fixspread()
 '   GoTo DontDoIt
'    Select Case screenW
'      Case 1280
'      If Screen.TwipsPerPixelX <> 12 Then
'        coladj = 27.3
'        vaSpread1.RowHeight(-1) = 22.5
'        vaSpread1.RowHeight(0) = 22.5
'      Else
'        coladj = 20
'        vaSpread1.RowHeight(-1) = 18
'        vaSpread1.RowHeight(0) = 18
'      End If
'      Case 1152
'      If Screen.TwipsPerPixelX <> 12 Then
'        coladj = 24
'        vaSpread1.RowHeight(-1) = 19.2
'        vaSpread1.RowHeight(0) = 19.2
'      Else
'        coladj = 17
'        vaSpread1.RowHeight(-1) = 15
'        vaSpread1.RowHeight(0) = 15
'      End If
'      Case 1024
'      If Screen.TwipsPerPixelX <> 12 Then
'        coladj = 20
'        vaSpread1.RowHeight(0) = 18
'        vaSpread1.RowHeight(-1) = 18
'      Else
'        coladj = 14
'      End If
'      Case 800
'        coladj = 13.8
'        'vaSpread1.Font.Size = 8
'        'vaSpread1.RowHeight(-1) = 13
'      Case Else
'        'don't worry be happpy
'    End Select
    vaSpread1.Font.Size = 6
    vaSpread1.RowHeight(0) = 18
    vaSpread1.RowHeight(-1) = 18

    'vaSpread1.ColWidth(-1) = vaSpread1.ColWidth(-1) + coladj
DontDoIt:
End Function

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

'**************************
'Use Emode to determine if New record or Editing, so true if editing.
'RecNum passed from Listing to load chosen record to form
'******************************
Private Sub cmdDelete_Click()
  Dim CJREditFile As Integer, CJEditFileNum As Integer, NumEdTrans As Integer
  Dim CRBusy As Boolean
  If Exist("GLCREd.DAT") Then CRBusy = GetAttr("GLCREd.DAT") And vbReadOnly
  If Not CRBusy Then
    If MsgBox("Are you sure you wish to delete this entry?", vbYesNo, "Delete GJEntry") = vbYes Then
      If Emode = True Then
        OpenCJEditFile CJEditFileNum, NumEdTrans, CJType
        GLCREd(1).DelFlag = -1
        GLCREd(1).LOCKED = False
        Put CJEditFileNum, RecNum, GLCREd(1)
        Close CJEditFileNum
        GLCREd(1).DelFlag = 0
        Call MainLog("CR Deleted.")
        Call NextNew
      Else
        Call NextNew
      End If
    Else
      txtDate.SetFocus
    End If
  Else
    MsgBox "Posting Is In Progress, Editing Not Allowed At This Time.", vbOKOnly, "Request Canceled"
    frmCashReceiptsMenu.Show
    Unload frmCashReceiptEntry
  End If
End Sub

Private Sub cmdEdit_Click()
 If Changed = True Then
    If MsgBox("Changes Were Made to the Current Information on Screen and Not Saved." & Chr(13) & "Select OK to View Edit List," & Chr(13) & "or Cancel to Remain on Current Record.", vbOKCancel, "View Edit List?") = vbCancel Then
      txtDate.SetFocus
      Exit Sub
    End If
  End If
  Undolok RecNum
  NextNew
  If Check4Trans = True Then
    frmCRListing.Show 1, frmCashReceiptEntry
    If Emode = True Then
      SetScreen
      'DisplayTotals
      txtDate.SetFocus
      cmdDelete.Enabled = True
    End If
  Else
    MsgBox "No Entries To Display.", vbOKOnly, "No Entries"
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
    frmCRListing.Show 1, frmCashReceiptEntry
    If Emode = True Then
      SetScreen
      txtDate.SetFocus
    End If
  Else
    'RetValue = sndPlaySound("noentries.wav", SND_ASYNC Or SND_NODEFAULT)
    MsgBox "No Entries To Display.", vbOKOnly, "No Entries"
    txtDate.SetFocus
  End If
End Sub

Private Sub cmdNew_Click()
  Dim CJREditFile As Integer, CJEditFileNum As Integer, NumEdTrans As Integer
  Dim CRBusy As Boolean
  If Exist("GLCREd.DAT") Then CRBusy = GetAttr("GLCREd.DAT") And vbReadOnly
  If Not CRBusy Then
    OpenCJEditFile CJEditFileNum, NumEdTrans, CJType
    If Changed = True Then
      If MsgBox("Changes Have Been Made to the Current Record." & Chr(13) & "Select OK to Abandon," & Chr(13) & "or Cancel to Remain on Current Record.", vbOKCancel, "Abandon Changes?") = vbCancel Then
        txtDate.SetFocus
        Exit Sub
      End If
    End If
    Get CJEditFileNum, RecNum, GLCREd(1)
    GLCREd(1).LOCKED = False
    Put CJEditFileNum, RecNum, GLCREd(1)
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
    MsgBox "Posting Is In Progress, Editing Not Allowed At This Time.", vbOKOnly, "Request Canceled"
    frmCashReceiptsMenu.Show
    Unload frmCashReceiptEntry
  End If
End Sub

Private Sub cmdAddDist_Click()
  If VerifyEntered = False Then
    MsgBox "The Information In The Top Section Must Be Completed Before Adding Distributions.", vbOKOnly, "Cash Receipt"
  Else
    If fpcboAcctNumNa.Text <> "" And txtCrAmt.DoubleValue <> 0 Then
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
      txtDist = Round#(txtCrAmt.DoubleValue + txtDist.DoubleValue)
      vaSpread1.Text = txtCrAmt
      fpcboAcctNumNa.ListIndex = -1
      txtCrAmt = 0
      fpcboAcctNumNa.SetFocus
      txtUndist = Round#(txtAmount.DoubleValue - txtDist.DoubleValue)
    Else
      MsgBox "Only 36 Distributions Allowed.", vbOKOnly, "Limit Reached"
    End If
    Else
      MsgBox "The Account and Amount Must Be Entered Before Adding To The Distribution List.", vbOKOnly, "Add Distribution Denied"
    End If
  End If
End Sub

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
      MsgBox "The Total Distributed Does Not Equal The Amount of The Receipt." & Chr$(13) & "Please Correct Before Saving.", vbOKOnly, "Cash Receipt"
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
        MsgBox "Totals Are Not In Balance. Please Correct.", vbOKOnly, "Cash Receipts"
        Ready2Save = False
        Exit Function
      Else
        Ready2Save = True
      End If
     End If
  Else
    MsgBox "You May Not Save A Cash Receipt With A $0.00 Total.", vbOKOnly, "Cash Receipts"
    Ready2Save = False
  End If
End Function

Private Sub cmdSave_Click()
  If Ready2Save = True Then
    SaveCashRecpt
    Call NextNew
  Else
    MsgBox "        Save Canceled.", vbOKOnly, "Cash Receipt"
  End If
End Sub

Private Sub fpcboAcctNumNa_LostFocus()
  fpcboAcctNumNa.Action = ActionClearSearchBuffer
End Sub


Private Sub mnuPrnScn_Click()
  PrintForm
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
    Case vbKeyF6:
      SendKeys "%i"
      KeyCode = 0
    Case vbKeyF3:
      SendKeys "%D"
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

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub cmdExit_Click()
  If Changed = False Then
    Undolok RecNum
    frmCashReceiptsMenu.Show
    Unload frmCashReceiptEntry
  Else
    If MsgBox("Changes Have Been Made to the Current Record." & Chr(13) & Chr(13) & "                          Select OK to Abandon," & Chr(13) & Chr(13) & "       or Cancel to Remain on Entry/Edit Screen.", vbOKCancel, "Abandon Changes?") = vbOK Then
      Undolok RecNum
      frmCashReceiptsMenu.Show
      Unload frmCashReceiptEntry
    Else
      txtDate.SetFocus
    End If
  End If
End Sub
Private Sub Undolok(OldRec)
  Dim CJEditFile As Integer, CJEditFileNum As Integer, NumEdTrans As Integer, cnt As Integer
  Dim CRBusy As Boolean
  If Exist("GLCREd.DAT") Then CRBusy = GetAttr("GLCREd.DAT") And vbReadOnly
  If Not CRBusy Then
    OpenCJEditFile CJEditFileNum, NumEdTrans, CJType
    If OldRec <= NumEdTrans Then
      Get CJEditFileNum, OldRec, GLCREd(1)
      GLCREd(1).LOCKED = False
      Put CJEditFileNum, OldRec, GLCREd(1)
    End If
    Close CJEditFileNum
  Else
    MsgBox "Posting Is In Progress, Editing Not Allowed At This Time.", vbOKOnly, "Request Canceled"
    frmCashReceiptsMenu.Show
    Unload frmCashReceiptEntry
  End If
End Sub

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
  CJType = 1 'This passes to Open file procedure to determine File name to open for Receipt
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  ActLen = (GLFundLen + GLAcctLen + GLDetLen + 2)
  GetPostDates LPDate, HPDate  'In Main Module to get dates from setup
  StatusBar1.Panels.Item(1).Text = GLUserName
  Me.HelpContextID = hlpEnterEditCash
  GetCentDep CDActive, CashAcct, CDCash, CDDue
  Fixspread
  FillAcctNumName fpcboAcctNumNa
'  If CDActive = "Y" Then
'    txtBanks.AddItem ("99    System")
'  Else
    GetBankList txtBanks
  SetDefBank "R", DefBnk
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
      Get CJEditFileNum, Rec, GLCREd(1)
      If GLCREd(1).DelFlag = 0 Then
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
    txtDocRef = ""
    txtAmount = 0
    txtBatch = ""
    'txtBanks.ListIndex =
    fpcboAcctNumNa.ListIndex = -1
    txtCrAmt = 0
'*****What about the spreadsheet do I have to set blank fields ??? Not on load..
    txtUndist = 0
    txtDist = 0
End Sub
Public Sub FirstOpenCR()
  If RecLok = True Then
    frmCRListing.Show 1, frmCashReceiptEntry
  End If
End Sub

Private Sub mnuExit_Click()
  Call cmdExit_Click
End Sub

Private Sub SaveCashRecpt()
  Dim CJREditFile As Integer, CJEditFileNum As Integer, NumEdTrans As Integer, cnt As Integer
  Dim CRBusy As Boolean
  If Exist("GLCREd.DAT") Then CRBusy = GetAttr("GLCREd.DAT") And vbReadOnly
  If Not CRBusy Then
    OpenCJEditFile CJEditFileNum, NumEdTrans, CJType
    If Emode = False Then
      If NumEdTrans > 0 Then
        RecNum = NumEdTrans + 1
      Else
        RecNum = 1
      End If
    Else
      Get CJEditFileNum, RecNum, GLCREd(1)
    End If
    GLCREd(1).DelFlag = 0
    GLCREd(1).LOCKED = False
    GLCREd(1).TRDATE = DateDiff("d", "12/31/1979", txtDate)
    GLCREd(1).Desc = Trim(txtDesc)
    GLCREd(1).LDesc = Trim(txtAdditDesc)
    GLCREd(1).DOCREF = Trim(txtDocRef)
    GLCREd(1).Amt = txtAmount.DoubleValue()
    GLCREd(1).BATCHNUM = Trim(txtBatch)
    GLCREd(1).RECCODE = txtBanks.Text
    For cnt = 1 To 36
      vaSpread1.Row = cnt
      vaSpread1.Col = 1
      If vaSpread1.Text = "" Then
        GLCREd(1).Dist(cnt).DACREC = 0
        GLCREd(1).Dist(cnt).DACN = 0
        GLCREd(1).Dist(cnt).DACNM = 0
        GLCREd(1).Dist(cnt).DAMT = 0
      Else
        GLCREd(1).Dist(cnt).DACREC = vaSpread1.Text
        vaSpread1.Col = 2
        GLCREd(1).Dist(cnt).DACN = Trim(vaSpread1.Text)
        vaSpread1.Col = 3
        GLCREd(1).Dist(cnt).DACNM = Trim(vaSpread1.Text)
        vaSpread1.Col = 4
        GLCREd(1).Dist(cnt).DAMT = vaSpread1.Text
      
      End If
    Next
    Put CJEditFileNum, RecNum, GLCREd(1)
    Close CJEditFileNum
    Call MainLog("CR Saved.")
  Else
    MsgBox "Posting Is In Progress, Editing Not Allowed At This Time.", vbOKOnly, "Request Canceled"
    frmCashReceiptsMenu.Show
    Unload frmCashReceiptEntry
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
  Dim CJREditFile As Integer, CJEditFileNum As Integer, NumEdTrans As Integer
  Dim CurrRec As Integer, NextRec As Integer, cnt As Integer, Last As Integer
  Dim CRBusy As Boolean
  If Exist("GLCREd.DAT") Then CRBusy = GetAttr("GLCREd.DAT") And vbReadOnly
  If Not CRBusy Then
    OpenCJEditFile CJEditFileNum, NumEdTrans, CJType
    'If used list or edit the temprec was selected there, need to transfer to recnum
    OldRec = RecNum
    RecNum = TempRec
    Get CJEditFileNum, RecNum, GLCREd(1)
    If GLCREd(1).LOCKED = False Then
      txtDate = Format(DateAdd("d", (GLCREd(1).TRDATE), "12-31-1979"), "mm/dd/yyyy")
      txtDesc = GLCREd(1).Desc
      txtAdditDesc = GLCREd(1).LDesc
      txtDocRef = GLCREd(1).DOCREF
      txtAmount = GLCREd(1).Amt
      txtBatch = GLCREd(1).BATCHNUM
      txtBanks.Text = GLCREd(1).RECCODE
      Last = UBound(GLCREd(1).Dist)
    
      For cnt = 1 To Last
        If GLCREd(1).Dist(cnt).DACREC <> 0 Then
          vaSpread1.Row = vaSpread1.DataRowCnt + 1
  '***** Until complete testing display the record num here to make sure correct.
  '****** Fixed the Spread to Hide this column so can leave
          vaSpread1.Col = 1
          vaSpread1.Text = GLCREd(1).Dist(cnt).DACREC
          vaSpread1.Col = 2
          vaSpread1.Text = GLCREd(1).Dist(cnt).DACN
          vaSpread1.Col = 3
          vaSpread1.Text = GLCREd(1).Dist(cnt).DACNM
          vaSpread1.Col = 4
          vaSpread1.Text = GLCREd(1).Dist(cnt).DAMT
          txtDist = Round#(txtDist.DoubleValue + GLCREd(1).Dist(cnt).DAMT)
        Else
          Exit For
        End If
      Next
      txtUndist = Round#(txtAmount.DoubleValue - txtDist.DoubleValue)
      GLCREd(1).LOCKED = True
      Put CJEditFileNum, RecNum, GLCREd(1)
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
    MsgBox "Posting Is In Progress, Editing Not Allowed At This Time.", vbOKOnly, "Request Canceled"
    frmCashReceiptsMenu.Show
    Unload frmCashReceiptEntry
  End If
End Function
Private Function Changed()
  Dim CJREditFile As Integer, CJEditFileNum As Integer, NumEdTrans As Integer
  Dim cnt As Integer
  If Emode = False Then
    If Val(txtDesc) <> 0 Then
      Changed = True
      Exit Function
    Else
    If Val(txtAdditDesc) <> 0 Then
      Changed = True
      Exit Function
    Else

      If Val(txtDocRef) <> 0 Then
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
'            If Val(txtBanks.Text) <> 0 Then
'              Changed = True
'              Exit Function
'            Else
              Changed = False
'            End If
          End If
        End If
      End If
      End If
    End If
  Else
    OpenCJEditFile CJEditFileNum, NumEdTrans, CJType
    Get CJEditFileNum, RecNum, GLCREd(1)
    If txtDate <> Format(DateAdd("d", (GLCREd(1).TRDATE), "12-31-1979"), "mm/dd/yyyy") Then
      Changed = True
      Close CJEditFileNum
      Exit Function
    Else
      If txtDesc <> GLCREd(1).Desc Then
        Changed = True
        Close CJEditFileNum
        Exit Function
      Else
      If txtAdditDesc <> GLCREd(1).LDesc Then
        Changed = True
        Close CJEditFileNum
        Exit Function
      Else
        
        If txtDocRef <> GLCREd(1).DOCREF Then
          Changed = True
          Close CJEditFileNum
          Exit Function
        Else
          If txtAmount.DoubleValue() <> GLCREd(1).Amt Then
            Changed = True
            Close CJEditFileNum
            Exit Function
          Else
            If txtBatch <> GLCREd(1).BATCHNUM Then
              Changed = True
              Close CJEditFileNum
              Exit Function
            Else
              If txtBanks.Text <> GLCREd(1).RECCODE Then
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
        If Val(vaSpread1.Text) = GLCREd(1).Dist(cnt).DACREC Then
          If Val(vaSpread1.Text) = 0 Then
            Changed = False
            Exit For
          Else
            vaSpread1.Col = 2
            If vaSpread1.Text = GLCREd(1).Dist(cnt).DACN Then
              vaSpread1.Col = 3
              If vaSpread1.Text = GLCREd(1).Dist(cnt).DACNM Then
                vaSpread1.Col = 4
                If vaSpread1.Text = GLCREd(1).Dist(cnt).DAMT Then
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
  Close CJEditFileNum
End Function

Private Sub NextNew()
  Dim CJREditFile As Integer, CJEditFileNum As Integer, NumEdTrans As Integer
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

'Private Sub fpcboAcctNumNa_GotFocus()
'  If VerifyEntered = False Then
'    MsgBox "The Information In The Top Section Must Be Completed Before The Distributions.", vbOKOnly, "Cash Receipt"
'
'  End If
'End Sub

Private Sub txtAmount_Change()
  txtUndist = Round#(txtAmount - txtDist)
End Sub

Private Sub txtDate_LostFocus()
  If CheckValDate(txtDate) = False Then
    MsgBox "Invalid Date, Please Correct.", vbOKOnly, "Invalid Date"
    txtDate.SetFocus
  End If
End Sub
Private Function Check4Trans()
  Dim CJREditFile As Integer, CJEditFileNum As Integer, NumEdTrans As Integer
  Dim cnt As Integer, Good As Integer
  Good = 0
  If Exist("GLCREd.dat") Then
    OpenCJEditFile CJEditFileNum, NumEdTrans, CJType
    If NumEdTrans > 0 Then
      For cnt = 1 To NumEdTrans
        Get CJEditFileNum, cnt, GLCREd(1)
        If GLCREd(1).DelFlag = 0 Then
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
'****Verify the info in top section has been filled out before
'******allowing distributions
Private Function VerifyEntered()
  If txtDesc <> "" Or txtDesc <> " " Then
    If txtDocRef <> "" Or txtDocRef <> " " Then
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
      txtDocRef.SetFocus
      Exit Function
    End If
  Else
    VerifyEntered = False
    txtDesc.SetFocus
    Exit Function
  End If
End Function


'Private Sub txtCrAmt_GotFocus()
'  If VerifyEntered = False Then
'    MsgBox "The Information In The Top Section Must Be Filled Out Before The Distributions.", vbOKOnly, "Cash Receipt"
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
    fpcboAcctNumNa.Col = 1
    TempAcct = QPTrim(vaSpread1.Text)
    If vaSpread1.Text <> "" Then
      If fpcboAcctNumNa.ListIndex <> -1 Or txtCrAmt <> 0 Then
        If MsgBox("Do You Wish To Abandon Current Distribution?, 'Yes' or 'No' Complete Distribution Entry.", vbYesNo, "Clear??") = vbNo Then
          cmdAddDist.SetFocus
          Exit Sub
        Else
          fpcboAcctNumNa.ListIndex = -1
          txtCrAmt = 0
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
        txtCrAmt = vaSpread1.Text
        txtDist = Round#(txtDist.DoubleValue - txtCrAmt.DoubleValue)
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
  txtDocRef = ""
  txtAmount = 0
  txtBatch = ""
  txtBanks.Text = DefBnk
  fpcboAcctNumNa.ListIndex = -1
  txtCrAmt = 0
'*****Clear data in the spreadsheet
  vaSpread1.ClearRange 1, 1, 4, 36, True
  txtUndist = 0
  txtDist = 0
'  txtDate.SetFocus
End Sub


