VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Begin VB.Form frmGenJournalEntry 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Journal Entry"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11655
   ClipControls    =   0   'False
   ForeColor       =   &H80000007&
   Icon            =   "frmGenJournalEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboAcctNumNa 
      Height          =   405
      Left            =   2310
      TabIndex        =   5
      Top             =   2445
      Width           =   5850
      _Version        =   196608
      _ExtentX        =   10319
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
      SearchMethod    =   1
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
      BorderDropShadowColor=   8421504
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
      ColDesigner     =   "frmGenJournalEntry.frx":08CA
   End
   Begin LpLib.fpCombo txtEType 
      Height          =   405
      Left            =   6930
      TabIndex        =   2
      Top             =   1485
      Width           =   1215
      _Version        =   196608
      _ExtentX        =   2143
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
      ForeColor       =   0
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
      SearchMethod    =   2
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
      BorderDropShadowColor=   8421504
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
      GrayAreaColor   =   14737632
      ListLeftOffset  =   0
      ComboGap        =   -2
      MaxEditLen      =   150
      VirtualPageSize =   0
      VirtualPagesAhead=   0
      ExtendCol       =   0
      ColumnLevels    =   1
      ListGrayAreaColor=   14737632
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
      ColDesigner     =   "frmGenJournalEntry.frx":0D2C
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
      Height          =   540
      Left            =   9264
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7632
      Width           =   1668
   End
   Begin VB.CommandButton cmdDelDist 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F6 Del &Entry From List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   5184
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7632
      Width           =   1668
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F9 &Add Entry To List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   564
      Left            =   8304
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   1596
   End
   Begin EditLib.fpText txtRef 
      Height          =   348
      Left            =   4176
      TabIndex        =   1
      Top             =   1488
      Width           =   1284
      _Version        =   196608
      _ExtentX        =   2265
      _ExtentY        =   614
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
      ThreeDTextShadowColor=   12632256
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
      BorderDropShadowColor=   8421504
      BorderDropShadowWidth=   3
      ButtonColor     =   12632256
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpCurrency txtCredits 
      Height          =   372
      Left            =   9408
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   7056
      Width           =   1332
      _Version        =   196608
      _ExtentX        =   2350
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
      MaxValue        =   "9999999999.99"
      MinValue        =   "-9999999999.99"
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
   Begin EditLib.fpCurrency txtDebits 
      Height          =   372
      Left            =   7968
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   7056
      Width           =   1428
      _Version        =   196608
      _ExtentX        =   2519
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
      MaxValue        =   "99999999999.99"
      MinValue        =   "-9999999999.99"
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
      Height          =   348
      Left            =   5496
      TabIndex        =   4
      Top             =   1992
      Width           =   2676
      _Version        =   196608
      _ExtentX        =   4720
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
      Left            =   1920
      TabIndex        =   3
      Top             =   1968
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
      MaxValue        =   "999999999.99"
      MinValue        =   "-999999999.99"
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
      IncDec          =   1
      BorderGrayAreaColor=   14737632
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   14737632
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
      Left            =   1632
      TabIndex        =   0
      Top             =   1488
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
      ButtonStyle     =   2
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   16777215
      ThreeDTextShadowColor=   8421504
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
      BorderGrayAreaColor=   12632256
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   12632256
      Appearance      =   2
      BorderDropShadow=   0
      BorderDropShadowColor=   8421504
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
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F3 &Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   564
      Left            =   9960
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2280
      Width           =   876
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   20
      Top             =   8400
      Width           =   11652
      _ExtentX        =   20558
      _ExtentY        =   423
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
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00D0D0D0&
      Caption         =   "F10 &Save GJ Entries"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   7224
      MaskColor       =   &H00D0D0D0&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7632
      Width           =   1668
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   3585
      Left            =   840
      TabIndex        =   7
      Top             =   3240
      Width           =   10050
      _Version        =   196613
      _ExtentX        =   17727
      _ExtentY        =   6324
      _StockProps     =   64
      BackColorStyle  =   1
      ButtonDrawMode  =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridColor       =   8421504
      MaxCols         =   10
      OperationMode   =   3
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   14737632
      ShadowDark      =   8421504
      ShadowText      =   0
      SpreadDesigner  =   "frmGenJournalEntry.frx":10BA
      VisibleCols     =   7
      VisibleRows     =   12
   End
   Begin EditLib.fpText fptxtDesc2 
      Height          =   300
      Left            =   8280
      TabIndex        =   12
      Top             =   1824
      Width           =   2556
      _Version        =   196608
      _ExtentX        =   4508
      _ExtentY        =   529
      Enabled         =   -1  'True
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
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   765
      Left            =   240
      TabIndex        =   26
      Top             =   240
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label Label3b 
      BackStyle       =   0  'Transparent
      Caption         =   "(Opt.)Additional Desc."
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
      Left            =   8280
      TabIndex        =   25
      Top             =   1536
      Width           =   2556
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      FillColor       =   &H00D0D0D0&
      Height          =   3780
      Left            =   765
      Top             =   3120
      Width           =   10170
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Double-Click Row To Edit Information."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   300
      Left            =   624
      TabIndex        =   24
      Top             =   7776
      Width           =   3900
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      Height          =   516
      Left            =   792
      Top             =   6960
      Width           =   10140
   End
   Begin VB.Label lblTotals 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Totals"
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
      Index           =   4
      Left            =   6384
      TabIndex        =   23
      Top             =   7104
      Width           =   852
   End
   Begin VB.Label Label3b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ref"
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
      Left            =   3456
      TabIndex        =   22
      Top             =   1536
      Width           =   564
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
      Left            =   864
      TabIndex        =   21
      Top             =   1536
      Width           =   660
   End
   Begin VB.Label Label3b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Left            =   3912
      TabIndex        =   19
      Top             =   1992
      Width           =   1428
   End
   Begin VB.Label Label4b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Entry Type"
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
      Left            =   5592
      TabIndex        =   18
      Top             =   1536
      Width           =   1284
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
      Left            =   768
      TabIndex        =   17
      Top             =   2016
      Width           =   1068
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000016&
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      Height          =   1644
      Index           =   0
      Left            =   768
      Top             =   1392
      Width           =   10140
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "GL Account"
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
      Left            =   768
      TabIndex        =   16
      Top             =   2496
      Width           =   1428
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "General Journal Entry/Edit"
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
      Left            =   3840
      TabIndex        =   15
      Top             =   576
      Width           =   3972
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000005&
      Height          =   852
      Left            =   2760
      Top             =   336
      Width           =   6132
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00D0D0D0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D0D0D0&
      FillColor       =   &H00D0D0D0&
      Height          =   972
      Left            =   2760
      Top             =   216
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
Attribute VB_Name = "frmGenJournalEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GLSetup As GLSetupRecType
Dim GLAcct As GLAcctRecType
Dim GJEdit As TrEditRecType
Dim GLTrans As GLTransRecType
Dim Over As clsTextBoxOverRider
Dim LPDate As Integer, HPDate As Integer
Dim GLUserName As String, GLFundLen As Integer, GLAcctLen As Integer, GLDetLen As Integer
Dim GJEditFileNum As Integer, NumEdTrans As Integer
Dim Emode As Boolean, RecNum As Integer
Private Temp_Class As Resize_Class
Dim GLAcctidx As GLAcctIndexType
Dim Verify As String, Change As Boolean
'This is to fix spreadsheet for various resolutions
Public Function Fixspread()
'    Select Case screenW
'      Case 1280
'      If Screen.TwipsPerPixelX <> 12 Then
'        coladj = 10.7
'        vaSpread1.RowHeight(-1) = 22.2
'      Else
'        coladj = 6.8
'        vaSpread1.RowHeight(-1) = 18
'      End If
'      Case 1152
'      If Screen.TwipsPerPixelX <> 12 Then
'        coladj = 8.8
'        vaSpread1.RowHeight(-1) = 19.2
'      Else
'        coladj = 5.2
'        vaSpread1.RowHeight(-1) = 16
'      End If
'      Case 1024
'      If Screen.TwipsPerPixelX <> 12 Then
'        coladj = 6.75
'        vaSpread1.RowHeight(0) = 18
'        vaSpread1.RowHeight(-1) = 18
'      Else
'        coladj = 3.7
'      End If
'      Case 800
'        coladj = 3.4
'        'vaSpread1.Font.Size = 8
'        vaSpread1.RowHeight(-1) = 13
'      Case 1400
'        coladj = 15.7
'      Case Else
'        'don't worry be happpy
'    End Select
    vaSpread1.Font.Size = 11
    'vaSpread1.RowHeight(0) = 18
    'vaSpread1.RowHeight(-1) = 18

    'vaSpread1.ColWidth(-1) = vaSpread1.ColWidth(-1) + coladj
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      If VerifyEntered = True Then
        If MsgBox("Abandon Entry and Close.", vbYesNo, "Abandon?") = vbNo Then
          Cancel = True
        End If
      Else
      If Change = True Then
        If MsgBox("Close Program Without Saving?", vbYesNo, "GJ Entry") = vbNo Then
          Cancel = True
        End If
      End If
      End If
      Call MainLog("Close via GenJournalEntry.")
      Close GJEditFileNum
      KillFileD "GJEdit.opn"
      ClearInUse PWcnt
    End If
 End If
End Sub

Private Sub Form_Load()
  Dim cnt As Integer, CntB As Integer
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  GetPostDates LPDate, HPDate
  GetAcctStruct GLUserName, GLFundLen, GLAcctLen, GLDetLen
  StatusBar1.Panels.Item(1).Text = GLUserName
  FillAcctNumName fpcboAcctNumNa
  OpenGJEditFile GJEditFileNum, NumEdTrans
'  If GJEditFileNum = -1 Then
'    Exit Sub
'  End If
  Me.HelpContextID = hlpGJEdit
  Fixspread
  Label4.Caption = Screen.Width & "   " & Screen.TwipsPerPixelX
  If GJEditFileNum < 0 Then
    frmGenJournalMenu.Show
    Unload frmGenJournalEntry
    Exit Sub
  End If
  Change = False
  If NumEdTrans > 0 Then
    For cnt = 1 To NumEdTrans
      Get GJEditFileNum, cnt, GJEdit
      If GJEdit.Deleted <> 0 Then
        CntB = CntB + 1
      Else
        Emode = True
        Exit For
      End If
    Next
  Else
    Emode = False
  End If
  If CntB = NumEdTrans Then
    Close GJEditFileNum
    KillFile "GJEdit.dat"
    OpenGJEditFile GJEditFileNum, NumEdTrans
    If GJEditFileNum < 0 Then
      frmGenJournalMenu.Show
      Unload frmGenJournalEntry
      Exit Sub
    End If
  End If
  If Emode = True Then
    Rec2Form
  Else
    RecNum = 1
    txtDebits = 0
    txtCredits = 0
  End If
  txtDate.Text = Format(Now, "mm/dd/yyyy")
  txtDesc = ""
  fptxtDesc2 = ""
  txtRef = ""
  txtAmount = 0
  txtEType.AddItem "Debit"
  txtEType.AddItem "Credit"
  txtEType.ListIndex = -1
  fpcboAcctNumNa.ListIndex = -1
  
  '***** spreadsheet do Not have to set blank fields on load ..
End Sub

Public Function Rec2Form()
  Dim AcctRec As Integer, Last As Integer
  Dim CurrRec As Integer, NextRec As Integer, cnt As Integer, BadCnt&, zzz&, GotIt As Boolean
  ReDim AcctList(1 To 1) As String
  'OpenGJEditFile GJEditFileNum, NumEdTrans
  For RecNum = 1 To NumEdTrans
    Get GJEditFileNum, RecNum, GJEdit
    If GJEdit.Deleted = 0 Then
      AcctRec = AcctFind(GJEdit.AcctNum)
      If AcctRec > 0 Then
        vaSpread1.Row = vaSpread1.DataRowCnt + 1
        vaSpread1.Col = 1
        vaSpread1.Text = Format(DateAdd("d", (GJEdit.TRDATE), "12-31-1979"), "mm/dd/yyyy")
        vaSpread1.Col = 2
        vaSpread1.Text = AcctRec
        vaSpread1.Col = 3
        vaSpread1.Text = GJEdit.AcctNum
        vaSpread1.Col = 4
        vaSpread1.Text = GJEdit.AcctName
        vaSpread1.Col = 5
        vaSpread1.Text = GJEdit.Ref
        vaSpread1.Col = 6
        vaSpread1.Text = GJEdit.Desc
        vaSpread1.Col = 7
        vaSpread1.Text = GJEdit.EType
        vaSpread1.Col = 8
        vaSpread1.Text = GJEdit.DrAmt
        vaSpread1.Col = 9
        vaSpread1.Text = GJEdit.CrAmt
        vaSpread1.Col = 10
        vaSpread1.Text = GJEdit.LDesc
        txtDebits = Round#(txtDebits.DoubleValue + GJEdit.DrAmt)
        txtCredits = Round#(txtCredits.DoubleValue + GJEdit.CrAmt)
      Else
'        GotIt = False
'        If BadCnt& = 0 Then
'          BadCnt& = BadCnt& + 1
'          AcctList(1) = GJEdit.AcctNum
'        Else
'          For zzz& = 1 To BadCnt&
'            If AcctList(zzz&) = GJEdit.AcctNum Then
'              GotIt = True
'              Exit For
'            End If
'          Next
'          If Not GotIt Then
'            BadCnt& = BadCnt& + 1
'            ReDim Preserve AcctList(1 To BadCnt&) As String
'            AcctList(BadCnt) = GJEdit.AcctNum
'          End If
'        End If
        MsgBox "An Invalid Account Record Was Encountered And Will Not Be Loaded.", vbOKOnly, "Invalid Account"
      End If
    End If
  Next
 
'  Open "badacct.txt" For Output As #50
'  For zzz& = 1 To BadCnt
'    Print #50, AcctList(zzz&)
'  Next
'  Close 50
  Emode = True
End Function



Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    ''Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
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
      cmdUpdate.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        txtDesc.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub
Private Sub fpcboAcctNumNa_LostFocus()
  fpcboAcctNumNa.Action = ActionClearSearchBuffer
End Sub

Private Sub fptxtDesc2_ChangeMode(EditMode As Integer)
  EditMode = True
End Sub

Private Sub fptxtDesc2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
   cmdUpdate.SetFocus
  End If
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
End Sub

Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    txtDesc.SetFocus
  End If
End Sub

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    txtRef.SetFocus
  End If
End Sub

Private Sub txtDate_LostFocus()
  If CheckValDate(txtDate) = False Then
    MsgBox "Invalid Date, Please Retry.", vbOKOnly, "GJ Entry"
    txtDate.SetFocus
  End If
End Sub
Private Sub cmdDelDist_Click()
  If vaSpread1.ActiveRow > 0 Then
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 4
    If vaSpread1.Text <> "" Then
      If MsgBox("You Wish to Delete Selected Entry?", vbYesNo, "Delete Entry") = vbYes Then
        Change = True
        vaSpread1.Col = 8
        txtDebits = Round#(txtDebits.DoubleValue - vaSpread1.Text)
        vaSpread1.Col = 9
        txtCredits = Round#(txtCredits.DoubleValue - vaSpread1.Text)
        vaSpread1.DeleteRows vaSpread1.Row, 1
        fpcboAcctNumNa.SetFocus
      End If
    End If
  End If

End Sub

Private Sub cmdDelete_Click()
  If VerifyEntered = True Then
   Change = True
   If MsgBox("Are you sure you wish to delete this entry?", vbYesNo, "Delete GJEntry") = vbYes Then
      Call NextNew
      txtDesc = ""
      txtRef = ""
      fptxtDesc2 = ""
    Else
      txtDate.SetFocus
      Exit Sub
    End If
  Else
    MsgBox "You Must First Select An Entry To Delete.", vbOKOnly, "Deletion Denied"
  End If
End Sub

Private Sub cmdUpdate_Click()
  Dim AcctRec As Integer, EType As String
  If VerifyEntered = True Then
    If fpcboAcctNumNa.Text <> "" And txtAmount.DoubleValue <> 0 Then
      vaSpread1.Row = vaSpread1.DataRowCnt + 1
      vaSpread1.Col = 1
      vaSpread1.Text = txtDate
      vaSpread1.Col = 2
      fpcboAcctNumNa.Col = 0
      AcctRec = Val(fpcboAcctNumNa.ColText)
      vaSpread1.Text = fpcboAcctNumNa.ColText
      vaSpread1.Col = 3
      fpcboAcctNumNa.Col = 1
      vaSpread1.Text = fpcboAcctNumNa.ColText
      vaSpread1.Col = 4
      fpcboAcctNumNa.Col = 2
      vaSpread1.Text = fpcboAcctNumNa.ColText
      vaSpread1.Col = 5
      vaSpread1.Text = txtRef
      vaSpread1.Col = 6
      vaSpread1.Text = txtDesc
      vaSpread1.Col = 7
      vaSpread1.Text = Mid$(txtEType.Text, 1, 1)
      EType = vaSpread1.Text
      If EType = "D" Then
        vaSpread1.Col = 8
        vaSpread1.Text = txtAmount.DoubleValue
        txtDebits = Round#(txtDebits.DoubleValue + txtAmount.DoubleValue)
        vaSpread1.Col = 9
        vaSpread1.Text = 0
      Else
        vaSpread1.Col = 8
        vaSpread1.Text = 0
        vaSpread1.Col = 9
        vaSpread1.Text = txtAmount.DoubleValue
        txtCredits = Round#(txtCredits.DoubleValue + txtAmount.DoubleValue)
      End If
      vaSpread1.Col = 10
      vaSpread1.Text = fptxtDesc2
      Change = True
      NextNew
    End If
  Else
    MsgBox Verify$, vbOKOnly, "General Journal Entry Error"
  End If
End Sub
Public Function GetType(AcctRec As Integer, EType As String)
  Dim AcctFile As Integer, NumAccts As Integer
  OpenAcctFile AcctFile
  NumAccts = LOF(AcctFile) / Len(GLAcct)
    Get AcctFile, AcctRec, GLAcct
    If GLAcct.Deleted = 0 Then
      If EType = "I" Then
        If GLAcct.Typ = "E" Then
          GetType = True
        Else
          GetType = False
        End If
      End If
      If EType = "D" Then
        If GLAcct.Typ = "E" Then
          GetType = False
        Else
          GetType = True
        End If
      End If
       
    End If
  Close AcctFile
End Function
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
'    Case vbKeyDown, vbKeyReturn:
'      SendKeys "{Tab}"
'      KeyCode = 0
'    Case vbKeyUp:
'      SendKeys "+{Tab}"
'      KeyCode = 0
    Case vbKeyEscape:
      cmdExit_Click
      KeyCode = 0
    Case vbKeyF3:
      cmdDelete_Click
      KeyCode = 0
    Case vbKeyF6:
      KeyCode = 0
      cmdDelDist_Click
    Case vbKeyF10:
      cmdSave_Click
      KeyCode = 0
    Case vbKeyF9:
      cmdUpdate_Click
      KeyCode = 0
    Case Else:
  End Select
End Sub
Private Sub cmdExit_Click()
  If MsgBox("Are You Sure You Wish To Exit?", vbYesNo, "Journal Entry") = vbYes Then

  If VerifyEntered = True Then
    If MsgBox("          Select OK To Abandon Entry." & Chr(13) & Chr(13) & "CANCEL to Remain on Current Screen.", vbOKCancel, "Abandon Current Entry?") = vbCancel Then
      txtDate.SetFocus
      Exit Sub
    Else
      cmdDelete_Click
      Exit Sub
    End If
  End If
  If Change = False Then
    frmGenJournalMenu.Show
    Unload frmGenJournalEntry
    Close GJEditFileNum
  Else
    If MsgBox("Save General Journal Entry List Before Exiting, Yes or No?", vbYesNo, "GJ Entry") = vbYes Then
       SaveGJList
    End If
    frmGenJournalMenu.Show
    Unload frmGenJournalEntry
    Close GJEditFileNum
  End If
  KillFileD "GJEdit.opn"
  End If
End Sub
Private Sub cmdSave_Click()
  If VerifyEntered = True Then
    If MsgBox("         Select OK to Abandon Entry." & Chr(13) & Chr(13) & "CANCEL to Remain on Current Screen.", vbOKCancel, "Abandon Current Entry?") = vbCancel Then
      txtDate.SetFocus
      Exit Sub
    Else
      cmdDelete_Click
      Exit Sub
    End If
  End If
  
'    If vaSpread1.DataRowCnt = 0 Then
'      If MsgBox("                  The Budget List Is Empty." & Chr(13) & Chr(13) & "Select YES And The Edit File Will Be Cleared" & Chr(13) & Chr(13) & "              Or NO To Abandon Changes.", vbYesNo, "Budget Entry") = vbNo Then
'        frmBudgetMaintMenu.Show
'        Unload frmBudgetEntEdit
'        Close BgtEditFileNum
'        Exit Sub
'      End If
'    End If
    SaveGJList
    MsgBox "Your General Journal Entries Have Been Saved.", vbOKOnly, "GJ Saved"
  
  frmGenJournalMenu.Show
  Unload frmGenJournalEntry
  'Close is done in SaveBudgetList procedure
  KillFileD "GJEdit.opn"
End Sub

Private Sub SaveGJList()
  Dim cnt As Integer
 ' If Exist("BGTED.DAT") Then
    Close GJEditFileNum
    KillFile "GJEdit.DAT"
  'End If
  OpenGJEditFile GJEditFileNum, NumEdTrans
  If GJEditFileNum < 0 Then
    frmGenJournalMenu.Show
    Unload frmGenJournalEntry
    Exit Sub
  End If
  If vaSpread1.DataRowCnt > 0 Then
    For cnt = 1 To vaSpread1.DataRowCnt
      vaSpread1.Row = cnt
      GJEdit.Deleted = 0
      vaSpread1.Col = 1
      GJEdit.TRDATE = DateDiff("d", "12/31/1979", vaSpread1.Text)
      vaSpread1.Col = 2
      GJEdit.AcctRec = QPTrim(vaSpread1.Text)
      vaSpread1.Col = 3
      GJEdit.AcctNum = QPTrim(vaSpread1.Text)
      vaSpread1.Col = 4
      GJEdit.AcctName = QPTrim(vaSpread1.Text)
      vaSpread1.Col = 5
      GJEdit.Ref = QPTrim(vaSpread1.Text)
      vaSpread1.Col = 6
      GJEdit.Desc = QPTrim(vaSpread1.Text)
      vaSpread1.Col = 7
      GJEdit.EType = QPTrim(vaSpread1.Text)
      vaSpread1.Col = 8
      GJEdit.DrAmt = QPTrim(vaSpread1.Text)
      vaSpread1.Col = 9
      GJEdit.CrAmt = QPTrim(vaSpread1.Text)
      vaSpread1.Col = 10
      GJEdit.LDesc = QPTrim(vaSpread1.Text)
      RecNum = cnt
      Put GJEditFileNum, RecNum, GJEdit
    Next
  Close GJEditFileNum
  Call MainLog("GJ Saved.")
  Else
    Close GJEditFileNum
    
  End If
End Sub


Private Sub NextNew()
'  Dim BgtEditFileNum As Integer, NumEdTrans As Integer
'  OpenBgtEditFile BgtEditFileNum, NumEdTrans
'  Close BgtEditFileNum
'   If NumEdTrans > 0 Then
'     RecNum = NumEdTrans + 1
'   Else
'     RecNum = 1
'   End If

   Emode = False
   ClearFields
   txtDate.SetFocus
End Sub

Private Function VerifyEntered()
  Dim TempDate As Integer
  If CheckValDate(txtDate) = True Then

  TempDate = DateDiff("d", "12/31/1979", txtDate)
    If txtEType.Text <> "" Then
      If txtAmount <> 0 Then
        If txtDesc <> "" Then
          If fpcboAcctNumNa.Text <> "" Then
            VerifyEntered = True
            'Also compare date with First Fiscal Year range
            If (TempDate < LPDate) Or (TempDate > HPDate) Then
              Verify$ = "This Date Is Not Within Valid Posting Date Range. Please Correct."
              VerifyEntered = False
              txtDate.SetFocus
              Exit Function
            End If
          Else
            VerifyEntered = False
            Verify$ = "An Account Must Be Selected Before Adding To List."
            fpcboAcctNumNa.SetFocus
            Exit Function
          End If
        Else
          VerifyEntered = False
          Verify$ = "Description May Not Be Left Blank."
          txtDesc.SetFocus
          Exit Function
        End If
      Else
        VerifyEntered = False
        Verify$ = "Amount May Not Be Zero."
        txtAmount.SetFocus
        Exit Function
      End If
    Else
      VerifyEntered = False
      Verify$ = "You Must Select Either Debit or Credit."
      txtEType.SetFocus
      Exit Function
    End If
  Else
    VerifyEntered = False
    Verify$ = "Invalid Date."
    txtDate.SetFocus
    Exit Function
  End If
  
End Function




Private Sub txtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    fpcboAcctNumNa.SetFocus
  End If
End Sub

Private Sub txtEType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    txtEType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    txtEType.ListIndex = -1
    txtEType.Action = ActionClearSearchBuffer
  End If
  If txtEType.ListDown <> True Then
    If KeyCode = vbKeyDown Or KeyCode = vbKeyReturn Then
      txtAmount.SetFocus
      KeyCode = 0
    Else
      If KeyCode = vbKeyUp Then
        txtRef.SetFocus
        KeyCode = 0
      End If
    End If
  End If

End Sub

Private Sub txtRef_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    txtEType.SetFocus
  End If
End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
  Dim TempAcct As String
  Dim TempCol As Long, TempRow As Long
  TempRow = Row
  TempCol = Col
  vaSpread1.Col = 1
  vaSpread1.Row = Row
  If TempRow > 0 And vaSpread1.Text <> "" Then
    If fpcboAcctNumNa.Text <> "" Then
'      frmMsgBlank.Label1 = "The Entry in Edit Mode Has Not Been Saved." & Chr(13) & "Select 'Ok' to Abandon And Edit Another Entry." & Chr(13) & "or 'Cancel' to Complete Current Entry ?"
'      frmMsgBlank.cmd1.Caption = "F10 &Ok"
'      frmMsgBlank.cmd2.Caption = "Esc &Cancel"
'      frmMsgBlank.Caption = "GJ Entry"
'      frmMsgBlank.Show , Me
      If MsgBox("Overwrite and Delete Current Entry, 'OK' or 'Cancel'", vbOKCancel, "GJ Entry") = vbCancel Then
        Exit Sub
      Else
        cmdDelete_Click
      End If
    End If
    vaSpread1.Row = TempRow
    vaSpread1.Col = 3
    TempAcct = QPTrim(vaSpread1.Text)
    If vaSpread1.Text <> "" Then
      fpcboAcctNumNa.SearchText = QPStrip(TempAcct)
      fpcboAcctNumNa.Action = 0
      If fpcboAcctNumNa.SearchIndex <> -1 Then
        fpcboAcctNumNa.ListIndex = fpcboAcctNumNa.SearchIndex
      End If
      vaSpread1.Col = 2
      fpcboAcctNumNa.Col = 0
      fpcboAcctNumNa.ColText = vaSpread1.Text
      vaSpread1.Col = 3
      fpcboAcctNumNa.Col = 1
      fpcboAcctNumNa.ColText = vaSpread1.Text
      vaSpread1.Col = 4
      fpcboAcctNumNa.Col = 2
      fpcboAcctNumNa.ColText = vaSpread1.Text
      vaSpread1.Col = 1
      txtDate = vaSpread1.Text
      vaSpread1.Col = 5
      txtRef = vaSpread1.Text
      vaSpread1.Col = 6
      txtDesc = vaSpread1.Text
      vaSpread1.Col = 7
      If vaSpread1.Text = "D" Then
        txtEType.Text = "Debit"
      Else
        txtEType.Text = "Credit"
      End If
      vaSpread1.Col = 8
      If vaSpread1.Value > 0 Then
        txtAmount = vaSpread1.Text
        txtDebits = Round#(txtDebits.DoubleValue - txtAmount.DoubleValue)
      Else
        vaSpread1.Col = 9
        txtAmount = vaSpread1.Text
        txtCredits = Round#(txtCredits.DoubleValue - txtAmount.DoubleValue)
      End If
      vaSpread1.Col = 10
      fptxtDesc2 = vaSpread1.Text
      'vaSpread1.ClearRange TempCol, TempRow, 4, TempRow, True
      vaSpread1.DeleteRows TempRow, 1
      fpcboAcctNumNa.SetFocus
    End If
  End If
End Sub
Public Sub ClearFields()
  'txtDesc = ""
  'txtRef = ""
  txtAmount = 0
  txtEType.Text = ""
  fpcboAcctNumNa.ListIndex = -1
  'vaSpread1.ClearRange 1, 1, 4, 36, True
End Sub



