VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmTaxPayEntry 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Payment Entry"
   ClientHeight    =   8856
   ClientLeft      =   36
   ClientTop       =   540
   ClientWidth     =   12192
   ClipControls    =   0   'False
   Icon            =   "frmTaxPayEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8856
   ScaleWidth      =   12192
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcboTenderType 
      Height          =   360
      Left            =   2880
      TabIndex        =   2
      Top             =   3864
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3831
      _ExtentY        =   635
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
      Columns         =   2
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
      ThreeDInsideStyle=   0
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
      Appearance      =   0
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
      ColDesigner     =   "frmTaxPayEntry.frx":08CA
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "F6 Chec&k"
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
      Left            =   2544
      TabIndex        =   50
      Top             =   7512
      Width           =   1332
   End
   Begin VB.CommandButton cmdCash 
      Caption         =   "F5 &Cash"
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
      Left            =   620
      TabIndex        =   49
      Top             =   7512
      Width           =   1332
   End
   Begin EditLib.fpText fptxtBill 
      Height          =   432
      Left            =   1560
      TabIndex        =   48
      Top             =   2580
      Width           =   2244
      _Version        =   196608
      _ExtentX        =   3958
      _ExtentY        =   762
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
      ThreeDInsideStyle=   0
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
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpLongInteger fplngAcct 
      Height          =   372
      Left            =   3600
      TabIndex        =   0
      Top             =   1320
      Width           =   1872
      _Version        =   196608
      _ExtentX        =   3302
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
      ControlType     =   0
      Text            =   "0"
      MaxValue        =   "2147483647"
      MinValue        =   "-2147483648"
      NegFormat       =   1
      NegToggle       =   0   'False
      Separator       =   ""
      UseSeparator    =   0   'False
      IncInt          =   1
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
   Begin VB.CommandButton cmdBills 
      Caption         =   "F8 &Bills"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Left            =   3792
      TabIndex        =   1
      Top             =   2580
      Width           =   1692
   End
   Begin VB.CommandButton cmdDist 
      Caption         =   "F9 &Dist"
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
      Left            =   6392
      TabIndex        =   13
      Top             =   7512
      Width           =   1332
   End
   Begin VB.CommandButton cmdLookup 
      Caption         =   "F7 &Lookup"
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
      Left            =   4468
      TabIndex        =   12
      Top             =   7512
      Width           =   1332
   End
   Begin VB.CommandButton cmdExit 
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
      Left            =   10241
      TabIndex        =   11
      Top             =   7512
      Width           =   1332
   End
   Begin VB.CommandButton cmdSave 
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
      Left            =   8316
      TabIndex        =   10
      Top             =   7512
      Width           =   1332
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   14
      Top             =   8496
      Width           =   12192
      _ExtentX        =   21505
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7133
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "5:02 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   7133
            TextSave        =   "11/18/2003"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin EditLib.fpDateTime txtPaymentDate 
      Height          =   372
      Left            =   9960
      TabIndex        =   21
      Top             =   720
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
   Begin EditLib.fpCurrency fpChkAmt 
      Height          =   384
      Left            =   2880
      TabIndex        =   4
      Top             =   4608
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3831
      _ExtentY        =   677
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
      ThreeDInsideStyle=   0
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
   Begin EditLib.fpText fptxtName 
      Height          =   396
      Left            =   1560
      TabIndex        =   23
      Top             =   1800
      Width           =   3924
      _Version        =   196608
      _ExtentX        =   6921
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
      ThreeDInsideStyle=   0
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
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtDesc 
      Height          =   372
      Left            =   1740
      TabIndex        =   6
      Top             =   6480
      Width           =   3720
      _Version        =   196608
      _ExtentX        =   6562
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
   Begin EditLib.fpText fptxtAddress 
      Height          =   396
      Left            =   1560
      TabIndex        =   28
      Top             =   2196
      Width           =   3924
      _Version        =   196608
      _ExtentX        =   6921
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
      ThreeDInsideStyle=   0
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
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpCurrency fpDiscAmt 
      Height          =   384
      Left            =   2880
      TabIndex        =   5
      Top             =   4992
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3831
      _ExtentY        =   677
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
      ThreeDInsideStyle=   0
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
   Begin EditLib.fpCurrency fpCashAmt 
      Height          =   384
      Left            =   2880
      TabIndex        =   3
      Top             =   4224
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3831
      _ExtentY        =   677
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
      ThreeDInsideStyle=   0
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
   Begin EditLib.fpCurrency fpChangeDue 
      Height          =   384
      Left            =   2880
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   5880
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3831
      _ExtentY        =   677
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ThreeDInsideStyle=   0
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
      ControlType     =   2
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
   Begin EditLib.fpCurrency fpTotReceived 
      Height          =   384
      Left            =   2880
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   5520
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3831
      _ExtentY        =   677
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ThreeDInsideStyle=   0
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
      ControlType     =   2
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
   Begin EditLib.fpCurrency fpTaxOwed 
      Height          =   384
      Left            =   7320
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   1920
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3831
      _ExtentY        =   677
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ThreeDInsideStyle=   0
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
      ControlType     =   2
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
   Begin EditLib.fpCurrency fpIntOwed 
      Height          =   384
      Left            =   7320
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   2304
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3831
      _ExtentY        =   677
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ThreeDInsideStyle=   0
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
      ControlType     =   2
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
   Begin EditLib.fpCurrency fpAdColOwed 
      Height          =   384
      Left            =   7320
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   2688
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3831
      _ExtentY        =   677
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ThreeDInsideStyle=   0
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
      ControlType     =   2
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
   Begin EditLib.fpCurrency fpTaxPaid 
      Height          =   384
      Left            =   9600
      TabIndex        =   7
      Top             =   1920
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3831
      _ExtentY        =   677
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
      ThreeDInsideStyle=   0
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
   Begin EditLib.fpCurrency fpIntPaid 
      Height          =   384
      Left            =   9600
      TabIndex        =   8
      Top             =   2304
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3831
      _ExtentY        =   677
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
      ThreeDInsideStyle=   0
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
   Begin EditLib.fpCurrency fpAdColPaid 
      Height          =   384
      Left            =   9600
      TabIndex        =   9
      Top             =   2688
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3831
      _ExtentY        =   677
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
      ThreeDInsideStyle=   0
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
   Begin EditLib.fpCurrency fpTotOwed 
      Height          =   384
      Left            =   7320
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   6480
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3831
      _ExtentY        =   677
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ThreeDInsideStyle=   0
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
      ControlType     =   2
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
   Begin EditLib.fpCurrency fpTotPaid 
      Height          =   384
      Left            =   9600
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   6480
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3831
      _ExtentY        =   677
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ThreeDInsideStyle=   0
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
      ControlType     =   2
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
   Begin EditLib.fpCurrency fptxtAmtOwed 
      Height          =   384
      Left            =   2880
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   3480
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3831
      _ExtentY        =   677
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   0   'False
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
      ThreeDInsideStyle=   0
      ThreeDInsideHighlightColor=   -2147483633
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
      ThreeDTextHighlightColor=   -2147483633
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
      OnFocusNoSelect =   -1  'True
      OnFocusPosition =   1
      ControlType     =   2
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   2
      CurrencyNegFormat=   0
      CurrencyPlacement=   0
      CurrencySymbol  =   "$"
      DecimalPoint    =   "."
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
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin VB.Label lblSource 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2160
      TabIndex        =   52
      Top             =   720
      Width           =   1032
   End
   Begin VB.Label lblOperator 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6300
      TabIndex        =   51
      Top             =   720
      Width           =   732
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Totals:"
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
      Left            =   6000
      TabIndex        =   45
      Top             =   6540
      Width           =   1212
   End
   Begin VB.Line Line5 
      X1              =   240
      X2              =   12060
      Y1              =   6900
      Y2              =   6900
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Source:"
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
      Height          =   312
      Left            =   300
      TabIndex        =   44
      Top             =   780
      Width           =   3192
   End
   Begin VB.Line Line4 
      X1              =   9540
      X2              =   9540
      Y1              =   1560
      Y2              =   6840
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "Amount Paid"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   9600
      TabIndex        =   40
      Top             =   1380
      Width           =   2172
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "Amount Owed"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   312
      Left            =   7320
      TabIndex        =   39
      Top             =   1380
      Width           =   2172
   End
   Begin VB.Line Line3 
      X1              =   5700
      X2              =   5700
      Y1              =   1440
      Y2              =   6900
   End
   Begin VB.Line Line2 
      X1              =   2640
      X2              =   5340
      Y1              =   5460
      Y2              =   5460
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Adv/Collect:"
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
      Left            =   5760
      TabIndex        =   38
      Top             =   2760
      Width           =   1452
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Interest:"
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
      Left            =   5940
      TabIndex        =   37
      Top             =   2340
      Width           =   1272
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tax:"
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
      Left            =   6300
      TabIndex        =   36
      Top             =   1980
      Width           =   912
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
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
      Left            =   120
      TabIndex        =   35
      Top             =   6540
      Width           =   1512
   End
   Begin VB.Label Lbl12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Discount Amt:"
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
      Left            =   816
      TabIndex        =   34
      Top             =   5040
      Width           =   1872
   End
   Begin VB.Label Lbl11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Check Amt Paid:"
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
      Left            =   816
      TabIndex        =   33
      Top             =   4620
      Width           =   1872
   End
   Begin VB.Label lblchange 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Change Due:"
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
      Left            =   816
      TabIndex        =   32
      Top             =   5940
      Width           =   1872
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
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
      Left            =   24
      TabIndex        =   29
      Top             =   2250
      Width           =   1368
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bill:"
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
      Left            =   468
      TabIndex        =   27
      Top             =   2640
      Width           =   924
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tender Type:"
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
      Left            =   1104
      TabIndex        =   26
      Top             =   3900
      Width           =   1584
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cash/Other Amt Paid:"
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
      Left            =   180
      TabIndex        =   25
      Top             =   4260
      Width           =   2508
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Tax  Payment"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   180
      TabIndex        =   24
      Top             =   3120
      Width           =   5520
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Received:"
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
      Height          =   336
      Left            =   396
      TabIndex        =   22
      Top             =   5580
      Width           =   2292
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Date:"
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
      Left            =   7800
      TabIndex        =   20
      Top             =   780
      Width           =   2040
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Owed:"
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
      Height          =   336
      Index           =   0
      Left            =   960
      TabIndex        =   19
      Top             =   3540
      Width           =   1728
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
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
      Left            =   420
      TabIndex        =   18
      Top             =   1860
      Width           =   972
   End
   Begin VB.Label Label2b 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Account Number:"
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
      Left            =   120
      TabIndex        =   17
      Top             =   1380
      Width           =   3288
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      X1              =   120
      X2              =   11760
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operator Number:"
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
      Height          =   336
      Left            =   4020
      TabIndex        =   16
      Top             =   780
      Width           =   2208
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   456
      Left            =   2580
      Top             =   108
      Width           =   7020
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter/Edit Tax Payments"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4092
      TabIndex        =   15
      Top             =   192
      Width           =   4020
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000B&
      FillStyle       =   0  'Solid
      Height          =   600
      Left            =   2592
      Top             =   -12
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
Attribute VB_Name = "frmTaxPayEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Dim TXUserName As String, TXCity As String, TXZip As String, TXState As String
Private Temp_Class As Resize_Class
Dim TaxCust As TaxCustType
Dim LTranNum As Long, Discount As Integer, DisPct As Double
'
'Code from Dos program '&*'''
'Also some code from this form was copied from Inv enter/edit still here and about
'&*&*&*&*&*&*&*&*&*&*&*&*&**&*&*
'  ReDim TaxPaymentRec(1) As TaxPaymentRecType
'  ReDim PayList(1 To 1) As PayListType
'  ReDim TaxCustRec(1) As TaxCustType
'  TaxPayRecLen = Len(TaxPaymentRec(1))
'  PayListLen = Len(PayList(1))
'  TaxCustRecLen = Len(TaxCustRec(1))
'
'  ReDim TaxSetup(1) As TaxMasterType
'  TaxSetupLen = Len(TaxSetup(1))
'  FGetAH "TAXSETUP.DAT", TaxSetup(1), TaxSetupLen, 1            'load it
'
'  RcptPort = TaxSetup(1).RcptPort
'  If RcptPort < 1 Then
'    RcptPort = 1
'  ElseIf RcptPort > 2 Then
'    RcptPort = 2
'  End If
'  GoSub LoadCustPayList
'
'  If RecpPort < 1 Or RecpPort > 2 Then
'    RecpPort = 1
'  End If
'
'  TOWNNAME$ = UCase$(TaxSetup(1).Name)
'
'  If InStr(TOWNNAME$, "HAMLET") > 0 Then
'    HamFlag = True
'  End If
'left some out here
'
'~~~~~~~~~~~~~~~~~~~
'What to do here?  LOADCUSTPAYLIST  Should these variables be global or not ?
'Check with Dale - Same files will be needed for Cash Management....
'LoadCustPayList:
'  Oper$ = QPTrim$(Str$(OperNum))
'  PayRecpName$ = "C:\TAXRCP" + Oper$ + ".RPT"
'  TaxCPRFileName$ = "TAXCPR" + Oper$ + ".DAT"   'Customers Payment Record file
'  TaxLOPFileName$ = "TAXLOP" + Oper$ + ".DAT"   'List Of Payments customers
'  PayRecFile = FreeFile
'  Open TaxCPRFileName$ For Random Shared As PayRecFile Len = TaxPayRecLen
'  NumofRecs& = LOF(PayRecFile) \ TaxPayRecLen
'  If NumofRecs& > 0 Then
'    ReDim CustList(1 To NumofRecs&) As CustPayListType
'    For Cnt& = 1 To NumofRecs&
'      Get #PayRecFile, Cnt&, TaxPaymentRec(1)
'      CustList(Cnt&).CustAcct = TaxPaymentRec(1).CustAcct
'      CustList(Cnt&).LastPayRec = TaxPaymentRec(1).LastPayRec
'      CustList(Cnt&).NumPayRec = TaxPaymentRec(1).NumPayRec
'    Next
'  End If
'  Close PayRecFile
'  CustListCnt& = NumofRecs&
'Return
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'SetOperInfo:
'  LSet Form$(1, 0) = FUsing$(Str$(OperNum), "##")
'  LSet Form$(2, 0) = PostDate$
'  Action = 2
'Return
'ChkCustList:
'  EditFlag = False
'  If CustListCnt& > 0 Then
'    For Cnt = 1 To CustListCnt&
'      If CustList(Cnt).CustAcct = CustAcct& Then
'        CustPayRec& = Cnt
'        NPicked = CustList(Cnt).NumPayRec
'        LastPayRec& = CustList(Cnt).LastPayRec
'        GoSub LoadEditCustPayList
'        EditFlag = True
'        Exit For
'      End If
'    Next
'  End If
'Return
'LoadEditCustPayList:
'  TPrinciple# = 0
'  TInterest# = 0
'  TCollection# = 0
'  LCnt = 0
'  ReDim TPayList(1) As PayListType
'  ReDim PayList(1 To 1) As PayListType
'  PayListFile = FreeFile
'  Open TaxLOPFileName$ For Random As PayListFile Len = PayListLen
'  ThisPayRec& = LastPayRec&
'  Do While ThisPayRec& > 0
'    LCnt = LCnt + 1
'    ReDim Preserve PayList(1 To LCnt) As PayListType
'    Get #PayListFile, ThisPayRec&, TPayList(1)
'    PayList(LCnt).BillRec = TPayList(1).BillRec
'    PayList(LCnt).CustRec = TPayList(1).CustRec
'    PayList(LCnt).Principle1 = TPayList(1).Principle1
'    TPrinciple# = Round#(TPrinciple# + TPayList(1).Principle1)
'    PayList(LCnt).Interest1 = TPayList(1).Interest1
'    TInterest# = Round#(TInterest# + TPayList(1).Interest1)
'    PayList(LCnt).Collection = TPayList(1).Collection
'    TCollection# = Round#(TCollection# + TPayList(1).Collection)
'    ThisPayRec& = TPayList(1).PrevListRec
'  Loop
'  NPicked = LCnt
'  Close PayListFile
'Return
'GetCustInfo:
'  GoSub ClearForm
'  NumOfCustRecs& = FileSize("TAXCUST.DAT") \ TaxCustRecLen
'  If CustAcct& > NumOfCustRecs& Or CustAcct& = 0 Then
'    CustAcct& = 0
'    ok = MsgBox%("TAX.QSL", "BADACCTN")
'    frm(1).FldNo = 1
'    GoSub SetOperInfo
'    GoTo SkipCustInfo
'  ElseIf IsCustDeleted(CustAcct&) Then
'    CustAcct& = 0
'    ok = MsgBox%("TAX.QSL", "DELACCTN")
'    frm(1).FldNo = 1
'    GoSub SetOperInfo
'    GoTo SkipCustInfo
'  End If
'
'  CustFile = FreeFile
'  Open "TAXCUST.DAT" For Random Shared As CustFile Len = TaxCustRecLen
'  Get CustFile, CustAcct&, TaxCustRec(1)
'  Close CustFile
'  If Not EditFlag Then
'    If Not DoesCustOwe%(TaxCustRec(1)) Then
'      CustAcct& = 0
'      SaveScrn TempScrn()
'      DisplayTaxScrn "ERRSCRN1"
'      QPrintRC "This customer has NO BALANCE!", 10, 26, -1
'      QPrintRC "Press any key to continue.", 13, 28, -1
'      WaitForAction
'      RestScrn TempScrn()
'      frm(1).FldNo = 1
'      GoSub SetOperInfo
'      GoTo SkipCustInfo
'    End If
'    LSet Form$(CustAcctFld, 0) = Str$(CustAcct&)
'    CustName$ = QPTrim$(TaxCustRec(1).FName) + " " + QPTrim$(TaxCustRec(1).LName)
'     LSet Form$(4, 0) = CustName$
'    LSet Form$(5, 0) = TaxCustRec(1).Addr1
'  Else
'    PayRecFile = FreeFile
'    Open TaxCPRFileName$ For Random Shared As PayRecFile Len = TaxPayRecLen
'    Get PayRecFile, CustPayRec&, TaxPaymentRec(1)
'    Close PayRecFile
'    BCopy VARSEG(TaxPaymentRec(1)), VarPtr(TaxPaymentRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), Len(Form$(0, 0)), 0
'    UnPackBuffer 0, 0, Form$(), Fld()
'  End If
'  CustAcct& = QPValL(Form$(CustAcctFld, 0))
'  FirstTime = True
'
'SkipCustInfo:
'  Action = 1
'Return
'    '--Check for Key presses
'    Select Case frm(1).KeyCode
'    Case F4KEY
'      If CustAcct& > 0 Then
'        ShowCustHistory CustAcct&
'      End If
'    Case EscKey
'      If BeenEditedFlag Then
'        SaveFlag = PromptSaveData
'        Select Case SaveFlag
'        Case True               'user wants to save
'          StuffBuf Chr$(0) + Chr$(Abs(F10Key))
'        Case False              'user wants to abandon
'          ExitFlag = True
'        Case Else               'continue editing
'        End Select
'        Action = 1
'      Else
'        ExitFlag = True
'      End If
'
'    Case F7KEY  'Lookup Customer
'       If frm(1).FldNo = 3 Then  'if user is on the Customer field
'        SaveScrn TempScrn()     'and F7key then do lookup routine
'        MPaintBox 4, 5, 22, 75, 8
'        LastCust& = CustAcct&
'        LookUp CustAcct&, "Payment", 0, False, False
'        RestScrn TempScrn()
'        If CustAcct& > 0 Then   'if this is a valid customer
'          GoSub ChkCustList
'          GoSub GetCustInfo     'go get customer info
'          frm(1).FldNo = 4
'          Action = 1
'        ElseIf LastCust& = CustAcct& Then
'          frm(1).FldNo = 1
'          Action = 1
'          ' don't do anything
'        Else
'          GoSub ClearForm
'          frm(1).FldNo = 1
'          Action = 1
'        End If
'      End If
'       Case F8KEY  'Select the bills being paid
'      If frm(1).FldNo = BillsFld Then
'        GoSub SelectBills2Pay
'      End If
'
'    Case F9KEY
'      TempAmtRecv# = Value#(Form$(AmtRecvFld, 0), ECode)
'      If TempAmtRecv# > 0 Then
'        GoSub AutoDistribute
'      End If
'
'    Case F10Key 'Save
'      GoSub CheckPaymentInfo
'      If PaymentOKFlag Then
'        Select Case AskSavePayment("Y")
'        Case 1  'Save trans print receipt
'          GoSub SaveTransaction
'          GoSub PrintReceipt
'          GoSub ClearForm
'          frm(1).FldNo = 1
'          Action = 1
'          EditFlag = False
'            GoSub LoadCustPayList
'        Case True               'Save trans no receipt
'          ReceiptFlag = False
'          GoSub SaveTransaction
'          GoSub ClearForm
'          frm(1).FldNo = 1
'          Action = 1
'          EditFlag = False
'          GoSub LoadCustPayList
'        Case False              'oops, just keep editing
'          Action = 2
'        End Select
'      End If
'    Case Is <> 0
'      'STOP
'    End Select
'
'    '--check for mouse clicks on buttons not attached to the form
'    If frm(1).Presses Then
'      Select Case frm(1).MRow
'      Case 22   'Look for the f10 or esc button
'        Select Case frm(1).MCol
'                Case 19 To 29           'f7 Look-Up
'          PressButton F7KEY, 22, 19, 29
'        Case 31 To 40           'f8 Bill select
'          PressButton F8KEY, 22, 31, 40
'        Case 42 To 50           'f9 Distrubt
'          PressButton F9KEY, 22, 42, 50
'        Case 54 To 63           'f10 Save
'          PressButton F10Key, 22, 54, 63
'        Case 65 To 75           '--cancel button
'          PressButton EscKey, 22, 65, 75
'        End Select
'      End Select                'row
'    End If
'  Loop Until ExitFlag
'
'  Erase TempScrn, TaxPaymentRec, TaxSetup
'
'  HideCursor
'
'  Close
'  Exit Sub
''*&*&*&*&*&*&*&*&*&*&*&*&*&*&*&*&*&*&*&*&*&*&*&*&*&*&*
Private Sub cmdBills_Click()
  Dim TaxCustFile As Integer
  Dim TaxFile As Integer, TaxRecLen As Integer
  If fplngAcct > 0 Then
  ReDim TaxCustT(1) As TaxCustType
  TaxRecLen = Len(TaxCustT(1))
  TaxFile = FreeFile
  Open "TaxCust.dat" For Random Shared As TaxFile Len = TaxRecLen
  Get #TaxFile, fplngAcct, TaxCustT(1)
  Close

  If DoesCustOwe%(TaxCustT(1)) Then
    SetBillList
  Else
    MsgBox "This Customer Does Not Owe A Balance.", vbOKOnly, "No Balance"
    fplngAcct.SetFocus
  End If

'  If Changed = True Then
'    If MsgBox("Changes Were Made to the Current Information on Screen and Not Saved." & Chr(13) & "Select OK to View List," & Chr(13) & "or Cancel to Remain on Current Record.", vbOKCancel, "View List?") = vbCancel Then
'      fpcboAcctNumNa.SetFocus
'      Exit Sub
'    End If
'  End If
'  Undolok RecNum
'  NextNew
'  If Check4Trans = True Then
'    frmInvListing.Show 1, frmInvEnterEdit
'    If Emode = True Then
'      SetScreen
'      fpcboAcctNumNa.SetFocus
'    End If
'  Else
'    MsgBox "No Entries To Display.", vbOKOnly, "No Entries"
'    fpcboAcctNumNa.SetFocus
  End If
    
End Sub

Public Sub SetBillList()
  'LTranNum is lasttrans from customer record
  Dim BillCnt As Integer, TransRecord As Long
  Dim tempstr As String, TransFile As Integer
  Dim Balance As Double
  BillCnt = 0
  ReDim taxtrans(1) As TaxTransactionType
  If LTranNum > 0 Then
    TransFile = FreeFile
    Open "TaxTrans.dat" For Random Shared As TransFile Len = Len(taxtrans(1))
    TransRecord& = LTranNum
    Do While TransRecord& <> 0
      Get TransFile, TransRecord&, taxtrans(1)
      If taxtrans(1).TranType = 1 Then
        Balance# = Round#(taxtrans(1).Revenue.Principle1 + taxtrans(1).Revenue.Principle2 + taxtrans(1).Revenue.Principle3 + taxtrans(1).Revenue.Principle4 + taxtrans(1).Revenue.Principle5)
        Balance# = Round#(Balance# + taxtrans(1).Revenue.Interest + taxtrans(1).Revenue.Penalty + taxtrans(1).Revenue.Collection)
        Balance# = Round#(Balance# - (taxtrans(1).Revenue.Principle1Pd + taxtrans(1).Revenue.Principle2Pd + taxtrans(1).Revenue.Principle3Pd + taxtrans(1).Revenue.Principle4Pd + taxtrans(1).Revenue.Principle5Pd))
        Balance# = Round#(Balance# - (taxtrans(1).Revenue.InterestPd + taxtrans(1).Revenue.PenaltyPd + taxtrans(1).Revenue.CollectionPd))
        If Balance# > 0 Then
          BillCnt = BillCnt + 1
          tempstr = Space$(80)
          'ReDim Preserve Items(1 To BillCnt) As FLen2
          LSet tempstr = Num2Date(taxtrans(1).TransDate)
          Mid$(tempstr, 25) = taxtrans(1).TAXYEAR 'Using$("####", taxtrans(1).TAXYEAR)
          Mid$(tempstr, 30, 8) = Str(TransRecord&)
          Mid$(tempstr, 52) = Using$("######.##", Str(taxtrans(1).Amount))
          Mid$(tempstr, 64) = Using$("######.##", Str(Balance#))
          ''''Mid$(Items(BillCnt).V, 61) = MKL$(TransRecord&)
          frmBillListing.lstTaxBills.AddItem tempstr$
        End If
      End If
      TransRecord& = taxtrans(1).LastTrans
    Loop

  Else
    Close
   'Unload frmLoadingRpt
    MsgBox "No Bills to Display", vbOKOnly, "No Bills"
    Exit Sub
  End If
  'Unload frmLoadingRpt
  
  Close
frmBillListing.Show 1
fpcboTenderType.SetFocus
End Sub

Public Sub Bill2Screen(BILLNUM() As Long)
  Dim TPrinciple As Double, TInterest As Double, TCollection As Double
  Dim TransFile As Integer, TAmtOwed As Double, upper As Integer
  Dim cnt As Integer, Disc As Integer, curyr As String
  ReDim taxtrans(1) As TaxTransactionType
    TPrinciple# = 0
    TInterest# = 0
    TCollection# = 0
    Disc = 0
    curyr = Year(Now)
    
    TransFile = FreeFile
    upper = UBound(BILLNUM())
    ReDim PayList(1 To upper) As PayListType
    Open "TaxTrans.dat" For Random Shared As TransFile Len = Len(taxtrans(1))
    
    For cnt = 1 To upper
    Get #TransFile, BILLNUM(cnt), taxtrans(1)
      PayList(cnt).BillRec = BILLNUM(cnt)
      PayList(cnt).CustRec = fplngAcct
      PayList(cnt).Principle1 = Round#(taxtrans(1).Revenue.Principle1 - taxtrans(1).Revenue.Principle1Pd)
      TPrinciple# = Round#(TPrinciple# + (taxtrans(1).Revenue.Principle1 - taxtrans(1).Revenue.Principle1Pd))
      PayList(cnt).Interest1 = Round#(taxtrans(1).Revenue.Interest - taxtrans(1).Revenue.InterestPd)
      TInterest# = Round#(TInterest# + (taxtrans(1).Revenue.Interest - taxtrans(1).Revenue.InterestPd))
      PayList(cnt).Collection = Round#(taxtrans(1).Revenue.Collection - taxtrans(1).Revenue.CollectionPd)
      TCollection# = Round#(TCollection# + (taxtrans(1).Revenue.Collection - taxtrans(1).Revenue.CollectionPd))
    Next
'Compare taxbill year with current year and check to see if only one
'bill was selected - current bill - and if so then set flag to
'calc discount below if setup has been set for discount
      If taxtrans(1).TAXYEAR = curyr And upper = 1 Then
        Disc = 1
      End If
    Close
    fpTaxOwed = Str$(TPrinciple#)
    fpIntOwed = Str$(TInterest#)
    fpAdColOwed = Str$(TCollection#)
    TAmtOwed# = Round#(TPrinciple# + TInterest# + TCollection#)
    fptxtAmtOwed = Str$(TAmtOwed#)
    fpTotOwed = Round#(fpTaxOwed.DoubleValue + fpIntOwed.DoubleValue + fpAdColOwed.DoubleValue)
    If Discount = 1 And DisPct > 0 Then
      If Disc = 1 Then
        fpDiscAmt = Round#(DisPct * fpTaxOwed)
      Else
        fpDiscAmt = 0
      End If
    Else
      fpDiscAmt = 0
    End If
    fpcboTenderType.ListIndex = -1
    fpCashAmt = 0
    fpChkAmt = 0
    
End Sub

Private Sub Autodist()
  Dim left As Double
  left = 0
  If fpTotReceived.DoubleValue > 0 Then
  If fpTotReceived.DoubleValue >= fpTotOwed.DoubleValue Then
    fpTaxPaid = fpTaxOwed
    fpIntPaid = fpIntOwed
    fpAdColPaid = fpAdColOwed
  Else
    left = fpTotReceived
    If left <= fpAdColOwed Then
      fpAdColPaid = left
      fpIntPaid = 0
      fpTaxPaid = 0
    Else
      fpAdColPaid = fpAdColOwed
      left = Round#(fpTotReceived.DoubleValue - fpAdColPaid.DoubleValue)
      If left <= fpIntOwed Then
        fpIntPaid = left
        fpTaxPaid = 0
      Else
        fpIntPaid = fpIntOwed
        left = Round#(fpTotReceived.DoubleValue - fpIntPaid.DoubleValue)
        If left <= fpTaxOwed Then
          fpTaxPaid = left
        Else
          fpTaxPaid = fpTaxOwed
        End If
      End If
    End If
    
  End If
    fpTotPaid = Round#(fpTaxPaid.DoubleValue + fpIntPaid.DoubleValue + fpAdColPaid.DoubleValue)
  End If
End Sub
Private Function SetScreen()
'  If Emode = False Then  'This is in New Mode
'    cmdNew.Enabled = False
'    cmdEdit.Enabled = True
'    lblNew.Visible = True
'    lblEdit.Visible = False
'  Else               'This is in Edit Mode
'    cmdNew.Enabled = True
'    cmdEdit.Enabled = False
'    cmdDelete.Enabled = True
'    lblNew.Visible = False
'    lblEdit.Visible = True
'  End If
  
End Function

Private Sub cmdCash_Click()
  fpcboTenderType.ListIndex = 0
  fpChkAmt.Enabled = False
  fpCashAmt.Enabled = True
  fpChkAmt = 0
  fpCashAmt.SetFocus
End Sub

Private Sub cmdCheck_Click()
  fpcboTenderType.ListIndex = 1
  fpCashAmt.Enabled = False
  fpChkAmt.Enabled = True
  fpCashAmt = 0
  fpChkAmt.SetFocus
End Sub

Private Sub cmdDist_Click()
  Autodist
End Sub

Private Sub fpCashAmt_Change()
fpTotReceived = Round#(fpCashAmt.DoubleValue + fpChkAmt.DoubleValue + fpDiscAmt.DoubleValue)
fpChangeDue = Round#(fpTotReceived.DoubleValue - fptxtAmtOwed)
End Sub


Private Sub fpChkAmt_Change()
fpTotReceived = Round#(fpCashAmt.DoubleValue + fpChkAmt.DoubleValue + fpDiscAmt.DoubleValue)
fpChangeDue = Round#(fpTotReceived.DoubleValue - fptxtAmtOwed)
End Sub

Private Sub fpDiscAmt_Change()
fpTotReceived = Round#(fpCashAmt.DoubleValue + fpChkAmt.DoubleValue + fpDiscAmt.DoubleValue)
fpChangeDue = Round#(fpTotReceived.DoubleValue - fptxtAmtOwed)
End Sub

'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'  If ((UnloadMode = vbFormControlMenu)) Then
'    If Changed = False Then
'      If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
'        Cancel = True
'      Else
'        MainLog "Close Tax via Payment"
'        ClearInUse PWcnt
'      End If
'    Else
'      If MsgBox("Changes Have Been Made to the Current Record." & Chr(13) & Chr(13) & "                          Select OK to Abandon and Close Program," & Chr(13) & Chr(13) & "       or Cancel to Remain on Entry/Edit Screen.", vbOKCancel, "Abandon Changes?") = vbOK Then
'        MainLog "Close Tax via Payment"
'        ClearInUse PWcnt
'      Else
'        Cancel = True
'      End If
'    End If
'  End If
'End Sub
Private Sub fplngAcct_Change()
Dim Acct As Long
    Acct = fplngAcct
    If Acct > 0 Then
      If Acct > GetTaxCustCnt Then
        MsgBox "Bad Account Number.", vbOKOnly, "Invalid Account"
        fplngAcct.SetFocus
        Exit Sub
      ElseIf IsCustDeleted(Acct) Then
        MsgBox "Deleted Account.", vbOKOnly, "Deleted Account"
        fplngAcct.SetFocus
        Exit Sub
      Else
       'If DoesCustOwe(Acct) Then
          Cust2Screen (Acct)
       ' Else
       '   MsgBox "This Customer Does Not Owe A Balance.", vbOKOnly, "No Balance"
      End If
    Else
      MsgBox "Bad Account Number.", vbOKOnly, "Invalid Account"
      fplngAcct.SetFocus
      Exit Sub
    End If
End Sub
Public Function Cust2Screen(TempRec)
  Dim TaxCustFile As Integer, NumofCust As Long, Rec As Long
  Dim TaxFileNum As Integer
  If Exist("Taxcust.dat") Then
  If TempRec > 0 Then
    OpenTaxCustFile TaxFileNum, NumofCust
    Get TaxFileNum, TempRec, TaxCust
    If Not TaxCust.Deleted Then
      If TaxCust.LastTrans > 0 Then
        LTranNum = TaxCust.LastTrans
      Else
        LTranNum = 0
      End If
      fptxtName = QPTrim(TaxCust.FName) & " " & QPTrim(TaxCust.LName)
      fptxtAddress = QPTrim(TaxCust.ADDR1)
      Close TaxCustFile
      'End If
      Else
        MsgBox "Customer Deleted", vbOKOnly, "Request Denied"
        Close
      End If
 End If
 End If
End Function

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  GetTaxInfo TXUserName, TXState, TXCity, TXZip
  StatusBar1.Panels.Item(1).Text = TXUserName
  IsDiscSet Discount, DisPct
  If Discount = 1 Then
    fpDiscAmt.Enabled = True
  Else
    fpDiscAmt.Enabled = False
  End If
  fpcboTenderType.InsertRow = "1" & Chr$(9) & "Cash"
  fpcboTenderType.InsertRow = "2" & Chr$(9) & "Check"
  fpcboTenderType.InsertRow = "3" & Chr$(9) & "Cash & Check"
  fpcboTenderType.InsertRow = "4" & Chr$(9) & "Other"
End Sub
Private Sub ChkCustList()
'  EditFlag = False
'  If CustListCnt& > 0 Then
'    For Cnt = 1 To CustListCnt&
'      If CustList(Cnt).CustAcct = CustAcct& Then
'        CustPayRec& = Cnt
'        NPicked = CustList(Cnt).NumPayRec
'        LastPayRec& = CustList(Cnt).LastPayRec
'        GoSub LoadEditCustPayList
'        EditFlag = True
'        Exit For
'      End If
'    Next
'  End If
End Sub
'Private Sub ClearScn()
'    Emode = False
'    SetScreen
'    LoadControl
'    txtDate.Text = Format(Now, "mm/dd/yyyy")
'
'    fpcboVendCode.ListIndex = -1
'    fplstVendor.Clear
'    txtTotPOAmt = 0
'    fpcboDepartment.ListIndex = -1
'    vaSpread1.ClearRange 1, 1, 7, 36, True
'    txtTotDistAmt = 0
'    ClearBuds
'End Sub
'Private Sub LoadControl()
''loads info from SETUP FILE
' SUCH AS DISCOUNT ?
'AND ANY SPECIAL PAYMENT INFO ABOUT THIS CUSTOMER

'End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    DoEvents
    Temp_Class.ResizeControls Me
    'DoEvents
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
    Case vbKeyUp:
      SendKeys "+{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      cmdExit_Click
      KeyCode = 0
'    Case vbKeyF10:
'      cmdSave_Click
'      KeyCode = 0
    Case vbKeyF9:
      cmdDist_Click
      KeyCode = 0
    Case vbKeyF8:
      cmdBills_Click
      KeyCode = 0
'    Case vbKeyF4:
'      cmdEdit_Click
'      KeyCode = 0
    Case vbKeyF5:
      cmdCash_Click
      KeyCode = 0
'    Case vbKeyF3:
'      cmdDelete_Click
'      KeyCode = 0
    Case vbKeyF6:
      cmdCheck_Click
      KeyCode = 0
'    Case vbKeyPageDown:
'      Call cmdDist_Click
'      KeyCode = 0
'    Case vbKeyPageUp:
'      Call cmdPage1_Click
'      KeyCode = 0
    Case Else:
  End Select
End Sub

Private Sub cmdExit_Click()
  frmTaxPayMenu.Show
  Unload frmTaxPayEntry

'  If Changed = False Then
'    Undolok RecNum
'    frmPOProcessMenu.Show
'    Unload frmPOEnterEdit
'  Else
'    If MsgBox("Changes Have Been Made to the Current Record." & Chr(13) & Chr(13) & "                          Select OK to Abandon," & Chr(13) & Chr(13) & "       or Cancel to Remain on Entry/Edit Screen.", vbOKCancel, "Abandon Changes?") = vbOK Then
'      Undolok RecNum
'      frmPOProcessMenu.Show
'      Unload frmPOEnterEdit
'    End If
''*****      'figure out how to get focus to proper place.
'     ' fpcbo.SetFocus
'
'  End If
End Sub
'Private Function Changed()
'  Dim POFile As Integer, POFileLen As Integer, NumRecs As Integer
'  Dim POEditFile As Integer, NumEdTrans As Integer
'  Dim Cnt As Integer
'  ReDim POCont(1) As POControlRecType
'
'  If Val(txtStock) <> 0 Then GoTo DoChange
'  If Val(txtDesc) <> 0 Then GoTo DoChange
'  If Val(txtQty) <> 0 Then GoTo DoChange
'  If txtPrice <> 0 Then GoTo DoChange
'  If fpcboAcctNumNa.ListIndex <> -1 Then GoTo DoChange
'  If Emode = False Then
'    If fpcboDepartment.ListIndex <> -1 Then GoTo DoChange
'    If fpcboVendCode.ListIndex <> -1 Then GoTo DoChange
'    If txtTotPOAmt <> 0 Then GoTo DoChange
'    If Val(txtShipOn) <> 0 Then GoTo DoChange
'    OpenPOFile POFile, NumRecs
'    If LOF(POFile) > 0 Then
'      Get POFile, 1, POCont(1)
'      If QPTrim(txtFOB) <> QPTrim(POCont(1).FOB) Then GoTo DoChange
'      If QPTrim(txtShipVia) <> QPTrim(POCont(1).Shipvia) Then GoTo DoChange
'      If QPTrim(txtTerms) <> QPTrim(POCont(1).Terms) Then GoTo DoChange
'      If QPTrim(txtShipTo1) <> QPTrim(POCont(1).Shipto1) Then GoTo DoChange
'      If QPTrim(txtShipTo2) <> QPTrim(POCont(1).Shipto2) Then GoTo DoChange
'      If QPTrim(txtShipTo3) <> QPTrim(POCont(1).Shipto3) Then GoTo DoChange
'      If QPTrim(txtShipTo4) <> QPTrim(POCont(1).Shipto4) Then GoTo DoChange
'      If QPTrim(txtShipTo5) <> QPTrim(POCont(1).Shipto5) Then GoTo DoChange
'      If QPTrim(txtAddinst1) <> QPTrim(POCont(1).Addinst1) Then GoTo DoChange
'      If QPTrim(txtAddinst2) <> QPTrim(POCont(1).Addinst2) Then GoTo DoChange
'      If QPTrim(txtAddinst3) <> QPTrim(POCont(1).Addinst3) Then GoTo DoChange
'    End If
'    Close POFile
'    vaSpread1.Row = 1
'    vaSpread1.col = 1
'    If Val(vaSpread1.Text) <> 0 Then GoTo DoChange
'    Changed = False
'  Else
'    OpenPOEditFile POEditFile, NumEdTrans
'    Get POEditFile, RecNum, POEdit
'    If txtDate <> Format(DateAdd("d", (POEdit.PODATE), "12-31-1979"), "mm/dd/yyyy") Then GoTo DoChangeClose
'    fpcboDepartment.col = 1
'    If fpcboDepartment.ColText <> POEdit.REQNUM Then GoTo DoChangeClose
'    fpcboVendCode.col = 1
'    If fpcboVendCode.ColText <> POEdit.VNDRREC Then GoTo DoChangeClose
'    If QPTrim(txtFOB) <> QPTrim(POEdit.FOB) Then GoTo DoChangeClose
'    If QPTrim(txtShipVia) <> QPTrim(POEdit.Shipvia) Then GoTo DoChangeClose
'    If QPTrim(txtTerms) <> QPTrim(POEdit.Terms) Then GoTo DoChangeClose
'    If QPTrim(txtShipOn) <> QPTrim(POEdit.SHIPON) Then GoTo DoChangeClose
'    If txtTotPOAmt.DoubleValue <> POEdit.POAmt Then GoTo DoChangeClose
'    If QPTrim(txtShipTo1) <> QPTrim(POEdit.SHPLINE1) Then GoTo DoChangeClose
'    If QPTrim(txtShipTo2) <> QPTrim(POEdit.SHPLINE2) Then GoTo DoChangeClose
'    If QPTrim(txtShipTo3) <> QPTrim(POEdit.SHPLINE3) Then GoTo DoChangeClose
'    If QPTrim(txtShipTo4) <> QPTrim(POEdit.SHPLINE4) Then GoTo DoChangeClose
'    If QPTrim(txtShipTo5) <> QPTrim(POEdit.SHPLINE5) Then GoTo DoChangeClose
'    If QPTrim(txtAddinst1) <> QPTrim(POEdit.Addinst1) Then GoTo DoChangeClose
'    If QPTrim(txtAddinst2) <> QPTrim(POEdit.Addinst2) Then GoTo DoChangeClose
'    If QPTrim(txtAddinst3) <> QPTrim(POEdit.Addinst3) Then GoTo DoChangeClose
'
'      For Cnt = 1 To 36
'        vaSpread1.Row = Cnt
'        vaSpread1.col = 1
'        If Val(vaSpread1.Text) <> POEdit.ITEMS(Cnt).AcctRec Then GoTo DoChangeClose
'        If Val(vaSpread1.Text) = 0 Then
'          Changed = False
'          Exit For
'        End If
'        vaSpread1.col = 2
'        If vaSpread1.Text <> QPTrim(POEdit.ITEMS(Cnt).STKNO) Then GoTo DoChangeClose
'        vaSpread1.col = 3
'        If vaSpread1.Text <> QPTrim(POEdit.ITEMS(Cnt).Desc) Then GoTo DoChangeClose
'        vaSpread1.col = 4
'        If vaSpread1.Text <> POEdit.ITEMS(Cnt).QUAN Then GoTo DoChangeClose
'        vaSpread1.col = 5
'        If vaSpread1.Text <> POEdit.ITEMS(Cnt).PRICE Then GoTo DoChangeClose
'        vaSpread1.col = 7
'        If vaSpread1.Text <> QPTrim(POEdit.ITEMS(Cnt).ACCTNO) Then GoTo DoChangeClose
'        Changed = False
'      Next
'    Close POEditFile
'    End If
'    Exit Function
'DoChange:
'   Changed = True
'   Exit Function
'DoChangeClose:
'   Changed = True
'   Close POEditFile
'   Exit Function
'
'End Function
'Private Function Check4Trans()
'  Dim POEditFile As Integer, NumEdTrans As Integer
'  Dim Cnt As Integer, Good As Integer
'  Good = 0
'  If Exist("APPED.dat") Then
'    OpenPOEditFile POEditFile, NumEdTrans
'    If NumEdTrans > 0 Then
'      For Cnt = 1 To NumEdTrans
'        Get POEditFile, Cnt, POEdit
'        If POEdit.Deleted <> True Then
'          If QPTrim(POEdit.PONum) = "N/A" Then
'            Good = Good + 1
'          End If
'        End If
'      Next
'    Else
'      Check4Trans = False
'    End If
'  Else
'    Check4Trans = False
'  End If
'  If Good > 0 Then
'    Check4Trans = True
'  Else
'    Check4Trans = False
'  End If
' Close POEditFile
' End Function

'Private Sub cmdEdit_Click()
' If Changed = True Then
'    If MsgBox("Changes Were Made to the Current Information on Screen and Not Saved." & Chr(13) & "Select OK to View Edit List," & Chr(13) & "or Cancel to Remain on Current Record.", vbOKCancel, "View Edit List?") = vbCancel Then
'      vaTabPro1.ActivePage = 0
'      txtDate.SetFocus
'      Exit Sub
'    End If
'  End If
'  Undolok RecNum
'  NextNew
'  If Check4Trans = True Then
'    frmPOListing.Show 1, frmPOEnterEdit
'    If Emode = True Then
'      SetScreen
'      'DisplayTotals
'      txtDate.SetFocus
'      cmdDelete.Enabled = True
'    End If
'  Else
'    MsgBox "No Entries To Display.", vbOKOnly, "No Entries"
'    txtDate.SetFocus
'  End If
'End Sub

'Private Sub cmdList_Click()
'  If Changed = True Then
'    If MsgBox("Changes Were Made to the Current Information on Screen and Not Saved." & Chr(13) & "Select OK to View List," & Chr(13) & "or Cancel to Remain on Current Record.", vbOKCancel, "View List?") = vbCancel Then
'      vaTabPro1.ActivePage = 0
'      txtDate.SetFocus
'      Exit Sub
'    End If
'  End If
'  Undolok RecNum
'  NextNew
'  If Check4Trans = True Then
'    frmPOListing.Show 1, frmPOEnterEdit
'    If Emode = True Then
'      SetScreen
'      vaTabPro1.ActivePage = 0
'      cmdDelDist.Enabled = False
'      txtDate.SetFocus
'    End If
'  Else
'    MsgBox "No Entries To Display.", vbOKOnly, "No Entries"
'    vaTabPro1.ActivePage = 0
'    cmdDelDist.Enabled = False
'    txtDate.SetFocus
'  End If
'End Sub

'Private Sub cmdNew_Click()
'  Dim POBusy As Boolean
'  POBusy = False
'  If Exist("APPED.DAT") Then POBusy = GetAttr("APPED.DAT") And vbReadOnly
'  If Not POBusy Then
'    If Changed = True Then
'      If MsgBox("Changes Have Been Made to the Current Record." & Chr(13) & "Select OK to Abandon," & Chr(13) & "or Cancel to Remain on Current Record.", vbOKCancel, "Abandon Changes?") = vbCancel Then
'        txtDate.SetFocus
'        Exit Sub
'      End If
'    End If
'    Undolok RecNum
'    NextNew
'    vaTabPro1.ActivePage = 0
'    cmdDelDist.Enabled = False
'    txtDate.SetFocus
'  Else
'    MsgBox "Posting Is In Progress, Editing Not Allowed At This Time.", vbOKOnly, "Canceled"
'    frmPOProcessMenu.Show
'    Unload frmPOEnterEdit
'  End If
'End Sub

'Private Function VerifyEntered()
'  If txtQty > 0 Then
'    If txtPrice > 0 Then
'      If fpcboAcctNumNa.ListIndex <> -1 Then
'        VerifyEntered = True
'      Else
'        VerifyEntered = False
'        fpcboAcctNumNa.SetFocus
'        Exit Function
'      End If
'    Else
'      VerifyEntered = False
'      txtPrice.SetFocus
'      Exit Function
'    End If
'  Else
'    VerifyEntered = False
'    txtQty.SetFocus
'    Exit Function
'  End If
'End Function

 
'Private Function Ready2Save()
'  Dim TempDate As Integer, Cnt As Integer
'  Dim TempDist As Double
'  TempDist = 0
'  'Take care of Invalid Data and Messages in this Section
'  'CheckValDate is in main module to verify date entered w/correct format
'  If CheckValDate(txtDate) = True Then
'    TempDate = DateDiff("d", "12/31/1979", txtDate)
'  'Also compare date with Hi/Lo range
'    If (TempDate < LPDate) Or (TempDate > HPDate) Then
'      MsgBox "This Date Is Not Within Allowable Posting Range. Please Correct or Change Setup.", vbOKOnly, "Invalid Date"
'      Ready2Save = False
'      Exit Function
'    Else
'      Ready2Save = True
'    End If
'  Else
'    MsgBox "This Date Is Not Valid. Please Correct.", vbOKOnly, "Invalid Date"
'    Ready2Save = False
'    Exit Function
'  End If
'  'Not allow Zero Total or Unequal Distritbutions
'  If txtTotPOAmt <> 0 Then
'    If txtTotDistAmt = 0 Or txtTotPOAmt <> txtTotDistAmt Then
'      MsgBox "The Total Purchase Order Does Not Equal The Amount of The Distributions." & Chr$(13) & "Please Correct Before Saving.", vbOKOnly, "PO Entry"
'      Ready2Save = False
'      Exit Function
'    Else
'
'      For Cnt = 1 To 36
'        vaSpread1.col = 6
'        vaSpread1.Row = Cnt
'        If vaSpread1.Text <> "" Then
'          TempDist = Round(vaSpread1.Text + TempDist)
'        Else
'          Exit For
'        End If
'      Next
'      If TempDist <> txtTotDistAmt Or TempDist <> txtTotPOAmt Then
'        MsgBox "Totals Are Not In Balance. Please Correct.", vbOKOnly, "PO Entry"
'        Ready2Save = False
'        Exit Function
'      Else
'        Ready2Save = True
'      End If
'     End If
'  Else
'    MsgBox "You May Not Save A Purchase Order With A $0.00 Total.", vbOKOnly, "PO Entry"
'    Ready2Save = False
'  End If
'End Function
'Private Sub fpcboAcctNumNa_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeySpace Then
'    fpcboAcctNumNa.ListDown = True
'  End If
'  If KeyCode = vbKeyDelete Then
'    ClearBuds
'    fpcboAcctNumNa.ListIndex = -1
'    fpcboAcctNumNa.Action = ActionClearSearchBuffer
'  End If
'  If fpcboAcctNumNa.ListDown <> True Then
'    If KeyCode = vbKeyDown Then
'        SendKeys "{Tab}"
'        KeyCode = 0
'    Else
'      If KeyCode = vbKeyUp Then
'        SendKeys "+{Tab}"
'        KeyCode = 0
'      End If
'    End If
'  End If
'
'End Sub


'Private Sub cmdSave_Click()
'  If Ready2Save = True Then
'    SavePO
'    Call NextNew
'  Else
'    MsgBox "             Save Canceled.", vbOKOnly, "PO Entry"
'  End If
'End Sub

'Private Sub SavePO()
'  Dim POEditFile As Integer, NumEdTrans As Integer, Cnt As Integer
'  Dim POBusy As Boolean
'  POBusy = False
'  If Exist("APPED.DAT") Then POBusy = GetAttr("APPED.DAT") And vbReadOnly
'  If Not POBusy Then
'    OpenPOEditFile POEditFile, NumEdTrans
'    POEdit.Deleted = 0
'    POEdit.Locked = False
'    POEdit.PONum = Trim(txtPONumber)
'    fpcboDepartment.col = 1
'    POEdit.REQNUM = QPTrim(fpcboDepartment.ColText)
'    'POEdit.REQNUM = QPTrim(fpcboDepartment.Text)
'    POEdit.PODATE = DateDiff("d", "12/31/1979", txtDate)
'    fpcboVendCode.col = 0
'    POEdit.VNDRCODE = fpcboVendCode.ColText
'    fpcboVendCode.col = 1
'    POEdit.VNDRREC = fpcboVendCode.ColText
'    fplstVendor.col = -1
'    fplstVendor.Selected(0) = True
'    POEdit.VNDRINF1 = QPTrim(fplstVendor.Text)
'    fplstVendor.Selected(1) = True
'    POEdit.VNDRINF2 = QPTrim(fplstVendor.Text)
'    fplstVendor.Selected(2) = True
'    POEdit.VNDRINF3 = QPTrim(fplstVendor.Text)
'    fplstVendor.Selected(3) = True
'    POEdit.VNDRINF4 = QPTrim(fplstVendor.Text)
'    fplstVendor.Selected(4) = True
'    POEdit.VNDRINF5 = QPTrim(fplstVendor.Text)
'    POEdit.FOB = QPTrim(txtFOB)
'    POEdit.Shipvia = QPTrim(txtShipVia)
'    POEdit.Terms = QPTrim(txtTerms)
'    POEdit.SHIPON = QPTrim(txtShipOn)
'    POEdit.POAmt = txtTotPOAmt
'    POEdit.SHPLINE1 = QPTrim(txtShipTo1)
'    POEdit.SHPLINE2 = QPTrim(txtShipTo2)
'    POEdit.SHPLINE3 = QPTrim(txtShipTo3)
'    POEdit.SHPLINE4 = QPTrim(txtShipTo4)
'    POEdit.SHPLINE5 = QPTrim(txtShipTo5)
'    POEdit.Addinst1 = QPTrim(txtAddinst1)
'    POEdit.Addinst2 = QPTrim(txtAddinst2)
'    POEdit.Addinst3 = QPTrim(txtAddinst3)
'
'    For Cnt = 1 To 36
'      vaSpread1.Row = Cnt
'      vaSpread1.col = 1
'      If vaSpread1.Text = "" Then
'        POEdit.ITEMS(Cnt).AcctRec = 0
'        POEdit.ITEMS(Cnt).STKNO = ""
'        POEdit.ITEMS(Cnt).ACCTNO = ""
'        POEdit.ITEMS(Cnt).Desc = ""
'        POEdit.ITEMS(Cnt).EXT = 0
'        POEdit.ITEMS(Cnt).PRICE = 0
'        POEdit.ITEMS(Cnt).QUAN = 0
'      Else
'        POEdit.ITEMS(Cnt).AcctRec = vaSpread1.Text
'        vaSpread1.col = 2
'        POEdit.ITEMS(Cnt).STKNO = QPTrim(vaSpread1.Text)
'        vaSpread1.col = 3
'        POEdit.ITEMS(Cnt).Desc = QPTrim(vaSpread1.Text)
'        vaSpread1.col = 4
'        POEdit.ITEMS(Cnt).QUAN = vaSpread1.Text
'        vaSpread1.col = 5
'        POEdit.ITEMS(Cnt).PRICE = vaSpread1.Text
'        vaSpread1.col = 6
'        POEdit.ITEMS(Cnt).EXT = vaSpread1.Text
'        vaSpread1.col = 7
'        POEdit.ITEMS(Cnt).ACCTNO = QPTrim(vaSpread1.Text)
'      End If
'    Next
'    If Emode = False Then
'      If NumEdTrans > 0 Then
'        RecNum = NumEdTrans + 1
'      Else
'        RecNum = 1
'      End If
'    End If
'    Put POEditFile, RecNum, POEdit
'    Close POEditFile
'  Else
'    MsgBox "Posting Is In Progress, Editing Not Allowed At This Time.", vbOKOnly, "Canceled"
'    frmPOProcessMenu.Show
'    Unload frmPOEnterEdit
'  End If
'End Sub
'

Private Sub mnuExit_Click()
  Call cmdExit_Click
End Sub

'Private Sub NextNew()
'  Dim POEditFile As Integer, NumEdTrans As Integer
'  OpenPOEditFile POEditFile, NumEdTrans
'  Close POEditFile
'   If NumEdTrans > 0 Then
'     RecNum = NumEdTrans + 1
'   Else
'     RecNum = 1
'   End If
'
'   Emode = False
'   ClearFields
'   SetScreen
'   LoadControl
'   vaTabPro1.ActivePage = 0
'   cmdDelDist.Enabled = False
'   fpcboDepartment.SetFocus
'End Sub
'
Private Sub fpcboTenderType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcboTenderType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcboTenderType.ListIndex = -1
    fpcboTenderType.Action = ActionClearSearchBuffer
    fpCashAmt = 0
    fpChkAmt = 0
  End If
  If fpcboTenderType.ListDown <> True Then
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
Private Sub fpcboTenderType_LostFocus()
  fpcboTenderType.Action = ActionClearSearchBuffer
  If fpcboTenderType.ListIndex = 0 Then
    fpCashAmt.Enabled = True
    fpChkAmt.Enabled = False
    fpCashAmt.SetFocus
  ElseIf fpcboTenderType.ListIndex = 1 Then
    fpCashAmt.Enabled = False
    fpChkAmt.Enabled = True
    fpChkAmt.SetFocus
  ElseIf fpcboTenderType.ListIndex = 2 Then
    fpCashAmt.Enabled = True
    fpChkAmt.Enabled = True
    fpCashAmt.SetFocus
  ElseIf fpcboTenderType.ListIndex = 3 Then
    fpCashAmt.Enabled = True
    fpChkAmt.Enabled = False
    fpCashAmt.SetFocus
  End If
End Sub
'Private Sub mnuPrnScn_Click()
'  PrintForm
'End Sub
'
'
'Private Sub txtDate_LostFocus()
'  If CheckValDate(txtDate) = False Then
'    MsgBox "Invalid Date, Please Correct.", vbOKOnly, "Invalid Date"
'    txtDate.SetFocus
'  End If
'End Sub
'
'Private Sub txtFOB_ChangeMode(EditMode As Integer)
'  EditMode = True
'End Sub
'
'
'Private Sub txtShipVia_ChangeMode(EditMode As Integer)
'  EditMode = True
'End Sub
'Private Sub txtTerms_ChangeMode(EditMode As Integer)
'  EditMode = True
'End Sub
'Private Sub txtStock_ChangeMode(EditMode As Integer)
'  EditMode = True
'End Sub
'Private Sub txtDesc_ChangeMode(EditMode As Integer)
'  EditMode = True
'End Sub
'
'Private Sub txtAddinst1_ChangeMode(EditMode As Integer)
'  EditMode = True
'End Sub
'Private Sub txtAddinst2_ChangeMode(EditMode As Integer)
'  EditMode = True
'End Sub
'Private Sub txtAddinst3_ChangeMode(EditMode As Integer)
'  EditMode = True
'End Sub
'Private Sub txtShipOn_ChangeMode(EditMode As Integer)
'  EditMode = True
'End Sub
'Private Sub txtShipTo1_ChangeMode(EditMode As Integer)
'  EditMode = True
'End Sub
'Private Sub txtShipTo2_ChangeMode(EditMode As Integer)
'  EditMode = True
'End Sub
'Private Sub txtShipTo3_ChangeMode(EditMode As Integer)
'  EditMode = True
'End Sub
'Private Sub txtShipTo4_ChangeMode(EditMode As Integer)
'  EditMode = True
'End Sub
'Private Sub txtShipTo5_ChangeMode(EditMode As Integer)
'  EditMode = True
'End Sub
'
'Private Sub txtTotPOAmt_Change()
'  txtTot2 = txtTotPOAmt
'End Sub
'
'CheckPaymentInfo:
'
'  'Parse and move data to Paylist records here
'  PaymentOKFlag = True
'  PrinceOw# = fpTaxOwed
'  PrincePD# = fpTaxPaid
'  InterestOw# = fpIntOwed
'  InterestPd# = fpIntPaid
'  CollectOw# = fpAdColOwed
'  CollectPd# = fpAdColPaid
'  TDiscAmt# = fpDiscAmt
'  TAmtRecv# = fpTotReceived
'  TAmtPaid# = fpTotPaid
'  ChangeAmt# = fpChangeDue
'
'  If TAmtPaid# = 0 Then
'    MsgBox "Amount Paid Must Be more than 0.", vbOKOnly, "Invalid Amount"        'show bad scrn
'    'Action = 2
'    PaymentOKFlag = False
'    'frm(1).FldNo = frm(1).PrevFld
'    GoTo BadPayment
'  End If
'  If TAmtRecv# = Round#(TAmtPaid# + ChangeAmt#) And TAmtRecv# > 0 And ChangeAmt# >= 0 Then
'      PaymentOKFlag = True
'  Else
'    'ok = MsgBox%("TAX.QSL", "BADPYTOT")         'show bad scrn
'    'Action = 2
'    PaymentOKFlag = False
'    'frm(1).FldNo = frm(1).PrevFld
'    GoTo BadPayment
'  End If
' ' TenderType$ = QPTrim$(Form$(TenderFld, 0))
'  'If Len(TenderType$) = 0 Then
'  If fpcboTenderType.ListIndex = -1 Then
'    'ok = MsgBox%("TAX.QSL", "BADTENDR")
'    'Action = 2
'    PaymentOKFlag = False
'    'frm(1).FldNo = TenderFld
'    GoTo BadPayment
'  End If
'
'  If (PrincePD# > PrinceOw#) Or (InterestPd# > InterestOw#) Or (CollectPd# > CollectOw#) Then
''     SaveScrn TempScrn()
''    DisplayTaxScrn "ERRSCRN1"
''    QPrintRC "Can not overpay Tax Payments.", 10, 27, -1
''    QPrintRC "Correct and Save transaction Again.", 12, 24, -1
''    WaitForAction
''    RestScrn TempScrn()
''    Action = 2
''    PaymentOKFlag = False
''    frm(1).FldNo = AmtFlds(1)
'    GoTo BadPayment
'  End If
'
'  For Cnt = 1 To NPicked
'    PPrinciple# = Round#(PayList(Cnt).Principle1)
'    If (PrincePD# >= PPrinciple#) And (PrincePD# > 0) Then
'      PrincePD# = Round#(PrincePD# - PPrinciple#)
'    Else
'      If PrincePD# > 0 Then
'        PayList(Cnt).Principle1 = Round(PrincePD#)
'        PrincePD# = 0
'      Else
'        PayList(Cnt).Principle1 = 0
'      End If
'    End If
'    PInterest# = Round(PayList(Cnt).Interest1)
'    If (InterestPd# >= PInterest#) And (InterestPd# > 0) Then
'      InterestPd# = Round#(InterestPd# - PInterest#)
'    Else
'      If InterestPd# <> 0 Then
'        PayList(Cnt).Interest1 = Round(InterestPd#)
'        InterestPd# = 0
'      Else
'        PayList(Cnt).Interest1 = 0
'      End If
'    End If
'    PCollect# = Round(PayList(Cnt).Collection)
'    If (CollectPd# >= PCollect#) And (CollectPd# > 0) Then
'      CollectPd# = Round#(CollectPd# - PCollect#)
'    Else
'      If CollectPd# > 0 Then
'        PayList(Cnt).Collection = Round(CollectPd#)
'        CollectPd# = 0
'      Else
'        PayList(Cnt).Collection = 0
'      End If
'    End If
'  Next
'BadPayment:
'
'Return
