VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{8DDE6232-1BB0-11D0-81C3-0080C7A2EF7D}#3.0#0"; "Flp32a30.ocx"
Begin VB.Form frmVATaxPersPaymentEntry 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CM Personal Property Tax Payment Entry"
   ClientHeight    =   8760
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   11655
   Icon            =   "frmVATaxPersPaymentEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin LpLib.fpCombo fpcmbTenderType 
      Height          =   390
      Left            =   3060
      TabIndex        =   4
      Top             =   4695
      Width           =   2175
      _Version        =   196608
      _ExtentX        =   3836
      _ExtentY        =   688
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
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
      ColDesigner     =   "frmVATaxPersPaymentEntry.frx":08CA
   End
   Begin EditLib.fpCurrency fpCurrOpt1Owed 
      Height          =   372
      Left            =   8040
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   5330
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2773
      _ExtentY        =   656
      Enabled         =   0   'False
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
      NoSpecialKeys   =   2
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
   Begin VB.Timer Timer1 
      Left            =   185
      Top             =   240
   End
   Begin VB.Timer MsgAlertTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   185
      Top             =   720
   End
   Begin EditLib.fpLongInteger fpLongAcctNum 
      Height          =   372
      Left            =   2760
      TabIndex        =   0
      Top             =   1800
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
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
   Begin EditLib.fpText fptxtInterest 
      Height          =   372
      Left            =   5868
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4560
      Width           =   2052
      _Version        =   196608
      _ExtentX        =   3619
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
      NoSpecialKeys   =   2
      AutoAdvance     =   0   'False
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
      Text            =   "INTEREST"
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin EditLib.fpCurrency fpCurrAmtOwed 
      Height          =   372
      Left            =   3060
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4320
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3836
      _ExtentY        =   661
      Enabled         =   0   'False
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
      NoSpecialKeys   =   2
      AutoAdvance     =   0   'False
      AutoBeep        =   0   'False
      CaretInsert     =   0
      CaretOverWrite  =   3
      UserEntry       =   0
      HideSelection   =   -1  'True
      InvalidColor    =   -2147483642
      InvalidOption   =   0
      MarginLeft      =   3
      MarginTop       =   3
      MarginRight     =   3
      MarginBottom    =   3
      NullColor       =   -2147483641
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
   Begin EditLib.fpText fptxtState 
      Height          =   372
      Left            =   1500
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3552
      Width           =   612
      _Version        =   196608
      _ExtentX        =   1085
      _ExtentY        =   661
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
      NoSpecialKeys   =   2
      AutoAdvance     =   0   'False
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
      MaxLength       =   2
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
   Begin EditLib.fpDateTime fptxtPayDate 
      Height          =   348
      Left            =   8580
      TabIndex        =   3
      Tag             =   "The date you enter here will be the date that appears on the 'Payment Entry' screen. The date on that screen is not editable."
      Top             =   1200
      Width           =   1740
      _Version        =   196608
      _ExtentX        =   3069
      _ExtentY        =   609
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
      ControlType     =   1
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
      ButtonColor     =   13684944
      AutoMenu        =   0   'False
      StartMonth      =   4
      ButtonAlign     =   0
      BoundDataType   =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpText fptxtName 
      Height          =   372
      Left            =   1500
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2400
      Width           =   4212
      _Version        =   196608
      _ExtentX        =   7429
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
      NoSpecialKeys   =   2
      AutoAdvance     =   0   'False
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
      MaxLength       =   255
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
      Height          =   372
      Left            =   1500
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2784
      Width           =   4212
      _Version        =   196608
      _ExtentX        =   7429
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
      NoSpecialKeys   =   2
      AutoAdvance     =   0   'False
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
      MaxLength       =   255
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
   Begin EditLib.fpText fptxtCity 
      Height          =   372
      Left            =   1500
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3168
      Width           =   4212
      _Version        =   196608
      _ExtentX        =   7429
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
      NoSpecialKeys   =   2
      AutoAdvance     =   0   'False
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
      MaxLength       =   255
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
   Begin EditLib.fpMask fptxtZip 
      Height          =   372
      Left            =   4260
      TabIndex        =   24
      TabStop         =   0   'False
      Tag             =   "This field contains the postal code for this business. This field cannot be edited."
      Top             =   3552
      Width           =   1452
      _Version        =   196608
      _ExtentX        =   2566
      _ExtentY        =   661
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   1
      AlignTextH      =   1
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
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
      ControlType     =   1
      AllowOverflow   =   0   'False
      BestFit         =   0   'False
      ClipMode        =   0
      DataFormatEx    =   0
      Mask            =   "#####-####"
      PromptChar      =   "_"
      PromptInclude   =   0   'False
      RequireFill     =   0   'False
      BorderGrayAreaColor=   -2147483637
      NoPrefix        =   0   'False
      ThreeDOnFocusInvert=   0   'False
      ThreeDFrameColor=   -2147483633
      Appearance      =   0
      BorderDropShadow=   0
      BorderDropShadowColor=   -2147483632
      BorderDropShadowWidth=   3
      AutoTab         =   0   'False
      ButtonColor     =   -2147483633
      AutoMenu        =   0   'False
      ButtonAlign     =   0
      OLEDropMode     =   0
      OLEDragMode     =   0
   End
   Begin EditLib.fpCurrency fpCurrCashPd 
      Height          =   372
      Left            =   3060
      TabIndex        =   5
      Top             =   5088
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3836
      _ExtentY        =   661
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
      MinValue        =   "0"
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
   Begin EditLib.fpCurrency fpCurrChkChrgPd 
      Height          =   372
      Left            =   3060
      TabIndex        =   6
      Top             =   5470
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3836
      _ExtentY        =   661
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
      MinValue        =   "0"
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
   Begin EditLib.fpCurrency fpCurrTotRecd 
      Height          =   372
      Left            =   3060
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6372
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3836
      _ExtentY        =   661
      Enabled         =   0   'False
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
      NoSpecialKeys   =   2
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
   Begin EditLib.fpCurrency fpCurrChngDue 
      Height          =   372
      Left            =   3060
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   6766
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3836
      _ExtentY        =   661
      Enabled         =   0   'False
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
      NoSpecialKeys   =   2
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
   Begin EditLib.fpText fptxtDescription 
      Height          =   372
      Left            =   1680
      TabIndex        =   17
      Top             =   7420
      Width           =   3972
      _Version        =   196608
      _ExtentX        =   7006
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
      AutoAdvance     =   0   'False
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
      CharValidationText=   ""
      MaxLength       =   19
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
   Begin EditLib.fpText fptxtPers 
      Height          =   372
      Left            =   5880
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2640
      Width           =   2052
      _Version        =   196608
      _ExtentX        =   3625
      _ExtentY        =   661
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
      NoSpecialKeys   =   2
      AutoAdvance     =   0   'False
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
      Text            =   "PERSONAL"
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin EditLib.fpText fptxtMachTools 
      Height          =   372
      Left            =   5868
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   3022
      Width           =   2052
      _Version        =   196608
      _ExtentX        =   3625
      _ExtentY        =   661
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
      NoSpecialKeys   =   2
      AutoAdvance     =   0   'False
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
      Text            =   "MACHINE TOOLS"
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin EditLib.fpText fptxtMerchCap 
      Height          =   372
      Left            =   5868
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   3408
      Width           =   2052
      _Version        =   196608
      _ExtentX        =   3619
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
      NoSpecialKeys   =   2
      AutoAdvance     =   0   'False
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
      Text            =   "MERCHANT CAP"
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin EditLib.fpText fptxtFarmEquip 
      Height          =   372
      Left            =   5868
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   3792
      Width           =   2052
      _Version        =   196608
      _ExtentX        =   3619
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
      NoSpecialKeys   =   2
      AutoAdvance     =   0   'False
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
      Text            =   "FARM EQUIPMENT"
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin EditLib.fpText fptxtMobHomes 
      Height          =   372
      Left            =   5868
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   4178
      Width           =   2052
      _Version        =   196608
      _ExtentX        =   3619
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
      NoSpecialKeys   =   2
      AutoAdvance     =   0   'False
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
      Text            =   "MOBILE HOME"
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin EditLib.fpText fptxtPenalty 
      Height          =   372
      Left            =   5868
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   4946
      Width           =   2052
      _Version        =   196608
      _ExtentX        =   3619
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
      NoSpecialKeys   =   2
      AutoAdvance     =   0   'False
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
      Text            =   "PENALTY"
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin EditLib.fpCurrency fpCurrPersOwed 
      Height          =   372
      Left            =   8028
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2640
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
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
      NoSpecialKeys   =   2
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
   Begin EditLib.fpCurrency fpCurrPersPaid 
      Height          =   372
      Left            =   9708
      TabIndex        =   7
      Top             =   2640
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
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
      MinValue        =   "0"
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
   Begin EditLib.fpCurrency fpCurrMTOwed 
      Height          =   372
      Left            =   8028
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   3022
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
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
      ButtonStyle     =   0
      ButtonWidth     =   0
      ButtonWrap      =   -1  'True
      ButtonDefaultAction=   -1  'True
      ThreeDText      =   0
      ThreeDTextHighlightColor=   -2147483633
      ThreeDTextShadowColor=   -2147483632
      ThreeDTextOffset=   2
      AlignTextH      =   2
      AlignTextV      =   0
      AllowNull       =   0   'False
      NoSpecialKeys   =   2
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
   Begin EditLib.fpCurrency fpCurrMTPaid 
      Height          =   372
      Left            =   9708
      TabIndex        =   8
      Top             =   3022
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
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
      MinValue        =   "0"
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
   Begin EditLib.fpCurrency fpCurrMCOwed 
      Height          =   372
      Left            =   8028
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   3408
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2773
      _ExtentY        =   656
      Enabled         =   0   'False
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
      NoSpecialKeys   =   2
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
   Begin EditLib.fpCurrency fpCurrMCPaid 
      Height          =   372
      Left            =   9708
      TabIndex        =   9
      Top             =   3408
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
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
      MinValue        =   "0"
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
   Begin EditLib.fpCurrency fpCurrFEOwed 
      Height          =   372
      Left            =   8028
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   3792
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2773
      _ExtentY        =   656
      Enabled         =   0   'False
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
      NoSpecialKeys   =   2
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
   Begin EditLib.fpCurrency fpCurrFEPaid 
      Height          =   372
      Left            =   9708
      TabIndex        =   10
      Top             =   3792
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
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
      MinValue        =   "0"
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
   Begin EditLib.fpCurrency fpCurrMHOwed 
      Height          =   372
      Left            =   8040
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   4178
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2773
      _ExtentY        =   656
      Enabled         =   0   'False
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
      NoSpecialKeys   =   2
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
   Begin EditLib.fpCurrency fpCurrMHPaid 
      Height          =   372
      Left            =   9708
      TabIndex        =   11
      Top             =   4178
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
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
      MinValue        =   "0"
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
   Begin EditLib.fpCurrency fpCurrIntOwed 
      Height          =   372
      Left            =   8028
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   4560
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2773
      _ExtentY        =   656
      Enabled         =   0   'False
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
      NoSpecialKeys   =   2
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
   Begin EditLib.fpCurrency fpCurrIntPaid 
      Height          =   372
      Left            =   9708
      TabIndex        =   12
      Top             =   4560
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
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
      MinValue        =   "0"
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
   Begin EditLib.fpCurrency fpCurrPenOwed 
      Height          =   372
      Left            =   8028
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   4946
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2773
      _ExtentY        =   656
      Enabled         =   0   'False
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
      NoSpecialKeys   =   2
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
   Begin EditLib.fpCurrency fpCurrPenPaid 
      Height          =   372
      Left            =   9708
      TabIndex        =   13
      Tag             =   "1"
      Top             =   4946
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
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
      MinValue        =   "0"
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
   Begin EditLib.fpCurrency fpCurrTotOwed 
      Height          =   372
      Left            =   8028
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   6720
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
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
      NoSpecialKeys   =   2
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
   Begin EditLib.fpCurrency fpCurrTotPaid 
      Height          =   372
      Left            =   9708
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   6720
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
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
      BackColor       =   16777215
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
      NoSpecialKeys   =   2
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
   Begin EditLib.fpCurrency fpCurrDisc 
      Height          =   372
      Left            =   3060
      TabIndex        =   42
      Top             =   5854
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3836
      _ExtentY        =   661
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
      MinValue        =   "0"
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
   Begin EditLib.fpCurrency fpCurrTotWDisc 
      Height          =   372
      Left            =   9708
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   7420
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2773
      _ExtentY        =   656
      Enabled         =   0   'False
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
      BackColor       =   16777215
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
      NoSpecialKeys   =   2
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
   Begin EditLib.fpCurrency fpCurrPrePay 
      Height          =   372
      Left            =   9708
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2778
      _ExtentY        =   661
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
      MinValue        =   "0"
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
   Begin fpBtnAtlLibCtl.fpBtn cmdLookup 
      Height          =   372
      Left            =   4500
      TabIndex        =   2
      Top             =   1800
      Width           =   1812
      _Version        =   131072
      _ExtentX        =   3196
      _ExtentY        =   656
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
      ButtonDesigner  =   "frmVATaxPersPaymentEntry.frx":0BC1
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdBills 
      Height          =   372
      Left            =   6420
      TabIndex        =   1
      Top             =   1800
      Width           =   1452
      _Version        =   131072
      _ExtentX        =   2561
      _ExtentY        =   656
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
      ButtonDesigner  =   "frmVATaxPersPaymentEntry.frx":0DA3
   End
   Begin EditLib.fpText fptxtOpt1 
      Height          =   372
      Left            =   5880
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   5330
      Width           =   2052
      _Version        =   196608
      _ExtentX        =   3619
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
      NoSpecialKeys   =   2
      AutoAdvance     =   0   'False
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
      Text            =   "OPT REV 1"
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin EditLib.fpText fptxtOpt2 
      Height          =   372
      Left            =   5880
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   5710
      Width           =   2052
      _Version        =   196608
      _ExtentX        =   3619
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
      NoSpecialKeys   =   2
      AutoAdvance     =   0   'False
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
      Text            =   "OPT REV 2"
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin EditLib.fpText fptxtOpt3 
      Height          =   372
      Left            =   5880
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   6100
      Width           =   2052
      _Version        =   196608
      _ExtentX        =   3619
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
      NoSpecialKeys   =   2
      AutoAdvance     =   0   'False
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
      Text            =   "OPT REV 3"
      CharValidationText=   ""
      MaxLength       =   255
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
   Begin EditLib.fpCurrency fpCurrOpt1Paid 
      Height          =   372
      Left            =   9720
      TabIndex        =   14
      Tag             =   "1"
      Top             =   5330
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2773
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
      MinValue        =   "0"
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
   Begin EditLib.fpCurrency fpCurrOpt2Owed 
      Height          =   372
      Left            =   8040
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   5710
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2773
      _ExtentY        =   656
      Enabled         =   0   'False
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
      NoSpecialKeys   =   2
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
   Begin EditLib.fpCurrency fpCurrOpt2Paid 
      Height          =   372
      Left            =   9720
      TabIndex        =   15
      Tag             =   "1"
      Top             =   5710
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2773
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
      MinValue        =   "0"
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
   Begin EditLib.fpCurrency fpCurrOpt3Owed 
      Height          =   372
      Left            =   8040
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   6100
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2773
      _ExtentY        =   656
      Enabled         =   0   'False
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
      NoSpecialKeys   =   2
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
   Begin EditLib.fpCurrency fpCurrOpt3Paid 
      Height          =   372
      Left            =   9720
      TabIndex        =   16
      Tag             =   "1"
      Top             =   6100
      Width           =   1572
      _Version        =   196608
      _ExtentX        =   2773
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
      MinValue        =   "0"
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
   Begin fpBtnAtlLibCtl.fpBtn cmdCash 
      Height          =   420
      Left            =   3048
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   8184
      Width           =   1284
      _Version        =   131072
      _ExtentX        =   2265
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmVATaxPersPaymentEntry.frx":0F7F
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCheck 
      Height          =   420
      Left            =   4476
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   8184
      Width           =   1284
      _Version        =   131072
      _ExtentX        =   2265
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmVATaxPersPaymentEntry.frx":115A
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdCharge 
      Height          =   420
      Left            =   5895
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   8190
      Width           =   1275
      _Version        =   131072
      _ExtentX        =   2249
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmVATaxPersPaymentEntry.frx":1336
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDist 
      Height          =   420
      Left            =   7320
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   8184
      Width           =   1284
      _Version        =   131072
      _ExtentX        =   2265
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmVATaxPersPaymentEntry.frx":1513
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   420
      Left            =   8748
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   8184
      Width           =   1284
      _Version        =   131072
      _ExtentX        =   2265
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmVATaxPersPaymentEntry.frx":16EE
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   420
      Left            =   10176
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   8184
      Width           =   1284
      _Version        =   131072
      _ExtentX        =   2265
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmVATaxPersPaymentEntry.frx":18CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdInfo 
      Height          =   420
      Left            =   1620
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   8184
      Width           =   1284
      _Version        =   131072
      _ExtentX        =   2265
      _ExtentY        =   741
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
      ButtonDesigner  =   "frmVATaxPersPaymentEntry.frx":1AA6
   End
   Begin fpBtnAtlLibCtl.fpBtn fpcmdDrawer 
      Height          =   420
      Left            =   195
      TabIndex        =   86
      Top             =   8190
      Width           =   1275
      _Version        =   131072
      _ExtentX        =   2249
      _ExtentY        =   741
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      GrayAreaColor   =   12632256
      BorderShowDefault=   -1  'True
      ButtonType      =   0
      NoPointerFocus  =   -1  'True
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
      ButtonDesigner  =   "frmVATaxPersPaymentEntry.frx":1C81
   End
   Begin VB.Line Line10 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   180
      X2              =   11340
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Line Line11 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   180
      X2              =   5790
      Y1              =   7280
      Y2              =   7280
   End
   Begin VB.Line Line9 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   5820
      X2              =   11340
      Y1              =   7280
      Y2              =   7280
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CM Personal Tax Payment  Entry"
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
      Left            =   3144
      TabIndex        =   72
      Top             =   432
      Width           =   5364
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000009&
      Height          =   456
      Left            =   2256
      Top             =   372
      Width           =   7008
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Acct Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   228
      TabIndex        =   71
      Top             =   1920
      Width           =   2412
   End
   Begin VB.Label lblOperNum 
      BackStyle       =   0  'Transparent
      Caption         =   "Operator Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   1260
      TabIndex        =   70
      Top             =   1280
      Width           =   2412
   End
   Begin VB.Label lblOperName 
      BackStyle       =   0  'Transparent
      Caption         =   "Operator Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   4260
      TabIndex        =   69
      Top             =   1280
      Width           =   3012
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Source: Tax"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5424
      TabIndex        =   68
      Top             =   984
      Width           =   2532
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   7860
      TabIndex        =   67
      Top             =   1280
      Width           =   612
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   420
      TabIndex        =   66
      Top             =   2520
      Width           =   852
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   420
      TabIndex        =   65
      Top             =   2888
      Width           =   852
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "City:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   420
      TabIndex        =   64
      Top             =   3278
      Width           =   852
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "State:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   420
      TabIndex        =   63
      Top             =   3648
      Width           =   852
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Zip:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   3180
      TabIndex        =   62
      Top             =   3648
      Width           =   852
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Tax  Billing Payment"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   292
      Left            =   180
      TabIndex        =   61
      Top             =   4040
      Width           =   2412
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Owed:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   1140
      TabIndex        =   60
      Top             =   4440
      Width           =   1692
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tender Type:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   1140
      TabIndex        =   59
      Top             =   4810
      Width           =   1692
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Amount Paid:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   660
      TabIndex        =   58
      Top             =   5172
      Width           =   2172
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Check/Charge Amt Paid:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   300
      TabIndex        =   57
      Top             =   5578
      Width           =   2532
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Received:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   420
      TabIndex        =   56
      Top             =   6492
      Width           =   2412
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Change Due:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   420
      TabIndex        =   55
      Top             =   6862
      Width           =   2412
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   5820
      X2              =   5820
      Y1              =   2280
      Y2              =   7920
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Revenue"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   6324
      TabIndex        =   54
      Top             =   2280
      Width           =   1092
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Amt Owed"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   8148
      TabIndex        =   53
      Top             =   2280
      Width           =   1332
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Amt Paid"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   9828
      TabIndex        =   52
      Top             =   2280
      Width           =   1332
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Totals:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   6060
      TabIndex        =   51
      Top             =   6792
      Width           =   1692
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   7980
      X2              =   7980
      Y1              =   1680
      Y2              =   7280
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   8028
      X2              =   9588
      Y1              =   6608
      Y2              =   6608
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   9708
      X2              =   11268
      Y1              =   6608
      Y2              =   6608
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   240
      TabIndex        =   50
      Top             =   7480
      Width           =   1452
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Discount:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   1860
      TabIndex        =   49
      Top             =   5974
      Width           =   972
   End
   Begin VB.Line Line6 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   9660
      X2              =   9660
      Y1              =   2280
      Y2              =   7280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   180
      X2              =   180
      Y1              =   2280
      Y2              =   7920
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   3060
      X2              =   5220
      Y1              =   6310
      Y2              =   6310
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount Paid With Discount:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   372
      Left            =   6060
      TabIndex        =   48
      Top             =   7480
      Width           =   3492
   End
   Begin VB.Line Line12 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   11340
      X2              =   11340
      Y1              =   7920
      Y2              =   2280
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Prepay Amt:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   252
      Left            =   8220
      TabIndex        =   47
      Top             =   1920
      Width           =   1332
   End
   Begin VB.Label lblCurrTaxYr 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Tax Year:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   1320
      TabIndex        =   46
      Top             =   936
      Width           =   2412
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   9420
      TabIndex        =   45
      Top             =   120
      Visible         =   0   'False
      Width           =   2052
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   180
      X2              =   5820
      Y1              =   4040
      Y2              =   4040
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   612
      Left            =   180
      Top             =   1680
      Width           =   11172
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   612
      Left            =   2268
      Top             =   240
      Width           =   7020
   End
End
Attribute VB_Name = "frmVATaxPersPaymentEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim ScrWidth As Long
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim FirstBillRec As Long
  Dim BtnFnt As Double, Lunney As Boolean
  Public NotFirstLoad As Boolean
  Public TempAcctNum As Long
  Dim DiscPXDate As Integer
  Dim ThisDiscAmt As Double
  Dim ThisDiscPct As Double
  Dim CustList() As VACustPayListType
  Dim CustListCnt&
  Public EditFlag As Boolean
  Public GetNewCust As Boolean
  Dim ExitFlag As Boolean
  Dim LastPayRec&, CustPayRec&
  Dim TempBillList() As VAPersPayListType
  Dim TempBillListCnt As Integer
  Dim TempPayDate As Integer
  Dim TempAmtOwed As Double
  Dim TempTenderTY As String
  Dim TempCashAmt As Double
  Dim TempChkAmt As Double
  Dim TempChrgAmt As Double
  Dim TempAmtRecd As Double
  Dim TempChange As Double
  Dim TempDesc As String
  Dim TempPreYear As Integer
  Dim TempPaidOwed1AmtOwed As Double
  Dim TempPaidOwed2AmtOwed As Double
  Dim TempPaidOwed3AmtOwed As Double
  Dim TempPaidOwed4AmtOwed As Double
  Dim TempPaidOwed5AmtOwed As Double
  Dim TempPaidOwed6AmtOwed As Double
  Dim TempPaidOwed7AmtOwed As Double
  Dim TempPaidOwed8AmtOwed As Double
  Dim TempPaidOwed9AmtOwed As Double
  Dim TempPaidOwed10AmtOwed As Double
  Dim TempPaidOwed1AmtPaid As Double
  Dim TempPaidOwed2AmtPaid As Double
  Dim TempPaidOwed3AmtPaid As Double
  Dim TempPaidOwed4AmtPaid As Double
  Dim TempPaidOwed5AmtPaid As Double
  Dim TempPaidOwed6AmtPaid As Double
  Dim TempPaidOwed7AmtPaid As Double
  Dim TempPaidOwed8AmtPaid As Double
  Dim TempPaidOwed9AmtPaid As Double
  Dim TempPaidOwed10AmtPaid As Double
  Dim TempTotOwed As Double
  Dim TempAmtPaid As Double
  Dim TempTotPaid As Double
  Dim DontExit As Boolean
  Dim CustDiscAmt As Double
  Dim DiscAmtAry() As Double
  Dim DiscRecAry() As Long
  Dim DiscAryCnt As Integer
  Dim StopWarn As Boolean
  Dim WhichRec() As Integer
  Dim DiscCnt As Integer
  Dim InClear As Boolean
  Dim TempPersPaid As Double
  Dim TempMTPaid As Double
  Dim TempMCPaid As Double
  Dim TempFEPaid As Double
  Dim TempMHPaid As Double
  Dim TempIntPaid As Double
  Dim TempPenPaid As Double
  Dim TempOpt1Paid As Double
  Dim TempOpt2Paid As Double
  Dim TempOpt3Paid As Double
  Dim TempDisc As Double
  Dim TempTotPd As Double
  Dim TempPrePay As Double
  Dim MaxDisc As Double
  Dim InOverRideDist As Boolean
  Dim InSave As Boolean
  Dim OverPay As Boolean
  Dim BillHasFocus As Boolean
  Dim RctValidate As Boolean
  Dim RecpPort As String
  Dim DiscYN As Boolean
  Dim SaveMode As Boolean
  Dim WasSaved As Boolean
  Dim TaxYear As Integer
  Dim Opt1Desc$
  Dim Opt2Desc$
  Dim Opt3Desc$
  Public ThisBillType$
  Public LookUp As Boolean '2/14/06
  Dim PayOrder() As Integer
Dim fromform As Form, toform As Form, codeopt As Integer
Dim DefPayDate As String
  Dim BegAmount As Double
  'OpenTempPayFile is the same as open TaxCPRFileName
  'OpenPayListFile is the same as open TaxLOPFileName
Public Sub Wheretogo(xfrm As Form, tfrm As Form, Optional opt As Integer, Optional DDate As String)
  Set fromform = xfrm
  Set toform = tfrm
  If opt <> 0 Then
    codeopt = opt
  Else
    codeopt = 0
  End If
  If DDate <> "" Then
    DefPayDate = DDate
  End If
End Sub

Public Sub cmdBills_Click()
  Dim TaxRec As VATaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long, luncnt As Long, Laugh As Long
  Dim ThisAmtOwed As Double
  Dim ThisCust As Long
  
  On Local Error GoTo ERRORSTUFF
  
  ThisAmtOwed = CDbl(fpCurrAmtOwed.Value)
  If fpLongAcctNum.Value = 0 Then
    frmVATaxMsg.Label1.Caption = "Please supply a customer first before accessing bill information."
    frmVATaxMsg.Label1.Top = 800
    frmVATaxMsg.Show vbModal
    If fpLongAcctNum.Enabled = True Then
      fpLongAcctNum.SetFocus
    End If
    Exit Sub
  End If
  If Lunney = True Then
    luncnt = 8000000
    StandBy (luncnt)
  End If

  OpenVATaxCustFile CHandle, NumOfCRecs
  Get CHandle, GCustNum, TaxRec
  Close CHandle
  
  If VAGetCustPersBalance(GCustNum, -1) = 0 Then
    frmVATaxMsg.Label1.Caption = "This customer has a zero personal balance."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
    fpLongAcctNum.SetFocus
    Exit Sub
  End If
  
  frmVATaxBillList.Show vbModal
  DoEvents
'  If EditFlag = True Then Exit Sub 'critical

  ThisCust = 0
  If BillCnt > 0 Or Exist(VATempPersBillRecs) Then 'BillCnt is a temporary value representing
  'the number of bills tagged for payment. TempBillRecs is a temporary file that keeps up
  'with the currently tagged bill and serves as a backup to BillCnt
    Call LoadAmtOwed
  Else
    fpCurrDisc = 0 '9/20/05
    Call ResetLeaveName
  End If
  
  GetNewCust = False
  fpcmbTenderType.SetFocus
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersPaymentEntry", "cmdBills_Click", Erl)
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
 '   ClearInUse PWcnt
 '  CMTerminate

End Sub

Private Sub cmdBills_GotFocus()
  BillHasFocus = True
End Sub

Private Sub cmdBills_LostFocus()
  BillHasFocus = False
End Sub

Private Sub cmdCash_Click()
  On Local Error GoTo ERRORSTUFF
  
  If fpCurrAmtOwed.DoubleValue > 0 Then
    fpcmbTenderType.Text = "CASH"
    Call ClearPaidFields
    If fpCurrDisc.Value = 0 Then
      fpCurrDisc.Value = 0
      fpCurrCashPd = OldRound(fpCurrAmtOwed.DoubleValue)
      fpCurrTotRecd = OldRound(fpCurrAmtOwed.DoubleValue)
    Else
      fpCurrDisc.Value = MaxDisc
      fpCurrCashPd = OldRound(fpCurrAmtOwed.DoubleValue - MaxDisc)
      fpCurrTotRecd = OldRound(fpCurrAmtOwed.DoubleValue - MaxDisc)
    End If
    Call cmdDist_Click
    Call ReFigure
    fptxtDescription.SetFocus
  ElseIf fpCurrAmtOwed.DoubleValue = 0 Then
    Call VATaxMsg(900, "Automatic distribution can only take place if there is an amount owed.")
    If fpCurrCashPd.Enabled = True Then
      fpCurrCashPd.SetFocus
    End If
  End If
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersPaymentEntry", "cmdCash_Click", Erl)
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
'    ClearInUse PWcnt
'    CMTerminate

End Sub

Private Sub cmdCharge_Click()
  On Local Error GoTo ERRORSTUFF
  
  If fpCurrAmtOwed.DoubleValue > 0 Then
    fpcmbTenderType.Text = "CHARGE"
    Call ClearPaidFields
    If fpCurrDisc.Value = 0 Then
      fpCurrDisc.Value = 0
      fpCurrChkChrgPd = OldRound(fpCurrAmtOwed.DoubleValue)
      fpCurrTotRecd = OldRound(fpCurrAmtOwed.DoubleValue)
    Else
      fpCurrDisc.Value = MaxDisc
      fpCurrChkChrgPd = OldRound(fpCurrAmtOwed.DoubleValue - MaxDisc)
      fpCurrTotRecd = OldRound(fpCurrAmtOwed.DoubleValue - MaxDisc)
    End If
    Call cmdDist_Click
    Call ReFigure
    fptxtDescription.SetFocus
  ElseIf fpCurrAmtOwed.DoubleValue = 0 Then
    Call VATaxMsg(900, "Automatic distribution can only take place if there is an amount owed.")
    If fpCurrCashPd.Enabled = True Then
      fpCurrCashPd.SetFocus
    End If
  End If
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersPaymentEntry", "cmdCharge_Click", Erl)
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
 '   ClearInUse PWcnt
 '   CMTerminate

End Sub

Private Sub cmdCheck_Click()
  On Local Error GoTo ERRORSTUFF
  
  If fpCurrAmtOwed.DoubleValue > 0 Then
    fpcmbTenderType.Text = "CHECK"
    Call ClearPaidFields
    If fpCurrDisc.Value = 0 Then
      fpCurrDisc.Value = 0
      fpCurrChkChrgPd = OldRound(fpCurrAmtOwed.DoubleValue)
      fpCurrTotRecd = OldRound(fpCurrAmtOwed.DoubleValue)
    Else
      fpCurrDisc.Value = MaxDisc
      fpCurrChkChrgPd = OldRound(fpCurrAmtOwed.DoubleValue - MaxDisc)
      fpCurrTotRecd = OldRound(fpCurrAmtOwed.DoubleValue - MaxDisc)
    End If
    Call cmdDist_Click
    Call ReFigure
    fptxtDescription.SetFocus
  ElseIf fpCurrAmtOwed.DoubleValue = 0 Then
    Call VATaxMsg(900, "Automatic distribution can only take place if there is an amount owed.")
    If fpCurrCashPd.Enabled = True Then
      fpCurrCashPd.SetFocus
    End If
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersPaymentEntry", "cmdCheck_Click", Erl)
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
 '   ClearInUse PWcnt
 '   CMTerminate

End Sub

Private Sub cmdCustHist_Click()
  If GCustNum = 0 Then
    Exit Sub
  End If
  frmVATaxCustInfoTHist.Show vbModal
  DoEvents
  Me.Hide
End Sub

Private Sub cmdDist_Click()
  Dim TaxRec As VATaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim SetUpRec As VATaxMasterType
  Dim SHandle As Integer
  Dim x As Integer
  Dim TotRecd As Double
  Dim TaxTrans As VATaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TransRecord&
  Dim WhatsLeft As Double
  Dim PaidDif As Double
  Dim ThisDif As Double
  Dim TPayRec As VAPersPayListType
  Dim PayRec As VAPersPayListType
  Dim PHandle As Integer
  Dim NumOfPRecs As Integer
  Dim y As Integer, z As Integer
  Dim ThisPrevRec As Long
  Dim NewRec As Integer
  Dim Operator$
  Dim TempRec As TempPayList
  Dim NumOfTemps As Integer
  Dim TempHandle As Integer
  Dim SmallNum As Integer
  Dim HoldDate As Integer
  Dim HoldNum As Long
  Dim Nextx As Integer
  Dim Thisx As Integer
  
  On Local Error GoTo ERRORSTUFF
  
  If CDbl(fpCurrAmtOwed.Value) = 0 And Val(fpLongAcctNum.Text) = 0 Then
    Exit Sub
  End If
  
  If fpCurrCashPd.Value = 0 And fpCurrChkChrgPd.Value = 0 And fpCurrTotPaid.Value = 0 Then
    Call VATaxMsg(900, "Please enter an amount paid.")
    If fpcmbTenderType = "CASH" Then
      If fpCurrCashPd.Enabled = True Then
        fpCurrCashPd.BackColor = &H8080FF
        fpCurrCashPd.SetFocus
      End If
    Else
      If fpCurrChkChrgPd.Enabled = True Then
        fpCurrChkChrgPd.BackColor = &H8080FF
        fpCurrChkChrgPd.SetFocus
      End If
    End If
    
    Exit Sub
    DoEvents
  ElseIf fpCurrCashPd.Value = 0 And fpCurrChkChrgPd.Value = 0 And fpCurrTotPaid.Value > 0 Then
    fpCurrPersPaid.Value = 0
    fpCurrMTPaid.Value = 0
    fpCurrMCPaid.Value = 0
    fpCurrFEPaid.Value = 0
    fpCurrMHPaid.Value = 0
    fpCurrIntPaid.Value = 0
    fpCurrPenPaid.Value = 0
    fpCurrOpt1Paid.Value = 0
    fpCurrOpt2Paid.Value = 0
    fpCurrOpt3Paid.Value = 0
    fpCurrTotPaid.Value = 0
  End If
    
  TotRecd = fpCurrTotRecd.Value
'  WhatsLeft = OldRound(CDbl(fpCurrTotRecd.Value) + CDbl(fpCurrDisc.Value))
  
  OpenVATaxSetUpFile SHandle
  Get SHandle, 1, SetUpRec
  Close SHandle
  
  If BillCnt = 0 And EditFlag = True Then 'user is editing and is not accessing
  'the bill list
    ReDim BillTrans(1 To 1) As Long
    ReDim BillDate(1 To 1) As Integer
    ThisPrevRec = 0
    NewRec = 0
    Operator$ = CStr(OperNum)
    Operator$ = QPTrim$(Operator$)
    OpenPersPayListFile PHandle, OperNum 'saved by getting data from temporary
    'bill record
    NumOfPRecs = LOF(PHandle) / Len(PayRec)
    For x = 1 To NumOfPRecs
      Get PHandle, x, PayRec
      If PayRec.CustRec = GCustNum Then
        BillCnt = BillCnt + 1
        ReDim Preserve BillTrans(1 To BillCnt) As Long
        BillTrans(BillCnt) = PayRec.BillRec
        ReDim Preserve BillDate(1 To BillCnt) As Integer
        BillDate(BillCnt) = TempRec.BillDate
      End If
    Next x
    Close PHandle
  ElseIf Exist(VATempPersBillRecs) Then
    ReDim BillTrans(1 To 1) As Long
    ReDim BillDate(1 To 1) As Integer
    BillCnt = 0
    OpenVAPersTempBillRecs TempHandle, NumOfTemps
    For x = 1 To NumOfTemps
      Get TempHandle, x, TempRec
      If TempRec.BillRec > 0 Then
        BillCnt = BillCnt + 1
        ReDim Preserve BillTrans(1 To BillCnt) As Long
        ReDim Preserve BillDate(1 To BillCnt) As Integer
        BillTrans(BillCnt) = TempRec.BillPtr
        BillDate(BillCnt) = TempRec.BillDate
        'this data should be the sane data as that where PaySeq() are loaded
      End If
    Next x
    Close TempHandle
  End If
  WhatsLeft = OldRound(CDbl(fpCurrTotRecd.Value) + CDbl(fpCurrDisc.Value))
  
  Call Distribute(WhatsLeft)
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersPaymentEntry", "cmdDist_Click", Erl)
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
  '  ClearInUse PWcnt
   ' CMTerminate

End Sub



Private Sub cmdExit_Click()
  Dim PayRec As VATaxPaymentRecType
  Dim PayHandle As Integer
  Dim NumOfPRecs As Integer
  Dim x As Integer
  Dim ThisCust As Integer
 ' Dim Handle As Integer
  
  On Local Error GoTo ERRORSTUFF
  
'  Handle = FreeFile
  If CDbl(fpLongAcctNum.Value) <> 0 Then
    If Check4Changes = True Then
      Exit Sub
    End If
  End If
  
  If DontExit = True Then 'DontExit is set in the save routine if after
  'checking for changes is true a problem is caught in the save routine
  'which would then trigger this sub because the save routine was triggered
  'from Check4Changes
    DontExit = False
    Exit Sub
  End If
  
  BillCnt = 0
  PPayEntry = False 'a global that tells frmCustomerLookup that this form is
  'where to return when frmCustomerLookup is used
  KillFile VATempPersBillRecs 'TempBillRecs is the filename for the temporary file
  'created when a bill is tagged
  ExitFlag = True
  Close
  KillFile "C:\CPWork\txperspyment.dat" 'could be used to identify this form as being opened...
  'currently (4/6/05) not being used
  Call ClearTemps
  GPayNum = 0
  GCustNum = 0
    Load frmCMPaySource
    DoEvents
    frmCMPaySource.Show
'  If Not Exist("editpyment.dat") Then
'    frmVATaxPayMenu.Show
'    DoEvents
'  Else
'    OpenTempPersPayFile PayHandle, OperNum
'    NumOfPRecs = LOF(PayHandle) / Len(PayRec)
'
''    If frmVATaxPayEditList.fpListPPay.ListCount <> NumOfPRecs Then
'      frmVATaxPayEditList.fpListPPay.Clear
'      For x = 1 To NumOfPRecs
'        Get PayHandle, x, PayRec
'        frmVATaxPayEditList.fpListPPay.InsertRow = CStr(PayRec.CustAcct) + Chr(9) + QPTrim$(PayRec.CustName) + Chr(9) + QPTrim$(Using$("$###,##0.00", PayRec.AmtPaid)) + Chr(9) + QPTrim$(Using$("$###,##0.00", PayRec.AmtOwed))
''        Debug.Print CStr(PayRec.CustAcct)
'        If PayRec.CustAcct = fpLongAcctNum.Value Then
'          frmVATaxPayEditList.fpListPPay.ListIndex = x
'        End If
'        DoEvents
'      Next x
'      Close PayHandle
''    End If
'    frmVATaxPayEditList.fpListPPay.Action = ActionForceUpdate
'
'    frmVATaxPayEditList.Show
'    DoEvents
'  End If
  
  Unload Me
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersPaymentEntry", "cmdExit_Click", Erl)
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
  '  ClearInUse PWcnt
   ' CMTerminate

End Sub

Private Sub cmdInfo_Click()
  If GCustNum = 0 Then
    Exit Sub
  End If
  
  Call frmVATaxCustInq.LoadCust
  frmVATaxCustInq.Show vbModal

End Sub

Private Sub cmdLookup_Click()
  Dim TaxRec As VATaxCustType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  PPayEntry = True
  LookUp = True '2/14/06
  frmVATaxCustLookup.Show
  DoEvents
End Sub

Private Sub cmdSave_Click()
  Dim TaxCustRec As VATaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim TaxPayRec As VATaxPaymentRecType
  Dim PayHandle As Integer
  Dim NumOfPayRecs As Long
  Dim TaxSetupRec As VATaxMasterType
  Dim MHandle As Integer
  Dim PrinceOw#
  Dim PrincePD#
  Dim InterestOw#
  Dim InterestPd#
  Dim CollectOw#
  Dim CollectPd#
  Dim TDiscAmt#
  Dim TAmtRecv#
  Dim TAmtPaid#
  Dim ChangeAmt#
  Dim Oper$, TestDate As Integer
  Dim NextListRec&
  Dim PayRecFile As Integer
  Dim NumOfRecs&
  Dim cnt&
  Dim n As Integer
  Dim Num As Integer
  Dim ThisBal As Double
  Dim SaveFlag As Integer
  Dim Message$
  Dim ThisDisc As Double
  
  On Local Error GoTo ERRORSTUFF
  WasSaved = False
  SaveMode = True
  SaveFlag = 2
  If GCustNum <= 0 Then
    fpLongAcctNum.BackColor = &H8080FF
    frmVATaxMsg.Label1.Caption = "No customer data has been supplied. No payment can be saved."
    frmVATaxMsg.Label1.Top = 800
    frmVATaxMsg.Show vbModal
    fpLongAcctNum.SetFocus
    DontExit = True
    If frmVATaxCustLookup.Visible = True Then
      Unload frmVATaxCustLookup
    End If
    fpLongAcctNum.Text = CStr(TempAcctNum)
    SaveMode = False
    Exit Sub
  End If
  TestDate = Date2Num(fptxtPayDate)
  If TestDate < 0 Then
    MsgBox "Invalid Date.", vbOKOnly, "Request Canceled."
    Exit Sub
  End If

  If Check4ValidPaidEntries = False Then  '8/12/05'checks to make sure
  'no payments are more that the amounts owed
    Exit Sub
  End If
  
  InSave = True
  If AllTaggedPaid = False Then
    If CDbl(fpCurrDisc.Value) > 0 Then
      If CDbl(fpCurrTotOwed.Value) > OldRound(CDbl(fpCurrTotRecd.Value) + CDbl(fpCurrDisc.Value)) Then
        Message = "This customer cannot receive the discount entered because the bills tagged are not being paid in full. To correct this situation you can eliminate the discount or have the customer pay in full all bills tagged."
        Call VATaxMsg(600, Message)
        If fpCurrDisc.Enabled = True Then
          fpCurrDisc.SetFocus
        End If
        InSave = False
        If frmVATaxCustLookup.Visible = True Then
          Unload frmVATaxCustLookup
        End If
        SaveMode = False
        Exit Sub
      End If
    ElseIf CDbl(fpCurrPrePay.Value) > 0 Then
      If CDbl(fpCurrTotOwed.Value) > 0 Then
        Message = "This customer cannot prepay because all bills tagged for payment are not paid in full."
        Call VATaxMsg(800, Message)
        If fpCurrPrePay.Enabled = True Then
          fpCurrPrePay.SetFocus
        End If
        InSave = False
        If frmVATaxCustLookup.Visible = True Then
          Unload frmVATaxCustLookup
        End If
        SaveMode = False
        Exit Sub
      End If
    End If
  End If
  
  InSave = False
  
  If CDbl(fpCurrPrePay.Value) > 0 And CDbl(fpCurrDisc.Value) > 0 Then
    fpCurrPrePay.BackColor = &H8080FF
    fpCurrDisc.BackColor = &H8080FF
    Call VATaxMsg(800, "Overpayment is not allowed when discounts are being used. Please eliminate the prepayment or eliminate the discount amount.")
    If fpCurrPrePay.Enabled = True Then
      fpCurrPrePay.SetFocus
    End If
    SaveMode = False
    Exit Sub
  End If
      
  If CDbl(fpCurrTotPaid.Value) = 0 Then
    fpCurrTotPaid.BackColor = &H8080FF
    frmVATaxMsg.Label1.Caption = "There is a zero value in the 'Total Amount Paid' field. Please make sure an amount paid has been entered and distributed among the 'Revenue Amount Paid' fields."
    frmVATaxMsg.Label1.Top = 700
    frmVATaxMsg.Show vbModal
    If fpcmbTenderType.Text = "CASH" Or fpcmbTenderType.Text = "CASH AND CHECK" Then
      If fpCurrCashPd.Enabled = True Then
        fpCurrCashPd.SetFocus
      End If
    Else
      If fpCurrChkChrgPd.Enabled = True Then
        fpCurrChkChrgPd.SetFocus
      End If
    End If
    DontExit = True
    If frmVATaxCustLookup.Visible = True Then
      Unload frmVATaxCustLookup
    End If
    fpLongAcctNum.Text = CStr(TempAcctNum)
    SaveMode = False
    Exit Sub
  End If
  
  If fpCurrTotRecd.Value = 0 Then
    fpCurrTotRecd.BackColor = &H8080FF
    frmVATaxMsg.Label1.Caption = "Please enter an amount paid."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
    If fpcmbTenderType.Text = "CASH" Or fpcmbTenderType.Text = "CASH AND CHECK" Then
      If fpCurrCashPd.Enabled = True Then
        fpCurrCashPd.SetFocus
      End If
    ElseIf fpcmbTenderType.Text = "CHECK" Then
      If fpCurrChkChrgPd.Enabled = True Then
        fpCurrChkChrgPd.SetFocus
      End If
    ElseIf fpcmbTenderType.Text = "CHARGE" Then
      If fpCurrChkChrgPd.Enabled = True Then
        fpCurrChkChrgPd.SetFocus
      End If
    End If
    DontExit = True
    If frmVATaxCustLookup.Visible = True Then
      Unload frmVATaxCustLookup
    End If
    fpLongAcctNum.Text = CStr(TempAcctNum)
    SaveMode = False
    Exit Sub
  End If
  
  If CDbl(fpCurrAmtOwed.Value) <> CDbl(fpCurrTotOwed.Value) Then
    fpCurrAmtOwed.BackColor = &H8080FF
    fpCurrTotOwed.BackColor = &H8080FF
    frmVATaxMsg.Label1.Caption = "The 'Tax Billing Amount Owed' and the 'Revenue Amount Owed' are not equal. Please make sure you distribute funds received by pressing F9 each time payment data has been updated."
    frmVATaxMsg.Label1.Top = 600
    frmVATaxMsg.Show vbModal
    Close
    If fpcmbTenderType.Text = "CASH" Or fpcmbTenderType.Text = "CASH AND CHECK" Then
      fpCurrCashPd.SetFocus
    ElseIf fpcmbTenderType.Text = "CHECK" Then
      fpCurrChkChrgPd.SetFocus
    ElseIf fpcmbTenderType.Text = "CHARGE" Then
      fpCurrChkChrgPd.SetFocus
    End If
    DontExit = True
    If frmVATaxCustLookup.Visible = True Then
      Unload frmVATaxCustLookup
    End If
    fpLongAcctNum.Text = CStr(TempAcctNum)
    SaveMode = False
    Exit Sub
  End If
    
  If CDbl(fpCurrAmtOwed.Value) <= CDbl(fpCurrTotRecd.Value) Then
    If CDbl(fpCurrTotWDisc.Value) <> OldRound((CDbl(fpCurrTotRecd.Value) + CDbl(fpCurrDisc.Value)) - CDbl(fpCurrChngDue.Value)) Then
      fpCurrTotPaid.BackColor = &H8080FF
      fpCurrChngDue.BackColor = &H8080FF
      fpCurrTotRecd.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The 'Total Amount Paid' does not equal the 'Total Amount Received' plus the 'Change Due'. Please make sure the funds distribution is accurate."
      frmVATaxMsg.Label1.Top = 700
      frmVATaxMsg.Show vbModal
      If fpcmbTenderType.Text = "CASH" Or fpcmbTenderType.Text = "CASH AND CHECK" Then
        If fpCurrCashPd.Enabled = True Then
          fpCurrCashPd.SetFocus
        End If
      ElseIf fpcmbTenderType.Text = "CHECK" Then
        If fpCurrChkChrgPd.Enabled = True Then
          fpCurrChkChrgPd.SetFocus
        End If
      ElseIf fpcmbTenderType.Text = "CHARGE" Then
        If fpCurrChkChrgPd.Enabled = True Then
          fpCurrChkChrgPd.SetFocus
        End If
      End If
      DontExit = True
      If frmVATaxCustLookup.Visible = True Then
        Unload frmVATaxCustLookup
      End If
      fpLongAcctNum.Text = CStr(TempAcctNum)
      SaveMode = False
      Exit Sub
    End If
  ElseIf CDbl(fpCurrAmtOwed.Value) > CDbl(fpCurrTotRecd.Value) And CDbl(fpCurrChngDue.Value) > 0 Then
    fpCurrAmtOwed.BackColor = &H8080FF
    fpCurrTotRecd.BackColor = &H8080FF
    fpCurrChngDue.BackColor = &H8080FF
    If CDbl(fpCurrChngDue.Value) <> 0 Then
      frmVATaxMsg.Label1.Caption = "ERROR: The amount owed is more than the amount received so no change should be returned. Please re-distribute your data to try to fix this issue."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      Close
      If fpcmbTenderType.Text = "CASH" Or fpcmbTenderType.Text = "CASH AND CHECK" Then
        If fpCurrCashPd.Enabled = True Then
          fpCurrCashPd.SetFocus
        End If
      ElseIf fpcmbTenderType.Text = "CHECK" Then
        If fpCurrChkChrgPd.Enabled = True Then
          fpCurrChkChrgPd.SetFocus
        End If
      ElseIf fpcmbTenderType.Text = "CHARGE" Then
        If fpCurrChkChrgPd.Enabled = True Then
          fpCurrChkChrgPd.SetFocus
        End If
      End If
      DontExit = True
      If frmVATaxCustLookup.Visible = True Then
        Unload frmVATaxCustLookup
      End If
      fpLongAcctNum.Text = CStr(TempAcctNum)
      SaveMode = False
      Exit Sub
    ElseIf CDbl(fpCurrTotPaid.Value) <> CDbl(fpCurrTotRecd.Value) Then
      fpCurrTotPaid.BackColor = &H8080FF
      fpCurrTotRecd.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The total amount received does not equal the total amount paid. Please re-distribute amounts to try and fix this issue or call Southern Software at 1-800-842-8190 for assistance."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      Close
      If fpcmbTenderType.Text = "CASH" Or fpcmbTenderType.Text = "CASH AND CHECK" Then
        If fpCurrCashPd.Enabled = True Then
          fpCurrCashPd.SetFocus
        End If
      ElseIf fpcmbTenderType.Text = "CHECK" Then
        If fpCurrChkChrgPd.Enabled = True Then
          fpCurrChkChrgPd.SetFocus
        End If
      ElseIf fpcmbTenderType.Text = "CHARGE" Then
        If fpCurrChkChrgPd.Enabled = True Then
          fpCurrChkChrgPd.SetFocus
        End If
      End If
      DontExit = True
      If frmVATaxCustLookup.Visible = True Then
        Unload frmVATaxCustLookup
      End If
      fpLongAcctNum.Text = CStr(TempAcctNum)
      SaveMode = False
      Exit Sub
    End If
  End If
   
  If CDbl(fpCurrTotPaid.Value) > CDbl(fpCurrTotOwed.Value) And (CDbl(fpCurrTotPaid.Value) = CDbl(fpCurrTotRecd.Value)) Then
    fpCurrTotPaid.BackColor = &H8080FF
    fpCurrTotOwed.BackColor = &H8080FF
    frmVATaxMsgWOpts.Label1.Caption = "The total amount paid exceeds the total amount owed. Press F10 to continue saving allowing this customer to have a credit balance. Otherwise, press ESC to review."
    frmVATaxMsgWOpts.Label1.Top = 700
    frmVATaxMsgWOpts.cmdCont.Text = "F10 Continue"
    frmVATaxMsgWOpts.cmdExit.Text = "ESC Review"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
    ElseIf frmVATaxMsgWOpts.fptxtChoice.Text = "abort" Then
      Unload frmVATaxMsgWOpts
      Close
      If fpcmbTenderType.Text = "CASH" Or fpcmbTenderType.Text = "CASH AND CHECK" Then
        If fpCurrCashPd.Enabled = True Then
          fpCurrCashPd.SetFocus
        End If
      ElseIf fpcmbTenderType.Text = "CHECK" Then
        If fpCurrChkChrgPd.Enabled = True Then
          fpCurrChkChrgPd.SetFocus
        End If
      ElseIf fpcmbTenderType.Text = "CHARGE" Then
        If fpCurrChkChrgPd.Enabled = True Then
          fpCurrChkChrgPd.SetFocus
        End If
      End If
      DontExit = True
      If frmVATaxCustLookup.Visible = True Then
        Unload frmVATaxCustLookup
      End If
      fpLongAcctNum.Text = CStr(TempAcctNum)
      SaveMode = False
      Exit Sub
    End If
  End If
  
  If OldRound(CDbl(fpCurrTotPaid.Value) + CDbl(fpCurrChngDue.Value)) <> CDbl(fpCurrTotRecd.Value) Then
    fpCurrTotPaid.BackColor = &H8080FF
    fpCurrTotRecd.BackColor = &H8080FF
    fpCurrChngDue.BackColor = &H8080FF
    frmVATaxMsg.Label1.Caption = "The total amount paid plus the total change amount do not equal the total amount received. To try to correct this issue (1) Re-distribute funds or (2) Re-select bills and then re-distribute funds." '1/25/07
    frmVATaxMsg.Label1.Top = 700
    frmVATaxMsg.Show vbModal
    Close
    If fpcmbTenderType.Text = "CASH" Or fpcmbTenderType.Text = "CASH AND CHECK" Then
      If fpCurrCashPd.Enabled = True Then
        fpCurrCashPd.SetFocus
      End If
    ElseIf fpcmbTenderType.Text = "CHECK" Then
      If fpCurrChkChrgPd.Enabled = True Then
        fpCurrChkChrgPd.SetFocus
      End If
    ElseIf fpcmbTenderType.Text = "CHARGE" Then
      If fpCurrChkChrgPd.Enabled = True Then
        fpCurrChkChrgPd.SetFocus
      End If
    End If
    DontExit = True
    If frmVATaxCustLookup.Visible = True Then 'when the customer lookup is used
    'the user can change the data on this screen by selecting a new customer...
    'this means that a check4changes must be activate from the lookup screen
    'to make sure the customer change will not leave the current customer's data
    'unsaved...if this trap catches a problem then that problem must be
    'dealt with before changing to the new customer...so if we are here by way
    'of that scenarion then the customer lookup needs unloading now
      Unload frmVATaxCustLookup
    End If
    fpLongAcctNum.Text = CStr(TempAcctNum) 'reset the customer number to the
    'what it was before the switch was attempted
    SaveMode = False
    Exit Sub
  End If
  ThisDisc = CDbl(fpCurrDisc.Value) '9/22/05
  ThisBal = 0
  If CDbl(fpCurrAmtOwed.Value) = 0 Then
    ThisBal = VAGetCustPersBalance(GCustNum, -1)
    If ThisBal > 0 And CDbl(fpCurrPrePay.Value) > 0 Then
      fpCurrAmtOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The total amount owed field is zero. This customer has an outstanding balance of " + QPTrim$(Using$("$###,##0.00", ThisBal)) + ". Customers with outstanding balances must meet these obligations before prepaying."
      frmVATaxMsg.Label1.Top = 700
      frmVATaxMsg.Show vbModal
      Close
      DontExit = True
      If frmVATaxCustLookup.Visible = True Then
        Unload frmVATaxCustLookup
      End If
      fpLongAcctNum.Text = CStr(TempAcctNum)
      SaveMode = False
      Exit Sub
    End If
  End If
  
  If CDbl(fpCurrAmtOwed.Value) < CDbl(fpCurrTotPaid.Value) Then
    ThisBal = VAGetCustPersBalance(GCustNum, -1)
    If ThisBal > CDbl(fpCurrAmtOwed.Value) Then
      fpCurrAmtOwed.BackColor = &H8080FF
      fpCurrTotPaid.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "This customer has an outstanding personal balance of " + QPTrim$(Using$("$###,##0.00", ThisBal)) + ". Customers with outstanding balances greater then the displayed amount owed, " + QPTrim$(Using$("$###,##0.00", CDbl(fpCurrAmtOwed.Value))) + ", cannot pay more than the displayed amount owed until all prior obligations have been fulfilled."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      Close
      DontExit = True
      If frmVATaxCustLookup.Visible = True Then
        Unload frmVATaxCustLookup
      End If
      fpLongAcctNum.Text = CStr(TempAcctNum)
      SaveMode = False
      Exit Sub
    End If
  End If
  
  If OldRound(CDbl(fpCurrCashPd.Value) + CDbl(fpCurrChkChrgPd.Value)) <> CDbl(fpCurrTotRecd.Value) Then
    fpCurrCashPd.BackColor = &H8080FF
    fpCurrChkChrgPd.BackColor = &H8080FF
    fpCurrTotRecd.BackColor = &H8080FF
    frmVATaxMsg.Label1.Caption = "The total amount received does not equal the total cash paid plus the total check/charges paid. Please correct this situation."
    frmVATaxMsg.Label1.Top = 700
    frmVATaxMsg.Show vbModal
    If fpcmbTenderType.Text = "CASH" Or fpcmbTenderType.Text = "CASH AND CHECK" Then
      If fpCurrCashPd.Enabled = True Then
        fpCurrCashPd.SetFocus
      End If
    ElseIf fpcmbTenderType.Text = "CHECK" Then
      If fpCurrChkChrgPd.Enabled = True Then
        fpCurrChkChrgPd.SetFocus
      End If
    ElseIf fpcmbTenderType.Text = "CHARGE" Then
      If fpCurrChkChrgPd.Enabled = True Then
        fpCurrChkChrgPd.SetFocus
      End If
    End If
    DontExit = True
    If frmVATaxCustLookup.Visible = True Then
      Unload frmVATaxCustLookup
    End If
    fpLongAcctNum.Text = CStr(TempAcctNum)
    SaveMode = False
    Exit Sub
  End If
  
  If CDbl(fpCurrTotWDisc.Value) > OldRound(CDbl(fpCurrTotRecd.Value) - CDbl(fpCurrChngDue.Value) + CDbl(fpCurrDisc.Value)) Then
    fpCurrTotRecd.BackColor = &H8080FF
    fpCurrTotWDisc.BackColor = &H8080FF
    fpCurrChngDue.BackColor = &H8080FF
    fpCurrDisc.BackColor = &H8080FF
    Message = "The total amount received (minus change due) plus discount is less than the total credited. Please increase the amount received or decrease the amount credited."
    Call VATaxMsg(700, Message)
    If fpcmbTenderType.Text = "CASH" Or fpcmbTenderType.Text = "CASH AND CHECK" Then
      If fpCurrCashPd.Enabled = True Then
        fpCurrCashPd.SetFocus
      End If
    ElseIf fpcmbTenderType.Text = "CHECK" Then
      If fpCurrChkChrgPd.Enabled = True Then
        fpCurrChkChrgPd.SetFocus
      End If
    ElseIf fpcmbTenderType.Text = "CHARGE" Then
      If fpCurrChkChrgPd.Enabled = True Then
        fpCurrChkChrgPd.SetFocus
      End If
    End If
    DontExit = True
    If frmVATaxCustLookup.Visible = True Then
      Unload frmVATaxCustLookup
    End If
    fpLongAcctNum.Text = CStr(TempAcctNum)
    SaveMode = False
    Exit Sub
  End If
  
  
'  If CheckOverPay = True Then 'this routine looks to see if the customer is trying
'  'to overpay one revenue before completely paying all revenues...this is not allowed
'    If frmVATaxCustLookup.Visible = True Then
'      Unload frmVATaxCustLookup
'    End If
'    fpLongAcctNum.Text = CStr(TempAcctNum)
'    SaveMode = False
'    Exit Sub
'  End If
  
  OpenVATaxSetUpFile MHandle
  Get MHandle, 1, TaxSetupRec
  Close MHandle
  If GetNewTot <> Val(fpCurrAmtOwed.Value) Then
      frmTaxMsgGeneral.Label2.Caption = "This Customers Balance information has changed, and is no longer valid."
      frmTaxMsgGeneral.Label3.Caption = "The screen will be cleared so you may try again."
      frmTaxMsgGeneral.Show vbModal
      SaveMode = False
      Exit Sub
  End If
'  OpenTempPersPayFile PayHandle, OperNum
'  Num = LOF(PayHandle) / Len(TaxPayRec)
'
'  If EditFlag = True Then
'    GPayNum = 0
'    For n = 1 To Num
'    Get PayHandle, n, TaxPayRec
'      If TaxPayRec.CustAcct = GCustNum Then
'        GPayNum = n
'        Exit For
'      End If
'    Next n
'  End If
  
  If fpcmbTenderType.ListIndex = 1 Or fpcmbTenderType.ListIndex = 2 Then
    frmPrintReceipt.setvallist = 1
  Else
    frmPrintReceipt.setvallist = 0
  End If

  frmPrintReceipt.Show vbModal
  If SavePay = True Then
    'do the save here
     GoSub Dothesave
     If WasSaved = True Then
        If PrnRecp = True Or PrnVali = True Then
          Call PrintReceipt
          TXLog ("Receipt printed for " + QPTrim$(fptxtName.Text) + ".")
        End If
      
        MsgBox "Transaction Complete.", vbOKOnly, "Complete"
      Else
        frmTaxMsgGeneral.Label2.Caption = "This Customers Balance information has changed, and is no longer valid."
        frmTaxMsgGeneral.Label3.Caption = "The screen will be cleared so you may try again."
        frmTaxMsgGeneral.Show vbModal
      End If
    Call Clearscreen
      TempAcctNum = 0
      DoEvents
     ' fpLongAcctNum.SetFocus
      SaveMode = False
    End If
    Oper$ = CStr(OperNum)
      KillFile ("CMXPCPR" + Oper$ + ".DAT")
      KillFile ("CMXLOP" + Oper$ + ".DAT")
  
'  If EditFlag = True And GPayNum = 0 Then
'    frmVATaxMsg.Label1.Caption = "ERROR: The program was not able to locate the customer record being edited in the save procedure. Save attempt aborted. Please call Southern Software @ 1-800-842-8190 for assistance."
'    frmVATaxMsg.Label1.Top = 800
'    frmVATaxMsg.Show vbModal
'    Close
'    SaveMode = False
'    Exit Sub
'  End If
  Exit Sub
Dothesave:
  OpenVATaxCustFile CHandle, NumOfCRecs
  Get CHandle, GCustNum, TaxCustRec
  Close CHandle
  
  If CDbl(fpCurrDisc.Value) = 0 Then Call RemoveDiscount
  
  Call UPDateListOfPayments
  OpenTempPersPayFile PayHandle, OperNum
  Num = LOF(PayHandle) / Len(TaxPayRec)

  TaxPayRec.AmtOwed = fpCurrAmtOwed.Value
  TaxPayRec.AmtPaid = OldRound(CDbl(fpCurrCashPd.Value) + CDbl(fpCurrChkChrgPd.Value))
  TaxPayRec.AmtRecd = CDbl(fpCurrTotRecd.Value)
  TaxPayRec.CashAmt = CDbl(fpCurrCashPd.Value)
  TaxPayRec.Change = CDbl(fpCurrChngDue.Value)
  If fpcmbTenderType.Text = "CHARGE" Then
    TaxPayRec.ChkAmt = 0
    TaxPayRec.ChrgAmt = CDbl(fpCurrChkChrgPd.Value)
  ElseIf fpcmbTenderType.Text = "CHECK" Then
    TaxPayRec.ChkAmt = CDbl(fpCurrChkChrgPd.Value)
    TaxPayRec.ChrgAmt = 0
  ElseIf fpcmbTenderType.Text = "CASH AND CHECK" Then
    TaxPayRec.ChkAmt = CDbl(fpCurrChkChrgPd.Value)
    TaxPayRec.ChrgAmt = 0
  Else
    TaxPayRec.ChkAmt = 0
    TaxPayRec.ChrgAmt = 0
  End If
  TaxPayRec.CustAcct = GCustNum 'CLng(fpLongAcctNum.Value)
  TaxPayRec.CustAddr = QPTrim$(fptxtAddress.Text)
  TaxPayRec.CustName = QPTrim$(fptxtName.Text)
  TaxPayRec.CustPin = TaxCustRec.PIN
  TaxPayRec.Desc = QPTrim$(fptxtDescription.Text)
  TaxPayRec.DiscAmt = CDbl(fpCurrDisc.Value)
  If BillCnt > 0 Or CDbl(fpCurrPrePay.Value) > 0 Then
    TaxPayRec.LastPayRec = LastPayRec
    TaxPayRec.NumPayRec = BillCnt
  End If
  TaxPayRec.OperNum = OperNum
  TaxPayRec.PaidOwed(1).AmtOwed = CDbl(fpCurrPersOwed.Value)
  TaxPayRec.PaidOwed(1).AmtPaid = CDbl(fpCurrPersPaid.Value)
  TaxPayRec.PaidOwed(2).AmtOwed = CDbl(fpCurrMTOwed.Value)
  TaxPayRec.PaidOwed(2).AmtPaid = CDbl(fpCurrMTPaid.Value)
  TaxPayRec.PaidOwed(3).AmtOwed = CDbl(fpCurrMCOwed.Value)
  TaxPayRec.PaidOwed(3).AmtPaid = CDbl(fpCurrMCPaid.Value)
  TaxPayRec.PaidOwed(4).AmtOwed = CDbl(fpCurrFEOwed.Value)
  TaxPayRec.PaidOwed(4).AmtPaid = CDbl(fpCurrFEPaid.Value)
  TaxPayRec.PaidOwed(5).AmtOwed = CDbl(fpCurrMHOwed.Value)
  TaxPayRec.PaidOwed(5).AmtPaid = CDbl(fpCurrMHPaid.Value)
  TaxPayRec.PaidOwed(6).AmtOwed = CDbl(fpCurrIntOwed.Value)
  TaxPayRec.PaidOwed(6).AmtPaid = CDbl(fpCurrIntPaid.Value)
  TaxPayRec.PaidOwed(7).AmtOwed = CDbl(fpCurrPenOwed.Value)
  TaxPayRec.PaidOwed(7).AmtPaid = CDbl(fpCurrPenPaid.Value)
  TaxPayRec.PaidOwed(8).AmtOwed = CDbl(fpCurrOpt1Owed.Value)
  TaxPayRec.PaidOwed(8).AmtPaid = CDbl(fpCurrOpt1Paid.Value)
  TaxPayRec.PaidOwed(9).AmtOwed = CDbl(fpCurrOpt2Owed.Value)
  TaxPayRec.PaidOwed(9).AmtPaid = CDbl(fpCurrOpt2Paid.Value)
  TaxPayRec.PaidOwed(10).AmtOwed = CDbl(fpCurrOpt3Owed.Value)
  TaxPayRec.PaidOwed(10).AmtPaid = CDbl(fpCurrOpt3Paid.Value)
  TaxPayRec.payDate = Date2Num(fptxtPayDate.Text)
  TaxPayRec.TenderTY = QPTrim$(fpcmbTenderType.Text)
  TaxPayRec.TotOwed = fpCurrAmtOwed.Value
  TaxPayRec.TotPaid = OldRound(CDbl(fpCurrPersPaid.Value) + CDbl(fpCurrMTPaid.Value) + CDbl(fpCurrMCPaid))
  TaxPayRec.TotPaid = OldRound(TaxPayRec.TotPaid + CDbl(fpCurrFEPaid.Value) + CDbl(fpCurrMHPaid.Value))
  TaxPayRec.TotPaid = OldRound(TaxPayRec.TotPaid + CDbl(fpCurrIntPaid.Value) + CDbl(fpCurrPenPaid))
  TaxPayRec.TotPaid = OldRound(TaxPayRec.TotPaid + CDbl(fpCurrOpt1Paid.Value) + CDbl(fpCurrOpt2Paid.Value))
  TaxPayRec.TotPaid = OldRound(TaxPayRec.TotPaid + CDbl(fpCurrOpt3Paid.Value)) + CDbl(fpCurrPrePay.Value)
  TaxPayRec.PrePayAmt = CDbl(fpCurrPrePay.Value)
  TaxPayRec.CustPin = TaxCustRec.PIN
  TaxPayRec.BillType = "P"
    Put PayHandle, 1, TaxPayRec
  
'  KillFile VATempPersBillRecs 'get rid of all temporary files and records in
'  'preparation for the next customer
'  BillCnt = 0
'  ReDim BillTrans(0 To 0) As Long
'
'  Call LoadTemps 'save new temps in case a new save takes place for the
'  'same customer
'  If CLng(fpLongAcctNum.Value) = GCustNum Then
'    EditFlag = True
'  End If
  
  Close PayHandle
  
  DontExit = False
 ' Call VASavemsg(900, "This personal tax payment has been saved successfully.")
  PostEmTax
  Return
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersPaymentEntry", "cmdSave_Click", Erl)
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
'    ClearInUse PWcnt
 '   CMTerminate

End Sub

Private Sub Form_Click()
  Call MakeEmWhite
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Call MakeEmWhite ' done to white out any fields that were reddened when
  'errors were flagged
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      If BillHasFocus = True Then Call cmdBills_Click
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      'SendKeys "%E"
      Call cmdExit_Click
      KeyCode = 0
    Case vbKeyF9:
      'SendKeys "%D"
      Call cmdDist_Click
      KeyCode = 0
    Case vbKeyF10:
      'SendKeys "%S"
      Call cmdSave_Click
      KeyCode = 0
    Case vbKeyF5:
     ' SendKeys "%C"
      Call cmdCash_Click
      KeyCode = 0
    Case vbKeyF6:
     ' SendKeys "%k"
      Call cmdCheck_Click
      KeyCode = 0
    Case vbKeyF8:
     ' SendKeys "%g"
      Call cmdCharge_Click
      KeyCode = 0
    Case vbKeyF3:
      SendKeys "{Tab}"
      DoEvents
      Call cmdBills_Click
      KeyCode = 0
    Case vbKeyF2:
     ' SendKeys "%w"
      Call fpcmdDrawer_Click
      KeyCode = 0
    Case vbKeyF4:
      'SendKeys "%I"
      Call cmdInfo_Click
      KeyCode = 0
    Case vbKeyF7:
      'SendKeys "%L"
      Call cmdLookup_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  Call MakeEmWhite
End Sub

Private Sub Form_Load()
  ScrWidth = Screen.Width / Screen.TwipsPerPixelX
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ScreenW = (Screen.Width / Screen.TwipsPerPixelX)
  NotFirstLoad = False
  StopWarn = False
  InClear = False
  InOverRideDist = False
  InSave = False
  SaveMode = False
  Call LoadMe
  Call GetRcpInfo
  If InStr(TownName$, "LUNENBURG") Then
    Lunney = True
  Else
    Lunney = False
  End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Call MakeEmWhite
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      PPayEntry = False
      KillFile "c:\CPWork\editpyment.dat"
      KillFile "C:\CPWork\txperspyment.dat"
      'TxLog ("CitiTaxes.exe terminated via menu bar on frmVATaxPersPaymentEntry.")
      Call CMTerminate
      End
    End If
  End If

End Sub
Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
   ''' Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
    DoEvents
  End If
End Sub

Private Sub fpcmbTenderType_Change()
  If QPTrim$(fpcmbTenderType.Text) = "" Then
    fpcmbTenderType.Text = "CASH"
    fpCurrChkChrgPd.Enabled = False
    fpCurrChkChrgPd.Value = 0
  End If
  
  If QPTrim$(fpcmbTenderType.Text) = "CASH" Then
    fpCurrChkChrgPd.Enabled = False
    fpCurrChkChrgPd.Value = 0
    fpCurrCashPd.Enabled = True
  ElseIf QPTrim$(fpcmbTenderType.Text) = "CHECK" Then
    fpCurrChkChrgPd.Enabled = True
    fpCurrCashPd.Value = 0
    fpCurrCashPd.Enabled = False
  ElseIf QPTrim$(fpcmbTenderType.Text) = "CHARGE" Then
    fpCurrChkChrgPd.Enabled = True
    fpCurrCashPd.Value = 0
    fpCurrCashPd.Enabled = False
  ElseIf QPTrim$(fpcmbTenderType.Text) = "CASH AND CHECK" Then
    fpCurrChkChrgPd.Enabled = True
    fpCurrCashPd.Enabled = True
  End If
  Call ClearPaidFields
  ReFigure

End Sub

Private Sub fpcmbTenderType_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeySpace Then
    fpcmbTenderType.ListDown = True
  End If
  If KeyCode = vbKeyDelete Then
    fpcmbTenderType.ListIndex = -1
  End If
  If fpcmbTenderType.ListDown <> True Then
    If KeyCode = vbKeyDown Then
      If fpCurrCashPd.Enabled = True Then
        fpCurrCashPd.SetFocus
      ElseIf fpCurrChkChrgPd.Enabled = True Then
        fpCurrChkChrgPd.SetFocus
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
Private Sub fpcmdDrawer_Click()
  Dim Port As String, PortFile As Integer ', DPName As String, DefPrinter As String
  On Local Error Resume Next
  If RecpDef = 99 Then Exit Sub
  Port$ = QPTrim$(RecpPort)
  CMLog "Oper: " + Str(OperNum) + "CMTax Pay-Open Drawer"
  TXLog "Oper: " + Str(OperNum) + "CMTax Pay-Open Drawer"
  PortFile = FreeFile
  Open Port$ For Output As #PortFile
  Print #PortFile, Chr$(27); "p"; Chr$(0); Chr$(25); Chr$(250)
  Print #PortFile, Chr$(7)
  Close PortFile
End Sub

Private Sub fpCurrMCPaid_LostFocus()
  If fpCurrMCPaid.ControlType = ControlTypeReadOnly Then Exit Sub '1/25/07
  If NotFirstLoad = False Then Exit Sub
  If InOverRideDist = True Then Exit Sub
  
  If TempMCPaid <> CDbl(fpCurrMCPaid.Value) Then
    Call OverRideDist
  End If
 
  Call AddUpPaidCol
  Call LoadTempPayList
End Sub

Private Sub fpCurrCashPd_LostFocus()
  Call ReFigure

End Sub

Private Sub fpCurrChkChrgPd_LostFocus()
  On Local Error GoTo ERRORSTUFF
  
  Call ReFigure
  If StopWarn = True Then Exit Sub
  StopWarn = True
  If EditFlag = False Then
    If fpCurrAmtOwed.DoubleValue = 0 Then
      If fpCurrTotRecd.DoubleValue > 0 Then
        fpCurrAmtOwed.BackColor = &H8080FF
        frmVATaxMsg.Label1.Caption = "Automatic amount distribution does not take place if the Amount Due equals $0.00. Please make sure the appropriate amounts are entered manually."
        frmVATaxMsg.Label1.Top = 700
        frmVATaxMsg.Show vbModal
        If fpcmbTenderType = "CASH" Then
          If fpCurrCashPd.Enabled = True Then
            fpCurrCashPd.BackColor = &H8080FF
            fpCurrCashPd.SetFocus
          End If
        ElseIf fpcmbTenderType = "CASH AND CHECK" Then
          If fpCurrCashPd.Enabled = True Then
            fpCurrCashPd.BackColor = &H8080FF
            fpCurrChkChrgPd.BackColor = &H8080FF
            fpCurrCashPd.SetFocus
          End If
        Else
          If fpCurrChkChrgPd.Enabled = True Then
            fpCurrChkChrgPd.BackColor = &H8080FF
            fpCurrChkChrgPd.SetFocus
          End If
        End If
        TXLog ("The amount due is zero for this customer but a value has been entered for amount received. The user was warned to make sure the appropriate amounts were manually entered.")
      End If
    End If
  End If
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersPaymentEntry", "fpCurrChkChrgPd_LostFocus", Erl)
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
'    ClearInUse PWcnt
'    CMTerminate


End Sub

Private Sub fpCurrDisc_Click(Button As Integer)
  If NotFirstLoad = False Then Exit Sub
  If MaxDisc = 0 Then
    Call VATaxMsg(900, "This customer is not eligible for a discount")
    fpCurrDisc.ControlType = ControlTypeReadOnly
  ElseIf MaxDisc > 0 Then
    fpCurrDisc.ControlType = ControlTypeNormal
  End If
End Sub
Private Sub fpCurrDisc_Change() '1/25/07
  If fpCurrDisc.Value > 0 Then
    fpCurrMTPaid.ControlType = ControlTypeReadOnly
    fpCurrMCPaid.ControlType = ControlTypeReadOnly
    fpCurrMHPaid.ControlType = ControlTypeReadOnly
    fpCurrFEPaid.ControlType = ControlTypeReadOnly
    fpCurrPersPaid.ControlType = ControlTypeReadOnly
    fpCurrIntPaid.ControlType = ControlTypeReadOnly
    fpCurrPenPaid.ControlType = ControlTypeReadOnly
    fptxtOpt1.ControlType = ControlTypeReadOnly
    fptxtOpt2.ControlType = ControlTypeReadOnly
    fptxtOpt3.ControlType = ControlTypeReadOnly
    fpCurrPrePay.Value = 0
    fpCurrPrePay.ControlType = ControlTypeReadOnly
  Else
    fpCurrMTPaid.ControlType = ControlTypeNormal
    fpCurrMCPaid.ControlType = ControlTypeNormal
    fpCurrMHPaid.ControlType = ControlTypeNormal
    fpCurrFEPaid.ControlType = ControlTypeNormal
    fpCurrPersPaid.ControlType = ControlTypeNormal
    fpCurrIntPaid.ControlType = ControlTypeNormal
    fpCurrPenPaid.ControlType = ControlTypeNormal
    fptxtOpt1.ControlType = ControlTypeNormal
    fptxtOpt2.ControlType = ControlTypeNormal
    fptxtOpt3.ControlType = ControlTypeNormal
    fpCurrPrePay.ControlType = ControlTypeNormal
  End If
End Sub

Private Sub fpCurrDisc_LostFocus()
  Dim ThisAmt As Double
  
  On Local Error GoTo ERRORSTUFF
  
  If CDbl(fpCurrTotOwed.Value) = 0 Then Exit Sub
  
  If CDbl(fpCurrDisc.Value) > MaxDisc Then
    Call VATaxMsg(800, "The maximum personal discount allowed for this customer is " + QPTrim$(Using$("$##,##0.00", MaxDisc)) + ". The program will reset the discount to the maximum allowed.")
    fpCurrDisc = MaxDisc
  ElseIf fpCurrDisc.Value = 0 Then 'added 1/25/07
    fpCurrMTPaid = fpCurrMTOwed
    fpCurrMCPaid = fpCurrMCOwed
    fpCurrMHPaid = fpCurrMHOwed
    fpCurrFEPaid = fpCurrFEOwed
    fpCurrPersPaid = fpCurrPersOwed
    fpCurrPenPaid = fpCurrPenOwed
    fpCurrOpt1Paid = fpCurrOpt1Owed
    fpCurrOpt2Paid = fpCurrOpt2Owed
    fpCurrOpt3Paid = fpCurrOpt3Owed
  ElseIf CDbl(fpCurrDisc.Value) < MaxDisc Then
    If BillCnt > 0 Then 'added 7/18/07*****************************************
      Call ReassignDiscount
      Call Distribute(OldRound(CDbl(fpCurrTotRecd.Value) + CDbl(fpCurrDisc.Value)))
    End If
  End If
  
  fpCurrTotPaid = AddUpPaidCol
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersPaymentEntry", "fpCurrDisc_LostFocus", Erl)
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
 '   ClearInUse PWcnt
 '   CMTerminate

End Sub

Private Sub fpCurrMTPaid_LostFocus()
  If fpCurrMTPaid.ControlType = ControlTypeReadOnly Then Exit Sub '1/25/07
  If NotFirstLoad = False Then Exit Sub
  If InOverRideDist = True Then Exit Sub
  
  If TempMTPaid <> CDbl(fpCurrMTPaid.Value) Then
    Call OverRideDist
  End If
 
  Call AddUpPaidCol
  Call LoadTempPayList
  
End Sub

Private Sub fpCurrFEPaid_LostFocus()
  If fpCurrFEPaid.ControlType = ControlTypeReadOnly Then Exit Sub '1/25/07
  If NotFirstLoad = False Then Exit Sub
  If InOverRideDist = True Then Exit Sub
  
  If TempFEPaid <> CDbl(fpCurrFEPaid.Value) Then
    Call OverRideDist
  End If
 
  Call AddUpPaidCol
  Call LoadTempPayList
End Sub

Private Sub fpCurrOpt1Paid_LostFocus()
    If fpCurrOpt1Paid.ControlType = ControlTypeReadOnly Then Exit Sub '1/25/07
If NotFirstLoad = False Then Exit Sub
  If InOverRideDist = True Then Exit Sub
  If TempOpt1Paid <> CDbl(fpCurrOpt1Paid.Value) Then
    Call OverRideDist
  End If
 
  Call AddUpPaidCol
  Call LoadTempPayList

End Sub

Private Sub fpCurrOpt2Paid_LostFocus()
   If fpCurrOpt2Paid.ControlType = ControlTypeReadOnly Then Exit Sub '1/25/07
 If NotFirstLoad = False Then Exit Sub
  If InOverRideDist = True Then Exit Sub
  If TempOpt2Paid <> CDbl(fpCurrOpt2Paid.Value) Then
    Call OverRideDist
  End If
 
  Call AddUpPaidCol
  Call LoadTempPayList

End Sub

Private Sub fpCurrOpt3Paid_LostFocus()
    If fpCurrOpt3Paid.ControlType = ControlTypeReadOnly Then Exit Sub '1/25/07
If NotFirstLoad = False Then Exit Sub
  If InOverRideDist = True Then Exit Sub
  If TempOpt3Paid <> CDbl(fpCurrOpt3Paid.Value) Then
    Call OverRideDist
  End If
 
  Call AddUpPaidCol
  Call LoadTempPayList

End Sub

Private Sub fpCurrPrePay_LostFocus()
  On Local Error GoTo ERRORSTUFF
  
  If SaveMode = True Then Exit Sub
  If QPTrim$(fptxtName.Text) <> "" Then NotFirstLoad = True 'added 6/1/06
  If CDbl(fpCurrPersPaid.Value) = 0 Then
    If CDbl(fpCurrPrePay.Value) > 0 And CDbl(fpCurrTotRecd.Value) = 0 Then
      fpCurrPrePay = 0
      Call VATaxMsg(900, "No payment has been entered. Prepayment will be reset to zero.")
      If fpcmbTenderType.Text = "CASH" Or fpcmbTenderType.Text = "CASH AND CHECK" Then
        If fpCurrCashPd.Enabled = True Then
          fpCurrCashPd.SetFocus
        End If
      Else
        If fpCurrChkChrgPd.Enabled = True Then
          fpCurrChkChrgPd.SetFocus
        End If
      End If
      Exit Sub
    End If
  End If
    
  If CDbl(fpCurrPrePay.Value) > 0 And CDbl(fpCurrAmtOwed.Value) <> 0 Then
    If CDbl(fpCurrTotWDisc.Value) < CDbl(fpCurrAmtOwed.Value) Then
      fpCurrPrePay.Value = 0
      Call VATaxMsg(700, "Prepayment amounts can only be added if the total amounts owed are paid in full. Applying discounts also prevent prepayments.")
      Call ReLoadPaidTemps
    End If
  End If
  Call AddUpPaidCol
  Call LoadTempPayList
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersPaymentEntry", "fpCurrPrePay_LostFocus", Erl)
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
 '   ClearInUse PWcnt
 '   CMTerminate

End Sub

Private Sub fpCurrPersPaid_LostFocus()
  If fpCurrPersPaid.ControlType = ControlTypeReadOnly Then Exit Sub '1/25/07
  If NotFirstLoad = False Then Exit Sub
  If InOverRideDist = True Then Exit Sub
  If TempPersPaid > CDbl(fpCurrPersPaid.Value) Then
    Call OverRideDist
  End If
 
  Call AddUpPaidCol
  Call LoadTempPayList
End Sub

Private Sub fpCurrMHPaid_LostFocus()
  If fpCurrMHPaid.ControlType = ControlTypeReadOnly Then Exit Sub '1/25/07
  If NotFirstLoad = False Then Exit Sub
  If InOverRideDist = True Then Exit Sub
  If TempMHPaid > CDbl(fpCurrMHPaid.Value) Then
    Call OverRideDist
  End If
  
  Call AddUpPaidCol
  Call LoadTempPayList
End Sub

Private Sub fpCurrIntPaid_LostFocus()
  If fpCurrIntPaid.ControlType = ControlTypeReadOnly Then Exit Sub '1/25/07
  If NotFirstLoad = False Then Exit Sub
  If InOverRideDist = True Then Exit Sub
  If TempIntPaid > CDbl(fpCurrIntPaid.Value) Then
    Call OverRideDist
  End If
 
  Call AddUpPaidCol
  Call LoadTempPayList
End Sub

Private Sub fpCurrPenPaid_LostFocus()
  If fpCurrPenPaid.ControlType = ControlTypeReadOnly Then Exit Sub '1/25/07
  If NotFirstLoad = False Then Exit Sub
  If InOverRideDist = True Then Exit Sub
  If TempPenPaid <> CDbl(fpCurrPenPaid.Value) Then
    Call OverRideDist
  End If
 
  Call AddUpPaidCol
  Call LoadTempPayList
End Sub

Private Sub fpCurrTotPaid_Change()
  Call ReFigure
End Sub

Private Sub fpLongAcctNum_LostFocus()
  Dim ThisAcctNum As Long
  
  On Local Error GoTo ERRORSTUFF
  
  If TempAcctNum = CLng(fpLongAcctNum.Value) Then Exit Sub
  
  'If frmVATaxPayMenu.Visible = True Then Exit Sub
  If fpLongAcctNum.Value = 0 And BillCnt = 0 Then Exit Sub
  
  If fpLongAcctNum.Value = 0 Then
    Call Clearscreen
    Exit Sub
  End If
  
  If fpLongAcctNum.Value <> GCustNum Then
    StopWarn = False 'used to keep the warning popup regarding the need for
    'manual distribution if the customer owes nothing and prepays
    NotFirstLoad = False
    GetNewCust = True 'used to keep the _Change sub from being used needlessly
  Else
    If fpLongAcctNum.Value = GCustNum Then
      Exit Sub
    End If
  End If
  
  If VACheck4ValidCustNum(fpLongAcctNum.Value) = False Then
    fpLongAcctNum.BackColor = &H8080FF
    frmVATaxMsg.Label1.Caption = "The customer number is not valid. Please enter a valid customer number."
    frmVATaxMsg.Label1.Top = 800
    frmVATaxMsg.Show vbModal
    GetNewCust = False
    Call Clearscreen
    fpLongAcctNum.SetFocus
    Exit Sub
  End If

  ThisAcctNum = fpLongAcctNum.Value
  If NotFirstLoad = True Then 'screen is already loaded
    If TempAcctNum = 0 Then
      If CDbl(fpLongAcctNum.Value) = 0 Then
        Exit Sub 'no need to proceed if the acct num is 0 or unchanged
      ElseIf GCustNum = CDbl(fpLongAcctNum.Value) Then
        Exit Sub
      Else
        Call Check4Changes
        Call Clearscreen
        fpLongAcctNum = ThisAcctNum
        Call VAGetCust
      End If
    Else
      If TempAcctNum > 0 And TempAcctNum <> fpLongAcctNum.Value Then 'new acct num entered
        If Check4Changes = True Then Exit Sub 'stop to allow Check4Changes to handle
        'any changes
        If DontExit = True Then Exit Sub 'DontExit is set in the save routine if a trap caught
        'something
        Call Clearscreen 'All is OK so reset the data on the screen
      End If
      GetNewCust = True
      fpLongAcctNum = ThisAcctNum 'start the loading process for the new number
      Call LostFocusCheck
    End If
  Else 'brand new screen being loaded
    If TempAcctNum > 0 And TempAcctNum <> fpLongAcctNum.Value Then
      If Check4Changes = True Then Exit Sub
      Call Clearscreen
    End If
    GetNewCust = True
    fpLongAcctNum = ThisAcctNum
    Call LostFocusCheck
  End If
  
  If fpLongAcctNum.Value <> 0 Then 'we are loaded now so turn off the load flag
    NotFirstLoad = True
  End If
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersPaymentEntry", "fpLongAcctNum_LostFocus", Erl)
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
 '   ClearInUse PWcnt
 '   CMTerminate

End Sub

Public Sub Clearscreen()
  InClear = True
  NotFirstLoad = False
  Label25.Visible = False
  fpLongAcctNum.Value = 0
  fptxtPayDate = DefPayDate
  fptxtName.Text = ""
  fptxtAddress.Text = ""
  fptxtCity.Text = ""
  fptxtState.Text = ""
  fptxtZip.Text = ""
  fpCurrAmtOwed.Value = 0
  fpcmbTenderType.Text = "CASH"
  fpCurrCashPd.Value = 0
  fpCurrChkChrgPd.Value = 0
  fpCurrTotRecd.Value = 0
  fpCurrChngDue.Value = 0
  fpCurrDisc.Value = 0
  fpCurrPersOwed.Value = 0
  fpCurrPersPaid.Value = 0
  fpCurrMTOwed.Value = 0
  fpCurrMTPaid.Value = 0
  fpCurrMCOwed.Value = 0
  fpCurrMCPaid.Value = 0
  fpCurrFEOwed.Value = 0
  fpCurrFEPaid.Value = 0
  fpCurrMHOwed.Value = 0
  fpCurrMHPaid.Value = 0
  fpCurrIntOwed.Value = 0
  fpCurrIntPaid.Value = 0
  fpCurrPenOwed.Value = 0
  fpCurrPenPaid.Value = 0
  fpCurrOpt1Owed.Value = 0
  fpCurrOpt1Paid.Value = 0
  fpCurrOpt2Owed.Value = 0
  fpCurrOpt2Paid.Value = 0
  fpCurrOpt3Owed.Value = 0
  fpCurrOpt3Paid.Value = 0
  fpCurrTotOwed.Value = 0
  fpCurrTotPaid.Value = 0
  fpCurrPrePay.Value = 0
  fpCurrTotWDisc.Value = 0
  fptxtDescription.Text = ""
  NotFirstLoad = False
  ThisDiscAmt = 0
  Call LoadTemps
  Call AssignPaidTemps
  InClear = False
  BillCnt = 0
  GCustNum = 0
  GPayNum = 0
  ReDim TempBillList(1 To 1) As VAPersPayListType
  TempBillListCnt = 0
  BegAmount = 0 'added 7/18/07******************************************
  fpLongAcctNum.SetFocus
End Sub

Private Sub LoadMe()
  Dim TaxMasterRec As VATaxMasterType
  Dim TMHandle As Integer
  Dim x As Integer
  Dim One As Integer
  Dim ThisFile As Integer
  Dim FileName$
  
  On Local Error GoTo ERRORSTUFF
  LookUp = False '2/14/06
  ThisBillType = "P"
  BillHasFocus = False
  DiscYN = False
  ClearTemps
  OverPay = False
  MaxDisc = 0
  FileName = "C:\CPWork\txperspyment.dat" 'used when using the transaction history report
  ThisFile = FreeFile
  Open FileName For Output As ThisFile
  One = 1
  Print #ThisFile, One
  Close ThisFile
  ReDim TempBillList(1 To 1) As VAPersPayListType
  TempBillListCnt = 0
  
  fptxtPayDate.Text = DefPayDate
  ThisDiscAmt = 0 'reset this global for new customer
  DiscPXDate = TaxMasterRec.DiscPXDate
  DiscPXDate = Date2Num(fptxtPayDate.Text) 'corrected 9/20/05 ...was 'Date' instead of fptxtPayDate.text
'  DiscXDate = DiscXDate + 1 'remarked 9/20/05
  
  OpenVATaxSetUpFile TMHandle
  Get TMHandle, 1, TaxMasterRec
  Close TMHandle
  
  ReDim PayOrder(1 To 10) As Integer
  PayOrder(1) = TaxMasterRec.PersPayOrder
  PayOrder(2) = TaxMasterRec.MTPayOrder
  PayOrder(3) = TaxMasterRec.MCPayOrder
  PayOrder(4) = TaxMasterRec.FEPayOrder
  PayOrder(5) = TaxMasterRec.MHPayOrder
  PayOrder(6) = TaxMasterRec.PIntPayOrder
  PayOrder(7) = TaxMasterRec.PPenPayOrder
  PayOrder(8) = TaxMasterRec.POpt1PayOrder
  PayOrder(9) = TaxMasterRec.POpt2PayOrder
  PayOrder(10) = TaxMasterRec.POpt3PayOrder
  
  If QPTrim$(TaxMasterRec.POptRev1) = "" Then
    fptxtOpt1.Text = "NOT IN USE"
    fpCurrOpt1Paid.Enabled = False
  Else
    fptxtOpt1.Text = QPTrim$(TaxMasterRec.POptRev1)
  End If
  If QPTrim$(TaxMasterRec.POptRev2) = "" Then
    fptxtOpt2.Text = "NOT IN USE"
    fpCurrOpt2Paid.Enabled = False
  Else
    fptxtOpt2.Text = QPTrim$(TaxMasterRec.POptRev2)
  End If
  If QPTrim$(TaxMasterRec.POptRev3) = "" Then
    fptxtOpt3.Text = "NOT IN USE"
    fpCurrOpt3Paid.Enabled = False
  Else
    fptxtOpt3.Text = QPTrim$(TaxMasterRec.POptRev3)
  End If
  
  lblCurrTaxYr.Caption = "Current Tax Year: " + CStr(TaxMasterRec.RTaxYear)
  TempPreYear = TaxMasterRec.RTaxYear
  fpcmbTenderType.Clear
  fpcmbTenderType.Text = "CASH"
  fpcmbTenderType.AddItem "CASH"
  fpcmbTenderType.AddItem "CHECK"
  fpcmbTenderType.AddItem "CASH AND CHECK"
  fpcmbTenderType.AddItem "CHARGE"
  fpCurrChkChrgPd.Enabled = False 'this field is disabled when 'CASH' is
  'displayed on the tender type drop down
  EditFlag = False
  ExitFlag = False
'  OPERNUM = 1 'used for testing only
  lblOperNum.Caption = "Operator Number: " + CStr(OperNum)
  lblOperName.Caption = "Operator Name: " + PWUser
  ThisDiscAmt = 0
  GetNewCust = False
  NotFirstLoad = True
  If GCustNum > 0 Then 'coming from edit lookup
    Call LoadHerUpEdit
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersPaymentEntry", "LoadMe", Erl)
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
'    ClearInUse PWcnt
'    CMTerminate

End Sub

Private Function VACheck4ValidCustNum(ThisCust As Long) As Boolean
  Dim TaxRec As VATaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim x As Long
  Dim Number$
  Dim Name$
  Dim Found As Boolean
  
  On Local Error GoTo ERRORSTUFF
  
  VACheck4ValidCustNum = True
  
  If fpLongAcctNum.Value = 0 Then
    VACheck4ValidCustNum = False
    Exit Function
  End If
  
  OpenVATaxCustFile CHandle, NumOfCRecs
  
  If NumOfCRecs = 0 Then
    frmVATaxMsg.Label1.Caption = "There are no tax customers saved."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
    Close CHandle
    Exit Function
  End If
  
  For x = 1 To NumOfCRecs
    Get CHandle, x, TaxRec
    If ThisCust = TaxRec.Acct Then
      If TaxRec.Deleted <> 0 Then
        VACheck4ValidCustNum = False
      End If
      Exit For
    End If
  Next x

  Close CHandle

  If x > NumOfCRecs Then
    Call Clearscreen
    VACheck4ValidCustNum = False
  End If
  
  Exit Function

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersPaymentEntry", "VACheck4ValidCustNum", Erl)
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
 '   ClearInUse PWcnt
 '   CMTerminate

End Function

Public Sub VAGetCust()
  Dim TaxRec As VATaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim x As Integer
  Dim Number As Long
  Dim Name$
  Dim Found As Boolean
  
  On Local Error Resume Next
  If fpLongAcctNum.Value = 0 Then
    fpLongAcctNum.BackColor = &H8080FF
    frmVATaxMsg.Label1.Caption = "Please enter a customer account number."
    frmVATaxMsg.Label1.Top = 800
    frmVATaxMsg.Show vbModal
    If fpLongAcctNum.Enabled = True Then
      fpLongAcctNum.SetFocus
    End If
    Exit Sub
  End If
  
  Number = fpLongAcctNum.Value
  
  OpenVATaxCustFile CHandle, NumOfCRecs
  
  If NumOfCRecs = 0 Then
    frmVATaxMsg.Label1.Caption = "There are no tax customers saved."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
    Close
    Exit Sub
  End If
  
  For x = 1 To NumOfCRecs
    Get CHandle, x, TaxRec
    If Number = TaxRec.Acct Then 'match the selected
    'row with the right code
      Found = True
      GCustNum = x 'now you can assign the correct global
      Exit For
    Else
      Found = False
      GoTo NotAMatch
    End If
      
NotAMatch:
  Next x
  
  Close CHandle
  
  If Found = True Then
    If QPTrim$(fptxtName.Text) <> "" And GetNewCust = False Then
      Call Clearscreen
    End If
    KillFile VATempPersBillRecs 'get rid of all temporary files and records in
    'preparation for the this customer
    BillCnt = 0
    ReDim BillTrans(0 To 0) As Long
    Call EnterEditChk
  End If
  
  GetNewCust = False
  
End Sub

Public Sub EnterEditChk()
  Dim ONum$
  Dim ThisRec As Integer
  Dim CustNum As Long
  Dim FindStr As String
  On Local Error GoTo ERRORSTUFF
  
  'in conjunction with BegBalCheck this set of code determines the
  'current status of the customer the user is attempting to bring up
  'on the screen
  ONum = OperNum
  ThisRec = 0
  CustNum = GCustNum
  If CustNum > 0 Then
    If CustHasMsg(GCustNum) Then
      MsgAlertTimer.Enabled = True
    Else
      MsgAlertTimer.Enabled = False
      cmdInfo.ForeColor = &H80000012
    End If
  End If
  
  Select Case VABegBalCheck(CustNum, ONum$, ThisRec, ThisBillType)
    Case 1 'normal first time transaction for this customer
      EditFlag = False
      TempAcctNum = CustNum
      Call LoadHerUpWOEdit
      FindStr = FindVACustInBatchFile(CustNum, "P")
      If FindStr <> "0" Then
         frmVATaxInBatchList.ListStr = FindStr
         frmVATaxInBatchList.Show vbModal
         TXLog ("User informed this customer, " + CStr(CustNum) + ", is included in            the following unposted batch files: " + FindStr + ".")
      End If
      Exit Sub
    Case 2 'edit a transaction that is in progress
      EditFlag = True
      TempAcctNum = CustNum
      GPayNum = ThisRec
      Call LoadHerUpEdit
      NotFirstLoad = True
      If VAGetCustPersBalance(GCustNum, -1) < 0 Then
        Call DisablePayFields
      Else
        Call EnablePayFields
      End If
      Exit Sub
    Case 4 'a transaction for this customer is already in progress
    'so abort this attempt
      GCustNum = 0
      TempAcctNum = 0
      Call Clearscreen
      EditFlag = False
      Call LoadMe
      Exit Sub
    Case Else
      frmVATaxMsg.Label1.Caption = "Error: This customer's data could not be retrieved."
      frmVATaxMsg.Label1.Top = 700
      frmVATaxMsg.Show vbModal
      Close
      Exit Sub
  End Select
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersPaymentEntry", "EnterEditChk", Erl)
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
 '   ClearInUse PWcnt
 '   CMTerminate

End Sub

Private Sub LoadHerUpWOEdit()
  Dim TaxRec As VATaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim ThisBalance#
  
  On Local Error GoTo ERRORSTUFF
  
  KillFile VATempPersBillRecs
  NotFirstLoad = False
  OpenVATaxCustFile CHandle, NumOfCRecs
  Get CHandle, GCustNum, TaxRec
  Close CHandle
  
  Label25.Visible = False

  DiscYN = False
  ThisBalance = VAGetCustPersBalance(GCustNum, -1)
  If ThisBalance = 0 Then
    frmVATaxMsg.Label1.Caption = "This customer has a zero balance."
    frmVATaxMsg.Label1.Top = 900
    frmVATaxMsg.Show vbModal
    Call DisablePayFields
  ElseIf ThisBalance < 0 Then
    Call VATaxMsg(900, "This customer has a personal balance of -" + QPTrim$(Using$("$##,##0.00", Abs(ThisBalance))) + ".")
    Label25.Visible = True
    Label25.Caption = "This customer has a personal balance of -" + QPTrim$(Using$("$##,##0.00", Abs(ThisBalance))) + "."
    Call DisablePayFields
  Else
    Call EnablePayFields
  End If
  
  TempAcctNum = TaxRec.Acct
  fptxtName.Text = QPTrim$(TaxRec.CustName)
  If QPTrim$(TaxRec.Addr1) <> "" Then
    fptxtAddress.Text = QPTrim$(TaxRec.Addr1)
  Else
    fptxtAddress.Text = QPTrim$(TaxRec.Addr2)
  End If
  fptxtCity.Text = QPTrim$(TaxRec.City)
  fptxtState.Text = QPTrim$(TaxRec.State)
  fptxtZip = QPTrim$(TaxRec.Zip)
  
  ThisDiscAmt = 0
  GetNewCust = False
  Call LoadTemps
  EditFlag = False
'  NotFirstLoad = True
'  cmdBills.Enabled = True

  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersPaymentEntry", "LoadHerUpWOEdit", Erl)
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
 '   ClearInUse PWcnt
 '   CMTerminate

End Sub

Private Sub LostFocusCheck()
  On Local Error GoTo ERRORSTUFF
  
  If fpLongAcctNum.Value = 0 Then
    Call Clearscreen
    Exit Sub
  End If
  
  If ExitFlag = False Then
    Call VAGetCust
  End If
  
  If GCustNum = 0 Then
    Call Clearscreen
    Exit Sub
  End If
  
  If EditFlag = False And GCustNum <> fpLongAcctNum.Value Then
    Call LoadHerUpWOEdit
  End If
  
  NotFirstLoad = True
  
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersPaymentEntry", "LostFocusCheck", Erl)
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
 '   ClearInUse PWcnt
 '   CMTerminate


End Sub
Private Sub fpCurrCashPd_Change()
  Dim TotAmt As Double
  Dim Cash As Double
  Dim Chrg As Double
  Dim Disc As Double
  
  Disc = CDbl(fpCurrDisc.Value)
  Cash = CDbl(fpCurrCashPd.Value)
  Chrg = CDbl(fpCurrChkChrgPd.Value)
  TotAmt = OldRound(Cash + Chrg)
  fpCurrTotRecd.Value = TotAmt
End Sub

Private Sub fpCurrChkChrgPd_Change()
  Dim TotAmt As Double
  Dim Cash As Double
  Dim Chrg As Double
  
  Cash = CDbl(fpCurrCashPd.Value)
  Chrg = CDbl(fpCurrChkChrgPd.Value)
  TotAmt = OldRound(Cash + Chrg)
  fpCurrTotRecd.Value = TotAmt

End Sub

Private Sub LoadAmtOwed()
  Dim x As Integer
  Dim PersOwed#
  Dim MTOwed#
  Dim MCOwed#
  Dim FEOwed#
  Dim MHOwed#
  Dim IntOwed#
  Dim PenOwed#
  Dim Opt1Owed#
  Dim Opt2Owed#
  Dim Opt3Owed#
  Dim TransRec As VATaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim MasterRec As VATaxMasterType
  Dim MHandle As Integer
  Dim ThisTaxYear As Integer
  Dim Message$
  Dim ThisBal As Double
  Dim DiscCheck As Integer
  Dim Dif As Double
  
  On Local Error GoTo ERRORSTUFF
  
  OpenVATaxSetUpFile MHandle
  Get MHandle, 1, MasterRec
  Close MHandle
  ThisTaxYear = MasterRec.RTaxYear
  
  ThisDiscPct = MasterRec.DisPPct
  
  PersOwed# = 0
  MTOwed# = 0
  MCOwed# = 0
  FEOwed# = 0
  MHOwed# = 0
  IntOwed# = 0
  PenOwed# = 0
  Opt1Owed# = 0
  Opt2Owed# = 0
  Opt3Owed# = 0
  OpenVATaxTransFile THandle, NumOfTRecs
  
  For x = 1 To BillCnt
    Get THandle, BillTrans(x), TransRec
      If TransRec.BillType = "P" Then
        PersOwed# = OldRound(PersOwed# + TransRec.Revenue.Principle1 - TransRec.PPTRADisc + TransRec.PPTRARmvl)
        PersOwed# = OldRound(PersOwed# - TransRec.Revenue.Principle1Pd)
        MTOwed# = OldRound(MTOwed# + TransRec.Revenue.Principle2)
        MTOwed# = OldRound(MTOwed# - TransRec.Revenue.Principle2Pd)
        MCOwed# = OldRound(MCOwed# + TransRec.Revenue.Principle3)
        MCOwed# = OldRound(MCOwed# - TransRec.Revenue.Principle3Pd)
        FEOwed# = OldRound(FEOwed# + TransRec.Revenue.Principle4)
        FEOwed# = OldRound(FEOwed# - TransRec.Revenue.Principle4Pd)
        MHOwed# = OldRound(MHOwed# + TransRec.Revenue.Principle5)
        MHOwed# = OldRound(MHOwed# - TransRec.Revenue.Principle5Pd)
        IntOwed# = OldRound(IntOwed# + TransRec.Revenue.Interest)
        IntOwed# = OldRound(IntOwed# - TransRec.Revenue.InterestPd)
        PenOwed# = OldRound(PenOwed# + TransRec.Revenue.Penalty)
        PenOwed# = OldRound(PenOwed# - TransRec.Revenue.PenaltyPd)
        Opt1Owed# = OldRound(Opt1Owed# + TransRec.Revenue.RevOpt1)
        Opt1Owed# = OldRound(Opt1Owed# - TransRec.Revenue.RevOpt1Pd)
        Opt2Owed# = OldRound(Opt2Owed# + TransRec.Revenue.RevOpt2)
        Opt2Owed# = OldRound(Opt2Owed# - TransRec.Revenue.RevOpt2Pd)
        Opt3Owed# = OldRound(Opt3Owed# + TransRec.Revenue.RevOpt3)
        Opt3Owed# = OldRound(Opt3Owed# - TransRec.Revenue.RevOpt3Pd)
      End If
  Next x

  fpCurrPersOwed = PersOwed#
  fpCurrMTOwed = MTOwed#
  fpCurrMCOwed = MCOwed#
  fpCurrFEOwed = FEOwed#
  fpCurrMHOwed = MHOwed#
  fpCurrIntOwed = IntOwed#
  fpCurrPenOwed = PenOwed#
  fpCurrOpt1Owed = Opt1Owed#
  fpCurrOpt2Owed = Opt2Owed#
  fpCurrOpt3Owed = Opt3Owed#
  fpCurrTotOwed = OldRound(PersOwed# + MTOwed# + MCOwed# + FEOwed# + MHOwed# + IntOwed# + PenOwed# + Opt1Owed# + Opt2Owed# + Opt3Owed#)
  fpCurrAmtOwed = OldRound(PersOwed# + MTOwed# + MCOwed# + FEOwed# + MHOwed# + IntOwed# + PenOwed# + Opt1Owed# + Opt2Owed# + Opt3Owed#)
  
  Close THandle
  
  MaxDisc = 0
  Call GetMaxDisc
  If MaxDisc > 0 Then
    If VAGetCustPersBalance(GCustNum, ThisTaxYear) > 0 Then
      Message = "This customer is eligible for a maximum personal discount of " + QPTrim$(Using$("$##,##0.00", MaxDisc)) + " but still owes money for past personal tax bills. If you wish to apply the discount anyway then press F10. Otherwise, press ESC to override the discount."
      If VATaxMsgWOpts(600, Message, "F10 Discount OK", "ESC NO Discount") = "abort" Then
        Unload frmVATaxMsgWOpts
        Call RemoveDiscount
      Else
        Unload frmVATaxMsgWOpts
        fpCurrDisc = MaxDisc
        Dif = OldRound(CDbl(fpCurrTotOwed.Value) - MaxDisc)
        Call VATaxMsg(900, "The total personal amount owed including the discount will be " + QPTrim$(Using$("$##,##0.00", Dif)) + ".")
      End If
    Else
      Message = "This customer is eligible for a maximum personal discount of " + QPTrim$(Using$("$##,##0.00", MaxDisc)) + ". If you wish to apply this discount then press F10. Otherwise, press ESC to override the discount."
      If VATaxMsgWOpts(700, Message, "F10 Discount OK", "ESC NO Discount") = "abort" Then
        Unload frmVATaxMsgWOpts
        Call RemoveDiscount
      Else
        Unload frmVATaxMsgWOpts
        fpCurrDisc = MaxDisc
        Dif = OldRound(CDbl(fpCurrTotOwed.Value) - MaxDisc)
        Call VATaxMsg(900, "The total personal amount owed including the discount will be " + QPTrim$(Using$("$##,##0.00", Dif)) + ".")
      End If
    End If
  Else
    Call RemoveDiscount
  End If
  fpcmbTenderType.SetFocus
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersPaymentEntry", "LoadAmtOwed", Erl)
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
'    ClearInUse PWcnt
 '   CMTerminate

End Sub

Private Function VACheck4Discounts() As Integer
  Dim TaxSURec As VATaxMasterType
  Dim MHandle As Integer
  Dim TaxTrans As VATaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim Balance#
  Dim x As Integer
  Dim TaxYear As Integer
  Dim ListRec As VAPersPayListType
  Dim ListHandle As Integer
  Dim NumOfLRecs As Integer
  Dim Operator$
  Dim PayTally As Double
  Dim DiscOK As Integer
  
  On Local Error GoTo ERRORSTUFF
  
  VACheck4Discounts = 0
  If DiscPXDate < Date2Num(fptxtPayDate.Text) Then
    fpCurrDisc.Value = 0
    Exit Function
  ElseIf DiscPXDate = 0 Then
    fpCurrDisc.Value = 0
    Exit Function
  End If
  
  ReDim DiscAmtAry(1 To 1) As Double
  ReDim DiscRecAry(1 To 1) As Long
  DiscAryCnt = 0
  
  Operator$ = CStr(OperNum)
  
  ReDim WhichRec(1 To 1) As Integer
  DiscCnt = 0
  
  OpenVATaxSetUpFile MHandle
  Get MHandle, 1, TaxSURec
  Close MHandle
  
  If TaxSURec.DisPPct = 0 Then
    Exit Function
  End If
  TaxYear = TaxSURec.RTaxYear
  
  ThisDiscAmt = 0
  PayTally = 0
  DiscOK = 0 '0 - no discounts allowed  1 - discounts allowed no warnings needed
  '2 - discounts can be allowed but...
  For x = 1 To BillCnt
    If TempBillList(x).TaxYear = TaxYear Then
      PayTally = OldRound(TempBillList(x).DiscAmt + TempBillList(x).MachTools + TempBillList(x).MerchCap + TempBillList(x).FarmEquip + TempBillList(x).Personal + TempBillList(x).MobHomes + TempBillList(x).Interest + TempBillList(x).Penalty)
      If TempBillList(x).TotOwed > PayTally Then  'no discount if nothing is paid for this tax year
        Exit Function
      End If
    End If
  Next x
  
  PayTally = 0
  For x = 1 To BillCnt
    PayTally = OldRound(PayTally + TempBillList(x).TotPaid)
  Next x
  If PayTally = CDbl(fpCurrTotOwed.Value) Then
    If VAGetCustPersBalance(GCustNum, TaxYear) > PayTally Then 'if true than old balance exists
      DiscOK = 2 'discount allowed but warn that some old balance is outstanding
    Else
      DiscOK = 1 'discount allowed...no warnings necessary
    End If
  End If
  
  If InSave = True And CDbl(fpCurrTotWDisc.Value) < CDbl(fpCurrTotOwed.Value) Then
    If VAGetCustPersBalance(GCustNum, TaxYear) > 0 Then 'if true than old balance exists
      DiscOK = 2 'discount allowed but warn that some old balance is outstanding
    Else
      DiscOK = 1 'discount allowed...no warnings necessary
    End If
  End If
  
DoOver1:
  OpenVATaxTransFile THandle, NumOfTRecs
  If BillCnt > 0 Then
    For x = 1 To BillCnt
      Get THandle, BillTrans(x), TaxTrans
      TempBillList(x).DiscAmt = 0
      TempBillList(x).DiscXDate = 0
'      TaxTrans.DiscXDate = DiscXDate 'for testing only
      If TaxTrans.DiscXDate > Date2Num(fptxtPayDate.Text) And TaxTrans.TaxYear = TaxYear Then
'      If TaxTrans.DiscXDate > Date2Num(fptxtPayDate.Text) And CInt(Mid(MakeRegDate(TaxTrans.TransDate), 7, 4)) = TAXYEAR Then
        Balance# = OldRound(Balance# + TaxTrans.Revenue.Principle1 + TaxTrans.Revenue.Principle2 + TaxTrans.Revenue.Principle3)
        Balance# = OldRound(Balance# + TaxTrans.Revenue.Principle4 + TaxTrans.Revenue.Principle5)
        Balance# = OldRound(Balance# - (TaxTrans.Revenue.Principle1Pd + TaxTrans.Revenue.Principle2Pd + TaxTrans.Revenue.Principle3Pd))
        Balance# = OldRound(Balance# - (TaxTrans.Revenue.Principle4Pd + TaxTrans.Revenue.Principle5Pd + TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
        If Balance# > 0 Then 'save which transaction the discount is applied to
          DiscCnt = DiscCnt + 1
          ReDim Preserve WhichRec(1 To DiscCnt) As Integer
          WhichRec(DiscCnt) = x
          ThisDiscAmt = ThisDiscAmt + OldRound(Balance# * (TaxSURec.DisPPct * 0.01))
          MaxDisc = ThisDiscAmt
          TempBillList(x).DiscAmt = ThisDiscAmt
          TempBillList(x).DiscXDate = TaxTrans.DiscXDate
          DiscAryCnt = DiscAryCnt + 1
          ReDim Preserve DiscAmtAry(1 To DiscAryCnt) As Double
          DiscAmtAry(DiscAryCnt) = ThisDiscAmt
          ReDim Preserve DiscRecAry(1 To DiscAryCnt) As Long
          DiscRecAry(DiscAryCnt) = BillTrans(x)
        End If
      End If
    Next x
    
    ThisDiscPct = TaxSURec.DisPPct 'assign to global
    If ThisDiscAmt > 0 Then
      VACheck4Discounts = DiscOK
    End If
    
    Close THandle
    
 End If
 
 Exit Function
 
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersPaymentEntry", "VACheck4Discounts", Erl)
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
 '   ClearInUse PWcnt
'    CMTerminate
 
End Function

Private Function Unique$(Path$)
  Dim TempName$
  Dim Seed&
  
  If Len(Path$) And Right$(Path$, 1) <> "\" Then Path$ = Path$ + "\"
  Seed& = Abs(Timer1)            'use the TIMER as a seed
  Do
    TempName$ = Path$ + Mid$(Str$(Seed&), 2)    'make a string out of it
    TempName$ = TempName$ + ".RPT"
    Seed& = Seed& + 1           'increment for next time
  Loop Until Not Exist(TempName$)              'loop and try another name
  Unique$ = TempName$           'this is the function output
  
End Function

Private Sub LoadHerUpEdit()
  Dim PayRec As VATaxPaymentRecType
  Dim PayHandle As Integer
  Dim TaxRec As VATaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  
  On Local Error GoTo ERRORSTUFF
  
  KillFile VATempPersBillRecs
  Label25.Visible = False
  NotFirstLoad = False
  BillCnt = 0

  OpenTempPersPayFile PayHandle, OperNum
  Get PayHandle, GPayNum, PayRec
  Close PayHandle
  
  OpenVATaxCustFile CHandle, NumOfCRecs
  Get CHandle, GCustNum, TaxRec
  Close CHandle
  
  TempAcctNum = TaxRec.Acct
  fpLongAcctNum = TaxRec.Acct
  fptxtName.Text = QPTrim$(TaxRec.CustName)
  If QPTrim$(TaxRec.Addr1) <> "" Then
    fptxtAddress.Text = QPTrim$(TaxRec.Addr1)
  Else
    fptxtAddress.Text = QPTrim$(TaxRec.Addr2)
  End If
  fptxtCity.Text = QPTrim$(TaxRec.City)
  fptxtState.Text = QPTrim$(TaxRec.State)
  fptxtZip = QPTrim$(TaxRec.Zip)
  
  fpCurrDisc = PayRec.DiscAmt
  fpCurrPersOwed = PayRec.PaidOwed(1).AmtOwed
  fpCurrPersPaid = PayRec.PaidOwed(1).AmtPaid
  fpCurrMTOwed = PayRec.PaidOwed(2).AmtOwed
  fpCurrMTPaid = PayRec.PaidOwed(2).AmtPaid
  fpCurrMCOwed = PayRec.PaidOwed(3).AmtOwed
  fpCurrMCPaid = PayRec.PaidOwed(3).AmtPaid
  fpCurrFEOwed = PayRec.PaidOwed(4).AmtOwed
  fpCurrFEPaid = PayRec.PaidOwed(4).AmtPaid
  fpCurrMHOwed = PayRec.PaidOwed(5).AmtOwed
  fpCurrMHPaid = PayRec.PaidOwed(5).AmtPaid
  fpCurrIntOwed = PayRec.PaidOwed(6).AmtOwed
  fpCurrIntPaid = PayRec.PaidOwed(6).AmtPaid
  fpCurrPenOwed = PayRec.PaidOwed(7).AmtOwed
  fpCurrPenPaid = PayRec.PaidOwed(7).AmtPaid
  fpCurrOpt1Owed = PayRec.PaidOwed(8).AmtOwed
  fpCurrOpt1Paid = PayRec.PaidOwed(8).AmtPaid
  fpCurrOpt2Owed = PayRec.PaidOwed(9).AmtOwed
  fpCurrOpt2Paid = PayRec.PaidOwed(9).AmtPaid
  fpCurrOpt3Owed = PayRec.PaidOwed(10).AmtOwed
  fpCurrOpt3Paid = PayRec.PaidOwed(10).AmtPaid
  fpCurrTotOwed = PayRec.TotOwed
  fpCurrAmtOwed = PayRec.AmtOwed
  fpCurrCashPd = PayRec.CashAmt
  If PayRec.ChkAmt > 0 Then
    fpCurrChkChrgPd = PayRec.ChkAmt
  ElseIf PayRec.ChrgAmt > 0 Then
    fpCurrChkChrgPd = PayRec.ChrgAmt
  Else
    fpCurrChkChrgPd = 0
  End If
  fpcmbTenderType.Text = QPTrim$(PayRec.TenderTY)
  fptxtPayDate.Text = MakeRegDate(PayRec.payDate)
  fpCurrTotRecd = PayRec.AmtRecd
  fpCurrChngDue = PayRec.Change
  fptxtDescription.Text = QPTrim$(PayRec.Desc)
  fpCurrTotPaid = OldRound(PayRec.TotPaid + PayRec.PrePayAmt)
  fpCurrPrePay = PayRec.PrePayAmt
  fpCurrTotWDisc = OldRound(CDbl(fpCurrTotPaid.Value) + PayRec.DiscAmt)
  GetNewCust = False
  Call GetMaxDisc
  Call LoadTempPayList
  Call LoadTemps
  Call AssignPaidTemps
  If OverPay = False Then
    Call EnablePayFields
  End If
  EditFlag = True
'  cmdBills.Enabled = False
  BegAmount = fpCurrTotRecd.Value 'added 7/18/07***********************************
  If CDbl(fpCurrDisc.Value) = 0 Then Call RemoveDiscount
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersPaymentEntry", "LoadHerUpEdit", Erl)
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
 '   ClearInUse PWcnt
 '   CMTerminate

End Sub


Private Sub UPDateListOfPayments()
  'Keeps up with which tagged bills go with which customer
  'If the bill list is not accessed then this sub is not used
  Dim TempRec As TempPayList
  Dim NumOfTemps As Integer
  Dim TempHandle As Integer
  Dim Operator$
  Dim TPayRec As VAPersPayListType
  Dim PayRec As VAPersPayListType
  Dim PHandle As Integer
  Dim NumOfPRecs As Integer
  Dim x As Integer, y As Integer, z As Integer
  Dim ThisPrevRec As Long
  Dim NewRec As Integer
  Dim TotPaid#
  Dim PrevCnt As Integer
  Dim FoundCnt As Integer
  Dim TaxMasterRec As VATaxMasterType
  Dim MHandle As Integer
  Dim ThisPrePay As Double
  Dim Nextx As Integer
  Dim tottotpaid As Double
  On Local Error GoTo ERRORSTUFF
  ThisPrePay = CDbl(fpCurrPrePay.Value)
  ThisPrevRec = 0
  NewRec = 0
  Operator$ = CStr(OperNum)
  Operator$ = QPTrim$(Operator$)
  OpenPersPayListFile PHandle, OperNum 'saved by getting data from temporary
  NumOfPRecs = LOF(PHandle) / Len(PayRec)
  LastPayRec = 0
  
  If BillCnt = 0 And VAGetCustPersBalance(GCustNum, -1) <= 0 Then 'customer
  'owes nothing and wants to prepay
    OpenVATaxSetUpFile MHandle
    Get MHandle, 1, TaxMasterRec
    Close MHandle
    
    TotPaid# = 0
    PayRec.BillRec = -GCustNum
    PayRec.CustRec = GCustNum
    PayRec.PrevListRec = 0
    'the following should always be zero
    PayRec.Personal = CDbl(fpCurrPersPaid.Value)
    TotPaid# = CDbl(fpCurrPersPaid.Value)
    PayRec.MachTools = CDbl(fpCurrMTPaid.Value)
    TotPaid# = OldRound(TotPaid# + CDbl(fpCurrMTPaid.Value))
    PayRec.MerchCap = CDbl(fpCurrMCPaid.Value)
    TotPaid# = OldRound(TotPaid# + CDbl(fpCurrMCPaid.Value))
    PayRec.FarmEquip = CDbl(fpCurrFEPaid.Value)
    TotPaid# = OldRound(TotPaid# + CDbl(fpCurrFEPaid.Value))
    PayRec.MobHomes = CDbl(fpCurrMHPaid.Value)
    TotPaid# = OldRound(TotPaid# + CDbl(fpCurrMHPaid.Value))
    PayRec.Interest = CDbl(fpCurrIntPaid.Value)
    TotPaid# = OldRound(TotPaid# + CDbl(fpCurrIntPaid.Value))
    PayRec.Penalty = CDbl(fpCurrPenPaid.Value)
    TotPaid# = OldRound(TotPaid# + CDbl(fpCurrPenPaid.Value))
    PayRec.DiscAmt = CDbl(fpCurrDisc.Value) 'should be zero always
    TotPaid# = OldRound(TotPaid# + CDbl(fpCurrDisc.Value))
    PayRec.Opt1 = CDbl(fpCurrOpt1Paid.Value)
    TotPaid# = OldRound(TotPaid# + CDbl(fpCurrOpt1Paid.Value))
    PayRec.Opt2 = CDbl(fpCurrOpt2Paid.Value)
    TotPaid# = OldRound(TotPaid# + CDbl(fpCurrOpt2Paid.Value))
    PayRec.Opt3 = CDbl(fpCurrOpt3Paid.Value)
    TotPaid# = OldRound(TotPaid# + CDbl(fpCurrOpt3Paid.Value))
    PayRec.TaxYear = TaxMasterRec.RTaxYear
    PayRec.Description = QPTrim$(fptxtDescription.Text)
    'ThisPrePay will have a value
    PayRec.PrePayAmt = ThisPrePay
    ThisPrePay = 0
    NumOfPRecs = NumOfPRecs + 1
    LastPayRec = NumOfPRecs
    PayRec.PrevListRec = 0
    Put PHandle, NumOfPRecs, PayRec
    Close PHandle
    Exit Sub
  End If
    
  If NumOfPRecs = 0 Then 'first record saved in the bills list
    For y = 1 To BillCnt
      If TempBillList(y).BillRec > 0 Then
      'this is saving the totals of all bills tagged
        TotPaid# = 0
        PayRec.BillRec = TempBillList(y).BillRec  'TempRec.BillPtr
        PayRec.CustRec = GCustNum
        PayRec.PrevListRec = y - 1
        PayRec.Personal = TempBillList(y).Personal  'CDbl(fpCurrPersPaid.Value)
        TotPaid# = PayRec.Personal
        PayRec.MachTools = TempBillList(y).MachTools  ' CDbl(fpCurrMTPaid.Value)
        TotPaid# = OldRound(TotPaid# + PayRec.MachTools)
        PayRec.MerchCap = TempBillList(y).MerchCap  'CDbl(fpCurrMCPaid.Value)
        TotPaid# = OldRound(TotPaid# + PayRec.MerchCap)
        PayRec.FarmEquip = TempBillList(y).FarmEquip  'CDbl(fpCurrFEPaid.Value)
        TotPaid# = OldRound(TotPaid# + PayRec.FarmEquip)
        PayRec.MobHomes = TempBillList(y).MobHomes  'CDbl(fpCurrMHPaid.Value)
        TotPaid# = OldRound(TotPaid# + PayRec.MobHomes)
        PayRec.Interest = TempBillList(y).Interest  ' CDbl(fpCurrIntPaid.Value)
        TotPaid# = OldRound(TotPaid# + PayRec.Interest)
        PayRec.Penalty = TempBillList(y).Penalty  'CDbl(fpCurrPenPaid.Value)
        TotPaid# = OldRound(TotPaid# + PayRec.Penalty)
        PayRec.Opt1 = TempBillList(y).Opt1
        TotPaid# = OldRound(TotPaid# + PayRec.Opt1)
        PayRec.Opt2 = TempBillList(y).Opt2
        TotPaid# = OldRound(TotPaid# + PayRec.Opt2)
        PayRec.Opt3 = TempBillList(y).Opt3
        TotPaid# = OldRound(TotPaid# + PayRec.Opt3)
        PayRec.TotPaid = TotPaid#  'CDbl(fpCurrTotPaid.Value)
        PayRec.DiscAmt = TempBillList(y).DiscAmt  'CDbl(fpCurrDisc.Value)
        PayRec.TaxYear = TempBillList(y).TaxYear
        PayRec.Description = QPTrim$(fptxtDescription.Text)
        PayRec.PrePayAmt = ThisPrePay 'prepay applies once
        ThisPrePay = 0
        NewRec = NewRec + 1
        LastPayRec = NewRec
        Put PHandle, NewRec, PayRec
      End If
    Next y
    Close PHandle
    Exit Sub
  End If
  
  If BillCnt > 0 And EditFlag = False Then 'bill list has been processed
  'for a new customer
    For y = 1 To BillCnt 'NumOfTemps 'NumOfTemps is how many bill records there are
    'fpr thios customer after having accessed the bill tag screen
      If TempBillList(y).BillRec > 0 Then
        PayRec.BillRec = TempBillList(y).BillRec
        PayRec.CustRec = GCustNum
        PayRec.Personal = TempBillList(y).Personal
        TotPaid# = PayRec.Personal
        PayRec.MachTools = TempBillList(y).MachTools
        TotPaid# = OldRound(TotPaid# + PayRec.MachTools)
        PayRec.MerchCap = TempBillList(y).MerchCap
        TotPaid# = OldRound(TotPaid# + PayRec.MerchCap)
        PayRec.FarmEquip = TempBillList(y).FarmEquip
        TotPaid# = OldRound(TotPaid# + PayRec.FarmEquip)
        PayRec.MobHomes = TempBillList(y).MobHomes
        TotPaid# = OldRound(TotPaid# + PayRec.MobHomes)
        PayRec.Interest = TempBillList(y).Interest
        TotPaid# = OldRound(TotPaid# + PayRec.Interest)
        PayRec.Penalty = TempBillList(y).Penalty
        TotPaid# = OldRound(TotPaid# + PayRec.Penalty)
        PayRec.Opt1 = TempBillList(y).Opt1
        TotPaid# = OldRound(TotPaid# + PayRec.Opt1)
        PayRec.Opt2 = TempBillList(y).Opt2
        TotPaid# = OldRound(TotPaid# + PayRec.Opt2)
        PayRec.Opt3 = TempBillList(y).Opt3
        TotPaid# = OldRound(TotPaid# + PayRec.Opt3)
        PayRec.TotPaid = TotPaid#
        PayRec.DiscAmt = TempBillList(y).DiscAmt
        PayRec.TaxYear = TempBillList(y).TaxYear
        PayRec.Description = QPTrim$(fptxtDescription.Text)
        PayRec.PrePayAmt = ThisPrePay
        ThisPrePay = 0
        NumOfPRecs = NumOfPRecs + 1
        LastPayRec = NumOfPRecs
        If y = 1 Then
          PayRec.PrevListRec = 0
        Else
          PayRec.PrevListRec = NumOfPRecs - 1
        End If
        Put PHandle, NumOfPRecs, PayRec
      End If
    Next y
    Close PHandle
    Exit Sub
  End If
   
  FoundCnt = 0
  For x = 1 To NumOfPRecs 'number of bills tagged
    Get PHandle, x, PayRec
    If PayRec.CustRec = GCustNum Then
      For y = 1 To BillCnt
        If PayRec.BillRec = TempBillList(y).BillRec And TempBillList(y).BillRec > 0 Then
          FoundCnt = FoundCnt + 1
          PayRec.Personal = TempBillList(y).Personal
          TotPaid# = PayRec.Personal
          PayRec.MachTools = TempBillList(y).MachTools
          TotPaid# = OldRound(TotPaid# + PayRec.MachTools)
          PayRec.MerchCap = TempBillList(y).MerchCap
          TotPaid# = OldRound(TotPaid# + PayRec.MerchCap)
          PayRec.FarmEquip = TempBillList(y).FarmEquip
          TotPaid# = OldRound(TotPaid# + PayRec.FarmEquip)
          PayRec.MobHomes = TempBillList(y).MobHomes
          TotPaid# = OldRound(TotPaid# + PayRec.MobHomes)
          PayRec.Interest = TempBillList(y).Interest
          TotPaid# = OldRound(TotPaid# + PayRec.Interest)
          PayRec.Penalty = TempBillList(y).Penalty
          TotPaid# = OldRound(TotPaid# + PayRec.Penalty)
          PayRec.Opt1 = TempBillList(y).Opt1
          TotPaid# = OldRound(TotPaid# + PayRec.Opt1)
          PayRec.Opt2 = TempBillList(y).Opt2
          TotPaid# = OldRound(TotPaid# + PayRec.Opt2)
          PayRec.Opt3 = TempBillList(y).Opt3
          TotPaid# = OldRound(TotPaid# + PayRec.Opt3)
          PayRec.TotPaid = TotPaid#
          PayRec.DiscAmt = TempBillList(y).DiscAmt
          PayRec.BillRec = TempBillList(y).BillRec
          PayRec.TaxYear = TempBillList(y).TaxYear
          PayRec.Description = QPTrim$(fptxtDescription.Text)
          PayRec.PrePayAmt = ThisPrePay
          ThisPrePay = 0
          LastPayRec = x
          Put PHandle, x, PayRec
        End If
      Next y
    End If
  Next x
  
  Nextx = 1
  If FoundCnt < BillCnt Then
    For z = NumOfPRecs To 1 Step -1
      Get PHandle, z, PayRec
      If PayRec.CustRec = GCustNum Then
        ThisPrevRec = z
      End If
    Next z
    
    For y = 1 To BillCnt 'take one bill at a time
    'if the bill rec #s match then they were already saved above
    'if they don't match then we have a new billrec and it will
    'be saved below
      For x = 1 To NumOfPRecs
        Get PHandle, x, PayRec
        If PayRec.BillRec = TempBillList(y).BillRec Then
          GoTo NextOne
        End If
      Next x
      
      PrevCnt = 0
      PayRec.CustRec = GCustNum
      PayRec.Personal = TempBillList(y).Personal
      TotPaid# = PayRec.Personal
      PayRec.MachTools = TempBillList(y).MachTools
      TotPaid# = OldRound(TotPaid# + PayRec.MachTools)
      PayRec.MerchCap = TempBillList(y).MerchCap
      TotPaid# = OldRound(TotPaid# + PayRec.MerchCap)
      PayRec.FarmEquip = TempBillList(y).FarmEquip
      TotPaid# = OldRound(TotPaid# + PayRec.FarmEquip)
      PayRec.MobHomes = TempBillList(y).MobHomes
      TotPaid# = OldRound(TotPaid# + PayRec.MobHomes)
      PayRec.Interest = TempBillList(y).Interest
      TotPaid# = OldRound(TotPaid# + PayRec.Interest)
      PayRec.Penalty = TempBillList(y).Penalty
      TotPaid# = OldRound(TotPaid# + PayRec.Penalty)
      PayRec.Opt1 = TempBillList(y).Opt1
      TotPaid# = OldRound(TotPaid# + PayRec.Opt1)
      PayRec.Opt2 = TempBillList(y).Opt2
      TotPaid# = OldRound(TotPaid# + PayRec.Opt2)
      PayRec.Opt3 = TempBillList(y).Opt3
      TotPaid# = OldRound(TotPaid# + PayRec.Opt3)
      PayRec.TotPaid = TotPaid#
      PayRec.DiscAmt = TempBillList(y).DiscAmt
      PayRec.BillRec = TempBillList(y).BillRec
      PayRec.TaxYear = TempBillList(y).TaxYear
      PayRec.PrePayAmt = ThisPrePay
      ThisPrePay = 0
      NumOfPRecs = NumOfPRecs + 1
      LastPayRec = NumOfPRecs
      PayRec.PrevListRec = ThisPrevRec
      Put PHandle, NumOfPRecs, PayRec
      ThisPrevRec = NumOfPRecs
NextOne:
    Next y
  End If
  Close PHandle
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersPaymentEntry", "UPDateListOfPayments", Erl)
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
 '   ClearInUse PWcnt
 '   CMTerminate

End Sub
  
Private Function AddUpPaidCol() As Double
  Dim ThisAdd As Double
  Dim MatchAdd As Double
  Dim Message$
  
  On Local Error GoTo ERRORSTUFF
  
  MatchAdd = CDbl(fpCurrTotPaid.Value)
  If NotFirstLoad = False Then Exit Function
  AddUpPaidCol = 0
  ThisAdd = OldRound(CDbl(fpCurrPersPaid.Value) + CDbl(fpCurrMTPaid.Value) + CDbl(fpCurrMCPaid.Value))
  ThisAdd = OldRound(ThisAdd + CDbl(fpCurrFEPaid.Value) + CDbl(fpCurrMHPaid.Value))
  ThisAdd = OldRound(ThisAdd + CDbl(fpCurrIntPaid.Value) + CDbl(fpCurrPenPaid.Value) + CDbl(fpCurrPrePay.Value))
  ThisAdd = OldRound(ThisAdd + CDbl(fpCurrOpt1Paid.Value) + CDbl(fpCurrOpt2Paid.Value) + CDbl(fpCurrOpt3Paid.Value))
  If ThisAdd = MatchAdd Then
    AddUpPaidCol = ThisAdd
    fpCurrTotPaid = ThisAdd
    fpCurrChngDue = OldRound(CDbl(fpCurrTotRecd.Value) - ThisAdd)
    If CDbl(fpCurrChngDue.Value) < 0 Then fpCurrChngDue = 0
    fpCurrTotWDisc = OldRound(CDbl(fpCurrTotPaid.Value) + CDbl(fpCurrDisc.Value))
    Exit Function
  End If
  
  fpCurrChngDue = OldRound(CDbl(fpCurrTotRecd.Value) - ThisAdd)
  If CDbl(fpCurrChngDue.Value) < 0 Then fpCurrChngDue = 0
  If ThisAdd > OldRound(CDbl(fpCurrTotRecd.Value) - fpCurrChngDue.Value) Then
'    fpCurrChngDue = OldRound(CDbl(fpCurrTotRecd.Value) - ThisAdd)
    fpCurrTotPaid = ThisAdd
    fpCurrTotRecd.BackColor = &H8080FF
    fpCurrTotPaid.BackColor = &H8080FF
    frmVATaxMsg.Label1.Caption = "The amount distributed exceeds the amount received. Please re-distribute the amounts paid or add funds to the 'Cash Amount Paid' or the 'Check/Charge Amount Paid' fields."
    frmVATaxMsg.Label1.Top = 700
    frmVATaxMsg.Show vbModal
    If fpcmbTenderType.Text = "CASH" Or fpcmbTenderType.Text = "CASH AND CHECK" Then
      If fpCurrCashPd.Enabled = True Then
        fpCurrCashPd.SetFocus
      End If
    Else
      If fpCurrChkChrgPd.Enabled = True Then
        fpCurrChkChrgPd.SetFocus
      End If
    End If
  End If
  AddUpPaidCol = ThisAdd
  fpCurrTotPaid = ThisAdd
'  fpCurrChngDue = OldRound(CDbl(fpCurrTotRecd.Value) - ThisAdd)
'  If CDbl(fpCurrChngDue.Value) < 0 Then fpCurrChngDue = 0
  fpCurrTotWDisc = OldRound(CDbl(fpCurrTotPaid.Value) + CDbl(fpCurrDisc.Value))
  
  Exit Function
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersPaymentEntry", "AddUpPaidCol", Erl)
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
 '   ClearInUse PWcnt
 '   CMTerminate

End Function

Private Sub MakeEmWhite()
  fpLongAcctNum.BackColor = &H80000005
  fpCurrAmtOwed.BackColor = &H80000005
  fpCurrCashPd.BackColor = &H80000005
  fpCurrChkChrgPd.BackColor = &H80000005
  fpCurrTotRecd.BackColor = &H80000005
  fpCurrChngDue.BackColor = &H80000005
  fpCurrTotOwed.BackColor = &H80000005
  fpCurrTotPaid.BackColor = &H80000005
  fpCurrPersPaid.BackColor = &H80000005
  fpCurrPersOwed.BackColor = &H80000005
  fpCurrMTPaid.BackColor = &H80000005
  fpCurrMTOwed.BackColor = &H80000005
  fpCurrMCPaid.BackColor = &H80000005
  fpCurrMCOwed.BackColor = &H80000005
  fpCurrFEPaid.BackColor = &H80000005
  fpCurrFEOwed.BackColor = &H80000005
  fpCurrMHPaid.BackColor = &H80000005
  fpCurrMHOwed.BackColor = &H80000005
  fpCurrIntPaid.BackColor = &H80000005
  fpCurrIntOwed.BackColor = &H80000005
  fpCurrPenPaid.BackColor = &H80000005
  fpCurrPenOwed.BackColor = &H80000005
  fpCurrOpt1Paid.BackColor = &H80000005
  fpCurrOpt1Owed.BackColor = &H80000005
  fpCurrOpt2Paid.BackColor = &H80000005
  fpCurrOpt2Owed.BackColor = &H80000005
  fpCurrOpt3Paid.BackColor = &H80000005
  fpCurrOpt3Owed.BackColor = &H80000005
  fpCurrDisc.BackColor = &H80000005
  fpCurrTotWDisc.BackColor = &H80000005
  fpCurrPrePay.BackColor = &H80000005
End Sub

Private Sub ReFigure()
  fpCurrChngDue = 0
  fpCurrChngDue = OldRound(CDbl(fpCurrTotRecd.Value) - CDbl(fpCurrTotPaid.Value))
  If CDbl(fpCurrChngDue.Value) < 0 Then fpCurrChngDue = 0
End Sub

Private Sub LoadTemps()
  'Temp variables are used to reset changes back to what the values were
  'before changes were made...used extensively in Check4Changes
  Dim PayRec As VATaxPaymentRecType
  Dim PayHandle As Integer
  Dim x As Integer
  Dim NumOfPayRecs As Integer
  
  TempAcctNum = fpLongAcctNum.Value
  OpenTempPersPayFile PayHandle, OperNum
  NumOfPayRecs = LOF(PayHandle) / Len(PayRec)
  For x = 1 To NumOfPayRecs
    Get PayHandle, x, PayRec
    If PayRec.CustAcct = CDbl(fpLongAcctNum.Text) Then
      TempPayDate = PayRec.payDate
      TempAmtOwed = PayRec.AmtOwed
      TempTenderTY = QPTrim$(PayRec.TenderTY)
      TempCashAmt = PayRec.CashAmt
      TempChkAmt = PayRec.ChkAmt
      TempChrgAmt = PayRec.ChrgAmt
      TempAmtRecd = PayRec.AmtRecd
      TempChange = PayRec.Change
      TempDesc = QPTrim$(PayRec.Desc)
      TempPaidOwed1AmtOwed = PayRec.PaidOwed(1).AmtOwed
      TempPaidOwed2AmtOwed = PayRec.PaidOwed(2).AmtOwed
      TempPaidOwed3AmtOwed = PayRec.PaidOwed(3).AmtOwed
      TempPaidOwed4AmtOwed = PayRec.PaidOwed(4).AmtOwed
      TempPaidOwed5AmtOwed = PayRec.PaidOwed(5).AmtOwed
      TempPaidOwed6AmtOwed = PayRec.PaidOwed(6).AmtOwed
      TempPaidOwed7AmtOwed = PayRec.PaidOwed(7).AmtOwed
      TempPaidOwed8AmtOwed = PayRec.PaidOwed(8).AmtOwed
      TempPaidOwed9AmtOwed = PayRec.PaidOwed(9).AmtOwed
      TempPaidOwed10AmtOwed = PayRec.PaidOwed(10).AmtOwed
      TempPaidOwed1AmtPaid = PayRec.PaidOwed(1).AmtPaid
      TempPaidOwed2AmtPaid = PayRec.PaidOwed(2).AmtPaid
      TempPaidOwed3AmtPaid = PayRec.PaidOwed(3).AmtPaid
      TempPaidOwed4AmtPaid = PayRec.PaidOwed(4).AmtPaid
      TempPaidOwed5AmtPaid = PayRec.PaidOwed(5).AmtPaid
      TempPaidOwed6AmtPaid = PayRec.PaidOwed(6).AmtPaid
      TempPaidOwed7AmtPaid = PayRec.PaidOwed(7).AmtPaid
      TempPaidOwed8AmtPaid = PayRec.PaidOwed(8).AmtPaid
      TempPaidOwed9AmtPaid = PayRec.PaidOwed(9).AmtPaid
      TempPaidOwed10AmtPaid = PayRec.PaidOwed(10).AmtPaid
      TempTotOwed = PayRec.TotOwed
      TempAmtPaid = PayRec.AmtPaid
      TempTotPaid = PayRec.TotPaid
'      TempAcctNum = CLng(fpLongAcctNum.Value)
      Exit For
    End If
  Next x
  If x > NumOfPayRecs Then
    TempPayDate = Date2Num(payDate) '2/14/06
    TempAmtOwed = 0
    TempTenderTY = "CASH"
    TempCashAmt = 0
    TempChkAmt = 0
    TempChrgAmt = 0
    TempAmtRecd = 0
    TempChange = 0
    TempDesc = ""
    TempPaidOwed1AmtOwed = 0
    TempPaidOwed2AmtOwed = 0
    TempPaidOwed3AmtOwed = 0
    TempPaidOwed4AmtOwed = 0
    TempPaidOwed5AmtOwed = 0
    TempPaidOwed6AmtOwed = 0
    TempPaidOwed7AmtOwed = 0
    TempPaidOwed8AmtOwed = 0
    TempPaidOwed9AmtOwed = 0
    TempPaidOwed10AmtOwed = 0
    TempPaidOwed1AmtPaid = 0
    TempPaidOwed2AmtPaid = 0
    TempPaidOwed3AmtPaid = 0
    TempPaidOwed4AmtPaid = 0
    TempPaidOwed5AmtPaid = 0
    TempPaidOwed6AmtPaid = 0
    TempPaidOwed7AmtPaid = 0
    TempPaidOwed8AmtPaid = 0
    TempPaidOwed9AmtPaid = 0
    TempPaidOwed10AmtPaid = 0
    TempTotOwed = 0
    TempAmtPaid = 0
    TempTotPaid = 0
'    TempAcctNum = 0
  End If
  
  Close PayHandle
End Sub

Public Function Check4Changes() As Boolean
  Dim PayRec As VATaxPaymentRecType
  Dim PayHandle As Integer
  Dim x As Integer
  Dim NumOfPayRecs As Integer
  Dim Operator$
  Dim choice As String
  Dim ThisControl As Control
  Dim ThisDesc As String
  Dim ThatDesc As String
  Dim ThisText As String
  Dim ThisDbl As Double
  Dim ThatDbl As Double
  Dim ThisInt As Integer
  Dim ThatInt As Integer
  
  On Local Error GoTo ERRORSTUFF
  Check4Changes = False
  If fpLongAcctNum.Value = 0 And BillCnt = 0 Then Exit Function
  
  Set ThisControl = fptxtPayDate
  ThisText = fptxtPayDate.Text
  ThisDesc = MakeRegDate(TempPayDate)
  If ThisText <> ThisDesc Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
'  Set ThisControl = fpCurrAmtOwed
'  thisDbl = CDbl(fpCurrAmtOwed.Value)
'  thatDbl = TempAmtOwed
'  If thisDbl <> thatDbl Then
'    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
'    frmVATaxMsgW3Opts.Label1.Top = 800
'    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
'    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
'    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
'    frmVATaxMsgW3Opts.Show vbModal
'    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
'    Unload frmVATaxMsgW3Opts
'    If choice = "continue" Then
'      DontExit = False
''      Close PayHandle
'      Call cmdSave_Click
'      Exit Function
'    Else
'      GoSub HandleChoice
'    End If
'  End If
  
  Set ThisControl = fpcmbTenderType
  ThisDesc = QPTrim$(fpcmbTenderType.Text)
  ThatDesc = TempTenderTY
  If ThisDesc <> ThatDesc Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpCurrCashPd
  ThisDbl = CDbl(fpCurrCashPd.Value)
  ThatDbl = TempCashAmt
  If ThisDbl <> ThatDbl Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpCurrChkChrgPd
  ThisDbl = CDbl(fpCurrChkChrgPd.Value)
  If TempTenderTY = "CHECK" Or TempTenderTY = "CASH AND CHECK" Then
    ThatDbl = TempChkAmt
  ElseIf TempTenderTY = "CHARGE" Then
    ThatDbl = TempChrgAmt
  Else
    ThatDbl = 0
  End If
  If ThisDbl <> ThatDbl Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
'  Set ThisControl = fpCurrTotRecd
'  thisDbl = CDbl(fpCurrTotRecd.Value)
'  thatDbl = TempAmtRecd
'  If thisDbl <> thatDbl Then
'    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
'    frmVATaxMsgW3Opts.Label1.Top = 800
'    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
'    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
'    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
'    frmVATaxMsgW3Opts.Show vbModal
'    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
'    Unload frmVATaxMsgW3Opts
'    If choice = "continue" Then
'      DontExit = False
''      Close PayHandle
'      Call cmdSave_Click
'      Exit Function
'    Else
'      GoSub HandleChoice
'    End If
'  End If
  
  Set ThisControl = fpCurrChngDue
  ThisDbl = CDbl(fpCurrChngDue.Value)
  ThatDbl = TempChange
  If ThisDbl <> ThatDbl Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
'      Close PayHandle
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fptxtDescription
  ThisDesc = QPTrim$(fptxtDescription.Text)
  If ThisDesc = "" Then ThisDesc = "BLANK"
  ThatDesc = TempDesc
  If ThatDesc = "" Then ThatDesc = "BLANK"
  If ThisDesc <> ThatDesc Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.Show vbModal
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
'      Close PayHandle
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpCurrPrePay
  ThisDbl = CDbl(fpCurrPrePay.Value)
  ThatDbl = TempPrePay
  If ThisDbl <> ThatDbl Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpCurrMTPaid
  ThisDbl = CDbl(fpCurrMTPaid.Value)
  ThatDbl = TempPaidOwed2AmtPaid
  If ThisDbl <> ThatDbl Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpCurrMCPaid
  ThisDbl = CDbl(fpCurrMCPaid.Value)
  ThatDbl = TempPaidOwed3AmtPaid
  If ThisDbl <> ThatDbl Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
  Set ThisControl = fpCurrFEPaid
  ThisDbl = CDbl(fpCurrFEPaid.Value)
  ThatDbl = TempPaidOwed4AmtPaid
  If ThisDbl <> ThatDbl Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
      
  Set ThisControl = fpCurrPersPaid
  ThisDbl = CDbl(fpCurrPersPaid.Value)
  ThatDbl = TempPaidOwed1AmtPaid
  If ThisDbl <> ThatDbl Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
      
  Set ThisControl = fpCurrMHPaid
  ThisDbl = CDbl(fpCurrMHPaid.Value)
  ThatDbl = TempPaidOwed5AmtPaid
  If ThisDbl <> ThatDbl Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
      
  Set ThisControl = fpCurrIntPaid
  ThisDbl = CDbl(fpCurrIntPaid.Value)
  ThatDbl = TempPaidOwed6AmtPaid
  If ThisDbl <> ThatDbl Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
      
  Set ThisControl = fpCurrPenPaid
  ThisDbl = CDbl(fpCurrPenPaid.Value)
  ThatDbl = TempPaidOwed7AmtPaid
  If ThisDbl <> ThatDbl Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
      
  Set ThisControl = fpCurrOpt1Paid
  ThisDbl = CDbl(fpCurrOpt1Paid.Value)
  ThatDbl = TempPaidOwed8AmtPaid
  If ThisDbl <> ThatDbl Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
      
  Set ThisControl = fpCurrOpt2Paid
  ThisDbl = CDbl(fpCurrOpt2Paid.Value)
  ThatDbl = TempPaidOwed9AmtPaid
  If ThisDbl <> ThatDbl Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
      
  Set ThisControl = fpCurrOpt3Paid
  ThisDbl = CDbl(fpCurrOpt3Paid.Value)
  ThatDbl = TempPaidOwed10AmtPaid
  If ThisDbl <> ThatDbl Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
      
  Set ThisControl = fpCurrTotOwed
  ThisDbl = CDbl(fpCurrTotOwed.Value)
  ThatDbl = TempTotOwed
  If ThisDbl <> ThatDbl Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
      
  Set ThisControl = fpCurrTotRecd
  ThisDbl = CDbl(fpCurrTotRecd.Value)
  ThatDbl = TempAmtRecd 'Paid
  If ThisDbl <> ThatDbl Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
      
  Set ThisControl = fpCurrTotPaid
  ThisDbl = CDbl(fpCurrTotPaid.Value)
  ThatDbl = TempTotPaid + TempPrePay
  If ThisDbl <> ThatDbl Then
    frmVATaxMsgW3Opts.Label1.Caption = "Changes have been made. Do you wish to save these changes. Press F10 to save, press F5 to review or press ESC to abandon all changes."
    frmVATaxMsgW3Opts.Label1.Top = 800
    frmVATaxMsgW3Opts.cmdCont.Text = "F10 Save"
    frmVATaxMsgW3Opts.cmdExit.Text = "ESC Don't Save"
    frmVATaxMsgW3Opts.cmdOption.Text = "F5 Review"
    frmVATaxMsgW3Opts.Show vbModal
    choice = frmVATaxMsgW3Opts.fptxtChoice.Text
    Unload frmVATaxMsgW3Opts
    If choice = "continue" Then
      DontExit = False
      Call cmdSave_Click
      Exit Function
    Else
      GoSub HandleChoice
    End If
  End If
  
  Exit Function
  
HandleChoice:
  Close PayHandle
  Select Case choice
    Case "abort" 'don't save
'      If Exist("editpyment.dat") Or Lookup = True Then '2/14/06 added Or Lookup
'        Exit Function 'trying to access another customer
'      ElseIf TempAcctNum = CLng(fpLongAcctNum.Value) Then
'        frmVATaxPayMenu.Show
'        DoEvents
'        KillFile "txperspyment.dat"
'        Unload Me
'      End If
      Exit Function
    Case "option" 'review
      fpLongAcctNum = TempAcctNum
      If ThisControl.Enabled = True Then
        ThisControl.SetFocus
      Else
        ThisControl.BackColor = &H8080FF
      End If
      Close PayHandle
      Check4Changes = True
      Exit Function
    Case Else
  End Select
      
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersPaymentEntry", "Check4Changes", Erl)
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

End Function

Private Sub fptxtDescription_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDown Or KeyCode = vbKeyTab Then
    fpLongAcctNum.SetFocus
  ElseIf KeyCode = vbKeyUp Then
    If fpCurrOpt3Paid.Enabled = True Then
      fpCurrOpt3Paid.SetFocus
    ElseIf fpCurrOpt2Paid.Enabled = True Then
      fpCurrOpt2Paid.SetFocus
    ElseIf fpCurrOpt1Paid.Enabled = True Then
      fpCurrOpt1Paid.SetFocus
    Else
      fpCurrMHPaid.SetFocus
    End If
  End If
End Sub

'Private Sub fptxtPayDate_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeyDown Then
'    fpLongAcctNum.SetFocus
'  ElseIf KeyCode = vbKeyUp Then
'    fptxtDescription.SetFocus
'  End If
'End Sub

Private Function CheckOverPay() As Boolean

  On Local Error GoTo ERRORSTUFF
  
  'looks for overpayment of revenues if others are not fully paid...not allowed
  CheckOverPay = False
  If CDbl(fpCurrMTPaid.Value) = CDbl(fpCurrMTOwed.Value) Then
    GoTo MTOK
  ElseIf CDbl(fpCurrMTOwed.Value) > CDbl(fpCurrMTPaid.Value) Then
    If CDbl(fpCurrPersPaid.Value) > CDbl(fpCurrPersOwed.Value) Then
      fpCurrMTOwed.BackColor = &H8080FF
      fpCurrMTPaid.BackColor = &H8080FF
      fpCurrPersPaid.BackColor = &H8080FF
      fpCurrPersOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Personal while underpaying Machine Tools. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrMCPaid.Value) > CDbl(fpCurrMCOwed.Value) Then
      fpCurrMTOwed.BackColor = &H8080FF
      fpCurrMTPaid.BackColor = &H8080FF
      fpCurrMCPaid.BackColor = &H8080FF
      fpCurrMCOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Merchant Capital while underpaying Machine Tools. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrFEPaid.Value) > CDbl(fpCurrFEOwed.Value) Then
      fpCurrMTOwed.BackColor = &H8080FF
      fpCurrMTPaid.BackColor = &H8080FF
      fpCurrFEPaid.BackColor = &H8080FF
      fpCurrFEOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Farm Equipment while underpaying Machine Tools. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrMHPaid.Value) > CDbl(fpCurrMHOwed.Value) Then
      fpCurrMTOwed.BackColor = &H8080FF
      fpCurrMTPaid.BackColor = &H8080FF
      fpCurrMHPaid.BackColor = &H8080FF
      fpCurrMHOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Mobile Homes while underpaying Machine Tools. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrIntPaid.Value) > CDbl(fpCurrIntOwed.Value) Then
      fpCurrMTOwed.BackColor = &H8080FF
      fpCurrMTPaid.BackColor = &H8080FF
      fpCurrIntPaid.BackColor = &H8080FF
      fpCurrIntOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Interest while underpaying Machine Tools. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrPenPaid.Value) > CDbl(fpCurrPenOwed.Value) Then
      fpCurrMTOwed.BackColor = &H8080FF
      fpCurrMTPaid.BackColor = &H8080FF
      fpCurrPenPaid.BackColor = &H8080FF
      fpCurrPenOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Penalty while underpaying Machine Tools. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrOpt1Paid.Value) > CDbl(fpCurrOpt1Owed.Value) Then
      fpCurrMTOwed.BackColor = &H8080FF
      fpCurrMTPaid.BackColor = &H8080FF
      fpCurrOpt1Paid.BackColor = &H8080FF
      fpCurrOpt1Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay " + Opt1Desc$ + " while underpaying Machine Tools. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrOpt2Paid.Value) > CDbl(fpCurrOpt2Owed.Value) Then
      fpCurrMTOwed.BackColor = &H8080FF
      fpCurrMTPaid.BackColor = &H8080FF
      fpCurrOpt2Paid.BackColor = &H8080FF
      fpCurrOpt2Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay " + Opt2Desc$ + " while underpaying Machine Tools. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrOpt3Paid.Value) > CDbl(fpCurrOpt3Owed.Value) Then
      fpCurrMTOwed.BackColor = &H8080FF
      fpCurrMTPaid.BackColor = &H8080FF
      fpCurrOpt3Paid.BackColor = &H8080FF
      fpCurrOpt3Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay " + Opt3Desc$ + " while underpaying Machine Tools. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    End If
  End If
  
MTOK:
      
  If CDbl(fpCurrPersPaid.Value) = CDbl(fpCurrPersOwed.Value) Then
    GoTo PrincOK
  ElseIf CDbl(fpCurrPersOwed.Value) > CDbl(fpCurrPersPaid.Value) Then
    If CDbl(fpCurrMTPaid.Value) > CDbl(fpCurrMTOwed.Value) Then
      fpCurrMTOwed.BackColor = &H8080FF
      fpCurrMTPaid.BackColor = &H8080FF
      fpCurrPersPaid.BackColor = &H8080FF
      fpCurrPersOwed.BackColor = &H8080FF
      If CDbl(fpCurrDisc.Value) <= 0 Then
        frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Machine Tools while underpaying Personal. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      Else
        frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Machine Tools while underpaying Personal (plus Discount). Please only allow overpayment for a revenue if all other revenue obligations have been met."
      End If
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrMCPaid.Value) > CDbl(fpCurrMCOwed.Value) Then
      fpCurrPersOwed.BackColor = &H8080FF
      fpCurrPersPaid.BackColor = &H8080FF
      fpCurrMCPaid.BackColor = &H8080FF
      fpCurrMCOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Merchant Capital while underpaying Personal. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrFEPaid.Value) > CDbl(fpCurrFEOwed.Value) Then
      fpCurrPersOwed.BackColor = &H8080FF
      fpCurrPersPaid.BackColor = &H8080FF
      fpCurrFEPaid.BackColor = &H8080FF
      fpCurrFEOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Farm Equipment while underpaying Personal. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrMHPaid.Value) > CDbl(fpCurrMHOwed.Value) Then
      fpCurrPersOwed.BackColor = &H8080FF
      fpCurrPersPaid.BackColor = &H8080FF
      fpCurrMHPaid.BackColor = &H8080FF
      fpCurrMHOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Mobile Homes while underpaying Personal. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrIntPaid.Value) > CDbl(fpCurrIntOwed.Value) Then
      fpCurrPersOwed.BackColor = &H8080FF
      fpCurrPersPaid.BackColor = &H8080FF
      fpCurrIntPaid.BackColor = &H8080FF
      fpCurrIntOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Interest while underpaying Personal. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrPenPaid.Value) > CDbl(fpCurrPenOwed.Value) Then
      fpCurrPersOwed.BackColor = &H8080FF
      fpCurrPersPaid.BackColor = &H8080FF
      fpCurrPenPaid.BackColor = &H8080FF
      fpCurrPenOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Penalty while underpaying Personal. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrOpt1Paid.Value) > CDbl(fpCurrOpt1Owed.Value) Then
      fpCurrPersOwed.BackColor = &H8080FF
      fpCurrPersPaid.BackColor = &H8080FF
      fpCurrOpt1Paid.BackColor = &H8080FF
      fpCurrOpt1Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay " + Opt1Desc$ + " while underpaying Personal. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrOpt2Paid.Value) > CDbl(fpCurrOpt2Owed.Value) Then
      fpCurrPersOwed.BackColor = &H8080FF
      fpCurrPersPaid.BackColor = &H8080FF
      fpCurrOpt2Paid.BackColor = &H8080FF
      fpCurrOpt2Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay " + Opt2Desc$ + " while underpaying Personal. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrOpt3Paid.Value) > CDbl(fpCurrOpt3Owed.Value) Then
      fpCurrPersOwed.BackColor = &H8080FF
      fpCurrPersPaid.BackColor = &H8080FF
      fpCurrOpt3Paid.BackColor = &H8080FF
      fpCurrOpt3Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay " + Opt3Desc$ + " while underpaying Personal. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    End If
  End If
PrincOK:

  If CDbl(fpCurrMCPaid.Value) = CDbl(fpCurrMCOwed.Value) Then
    GoTo AdvColOK
  ElseIf CDbl(fpCurrMCOwed.Value) > CDbl(fpCurrMCPaid.Value) Then
    If CDbl(fpCurrMTPaid.Value) > CDbl(fpCurrMTOwed.Value) Then
      fpCurrMCOwed.BackColor = &H8080FF
      fpCurrMCPaid.BackColor = &H8080FF
      fpCurrMTPaid.BackColor = &H8080FF
      fpCurrMTOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Machine Tools while underpaying Merchant Capital. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrPersPaid.Value) > CDbl(fpCurrPersOwed.Value) Then
      fpCurrMCPaid.BackColor = &H8080FF
      fpCurrMCOwed.BackColor = &H8080FF
      fpCurrPersOwed.BackColor = &H8080FF
      fpCurrPersPaid.BackColor = &H8080FF
      If CDbl(fpCurrDisc.Value) <= 0 Then
        frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Personal while underpaying Merchant Capital. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      Else
        frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Personal (plus Discount) while underpaying Merchant Capital. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      End If
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrFEPaid.Value) > CDbl(fpCurrFEOwed.Value) Then
      fpCurrMCPaid.BackColor = &H8080FF
      fpCurrMCOwed.BackColor = &H8080FF
      fpCurrFEPaid.BackColor = &H8080FF
      fpCurrFEOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Farm Equipment while underpaying Merchant Capital. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrMHPaid.Value) > CDbl(fpCurrMHOwed.Value) Then
      fpCurrMCPaid.BackColor = &H8080FF
      fpCurrMCOwed.BackColor = &H8080FF
      fpCurrMHPaid.BackColor = &H8080FF
      fpCurrMHOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Mobile Homes while underpaying Merchant Capital. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrIntPaid.Value) > CDbl(fpCurrIntOwed.Value) Then
      fpCurrMCPaid.BackColor = &H8080FF
      fpCurrMCOwed.BackColor = &H8080FF
      fpCurrIntPaid.BackColor = &H8080FF
      fpCurrIntOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Interest while underpaying Merchant Capital. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrPenPaid.Value) > CDbl(fpCurrPenOwed.Value) Then
      fpCurrMCPaid.BackColor = &H8080FF
      fpCurrMCOwed.BackColor = &H8080FF
      fpCurrPenPaid.BackColor = &H8080FF
      fpCurrPenOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Penalty while underpaying Merchant Capital. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrOpt1Paid.Value) > CDbl(fpCurrOpt1Owed.Value) Then
      fpCurrMCOwed.BackColor = &H8080FF
      fpCurrMCPaid.BackColor = &H8080FF
      fpCurrOpt1Paid.BackColor = &H8080FF
      fpCurrOpt1Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay " + Opt1Desc$ + " while underpaying Merchant Capital. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrOpt2Paid.Value) > CDbl(fpCurrOpt2Owed.Value) Then
      fpCurrMCOwed.BackColor = &H8080FF
      fpCurrMCPaid.BackColor = &H8080FF
      fpCurrOpt1Paid.BackColor = &H8080FF
      fpCurrOpt1Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay " + Opt2Desc$ + " while underpaying Merchant Capital. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrOpt3Paid.Value) > CDbl(fpCurrOpt3Owed.Value) Then
      fpCurrMCOwed.BackColor = &H8080FF
      fpCurrMCPaid.BackColor = &H8080FF
      fpCurrOpt1Paid.BackColor = &H8080FF
      fpCurrOpt1Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay " + Opt3Desc$ + " while underpaying Merchant Capital. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    End If
  End If
AdvColOK:

  If CDbl(fpCurrFEPaid.Value) = CDbl(fpCurrFEOwed.Value) Then
    GoTo FarmOK
  ElseIf CDbl(fpCurrFEOwed.Value) > CDbl(fpCurrFEPaid.Value) Then
    If CDbl(fpCurrMTPaid.Value) > CDbl(fpCurrMTOwed.Value) Then
      fpCurrFEOwed.BackColor = &H8080FF
      fpCurrFEPaid.BackColor = &H8080FF
      fpCurrMTPaid.BackColor = &H8080FF
      fpCurrMTOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Machine Tools while underpaying Farm Equipment. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrPersPaid.Value) > CDbl(fpCurrPersOwed.Value) Then
      fpCurrFEOwed.BackColor = &H8080FF
      fpCurrFEPaid.BackColor = &H8080FF
      fpCurrPersOwed.BackColor = &H8080FF
      fpCurrPersPaid.BackColor = &H8080FF
      If CDbl(fpCurrDisc.Value) <= 0 Then
        frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Personal while underpaying Farm Equipment. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      Else
        frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Personal (and Discount) while underpaying Farm Equipment. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      End If
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrMCPaid.Value) > CDbl(fpCurrMCOwed.Value) Then
      fpCurrFEOwed.BackColor = &H8080FF
      fpCurrFEPaid.BackColor = &H8080FF
      fpCurrMCPaid.BackColor = &H8080FF
      fpCurrMCOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Merchant Capital while underpaying Farm Equipment. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrMHPaid.Value) > CDbl(fpCurrMHOwed.Value) Then
      fpCurrFEOwed.BackColor = &H8080FF
      fpCurrFEPaid.BackColor = &H8080FF
      fpCurrMHPaid.BackColor = &H8080FF
      fpCurrMHOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Mobile Homes while underpaying Farm Equipment. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrIntPaid.Value) > CDbl(fpCurrIntOwed.Value) Then
      fpCurrFEOwed.BackColor = &H8080FF
      fpCurrFEPaid.BackColor = &H8080FF
      fpCurrIntPaid.BackColor = &H8080FF
      fpCurrIntOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Interest while underpaying Farm Equipment. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrPenPaid.Value) > CDbl(fpCurrPenOwed.Value) Then
      fpCurrFEOwed.BackColor = &H8080FF
      fpCurrFEPaid.BackColor = &H8080FF
      fpCurrPenPaid.BackColor = &H8080FF
      fpCurrPenOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Penalty while underpaying Farm Equipment. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrOpt1Paid.Value) > CDbl(fpCurrOpt1Owed.Value) Then
      fpCurrFEOwed.BackColor = &H8080FF
      fpCurrFEPaid.BackColor = &H8080FF
      fpCurrOpt1Paid.BackColor = &H8080FF
      fpCurrOpt1Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay " + Opt1Desc$ + " while underpaying Farm Equipment. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrOpt2Paid.Value) > CDbl(fpCurrOpt2Owed.Value) Then
      fpCurrFEOwed.BackColor = &H8080FF
      fpCurrFEPaid.BackColor = &H8080FF
      fpCurrOpt2Paid.BackColor = &H8080FF
      fpCurrOpt2Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay " + Opt2Desc$ + " while underpaying Farm Equipment. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrOpt3Paid.Value) > CDbl(fpCurrOpt3Owed.Value) Then
      fpCurrFEOwed.BackColor = &H8080FF
      fpCurrFEPaid.BackColor = &H8080FF
      fpCurrOpt3Paid.BackColor = &H8080FF
      fpCurrOpt3Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay " + Opt3Desc$ + " while underpaying Farm Equipment. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    End If
  End If
FarmOK:

  If CDbl(fpCurrMHPaid.Value) = CDbl(fpCurrMHOwed.Value) Then
    GoTo MHOK
  ElseIf CDbl(fpCurrMHOwed.Value) > CDbl(fpCurrMHPaid.Value) Then
    If CDbl(fpCurrMTPaid.Value) > CDbl(fpCurrMTOwed.Value) Then
      fpCurrMHOwed.BackColor = &H8080FF
      fpCurrMHPaid.BackColor = &H8080FF
      fpCurrMTPaid.BackColor = &H8080FF
      fpCurrMTOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Machine Tools while underpaying Mobile Homes. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrPersPaid.Value) > CDbl(fpCurrPersOwed.Value) Then
      fpCurrMHOwed.BackColor = &H8080FF
      fpCurrMHPaid.BackColor = &H8080FF
      fpCurrPersOwed.BackColor = &H8080FF
      fpCurrPersPaid.BackColor = &H8080FF
      If CDbl(fpCurrDisc.Value) <= 0 Then
        frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Personal while underpaying Mobile Homes. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      Else
        frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Personal (plus Discount) while underpaying Mobile Homes. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      End If
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrMCPaid.Value) > CDbl(fpCurrMCOwed.Value) Then
      fpCurrMHOwed.BackColor = &H8080FF
      fpCurrMHPaid.BackColor = &H8080FF
      fpCurrMCPaid.BackColor = &H8080FF
      fpCurrMCOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Merchant Capital while underpaying Mobile Homes. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrFEPaid.Value) > CDbl(fpCurrFEOwed.Value) Then
      fpCurrFEOwed.BackColor = &H8080FF
      fpCurrFEPaid.BackColor = &H8080FF
      fpCurrMHPaid.BackColor = &H8080FF
      fpCurrMHOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Farm Equipment while underpaying Mobile Homes. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrIntPaid.Value) > CDbl(fpCurrIntOwed.Value) Then
      fpCurrMHOwed.BackColor = &H8080FF
      fpCurrMHPaid.BackColor = &H8080FF
      fpCurrIntPaid.BackColor = &H8080FF
      fpCurrIntOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Interest while underpaying Mobile Homes. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrPenPaid.Value) > CDbl(fpCurrPenOwed.Value) Then
      fpCurrMHOwed.BackColor = &H8080FF
      fpCurrMHPaid.BackColor = &H8080FF
      fpCurrPenPaid.BackColor = &H8080FF
      fpCurrPenOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Penalty while underpaying Mobile Homes. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrOpt1Paid.Value) > CDbl(fpCurrOpt1Owed.Value) Then
      fpCurrMHOwed.BackColor = &H8080FF
      fpCurrMHPaid.BackColor = &H8080FF
      fpCurrOpt1Paid.BackColor = &H8080FF
      fpCurrOpt1Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay " + Opt1Desc$ + " while underpaying Mobile Homes. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrOpt2Paid.Value) > CDbl(fpCurrOpt2Owed.Value) Then
      fpCurrMHOwed.BackColor = &H8080FF
      fpCurrMHPaid.BackColor = &H8080FF
      fpCurrOpt2Paid.BackColor = &H8080FF
      fpCurrOpt2Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay " + Opt2Desc$ + " while underpaying Mobile Homes. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrOpt3Paid.Value) > CDbl(fpCurrOpt3Owed.Value) Then
      fpCurrMHOwed.BackColor = &H8080FF
      fpCurrMHPaid.BackColor = &H8080FF
      fpCurrOpt3Paid.BackColor = &H8080FF
      fpCurrOpt3Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay " + Opt3Desc$ + " while underpaying Mobile Homes. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    End If
  End If
MHOK:

  If CDbl(fpCurrIntPaid.Value) = CDbl(fpCurrIntOwed.Value) Then
    GoTo IntOK
  ElseIf CDbl(fpCurrIntOwed.Value) > CDbl(fpCurrIntPaid.Value) Then
    If CDbl(fpCurrMTPaid.Value) > CDbl(fpCurrMTOwed.Value) Then
      fpCurrIntOwed.BackColor = &H8080FF
      fpCurrIntPaid.BackColor = &H8080FF
      fpCurrMTPaid.BackColor = &H8080FF
      fpCurrMTOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Machine Tools while underpaying Interest. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrPersPaid.Value) > CDbl(fpCurrPersOwed.Value) Then
      fpCurrIntOwed.BackColor = &H8080FF
      fpCurrIntPaid.BackColor = &H8080FF
      fpCurrPersOwed.BackColor = &H8080FF
      fpCurrPersPaid.BackColor = &H8080FF
      If CDbl(fpCurrDisc.Value) <= 0 Then
        frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Personal while underpaying Interest. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      Else
        frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Personal (plus Discount) while underpaying Interest. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      End If
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrMCPaid.Value) > CDbl(fpCurrMCOwed.Value) Then
      fpCurrIntOwed.BackColor = &H8080FF
      fpCurrIntPaid.BackColor = &H8080FF
      fpCurrMCPaid.BackColor = &H8080FF
      fpCurrMCOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Merchant Capital while underpaying Interest. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrFEPaid.Value) > CDbl(fpCurrFEOwed.Value) Then
      fpCurrFEOwed.BackColor = &H8080FF
      fpCurrFEPaid.BackColor = &H8080FF
      fpCurrIntPaid.BackColor = &H8080FF
      fpCurrIntOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Farm Equipment while underpaying Interest. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrMHPaid.Value) > CDbl(fpCurrMHOwed.Value) Then
      fpCurrMHOwed.BackColor = &H8080FF
      fpCurrMHPaid.BackColor = &H8080FF
      fpCurrIntPaid.BackColor = &H8080FF
      fpCurrIntOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Mobile Homes while underpaying Interest. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrPenPaid.Value) > CDbl(fpCurrPenOwed.Value) Then
      fpCurrIntOwed.BackColor = &H8080FF
      fpCurrIntPaid.BackColor = &H8080FF
      fpCurrPenPaid.BackColor = &H8080FF
      fpCurrPenOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Penalty while underpaying Interest. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrOpt1Paid.Value) > CDbl(fpCurrOpt1Owed.Value) Then
      fpCurrIntOwed.BackColor = &H8080FF
      fpCurrIntPaid.BackColor = &H8080FF
      fpCurrOpt1Paid.BackColor = &H8080FF
      fpCurrOpt1Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay " + Opt1Desc$ + " while underpaying Interest. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrOpt2Paid.Value) > CDbl(fpCurrOpt2Owed.Value) Then
      fpCurrIntOwed.BackColor = &H8080FF
      fpCurrIntPaid.BackColor = &H8080FF
      fpCurrOpt2Paid.BackColor = &H8080FF
      fpCurrOpt2Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay " + Opt2Desc$ + " while underpaying Interest. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrOpt3Paid.Value) > CDbl(fpCurrOpt3Owed.Value) Then
      fpCurrIntOwed.BackColor = &H8080FF
      fpCurrIntPaid.BackColor = &H8080FF
      fpCurrOpt3Paid.BackColor = &H8080FF
      fpCurrOpt3Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay " + Opt3Desc$ + " while underpaying Interest. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    End If
  End If
IntOK:

  If CDbl(fpCurrPenPaid.Value) = CDbl(fpCurrPenOwed.Value) Then
    GoTo PenOK
  ElseIf CDbl(fpCurrPenOwed.Value) > CDbl(fpCurrPenPaid.Value) Then
    If CDbl(fpCurrMTPaid.Value) > CDbl(fpCurrMTOwed.Value) Then
      fpCurrPenOwed.BackColor = &H8080FF
      fpCurrPenPaid.BackColor = &H8080FF
      fpCurrMTPaid.BackColor = &H8080FF
      fpCurrMTOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Machine Tools while underpaying Penalty. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrPersPaid.Value) > CDbl(fpCurrPersOwed.Value) Then
      fpCurrPenOwed.BackColor = &H8080FF
      fpCurrPenPaid.BackColor = &H8080FF
      fpCurrPersOwed.BackColor = &H8080FF
      fpCurrPersPaid.BackColor = &H8080FF
      If CDbl(fpCurrDisc.Value) <= 0 Then
        frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Personal while underpaying Penalty. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      Else
        frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Personal (plus Discount) while underpaying Penalty. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      End If
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrMCPaid.Value) > CDbl(fpCurrMCOwed.Value) Then
      fpCurrPenOwed.BackColor = &H8080FF
      fpCurrPenPaid.BackColor = &H8080FF
      fpCurrMCPaid.BackColor = &H8080FF
      fpCurrMCOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Merchant Capital while underpaying Penalty. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrFEPaid.Value) > CDbl(fpCurrFEOwed.Value) Then
      fpCurrFEOwed.BackColor = &H8080FF
      fpCurrFEPaid.BackColor = &H8080FF
      fpCurrPenPaid.BackColor = &H8080FF
      fpCurrPenOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Farm Equipment while underpaying Penalty. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrMHPaid.Value) > CDbl(fpCurrMHOwed.Value) Then
      fpCurrMHOwed.BackColor = &H8080FF
      fpCurrMHPaid.BackColor = &H8080FF
      fpCurrPenPaid.BackColor = &H8080FF
      fpCurrPenOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Mobile Homes while underpaying Penalty. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrIntPaid.Value) > CDbl(fpCurrIntOwed.Value) Then
      fpCurrIntOwed.BackColor = &H8080FF
      fpCurrIntPaid.BackColor = &H8080FF
      fpCurrPenPaid.BackColor = &H8080FF
      fpCurrPenOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Interest while underpaying Penalty. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrOpt1Paid.Value) > CDbl(fpCurrOpt1Owed.Value) Then
      fpCurrPenOwed.BackColor = &H8080FF
      fpCurrPenPaid.BackColor = &H8080FF
      fpCurrOpt1Paid.BackColor = &H8080FF
      fpCurrOpt1Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay " + Opt1Desc$ + " while underpaying Penalty. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrOpt2Paid.Value) > CDbl(fpCurrOpt2Owed.Value) Then
      fpCurrPenOwed.BackColor = &H8080FF
      fpCurrPenPaid.BackColor = &H8080FF
      fpCurrOpt2Paid.BackColor = &H8080FF
      fpCurrOpt2Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay " + Opt2Desc$ + " while underpaying Penalty. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrOpt3Paid.Value) > CDbl(fpCurrOpt3Owed.Value) Then
      fpCurrPenOwed.BackColor = &H8080FF
      fpCurrPenPaid.BackColor = &H8080FF
      fpCurrOpt3Paid.BackColor = &H8080FF
      fpCurrOpt3Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay " + Opt3Desc$ + " while underpaying Penalty. Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    End If
  End If
PenOK:

  If CDbl(fpCurrOpt1Paid.Value) = CDbl(fpCurrOpt1Owed.Value) Then
    GoTo Opt1OK
  ElseIf CDbl(fpCurrOpt1Owed.Value) > CDbl(fpCurrOpt1Paid.Value) Then
    If CDbl(fpCurrMTPaid.Value) > CDbl(fpCurrMTOwed.Value) Then
      fpCurrOpt1Owed.BackColor = &H8080FF
      fpCurrOpt1Paid.BackColor = &H8080FF
      fpCurrMTPaid.BackColor = &H8080FF
      fpCurrMTOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Machine Tools while underpaying " + Opt1Desc + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrPersPaid.Value) > CDbl(fpCurrPersOwed.Value) Then
      fpCurrOpt1Owed.BackColor = &H8080FF
      fpCurrOpt1Paid.BackColor = &H8080FF
      fpCurrPersOwed.BackColor = &H8080FF
      fpCurrPersPaid.BackColor = &H8080FF
      If CDbl(fpCurrDisc.Value) <= 0 Then
        frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Personal while underpaying " + Opt1Desc + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      Else
        frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Personal (plus Discount) while underpaying " + Opt1Desc + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      End If
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrMCPaid.Value) > CDbl(fpCurrMCOwed.Value) Then
      fpCurrOpt1Owed.BackColor = &H8080FF
      fpCurrOpt1Paid.BackColor = &H8080FF
      fpCurrMCPaid.BackColor = &H8080FF
      fpCurrMCOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Merchant Capital while underpaying " + Opt1Desc + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrFEPaid.Value) > CDbl(fpCurrFEOwed.Value) Then
      fpCurrFEOwed.BackColor = &H8080FF
      fpCurrFEPaid.BackColor = &H8080FF
      fpCurrOpt1Paid.BackColor = &H8080FF
      fpCurrOpt1Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Farm Equipment while underpaying " + Opt1Desc + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrMHPaid.Value) > CDbl(fpCurrMHOwed.Value) Then
      fpCurrMHOwed.BackColor = &H8080FF
      fpCurrMHPaid.BackColor = &H8080FF
      fpCurrOpt1Paid.BackColor = &H8080FF
      fpCurrOpt1Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Mobile Homes while underpaying " + Opt1Desc + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrIntPaid.Value) > CDbl(fpCurrIntOwed.Value) Then
      fpCurrIntOwed.BackColor = &H8080FF
      fpCurrIntPaid.BackColor = &H8080FF
      fpCurrOpt1Paid.BackColor = &H8080FF
      fpCurrOpt1Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Interest while underpaying " + Opt1Desc + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrPenPaid.Value) > CDbl(fpCurrPenOwed.Value) Then
      fpCurrOpt1Owed.BackColor = &H8080FF
      fpCurrOpt1Paid.BackColor = &H8080FF
      fpCurrPenPaid.BackColor = &H8080FF
      fpCurrPenOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Penalty while underpaying " + Opt1Desc + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrOpt2Paid.Value) > CDbl(fpCurrOpt2Owed.Value) Then
      fpCurrOpt1Owed.BackColor = &H8080FF
      fpCurrOpt1Paid.BackColor = &H8080FF
      fpCurrOpt2Paid.BackColor = &H8080FF
      fpCurrOpt2Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay " + Opt2Desc$ + " while underpaying " + Opt1Desc + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrOpt3Paid.Value) > CDbl(fpCurrOpt3Owed.Value) Then
      fpCurrOpt1Owed.BackColor = &H8080FF
      fpCurrOpt1Paid.BackColor = &H8080FF
      fpCurrOpt3Paid.BackColor = &H8080FF
      fpCurrOpt3Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay " + Opt3Desc$ + " while underpaying " + Opt1Desc + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    End If
  End If
Opt1OK:

  If CDbl(fpCurrOpt2Paid.Value) = CDbl(fpCurrOpt2Owed.Value) Then
    GoTo Opt2OK
  ElseIf CDbl(fpCurrOpt2Owed.Value) > CDbl(fpCurrOpt2Paid.Value) Then
    If CDbl(fpCurrMTPaid.Value) > CDbl(fpCurrMTOwed.Value) Then
      fpCurrOpt2Owed.BackColor = &H8080FF
      fpCurrOpt2Paid.BackColor = &H8080FF
      fpCurrMTPaid.BackColor = &H8080FF
      fpCurrMTOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Machine Tools while underpaying " + Opt2Desc + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrPersPaid.Value) > CDbl(fpCurrPersOwed.Value) Then
      fpCurrOpt2Owed.BackColor = &H8080FF
      fpCurrOpt2Paid.BackColor = &H8080FF
      fpCurrPersOwed.BackColor = &H8080FF
      fpCurrPersPaid.BackColor = &H8080FF
      If CDbl(fpCurrDisc.Value) <= 0 Then
        frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Personal while underpaying " + Opt2Desc + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      Else
        frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Personal (plus Discount) while underpaying " + Opt2Desc + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      End If
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrMCPaid.Value) > CDbl(fpCurrMCOwed.Value) Then
      fpCurrOpt2Owed.BackColor = &H8080FF
      fpCurrOpt2Paid.BackColor = &H8080FF
      fpCurrMCPaid.BackColor = &H8080FF
      fpCurrMCOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Merchant Capital while underpaying " + Opt2Desc + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrFEPaid.Value) > CDbl(fpCurrFEOwed.Value) Then
      fpCurrFEOwed.BackColor = &H8080FF
      fpCurrFEPaid.BackColor = &H8080FF
      fpCurrOpt2Paid.BackColor = &H8080FF
      fpCurrOpt2Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Farm Equipment while underpaying " + Opt2Desc + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrMHPaid.Value) > CDbl(fpCurrMHOwed.Value) Then
      fpCurrMHOwed.BackColor = &H8080FF
      fpCurrMHPaid.BackColor = &H8080FF
      fpCurrOpt2Paid.BackColor = &H8080FF
      fpCurrOpt2Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Mobile Homes while underpaying " + Opt2Desc + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrIntPaid.Value) > CDbl(fpCurrIntOwed.Value) Then
      fpCurrIntOwed.BackColor = &H8080FF
      fpCurrIntPaid.BackColor = &H8080FF
      fpCurrOpt2Paid.BackColor = &H8080FF
      fpCurrOpt2Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Interest while underpaying " + Opt2Desc + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrPenPaid.Value) > CDbl(fpCurrPenOwed.Value) Then
      fpCurrOpt2Owed.BackColor = &H8080FF
      fpCurrOpt2Paid.BackColor = &H8080FF
      fpCurrPenPaid.BackColor = &H8080FF
      fpCurrPenOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Penalty while underpaying " + Opt2Desc + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrOpt2Paid.Value) > CDbl(fpCurrOpt2Owed.Value) Then
      fpCurrOpt1Owed.BackColor = &H8080FF
      fpCurrOpt1Paid.BackColor = &H8080FF
      fpCurrOpt2Paid.BackColor = &H8080FF
      fpCurrOpt2Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay " + Opt1Desc$ + " while underpaying " + Opt2Desc + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrOpt3Paid.Value) > CDbl(fpCurrOpt3Owed.Value) Then
      fpCurrOpt2Owed.BackColor = &H8080FF
      fpCurrOpt2Paid.BackColor = &H8080FF
      fpCurrOpt3Paid.BackColor = &H8080FF
      fpCurrOpt3Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay " + Opt3Desc$ + " while underpaying " + Opt2Desc + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    End If
  End If
Opt2OK:

  If CDbl(fpCurrOpt3Paid.Value) = CDbl(fpCurrOpt3Owed.Value) Then
    GoTo Opt3OK
  ElseIf CDbl(fpCurrOpt3Owed.Value) > CDbl(fpCurrOpt3Paid.Value) Then
    If CDbl(fpCurrMTPaid.Value) > CDbl(fpCurrMTOwed.Value) Then
      fpCurrOpt3Owed.BackColor = &H8080FF
      fpCurrOpt3Paid.BackColor = &H8080FF
      fpCurrMTPaid.BackColor = &H8080FF
      fpCurrMTOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Machine Tools while underpaying " + Opt3Desc + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrPersPaid.Value) > CDbl(fpCurrPersOwed.Value) Then
      fpCurrOpt3Owed.BackColor = &H8080FF
      fpCurrOpt3Paid.BackColor = &H8080FF
      fpCurrPersOwed.BackColor = &H8080FF
      fpCurrPersPaid.BackColor = &H8080FF
      If CDbl(fpCurrDisc.Value) <= 0 Then
        frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Personal while underpaying " + Opt3Desc + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      Else
        frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Personal (plus Discount) while underpaying " + Opt3Desc + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      End If
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrMCPaid.Value) > CDbl(fpCurrMCOwed.Value) Then
      fpCurrOpt3Owed.BackColor = &H8080FF
      fpCurrOpt3Paid.BackColor = &H8080FF
      fpCurrMCPaid.BackColor = &H8080FF
      fpCurrMCOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Merchant Capital while underpaying " + Opt3Desc + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrFEPaid.Value) > CDbl(fpCurrFEOwed.Value) Then
      fpCurrFEOwed.BackColor = &H8080FF
      fpCurrFEPaid.BackColor = &H8080FF
      fpCurrOpt3Paid.BackColor = &H8080FF
      fpCurrOpt3Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Farm Equipment while underpaying " + Opt3Desc + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrMHPaid.Value) > CDbl(fpCurrMHOwed.Value) Then
      fpCurrMHOwed.BackColor = &H8080FF
      fpCurrMHPaid.BackColor = &H8080FF
      fpCurrOpt3Paid.BackColor = &H8080FF
      fpCurrOpt3Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Mobile Homes while underpaying " + Opt3Desc + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrIntPaid.Value) > CDbl(fpCurrIntOwed.Value) Then
      fpCurrIntOwed.BackColor = &H8080FF
      fpCurrIntPaid.BackColor = &H8080FF
      fpCurrOpt3Paid.BackColor = &H8080FF
      fpCurrOpt3Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Interest while underpaying " + Opt3Desc + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrPenPaid.Value) > CDbl(fpCurrPenOwed.Value) Then
      fpCurrOpt3Owed.BackColor = &H8080FF
      fpCurrOpt3Paid.BackColor = &H8080FF
      fpCurrPenPaid.BackColor = &H8080FF
      fpCurrPenOwed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay Penalty while underpaying " + Opt3Desc + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrOpt1Paid.Value) > CDbl(fpCurrOpt1Owed.Value) Then
      fpCurrOpt1Owed.BackColor = &H8080FF
      fpCurrOpt1Paid.BackColor = &H8080FF
      fpCurrOpt3Paid.BackColor = &H8080FF
      fpCurrOpt3Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay " + Opt1Desc$ + " while underpaying " + Opt3Desc + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    ElseIf CDbl(fpCurrOpt2Paid.Value) > CDbl(fpCurrOpt2Owed.Value) Then
      fpCurrOpt2Owed.BackColor = &H8080FF
      fpCurrOpt2Paid.BackColor = &H8080FF
      fpCurrOpt3Paid.BackColor = &H8080FF
      fpCurrOpt3Owed.BackColor = &H8080FF
      frmVATaxMsg.Label1.Caption = "The customer is attempting to overpay " + Opt2Desc$ + " while underpaying " + Opt3Desc + ". Please only allow overpayment for a revenue if all other revenue obligations have been met."
      frmVATaxMsg.Label1.Top = 600
      frmVATaxMsg.Show vbModal
      CheckOverPay = True
      Exit Function
    End If
  End If
Opt3OK:

  Exit Function
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersPaymentEntry", "CheckOverPay", Erl)
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
  '  ClearInUse PWcnt
  '  CMTerminate

End Function

Private Sub PrintReceipt()
  Dim PayRec As VATaxPaymentRecType
  Dim PHandle As Integer
  Dim RHandle As Integer
  Dim Oper$
  Dim MasterRec As VATaxMasterType
  Dim MHandle As Integer
  Dim TownName$
  Dim PostDate$
  Dim RptHandle As Integer, LPTHandle As Integer
  Dim RptA As Integer, LPTA As Integer, ToPrintA As String
  Dim ToPrint As String, CopyLoop As Integer, DefPrinter As String
  Dim PayRecpName$
  Dim RHandle2 As Integer, PayRecpName2 As String, RptHandle2 As Integer
  On Local Error GoTo ERRORSTUFF
  Oper$ = CStr(OperNum)
  PayRecpName2$ = UBPath$ + "CMXVLD" + Oper$ + ".Rpt"
  If PrnRecp = False And PrnVali = True Then GoTo Validationthing
  
  OpenVATaxSetUpFile MHandle
  Get MHandle, 1, MasterRec
  Close MHandle
  
  OpenTempPersPayFile PHandle, OperNum
  Get PHandle, 1, PayRec
  Close PHandle
  
  TownName$ = QPTrim$(MasterRec.Name)
  PostDate$ = MakeRegDate(PayRec.payDate)
  PayRecpName$ = "C:\CPWork\CMXRCP" + Oper$ + ".RPT"
  RHandle = FreeFile
  Open PayRecpName$ For Output As RHandle
  'Print #RHandle, Chr$(27); "p"; Chr$(0); Chr$(25); Chr$(250)
  'Print #RHandle, Chr$(7)
  Print #RHandle, TownName$
  Print #RHandle, "CM-TAX PAYMENT"
  Print #RHandle, "Date: "; PostDate$
  Print #RHandle,
  Print #RHandle, "CUSTOMER NAME & DESC. OF PAYMENT"
  Print #RHandle, QPTrim$(PayRec.CustName)
  Print #RHandle, QPTrim$(PayRec.CustAddr)
  Print #RHandle, QPTrim$(PayRec.Desc)
  Print #RHandle, "Acct. No. "; PayRec.CustAcct
  Print #RHandle,
  Print #RHandle, "Total Owed: "; QPTrim$(Using("$##,##0.00", PayRec.AmtOwed))
  Print #RHandle, "Total Paid: "; QPTrim$(Using("$##,##0.00", PayRec.AmtPaid))
  Print #RHandle, "  Discount: "; QPTrim$(Using("$##,##0.00", PayRec.DiscAmt))
  Print #RHandle, "Change Due: "; QPTrim$(Using("$##,##0.00", PayRec.Change))
  Print #RHandle,
  Print #RHandle,
  Print #RHandle, "Operator: "; CStr(OperNum)
  Print #RHandle, '"Receipt#: "; Using("$##,##0.00", FileSize("TAXCPR" + Oper$ + ".DAT") \ TaxPayRecLen)
  Print #RHandle,
  Print #RHandle, "       T H A N K   Y O U !"
  Print #RHandle,
  Print #RHandle,
  Print #RHandle,
  Print #RHandle,
  Print #RHandle,
  Print #RHandle,

  Close RHandle
  If CntrlDef = 1 Then
    Call fpcmdDrawer_Click
  End If
10:
  DefPrinter = RecpPort
20:
  
  For CopyLoop = 1 To 1 'Copies
    LPTHandle = FreeFile
    Open DefPrinter For Output As LPTHandle
    RptHandle = FreeFile
30:
    Open PayRecpName$ For Input As RptHandle
40:
    Do
     ' If frmVATaxPrint.cmdCancel = False Then
45:
        Line Input #RptHandle, ToPrint$
        
        ToPrint$ = RTrim$(ToPrint$)
        Print #LPTHandle, ToPrint$
   '   Else
50:
    '    Exit Do
        'Printer.EndDoc
   '   End If
  Loop Until eof(RptHandle)
60:
  Close RptHandle
62:
  Close LPTHandle
65:
  Next CopyLoop
68:
  Printer.EndDoc
  
69:
Validationthing:
  If QPTrim(PayRec.TenderTY) = "CHECK" Or QPTrim(PayRec.TenderTY) = "CASH AND CHECK" Then
   If RctValidate Then
     RHandle2 = FreeFile
     Open PayRecpName2$ For Output As RHandle2
     Print #LPTHandle, Chr$(27); Chr$(&H63); Chr$(&H30); Chr$(&H4)
     Print #LPTHandle, Chr$(13); Chr$(10)
     Print #RHandle2, TownName$
     Print #RHandle2, "FOR DEPOSIT ONLY"
     Print #RHandle2, "Acct. No. "; PayRec.CustAcct
     Print #RHandle2, "Date: "; PostDate$
     Print #RHandle2, "Time: "; Time
     Print #RHandle2,
     Print #LPTHandle, Chr$(12)
     Close RHandle2
     LPTHandle = FreeFile
     Open DefPrinter For Output As LPTHandle
     RptHandle2 = FreeFile
     Open PayRecpName2$ For Input As RptHandle2
     Do
     ' If frmVATaxPrint.cmdCancel = False Then
         Line Input #RptHandle2, ToPrint$
         ToPrint$ = RTrim$(ToPrint$)
         Print #LPTHandle, ToPrint$
    '   Else
    '     Exit Do
    '   End If
     Loop Until eof(RptHandle2)
     Close RptHandle2
     Close LPTHandle
    Printer.EndDoc
    TXLog "Oper: " + Oper$ + " Print Validation Acct:" + Str(PayRec.CustAcct)
  End If
 End If

70:
If PrnRecp = True Then
 TXLog "Oper: " + Oper$ + " Print receipt Acct:" + Str(PayRec.CustAcct)
 CMLog "Oper: " + Oper$ + " Print receipt Acct:" + Str(PayRec.CustAcct)
 KillFile PayRecpName$
 'KillFile PayFileName$
End If
80:
  Exit Sub

ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersPaymentEntry", "PrintReceipt", Erl)
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
  '  ClearInUse PWcnt
  '  CMTerminate

End Sub

Private Sub fptxtPayDate_LostFocus()
  Dim WhichRec() As Integer
  
  On Local Error GoTo ERRORSTUFF
  
  'warns user if the date changes causing an existing discount to change
  If NotFirstLoad = False Then Exit Sub
  If CDbl(fpCurrDisc.Value) > 0 And VACheck4Discounts = 0 Then
    frmVATaxMsgWOpts.Label1.Caption = "Changing the date from " + MakeRegDate(TempPayDate) + " to " + QPTrim$(fptxtPayDate.Text) + " will disqualify this customer from an existing discount. If you wish to continue with the new date which will automatically recalculate the amounts owed then press F10. Otherwise, press ESC to leave the date untouched."
    frmVATaxMsgWOpts.Label1.Top = 600
    frmVATaxMsgWOpts.cmdExit.Text = "ESC Restore Date"
    frmVATaxMsgWOpts.cmdCont.Text = "F10 New Date OK"
    frmVATaxMsgWOpts.Show vbModal
    If frmVATaxMsgWOpts.fptxtChoice.Text = "abort" Then
      Unload frmVATaxMsgWOpts
      fptxtPayDate = DefPayDate
    ElseIf frmVATaxMsgWOpts.fptxtChoice.Text = "continue" Then
      Unload frmVATaxMsgWOpts
      Call RemoveDiscount
    End If
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersPaymentEntry", "fptxtPayDate_LostFocus", Erl)
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
    'ClearInUse PWcnt
    'CMTerminate

End Sub

Private Sub ApplyDiscount()
'  Dim ColTot As Double
'  Dim x As Integer
'
'  On Error GoTo ERRORSTUFF
'
'  For x = 1 To BillCnt
'    TempBillList(x).Personal = OldRound(TempBillList(x).Personal - TempBillList(x).DiscAmt)
'  Next x
'
'  ColTot = 0
'  fpCurrDisc = MaxDisc
'  fpCurrPersPaid = OldRound(CDbl(fpCurrPersOwed.Value) - CDbl(fpCurrDisc.Value))
'  GoSub AddCol
'
'  If CDbl(fpCurrPersPaid.Value) < ThisDiscAmt Then
'    fpCurrDisc = CDbl(fpCurrPersPaid.Value)
'    fpCurrPersPaid = 0
'    Call AddUpPaidCol
'  ElseIf OldRound(CDbl(fpCurrPersPaid.Value) + CDbl(fpCurrDisc.Value)) > CDbl(fpCurrPersOwed.Value) Then
'    fpCurrPersPaid = OldRound(CDbl(fpCurrPersOwed.Value) - CDbl(fpCurrDisc.Value))
'  Else
'    If ColTot > CDbl(fpCurrTotRecd.Value) Then
'      If OldRound(ColTot - CDbl(fpCurrTotRecd.Value)) = CDbl(fpCurrDisc.Value) Then
'        fpCurrPersPaid = OldRound(CDbl(fpCurrPersPaid.Value) - CDbl(fpCurrDisc.Value))
'      End If
'    End If
'  End If
'
'  Call ReFigure
'
'  TxLog ("frmVATaxPersPaymentEntry: Customer, " + fptxtName.Text + ", is eligible for a discount of " + QPTrim$(Using$("$#,##0.00", ThisDiscAmt)) + " and the user allowed the discount to apply.")
'
'  Exit Sub
'
'AddCol:
''  ColTot = OldRound(CDbl(fpCurrDisc.Value) + CDbl(fpCurrPersPaid.Value) + CDbl(fpCurrMTPaid.Value))
'  ColTot = OldRound(CDbl(fpCurrPersPaid.Value) + CDbl(fpCurrMTPaid.Value))
'  ColTot = OldRound(ColTot# + CDbl(fpCurrMCPaid.Value) + CDbl(fpCurrFEPaid.Value))
'  ColTot = OldRound(ColTot# + CDbl(fpCurrMHPaid.Value) + CDbl(fpCurrIntPaid.Value) + CDbl(fpCurrPenPaid.Value))
'  fpCurrTotPaid = ColTot
'  fpCurrTotWDisc = OldRound(ColTot + CDbl(fpCurrDisc.Value))
'  Return
'
'ERRORSTUFF:
'   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersPaymentEntry", "ApplyDiscount", Erl)
'     Case emrExitProc:
'       Resume Proc_Exit
'     Case emrResume:
'       Resume
'     Case emrResumeNext:
'       Resume Next
'     Case Else
'      '--- Technically, this should never happen.
'       Resume Proc_Exit
'   End Select
'
'Proc_Exit:
'  '--- Cleanup code goes here...
'    Close
'    ClearInUse PWcnt
'    CMTerminate

End Sub

Private Sub Distribute(WhatsLeft As Double)
  Dim SetUpRec As VATaxMasterType
  Dim SHandle As Integer
  Dim x As Integer, y As Integer
  Dim TotRecd As Double
  Dim TaxTrans As VATaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim TransRecord&
  Dim PaidDif As Double
  Dim ThisDif As Double
  Dim TPayRec As VAPersPayListType
  Dim PayRec As VAPersPayListType
  Dim PHandle As Integer
  Dim NumOfPRecs As Integer
  Dim z As Integer
  Dim ThisPrevRec As Long
  Dim NewRec As Integer
  Dim Nextx As Integer
  Dim SmallNum As Integer
  Dim HoldNum As Long
  Dim HoldDate As Integer
  Dim Thisx As Integer
  Dim TaxMasterRec As VATaxMasterType
  Dim MHandle As Integer
  Dim ThisTaxYear As Integer
  Dim Message$
  Dim ThisBal As Double
  Dim DiscCheck As Integer
  Dim ThisPct As Double
  Dim TotPaid#
  Dim TotDisc As Double
  Dim Disc1 As Double '1/25/2007
  Dim Disc2 As Double '1/25/2007
  Dim Disc3 As Double '1/25/2007
  Dim Disc4 As Double '1/25/2007
  Dim Disc5 As Double '1/25/2007
  Dim Disc6 As Double '1/25/2007
  Dim Disc7 As Double '1/25/2007
  Dim Disc8 As Double '1/25/2007
  Dim DiscApplied As Boolean '1/25/2007
  Dim SaveAmt As Double '1/25/2007
  Dim DumpPenny As Double '1/25/2007
  
  On Local Error GoTo ERRORSTUFF
  
  If CDbl(fpCurrAmtOwed.Value) = 0 Then '8/12/05
    If VATaxMsgWOpts(800, "Since this customer does not owe any money automatic distribution will place the amount paid in the 'Prepay Amt' field. Press F10 to OK this distribution.", "F10 OK", "ESC Abort") = "abort" Then
      Unload frmVATaxMsgWOpts
      Exit Sub
    Else
      fpCurrPrePay = WhatsLeft
      fpCurrChngDue = 0
      NotFirstLoad = True 'added 6/1/06
      Call AddUpPaidCol
      Call LoadTempPayList
      Exit Sub
    End If
  End If
    If fpCurrDisc > 0 Then
    If CDbl(fpCurrAmtOwed) < CDbl(fpCurrCashPd) + CDbl(fpCurrChkChrgPd) + CDbl(fpCurrDisc) Then '1/25/07
      Call TaxMsg(900, "Overpayments are not allowed when applying discounts.")
      fpCurrDisc.SetFocus
      Exit Sub
    End If
  End If

  OpenVATaxSetUpFile MHandle
  Get MHandle, 1, TaxMasterRec
  Close MHandle
  
  ThisTaxYear = TaxMasterRec.RTaxYear
  
  fpCurrMTPaid.Value = 0
  fpCurrMCPaid.Value = 0
  fpCurrFEPaid.Value = 0
  fpCurrPersPaid.Value = 0
  fpCurrMHPaid.Value = 0
  fpCurrIntPaid.Value = 0
  fpCurrPenPaid.Value = 0
  fpCurrOpt1Paid.Value = 0
  fpCurrOpt2Paid.Value = 0
  fpCurrOpt3Paid.Value = 0
'  PrePayAmt = CDbl(fpCurrPrePay.Value)
  Nextx = 1
  SmallNum = 30000
  ReDim TransSeq(1 To BillCnt) As Long
  Do While Nextx <= BillCnt
    For x = Nextx To BillCnt
      If BillDate(x) <= SmallNum Then
        SmallNum = BillDate(x)
        Thisx = x
      End If
    Next x
    HoldNum = BillTrans(Nextx)
    HoldDate = BillDate(Nextx)
    BillTrans(Nextx) = BillTrans(Thisx)
    BillDate(Nextx) = BillDate(Thisx)
    BillTrans(Thisx) = HoldNum
    BillDate(Thisx) = HoldDate
    Nextx = Nextx + 1
    SmallNum = 30000
  Loop
  
  ReDim Preserve TempBillList(1 To BillCnt) As VAPersPayListType
  TempBillListCnt = 0
  For x = 1 To BillCnt
    TempBillList(x).MachTools = 0
    TempBillList(x).MerchCap = 0
    TempBillList(x).FarmEquip = 0
    TempBillList(x).Personal = 0
    TempBillList(x).MobHomes = 0
    TempBillList(x).Interest = 0
    TempBillList(x).Penalty = 0
    TempBillList(x).BillRec = 0
    TempBillList(x).Opt1 = 0
    TempBillList(x).Opt2 = 0
    TempBillList(x).Opt3 = 0
    TempBillList(x).CustRec = 0
    TempBillList(x).TaxYear = 0
    TempBillList(x).TotPaid = 0
  Next x
   
  ReDim PaySeq(1 To BillCnt, 1 To 10) As Double 'Payments are applied by priority. The first
  '4 are hard coded. The final 3 are determined by the order the user enters
  'them on the System Setup screen (last tab)
  If EditFlag = False Or (EditFlag = True And BillCnt > 0) Then 'If EditFlag is
  'false then this is a new customer and BillCnt will be > 0 since this function
  'is not accessible unless there is an amount in the amount owed field
    OpenVATaxTransFile THandle, NumOfTRecs
    For x = 1 To BillCnt
      Get THandle, BillTrans(x), TaxTrans
        TempBillList(x).BillRec = BillTrans(x)
        TempBillList(x).CustRec = GCustNum
        TempBillList(x).TaxYear = TaxTrans.TaxYear
        PaySeq(x, 1) = OldRound(PaySeq(x, 1) + TaxTrans.Revenue.Principle2)
        PaySeq(x, 1) = OldRound(PaySeq(x, 1) - TaxTrans.Revenue.Principle2Pd)
        TempBillList(x).TotOwed = PaySeq(x, 1)
        PaySeq(x, 2) = OldRound(PaySeq(x, 2) + TaxTrans.Revenue.Principle3)
        PaySeq(x, 2) = OldRound(PaySeq(x, 2) - TaxTrans.Revenue.Principle3Pd)
        TempBillList(x).TotOwed = OldRound(TempBillList(x).TotOwed + PaySeq(x, 2))
        PaySeq(x, 3) = OldRound(PaySeq(x, 3) + TaxTrans.Revenue.Principle4)
        PaySeq(x, 3) = OldRound(PaySeq(x, 3) - TaxTrans.Revenue.Principle4Pd)
        TempBillList(x).TotOwed = OldRound(TempBillList(x).TotOwed + PaySeq(x, 3))
        PaySeq(x, 4) = OldRound(PaySeq(x, 4) + TaxTrans.Revenue.Principle1)
        PaySeq(x, 4) = OldRound(PaySeq(x, 4) - OldRound(TaxTrans.Revenue.Principle1Pd + TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
        TempBillList(x).TotOwed = OldRound(TempBillList(x).TotOwed + PaySeq(x, 4))
        PaySeq(x, 5) = OldRound(PaySeq(x, 5) + TaxTrans.Revenue.Principle5)
        PaySeq(x, 5) = OldRound(PaySeq(x, 5) - TaxTrans.Revenue.Principle5Pd)
        TempBillList(x).TotOwed = OldRound(TempBillList(x).TotOwed + PaySeq(x, 5))
        PaySeq(x, 6) = OldRound(PaySeq(x, 6) + TaxTrans.Revenue.Interest)
        PaySeq(x, 6) = OldRound(PaySeq(x, 6) - TaxTrans.Revenue.InterestPd)
        TempBillList(x).TotOwed = OldRound(TempBillList(x).TotOwed + PaySeq(x, 6))
        PaySeq(x, 7) = OldRound(PaySeq(x, 7) + TaxTrans.Revenue.Penalty)
        PaySeq(x, 7) = OldRound(PaySeq(x, 7) - TaxTrans.Revenue.PenaltyPd)
        TempBillList(x).TotOwed = OldRound(TempBillList(x).TotOwed + PaySeq(x, 7))
        PaySeq(x, 8) = OldRound(PaySeq(x, 8) + TaxTrans.Revenue.RevOpt1)
        PaySeq(x, 8) = OldRound(PaySeq(x, 8) - TaxTrans.Revenue.RevOpt1Pd)
        TempBillList(x).TotOwed = OldRound(TempBillList(x).TotOwed + PaySeq(x, 8))
        PaySeq(x, 9) = OldRound(PaySeq(x, 9) + TaxTrans.Revenue.RevOpt2)
        PaySeq(x, 9) = OldRound(PaySeq(x, 9) - TaxTrans.Revenue.RevOpt2Pd)
        TempBillList(x).TotOwed = OldRound(TempBillList(x).TotOwed + PaySeq(x, 9))
        PaySeq(x, 10) = OldRound(PaySeq(x, 10) + TaxTrans.Revenue.RevOpt3)
        PaySeq(x, 10) = OldRound(PaySeq(x, 10) - TaxTrans.Revenue.RevOpt3Pd)
        TempBillList(x).TotOwed = OldRound(TempBillList(x).TotOwed + PaySeq(x, 10))
   Next x
   
   For x = 1 To BillCnt
        If TempBillList(x).DiscAmt > 0 Then GoSub ApplyDisc '1/25/07

'     WhatsLeft = OldRound(WhatsLeft - TempBillList(x).DiscAmt)
'     For y = 1 To BillCnt
'       WhatsLeft = WhatsLeft - TempBillList(y).DiscAmt
'       TempBillList(y).DiscAmt = 0
'     Next y
     For y = 1 To 10
       TotDisc = OldRound(TotDisc + TempBillList(x).DiscAmt)
       If y = PayOrder(1) Then
         If WhatsLeft >= PaySeq(x, 1) Then
           fpCurrMTPaid.Value = CDbl(fpCurrMTPaid.Value) + PaySeq(x, 1)
           TempBillList(x).MachTools = PaySeq(x, 1)
           TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).MachTools)
         Else
           fpCurrMTPaid.Value = CDbl(fpCurrMTPaid.Value) + WhatsLeft
           TempBillList(x).MachTools = WhatsLeft
           TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).MachTools)
         End If
         WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 1))
         If WhatsLeft <= 0 Then GoTo PlayedOut
       End If
       
       If y = PayOrder(2) Then
         If WhatsLeft >= PaySeq(x, 2) Then
           fpCurrMCPaid.Value = CDbl(fpCurrMCPaid.Value) + PaySeq(x, 2)
           TempBillList(x).MerchCap = PaySeq(x, 2)
           TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).MerchCap)
         Else
           fpCurrMCPaid.Value = CDbl(fpCurrMCPaid.Value) + WhatsLeft
           TempBillList(x).MerchCap = WhatsLeft
           TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).MerchCap)
         End If
         WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 2))
         If WhatsLeft <= 0 Then GoTo PlayedOut
       End If
       
       If y = PayOrder(3) Then
         If WhatsLeft >= PaySeq(x, 3) Then
           fpCurrFEPaid.Value = CDbl(fpCurrFEPaid.Value) + PaySeq(x, 3)
           TempBillList(x).FarmEquip = PaySeq(x, 3)
           TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).FarmEquip)
         Else
           fpCurrFEPaid.Value = CDbl(fpCurrFEPaid.Value) + WhatsLeft
           TempBillList(x).FarmEquip = WhatsLeft
           TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).FarmEquip)
         End If
         WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 3))
         If WhatsLeft <= 0 Then GoTo PlayedOut
       End If
     
       If y = PayOrder(4) Then
         If WhatsLeft >= PaySeq(x, 4) Then
           fpCurrPersPaid.Value = CDbl(fpCurrPersPaid.Value) + PaySeq(x, 4)
           TempBillList(x).Personal = PaySeq(x, 4)
           TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Personal)
         Else
           fpCurrPersPaid.Value = CDbl(fpCurrPersPaid.Value) + WhatsLeft
           TempBillList(x).Personal = WhatsLeft
           TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Personal)
         End If
         WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 4))
         If WhatsLeft <= 0 Then GoTo PlayedOut
       End If
       
       If y = PayOrder(5) Then
         If WhatsLeft >= PaySeq(x, 5) Then
           fpCurrMHPaid.Value = CDbl(fpCurrMHPaid.Value) + PaySeq(x, 5)
           TempBillList(x).MobHomes = PaySeq(x, 5)
           TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).MobHomes)
         Else
           fpCurrMHPaid.Value = CDbl(fpCurrMHPaid.Value) + WhatsLeft
           TempBillList(x).MobHomes = WhatsLeft
           TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).MobHomes)
         End If
         WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 5))
         If WhatsLeft <= 0 Then GoTo PlayedOut
       End If
       
       If y = PayOrder(6) Then
         If WhatsLeft >= PaySeq(x, 6) Then
           fpCurrIntPaid.Value = CDbl(fpCurrIntPaid.Value) + PaySeq(x, 6)
           TempBillList(x).Interest = PaySeq(x, 6)
           TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Interest)
         Else
           fpCurrIntPaid.Value = CDbl(fpCurrIntPaid.Value) + WhatsLeft
           TempBillList(x).Interest = WhatsLeft
           TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Interest)
         End If
         WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 6))
         If WhatsLeft <= 0 Then GoTo PlayedOut
       End If
       
       If y = PayOrder(7) Then
         If WhatsLeft >= PaySeq(x, 7) Then
           fpCurrPenPaid.Value = CDbl(fpCurrPenPaid.Value) + PaySeq(x, 7)
           TempBillList(x).Penalty = PaySeq(x, 7)
           TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Penalty)
         Else
           fpCurrPenPaid.Value = CDbl(fpCurrPenPaid.Value) + WhatsLeft
           TempBillList(x).Penalty = WhatsLeft
           TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Penalty)
         End If
         WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 7))
         If WhatsLeft <= 0 Then GoTo PlayedOut
       End If
         
       If y = PayOrder(8) Then
         If WhatsLeft >= PaySeq(x, 8) Then
           fpCurrOpt1Paid.Value = CDbl(fpCurrOpt1Paid.Value) + PaySeq(x, 8)
           TempBillList(x).Opt1 = PaySeq(x, 8)
           TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Opt1)
         Else
           fpCurrOpt1Paid.Value = CDbl(fpCurrOpt1Paid.Value) + WhatsLeft
           TempBillList(x).Opt1 = WhatsLeft
           TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Opt1)
         End If
         WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 8))
         If WhatsLeft <= 0 Then GoTo PlayedOut
       End If
         
       If y = PayOrder(9) Then
         If WhatsLeft >= PaySeq(x, 9) Then
           fpCurrOpt2Paid.Value = CDbl(fpCurrOpt2Paid.Value) + PaySeq(x, 9)
           TempBillList(x).Opt2 = PaySeq(x, 9)
           TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Opt2)
         Else
           fpCurrOpt2Paid.Value = CDbl(fpCurrOpt2Paid.Value) + WhatsLeft
           TempBillList(x).Opt2 = WhatsLeft
           TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Opt2)
         End If
         WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 9))
         If WhatsLeft <= 0 Then GoTo PlayedOut
       End If
         
       If y = PayOrder(10) Then
         If WhatsLeft >= PaySeq(x, 10) Then
           fpCurrOpt3Paid.Value = CDbl(fpCurrOpt3Paid.Value) + PaySeq(x, 10)
           TempBillList(x).Opt3 = PaySeq(x, 10)
           TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Opt3)
         Else
           fpCurrOpt3Paid.Value = CDbl(fpCurrOpt3Paid.Value) + WhatsLeft
           TempBillList(x).Opt3 = WhatsLeft
           TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Opt3)
         End If
         WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 10))
       End If
       
     Next y
   Next x
   
   Call AssignPaidTemps
   
 End If
 
PlayedOut:
  TotPaid# = OldRound(CDbl(fpCurrMTPaid.Value) + CDbl(fpCurrMCPaid.Value) + CDbl(fpCurrFEPaid.Value))
  TotPaid# = OldRound(TotPaid# + CDbl(fpCurrPersPaid.Value) + CDbl(fpCurrMHPaid.Value) + CDbl(fpCurrIntPaid.Value) + CDbl(fpCurrPenPaid.Value))
  TotPaid# = OldRound(TotPaid# + CDbl(fpCurrOpt1Paid.Value) + CDbl(fpCurrOpt2Paid.Value) + CDbl(fpCurrOpt3Paid.Value) + CDbl(fpCurrPrePay.Value))
  fpCurrTotPaid = TotPaid#
  fpCurrChngDue.Value = OldRound(CDbl(fpCurrTotRecd.Value) - CDbl(fpCurrTotPaid.Value) + CDbl(fpCurrDisc.Value))
  If CDbl(fpCurrChngDue.Value) < 0 Then
    fpCurrChngDue = 0
  End If
  Close THandle 'added THandle 7/18/07****************************************
 
  fpCurrTotWDisc = OldRound(CDbl(fpCurrTotPaid.Value) + CDbl(fpCurrDisc.Value))

  GetNewCust = False
  
'  If MaxDisc > 0 And CDbl(fpCurrDisc.Value) = 0 Then 'remarked 9/20/05
'    If VATaxMsgWOpts(800, "This customer is eligible for a discount. Press F10 if you wish to apply the discount. Otherwise, press ESC to override the discount.", "F10 Apply Discount", "ESC NO Discount") = "abort" Then
'      Unload frmVATaxMsgWOpts
'      If CDbl(fpCurrDisc.Value) > 0 Then
'        Call RemoveDiscount
'      End If
'      Exit Sub
'    Else
'      Unload frmVATaxMsgWOpts
'      ApplyDiscount
'      Call ReassignDiscount
'    End If
'  End If
    
  Exit Sub
ApplyDisc: 'added 1/25/07
  Disc1 = 0
  Disc2 = 0
  Disc3 = 0
  Disc4 = 0
  Disc5 = 0
  Disc6 = 0
  Disc7 = 0
  Disc8 = 0
  
  SaveAmt = TempBillList(x).TotOwed
  Disc1 = PaySeq(x, 1) / SaveAmt
  Disc1 = Disc1 * TempBillList(x).DiscAmt
  Disc2 = PaySeq(x, 2) / SaveAmt
  Disc2 = Disc2 * TempBillList(x).DiscAmt
  Disc3 = PaySeq(x, 3) / SaveAmt
  Disc3 = Disc3 * TempBillList(x).DiscAmt
  Disc4 = PaySeq(x, 4) / SaveAmt
  Disc4 = Disc4 * TempBillList(x).DiscAmt
  Disc5 = PaySeq(x, 5) / SaveAmt
  Disc5 = Disc5 * TempBillList(x).DiscAmt
  Disc6 = PaySeq(x, 8) / SaveAmt
  Disc6 = Disc6 * TempBillList(x).DiscAmt
  Disc7 = PaySeq(x, 9) / SaveAmt
  Disc7 = Disc7 * TempBillList(x).DiscAmt
  Disc8 = PaySeq(x, 10) / SaveAmt
  Disc8 = Disc8 * TempBillList(x).DiscAmt
  
  PaySeq(x, 1) = OldRound(PaySeq(x, 1) - Disc1)
  PaySeq(x, 2) = OldRound(PaySeq(x, 2) - Disc2)
  PaySeq(x, 3) = OldRound(PaySeq(x, 3) - Disc3)
  PaySeq(x, 4) = OldRound(PaySeq(x, 4) - Disc4)
  PaySeq(x, 5) = OldRound(PaySeq(x, 5) - Disc5)
  PaySeq(x, 8) = OldRound(PaySeq(x, 8) - Disc6)
  PaySeq(x, 9) = OldRound(PaySeq(x, 9) - Disc7)
  PaySeq(x, 10) = OldRound(PaySeq(x, 10) - Disc8)
  DiscApplied = True
  DumpPenny = OldRound(PaySeq(x, 1) + PaySeq(x, 2) + PaySeq(x, 3) + PaySeq(x, 4) + PaySeq(x, 5) + PaySeq(x, 8) + PaySeq(x, 9) + PaySeq(x, 10))
  If DumpPenny + TempBillList(x).DiscAmt < TempBillList(x).TotOwed Then
    If PaySeq(x, 1) > TempBillList(x).TotOwed - OldRound(DumpPenny + TempBillList(x).DiscAmt) Then
      PaySeq(x, 1) = PaySeq(x, 1) + OldRound(TempBillList(x).TotOwed - (DumpPenny + TempBillList(x).DiscAmt))
    ElseIf PaySeq(x, 2) > TempBillList(x).TotOwed - OldRound(DumpPenny + TempBillList(x).DiscAmt) Then
      PaySeq(x, 2) = PaySeq(x, 2) + OldRound(TempBillList(x).TotOwed - (DumpPenny + TempBillList(x).DiscAmt))
    ElseIf PaySeq(x, 3) > TempBillList(x).TotOwed - OldRound(DumpPenny + TempBillList(x).DiscAmt) Then
      PaySeq(x, 3) = PaySeq(x, 3) + OldRound(TempBillList(x).TotOwed - (DumpPenny + TempBillList(x).DiscAmt))
    ElseIf PaySeq(x, 4) > TempBillList(x).TotOwed - OldRound(DumpPenny + TempBillList(x).DiscAmt) Then
      PaySeq(x, 4) = PaySeq(x, 4) + OldRound(TempBillList(x).TotOwed - (DumpPenny + TempBillList(x).DiscAmt))
    ElseIf PaySeq(x, 5) > TempBillList(x).TotOwed - OldRound(DumpPenny + TempBillList(x).DiscAmt) Then
      PaySeq(x, 5) = PaySeq(x, 5) + OldRound(TempBillList(x).TotOwed - (DumpPenny + TempBillList(x).DiscAmt))
    ElseIf PaySeq(x, 8) > TempBillList(x).TotOwed - OldRound(DumpPenny + TempBillList(x).DiscAmt) Then
      PaySeq(x, 8) = PaySeq(x, 8) + OldRound(TempBillList(x).TotOwed - (DumpPenny + TempBillList(x).DiscAmt))
    ElseIf PaySeq(x, 9) > TempBillList(x).TotOwed - OldRound(DumpPenny + TempBillList(x).DiscAmt) Then
      PaySeq(x, 9) = PaySeq(x, 9) + OldRound(TempBillList(x).TotOwed - (DumpPenny + TempBillList(x).DiscAmt))
    ElseIf PaySeq(x, 10) > TempBillList(x).TotOwed - OldRound(DumpPenny + TempBillList(x).DiscAmt) Then
      PaySeq(x, 10) = PaySeq(x, 10) + OldRound(TempBillList(x).TotOwed - (DumpPenny + TempBillList(x).DiscAmt))
    End If
  ElseIf DumpPenny + TempBillList(x).DiscAmt > TempBillList(x).TotOwed Then
    If PaySeq(x, 1) > OldRound(DumpPenny + TempBillList(x).DiscAmt) - TempBillList(x).TotOwed Then
      PaySeq(x, 1) = PaySeq(x, 1) - OldRound((DumpPenny + TempBillList(x).DiscAmt) - TempBillList(x).TotOwed)
    ElseIf PaySeq(x, 2) > OldRound(DumpPenny + TempBillList(x).DiscAmt) - TempBillList(x).TotOwed Then
      PaySeq(x, 2) = PaySeq(x, 2) - OldRound((DumpPenny + TempBillList(x).DiscAmt) - TempBillList(x).TotOwed)
    ElseIf PaySeq(x, 3) > OldRound(DumpPenny + TempBillList(x).DiscAmt) - TempBillList(x).TotOwed Then
      PaySeq(x, 3) = PaySeq(x, 3) - OldRound((DumpPenny + TempBillList(x).DiscAmt) - TempBillList(x).TotOwed)
    ElseIf PaySeq(x, 4) > OldRound(DumpPenny + TempBillList(x).DiscAmt) - TempBillList(x).TotOwed Then
      PaySeq(x, 4) = PaySeq(x, 4) - OldRound((DumpPenny + TempBillList(x).DiscAmt) - TempBillList(x).TotOwed)
    ElseIf PaySeq(x, 5) > OldRound(DumpPenny + TempBillList(x).DiscAmt) - TempBillList(x).TotOwed Then
      PaySeq(x, 5) = PaySeq(x, 5) - OldRound((DumpPenny + TempBillList(x).DiscAmt) - TempBillList(x).TotOwed)
    ElseIf PaySeq(x, 8) > OldRound(DumpPenny + TempBillList(x).DiscAmt) - TempBillList(x).TotOwed Then
      PaySeq(x, 8) = PaySeq(x, 8) - OldRound((DumpPenny + TempBillList(x).DiscAmt) - TempBillList(x).TotOwed)
    ElseIf PaySeq(x, 9) > OldRound(DumpPenny + TempBillList(x).DiscAmt) - TempBillList(x).TotOwed Then
      PaySeq(x, 9) = PaySeq(x, 9) - OldRound((DumpPenny + TempBillList(x).DiscAmt) - TempBillList(x).TotOwed)
    ElseIf PaySeq(x, 10) > OldRound(DumpPenny + TempBillList(x).DiscAmt) - TempBillList(x).TotOwed Then
      PaySeq(x, 10) = PaySeq(x, 10) - OldRound((DumpPenny + TempBillList(x).DiscAmt) - TempBillList(x).TotOwed)
    End If
  End If
  
  Return
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersPaymentEntry", "Distribute", Erl)
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
   ' ClearInUse PWcnt
    'CMTerminate

End Sub

Private Sub RemoveDiscount()
  Dim x As Integer
'  Dim TempPayRec As VAPersPayListType
'  Dim THandle As Integer
'  Dim NumOfTRecs As Integer
'  Dim DiscAmt As Double
'
'  If BillCnt = 0 Then
'    OpenPersPayListFile THandle, OperNum
'    NumOfTRecs = LOF(THandle) / Len(TempPayRec)
'    DiscAmt = 0
'    For x = 1 To NumOfTRecs
'      Get THandle, x, TempPayRec
'      If TempPayRec.CustRec = GCustNum Then
'        DiscAmt = DiscAmt + TempPayRec.DiscAmt
'        TempPayRec.DiscAmt = 0
'        Put THandle, x, TempPayRec
'      End If
'    Next x
'    Close THandle
'    If DiscAmt > 0 Then Call Distribute(OldRound(CDbl(fpCurrTotRecd.Value)))
'  Else
    For x = 1 To BillCnt
      TempBillList(x).DiscAmt = 0
    Next x
'  End If
  
  ThisDiscAmt = 0
  fpCurrDisc = 0
  
  Call AddUpPaidCol
  
  Call AssignPaidTemps
End Sub

Private Sub OverRideDist()
  'this sub handles user entered overrides after auto distribution takes place
  Dim Message$
  Dim Top As Integer
  Dim OptRev$
  Dim x As Integer, y As Integer
  On Local Error GoTo ERRORSTUFF
  
  'we are looking for amounts owed versus amounts paid to see if there are
  'any shortfalls...if found then shortfalls negate discounts
  If CDbl(fpCurrDisc.Value) <= 0 Then Exit Sub
  
  InOverRideDist = True
  
  If CDbl(fpCurrMTOwed.Value) > CDbl(fpCurrMTPaid.Value) Then
    Message = "This personal payment configuration eliminates the discount because now the machine tools portion of the amount owed is underpaid. If you wish to continue with this payment configuation then press F10. Otherwise. press ESC to restore the former values."
    If VATaxMsgWOpts(600, Message, "F10 Continue", "ESC Restore") = "abort" Then
      Unload frmVATaxMsgWOpts
      Call ReLoadPaidTemps
      Close
      InOverRideDist = False
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
      fpCurrDisc.Value = 0
      TXLog ("WARNING: User warned that overriding the automatic machine tools payment of " + QPTrim$(Using("$##,##0.00", TempMTPaid)) + " would eliminate the discount of " + QPTrim$(Using$("$##,##0.00", TempDisc)) + " but the user continued with the new machine tools amount of " + fpCurrMTPaid.Text + ".")
    End If
  End If
  
  If CDbl(fpCurrMCOwed.Value) > CDbl(fpCurrMCPaid.Value) Then
    Message = "This personal payment configuration eliminates the discount because now the merchant capital portion of the amount owed is underpaid. If you wish to continue with this payment configuation then press F10. Otherwise. press ESC to restore the former values."
    If VATaxMsgWOpts(600, Message, "F10 Continue", "ESC Restore") = "abort" Then
      Unload frmVATaxMsgWOpts
      Call ReLoadPaidTemps
      Close
      InOverRideDist = False
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
      fpCurrDisc.Value = 0
      TXLog ("WARNING: User warned that overriding the automatic merchant capital payment of " + QPTrim$(Using("$##,##0.00", TempMCPaid)) + " would eliminate the discount of " + QPTrim$(Using$("$##,##0.00", TempDisc)) + " but the user continued with the new merchant capital amount of " + fpCurrMCPaid.Text + ".")
    End If
  End If
  
  If CDbl(fpCurrFEOwed.Value) > CDbl(fpCurrFEPaid.Value) Then
    Message = "This payment configuration eliminates the discount because now the farm equipment portion of the amount owed is underpaid. If you wish to continue with this payment configuation then press F10. Otherwise. press ESC to restore the former values."
    If VATaxMsgWOpts(600, Message, "F10 Continue", "ESC Restore") = "abort" Then
      Unload frmVATaxMsgWOpts
      Call ReLoadPaidTemps
      Close
      InOverRideDist = False
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
      fpCurrDisc.Value = 0
      TXLog ("WARNING: User warned that overriding the automatic farm equipment payment of " + QPTrim$(Using("$##,##0.00", TempFEPaid)) + " would eliminate the discount of " + QPTrim$(Using$("$##,##0.00", TempDisc)) + " but the user continued with the new farm equipment amount of " + fpCurrFEPaid.Text + ".")
    End If
  End If
  
  If CDbl(fpCurrPersOwed.Value) > OldRound(CDbl(fpCurrPersPaid.Value) + CDbl(fpCurrDisc.Value)) Then
    Message = "This payment configuration eliminates the discount because now the personal portion of the amount owed is underpaid. If you wish to continue with this payment configuation then press F10. Otherwise. press ESC to restore the former values."
    If VATaxMsgWOpts(600, Message, "F10 Continue", "ESC Restore") = "abort" Then
      Unload frmVATaxMsgWOpts
      Call ReLoadPaidTemps
      Close
      InOverRideDist = False
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
      fpCurrDisc.Value = 0
      TXLog ("WARNING: User warned that overriding the automatic personal payment of " + QPTrim$(Using("$##,##0.00", TempPersPaid)) + " would eliminate the discount of " + QPTrim$(Using$("$##,##0.00", TempDisc)) + " but the user continued with the new personal amount of " + fpCurrPersPaid.Text + ".")
    End If
  End If
  
'  If fpCurrMHPaid.Enabled = False Then GoTo Next2
'  OptRev = QPTrim$(fptxtOpt1.Text)
  If CDbl(fpCurrMHOwed.Value) > CDbl(fpCurrMHPaid.Value) Then
    Message = "This payment configuration eliminates the discount because now the mobile home portion of the amount owed is underpaid. If you wish to continue with this payment configuation then press F10. Otherwise. press ESC to restore the former values."
    If VATaxMsgWOpts(600, Message, "F10 Continue", "ESC Restore") = "abort" Then
      Unload frmVATaxMsgWOpts
      Call ReLoadPaidTemps
      Close
      InOverRideDist = False
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
      fpCurrDisc.Value = 0
      TXLog ("WARNING: User warned that overriding the automatic mobile home payment of " + QPTrim$(Using("$##,##0.00", TempMHPaid)) + " would eliminate the discount of " + QPTrim$(Using$("$##,##0.00", TempDisc)) + " but the user continued with the new mobile home amount of " + fpCurrMHPaid.Text + ".")
    End If
  End If
  
'  If fpCurrIntPaid.Enabled = False Then GoTo Next3
'  OptRev = QPTrim$(fptxtOpt2.Text)
  If CDbl(fpCurrIntOwed.Value) > CDbl(fpCurrIntPaid.Value) Then
    Message = "This payment configuration eliminates the discount because now the interest portion of the amount owed is underpaid. If you wish to continue with this payment configuation then press F10. Otherwise. press ESC to restore the former values."
    If VATaxMsgWOpts(600, Message, "F10 Continue", "ESC Restore") = "abort" Then
      Unload frmVATaxMsgWOpts
      Call ReLoadPaidTemps
      Close
      InOverRideDist = False
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
      fpCurrDisc.Value = 0
      TXLog ("WARNING: User warned that overriding the automatic interest payment of " + QPTrim$(Using("$##,##0.00", TempIntPaid)) + " would eliminate the discount of " + QPTrim$(Using$("$##,##0.00", TempDisc)) + " but the user continued with the new interest amount of " + fpCurrIntPaid.Text + ".")
    End If
  End If
  
'  If fpCurrIntPaid.Enabled = False Then GoTo Next4
'  OptRev = QPTrim$(fptxtOpt3.Text)
  If CDbl(fpCurrPenOwed.Value) > CDbl(fpCurrPenPaid.Value) Then
    Message = "This payment configuration eliminates the discount because now the penalty portion of the amount owed is underpaid. If you wish to continue with this payment configuation then press F10. Otherwise. press ESC to restore the former values."
    If VATaxMsgWOpts(600, Message, "F10 Continue", "ESC Restore") = "abort" Then
      Unload frmVATaxMsgWOpts
      Call ReLoadPaidTemps
      Close
      InOverRideDist = False
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
      fpCurrDisc.Value = 0
      TXLog ("WARNING: User warned that overriding the automatic penalty payment of " + QPTrim$(Using("$##,##0.00", TempPenPaid)) + " would eliminate the discount of " + QPTrim$(Using$("$##,##0.00", TempDisc)) + " but the user continued with the new penalty amount of " + fpCurrPenPaid.Text + ".")
    End If
  End If
  
  If fpCurrOpt1Paid.Enabled = False Then GoTo Next2
  If CDbl(fpCurrOpt1Owed.Value) > CDbl(fpCurrOpt1Paid.Value) Then
    Message = "This payment configuration eliminates the discount because now the " + Opt1Desc + " portion of the amount owed is underpaid. If you wish to continue with this payment configuation then press F10. Otherwise. press ESC to restore the former values."
    If VATaxMsgWOpts(600, Message, "F10 Continue", "ESC Restore") = "abort" Then
      Unload frmVATaxMsgWOpts
      Call ReLoadPaidTemps
      Close
      InOverRideDist = False
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
      fpCurrDisc.Value = 0
      TXLog ("WARNING: User warned that overriding the automatic " + Opt1Desc + " payment of " + QPTrim$(Using("$##,##0.00", TempOpt1Paid)) + " would eliminate the discount of " + QPTrim$(Using$("$##,##0.00", TempDisc)) + " but the user continued with the new penalty amount of " + fpCurrOpt1Paid.Text + ".")
    End If
  End If
  
Next2:
  If fpCurrOpt2Paid.Enabled = False Then GoTo Next3
  If CDbl(fpCurrOpt2Owed.Value) > CDbl(fpCurrOpt2Paid.Value) Then
    Message = "This payment configuration eliminates the discount because now the " + Opt2Desc + " portion of the amount owed is underpaid. If you wish to continue with this payment configuation then press F10. Otherwise. press ESC to restore the former values."
    If VATaxMsgWOpts(600, Message, "F10 Continue", "ESC Restore") = "abort" Then
      Unload frmVATaxMsgWOpts
      Call ReLoadPaidTemps
      Close
      InOverRideDist = False
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
      fpCurrDisc.Value = 0
      TXLog ("WARNING: User warned that overriding the automatic " + Opt2Desc + " payment of " + QPTrim$(Using("$##,##0.00", TempOpt2Paid)) + " would eliminate the discount of " + QPTrim$(Using$("$##,##0.00", TempDisc)) + " but the user continued with the new " + Opt2Desc + " amount of " + fpCurrOpt2Paid.Text + ".")
    End If
  End If
  
Next3:
  If fpCurrOpt3Paid.Enabled = False Then GoTo Next4
  If CDbl(fpCurrOpt3Owed.Value) > CDbl(fpCurrOpt3Paid.Value) Then
    Message = "This payment configuration eliminates the discount because now the " + Opt3Desc + " portion of the amount owed is underpaid. If you wish to continue with this payment configuation then press F10. Otherwise. press ESC to restore the former values."
    If VATaxMsgWOpts(600, Message, "F10 Continue", "ESC Restore") = "abort" Then
      Unload frmVATaxMsgWOpts
      Call ReLoadPaidTemps
      Close
      InOverRideDist = False
      Exit Sub
    Else
      Unload frmVATaxMsgWOpts
      fpCurrDisc.Value = 0
      TXLog ("WARNING: User warned that overriding the automatic " + Opt3Desc + " payment of " + QPTrim$(Using("$##,##0.00", TempOpt3Paid)) + " would eliminate the discount of " + QPTrim$(Using$("$##,##0.00", TempDisc)) + " but the user continued with the new " + Opt3Desc + "  amount of " + fpCurrOpt3Paid.Text + ".")
    End If
  End If
  
Next4:

  Call AddUpPaidCol
  InOverRideDist = False
  Call AssignPaidTemps
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersPaymentEntry", "OverRideDist", Erl)
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
  '  ClearInUse PWcnt
   ' CMTerminate

End Sub

Private Sub AssignPaidTemps()
  TempPersPaid = CDbl(fpCurrPersPaid.Value)
  TempMTPaid = CDbl(fpCurrMTPaid.Value)
  TempMCPaid = CDbl(fpCurrMCPaid.Value)
  TempFEPaid = CDbl(fpCurrFEPaid.Value)
  TempMHPaid = CDbl(fpCurrMHPaid.Value)
  TempIntPaid = CDbl(fpCurrIntPaid.Value)
  TempPenPaid = CDbl(fpCurrPenPaid.Value)
  TempOpt1Paid = CDbl(fpCurrOpt1Paid.Value)
  TempOpt2Paid = CDbl(fpCurrOpt2Paid.Value)
  TempOpt3Paid = CDbl(fpCurrOpt3Paid.Value)
  TempDisc = CDbl(fpCurrDisc.Value)
  TempTotPd = CDbl(fpCurrTotPaid.Value)
  TempPrePay = CDbl(fpCurrPrePay.Value)
End Sub

Private Sub ReLoadPaidTemps()
  fpCurrPersPaid = TempPersPaid
  fpCurrMTPaid = TempMTPaid
  fpCurrMCPaid = TempMCPaid
  fpCurrFEPaid = TempFEPaid
  fpCurrMHPaid = TempMHPaid
  fpCurrIntPaid = TempIntPaid
  fpCurrPenPaid = TempPenPaid
  fpCurrOpt1Paid = TempOpt1Paid
  fpCurrOpt2Paid = TempOpt2Paid
  fpCurrOpt3Paid = TempOpt3Paid
  fpCurrDisc = TempDisc
  fpCurrTotPaid = TempTotPd
  fpCurrPrePay = TempPrePay
End Sub

Private Sub GetMaxDisc()
  Dim TPayRec As VAPersPayListType
  Dim PayRec As VAPersPayListType
  Dim PHandle As Integer
  Dim NumOfPRecs As Integer
  Dim ThisPrevRec As Long
  Dim NewRec As Integer
  Dim Operator$
  Dim TempRec As TempPayList
  Dim NumOfTemps As Integer
  Dim TempHandle As Integer
  Dim x As Integer
  Dim TaxMasterRec As VATaxMasterType
  Dim MHandle As Integer
  Dim ThisDiscPct As Double
  Dim TaxTRec As VATaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim ThisTaxYear As Integer
  Dim Balance#
  Dim Nextx As Integer
  Dim SmallNum As Integer
  Dim HoldNum As Long
  Dim HoldDate As Integer
  Dim Thisx As Integer
  
  On Local Error GoTo ERRORSTUFF
  
  OpenVATaxSetUpFile MHandle
  Get MHandle, 1, TaxMasterRec
  Close MHandle
  ThisDiscPct = TaxMasterRec.DisPPct
  MaxDisc = 0
'  If TaxMasterRec.DiscPXDate > Date2Num(fptxtPayDate.Text) Then
'    DiscYN = True
'  End If
  
  ThisTaxYear = TaxMasterRec.PTaxYear
  
  If BillCnt = 0 And EditFlag = True Then 'Or Exist("editpyment.dat") Then 'user is editing and is not accessing
  'the bill list
    ReDim BillTrans(1 To 1) As Long
    ReDim BillDate(1 To 1) As Integer
    ThisPrevRec = 0
    NewRec = 0
    Operator$ = CStr(OperNum)
    Operator$ = QPTrim$(Operator$)
    OpenPersPayListFile PHandle, OperNum 'saved by getting data from temporary
    'bill record
    NumOfPRecs = LOF(PHandle) / Len(PayRec)
    For x = 1 To NumOfPRecs
      Get PHandle, x, PayRec
      If PayRec.CustRec = GCustNum Then
        BillCnt = BillCnt + 1
        ReDim Preserve BillTrans(1 To BillCnt) As Long
        BillTrans(BillCnt) = PayRec.BillRec
        ReDim Preserve BillDate(1 To BillCnt) As Integer
        BillDate(BillCnt) = TempRec.BillDate
      End If
    Next x
    Close PHandle
  ElseIf Exist(VATempPersBillRecs) Then
    ReDim BillTrans(1 To 1) As Long
    ReDim BillDate(1 To 1) As Integer
    BillCnt = 0
    OpenVAPersTempBillRecs TempHandle, NumOfTemps
    For x = 1 To NumOfTemps
      Get TempHandle, x, TempRec
      If TempRec.BillRec > 0 Then
        BillCnt = BillCnt + 1
        ReDim Preserve BillTrans(1 To BillCnt) As Long
        ReDim Preserve BillDate(1 To BillCnt) As Integer
        BillTrans(BillCnt) = TempRec.BillPtr
        BillDate(BillCnt) = TempRec.BillDate
        'this data should be the same data as that where PaySeq() are loaded
      End If
    Next x
    Close TempHandle
  End If
  
  If BillCnt = 0 Then Exit Sub
  Nextx = 1
  SmallNum = 30000
  ReDim TransSeq(1 To BillCnt) As Long
  Do While Nextx <= BillCnt
    For x = Nextx To BillCnt
      If BillDate(x) <= SmallNum Then
        SmallNum = BillDate(x)
        Thisx = x
      End If
    Next x
    HoldNum = BillTrans(Nextx)
    HoldDate = BillDate(Nextx)
    BillTrans(Nextx) = BillTrans(Thisx)
    BillDate(Nextx) = BillDate(Thisx)
    BillTrans(Thisx) = HoldNum
    BillDate(Thisx) = HoldDate
    Nextx = Nextx + 1
    SmallNum = 30000
  Loop
  
  ReDim Preserve TempBillList(1 To BillCnt) As VAPersPayListType
  
  OpenVATaxTransFile THandle, NumOfTRecs
  For x = 1 To BillCnt
    If BillTrans(x) <= 0 Then
      Call DisablePayFields
      OverPay = True
      GoTo OverPay
    End If
    Get THandle, BillTrans(x), TaxTRec
    Balance = 0
    If TaxTRec.BillType = "P" And TaxTRec.TaxYear = ThisTaxYear And TaxTRec.DiscXDate > 0 And DiscPXDate <= TaxTRec.DiscXDate Then ' Date2Num(fptxtPayDate.Text) Then
'      Balance# = OldRound(Balance# + TaxTRec.Revenue.Principle1 + TaxTRec.Revenue.Principle2 + TaxTRec.Revenue.Principle3)'remmed out on 2/9/07
'      Balance# = OldRound(Balance# + TaxTRec.Revenue.Principle4 + TaxTRec.Revenue.Principle5)
'      Balance# = OldRound(Balance# + TaxTRec.Revenue.RevOpt1 + TaxTRec.Revenue.RevOpt2 + TaxTRec.Revenue.RevOpt3)
'
'      Balance# = OldRound(Balance# - (TaxTRec.Revenue.Principle1Pd + TaxTRec.Revenue.Principle2Pd + TaxTRec.Revenue.Principle3Pd))
'      Balance# = OldRound(Balance# - (TaxTRec.Revenue.Principle4Pd + TaxTRec.Revenue.Principle5Pd))
'      Balance# = OldRound(Balance# - (TaxTRec.Revenue.RevOpt1Pd + TaxTRec.Revenue.RevOpt2Pd + TaxTRec.Revenue.RevOpt3Pd + TaxTRec.PPTRADisc - TaxTRec.PPTRARmvl))
      Balance = TaxTRec.Amount 'added 2/9/07
      If Balance# > 0 Then 'save which transaction the discount is applied to
  '      If DiscYN = True Then
          MaxDisc = MaxDisc + OldRound(Balance# * ThisDiscPct * 0.01)
  '      Else
  '        MaxDisc = 0
  '      End If
        TempBillList(x).DiscAmt = OldRound(Balance# * ThisDiscPct * 0.01)
      End If
    End If
  Next x
OverPay:
  Close THandle
  
  If MaxDisc = 0 Then
    fpCurrDisc.ControlType = ControlTypeReadOnly
    If CDbl(fpCurrDisc.Value) > 0 Then
      fpCurrDisc = 0
    End If
  Else
    fpCurrDisc.ControlType = ControlTypeNormal
  End If
  
  Exit Sub
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersPaymentEntry", "GetMaxDisc", Erl)
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
  '  ClearInUse PWcnt
  '  CMTerminate


End Sub

Private Sub LoadTempPayList()
  Dim TaxTrans As VATaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim WhatsLeft As Double
  Dim x As Integer
  
  On Local Error GoTo ERRORSTUFF
  
  'this sub loads the tempbilllist with amounts that are
  'not generated by the automatic distribution such as when
  'an existing edit is brought up or when the user overrides
  'automatically distributed amounts
  If BillCnt = 0 Then Exit Sub
  ReDim PaySeq(1 To BillCnt, 1 To 10) As Double
  ReDim Preserve TempBillList(1 To BillCnt) As VAPersPayListType
  TempBillListCnt = BillCnt
  'BillTrans are in oldest first order
  OpenVATaxTransFile THandle, NumOfTRecs
  For x = 1 To BillCnt
    If BillTrans(x) <= 0 Then
      Call DisablePayFields
      OverPay = True
      GoTo OverPay
    End If
    Get THandle, BillTrans(x), TaxTrans
OverPay:
      TempBillList(x).BillRec = BillTrans(x)
      TempBillList(x).CustRec = GCustNum
      TempBillList(x).TaxYear = TaxTrans.TaxYear
      PaySeq(x, 1) = OldRound(PaySeq(x, 1) + TaxTrans.Revenue.Principle2)
      PaySeq(x, 1) = OldRound(PaySeq(x, 1) - TaxTrans.Revenue.Principle2Pd)
      PaySeq(x, 2) = OldRound(PaySeq(x, 2) + TaxTrans.Revenue.Principle3)
      PaySeq(x, 2) = OldRound(PaySeq(x, 2) - TaxTrans.Revenue.Principle3Pd)
      PaySeq(x, 3) = OldRound(PaySeq(x, 3) + TaxTrans.Revenue.Principle4)
      PaySeq(x, 3) = OldRound(PaySeq(x, 3) - TaxTrans.Revenue.Principle4Pd)
      PaySeq(x, 4) = OldRound(PaySeq(x, 4) + TaxTrans.Revenue.Principle1)
      PaySeq(x, 4) = OldRound(PaySeq(x, 4) - OldRound(TaxTrans.Revenue.Principle1Pd + TaxTrans.PPTRADisc - TaxTrans.PPTRARmvl))
      PaySeq(x, 5) = OldRound(PaySeq(x, 5) + TaxTrans.Revenue.Principle5)
      PaySeq(x, 5) = OldRound(PaySeq(x, 5) - TaxTrans.Revenue.Principle5Pd)
      PaySeq(x, 6) = OldRound(PaySeq(x, 6) + TaxTrans.Revenue.Interest)
      PaySeq(x, 6) = OldRound(PaySeq(x, 6) - TaxTrans.Revenue.InterestPd)
      PaySeq(x, 7) = OldRound(PaySeq(x, 7) + TaxTrans.Revenue.Penalty)
      PaySeq(x, 7) = OldRound(PaySeq(x, 7) - TaxTrans.Revenue.PenaltyPd)
      PaySeq(x, 8) = OldRound(PaySeq(x, 8) + TaxTrans.Revenue.RevOpt1)
      PaySeq(x, 8) = OldRound(PaySeq(x, 8) - TaxTrans.Revenue.RevOpt1Pd)
      PaySeq(x, 9) = OldRound(PaySeq(x, 9) + TaxTrans.Revenue.RevOpt2)
      PaySeq(x, 9) = OldRound(PaySeq(x, 9) - TaxTrans.Revenue.RevOpt2Pd)
      PaySeq(x, 10) = OldRound(PaySeq(x, 10) + TaxTrans.Revenue.RevOpt3)
      PaySeq(x, 10) = OldRound(PaySeq(x, 10) - TaxTrans.Revenue.RevOpt3Pd)
   Next x
   
   WhatsLeft = CDbl(fpCurrMTPaid.Value)
   For x = 1 To BillCnt
     If WhatsLeft >= PaySeq(x, 1) Then
       TempBillList(x).MachTools = PaySeq(x, 1)
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).MachTools)
     Else
       TempBillList(x).MachTools = WhatsLeft
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).MachTools)
     End If
     WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 1))
     If WhatsLeft < 0 Then WhatsLeft = 0
   Next x
   
   WhatsLeft = CDbl(fpCurrMCPaid.Value)
   For x = 1 To BillCnt
     If WhatsLeft >= PaySeq(x, 2) Then
       TempBillList(x).MerchCap = PaySeq(x, 2)
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).MerchCap)
     Else
       fpCurrMCPaid.Value = CDbl(fpCurrMCPaid.Value) + WhatsLeft
       TempBillList(x).MerchCap = WhatsLeft
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).MerchCap)
     End If
     WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 2))
     If WhatsLeft < 0 Then WhatsLeft = 0
   Next x

   WhatsLeft = CDbl(fpCurrFEPaid.Value)
   For x = 1 To BillCnt
     If WhatsLeft >= PaySeq(x, 3) Then
       TempBillList(x).FarmEquip = PaySeq(x, 3)
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).FarmEquip)
     Else
       TempBillList(x).FarmEquip = WhatsLeft
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).FarmEquip)
     End If
     WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 3))
     If WhatsLeft < 0 Then WhatsLeft = 0
   Next x

   WhatsLeft = CDbl(fpCurrPersPaid.Value)
   For x = 1 To BillCnt
     If WhatsLeft >= PaySeq(x, 4) Then
       TempBillList(x).Personal = PaySeq(x, 4)
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Personal)
     Else
       TempBillList(x).Personal = WhatsLeft
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Personal)
     End If
     WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 4))
     If WhatsLeft < 0 Then WhatsLeft = 0
   Next x

   WhatsLeft = CDbl(fpCurrMHPaid.Value)
   For x = 1 To BillCnt
     If WhatsLeft >= PaySeq(x, 5) Then
       TempBillList(x).MobHomes = PaySeq(x, 5)
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).MobHomes)
     Else
       TempBillList(x).MobHomes = WhatsLeft
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).MobHomes)
     End If
     WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 5))
     If WhatsLeft < 0 Then WhatsLeft = 0
   Next x

   WhatsLeft = CDbl(fpCurrIntPaid.Value)
   For x = 1 To BillCnt
     If WhatsLeft >= PaySeq(x, 6) Then
       TempBillList(x).Interest = PaySeq(x, 6)
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Interest)
     Else
       TempBillList(x).Interest = WhatsLeft
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Interest)
     End If
     WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 6))
     If WhatsLeft < 0 Then WhatsLeft = 0
   Next x

   WhatsLeft = CDbl(fpCurrPenPaid.Value)
   For x = 1 To BillCnt
     If WhatsLeft >= PaySeq(x, 7) Then
       TempBillList(x).Penalty = PaySeq(x, 7)
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Penalty)
     Else
       TempBillList(x).Penalty = WhatsLeft
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Penalty)
     End If
     WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 7))
     If WhatsLeft < 0 Then WhatsLeft = 0
   Next x
   
   WhatsLeft = CDbl(fpCurrOpt1Paid.Value)
   For x = 1 To BillCnt
     If WhatsLeft >= PaySeq(x, 8) Then
       TempBillList(x).Opt1 = PaySeq(x, 8)
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Opt1)
     Else
       TempBillList(x).Opt1 = WhatsLeft
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Opt1)
     End If
     WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 8))
     If WhatsLeft < 0 Then WhatsLeft = 0
   Next x
   
   WhatsLeft = CDbl(fpCurrOpt2Paid.Value)
   For x = 1 To BillCnt
     If WhatsLeft >= PaySeq(x, 9) Then
       TempBillList(x).Opt2 = PaySeq(x, 9)
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Opt2)
     Else
       TempBillList(x).Opt2 = WhatsLeft
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Opt2)
     End If
     WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 9))
     If WhatsLeft < 0 Then WhatsLeft = 0
   Next x
   
   WhatsLeft = CDbl(fpCurrOpt3Paid.Value)
   For x = 1 To BillCnt
     If WhatsLeft >= PaySeq(x, 10) Then
       TempBillList(x).Opt3 = PaySeq(x, 10)
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Opt3)
     Else
       TempBillList(x).Opt3 = WhatsLeft
       TempBillList(x).TotPaid = OldRound(TempBillList(x).TotPaid + TempBillList(x).Opt3)
     End If
     WhatsLeft = OldRound(WhatsLeft - PaySeq(x, 10))
     If WhatsLeft < 0 Then WhatsLeft = 0
   Next x
   
   Exit Sub
   
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersPaymentEntry", "LoadTempPayList", Erl)
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
  '  ClearInUse PWcnt
   ' CMTerminate

End Sub

Private Sub DisablePayFields()
  fpCurrPersPaid.Enabled = False
  fpCurrMTPaid.Enabled = False
  fpCurrMCPaid.Enabled = False
  fpCurrFEPaid.Enabled = False
  fpCurrMHPaid.Enabled = False
  fpCurrIntPaid.Enabled = False
  fpCurrPenPaid.Enabled = False
  fpCurrOpt1Paid.Enabled = False
  fpCurrOpt2Paid.Enabled = False
  fpCurrOpt3Paid.Enabled = False
End Sub

Private Sub EnablePayFields()
  fpCurrPersPaid.Enabled = True
  fpCurrMTPaid.Enabled = True
  fpCurrMCPaid.Enabled = True
  fpCurrFEPaid.Enabled = True
  fpCurrMHPaid.Enabled = True
  fpCurrIntPaid.Enabled = True
  fpCurrPenPaid.Enabled = True
  fpCurrOpt1Paid.Enabled = True
  fpCurrOpt2Paid.Enabled = True
  fpCurrOpt3Paid.Enabled = True
End Sub

Private Sub ClearTemps()
   TempPayDate = Date2Num(payDate) '2/14/06
   TempAmtOwed = 0
   TempTenderTY = 0
   TempCashAmt = 0
   TempChkAmt = 0
   TempChrgAmt = 0
   TempAmtRecd = 0
   TempChange = 0
   TempDesc = 0
   TempPaidOwed1AmtOwed = 0
   TempPaidOwed2AmtOwed = 0
   TempPaidOwed3AmtOwed = 0
   TempPaidOwed4AmtOwed = 0
   TempPaidOwed5AmtOwed = 0
   TempPaidOwed6AmtOwed = 0
   TempPaidOwed7AmtOwed = 0
   TempPaidOwed8AmtOwed = 0
   TempPaidOwed9AmtOwed = 0
   TempPaidOwed10AmtOwed = 0
   TempPaidOwed1AmtPaid = 0
   TempPaidOwed2AmtPaid = 0
   TempPaidOwed3AmtPaid = 0
   TempPaidOwed4AmtPaid = 0
   TempPaidOwed5AmtPaid = 0
   TempPaidOwed6AmtPaid = 0
   TempPaidOwed7AmtPaid = 0
   TempPaidOwed8AmtPaid = 0
   TempPaidOwed9AmtPaid = 0
   TempPaidOwed10AmtPaid = 0
   TempTotOwed = 0
   TempAmtPaid = 0
   TempTotPaid = 0
   TempPersPaid = 0
   TempMTPaid = 0
   TempMCPaid = 0
   TempFEPaid = 0
   TempMHPaid = 0
   TempIntPaid = 0
   TempPenPaid = 0
   TempOpt1Paid = 0
   TempOpt2Paid = 0
   TempOpt3Paid = 0
   TempDisc = 0
   TempTotPd = 0
   TempPrePay = 0
   TempAcctNum = 0
End Sub

Private Sub ReassignDiscount()
  Dim x As Integer
  Dim DiscEntered As Double
  
  DiscEntered = CDbl(fpCurrDisc.Value)
  ReDim ThisPct(1 To BillCnt) As Double
  Call GetMaxDisc
'  fpCurrPersPaid = fpCurrPersPaid
  For x = 1 To BillCnt
    If TempBillList(x).DiscAmt > 0 Then
      ThisPct(x) = OldRound(TempBillList(x).DiscAmt / MaxDisc)
    Else
      ThisPct(x) = 0
    End If
  Next x
  
  For x = 1 To BillCnt
    TempBillList(x).DiscAmt = OldRound(ThisPct(x) * DiscEntered)
  Next x
  
End Sub

Private Function AllTaggedPaid() As Boolean
  Dim x As Integer
  Dim ThisTot#
  
  AllTaggedPaid = True
  
  For x = 1 To BillCnt
    If TempBillList(x).TotOwed > OldRound(TempBillList(x).TotPaid + TempBillList(x).DiscAmt) Then
      AllTaggedPaid = False
      Exit For
    End If
  Next x
  
End Function

Private Sub ClearPaidFields()
  fpCurrPrePay = 0
  fpCurrPersPaid = 0
  fpCurrMCPaid = 0
  fpCurrFEPaid = 0
  fpCurrMHPaid = 0
  fpCurrIntPaid = 0
  fpCurrPenPaid = 0
  fpCurrOpt1Paid = 0
  fpCurrOpt2Paid = 0
  fpCurrOpt3Paid = 0
  fpCurrTotPaid = 0
  fpCurrTotWDisc = 0
'  fpCurrDisc = 0
End Sub

Private Sub ResetLeaveName()
  Call ClearPaidFields
  fpCurrAmtOwed = 0
  fpCurrTotRecd = 0
  fpCurrCashPd = 0
  fpCurrChkChrgPd = 0
  fpcmbTenderType.Text = "CASH"
  fpCurrChngDue = 0
  fpCurrPersOwed = 0
  fpCurrPersPaid = 0
  fpCurrMTOwed = 0
  fpCurrMTPaid = 0
  fpCurrMCOwed = 0
  fpCurrMCPaid = 0
  fpCurrFEOwed = 0
  fpCurrFEPaid = 0
  fpCurrMHOwed = 0
  fpCurrMHPaid = 0
  fpCurrIntOwed = 0
  fpCurrIntPaid = 0
  fpCurrPenOwed = 0
  fpCurrPenPaid = 0
  fpCurrOpt1Owed = 0
  fpCurrOpt1Paid = 0
  fpCurrOpt2Owed = 0
  fpCurrOpt2Paid = 0
  fpCurrOpt3Owed = 0
  fpCurrOpt3Paid = 0
  fpCurrTotOwed = 0
  TempPayDate = Date2Num(DefPayDate)
  TempAmtOwed = 0
  TempTenderTY = "CASH"
  TempCashAmt = 0
  TempChkAmt = 0
  TempChrgAmt = 0
  TempAmtRecd = 0
  TempChange = 0
  TempDesc = ""
  TempPaidOwed1AmtOwed = 0
  TempPaidOwed2AmtOwed = 0
  TempPaidOwed3AmtOwed = 0
  TempPaidOwed4AmtOwed = 0
  TempPaidOwed5AmtOwed = 0
  TempPaidOwed6AmtOwed = 0
  TempPaidOwed7AmtOwed = 0
  TempPaidOwed8AmtOwed = 0
  TempPaidOwed9AmtOwed = 0
  TempPaidOwed10AmtOwed = 0
  TempPaidOwed1AmtPaid = 0
  TempPaidOwed2AmtPaid = 0
  TempPaidOwed3AmtPaid = 0
  TempPaidOwed4AmtPaid = 0
  TempPaidOwed5AmtPaid = 0
  TempPaidOwed6AmtPaid = 0
  TempPaidOwed7AmtPaid = 0
  TempPaidOwed8AmtPaid = 0
  TempPaidOwed9AmtPaid = 0
  TempPaidOwed10AmtPaid = 0
  TempTotOwed = 0
  TempAmtPaid = 0
  TempTotPaid = 0
End Sub

Public Sub MsgAlertTimer_Timer()
  Static tog As Double
  Static TogState As Boolean
  If Me.Visible Then
    If BtnFnt# = 0 Then
      BtnFnt# = cmdInfo.FontSize
    End If
    If TogState Then
      tog = tog + 1
    Else
      tog = tog - 1
    End If
    Select Case tog
    Case 1
      cmdInfo.ForeColor = &H80000012
      cmdInfo.FontSize = BtnFnt
    Case 2
      cmdInfo.ForeColor = &H80000011
      cmdInfo.FontSize = BtnFnt - 0.7
    Case 3
      cmdInfo.ForeColor = &H80000011
      cmdInfo.FontSize = BtnFnt - 1.4
    Case 4
      cmdInfo.ForeColor = &H80000010
      cmdInfo.FontSize = BtnFnt - 2.1
    Case 5
      cmdInfo.ForeColor = &H80000010
      cmdInfo.FontSize = BtnFnt - 2.8
    Case 6
      cmdInfo.ForeColor = &H8000000F
      cmdInfo.FontSize = BtnFnt - 3.5
    Case 7
      cmdInfo.ForeColor = &H8000000F
      cmdInfo.FontSize = BtnFnt - 4.2
    Case 8
      cmdInfo.ForeColor = &H8000000E
      cmdInfo.FontSize = BtnFnt - 4.9
    Case 9
      cmdInfo.ForeColor = &H8000000E
      cmdInfo.FontSize = BtnFnt - 5.6
    End Select
    Select Case tog
    Case Is < 0, Is > 9
      TogState = Not TogState
    End Select
  End If
''  DoEvents
End Sub

Private Sub GetRcpInfo()
  Dim RP As Integer, lenRP As Integer, RP1 As Integer
  Dim RcptPrnFile As ReceiptPRNType
  
  On Local Error GoTo ERRORSTUFF
  
  RP1 = FreeFile
  lenRP = Len(RcptPrnFile)
  If Exist(RcptFileName$) Then
    Open RcptFileName$ For Random Shared As RP1 Len = lenRP
    Get RP1, 1, RcptPrnFile
    RecpPort = QPTrim(RcptPrnFile.RcpPort)
    If RcptPrnFile.PrnDefYN = 0 Then
      RecpDef = 0
    Else
      On Local Error GoTo nofound
      RP = FreeFile
      Open RecpPort For Output As RP
      Close RP
      RecpDef = 1
    End If
    If RcptPrnFile.CtlDefYN = 0 Then
      CntrlDef = 0
    Else
      CntrlDef = 1
    End If
    If RcptPrnFile.RValidate = 1 Then
      ValiDef = 1
      RctValidate = True
      GetUBBankINfo
    Else
      ValiDef = 0
      RctValidate = False
    End If
  Close RP1
  Else
    RecpDef = 99
    ValiDef = 0
  End If
Exit Sub
nofound:
  RecpDef = 99
  ValiDef = 0
 Exit Sub
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPaymentEntry", "GetRcpInfo", Erl)
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
 '   ClearInUse PWcnt
 '   CMTerminate

End Sub

Private Function Check4ValidPaidEntries() As Boolean '8/12/05

  On Local Error GoTo ERRORSTUFF
  
  Check4ValidPaidEntries = True
  If CDbl(fpCurrPersPaid.Value) > CDbl(fpCurrPersOwed.Value) Then
    Call VATaxMsg(800, "The amount entered for personal paid cannot be more than the amount owed for personal. Press F9 to correct this problem.")
    fpCurrPersPaid.SetFocus
    Check4ValidPaidEntries = False
    Exit Function
  End If
  If CDbl(fpCurrMTPaid.Value) > CDbl(fpCurrMTOwed.Value) Then
    Call VATaxMsg(800, "The amount entered for machine tools paid cannot be more than the amount owed for machine tools. Press F9 to correct this problem.")
    fpCurrMTPaid.SetFocus
    Check4ValidPaidEntries = False
    Exit Function
  End If
  
  If CDbl(fpCurrMCPaid.Value) > CDbl(fpCurrMCOwed.Value) Then
    Call VATaxMsg(800, "The amount entered for merchant capital paid cannot be more than the amount owed for merchant capital. Press F9 to correct this problem.")
    fpCurrMCPaid.SetFocus
    Check4ValidPaidEntries = False
    Exit Function
  End If
  
  If CDbl(fpCurrFEPaid.Value) > CDbl(fpCurrFEOwed.Value) Then
    Call VATaxMsg(800, "The amount entered for farm equipment paid cannot be more than the amount owed for farm equipment. Press F9 to correct this problem.")
    fpCurrFEPaid.SetFocus
    Check4ValidPaidEntries = False
    Exit Function
  End If
  
  If CDbl(fpCurrMHPaid.Value) > CDbl(fpCurrMHOwed.Value) Then
    Call VATaxMsg(800, "The amount entered for mobile homes paid cannot be more than the amount owed for mobile homes. Press F9 to correct this problem.")
    fpCurrMHPaid.SetFocus
    Check4ValidPaidEntries = False
    Exit Function
  End If
  
  If CDbl(fpCurrIntPaid.Value) > CDbl(fpCurrIntOwed.Value) Then
    Call VATaxMsg(800, "The amount entered for interest paid cannot be more than the amount owed for interest. Press F9 to correct this problem.")
    fpCurrIntPaid.SetFocus
    Check4ValidPaidEntries = False
    Exit Function
  End If
  
  If CDbl(fpCurrPenPaid.Value) > CDbl(fpCurrPenOwed.Value) Then
    Call VATaxMsg(800, "The amount entered for penalty paid cannot be more than the amount owed for penalty. Press F9 to correct this problem.")
    fpCurrPenPaid.SetFocus
    Check4ValidPaidEntries = False
    Exit Function
  End If
  
  If CDbl(fpCurrOpt1Paid.Value) > CDbl(fpCurrOpt1Owed.Value) Then
    Call VATaxMsg(800, "The amount entered for " + Opt1Desc + " paid cannot be more than the amount owed for " + Opt1Desc + ". Press F9 to correct this problem.")
    fpCurrOpt1Paid.SetFocus
    Check4ValidPaidEntries = False
    Exit Function
  End If
  
  If CDbl(fpCurrOpt2Paid.Value) > CDbl(fpCurrOpt2Owed.Value) Then
    Call VATaxMsg(800, "The amount entered for " + Opt2Desc + " paid cannot be more than the amount owed for " + Opt2Desc + ". Press F9 to correct this problem.")
    fpCurrOpt2Paid.SetFocus
    Check4ValidPaidEntries = False
    Exit Function
  End If
  
  If CDbl(fpCurrOpt3Paid.Value) > CDbl(fpCurrOpt3Owed.Value) Then
    Call VATaxMsg(800, "The amount entered for " + Opt3Desc + " paid cannot be more than the amount owed for " + Opt3Desc + ". Press F9 to correct this problem.")
    fpCurrOpt3Paid.SetFocus
    Check4ValidPaidEntries = False
    Exit Function
  End If
  
  Exit Function
  
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPersPaymentEntry", "Check4ValidPaidEntries", Erl)
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
 '   ClearInUse PWcnt
  '  CMTerminate


End Function
Private Sub PostEmTax()
  Dim Oper$
  Dim TaxPaymentRec As VATaxPaymentRecType
  Dim PHandle As Integer
  Dim NumOfPRecs As Integer
  Dim PayListRec As VAPersPayListType
  Dim LHandle As Integer
  Dim NumOfLRecs As Integer
  Dim TaxCustRec As VATaxCustType
  Dim CHandle As Integer
  Dim NumOfCRecs As Long
  Dim TaxTranRec As VATaxTransactionType
  Dim TaxTranHandle As Integer
  Dim NumOfTaxTranRecs As Long, CmNum As Long
  Dim PayTranRec As VATaxTransactionType
  Dim PayTranHandle As Integer
  Dim NumOfPayTranRecs As Long
  Dim EmptyPay As VATaxTransactionType
  Dim cnt&, TotalPaid#
  Dim ThisListRec&
  Dim NextTransRec&
  Dim CMTrRecLen As Integer, CMHandle As Integer
 ' tottotpaid# = 0
  On Local Error GoTo ERRORSTUFF
  Oper$ = QPTrim$(Str$(OperNum))
  If Not Exist("CMXPCPR" + Oper$ + ".DAT") Then Exit Sub
  TXLog ("OpenTempPersPay CMXPCPR" + Oper$)
  OpenTempPersPayFile PHandle, OperNum ' is the same as open TaxCPRFileName
  NumOfPRecs = LOF(PHandle) / Len(TaxPaymentRec)
  
  OpenPersPayListFile LHandle, OperNum 'is the same as open TaxLOPFileName
  NumOfLRecs = LOF(LHandle) / Len(PayListRec)
  
  OpenVATaxCustFile CHandle, NumOfCRecs
  OpenVATaxTransFile TaxTranHandle, NumOfTaxTranRecs
  
  ReDim CMTrRec(1) As CMTransRecType
  CMTrRecLen = Len(CMTrRec(1))
  CMHandle = FreeFile
  Open UBPath$ + "CMTRANS.DAT" For Random Access Read Write Shared As CMHandle Len = CMTrRecLen
    TXLog (" PayRecNo Check totals " + Str$(ThisListRec&) + " for " + Oper$)
   'do this to make sure balance has not changed since brought cust on screen....
  Get PHandle, 1, TaxPaymentRec
    If GetNewTot <> Val(fpCurrAmtOwed.Value) Then
      GoTo ErrorNOGo
    End If
    ThisListRec& = TaxPaymentRec.LastPayRec
    Do While ThisListRec& > 0
      Get LHandle, ThisListRec&, PayListRec 'listrec is the list of bills being paid (some could be
      'multiple tags for a single customer
      'get paylist rec
      Get CHandle, TaxPaymentRec.CustAcct, TaxCustRec
      'get cust rec
      
      If PayListRec.BillRec < 0 Then
        TXLog (" PayRecNo " + Str$(ThisListRec&) + " for " + Oper$ + " sent to prepay")
        GoSub PrePay
        GoTo SkipThisRec
      End If
      
      Get TaxTranHandle, PayListRec.BillRec, TaxTranRec
      'get bill trans this payrec is for
      TotalPaid# = 0
      PayTranRec = EmptyPay
      'make a new clean payment trans
'      TotalPaid# = OldRound#(PayListRec.DiscAmt + PayListRec.Principle1 + PayListRec.Interest1 + PayListRec.Collection + PayListRec.LateList)
      TotalPaid# = OldRound#(PayListRec.Personal + PayListRec.MachTools + PayListRec.MerchCap + PayListRec.FarmEquip) '4/22/05
      TotalPaid# = OldRound(TotalPaid# + PayListRec.MobHomes + PayListRec.Interest + PayListRec.Penalty + PayListRec.PrePayAmt)
      TotalPaid# = OldRound(TotalPaid# + PayListRec.Opt1 + PayListRec.Opt2 + PayListRec.Opt3)
      If TotalPaid# = 0 Then
        GoTo SkipThisRec
      End If
      'PayTranRec = the new record for tax transaction records
      PayTranRec.TransDate = TaxPaymentRec.payDate
      If PayListRec.PrePayAmt > 0 Then
        PayTranRec.TranType = 21 'overpay and bill pay combined
      Else
        PayTranRec.TranType = 2
      End If
'      PayTranRec.Revenue.Principle1Pd = OldRound(PayListRec.Principle1 + PayListRec.DiscAmt)
      PayTranRec.Revenue.Principle1Pd = PayListRec.Personal '4/22/05
      PayTranRec.Revenue.Principle2Pd = PayListRec.MachTools
      PayTranRec.Revenue.Principle3Pd = PayListRec.MerchCap
      PayTranRec.Revenue.Principle4Pd = PayListRec.FarmEquip
      PayTranRec.Revenue.Principle5Pd = PayListRec.MobHomes
      PayTranRec.Revenue.InterestPd = PayListRec.Interest
      PayTranRec.Revenue.PenaltyPd = PayListRec.Penalty
      PayTranRec.Revenue.CollectionPd = 0
      PayTranRec.Revenue.LateListPd = 0
      PayTranRec.Revenue.RevOpt1Pd = PayListRec.Opt1
      PayTranRec.Revenue.RevOpt2Pd = PayListRec.Opt2
      PayTranRec.Revenue.RevOpt3Pd = PayListRec.Opt3
      PayTranRec.CustPin = TaxCustRec.PIN
      PayTranRec.DiscXDate = TaxTranRec.DiscXDate
      PayTranRec.RealPin = ""
      PayTranRec.PersPin = QPTrim$(TaxTranRec.PersPin)
      PayTranRec.Posted2GL = "N"
      PayTranRec.TaxYear = TaxTranRec.TaxYear
      PayTranRec.DiscAmt = PayListRec.DiscAmt
      PayTranRec.OperNum = OperNum
      PayTranRec.Amount = TotalPaid#
      If QPTrim$(PayListRec.Description) = "" Then
        PayTranRec.Description = "CM-" + TaxTranRec.Description
      Else
        PayTranRec.Description = "CM-" + QPTrim$(PayListRec.Description) + " " + ParseBillNum$(TaxTranRec.Description)
      End If
      PayTranRec.CustomerRec = TaxPaymentRec.CustAcct
      PayTranRec.LastTrans = TaxCustRec.LastTrans
      PayTranRec.BelongTo = PayListRec.BillRec
      PayTranRec.Revenue.PrePaidAmt = PayListRec.PrePayAmt
      PayTranRec.Revenue.PrePaidUsed = 0
      PayTranRec.Revenue.PrePaidBal = OldRound(VAGetOverPayBalance(TaxPaymentRec.CustAcct, "N") + PayTranRec.Revenue.PrePaidAmt)
      PayTranRec.InternalPin = TaxTranRec.InternalPin
      PayTranRec.BillType = TaxPaymentRec.BillType
      'TaxTranRec is the update to the existing tax record
'      TaxTranRec.Revenue.Principle1Pd = OldRound#(TaxTranRec.Revenue.Principle1Pd + PayListRec.Principle1 + PayListRec.DiscAmt)
      TaxTranRec.Revenue.Principle1Pd = OldRound#(TaxTranRec.Revenue.Principle1Pd + PayListRec.Personal) '4/22/05
      TaxTranRec.Revenue.Principle2Pd = OldRound#(TaxTranRec.Revenue.Principle2Pd + PayListRec.MachTools)
      TaxTranRec.Revenue.Principle3Pd = OldRound#(TaxTranRec.Revenue.Principle3Pd + PayListRec.MerchCap)
      TaxTranRec.Revenue.Principle4Pd = OldRound#(TaxTranRec.Revenue.Principle4Pd + PayListRec.FarmEquip)
      TaxTranRec.Revenue.Principle5Pd = OldRound#(TaxTranRec.Revenue.Principle5Pd + PayListRec.MobHomes)
      TaxTranRec.Revenue.InterestPd = OldRound#(TaxTranRec.Revenue.InterestPd + PayListRec.Interest)
      TaxTranRec.Revenue.PenaltyPd = OldRound#(TaxTranRec.Revenue.PenaltyPd + PayListRec.Penalty)
      TaxTranRec.Revenue.Future1Pd = OldRound#(TaxTranRec.DiscAmt + PayListRec.DiscAmt)
      
      TaxTranRec.DiscAmt = OldRound#(TaxTranRec.DiscAmt + PayListRec.DiscAmt)
      TaxTranRec.Revenue.RevOpt1Pd = OldRound#(TaxTranRec.Revenue.RevOpt1Pd + PayListRec.Opt1)
      TaxTranRec.Revenue.RevOpt2Pd = OldRound#(TaxTranRec.Revenue.RevOpt2Pd + PayListRec.Opt2)
      TaxTranRec.Revenue.RevOpt3Pd = OldRound#(TaxTranRec.Revenue.RevOpt3Pd + PayListRec.Opt3)
      Put TaxTranHandle, PayListRec.BillRec, TaxTranRec
      NextTransRec& = (LOF(TaxTranHandle) \ Len(TaxTranRec)) + 1

      Put TaxTranHandle, NextTransRec&, PayTranRec
      TaxCustRec.LastTrans = NextTransRec&
      Put CHandle, TaxPaymentRec.CustAcct, TaxCustRec

SkipThisRec:
      ThisListRec& = PayListRec.PrevListRec
    Loop
 ' Next

  GoSub dotheCMsave
  Close
  WasSaved = True
  Call VASavemsg(900, "Personal transaction posting has completed successfully.")

'  KillFile ("TAXPCPR" + Oper$ + ".DAT")
'  KillFile ("TAXPLOP" + Oper$ + ".DAT")
'  MainLog ("Personal payment post completed successfully.")
'  Call cmdExit_Click
  
  Exit Sub

PrePay:
  TotalPaid# = 0
  PayTranRec = EmptyPay
  'make a new clean payment trans
  TotalPaid# = OldRound#(PayListRec.DiscAmt + PayListRec.Personal + PayListRec.MachTools + PayListRec.MerchCap + PayListRec.FarmEquip)
  TotalPaid# = OldRound(TotalPaid# + PayListRec.MobHomes + PayListRec.Interest + PayListRec.Penalty + PayListRec.PrePayAmt)
  If TotalPaid# = 0 Then
    GoTo SkipThisRec
  End If
  'PayTranRec = the new record for tax transaction records
  PayTranRec.TransDate = TaxPaymentRec.payDate
  PayTranRec.TranType = 22 'overpay only
  PayTranRec.Revenue.Principle1Pd = OldRound(PayListRec.Personal + PayListRec.DiscAmt)
  PayTranRec.Revenue.Principle2Pd = PayListRec.MachTools
  PayTranRec.Revenue.Principle3Pd = PayListRec.MerchCap
  PayTranRec.Revenue.Principle4Pd = PayListRec.FarmEquip
  PayTranRec.Revenue.Principle5Pd = PayListRec.MobHomes
  PayTranRec.Revenue.InterestPd = PayListRec.Interest
  PayTranRec.Revenue.PenaltyPd = PayListRec.Penalty
  PayTranRec.Revenue.CollectionPd = 0
  PayTranRec.Revenue.LateListPd = 0
  PayTranRec.Revenue.RevOpt1Pd = 0
  PayTranRec.Revenue.RevOpt2Pd = 0
  PayTranRec.Revenue.RevOpt3Pd = 0
  PayTranRec.CustPin = TaxCustRec.PIN
  PayTranRec.DiscXDate = TaxTranRec.DiscXDate
  PayTranRec.RealPin = ""
  PayTranRec.PersPin = QPTrim$(TaxTranRec.PersPin)
  PayTranRec.Posted2GL = "N"
  PayTranRec.TaxYear = TempPreYear
  PayTranRec.DiscAmt = PayListRec.DiscAmt
  PayTranRec.OperNum = OperNum
  PayTranRec.Amount = TotalPaid#
  If QPTrim$(PayListRec.Description) = "" Then
    PayTranRec.Description = "CM-Prepay"
  Else
    PayTranRec.Description = "CM-" + QPTrim$(PayListRec.Description)
  End If
  PayTranRec.CustomerRec = TaxPaymentRec.CustAcct
  PayTranRec.LastTrans = TaxCustRec.LastTrans
  PayTranRec.BelongTo = 0
  PayTranRec.Revenue.PrePaidAmt = PayListRec.PrePayAmt
  PayTranRec.Revenue.PrePaidUsed = 0
  PayTranRec.Revenue.PrePaidBal = OldRound(VAGetOverPayBalance(TaxPaymentRec.CustAcct, "N") + PayTranRec.Revenue.PrePaidAmt)
  PayTranRec.BillType = TaxPaymentRec.BillType
  NextTransRec& = (LOF(TaxTranHandle) \ Len(TaxTranRec)) + 1
  Put TaxTranHandle, NextTransRec&, PayTranRec
  
  TaxCustRec.LastTrans = NextTransRec&
  Put CHandle, TaxPaymentRec.CustAcct, TaxCustRec

  Return
dotheCMsave:
  CMTrRec(1).TransDate = TaxPaymentRec.payDate
  If Len(QPTrim$(TaxPaymentRec.Desc)) = 0 Then
    CMTrRec(1).TransDesc = "Tax P Billing Payment"
  Else
    CMTrRec(1).TransDesc = TaxPaymentRec.Desc
  End If
  If fpcmbTenderType.Text = "CHARGE" Then
    CMTrRec(1).TransCheck = TaxPaymentRec.ChrgAmt
    CMTrRec(1).TransTender = 4
  ElseIf fpcmbTenderType.Text = "CHECK" Then
    CMTrRec(1).TransCheck = TaxPaymentRec.ChkAmt
    CMTrRec(1).TransTender = 2
  ElseIf fpcmbTenderType.Text = "CASH AND CHECK" Then
    CMTrRec(1).TransCheck = TaxPaymentRec.ChkAmt
    CMTrRec(1).TransTender = 3
  Else
    CMTrRec(1).TransTender = 1
    CMTrRec(1).TransCheck = 0
  End If
 
  CMTrRec(1).TransCash = TaxPaymentRec.CashAmt
  CMTrRec(1).TransAmount = TaxPaymentRec.TotPaid
  CMTrRec(1).TransAmtOwed = TaxPaymentRec.AmtOwed
  CMTrRec(1).TransSource = 171
  CMTrRec(1).TransName = TaxPaymentRec.CustName
  CMTrRec(1).TransAcctNum = TaxPaymentRec.CustAcct
  CMTrRec(1).TransDetNum = 0
  CMTrRec(1).TransOperNum = OperNum
  CMTrRec(1).TransPad = ""
  CMTrRec(1).TransRevAmt(1) = TaxPaymentRec.PaidOwed(1).AmtPaid    'Princeple1
  CMTrRec(1).TransRevAmt(2) = TaxPaymentRec.PaidOwed(2).AmtPaid    'Princeple2
  CMTrRec(1).TransRevAmt(3) = TaxPaymentRec.PaidOwed(3).AmtPaid    'Princeple3
  CMTrRec(1).TransRevAmt(4) = TaxPaymentRec.PaidOwed(4).AmtPaid    'Princeple4
  CMTrRec(1).TransRevAmt(5) = TaxPaymentRec.PaidOwed(5).AmtPaid    'Princeple5
  CMTrRec(1).TransRevAmt(6) = TaxPaymentRec.PaidOwed(6).AmtPaid    'Interest
  CMTrRec(1).TransRevAmt(7) = TaxPaymentRec.PaidOwed(7).AmtPaid    'Penalty
  CMTrRec(1).TransRevAmt(8) = TaxPaymentRec.PaidOwed(8).AmtPaid    'Option1
  CMTrRec(1).TransRevAmt(9) = TaxPaymentRec.PaidOwed(9).AmtPaid    'Option2
  CMTrRec(1).TransRevAmt(10) = TaxPaymentRec.PaidOwed(10).AmtPaid    'Option3
  CMTrRec(1).TransRevAmt(11) = TaxPaymentRec.DiscAmt
  CMTrRec(1).TransRevAmt(12) = TaxPaymentRec.PrePayAmt
  CMTrRec(1).TransRevAmt(13) = CDbl(TaxPaymentRec.NumPayRec)  'numofbills paid
  CMTrRec(1).TransVoidNum = 0
  'CMTrRec(1).Post2GL = "Y"
  CMTrRec(1).ChkByte = Chr$(1)
  CmNum = (LOF(CMHandle) / CMTrRecLen) + 1
  Put CMHandle, CmNum, CMTrRec(1)
  CMLog ("Oper " + Str$(OperNum) + " TX Pay postedCM for trans# " + Str$(CmNum))
  Close CMHandle
  TXLog ("CMTrans posted. " + Str$(OperNum) + "," + Str$(TaxPaymentRec.CustAcct))
  
Return
ErrorNOGo:
    
    TXLog ("NOGo encountered. Oper: " + Str$(OperNum) + "," + Str$(TaxPaymentRec.CustAcct))
    CMLog ("NOGo encountered. Oper: " + Str$(OperNum) + "," + Str$(TaxPaymentRec.CustAcct))
    WasSaved = False
    Close
    Exit Sub
Return
ERRORSTUFF:
   Select Case ErrorMessage(Err.Number, Err.Description, Err.Source, "frmVATaxPayPost", "cmdPostPers_Click", Erl)
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
 '   ClearInUse PWcnt
  '  CMTerminate

End Sub

Private Function GetNewTot()
  Dim x As Integer
  Dim TaxOwed#
  Dim IntOwed#
  Dim ColOwed#
  Dim LLOwed#
  Dim PenOwed#
  Dim RevOpt1#
  Dim RevOpt2#
  Dim RevOpt3#
  Dim TransRec As VATaxTransactionType
  Dim THandle As Integer
  Dim NumOfTRecs As Long
  Dim MasterRec As VATaxMasterType
  Dim MHandle As Integer
  Dim ThisTaxYear As Integer
  Dim Message$
  Dim ThisBal As Double
  Dim DiscCheck As Integer
  Dim Dif As Double
  Dim xx As Double
  xx = 0
  OpenVATaxSetUpFile MHandle
  Get MHandle, 1, MasterRec
  Close MHandle
  ThisTaxYear = MasterRec.PTaxYear
  
  ThisDiscPct = MasterRec.DisPPct
  
  TaxOwed# = 0
  IntOwed# = 0
  ColOwed# = 0
  LLOwed# = 0
  PenOwed# = 0
  RevOpt1# = 0
  RevOpt2# = 0
  RevOpt3# = 0
  OpenVATaxTransFile THandle, NumOfTRecs
  
  For x = 1 To BillCnt
    Get THandle, BillTrans(x), TransRec
      TaxOwed# = OldRound(TaxOwed# + TransRec.Revenue.Principle1 + TransRec.Revenue.Principle2 + TransRec.Revenue.Principle3 + TransRec.Revenue.Principle4 + TransRec.Revenue.Principle5)
      TaxOwed# = OldRound(TaxOwed# - TransRec.PPTRADisc + TransRec.PPTRARmvl)
      TaxOwed# = OldRound(TaxOwed# - (TransRec.Revenue.Principle1Pd + TransRec.Revenue.Principle2Pd + TransRec.Revenue.Principle3Pd + TransRec.Revenue.Principle4Pd + TransRec.Revenue.Principle5Pd))
      IntOwed# = OldRound(IntOwed# + TransRec.Revenue.Interest)
      IntOwed# = OldRound(IntOwed# - TransRec.Revenue.InterestPd)
      ColOwed# = OldRound(ColOwed# + TransRec.Revenue.Collection)
      ColOwed# = OldRound(ColOwed# - TransRec.Revenue.CollectionPd)
      LLOwed# = OldRound(LLOwed# + TransRec.Revenue.LateList)
      LLOwed# = OldRound(LLOwed# - TransRec.Revenue.LateListPd)
      PenOwed# = OldRound(PenOwed# + TransRec.Revenue.Penalty)
      PenOwed# = OldRound(PenOwed# - TransRec.Revenue.PenaltyPd)
      RevOpt1# = OldRound(RevOpt1# + TransRec.Revenue.RevOpt1)
      RevOpt1# = OldRound(RevOpt1# - TransRec.Revenue.RevOpt1Pd)
      RevOpt2# = OldRound(RevOpt2# + TransRec.Revenue.RevOpt2)
      RevOpt2# = OldRound(RevOpt2# - TransRec.Revenue.RevOpt2Pd)
      RevOpt3# = OldRound(RevOpt3# + TransRec.Revenue.RevOpt3)
      RevOpt3# = OldRound(RevOpt3# - TransRec.Revenue.RevOpt3Pd)
  Next x
  
  xx = OldRound(TaxOwed# + IntOwed# + ColOwed# + LLOwed# + PenOwed# + RevOpt1# + RevOpt2# + RevOpt3#)
  GetNewTot = Val(xx)
  Close THandle
End Function


