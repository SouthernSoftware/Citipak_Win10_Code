VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#3.5#0"; "SPR32X35.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmPRDCalcScr3 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PayRoll Calculation"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   585
   ClientWidth     =   11655
   ControlBox      =   0   'False
   Icon            =   "frmPRDCalcScr3.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   355
      Left            =   0
      Top             =   0
   End
   Begin FPSpread.vaSpread vaSpreadDedData 
      Height          =   1815
      Left            =   675
      TabIndex        =   2
      Top             =   3150
      Width           =   10320
      _Version        =   196613
      _ExtentX        =   18203
      _ExtentY        =   3201
      _StockProps     =   64
      ColsFrozen      =   6
      DisplayRowHeaders=   0   'False
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   6
      MaxRows         =   17
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      ShadowColor     =   13684944
      SpreadDesigner  =   "frmPRDCalcScr3.frx":08CA
   End
   Begin EditLib.fpCurrency fpcurrGrossPay 
      Height          =   390
      Left            =   765
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6915
      Width           =   1830
      _Version        =   196608
      _ExtentX        =   3238
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
      CurrencyDecimalPlaces=   2
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
   Begin EditLib.fpCurrency fpcurrTaxRetire 
      Height          =   390
      Left            =   8880
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5910
      Width           =   1590
      _Version        =   196608
      _ExtentX        =   2805
      _ExtentY        =   688
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
      CurrencyDecimalPlaces=   2
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
   Begin EditLib.fpCurrency fpcurrTaxFed 
      Height          =   390
      Left            =   4890
      TabIndex        =   3
      Top             =   5910
      Width           =   1590
      _Version        =   196608
      _ExtentX        =   2805
      _ExtentY        =   688
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   2
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
   Begin EditLib.fpCurrency fpcurrEarnings 
      Height          =   390
      Index           =   1
      Left            =   1020
      TabIndex        =   1
      Top             =   2160
      Width           =   1590
      _Version        =   196608
      _ExtentX        =   2805
      _ExtentY        =   688
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   2
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
   Begin EditLib.fpText fptxtGrossPay 
      Height          =   396
      Left            =   1344
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   15264
      Width           =   1836
      _Version        =   196608
      _ExtentX        =   3238
      _ExtentY        =   698
      Enabled         =   -1  'True
      MousePointer    =   0
      Object.TabStop         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   13.5
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
   Begin EditLib.fpText fptxtEmpName 
      Height          =   345
      Left            =   4845
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Width           =   4335
      _Version        =   196608
      _ExtentX        =   7646
      _ExtentY        =   609
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
   Begin EditLib.fpText fptxtEmpNum 
      Height          =   345
      Left            =   3435
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   720
      Width           =   1410
      _Version        =   196608
      _ExtentX        =   2487
      _ExtentY        =   609
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
   Begin EditLib.fpCurrency fpcurrEarnings 
      Height          =   390
      Index           =   2
      Left            =   3045
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1590
      _Version        =   196608
      _ExtentX        =   2805
      _ExtentY        =   688
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
      CurrencyDecimalPlaces=   2
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
   Begin EditLib.fpCurrency fpcurrEarnings 
      Height          =   390
      Index           =   3
      Left            =   5055
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1590
      _Version        =   196608
      _ExtentX        =   2805
      _ExtentY        =   688
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
      CurrencyDecimalPlaces=   2
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
   Begin EditLib.fpCurrency fpcurrEarnings 
      Height          =   390
      Index           =   4
      Left            =   7095
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1590
      _Version        =   196608
      _ExtentX        =   2805
      _ExtentY        =   688
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
      CurrencyDecimalPlaces=   2
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
   Begin EditLib.fpCurrency fpcurrEarnings 
      Height          =   390
      Index           =   5
      Left            =   9045
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1590
      _Version        =   196608
      _ExtentX        =   2805
      _ExtentY        =   688
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
      CurrencyDecimalPlaces=   2
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
   Begin EditLib.fpCurrency fpcurrTaxSS 
      Height          =   390
      Left            =   870
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5910
      Width           =   1590
      _Version        =   196608
      _ExtentX        =   2805
      _ExtentY        =   688
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
      CurrencyDecimalPlaces=   2
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
   Begin EditLib.fpCurrency fpcurrTaxMed 
      Height          =   390
      Left            =   2880
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5910
      Width           =   1590
      _Version        =   196608
      _ExtentX        =   2805
      _ExtentY        =   688
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
      CurrencyDecimalPlaces=   2
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
   Begin EditLib.fpCurrency fpcurrTaxState 
      Height          =   390
      Left            =   6915
      TabIndex        =   4
      Top             =   5910
      Width           =   1590
      _Version        =   196608
      _ExtentX        =   2805
      _ExtentY        =   688
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
      Text            =   "$0.00"
      CurrencyDecimalPlaces=   2
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
   Begin EditLib.fpCurrency fpcurrTotDed 
      Height          =   390
      Left            =   3510
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6915
      Width           =   1830
      _Version        =   196608
      _ExtentX        =   3238
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
      CurrencyDecimalPlaces=   2
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
   Begin EditLib.fpCurrency fpcurrAdvEIC 
      Height          =   390
      Left            =   6195
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6915
      Width           =   1830
      _Version        =   196608
      _ExtentX        =   3238
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
      CurrencyDecimalPlaces=   2
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
   Begin EditLib.fpCurrency fpcurrNetPay 
      Height          =   390
      Left            =   8970
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6915
      Width           =   1830
      _Version        =   196608
      _ExtentX        =   3238
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
      CurrencyDecimalPlaces=   2
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
   Begin fpBtnAtlLibCtl.fpBtn cmdCont 
      Height          =   540
      Left            =   8640
      TabIndex        =   37
      TabStop         =   0   'False
      ToolTipText     =   "Press to complete this employee's payroll process."
      Top             =   7650
      Width           =   1695
      _Version        =   131072
      _ExtentX        =   2990
      _ExtentY        =   952
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
      ButtonDesigner  =   "frmPRDCalcScr3.frx":0EC9
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdDelete 
      Height          =   540
      Left            =   6240
      TabIndex        =   38
      TabStop         =   0   'False
      ToolTipText     =   "Press to remove this employee from the current payroll process."
      Top             =   7650
      Width           =   1695
      _Version        =   131072
      _ExtentX        =   2990
      _ExtentY        =   952
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
      ButtonDesigner  =   "frmPRDCalcScr3.frx":10E1
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Taxes"
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
      Height          =   345
      Left            =   570
      TabIndex        =   27
      Top             =   5160
      Width           =   780
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      Height          =   1305
      Left            =   570
      Top             =   5175
      Width           =   10530
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Deductions"
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
      Height          =   345
      Left            =   570
      TabIndex        =   26
      Top             =   2729
      Width           =   1710
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Earnings"
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
      Height          =   345
      Left            =   570
      TabIndex        =   20
      Top             =   1440
      Width           =   1170
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Net Pay"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   345
      Left            =   8925
      TabIndex        =   36
      Top             =   6525
      Width           =   1830
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Adv EIC"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   345
      Left            =   6195
      TabIndex        =   35
      Top             =   6525
      Width           =   1830
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Total Ded"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   345
      Left            =   3510
      TabIndex        =   34
      Top             =   6525
      Width           =   1830
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Gross Pay"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   345
      Left            =   810
      TabIndex        =   33
      Top             =   6525
      Width           =   1830
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Retire"
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
      Left            =   8970
      TabIndex        =   32
      Top             =   5565
      Width           =   1590
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "State W/H"
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
      Left            =   6915
      TabIndex        =   31
      Top             =   5565
      Width           =   1590
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fed W/H"
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
      Left            =   4890
      TabIndex        =   30
      Top             =   5565
      Width           =   1590
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Medicare"
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
      Left            =   2880
      TabIndex        =   29
      Top             =   5565
      Width           =   1590
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Soc Sec"
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
      TabIndex        =   28
      Top             =   5565
      Width           =   1590
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      Height          =   2460
      Left            =   570
      Top             =   2725
      Width           =   10530
   End
   Begin VB.Label lblEarnDesc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "EarnCode3"
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
      Height          =   300
      Index           =   3
      Left            =   9135
      TabIndex        =   25
      Top             =   1830
      Width           =   1455
   End
   Begin VB.Label lblEarnDesc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "EarnCode2"
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
      Height          =   300
      Index           =   2
      Left            =   7185
      TabIndex        =   24
      Top             =   1830
      Width           =   1455
   End
   Begin VB.Label lblEarnDesc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "EarnCode1"
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
      Height          =   300
      Index           =   1
      Left            =   5160
      TabIndex        =   23
      Top             =   1830
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OT Earnings"
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
      Height          =   300
      Left            =   3135
      TabIndex        =   22
      Top             =   1830
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reg Earnings"
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
      Height          =   300
      Index           =   0
      Left            =   1110
      TabIndex        =   21
      Top             =   1830
      Width           =   1455
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      Height          =   1308
      Left            =   564
      Top             =   1440
      Width           =   10524
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Payroll Calculation"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   3408
      TabIndex        =   19
      Top             =   360
      Width           =   4620
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      Height          =   972
      Index           =   1
      Left            =   1392
      Top             =   252
      Width           =   8652
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Employee:"
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
      Left            =   2040
      TabIndex        =   18
      Top             =   768
      Width           =   1260
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D0D0D0&
      FillStyle       =   0  'Solid
      Height          =   1044
      Left            =   1392
      Top             =   192
      Width           =   8652
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuPrnScn 
         Caption         =   "Prin&t Screen"
      End
      Begin VB.Menu mnuCont 
         Caption         =   "&Continue"
      End
   End
End
Attribute VB_Name = "frmPRDCalcScr3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim Over As clsTextBoxOverRider
  Private Temp_Class As Resize_Class
  Dim TransRec(1) As TransRecType
  Dim THandle As Integer
  Dim OldFldVal(1 To 50) As Double
  Dim PayType$
  Dim OldRegWages#
  Dim OldRetAmt#
  Dim OldAdvEIC#
  Dim OldFedTax#
  Dim OldStateTax# 'added 8/22
  Dim DifRetAmt#
  Dim ScrnCalc(1) As ScrnCalcType
  Dim ZeroFlag As Boolean
  Dim DedCnt As Integer
Private Sub cmdCont_Click()
  Dim ECnt As Integer
  Dim NumOfErns As Integer
  Dim ErnHandle As Integer
  Dim ErnRec As ErnCodeRecType
  Dim x As Integer
   
  If fpcurrNetPay < 0 Then 'negative ney pay is not permitted
    frmWarnNegNetPay.Show vbModal, Me
    Exit Sub
  End If
  
  If fpcurrNetPay = 0 Then 'this if statement alerts the
  'user to a total of zero net pay...although acceptable this
  'situation is unusual...it is set up so the screen returns
  'as is but allows the user to exit with the next Continue
  'activation
    If ZeroFlag = True Then GoTo ZFlagOn 'ZFlagOn is true
    'if we've already been alerted once so we allow the
    'program to proceed
    frmWarnZeroNetPay.Show vbModal, Me
    ZeroFlag = True
    Exit Sub
  End If
ZFlagOn:
  OpenTransWorkFile TRHandle
  Get TRHandle, RecNum, TransRec(1)
  ParseScrnCalc2Trans TransRec(), ScrnCalc()
  
  If OldRetAmt# <> fpcurrTaxRetire.Text Then 'a change has been made
    DifRetAmt# = OldRetAmt# - fpcurrTaxRetire.Text
    TransRec(1).FedGrossPay = OldRound(TransRec(1).FedGrossPay + DifRetAmt#)
    TransRec(1).StaGrossPay = OldRound(TransRec(1).StaGrossPay + DifRetAmt#)
  End If
  
  TransRec(1).TActive = True 'this employee is now officially
  'set up to be paid (TActive = Transaction is Active)
  TransRec(1).Less401k(1) = False 'false means ok to match
  TransRec(1).Less401k(2) = False
  TransRec(1).Less401k(3) = False
  If Exist("PRDATA\PRERNCOD.DAT") Then
    OpenErnCodeFile ErnHandle
    NumOfErns = LOF(ErnHandle) / Len(ErnRec)
    For x = 1 To 3
      Get ErnHandle, x, ErnRec
      If QPTrim$(ErnRec.EarnYN) = "N" Then 'added 10/07
        TransRec(1).Less401k(x) = True
      End If
    Next x
  End If
  Close ErnHandle
  
  Put TRHandle, RecNum, TransRec(1)
  Close TRHandle
  Close
  NewListFlag = True 'tells the lookup process to reload
  'the list with this employee as active
  Call frmPRTPrevEmpLookUp.cmdPickList_Click
  frmPRTPrevEmpLookUp.Show
  frmPRTPrevEmpLookUp.fpList.SetFocus 'added 9/8/04
  DoEvents
  Unload frmPRDCalcScr3
  MainLog ("Edit transaction completed in frmPRDCalcScr3.")
End Sub

Private Sub cmdDelete_Click()
  Dim DoWhatFlag As PRTRemove
  Dim THandle As Integer, TRec As TransRecType
  
  DoWhatFlag = PromptPRTRemove(Me) 'are you sure you want to delete?
  Select Case DoWhatFlag
  Case PRTRemove.prtrEscape
     Exit Sub
  Case PRTRemove.prtrDelete
     Call DeleteThisEmp
  End Select
  
  NewListFlag = True
  Call frmPRTPrevEmpLookUp.cmdPickList_Click
  frmPRTPrevEmpLookUp.Show
  DoEvents
  Unload frmPRDCalcScr3
  MainLog ("Edit transaction deleted in frmPRDCalcScr3.")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%C"
      Call cmdCont_Click
      KeyCode = 0
    Case vbKeyF3:
      SendKeys "%D"
      Call cmdDelete_Click
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_Load()
  Dim PauseTime, Start, Finish, TotalTime
  PauseTime = 0.5  ' Set duration.
  Start = Timer   ' Set start time.
  DoEvents
  Do While Timer < Start + PauseTime
     DoEvents   ' Yield to other processes.
  Loop
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  ZeroFlag = False
  Call FixSpread
  DoEvents
  Call LoadThisForm
  DoEvents
  Me.Show
  
End Sub

Private Sub Form_Resize()
  If Me.Visible = True Then
    If Me.WindowState <> vbMinimized Then
      Me.Visible = False
      Temp_Class.ResizeControls Me
      Me.Visible = True
      Me.SetFocus
    End If
  End If
End Sub

Private Sub LoadThisForm()
  
  Dim cnt As Integer, Nextx As Integer
  Dim EarnDes$(1 To 3)
  Dim DedDes$(1 To 50)
  Dim ErnCodes(1 To 3) As ErnCodeRecType
  Dim ErnHandle As Integer
  Dim DedCodes(1 To 50) As DedCodeRecType
  Dim DedRec As DedCodeRecType
  Dim DedHandle As Integer
  Dim x As Integer
  Dim EmpHandle As Integer
  Dim Emp2Rec(1) As EmpData2Type
  Dim EmpNum$, EmpName$
  Dim TransRec(1) As TransRecType
  Dim Image$
  
  Image = "$###,##0.00"
  OpenEmpData2File EmpHandle
  Get EmpHandle, RecNum, Emp2Rec(1)
  Close EmpHandle
  
  EmpNum$ = QPTrim$(Emp2Rec(1).EmpNo)
  fptxtEmpNum.Text = EmpNum$
  EmpName$ = QPTrim$(Emp2Rec(1).EmpFName) & " " & QPTrim$(Emp2Rec(1).EmpLName)
  fptxtEmpName.Text = EmpName$
  
  OpenTransWorkFile TRHandle
  Get TRHandle, RecNum, TransRec(1)

  PayType$ = TransRec(1).PayType
  
  OpenEmpData2File EHandle
  Get EHandle, RecNum, Emp2Rec(1)
  Close EHandle
  MainLog ("Payroll Calculation screen for " + QPTrim(Emp2Rec(1).EmpFName) + " " + QPTrim(Emp2Rec(1).EmpLName) + " opened.")
  
  If TransRec(1).TActive = False Then 'CalcPay examines
  'the default settings if TActive is false because it
  'means this employee is being processed for the first time
  'in this pay period
    CalcPay TransRec(1), RecNum, False
  Else
    CalcPay TransRec(1), RecNum, True 'true means skip default
    'examination...it's already been done
  End If
  Put TRHandle, RecNum, TransRec(1)
  Close TRHandle
  
  ParseTrans2ScrnCalc TransRec(), ScrnCalc()
  
  fpcurrEarnings(1).Text = ScrnCalc(1).REGEARN
  OldRegWages# = fpcurrEarnings(1).Text 'set this so we
  'can compare it later against the current Earnings
  'value to see if any changes have taken place
  fpcurrEarnings(2).Text = ScrnCalc(1).OTEARN
  fpcurrEarnings(3).Text = ScrnCalc(1).ALTEARN1
  fpcurrEarnings(4).Text = ScrnCalc(1).ALTEARN2
  fpcurrEarnings(5).Text = ScrnCalc(1).ALTEARN3
  
  OpenDedCodeFile DedHandle
  DedCnt = LOF(DedHandle) \ Len(DedRec)
  For x = 1 To DedCnt
    Get DedHandle, x, DedCodes(x) 'load DedCodes array
    'with data for each deduction on file
  Next x
  Close DedHandle
  
  Nextx = 1 'this spreadsheet has 6 columns, 2 for each deduction.
  'we are actually processing all 6 columns of each row with
  'each iteration...when full the last row would only contain
  '4 columns so we would want to exit before we finish the
  'whole process
  For x = 1 To 52 'we want this code iterating thru
  'all 50 possible deductions even if there are less
  'than 50 because in the process we also make any
  'unused cells "CellTypeStaticText" which means they
  'cannot be edited...inadvertantly entering a value
  'in a column with no deduction description means
  'we'll have a value we don't know what to do with
  
    vaSpreadDedData.Col = 1 'column 1 gets description
    'which cannot be edited
    vaSpreadDedData.Row = Nextx
    vaSpreadDedData.Text = QPTrim$(DedCodes(x).DCDESC1)
    vaSpreadDedData.CellType = CellTypeStaticText
    If QPTrim$(DedCodes(x).DCDESC1) = "" Then 'if no
    'description then we have gone past the number
    'of descriptions on file...make this cell uneditable
      vaSpreadDedData.Col = 2
      vaSpreadDedData.CellType = CellTypeStaticText
      GoTo NoDesc1
    End If
    vaSpreadDedData.Col = 2 'column 2 gets values which
    'can be edited
    vaSpreadDedData.Row = Nextx
    vaSpreadDedData.Text = ScrnCalc(1).Ded(x)
    If ScrnCalc(1).Ded(x) > 0 Then 'make the value bold if
    'it is greater than zero
      vaSpreadDedData.Col = 1 'description goes bold too
      vaSpreadDedData.Row = Nextx
      vaSpreadDedData.FontBold = True
      vaSpreadDedData.Col = 2
      vaSpreadDedData.Row = Nextx
      vaSpreadDedData.FontBold = True
    End If
NoDesc1:
    vaSpreadDedData.Col = 3
    vaSpreadDedData.Row = Nextx
    vaSpreadDedData.Text = QPTrim$(DedCodes(x + 1).DCDESC1)
    vaSpreadDedData.CellType = CellTypeStaticText
    If QPTrim$(DedCodes(x + 1).DCDESC1) = "" Then
      vaSpreadDedData.Col = 4
      vaSpreadDedData.CellType = CellTypeStaticText
      GoTo NoDesc2
    End If
    vaSpreadDedData.Col = 4
    vaSpreadDedData.Row = Nextx
    vaSpreadDedData.Text = ScrnCalc(1).Ded(x + 1)
    If ScrnCalc(1).Ded(x + 1) > 0 Then
      vaSpreadDedData.Col = 3
      vaSpreadDedData.Row = Nextx
      vaSpreadDedData.FontBold = True
      vaSpreadDedData.Col = 4
      vaSpreadDedData.Row = Nextx
      vaSpreadDedData.FontBold = True
    End If
NoDesc2:
    If x + 2 > 50 Then Exit For
    vaSpreadDedData.Col = 5
    vaSpreadDedData.Row = Nextx
    vaSpreadDedData.CellType = CellTypeStaticText
    vaSpreadDedData.Text = QPTrim$(DedCodes(x + 2).DCDESC1)
    If QPTrim$(DedCodes(x + 2).DCDESC1) = "" Then
      vaSpreadDedData.Col = 6
      vaSpreadDedData.Row = Nextx
      vaSpreadDedData.CellType = CellTypeStaticText
      GoTo NoDesc3
    End If
    vaSpreadDedData.Col = 6
    vaSpreadDedData.Row = Nextx
    vaSpreadDedData.Text = ScrnCalc(1).Ded(x + 2)
    If ScrnCalc(1).Ded(x + 2) > 0 Then
      vaSpreadDedData.Col = 5
      vaSpreadDedData.Row = Nextx
      vaSpreadDedData.FontBold = True
      vaSpreadDedData.Col = 6
      vaSpreadDedData.Row = Nextx
      vaSpreadDedData.FontBold = True
    End If
NoDesc3:
    x = x + 2
    Nextx = Nextx + 1
  Next x
  
  'these last 2 columns will always be uneditable as
  'long as there are a maximum of 50 deductions
  vaSpreadDedData.Col = 5
  vaSpreadDedData.Row = 17
  vaSpreadDedData.CellType = CellTypeStaticText
  vaSpreadDedData.Col = 6
  vaSpreadDedData.Row = 17
  vaSpreadDedData.CellType = CellTypeStaticText
  
  Nextx = 1
  For x = 1 To 52 'load up OldFldVal array with current
  'data to be compared later to see if a change has been
  'made
    vaSpreadDedData.Col = 2
    vaSpreadDedData.Row = Nextx
    OldFldVal(x) = Val(ReplaceString(vaSpreadDedData.Text, "$", "")) 'RemoveDollarMark
    vaSpreadDedData.Col = 4
    vaSpreadDedData.Row = Nextx
    OldFldVal(x + 1) = Val(ReplaceString(vaSpreadDedData.Text, "$", "")) 'RemoveDollarMark
    vaSpreadDedData.Col = 6
    vaSpreadDedData.Row = Nextx
    If x + 2 > 50 Then Exit For
    OldFldVal(x + 2) = Val(ReplaceString(vaSpreadDedData.Text, "$", "")) 'RemoveDollarMark
    Nextx = Nextx + 1
    x = x + 2
  Next x
  
  'these text boxes are being loaded with the latest
  'pay calculations since we just ran CalcPay earlier
  'in this procedure
  fpcurrTaxSS.Text = ScrnCalc(1).SOCTAX
  fpcurrTaxMed.Text = ScrnCalc(1).MEDTAX
  fpcurrTaxFed.Text = ScrnCalc(1).FEDTAX
  OldFedTax# = fpcurrTaxFed.Text
  fpcurrTaxState.Text = ScrnCalc(1).STATAX
  OldStateTax# = fpcurrTaxState.Text 'added 8/22
  fpcurrTaxRetire.Text = ScrnCalc(1).RETIRE
  OldRetAmt# = fpcurrTaxRetire.Text
  fpcurrGrossPay.Text = ScrnCalc(1).GrossPay
  fpcurrTotDed.Text = ScrnCalc(1).TOTDED
  fpcurrAdvEIC.Text = ScrnCalc(1).EIC
  OldAdvEIC# = fpcurrAdvEIC.Text
  fpcurrNetPay.Text = ScrnCalc(1).NetPay

  If PayType$ = "H" Then
    fpcurrEarnings(1).Enabled = False
  End If

  OpenErnCodeFile ErnHandle
  For x = 1 To 3
    Get ErnHandle, x, ErnCodes(x)
    lblEarnDesc(x).Caption = QPTrim$(ErnCodes(x).ERNCODE1)
    EarnDes$(x) = QPTrim$(ErnCodes(x).ERNCODE1)
  Next x
  Close ErnHandle
 
  Exit Sub

End Sub

Private Sub ReCalcPay(TransRec() As TransRecType, ScrnCalc() As ScrnCalcType)
  Dim x As Integer
  'make sure all fields in the form are calculated
  Call CalcFields
  'copy all screen fields to scrncalc type variable
  Screen2ScrnCalc ScrnCalc()
  'copy scrn calcs fields to the transaction fields
  ParseScrnCalc2Trans TransRec(), ScrnCalc()
  'recalc tax's, net pay, deductions, etc.
  CalcPay TransRec(1), RecNum, True
  'copy transaction data back to scrn calc type variable
  ParseTrans2ScrnCalc TransRec(), ScrnCalc()
  'copy scrn calc type back to scrn calc form
  ScrnCalc2Screen ScrnCalc()
End Sub

Private Sub Screen2ScrnCalc(ScrnCalc() As ScrnCalcType)
   Dim x As Integer
   Dim Nextx As Integer
   Dim Change As Boolean
   
  '8/26/04...discovered that if a change is made after registers are run
  '(TEMPIF.DAT exists after registers are run but not before) then you are
  'allowed to print checks without re-running registers...this means that
  'transactions are OK but the post to the GL will be wrong...hence this
  'additional code to check for changes and if one is found then the TEMPIF.DAT
  'file is destroyed so that registers will have to be re-run before checks
  'can be reprinted or posting can take place
   Change = False
   Nextx = 1
   'load up ScrnCalc(1) with current screen data
   For x = 1 To 52
     vaSpreadDedData.Col = 2
     vaSpreadDedData.Row = Nextx
     If ScrnCalc(1).Ded(x) <> Val(ReplaceString$(vaSpreadDedData.Text, "$", "")) Then Change = True '8/26/04
     ScrnCalc(1).Ded(x) = Val(ReplaceString$(vaSpreadDedData.Text, "$", ""))
     
     vaSpreadDedData.Col = 4
     vaSpreadDedData.Row = Nextx
     If ScrnCalc(1).Ded(x + 1) <> Val(ReplaceString$(vaSpreadDedData.Text, "$", "")) Then Change = True '8/26/04
     ScrnCalc(1).Ded(x + 1) = Val(ReplaceString$(vaSpreadDedData.Text, "$", ""))
     
     vaSpreadDedData.Col = 6
     vaSpreadDedData.Row = Nextx
     If x + 2 > 50 Then Exit For
     If ScrnCalc(1).Ded(x + 2) <> Val((ReplaceString$(vaSpreadDedData.Text, "$", ""))) Then Change = True '8/26/04
     ScrnCalc(1).Ded(x + 2) = Val(ReplaceString$(vaSpreadDedData.Text, "$", ""))
     Nextx = Nextx + 1
     x = x + 2
   Next x
   
   If ScrnCalc(1).REGEARN <> CDbl(fpcurrEarnings(1).Text) Then Change = True '8/26/04
   ScrnCalc(1).REGEARN = fpcurrEarnings(1).Text
   
   If ScrnCalc(1).OTEARN <> CDbl(fpcurrEarnings(2).Text) Then Change = True '8/26/04
   ScrnCalc(1).OTEARN = fpcurrEarnings(2).Text
   
   If ScrnCalc(1).ALTEARN1 <> CDbl(fpcurrEarnings(3).Text) Then Change = True '8/26/04
   ScrnCalc(1).ALTEARN1 = fpcurrEarnings(3).Text
   
   If ScrnCalc(1).ALTEARN2 <> CDbl(fpcurrEarnings(4).Text) Then Change = True '8/26/04
   ScrnCalc(1).ALTEARN2 = fpcurrEarnings(4).Text
   
   If ScrnCalc(1).ALTEARN3 <> CDbl(fpcurrEarnings(5).Text) Then Change = True '8/26/04
   ScrnCalc(1).ALTEARN3 = fpcurrEarnings(5).Text
   
   If ScrnCalc(1).SOCTAX <> CDbl(fpcurrTaxSS.Text) Then Change = True '8/26/04
   ScrnCalc(1).SOCTAX = fpcurrTaxSS.Text
   
   If ScrnCalc(1).MEDTAX <> CDbl(fpcurrTaxMed.Text) Then Change = True '8/26/04
   ScrnCalc(1).MEDTAX = fpcurrTaxMed.Text
   
   If ScrnCalc(1).FEDTAX <> CDbl(fpcurrTaxFed.Text) Then Change = True '8/26/04
   ScrnCalc(1).FEDTAX = fpcurrTaxFed.Text
   
   If ScrnCalc(1).STATAX <> CDbl(fpcurrTaxState.Text) Then Change = True '8/26/04
   ScrnCalc(1).STATAX = fpcurrTaxState.Text
   
   If ScrnCalc(1).RETIRE <> CDbl(fpcurrTaxRetire.Text) Then Change = True '8/26/04
   ScrnCalc(1).RETIRE = fpcurrTaxRetire.Text
   
   If ScrnCalc(1).GrossPay <> CDbl(fpcurrGrossPay.Text) Then Change = True '8/26/04
   ScrnCalc(1).GrossPay = fpcurrGrossPay.Text
   
   If ScrnCalc(1).TOTDED <> CDbl(fpcurrTotDed.Text) Then Change = True '8/26/04
   ScrnCalc(1).TOTDED = fpcurrTotDed.Text
   
   If ScrnCalc(1).EIC <> CDbl(fpcurrAdvEIC.Text) Then Change = True '8/26/04
   ScrnCalc(1).EIC = fpcurrAdvEIC.Text
   
   If ScrnCalc(1).NetPay <> CDbl(fpcurrNetPay.Text) Then Change = True '8/26/04
   ScrnCalc(1).NetPay = fpcurrNetPay.Text

  If Change = True Then '8/26/04
    If Exist("TEMPIF.DAT") Then '8/26/04
      KillFile "TEMPIF.DAT" '8/26/04
    End If '8/26/04
  End If '8/26/04

End Sub
Private Sub ScrnCalc2Screen(ScrnCalc() As ScrnCalcType)
   Dim x As Integer
   Dim Nextx As Integer
   Dim Image$
   
   Image = "$###,##0.00"
   Nextx = 1
   For x = 1 To 52
     If x > DedCnt Then Exit For
     vaSpreadDedData.Col = 2
     vaSpreadDedData.Row = Nextx
     vaSpreadDedData.Text = Using(Image, ScrnCalc(1).Ded(x))
     If x + 1 > DedCnt Then Exit For
     vaSpreadDedData.Col = 4
     vaSpreadDedData.Row = Nextx
     vaSpreadDedData.Text = Using(Image, ScrnCalc(1).Ded(x + 1))
     If x + 2 > DedCnt Then Exit For
     vaSpreadDedData.Col = 6
     vaSpreadDedData.Row = Nextx
     If x + 2 > 50 Then Exit For
     vaSpreadDedData.Text = Using(Image, ScrnCalc(1).Ded(x + 2))
     Nextx = Nextx + 1
     x = x + 2
   Next x
   
   fpcurrEarnings(1).Text = ScrnCalc(1).REGEARN
   fpcurrEarnings(2).Text = ScrnCalc(1).OTEARN
   fpcurrEarnings(3).Text = ScrnCalc(1).ALTEARN1
   fpcurrEarnings(4).Text = ScrnCalc(1).ALTEARN2
   fpcurrEarnings(5).Text = ScrnCalc(1).ALTEARN3
   fpcurrTaxSS.Text = ScrnCalc(1).SOCTAX
   fpcurrTaxMed.Text = ScrnCalc(1).MEDTAX
   fpcurrTaxFed.Text = ScrnCalc(1).FEDTAX
   fpcurrTaxState.Text = ScrnCalc(1).STATAX
   fpcurrTaxRetire.Text = ScrnCalc(1).RETIRE
   fpcurrGrossPay.Text = ScrnCalc(1).GrossPay
   fpcurrTotDed.Text = ScrnCalc(1).TOTDED
   fpcurrAdvEIC.Text = ScrnCalc(1).EIC
   fpcurrNetPay.Text = ScrnCalc(1).NetPay
End Sub

Private Sub fpcurrAdvEIC_LostFocus()
  'this code should never be used because
  'this textbox is read only...left it in
  'in case the read only status ever changes
  Dim Amt As Double
  Amt = Val(fpcurrAdvEIC.Text)
  
  If fpcurrAdvEIC.Text = "" Then
    fpcurrAdvEIC = ".00"
  Else
    fpcurrAdvEIC.Text = Amt
  End If
End Sub

Private Sub fpcurrEarnings_LostFocus(Index As Integer)
  Dim x As Integer
  Dim Amt As Double
  Dim Nextx As Integer
  Dim Image$
  Dim TransRec(1) As TransRecType
  Dim THandle As Integer
  
  Image = "$###,##0.00"
  If OldRegWages# <> fpcurrEarnings(1).Text Then
    OpenTransWorkFile THandle
    Get THandle, RecNum, TransRec(1)
    ReCalcPay TransRec(), ScrnCalc()
    Put THandle, RecNum, TransRec(1)
    Close THandle
    Nextx = 1
    For x = 1 To 52 'deduction amounts could change
    'if they are a percent of Earnings
      If x > DedCnt Then Exit For '7/25
      vaSpreadDedData.Col = 2
      vaSpreadDedData.Row = Nextx
      Amt = Val(ReplaceString$(vaSpreadDedData.Text, "$", "")) '7/25
      If vaSpreadDedData.Text = "" Then
        vaSpreadDedData.Text = "0.00"
      Else
        vaSpreadDedData.Text = Using(Image, Amt)
      End If
      OldFldVal(x) = Val(ReplaceString$(vaSpreadDedData.Text, "$", "")) 'RemoveDollarMark
      
      If x + 1 > DedCnt Then Exit For '7/25
      
      vaSpreadDedData.Col = 4
      vaSpreadDedData.Row = Nextx
      Amt = Val(ReplaceString$(vaSpreadDedData.Text, "$", "")) '7/25
      If vaSpreadDedData.Text = "" Then
        vaSpreadDedData.Text = "0.00"
      Else
        vaSpreadDedData.Text = Using(Image, Amt)
      End If
      OldFldVal(x + 1) = Val(ReplaceString$(vaSpreadDedData.Text, "$", "")) 'RemoveDollarMark
      
      If x + 2 > DedCnt Then Exit For '7/25
      
      vaSpreadDedData.Col = 6
      vaSpreadDedData.Row = Nextx
      If x + 2 > 50 Then Exit For
      Amt = Val(ReplaceString$(vaSpreadDedData.Text, "$", "")) '7/25
      If vaSpreadDedData.Text = "" Then
        vaSpreadDedData.Text = "0.00"
      Else
        vaSpreadDedData.Text = Using(Image, Amt)
      End If
      OldFldVal(x + 2) = Val(ReplaceString$(vaSpreadDedData.Text, "$", "")) 'RemoveDollarMark
      x = x + 2
      Nextx = Nextx + 1
    Next x
      
    OldRetAmt# = fpcurrTaxRetire.Text
    OldRegWages# = fpcurrEarnings(1).Text 'save it to have
    'a constant with the current value
    OldFedTax# = fpcurrTaxFed.Text 'constant value
    OldStateTax# = fpcurrTaxState.Text 'added 8/22...constant value
  End If
End Sub

Private Sub fpcurrGrossPay_LostFocus()
  'this code should never be used because
  'this textbox is read only...left it in
  'in case the read only status ever changes
  Dim Amt As Double
  Amt = fpcurrGrossPay.Text
  If fpcurrGrossPay.Text = "" Then
    fpcurrGrossPay.Text = ".00"
  Else
    fpcurrGrossPay.Text = Amt
  End If
  
  Call CalcFields
End Sub

Private Sub fpcurrNetPay_LostFocus()
  'this code should never be used because
  'this textbox is read only...left it in
  'in case the read only status ever changes
  Dim Amt As Double
  Amt = fpcurrNetPay.Text
  If fpcurrNetPay.Text = "" Then
    fpcurrNetPay.Text = ".00"
  Else
    fpcurrNetPay.Text = Amt
  End If
  
  Call CalcFields
  If fpcurrNetPay < 0 Then
    frmWarnNegNetPay.Show vbModal, Me
  End If
End Sub

Private Sub fpcurrTaxFed_LostFocus()
  Dim Amt As Double
  Amt = fpcurrTaxFed.Text

  If fpcurrTaxFed.Text = "" Then
    fpcurrTaxFed.Text = ".00"
  Else
    fpcurrTaxFed.Text = Amt
  End If
  
  If OldFedTax# <> fpcurrTaxFed.Text Then '8/22 OldFedTax
  'always refigures based on Gross Pay when this screen is
  'loaded
    OpenTransWorkFile THandle '8/22
    Get THandle, RecNum, TransRec(1) '8/22
    Call CalcFields 'refigures net pay based on screen data
    '...it does not go thru the tax calculations in CalcPay
    '...we only need for Net Pay to be changed by the
    'difference in old fed tax to new fed tax
    Screen2ScrnCalc ScrnCalc() '8/22
    'copy scrn calcs fields to the transaction fields
    ParseScrnCalc2Trans TransRec(), ScrnCalc() '8/22 save
    'new Net Pay and Fed Tax values...all else stays the same
    Put THandle, RecNum, TransRec(1) '8/22
    Close THandle '8/22
    OldRetAmt# = fpcurrTaxRetire.Text
    OldRegWages# = fpcurrEarnings(1).Text
    OldFedTax# = fpcurrTaxFed.Text
  End If '8/22
End Sub

Private Sub fpcurrTaxMed_LostFocus()
  'this code should never be used because
  'this textbox is read only...left it in
  'in case the read only status ever changes
  Dim Amt As Double
  Amt = fpcurrTaxMed.Text

  If fpcurrTaxMed.Text = "" Then
    fpcurrTaxMed.Text = ".00"
  Else
    fpcurrTaxMed.Text = Amt
  End If
  Call CalcFields
End Sub

Private Sub fpcurrTaxRetire_LostFocus()
  'this code should never be used because
  'this textbox is read only...left it in
  'in case the read only status ever changes
  Dim Amt As Double
  Amt = fpcurrTaxRetire.Text

  If fpcurrTaxRetire.Text = "" Then
    fpcurrTaxRetire.Text = ".00"
  Else
    fpcurrTaxRetire.Text = Amt
  End If
  
  Call CalcFields
End Sub

Private Sub fpcurrTaxSS_LostFocus()
  'this code should never be used because
  'this textbox is read only...left it in
  'in case the read only status ever changes
  Dim Amt As Double
  Amt = fpcurrTaxSS.Text

  If fpcurrTaxSS.Text = "" Then
    fpcurrTaxSS.Text = "0"
  Else
    fpcurrTaxSS.Text = Amt
  End If
  
  Call CalcFields
End Sub

Private Sub fpcurrTaxState_LostFocus()
  Dim Amt As Double
  Amt = fpcurrTaxState.Text

  If fpcurrTaxState.Text = "" Then
    fpcurrTaxState.Text = ".00"
  Else
    fpcurrTaxState.Text = Amt
  End If
  
  If OldStateTax# <> fpcurrTaxState.Text Then '8/22 OldStateTax
  'always refigures based on Gross Pay when this screen is
  'loaded
  OpenTransWorkFile THandle '8/22
    Get THandle, RecNum, TransRec(1) '8/22
    Call CalcFields 'refigures net pay based on screen data
    '...it does not go thru the tax calculations in CalcPay
    '...we only need for Net Pay to be changed by the
    'difference in old fed tax to new fed tax
    Screen2ScrnCalc ScrnCalc() '8/22
    'copy scrn calcs fields to the transaction fields
    ParseScrnCalc2Trans TransRec(), ScrnCalc() '8/22 save
    'new Net Pay and Fed Tax values...all else stays the same
    Put THandle, RecNum, TransRec(1) '8/22
    Close THandle '8/22
    OldRetAmt# = fpcurrTaxRetire.Text
    OldRegWages# = fpcurrEarnings(1).Text
    OldFedTax# = fpcurrTaxFed.Text
    OldStateTax# = fpcurrTaxState.Text
  End If '8/22
End Sub

Private Sub fpcurrTotDed_LostFocus()
  'this code should never be used because
  'this textbox is read only...left it in
  'in case the read only status ever changes
  Dim Amt As Double
  Amt = fpcurrTotDed.Text

  If fpcurrTotDed.Text = "" Then
    fpcurrTotDed.Text = ".00"
  Else
    fpcurrTotDed.Text = Amt
  End If
    
  Call CalcFields
End Sub

Private Sub CalcFields()

  Dim x As Integer
  Dim GrossPay As Double
  Dim TotDeds As Double
  Dim NetPayTotal As Double
  Dim Nextx As Integer
  
  For x = 1 To 5 'add up all earnings
    GrossPay = CDbl(GrossPay) + CDbl(fpcurrEarnings(x))
  Next x
  fpcurrGrossPay.Text = GrossPay
  TotDeds = CDbl(fpcurrTaxSS) + CDbl(fpcurrTaxMed) + CDbl(fpcurrTaxFed) + CDbl(fpcurrTaxState) + CDbl(fpcurrTaxRetire)
'  fpcurrTotDed.Text = TotDeds '8/22 also below
  Nextx = 1
  For x = 1 To 52 'add up all deduction amounts
    vaSpreadDedData.Col = 2
    vaSpreadDedData.Row = Nextx
    If Len(QPTrim$(vaSpreadDedData.Text)) = 0 Then Exit For
    TotDeds = CDbl(TotDeds) + CDbl(vaSpreadDedData.Text) 'RemoveDollarMark
    vaSpreadDedData.Col = 4
    vaSpreadDedData.Row = Nextx
    If Len(QPTrim$(vaSpreadDedData.Text)) = 0 Then Exit For
    TotDeds = CDbl(TotDeds) + CDbl(vaSpreadDedData.Text) 'RemoveDollarMark
    If x + 2 > 50 Then Exit For
    vaSpreadDedData.Col = 6
    vaSpreadDedData.Row = Nextx
    If Len(QPTrim$(vaSpreadDedData.Text)) = 0 Then Exit For
    TotDeds = CDbl(TotDeds) + CDbl(vaSpreadDedData.Text) 'RemoveDollarMark
    x = x + 2
    Nextx = Nextx + 1
  Next x
  
  fpcurrTotDed.Text = TotDeds
  NetPayTotal = OldRound(GrossPay) - OldRound(TotDeds) + fpcurrAdvEIC.Text
  fpcurrNetPay.Text = NetPayTotal
End Sub

Private Sub DeleteThisEmp()
  Dim THandle As Integer
  Dim TransRec As TransRecType
  
  OpenTransWorkFile THandle
  Get THandle, RecNum, TransRec
  TransRec.TActive = 0 'deactivate this employee
  Put THandle, RecNum, TransRec
  
  Close THandle
  KillFile "TEMPIF.DAT" '2/3/04
  
End Sub

Private Sub mnuCont_Click()
  Call cmdCont_Click
End Sub

Private Sub mnuPrnScn_Click()
  PrintForm
  MainLog ("Payroll Calculations (PRDCalcScr3) screen printed.")
End Sub

Private Sub vaSpreadDedData_Click(ByVal Col As Long, ByVal Row As Long)
  '8/7 added "Replace Existing Text" option in spreadsheet, located
  'in the spreadsheet edit under General
  vaSpreadDedData.EditMode = True '8/7

End Sub

Private Sub vaSpreadDedData_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
  Dim x As Integer, Nextx As Integer
  Dim changeFlag As Boolean
  Dim Amt As Double
  Dim TransRec(1) As TransRecType
  Dim Image$
  OpenTransWorkFile TRHandle
  Get TRHandle, RecNum, TransRec(1)
  
  Image = "$###,##0.00"
  changeFlag = False
  Nextx = 1
  For x = 1 To 52 '52 is used because this spreadsheet is dynamic
  'if all 50 deductions are used then the last cell to be filled
  'will be in the second column...there are 3 columns, so 48/3
  '= 16 rows...that means 50 entries will only use 2 cells of the last
  'row
    If x > DedCnt Then Exit For
    vaSpreadDedData.Col = 2
    vaSpreadDedData.Row = Nextx
    Amt = Val(ReplaceString$(vaSpreadDedData.Text, "$", "")) 'RemoveDollarMark
    If vaSpreadDedData.Text = "" Then
      vaSpreadDedData.Text = ".00"
    Else
      vaSpreadDedData.Text = Amt
      If Amt > 0 Then
        vaSpreadDedData.FontBold = True 'cells with non-zero values
        'are displayed as bold for the user to find them easier
        vaSpreadDedData.Col = 1
        vaSpreadDedData.Row = Nextx
        vaSpreadDedData.FontBold = True
      ElseIf Amt = 0 Then
        vaSpreadDedData.FontBold = False
        vaSpreadDedData.Col = 1
        vaSpreadDedData.Row = Nextx
        vaSpreadDedData.FontBold = False
      End If
    End If
    vaSpreadDedData.Col = 2
    If OldFldVal(x) <> Val(ReplaceString$(vaSpreadDedData.Text, "$", "")) Then 'RemoveDollarMark
      changeFlag = True
    End If
    If x + 1 > DedCnt Then Exit For 'drop out of for loop
    'if we've passed the number of deductions saved
    vaSpreadDedData.Col = 4
    vaSpreadDedData.Row = Nextx
    Amt = Val(ReplaceString$(vaSpreadDedData.Text, "$", "")) 'RemoveDollarMark
    If vaSpreadDedData.Text = "" Then
      vaSpreadDedData.Text = ".00"
    Else
      vaSpreadDedData.Text = Amt
      If Amt > 0 Then
        vaSpreadDedData.FontBold = True
        vaSpreadDedData.Col = 3
        vaSpreadDedData.Row = Nextx
        vaSpreadDedData.FontBold = True
      ElseIf Amt = 0 Then
        vaSpreadDedData.FontBold = False
        vaSpreadDedData.Col = 3
        vaSpreadDedData.Row = Nextx
        vaSpreadDedData.FontBold = False
      End If
    End If
    vaSpreadDedData.Col = 4
    If OldFldVal(x + 1) <> Val(ReplaceString$(vaSpreadDedData.Text, "$", "")) Then 'RemoveDollarMark
      changeFlag = True
    End If
    If x + 2 > 50 Then Exit For
    If x + 2 > DedCnt Then Exit For
    vaSpreadDedData.Col = 6
    vaSpreadDedData.Row = Nextx
    Amt = Val(ReplaceString$(vaSpreadDedData.Text, "$", "")) 'RemoveDollarMark
    If vaSpreadDedData.Text = "" Then
      vaSpreadDedData.Text = ".00"
    Else
      vaSpreadDedData.Text = Amt
      If Amt > 0 Then
        vaSpreadDedData.FontBold = True
        vaSpreadDedData.Col = 5
        vaSpreadDedData.Row = Nextx
        vaSpreadDedData.FontBold = True
      ElseIf Amt = 0 Then
        vaSpreadDedData.FontBold = False
        vaSpreadDedData.Col = 5
        vaSpreadDedData.Row = Nextx
        vaSpreadDedData.FontBold = False
      End If
    End If
    vaSpreadDedData.Col = 6
    If OldFldVal(x + 2) <> Val(ReplaceString$(vaSpreadDedData.Text, "$", "")) Then 'RemoveDollarMark
      changeFlag = True
    End If
    Nextx = Nextx + 1
    x = x + 2
  Next x
  
  Nextx = 1
  If changeFlag = True Then
    ReCalcPay TransRec(), ScrnCalc()
    For x = 1 To 52
      vaSpreadDedData.Col = 2
      vaSpreadDedData.Row = Nextx
      OldFldVal(x) = Val(ReplaceString$(vaSpreadDedData.Text, "$", "")) 'RemoveDollarMark
      vaSpreadDedData.Col = 4
      vaSpreadDedData.Row = Nextx
      OldFldVal(x + 1) = Val(ReplaceString$(vaSpreadDedData.Text, "$", "")) 'RemoveDollarMark
      If x + 2 > 50 Then Exit For
      vaSpreadDedData.Col = 6
      vaSpreadDedData.Row = Nextx
      OldFldVal(x + 2) = Val(ReplaceString$(vaSpreadDedData.Text, "$", "")) 'RemoveDollarMark
    Nextx = Nextx + 1
    x = x + 2
    Next x
    OldRetAmt# = fpcurrTaxRetire.Text
    OldRegWages# = fpcurrEarnings(1).Text
    OldFedTax# = fpcurrTaxFed.Text
    OldStateTax = fpcurrTaxState.Text 'added 8/22
  End If
  Put TRHandle, RecNum, TransRec(1) 'added 8/22 to
  'fix a bug where pertinent data figured in this procedure
  'was not being saved
  Close TRHandle
End Sub

Private Function FixSpread()
  Dim COne As Integer
  Dim CTwo As Integer
  Dim CThree As Integer
  Dim CFour As Integer
  Dim CFive As Integer
  Dim CSix As Integer
  '-1 means all rows or all columns....0 means headers
    Select Case ScreenW
      Case 1280
      If Screen.TwipsPerPixelX <> 12 Then
        COne = 13
        coladj = 4
        vaSpreadDedData.FontSize = 18
        vaSpreadDedData.RowHeight(-1) = 22
        vaSpreadDedData.RowHeight(0) = 22
      Else
        COne = 6
        coladj = 3.1
        vaSpreadDedData.RowHeight(-1) = 19
        vaSpreadDedData.RowHeight(0) = 19
      End If
      Case 1152
      If Screen.TwipsPerPixelX <> 12 Then
        COne = 8.5
        coladj = 4.5
        vaSpreadDedData.FontSize = 14
        vaSpreadDedData.RowHeight(0) = 18.5
        vaSpreadDedData.RowHeight(-1) = 18.5
      Else
        COne = 3
        coladj = 2.3
        vaSpreadDedData.RowHeight(0) = 15
        vaSpreadDedData.RowHeight(-1) = 15
      End If
      Case 1024
      If Screen.TwipsPerPixelX <> 12 Then
        COne = 5
        coladj = 4#
        vaSpreadDedData.RowHeight(0) = 17.5
        vaSpreadDedData.FontBold = True
        vaSpreadDedData.RowHeight(-1) = 17.5
      Else
        COne = 0.5
        coladj = 1.6
      End If
      Case 800
        COne = -0.6
        coladj = 1.85
        vaSpreadDedData.Font.Size = 10
        vaSpreadDedData.RowHeight(-1) = 12.2
      Case Else
       
    End Select
    vaSpreadDedData.ColWidth(1) = vaSpreadDedData.ColWidth(1) + COne
    vaSpreadDedData.ColWidth(2) = vaSpreadDedData.ColWidth(2) + coladj
    vaSpreadDedData.ColWidth(3) = vaSpreadDedData.ColWidth(3) + COne
    vaSpreadDedData.ColWidth(4) = vaSpreadDedData.ColWidth(4) + coladj
    vaSpreadDedData.ColWidth(5) = vaSpreadDedData.ColWidth(5) + COne
    vaSpreadDedData.ColWidth(6) = vaSpreadDedData.ColWidth(6) + coladj

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      MainLog ("Payroll.exe terminated via menu bar on frmPRDCalcScr3.")
      Call Terminate
      End
    End If
  End If
End Sub


