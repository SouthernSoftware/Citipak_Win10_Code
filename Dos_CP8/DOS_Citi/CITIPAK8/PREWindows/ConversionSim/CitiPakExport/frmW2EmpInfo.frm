VERSION 5.00
Object = "{FD2FB1F1-D4FC-11CE-A335-A8D5ECAE5B02}#2.0#0"; "btn32a20.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form frmW2EmpInfo 
   BackColor       =   &H008F8265&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee W2 Information"
   ClientHeight    =   8865
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   11655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmW2EmpInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8640
   ScaleMode       =   0  'User
   ScaleWidth      =   11652
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin EditLib.fpCurrency fpcurrSeeInstr12c 
      Height          =   345
      Left            =   4320
      TabIndex        =   20
      ToolTipText     =   "Box 12 amount."
      Top             =   4410
      Width           =   1830
      _Version        =   196608
      _ExtentX        =   3228
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
   Begin EditLib.fpText fptxtOtherb 
      Height          =   345
      Left            =   9780
      TabIndex        =   27
      ToolTipText     =   "Box 14 other text."
      Top             =   4800
      Width           =   870
      _Version        =   196608
      _ExtentX        =   1545
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
      MaxLength       =   4
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
   Begin EditLib.fpText fptxtOthera 
      Height          =   345
      Left            =   9780
      TabIndex        =   25
      ToolTipText     =   "Box 14 other text."
      Top             =   4365
      Width           =   870
      _Version        =   196608
      _ExtentX        =   1545
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
      MaxLength       =   4
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
   Begin EditLib.fpText fptxtBox12b 
      Height          =   345
      Left            =   3030
      TabIndex        =   19
      ToolTipText     =   "Box 12 text."
      Top             =   4845
      Width           =   870
      _Version        =   196608
      _ExtentX        =   1545
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
      MaxLength       =   4
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
   Begin EditLib.fpCurrency fpcurrOtherb 
      Height          =   345
      Left            =   7815
      TabIndex        =   26
      ToolTipText     =   "Box 14 other amount."
      Top             =   4800
      Width           =   1830
      _Version        =   196608
      _ExtentX        =   3238
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
   Begin EditLib.fpCurrency fpcurrSeeInstr12b 
      Height          =   345
      Left            =   1095
      TabIndex        =   18
      ToolTipText     =   "Box 12 amount."
      Top             =   4845
      Width           =   1830
      _Version        =   196608
      _ExtentX        =   3238
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
   Begin EditLib.fpText fptxtLocality19 
      Height          =   348
      Left            =   1212
      TabIndex        =   34
      ToolTipText     =   "Locality name."
      Top             =   7212
      Width           =   2172
      _Version        =   196608
      _ExtentX        =   3831
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
      MaxLength       =   36
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
   Begin EditLib.fpText fptxtState16 
      Height          =   348
      Left            =   1212
      TabIndex        =   31
      ToolTipText     =   "State"
      Top             =   6540
      Width           =   876
      _Version        =   196608
      _ExtentX        =   1545
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
      CharValidationText=   "a b c d e f g h i j k l m n o p q r s t u v w x y z A B C D E F G H I J K L M N O P Q R S T U V W X Y Z"
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
   Begin EditLib.fpCurrency fpcurrLocalIncTax21 
      Height          =   348
      Left            =   8796
      TabIndex        =   36
      ToolTipText     =   "Local income tax."
      Top             =   7212
      Width           =   1836
      _Version        =   196608
      _ExtentX        =   3238
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
   Begin EditLib.fpCurrency fpcurrLocalWages20 
      Height          =   348
      Left            =   5184
      TabIndex        =   35
      ToolTipText     =   "Local wages, tips, etc."
      Top             =   7212
      Width           =   1836
      _Version        =   196608
      _ExtentX        =   3238
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
   Begin EditLib.fpCurrency fpcurrStateIncTax18 
      Height          =   348
      Left            =   8796
      TabIndex        =   33
      ToolTipText     =   "State income tax"
      Top             =   6540
      Width           =   1836
      _Version        =   196608
      _ExtentX        =   3238
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
   Begin EditLib.fpCurrency fpcurrOthera 
      Height          =   345
      Left            =   7800
      TabIndex        =   24
      ToolTipText     =   "Box 14 other amount."
      Top             =   4380
      Width           =   1830
      _Version        =   196608
      _ExtentX        =   3238
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
   Begin EditLib.fpCurrency fpcurrSeeInstr12a 
      Height          =   345
      Left            =   1080
      TabIndex        =   16
      ToolTipText     =   "Box 12 amount."
      Top             =   4425
      Width           =   1830
      _Version        =   196608
      _ExtentX        =   3238
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
   Begin EditLib.fpCurrency fpcurrDepCareBene10 
      Height          =   348
      Left            =   3744
      TabIndex        =   14
      ToolTipText     =   "Dependent care benefits"
      Top             =   3564
      Width           =   1644
      _Version        =   196608
      _ExtentX        =   2900
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
   Begin EditLib.fpCurrency fpcurrAdvanceEIC9 
      Height          =   348
      Left            =   1212
      TabIndex        =   13
      ToolTipText     =   "Advance Earned Income Credit payments"
      Top             =   3564
      Width           =   1644
      _Version        =   196608
      _ExtentX        =   2900
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
   Begin EditLib.fpCurrency fpcurrAlloTips8 
      Height          =   348
      Left            =   8988
      TabIndex        =   12
      ToolTipText     =   "Allocated tips"
      Top             =   2832
      Width           =   1644
      _Version        =   196608
      _ExtentX        =   2900
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
   Begin EditLib.fpCurrency fpcurrSocSecTips7 
      Height          =   348
      Left            =   6336
      TabIndex        =   11
      ToolTipText     =   "Social Security tips"
      Top             =   2844
      Width           =   1644
      _Version        =   196608
      _ExtentX        =   2900
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
   Begin EditLib.fpCurrency fpcurrMedTaxWH6 
      Height          =   348
      Left            =   3744
      TabIndex        =   10
      ToolTipText     =   "Medicare taxes withheld"
      Top             =   2844
      Width           =   1644
      _Version        =   196608
      _ExtentX        =   2900
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
   Begin EditLib.fpCurrency fpcurrMedWages5 
      Height          =   348
      Left            =   1212
      TabIndex        =   9
      ToolTipText     =   "Medicare wages"
      Top             =   2844
      Width           =   1644
      _Version        =   196608
      _ExtentX        =   2900
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
   Begin EditLib.fpCurrency fpcurrSocSecWH4 
      Height          =   348
      Left            =   8988
      TabIndex        =   8
      ToolTipText     =   "Social Security taxes withheld"
      Top             =   2172
      Width           =   1644
      _Version        =   196608
      _ExtentX        =   2900
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
   Begin EditLib.fpCurrency fpcurrFedIncTax2 
      Height          =   348
      Left            =   3744
      TabIndex        =   6
      ToolTipText     =   "Federal income tax withheld"
      Top             =   2172
      Width           =   1644
      _Version        =   196608
      _ExtentX        =   2900
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
   Begin EditLib.fpCurrency fpcurrSocSecWages3 
      Height          =   348
      Left            =   6336
      TabIndex        =   7
      ToolTipText     =   "Social Security wages"
      Top             =   2172
      Width           =   1644
      _Version        =   196608
      _ExtentX        =   2900
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
   Begin VB.CheckBox chkThirdParty3 
      BackColor       =   &H008F8265&
      Height          =   288
      Left            =   10428
      TabIndex        =   30
      ToolTipText     =   "Deferred Compensation."
      Top             =   5664
      Width           =   204
   End
   Begin VB.CheckBox chkRetirePlan 
      BackColor       =   &H008F8265&
      Height          =   288
      Left            =   6636
      TabIndex        =   29
      ToolTipText     =   "Pension plan."
      Top             =   5664
      Width           =   204
   End
   Begin VB.CheckBox chkStatEmp 
      BackColor       =   &H008F8265&
      Height          =   288
      Left            =   2700
      TabIndex        =   28
      ToolTipText     =   "Statutory employee."
      Top             =   5664
      Width           =   204
   End
   Begin EditLib.fpCurrency fpcurrWagesTips1 
      Height          =   348
      Left            =   1212
      TabIndex        =   5
      ToolTipText     =   "Wages, tips, other compensation"
      Top             =   2172
      Width           =   1644
      _Version        =   196608
      _ExtentX        =   2900
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
   Begin EditLib.fpCurrency fpcurrStateWages17 
      Height          =   348
      Left            =   5196
      TabIndex        =   32
      ToolTipText     =   "State wages, tips, etc."
      Top             =   6528
      Width           =   1836
      _Version        =   196608
      _ExtentX        =   3238
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
   Begin EditLib.fpCurrency fpcurrNonQPlans11 
      Height          =   348
      Left            =   6336
      TabIndex        =   15
      ToolTipText     =   "Nonqualified plans"
      Top             =   3564
      Width           =   1644
      _Version        =   196608
      _ExtentX        =   2900
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
   Begin EditLib.fpText fptxtBox12a 
      Height          =   345
      Left            =   3030
      TabIndex        =   17
      ToolTipText     =   "Box 12 text."
      Top             =   4410
      Width           =   870
      _Version        =   196608
      _ExtentX        =   1545
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
      MaxLength       =   4
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
   Begin EditLib.fpText fptxtFirstName 
      Height          =   348
      Left            =   3660
      TabIndex        =   0
      ToolTipText     =   "Enter the employee's first name as it appears on their social security card."
      Top             =   720
      Width           =   2796
      _Version        =   196608
      _ExtentX        =   4932
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
      MaxLength       =   15
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
   Begin EditLib.fpText fptxtLastName 
      Height          =   348
      Left            =   3660
      TabIndex        =   2
      ToolTipText     =   "Enter the employee's last name as it appears on their social security card."
      Top             =   1152
      Width           =   3660
      _Version        =   196608
      _ExtentX        =   6456
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
   Begin EditLib.fpText fptxtSuffix 
      Height          =   348
      Left            =   9696
      TabIndex        =   3
      ToolTipText     =   "Enter the employee's suffix (Jr. or II, etc.) as it appears on their social security card."
      Top             =   1152
      Width           =   924
      _Version        =   196608
      _ExtentX        =   1630
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
      MaxLength       =   4
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
   Begin EditLib.fpText fptxtMiddleName 
      Height          =   348
      Left            =   7836
      TabIndex        =   1
      ToolTipText     =   "Enter the employee's middle name as it appears on their social security card."
      Top             =   720
      Width           =   2796
      _Version        =   196608
      _ExtentX        =   4932
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
      MaxLength       =   15
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
   Begin fpBtnAtlLibCtl.fpBtn cmdExit 
      Height          =   495
      Left            =   6828
      TabIndex        =   64
      ToolTipText     =   "Press to exit to the main W2 menu."
      Top             =   8016
      Width           =   1410
      _Version        =   131072
      _ExtentX        =   2487
      _ExtentY        =   873
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
      DrawFocusRect   =   4
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
      ButtonDesigner  =   "frmW2EmpInfo.frx":08CA
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdSave 
      Height          =   495
      Left            =   8508
      TabIndex        =   65
      ToolTipText     =   "Press to commit this data to memory."
      Top             =   8016
      Width           =   1410
      _Version        =   131072
      _ExtentX        =   2487
      _ExtentY        =   873
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
      DrawFocusRect   =   4
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
      ButtonDesigner  =   "frmW2EmpInfo.frx":0AA6
   End
   Begin fpBtnAtlLibCtl.fpBtn cmdPrintForm 
      Height          =   495
      Left            =   4464
      TabIndex        =   66
      Top             =   8016
      Width           =   2070
      _Version        =   131072
      _ExtentX        =   3651
      _ExtentY        =   873
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
      DrawFocusRect   =   4
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
      ButtonDesigner  =   "frmW2EmpInfo.frx":0C82
   End
   Begin EditLib.fpText fptxtBox12c 
      Height          =   345
      Left            =   6270
      TabIndex        =   21
      ToolTipText     =   "Box 12 text."
      Top             =   4410
      Width           =   870
      _Version        =   196608
      _ExtentX        =   1535
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
      MaxLength       =   4
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
   Begin EditLib.fpCurrency fpcurrSeeInstr12d 
      Height          =   345
      Left            =   4320
      TabIndex        =   22
      ToolTipText     =   "Box 12 amount."
      Top             =   4845
      Width           =   1830
      _Version        =   196608
      _ExtentX        =   3228
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
   Begin EditLib.fpText fptxtBox12d 
      Height          =   345
      Left            =   6270
      TabIndex        =   23
      ToolTipText     =   "Box 12 text."
      Top             =   4845
      Width           =   870
      _Version        =   196608
      _ExtentX        =   1535
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
      MaxLength       =   4
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
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "b."
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
      Height          =   300
      Left            =   7560
      TabIndex        =   72
      Top             =   4920
      Width           =   210
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "a."
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
      Height          =   300
      Left            =   7560
      TabIndex        =   71
      Top             =   4515
      Width           =   210
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "d."
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
      Height          =   300
      Left            =   4080
      TabIndex        =   70
      Top             =   4920
      Width           =   210
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "c."
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
      Height          =   300
      Left            =   4080
      TabIndex        =   69
      Top             =   4515
      Width           =   210
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "b."
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
      Height          =   300
      Left            =   840
      TabIndex        =   68
      Top             =   4920
      Width           =   210
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "a."
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
      Height          =   300
      Left            =   840
      TabIndex        =   67
      Top             =   4515
      Width           =   210
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   567.854
      X2              =   11022.16
      Y1              =   5939.33
      Y2              =   5939.33
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   567.854
      X2              =   11020.16
      Y1              =   1636.386
      Y2              =   1636.386
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Names For Electronic W2 Submission Only"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   828
      Left            =   828
      TabIndex        =   63
      Top             =   720
      Width           =   1644
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Suffix"
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
      Height          =   300
      Left            =   9084
      TabIndex        =   62
      Top             =   1200
      Width           =   636
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
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
      Height          =   300
      Left            =   2556
      TabIndex        =   61
      Top             =   1200
      Width           =   1020
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Middle Name"
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
      Height          =   300
      Left            =   6588
      TabIndex        =   60
      Top             =   768
      Width           =   1164
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
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
      Height          =   300
      Left            =   2556
      TabIndex        =   59
      Top             =   768
      Width           =   1020
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "8. Allocated tips"
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
      Height          =   300
      Left            =   9084
      TabIndex        =   43
      Top             =   2556
      Width           =   1596
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "W2 Information for"
      ForeColor       =   &H0080FFFF&
      Height          =   300
      Left            =   732
      TabIndex        =   58
      Top             =   96
      Width           =   7212
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   567.854
      X2              =   11020.16
      Y1              =   5331.167
      Y2              =   5331.167
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      Height          =   7215
      Left            =   570
      Top             =   480
      Width           =   10470
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Third Party Sick Pay"
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
      Height          =   345
      Left            =   8280
      TabIndex        =   57
      Top             =   5670
      Width           =   1935
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Retirement Plan"
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
      Height          =   300
      Left            =   4860
      TabIndex        =   56
      Top             =   5676
      Width           =   1548
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "20. Local income taxes"
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
      Left            =   8700
      TabIndex        =   55
      Top             =   6924
      Width           =   2124
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "19. Local wages"
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
      Height          =   300
      Left            =   5388
      TabIndex        =   54
      Top             =   6924
      Width           =   1548
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "18. Locality Name"
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
      Height          =   300
      Left            =   1488
      TabIndex        =   53
      Top             =   6924
      Width           =   1692
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "17. State income taxes"
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
      Height          =   300
      Left            =   8688
      TabIndex        =   52
      Top             =   6240
      Width           =   2124
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "16. State wages, tips, etc"
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
      Height          =   300
      Left            =   4992
      TabIndex        =   51
      Top             =   6240
      Width           =   2316
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "15. State"
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
      Height          =   300
      Left            =   1248
      TabIndex        =   50
      Top             =   6240
      Width           =   876
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "13. Stat Emp"
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
      Height          =   300
      Left            =   1248
      TabIndex        =   49
      Top             =   5676
      Width           =   1212
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "14. Other"
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
      Height          =   300
      Left            =   7440
      TabIndex        =   48
      Top             =   4080
      Width           =   930
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "12. See instrs. for box 12"
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
      Height          =   300
      Left            =   705
      TabIndex        =   47
      Top             =   4125
      Width           =   2370
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "11. Nonqualified plans"
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
      Height          =   300
      Left            =   6144
      TabIndex        =   46
      Top             =   3276
      Width           =   2124
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "10. Dependent care benefits"
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
      Height          =   300
      Left            =   3324
      TabIndex        =   45
      Top             =   3276
      Width           =   2604
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "9. Advance EIC payment"
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
      Height          =   300
      Left            =   828
      TabIndex        =   44
      Top             =   3276
      Width           =   2316
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "7. Social Security tips"
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
      Height          =   300
      Left            =   6192
      TabIndex        =   42
      Top             =   2556
      Width           =   2028
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "6. Medicare tax w/h"
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
      Height          =   300
      Left            =   3612
      TabIndex        =   41
      Top             =   2556
      Width           =   1884
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "5. Medicare wages and tips"
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
      Left            =   816
      TabIndex        =   40
      Top             =   2556
      Width           =   2556
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "4. Social Security w/h"
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
      Height          =   300
      Left            =   8748
      TabIndex        =   39
      Top             =   1884
      Width           =   2028
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "3. Social Security wages"
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
      Height          =   300
      Left            =   6048
      TabIndex        =   38
      Top             =   1884
      Width           =   2268
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "2. Federal income tax w/h"
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
      Height          =   300
      Left            =   3408
      TabIndex        =   37
      Top             =   1884
      Width           =   2364
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "1. Wages, tips, other comp"
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
      Height          =   300
      Left            =   768
      TabIndex        =   4
      Top             =   1884
      Width           =   2460
   End
End
Attribute VB_Name = "frmW2EmpInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Over As clsTextBoxOverRider
Private Temp_Class As Resize_Class

Private Sub chkRetirePlan_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub chkStatEmp_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub chkThirdParty3_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub cmdExit_Click()
  Dim DoWhatFlag As SaveChangeOptions1
  Dim changeFlag As Boolean
  Dim W2Handle As Integer
  Dim W2FormRec(1) As W2FormType
  Dim x As Integer
  
  changeFlag = False
  
  OpenW2Info W2Handle
  Get W2Handle, RecNum, W2FormRec(1)
  Close W2Handle
  
  If fpcurrWagesTips1.Text <> W2FormRec(1).FEDWAGE Then
    changeFlag = True
    fpcurrWagesTips1.SetFocus
    GoTo changeFound
  End If
  
  If fpcurrFedIncTax2.Text <> W2FormRec(1).FEDTAXWH Then
    changeFlag = True
    fpcurrFedIncTax2.SetFocus
    GoTo changeFound
  End If
  
  If fpcurrSocSecWages3.Text <> W2FormRec(1).SOCWAGE Then
    changeFlag = True
    fpcurrSocSecWages3.SetFocus
    GoTo changeFound
  End If
  If fpcurrSocSecWH4.Text <> W2FormRec(1).SOCTAXWH Then
    changeFlag = True
    fpcurrSocSecWH4.SetFocus
    GoTo changeFound
  End If
  If fpcurrMedWages5.Text <> W2FormRec(1).MedWages Then
    changeFlag = True
    fpcurrMedWages5.SetFocus
    GoTo changeFound
  End If
  If fpcurrMedTaxWH6.Text <> W2FormRec(1).MEDTAXWH Then
    changeFlag = True
    fpcurrMedTaxWH6.SetFocus
    GoTo changeFound
  End If
  If fpcurrSocSecTips7.Text <> W2FormRec(1).SocTips Then
    changeFlag = True
    fpcurrSocSecTips7.SetFocus
    GoTo changeFound
  End If
  If fpcurrAlloTips8.Text <> W2FormRec(1).ALLOCTIP Then
    changeFlag = True
    fpcurrAlloTips8.SetFocus
    GoTo changeFound
  End If
  If fpcurrAdvanceEIC9.Text <> W2FormRec(1).AdvEIC Then
    changeFlag = True
    fpcurrAdvanceEIC9.SetFocus
    GoTo changeFound
  End If
  If fpcurrDepCareBene10.Text <> W2FormRec(1).DEPNDCAR Then
    changeFlag = True
    fpcurrDepCareBene10.SetFocus
    GoTo changeFound
  End If
  If fpcurrNonQPlans11.Text <> W2FormRec(1).NQPLAN Then
    changeFlag = True
    fpcurrNonQPlans11.SetFocus
    GoTo changeFound
  End If
  
  If fpcurrSeeInstr12a.Text <> W2FormRec(1).BOX13AMT Then
    changeFlag = True
    fpcurrSeeInstr12a.SetFocus
    GoTo changeFound
  End If
  If QPTrim$(fptxtBox12a.Text) <> QPTrim$(W2FormRec(1).BOX13TXt) Then
    changeFlag = True
    fptxtBox12a.SetFocus
    GoTo changeFound
  End If
  
  If fpcurrSeeInstr12b.Text <> W2FormRec(1).BOX13AM1 Then
    changeFlag = True
    fpcurrSeeInstr12b.SetFocus
    GoTo changeFound
  End If
  If QPTrim$(fptxtBox12b.Text) <> QPTrim$(W2FormRec(1).BOX13TX1) Then
    changeFlag = True
    fptxtBox12b.SetFocus
    GoTo changeFound
  End If
  
  If fpcurrSeeInstr12c.Text <> W2FormRec(1).BOX13AM2 Then
    changeFlag = True
    fpcurrSeeInstr12c.SetFocus
    GoTo changeFound
  End If
  If QPTrim$(fptxtBox12c.Text) <> QPTrim$(W2FormRec(1).BOX13TX2) Then
    changeFlag = True
    fptxtBox12c.SetFocus
    GoTo changeFound
  End If
  
  If fpcurrSeeInstr12d.Text <> W2FormRec(1).BOX13AM3 Then
    changeFlag = True
    fpcurrSeeInstr12d.SetFocus
    GoTo changeFound
  End If
  If QPTrim$(fptxtBox12d.Text) <> QPTrim$(W2FormRec(1).BOX13TX3) Then
    changeFlag = True
    fptxtBox12d.SetFocus
    GoTo changeFound
  End If
  
  If fpcurrOthera.Text <> W2FormRec(1).BOX14AMT Then
    changeFlag = True
    fpcurrOthera.SetFocus
    GoTo changeFound
  End If
  If QPTrim$(fptxtOthera.Text) <> QPTrim$(W2FormRec(1).BOX14TXT) Then
    changeFlag = True
    fptxtOthera.SetFocus
    GoTo changeFound
  End If
  If fpcurrOtherb.Text <> W2FormRec(1).BOX14AM1 Then
    changeFlag = True
    fpcurrOtherb.SetFocus
    GoTo changeFound
  End If
  If QPTrim$(fptxtOtherb.Text) <> QPTrim$(W2FormRec(1).BOX14TX1) Then
    changeFlag = True
    fptxtOtherb.SetFocus
    GoTo changeFound
  End If
  
  If chkStatEmp.Value = 1 Then
    If W2FormRec(1).BOX15A <> "X" Then
      changeFlag = True
      chkStatEmp.SetFocus
      GoTo changeFound
    End If
  ElseIf chkStatEmp.Value <> 1 Then
    If W2FormRec(1).BOX15A = "X" Then
      changeFlag = True
      chkStatEmp.SetFocus
      GoTo changeFound
    End If
  End If
  
  If chkRetirePlan.Value = 1 Then
    If W2FormRec(1).BOX15c <> "X" Then
      changeFlag = True
      chkRetirePlan.SetFocus
      GoTo changeFound
    End If
  ElseIf chkRetirePlan.Value <> 1 Then
    If W2FormRec(1).BOX15c = "X" Then
      changeFlag = True
      chkRetirePlan.SetFocus
      GoTo changeFound
    End If
  End If
  
  If chkThirdParty3.Value = 1 Then
    If W2FormRec(1).BOX15G <> "X" Then
      changeFlag = True
      chkThirdParty3.SetFocus
      GoTo changeFound
    End If
  ElseIf chkThirdParty3.Value <> 1 Then
    If W2FormRec(1).BOX15G = "X" Then
      changeFlag = True
      chkThirdParty3.SetFocus
      GoTo changeFound
    End If
  End If
  
  If QPTrim$(fptxtState16.Text) <> QPTrim$(W2FormRec(1).State) Then
    changeFlag = True
    fptxtState16.SetFocus
    GoTo changeFound
  End If
  If fpcurrStateWages17.Text <> W2FormRec(1).STAWAGE Then
    changeFlag = True
    fpcurrStateWages17.SetFocus
    GoTo changeFound
  End If
  If fpcurrStateIncTax18.Text <> W2FormRec(1).STATAXWH Then
    changeFlag = True
    fpcurrStateIncTax18.SetFocus
    GoTo changeFound
  End If
  If QPTrim$(fptxtLocality19.Text) <> QPTrim$(W2FormRec(1).LOCALNAM) Then
    changeFlag = True
    fptxtLocality19.SetFocus
    GoTo changeFound
  End If
  If fpcurrLocalWages20.Text <> W2FormRec(1).LocWage Then
    changeFlag = True
    fpcurrLocalWages20.SetFocus
    GoTo changeFound
  End If
  If fpcurrLocalIncTax21.Text <> W2FormRec(1).LOCALTAX Then
    changeFlag = True
    fpcurrLocalIncTax21.SetFocus
    GoTo changeFound
  End If
  
changeFound:
  If changeFlag = True Then
    DoWhatFlag = PromptSaveChanges(Me)
    Select Case DoWhatFlag
    Case SaveChangeOptions1.scoSaveChanges 'save changes
      Call cmdSave_Click
    Case SaveChangeOptions1.scoReviewChanges 'review is just bringing back the current form
      Exit Sub
    Case SaveChangeOptions1.scoAbandonChanges 'abandon
      frmW2Processing.Show
      DoEvents
      Unload frmW2EmpInfo
    Case Else:
       'Do nothing because we don't know about any options except
       'save, review or abandon...used as a placeholder for adding
       'other options at a later date
    End Select
  Else
    changeFlag = False
    frmEditReviewEmpW2.Show
'    frmW2Processing.Show
    DoEvents
    Unload frmW2EmpInfo
  End If
  
End Sub

Private Sub cmdPrintForm_Click()
  PrintForm
End Sub

Private Sub cmdSave_Click()
  Dim W2FormRecLen As Integer
  Dim WHandle As Integer
  Dim W2RWRec As W2ElectronicSubRW
  Dim RWHandle As Integer
  Dim OHandle As Integer
  Dim W2RORec As W2ElectronicSubRO
  Dim ROHandle As Integer
  Dim EmpRec2 As EmpData2Type
  Dim EHandle As Integer
  Dim NumOfORecs As Integer
  Dim x As Integer
  
  ReDim W2FormRec(1) As W2FormType

  W2FormRecLen = Len(W2FormRec(1))
  
  W2FormRec(1).FEDWAGE = fpcurrWagesTips1.Text
  W2RWRec.WageTips = Currency2String(W2FormRec(1).FEDWAGE, 11) '12/03
  
  W2FormRec(1).FEDTAXWH = fpcurrFedIncTax2.Text
  W2RWRec.FedTax = Currency2String(W2FormRec(1).FEDTAXWH, 11) '12/03
  
  W2FormRec(1).SOCWAGE = fpcurrSocSecWages3.Text
  W2RWRec.SSWages = Currency2String(W2FormRec(1).SOCWAGE, 11) '12/03
  
  W2FormRec(1).SOCTAXWH = fpcurrSocSecWH4.Text
  W2RWRec.SSTax = Currency2String(W2FormRec(1).SOCTAXWH, 11) '12/03
  
  W2FormRec(1).MedWages = fpcurrMedWages5.Text
  W2RWRec.MedWages = Currency2String(W2FormRec(1).MedWages, 11) '12/03
  
  W2FormRec(1).MEDTAXWH = fpcurrMedTaxWH6.Text
  W2RWRec.MedTax = Currency2String(W2FormRec(1).MEDTAXWH, 11) '12/03
  
  W2FormRec(1).SocTips = fpcurrSocSecTips7.Text
  W2RWRec.SSTips = Currency2String(fpcurrSocSecTips7.Text, 11)
  
  W2FormRec(1).ALLOCTIP = fpcurrAlloTips8.Text
  If W2FormRec(1).ALLOCTIP > 0 Then
    If Exist("PRDATA\W2ESUBRO.DAT") Then
      OpenW2ESubRO OHandle
      NumOfORecs = (LOF(OHandle) / Len(W2RORec))
      For x = 1 To NumOfORecs
        Get OHandle, x, W2RORec
        If RecNum = W2RORec.RecNum Then
'          W2RORec.AdoptionX = "00000000000"
          W2RORec.AllocTips = Currency2String(W2FormRec(1).ALLOCTIP, 11)
'          W2RORec.MedSavings = "00000000000"
'          W2RORec.RetAcct = "00000000000"
'          W2RORec.TaxOnTips = "00000000000"
'          W2RORec.UnMedLife = "00000000000"
'          W2RORec.UnSSLife = "00000000000"
          Put OHandle, x, W2RORec
          W2RWRec.RONum = CStr(x) 'this is saved as RW not RO
          Exit For
        End If
      Next x
      If x > NumOfORecs Then
        W2RORec.AdoptionX = "00000000000"
        W2RORec.AllocTips = Currency2String(W2FormRec(1).ALLOCTIP, 11)
        W2RORec.MedSavings = "00000000000"
        W2RORec.RetAcct = "00000000000"
        W2RORec.TaxOnTips = "00000000000"
        W2RORec.UnMedLife = "00000000000"
        W2RORec.UnSSLife = "00000000000"
        W2RORec.RecNum = RecNum
        Put OHandle, NumOfORecs + 1, W2RORec
        W2RWRec.RONum = CStr(NumOfORecs + 1) 'this is saved as RW not RO
      End If
    Else
      OpenW2ESubRO OHandle
      W2RORec.AdoptionX = "00000000000"
      W2RORec.AllocTips = Currency2String(W2FormRec(1).ALLOCTIP, 11)
      W2RORec.MedSavings = "00000000000"
      W2RORec.RetAcct = "00000000000"
      W2RORec.TaxOnTips = "00000000000"
      W2RORec.UnMedLife = "00000000000"
      W2RORec.UnSSLife = "00000000000"
      W2RORec.RecNum = RecNum
      Put OHandle, 1, W2RORec
      W2RWRec.RONum = "1" 'this is saved as RW not RO
    End If
  End If
  Close OHandle
  W2FormRec(1).AdvEIC = fpcurrAdvanceEIC9.Text
  W2RWRec.AdvEIC = Currency2String(W2FormRec(1).AdvEIC, 11)
  
  W2FormRec(1).DEPNDCAR = fpcurrDepCareBene10.Text
  W2RWRec.DepCare = Currency2String(W2FormRec(1).DEPNDCAR, 11)
  
  W2FormRec(1).NQPLAN = fpcurrNonQPlans11.Text
  W2RWRec.NQPlan457 = Currency2String(W2FormRec(1).NQPLAN, 11)
  
  W2FormRec(1).BOX13AMT = fpcurrSeeInstr12a.Text
  W2FormRec(1).BOX13TXt = QPTrim$(fptxtBox12a.Text)
  W2FormRec(1).BOX13AM1 = fpcurrSeeInstr12b.Text
  W2FormRec(1).BOX13TX1 = QPTrim$(fptxtBox12b.Text)
  W2FormRec(1).BOX13AM2 = fpcurrSeeInstr12c.Text
  W2FormRec(1).BOX13TX2 = QPTrim$(fptxtBox12c.Text)
  W2FormRec(1).BOX13AM3 = fpcurrSeeInstr12d.Text
  W2FormRec(1).BOX13TX3 = QPTrim$(fptxtBox12d.Text)
  W2FormRec(1).BOX14AMT = fpcurrOthera.Text
  W2FormRec(1).BOX14TXT = QPTrim$(fptxtOthera.Text)
  W2FormRec(1).BOX14AM1 = fpcurrOtherb.Text
  W2FormRec(1).BOX14TX1 = QPTrim$(fptxtOtherb.Text)
  
  '12/03 from here to....
  Select Case QPTrim$(W2FormRec(1).BOX13TXt)
    Case "C"
      W2RWRec.LifeIns = Currency2String(W2FormRec(1).BOX13AMT, 11)
    Case "D"
      W2RWRec.Defr401k = Currency2String(W2FormRec(1).BOX13AMT, 11)
    Case "E"
      W2RWRec.Defr403b = Currency2String(W2FormRec(1).BOX13AMT, 11)
    Case "F"
      W2RWRec.Defr408k6 = Currency2String(W2FormRec(1).BOX13AMT, 11)
    Case "G"
      W2RWRec.Defr457b = Currency2String(W2FormRec(1).BOX13AMT, 11)
    Case "H"
      W2RWRec.Defr501c18D = Currency2String(W2FormRec(1).BOX13AMT, 11)
    Case "V"
      W2RWRec.NonStaStcks = Currency2String(W2FormRec(1).BOX13AMT, 11)
    Case "AA"
      W2RWRec.Roth401K = Currency2String(W2FormRec(1).BOX13AMT, 11) 'added 1/22/07
  End Select
  
  Select Case QPTrim$(W2FormRec(1).BOX13TX1)
    Case "C"
      W2RWRec.LifeIns = Currency2String(W2FormRec(1).BOX13AM1, 11)
    Case "D"
      W2RWRec.Defr401k = Currency2String(W2FormRec(1).BOX13AM1, 11)
    Case "E"
      W2RWRec.Defr403b = Currency2String(W2FormRec(1).BOX13AM1, 11)
    Case "F"
      W2RWRec.Defr408k6 = Currency2String(W2FormRec(1).BOX13AM1, 11)
    Case "G"
      W2RWRec.Defr457b = Currency2String(W2FormRec(1).BOX13AM1, 11)
    Case "H"
      W2RWRec.Defr501c18D = Currency2String(W2FormRec(1).BOX13AM1, 11)
    Case "V"
      W2RWRec.NonStaStcks = Currency2String(W2FormRec(1).BOX13AM1, 11)
    Case "AA"
      W2RWRec.Roth401K = Currency2String(W2FormRec(1).BOX13AM1, 11) 'added 1/22/07
  End Select
  
  Select Case QPTrim$(W2FormRec(1).BOX13TX2) 'added fall 04
    Case "C"
      W2RWRec.LifeIns = Currency2String(W2FormRec(1).BOX13AM2, 11)
    Case "D"
      W2RWRec.Defr401k = Currency2String(W2FormRec(1).BOX13AM2, 11)
    Case "E"
      W2RWRec.Defr403b = Currency2String(W2FormRec(1).BOX13AM2, 11)
    Case "F"
      W2RWRec.Defr408k6 = Currency2String(W2FormRec(1).BOX13AM2, 11)
    Case "G"
      W2RWRec.Defr457b = Currency2String(W2FormRec(1).BOX13AM2, 11)
    Case "H"
      W2RWRec.Defr501c18D = Currency2String(W2FormRec(1).BOX13AM2, 11)
    Case "V"
      W2RWRec.NonStaStcks = Currency2String(W2FormRec(1).BOX13AM2, 11)
    Case "AA"
      W2RWRec.Roth401K = Currency2String(W2FormRec(1).BOX13AM2, 11) 'added 1/22/07
  End Select
  
  Select Case QPTrim$(W2FormRec(1).BOX13TX3) 'added fall 04
    Case "C"
      W2RWRec.LifeIns = Currency2String(W2FormRec(1).BOX13AM3, 11)
    Case "D"
      W2RWRec.Defr401k = Currency2String(W2FormRec(1).BOX13AM3, 11)
    Case "E"
      W2RWRec.Defr403b = Currency2String(W2FormRec(1).BOX13AM3, 11)
    Case "F"
      W2RWRec.Defr408k6 = Currency2String(W2FormRec(1).BOX13AM3, 11)
    Case "G"
      W2RWRec.Defr457b = Currency2String(W2FormRec(1).BOX13AM3, 11)
    Case "H"
      W2RWRec.Defr501c18D = Currency2String(W2FormRec(1).BOX13AM3, 11)
    Case "V"
      W2RWRec.NonStaStcks = Currency2String(W2FormRec(1).BOX13AM3, 11)
    Case "AA"
      W2RWRec.Roth401K = Currency2String(W2FormRec(1).BOX13AM3, 11) 'added 1/22/07
  End Select
  
  Select Case QPTrim$(W2FormRec(1).BOX14TXT)
    Case "C"
      W2RWRec.LifeIns = Currency2String(W2FormRec(1).BOX14AMT, 11)
    Case "D"
      W2RWRec.Defr401k = Currency2String(W2FormRec(1).BOX14AMT, 11)
    Case "E"
      W2RWRec.Defr403b = Currency2String(W2FormRec(1).BOX14AMT, 11)
    Case "F"
      W2RWRec.Defr408k6 = Currency2String(W2FormRec(1).BOX14AMT, 11)
    Case "G"
      W2RWRec.Defr457b = Currency2String(W2FormRec(1).BOX14AMT, 11)
    Case "H"
      W2RWRec.Defr501c18D = Currency2String(W2FormRec(1).BOX14AMT, 11)
    Case "V"
      W2RWRec.NonStaStcks = Currency2String(W2FormRec(1).BOX14AMT, 11)
    Case "AA"
      W2RWRec.Roth401K = Currency2String(W2FormRec(1).BOX14AMT, 11) 'added 1/22/07
  End Select
  
  Select Case QPTrim$(W2FormRec(1).BOX14TX1)
    Case "C"
      W2RWRec.LifeIns = Currency2String(W2FormRec(1).BOX14AM1, 11)
    Case "D"
      W2RWRec.Defr401k = Currency2String(W2FormRec(1).BOX14AM1, 11)
    Case "E"
      W2RWRec.Defr403b = Currency2String(W2FormRec(1).BOX14AM1, 11)
    Case "F"
      W2RWRec.Defr408k6 = Currency2String(W2FormRec(1).BOX14AM1, 11)
    Case "G"
      W2RWRec.Defr457b = Currency2String(W2FormRec(1).BOX14AM1, 11)
    Case "H"
      W2RWRec.Defr501c18D = Currency2String(W2FormRec(1).BOX14AM1, 11)
    Case "V"
      W2RWRec.NonStaStcks = Currency2String(W2FormRec(1).BOX14AM1, 11)
    Case "AA"
      W2RWRec.Roth401K = Currency2String(W2FormRec(1).BOX14AM1, 11) 'added 1/22/07
  End Select
  
  If Len(QPTrim$(W2RWRec.LifeIns)) = 0 Then
    W2RWRec.LifeIns = ""
  End If
  If Len(QPTrim$(W2RWRec.Defr401k)) = 0 Then
    W2RWRec.Defr401k = "0"
  End If
  If Len(QPTrim$(W2RWRec.Defr403b)) = 0 Then
    W2RWRec.Defr403b = ""
  End If
  If Len(QPTrim$(W2RWRec.Defr408k6)) = 0 Then
    W2RWRec.Defr408k6 = ""
  End If
  If Len(QPTrim$(W2RWRec.Defr457b)) = 0 Then
    W2RWRec.Defr457b = ""
  End If
  If Len(QPTrim$(W2RWRec.Defr501c18D)) = 0 Then
    W2RWRec.Defr501c18D = ""
  End If
  If Len(QPTrim$(W2RWRec.NonStaStcks)) = 0 Then
    W2RWRec.NonStaStcks = ""
  End If
  If Len(QPTrim$(W2RWRec.Roth401K)) = 0 Then 'added 1/22/07
    W2RWRec.Roth401K = "0"
  End If
  
  '...here
  
  If chkStatEmp.Value = 1 Then
    W2FormRec(1).BOX15A = "X"
    W2RWRec.StatuEmp = "1"
  Else
    W2FormRec(1).BOX15A = ""
    W2RWRec.StatuEmp = "0"
  End If
  
  If chkRetirePlan.Value = 1 Then
    W2FormRec(1).BOX15c = "X"
    W2RWRec.RetPlan = "1" '12/03
  Else
    W2FormRec(1).BOX15c = ""
    W2RWRec.RetPlan = "0" '12/03
  End If
  
  W2FormRec(1).BOX14AM1 = W2FormRec(1).BOX14AM1
  W2FormRec(1).BOX14TX1 = W2FormRec(1).BOX14TX1
  If chkThirdParty3.Value = 1 Then
    W2FormRec(1).BOX15G = "X"
    W2RWRec.ThrdSckPay = "1"
'    W2RWRec.ThrdSckAmt = "0"
  Else
    W2FormRec(1).BOX15G = ""
    W2RWRec.ThrdSckPay = "0"
  End If
  W2RWRec.ThrdSckAmt = "0"
    
    W2FormRec(1).BOX15B = W2FormRec(1).BOX15B
    W2FormRec(1).BOX15c = W2FormRec(1).BOX15c
    W2FormRec(1).BOX15D = W2FormRec(1).BOX15D
    W2FormRec(1).BOX15E = W2FormRec(1).BOX15E
    W2FormRec(1).BOX15F = W2FormRec(1).BOX15F
    W2FormRec(1).BOX15G = W2FormRec(1).BOX15G
  W2FormRec(1).State = fptxtState16.Text
  W2FormRec(1).STAWAGE = fpcurrStateWages17.Text
  W2FormRec(1).STATAXWH = fpcurrStateIncTax18.Text
  W2FormRec(1).LOCALNAM = fptxtLocality19.Text
  W2FormRec(1).LocWage = fpcurrLocalWages20.Text
  W2FormRec(1).LOCALTAX = fpcurrLocalIncTax21.Text
  
  OpenW2Info WHandle
  Put WHandle, RecNum, W2FormRec(1)
  Close WHandle
  
  OpenEmpData2File EHandle
  Get EHandle, RecNum, EmpRec2
  Close EHandle
  W2RWRec.EmpSSN = ReplaceString(EmpRec2.EmpSSN, "-", "")
  W2RWRec.EmpFName = fptxtFirstName.Text
  W2RWRec.EmpLName = fptxtLastName.Text
  W2RWRec.EmpMName = fptxtMiddleName.Text
  W2RWRec.EmpSuffix = fptxtSuffix.Text
  W2RWRec.EmpAdd1 = QPTrim$(EmpRec2.EmpAddr1)
  W2RWRec.EmpAdd2 = QPTrim$(EmpRec2.EMPADDR2)
  W2RWRec.EmpCity = QPTrim$(EmpRec2.EmpCity)
  W2RWRec.EmpState = QPTrim$(EmpRec2.EmpState)
  W2RWRec.EmpZip = Mid(EmpRec2.EmpZip, 1, 5)
  W2RWRec.EmpZipX = Mid(EmpRec2.EmpZip, 7, 4)
  W2RWRec.NQPNot457 = ""
  OpenW2ESubRW RWHandle '12/03
  Put RWHandle, RecNum, W2RWRec
  Close RWHandle
  
  MsgBox "Your information has been saved."
'  frmW2Processing.Show
  frmEditReviewEmpW2.Show
  DoEvents
  Unload frmW2EmpInfo
  MainLog ("Employee W2 data saved.")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown, vbKeyReturn:
      SendKeys "{Tab}"
      KeyCode = 0
    Case vbKeyEscape:
      SendKeys "%x"
      KeyCode = 0
    Case vbKeyF10:
      SendKeys "%S"
      KeyCode = 0
    Case vbKeyF11:
      SendKeys "%P"
      KeyCode = 0
    Case Else:
  End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub Form_Load()
  Set Over = New clsTextBoxOverRider
  Over.OverRide Me
  Set Temp_Class = New Resize_Class
  Temp_Class.InitResizeClass Me
  Call LoadThisForm
  Me.HelpContextID = hlpEditReview
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    Me.Visible = False
    Temp_Class.ResizeControls Me
    Me.Visible = True
    Me.SetFocus
  End If
End Sub

Private Sub LoadThisForm()
  Dim Emp2Rec As EmpData2Type
  Dim EHandle As Integer
  Dim W2FormRecLen As Integer
  Dim WHandle As Integer
  Dim W2RWRec As W2ElectronicSubRW
  Dim RWHandle As Integer
  Dim NumOfRWRecs As Integer
  Dim Extract As Boolean
  
  Extract = True
  If Exist("PRDATA\W2ESUBRW.DAT") Then
    OpenW2ESubRW RWHandle
    NumOfRWRecs = LOF(RWHandle) / Len(W2RWRec)
    If NumOfRWRecs = 0 Then
      MsgBox "There are no employee W2 records saved for processing. Be sure to extract this year's W2 data before continuing."
      Close RWHandle
      Exit Sub
    End If
    Get RWHandle, RecNum, W2RWRec
    Close RWHandle
    fptxtFirstName.Text = QPTrim$(W2RWRec.EmpFName)
    fptxtMiddleName.Text = QPTrim$(W2RWRec.EmpMName)
    fptxtLastName.Text = QPTrim$(W2RWRec.EmpLName)
    fptxtSuffix.Text = QPTrim$(W2RWRec.EmpSuffix)
  End If
  
  OpenEmpData2File EHandle
  Get EHandle, RecNum, Emp2Rec
  Label26.Caption = "W2 Information For: " + QPTrim$(Emp2Rec.EmpFName) & " " & QPTrim$(Emp2Rec.EmpLName)
  MainLog ("W2 personal information for " + QPTrim(Emp2Rec.EmpFName) + " " + QPTrim$(Emp2Rec.EmpLName) + " opened.")
  Close EHandle

  ReDim W2FormRec(1) As W2FormType

  W2FormRecLen = Len(W2FormRec(1))

  OpenW2Info WHandle
  Get WHandle, RecNum, W2FormRec(1)
  
  Close WHandle
  
  If Not IsNumeric(W2FormRec(1).FEDWAGE) Or InStr(CStr(W2FormRec(1).FEDWAGE), "E") > 0 Then Extract = False
  fpcurrWagesTips1.Text = W2FormRec(1).FEDWAGE
  If Not IsNumeric(W2FormRec(1).FEDTAXWH) Or InStr(CStr(W2FormRec(1).FEDTAXWH), "E") > 0 Then Extract = False
  fpcurrFedIncTax2.Text = W2FormRec(1).FEDTAXWH
  If Not IsNumeric(W2FormRec(1).SOCWAGE) Or InStr(CStr(W2FormRec(1).SOCWAGE), "E") > 0 Then Extract = False
  fpcurrSocSecWages3.Text = W2FormRec(1).SOCWAGE
  If Not IsNumeric(W2FormRec(1).SOCTAXWH) Or InStr(CStr(W2FormRec(1).SOCTAXWH), "E") > 0 Then Extract = False
  fpcurrSocSecWH4.Text = W2FormRec(1).SOCTAXWH
  If Not IsNumeric(W2FormRec(1).MedWages) Or InStr(CStr(W2FormRec(1).MedWages), "E") > 0 Then Extract = False
  fpcurrMedWages5.Text = W2FormRec(1).MedWages
  If Not IsNumeric(W2FormRec(1).MEDTAXWH) Or InStr(CStr(W2FormRec(1).MEDTAXWH), "E") > 0 Then Extract = False
  fpcurrMedTaxWH6.Text = W2FormRec(1).MEDTAXWH
  If Not IsNumeric(W2FormRec(1).SocTips) Or InStr(CStr(W2FormRec(1).SocTips), "E") > 0 Then Extract = False
  fpcurrSocSecTips7.Text = W2FormRec(1).SocTips
  If Not IsNumeric(W2FormRec(1).ALLOCTIP) Or InStr(CStr(W2FormRec(1).ALLOCTIP), "E") > 0 Then Extract = False
  fpcurrAlloTips8.Text = W2FormRec(1).ALLOCTIP
  If Not IsNumeric(W2FormRec(1).AdvEIC) Or InStr(CStr(W2FormRec(1).AdvEIC), "E") > 0 Then Extract = False
  fpcurrAdvanceEIC9.Text = W2FormRec(1).AdvEIC
  If Not IsNumeric(W2FormRec(1).DEPNDCAR) Or InStr(CStr(W2FormRec(1).DEPNDCAR), "E") > 0 Then Extract = False
  fpcurrDepCareBene10.Text = W2FormRec(1).DEPNDCAR
  If Not IsNumeric(W2FormRec(1).NQPLAN) Or InStr(CStr(W2FormRec(1).NQPLAN), "E") > 0 Then Extract = False
  fpcurrNonQPlans11.Text = W2FormRec(1).NQPLAN
  If Not IsNumeric(W2FormRec(1).BOX13AMT) Or InStr(CStr(W2FormRec(1).BOX13AMT), "E") > 0 Then Extract = False
  fpcurrSeeInstr12a.Text = W2FormRec(1).BOX13AMT
  fptxtBox12a.Text = QPTrim$(W2FormRec(1).BOX13TXt)
  If Not IsNumeric(W2FormRec(1).BOX13AM1) Or InStr(CStr(W2FormRec(1).BOX13AM1), "E") > 0 Then Extract = False
  fpcurrSeeInstr12b.Text = W2FormRec(1).BOX13AM1
  fptxtBox12b.Text = QPTrim$(W2FormRec(1).BOX13TX1)
  
  
  If Not IsNumeric(W2FormRec(1).BOX13AM2) Or InStr(CStr(W2FormRec(1).BOX13AM2), "E") > 0 Then Extract = False
  fpcurrSeeInstr12c.Text = W2FormRec(1).BOX13AM2 'added fall 04
  fptxtBox12c.Text = QPTrim$(W2FormRec(1).BOX13TX2) 'added fall 04
  If Not IsNumeric(W2FormRec(1).BOX13AM3) Or InStr(CStr(W2FormRec(1).BOX13AM3), "E") > 0 Then Extract = False
  fpcurrSeeInstr12d.Text = W2FormRec(1).BOX13AM3 'added fall 04
  fptxtBox12d.Text = QPTrim$(W2FormRec(1).BOX13TX3) 'added fall 04
  
  If Not IsNumeric(W2FormRec(1).BOX14AMT) Or InStr(CStr(W2FormRec(1).BOX14AMT), "E") > 0 Then Extract = False
  fpcurrOthera.Text = W2FormRec(1).BOX14AMT
  fptxtOthera.Text = QPTrim$(W2FormRec(1).BOX14TXT)
  If Not IsNumeric(W2FormRec(1).BOX14AM1) Or InStr(CStr(W2FormRec(1).BOX14AM1), "E") > 0 Then Extract = False
  fpcurrOtherb.Text = W2FormRec(1).BOX14AM1
  fptxtOtherb.Text = QPTrim$(W2FormRec(1).BOX14TX1)
  If W2FormRec(1).BOX15A = "X" Then
    chkStatEmp.Value = 1
  Else
    chkStatEmp.Value = 0
  End If
  
  If W2FormRec(1).BOX15c = "X" Then
    chkRetirePlan.Value = 1
  Else
    chkRetirePlan.Value = 0
  End If
  
  If W2FormRec(1).BOX15G = "X" Then
    chkThirdParty3.Value = 1
  Else
    chkThirdParty3.Value = 0
  End If
'    W2FormRec(1).BOX15B = W2FormRec(1).BOX15B
'    W2FormRec(1).BOX15c = W2FormRec(1).BOX15c
'    W2FormRec(1).BOX15D = W2FormRec(1).BOX15D
'    W2FormRec(1).BOX15E = W2FormRec(1).BOX15E
'    W2FormRec(1).BOX15F = W2FormRec(1).BOX15F
'    W2FormRec(1).BOX15G = W2FormRec(1).BOX15G
  fptxtState16.Text = W2FormRec(1).State
  If Not IsNumeric(W2FormRec(1).STAWAGE) Or InStr(CStr(W2FormRec(1).STAWAGE), "E") > 0 Then Extract = False
  fpcurrStateWages17.Text = W2FormRec(1).STAWAGE
  If Not IsNumeric(W2FormRec(1).STATAXWH) Or InStr(CStr(W2FormRec(1).STATAXWH), "E") > 0 Then Extract = False
  fpcurrStateIncTax18.Text = W2FormRec(1).STATAXWH
  fptxtLocality19.Text = W2FormRec(1).LOCALNAM
  fpcurrLocalWages20.Text = W2FormRec(1).LocWage
  fpcurrLocalIncTax21.Text = W2FormRec(1).LOCALTAX
  
  If Extract = False Then
    frmW2Message.Label1.Caption = "Some of the currency amounts displayed are faulty. This is probably because extraction of the latest data has not yet taken place. Please extract the latest data to correct these amounts."
    frmW2Message.Label1.Top = 700
    frmW2Message.Show vbModal
  End If
  
  MainLog ("Employee W2 data screen loaded for " + QPTrim$(Emp2Rec.EmpFName) + " " + QPTrim$(Emp2Rec.EmpLName) + ".")
End Sub

Private Sub fpcurrAdvanceEIC9_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If
End Sub

Private Sub fpcurrAlloTips8_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fpcurrDepCareBene10_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fpcurrFedIncTax2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fpcurrLocalIncTax21_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    fptxtFirstName.SetFocus
    KeyCode = 0
  End If

End Sub

Private Sub fpcurrLocalWages20_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fpcurrMedTaxWH6_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fpcurrMedWages5_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fpcurrNonQPlans11_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fpcurrOthera_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fpcurrOtherb_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fpcurrSeeInstr12a_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fpcurrSeeInstr12b_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fpcurrSocSecTips7_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fpcurrSocSecWages3_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fpcurrSocSecWH4_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fpcurrStateIncTax18_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fpcurrStateWages17_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fpcurrWagesTips1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    fpcurrLocalIncTax21.SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If
End Sub

Private Sub fptxtBox12a_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fptxtBox12b_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fptxtLocality19_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fptxtOthera_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fptxtOtherb_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub fptxtState16_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    SendKeys "+{Tab}"
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    SendKeys "{Tab}"
    KeyCode = 0
  End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If ((UnloadMode = vbFormControlMenu)) Then
    If cmdExit.Enabled = False Then
      Cancel = True
    ElseIf MsgBox("Are You Sure You Wish To Close Program?", vbYesNo, "Close?") = vbNo Then
      Cancel = True
    Else
      Call UnloadAllFormsAndOpn(RegExit)
      MainLog ("Payroll.exe terminated via menu bar on frmW2EmpInfo.")
      End
    End If
  End If
End Sub

Private Sub fptxtSuffix_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then
    fptxtLastName.SetFocus
    KeyCode = 0
  ElseIf KeyCode = vbKeyDown Then
    fpcurrWagesTips1.SetFocus
    KeyCode = 0
  End If

End Sub

